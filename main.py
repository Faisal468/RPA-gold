
import configparser
import json
import logging
import re
import sys
import time
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path 

import openpyxl
from copy import copy
from openpyxl.formula.translate import Translator
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

# When packed as an exe, config.ini, browser\ and the Excel log sit next to the exe.
BASE = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
log = logging.getLogger("Gold Rate")

class SetupError(Exception):
    """Something is wrong with config.ini or the browser files."""

class SiteError(Exception):
    """One site could not give a valid rate."""

def lines(value):
    return [ln.strip() for ln in value.splitlines() if ln.strip()]


def load_config(path=BASE / "config.ini"):
    parser = configparser.ConfigParser(interpolation=None)  # keeps % and \ as typed
    if not parser.read(path, encoding="utf-8"):
        raise SetupError(f"Cannot find {path}")
    try:
        for name in ("run", "browser", "target", "units", "validation"):
            parser[name]
        for site in site_names(parser):
            for key in ("name", "url", "rate_patterns"):
                if not parser[site].get(key):
                    raise SetupError(f"[{site}] is missing '{key}'")
    except KeyError as exc:
        raise SetupError(f"config.ini is missing the section {exc}") from exc

    for key in ("chrome_path", "driver_path"):
        file = BASE / parser["browser"][key]
        if not file.exists():
            raise SetupError(f"[browser] {key} points to a file that does not exist: {file}")
    return parser


def site_names(cfg):
    return [s.strip() for s in cfg["run"]["sites"].split(",") if s.strip()]


def start_browser(cfg):
    b = cfg["browser"]
    options = webdriver.ChromeOptions()
    options.binary_location = str(BASE / b["chrome_path"])
    options.page_load_strategy = "eager"  # do not wait for ads and trackers
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    if b.getboolean("headless"):
        options.add_argument("--headless=new")
    if b.getboolean("maximize"):
        options.add_argument("--start-maximized")

    service = Service(executable_path=str(BASE / b["driver_path"]))
    try:
        driver = webdriver.Chrome(service=service, options=options)
    except WebDriverException as exc:
        raise SetupError(
            "Chrome did not start. Check that chrome-win64 and chromedriver-win64 "
            f"are the same version.\n{str(exc).splitlines()[0]}"
        ) from exc
    driver.set_page_load_timeout(b.getint("page_timeout_seconds"))
    return driver


def open_tabs(driver, cfg):
    """Open every site in its own tab. Returns {site key: window handle}."""
    handles = {}
    for i, site in enumerate(site_names(cfg)):
        if i > 0:
            driver.switch_to.new_window("tab")
        handles[site] = driver.current_window_handle
        log.info("Opening %s", cfg[site]["url"])
        try:
            driver.get(cfg[site]["url"])
        except TimeoutException:
            log.warning("%s: page is slow, will read what has loaded", cfg[site]["name"])
        except WebDriverException as exc:
            log.error("%s: could not open page: %s", cfg[site]["name"], str(exc).splitlines()[0])
    return handles


def page_text(driver):
    return driver.find_element(By.TAG_NAME, "body").text


def find_number(patterns, text):
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return Decimal(match.group(1).replace(",", ""))
    return None


def read_values(text, site_cfg):
    """Return (rate per 10 gram, rate per tola or None) found in the page text."""
    rate = find_number(lines(site_cfg["rate_patterns"]), text)
    check = find_number(lines(site_cfg.get("check_patterns", "")), text)
    return rate, check


def validate(rate, tola, cfg):
    v, units = cfg["validation"], cfg["units"]
    low, high = Decimal(v["min_rate"]), Decimal(v["max_rate"])
    if not (low <= rate <= high):
        raise SiteError(f"Rate {rate:,} is outside the allowed range {low:,} to {high:,}")
    if tola is not None:
        expected = tola * 10 / Decimal(units["tola_grams"])
        gap = abs(rate - expected) / expected * 100
        if gap > Decimal(v["check_tolerance_pct"]):
            raise SiteError(
                f"Per-10-gram value {rate:,} does not match the per-tola value {tola:,} "
                f"(off by {gap:.1f}%). Probably the wrong column was read."
            )


def read_site(driver, handle, site, cfg):
    site_cfg, b = cfg[site], cfg["browser"]
    entry = {"source": site_cfg["name"], "url": site_cfg["url"], "status": "failed"}
    driver.switch_to.window(handle)

    attempts = b.getint("retries") + 1
    rate = tola = None
    for attempt in range(1, attempts + 1):
        deadline = time.time() + b.getint("page_timeout_seconds")
        text = ""
        while time.time() < deadline:  
            text = page_text(driver)
            rate, tola = read_values(text, site_cfg)
            if rate is not None:
                break
            time.sleep(1)
        if rate is not None:
            break
        log.warning("%s: rate not found (attempt %d/%d)", site_cfg["name"], attempt, attempts)
        if attempt < attempts:
            time.sleep(b.getint("retry_wait_seconds"))
            driver.refresh()

    try:
        if rate is None:
            raise SiteError("The 24K per-10-gram rate was not found on the page")
        validate(rate, tola, cfg)
        decimals = Decimal(1).scaleb(-cfg["target"].getint("decimals"))
        entry.update(
            status="ok",
            rate=float(rate.quantize(decimals, rounding=ROUND_HALF_UP)),
            tola_rate=float(tola) if tola is not None else None,
        )
    except SiteError as exc:
        entry["error"] = str(exc)
        log.error("%s: %s", site_cfg["name"], exc)
        if b.getboolean("save_debug_on_error"):
            save_debug(driver, site, text, cfg)
    return entry


def save_debug(driver, site, text, cfg):
    folder = BASE / cfg["run"]["output_folder"] / "debug"
    folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    (folder / f"{site}_{stamp}.txt").write_text(text, encoding="utf-8")
    driver.save_screenshot(str(folder / f"{site}_{stamp}.png"))
    log.info("Saved page text and screenshot to %s", folder)


def run(cfg):
    driver = start_browser(cfg)
    try:
        handles = open_tabs(driver, cfg)
        entries = [read_site(driver, handles[s], s, cfg) for s in site_names(cfg)]
        time.sleep(cfg["browser"].getint("close_delay_seconds"))
    finally:
        driver.quit()

    good = [e for e in entries if e["status"] == "ok"]
    warnings = []
    if len(good) >= 2:
        rates = [e["rate"] for e in good]
        gap = (max(rates) - min(rates)) / min(rates) * 100
        limit = float(cfg["validation"]["max_source_gap_pct"])
        if gap > limit:
            warnings.append(f"Sites differ by {gap:.1f}%, above the {limit:g}% limit")

    enough = len(good) >= cfg["run"].getint("min_sources_required")
    now = datetime.now().astimezone()
    result = {
        "run_date": now.date().isoformat(),
        "run_time": now.isoformat(timespec="seconds"),
        "basis": {k: cfg["target"][k] for k in ("karat", "unit", "currency")},
        "status": "ok" if enough else "failed",
        "sources_ok": len(good),
        "sources_total": len(entries),
        "warnings": warnings,
        "rates": entries,
    }
    out_dir = BASE / cfg["run"]["output_folder"]
    out_dir.mkdir(parents=True, exist_ok=True)
    out_file = out_dir / f"gold_rates_{result['run_date']}.json"
    out_file.write_text(json.dumps(result, indent=2), encoding="utf-8")
    log.info("Saved %s", out_file)
    return result, enough


def setup_logging(cfg=None):
    handlers = [logging.StreamHandler()]
    if cfg:
        log_file = BASE / cfg["run"]["log_file"]
        log_file.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))
    logging.basicConfig(level=logging.INFO, handlers=handlers, force=True,
                        format="%(asctime)s %(levelname)s %(message)s")


# Websites for open-only mode (no config.ini needed)
URLS = [
    "https://www.forex.pk/bullion-rates.php",
    "https://pakbiz.com/gold-rate/",
]


def set_forex_date_today(driver):
    """forex.pk does not update its date by itself: pick today's date and press Show."""
    today = datetime.now()
    for field, value in (("day", f"{today.day:02d}"), ("month", f"{today.month:02d}"),
                         ("year", str(today.year))):
        Select(driver.find_element(By.ID, field)).select_by_value(value)
    old_page = driver.find_element(By.TAG_NAME, "html")
    driver.find_element(By.ID, "btnShow").click()
    WebDriverWait(driver, 30).until(EC.staleness_of(old_page))


def read_forex(driver):
    text = WebDriverWait(driver, 30).until(
        lambda d: page_text(d) if re.search(r"per\s+10\s+Gram", page_text(d)) else None)
    # "Gold Price in Pakistan" table: 24K Gold column, per 10 Gram row
    return re.search(r"per\s+10\s+Gram\s+Rs\.?\s*([\d,]+)", text).group(1)


def read_pakbiz(driver):
    text = WebDriverWait(driver, 30).until(
        lambda d: page_text(d) if re.search(r"24K", page_text(d)) else None)
    return re.search(r"24K\s+Rs\.?\s*[\d,]+\s+Rs\.?\s*([\d,]+)", text).group(1)


LOG_FILE = BASE / "gold rate logs.xlsx"
LOG_SHEET = "Gold Rate Log"

def log_to_excel(forex, pakbiz, extra=None):
    """Write today's time and rates into today's row of the log sheet (adds the row if missing).

    Columns are found by their header names, so moving or inserting columns in Excel is safe.
    `extra` is {header: value}; a header that does not exist yet is added after the last one.
    """
    now = datetime.now()
    wb = openpyxl.load_workbook(LOG_FILE)
    ws = wb[LOG_SHEET] if LOG_SHEET in wb.sheetnames else wb.active
    header = next((r for r in range(1, 11)
                   if {"Date", "Forex rate"} <= {str(c.value).strip() for c in ws[r]}), None)
    if header is None:
        raise SetupError(f"{LOG_FILE.name}: cannot find the header row (Date, Forex rate, ...)")
    col = {str(c.value).strip(): c.column for c in ws[header] if c.value}
    missing = [h for h in ("Time", "Date", "Day", "Forex rate", "PakBiz rate") if h not in col]
    if missing:
        raise SetupError(f"{LOG_FILE.name}: is missing the column(s) {', '.join(missing)}")

    row = None
    last = header
    for r in range(header + 1, ws.max_row + 1):
        value = ws.cell(r, col["Date"]).value
        if value is None:
            break
        last = r
        if isinstance(value, datetime) and value.date() == now.date():
            row = r
    if row is None:  # new day: copy the look and formulas of the row above
        row = last + 1
        for c in range(1, ws.max_column + 1):
            src, dst = ws.cell(last, c), ws.cell(row, c)
            dst._style = copy(src._style)
            if isinstance(src.value, str) and src.value.startswith("="):
                dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
        ws.row_dimensions[row].height = ws.row_dimensions[last].height
        ws.cell(row, col["Date"], datetime(now.year, now.month, now.day))

    for c in ("Time", "Forex rate", "PakBiz rate"):  # keep the sheet's look if the cell is unformatted
        cell, above = ws.cell(row, col[c]), ws.cell(row - 1, col[c])
        if not cell.has_style and row - 1 > header:
            cell._style = copy(above._style)
    ws.cell(row, col["Time"], now.time().replace(microsecond=0))
    ws.cell(row, col["Forex rate"], int(forex.replace(",", "")))
    ws.cell(row, col["PakBiz rate"], int(pakbiz.replace(",", "")))
    for name, value in (extra or {}).items():
        if name not in col:
            col[name] = max(col.values()) + 1
            head, prev = ws.cell(header, col[name]), ws.cell(header, col[name] - 1)
            head.value, head._style = name, copy(prev._style)
            ws.column_dimensions[head.column_letter].width = 16
        cell = ws.cell(row, col[name], value)
        cell._style = copy(ws.cell(row, 4)._style) if isinstance(value, (int, float)) else cell._style
    try:
        wb.save(LOG_FILE)
    except PermissionError:
        raise SetupError(f"Cannot save {LOG_FILE.name}. Close it in Excel and run again.")
    log.info("Saved %s to row %d of %s", now.strftime("%A %d-%b-%Y %H:%M"), row, LOG_FILE.name)


def open_only():
    """Start the bundled Chrome, open both sites, set the forex date to today, read 24K gold."""
    options = webdriver.ChromeOptions()
    options.binary_location = str(BASE / "browser" / "chrome-win64" / "chrome.exe")
    options.page_load_strategy = "eager"
    options.add_argument("--start-maximized")
    service = Service(executable_path=str(BASE / "browser" / "chromedriver-win64" / "chromedriver.exe"))
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(60)
    try:
        tabs = []
        for i, url in enumerate(URLS):
            if i > 0:
                driver.switch_to.new_window("tab")
            tabs.append(driver.current_window_handle)
            log.info("Opening %s", url)
            try:
                driver.get(url)
            except TimeoutException:
                log.warning("%s: page is slow, continuing", url)

        driver.switch_to.window(tabs[0])
        print()
        print(f"Date: {datetime.now():%d %b %Y}  (24K gold, per 10 gram)")
        forex = read_forex(driver)
        driver.switch_to.window(tabs[1])
        pakbiz = read_pakbiz(driver)
        rates = [Decimal(forex.replace(",", "")), Decimal(pakbiz.replace(",", ""))]
        average = (sum(rates) / len(rates)).quantize(Decimal(1), rounding=ROUND_HALF_UP)
        print("Forex.pk :", forex)
        print("Pakbiz   :", pakbiz)
        print(f"Average  : {average:,}")
        log_to_excel(forex, pakbiz)
        input("Press Enter here to close the browser...")
    finally:
        driver.quit()
    return 0


def main():
    setup_logging()
    if "--scrape" not in sys.argv:  # default: open the sites, read rates, log to Excel
        if datetime.now().weekday() == 6:
            log.info("Today is %s. Sundays are skipped, nothing to do.", datetime.now().strftime("%A"))
            return 0
        try:
            return open_only()
        except SetupError as exc:
            log.error("%s", exc)
            return 2
    try:
        cfg = load_config()
        setup_logging(cfg)
        result, enough = run(cfg)
    except SetupError as exc:
        log.error("%s", exc)
        return 2

    print()
    for e in result["rates"]:
        shown = f"Rs. {e['rate']:,.2f}" if e["status"] == "ok" else f"FAILED - {e['error']}"
        print(f"{e['source']:<10} {shown}")
    for w in result["warnings"]:
        print("WARNING:", w)
    return 0 if enough else 1


if __name__ == "__main__":
    sys.exit(main())