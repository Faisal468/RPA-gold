"""Open the bank portals listed in config.ini [portals] in a visible browser.

Run:  python portals.py
"""
import configparser
import sys

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException

from main import BASE, SetupError

CONFIG = BASE / "config.ini"


def load_portals():
    parser = configparser.ConfigParser(interpolation=None)
    if not parser.read(CONFIG, encoding="utf-8") or "portals" not in parser:
        raise SetupError(f"{CONFIG} needs a [portals] section (name = url)")
    return {name: url.strip() for name, url in parser["portals"].items() if url.strip()}


def start_browser(keep_open=True):
    options = webdriver.EdgeOptions()  # Edge is on every Windows PC and is signed
    options.page_load_strategy = "eager"
    options.add_argument("--start-maximized")
    if keep_open:
        options.add_experimental_option("detach", True)  # browser stays open after the script ends
    return webdriver.Edge(options=options)


def open_portals(keep_open=True):
    """Open every portal in its own tab. Returns (driver, {name: window handle})."""
    portals = load_portals()
    if not portals:
        raise SetupError("No portals are listed under [portals] in config.ini")
    driver = start_browser(keep_open)
    driver.set_page_load_timeout(60)
    handles = {}
    for i, (name, url) in enumerate(portals.items()):
        if i > 0:
            driver.switch_to.new_window("tab")
        handles[name] = driver.current_window_handle
        print(f"Opening {name}: {url}")
        try:
            driver.get(url)
        except TimeoutException:
            print(f"  {name}: page is slow, it may still be loading")
        except WebDriverException as exc:
            print(f"  {name}: could not be reached ({str(exc).splitlines()[0]}). "
                  "Check your internet or VPN.")
    return driver, handles


def main():
    try:
        open_portals()
    except (SetupError, WebDriverException) as exc:
        print("Error:", str(exc).splitlines()[0])
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
