from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import InvalidSessionIdException, WebDriverException
from urllib.parse import urlparse, parse_qs
import os
import pyotp
import time
import json
import xlwings as xw

# ─── CONFIG ───────────────────────────────────────────────────────────────────
CLIENT_ID = None
USER_ID = None
PASSWORD = None
TOTP_SECRET = None
LOGIN_URL = "https://trade.shoonya.com/OAuthlogin/investor-entry-level/login"
EXCEL_PATH = r"C:\python_trader\Finvasia_Trade_Terminal_v3.xlsm"
EXCEL_SHEET = "User_Credential"
EXCEL_CELL = "B13"

CREDENTIAL_CELL_MAP = {
    "client_id": "B2",
    "password": "B3",
    "totp_secret": "B5",
    "user_id": "B6",
}


def scan_network_for_code(driver):
    try:
        logs = driver.get_log("performance")
        for entry in logs:
            try:
                message = json.loads(entry["message"])["message"]
                if message.get("method") == "Network.requestWillBeSent":
                    url = message.get("params", {}).get("request", {}).get("url", "")
                    if "code=" in url:
                        parsed = urlparse(url)
                        code = parse_qs(parsed.query).get("code", [None])[0]
                        if code:
                            return url, code
            except Exception:
                continue
    except Exception:
        pass
    return None, None


def fast_fill(element, value):
    element.click()
    time.sleep(0.1)
    element.clear()
    element.send_keys(value)
    time.sleep(0.1)


def extract_code_from_url(url):
    parsed = urlparse(url)
    return parse_qs(parsed.query).get("code", [None])[0]


def find_open_workbook(excel_path):
    excel_path = os.path.abspath(excel_path).lower()
    for app in xw.apps:
        for wb in app.books:
            try:
                if os.path.abspath(wb.fullname).lower() == excel_path:
                    return app, wb
            except Exception:
                continue
    return None, None


def read_credentials_from_excel(excel_path, sheet_name):
    excel_path = os.path.abspath(excel_path)
    app = None
    workbook = None
    created_app = False
    workbook_already_open = False

    try:
        try:
            app, workbook = find_open_workbook(excel_path)
        except Exception:
            app = None
            workbook = None

        if workbook is not None:
            workbook_already_open = True
        else:
            if app is None:
                app = xw.App(visible=False, add_book=False)
                created_app = True
            workbook = app.books.open(excel_path)

        sheet = workbook.sheets[sheet_name]
        return {
            "client_id": sheet.range(CREDENTIAL_CELL_MAP["client_id"]).value,
            "password": sheet.range(CREDENTIAL_CELL_MAP["password"]).value,
            "totp_secret": sheet.range(CREDENTIAL_CELL_MAP["totp_secret"]).value,
            "user_id": sheet.range(CREDENTIAL_CELL_MAP["user_id"]).value,
        }
    finally:
        if workbook is not None and not workbook_already_open:
            workbook.close()
        if created_app and app is not None:
            app.quit()


def save_code_to_excel(auth_code, excel_path, sheet_name, cell="B13"):
    excel_path = os.path.abspath(excel_path)
    app = None
    workbook = None
    created_app = False
    workbook_already_open = False

    try:
        try:
            app, workbook = find_open_workbook(excel_path)
        except Exception:
            app = None
            workbook = None

        if workbook is not None:
            workbook_already_open = True
        else:
            if app is None:
                app = xw.App(visible=False, add_book=False)
                created_app = True
            workbook = app.books.open(excel_path)

        try:
            sheet = workbook.sheets[sheet_name]
        except Exception:
            sheet = workbook.sheets.add(sheet_name)

        sheet.range(cell).value = auth_code
        workbook.save()
        print(f"Auth code saved to Excel: {excel_path} -> {sheet_name}!{cell}")
    finally:
        if workbook is not None and not workbook_already_open:
            workbook.close()
        if created_app and app is not None:
            app.quit()


def main():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    try:
        credentials = read_credentials_from_excel(EXCEL_PATH, EXCEL_SHEET)
        CLIENT_ID = credentials.get("client_id")
        USER_ID = credentials.get("user_id")
        PASSWORD = credentials.get("password")
        TOTP_SECRET = credentials.get("totp_secret")

        print("Read credentials from Excel:")
        print(f"  client_id = {CLIENT_ID}")
        print(f"  user_id   = {USER_ID}")
        print(f"  password  = {'*' * len(PASSWORD) if PASSWORD else None}")
        print(f"  totp_secret = {'*' * len(TOTP_SECRET) if TOTP_SECRET else None}")

        if not (CLIENT_ID and USER_ID and PASSWORD and TOTP_SECRET):
            raise RuntimeError("Missing credentials in Excel. Ensure B2, B3, B5, and B6 are populated.")

        login_url = f"https://trade.shoonya.com/OAuthlogin/investor-entry-level/login?api_key={USER_ID}&route_to={CLIENT_ID}"
        print("Opening login page...")
        driver.get(login_url)

        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='password']")))
        time.sleep(1)

        all_inputs = driver.find_elements(By.CSS_SELECTOR, "input:not([type='hidden']):not([type='checkbox']):not([type='radio'])")
        visible_inputs = [inp for inp in all_inputs if inp.is_displayed()]
        if len(visible_inputs) < 3:
            raise RuntimeError("Could not find the expected login input fields.")

        fast_fill(visible_inputs[0], CLIENT_ID)
        fast_fill(visible_inputs[1], PASSWORD)

        otp_value = pyotp.TOTP(TOTP_SECRET).now()
        fast_fill(visible_inputs[2], otp_value)

        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='LOGIN']"))).click()
        print("Login submitted. Waiting for redirect...")

        start = time.time()
        auth_url = None
        auth_code = None

        while time.time() - start < 60:
            current_url = driver.current_url
            if current_url and current_url != login_url:
                if "code=" in current_url or "error=" in current_url:
                    auth_url = current_url
                    print(f"Redirected URL detected: {current_url}")
                    break
                else:
                    print(f"Current URL changed, still waiting for auth code: {current_url}")

            auth_url, auth_code = scan_network_for_code(driver)
            if auth_code:
                break

            time.sleep(0.5)

        if auth_url:
            print("Output URL:", auth_url)
            if not auth_code:
                auth_code = extract_code_from_url(auth_url)

        if auth_code:
            print("Extracted auth code:", auth_code)
            save_code_to_excel(auth_code, EXCEL_PATH, EXCEL_SHEET, EXCEL_CELL)
        else:
            print("Unable to extract auth code. Please verify the login flow and OTP.")

    except (InvalidSessionIdException, WebDriverException) as e:
        print(f"[ERROR] Browser issue: {e}")
    except Exception as e:
        print(f"[ERROR] {e}")
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
