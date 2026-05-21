import os
import re
import time
try:
    import tkinter as tk
    from tkinter import messagebox, filedialog
    GUI_AVAILABLE = True
except Exception:
    tk = None
    messagebox = None
    filedialog = None
    GUI_AVAILABLE = False
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import Workbook, load_workbook

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# =========================================================
# CONFIG
# =========================================================

URL = "https://aca-prod.accela.com/LEECO/Cap/CapHome.aspx?module=Permitting&TabName=Permitting"

OUTPUT_FILE = os.getenv("OUTPUT_FILE", "LehighAcres_City_FL_Permits_DateRange_Output.xlsx")

COUNTY_VALUE = "Lee"
STATE_VALUE = "FL"
CITY_VALUE = "Lehigh Acres"

WAIT_TIME = int(os.getenv("WAIT_TIME", "30"))

SEARCH_INPUT_ID = "ctl00_PlaceHolderMain_generalSearchForm_txtGSPermitNumber"
SEARCH_BUTTON_XPATH = '//*[@id="ctl00_PlaceHolderMain_btnNewSearch"]'

START_DATE = os.getenv("START_DATE", "04/26/2026")
END_DATE = os.getenv("END_DATE", "04/31/2026")

DATE_FROM_SELECTORS = [
    (By.ID, "ctl00_PlaceHolderMain_generalSearchForm_txtGSFromDate"),
    (By.ID, "ctl00_PlaceHolderMain_generalSearchForm_txtGSPermitDateFrom"),
    (By.ID, "ctl00_PlaceHolderMain_generalSearchForm_txtGSFromDate1"),
    (By.NAME, "ctl00$PlaceHolderMain$generalSearchForm$txtGSFromDate"),
    (By.NAME, "ctl00$PlaceHolderMain$generalSearchForm$txtGSPermitDateFrom"),
]

DATE_TO_SELECTORS = [
    (By.ID, "ctl00_PlaceHolderMain_generalSearchForm_txtGSToDate"),
    (By.ID, "ctl00_PlaceHolderMain_generalSearchForm_txtGSPermitDateTo"),
    (By.ID, "ctl00_PlaceHolderMain_generalSearchForm_txtGSToDate1"),
    (By.NAME, "ctl00$PlaceHolderMain$generalSearchForm$txtGSToDate"),
    (By.NAME, "ctl00$PlaceHolderMain$generalSearchForm$txtGSPermitDateTo"),
]

HEADERS = [
    "S.No",
    "County",
    "State",
    "City",

    "Date",
    "Input Record Number",
    "Record Number",
    "Record Type",
    "Record Status",
    "Project Name",
    "Action",
    "Grid Address",
    "Grid Status",
    "Related Records",
    "Submittal Type",

    "Work Location",
    "Work City",
    "Work State",
    "Work Zip",

    "Applicant Name",
    "Applicant Company",
    "Applicant Address",
    "Applicant City",
    "Applicant State",
    "Applicant Zip",
    "Applicant Phone",
    "Applicant Email",

    "Licensed Professional Name",
    "Licensed Professional Company",
    "Licensed Professional Address",
    "Licensed Professional City",
    "Licensed Professional State",
    "Licensed Professional Zip",
    "Licensed Professional Phone",
    "Licensed Professional Fax",
    "Licensed Professional Contractor Type",
    "Licensed Professional Contractor ID",

    "Related Contact Name",
    "Related Contact Company",
    "Related Contact Address",
    "Related Contact City",
    "Related Contact State",
    "Related Contact Zip",
    "Related Contact Phone",
    "Related Contact Email",

    "Project Description",
    "Application Expiration Date",
    "Job Value",
    "Parcel Information"
]


# =========================================================
# POPUPS
# =========================================================

def show_message_box():
    if not GUI_AVAILABLE:
        print("INFO: (non-interactive) select date range on site, then allow script to continue.")
        return

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo(
        "Info",
        "Select Date Range (or enter search criteria) on the website and click Search.\n\n"
        "After the search results grid loads, click OK to begin automated scraping."
    )
    root.destroy()


def select_input_file():
    if not GUI_AVAILABLE:
        raise RuntimeError("File selection dialog is not available in headless mode. Provide input file via arguments or environment variables.")

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    file_path = filedialog.askopenfilename(
        title="Select input text file containing record numbers",
        filetypes=[
            ("Text Files", "*.txt"),
            ("All Files", "*.*")
        ]
    )

    root.destroy()
    return file_path


def find_date_input(driver, selectors, fallback_keywords):
    for by, value in selectors:
        try:
            return driver.find_element(by, value)
        except Exception:
            pass

    for keyword in fallback_keywords:
        try:
            return driver.find_element(
                By.XPATH,
                "//input[" +
                "contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '" + keyword + "') or " +
                "contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '" + keyword + "') or " +
                "contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '" + keyword + "')]"
            )
        except Exception:
            pass

    return None


def set_input_value(driver, element, value):
    try:
        element.clear()
        element.send_keys(value)
        element.send_keys(Keys.TAB)
        time.sleep(0.3)
        return True
    except Exception:
        try:
            driver.execute_script(
                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
                element,
                value
            )
            return True
        except Exception:
            return False


def set_date_range(driver, start_date, end_date):
    if not start_date or not end_date:
        return False

    from_input = find_date_input(
        driver,
        DATE_FROM_SELECTORS,
        ["fromdate", "from date", "startdate", "start date"]
    )
    to_input = find_date_input(
        driver,
        DATE_TO_SELECTORS,
        ["todate", "to date", "enddate", "end date"]
    )

    if not from_input or not to_input:
        return False

    success_from = set_input_value(driver, from_input, start_date)
    success_to = set_input_value(driver, to_input, end_date)

    if success_from and success_to:
        print(f"Set date range: {start_date} to {end_date}")
        return True

    return False


def apply_date_range_search(driver):
    try:
        if set_date_range(driver, START_DATE, END_DATE):
            if click_search(driver):
                time.sleep(5)
                return True
            print("Date range set but search button click failed.")
    except Exception as e:
        print(f"Date range automation error: {e}")

    return False


# =========================================================
# DRIVER
# =========================================================

def setup_driver():
    options = Options()

    # Headless when running in CI or when HEADLESS env var is set to 1/true
    headless_env = os.getenv("HEADLESS", "true").lower()
    is_headless = headless_env in ("1", "true", "yes", "on") or os.getenv("GITHUB_ACTIONS")

    if is_headless:
        # Use new headless mode when available
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-software-rasterizer")
        options.add_argument("--remote-allow-origins=*")
    else:
        options.add_argument("--start-maximized")

    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-blink-features=AutomationControlled")

    # Create driver using webdriver-manager so CI doesn't need preinstalled chromedriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver


# =========================================================
# EXCEL
# =========================================================

def setup_excel():
    if os.path.exists(OUTPUT_FILE):
        wb = load_workbook(OUTPUT_FILE)
        ws = wb.active

        existing_headers = [cell.value for cell in ws[1]]
        if existing_headers != HEADERS:
            print("Existing Excel headers do not match current HEADERS. Creating a fresh file.")
            wb = Workbook()
            ws = wb.active
            ws.title = "Lee FL Records"
            ws.append(HEADERS)
            wb.save(OUTPUT_FILE)
            return wb, ws, set(), 1

        existing_records = set()
        record_col = HEADERS.index("Record Number") + 1

        for row in range(2, ws.max_row + 1):
            value = ws.cell(row=row, column=record_col).value
            if value:
                existing_records.add(str(value).strip())

        next_sno = ws.max_row
        return wb, ws, existing_records, next_sno

    wb = Workbook()
    ws = wb.active
    ws.title = "Lee FL Records"
    ws.append(HEADERS)
    wb.save(OUTPUT_FILE)

    return wb, ws, set(), 1


def append_record(wb, ws, data):
    ws.append([data.get(header, "") for header in HEADERS])
    wb.save(OUTPUT_FILE)


# =========================================================
# BASIC HELPERS
# =========================================================

def clean_text(text):
    if not text:
        return ""

    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)
    return text.strip()


def safe_text(driver, by, value, timeout=5):
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
        return clean_text(element.text)
    except Exception:
        return ""


def safe_find_text(container, by, value):
    try:
        elem = container.find_element(by, value)
        return clean_text(elem.text)
    except Exception:
        return ""


def clean_phone(text):
    if not text:
        return ""
    digits = re.sub(r"\D", "", text)
    return digits if digits else clean_text(text)


def wait_for_overlay(driver, timeout=20):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script(
                "return document.getElementById('iframeBlocker')?.style.display !== 'block'"
            )
        )
    except Exception:
        pass


def scroll_down(driver):
    driver.execute_script("window.scrollBy({top: 700, behavior: 'smooth'});")
    time.sleep(1)


def scroll_down_smooth(driver, steps=2, amount=700):
    for _ in range(steps):
        driver.execute_script(f"window.scrollBy({{top: {amount}, behavior: 'smooth'}});")
        time.sleep(0.5)


def scroll_up(driver):
    driver.execute_script("window.scrollTo({top: 0, behavior: 'smooth'});")
    time.sleep(1)


def js_click(driver, element):
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center', inline:'center'});",
        element
    )
    time.sleep(0.5)
    driver.execute_script("arguments[0].click();", element)
    time.sleep(2)


def safe_click_xpath(driver, xpath, timeout=10):
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center', inline:'center'});",
            element
        )
        time.sleep(0.5)

        try:
            driver.execute_script("arguments[0].click();", element)
            time.sleep(2)
            return True
        except Exception:
            pass

        try:
            element.click()
            time.sleep(2)
            return True
        except Exception:
            pass

        return False

    except Exception:
        return False


def read_record_numbers(file_path):
    records = []

    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            value = line.strip()
            if value:
                records.append(value)

    return records


# =========================================================
# SEARCH LOGIC (Legacy / Helper)
# =========================================================

def wait_for_search_page(driver):
    WebDriverWait(driver, WAIT_TIME).until(
        EC.presence_of_element_located((By.ID, SEARCH_INPUT_ID))
    )


def clear_and_enter_record_number(driver, record_number):
    wait_for_search_page(driver)

    search_input = WebDriverWait(driver, WAIT_TIME).until(
        EC.element_to_be_clickable((By.ID, SEARCH_INPUT_ID))
    )

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center', inline:'center'});",
        search_input
    )
    time.sleep(0.5)

    search_input.click()
    time.sleep(0.2)
    search_input.send_keys(Keys.CONTROL, "a")
    search_input.send_keys(Keys.DELETE)
    time.sleep(0.2)
    search_input.send_keys(record_number)
    time.sleep(0.5)


def click_search(driver):
    try:
        button = WebDriverWait(driver, WAIT_TIME).until(
            EC.presence_of_element_located((By.XPATH, SEARCH_BUTTON_XPATH))
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center', inline:'center'});",
            button
        )
        time.sleep(1)

        # Normal click
        try:
            button.click()
            print("Search clicked using Selenium click.")
            time.sleep(5)
            return True
        except Exception as e:
            print("Normal click failed:", e)

        # JavaScript click
        try:
            driver.execute_script("arguments[0].click();", button)
            print("Search clicked using JavaScript click.")
            time.sleep(5)
            return True
        except Exception as e:
            print("JavaScript click failed:", e)

        # ENTER key
        try:
            button.send_keys(Keys.ENTER)
            print("Search clicked using ENTER key.")
            time.sleep(5)
            return True
        except Exception as e:
            print("ENTER key failed:", e)

        # Direct Accela postback
        try:
            driver.execute_script("""
                WebForm_DoPostBackWithOptions(
                    new WebForm_PostBackOptions(
                        "ctl00$PlaceHolderMain$btnNewSearch",
                        "",
                        true,
                        "",
                        "",
                        false,
                        true
                    )
                );
            """)
            print("Search clicked using direct Accela postback.")
            time.sleep(5)
            return True
        except Exception as e:
            print("Direct postback failed:", e)

    except Exception as e:
        print("Search button not found:", e)

    return False


def wait_for_detail_page(driver, timeout=20):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(
                (By.ID, "ctl00_PlaceHolderMain_lblPermitNumber")
            )
        )
        return True
    except TimeoutException:
        return False


def is_no_result_or_same_search_page(driver):
    try:
        body_text = clean_text(driver.find_element(By.TAG_NAME, "body").text).lower()

        no_result_keywords = [
            "your search returned no results",
            "no records found",
            "no result",
            "no records"
        ]

        for keyword in no_result_keywords:
            if keyword in body_text:
                return True

        try:
            driver.find_element(By.ID, "ctl00_PlaceHolderMain_lblPermitNumber")
            return False
        except Exception:
            return True

    except Exception:
        return True


def go_back_to_search(driver):
    scroll_up(driver)

    xpaths = [
        '//a[contains(normalize-space(.), "Search Applications")]',
        '//a[contains(normalize-space(.), "Search Records")]',
        '//a[contains(@href, "CapHome.aspx")]'
    ]

    for xpath in xpaths:
        try:
            element = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            js_click(driver, element)
            time.sleep(5)
            wait_for_search_page(driver)
            return True
        except Exception:
            pass

    driver.get(URL)
    time.sleep(6)
    wait_for_search_page(driver)
    return True


# =========================================================
# PARSING HELPERS
# =========================================================

def extract_email(text):
    match = re.search(r"[\w\.-]+@[\w\.-]+\.\w+", text or "", re.I)
    return match.group(0).strip() if match else ""


def extract_phone_after_label(text, label):
    if not text:
        return ""

    pattern = rf"{re.escape(label)}\s*:?\s*([0-9\-\(\)\s\.]+)"
    match = re.search(pattern, text, re.I)

    if match:
        return re.sub(r"\D", "", match.group(1))

    return ""


def extract_all_phones(text):
    return re.findall(r"\b\d{10}\b", text or "")


def looks_like_company(line):
    if not line:
        return False

    if looks_like_address(line):
        return False

    company_keywords = [
        r"\bLLC\b", r"\bINC\b", r"\bCORP\b", r"\bCORPORATION\b", r"\bCO\b", r"\bCOMPANY\b",
        r"\bLTD\b", r"\bLP\b", r"\bLLP\b", r"\bMECHANICAL\b", r"\bCONSTRUCTION\b",
        r"\bPLUMBING\b", r"\bELECTRIC\b", r"\bELECTRICAL\b", r"\bROOFING\b",
        r"\bBUILDERS\b", r"\bBUILDING\b", r"\bCONTRACTOR\b", r"\bCONTRACTORS\b",
        r"\bSERVICES\b", r"\bSERVICE\b", r"\bHOMES\b", r"\bENTERPRISES\b", r"\bGROUP\b",
        r"\bDEVELOPMENT\b", r"\bDEVELOPERS\b", r"\bREALTY\b", r"\bPROPERTIES\b", r"\bPROPERTY\b",
        r"\bASSOCIATES\b", r"\bPARTNERS\b", r"\bINSTALLATIONS\b", r"\bDESIGNS\b", r"\bSYSTEMS\b"
    ]

    return any(re.search(keyword, line, re.I) for keyword in company_keywords)


def looks_like_address(line):
    if not line:
        return False

    pattern = r"^(?:\d+[A-Z]?\s+[A-Z0-9#\.\-\s]+|P\.?O\.?\s+BOX\s+\d+)"
    if not re.search(pattern, line, re.I):
        return False

    address_suffixes = [
        r"\bST\b", r"\bSTREET\b", r"\bAVE\b", r"\bAVENUE\b", r"\bRD\b", r"\bROAD\b",
        r"\bDR\b", r"\bDRIVE\b", r"\bCIR\b", r"\bCIRCLE\b", r"\bCT\b", r"\bCOURT\b",
        r"\bLN\b", r"\bLANE\b", r"\bBLVD\b", r"\bWAY\b", r"\bTRL\b", r"\bTRAIL\b",
        r"\bPKWY\b", r"\bHWY\b", r"\bHIGHWAY\b", r"\bPL\b", r"\bPLACE\b",
        r"\bTER\b", r"\bTERRACE\b", r"\bLOOP\b", r"\bPIKE\b", r"\bRUN\b",
        r"\bSTE\b", r"\bSUITE\b", r"\bAPT\b", r"\bUNIT\b", r"\bRM\b", r"\bROOM\b",
        r"\bBOX\b", r"\b#\s*\d+"
    ]

    return any(re.search(suffix, line, re.I) for suffix in address_suffixes)


def parse_city_state_zip(line):
    result = {
        "city": "",
        "state": "",
        "zip": ""
    }

    if not line:
        return result

    text = clean_text(line)

    text = text.replace(" ,", ",")
    text = re.sub(r"\s*,\s*", ", ", text)

    STATE_MAP = {
        "ALABAMA": "AL", "ALASKA": "AK", "ARIZONA": "AZ", "ARKANSAS": "AR",
        "CALIFORNIA": "CA", "COLORADO": "CO", "CONNECTICUT": "CT", "DELAWARE": "DE",
        "FLORIDA": "FL", "GEORGIA": "GA", "HAWAII": "HI", "IDAHO": "ID",
        "ILLINOIS": "IL", "INDIANA": "IN", "IOWA": "IA", "KANSAS": "KS",
        "KENTUCKY": "KY", "LOUISIANA": "LA", "MAINE": "ME", "MARYLAND": "MD",
        "MASSACHUSETTS": "MA", "MICHIGAN": "MI", "MINNESOTA": "MN", "MISSISSIPPI": "MS",
        "MISSOURI": "MO", "MONTANA": "MT", "NEBRASKA": "NE", "NEVADA": "NV",
        "NEW HAMPSHIRE": "NH", "NEW JERSEY": "NJ", "NEW MEXICO": "NM", "NEW YORK": "NY",
        "NORTH CAROLINA": "NC", "NORTH DAKOTA": "ND", "OHIO": "OH", "OKLAHOMA": "OK",
        "OREGON": "OR", "PENNSYLVANIA": "PA", "RHODE ISLAND": "RI", "SOUTH CAROLINA": "SC",
        "SOUTH DAKOTA": "SD", "TENNESSEE": "TN", "TEXAS": "TX", "UTAH": "UT",
        "VERMONT": "VT", "VIRGINIA": "VA", "WASHINGTON": "WA", "WEST VIRGINIA": "WV",
        "WISCONSIN": "WI", "WYOMING": "WY"
    }

    # FULL STATE NAME
    match = re.search(
        r"(.+?),\s*([A-Za-z ]+),\s*(\d{5}(?:-\d{4})?)",
        text
    )

    if match:
        city = match.group(1).strip(" ,")
        state = match.group(2).strip().upper()
        zip_code = match.group(3).strip()

        if state in STATE_MAP:
            state = STATE_MAP[state]

        result["city"] = city
        result["state"] = state
        result["zip"] = zip_code

        return result

    # SHORT STATE
    match = re.search(
        r"(.+?),\s*([A-Z]{2}),?\s*(\d{5}(?:-\d{4})?)",
        text
    )

    if match:
        result["city"] = match.group(1).strip(" ,")
        result["state"] = match.group(2).strip()
        result["zip"] = match.group(3).strip()

        return result

    # CITY + ZIP ONLY
    match = re.search(
        r"(.+?),\s*(\d{5}(?:-\d{4})?)",
        text
    )

    if match:
        result["city"] = match.group(1).strip(" ,")
        result["state"] = "FL"
        result["zip"] = match.group(2).strip()

        return result

    return result





# =========================================================
# DETAIL SCRAPING
# =========================================================

def scrape_record_summary(driver):
    result = {
        "record_number": "",
        "record_type": "",
        "record_status": ""
    }

    try:
        summary = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.record-summary"))
        )

        # Record Number
        try:
            result["record_number"] = clean_text(
                summary.find_element(By.ID, "ctl00_PlaceHolderMain_lblPermitNumber").text
            )
        except Exception:
            pass

        # Record Type
        try:
            result["record_type"] = clean_text(
                summary.find_element(By.ID, "ctl00_PlaceHolderMain_lblPermitType").text
            )
        except Exception:
            pass

        # Record Status
        try:
            result["record_status"] = clean_text(
                summary.find_element(By.ID, "ctl00_PlaceHolderMain_lblRecordStatus").text
            )
        except Exception:
            pass

    except Exception as e:
        print("Record summary scrape error:", e)

    return result


def scrape_work_location(driver):
    result = {
        "location": "",
        "city": "",
        "state": "",
        "zip": ""
    }

    try:
        table = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "tbl_worklocation"))
        )

        value_td = table.find_element(
            By.XPATH,
            ".//td[contains(@class,'NotBreakWord')]"
        )

        raw_text = clean_text(value_td.text)
        print("WORK LOCATION RAW:", raw_text)

        lines = [clean_text(x) for x in raw_text.splitlines() if clean_text(x)]

        if len(lines) >= 1:
            result["location"] = lines[0]

        # Search city/state/zip line
        for line in lines:
            match = re.search(r"(.+?)\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)", line)
            if match:
                result["city"] = match.group(1).strip()
                result["state"] = match.group(2).strip()
                result["zip"] = match.group(3).strip()
                break

    except Exception as e:
        print("WORK LOCATION ERROR:", e)

    return result


def get_section_text_by_h1_label(driver, label_keyword):
    try:
        h1_elements = driver.find_elements(By.XPATH, "//h1")

        for h1 in h1_elements:
            label = clean_text(h1.text).lower()

            if label_keyword.lower() in label:
                try:
                    row = h1.find_element(By.XPATH, "./ancestor::tr[1]")
                    row_text = clean_text(row.text)
                    row_text = re.sub(
                        re.escape(h1.text),
                        "",
                        row_text,
                        flags=re.I
                    ).strip()

                    if row_text:
                        return row_text
                except Exception:
                    pass

                try:
                    parent_td = h1.find_element(By.XPATH, "./ancestor::td[1]")
                    td_text = clean_text(parent_td.text)
                    td_text = td_text.replace(h1.text, "").strip()

                    if td_text:
                        return td_text
                except Exception:
                    pass

    except Exception:
        pass

    return ""


def clean_phone(text):
    if not text:
        return ""
    only_digits = re.sub(r"\D", "", text)
    return only_digits if len(only_digits) >= 7 else ""


def extract_all_phone_numbers(container):
    numbers = []

    try:
        phone_divs = container.find_elements(
            By.CSS_SELECTOR,
            ".ACA_PhoneNumberLTR"
        )

        for div in phone_divs:
            txt = clean_text(div.text)
            digits = re.sub(r"\D", "", txt)

            if digits and digits not in numbers:
                numbers.append(digits)

    except Exception:
        pass

    return ", ".join(numbers)


def scrape_applicant(driver):
    result = {
        "name": "",
        "company": "",
        "address": "",
        "city": "",
        "state": "",
        "zip": "",
        "phone": "",
        "email": ""
    }

    try:
        container = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.XPATH, "//span[contains(.,'Applicant:')]/ancestor::td[1]")
            )
        )

        # Name
        first = safe_find_text(container, By.CSS_SELECTOR, ".contactinfo_firstname")
        last = safe_find_text(container, By.CSS_SELECTOR, ".contactinfo_lastname")
        result["name"] = f"{first} {last}".strip()

        # Company
        result["company"] = safe_find_text(container, By.CSS_SELECTOR, ".contactinfo_businessname")

        # Address
        addr1 = safe_find_text(container, By.CSS_SELECTOR, ".contactinfo_addressline1")
        addr2 = safe_find_text(container, By.CSS_SELECTOR, ".contactinfo_addressline2")
        result["address"] = " ".join([x for x in [addr1, addr2] if x])

        # Region
        regions = container.find_elements(By.CSS_SELECTOR, ".contactinfo_region")
        region_text = " ".join([clean_text(x.text) for x in regions if clean_text(x.text)])
        csz = parse_city_state_zip(region_text)
        result["city"] = csz["city"]
        result["state"] = csz["state"]
        result["zip"] = csz["zip"]

        # Phone
        result["phone"] = extract_all_phone_numbers(container)

        # Email
        try:
            email_elem = container.find_element(By.CSS_SELECTOR, ".contactinfo_email")
            result["email"] = extract_email(email_elem.text)
        except Exception:
            pass

    except Exception as e:
        print("Applicant scrape error:", e)

    return result


def scrape_licensed_professional(driver):
    result = {
        "name": "",
        "company": "",
        "address": "",
        "city": "",
        "state": "",
        "zip": "",
        "phone": "",
        "fax": "",
        "contractor_type": "",
        "contractor_id": ""
    }

    try:
        container = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "tbl_licensedps"))
        )

        raw = clean_text(container.text)
        lines = [clean_text(x) for x in raw.splitlines() if clean_text(x)]

        if len(lines) >= 1:
            result["name"] = lines[0]

        if len(lines) >= 2:
            result["company"] = lines[1]

        if len(lines) >= 3:
            result["address"] = lines[2]

        # City State Zip
        regions = container.find_elements(By.CSS_SELECTOR, ".contactinfo_region")
        region_text = " ".join([clean_text(x.text) for x in regions if clean_text(x.text)])
        csz = parse_city_state_zip(region_text)
        if csz["zip"] or csz["city"]:
            result["city"] = csz["city"]
            result["state"] = csz["state"]
            result["zip"] = csz["zip"]
        else:
            for line in lines:
                csz = parse_city_state_zip(line)
                if csz["zip"]:
                    result["city"] = csz["city"]
                    result["state"] = csz["state"]
                    result["zip"] = csz["zip"]
                    break

        # Phone and Fax
        all_numbers = extract_all_phone_numbers(container)

        if all_numbers:
            split_numbers = [
                x.strip()
                for x in all_numbers.split(",")
                if x.strip()
            ]

            if len(split_numbers) >= 1:
                result["phone"] = split_numbers[0]

            if len(split_numbers) >= 2:
                result["fax"] = split_numbers[1]

        # Contractor
        for line in lines:
            lic_match = re.search(r'(.+?)\s+([A-Z]{2,5}\d{5,})', line)
            if lic_match:
                result["contractor_type"] = lic_match.group(1).strip()
                result["contractor_id"] = lic_match.group(2).strip()
                break

    except Exception as e:
        print("Licensed Professional scrape error:", e)

    return result


def scrape_project_description(driver):
    try:
        row = driver.find_element(
            By.XPATH,
            "//span[contains(.,'Project Description:')]/ancestor::td[1]"
        )

        text = clean_text(row.text)
        text = re.sub(r"Project Description\s*:?", "", text, flags=re.I).strip()
        return text

    except Exception:
        return ""


def expand_more_details(driver):
    try:
        elems = driver.find_elements(By.XPATH, '//*[@id="lnkASI" or @id="imgParcel" or @id="lnkRc"]')
        if elems and elems[0].is_displayed():
            return True
    except Exception:
        pass

    scroll_down_smooth(driver, steps=1, amount=500)
    wait_for_overlay(driver)
    clicked = safe_click_xpath(driver, '//*[@id="lnkMoreDetail"]', timeout=8)
    if clicked:
        wait_for_overlay(driver)
        time.sleep(2)

    return clicked


def open_application_information(driver):
    expand_more_details(driver)
    wait_for_overlay(driver)

    clicked = safe_click_xpath(driver, '//*[@id="lnkASI"]', timeout=8)
    if clicked:
        wait_for_overlay(driver)
        time.sleep(2)
        scroll_down_smooth(driver, steps=1, amount=500)
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(@id, 'PermitDetailList1')]"))
            )
        except Exception:
            pass

    return clicked


def get_job_value(driver):
    for container_id in ["trASITList", "trASIList", "PermitDetailList1"]:
        try:
            container = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, f"//*[contains(@id, '{container_id}')]"))
            )
            cols = container.find_elements(By.CSS_SELECTOR, "div.MoreDetail_ItemCol")
            for i in range(len(cols)):
                try:
                    label = clean_text(cols[i].text).lower()
                    if label in ["job value:", "job valuation:", "valuation:"]:
                        if i + 1 < len(cols):
                            val = clean_text(cols[i + 1].text)
                            if val:
                                return val
                except Exception:
                    pass
        except Exception:
            pass

    return ""


def open_parcel_information(driver):
    expand_more_details(driver)
    wait_for_overlay(driver)

    clicked = safe_click_xpath(driver, '//*[@id="imgParcel"]', timeout=8)
    if clicked:
        wait_for_overlay(driver)
        time.sleep(2)
        scroll_down_smooth(driver, steps=1, amount=500)
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(@id, 'Parcel') or contains(@id, 'parcel')]"))
            )
        except Exception:
            pass

    return clicked


def get_parcel_information(driver):
    open_parcel_information(driver)

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(@id, 'Parcel') or contains(@id, 'parcel')]"))
        )
    except Exception:
        pass

    possible_selectors = [
        "#ctl00_PlaceHolderMain_PermitDetailList1_tbParcel",
        "#tblparcel",
        "#tbl_parcel",
        "[id*='Parcel'] table",
        "table[id*='parcel']",
        "table[id*='Parcel']"
    ]

    raw_parcel_text = ""
    for selector in possible_selectors:
        try:
            elems = driver.find_elements(By.CSS_SELECTOR, selector)
            for elem in elems:
                txt = clean_text(elem.text)
                if txt and len(txt) > 3:
                    raw_parcel_text = txt
                    break
            if raw_parcel_text:
                break
        except Exception:
            pass

    if not raw_parcel_text or "no records found" in raw_parcel_text.lower():
        return ""

    clean_raw = re.sub(r"\s+", " ", raw_parcel_text).strip().lower()
    if clean_raw in ["parcel information", "parcel number", "parcel number parcel information", "parcel", "strap"]:
        return ""

    parcel_num = ""
    legal_desc = ""

    parcel_number_match = re.search(
        r'\b\d{2}-\d{2}-\d{2}-\d{2}-\d{5}\.\d{4}\b',
        raw_parcel_text
    )
    if parcel_number_match:
        parcel_num = parcel_number_match.group(0)
    else:
        for line in raw_parcel_text.splitlines():
            line = clean_text(line)
            if re.search(r"\bPARCEL\b|\bSTRAP\b", line, re.I):
                parcel_num = re.sub(r"^(?:Parcel Number|Parcel|STRAP)\s*:?", "", line, flags=re.I).strip()
                break

    if parcel_num and not re.search(r"\d", parcel_num):
        parcel_num = ""

    legal_match = re.search(
        r'Legal Description\s*:?\s*(.+?)(?:\n|$)',
        raw_parcel_text,
        re.I
    )
    if legal_match and legal_match.group(1).strip():
        legal_desc = legal_match.group(1).strip()
    else:
        for line in raw_parcel_text.splitlines():
            line = clean_text(line)
            if "LEGAL" in line.upper() or "DESCRIPTION" in line.upper():
                cleaned = re.sub(r"^(?:Legal Description|Legal|Description)\s*:?", "", line, flags=re.I).strip()
                if cleaned and cleaned != parcel_num:
                    legal_desc = cleaned
                    break

    if legal_desc.lower() in ["parcel information", "information", "description", "legal description", "none", "n/a"]:
        legal_desc = ""

    parts = []
    if parcel_num:
        parts.append(f"Parcel Number: {parcel_num}")
    if legal_desc:
        parts.append(f"Legal Description: {legal_desc}")

    if parts:
        return " | ".join(parts)

    return ""


def scrape_related_contact(driver):
    result = {
        "name": "",
        "company": "",
        "address": "",
        "city": "",
        "state": "",
        "zip": "",
        "phone": "",
        "email": ""
    }

    try:
        # Expand Related Contacts
        wait_for_overlay(driver)
        clicked = safe_click_xpath(
            driver,
            '//*[@id="lnkRc"]',
            timeout=10
        )

        if not clicked:
            return result

        wait_for_overlay(driver)
        time.sleep(2)
        scroll_down_smooth(driver, steps=1, amount=500)

        # Wait section
        container = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "trRCList"))
        )

        # First contact block
        first_contact = container.find_element(
            By.CSS_SELECTOR,
            "div.MoreDetail_ItemCol"
        )

        # Name
        first = safe_find_text(first_contact, By.CSS_SELECTOR, ".contactinfo_firstname")
        last = safe_find_text(first_contact, By.CSS_SELECTOR, ".contactinfo_lastname")
        result["name"] = f"{first} {last}".strip()

        # Company
        result["company"] = safe_find_text(first_contact, By.CSS_SELECTOR, ".contactinfo_businessname")

        # Address
        addr1 = safe_find_text(first_contact, By.CSS_SELECTOR, ".contactinfo_addressline1")
        addr2 = safe_find_text(first_contact, By.CSS_SELECTOR, ".contactinfo_addressline2")
        result["address"] = " ".join([x for x in [addr1, addr2] if x])

        # Region
        regions = first_contact.find_elements(By.CSS_SELECTOR, ".contactinfo_region")
        region = " ".join([clean_text(x.text) for x in regions if clean_text(x.text)])
        csz = parse_city_state_zip(region)

        result["city"] = csz["city"]
        result["state"] = csz["state"]
        result["zip"] = csz["zip"]

        # Phone
        result["phone"] = extract_all_phone_numbers(first_contact)

        # Email
        try:
            email_elem = first_contact.find_element(By.CSS_SELECTOR, ".contactinfo_email")
            result["email"] = extract_email(email_elem.text)
        except Exception:
            pass

    except Exception as e:
        print("Related Contact scrape error:", e)

    return result


def scrape_application_expiration_date(driver):
    result = ""

    try:
        # Expand Application Information Table
        wait_for_overlay(driver)
        clicked = safe_click_xpath(
            driver,
            '//*[@id="lnkASITableList"]',
            timeout=10
        )

        if not clicked:
            return result

        wait_for_overlay(driver)
        time.sleep(2)
        scroll_down_smooth(driver, steps=1, amount=500)

        # Wait section
        container = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "trASITList"))
        )

        # Find all label/value cols
        cols = container.find_elements(
            By.CSS_SELECTOR,
            "div.MoreDetail_ItemCol"
        )

        for i in range(len(cols)):
            try:
                label = clean_text(cols[i].text)
                if label.lower() == "expiration date:":
                    if i + 1 < len(cols):
                        result = clean_text(cols[i + 1].text)
                        break
            except Exception:
                pass

    except Exception as e:
        print("Application Expiration Date scrape error:", e)

    return result


def scrape_current_detail_page(driver, input_record, sno, grid_data):
    data = {header: "" for header in HEADERS}

    data["S.No"] = sno
    data["County"] = COUNTY_VALUE
    data["State"] = STATE_VALUE
    data["City"] = CITY_VALUE
    data["Input Record Number"] = input_record

    # Populate grid data
    data["Date"] = grid_data.get("Date", "")
    data["Record Type"] = grid_data.get("Record Type", "")
    data["Project Name"] = grid_data.get("Project Name", "")
    data["Action"] = grid_data.get("Action", "")
    data["Grid Address"] = grid_data.get("Address", "")
    data["Grid Status"] = grid_data.get("Status", "")
    data["Related Records"] = grid_data.get("Related Records", "")
    data["Submittal Type"] = grid_data.get("Submittal Type", "")

    try:
        summary = scrape_record_summary(driver)
        data["Record Number"] = summary["record_number"] or input_record
        data["Record Type"] = summary["record_type"] or grid_data.get("Record Type", "")
        data["Record Status"] = summary["record_status"] or grid_data.get("Status", "")

        # Work Location
        work = scrape_work_location(driver)

        data["Work Location"] = work["location"]
        data["Work City"] = work["city"]
        data["Work State"] = work["state"]
        data["Work Zip"] = work["zip"]

        scroll_down_smooth(driver, steps=2, amount=700)

        # Applicant
        applicant = scrape_applicant(driver)

        data["Applicant Name"] = applicant["name"]
        data["Applicant Company"] = applicant["company"]
        data["Applicant Address"] = applicant["address"]
        data["Applicant City"] = applicant["city"]
        data["Applicant State"] = applicant["state"]
        data["Applicant Zip"] = applicant["zip"]
        data["Applicant Phone"] = applicant["phone"]
        data["Applicant Email"] = applicant["email"]

        # Licensed Professional
        licensed = scrape_licensed_professional(driver)

        data["Licensed Professional Name"] = licensed["name"]
        data["Licensed Professional Company"] = licensed["company"]
        data["Licensed Professional Address"] = licensed["address"]
        data["Licensed Professional City"] = licensed["city"]
        data["Licensed Professional State"] = licensed["state"]
        data["Licensed Professional Zip"] = licensed["zip"]
        data["Licensed Professional Phone"] = licensed["phone"]
        data["Licensed Professional Fax"] = licensed["fax"]
        data["Licensed Professional Contractor Type"] = licensed["contractor_type"]
        data["Licensed Professional Contractor ID"] = licensed["contractor_id"]

        # Project Description
        data["Project Description"] = scrape_project_description(driver)

        # Application Expiration Date
        data["Application Expiration Date"] = scrape_application_expiration_date(driver)

        # Job Value
        open_application_information(driver)
        data["Job Value"] = get_job_value(driver)

        # Parcel Information
        data["Parcel Information"] = get_parcel_information(driver)

        # Related Contacts
        related = scrape_related_contact(driver)

        data["Related Contact Name"] = related["name"]
        data["Related Contact Company"] = related["company"]
        data["Related Contact Address"] = related["address"]
        data["Related Contact City"] = related["city"]
        data["Related Contact State"] = related["state"]
        data["Related Contact Zip"] = related["zip"]
        data["Related Contact Phone"] = related["phone"]
        data["Related Contact Email"] = related["email"]

    except Exception as e:
        print(f"Error during detail page scraping for {input_record}: {e}")
        data["Record Number"] = ""  # Mark as failed so it's not added to existing_records
        if not data["Record Status"]:
            data["Record Status"] = grid_data.get("Status", "")

    return data


# =========================================================
# MAIN LOOP (Date Range Grid Scraper Architecture)
# =========================================================

def main():
    driver = None

    try:
        wb, ws, existing_records, next_sno = setup_excel()

        driver = setup_driver()
        driver.get(URL)

        wait_for_overlay(driver)

        WebDriverWait(driver, WAIT_TIME).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )

        time.sleep(6)

        # Try to set the requested date range and search before scraping
        applied_date_range = apply_date_range_search(driver)

        # Only show interactive prompt when not running in CI/headless mode and automated range wasn't applied
        headless_env = os.getenv("HEADLESS", "true").lower()
        is_headless = headless_env in ("1", "true", "yes", "on") or os.getenv("GITHUB_ACTIONS")
        if not is_headless and not applied_date_range:
            show_message_box()

        results_page_url = driver.current_url

        sno = next_sno
        page_number = 1
        seen_page_signatures = set()

        while True:
            print(f"\nPROCESSING PAGE: {page_number}")
            time.sleep(5)

            try:
                # WAIT GRID FULLY LOADED
                WebDriverWait(driver, 20).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "//a[contains(@href,'CapDetail.aspx')]")
                    )
                )
                time.sleep(3)

                rows = driver.find_elements(By.XPATH, "//a[contains(@href,'CapDetail.aspx')]")
                total_records = len(rows)
                print(f"FOUND RECORDS ON PAGE {page_number}: {total_records}")
            except Exception as e:
                print(f"Error finding records on page {page_number}: {e}")
                break

            if total_records == 0:
                print("No records found on current page. Exiting loop.")
                break

            current_url = driver.current_url
            current_records = tuple([r.text.strip() for r in rows[:5]])
            page_signature = (current_url, current_records)
            if page_signature in seen_page_signatures:
                print("DUPLICATE PAGE DETECTED - STOPPING")
                break
            seen_page_signatures.add(page_signature)

            col_map = {
                "Date": -1,
                "Record Number": 1,
                "Address": 2,
                "Project Name": 3,
                "Record Type": -1,
                "Status": 4,
                "Action": 5,
                "Related Records": 6,
                "Submittal Type": 7
            }
            try:
                parent_table = rows[0].find_element(By.XPATH, "./ancestor::table[1]")
                th_elems = parent_table.find_elements(By.XPATH, ".//th")
                for idx, th in enumerate(th_elems):
                    txt = clean_text(th.text).lower()
                    if not txt:
                        txt = clean_text(th.get_attribute("innerText")).lower()
                    
                    if any(k in txt for k in ["record number", "permit number", "application number", "record id", "permit id", "cap number"]):
                        col_map["Record Number"] = idx
                    elif any(k in txt for k in ["address", "location", "work location"]):
                        col_map["Address"] = idx
                    elif any(k in txt for k in ["project name", "project", "job name"]):
                        col_map["Project Name"] = idx
                    elif any(k in txt for k in ["submittal type", "submittal"]):
                        col_map["Submittal Type"] = idx
                    elif any(k in txt for k in ["record type", "permit type", "application type", "type"]):
                        col_map["Record Type"] = idx
                    elif any(k in txt for k in ["status", "record status"]):
                        col_map["Status"] = idx
                    elif any(k in txt for k in ["action", "actions"]):
                        col_map["Action"] = idx
                    elif any(k in txt for k in ["related records", "related"]):
                        col_map["Related Records"] = idx
                    elif "date" in txt:
                        col_map["Date"] = idx
                print("DYNAMIC COLUMN MAP:", col_map)
            except Exception as e:
                print("Error mapping table headers dynamically, using defaults:", e)

            for i in range(total_records):
                try:
                    # REFRESH ROWS EACH TIME
                    rows = driver.find_elements(By.XPATH, "//a[contains(@href,'CapDetail.aspx')]")
                    row = rows[i]

                    # SCROLL ROW INTO VIEW
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
                    time.sleep(1)

                    record_number = row.text.strip()
                    if not record_number:
                        continue

                    if record_number in existing_records:
                        print(f"Skipping existing record: {record_number}")
                        continue

                    print(f"\nSCRAPING GRID ROW: {record_number}")

                    # SCRAPE GRID DATA
                    parent_row = row.find_element(By.XPATH, "./ancestor::tr[1]")
                    cells = parent_row.find_elements(By.XPATH, "./td")

                    grid_data = {
                        "Date": cells[col_map["Date"]].text.strip() if col_map["Date"] >= 0 and len(cells) > col_map["Date"] else "",
                        "Record Number": cells[col_map["Record Number"]].text.strip() if col_map["Record Number"] >= 0 and len(cells) > col_map["Record Number"] else "",
                        "Address": cells[col_map["Address"]].text.strip() if col_map["Address"] >= 0 and len(cells) > col_map["Address"] else "",
                        "Project Name": cells[col_map["Project Name"]].text.strip() if col_map["Project Name"] >= 0 and len(cells) > col_map["Project Name"] else "",
                        "Record Type": cells[col_map["Record Type"]].text.strip() if col_map["Record Type"] >= 0 and len(cells) > col_map["Record Type"] else "",
                        "Status": cells[col_map["Status"]].text.strip() if col_map["Status"] >= 0 and len(cells) > col_map["Status"] else "",
                        "Action": cells[col_map["Action"]].text.strip() if col_map["Action"] >= 0 and len(cells) > col_map["Action"] else "",
                        "Related Records": cells[col_map["Related Records"]].text.strip() if col_map["Related Records"] >= 0 and len(cells) > col_map["Related Records"] else "",
                        "Submittal Type": cells[col_map["Submittal Type"]].text.strip() if col_map["Submittal Type"] >= 0 and len(cells) > col_map["Submittal Type"] else "",
                    }

                    print("GRID DATA:", grid_data)

                    # NOW OPEN RECORD (New tab navigation to preserve AJAX grid pagination state)
                    driver.execute_script("window.focus();")
                    time.sleep(0.5)

                    try:
                        detail_url = row.get_attribute("href")
                        driver.execute_script(f"window.open('{detail_url}', '_blank');")
                        time.sleep(1)

                        driver.switch_to.window(driver.window_handles[1])
                        wait_for_overlay(driver)

                        # Wait for detail page to load
                        try:
                            WebDriverWait(driver, 15).until(
                                EC.presence_of_element_located((By.ID, "ctl00_PlaceHolderMain_lblPermitNumber"))
                            )
                        except Exception:
                            print("Detail page element not found immediately, refreshing to fix potential black screen...")
                            driver.refresh()
                            wait_for_overlay(driver)
                            WebDriverWait(driver, 20).until(
                                EC.presence_of_element_located((By.ID, "ctl00_PlaceHolderMain_lblPermitNumber"))
                            )
                        wait_for_overlay(driver)
                        time.sleep(3)
                        detail_data = scrape_current_detail_page(driver, record_number, sno, grid_data)

                        # Close the detail tab and switch back to main grid tab
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        wait_for_overlay(driver)
                        time.sleep(1)

                    except Exception as e:
                        print(f"Timeout/Error clicking or loading detail page for {record_number}: {e}")
                        detail_data = {header: "" for header in HEADERS}
                        detail_data["S.No"] = sno
                        detail_data["County"] = COUNTY_VALUE
                        detail_data["State"] = STATE_VALUE
                        detail_data["City"] = CITY_VALUE
                        detail_data["Input Record Number"] = record_number
                        detail_data["Record Number"] = ""  # Leave empty so it's not marked as permanently scraped
                        detail_data["Date"] = grid_data.get("Date", "")
                        detail_data["Record Type"] = grid_data.get("Record Type", "")
                        detail_data["Project Name"] = grid_data.get("Project Name", "")
                        detail_data["Action"] = grid_data.get("Action", "")
                        detail_data["Grid Address"] = grid_data.get("Address", "")
                        detail_data["Grid Status"] = grid_data.get("Status", "")
                        detail_data["Record Status"] = grid_data.get("Status", "")
                        detail_data["Related Records"] = grid_data.get("Related Records", "")
                        detail_data["Submittal Type"] = grid_data.get("Submittal Type", "")

                        try:
                            if len(driver.window_handles) > 1:
                                driver.close()
                                driver.switch_to.window(driver.window_handles[0])
                                wait_for_overlay(driver)
                                time.sleep(1)
                            elif "CapDetail.aspx" in driver.current_url:
                                driver.back()
                                wait_for_overlay(driver)
                                time.sleep(3)
                        except Exception:
                            pass

                    append_record(wb, ws, detail_data)
                    if detail_data.get("Record Number"):
                        existing_records.add(record_number)
                    print(f"Saved Record: {record_number}")
                    sno += 1

                    # WAIT GRID FULLY LOADED AGAIN
                    WebDriverWait(driver, 20).until(
                        EC.presence_of_all_elements_located(
                            (By.XPATH, "//a[contains(@href,'CapDetail.aspx')]")
                        )
                    )

                except Exception as e:
                    print(f"ROW ERROR: {e}")
                    try:
                        if len(driver.window_handles) > 1:
                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])
                            wait_for_overlay(driver)
                            time.sleep(2)
                        else:
                            if "CapDetail.aspx" in driver.current_url:
                                driver.back()
                                wait_for_overlay(driver)
                                time.sleep(3)
                    except Exception:
                        pass

            # =====================================================
            # PAGINATION
            # =====================================================

            clicked_pagination = False

            try:
                wait_for_overlay(driver)

                # Store old page records
                old_rows = driver.find_elements(
                    By.XPATH,
                    "//a[contains(@href,'CapDetail.aspx')]"
                )

                old_records = [
                    r.text.strip()
                    for r in old_rows[:5]
                ]

                # Find NEXT button ONLY
                next_button = None

                next_candidates = driver.find_elements(
                    By.XPATH,
                    "//a[contains(.,'Next')]"
                )

                for btn in next_candidates:
                    try:
                        txt = clean_text(btn.text).lower()
                        href = (btn.get_attribute("href") or "").lower()
                        cls = (btn.get_attribute("class") or "").lower()

                        if not btn.is_displayed():
                            continue

                        if "disabled" in cls:
                            continue

                        if "javascript:void(0)" in href:
                            continue

                        if "next" in txt:
                            next_button = btn
                            break

                    except Exception:
                        pass

                if not next_button:
                    print("NO MORE PAGES")
                    break

                # Scroll to button
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});",
                    next_button
                )

                time.sleep(1)

                # Click next
                driver.execute_script(
                    "arguments[0].click();",
                    next_button
                )

                print(f"MOVING TO NEXT PAGE")

                wait_for_overlay(driver)

                # Wait old rows stale
                if old_rows:
                    WebDriverWait(driver, 20).until(
                        EC.staleness_of(old_rows[0])
                    )

                # Wait new rows
                WebDriverWait(driver, 20).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "//a[contains(@href,'CapDetail.aspx')]")
                    )
                )

                time.sleep(3)

                # Verify page changed
                new_rows = driver.find_elements(
                    By.XPATH,
                    "//a[contains(@href,'CapDetail.aspx')]"
                )

                new_records = [
                    r.text.strip()
                    for r in new_rows[:5]
                ]

                if old_records == new_records:
                    print("SAME PAGE DETECTED - STOPPING")
                    break

                clicked_pagination = True

            except Exception as e:
                print("PAGINATION ERROR:", e)
                break

            if clicked_pagination:
                page_number += 1
                results_page_url = driver.current_url

        print("\nScraping completed.")
        print(f"Output saved to: {OUTPUT_FILE}")

        if GUI_AVAILABLE and not (os.getenv("HEADLESS", "true").lower() in ("1", "true", "yes", "on") or os.getenv("GITHUB_ACTIONS")):
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            messagebox.showinfo(
                "Done",
                f"Scraping completed.\nOutput saved to:\n{OUTPUT_FILE}"
            )
            root.destroy()
        else:
            print(f"Scraping completed. Output saved to: {OUTPUT_FILE}")

    except Exception as e:
        print(f"Fatal error: {e}")

        if GUI_AVAILABLE and not (os.getenv("HEADLESS", "true").lower() in ("1", "true", "yes", "on") or os.getenv("GITHUB_ACTIONS")):
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            messagebox.showerror("Error", str(e))
            root.destroy()
        else:
            print(f"Fatal error: {e}")

    finally:
        if driver:
            driver.quit()


if __name__ == "__main__":
    main()
