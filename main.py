import calendar
import configparser
import logging
import time
import traceback
from datetime import date, datetime
from functools import lru_cache
from pathlib import Path
import sys

from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent  # folder of the .exe
else:
    BASE_DIR = Path(__file__).resolve().parent

CONFIG_PATH = BASE_DIR / "config.ini"
EXCEL_PATH = BASE_DIR / "Edap.xlsx"
LD_LIST_PATH = BASE_DIR / "LD_List.xlsx"  # known LDs: first column, data from row 2
LOG_PATH = BASE_DIR / "Execution_Log.txt"
BROWSER_DIR = BASE_DIR / "browser"
SUBMITTED_DIR = BASE_DIR / "Submitted Data"

EARLIEST_MONTH = (2025, 5)  # never go back before May 2025
MAX_MONTHS = 24  # 2 years

NO_DATA_XPATH = "//span[contains(text(), 'No data found')]"
GENERATE_REQUEST_XPATH = '//button[text()="Generate Request"]'

logger = logging.getLogger("EDAP")


def initialize_logging():
    """Initialize logging to the execution text file."""
    logger.setLevel(logging.INFO)
    logger.propagate = False

    for handler in logger.handlers[:]:
        handler.close()
        logger.removeHandler(handler)

    handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    logger.addHandler(handler)
    logger.info("EDAP automation started.")


def read_excel_function():
    """Read CNICs from the first column, starting from row 2."""
    if not EXCEL_PATH.is_file():
        raise FileNotFoundError(f"Excel file not found: {EXCEL_PATH}")

    workbook = load_workbook(EXCEL_PATH, read_only=True, data_only=True)

    try:
        worksheet = workbook.active
        records = []

        for (value,) in worksheet.iter_rows(
            min_row=2, min_col=1, max_col=1, values_only=True
        ):
            if value is None:
                continue

            if isinstance(value, float) and value.is_integer():
                value = int(value)

            cnic = str(value).strip()
            if cnic:
                records.append(cnic)

        logger.info("Total CNIC records loaded: %s", len(records))
        print(f"Total CNICs loaded: {len(records)}")
        return records

    finally:
        workbook.close()


def read_known_lds_function():

    if not LD_LIST_PATH.is_file():
        raise FileNotFoundError(f"LD list file not found: {LD_LIST_PATH}")

    workbook = load_workbook(LD_LIST_PATH, read_only=True, data_only=True)

    try:
        worksheet = workbook.active
        known_lds = set()

        for (value,) in worksheet.iter_rows(
            min_row=2, min_col=1, max_col=1, values_only=True
        ):
            if value is None:
                continue

            ld_number = " ".join(str(value).split()).upper()
            if ld_number:
                known_lds.add(ld_number)

        logger.info("Total LD numbers loaded: %s", len(known_lds))
        print(f"Total LD numbers loaded: {len(known_lds)}")
        return known_lds

    finally:
        workbook.close()


def browser_launch(element_wait_time):
    """Launch bundled Chrome and return the driver and explicit wait."""
    driver_files = sorted(BROWSER_DIR.rglob("chromedriver.exe"))
    chrome_files = sorted(BROWSER_DIR.rglob("chrome.exe"))

    if not driver_files: 
        raise FileNotFoundError("chromedriver.exe was not found in the browser folder.")

    options = Options()
    options.add_argument("--start-maximized")
    options.page_load_strategy = (
        "eager"  # Continue once the DOM is ready, don't wait for full page load
    )

    # Keep Chrome open after the Python script exits
    options.add_experimental_option("detach", True)

    if chrome_files:
        options.binary_location = str(chrome_files[0])

    service = Service(executable_path=str(driver_files[0]))
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.command_executor._client_config.timeout = 600
    except AttributeError:
        logger.warning("Could not raise the Selenium command timeout")

    element_wait = WebDriverWait(driver, element_wait_time)

    logger.info("Chrome browser launched successfully.")
    print("Chrome browser launched successfully.")
    return driver, element_wait


def activities(driver, element_wait, edap_tab):
    """Navigate to Add Amendment Request in the EDAP portal."""
    driver.switch_to.window(edap_tab)

    element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="B65401654847547531"]'))
    ).click()
    time.sleep(5)  # wait for the page to load

    logger.info("Add Amendment Request opened.")
    print("Add Amendment Request opened.")


def open_add_amendment_page(driver, element_wait, edap_tab, edap_url):
    """Go back to the EDAP start page and click Add Amendment Request again."""
    driver.switch_to.window(edap_tab)
    driver.get(edap_url)
    activities(driver, element_wait, edap_tab)


@lru_cache(maxsize=1)
def get_months():
    """Latest completed months, newest first.

    At most MAX_MONTHS (24) months, and never before EARLIEST_MONTH (May 2025).
    Calculated once per run so the whole run (and the Excel report) use the
    same window.
    """
    months = []
    today = date.today()

    year = today.year
    month = today.month - 1

    if month == 0:
        month = 12
        year -= 1

    for _ in range(MAX_MONTHS):
        if (year, month) < EARLIEST_MONTH:
            break

        months.append((year, month))
        month -= 1

        if month == 0:
            month = 12
            year -= 1

    return tuple(months)


def month_label(year, month):
    """Build the dropdown text: (2026, 10) -> '31-OCT-26', (2026, 2) -> '28-FEB-26'."""
    last_day = calendar.monthrange(year, month)[1]
    month_abbr = date(year, month, 1).strftime("%b").upper()
    return f"{last_day:02d}-{month_abbr}-{str(year)[-2:]}"


def prepare_report_path(run_started):
    folder = SUBMITTED_DIR / run_started.strftime("%d %b %Y")  # e.g. 23 Nov 2026
    folder.mkdir(parents=True, exist_ok=True)
    report_path = folder / f"Submitted_Data_{run_started.strftime('%H%M%S')}.xlsx"
    logger.info("Report file: %s", report_path)
    return report_path


def write_failed_log(report_path, index, total, cnic, statuses, error):
    """Append the failure reason to Failed_Log_<time>.txt next to the Excel report.

    The file is only created when something fails.
    """
    try:
        failed_path = report_path.with_name(
            report_path.name.replace("Submitted_Data_", "Failed_Log_").replace(
                ".xlsx", ".txt"
            )
        )
        failed_months = [label for label, s in statuses.items() if s == "failed"]
        message = str(error).strip().splitlines()
        reason = message[0] if message else ""

        with open(failed_path, "a", encoding="utf-8") as file:
            file.write(f"Time    : {datetime.now():%Y-%m-%d %H:%M:%S}\n")
            file.write(f"CNIC    : {cnic} (row {index}/{total})\n")
            file.write(f"Month(s): {', '.join(failed_months) or 'unknown'}\n")
            file.write(f"Reason  : {type(error).__name__}: {reason}\n")
            file.write("Traceback:\n")
            file.write(
                "".join(
                    traceback.format_exception(type(error), error, error.__traceback__)
                )
            )
            file.write("-" * 80 + "\n\n")

        logger.info("Failure details saved: %s", failed_path)
    except Exception:
        logger.exception("Could not write the failed log file.")


def save_report_function(report_path, report):

    if not report:
        return

    try:
        months = get_months()
        labels = [month_label(year, month) for year, month in months]
        headers = [date(year, month, 1).strftime("%b %y") for year, month in months]

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Submitted Data"

        sheet.cell(row=1, column=1, value="CNIC")
        for column, header in enumerate(headers, start=2):
            sheet.cell(row=1, column=column, value=header)

        for row, (cnic, statuses) in enumerate(report, start=2):
            sheet.cell(row=row, column=1, value=str(cnic))
            for column, label in enumerate(labels, start=2):
                sheet.cell(
                    row=row, column=column, value=statuses.get(label, "not processed")
                )

        sheet.column_dimensions["A"].width = 18
        for column in range(2, len(headers) + 2):
            letter = sheet.cell(row=1, column=column).column_letter
            sheet.column_dimensions[letter].width = 20
        sheet.freeze_panes = "B2"

        workbook.save(report_path)
        logger.info("Report saved: %s", report_path)

    except Exception:
        # Never let a report problem (for example the file open in Excel) stop the run
        logger.exception("Could not save the report file.")


DONE_PREFIXES = (
    "cleared",
    "LD already in list",
    "no LD account",
    "no information found",
)


def is_done_status(status):
    """A month is finished for good (nothing to redo) for these statuses."""
    return isinstance(status, str) and status.startswith(DONE_PREFIXES)


def load_previous_progress():
    """Read every earlier report and merge it: {cnic: {month_label: status}}.

    Files are read oldest to newest; only finished statuses are kept, so a
    month that was 'failed' / 'not processed' is retried on the next run.
    """
    progress = {}
    files = sorted(
        SUBMITTED_DIR.glob("*/Submitted_Data_*.xlsx"),
        key=lambda p: p.stat().st_mtime,
    )

    for file in files:
        try:
            workbook = load_workbook(file, read_only=True, data_only=True)
        except Exception:
            logger.exception("Could not read previous report: %s", file)
            continue

        try:
            rows = workbook.active.iter_rows(values_only=True)
            header = next(rows, None)
            if not header:
                continue

            labels = {}  # column index -> month label
            for col, text in enumerate(header):
                if col == 0 or not text:
                    continue
                try:
                    parsed = datetime.strptime(str(text), "%b %y")
                except ValueError:
                    continue
                labels[col] = month_label(parsed.year, parsed.month)

            for row in rows:
                if not row or row[0] is None:
                    continue
                cnic = str(row[0]).strip()
                for col, label in labels.items():
                    if col < len(row) and is_done_status(row[col]):
                        progress.setdefault(cnic, {})[label] = row[col]
        except Exception:
            logger.exception("Could not parse previous report: %s", file)
        finally:
            workbook.close()

    logger.info("Previous progress loaded for %s CNICs.", len(progress))
    print(f"Previous progress loaded for {len(progress)} CNICs.")
    return progress


def clear_p76_fields(driver, element_wait):

    product_code = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//select[@id="P76_PRODUCT_CODE"]'))
    )
    Select(product_code).select_by_value("")

    product_relationship = element_wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//input[@name="P76_CREDIT_RELATIONSHIP_DATE"]')
        )
    )
    product_relationship.clear()

    loan_data = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_LOAN_DATE"]'))
    )
    loan_data.clear()

    DISBURSEMENT_DATE = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_DISBURSEMENT_DATE"]'))
    )
    DISBURSEMENT_DATE.clear()

    LOAN_TYPE = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//select[@id="P76_LOAN_TYPE"]'))
    )
    Select(LOAN_TYPE).select_by_value("")

    MATURITY_DATE = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_MATURITY_DATE"]'))
    )
    MATURITY_DATE.clear()

    SECURED = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//select[@id="P76_SECURED"]'))
    )
    Select(SECURED).select_by_value("")

    TOTAL_LIMIT_AMT = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_TOTAL_LIMIT_AMT"]'))
    )
    TOTAL_LIMIT_AMT.clear()

    TOTAL_PRINCIPAL_OUTSTAND_A = element_wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="P76_TOTAL_PRINCIPAL_OUTSTAND_A"]')
        )
    )
    TOTAL_PRINCIPAL_OUTSTAND_A.clear()

    TOTAL_MARKUP_AMT = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_TOTAL_MARKUP_AMT"]'))
    )
    TOTAL_MARKUP_AMT.clear()

    TOTAL_OTHER_OUTSTAND_AMT = element_wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="P76_TOTAL_OTHER_OUTSTAND_AMT"]')
        )
    )
    TOTAL_OTHER_OUTSTAND_AMT.clear()

    PERIODICITY_OF_PAYMENT = element_wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="P76_PERIODICITY_OF_PAYMENT"]')
        )
    )
    PERIODICITY_OF_PAYMENT.clear()

    LAST_PAYMENT_DATE = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_LAST_PAYMENT_DATE"]'))
    )
    LAST_PAYMENT_DATE.clear()

    MINIMUM_AMT_DUE = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_MINIMUM_AMT_DUE"]'))
    )
    MINIMUM_AMT_DUE.clear()

    OVERDUE_30_PLUS_AMT = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_OVERDUE_30_PLUS_AMT"]'))
    )
    OVERDUE_30_PLUS_AMT.clear()

    OVERDUE_60_PLUS_AMT = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_OVERDUE_60_PLUS_AMT"]'))
    )
    OVERDUE_60_PLUS_AMT.clear()

    OVERDUE_90_PLUS_AMT = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_OVERDUE_90_PLUS_AMT"]'))
    )
    OVERDUE_90_PLUS_AMT.clear()

    OVERDUE_120_PLUS_AMT = element_wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="P76_OVERDUE_120_PLUS_AMT"]')
        )
    )
    OVERDUE_120_PLUS_AMT.clear()

    OVERDUE_150_PLUS_AMT = element_wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="P76_OVERDUE_150_PLUS_AMT"]')
        )
    )
    OVERDUE_150_PLUS_AMT.clear()

    OVERDUE_180_PLUS_AMT = element_wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="P76_OVERDUE_180_PLUS_AMT"]')
        )
    )
    OVERDUE_180_PLUS_AMT.clear()

    DUE_SINCE = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_DUE_SINCE"]'))
    )
    DUE_SINCE.clear()

    RESTRUCTURED_AMT = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_RESTRUCTURED_AMT"]'))
    )
    RESTRUCTURED_AMT.clear()

    LITIGATION_AMT = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_LITIGATION_AMT"]'))
    )
    LITIGATION_AMT.clear()

    UNPAID_DAYS_LATE_CODE = element_wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//select[@id="P76_UNPAID_DAYS_LATE_CODE"]')
        )
    )
    Select(UNPAID_DAYS_LATE_CODE).select_by_value("")

    INTEREST_RATE_TYPE = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//select[@id="P76_INTEREST_RATE_TYPE"]'))
    )
    Select(INTEREST_RATE_TYPE).select_by_value("")

    CLASSIFICATION_LOAN_TYPE = element_wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//select[@id="P76_CLASSIFICATION_LOAN_TYPE"]')
        )
    )
    Select(CLASSIFICATION_LOAN_TYPE).select_by_value("")

    LOAN_CLASS = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//select[@id="P76_LOAN_CLASS"]'))
    )
    Select(LOAN_CLASS).select_by_value("")

    LOAN_PROVISION_AMT = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_LOAN_PROVISION_AMT"]'))
    )
    LOAN_PROVISION_AMT.clear()

    TOTAL_BOUNCE_CHEQUE = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_TOTAL_BOUNCE_CHEQUE"]'))
    )
    TOTAL_BOUNCE_CHEQUE.clear()

    FACILITY_RATING = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//select[@id="P76_FACILITY_RATING"]'))
    )
    Select(FACILITY_RATING).select_by_value("")

    SBP_SCHEME_TYPE_CODE = element_wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//select[@id="P76_SBP_SCHEME_TYPE_CODE"]')
        )
    )
    Select(SBP_SCHEME_TYPE_CODE).select_by_value("")

    FINANCIAL_MODE = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//select[@id="P76_FINANCIAL_MODE"]'))
    )
    Select(FINANCIAL_MODE).select_by_value("")

    LOAN_PAID_IN_FULL = element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="P76_LOAN_PAID_IN_FULL"]'))
    )
    LOAN_PAID_IN_FULL.clear()


def amendment_form_function(driver, element_wait, known_lds):

    # Generate Request
    element_wait.until(
        EC.element_to_be_clickable((By.XPATH, GENERATE_REQUEST_XPATH))
    ).click()

    time.sleep(10)

    report_xpath = '//*[@id="report_R65726130771184273"]'

    # STEP 1: Identify all Loan Account IDs starting with LD.
    report = element_wait.until(
        EC.presence_of_element_located((By.XPATH, report_xpath))
    )
    rows = report.find_elements(By.XPATH, "./tbody/tr")
    ld_accounts = []

    for row in rows:
        cells = row.find_elements(By.XPATH, "./td")
        for cell in cells:
            account_id = " ".join(cell.text.split())
            if account_id.upper().startswith("LD"):
                ld_accounts.append(account_id)
                break

    logger.info("Total LD loan accounts identified: %s", len(ld_accounts))
    print(f"LD loan accounts identified: {len(ld_accounts)}")

    # Do not submit if no qualifying LD account exists.
    if not ld_accounts:
        logger.warning("No LD loan accounts found. No amendment was submitted.")
        print("No LD loan accounts found. Skipping submission.")
        return "no LD account"

    # LDs already in the known LD list are skipped; only the others are processed.
    skipped_lds = [ld for ld in ld_accounts if ld.upper() in known_lds]
    pending_lds = [ld for ld in ld_accounts if ld.upper() not in known_lds]

    if skipped_lds:
        logger.info("LD already in the list, skipped: %s", skipped_lds)
        print(f"LD already in the list, skipped: {skipped_lds}")

    if not pending_lds:
        logger.info(
            "All LD accounts are already in the list. Skipping this month, "
            "nothing submitted."
        )
        print("All LD accounts are already in the list. Skipping this month.")
        return "LD already in list"

    # STEP 2: Process each LD account row.
    for account_number, account_id in enumerate(pending_lds, start=1):
        logger.info(
            "Processing LD account %s/%s: %s",
            account_number,
            len(pending_lds),
            account_id,
        )
        print(
            f"Processing LD account {account_number}/{len(pending_lds)}: {account_id}"
        )

        # Re-find the report row after every Apply Changes.
        report = element_wait.until(
            EC.presence_of_element_located((By.XPATH, report_xpath))
        )
        rows = report.find_elements(By.XPATH, "./tbody/tr")
        account_cell = None

        for row in rows:
            cells = row.find_elements(By.XPATH, "./td")
            for cell in cells:
                if " ".join(cell.text.split()) == account_id:
                    account_cell = cell
                    break
            if account_cell is not None:
                break

        if account_cell is None:
            raise RuntimeError(
                f"Could not find the report row for LD account: {account_id}"
            )

        # Pencil link in the nearest preceding cell containing an anchor.
        pencil_links = account_cell.find_elements(
            By.XPATH,
            "./preceding-sibling::td[.//a][1]//a",
        )
        if len(pencil_links) != 1:
            raise RuntimeError(
                f"Expected one pencil link for {account_id}, "
                f"but found {len(pencil_links)}."
            )

        pencil_links[0].click()
        time.sleep(5)

        clear_p76_fields(driver, element_wait)

        element_wait.until(
            EC.element_to_be_clickable((By.XPATH, '//input[@id="B65744856223196882"]'))
        ).click()

        element_wait.until(EC.presence_of_element_located((By.XPATH, report_xpath)))
        time.sleep(5)
        logger.info("Apply Changes completed for LD account: %s", account_id)

    # STEP 3: Submit once after all LD accounts have been processed.
    element_wait.until(
        EC.element_to_be_clickable((By.XPATH, '//input[@id="B65521141841049322"]'))
    ).click()
    time.sleep(5)

    logger.info(
        "All %s LD loan accounts processed. Final submission completed.",
        len(pending_lds),
    )
    print("All LD accounts processed. Final submission completed.")

    if skipped_lds:
        return f"cleared ({len(skipped_lds)} LD in list skipped)"
    return "cleared"


def each_entry_activity_function(
    driver,
    element_wait,
    edap_tab,
    edap_url,
    cnic,
    index,
    total,
    statuses,
    known_lds,
    save_progress,
):

    logger.info("Processing CNIC %s/%s: %s", index, total, cnic)

    on_search_page = False
    months_done = []
    months_no_ld = []
    months_in_list = []

    for year, month in get_months():

        label = month_label(year, month)

        # Resume: month finished in an earlier run, do not repeat it
        if is_done_status(statuses.get(label)):
            logger.info(
                "CNIC %s: %s already done earlier (%s), skipping.",
                cnic,
                label,
                statuses[label],
            )
            print(f"CNIC {cnic}: {label} already done, skipping.")
            continue

        statuses[label] = "failed"

        # Every month that follows a submitted month starts from Add Amendment Request
        if not on_search_page:
            open_add_amendment_page(driver, element_wait, edap_tab, edap_url)
            on_search_page = True

        logger.info("Processing CNIC %s for %s", cnic, label)
        print(f"CNIC: {cnic} | Month: {label}")

        cnic_field = element_wait.until(
            EC.element_to_be_clickable((By.XPATH, '//input[@id="P70_NEW_NIC"]'))
        )

        cnic_field.clear()
        cnic_field.send_keys(cnic)

        date_dropdown = element_wait.until(
            EC.presence_of_element_located((By.XPATH, '//select[@id="P70_AS_OF_DATE"]'))
        )

        date_select = Select(date_dropdown)

        matching_option = next(
            (
                option
                for option in date_select.options
                if option.text.strip().upper() == label
            ),
            None,
        )

        if matching_option is None:
            available = [option.text.strip() for option in date_select.options]
            raise RuntimeError(
                f"{label} not in dropdown for CNIC {cnic}. Available: {available}"
            )

        date_select.select_by_visible_text(matching_option.text)

        # Search
        search_button = element_wait.until(
            EC.element_to_be_clickable((By.XPATH, '//input[@value="Search"]'))
        )
        search_button.click()

        # Make sure the old page is gone so an old 'No data found' is not read
        try:
            WebDriverWait(driver, 10).until(EC.staleness_of(search_button))
        except TimeoutException:
            pass

        # time.sleep(5)  # wait for the search result to load

        # Wait for either result: "No data found" or the Generate Request button
        element_wait.until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, NO_DATA_XPATH)),
                EC.element_to_be_clickable((By.XPATH, GENERATE_REQUEST_XPATH)),
            )
        )

        if driver.find_elements(By.XPATH, NO_DATA_XPATH):
            logger.info(
                "CNIC %s: 'No data found' for %s, moving to previous month.",
                cnic,
                label,
            )
            print(f"No data found for {label}, trying previous month.")
            statuses[label] = "no information found"
            save_progress()
            continue

        logger.info("Search submitted for entry %s. Data found for %s.", index, label)
        print(f"Search submitted for entry {index}. Data found for {label}.")

        result = amendment_form_function(driver, element_wait, known_lds)
        statuses[label] = result
        save_progress()

        if result.startswith("cleared"):
            months_done.append(label)
            logger.info("CNIC %s: %s submitted.", cnic, label)
            print(f"CNIC {cnic}: {label} submitted.")
        elif result == "LD already in list":
            months_in_list.append(label)
            logger.info(
                "CNIC %s: %s skipped, LD already in list. Moving to previous month.",
                cnic,
                label,
            )
            print(f"CNIC {cnic}: {label} skipped, LD already in list.")
        else:
            months_no_ld.append(label)
            logger.warning(
                "CNIC %s: %s has no LD account. Nothing submitted, moving to previous month.",
                cnic,
                label,
            )
            print(f"CNIC {cnic}: {label} has no LD account, trying previous month.")

        # The page has moved on; next month starts from Add Amendment Request again
        on_search_page = False

    logger.info(
        "CNIC %s finished. Submitted: %s | LD already in list (skipped): %s | "
        "No LD account (skipped): %s",
        cnic,
        months_done,
        months_in_list,
        months_no_ld,
    )

    if not months_done and not months_in_list and not months_no_ld:
        logger.warning("CNIC %s finished. No data found in any month.", cnic)
        print(f"CNIC {cnic}: no data found in any month.")


def each_entry_row_function(
    driver, element_wait, edap_tab, edap_url, records, report_path, known_lds
):

    total = len(records)
    logger.info("Starting CNIC processing. Total records: %s", total)

    succeeded = 0
    failed = 0
    report = []  # list of (cnic, {month_label: status})
    previous = load_previous_progress()
    months = [month_label(y, m) for y, m in get_months()]

    def save_progress():
        save_report_function(report_path, report)

    for index, cnic in enumerate(records, start=1):
        statuses = dict(previous.get(cnic, {}))  # carry over finished months
        report.append((cnic, statuses))

        # Resume: every month already finished -> skip this CNIC entirely
        if all(is_done_status(statuses.get(label)) for label in months):
            logger.info("CNIC %s/%s %s already complete, skipping.", index, total, cnic)
            print(f"CNIC {index}/{total} {cnic}: already complete, skipping.")
            succeeded += 1
            continue

        try:
            driver.switch_to.window(edap_tab)
            each_entry_activity_function(
                driver,
                element_wait,
                edap_tab,
                edap_url,
                cnic,
                index,
                total,
                statuses,
                known_lds,
                save_progress,
            )
            succeeded += 1

        except Exception as error:
            failed += 1
            logger.exception("EachEntry_Activity_Function failed at row %s.", index)
            write_failed_log(report_path, index, total, cnic, statuses, error)
            print(f"Entry {index} failed. Check Execution_Log.txt.")
            break

        finally:
            # Save after every CNIC (also after a failure) so nothing is lost
            save_report_function(report_path, report)

    not_attempted = total - succeeded - failed

    logger.info(
        "CNIC row processing finished: %s succeeded, %s failed, %s not attempted.",
        succeeded,
        failed,
        not_attempted,
    )
    return succeeded, failed, not_attempted


def main():

    initialize_logging()

    try:
        logger.info("Main function started.")

        if not CONFIG_PATH.is_file():
            raise FileNotFoundError(f"Configuration file not found: {CONFIG_PATH}")

        config = configparser.ConfigParser()
        config.read(CONFIG_PATH, encoding="utf-8")

        sbp_url = config.get("PORTALS", "SBP_Portal").strip()
        edap_url = config.get("PORTALS", "EDAP_Portal").strip()
        element_wait_time = config.getint("SETTINGS", "Element_Wait_Time")

        records = read_excel_function()
        if not records:
            logger.warning("No CNIC records found in Edap.xlsx.")
            print("No CNIC records found.")
            return

        known_lds = read_known_lds_function()

        driver, element_wait = browser_launch(element_wait_time)

        try:
            logger.info("Opening SBP portal.")
            driver.get(sbp_url)
            print("SBP portal opened. Waiting 3 minutes for login...")

            element_wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//*[contains(text(), 'Welcome to Data Submission')]")
                )
            )
            # set wait for element here
            logger.info("SBP login wait completed.")

            logger.info("Opening EDAP portal in a second tab.")
            driver.switch_to.new_window("tab")
            edap_tab = driver.current_window_handle
            driver.get(edap_url)

            print("EDAP portal opened.")
            logger.info("EDAP login confirmed.")

        except Exception:
            logger.exception("Portal login failed; login was unsuccessful.")
            print("Portal login was unsuccessful. Check Execution_Log.txt.")
            return

        report_path = prepare_report_path(datetime.now())

        # Add Amendment Request is opened per month inside each_entry_activity_function
        succeeded, failed, not_attempted = each_entry_row_function(
            driver, element_wait, edap_tab, edap_url, records, report_path, known_lds
        )

        if failed == 0 and not_attempted == 0:
            logger.info("Main function completed successfully.")
            print("\nEDAP processing completed.")
        else:
            logger.warning(
                "Main function completed with errors: %s succeeded, %s failed, %s not attempted.",
                succeeded,
                failed,
                not_attempted,
            )
            print(
                f"\nEDAP processing finished with errors: {succeeded} succeeded, "
                f"{failed} failed, {not_attempted} not attempted."
            )

    except Exception:
        logger.exception("Main function failed.")
        print("Automation failed. Check Execution_Log.txt.")

    finally:
        logger.info("Automation finished. Browser remains open for manual review.")
        print("Browser remains open for manual review.")


if __name__ == "__main__":
    main()
