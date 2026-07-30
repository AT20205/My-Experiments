import os
import re
import time
from datetime import datetime
from typing import Dict, Optional

from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait


def update_silver_log(new_log_entries: list) -> None:
    """Logs scraping metrics and errors to the Excel file defined by the SILVERLOG env var,

    placing the newest logs at the top.
    """
    log_file_path = os.getenv("SILVERLOG")
    if not log_file_path:
        print("⚠️ SILVERLOG environment variable not set. Skipping logging.")
        return

    # Ensure the path ends with .xlsx extension
    if not log_file_path.lower().endswith(".xlsx"):
        log_file_path += ".xlsx"

    # Ensure parent directory exists
    log_dir = os.path.dirname(log_file_path)
    if log_dir and not os.path.exists(log_dir):
        os.makedirs(log_dir, exist_ok=True)

    new_df = pd.DataFrame(new_log_entries)

    # Ensure column order is clean and intuitive
    columns_order = [
        "Timestamp",
        "State",
        "Activity",
        "Records Fetched",
        "Status",
        "Error Details",
    ]
    new_df = new_df.reindex(columns=columns_order)

    # Read existing log data if it exists, otherwise write new DataFrame
    if os.path.exists(log_file_path):
        try:
            existing_df = pd.read_excel(log_file_path)
            # Prepend new records ahead of existing records
            combined_df = pd.concat([new_df, existing_df], ignore_index=True)
        except Exception as e:
            print(f"⚠️ Error reading existing log file: {e}. Overwriting file.")
            combined_df = new_df
    else:
        combined_df = new_df

    combined_df.to_excel(log_file_path, index=False)
    print(f"📝 Scraping log updated at: {log_file_path}")


def wait_and_select(
    driver: webdriver.Chrome,
    by: By,
    locator: str,
    value: Optional[str] = None,
    text: Optional[str] = None,
    timeout: int = 15,
):
    """Ensures dropdown options are fully loaded in the DOM before attempting selection."""
    elem = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, locator))
    )

    def options_are_loaded(d):
        try:
            select_obj = Select(elem)
            opts = select_obj.options
            if len(opts) > 1 and any(opt.text.strip() != "" for opt in opts):
                if value is not None:
                    return any(opt.get_attribute("value") == str(value) for opt in opts)
                elif text is not None:
                    return any(opt.text.strip() == str(text) for opt in opts)
                return True
            return False
        except (StaleElementReferenceException, NoSuchElementException):
            return False

    WebDriverWait(driver, timeout).until(options_are_loaded)

    select_obj = Select(elem)
    if value is not None:
        select_obj.select_by_value(str(value))
    elif text is not None:
        select_obj.select_by_visible_text(str(text))

    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));"
        "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));",
        elem,
    )
    return elem


def run_scraper(
    portal_url: str,
    output_base_dir: str,
    states: Dict[str, str],
    activities: Dict[str, str],
    year: str = "2026",
    major_clearance_type: str = "1",
    issue_authority: str = "SEIAA",
    headless: bool = False,
) -> None:
    """Runs the scraper pipeline with runtime-specified configurations."""
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")

    driver = webdriver.Chrome(options=options)
    session_logs = []

    try:
        driver.get(portal_url)
        wait = WebDriverWait(driver, 20)

        # Click Advance Search
        advance_btn = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Show Advance Search')]")
            )
        )
        driver.execute_script("arguments[0].click();", advance_btn)

        # Ensure initial dynamic options are loaded before selecting
        wait.until(
            lambda d: len(
                Select(
                    d.find_element(
                        By.XPATH, "//select[@formcontrolname='majorClearanceType']"
                    )
                ).options
            )
            > 1
        )
        wait_and_select(
            driver,
            By.XPATH,
            "//select[@formcontrolname='majorClearanceType']",
            value=major_clearance_type,
        )

        wait.until(
            lambda d: len(
                Select(
                    d.find_element(
                        By.CSS_SELECTOR, "select[formcontrolname='issueAuthority']"
                    )
                ).options
            )
            > 1
        )
        wait_and_select(
            driver,
            By.CSS_SELECTOR,
            "select[formcontrolname='issueAuthority']",
            value=issue_authority,
        )

        # Nested Loops
        for state_value, state_name in states.items():
            print(
                f"\n🌍 Processing State: {state_name} (Value: {state_value})..."
            )

            all_table_data = []
            headers = []

            for act_value, act_desc in activities.items():
                print(
                    f"  └── ⚙️ Searching Activity ID: {act_value} ({act_desc[:30]}...)"
                )

                records_scraped_before = len(all_table_data)
                last_error_reason = ""

                selection_success = False
                for attempt in range(3):
                    try:
                        WebDriverWait(driver, 15).until(
                            lambda d: len(
                                Select(
                                    d.find_element(
                                        By.XPATH, "//select[@formcontrolname='state']"
                                    )
                                ).options
                            )
                            > 1
                        )
                        wait_and_select(
                            driver,
                            By.XPATH,
                            "//select[@formcontrolname='state']",
                            value=state_value,
                        )
                        time.sleep(0.3)

                        WebDriverWait(driver, 15).until(
                            lambda d: len(
                                Select(
                                    d.find_element(
                                        By.XPATH,
                                        "//select[@formcontrolname='activityId']",
                                    )
                                ).options
                            )
                            > 1
                        )
                        wait_and_select(
                            driver,
                            By.XPATH,
                            "//select[@formcontrolname='activityId']",
                            value=act_value,
                        )
                        time.sleep(0.3)

                        WebDriverWait(driver, 15).until(
                            lambda d: len(
                                Select(
                                    d.find_element(
                                        By.XPATH, "//select[@formcontrolname='year']"
                                    )
                                ).options
                            )
                            > 1
                        )
                        wait_and_select(
                            driver,
                            By.XPATH,
                            "//select[@formcontrolname='year']",
                            value=str(year),
                        )
                        time.sleep(0.3)

                        selection_success = True
                        break
                    except (
                        StaleElementReferenceException,
                        TimeoutException,
                    ) as ex:
                        last_error_reason = f"Selection attempt {attempt+1} failed: {type(ex).__name__} - {str(ex)}"
                        print(
                            f"  ⚠️ Attempt {attempt+1} failed due to loading delay: {ex}. Retrying..."
                        )
                        time.sleep(1)

                if not selection_success:
                    print(
                        f"  ❌ Failed to set selection for Activity {act_value}. Skipping..."
                    )
                    session_logs.append({
                        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "State": state_name,
                        "Activity": act_desc,
                        "Records Fetched": 0,
                        "Status": "Failed",
                        "Error Details": f"Dropdown selection failed after 3 attempts. Last error: {last_error_reason}",
                    })
                    continue

                search_button = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//button[@type='submit' and contains(.,'Search')]",
                        )
                    )
                )

                existing_tables = driver.find_elements(By.ID, "excel-table")
                old_table = existing_tables[0] if existing_tables else None

                driver.execute_script("arguments[0].click();", search_button)

                if old_table:
                    try:
                        wait.until(EC.staleness_of(old_table))
                    except Exception:
                        time.sleep(1)

                try:
                    wait.until(
                        EC.visibility_of_element_located((By.ID, "excel-table"))
                    )
                    time.sleep(0.5)
                except TimeoutException:
                    print(
                        f"  ℹ️ No results found for State: {state_value} + Activity: {act_value}. Proceeding..."
                    )
                    session_logs.append({
                        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "State": state_name,
                        "Activity": act_desc,
                        "Records Fetched": 0,
                        "Status": "No Records",
                        "Error Details": "Table 'excel-table' not visible within timeout (likely 0 results).",
                    })
                    continue

                # PAGINATION
                page_num = 1
                pagination_error = None

                while True:
                    rows_elements = driver.find_elements(
                        By.XPATH, "//table[@id='excel-table']/tbody/tr"
                    )

                    if not rows_elements:
                        break

                    first_row_text = rows_elements[0].text.lower()
                    if (
                        "no record" in first_row_text
                        or "no data" in first_row_text
                    ):
                        break

                    if not headers:
                        html = driver.page_source
                        soup = BeautifulSoup(html, "html.parser")
                        table = soup.find("table", {"id": "excel-table"})
                        if table and table.find("thead"):
                            for th in table.find("thead").find_all("th"):
                                headers.append(th.get_text(strip=True))

                    total_rows = len(rows_elements)
                    print(
                        f"  📊 Page {page_num}: Scraping {total_rows} records..."
                    )

                    for row in rows_elements:
                        try:
                            cols = row.find_elements(By.TAG_NAME, "td")
                            row_data = [col.text.strip() for col in cols]

                            if not row_data or len(row_data) <= 1:
                                continue

                            row_data.append(state_name)
                            row_data.append(act_desc)
                            all_table_data.append(row_data)
                        except Exception:
                            continue

                    # Pagination control
                    try:
                        next_buttons = driver.find_elements(
                            By.XPATH, "//button[@aria-label='Next page']"
                        )
                        if not next_buttons:
                            break

                        next_btn = next_buttons[0]

                        is_disabled = (
                            next_btn.get_attribute("disabled")
                            in ["true", "disabled", True]
                            or next_btn.get_attribute("aria-disabled") == "true"
                            or "mat-button-disabled"
                            in (next_btn.get_attribute("class") or "")
                        )

                        if is_disabled:
                            print(
                                f"  🎉 Reached final page ({page_num}) for Activity {act_value}."
                            )
                            break

                        current_signature = (
                            rows_elements[0].text if rows_elements else ""
                        )

                        driver.execute_script(
                            "arguments[0].click();", next_btn
                        )
                        page_num += 1

                        def wait_for_page_transition(d):
                            try:
                                new_rows = d.find_elements(
                                    By.XPATH,
                                    "//table[@id='excel-table']/tbody/tr",
                                )
                                if not new_rows:
                                    return False
                                return new_rows[0].text != current_signature
                            except (
                                StaleElementReferenceException,
                                NoSuchElementException,
                            ):
                                return False

                        WebDriverWait(driver, 20).until(
                            wait_for_page_transition
                        )
                        time.sleep(0.5)

                    except TimeoutException:
                        pagination_error = f"Timeout waiting for page {page_num} transition."
                        print(
                            f"  ⚠️ Timeout on page {page_num}. Moving to next activity..."
                        )
                        break
                    except Exception as ex:
                        pagination_error = f"Exception on page {page_num}: {type(ex).__name__} - {str(ex)}"
                        print(
                            f"  ⚠️ Pagination exception on page {page_num}: {ex}"
                        )
                        break

                activity_records_count = len(all_table_data) - records_scraped_before
                
                status = "Success" if not pagination_error else "Partial Success / Interrupted"
                session_logs.append({
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "State": state_name,
                    "Activity": act_desc,
                    "Records Fetched": activity_records_count,
                    "Status": status,
                    "Error Details": pagination_error if pagination_error else "None",
                })

                time.sleep(0.3)

            # Save Results Per State
            if all_table_data:
                expected_header_count = len(all_table_data[0])

                if len(headers) < expected_header_count:
                    headers.append("State_Name")
                if len(headers) < expected_header_count:
                    headers.append("Activity Description")

                df = pd.DataFrame(
                    all_table_data, columns=headers[:expected_header_count]
                )

                illegal_xml_chars_re = re.compile(
                    r"[\x00-\x08\x0B-\x0C\x0E-\x1F]"
                )
                df = df.map(
                    lambda x: (
                        illegal_xml_chars_re.sub("", x)
                        if isinstance(x, str)
                        else x
                    )
                )

                print(f"\n--- Extracted Dataset Preview for {state_name} ---")
                print(df.head())

                clean_state_name = (
                    re.sub(r"[^\w\s-]", "", state_name)
                    .strip()
                    .replace(" ", "_")
                )
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                state_filename = f"SEIAA_{clean_state_name}_{timestamp}.xlsx"

                date_folder = datetime.now().strftime("%Y-%m-%d")
                output_dir = os.path.join(output_base_dir, date_folder)
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir, exist_ok=True)

                file_path = os.path.join(output_dir, state_filename)
                df.to_excel(file_path, index=False)
                print(f"✅ Total {len(df)} rows scraped and saved to {file_path}")
            else:
                print(f"❌ Zero data entries found for State: {state_name}")

    finally:
        driver.quit()
        if session_logs:
            update_silver_log(session_logs)