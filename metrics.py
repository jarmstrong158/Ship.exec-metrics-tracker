# ================================
# MASTER PROGRAM – REPORT RUNNER
# ================================

import sys
import os
import re
import json
import time
import hashlib
from collections import Counter

import pandas as pd
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURATION ---

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "metrics_config.json")


def load_config():
    """Load configuration from JSON file, or return None if not found."""
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            print(f"Warning: Could not read config file: {e}")
    return None


def save_config(config):
    """Save configuration to JSON file."""
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=2)


def setup_config():
    """Interactive first-time configuration."""
    print("\n=== First-Time Setup ===")
    print("No configuration file found. Setting up regions.\n")

    config = {"regions": {}, "sort_preference": "last_initial"}

    print("Default regions: NE (ne@), SE (se@), CTR (ctr@), WC (wc@)")
    use_defaults = input("Press Enter to use defaults, or type 'custom' to set up manually: ").strip().lower()

    if use_defaults != "custom":
        config["regions"] = {
            "NE":  {"email_pattern": "ne@",  "order": []},
            "SE":  {"email_pattern": "se@",  "order": []},
            "CTR": {"email_pattern": "ctr@", "order": []},
            "WC":  {"email_pattern": "wc@",  "order": []}
        }
    else:
        print("\nEnter regions one at a time. Type 'done' when finished.")
        while True:
            name = input("Region name (e.g., NE): ").strip().upper()
            if name.lower() == "done":
                break
            if not name:
                continue
            pattern = input(f"Email pattern for {name} (e.g., ne@): ").strip().lower()
            config["regions"][name] = {"email_pattern": pattern, "order": []}

    print("\nSort preference for initials:")
    print("  1 = Alphabetical by last initial (default)")
    print("  2 = Custom order (you'll set per region after first scrape)")
    sort_choice = input("Choice (1/2, default=1): ").strip()
    if sort_choice == "2":
        config["sort_preference"] = "custom"

    save_config(config)
    print(f"Configuration saved to: {CONFIG_PATH}\n")
    return config


def manage_regions(config):
    """Interactive region management menu."""
    while True:
        print("\n=== Region Management ===")
        print("Current regions:")
        for name, info in config["regions"].items():
            order_str = ", ".join(info.get("order", [])) or "(not set yet — will auto-detect on next run)"
            print(f"  {name} — email pattern: {info['email_pattern']} — order: {order_str}")

        print(f"\nSort preference: {config.get('sort_preference', 'last_initial')}")
        print("\nOptions:")
        print("  1 = Add a region")
        print("  2 = Remove a region")
        print("  3 = Edit a region's initial order")
        print("  4 = Change sort preference")
        print("  5 = Back to main menu")

        choice = input("Choice: ").strip()

        if choice == "1":
            name = input("Region name: ").strip().upper()
            if not name:
                continue
            pattern = input(f"Email pattern for {name} (e.g., ne@): ").strip().lower()
            config["regions"][name] = {"email_pattern": pattern, "order": []}
            save_config(config)
            print(f"Region '{name}' added.")

        elif choice == "2":
            name = input("Region to remove: ").strip().upper()
            if name in config["regions"]:
                confirm = input(f"Remove region '{name}'? (y/n): ").strip().lower()
                if confirm == "y":
                    del config["regions"][name]
                    save_config(config)
                    print(f"Region '{name}' removed.")
            else:
                print(f"Region '{name}' not found.")

        elif choice == "3":
            name = input("Region to edit: ").strip().upper()
            if name in config["regions"]:
                current = ", ".join(config["regions"][name].get("order", []))
                print(f"Current order: {current or '(not set)'}")
                new_order = input("New order (space-separated initials, or Enter to keep): ").strip()
                if new_order:
                    config["regions"][name]["order"] = new_order.upper().split()
                    save_config(config)
                    print("Order updated.")
            else:
                print(f"Region '{name}' not found.")

        elif choice == "4":
            print(f"Current: {config.get('sort_preference', 'last_initial')}")
            print("  1 = Alphabetical by last initial")
            print("  2 = Custom order")
            pref = input("Choice: ").strip()
            if pref == "1":
                config["sort_preference"] = "last_initial"
            elif pref == "2":
                config["sort_preference"] = "custom"
            save_config(config)
            print("Preference updated.")

        elif choice == "5":
            break

    return config


# --- UTILITY FUNCTIONS ---

def safe_to_int(value):
    try:
        if pd.isna(value): return 0
        if isinstance(value, (int, float)): return int(value)
        s = str(value).replace(",", "").strip()
        if s == "": return 0
        if "." in s: return int(float(s))
        return int(s)
    except:
        m = re.search(r"-?\d+", str(value))
        return int(m.group()) if m else 0


def sort_by_last_initial(initials):
    """Sort initials alphabetically by last character, then by preceding characters."""
    return sorted(initials, key=lambda x: (x[-1].upper(), x[:-1].upper()) if x else ("", ""))


def find_similar_initials(init, known_list):
    """Find known initials that look similar to an unknown one."""
    similar = []
    init_upper = init.upper()
    for k in known_list:
        k_upper = k.upper()
        if not k_upper:
            continue
        # One is a prefix/suffix of the other
        if init_upper.startswith(k_upper) or k_upper.startswith(init_upper):
            similar.append(k)
        # Same length, differ by one character
        elif len(init_upper) == len(k_upper):
            diffs = sum(1 for a, b in zip(init_upper, k_upper) if a != b)
            if diffs == 1:
                similar.append(k)
    return similar


# --- REGION AND INITIAL FUNCTIONS ---

def region_from_email(email, config):
    """Determine region from email address using config's email patterns."""
    if not isinstance(email, str):
        return None
    e = email.strip().lower()
    for region, info in config["regions"].items():
        if info["email_pattern"].lower() in e:
            return region
    return None


def consignee_tokens_in_order(field):
    if not isinstance(field, str):
        return []
    found = re.findall(r"[A-Za-z]+", field)
    return [f.upper() for f in found]


def process_region_counts(df_final, config):
    """Count picked/packed per region from scraped data. Counts all initials found."""
    df_final["Region"] = df_final["Email"].apply(lambda e: region_from_email(e, config))

    regions = list(config["regions"].keys())
    packed_counts = {r: Counter() for r in regions}
    picked_counts = {r: Counter() for r in regions}

    for _, row in df_final.iterrows():
        region = row.get("Region")
        if region not in regions:
            continue

        codes = consignee_tokens_in_order(row.get("Consignee Reference", ""))
        if not codes:
            continue

        if len(codes) == 1:
            code = codes[0]
            packed_counts[region][code] += 1
            picked_counts[region][code] += 1
        else:
            packer = codes[0]
            picker = codes[1]
            packed_counts[region][packer] += 1
            picked_counts[region][picker] += 1

    return df_final, packed_counts, picked_counts


def reconcile_initials_with_config(packed_counts, picked_counts, config):
    """Check for new initials after scraping. Prompts user if new ones are found. Updates config."""
    updated = False

    for region in config["regions"]:
        saved_order = config["regions"][region].get("order", [])
        all_found = set(packed_counts[region].keys()) | set(picked_counts[region].keys())

        # Identify "real" new initials (appear 2+ times or in both packer and picker positions)
        real_new = []
        for init in all_found:
            if init in saved_order:
                continue
            total = packed_counts[region].get(init, 0) + picked_counts[region].get(init, 0)
            in_both = packed_counts[region].get(init, 0) > 0 and picked_counts[region].get(init, 0) > 0
            if total >= 2 or in_both:
                real_new.append(init)

        if not real_new:
            continue

        if not saved_order:
            # First run for this region — all qualifying initials are new
            default_order = sort_by_last_initial(real_new)
            print(f"\n[{region}] Discovered initials: {', '.join(default_order)} (sorted by last initial)")
            print(f"  Press Enter to accept this order, or type your preferred order (space-separated):")
            user_input = input("  > ").strip()
            if user_input:
                config["regions"][region]["order"] = user_input.upper().split()
            else:
                config["regions"][region]["order"] = default_order
            updated = True
        else:
            # Subsequent run — new initials appeared
            print(f"\n[{region}] New initials found: {', '.join(sorted(real_new))}")
            print(f"  Current order: {', '.join(saved_order)}")

            if config.get("sort_preference") == "custom":
                print(f"  Type the full order including new ones, or press Enter to append by last initial:")
                user_input = input("  > ").strip()
                if user_input:
                    config["regions"][region]["order"] = user_input.upper().split()
                else:
                    new_sorted = sort_by_last_initial(real_new)
                    config["regions"][region]["order"] = saved_order + new_sorted
            else:
                # Auto-insert maintaining alphabetical-by-last-initial sort
                new_sorted = sort_by_last_initial(real_new)
                config["regions"][region]["order"] = sort_by_last_initial(saved_order + new_sorted)
            updated = True

    if updated:
        save_config(config)
        print("Configuration updated.\n")

    return config


def classify_initials(packed_counts, picked_counts, config):
    """Separate discovered initials into 'known' (main section) and 'review' (edge cases).

    Returns:
        region_initials: {region: [ordered list of known initials]}
        review_entries:  {region: [(initial, context_string), ...]}
    """
    regions = list(config["regions"].keys())
    region_initials = {}
    review_entries = {}

    for region in regions:
        saved_order = config["regions"][region].get("order", [])
        all_found = set(packed_counts[region].keys()) | set(picked_counts[region].keys())

        known = []
        review = []

        for init in all_found:
            total = packed_counts[region].get(init, 0) + picked_counts[region].get(init, 0)
            in_both = packed_counts[region].get(init, 0) > 0 and picked_counts[region].get(init, 0) > 0

            if init in saved_order:
                known.append(init)
            elif total >= 2 or in_both:
                known.append(init)
            else:
                # Edge case — build context clues
                positions = []
                if packed_counts[region].get(init, 0) > 0:
                    positions.append("packer")
                if picked_counts[region].get(init, 0) > 0:
                    positions.append("picker")

                context = f"{total}x, {' + '.join(positions)} only"

                similar = find_similar_initials(init, known + saved_order)
                if similar:
                    context += f" -- similar to: {', '.join(similar)}"

                review.append((init, context))

        # Include configured initials that weren't found in data (show with 0 counts)
        for init in saved_order:
            if init not in known:
                known.append(init)

        # Sort known initials according to preference
        sort_pref = config.get("sort_preference", "last_initial")
        if sort_pref == "custom" and saved_order:
            # Preserve saved order for configured initials, append new ones sorted
            ordered = [i for i in saved_order if i in known]
            remaining = sort_by_last_initial([i for i in known if i not in saved_order])
            known = ordered + remaining
        else:
            known = sort_by_last_initial(known)

        region_initials[region] = known
        review_entries[region] = review

    return region_initials, review_entries


# --- SUMMARY BUILDING ---

def build_region_summary(picked_counts, packed_counts, boxcounts, region_initials, review_entries):
    """Build the region summary DataFrame with dynamic initials and review zone."""
    regions = list(region_initials.keys())
    max_rows = max((len(v) for v in region_initials.values()), default=0)
    rows = []

    for i in range(max_rows):
        row = []
        for region in regions:
            initials = region_initials[region]
            if i < len(initials):
                init = initials[i]
                row += [init, picked_counts[region].get(init, 0), packed_counts[region].get(init, 0), ""]
            else:
                row += ["", "", "", ""]
        rows.append(row)

    # Totals row (only sums known initials, not review)
    total_row = []
    for region in regions:
        known_picked = sum(picked_counts[region].get(i, 0) for i in region_initials[region])
        known_packed = sum(packed_counts[region].get(i, 0) for i in region_initials[region])
        total_row += ["TOTAL", known_picked, known_packed, ""]
    rows.append(total_row)

    # Boxcount row
    boxcount_row = []
    for region in regions:
        boxcount_row += ["Boxcount", boxcounts.get(region, 0), "", ""]
    rows.append(boxcount_row)

    # Review section
    max_review = max((len(v) for v in review_entries.values()), default=0)
    if max_review > 0:
        sep_row = []
        for region in regions:
            sep_row += ["--- Review ---", "", "", ""]
        rows.append(sep_row)

        for i in range(max_review):
            row = []
            for region in regions:
                entries = review_entries[region]
                if i < len(entries):
                    init, context = entries[i]
                    row += [init, picked_counts[region].get(init, 0), packed_counts[region].get(init, 0), context]
                else:
                    row += ["", "", "", ""]
            rows.append(row)

    columns = []
    for region in regions:
        columns += [region, f"{region}_Picked", f"{region}_Packed", ""]

    return pd.DataFrame(rows, columns=columns)


def get_boxcounts_from_summary(df_summary, config):
    """Extract boxcount values from a previously saved RegionSummary sheet.

    Finds the 'Boxcount' row by searching column 0, and finds each region's
    value column by searching the headers for the region name and taking the
    next column (where the count value lives).
    """
    regions = list(config["regions"].keys())

    # Find the Boxcount row
    box_rows = df_summary[
        df_summary.iloc[:, 0].astype(str).str.strip().str.lower() == "boxcount"
    ].index.tolist()

    if not box_rows:
        print("[Boxcount] Warning: No 'Boxcount' row found in summary.")
        return {r: 0 for r in regions}

    box_row_idx = box_rows[-1]

    # Find each region's value column by header label
    headers = [str(c).strip() for c in df_summary.columns]
    boxcounts = {}

    for region in regions:
        col_indices = [i for i, h in enumerate(headers) if h == region]
        if col_indices:
            value_col = col_indices[0] + 1  # value is in the column after the label
            if value_col < df_summary.shape[1]:
                boxcounts[region] = safe_to_int(df_summary.iloc[box_row_idx, value_col])
            else:
                boxcounts[region] = 0
        else:
            boxcounts[region] = 0

    return boxcounts


# --- SELENIUM HELPERS ---

def initialize_driver_and_login():
    """Launch Chrome, navigate to ShipExec, and wait for the user to log in manually."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    prefs = {
        "profile.default_content_setting_values.local_network": 1
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://thinclient.shipexec.com/#!/history")

    print("\n==========================================")
    print("  Browser launched. Please log in to")
    print("  ShipExec in the browser window.")
    print("==========================================")
    input("Press Enter here once you have logged in...")

    # Click History tab
    print("Waiting for History tab to become clickable...")
    for attempt in range(3):
        try:
            history_tab = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href="/#!/history"]'))
            )
            history_tab.click()
            print("History tab clicked successfully.")
            break
        except Exception as e:
            print(f"Attempt {attempt+1}: History tab not clickable yet. Retrying...")
            time.sleep(2)
    else:
        print("Failed to click History tab after multiple attempts. Exiting.")
        driver.quit()
        raise SystemExit

    return driver


def click_search_button(driver):
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1.5)
        search_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[ng-click='vm.changeReport()']"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
        time.sleep(0.5)
        search_button.click()
        print("Search button clicked.")
    except Exception as e:
        print(f"Could not click Search button. Error: {e}")
    time.sleep(5)


def increase_table_size_to_100(driver):
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1.5)
        button_100 = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//button[span[text()='100']]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", button_100)
        time.sleep(0.5)
        button_100.click()
        print("Set table size to 100 rows.")
        time.sleep(3)
    except Exception as e:
        print("Could not find or click the '100' button. Using default row count instead.")
        print(f"Error details: {e}")

    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)


def find_and_click_date(driver, target_day):
    """Click a specific day number in the currently open calendar picker."""
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.ng-binding"))
        )
        date_spans = driver.find_elements(By.CSS_SELECTOR, "span.ng-binding")
        for span in date_spans:
            span_day = span.text.strip().lstrip("0")
            if span_day == str(target_day) and "text-muted" not in span.get_attribute("class"):
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", span)
                time.sleep(0.3)
                driver.execute_script("arguments[0].click();", span)
                print(f"Clicked date {target_day}.")
                return True
        return False
    except Exception as e:
        print(f"Error clicking date {target_day}: {e}")
        return False


def click_date_or_fallback(driver, target_day):
    """Try to click target_day; if not found, select next month's '01'."""
    if not find_and_click_date(driver, target_day):
        print(f"Could not locate target day {target_day}, selecting next month's '01'.")
        date_spans = driver.find_elements(By.CSS_SELECTOR, "span.ng-binding")
        for span in date_spans:
            if span.text.strip() == "01":
                driver.execute_script("arguments[0].scrollIntoView(true);", span)
                time.sleep(0.3)
                driver.execute_script("arguments[0].click();", span)
                print("Selected next month's '01'.")
                break


def set_date_range(driver, day_offset):
    """Open the date pickers and set both END and START to today + day_offset."""
    # Set END date
    try:
        end_calendar_button = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "button[ng-click='vm.openendship()']"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", end_calendar_button)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", end_calendar_button)
        print("Clicked END date calendar button.")
    except Exception as e:
        print(f"Could not click END date calendar button: {e}")

    today_elem = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "span.ng-binding.text-info"))
    )
    today_day = int(today_elem.text.strip())
    target_day = today_day + day_offset

    click_date_or_fallback(driver, target_day)
    time.sleep(3)

    # Set START date
    try:
        start_calendar_button = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "button[ng-click='vm.openstartship()']"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", start_calendar_button)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", start_calendar_button)
        print("Clicked START date calendar button.")
    except Exception as e:
        print(f"Could not click START date calendar button: {e}")

    click_date_or_fallback(driver, target_day)
    time.sleep(3)
    print("Date range successfully set.")


# --- SCRAPING ---

def scrape_history_and_detailed(driver, config):
    """Scrape History table, then Detailed Report. Returns (df_final, boxcounts dict)."""
    regions = list(config["regions"].keys())

    # --- SCRAPE HISTORY TABLE ---
    all_tracking_numbers = []
    all_shipper_refs = []
    all_consignee_refs = []
    previous_table_hash = None

    def deduplicate_tracking_number(s):
        if not isinstance(s, str):
            return s
        length = len(s)
        for i in range(1, length // 2 + 1):
            sub = s[:i]
            multiplier = length // i
            if sub * multiplier == s:
                return sub
        return s

    while True:
        table_elem = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "table"))
        )
        time.sleep(1)
        html = driver.execute_script("return document.documentElement.outerHTML;")
        soup = BeautifulSoup(html, "html.parser")
        table = soup.find("table")
        if table is None:
            print("Table not found, retrying...")
            time.sleep(2)
            continue

        current_hash = hashlib.md5(str(table).encode("utf-8")).hexdigest()
        if previous_table_hash == current_hash:
            print("Table unchanged. No more pages.")
            break
        previous_table_hash = current_hash

        rows = table.find_all("tr")
        for row in rows:
            tds = row.find_all("td")
            if len(tds) >= 5:
                tracking_number = deduplicate_tracking_number(tds[2].get_text(strip=True))
                shipper_ref = tds[3].get_text(strip=True) or "N/A"
                consignee_ref = tds[4].get_text(strip=True)
                all_tracking_numbers.append(tracking_number)
                all_shipper_refs.append(shipper_ref)
                all_consignee_refs.append(consignee_ref)

        print(f"Collected {len(all_shipper_refs)} total rows so far...")

        try:
            next_button = driver.find_element(
                By.CSS_SELECTOR, 'a.page-link.ng-scope[ng-switch-when="next"][ng-click]'
            )
            driver.execute_script("arguments[0].click();", next_button)
            time.sleep(2)
        except:
            print("No more pages left.")
            break

    # Save history data
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents")
    history_path = os.path.join(documents_folder, "history.xlsx")
    df_history = pd.DataFrame({
        "Tracking Number": all_tracking_numbers,
        "Shipper Reference": all_shipper_refs,
        "Consignee Reference": all_consignee_refs
    })
    df_history = df_history.drop_duplicates(subset=["Shipper Reference", "Consignee Reference"], keep="first")
    df_history.to_excel(history_path, index=False)
    print(f"History data saved to: {history_path}")

    # --- NAVIGATE TO DETAILED REPORT ---
    report_dropdown = WebDriverWait(driver, 40).until(
        EC.element_to_be_clickable((By.ID, "reports"))
    )
    time.sleep(3)
    report_dropdown.click()

    detailed_option = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//option[text()="Detailed Report"]'))
    )
    time.sleep(2)
    detailed_option.click()
    time.sleep(4)

    # --- SCRAPE DETAILED REPORT + BOXCOUNTS ---
    boxcounts = {r: 0 for r in regions}

    all_detailed_tracking = []
    all_emails = []
    previous_table_hash = None

    while True:
        table_elem = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "table"))
        )
        time.sleep(1)
        html = driver.execute_script("return document.documentElement.outerHTML;")
        soup = BeautifulSoup(html, "html.parser")
        table = soup.find("table")
        if not table:
            print("Table not found, retrying...")
            time.sleep(2)
            continue

        current_hash = hashlib.md5(str(table).encode("utf-8")).hexdigest()
        if previous_table_hash == current_hash:
            print("Detailed report table unchanged. No more pages.")
            break
        previous_table_hash = current_hash

        headers = [h.get_text(strip=True) for h in table.find_all("th")]
        tracking_idx = next((i for i, h in enumerate(headers) if "Tracking" in h), 2)
        email_idx = next((i for i, h in enumerate(headers) if "Email" in h), -1)

        rows = table.find_all("tr")
        for row in rows:
            tds = row.find_all("td")
            if len(tds) > max(tracking_idx, email_idx):
                tracking_number = deduplicate_tracking_number(tds[tracking_idx].get_text(strip=True))
                email = tds[email_idx].get_text(strip=True)

                region = region_from_email(email, config)
                if region:
                    boxcounts[region] += 1

                all_detailed_tracking.append(tracking_number)
                all_emails.append(email)

        print(f"Collected {len(all_detailed_tracking)} detailed rows so far...")

        try:
            next_button = driver.find_element(
                By.CSS_SELECTOR, 'a.page-link.ng-scope[ng-switch-when="next"][ng-click]'
            )
            driver.execute_script("arguments[0].click();", next_button)
            time.sleep(3)
        except:
            print("No more pages in detailed report.")
            break

    df_detailed = pd.DataFrame({
        "Tracking Number": all_detailed_tracking,
        "Email": all_emails
    })

    df_final = pd.merge(
        df_history,
        df_detailed,
        on="Tracking Number",
        how="left"
    )

    return df_final, boxcounts


# --- SHARED WEEKEND SCRAPE ---

def run_weekend_scrape(day_offset, config):
    """Shared Saturday/Sunday flow: login, set calendar dates, scrape.

    Returns: (driver, df_final, boxcounts)
    """
    driver = initialize_driver_and_login()
    set_date_range(driver, day_offset)
    click_search_button(driver)
    print("Waiting for table to refresh with new date range...")
    time.sleep(5)
    increase_table_size_to_100(driver)
    print("Continuing scrape...")
    df_final, boxcounts = scrape_history_and_detailed(driver, config)
    return driver, df_final, boxcounts


# --- REPORT FUNCTIONS ---

def run_monday(config):
    print("\n[RUNNING MONDAY SCRIPT...]\n")
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents")

    driver = initialize_driver_and_login()
    increase_table_size_to_100(driver)

    df_final, boxcounts = scrape_history_and_detailed(driver, config)

    # Adjust boxcounts by subtracting Saturday and Sunday
    saturday_path = os.path.join(documents_folder, "history_with_email_and_summary_automated.xlsx")
    sunday_path = os.path.join(documents_folder, "history_with_email_and_summary_SUN.xlsx")

    boxcounts_adj = dict(boxcounts)

    for path in [saturday_path, sunday_path]:
        if os.path.exists(path):
            try:
                df_prev_summary = pd.read_excel(path, sheet_name="RegionSummary", dtype=object)
                prev_boxcounts = get_boxcounts_from_summary(df_prev_summary, config)
                for region in boxcounts_adj:
                    boxcounts_adj[region] -= prev_boxcounts.get(region, 0)
                print(f"[Boxcount] Subtracted from {os.path.basename(path)} -> {prev_boxcounts}")
            except Exception as e:
                print(f"Could not adjust box counts from {path}: {e}")

    print(f"[Boxcount] After subtraction -> {boxcounts_adj}")

    # Isolate Monday-only shipments
    prior_dfs = []

    if os.path.exists(saturday_path):
        print("Saturday history file found -- loading...")
        prior_dfs.append(pd.read_excel(saturday_path, sheet_name="History"))

    if os.path.exists(sunday_path):
        print("Sunday history file found -- loading...")
        prior_dfs.append(pd.read_excel(sunday_path, sheet_name="History"))

    if prior_dfs:
        print("Isolating Monday-only data using available prior files...")
        df_prior = pd.concat(prior_dfs, ignore_index=True)
        df_mon_history = df_final[
            ~df_final["Tracking Number"].isin(df_prior["Tracking Number"])
        ].copy()
        print(f"Isolated {len(df_mon_history)} Monday-only shipments.")
    else:
        print("No prior history files found -- using full scrape as Monday data.")
        df_mon_history = df_final.copy()

    df_mon_history, packed_counts, picked_counts = process_region_counts(df_mon_history, config)
    config = reconcile_initials_with_config(packed_counts, picked_counts, config)
    region_initials, review_entries = classify_initials(packed_counts, picked_counts, config)
    df_summary = build_region_summary(picked_counts, packed_counts, boxcounts_adj, region_initials, review_entries)

    output_path = os.path.join(documents_folder, "history_with_email_and_summary_MON.xlsx")
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df_mon_history.to_excel(writer, sheet_name="History", index=False)
        df_summary.to_excel(writer, sheet_name="RegionSummary", index=False)

    print(f"Monday-only data and summary saved to: {output_path}")
    driver.quit()


def run_tuesday_friday(config):
    print("\n[RUNNING TUESDAY-FRIDAY SCRIPT...]\n")
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents")

    driver = initialize_driver_and_login()
    increase_table_size_to_100(driver)
    print("Continuing scrape...")

    df_final, boxcounts = scrape_history_and_detailed(driver, config)
    df_final, packed_counts, picked_counts = process_region_counts(df_final, config)

    config = reconcile_initials_with_config(packed_counts, picked_counts, config)
    region_initials, review_entries = classify_initials(packed_counts, picked_counts, config)
    df_summary = build_region_summary(picked_counts, packed_counts, boxcounts, region_initials, review_entries)

    output_path = os.path.join(documents_folder, "history_with_email_and_summary_automated.xlsx")
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, sheet_name="History", index=False)
        df_summary.to_excel(writer, sheet_name="RegionSummary", index=False)

    print(f"Region summary saved to: {output_path}")
    driver.quit()


def run_saturday(config):
    print("\n[RUNNING SATURDAY SCRIPT...]\n")
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents")

    driver, df_final, boxcounts = run_weekend_scrape(2, config)

    df_final, packed_counts, picked_counts = process_region_counts(df_final, config)

    config = reconcile_initials_with_config(packed_counts, picked_counts, config)
    region_initials, review_entries = classify_initials(packed_counts, picked_counts, config)
    df_summary = build_region_summary(picked_counts, packed_counts, boxcounts, region_initials, review_entries)

    output_path = os.path.join(documents_folder, "history_with_email_and_summary_automated.xlsx")
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, sheet_name="History", index=False)
        df_summary.to_excel(writer, sheet_name="RegionSummary", index=False)

    print(f"Region summary saved to: {output_path}")
    driver.quit()


def run_sunday(config):
    print("\n[RUNNING SUNDAY SCRIPT...]\n")
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents")

    driver, df_final, boxcounts = run_weekend_scrape(1, config)

    # Save raw Sunday scrape
    sunday_raw_path = os.path.join(documents_folder, "history_with_email_and_summary_SUN_raw.xlsx")
    df_final.to_excel(sunday_raw_path, index=False)
    print(f"Raw Sunday scrape saved to: {sunday_raw_path}")

    # Isolate Sunday-only shipments by removing Saturday's tracking numbers
    saturday_path = os.path.join(documents_folder, "history_with_email_and_summary_automated.xlsx")
    if os.path.exists(saturday_path):
        df_hist_sat = pd.read_excel(saturday_path, sheet_name="History")
        df_sun_unique = df_final[~df_final["Tracking Number"].isin(df_hist_sat["Tracking Number"])].copy()
        print(f"Isolated {len(df_sun_unique)} Sunday-only shipments (new since Saturday).")
        df_final = df_sun_unique
    else:
        print("No Saturday file found -- using full Sunday dataset.")

    df_final, packed_counts, picked_counts = process_region_counts(df_final, config)

    # Subtract Saturday boxcounts
    if os.path.exists(saturday_path):
        try:
            df_sat_summary = pd.read_excel(saturday_path, sheet_name="RegionSummary", dtype=object)
            prev_boxcounts = get_boxcounts_from_summary(df_sat_summary, config)
            for region in boxcounts:
                boxcounts[region] = max(0, boxcounts[region] - prev_boxcounts.get(region, 0))
            print(f"[Boxcount] After subtraction -> {boxcounts}")
        except Exception as e:
            print(f"[Boxcount] Error reading Saturday RegionSummary: {e}")

    config = reconcile_initials_with_config(packed_counts, picked_counts, config)
    region_initials, review_entries = classify_initials(packed_counts, picked_counts, config)
    df_summary = build_region_summary(picked_counts, packed_counts, boxcounts, region_initials, review_entries)

    output_path = os.path.join(documents_folder, "history_with_email_and_summary_SUN.xlsx")
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, sheet_name="History", index=False)
        df_summary.to_excel(writer, sheet_name="RegionSummary", index=False)

    print(f"Sunday scrape saved to: {output_path}")
    driver.quit()
    print("Script completed successfully.")


# --- USER INTERFACE ---

def get_user_choice():
    while True:
        print("\nSelect the type of report to run:")
        print("  1 = Monday")
        print("  2 = Tuesday-Friday")
        print("  3 = Saturday")
        print("  4 = Sunday")
        print("  5 = Manage Regions")
        print("  6 = Exit")

        choice = input("Enter your choice (1-6): ").strip()

        if not choice.isdigit():
            print("Invalid input -- please enter a number from 1 to 6.")
            continue

        choice = int(choice)

        if choice not in (1, 2, 3, 4, 5, 6):
            print("Invalid choice -- please select 1, 2, 3, 4, 5, or 6.")
            continue

        return choice


# --- MAIN ---

def main():
    config = load_config()
    if config is None:
        config = setup_config()

    while True:
        choice = get_user_choice()

        if choice == 1:
            run_monday(config)
        elif choice == 2:
            run_tuesday_friday(config)
        elif choice == 3:
            run_saturday(config)
        elif choice == 4:
            run_sunday(config)
        elif choice == 5:
            config = manage_regions(config)
        elif choice == 6:
            print("\nExiting program.\n")
            sys.exit()


if __name__ == "__main__":
    main()
