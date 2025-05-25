import os
import re
import time
import sys
import signal
import json
import pickle
import traceback
import getpass
from tqdm import tqdm
from rapidfuzz import fuzz, process
from datetime import datetime, date
from collections import Counter, defaultdict
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from dotenv import load_dotenv, set_key

TASK_URL = "http://inside.sockettelecom.com/menu.php?tabid=45&tasktype=2&nID=1439&width=1440&height=731"
PAGE_TIMEOUT = 30

DRY_RUN = False

COOKIE_FIELDS = (
    "name", "value", "domain", "path", "secure", "httpOnly", "expiry",
)

LOG_FILE = os.path.join(os.path.join(os.path.dirname(__file__), "..", "Outputs"), "consultation_log.txt")
def normalize_string(s):
    return re.sub(r'[^a-z0-9 ]+', '', s.lower()).strip()

JOB_TYPE_CATEGORIES = {
    "Free": {
        normalize_string(x) for x in [
            "WiFi Survey", "NID/IW/CopperTest", "equipment check", "swap router",
            "ONT Swap", "STB to ONN Conversion", "Jack/FXS/Phone Check", "Blank",
            "Go-Live", "Install", "rouge ont", "onn swap", "ont dying", "stb swap",
            "Tie down", "onn"
        ]
    },
    "Billable": {
        normalize_string(x) for x in [
            "ONT Move", "ONT in Disco", "Fiber Cut", "Broken Fiber"
        ]
    },
    "Unknown": set()
}

# === Login & Session ===
def prompt_for_credentials():
    username = input("Username: ")
    password = getpass.getpass("Password: ")
    return username, password

def save_env_credentials(user, pw):
    path = ".env"
    if not os.path.exists(path): open(path, "w").close()
    set_key(path, "UNITY_USER", user)
    set_key(path, "PASSWORD", pw)

def check_env_or_prompt_login():
    load_dotenv()
    user = os.getenv("UNITY_USER")
    pw = os.getenv("PASSWORD")
    if user and pw:
        log_message("✅ Loaded credentials from .env")
        return user, pw
    user, pw = prompt_for_credentials()
    save_env_credentials(user, pw)
    return user, pw

def perform_login(driver, user, pw):
    driver.get("http://inside.sockettelecom.com/system/login.php")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "username")))
    driver.find_element(By.NAME, "username").send_keys(user)
    driver.find_element(By.NAME, "password").send_keys(pw)
    driver.find_element(By.ID, "login").click()
    WebDriverWait(driver, 2)

def handle_login(driver):
    driver.get("http://inside.sockettelecom.com/")
    if load_cookies(driver):
        driver.refresh()
        if "login.php" not in driver.current_url:
            log_message("✅ Session restored with cookies")
            clear_first_time_overlays(driver)
            return
    user, pw = check_env_or_prompt_login()
    perform_login(driver, user, pw)
    clear_first_time_overlays(driver)
    save_cookies(driver)
    log_message("✅ Logged in via credentials")

def save_cookies(driver, filepath="cookies.pkl"):
    raw = driver.get_cookies()
    filtered = [{k: c[k] for k in COOKIE_FIELDS if k in c} for c in raw]
    with open(filepath, "w") as f:
        json.dump(filtered, f, indent=2)
    print(f"Saved {len(filtered)} cookies (JSON) → {filepath}")

def load_cookies(driver, filepath="cookies.pkl") -> bool:
    if not os.path.exists(filepath):
        print(f"No cookie file at {filepath}")
        return False

    # 1) Attempt JSON load, else fall back to pickle
    try:
        with open(filepath, "r") as f:
            cookies = json.load(f)
    except (json.JSONDecodeError, UnicodeDecodeError):
        with open(filepath, "rb") as f:
            cookies = pickle.load(f)

    now = int(time.time())
    added = 0

    for c in cookies:
        exp = c.get("expiry")
        if exp and exp < now:
            # skip stale cookies
            continue
        # ensure only safe fields (in case pickle payload has extras)
        safe_cookie = {k: c[k] for k in COOKIE_FIELDS if k in c}
        try:
            driver.add_cookie(safe_cookie)
            added += 1
        except Exception as e:
            print(f"⚠️ Skipped cookie {c.get('name')}: {e}")

    print(f"Loaded {added} cookies ← {filepath}")
    return added > 0

def clear_first_time_overlays(driver):
    # Dismiss alert if present
    try:
        WebDriverWait(driver, 0.5).until(EC.alert_is_present())
        driver.switch_to.alert.dismiss()
    except:
        pass

    # Known popup buttons
    buttons = [
        "//form[@id='valueForm']//input[@type='button']",
        "//form[@id='f']//input[@type='button']"
    ]
    for xpath in buttons:
        try:
            WebDriverWait(driver, 0.5).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
        except:
            pass

def log_message(msg, also_print=False):
    timestamp = datetime.now().strftime("[%H:%M:%S]")
    full_msg = f"{timestamp} {msg}"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(full_msg + "\n")
    if also_print:
        print(full_msg)

def format_dispatch_summary(driver, job_type):
    ci = get_customer_and_ticket_info_from_task(driver)
    if not ci or not ci["ticket"]:
        return None

    driver.get(ci["customer_url"])
    wo_url, wo_number = get_dispatch_work_order_url(driver, ci["ticket"])
    if not wo_url:
        return None

    driver.get(wo_url)
    # wait for the form to load
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "AdditionalNotes")))

    status_el = driver.find_element(
        By.XPATH,
        "//td[@class='detailHeader' and normalize-space(text())='Status:']"
        "/following-sibling::td//span"
    )
    status = status_el.text.strip()

    if status.lower() != "completed":
        log_message(f"⚠️ WO {wo_number} is still uncompleted; skipping")
        return


    # pull raw strings
    arr_date = driver.find_element(By.ID, "ArrivalOnsite").get_attribute("value").strip()
    arr_time = driver.find_element(By.ID, "ArrivalTime").get_attribute("value").strip()
    dep_date = driver.find_element(By.ID, "CompletedDate").get_attribute("value").strip()
    dep_time = driver.find_element(By.ID, "CompletedTime").get_attribute("value").strip()

    # display values (blank → “not given”)
    if arr_date and re.match(r"\d{4}-\d{2}-\d{2}", arr_date) and arr_time and re.match(r"\d{1,2}:\d{2}", arr_time):
        arr_display = arr_time
    else:
        arr_display = "not given"

    if dep_date and re.match(r"\d{4}-\d{2}-\d{2}", dep_date) and dep_time and re.match(r"\d{1,2}:\d{2}", dep_time):
        dep_display = dep_time
    else:
        dep_display = "not given"

    # compute total_hours only if both arrival and departure are valid
    if arr_display != "not given" and dep_display != "not given":
        t_arr = datetime.strptime(f"{arr_date} {arr_time}", "%Y-%m-%d %H:%M")
        t_dep = datetime.strptime(f"{dep_date} {dep_time}", "%Y-%m-%d %H:%M")
        hours = (t_dep - t_arr).total_seconds() / 3600
        # enforce 1h min & quarter‐hour rounding
        hours = max(hours, 1.0)
        total_hours = round(hours * 4) / 4
        total_display = f"{total_hours:.2f}"
    else:
        total_display = "1.00"

    notes = extract_work_order_notes(driver)
    # WORK DONE — prefer the “AdditionalNotes” field if present
    work_done = notes["fields"].get("AdditionalNotes", notes["combined"]).strip()

    # EQUIPMENT USED — from the EquipmentInstalled textarea, split lines
    equipment_raw = notes["fields"].get("EquipmentInstalled", "")
    equipment = [line.strip() for line in equipment_raw.splitlines() if line.strip()]

    # RESPONSIBLE PARTY — look for “damage caused by X” or “X responsible,” else default
    responsible = "Customer"
    text = notes["combined"].lower() 
    # look for “damage caused by ...”
    m = re.search(r"damage caused by\s+([^.,\n]+)", text, re.IGNORECASE)
    if m:
        responsible = m.group(1).strip()
    else:
        # look for “... responsible”
        m2 = re.search(r"([^.,\n]+?)\s+responsible", text, re.IGNORECASE)
        if m2:
            responsible = m2.group(1).strip()

        if "brightspeed" in text or "bright speed" in text:
            responsible = "Brightspeed"

    summary = [
        f"CUSTOMER: {ci['customer_name']}",
        f"CID: {ci['cid']}",
        f"WORK ORDER NUMBER: {wo_number}",
        f"WORK ORDER LINK: {wo_url}",
        "",
        f"Arrival Time: {arr_display}",
        f"Departure Time: {dep_display}",
        "",
        f"Total time of DP: {total_display}",
        "",
        "WORK DONE (DISPATCH NOTES):",
        work_done,
        "",
        "EQUIPMENT USED:",
    ] + equipment + [
        "",
        f"RESPONSIBLE PARTY: {responsible}"
    ]
    summary_text = "\n".join(summary)
    log_message(summary_text)
    return summary_text

def is_free_job(job_type):
    if not job_type:
        return ("(blank)", True, 100)

    job_type_norm = normalize_string(job_type)
    candidates = JOB_TYPE_CATEGORIES["Free"]

    match, score, _ = process.extractOne(job_type_norm, candidates, scorer=fuzz.partial_ratio)

    log_message(f"🔍 Matching '{job_type}' → '{match}' (score: {score})")
    return (match, score > 90, score)

def is_billable_job(job_type):
    if not job_type:
        return ("(blank)", False, 0)

    job_type_norm = normalize_string(job_type)
    candidates = JOB_TYPE_CATEGORIES["Billable"]

    match, score, _ = process.extractOne(job_type_norm, candidates, scorer=fuzz.partial_ratio)

    log_message(f"💰 Matching '{job_type}' → '{match}' (score: {score})")
    return (match, score > 90, score)

def update_notes_only(driver, task_id, summary_text):
    try:
        # switch into the MainView frame
        driver.switch_to.default_content()
        WebDriverWait(driver, 5).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "MainView"))
        )
        expand_task(driver, task_id)

        # update the notes field
        notes = driver.find_element(By.ID, f"txtNotes{task_id}")
        notes.clear()
        notes.send_keys(summary_text)

        # click the Update button
        btn = driver.find_element(By.ID, f"sub_{task_id}")
        driver.execute_script("arguments[0].click()", btn)

        log_message(f"✏️  (DRY RUN) Updated notes for task {task_id}")
        return True

    except Exception as e:
        log_message(f"❌ update_notes_only failed: {type(e).__name__} - {e}")
        return False

# === Consultation Task Extraction ===
def create_driver():
    opts = webdriver.ChromeOptions()
    opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-usb-keyboard-detect")
    opts.add_argument("--disable-hid-detection")
    opts.add_argument("--log-level=3")
    opts.page_load_strategy = 'eager'
    return webdriver.Chrome(options=opts)

def parse_task_row(row):
    try:
        tds = row.find_elements(By.TAG_NAME, "td")
        if len(tds) < 6: return None
        return {
            "url": tds[0].find_element(By.TAG_NAME, "a").get_attribute("href"),
            "desc": tds[1].text.strip(),
            "assigned": tds[4].text.strip(),
            "company": tds[5].text.strip(),
        }
    except:
        return None

def get_customer_and_ticket_info_from_task(driver):
    try:
        try:
            driver.switch_to.default_content()
            WebDriverWait(driver, 5).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "MainView"))
            )
        except:
            log_message("⚠️ Already in MainView or frame not needed.")

        try:
            cid = driver.find_element(
                By.XPATH,
                "//td[normalize-space(text())='Customer ID']/following-sibling::td/b"
            ).text.strip()
            customer_name = driver.find_element(
                By.XPATH,
                "//td[normalize-space(text())='Customer Name']/following-sibling::td/b"
            ).text.strip()
        except Exception as e:
            log_message(f"❌ Failed to extract Customer ID/Name: {e}")
            return None

        try:
            desc_cell = driver.find_element(
                By.XPATH,
                "//td[contains(., 'Dispatch for Ticket')]"
            )
            m = re.search(r"Dispatch for Ticket\s+(\d+)", desc_cell.text)
            ticket_number = m.group(1) if m else None
            log_message(f"✅ Found Ticket #: {ticket_number}")
        except:
            ticket_number = None
            log_message("⚠️ Could not find Ticket # in dispatch description")
        customer_url = (
            "http://inside.sockettelecom.com/menu.php"
            "?coid=1&tabid=7&parentid=9&customerid=" + cid
        )
        return {
            "customer_name": customer_name,
            "cid":            cid,
            "ticket":         ticket_number,
            "customer_url":   customer_url
        }
    except Exception as e:
        log_message(f"❌ Failed to get customer/ticket info from task: {e}")
        return None

def get_dispatch_work_order_url(driver, ticket_number, log=None):
    try:
        driver.switch_to.default_content()
        driver.switch_to.frame("MainView")
    except Exception as e:
        log_message(f"❌ Could not switch to MainView iframe: {e}")
    try:
        clear_first_time_overlays(driver)
        WebDriverWait(driver, 1.5, poll_frequency=0.05).until(
            lambda d: d.find_element(By.ID, "workShow").is_displayed()
        )

        rows = driver.find_elements(By.XPATH, "//div[@id='workShow']//table//tr[position()>1]")

        dispatch_wos = []
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 5:
                continue
            wo_num = cols[0].text.strip()
            desc = cols[1].text.strip().lower()
            url = cols[4].find_element(By.TAG_NAME, "a").get_attribute("href")

            if re.search(rf"ticket\s*#?\s*{ticket_number}", desc, re.IGNORECASE):
                dispatch_wos.append((int(wo_num), url))



        if not dispatch_wos:
            log_message(f"⚠️ No dispatch WOs found for Ticket #{ticket_number}")
            return None, None

        # Return the WO with the highest number (most recent)
        wo_url, wo_number = max(dispatch_wos, key=lambda x: x[0])[1], max(dispatch_wos, key=lambda x: x[0])[0]
        return wo_url, wo_number

    except Exception as e:
        if log:
            log(f"❌ Error finding dispatch WO for ticket {ticket_number}: {e}")
        else:
            print(f"❌ Error finding dispatch WO for ticket {ticket_number}: {e}")
        return None, None

def extract_work_order_notes(driver):
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "AdditionalNotes")))
        #log_message("✅ WO page loaded, extracting notes...")

        fields = {
            "EquipmentInstalled": "",
            "AdditionalMaterials": "",
            "TestsPerformed": "",
            "AdditionalNotes": ""
        }

        for field_id in fields:
            try:
                textarea = driver.find_element(By.ID, field_id)
                fields[field_id] = textarea.get_attribute("value").strip()
                if fields[field_id]:
                    log_message(f"📄 {field_id} → {len(fields[field_id])} chars")
            except Exception as e:
                log_message(f"⚠️ Could not read {field_id}: {e}")

        # Combine all note fields into a single block
        combined_notes = "\n".join(
            f"{label.replace('Additional', 'Additional ').replace('Performed', 'Performed:')}: {text}"
            for label, text in fields.items() if text
        )

        return {
            "fields": fields,
            "combined": combined_notes.strip()
        }

    except Exception as e:
        log_message(f"❌ Failed to extract WO notes: {e}")
        return {
            "fields": {},
            "combined": ""
        }

def extract_due_consultation_tasks(driver):
    driver.get(TASK_URL)
    # wait for the MainView frame to load and switch into it
    WebDriverWait(driver, 30).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, "MainView"))
    ) 
    log_message("Loading Tasks…", also_print=True)

    # grab all task rows
    rows = WebDriverWait(driver, 30).until(
        EC.presence_of_all_elements_located(
            (By.XPATH, '//tr[contains(@class,"taskElement")]')
        )
    ) 

    today = date.today()
    due_consults = []

    for row in rows:
        task = parse_task_row(row)
        if not task or "consultation" not in task["desc"].lower():
            continue

        # parse the due-date from the 4th <td> nobr
        try:
            due_str = row.find_element(
                By.CSS_SELECTOR, "td:nth-child(4) nobr"
            ).text.strip()
            due_dt = datetime.strptime(due_str, "%Y-%m-%d").date()
        except Exception:
            log_message(f"⚠️ Couldn't parse due date '{due_str}'; skipping", also_print=True)
            continue 

        # skip tasks not yet due
        if due_dt > today:
            log_message(f"⏳ Skipping '{task['desc']}' (due {due_dt.isoformat()})", also_print=True)
            continue

        # this is a consultation task due today or overdue
        due_consults.append(task)

    log_message(f"✅ Found {len(due_consults)} due consultation tasks.", also_print=True)
    return due_consults

def extract_task_id_from_page(driver):
    try:
        # We're already in MainView frame
        task_id_input = driver.find_element(By.NAME, "nTaskID")
        return task_id_input.get_attribute("value")
    except:
        return None

def parse_job_type_from_task(driver, url):
    try:
        log_message(f"\n🔎 Opening task URL: {url}")
        driver.get(url)
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "MainView"))
        )
        #log_message("✅ Switched to MainView for task") #uncomment to enable

        textarea = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "Notes"))
        )
        raw_notes = textarea.get_attribute("value").strip()
        #log_message("✅ Found Notes textarea")

        match = re.search(r"PROBLEM STATEMENT:\s*<b>(.*?)</b>", raw_notes, re.IGNORECASE)
        if match:
            #log_message("✅ Found bolded problem statement")
            return match.group(1).strip()

        match = re.search(r"PROBLEM STATEMENT:\s*(.+)", raw_notes, re.IGNORECASE)
        if match:
            log_message(f"⚠️ Found plain problem statement")
            line = match.group(1).strip()
            line = re.sub(r"</?[^>]+>", "", line)
            return line[:100].strip()

        for line in raw_notes.splitlines()[:15]:
            if "ont" in line.lower() and 3 < len(line.strip()) < 100:
                log_message("⚠️ Using fallback ONT line")
                return line.strip()
            
        if raw_notes == "":
            log_message("⚠️ Blank WO Notes")
            return "Blank"

        log_message("❌ Could not identify job type — returning 'Unknown'")
        log_message(f"WO Notes: {raw_notes}")
        return "Unknown"
    except Exception as e:
        log_message(f"❌ Failed to parse job type from {url}: {e}")
        raise

def summarize_job_types(results):
    major_types = [
        "ONT In Disco", "ONT Move", "ONT Swap", "WiFi Survey",
        "Go-Live", "NID/IW/CopperTest", "IW Tie Down", "Onn Install", 
        "Equipment Check/ONT Swap"
    ]
    job_counter = Counter()
    other_types = defaultdict(list)

    for task in results:
        job_type = task["Job Type"].strip()
        normalized = job_type.lower()

        matched = False
        for major in major_types:
            if major.lower() in normalized:
                job_counter[major] += 1
                matched = True
                break

        if not matched:
            if normalized == "blank":
                job_counter["Blank"] += 1
            elif normalized in ["unknown", "error"]:
                job_counter["Unknown"] += 1
            elif job_type == "":
                job_counter["Blank"] += 1
            else:
                other_types[job_type].append(task)
                job_counter["Other"] += 1



    # Print summary
    log_message("\n📊 Job Type Summary:", True)
    for jt in major_types:
        if jt in job_counter:
            log_message(f"  {jt}: {job_counter[jt]}", True)
    if "Blank" in job_counter:
        log_message(f"  Blank: {job_counter['Blank']}", True)
    if "Unknown" in job_counter:
        log_message(f"  Unknown: {job_counter['Unknown']}", True)
    if "Other" in job_counter:
        log_message(f"  Other: {job_counter['Other']} (see below)", True)
        for other in other_types:
            log_message(f"    • {other} — {len(other_types[other])} task(s)", True)

    return job_counter, other_types

def complete_free_task(driver, task_id, job_type, screenshot_dir=None):
    def try_complete():
        def debug_element(label, by, value):
            try:
                el = WebDriverWait(driver, 5).until(EC.presence_of_element_located((by, value)))
                log_message(f"✔️ Found {label}: {value}")
                return el
            except Exception as e:
                log_message(f"❌ Failed to locate {label} using {by}={value}: {e}")
                raise

        checkbox = debug_element("checkbox", By.ID, f"completedcheck{task_id}")
        notes_box = debug_element("notes box", By.ID, f"txtNotes{task_id}")
        submit_btn = debug_element("submit button", By.ID, f"sub_{task_id}")

        if not checkbox.is_selected():
            log_message("Clicking 'Completed' checkbox...")
            driver.execute_script("arguments[0].click();", checkbox)
        else:
            log_message("Checkbox already selected.")

        log_message("Entering notes...")
        notes_box.clear()
        notes_box.send_keys(f"{job_type}, no charge")

        log_message("Clicking 'Update Task' button...")
        driver.execute_script("arguments[0].click();", submit_btn)

    for attempt in range(2):
        try:
            log_message(f"--- Attempt {attempt + 1} to complete task ---")
            try_complete()
            log_message(f"✅ Successfully completed as free ({job_type})")
            return True
        except Exception as e:
            log_message(f"⚠️ Attempt {attempt + 1} failed: {type(e).__name__} - {e}")
            if attempt == 0:
                WebDriverWait(driver, 1)
                continue
            if screenshot_dir:
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                fname = f"wo_{task_id}_fail_{ts}.png"
                fpath = os.path.join(screenshot_dir, fname)
                driver.save_screenshot(fpath)
                log_message(f"📸 Screenshot saved to: {fpath}")
            log_message("❌ Gave up after retrying")
            return False

def complete_charged_task(driver, task_id, summary_text, screenshot_dir=None):
    def click_if_needed(el, label):
        if not el.is_selected():
            log_message(f"✔️ Clicking {label}")
            driver.execute_script("arguments[0].click()", el)

    try:
        # wait for the form
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.NAME, "SpawnBillingTask"))
        )

        # your existing “free job” completion first…
        checkbox = driver.find_element(By.ID, f"completedcheck{task_id}")
        click_if_needed(checkbox, "Completed")

        # now tick spawn‐billing
        spawn = driver.find_element(By.NAME, "SpawnBillingTask")
        click_if_needed(spawn, "SpawnBillingTask")

        # overwrite notes
        notes = driver.find_element(By.ID, f"txtNotes{task_id}")
        notes.clear()
        notes.send_keys(summary_text)

        # submit
        btn = driver.find_element(By.ID, f"sub_{task_id}")
        log_message("✔️ Clicking Update Task")
        driver.execute_script("arguments[0].click()", btn)

        log_message("✅ Charged task completed")
        return True

    except Exception as e:
        log_message(f"❌ complete_charged_task failed: {e}")
        if screenshot_dir:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = os.path.join(screenshot_dir, f"fail_{task_id}_{ts}.png")
            driver.save_screenshot(path)
            log_message(f"📸 Screenshot: {path}")
        return False

def expand_task(driver, task_id):
    try:
        # Locate the span wrapping the form
        span = driver.find_element(By.ID, f"displaySpan{task_id}")
        # From there, get the parent fieldset (2 levels up: span → td → fieldset)
        fieldset = span.find_element(By.XPATH, "./ancestor::fieldset[1]")
        legend = fieldset.find_element(By.TAG_NAME, "legend")

        # If the span (the content) is not visible, we assume it needs expansion
        if not span.is_displayed() or span.size["height"] < 5:
            legend.click()
            WebDriverWait(driver, 5, poll_frequency=0.1).until(
                EC.visibility_of_element_located((By.ID, f"displaySpan{task_id}"))
            )
    except Exception as e:
        log_message(f"❌ expand_task(): Failed to expand Task ID {task_id}: {e}")

def dump_debug_html(driver, form_id, task_id):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = f"debug_form_{form_id}_{task_id}_{ts}.html"
    path = os.path.join("logs", fname)
    with open(path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    log_message(f"📄 HTML snapshot saved to: {fname}")

def handle_sigterm(signum, frame):
    print("Received SIGTERM, exiting gracefully.")
    sys.exit(0)

def run_with_progress(driver, complete_free=False):
    results = []
    errors = []
    summaries = []

    due_tasks = extract_due_consultation_tasks(driver)

    for task in tqdm(due_tasks, desc="Processing consultation tasks", unit="task"):
        try:
            job_type = parse_job_type_from_task(driver, task["url"])
            results.append({
                "Company": task["company"],
                "Description": task["desc"],
                "URL": task["url"],
                "Job Type": job_type
            })

            norm_type = job_type.lower().strip()
            match, is_free, score = is_free_job(norm_type)
            if is_free:
                log_message(f"[MATCH] '{job_type}' matched as free ({match}) with score {score}")
                task_id = extract_task_id_from_page(driver)
                if task_id:
                    log_message(f"🔍 Task ID {task_id} — preparing to expand and complete...")
                    expand_task(driver, task_id)
                    log_message(f"✅ Expanded task {task_id}")

                    note_text = f"{job_type}, no charge"
                    if DRY_RUN or not complete_free:
                        success = update_notes_only(driver, task_id, note_text)
                        if success:
                            log_message(f"✏️  (DRY RUN) Updated notes for free Task {task_id}")
                        else:
                            log_message(f"⚠️  Failed DRY RUN update for free Task {task_id}")
                    else:
                        success = complete_free_task(driver, task_id, job_type, screenshot_dir=None)
                        if success:
                            log_message(f"✔️  Completed free Task {task_id} — {job_type}")
                        else:
                            log_message(f"⚠️  Failed to complete free Task {task_id} — {job_type}")

                    continue

            match, is_billable, score = is_billable_job(norm_type)
            if is_billable:
                log_message(f"\n💰 '{job_type}' matched as billable ({match}) with score {score}")
                task_id = extract_task_id_from_page(driver)
                if not task_id:
                    continue

                # scrape & format everything
                info = format_dispatch_summary(driver, job_type)
                if info:
                    summaries.append(info)
                else:
                    log_message(f"❌ Could not format dispatch summary for task {task_id}, skipping.")
                    continue


                driver.get(task["url"])
                WebDriverWait(driver, 5).until(
                    EC.frame_to_be_available_and_switch_to_it((By.ID, "MainView"))
                )
                expand_task(driver, task_id)

                if DRY_RUN:
                    if update_notes_only(driver, task_id, info):
                        log_message(f"✔️  Charged Task {task_id} updated successfully")
                    else:
                        log_message(f"⚠️  Failed to update Charged Task {task_id}")
                else:
                    if complete_charged_task(driver, task_id, info):
                        log_message(f"✔️  Charged Task {task_id} closed successfully")
                    else:
                        log_message(f"⚠️  Failed to close Charged Task {task_id}")

        except Exception as e:
            tb = traceback.format_exc()
            errors.append({"Task": task, "Error": str(e), "Traceback": tb})
            log_message(f"❌ Error for {task['desc']} — {e}")
            log_message(tb)
            continue

    return results, errors

if __name__ == "__main__":
    signal.signal(signal.SIGTERM, handle_sigterm)
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("")
    
    try:
        driver = create_driver()
        handle_login(driver)
        clear_first_time_overlays(driver)
        results, errors = run_with_progress(driver, complete_free=True)
        log_message(f"\n✅ Done. Parsed {len(results)} tasks with {len(errors)} errors.", True)
        summarize_job_types(results)

    except KeyboardInterrupt:
        print("Keyboard Interupt caught")
        sys.exit(0)

    except Exception as e:
        print("Unexpected error, aborting: %s", e)
        sys.exit(1)

    finally:
        try:
            driver.quit()
        except Exception:
            pass
    




