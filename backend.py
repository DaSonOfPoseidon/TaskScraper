import os
import re
import time
import pickle
import traceback
from tqdm import tqdm
from rapidfuzz import fuzz, process
from datetime import datetime
from collections import Counter, defaultdict
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from dotenv import load_dotenv, set_key
from threading import Lock

cookie_lock = Lock()
TASK_URL = "http://inside.sockettelecom.com/menu.php?tabid=45&tasktype=2&nID=1439&width=1440&height=731"
PAGE_TIMEOUT = 30

LOG_FILE = os.path.join(os.path.join(os.path.dirname(__file__), "..", "Outputs"), "consultation_log.txt")
def normalize_string(s):
    return re.sub(r'[^a-z0-9 ]+', '', s.lower()).strip()

JOB_TYPE_CATEGORIES = {
    "Free": {
        normalize_string(x) for x in [
            "WiFi Survey", "NID/IW/CopperTest",
            "ONT Swap", "STB to ONN Conversion", "Jack/FXS/Phone Check", "Blank",
            "Go-Live", "ONT Swap", "Install"
        ]
    },
    "Billable": {
        normalize_string(x) for x in [
            "ONT Move"
        ]
    },
    "Unknown": set()
}


# === Login & Session ===
def prompt_for_credentials():
    from tkinter import Tk, simpledialog
    root = Tk()
    root.withdraw()
    user = simpledialog.askstring("Login", "Username:", parent=root)
    pw = simpledialog.askstring("Login", "Password:", parent=root, show="*")
    root.destroy()
    return user, pw

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
    time.sleep(2)

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

def save_cookies(driver, filename="cookies.pkl"):
    with cookie_lock:
        with open(filename, "wb") as f:
            pickle.dump(driver.get_cookies(), f)

def load_cookies(driver, filename="cookies.pkl"):
    if not os.path.exists(filename): return False
    try:
        driver.get("http://inside.sockettelecom.com/")
        with cookie_lock:
            cookies = pickle.load(open(filename, "rb"))
            for cookie in cookies:
                driver.add_cookie(cookie)
        return True
    except:
        return False

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

def is_free_job(job_type):
    if not job_type:
        return ("(blank)", True, 100)

    job_type_norm = normalize_string(job_type)
    candidates = JOB_TYPE_CATEGORIES["Free"]

    match, score, _ = process.extractOne(job_type_norm, candidates, scorer=fuzz.partial_ratio)

    log_message(f"🔍 Matching '{job_type}' → '{match}' (score: {score})")
    return (match, score > 90, score)

# === Consultation Task Extraction ===
def create_driver():
    opts = webdriver.ChromeOptions()
    opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.page_load_strategy = 'eager'
    return webdriver.Chrome(service=Service("chromedriver.exe"), options=opts)

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

def extract_consultation_tasks(driver):
    driver.get(TASK_URL)
    WebDriverWait(driver, 30).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "MainView")))
    log_message("Loading Tasks...", also_print=True)
    rows = WebDriverWait(driver, 30).until(
        EC.presence_of_all_elements_located((By.XPATH, '//tr[contains(@class,"taskElement")]'))
    )
    tasks = [parse_task_row(row) for row in rows]
    consultations = [t for t in tasks if t and "consultation" in t["desc"].lower()]
    log_message(f"✅ Found {len(consultations)} consultation tasks.", also_print=True)
    return consultations

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
        log_message("✅ Switched to MainView for task") #uncomment to enable

        textarea = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "Notes"))
        )
        raw_notes = textarea.get_attribute("value").strip()
        log_message("✅ Found Notes textarea")

        match = re.search(r"PROBLEM STATEMENT:\s*<b>(.*?)</b>", raw_notes, re.IGNORECASE)
        if match:
            log_message("✅ Found bolded problem statement")
            return match.group(1).strip()

        match = re.search(r"PROBLEM STATEMENT:\s*(.+)", raw_notes, re.IGNORECASE)
        if match:
            log_message(f"⚠️ Found plain problem statement: {raw_notes}")
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
    def log_step(step_desc):
        log_message(f"[{datetime.now().strftime('%H:%M:%S')}] WO {task_id} — {step_desc}")

    def try_complete():
        def debug_element(label, by, value):
            try:
                el = WebDriverWait(driver, 5).until(EC.presence_of_element_located((by, value)))
                log_step(f"✔️ Found {label}: {value}")
                return el
            except Exception as e:
                log_step(f"❌ Failed to locate {label} using {by}={value}: {e}")
                raise

        checkbox = debug_element("checkbox", By.ID, f"completedcheck{task_id}")
        notes_box = debug_element("notes box", By.ID, f"txtNotes{task_id}")
        submit_btn = debug_element("submit button", By.ID, f"sub_{task_id}")

        if not checkbox.is_selected():
            log_step("Clicking 'Completed' checkbox...")
            driver.execute_script("arguments[0].click();", checkbox)
        else:
            log_step("Checkbox already selected.")

        log_step("Entering notes...")
        notes_box.clear()
        notes_box.send_keys(f"{job_type}, no charge")

        log_step("Clicking 'Update Task' button...")
        driver.execute_script("arguments[0].click();", submit_btn)

    for attempt in range(2):
        try:
            log_step(f"--- Attempt {attempt + 1} to complete task ---")
            try_complete()
            log_step(f"✅ Successfully completed as free ({job_type})")
            return True
        except Exception as e:
            log_step(f"⚠️ Attempt {attempt + 1} failed: {type(e).__name__} - {e}")
            if attempt == 0:
                time.sleep(1)
                continue
            if screenshot_dir:
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                fname = f"wo_{task_id}_fail_{ts}.png"
                fpath = os.path.join(screenshot_dir, fname)
                driver.save_screenshot(fpath)
                log_step(f"📸 Screenshot saved to: {fpath}")
            log_step("❌ Gave up after retrying")
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
            time.sleep(0.3)
    except Exception as e:
        log_message(f"❌ expand_task(): Failed to expand Task ID {task_id}: {e}")

def dump_debug_html(driver, form_id, task_id):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = f"debug_form_{form_id}_{task_id}_{ts}.html"
    path = os.path.join("logs", fname)
    with open(path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    log_message(f"📄 HTML snapshot saved to: {fname}")

def run_with_progress(driver, tasks, complete_free=False):
    results = []
    errors = []

    for task in tqdm(tasks, desc="Parsing tasks", unit="task"):
        try:
            job_type = parse_job_type_from_task(driver, task["url"])
            results.append({
                "Company": task["company"],
                "Description": task["desc"],
                "URL": task["url"],
                "Job Type": job_type
            })

            norm_type = job_type.lower().strip()
            match, is_match, score = is_free_job(norm_type)
            if complete_free and is_match:
                log_message(f"[MATCH] '{job_type}' matched as free ({match}) with score {score}")
                task_id = extract_task_id_from_page(driver)
                if task_id:
                    log_message(f"🔍 Task ID {task_id} — preparing to expand and complete...")
                    expand_task(driver, task_id)
                    log_message(f"✅ Expanded task {task_id}")

                    form_id = f"TOSSTask{task_id}"
                    try:
                        WebDriverWait(driver, 12).until(
                            lambda d: d.find_element(By.ID, form_id).is_displayed()
                        )
                        log_message(f"✅ Form {form_id} is now visible")
                    except Exception as e:
                        log_message(f"❌ Form {form_id} did not become visible in time: {type(e).__name__}")
                        dump_debug_html(driver, form_id, task_id, log_message)
                        errors.append({"Task": task, "Error": f"Form {form_id} not visible", "Traceback": str(e)})
                        continue


                    except TimeoutException:
                        log_message(f"❌ Timeout waiting for form displayForm{task_id}")
                        continue

                    success = complete_free_task(driver, task_id, job_type, log_message)
                    if success:
                        log_message(f"✔️  Completed Task {task_id} — {job_type}")
                    else:
                        log_message(f"⚠️  Failed to complete Task {task_id} — {job_type}")

            else:
                log_message(f"[NO MATCH] '{job_type}' not identified as a free job")


        except Exception as e:
            tb = traceback.format_exc()
            errors.append({"Task": task, "Error": str(e), "Traceback": tb})
            log_message(f"❌ Error for {task['desc']} — {e}")
            log_message(tb)
            continue

    return results, errors


if __name__ == "__main__":
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("")
    driver = create_driver()
    handle_login(driver)
    clear_first_time_overlays(driver)
    
    tasks = extract_consultation_tasks(driver)
    results, errors = run_with_progress(driver, tasks, complete_free=True)
    log_message(f"\n✅ Done. Parsed {len(results)} tasks with {len(errors)} errors.", True)
    summarize_job_types(results)
    driver.quit()


