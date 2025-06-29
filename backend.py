import os
import re
from time import perf_counter
import sys
import signal
import traceback
import getpass
import subprocess
from pathlib import Path
from urllib.parse import urljoin
from tqdm import tqdm
from rapidfuzz import fuzz, process
from datetime import datetime, date
from collections import Counter, defaultdict
from dotenv import load_dotenv, set_key
from tkinter import messagebox, Tk
from playwright.sync_api import sync_playwright, Page, TimeoutError as PlaywrightTimeout, Error as PlaywrightError
import argparse

def get_project_root() -> str: #Returns the root directory of the project as a string path.
    # return string path for PROJECT_ROOT
    if getattr(sys, "frozen", False):
        exe_path = Path(sys.executable).resolve()
        parent = exe_path.parent
        if parent.name.lower() == "bin":
            root = parent.parent
        else:
            root = parent
    else:
        file_path = Path(__file__).resolve()
        parent = file_path.parent
        if parent.name.lower() == "bin":
            root = parent.parent
        else:
            root = parent.parent  # adjust as needed
    return str(root)

# Then:
PROJECT_ROOT = get_project_root()
OUTPUT_DIR  = os.path.join(PROJECT_ROOT, "Outputs")
MISC_DIR = os.path.join(PROJECT_ROOT, "Misc")
ENV_PATH    = os.path.join(MISC_DIR, ".env")
BROWSERS    = os.path.join(PROJECT_ROOT, "browsers")
LOG_FOLDER  = os.path.join(PROJECT_ROOT, "logs")
LOG_FILE    = os.path.join(LOG_FOLDER, "consulation_log.txt")
STATE_PATH = os.path.join(MISC_DIR, "state.json")


UPDATE_MODE = None

# ensure folders exist
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)
os.makedirs(BROWSERS, exist_ok=True)

# === CONFIGURATION ===
load_dotenv(dotenv_path=ENV_PATH)

TASK_URL = "http://inside.sockettelecom.com/menu.php?tabid=45&tasktype=2&nID=1439&width=1440&height=731"
PAGE_TIMEOUT = 15

DRY_RUN = False

__version__ = "0.1.2"

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
            "ONT Move", "ONT in Disco", "Fiber Cut", "Broken Fiber", "Fiber Move"
        ]
    },
    "Unknown": set()
}


class PlaywrightDriver:
    def __init__(self,
                 headless: bool = True,
                 playwright=None,
                 browser=None,
                 state_path: str = STATE_PATH):
        # If the caller passed us a playwright/browser, use those
        if playwright and browser:
            self._pw = playwright
            self.browser = browser
        else:
            # otherwise start our own
            self._pw = sync_playwright().start()
            self.browser = self._pw.chromium.launch(headless=headless)

        # load or create context
        if Path(state_path).exists():
            self.context = self.browser.new_context(storage_state=state_path)
        else:
            self.context = self.browser.new_context()

        self.page = self.context.new_page()
        self.page.on("dialog", lambda dlg: dlg.dismiss())
        self.page.route("**/*.{png,svg}", lambda route: route.abort())

    def goto(self, url: str, *, timeout: int = 5_000, wait_until: str = "load"):

        try:
            return self.page.goto(url, timeout=timeout, wait_until=wait_until)
        except PlaywrightTimeout:
            # fallback to load event if even DOMContentLoaded hung
            return self.page.goto(url, timeout=timeout, wait_until="load")
        

    def save_state(self):
        self.context.storage_state(path=STATE_PATH)

    def __getattr__(self, name):
        return getattr(self.page, name)

    def close(self):
        self.context.close()
        # only close the browser/playwright if *we* started it
        try:
            self.browser.close()
            self._pw.stop()
        except Exception:
            pass

def timed_goto(driver, url, **kwargs):
    start = perf_counter()
    driver.goto(url, **kwargs)
    elapsed = perf_counter() - start
    log_message(f"🕒 Navigated to {url!r} in {elapsed:.2f}s")

# === Login & Session ===
def prompt_for_credentials():
    username = input("Username: ")
    password = getpass.getpass("Password: ")
    return username, password

def save_env_credentials(user, pw):
    path = ENV_PATH
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

def handle_login(driver):
    # 1) Try to restore state
    timed_goto(driver, "http://inside.sockettelecom.com/")
    if "login.php" not in driver.page.url:
        log_message("✅ Session restored with stored state")
        clear_first_time_overlays(driver.page)
        return

    # 2) Otherwise, fall back to manual login
    user, pw = check_env_or_prompt_login()
    timed_goto(driver, "http://inside.sockettelecom.com/system/login.php")
    driver.page.fill("input[name='username']", user)
    driver.page.fill("input[name='password']", pw)
    driver.page.click("#login")
    # wait for the main iframe or dashboard to appear
    driver.page.wait_for_selector("iframe#MainView", timeout=10_000)
    clear_first_time_overlays(driver.page)

    # 3) Persist for next runs
    driver.save_state()
    log_message("✅ Logged in via credentials")

def clear_first_time_overlays(page):
    """
    Dismiss any first-time popups by clicking through known
    “Close” or “OK” buttons until they’re gone.
    """
    # a list of selectors for the various popups you might hit
    selectors = [
        # the specific “Close This” button you showed
        'xpath=//input[@id="valueForm1" and @type="button"]',
        # any button with the value text “Close This”
        'xpath=//input[@value="Close This" and @type="button"]',
        # legacy forms
        'xpath=//form[starts-with(@id,"valueForm")]//input[@type="button"]',
        'xpath=//form[@id="f"]//input[@type="button"]'
    ]

    for sel in selectors:
        # keep clicking until no more of that selector appear
        while True:
            try:
                btn = page.wait_for_selector(sel, timeout=500)
                btn.click()
                # give the UI a moment to re-render
                page.wait_for_timeout(200)
            except PlaywrightTimeout:
                break

def install_chromium(log=print):
    log("=== install_chromium started ===")
    try:
        log(f"sys.frozen={getattr(sys, 'frozen', False)}")
        if getattr(sys, "frozen", False):
            log("Frozen branch: importing playwright.__main__")
            try:
                import playwright.__main__ as pw_cli
                log("Imported playwright.__main__ successfully")
            except Exception as ie:
                log(f"ImportError playwright.__main__: {ie}\n{traceback.format_exc()}")
                raise RuntimeError("Playwright package not found in the frozen bundle.") from ie

            old_argv = sys.argv.copy()
            sys.argv = ["playwright", "install", "chromium"]
            try:
                log("Calling pw_cli.main()")
                try:
                    pw_cli.main()
                    log("pw_cli.main() returned normally")
                except SystemExit as se:
                    log(f"pw_cli.main() called sys.exit({se.code}); continuing")
                    # You may check se.code: 0 means success; non-zero means failure.
                    if se.code != 0:
                        raise RuntimeError(f"playwright install exited with code {se.code}")
                except Exception as e:
                    log(f"Exception inside pw_cli.main(): {e}\n{traceback.format_exc()}")
                    raise
            finally:
                sys.argv = old_argv
        else:
            log("Script mode branch: calling subprocess")
            cmd = [sys.executable, "-m", "playwright", "install", "chromium"]
            log(f"Subprocess command: {cmd}")
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            log(f"Subprocess return code: {proc.returncode}")
            if proc.stdout:
                log(f"Subprocess stdout: {proc.stdout.strip()}")
            if proc.stderr:
                log(f"Subprocess stderr: {proc.stderr.strip()}")
            if proc.returncode != 0:
                raise RuntimeError(f"playwright install failed, return code {proc.returncode}")
    except Exception as e:
        log(f"Exception in install_chromium: {e}\n{traceback.format_exc()}")
        # Show error to user
        try:
            from tkinter import messagebox, Tk
            root = Tk(); root.withdraw()
            messagebox.showerror("Playwright Error", f"Failed to install Chromium:\n{e}\nSee diagnostic.log")
            root.destroy()
        except Exception as gui_e:
            print(f"Playwright install error: {e}; plus GUI error: {gui_e}")
        # Re-raise so caller knows install failed
        raise
    log("=== install_chromium finished ===")

def is_chromium_installed():
    """
    Try launching Chromium headless via sync API. Returns True if successful.
    """
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            browser.close()
        return True
    except PlaywrightError:
        return False
    except Exception:
        return False

def ensure_playwright(log=print):
    """
    Sync check: if Chromium not installed or broken, run install_chromium().
    """
    try:
        if not is_chromium_installed():
            # Inform user
            try:
                root = Tk()
                root.withdraw()
                messagebox.showinfo("Playwright", "Chromium not found; downloading browser binaries now. This may take a few minutes.")
                root.destroy()
            except Exception:
                print("Chromium not found; downloading browser binaries now...")

            install_chromium()

            # After install, re-check
            if not is_chromium_installed():
                raise RuntimeError("Install completed but Chromium still not launchable.")
    except Exception as e:
        # Log and show error to user, referencing the log file
        err_msg = f"Playwright setup failed: {e}\nSee log file for details"
        log(err_msg)
        try:
            root = Tk()
            root.withdraw()
            messagebox.showerror("Playwright Error", err_msg)
            root.destroy()
        except Exception:
            print(err_msg)
        # Optionally exit or re-raise
        raise


def log_message(msg, also_print=False):
    timestamp = datetime.now().strftime("[%H:%M:%S]")
    full_msg = f"{timestamp} {msg}"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(full_msg + "\n")
    if also_print:
        print(full_msg)

def format_dispatch_summary(driver):
    ci = get_customer_and_ticket_info_from_task(driver)
    if not ci or not ci["ticket"]:
        return None

    timed_goto(driver, ci["customer_url"], wait_until="load")
    wo_url, wo_number = get_dispatch_work_order_url(driver, ci["ticket"])
    if not wo_url:
        return None

    timed_goto(driver, wo_url)
    # wait for the form to load
    driver.wait_for_selector("#AdditionalNotes", state="attached", timeout=10_000)

    el = driver.locator("xpath=//td[@class='detailHeader' and normalize-space(text())='Status:']"
        "/following-sibling::td//span"
    ).first
    status = el.inner_text().strip().lower()
    log_message(f"WO {wo_number} status → {status!r}")

    # accept both "complete" and "completed"
    if status not in ("complete", "completed"):
        log_message(f"⚠️ WO {wo_number} is still uncompleted; skipping")
        return


    # pull raw strings
    arr_date = driver.page.locator("#ArrivalOnsite").input_value().strip()
    arr_time = driver.page.locator("#ArrivalTime").input_value().strip()
    dep_date = driver.page.locator("#CompletedDate").input_value().strip()
    dep_time = driver.page.locator("#CompletedTime").input_value().strip()

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

def update_notes_only(frame, task_id, summary_text, log=log_message):
    try:
        # 1) re-expand in case something collapsed it
        expand_task(frame, task_id)

        # 2) clear & fill the notes textarea
        notes = frame.locator(f"#txtNotes{task_id}")
        notes.clear()
        notes.fill(summary_text)

        # 3) click the update button
        btn = frame.locator(f"#sub_{task_id}")
        btn.click()

        log(f"✏️  Updated notes for task {task_id}")
        return True

    except Exception as e:
        log_message(f"❌ update_notes_only failed: {type(e).__name__} - {e}")
        return False

def finalize_task(page: Page, task_id: int, summary_text: str, is_free: bool) -> bool:
    try:
        # Always use the MainView iframe context
        frame = page.frame(name="MainView")
        if not frame:
            raise Exception("MainView iframe not found!")

        form_sel = f'form#TOSSTask{task_id}'

        # Wait explicitly for form visibility IN FRAME
        frame.wait_for_selector(form_sel, timeout=10_000)

        # Fill the Notes box
        notes_sel = f'#{ "txtNotes" + str(task_id) }'
        frame.wait_for_selector(notes_sel, timeout=10_000)
        frame.fill(notes_sel, summary_text)

        # For billable tasks, tick "Spawn Billing Task"
        if not is_free:
            billing_sel = 'input[name="SpawnBillingTask"]'
            frame.wait_for_selector(billing_sel, timeout=5_000)
            billing_checkbox = frame.locator(billing_sel)
            if not billing_checkbox.is_checked():
                billing_checkbox.check()

        # Mark as completed
        completed_sel = f'#completedcheck{task_id}'
        frame.wait_for_selector(completed_sel, timeout=5_000)
        completed_checkbox = frame.locator(completed_sel)
        if not completed_checkbox.is_checked():
            completed_checkbox.check()

        # Click the Update Task button
        update_button_sel = f'#sub_{task_id}'
        frame.click(update_button_sel)

        log_message(f"✅ Task {task_id} successfully finalized")
        return True

    except PlaywrightTimeout as e:
        log_message(f"❌ Timeout in finalize_task for task {task_id}: {e}")
        return False
    except Exception as e:
        log_message(f"❌ Error in finalize_task for task {task_id}: {e}")
        return False

def attach_network_listeners(page):
    page.on("response", lambda response: _log_response(response))
    page.on("requestfailed", lambda request: _log_failure(request))

def _log_response(response):
    status = response.status
    url    = response.url
    if status == 429:
        log_message(f"🚫 RATE LIMIT hit on {url} (429 Too Many Requests)")
    elif status >= 400:
        log_message(f"⚠️ HTTP {status} on {url}")

def _log_failure(request):
    # only log XHR/fetch failures; skip images, CSS, etc.
    if request.resource_type != "xhr":
        return

    # request.failure is a string (or None), not a callable
    reason = request.failure or "<no error text>"
    log_message(f"❌ XHR to {request.url} failed: {reason}")

# === Consultation Task Extraction ===
def parse_task_row(row):
    try:
        # grab all the <td> cells
        tds = row.locator("td")
        count = tds.count()
        if count < 6:
            return None

        # first cell → <a href="…">
        url = tds.nth(0).locator("a").get_attribute("href")

        # text of the other cells
        desc     = tds.nth(1).inner_text().strip()
        assigned = tds.nth(4).inner_text().strip()
        company  = tds.nth(5).inner_text().strip()

        return {
            "url":      url,
            "desc":     desc,
            "assigned": assigned,
            "company":  company,
        }
    except Exception:
        return None

def get_customer_and_ticket_info_from_task(driver):
    # ── enter MainView frame if present ────────────────────────────
    try:
        driver.wait_for_selector("iframe#MainView", timeout=5_000)
        frame = driver.frame(name="MainView")
    except:
        log_message("⚠️ Already in MainView or frame not needed.")
        frame = driver.main_frame()

    # ── extract Customer ID & Name ───────────────────────────────────
    try:
        cid = frame.locator(
            "xpath=//td[normalize-space(text())='Customer ID']"
            "/following-sibling::td/b"
        ).inner_text().strip()
        customer_name = frame.locator(
            "xpath=//td[normalize-space(text())='Customer Name']"
            "/following-sibling::td/b"
        ).inner_text().strip()
    except Exception as e:
        log_message(f"❌ Failed to extract Customer ID/Name: {e}")
        return None

    # ── extract Ticket # ─────────────────────────────────────────────
    ticket_number = None
    # 1) primary: look for “Dispatch for Ticket 12345”
    try:
        dispatch_handle = frame.locator("b", has_text="Dispatch for Ticket")
        dispatch_handle.wait_for(timeout=10_000)
        dispatch_text = dispatch_handle.inner_text().strip()
        ticket_id    = dispatch_text.split()[-1]
    except Exception:
        log_message("⚠️ Could not find Ticket # in page or URL")

    customer_url = (
        "http://inside.sockettelecom.com/menu.php"
        f"?coid=1&tabid=7&parentid=9&customerid={cid}"
    )
    return {
        "customer_name": customer_name,
        "cid":            cid,
        "ticket":         ticket_id,
        "customer_url":   customer_url
    }

def get_dispatch_work_order_url(driver, ticket_number, log=None):
    # 2) Grab the right frame
    try:
        iframe_el = driver.wait_for_selector('iframe[name="MainView"]', timeout=10_000)
        frame = iframe_el.content_frame()
    except PlaywrightTimeout:
        debug_frame_html(driver.page)
        frame = driver.main_frame()

    # 3) Wait for the work orders table
    try:
        frame.wait_for_selector("#custWork #workShow table tr", timeout=10_000)
    except PlaywrightTimeout:
        log(f"⚠️ No Work Orders table found for ticket {ticket_number}")
        debug_frame_html(driver.page)        # ← debug here
        return None, None

    # 4) Collect every row, skipping the header
    rows = frame.locator("#custWork #workShow table tr")
    count = rows.count()
    if count == 0:
        log(f"⚠️ Found zero rows in Work Orders for ticket {ticket_number}")
        debug_frame_html(driver.page)        # ← and debug here too
        return None, None

    dispatch_wos = []
    for i in range(count):
        row = rows.nth(i)
        tds = row.locator("td")
        if tds.count() < 5:
            continue

        first = tds.nth(0).inner_text().strip()
        # skip header row
        if first == "#" or not first.isdigit():
            continue

        wo_num = int(first)
        desc   = tds.nth(1).inner_text().strip().lower()
        link   = tds.nth(4).locator("a").get_attribute("href")

        if re.search(rf"ticket\s*#?\s*{ticket_number}", desc, re.IGNORECASE):
            dispatch_wos.append((wo_num, link))

    if not dispatch_wos:
        log(f"⚠️ No dispatch WOs found for Ticket #{ticket_number}")
        debug_frame_html(driver.page)        # ← and debug here as well
        return None, None

    # 5) Return the most-recent WO
    wo_number, pwo_url = max(dispatch_wos, key=lambda x: x[0])
    wo_url = urljoin("http://inside.sockettelecom.com/", pwo_url)
    return wo_url, wo_number

def extract_work_order_notes(driver):
    try:
        driver.wait_for_selector("#AdditionalNotes", timeout=10_000)

        fields = {
            "EquipmentInstalled":  "",
            "AdditionalMaterials": "",
            "TestsPerformed":      "",
            "AdditionalNotes":     ""
        }

        for fid in fields:
            try:
                val = driver.locator(f"#{fid}").input_value().strip()
                fields[fid] = val
                if val:
                    log_message(f"📄 {fid} → {len(val)} chars")
            except Exception as e:
                log_message(f"⚠️ Could not read {fid}: {e}")

        combined = "\n".join(
            f"{label.replace('Additional','Additional ').replace('Performed','Performed:')}: {txt}"
            for label, txt in fields.items() if txt
        )

        return {"fields": fields, "combined": combined.strip()}

    except Exception as e:
        log_message(f"❌ Failed to extract WO notes: {e}")
        return {"fields": {}, "combined": ""}

def extract_due_consultation_tasks(driver):
    page = driver.page

    # 1) Navigate & wait
    log_message(f"\n🔎 Opening task URL: {TASK_URL}")
    page.goto(TASK_URL)
    page.wait_for_selector("iframe#MainView", timeout=10_000)

    # 2) Grab the frame by name
    frame = page.frame(name="MainView")
    if frame is None:
        frame = page.main_frame()

    log_message("Loading Tasks…", also_print=True)

    frame.wait_for_selector("//tr[contains(@class,'taskElement')]", timeout=30_000)
    rows = frame.locator("//tr[contains(@class,'taskElement')]")
    today = date.today()
    due = []

    for i in range(rows.count()):
        row = rows.nth(i)
        task = parse_task_row(row)
        if not task or "consultation" not in task["desc"].lower():
            continue

        try:
            due_str = row.locator("td:nth-child(4) nobr").inner_text().strip()
            due_dt  = datetime.strptime(due_str, "%Y-%m-%d").date()
        except Exception:
            log_message(f"⚠️ Couldn't parse due date '{due_str}'; skipping", also_print=True)
            continue

        if due_dt > today:
            log_message(f"⏳ Skipping '{task['desc']}' (due {due_dt.isoformat()})", also_print=True)
            continue

        due.append(task)

    log_message(f"✅ Found {len(due)} due consultation tasks.", also_print=True)
    return due

def extract_task_id_from_page(driver):
    """
    Returns the current Task ID by reading the hidden nTaskID input
    inside the MainView frame, or None if not found.
    """
    page = driver.page

    # 1) Make sure the frame is there
    try:
        page.wait_for_selector("iframe#MainView", timeout=5_000)
        frame = page.frame(name="MainView") or page.main_frame()
    except Exception:
        # no frame → fall back
        frame = page.main_frame()

    # 2) Look for the hidden input by name
    locator = frame.locator("[name=nTaskID]")
    if locator.count() == 0:
        return None

    # 3) Return its value
    try:
        task_id = locator.input_value().strip()
        return task_id or None
    except Exception:
        return None

def parse_job_type_from_task(driver, url):
    try:
        page = driver.page

        # 1) Navigate & wait
        log_message(f"\n🔎 Opening task URL: {url}")
        page.goto(url)
        page.wait_for_selector("iframe#MainView", timeout=10_000)

        # 2) Grab the frame by name
        frame = page.frame(name="MainView")
        if frame is None:
            frame = page.main_frame()

        # …now use `frame` for everything below…
        textarea = frame.wait_for_selector("[name=Notes]", timeout=10_000)
        raw_notes = textarea.input_value().strip()
        lower = raw_notes.lower()

        if "courtesy dispatch" in lower or "no charge" in lower:
            return "Consultation"

        for patt in [r"\bphone check\b", r"\bjack\b", r"\bfxs\b",
                     r"\bdial tone\b", r"\bno dial tone\b"]:
            if re.search(patt, lower):
                return "Phone Check"

        if any(k in lower for k in ["go live", "activate", "turn up"]):
            return "Go-Live"
        if any(k in lower for k in ["speed test", "throughput", "latency"]):
            return "Speed Test"
        if any(k in lower for k in ["nid", "modem swap"]):
            return "NID/IW/CopperTest"

        match = re.search(r"PROBLEM STATEMENT(?:\s*\(Statement\))?:\s*<b>(.*?)</b>",
                          raw_notes, re.IGNORECASE)
        if match:
            return match.group(1).strip()

        match = re.search(r"PROBLEM STATEMENT(?:\s*\(Statement\))?:\s*(.+)",
                          raw_notes, re.IGNORECASE)
        if match:
            log_message("⚠️ Found plain problem statement")
            text = re.sub(r"</?[^>]+>", "", match.group(1).strip())
            return text[:100].strip()

        for line in raw_notes.splitlines()[:15]:
            if "ont" in line.lower() and 3 < len(line.strip()) < 100:
                log_message("⚠️ Using fallback ONT line")
                return line.strip()

        log_message("❌ Could not identify job type — returning 'Unknown'")
        log_message(f"WO Notes: {raw_notes}")
        return "Unknown"

    except Exception as e:
        log_message(f"❌ Failed to parse job type from {url}: {e}")
        raise

def complete_free_task(page: Page, task_id: int, summary_text: str) -> bool:
    # 1) Expand the task details
    page.click(f'#displaySpan{task_id} + img')
    page.wait_for_selector(f'form#TOSSTask{task_id}', timeout=5000)

    # 2) Fill in the notes and submit
    page.fill(
        f'form#TOSSTask{task_id} textarea[name="Notes"]',
        summary_text
    )
    page.click(
        f'form#TOSSTask{task_id} input[type="submit"][value="Update Task"]'
    )
    page.wait_for_load_state('networkidle')

    # 3) Tick the Completed box
    completed = page.locator(
        f'form#TOSSTask{task_id} input[name="nCompleted"]'
    )
    completed.wait_for(state='visible', timeout=5000)
    if not completed.is_checked():
        completed.check()

    return True

def complete_charged_task(page: Page, task_id: int, summary_text: str) -> bool:
    # 1) Expand the task details
    page.click(f'#displaySpan{task_id} + img')
    page.wait_for_selector(f'form#TOSSTask{task_id}', timeout=5000)

    # 2) Fill in the notes and submit
    page.fill(
        f'form#TOSSTask{task_id} textarea[name="Notes"]',
        summary_text
    )
    page.click(
        f'form#TOSSTask{task_id} input[type="submit"][value="Update Task"]'
    )
    page.wait_for_load_state('networkidle')

    # 3) Spawn the billing sub-task
    billing = page.locator(
        f'form#TOSSTask{task_id} input[name="SpawnBillingTask"]'
    )
    billing.wait_for(state='visible', timeout=5000)
    if not billing.is_checked():
        billing.check()

    # 4) Tick the Completed box
    completed = page.locator(
        f'form#TOSSTask{task_id} input[name="nCompleted"]'
    )
    completed.wait_for(state='visible', timeout=5000)
    if not completed.is_checked():
        completed.check()

    return True

def normalize_note_content(text):
    if not text:
        return ""
    # Convert <br> and <br/> to \n
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    # Remove all other HTML tags
    text = re.sub(r"<[^>]+>", "", text)
    # Collapse all whitespace to single spaces, except for newlines
    text = re.sub(r"[ \t]+", " ", text)
    # Collapse multiple blank lines
    text = re.sub(r"\n+", "\n", text)
    return text.strip()

def extract_static_summary_block(summary_text):
    idx = summary_text.find("CUSTOMER:")
    if idx >= 0:
        return summary_text[idx:].strip()
    return summary_text.strip()

def notes_already_contain_summary(frame, task_id, summary_text, log=log_message):
    log(f"===> notes_already_contain_summary CALLED for {task_id}")
    try:
        # Look for the <td> that contains the history; use frame.locator
        notes_td = frame.locator("xpath=//td[contains(., 'CUSTOMER:')]")
        notes_html = notes_td.inner_html()

        notes   = normalize_note_content(notes_html)
        summary = normalize_note_content(
                      extract_static_summary_block(summary_text)
                  )

        found = bool(summary and summary in notes)
        log(f"🔍 SUMMARY FOUND IN NOTES? {found}")
        return found

    except Exception as e:
        log(f"⚠️ Failed to read notes for Task {task_id}: {e}")
        return False

def expand_task(frame, task_id, log=log_message):
    """
    Expands the hidden task form inside MainView iframe.
    `frame` should be driver.frame(name="MainView") or driver.main_frame().
    """
    span_sel   = f"#displaySpan{task_id}"
    # This locator jumps from the span to its legend in one shot:
    legend_sel = f"{span_sel} >> xpath=ancestor::fieldset[1]//legend"

    try:
        # 1) Ensure the legend is in the DOM
        legend = frame.locator(legend_sel)
        legend.wait_for(timeout=5_000)

        # 2) If the span is still hidden, click the legend
        if not frame.locator(span_sel).is_visible():
            legend.scroll_into_view_if_needed(timeout=5_000)
            legend.click(force=True)
            # 3) Wait for the span to become visible
            frame.wait_for_selector(f"{span_sel}:not([style*='display:none'])",
                                     timeout=5_000)
        log(f"✅ Task {task_id} expanded")
    except TimeoutError as e:
        log(f"❌ expand_task(): Timeout waiting for legend or span → {e}")
    except Exception as e:
        log(f"❌ expand_task(): Unexpected error expanding {task_id} → {e}")

def handle_sigterm(signum, frame):
    print("Received SIGTERM, exiting gracefully.")
    sys.exit(0)

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

def has_existing_notes(frame, task_id):
    # grab everything in the span *before* the <form> tag
    html = frame.locator(f"#displaySpan{task_id}").inner_html()
    before_form = html.split("<form", 1)[0]
    # strip out any tags and whitespace
    plain = re.sub(r"<[^>]+>", "", before_form).strip()
    return bool(plain)

def run_with_progress(driver, complete_free=False):
    results, errors = [], []
    due_tasks = extract_due_consultation_tasks(driver)

    for task in tqdm(due_tasks, desc="Processing consultation tasks", unit="task"):
        try:
            # 1) parse job type
            job_type = parse_job_type_from_task(driver, task["url"])
            is_free = is_free_job(job_type)[1]
            is_bill = is_billable_job(job_type)[1]
            task_id = None

            # 2) open & expand
            driver.goto(task["url"])
            frame = driver.frame(name="MainView") or driver.main_frame()
            expand_task(frame, task_id := extract_task_id_from_page(driver))

            # 3) skip if notes already exist
            if has_existing_notes(frame, task_id):
                log_message(f"⏭️ Task {task_id} already has notes, skipping")
                continue

            # 4) format summary
            summary_text = format_dispatch_summary(driver)
            if not summary_text:
                log_message(f"⚠️ No summary for Task {task_id}, skipping")
                continue


            driver.goto(task["url"])
            frame = driver.frame(name="MainView") or driver.main_frame()
            expand_task(frame, task_id)
            # 5) finalize in one shot
            success = finalize_task(driver.page, task_id, summary_text, is_free)
            if success:
                mode = "Free" if is_free else "Billable"
                log_message(f"✔️ Task {task_id} {mode} completed")
                results.append({
                    "Company":     task["company"],
                    "Description": task["desc"],
                    "URL":         task["url"],
                    "Job Type":    job_type,
                    "Task ID":     task_id,
                    "Mode":        mode
                })
            else:
                log_message(f"⚠️ Failed to finalize Task {task_id}")
                debug_frame_html(driver)

        except Exception as e:
            tb = traceback.format_exc()
            errors.append({"Task": task, "Error": str(e), "Traceback": tb})
            log_message(f"❌ Error for {task['desc']} — {e}")
            log_message(tb)

    return results, errors

def debug_frame_html(driver):
    """
    Prints the URL and outerHTML of whichever frame we're in
    (MainView if present, otherwise main_frame).
    """
    try:
        # try to grab the MainView iframe
        iframe_el = driver.wait_for_selector('iframe#MainView', timeout=5_000)
        frame = iframe_el.content_frame()
    except TimeoutError:
        frame = driver.main_frame()

    print(f"\n[DEBUG] Frame URL: {frame.url}\n")

    # grab the full HTML of that frame
    html = frame.content()  # returns the entire HTML as a string
    snippet = html[:2_000].replace("\n", " ")
    print(f"[DEBUG] Frame HTML (first 2000 chars):\n{snippet!r}\n...")
    print("[DEBUG] (truncated) end of dump\n")

    # 2. Count all <table> elements
    try:
        frame.wait_for_selector("table", timeout=5_000)
        tables = frame.locator("table")
        print(f"[DEBUG] Found {tables.count()} <table> elements")
        snippet = tables.nth(0).inner_html()[:200].replace("\n", " ")
        print(f"[DEBUG] First table snippet: {snippet!r}")
    except TimeoutError:
        print("[DEBUG] No <table> tags found in MainView")

    # 3. Look for your WO links pattern
    links = frame.locator("xpath=//a[contains(@href,'view.php?nCount=')]")
    print(f"[DEBUG] Found {links.count()} WO-style links")
    for i in range(min(5, links.count())):
        a = links.nth(i)
        print(f"  • {a.inner_text().strip()!r} → {a.get_attribute('href')}")

def check_for_update():
    return

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--update',
        action='store_true',
        help="Check for a new version and apply it"
    )
    parser.add_argument(
        '--version',
        action='store_true',
        help="Print current version and exit"
    )
    args, remaining = parser.parse_known_args()

    if args.version:
        print(__version__)
        sys.exit(0)

    if args.update:
        check_for_update()
        sys.exit(0)
    signal.signal(signal.SIGTERM, handle_sigterm)
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = BROWSERS
    print(f"PLAYWRIGHT_BROWSERS_PATH set to {BROWSERS}")
    ensure_playwright()

    # clear log
    with open(LOG_FILE, "w", encoding="utf-8"):
        pass
    
    PW = sync_playwright().start()
    browser = PW.chromium.launch(headless=True)

    try:
        driver = PlaywrightDriver(
            headless=True,
            playwright=PW,
            browser=browser,
        )
        page = driver.page
        attach_network_listeners(page)
        handle_login(driver)
        clear_first_time_overlays(driver.page)

        results, errors = run_with_progress(driver, complete_free=True)
        log_message(f"\n✅ Done. Parsed {len(results)} tasks with {len(errors)} errors.", True)
        summarize_job_types(results)

    except KeyboardInterrupt:
        print("Keyboard Interrupt caught")
        sys.exit(0)

    except Exception as e:
        print(f"Unexpected error, aborting: {e}")
        sys.exit(1)

    finally:
        try:
            driver.close()
        except Exception:
            pass
        try:
            browser.close()
            PW.stop()
        except Exception:
            pass
        input("Press Enter to exit...")
        sys.exit(0)
    