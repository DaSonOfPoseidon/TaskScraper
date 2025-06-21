import os
import re
from time import perf_counter
import sys
import signal
import traceback
import getpass
from pathlib import Path
from urllib.parse import urljoin
from tqdm import tqdm
from rapidfuzz import fuzz, process
from datetime import datetime, date
from collections import Counter, defaultdict
from dotenv import load_dotenv, set_key

HERE        = os.path.abspath(os.path.dirname(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(HERE, os.pardir))
STATE_PATH = os.path.join(PROJECT_ROOT, "state.json")
OUTPUT_DIR  = os.path.join(PROJECT_ROOT, "Outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)
ENV_PATH    = os.path.join(PROJECT_ROOT, ".env")
LOG_FILE = os.path.join(OUTPUT_DIR, "consultation_log.txt")

TASK_URL = "http://inside.sockettelecom.com/menu.php?tabid=45&tasktype=2&nID=1439&width=1440&height=731"
PAGE_TIMEOUT = 30

DRY_RUN = False

if getattr(sys, "frozen", False):
    # sys._MEIPASS is the temp folder where PyInstaller unpacks data files
    base_path = sys._MEIPASS
else:
    # running in “dev” mode, point at your source tree
    base_path = os.path.dirname(__file__)

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

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
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

    def goto(self, url: str, *, timeout: int = 30_000, wait_until: str = "networkidle"):

        try:
            return self.page.goto(url, timeout=timeout, wait_until=wait_until)
        except PlaywrightTimeout:
            # fallback to load event if even DOMContentLoaded hung
            return self.page.goto(url, timeout=timeout, wait_until="load")
        

    def save_state(self, path: str = STATE_PATH):
        self.context.storage_state(path=path)

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

    timed_goto(driver, ci["customer_url"])
    wo_url, wo_number = get_dispatch_work_order_url(driver, ci["ticket"])
    if not wo_url:
        return None

    timed_goto(driver, wo_url)
    # wait for the form to load
    driver.wait_for_selector("#AdditionalNotes", state="attached", timeout=10_000)

    el = driver.locator("xpath=//td[@class='detailHeader' and normalize-space(text())='Status:']"
        "/following-sibling::td//span"
    ).first
    status = el.inner_text()

    if status != "complete":
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

def update_notes_only(driver, task_id, summary_text):
    try:
        # switch into the MainView frame
        frame = driver.frame(name="MainView")
        if not frame:
            raise RuntimeError("…")
        expand_task(driver, task_id)

        # update the notes field
        notes = driver.locator(f"#txtNotes{task_id}")
        notes.clear()
        notes.send_keys(summary_text)

        # click the Update button
        btn = driver.locator(f"sub_{task_id}").first
        btn.click()

        log_message(f"✏️  (DRY RUN) Updated notes for task {task_id}")
        return True

    except Exception as e:
        log_message(f"❌ update_notes_only failed: {type(e).__name__} - {e}")
        return False

def finalize_task(driver, task_id, summary_text, is_free):
    if DRY_RUN:
        return update_notes_only(driver, task_id, summary_text)
    return (complete_free_task if is_free else complete_charged_task)(
        driver, task_id, summary_text, screenshot_dir=None
    )

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

def complete_free_task(driver, task_id, job_type, screenshot_dir=None):
    frame = page.frame(name="MainView")
    if frame is None:
        frame = page.main_frame()


    def try_complete():
        def debug_el(label, sel, timeout=5_000):
            el = frame.wait_for_selector(sel, timeout=timeout)
            log_message(f"✔️ Found {label}: {sel}")
            return el

        chk = debug_el("checkbox", f"#completedcheck{task_id}")
        if not chk.is_checked():
            log_message("Clicking 'Completed' checkbox...")
            chk.click()
        else:
            log_message("Checkbox already selected.")

        notes = debug_el("notes box", f"#txtNotes{task_id}")
        notes.fill(f"{job_type}, no charge")

        btn = debug_el("submit button", f"#sub_{task_id}")
        log_message("Clicking 'Update Task' button...")
        btn.click()

    for attempt in range(2):
        try:
            log_message(f"--- Attempt {attempt+1} to complete task ---")
            try_complete()
            log_message(f"✅ Successfully completed as free ({job_type})")
            return True
        except Exception as e:
            log_message(f"⚠️ Attempt {attempt+1} failed: {type(e).__name__} - {e}")
            if attempt == 0:
                frame.wait_for_timeout(1_000)
                continue
            if screenshot_dir:
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                path = os.path.join(screenshot_dir, f"wo_{task_id}_fail_{ts}.png")
                driver.screenshot(path=path)
                log_message(f"📸 Screenshot saved to: {path}")
            log_message("❌ Gave up after retrying")
            return False

def complete_charged_task(driver, task_id, summary_text, screenshot_dir=None):
    frame = page.frame(name="MainView")
    if frame is None:
        frame = page.main_frame()

    try:
        frame.wait_for_selector("[name=SpawnBillingTask]", timeout=5_000)

        def click_if_needed(el, label):
            if not el.is_checked():
                log_message(f"✔️ Clicking {label}")
                el.click()

        checkbox = frame.locator(f"#completedcheck{task_id}")
        click_if_needed(checkbox, "Completed")

        spawn = frame.locator("[name=SpawnBillingTask]")
        click_if_needed(spawn, "SpawnBillingTask")

        notes = frame.locator(f"#txtNotes{task_id}")
        notes.fill(summary_text)

        btn = frame.locator(f"#sub_{task_id}")
        log_message("✔️ Clicking Update Task")
        btn.click()

        log_message("✅ Charged task completed")
        return True

    except Exception as e:
        log_message(f"❌ complete_charged_task failed: {e}")
        if screenshot_dir:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = os.path.join(screenshot_dir, f"fail_{task_id}_{ts}.png")
            driver.screenshot(path=path)
            log_message(f"📸 Screenshot: {path}")
        return False

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

def notes_already_contain_summary(driver, task_id, summary_text):
    log_message(f"===> notes_already_contain_summary CALLED for {task_id}")
    try:
        # Grab the <td> that has all the rendered notes history
        notes_td = driver.locator("xpath=//td[contains(., 'CUSTOMER:')]")
        # Playwright Locator.inner_html() returns the element's HTML
        notes_html = notes_td.inner_html()

        # Normalize both notes and summary for robust comparison
        notes   = normalize_note_content(notes_html)
        summary = normalize_note_content(extract_static_summary_block(summary_text))

        found = bool(summary and summary in notes)
        log_message(f"🔍 SUMMARY FOUND IN NOTES? {found}")
        return found

    except Exception as e:
        log_message(f"⚠️ Failed to read notes for Task {task_id}: {e}")
        return False

def expand_task(driver, task_id):
    try:
        # Locate the span wrapping the form
        span = driver.locator(f"#displaySpan{task_id}")
        # Climb up to its <fieldset>
        fieldset = span.locator("xpath=ancestor::fieldset[1]").first
        legend   = fieldset.locator("legend").first

        # If it's hidden or effectively zero-height, click to expand
        is_vis = span.is_visible()
        box    = span.bounding_box() or {}
        height = box.get("height", 0)
        if not is_vis or height < 5:
            legend.click()
            # wait until our span becomes visible
            driver.wait_for_selector(f"#displaySpan{task_id}", state="visible", timeout=5_000)

    except Exception as e:
        log_message(f"❌ expand_task(): Failed to expand Task ID {task_id}: {e}")

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

def run_with_progress(driver, complete_free=False):
    results = []
    errors = []

    due_tasks = extract_due_consultation_tasks(driver)

    for task in tqdm(due_tasks, desc="Processing consultation tasks", unit="task"):
        try:
            job_type = parse_job_type_from_task(driver, task["url"])

            _, is_free, _ = is_free_job(job_type)
            _, is_bill, _ = is_billable_job(job_type)

            # ── Handle unknown jobs (notes-only) ─────────────────────────────
            if not (is_free or is_bill):
                log_message(
                    f"⚠️ Unknown job type '{job_type}', will update notes only (no complete/billing)."
                )
                task_id = extract_task_id_from_page(driver)
                if not task_id:
                    log_message(f"⚠️ No Task ID found for '{task['desc']}', skipping")
                    continue

                summary_text = format_dispatch_summary(driver)
                if not summary_text:
                    log_message(f"⚠️ No summary for Task {task_id}, skipping")
                    continue

                timed_goto(driver, task["url"])
                driver.wait_for_selector(
                    "iframe#MainView", timeout=PAGE_TIMEOUT * 1000
                )
                expand_task(driver, task_id)

                if notes_already_contain_summary(driver, task_id, summary_text):
                    log_message(
                        f"⏭️ Task {task_id} already has summary notes, skipping update"
                    )
                    continue

                update_notes_only(driver, task_id, summary_text)
                log_message(
                    f"✔️ Task {task_id} (Unknown job) notes updated only (no complete/bill)"
                )
                results.append({
                    "Company":     task["company"],
                    "Description": task["desc"],
                    "URL":         task["url"],
                    "Job Type":    job_type,
                    "Task ID":     task_id,
                    "Mode":        "Unknown-NotesOnly"
                })
                continue

            # ── Known jobs (free or billable) ────────────────────────────────
            results.append({
                "Company":     task["company"],
                "Description": task["desc"],
                "URL":         task["url"],
                "Job Type":    job_type
            })

            task_id = extract_task_id_from_page(driver)
            if not task_id:
                log_message(f"⚠️ No Task ID found for '{task['desc']}', skipping")
                continue

            summary_text = format_dispatch_summary(driver)
            if not summary_text:
                log_message(f"⚠️ No summary for Task {task_id}, skipping")
                continue

            timed_goto(driver, task["url"])
            driver.wait_for_selector(
                "iframe#MainView", timeout=PAGE_TIMEOUT * 1000
            )
            expand_task(driver, task_id)

            success = finalize_task(driver, task_id, summary_text, is_free)
            if success:
                action = (
                    "(DRY RUN) notes updated" if DRY_RUN else
                    "completed free"      if is_free else
                    "charged & closed"
                )
                log_message(f"✔️ Task {task_id} {action}")
            else:
                log_message(f"⚠️ Failed to finalize Task {task_id}")

            results.append({
                "Company":     task["company"],
                "Description": task["desc"],
                "URL":         task["url"],
                "Job Type":    job_type,
                "Task ID":     task_id,
                "Mode":        "Free" if is_free else "Billable"
            })

        except Exception as e:
            tb = traceback.format_exc()
            errors.append({
                "Task":      task,
                "Error":     str(e),
                "Traceback": tb
            })
            log_message(f"❌ Error for {task['desc']} — {e}")
            log_message(tb)
            continue

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


if __name__ == "__main__":
    signal.signal(signal.SIGTERM, handle_sigterm)

    # clear log
    with open(LOG_FILE, "w", encoding="utf-8"):
        pass
    
    PW = sync_playwright().start()
    browser = PW.chromium.launch(headless=False)

    try:
        driver = PlaywrightDriver(
            headless=False,
            playwright=PW,
            browser=browser,
            state_path=STATE_PATH
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
    