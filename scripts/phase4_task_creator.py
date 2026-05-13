from __future__ import annotations
import argparse, csv, queue, re, threading, time
from pathlib import Path
from urllib.parse import urljoin
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException, NoAlertPresentException
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver, WebElement
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

"""
PHASE 4 - Task Creator and Missing Assignee Fixer

What this script does:
1) Logs in the same way as the previous scripts.
2) Opens each ticket from a ticket list.
3) Opens Tasks/Time.
4) If the ticket already has tasks:
   - fills missing Assigned To values using manual overrides first, then the Phase 3 DB.
5) If the ticket has zero tasks:
   - creates the standard task set:
        Developer Code Review -> Code Review
        Code Review 1 -> Code Review
        Code Review 2 -> Code Review
        QA Development item Sign Off -> QA Review
        Release Sign Off -> BS Review
   - assigns names using manual overrides first, then the Phase 3 DB.
   - sets Status to Closed.
   - checks all visible/enabled checkboxes if they appear.
6) Saves an audit Excel/CSV file.

Install:
    python -m pip install selenium webdriver-manager pandas openpyxl

Run:
    python ticket_phase4_task_creator_gui.py
"""

S = {
    "login_user": (By.CSS_SELECTOR, "input#user"),
    "login_pw": (By.CSS_SELECTOR, "input#pw"),
    "login_submit": (By.CSS_SELECTOR, "input[name='ctl01']"),
    "ticket_ready": (By.CSS_SELECTOR, "form#ctl00"),
    "project": (By.CSS_SELECTOR, "select#project"),
    "tasks_link": (By.CSS_SELECTOR, "li#tasks a"),
    "iframe": (By.CSS_SELECTOR, "iframe"),
    "tasks_table": (By.CSS_SELECTOR, "table#mytable"),
    "add_task": (By.XPATH, "//a[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'add new task')]"),
    "form": (By.CSS_SELECTOR, "form#Form1"),
    "desc": (By.CSS_SELECTOR, "input#desc"),
    "assigned": (By.CSS_SELECTOR, "select#assigned_to"),
    "status": (By.CSS_SELECTOR, "select#status"),
    "task_type": (By.CSS_SELECTOR, "select#task_type"),
    "submit": (By.CSS_SELECTOR, "input[type='submit']"),
}

STANDARD_TASKS = [
    ("Developer Code Review", "Code Review"),
    ("Code Review 1", "Code Review"),
    ("Code Review 2", "Code Review"),
    ("QA Development item Sign Off", "QA Review"),
    ("Release Sign Off", "BS Review"),
]

# Manual override from your message. These are tried before the Phase 3 DB.
MANUAL_OVERRIDES = {
    ("saleshub", "qa release final sign-off"): "ASSIGNEE_A",
    ("saleshub", "qa release final sign off"): "ASSIGNEE_A",
    ("saleshub", "release sign off"): "ASSIGNEE_A",
    ("saleshub", "release sign-off"): "ASSIGNEE_A",
}

COLUMNS = [
    "ticket_number", "ticket_url", "project", "tasks_url",
    "task_id", "edit_url", "task_description",
    "previous_assigned_to", "assigned_to_used", "assignee_source",
    "task_type_used", "previous_status", "target_status",
    "action", "attempted", "success",
    "assigned_to_after", "status_after",
    "checkboxes_visible_enabled_count", "checkboxes_checked_after_count",
    "error",
]

def norm(v: str) -> str:
    return re.sub(r"\s+", " ", (v or "").replace("\xa0", " ")).strip()

def key(v: str) -> str:
    return norm(v).lower()

def missing_assignee(v: str) -> bool:
    return key(v) in {"", "[not assigned]", "not assigned", "none", "null", "-", "--"}

def ticket_url(base: str, num: str) -> str:
    base, num = base.strip(), str(num).strip()
    return f"{base}{num}" if base.endswith(("=", "/")) else f"{base.rstrip('/')}/{num}"

def read_tickets(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(path)
    suf = path.suffix.lower()
    if suf == ".txt":
        return [x.strip() for x in path.read_text(encoding="utf-8").splitlines() if x.strip()]
    if suf == ".csv":
        out = []
        with path.open("r", newline="", encoding="utf-8-sig") as f:
            for row in csv.reader(f):
                if row and row[0].strip():
                    out.append(row[0].strip())
        return out
    if suf in {".xlsx", ".xls"}:
        df = pd.read_excel(path)
        c = df.columns[0]
        return [str(x).strip() for x in df[c].dropna().tolist() if str(x).strip()]
    raise ValueError("Ticket list must be .txt, .csv, .xlsx, or .xls")

def driver_new(headless=False, user_data_dir="") -> WebDriver:
    opts = webdriver.ChromeOptions()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-notifications")
    if user_data_dir:
        opts.add_argument(f"--user-data-dir={user_data_dir}")
    return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opts)

def wait(driver: WebDriver, name: str, timeout=20) -> WebElement:
    by, sel = S[name]
    return WebDriverWait(driver, timeout).until(lambda d: d.find_element(by, sel))

def find(driver: WebDriver, name: str) -> WebElement | None:
    try:
        by, sel = S[name]
        return driver.find_element(by, sel)
    except NoSuchElementException:
        return None

def selected(el: WebElement) -> str:
    try:
        return norm(Select(el).first_selected_option.text)
    except Exception:
        return ""

def choose(el: WebElement, value: str, field: str):
    target = norm(value)
    sel = Select(el)
    try:
        sel.select_by_visible_text(target)
        return
    except Exception:
        pass
    for opt in sel.options:
        if key(opt.text) == key(target):
            opt.click()
            return
    for opt in sel.options:
        if key(target) in key(opt.text) or key(opt.text) in key(target):
            opt.click()
            return
    raise ValueError(f"Could not find '{value}' in {field} dropdown")

def is_login(driver: WebDriver) -> bool:
    return find(driver, "login_user") is not None and find(driver, "login_pw") is not None

def login_if_needed(driver: WebDriver, username: str, password: str, log=None) -> tuple[bool, bool]:
    if not is_login(driver):
        return False, True
    if log:
        log("Login page detected. Attempting login...")
    if not username or not password:
        return True, False
    u = wait(driver, "login_user", 30)
    p = wait(driver, "login_pw", 30)
    u.clear(); u.send_keys(username)
    p.clear(); p.send_keys(password)
    wait(driver, "login_submit", 30).click()
    try:
        WebDriverWait(driver, 30).until(lambda d: find(d, "ticket_ready") is not None or not is_login(d))
    except TimeoutException:
        pass
    ok = not is_login(driver)
    if log:
        log("Login appears successful." if ok else "Login still appears to be on the login page.")
    return True, ok

def open_ticket(driver: WebDriver, url: str, username: str, password: str, log=None) -> str:
    try:
        driver.get(url)
        attempted, ok = login_if_needed(driver, username, password, log)
        if attempted and ok:
            driver.get(url)
        if attempted and not ok:
            return "Login failed or login page did not clear."
        wait(driver, "ticket_ready", 30)
        WebDriverWait(driver, 30).until(lambda d: d.execute_script("return document.readyState") == "complete")
        return ""
    except Exception as e:
        return f"{type(e).__name__}: {e}"

def read_project(driver: WebDriver) -> str:
    return selected(wait(driver, "project", 20))

def read_tasks_url(driver: WebDriver, current_ticket_url: str) -> str:
    a = find(driver, "tasks_link")
    if not a:
        return ""
    return urljoin(current_ticket_url, norm(a.get_attribute("href") or ""))

def switch_to_tasks_context(driver: WebDriver):
    """
    Move into the Tasks/Time iframe if present.

    Important fix:
    Some tickets have zero tasks. On those pages, table#mytable may not exist,
    but the "add new task" link still exists. So this function must NOT require
    table#mytable to be present.
    """
    driver.switch_to.default_content()

    # If task table or add link is already in the main document, stay there.
    if find(driver, "tasks_table") or find(driver, "add_task"):
        return

    # Otherwise switch into the iframe.
    iframe = wait(driver, "iframe", 20)
    driver.switch_to.frame(iframe)

    # Do not wait for table#mytable here. No-task pages may not have one.
    WebDriverWait(driver, 20).until(lambda d: find(d, "tasks_table") is not None or find(d, "add_task") is not None)

def parse_tasks(driver: WebDriver, tnum: str, turl: str, project: str, tasks_url: str) -> list[dict]:
    table = find(driver, "tasks_table")
    tasks = []
    if table is None:
        return tasks

    for tr in table.find_elements(By.CSS_SELECTOR, "tbody tr"):
        tds = tr.find_elements(By.CSS_SELECTOR, "td")
        if len(tds) < 4:
            continue
        task_id = norm(tds[0].text)
        if not task_id.isdigit():
            continue
        edit_url = ""
        try:
            edit_url = urljoin(tasks_url, tds[1].find_element(By.CSS_SELECTOR, "a").get_attribute("href") or "")
        except NoSuchElementException:
            pass
        tasks.append({
            "ticket_number": tnum,
            "ticket_url": turl,
            "project": project,
            "tasks_url": tasks_url,
            "task_id": task_id,
            "edit_url": edit_url,
            "description": norm(tds[3].text) if len(tds) > 3 else "",
            "assigned_to": norm(tds[4].text) if len(tds) > 4 else "",
            "status": norm(tds[-1].text) if tds else "",
        })
    return tasks

def open_tasks(driver: WebDriver, tnum: str, turl: str, project: str, tasks_url: str) -> list[dict]:
    driver.switch_to.default_content()
    driver.get(tasks_url)
    WebDriverWait(driver, 20).until(lambda d: d.execute_script("return document.readyState") == "complete")
    switch_to_tasks_context(driver)
    time.sleep(0.25)
    return parse_tasks(driver, tnum, turl, project, tasks_url)

def add_task_url(driver: WebDriver, tasks_url: str) -> str:
    """
    Return the add-new-task URL.

    Primary method:
      use the page's "add new task" link.

    Fallback:
      If the link is not found but the tasks URL is known, construct:
      edit_task.aspx?id=0&bugid=<ticket>
    This matches the screenshot you provided for adding a new task:
      /edit_task.aspx?id=0&bugid=12345
    """
    driver.switch_to.default_content()
    if not find(driver, "add_task"):
        try:
            driver.switch_to.frame(wait(driver, "iframe", 5))
        except TimeoutException:
            pass

    a = find(driver, "add_task")
    if a:
        href = a.get_attribute("href") or ""
        if href:
            return urljoin(tasks_url, href)

    # Fallback based on tasks_frame.aspx?bugid=XXXXX or tasks.aspx?bugid=XXXXX
    match = re.search(r"bugid=([0-9]+)", tasks_url)
    if match:
        bugid = match.group(1)
        return urljoin(tasks_url, f"edit_task.aspx?id=0&bugid={bugid}")

    return ""

def load_db(path: Path) -> dict[tuple[str, str], str]:
    if not path.exists():
        raise FileNotFoundError(path)
    df = pd.read_excel(path, sheet_name="assignment_map")
    m = {}
    for _, r in df.iterrows():
        project = norm(str(r.get("project", "")))
        desc = norm(str(r.get("task_description", "")))
        assigned = norm(str(r.get("recommended_assigned_to", "")))
        if project and desc and assigned:
            m[(key(project), key(desc))] = assigned
    return m

def assignee_for(project: str, desc: str, db: dict[tuple[str, str], str]) -> tuple[str, str]:
    k = (key(project), key(desc))
    if k in MANUAL_OVERRIDES:
        return MANUAL_OVERRIDES[k], "manual_override"
    if k in db:
        return db[k], "assignment_db"
    return "", "no_mapping"

def open_edit(driver: WebDriver, url: str):
    driver.switch_to.default_content()
    driver.get(url)
    WebDriverWait(driver, 20).until(lambda d: d.execute_script("return document.readyState") == "complete")
    driver.switch_to.default_content()
    if not find(driver, "form"):
        try:
            driver.switch_to.frame(wait(driver, "iframe", 5))
        except TimeoutException:
            pass
    wait(driver, "form", 20)

def check_all_boxes(driver: WebDriver) -> tuple[int, int]:
    form = wait(driver, "form", 20)
    boxes = []
    for cb in form.find_elements(By.CSS_SELECTOR, "input[type='checkbox']"):
        try:
            if cb.is_displayed() and cb.is_enabled():
                boxes.append(cb)
        except WebDriverException:
            pass
    for cb in boxes:
        if not cb.is_selected():
            cb.click()
            time.sleep(0.05)
    checked = sum(1 for cb in boxes if cb.is_selected())
    return len(boxes), checked

def alert_text(driver: WebDriver) -> str:
    try:
        a = driver.switch_to.alert
        txt = a.text
        a.accept()
        return txt
    except (NoAlertPresentException, WebDriverException):
        return ""

def submit(driver: WebDriver):
    wait(driver, "submit", 20).click()

def blank_result(tnum, turl, project, tasks_url, action, error="", success=None) -> dict:
    return {
        "ticket_number": tnum, "ticket_url": turl, "project": project, "tasks_url": tasks_url,
        "task_id": "", "edit_url": "", "task_description": "",
        "previous_assigned_to": "", "assigned_to_used": "", "assignee_source": "",
        "task_type_used": "", "previous_status": "", "target_status": "",
        "action": action, "attempted": False, "success": success,
        "assigned_to_after": "", "status_after": "",
        "checkboxes_visible_enabled_count": None, "checkboxes_checked_after_count": None,
        "error": error,
    }

def update_assignee(driver: WebDriver, task: dict, assignee: str, source: str, dry: bool, log=None) -> dict:
    r = blank_result(task["ticket_number"], task["ticket_url"], task["project"], task["tasks_url"], "fill_missing_assigned_to")
    r.update({
        "task_id": task["task_id"], "edit_url": task["edit_url"], "task_description": task["description"],
        "previous_assigned_to": task["assigned_to"], "assigned_to_used": assignee,
        "assignee_source": source, "previous_status": task["status"],
    })
    if not assignee:
        r["success"] = False; r["error"] = "No assignee mapping found."
        return r
    if dry:
        r["action"] = "dry_run_would_fill_missing_assigned_to"; r["success"] = None; r["error"] = "Dry run only."
        return r
    try:
        open_edit(driver, task["edit_url"])
        choose(wait(driver, "assigned", 20), assignee, "Assigned to")
        r["attempted"] = True
        submit(driver)
        a = alert_text(driver)
        if a:
            r["error"] = f"Alert after submit: {a}"
        time.sleep(1)
        open_edit(driver, task["edit_url"])
        r["assigned_to_after"] = selected(wait(driver, "assigned", 20))
        r["success"] = key(r["assigned_to_after"]) == key(assignee)
        if not r["success"] and not r["error"]:
            r["error"] = f"Assigned To after update is '{r['assigned_to_after']}', expected '{assignee}'."
    except Exception as e:
        r["success"] = False; r["error"] = f"{type(e).__name__}: {e}"
    return r

def create_standard_task(driver: WebDriver, tnum: str, turl: str, project: str, tasks_url: str, add_url: str, desc: str, task_type: str, assignee: str, source: str, dry: bool) -> dict:
    r = blank_result(tnum, turl, project, tasks_url, "create_standard_task")
    r.update({"task_description": desc, "assigned_to_used": assignee, "assignee_source": source, "task_type_used": task_type, "target_status": "Closed"})
    if not add_url:
        r["success"] = False; r["error"] = "Add new task URL not found."
        return r
    if not assignee:
        r["success"] = False; r["error"] = "No assignee mapping found for this task."
        return r
    if dry:
        r["action"] = "dry_run_would_create_standard_task"; r["success"] = None; r["error"] = "Dry run only."
        return r
    try:
        open_edit(driver, add_url)
        d = wait(driver, "desc", 20)
        d.clear(); d.send_keys(desc)
        choose(wait(driver, "assigned", 20), assignee, "Assigned to")
        choose(wait(driver, "status", 20), "Closed", "Status")
        choose(wait(driver, "task_type", 20), task_type, "Task Type")
        box_count, checked_count = check_all_boxes(driver)
        r["checkboxes_visible_enabled_count"] = box_count
        r["checkboxes_checked_after_count"] = checked_count
        r["attempted"] = True
        submit(driver)
        a = alert_text(driver)
        if a:
            r["error"] = f"Alert after submit: {a}"
        time.sleep(1)
        tasks_after = open_tasks(driver, tnum, turl, project, tasks_url)
        matches = [x for x in tasks_after if key(x["description"]) == key(desc)]
        if not matches:
            r["success"] = False
            if not r["error"]:
                r["error"] = "Submitted but could not verify created task in list."
            return r
        m = matches[-1]
        r["task_id"] = m["task_id"]; r["edit_url"] = m["edit_url"]
        r["assigned_to_after"] = m["assigned_to"]; r["status_after"] = m["status"]
        r["success"] = key(m["assigned_to"]) == key(assignee) and key(m["status"]) == "closed"
        if not r["success"] and not r["error"]:
            r["error"] = f"Verification mismatch. Assigned='{m['assigned_to']}', Status='{m['status']}'."
    except Exception as e:
        r["success"] = False; r["error"] = f"{type(e).__name__}: {e}"
    return r

def save_audit(rows: list[dict], path: Path):
    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(columns=COLUMNS)
    else:
        for c in COLUMNS:
            if c not in df.columns:
                df[c] = ""
        df = df[COLUMNS]
    if path.suffix.lower() == ".csv":
        df.to_csv(path, index=False)
    else:
        df.to_excel(path, index=False)

class App:
    def __init__(self, root):
        self.root = root
        root.title("Phase 4 - Task Creator and Missing Assignee Fixer")
        root.geometry("1040x760")
        self.q = queue.Queue()
        self.pq = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None

        self.base = tk.StringVar(value="YOUR_TICKET_SYSTEM_URL/edit_bug.aspx?id=")
        self.tickets = tk.StringVar()
        self.db = tk.StringVar(value=str(Path.cwd() / "phase3_assignment_db.xlsx"))
        self.out = tk.StringVar(value=str(Path.cwd() / "phase4_task_create_audit.xlsx"))
        self.user = tk.StringVar()
        self.pw = tk.StringVar()
        self.profile = tk.StringVar()
        self.headless = tk.BooleanVar(value=False)
        self.dry = tk.BooleanVar(value=True)
        self.pause = tk.DoubleVar(value=0.5)
        self.layout()
        root.after(200, self.poll)

    def layout(self):
        f = ttk.Frame(self.root); f.pack(fill="both", expand=True, padx=10, pady=10)
        pad = {"padx": 8, "pady": 5}
        labels = [
            ("Base ticket URL:", self.base, None),
            ("Ticket list file:", self.tickets, self.pick_tickets),
            ("Phase 3 assignment DB:", self.db, self.pick_db),
            ("Output audit file:", self.out, self.pick_out),
            ("Username:", self.user, None),
            ("Password:", self.pw, None),
            ("Chrome user data dir optional:", self.profile, self.pick_profile),
        ]
        for i, (txt, var, cmd) in enumerate(labels):
            ttk.Label(f, text=txt).grid(row=i, column=0, sticky="w", **pad)
            show = "*" if txt == "Password:" else None
            ttk.Entry(f, textvariable=var, width=100 if i not in {4,5} else 40, show=show).grid(row=i, column=1, sticky="ew" if i not in {4,5} else "w", **pad)
            if cmd:
                ttk.Button(f, text="Browse" if txt != "Output audit file:" else "Save As", command=cmd).grid(row=i, column=2, **pad)

        opts = ttk.Frame(f); opts.grid(row=7, column=1, sticky="w", **pad)
        ttk.Checkbutton(opts, text="Headless mode not recommended", variable=self.headless).pack(side="left")
        ttk.Checkbutton(opts, text="Dry run only - do not update/create", variable=self.dry).pack(side="left", padx=(20,0))
        ttk.Label(opts, text="Pause seconds:").pack(side="left", padx=(20,5))
        ttk.Spinbox(opts, textvariable=self.pause, from_=0.0, to=10.0, increment=0.25, width=6).pack(side="left")

        btns = ttk.Frame(f); btns.grid(row=8, column=1, sticky="w", **pad)
        self.start_btn = ttk.Button(btns, text="Start Phase 4", command=self.start); self.start_btn.pack(side="left", padx=(0,8))
        self.stop_btn = ttk.Button(btns, text="Stop after current ticket", command=self.stop, state="disabled"); self.stop_btn.pack(side="left")

        info = "If a ticket has no tasks, creates the standard task set. If a ticket has tasks, fills missing Assigned To only. Assignee source: manual SalesHub override first, then Phase 3 DB."
        ttk.Label(f, text=info, foreground="gray", wraplength=940).grid(row=9, column=0, columnspan=3, sticky="w", padx=8, pady=(10,5))

        self.progress = ttk.Progressbar(f, orient="horizontal", mode="determinate")
        self.progress.grid(row=10, column=0, columnspan=3, sticky="ew", padx=8, pady=(12,5))
        self.status = ttk.Label(f, text="Ready"); self.status.grid(row=11, column=0, columnspan=3, sticky="w", padx=8, pady=5)
        ttk.Label(f, text="Live progress log:").grid(row=12, column=0, columnspan=3, sticky="w", padx=8, pady=(12,5))
        self.logbox = tk.Text(f, height=22, wrap="word"); self.logbox.grid(row=13, column=0, columnspan=3, sticky="nsew", padx=8, pady=5)
        sb = ttk.Scrollbar(f, orient="vertical", command=self.logbox.yview); sb.grid(row=13, column=3, sticky="ns")
        self.logbox.configure(yscrollcommand=sb.set)
        f.columnconfigure(1, weight=1); f.rowconfigure(13, weight=1)

    def pick_tickets(self):
        x = filedialog.askopenfilename(filetypes=[("Ticket files","*.txt *.csv *.xlsx *.xls"),("All files","*.*")])
        if x: self.tickets.set(x)
    def pick_db(self):
        x = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls"),("All files","*.*")])
        if x: self.db.set(x)
    def pick_out(self):
        x = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel file","*.xlsx"),("CSV file","*.csv")])
        if x: self.out.set(x)
    def pick_profile(self):
        x = filedialog.askdirectory()
        if x: self.profile.set(x)
    def log(self, msg):
        self.q.put(msg)
    def stop(self):
        self.stop_event.set()
        self.log("Stop requested. Will stop after current ticket.")
        self.stop_btn.configure(state="disabled")
    def start(self):
        if not self.base.get().strip() or not self.tickets.get().strip() or not self.db.get().strip() or not self.out.get().strip():
            messagebox.showerror("Missing info", "Fill Base URL, ticket list, Phase 3 DB, and output file.")
            return
        if self.worker and self.worker.is_alive():
            messagebox.showerror("Already running", "Phase 4 is already running.")
            return
        self.stop_event.clear()
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.logbox.delete("1.0", "end")
        self.progress.configure(value=0, maximum=1)
        self.worker = threading.Thread(target=self.run, daemon=True)
        self.worker.start()
    def poll(self):
        try:
            while True:
                self.logbox.insert("end", self.q.get_nowait() + "\n")
                self.logbox.see("end")
        except queue.Empty:
            pass
        try:
            while True:
                cur, total = self.pq.get_nowait()
                self.progress.configure(maximum=max(total,1), value=cur)
                self.status.configure(text=f"Progress: {cur}/{total}")
        except queue.Empty:
            pass
        if self.worker and not self.worker.is_alive():
            self.start_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")
            self.worker = None
        self.root.after(200, self.poll)

    def run(self):
        rows = []
        drv = None
        out_path = Path(self.out.get().strip())
        try:
            nums = read_tickets(Path(self.tickets.get().strip()))
            db = load_db(Path(self.db.get().strip()))
            self.log(f"Loaded {len(nums)} tickets and {len(db)} DB mappings.")
            if self.dry.get():
                self.log("DRY RUN is ON. No updates or creates will be submitted.")
            self.pq.put((0, len(nums)))
            drv = driver_new(self.headless.get(), self.profile.get().strip())
            for idx, num in enumerate(nums, start=1):
                if self.stop_event.is_set():
                    self.log("Stopped before next ticket.")
                    break
                turl = ticket_url(self.base.get(), num)
                self.log(f"[{idx}/{len(nums)}] Processing ticket {num}...")
                err = open_ticket(drv, turl, self.user.get(), self.pw.get(), self.log)
                if err:
                    self.log(f"  Ticket error: {err}")
                    rows.append(blank_result(num, turl, "", "", "ticket_open_failed", err, False))
                    save_audit(rows, out_path); self.pq.put((idx, len(nums))); continue

                try:
                    project = read_project(drv)
                    tasks_url = read_tasks_url(drv, turl)
                    if not tasks_url:
                        rows.append(blank_result(num, turl, project, "", "tasks_link_not_found", "Tasks/Time link not found.", False))
                        save_audit(rows, out_path); self.pq.put((idx, len(nums))); continue

                    tasks = open_tasks(drv, num, turl, project, tasks_url)
                    self.log(f"  Project='{project}'. Existing tasks found: {len(tasks)}.")

                    if not tasks:
                        add_url = add_task_url(drv, tasks_url)
                        self.log("  No tasks found. Creating standard task set.")
                        for desc, typ in STANDARD_TASKS:
                            assignee, source = assignee_for(project, desc, db)
                            self.log(f"  Create '{desc}' type='{typ}' assignee='{assignee}' source='{source}'")
                            res = create_standard_task(drv, num, turl, project, tasks_url, add_url, desc, typ, assignee, source, self.dry.get())
                            rows.append(res); save_audit(rows, out_path)
                            self.log("    Success." if res["success"] else f"    Not successful: {res['error']}")
                            time.sleep(float(self.pause.get()))
                    else:
                        missing = [t for t in tasks if missing_assignee(t["assigned_to"])]
                        self.log(f"  Tasks with missing Assigned To: {len(missing)}.")
                        if not missing:
                            rows.append(blank_result(num, turl, project, tasks_url, "no_missing_assigned_to_and_tasks_exist", "", True))
                            save_audit(rows, out_path)
                        for task in missing:
                            assignee, source = assignee_for(project, task["description"], db)
                            self.log(f"  Fill task {task['task_id']} '{task['description']}' assignee='{assignee}' source='{source}'")
                            res = update_assignee(drv, task, assignee, source, self.dry.get(), self.log)
                            rows.append(res); save_audit(rows, out_path)
                            self.log("    Success." if res["success"] else f"    Not successful: {res['error']}")
                            time.sleep(float(self.pause.get()))
                except Exception as e:
                    err = f"{type(e).__name__}: {e}"
                    self.log(f"  Processing error: {err}")
                    rows.append(blank_result(num, turl, "", "", "ticket_processing_error", err, False))
                    save_audit(rows, out_path)
                self.pq.put((idx, len(nums)))
            self.log(f"Audit saved to: {out_path}")
            self.log("Phase 4 finished.")
        except Exception as e:
            self.log(f"Fatal error: {type(e).__name__}: {e}")
            try: save_audit(rows, out_path)
            except Exception: pass
        finally:
            if drv:
                try: drv.quit()
                except Exception: pass

def main():
    argparse.ArgumentParser(description="Phase 4 Task Creator and Missing Assignee Fixer").parse_args()
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
