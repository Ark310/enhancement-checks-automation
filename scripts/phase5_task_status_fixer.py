
from __future__ import annotations

import argparse
import csv
import queue
import re
import threading
import time
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException, NoAlertPresentException
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver, WebElement
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

"""
PHASE 5 - Explicit Task Name + Closed Status Fixer

Checks existing tasks only. It does not create tasks.

For each selected ticket:
- Opens ticket and logs in if needed.
- Opens Tasks/Time.
- If Assigned To is missing, fills it using the explicit Phase 5 mapping.
- If Status is not Closed, changes it to Closed.
- Checks all visible/enabled checkboxes on the task edit page.
- Clicks Update.
- Saves an audit file.

Install:
    python -m pip install selenium webdriver-manager pandas openpyxl

Run:
    python ticket_phase5_task_fix_gui.py
"""

SELECTORS = {
    "login_user": (By.CSS_SELECTOR, "input#user"),
    "login_pw": (By.CSS_SELECTOR, "input#pw"),
    "login_submit": (By.CSS_SELECTOR, "input[name='ctl01']"),
    "ticket_ready": (By.CSS_SELECTOR, "form#ctl00"),
    "project": (By.CSS_SELECTOR, "select#project"),
    "tasks_link": (By.CSS_SELECTOR, "li#tasks a"),
    "iframe": (By.CSS_SELECTOR, "iframe"),
    "tasks_table": (By.CSS_SELECTOR, "table#mytable"),
    "task_form": (By.CSS_SELECTOR, "form#Form1"),
    "assigned": (By.CSS_SELECTOR, "select#assigned_to"),
    "status": (By.CSS_SELECTOR, "select#status"),
    "submit": (By.CSS_SELECTOR, "input[type='submit']"),
}

PHASE5_ASSIGNMENT_MAP = {
    ("tradedesk client server", "release sign off"): "ASSIGNEE_B",
    ("tradedesk client server", "release sign-off"): "ASSIGNEE_B",
    ("business modeling", "release sign off"): "ASSIGNEE_C",
    ("business modeling", "release sign-off"): "ASSIGNEE_C",
    ("contoso erp", "release sign off"): "ASSIGNEE_B",
    ("contoso erp", "release sign-off"): "ASSIGNEE_B",

    ("*", "qa release final sign-off"): "ASSIGNEE_A",
    ("*", "qa release final sign off"): "ASSIGNEE_A",

    ("*", "qa development item sign off"): "ASSIGNEE_D",
    ("*", "qa development item sign-off"): "ASSIGNEE_D",

    ("contoso erp", "developer code review"): "ASSIGNEE_E",
    ("business modeling", "developer code review"): "ASSIGNEE_F",
    ("rest api", "developer code review"): "ASSIGNEE_F",

    ("business modeling", "code review 2"): "ASSIGNEE_G",
    ("contoso erp", "code review 2"): "ASSIGNEE_H",

    ("contoso erp", "code review 1"): "ASSIGNEE_B",
    ("business modeling", "code review 1"): "ASSIGNEE_C",
}

AUDIT_COLUMNS = [
    "ticket_number", "ticket_url", "project", "tasks_url",
    "task_id", "edit_url", "task_description",
    "previous_assigned_to", "previous_status",
    "mapped_assigned_to", "mapping_source",
    "needed_assigned_to_update", "needed_status_update",
    "attempted", "success",
    "assigned_to_after", "status_after",
    "checkboxes_visible_enabled_count", "checkboxes_checked_after_count",
    "error", "action",
]


def normalize_text(value: str) -> str:
    value = value or ""
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_key(value: str) -> str:
    return normalize_text(value).lower()


def is_missing_assignee(value: str) -> bool:
    return normalize_key(value) in {"", "[not assigned]", "not assigned", "none", "null", "-", "--"}


def is_closed_status(value: str) -> bool:
    return normalize_key(value) == "closed"


def build_ticket_url(base_url: str, ticket_number: str) -> str:
    base_url = base_url.strip()
    ticket_number = str(ticket_number).strip()
    if base_url.endswith("=") or base_url.endswith("/"):
        return f"{base_url}{ticket_number}"
    return f"{base_url.rstrip('/')}/{ticket_number}"


def read_ticket_numbers(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"Ticket list not found: {path}")

    suffix = path.suffix.lower()
    if suffix == ".txt":
        return [line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]

    if suffix == ".csv":
        tickets = []
        with path.open("r", newline="", encoding="utf-8-sig") as f:
            for row in csv.reader(f):
                if row and row[0].strip():
                    tickets.append(row[0].strip())
        return tickets

    if suffix in {".xlsx", ".xls"}:
        df = pd.read_excel(path)
        first_col = df.columns[0]
        return [str(v).strip() for v in df[first_col].dropna().tolist() if str(v).strip()]

    raise ValueError("Ticket list must be .txt, .csv, .xlsx, or .xls")


def make_driver(headless: bool, user_data_dir: str) -> WebDriver:
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    if user_data_dir:
        options.add_argument(f"--user-data-dir={user_data_dir}")
    return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)


def wait_for(driver: WebDriver, name: str, timeout: int = 20) -> WebElement:
    by, selector = SELECTORS[name]
    return WebDriverWait(driver, timeout).until(lambda d: d.find_element(by, selector))


def safe_find(driver: WebDriver, name: str) -> WebElement | None:
    try:
        by, selector = SELECTORS[name]
        return driver.find_element(by, selector)
    except NoSuchElementException:
        return None


def selected_dropdown_text(element: WebElement) -> str:
    try:
        return normalize_text(Select(element).first_selected_option.text)
    except Exception:
        return ""


def is_login_page(driver: WebDriver) -> bool:
    return safe_find(driver, "login_user") is not None and safe_find(driver, "login_pw") is not None


def login_if_needed(driver: WebDriver, username: str, password: str, log=None) -> tuple[bool, bool]:
    if not is_login_page(driver):
        return False, True

    if log:
        log("Login page detected. Attempting login...")

    if not username or not password:
        return True, False

    user_el = wait_for(driver, "login_user", 30)
    pw_el = wait_for(driver, "login_pw", 30)
    user_el.clear()
    user_el.send_keys(username)
    pw_el.clear()
    pw_el.send_keys(password)
    wait_for(driver, "login_submit", 30).click()

    try:
        WebDriverWait(driver, 30).until(
            lambda d: safe_find(d, "ticket_ready") is not None or not is_login_page(d)
        )
    except TimeoutException:
        pass

    success = not is_login_page(driver)
    if log:
        log("Login appears successful." if success else "Login still appears to be on the login page.")
    return True, success


def open_ticket(driver: WebDriver, ticket_url: str, username: str, password: str, log=None) -> str:
    try:
        driver.get(ticket_url)
        attempted, success = login_if_needed(driver, username, password, log=log)
        if attempted and success:
            driver.get(ticket_url)
        if attempted and not success:
            return "Login failed or login page did not clear."
        wait_for(driver, "ticket_ready", 30)
        WebDriverWait(driver, 30).until(lambda d: d.execute_script("return document.readyState") == "complete")
        return ""
    except Exception as exc:
        return f"{type(exc).__name__}: {exc}"


def read_project(driver: WebDriver) -> str:
    return selected_dropdown_text(wait_for(driver, "project", 20))


def read_tasks_url(driver: WebDriver, ticket_url: str) -> str:
    link = safe_find(driver, "tasks_link")
    if link is None:
        return ""
    return urljoin(ticket_url, normalize_text(link.get_attribute("href") or ""))


def switch_to_tasks_context(driver: WebDriver) -> None:
    driver.switch_to.default_content()
    if safe_find(driver, "tasks_table") is not None:
        return
    iframe = wait_for(driver, "iframe", 20)
    driver.switch_to.frame(iframe)
    wait_for(driver, "tasks_table", 20)


def parse_tasks(driver: WebDriver, ticket_number: str, ticket_url: str, project: str, tasks_url: str) -> list[dict]:
    table = safe_find(driver, "tasks_table")
    if table is None:
        return []

    tasks = []
    for row in table.find_elements(By.CSS_SELECTOR, "tbody tr"):
        cells = row.find_elements(By.CSS_SELECTOR, "td")
        if len(cells) < 4:
            continue

        task_id = normalize_text(cells[0].text)
        if not task_id.isdigit():
            continue

        edit_url = ""
        try:
            edit_link = cells[1].find_element(By.CSS_SELECTOR, "a")
            edit_url = urljoin(tasks_url, edit_link.get_attribute("href") or "")
        except NoSuchElementException:
            pass

        tasks.append(
            {
                "ticket_number": ticket_number,
                "ticket_url": ticket_url,
                "project": project,
                "tasks_url": tasks_url,
                "task_id": task_id,
                "edit_url": edit_url,
                "description": normalize_text(cells[3].text) if len(cells) > 3 else "",
                "assigned_to": normalize_text(cells[4].text) if len(cells) > 4 else "",
                "status": normalize_text(cells[-1].text) if cells else "",
            }
        )
    return tasks


def open_tasks_and_read(driver: WebDriver, ticket_number: str, ticket_url: str, project: str, tasks_url: str) -> list[dict]:
    driver.switch_to.default_content()
    driver.get(tasks_url)
    WebDriverWait(driver, 20).until(lambda d: d.execute_script("return document.readyState") == "complete")
    switch_to_tasks_context(driver)
    time.sleep(0.25)
    return parse_tasks(driver, ticket_number, ticket_url, project, tasks_url)


def lookup_assignee(project: str, task_description: str) -> tuple[str, str]:
    project_key = normalize_key(project)
    desc_key = normalize_key(task_description)
    if (project_key, desc_key) in PHASE5_ASSIGNMENT_MAP:
        return PHASE5_ASSIGNMENT_MAP[(project_key, desc_key)], "project_task_mapping"
    if ("*", desc_key) in PHASE5_ASSIGNMENT_MAP:
        return PHASE5_ASSIGNMENT_MAP[("*", desc_key)], "global_task_mapping"
    return "", "no_mapping"


def open_task_edit_context(driver: WebDriver, edit_url: str) -> None:
    driver.switch_to.default_content()
    driver.get(edit_url)
    WebDriverWait(driver, 20).until(lambda d: d.execute_script("return document.readyState") == "complete")
    driver.switch_to.default_content()
    if safe_find(driver, "task_form") is None:
        try:
            iframe = wait_for(driver, "iframe", 5)
            driver.switch_to.frame(iframe)
        except TimeoutException:
            pass
    wait_for(driver, "task_form", 20)


def choose_dropdown_value(select_el: WebElement, value: str, field_name: str) -> None:
    select = Select(select_el)
    target = normalize_text(value)

    try:
        select.select_by_visible_text(target)
        return
    except Exception:
        pass

    for option in select.options:
        if normalize_key(option.text) == normalize_key(target):
            option.click()
            return

    target_loose = normalize_key(target).replace(".", "").replace(" ", "")
    for option in select.options:
        option_loose = normalize_key(option.text).replace(".", "").replace(" ", "")
        if target_loose == option_loose:
            option.click()
            return

    for option in select.options:
        option_loose = normalize_key(option.text).replace(".", "").replace(" ", "")
        if target_loose and target_loose in option_loose:
            option.click()
            return

    raise ValueError(f"Could not find '{value}' in {field_name} dropdown.")


def check_all_visible_checkboxes(driver: WebDriver) -> tuple[int, int]:
    form = wait_for(driver, "task_form", 20)
    checkboxes = []
    for cb in form.find_elements(By.CSS_SELECTOR, "input[type='checkbox']"):
        try:
            if cb.is_displayed() and cb.is_enabled():
                checkboxes.append(cb)
        except WebDriverException:
            pass

    for cb in checkboxes:
        if not cb.is_selected():
            cb.click()
            time.sleep(0.05)

    checked_count = sum(1 for cb in checkboxes if cb.is_selected())
    return len(checkboxes), checked_count


def handle_alert(driver: WebDriver) -> str:
    try:
        alert = driver.switch_to.alert
        text = alert.text
        alert.accept()
        return text
    except (NoAlertPresentException, WebDriverException):
        return ""


def audit_blank(ticket_number, ticket_url, project, tasks_url, action, success=None, error="") -> dict:
    return {
        "ticket_number": ticket_number,
        "ticket_url": ticket_url,
        "project": project,
        "tasks_url": tasks_url,
        "task_id": "",
        "edit_url": "",
        "task_description": "",
        "previous_assigned_to": "",
        "previous_status": "",
        "mapped_assigned_to": "",
        "mapping_source": "",
        "needed_assigned_to_update": False,
        "needed_status_update": False,
        "attempted": False,
        "success": success,
        "assigned_to_after": "",
        "status_after": "",
        "checkboxes_visible_enabled_count": None,
        "checkboxes_checked_after_count": None,
        "error": error,
        "action": action,
    }


def fix_task(driver: WebDriver, task: dict, dry_run: bool) -> dict:
    mapped_assignee, mapping_source = lookup_assignee(task["project"], task["description"])
    need_assignee = is_missing_assignee(task["assigned_to"])
    need_status = not is_closed_status(task["status"])

    result = audit_blank(task["ticket_number"], task["ticket_url"], task["project"], task["tasks_url"], "fix_existing_task")
    result.update(
        {
            "task_id": task["task_id"],
            "edit_url": task["edit_url"],
            "task_description": task["description"],
            "previous_assigned_to": task["assigned_to"],
            "previous_status": task["status"],
            "mapped_assigned_to": mapped_assignee,
            "mapping_source": mapping_source,
            "needed_assigned_to_update": need_assignee,
            "needed_status_update": need_status,
        }
    )

    if not need_assignee and not need_status:
        result["action"] = "already_good"
        result["success"] = True
        return result

    if need_assignee and not mapped_assignee:
        result["success"] = False
        result["error"] = "Assigned To is missing, but no Phase 5 mapping exists for this project/task."
        return result

    if not task["edit_url"]:
        result["success"] = False
        result["error"] = "No edit URL found for task."
        return result

    if dry_run:
        result["action"] = "dry_run_would_fix_existing_task"
        result["success"] = None
        result["error"] = "Dry run only. No update submitted."
        return result

    try:
        open_task_edit_context(driver, task["edit_url"])

        if need_assignee:
            choose_dropdown_value(wait_for(driver, "assigned", 20), mapped_assignee, "Assigned to")

        if need_status:
            choose_dropdown_value(wait_for(driver, "status", 20), "Closed", "Status")

        box_count, checked_count = check_all_visible_checkboxes(driver)
        result["checkboxes_visible_enabled_count"] = box_count
        result["checkboxes_checked_after_count"] = checked_count

        result["attempted"] = True
        wait_for(driver, "submit", 20).click()

        alert_text = handle_alert(driver)
        if alert_text:
            result["error"] = f"Alert after submit: {alert_text}"

        time.sleep(1.0)

        open_task_edit_context(driver, task["edit_url"])
        result["assigned_to_after"] = selected_dropdown_text(wait_for(driver, "assigned", 20))
        result["status_after"] = selected_dropdown_text(wait_for(driver, "status", 20))

        assignee_ok = True
        if need_assignee:
            assignee_ok = normalize_key(result["assigned_to_after"]) == normalize_key(mapped_assignee)

        status_ok = is_closed_status(result["status_after"])
        result["success"] = assignee_ok and status_ok

        if not result["success"] and not result["error"]:
            result["error"] = (
                f"Verification failed. Assigned To after='{result['assigned_to_after']}', "
                f"expected='{mapped_assignee if need_assignee else task['assigned_to']}'. "
                f"Status after='{result['status_after']}', expected='Closed'."
            )

    except Exception as exc:
        result["success"] = False
        result["error"] = f"{type(exc).__name__}: {exc}"

    return result


def save_audit(rows: list[dict], output_path: Path) -> None:
    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(columns=AUDIT_COLUMNS)
    else:
        for col in AUDIT_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df[AUDIT_COLUMNS]

    if output_path.suffix.lower() == ".csv":
        df.to_csv(output_path, index=False)
    else:
        df.to_excel(output_path, index=False)


class Phase5App:
    def __init__(self, root):
        self.root = root
        self.root.title("Phase 5 - Missing Names and Closed Status Checker")
        self.root.geometry("1040x760")

        self.worker = None
        self.stop_event = threading.Event()
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()

        self.base_url_var = tk.StringVar(value="YOUR_TICKET_SYSTEM_URL/edit_bug.aspx?id=")
        self.ticket_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar(value=str(Path.cwd() / "phase5_task_fix_audit.xlsx"))
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.user_data_dir_var = tk.StringVar()
        self.headless_var = tk.BooleanVar(value=False)
        self.dry_run_var = tk.BooleanVar(value=True)
        self.pause_var = tk.DoubleVar(value=0.5)

        self.build_layout()
        self.root.after(200, self.poll)

    def build_layout(self):
        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        pad = {"padx": 8, "pady": 5}

        ttk.Label(frame, text="Base ticket URL:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.base_url_var, width=100).grid(row=0, column=1, sticky="ew", **pad)

        ttk.Label(frame, text="Ticket list file:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.ticket_file_var, width=100).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Browse", command=self.pick_ticket_file).grid(row=1, column=2, **pad)

        ttk.Label(frame, text="Output audit file:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.output_file_var, width=100).grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Save As", command=self.pick_output_file).grid(row=2, column=2, **pad)

        ttk.Label(frame, text="Username:").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.username_var, width=40).grid(row=3, column=1, sticky="w", **pad)

        ttk.Label(frame, text="Password:").grid(row=4, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.password_var, width=40, show="*").grid(row=4, column=1, sticky="w", **pad)

        ttk.Label(frame, text="Chrome user data dir optional:").grid(row=5, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.user_data_dir_var, width=100).grid(row=5, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Browse", command=self.pick_profile).grid(row=5, column=2, **pad)

        options = ttk.Frame(frame)
        options.grid(row=6, column=1, sticky="w", **pad)
        ttk.Checkbutton(options, text="Headless mode not recommended", variable=self.headless_var).pack(side="left")
        ttk.Checkbutton(options, text="Dry run only - do not update", variable=self.dry_run_var).pack(side="left", padx=(20, 0))
        ttk.Label(options, text="Pause seconds:").pack(side="left", padx=(20, 5))
        ttk.Spinbox(options, textvariable=self.pause_var, from_=0.0, to=10.0, increment=0.25, width=6).pack(side="left")

        buttons = ttk.Frame(frame)
        buttons.grid(row=7, column=1, sticky="w", **pad)
        self.start_button = ttk.Button(buttons, text="Start Phase 5", command=self.start)
        self.start_button.pack(side="left", padx=(0, 8))
        self.stop_button = ttk.Button(buttons, text="Stop after current ticket", command=self.stop, state="disabled")
        self.stop_button.pack(side="left")

        info = (
            "Checks existing tasks only. If Assigned To is missing, fills from the Phase 5 explicit mapping. "
            "If Status is not Closed, sets it to Closed and checks all visible/enabled checkboxes."
        )
        ttk.Label(frame, text=info, foreground="gray", wraplength=940).grid(row=8, column=0, columnspan=3, sticky="w", padx=8, pady=(10, 5))

        self.progress = ttk.Progressbar(frame, orient="horizontal", mode="determinate")
        self.progress.grid(row=9, column=0, columnspan=3, sticky="ew", padx=8, pady=(12, 5))

        self.status_label = ttk.Label(frame, text="Ready")
        self.status_label.grid(row=10, column=0, columnspan=3, sticky="w", padx=8, pady=5)

        ttk.Label(frame, text="Live progress log:").grid(row=11, column=0, columnspan=3, sticky="w", padx=8, pady=(12, 5))

        self.log_text = tk.Text(frame, height=23, wrap="word")
        self.log_text.grid(row=12, column=0, columnspan=3, sticky="nsew", padx=8, pady=5)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=12, column=3, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(12, weight=1)

    def pick_ticket_file(self):
        filename = filedialog.askopenfilename(
            title="Choose ticket list",
            filetypes=[("Ticket files", "*.txt *.csv *.xlsx *.xls"), ("All files", "*.*")],
        )
        if filename:
            self.ticket_file_var.set(filename)

    def pick_output_file(self):
        filename = filedialog.asksaveasfilename(
            title="Save Phase 5 audit",
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx"), ("CSV file", "*.csv")],
        )
        if filename:
            self.output_file_var.set(filename)

    def pick_profile(self):
        dirname = filedialog.askdirectory(title="Choose Chrome user data directory")
        if dirname:
            self.user_data_dir_var.set(dirname)

    def log(self, message):
        self.log_queue.put(message)

    def start(self):
        if not self.base_url_var.get().strip():
            messagebox.showerror("Missing base URL", "Enter the base ticket URL.")
            return
        if not self.ticket_file_var.get().strip():
            messagebox.showerror("Missing ticket list", "Choose a ticket list file.")
            return
        if not self.output_file_var.get().strip():
            messagebox.showerror("Missing output file", "Choose where to save the audit file.")
            return
        if self.worker and self.worker.is_alive():
            messagebox.showerror("Already running", "Phase 5 is already running.")
            return

        self.stop_event.clear()
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.progress.configure(value=0, maximum=1)
        self.status_label.configure(text="Starting...")
        self.log_text.delete("1.0", "end")

        self.worker = threading.Thread(target=self.run, daemon=True)
        self.worker.start()

    def stop(self):
        self.stop_event.set()
        self.log("Stop requested. The script will stop after the current ticket finishes.")
        self.stop_button.configure(state="disabled")

    def poll(self):
        try:
            while True:
                self.log_text.insert("end", self.log_queue.get_nowait() + "\n")
                self.log_text.see("end")
        except queue.Empty:
            pass

        try:
            while True:
                current, total = self.progress_queue.get_nowait()
                self.progress.configure(maximum=max(total, 1), value=current)
                self.status_label.configure(text=f"Progress: {current}/{total}")
        except queue.Empty:
            pass

        if self.worker and not self.worker.is_alive():
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.worker = None

        self.root.after(200, self.poll)

    def run(self):
        driver = None
        rows = []
        output_path = Path(self.output_file_var.get().strip())

        try:
            tickets = read_ticket_numbers(Path(self.ticket_file_var.get().strip()))
            self.progress_queue.put((0, len(tickets)))
            self.log(f"Loaded {len(tickets)} tickets.")
            if self.dry_run_var.get():
                self.log("DRY RUN is ON. No updates will be submitted.")

            driver = make_driver(
                headless=self.headless_var.get(),
                user_data_dir=self.user_data_dir_var.get().strip(),
            )

            for idx, ticket_number in enumerate(tickets, start=1):
                if self.stop_event.is_set():
                    self.log("Stopped before starting next ticket.")
                    break

                t_url = build_ticket_url(self.base_url_var.get(), ticket_number)
                self.log(f"[{idx}/{len(tickets)}] Processing ticket {ticket_number}...")

                err = open_ticket(driver, t_url, self.username_var.get(), self.password_var.get(), log=self.log)
                if err:
                    self.log(f"  Ticket error: {err}")
                    rows.append(audit_blank(ticket_number, t_url, "", "", "ticket_open_failed", False, err))
                    save_audit(rows, output_path)
                    self.progress_queue.put((idx, len(tickets)))
                    continue

                try:
                    project = read_project(driver)
                    tasks_url = read_tasks_url(driver, t_url)

                    if not tasks_url:
                        self.log("  Tasks/Time link not found.")
                        rows.append(audit_blank(ticket_number, t_url, project, "", "tasks_link_not_found", False, "Tasks/Time link not found."))
                        save_audit(rows, output_path)
                        self.progress_queue.put((idx, len(tickets)))
                        continue

                    tasks = open_tasks_and_read(driver, ticket_number, t_url, project, tasks_url)
                    self.log(f"  Project='{project}'. Tasks found: {len(tasks)}.")

                    if not tasks:
                        rows.append(audit_blank(ticket_number, t_url, project, tasks_url, "no_tasks_found", True, "No existing tasks found. Phase 5 does not create tasks."))
                        save_audit(rows, output_path)
                        self.progress_queue.put((idx, len(tickets)))
                        continue

                    actionable = [
                        task for task in tasks
                        if is_missing_assignee(task["assigned_to"]) or not is_closed_status(task["status"])
                    ]
                    self.log(f"  Tasks needing fix: {len(actionable)}.")

                    if not actionable:
                        rows.append(audit_blank(ticket_number, t_url, project, tasks_url, "all_tasks_good", True, ""))
                        save_audit(rows, output_path)
                    else:
                        for task in actionable:
                            if self.stop_event.is_set():
                                self.log("  Stop requested. Finishing current ticket loop.")
                                break

                            mapped, source = lookup_assignee(project, task["description"])
                            self.log(
                                f"  Task {task['task_id']} '{task['description']}': "
                                f"Assigned='{task['assigned_to']}', Status='{task['status']}', "
                                f"Mapped='{mapped}', Source='{source}'"
                            )

                            result = fix_task(driver, task, self.dry_run_var.get())
                            rows.append(result)
                            save_audit(rows, output_path)

                            if result["success"] is True:
                                self.log("    Fixed successfully.")
                            elif result["success"] is None:
                                self.log(f"    Dry run: {result['error']}")
                            else:
                                self.log(f"    Not fixed: {result['error']}")

                            if float(self.pause_var.get()) > 0:
                                time.sleep(float(self.pause_var.get()))

                except Exception as exc:
                    err = f"{type(exc).__name__}: {exc}"
                    self.log(f"  Processing error: {err}")
                    rows.append(audit_blank(ticket_number, t_url, "", "", "ticket_processing_error", False, err))
                    save_audit(rows, output_path)

                self.progress_queue.put((idx, len(tickets)))

            self.log(f"Audit saved to: {output_path}")
            self.log("Phase 5 finished.")

        except Exception as exc:
            self.log(f"Fatal error: {type(exc).__name__}: {exc}")
            try:
                save_audit(rows, output_path)
            except Exception:
                pass

        finally:
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass


def main():
    argparse.ArgumentParser(description="Phase 5 Missing Names and Closed Status Checker").parse_args()
    root = tk.Tk()
    Phase5App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
