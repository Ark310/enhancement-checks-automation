from __future__ import annotations

import argparse
import csv
import queue
import re
import threading
import time
from collections import Counter, defaultdict
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional, List, Tuple, Dict, Any
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


SELECTORS = {
    "login_username": (By.CSS_SELECTOR, "input#user"),
    "login_password": (By.CSS_SELECTOR, "input#pw"),
    "login_submit": (By.CSS_SELECTOR, "input[name='ctl01']"),
    "ticket_page_ready": (By.CSS_SELECTOR, "form#ctl00"),
    "project_select": (By.CSS_SELECTOR, "select#project"),
    "tasks_link": (By.CSS_SELECTOR, "li#tasks a"),
    "tasks_iframe": (By.CSS_SELECTOR, "iframe"),
    "tasks_table": (By.CSS_SELECTOR, "table#mytable"),
    "task_edit_form": (By.CSS_SELECTOR, "form#Form1"),
    "task_assigned_to_select": (By.CSS_SELECTOR, "select#assigned_to"),
    "task_update_button": (By.CSS_SELECTOR, "input#sub"),
}


@dataclass
class TaskObservation:
    ticket_number: str
    ticket_url: str
    project: str
    task_id: str
    task_description: str
    assigned_to: str
    status: str
    tasks_url: str
    edit_url: str


@dataclass
class AssignmentMapRow:
    project: str
    task_description: str
    recommended_assigned_to: str
    observation_count: int
    unique_assignee_count: int
    all_assignees_with_counts: str
    confidence: str


@dataclass
class AssigneeUpdateResult:
    ticket_number: str
    ticket_url: str
    project: str
    tasks_url: str
    task_id: str = ""
    edit_url: str = ""
    task_description: str = ""
    previous_assigned_to: str = ""
    recommended_assigned_to: str = ""
    status: str = ""
    action: str = ""
    update_attempted: bool = False
    update_success: Optional[bool] = None
    assigned_to_after_update: str = ""
    error: str = ""


def normalize_text(value: str) -> str:
    value = value or ""
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_key(value: str) -> str:
    return normalize_text(value).lower()


def is_missing_assignee(value: str) -> bool:
    clean = normalize_key(value)
    return clean in {"", "[not assigned]", "not assigned", "none", "null", "-", "--"}


def build_ticket_url(base_url: str, ticket_number: str) -> str:
    base_url = base_url.strip()
    ticket_number = ticket_number.strip()
    if base_url.endswith("=") or base_url.endswith("/"):
        return f"{base_url}{ticket_number}"
    return f"{base_url.rstrip('/')}/{ticket_number}"


def read_ticket_numbers(path: Path) -> List[str]:
    if not path.exists():
        raise FileNotFoundError(f"Ticket file not found: {path}")
    suffix = path.suffix.lower()
    if suffix == ".txt":
        return [line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
    if suffix == ".csv":
        tickets: List[str] = []
        with path.open("r", newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            for row in reader:
                if row and row[0].strip():
                    tickets.append(row[0].strip())
        return tickets
    if suffix in {".xlsx", ".xls"}:
        df = pd.read_excel(path)
        first_col = df.columns[0]
        return [str(v).strip() for v in df[first_col].dropna().tolist() if str(v).strip()]
    raise ValueError("Unsupported ticket file type. Use .txt, .csv, or .xlsx")


def make_driver(headless: bool = False, user_data_dir: str = "") -> WebDriver:
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    if user_data_dir:
        options.add_argument(f"--user-data-dir={user_data_dir}")
    service = ChromeService(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def wait_for(driver: WebDriver, selector_key: str, timeout: int = 20) -> WebElement:
    by, selector = SELECTORS[selector_key]
    return WebDriverWait(driver, timeout).until(lambda d: d.find_element(by, selector))


def safe_find(driver: WebDriver, selector_key: str) -> Optional[WebElement]:
    try:
        by, selector = SELECTORS[selector_key]
        return driver.find_element(by, selector)
    except NoSuchElementException:
        return None


def selected_text(select_el: WebElement) -> str:
    try:
        return normalize_text(Select(select_el).first_selected_option.text)
    except Exception:
        return ""


def is_login_page(driver: WebDriver) -> bool:
    return safe_find(driver, "login_username") is not None and safe_find(driver, "login_password") is not None


def login_if_needed(driver: WebDriver, username: str, password: str, log_func=None, timeout: int = 30) -> Tuple[bool, bool]:
    if not is_login_page(driver):
        return False, True
    if log_func:
        log_func("Login page detected. Attempting login...")
    if not username or not password:
        return True, False
    user_el = wait_for(driver, "login_username", timeout=timeout)
    pw_el = wait_for(driver, "login_password", timeout=timeout)
    user_el.clear(); user_el.send_keys(username)
    pw_el.clear(); pw_el.send_keys(password)
    wait_for(driver, "login_submit", timeout=timeout).click()
    try:
        WebDriverWait(driver, timeout).until(lambda d: safe_find(d, "ticket_page_ready") is not None or not is_login_page(d))
    except TimeoutException:
        pass
    success = not is_login_page(driver)
    if log_func:
        log_func("Login appears successful." if success else "Login still appears to be on the login page.")
    return True, success


def open_ticket_and_login_if_needed(driver: WebDriver, ticket_url: str, username: str, password: str, log_func=None) -> Tuple[bool, bool, str]:
    try:
        driver.get(ticket_url)
        attempted, success = login_if_needed(driver, username, password, log_func=log_func)
        if attempted and success:
            if log_func:
                log_func("Reloading ticket after login...")
            driver.get(ticket_url)
        if attempted and not success:
            return attempted, success, "Login failed or login page did not clear."
        wait_for(driver, "ticket_page_ready", timeout=30)
        WebDriverWait(driver, 30).until(lambda d: d.execute_script("return document.readyState") == "complete")
        return attempted, success, ""
    except Exception as exc:
        return False, False, f"{type(exc).__name__}: {exc}"


def read_project(driver: WebDriver) -> str:
    return selected_text(wait_for(driver, "project_select", timeout=20))


def read_tasks_link(driver: WebDriver, ticket_url: str) -> Tuple[str, str]:
    link = safe_find(driver, "tasks_link")
    if link is None:
        return "", ""
    text = normalize_text(link.text)
    href = normalize_text(link.get_attribute("href") or "")
    return text, urljoin(ticket_url, href)


def switch_into_tasks_table_context(driver: WebDriver, timeout: int = 20) -> None:
    driver.switch_to.default_content()
    if safe_find(driver, "tasks_table") is not None:
        return
    iframe = wait_for(driver, "tasks_iframe", timeout=timeout)
    driver.switch_to.frame(iframe)
    wait_for(driver, "tasks_table", timeout=timeout)


def parse_tasks_table(driver: WebDriver, ticket_number: str, ticket_url: str, project: str, tasks_url: str) -> List[TaskObservation]:
    table = wait_for(driver, "tasks_table", timeout=20)
    rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
    observations: List[TaskObservation] = []
    for row in rows:
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
        description = normalize_text(cells[3].text) if len(cells) > 3 else ""
        assigned_to = normalize_text(cells[4].text) if len(cells) > 4 else ""
        status = normalize_text(cells[-1].text) if cells else ""
        observations.append(TaskObservation(ticket_number, ticket_url, project, task_id, description, assigned_to, status, tasks_url, edit_url))
    return observations


def open_tasks_page_and_get_observations(driver: WebDriver, ticket_number: str, ticket_url: str, project: str, tasks_url: str, timeout: int = 20) -> List[TaskObservation]:
    driver.switch_to.default_content()
    driver.get(tasks_url)
    WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")
    switch_into_tasks_table_context(driver, timeout=timeout)
    time.sleep(0.25)
    return parse_tasks_table(driver, ticket_number, ticket_url, project, tasks_url)


def build_assignment_map(observations: List[TaskObservation]) -> Tuple[List[AssignmentMapRow], List[Dict[str, Any]]]:
    grouped: Dict[Tuple[str, str], Counter] = defaultdict(Counter)
    display_lookup: Dict[Tuple[str, str], Tuple[str, str]] = {}
    for obs in observations:
        key = (normalize_key(obs.project), normalize_key(obs.task_description))
        if key not in display_lookup:
            display_lookup[key] = (obs.project, obs.task_description)
        if obs.project and obs.task_description and not is_missing_assignee(obs.assigned_to):
            grouped[key][obs.assigned_to] += 1
    map_rows: List[AssignmentMapRow] = []
    ambiguous: List[Dict[str, Any]] = []
    for key, counter in grouped.items():
        project_display, task_display = display_lookup.get(key, (key[0], key[1]))
        most_common = counter.most_common()
        recommended, _ = most_common[0]
        total = sum(counter.values())
        unique_count = len(counter)
        all_counts = "; ".join(f"{name}: {cnt}" for name, cnt in most_common)
        confidence = "high"
        if unique_count > 1:
            confidence = "tie" if most_common[0][1] == most_common[1][1] else "mixed"
            ambiguous.append({
                "project": project_display,
                "task_description": task_display,
                "recommended_assigned_to": recommended,
                "all_assignees_with_counts": all_counts,
                "confidence": confidence,
            })
        map_rows.append(AssignmentMapRow(project_display, task_display, recommended, total, unique_count, all_counts, confidence))
    map_rows.sort(key=lambda r: (r.project.lower(), r.task_description.lower()))
    return map_rows, ambiguous


def save_assignment_db(observations: List[TaskObservation], map_rows: List[AssignmentMapRow], ambiguous: List[Dict[str, Any]], output_path: Path) -> None:
    obs_df = pd.DataFrame([asdict(o) for o in observations])
    map_df = pd.DataFrame([asdict(r) for r in map_rows])
    ambiguous_df = pd.DataFrame(ambiguous)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        map_df.to_excel(writer, sheet_name="assignment_map", index=False)
        obs_df.to_excel(writer, sheet_name="observations", index=False)
        ambiguous_df.to_excel(writer, sheet_name="ambiguous_choices", index=False)


def load_assignment_map(mapping_path: Path) -> Dict[Tuple[str, str], AssignmentMapRow]:
    if not mapping_path.exists():
        raise FileNotFoundError(f"Mapping file not found: {mapping_path}")
    df = pd.read_excel(mapping_path, sheet_name="assignment_map")
    mapping: Dict[Tuple[str, str], AssignmentMapRow] = {}
    for _, row in df.iterrows():
        project = normalize_text(str(row.get("project", "")))
        desc = normalize_text(str(row.get("task_description", "")))
        assignee = normalize_text(str(row.get("recommended_assigned_to", "")))
        if not project or not desc or not assignee:
            continue
        item = AssignmentMapRow(
            project=project,
            task_description=desc,
            recommended_assigned_to=assignee,
            observation_count=int(row.get("observation_count", 0) or 0),
            unique_assignee_count=int(row.get("unique_assignee_count", 0) or 0),
            all_assignees_with_counts=normalize_text(str(row.get("all_assignees_with_counts", ""))),
            confidence=normalize_text(str(row.get("confidence", ""))),
        )
        mapping[(normalize_key(project), normalize_key(desc))] = item
    return mapping


def choose_assigned_to(assigned_select: WebElement, assignee_name: str) -> None:
    select = Select(assigned_select)
    target = normalize_text(assignee_name)
    try:
        select.select_by_visible_text(target)
        return
    except Exception:
        pass
    for option in select.options:
        if normalize_key(option.text) == normalize_key(target):
            option.click(); return
    for option in select.options:
        if normalize_key(target) in normalize_key(option.text) or normalize_key(option.text) in normalize_key(target):
            option.click(); return
    raise ValueError(f"Could not find assignee option '{assignee_name}' in Assigned to dropdown.")


def current_selected_assigned_to(driver: WebDriver) -> str:
    return selected_text(wait_for(driver, "task_assigned_to_select", timeout=20))


def handle_possible_alert(driver: WebDriver, log_func=None) -> str:
    try:
        alert = driver.switch_to.alert
        text = alert.text
        alert.accept()
        if log_func:
            log_func(f"Accepted alert: {text}")
        return text
    except NoAlertPresentException:
        return ""
    except WebDriverException:
        return ""


def open_task_edit_context(driver: WebDriver, edit_url: str, timeout: int = 20) -> None:
    driver.switch_to.default_content()
    driver.get(edit_url)
    WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")
    driver.switch_to.default_content()
    if safe_find(driver, "task_edit_form") is None:
        try:
            iframe = wait_for(driver, "tasks_iframe", timeout=5)
            driver.switch_to.frame(iframe)
        except TimeoutException:
            pass
    wait_for(driver, "task_edit_form", timeout=timeout)


def update_task_assignee(driver: WebDriver, obs: TaskObservation, recommended: str, timeout: int = 20, log_func=None) -> AssigneeUpdateResult:
    result = AssigneeUpdateResult(obs.ticket_number, obs.ticket_url, obs.project, obs.tasks_url,
                                  task_id=obs.task_id, edit_url=obs.edit_url, task_description=obs.task_description,
                                  previous_assigned_to=obs.assigned_to, recommended_assigned_to=recommended,
                                  status=obs.status, action="fill_missing_assigned_to")
    if not obs.edit_url:
        result.error = "No edit URL found for task."
        result.update_attempted = False
        result.update_success = False
        return result
    try:
        if log_func:
            log_func(f"    Opening task edit page for task {obs.task_id}...")
        open_task_edit_context(driver, obs.edit_url, timeout=timeout)
        choose_assigned_to(wait_for(driver, "task_assigned_to_select", timeout=timeout), recommended)
        result.update_attempted = True
        wait_for(driver, "task_update_button", timeout=timeout).click()
        alert_text = handle_possible_alert(driver, log_func=log_func)
        if alert_text:
            result.error = f"Alert after Update: {alert_text}"
        time.sleep(1.0)
        open_task_edit_context(driver, obs.edit_url, timeout=timeout)
        result.assigned_to_after_update = current_selected_assigned_to(driver)
        result.update_success = normalize_key(result.assigned_to_after_update) == normalize_key(recommended)
        if not result.update_success and not result.error:
            result.error = f"Assigned to after update is '{result.assigned_to_after_update}', expected '{recommended}'."
    except Exception as exc:
        result.update_success = False
        result.error = f"{type(exc).__name__}: {exc}"
    return result


def save_update_audit(results: List[AssigneeUpdateResult], output_path: Path) -> None:
    ordered_cols = ["ticket_number", "ticket_url", "project", "tasks_url", "task_id", "edit_url",
                    "task_description", "previous_assigned_to", "recommended_assigned_to", "status",
                    "action", "update_attempted", "update_success", "assigned_to_after_update", "error"]
    df = pd.DataFrame([asdict(r) for r in results])
    df = pd.DataFrame(columns=ordered_cols) if df.empty else df[ordered_cols]
    suffix = output_path.suffix.lower()
    if suffix == ".xlsx":
        df.to_excel(output_path, index=False)
    elif suffix == ".csv":
        df.to_csv(output_path, index=False)
    else:
        raise ValueError("Output must be .xlsx or .csv")


class Phase3GUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Phase 3 - Task Assignee Mapper and Updater")
        self.root.geometry("1040x820")
        self.worker_thread: Optional[threading.Thread] = None
        self.stop_event = threading.Event()
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.progress_queue: queue.Queue[Tuple[int, int, str]] = queue.Queue()
        self.base_url_var = tk.StringVar(value="YOUR_TICKET_SYSTEM_URL/edit_bug.aspx?id=")
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.user_data_dir_var = tk.StringVar()
        self.headless_var = tk.BooleanVar(value=False)
        self.pause_var = tk.DoubleVar(value=0.5)
        self.build_ticket_file_var = tk.StringVar()
        self.assignment_db_output_var = tk.StringVar(value=str(Path.cwd() / "phase3_assignment_db.xlsx"))
        self.update_ticket_file_var = tk.StringVar()
        self.assignment_db_input_var = tk.StringVar(value=str(Path.cwd() / "phase3_assignment_db.xlsx"))
        self.update_audit_output_var = tk.StringVar(value=str(Path.cwd() / "phase3_assignee_update_audit.xlsx"))
        self.dry_run_var = tk.BooleanVar(value=True)
        self.build_layout()
        self.root.after(200, self.process_queues)

    def build_layout(self) -> None:
        pad = {"padx": 8, "pady": 5}
        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        row = 0
        ttk.Label(frame, text="Shared settings", font=("Segoe UI", 11, "bold")).grid(row=row, column=0, columnspan=3, sticky="w", **pad)
        row += 1
        ttk.Label(frame, text="Base ticket URL:").grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.base_url_var, width=100).grid(row=row, column=1, sticky="ew", **pad)
        row += 1
        ttk.Label(frame, text="Username:").grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.username_var, width=40).grid(row=row, column=1, sticky="w", **pad)
        row += 1
        ttk.Label(frame, text="Password:").grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.password_var, width=40, show="*").grid(row=row, column=1, sticky="w", **pad)
        row += 1
        ttk.Label(frame, text="Chrome user data dir optional:").grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.user_data_dir_var, width=100).grid(row=row, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Browse", command=self.choose_user_data_dir).grid(row=row, column=2, **pad)
        row += 1
        options = ttk.Frame(frame); options.grid(row=row, column=1, sticky="w", **pad)
        ttk.Checkbutton(options, text="Headless mode not recommended", variable=self.headless_var).pack(side="left")
        ttk.Label(options, text="Pause seconds:").pack(side="left", padx=(20, 5))
        ttk.Spinbox(options, textvariable=self.pause_var, from_=0.0, to=10.0, increment=0.25, width=6).pack(side="left")
        row += 1
        ttk.Separator(frame, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1
        ttk.Label(frame, text="A) Build project-based assignment DB", font=("Segoe UI", 11, "bold")).grid(row=row, column=0, columnspan=3, sticky="w", **pad)
        row += 1
        ttk.Label(frame, text="Training ticket list:").grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.build_ticket_file_var, width=100).grid(row=row, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Browse", command=self.choose_build_ticket_file).grid(row=row, column=2, **pad)
        row += 1
        ttk.Label(frame, text="Save assignment DB:").grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.assignment_db_output_var, width=100).grid(row=row, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Save As", command=self.choose_assignment_db_output).grid(row=row, column=2, **pad)
        row += 1
        ttk.Button(frame, text="Build Assignment DB", command=self.start_build_db).grid(row=row, column=1, sticky="w", **pad)
        row += 1
        ttk.Separator(frame, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1
        ttk.Label(frame, text="B) Update missing Assigned To using DB", font=("Segoe UI", 11, "bold")).grid(row=row, column=0, columnspan=3, sticky="w", **pad)
        row += 1
        ttk.Label(frame, text="Update ticket list:").grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.update_ticket_file_var, width=100).grid(row=row, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Browse", command=self.choose_update_ticket_file).grid(row=row, column=2, **pad)
        row += 1
        ttk.Label(frame, text="Load assignment DB:").grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.assignment_db_input_var, width=100).grid(row=row, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Browse", command=self.choose_assignment_db_input).grid(row=row, column=2, **pad)
        row += 1
        ttk.Label(frame, text="Save update audit:").grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.update_audit_output_var, width=100).grid(row=row, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Save As", command=self.choose_update_audit_output).grid(row=row, column=2, **pad)
        row += 1
        update_options = ttk.Frame(frame); update_options.grid(row=row, column=1, sticky="w", **pad)
        ttk.Checkbutton(update_options, text="Dry run only - do not update", variable=self.dry_run_var).pack(side="left")
        ttk.Button(update_options, text="Update Missing Assigned To", command=self.start_update_missing).pack(side="left", padx=(20, 0))
        self.stop_button = ttk.Button(update_options, text="Stop after current ticket", command=self.stop, state="disabled")
        self.stop_button.pack(side="left", padx=(8, 0))
        row += 1
        self.progress = ttk.Progressbar(frame, orient="horizontal", mode="determinate")
        self.progress.grid(row=row, column=0, columnspan=3, sticky="ew", padx=8, pady=(12, 5))
        row += 1
        self.status_label = ttk.Label(frame, text="Ready")
        self.status_label.grid(row=row, column=0, columnspan=3, sticky="w", padx=8, pady=5)
        row += 1
        ttk.Label(frame, text="Live progress log:").grid(row=row, column=0, columnspan=3, sticky="w", padx=8, pady=(12, 5))
        row += 1
        self.log_text = tk.Text(frame, height=18, wrap="word")
        self.log_text.grid(row=row, column=0, columnspan=3, sticky="nsew", padx=8, pady=5)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=row, column=3, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(row, weight=1)

    def choose_file(self, title: str, variable: tk.StringVar, filetypes=None) -> None:
        if filetypes is None:
            filetypes = [("Ticket files", "*.txt *.csv *.xlsx *.xls"), ("All files", "*.*")]
        filename = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if filename:
            variable.set(filename)

    def save_file(self, title: str, variable: tk.StringVar, default_ext: str = ".xlsx") -> None:
        filename = filedialog.asksaveasfilename(title=title, defaultextension=default_ext, filetypes=[("Excel file", "*.xlsx"), ("CSV file", "*.csv"), ("All files", "*.*")])
        if filename:
            variable.set(filename)

    def choose_user_data_dir(self) -> None:
        dirname = filedialog.askdirectory(title="Choose Chrome user data directory")
        if dirname:
            self.user_data_dir_var.set(dirname)

    def choose_build_ticket_file(self) -> None: self.choose_file("Choose training ticket list", self.build_ticket_file_var)
    def choose_assignment_db_output(self) -> None:
        self.save_file("Save assignment DB", self.assignment_db_output_var, ".xlsx")
        self.assignment_db_input_var.set(self.assignment_db_output_var.get())
    def choose_update_ticket_file(self) -> None: self.choose_file("Choose update ticket list", self.update_ticket_file_var)
    def choose_assignment_db_input(self) -> None: self.choose_file("Choose assignment DB", self.assignment_db_input_var, [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
    def choose_update_audit_output(self) -> None: self.save_file("Save update audit", self.update_audit_output_var, ".xlsx")
    def log(self, message: str) -> None: self.log_queue.put(message)
    def stop(self) -> None:
        self.stop_event.set(); self.log("Stop requested. The script will stop after the current ticket finishes."); self.stop_button.configure(state="disabled")

    def validate_shared(self) -> bool:
        if not self.base_url_var.get().strip():
            messagebox.showerror("Missing base URL", "Enter the base ticket URL."); return False
        return True

    def start_worker(self, target) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showerror("Already running", "A process is already running."); return
        self.stop_event.clear(); self.stop_button.configure(state="normal")
        self.progress.configure(value=0, maximum=1); self.status_label.configure(text="Starting...")
        self.log_text.delete("1.0", "end")
        self.worker_thread = threading.Thread(target=target, daemon=True); self.worker_thread.start()

    def start_build_db(self) -> None:
        if not self.validate_shared(): return
        if not self.build_ticket_file_var.get().strip(): messagebox.showerror("Missing training ticket list", "Choose the training ticket list for Part A."); return
        if not self.assignment_db_output_var.get().strip(): messagebox.showerror("Missing DB output", "Choose where to save the assignment DB."); return
        self.start_worker(self.run_build_db)

    def start_update_missing(self) -> None:
        if not self.validate_shared(): return
        if not self.update_ticket_file_var.get().strip(): messagebox.showerror("Missing update ticket list", "Choose the update ticket list for Part B."); return
        if not self.assignment_db_input_var.get().strip(): messagebox.showerror("Missing assignment DB", "Choose the assignment DB from Part A."); return
        if not self.update_audit_output_var.get().strip(): messagebox.showerror("Missing audit output", "Choose where to save the update audit."); return
        self.start_worker(self.run_update_missing)

    def process_queues(self) -> None:
        try:
            while True:
                message = self.log_queue.get_nowait(); self.log_text.insert("end", message + "\n"); self.log_text.see("end")
        except queue.Empty:
            pass
        try:
            while True:
                current, total, label = self.progress_queue.get_nowait()
                self.progress.configure(maximum=max(total, 1), value=current); self.status_label.configure(text=f"{label}: {current}/{total}")
        except queue.Empty:
            pass
        if self.worker_thread and not self.worker_thread.is_alive():
            self.stop_button.configure(state="disabled"); self.worker_thread = None
        self.root.after(200, self.process_queues)

    def make_driver_from_settings(self) -> WebDriver:
        return make_driver(headless=self.headless_var.get(), user_data_dir=self.user_data_dir_var.get().strip())

    def run_build_db(self) -> None:
        base_url = self.base_url_var.get().strip(); tickets_path = Path(self.build_ticket_file_var.get().strip()); output_path = Path(self.assignment_db_output_var.get().strip())
        username = self.username_var.get(); password = self.password_var.get(); pause_seconds = float(self.pause_var.get())
        driver: Optional[WebDriver] = None; observations: List[TaskObservation] = []
        try:
            ticket_numbers = read_ticket_numbers(tickets_path); self.progress_queue.put((0, len(ticket_numbers), "Build DB")); self.log(f"Building DB from {len(ticket_numbers)} training tickets.")
            driver = self.make_driver_from_settings()
            for index, ticket_number in enumerate(ticket_numbers, start=1):
                if self.stop_event.is_set(): self.log("Stopped before starting next training ticket."); break
                ticket_url = build_ticket_url(base_url, ticket_number); self.log(f"[Build {index}/{len(ticket_numbers)}] Ticket {ticket_number}")
                _, _, error = open_ticket_and_login_if_needed(driver, ticket_url, username, password, log_func=self.log)
                if error:
                    self.log(f"  Ticket error: {error}"); self.progress_queue.put((index, len(ticket_numbers), "Build DB")); continue
                try:
                    project = read_project(driver); _, tasks_url = read_tasks_link(driver, ticket_url)
                    if not tasks_url:
                        self.log("  Tasks/Time link not found."); self.progress_queue.put((index, len(ticket_numbers), "Build DB")); continue
                    ticket_obs = open_tasks_page_and_get_observations(driver, ticket_number, ticket_url, project, tasks_url)
                    observations.extend(ticket_obs); usable = [o for o in ticket_obs if not is_missing_assignee(o.assigned_to)]
                    self.log(f"  Project='{project}'. Tasks found={len(ticket_obs)}, usable assigned rows={len(usable)}.")
                except Exception as exc:
                    self.log(f"  Error reading ticket/tasks: {type(exc).__name__}: {exc}")
                self.progress_queue.put((index, len(ticket_numbers), "Build DB"))
                if pause_seconds > 0: time.sleep(pause_seconds)
            map_rows, ambiguous = build_assignment_map(observations); save_assignment_db(observations, map_rows, ambiguous, output_path)
            self.log(f"Assignment DB saved to: {output_path}"); self.log(f"Observations: {len(observations)}"); self.log(f"Project/task mappings: {len(map_rows)}"); self.log(f"Ambiguous mappings: {len(ambiguous)}"); self.log("Build DB finished.")
        except Exception as exc:
            self.log(f"Fatal Build DB error: {type(exc).__name__}: {exc}")
        finally:
            if driver is not None:
                try: driver.quit()
                except Exception: pass

    def run_update_missing(self) -> None:
        base_url = self.base_url_var.get().strip(); tickets_path = Path(self.update_ticket_file_var.get().strip()); mapping_path = Path(self.assignment_db_input_var.get().strip()); output_path = Path(self.update_audit_output_var.get().strip())
        username = self.username_var.get(); password = self.password_var.get(); dry_run = self.dry_run_var.get(); pause_seconds = float(self.pause_var.get())
        driver: Optional[WebDriver] = None; results: List[AssigneeUpdateResult] = []
        try:
            ticket_numbers = read_ticket_numbers(tickets_path); mapping = load_assignment_map(mapping_path)
            self.progress_queue.put((0, len(ticket_numbers), "Update Missing Assigned To")); self.log(f"Loaded {len(mapping)} assignment mappings from: {mapping_path}"); self.log(f"Processing {len(ticket_numbers)} update tickets.")
            if dry_run: self.log("DRY RUN is ON. No updates will be submitted.")
            driver = self.make_driver_from_settings()
            for index, ticket_number in enumerate(ticket_numbers, start=1):
                if self.stop_event.is_set(): self.log("Stopped before starting next update ticket."); break
                ticket_url = build_ticket_url(base_url, ticket_number); self.log(f"[Update {index}/{len(ticket_numbers)}] Ticket {ticket_number}")
                _, _, error = open_ticket_and_login_if_needed(driver, ticket_url, username, password, log_func=self.log)
                if error:
                    self.log(f"  Ticket error: {error}"); results.append(AssigneeUpdateResult(ticket_number, ticket_url, "", "", action="ticket_open_failed", update_attempted=False, update_success=False, error=error)); save_update_audit(results, output_path); self.progress_queue.put((index, len(ticket_numbers), "Update Missing Assigned To")); continue
                try:
                    project = read_project(driver); _, tasks_url = read_tasks_link(driver, ticket_url)
                    if not tasks_url:
                        self.log("  Tasks/Time link not found."); results.append(AssigneeUpdateResult(ticket_number, ticket_url, project, "", action="tasks_link_not_found", update_attempted=False, update_success=False, error="Tasks/Time link not found.")); save_update_audit(results, output_path); self.progress_queue.put((index, len(ticket_numbers), "Update Missing Assigned To")); continue
                    observations = open_tasks_page_and_get_observations(driver, ticket_number, ticket_url, project, tasks_url); missing = [o for o in observations if is_missing_assignee(o.assigned_to)]
                    self.log(f"  Project='{project}'. Tasks found={len(observations)}, missing assigned-to={len(missing)}.")
                    if not missing:
                        results.append(AssigneeUpdateResult(ticket_number, ticket_url, project, tasks_url, action="no_missing_assigned_to", update_attempted=False, update_success=True)); save_update_audit(results, output_path); self.progress_queue.put((index, len(ticket_numbers), "Update Missing Assigned To")); continue
                    for obs in missing:
                        if self.stop_event.is_set(): self.log("  Stop requested. Finishing current ticket loop."); break
                        key = (normalize_key(obs.project), normalize_key(obs.task_description)); map_row = mapping.get(key)
                        if map_row is None:
                            self.log(f"  No mapping found for task '{obs.task_description}' in project '{obs.project}'.")
                            results.append(AssigneeUpdateResult(obs.ticket_number, obs.ticket_url, obs.project, obs.tasks_url, task_id=obs.task_id, edit_url=obs.edit_url, task_description=obs.task_description, previous_assigned_to=obs.assigned_to, status=obs.status, action="no_mapping_found", update_attempted=False, update_success=False, error="No assignment mapping found for Project + Task Description.")); save_update_audit(results, output_path); continue
                        recommended = map_row.recommended_assigned_to; self.log(f"  Task {obs.task_id}: '{obs.task_description}' -> assigning '{recommended}' (confidence={map_row.confidence})")
                        if dry_run:
                            results.append(AssigneeUpdateResult(obs.ticket_number, obs.ticket_url, obs.project, obs.tasks_url, task_id=obs.task_id, edit_url=obs.edit_url, task_description=obs.task_description, previous_assigned_to=obs.assigned_to, recommended_assigned_to=recommended, status=obs.status, action="dry_run_would_fill_assigned_to", update_attempted=False, update_success=None, error="Dry run only. No update submitted.")); save_update_audit(results, output_path); continue
                        result = update_task_assignee(driver, obs, recommended, log_func=self.log); results.append(result); save_update_audit(results, output_path)
                        self.log(f"    Task {obs.task_id} {'updated successfully.' if result.update_success else 'update failed: ' + result.error}")
                        if pause_seconds > 0: time.sleep(pause_seconds)
                except Exception as exc:
                    self.log(f"  Error processing ticket: {type(exc).__name__}: {exc}")
                self.progress_queue.put((index, len(ticket_numbers), "Update Missing Assigned To"))
            self.log(f"Update audit saved to: {output_path}"); self.log("Update Missing Assigned To finished.")
        except Exception as exc:
            self.log(f"Fatal Update error: {type(exc).__name__}: {exc}")
            try: save_update_audit(results, output_path)
            except Exception: pass
        finally:
            if driver is not None:
                try: driver.quit()
                except Exception: pass


def cli_main() -> int:
    parser = argparse.ArgumentParser(description="Launch Phase 3 Task Assignee Mapper and Updater GUI.")
    parser.parse_args()
    root = tk.Tk(); Phase3GUI(root); root.mainloop(); return 0


if __name__ == "__main__":
    raise SystemExit(cli_main())
