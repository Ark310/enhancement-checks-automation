from __future__ import annotations

import argparse
import csv
import queue
import sys
import threading
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional, List, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver, WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


SELECTORS = {
    # Login page selectors.
    "login_username": (By.CSS_SELECTOR, "input#user"),
    "login_password": (By.CSS_SELECTOR, "input#pw"),
    "login_submit": (By.CSS_SELECTOR, "input[name='ctl01']"),

    # Ticket page selectors.
    "ticket_page_ready": (By.CSS_SELECTOR, "form#ctl00"),
    "sqa_assign_to_row": (By.CSS_SELECTOR, "#SQAAssignTo_row"),
    "khi_qa_signoff_row": (By.CSS_SELECTOR, "#KHIQASignoff_row"),
    "tor_qa_signoff_row": (By.CSS_SELECTOR, "#TORQASignoff_row"),
    "csqa_owner_row": (By.CSS_SELECTOR, "#CSQAOwner_row"),
    "scope_approval_count_hidden": (By.CSS_SELECTOR, "#ScopeApprovalCount"),
    "scope_approval_link": (By.CSS_SELECTOR, "li#scope_approval a"),
    "tasks_link": (By.CSS_SELECTOR, "li#tasks a"),

    # Scope Approval page selectors.
    "scope_approval_ready": (By.CSS_SELECTOR, "form#scopeForm table#grdViewScopeDtl"),
    "scope_approval_table": (By.CSS_SELECTOR, "table#grdViewScopeDtl"),

    # Tasks/Time page selectors.
    "tasks_iframe": (By.CSS_SELECTOR, "iframe"),
    "tasks_table": (By.CSS_SELECTOR, "table#mytable"),
}


@dataclass
class TicketCheckResult:
    ticket_number: str
    ticket_url: str
    opened: bool = False
    login_attempted: bool = False
    login_success: Optional[bool] = None
    error: str = ""

    sqa_assign_to: str = ""
    khi_qa_signoff: str = ""
    tor_qa_signoff: str = ""
    csqa_owner: str = ""

    scope_approval_menu_text: str = ""
    scope_approval_hidden_count: str = ""
    scope_approval_clicked: bool = False
    scope_approval_url: str = ""
    scope_approval_record_count: Optional[int] = None
    scope_approval_has_records: Optional[bool] = None
    scope_approval_has_enhancement: Optional[bool] = None
    scope_approval_request_types: str = ""

    tasks_menu_text: str = ""
    tasks_clicked: bool = False
    tasks_url: str = ""
    tasks_record_count: Optional[int] = None
    tasks_closed_count: Optional[int] = None
    tasks_not_closed_count: Optional[int] = None
    tasks_all_closed: Optional[bool] = None
    tasks_statuses: str = ""
    tasks_details: str = ""


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
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, selector)))


def safe_find(driver: WebDriver, selector_key: str) -> Optional[WebElement]:
    try:
        by, selector = SELECTORS[selector_key]
        return driver.find_element(by, selector)
    except NoSuchElementException:
        return None


def is_login_page(driver: WebDriver) -> bool:
    return safe_find(driver, "login_username") is not None and safe_find(driver, "login_password") is not None


def login_if_needed(driver: WebDriver, username: str, password: str, log_func=None, timeout: int = 30) -> Tuple[bool, bool]:
    """
    Returns:
      (login_attempted, login_success)

    It only logs in if the current page has the login fields.
    """
    if not is_login_page(driver):
        return False, True

    if log_func:
        log_func("Login page detected. Attempting login...")

    if not username or not password:
        return True, False

    user_el = wait_for(driver, "login_username", timeout=timeout)
    pw_el = wait_for(driver, "login_password", timeout=timeout)

    user_el.clear()
    user_el.send_keys(username)

    pw_el.clear()
    pw_el.send_keys(password)

    submit_el = wait_for(driver, "login_submit", timeout=timeout)
    submit_el.click()

    try:
        WebDriverWait(driver, timeout).until(
            lambda d: safe_find(d, "ticket_page_ready") is not None or not is_login_page(d)
        )
    except TimeoutException:
        pass

    success = not is_login_page(driver)
    if log_func:
        if success:
            log_func("Login appears successful.")
        else:
            log_func("Login still appears to be on the login page. Check username/password or page behavior.")
    return True, success


def selected_text_from_select(select_el: WebElement) -> str:
    try:
        return Select(select_el).first_selected_option.text.strip()
    except Exception:
        return ""


def first_nonempty_value_in_row(row: WebElement) -> str:
    for select_el in row.find_elements(By.CSS_SELECTOR, "select"):
        value = selected_text_from_select(select_el)
        if value:
            return value

    for input_el in row.find_elements(By.CSS_SELECTOR, "input"):
        input_type = (input_el.get_attribute("type") or "").lower()
        if input_type in {"hidden", "button", "submit", "reset", "checkbox", "radio"}:
            continue
        value = (input_el.get_attribute("value") or "").strip()
        if value:
            return value

    for textarea in row.find_elements(By.CSS_SELECTOR, "textarea"):
        value = (textarea.get_attribute("value") or textarea.text or "").strip()
        if value:
            return value

    cells = row.find_elements(By.CSS_SELECTOR, "td")
    if len(cells) >= 2:
        return cells[-1].text.strip()

    return row.text.strip()


def read_field_from_row(driver: WebDriver, selector_key: str) -> str:
    row = safe_find(driver, selector_key)
    if row is None:
        return ""
    return first_nonempty_value_in_row(row)


def read_scope_hidden_count(driver: WebDriver) -> str:
    el = safe_find(driver, "scope_approval_count_hidden")
    if el is None:
        return ""
    return (el.get_attribute("value") or "").strip()


def read_link_text_and_url(driver: WebDriver, selector_key: str) -> Tuple[str, str]:
    el = safe_find(driver, selector_key)
    if el is None:
        return "", ""
    return el.text.strip(), (el.get_attribute("href") or "").strip()


def parse_scope_approval_table(driver: WebDriver) -> Tuple[int, bool, List[str]]:
    table = wait_for(driver, "scope_approval_table", timeout=20)
    rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")

    data_rows: List[WebElement] = []
    request_types: List[str] = []

    for row in rows:
        cells = row.find_elements(By.CSS_SELECTOR, "td")
        if len(cells) < 2:
            continue
        first_cell = cells[0].text.strip()
        request_type = cells[1].text.strip()
        if first_cell.isdigit():
            data_rows.append(row)
            request_types.append(request_type)

    has_enhancement = any(value.strip().lower() == "enhancement" for value in request_types)
    return len(data_rows), has_enhancement, request_types


def check_scope_approval(driver: WebDriver, result: TicketCheckResult, timeout: int = 20) -> None:
    menu_text, href = read_link_text_and_url(driver, "scope_approval_link")
    result.scope_approval_menu_text = menu_text
    result.scope_approval_url = href
    result.scope_approval_hidden_count = read_scope_hidden_count(driver)

    link = safe_find(driver, "scope_approval_link")
    if link is None:
        result.error += " Scope Approval link not found."
        return

    original_window = driver.current_window_handle
    windows_before = set(driver.window_handles)

    try:
        WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(SELECTORS["scope_approval_link"]))
        link.click()
        result.scope_approval_clicked = True

        try:
            WebDriverWait(driver, 5).until(lambda d: len(set(d.window_handles) - windows_before) > 0)
            new_windows = list(set(driver.window_handles) - windows_before)
            if new_windows:
                driver.switch_to.window(new_windows[0])
        except TimeoutException:
            pass

        WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")
        wait_for(driver, "scope_approval_ready", timeout=timeout)
        time.sleep(0.25)

        count, has_enhancement, request_types = parse_scope_approval_table(driver)
        result.scope_approval_record_count = count
        result.scope_approval_has_records = count > 0
        result.scope_approval_has_enhancement = has_enhancement
        result.scope_approval_request_types = ", ".join(request_types)

    except TimeoutException as exc:
        result.error += f" Scope Approval timeout: {exc}"
        result.scope_approval_record_count = None
        result.scope_approval_has_records = None
        result.scope_approval_has_enhancement = None
    except WebDriverException as exc:
        result.error += f" Scope Approval browser error: {exc}"
        result.scope_approval_record_count = None
        result.scope_approval_has_records = None
        result.scope_approval_has_enhancement = None
    finally:
        try:
            if driver.current_window_handle != original_window and original_window in driver.window_handles:
                driver.close()
                driver.switch_to.window(original_window)
        except WebDriverException:
            pass


def switch_into_tasks_table_context(driver: WebDriver, timeout: int = 20) -> None:
    """
    The Tasks/Time page in your screenshot contains an iframe whose src is tasks.aspx?bugid=...
    The table#mytable is inside that iframe. This function switches into the iframe if needed.
    """
    driver.switch_to.default_content()

    if safe_find(driver, "tasks_table") is not None:
        return

    iframe = wait_for(driver, "tasks_iframe", timeout=timeout)
    driver.switch_to.frame(iframe)
    wait_for(driver, "tasks_table", timeout=timeout)


def parse_tasks_table(driver: WebDriver) -> Tuple[int, int, int, bool, List[str], List[str]]:
    """
    Parse table#mytable from the Tasks/Time page.

    Based on your screenshot, columns appear to be:
      id, edit, delete, description, assigned to, actual end, actual duration,
      duration units, status

    The last column is status. Example status: Closed.
    """
    table = wait_for(driver, "tasks_table", timeout=20)
    rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")

    task_count = 0
    closed_count = 0
    not_closed_count = 0
    statuses: List[str] = []
    details: List[str] = []

    for row in rows:
        cells = row.find_elements(By.CSS_SELECTOR, "td")
        # Header row has th, not td. A valid data row has at least id, edit, delete, description, assigned to, ..., status.
        if len(cells) < 4:
            continue

        task_id = cells[0].text.strip()
        if not task_id.isdigit():
            continue

        description = cells[3].text.strip() if len(cells) > 3 else ""
        assigned_to = cells[4].text.strip() if len(cells) > 4 else ""
        status = cells[-1].text.strip() if cells else ""

        task_count += 1
        statuses.append(status)
        details.append(f"{task_id} | {description} | assigned to: {assigned_to} | status: {status}")

        if status.strip().lower() == "closed":
            closed_count += 1
        else:
            not_closed_count += 1

    all_closed = task_count > 0 and not_closed_count == 0
    return task_count, closed_count, not_closed_count, all_closed, statuses, details


def check_tasks_time(driver: WebDriver, result: TicketCheckResult, timeout: int = 20) -> None:
    """
    Open Tasks/Time and check whether every task is Closed.
    """
    menu_text, href = read_link_text_and_url(driver, "tasks_link")
    result.tasks_menu_text = menu_text
    result.tasks_url = href

    link = safe_find(driver, "tasks_link")
    if link is None:
        result.error += " Tasks/Time link not found."
        return

    original_window = driver.current_window_handle
    windows_before = set(driver.window_handles)

    try:
        WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(SELECTORS["tasks_link"]))
        link.click()
        result.tasks_clicked = True

        try:
            WebDriverWait(driver, 5).until(lambda d: len(set(d.window_handles) - windows_before) > 0)
            new_windows = list(set(driver.window_handles) - windows_before)
            if new_windows:
                driver.switch_to.window(new_windows[0])
        except TimeoutException:
            pass

        WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")
        switch_into_tasks_table_context(driver, timeout=timeout)
        time.sleep(0.25)

        count, closed_count, not_closed_count, all_closed, statuses, details = parse_tasks_table(driver)
        result.tasks_record_count = count
        result.tasks_closed_count = closed_count
        result.tasks_not_closed_count = not_closed_count
        result.tasks_all_closed = all_closed
        result.tasks_statuses = ", ".join(statuses)
        result.tasks_details = " ; ".join(details)

    except TimeoutException as exc:
        result.error += f" Tasks/Time timeout: {exc}"
        result.tasks_record_count = None
        result.tasks_closed_count = None
        result.tasks_not_closed_count = None
        result.tasks_all_closed = None
    except WebDriverException as exc:
        result.error += f" Tasks/Time browser error: {exc}"
        result.tasks_record_count = None
        result.tasks_closed_count = None
        result.tasks_not_closed_count = None
        result.tasks_all_closed = None
    finally:
        try:
            driver.switch_to.default_content()
            if driver.current_window_handle != original_window and original_window in driver.window_handles:
                driver.close()
                driver.switch_to.window(original_window)
        except WebDriverException:
            pass


def check_ticket(
    driver: WebDriver,
    base_url: str,
    ticket_number: str,
    username: str,
    password: str,
    pause_seconds: float = 0.5,
    log_func=None,
) -> TicketCheckResult:
    url = build_ticket_url(base_url, ticket_number)
    result = TicketCheckResult(ticket_number=ticket_number, ticket_url=url)

    try:
        driver.get(url)
        result.opened = True

        attempted, success = login_if_needed(driver, username, password, log_func=log_func)
        result.login_attempted = attempted
        result.login_success = success

        if attempted and success:
            if log_func:
                log_func(f"Reloading ticket after login: {ticket_number}")
            driver.get(url)

        if attempted and not success:
            result.error = "Login failed or login page did not clear."
            return result

        wait_for(driver, "ticket_page_ready", timeout=30)
        WebDriverWait(driver, 30).until(lambda d: d.execute_script("return document.readyState") == "complete")

        if pause_seconds > 0:
            time.sleep(pause_seconds)

        result.sqa_assign_to = read_field_from_row(driver, "sqa_assign_to_row")
        result.khi_qa_signoff = read_field_from_row(driver, "khi_qa_signoff_row")
        result.tor_qa_signoff = read_field_from_row(driver, "tor_qa_signoff_row")
        result.csqa_owner = read_field_from_row(driver, "csqa_owner_row")

        check_scope_approval(driver, result)

        # Make sure we are back on the ticket page before checking tasks/time.
        try:
            driver.switch_to.default_content()
        except WebDriverException:
            pass

        check_tasks_time(driver, result)

    except TimeoutException as exc:
        result.error = f"Ticket page timeout: {exc}"
    except WebDriverException as exc:
        result.error = f"Browser error: {exc}"
    except Exception as exc:
        result.error = f"Unexpected error: {type(exc).__name__}: {exc}"

    return result


def save_results(results: List[TicketCheckResult], output_path: Path) -> None:
    df = pd.DataFrame([asdict(r) for r in results])
    ordered_cols = [
        "ticket_number",
        "ticket_url",
        "opened",
        "login_attempted",
        "login_success",
        "sqa_assign_to",
        "khi_qa_signoff",
        "tor_qa_signoff",
        "csqa_owner",
        "scope_approval_menu_text",
        "scope_approval_hidden_count",
        "scope_approval_clicked",
        "scope_approval_url",
        "scope_approval_record_count",
        "scope_approval_has_records",
        "scope_approval_has_enhancement",
        "scope_approval_request_types",
        "tasks_menu_text",
        "tasks_clicked",
        "tasks_url",
        "tasks_record_count",
        "tasks_closed_count",
        "tasks_not_closed_count",
        "tasks_all_closed",
        "tasks_statuses",
        "tasks_details",
        "error",
    ]
    df = df[ordered_cols]

    suffix = output_path.suffix.lower()
    if suffix == ".xlsx":
        df.to_excel(output_path, index=False)
    elif suffix == ".csv":
        df.to_csv(output_path, index=False)
    else:
        raise ValueError("Output must be .xlsx or .csv")


class TicketCheckerGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Ticket Batch Checker")
        self.root.geometry("980x740")

        self.worker_thread: Optional[threading.Thread] = None
        self.stop_event = threading.Event()
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.progress_queue: queue.Queue[Tuple[int, int]] = queue.Queue()

        self.base_url_var = tk.StringVar(value="YOUR_TICKET_SYSTEM_URL/edit_bug.aspx?id=")
        self.ticket_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar(value=str(Path.cwd() / "ticket_check_results.xlsx"))
        self.user_data_dir_var = tk.StringVar()
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.headless_var = tk.BooleanVar(value=False)
        self.pause_var = tk.DoubleVar(value=0.5)

        self.build_layout()
        self.root.after(200, self.process_queues)

    def build_layout(self) -> None:
        pad = {"padx": 8, "pady": 5}

        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(frame, text="Base ticket URL:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.base_url_var, width=95).grid(row=0, column=1, sticky="ew", **pad)

        ttk.Label(frame, text="Ticket list file:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.ticket_file_var, width=95).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Browse", command=self.choose_ticket_file).grid(row=1, column=2, **pad)

        ttk.Label(frame, text="Output file:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.output_file_var, width=95).grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Save As", command=self.choose_output_file).grid(row=2, column=2, **pad)

        ttk.Label(frame, text="Username:").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.username_var, width=40).grid(row=3, column=1, sticky="w", **pad)

        ttk.Label(frame, text="Password:").grid(row=4, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.password_var, width=40, show="*").grid(row=4, column=1, sticky="w", **pad)

        ttk.Label(frame, text="Chrome user data dir optional:").grid(row=5, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.user_data_dir_var, width=95).grid(row=5, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Browse", command=self.choose_user_data_dir).grid(row=5, column=2, **pad)

        options_frame = ttk.Frame(frame)
        options_frame.grid(row=6, column=1, sticky="w", **pad)
        ttk.Checkbutton(options_frame, text="Headless mode not recommended for first run", variable=self.headless_var).pack(side="left")
        ttk.Label(options_frame, text="Pause seconds:").pack(side="left", padx=(20, 5))
        ttk.Spinbox(options_frame, textvariable=self.pause_var, from_=0.0, to=10.0, increment=0.25, width=6).pack(side="left")

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=7, column=1, sticky="w", **pad)
        self.start_button = ttk.Button(button_frame, text="Start", command=self.start)
        self.start_button.pack(side="left", padx=(0, 8))
        self.stop_button = ttk.Button(button_frame, text="Stop after current ticket", command=self.stop, state="disabled")
        self.stop_button.pack(side="left")

        self.progress = ttk.Progressbar(frame, orient="horizontal", mode="determinate")
        self.progress.grid(row=8, column=0, columnspan=3, sticky="ew", padx=8, pady=(12, 5))
        self.status_label = ttk.Label(frame, text="Ready")
        self.status_label.grid(row=9, column=0, columnspan=3, sticky="w", padx=8, pady=5)

        ttk.Label(frame, text="Live progress log:").grid(row=10, column=0, columnspan=3, sticky="w", padx=8, pady=(12, 5))
        self.log_text = tk.Text(frame, height=22, wrap="word")
        self.log_text.grid(row=11, column=0, columnspan=3, sticky="nsew", padx=8, pady=5)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=11, column=3, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        note = (
            "Checks included: ticket fields, Scope Approval Enhancement record, and Tasks/Time all-closed status. "
            "Password is only used in memory during this run and is not saved by this script."
        )
        ttk.Label(frame, text=note, foreground="gray").grid(row=12, column=0, columnspan=3, sticky="w", padx=8, pady=5)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(11, weight=1)

    def choose_ticket_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="Choose ticket list",
            filetypes=[("Ticket files", "*.txt *.csv *.xlsx *.xls"), ("All files", "*.*")],
        )
        if filename:
            self.ticket_file_var.set(filename)

    def choose_output_file(self) -> None:
        filename = filedialog.asksaveasfilename(
            title="Save results as",
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx"), ("CSV file", "*.csv")],
        )
        if filename:
            self.output_file_var.set(filename)

    def choose_user_data_dir(self) -> None:
        dirname = filedialog.askdirectory(title="Choose Chrome user data directory")
        if dirname:
            self.user_data_dir_var.set(dirname)

    def log(self, message: str) -> None:
        self.log_queue.put(message)

    def start(self) -> None:
        base_url = self.base_url_var.get().strip()
        ticket_file = self.ticket_file_var.get().strip()
        output_file = self.output_file_var.get().strip()

        if not base_url:
            messagebox.showerror("Missing base URL", "Enter the base ticket URL.")
            return
        if not ticket_file:
            messagebox.showerror("Missing ticket list", "Choose a TXT, CSV, or XLSX ticket list file.")
            return
        if not output_file:
            messagebox.showerror("Missing output file", "Choose where to save the results.")
            return

        self.stop_event.clear()
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.progress.configure(value=0, maximum=1)
        self.status_label.configure(text="Starting...")
        self.log_text.delete("1.0", "end")

        self.worker_thread = threading.Thread(target=self.run_checker, daemon=True)
        self.worker_thread.start()

    def stop(self) -> None:
        self.stop_event.set()
        self.log("Stop requested. The script will stop after the current ticket finishes.")
        self.stop_button.configure(state="disabled")

    def process_queues(self) -> None:
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.insert("end", message + "\n")
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

        if self.worker_thread and not self.worker_thread.is_alive():
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.worker_thread = None

        self.root.after(200, self.process_queues)

    def run_checker(self) -> None:
        base_url = self.base_url_var.get().strip()
        tickets_path = Path(self.ticket_file_var.get().strip())
        output_path = Path(self.output_file_var.get().strip())
        headless = self.headless_var.get()
        pause_seconds = float(self.pause_var.get())
        user_data_dir = self.user_data_dir_var.get().strip()
        username = self.username_var.get()
        password = self.password_var.get()

        driver: Optional[WebDriver] = None
        results: List[TicketCheckResult] = []

        try:
            ticket_numbers = read_ticket_numbers(tickets_path)
            if not ticket_numbers:
                self.log("No ticket numbers found in the selected file.")
                return

            self.progress_queue.put((0, len(ticket_numbers)))
            self.log(f"Loaded {len(ticket_numbers)} tickets.")
            self.log("Opening Chrome. Login will be attempted only if the login page is detected.")

            driver = make_driver(headless=headless, user_data_dir=user_data_dir)

            for index, ticket_number in enumerate(ticket_numbers, start=1):
                if self.stop_event.is_set():
                    self.log("Stopped before starting next ticket.")
                    break

                self.log(f"[{index}/{len(ticket_numbers)}] Checking ticket {ticket_number}...")
                result = check_ticket(
                    driver=driver,
                    base_url=base_url,
                    ticket_number=ticket_number,
                    username=username,
                    password=password,
                    pause_seconds=pause_seconds,
                    log_func=self.log,
                )
                results.append(result)
                save_results(results, output_path)

                if result.error:
                    self.log(f"  Ticket {ticket_number} completed with error: {result.error}")
                else:
                    self.log(
                        f"  Done. SQA='{result.sqa_assign_to}', KHI='{result.khi_qa_signoff}', "
                        f"TOR='{result.tor_qa_signoff}', CSQA='{result.csqa_owner}', "
                        f"scope_records={result.scope_approval_record_count}, "
                        f"has_enhancement={result.scope_approval_has_enhancement}, "
                        f"tasks={result.tasks_record_count}, tasks_all_closed={result.tasks_all_closed}"
                    )

                self.progress_queue.put((index, len(ticket_numbers)))

            self.log(f"Results saved to: {output_path}")
            self.log("Finished.")

        except Exception as exc:
            self.log(f"Fatal error: {type(exc).__name__}: {exc}")
        finally:
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass


def cli_main() -> int:
    parser = argparse.ArgumentParser(description="Launch the Ticket Batch Checker GUI.")
    parser.parse_args()

    root = tk.Tk()
    TicketCheckerGUI(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(cli_main())
