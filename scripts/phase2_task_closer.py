from __future__ import annotations

import argparse
import csv
import queue
import threading
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional, List, Tuple
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
    "tasks_link": (By.CSS_SELECTOR, "li#tasks a"),
    "tasks_iframe": (By.CSS_SELECTOR, "iframe"),
    "tasks_table": (By.CSS_SELECTOR, "table#mytable"),
    "task_edit_form": (By.CSS_SELECTOR, "form#Form1"),
    "task_status_select": (By.CSS_SELECTOR, "select#status"),
    "task_update_button": (By.CSS_SELECTOR, "input#sub"),
}


@dataclass
class TaskRow:
    ticket_number: str
    ticket_url: str
    tasks_url: str
    task_id: str
    edit_url: str
    description: str
    assigned_to: str
    previous_status: str


@dataclass
class TaskUpdateResult:
    ticket_number: str
    ticket_url: str
    tasks_url: str
    task_id: str = ""
    edit_url: str = ""
    description: str = ""
    assigned_to: str = ""
    previous_status: str = ""
    target_status: str = "Closed"
    action: str = ""
    update_attempted: bool = False
    update_success: Optional[bool] = None
    checkboxes_visible_enabled_count: Optional[int] = None
    checkboxes_checked_before_count: Optional[int] = None
    checkboxes_checked_after_count: Optional[int] = None
    status_after_update: str = ""
    error: str = ""


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
        log_func("Login appears successful." if success else "Login still appears to be on the login page.")
    return True, success


def open_ticket_and_login_if_needed(
    driver: WebDriver,
    ticket_url: str,
    username: str,
    password: str,
    log_func=None,
) -> Tuple[bool, bool, str]:
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


def read_tasks_link(driver: WebDriver, ticket_url: str) -> Tuple[str, str]:
    link = safe_find(driver, "tasks_link")
    if link is None:
        return "", ""

    text = link.text.strip()
    href = (link.get_attribute("href") or "").strip()
    absolute = urljoin(ticket_url, href)
    return text, absolute


def switch_into_tasks_table_context(driver: WebDriver, timeout: int = 20) -> None:
    driver.switch_to.default_content()

    if safe_find(driver, "tasks_table") is not None:
        return

    iframe = wait_for(driver, "tasks_iframe", timeout=timeout)
    driver.switch_to.frame(iframe)
    wait_for(driver, "tasks_table", timeout=timeout)


def parse_tasks_table(driver: WebDriver, ticket_number: str, ticket_url: str, tasks_url: str) -> List[TaskRow]:
    table = wait_for(driver, "tasks_table", timeout=20)
    rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
    tasks: List[TaskRow] = []

    for row in rows:
        cells = row.find_elements(By.CSS_SELECTOR, "td")
        if len(cells) < 4:
            continue

        task_id = cells[0].text.strip()
        if not task_id.isdigit():
            continue

        edit_url = ""
        try:
            edit_link = cells[1].find_element(By.CSS_SELECTOR, "a")
            edit_url = urljoin(tasks_url, edit_link.get_attribute("href") or "")
        except NoSuchElementException:
            pass

        description = cells[3].text.strip() if len(cells) > 3 else ""
        assigned_to = cells[4].text.strip() if len(cells) > 4 else ""
        status = cells[-1].text.strip() if cells else ""

        tasks.append(
            TaskRow(
                ticket_number=ticket_number,
                ticket_url=ticket_url,
                tasks_url=tasks_url,
                task_id=task_id,
                edit_url=edit_url,
                description=description,
                assigned_to=assigned_to,
                previous_status=status,
            )
        )

    return tasks


def is_task_unfinished(status: str) -> bool:
    return status.strip().lower() != "closed"


def open_tasks_page_and_get_tasks(
    driver: WebDriver,
    ticket_number: str,
    ticket_url: str,
    tasks_url: str,
    timeout: int = 20,
) -> List[TaskRow]:
    driver.switch_to.default_content()
    driver.get(tasks_url)
    WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")
    switch_into_tasks_table_context(driver, timeout=timeout)
    time.sleep(0.25)
    return parse_tasks_table(driver, ticket_number, ticket_url, tasks_url)


def choose_closed_status(status_select: WebElement) -> None:
    select = Select(status_select)

    try:
        select.select_by_visible_text("Closed")
        return
    except Exception:
        pass

    for option in select.options:
        if option.text.strip().lower() == "closed":
            option.click()
            return

    raise ValueError("Could not find a 'Closed' option in the status dropdown.")


def current_selected_status(driver: WebDriver) -> str:
    status_select = wait_for(driver, "task_status_select", timeout=20)
    try:
        return Select(status_select).first_selected_option.text.strip()
    except Exception:
        return ""


def visible_enabled_checkboxes(driver: WebDriver) -> List[WebElement]:
    form = wait_for(driver, "task_edit_form", timeout=20)
    checkboxes = form.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
    usable: List[WebElement] = []

    for checkbox in checkboxes:
        try:
            if checkbox.is_displayed() and checkbox.is_enabled():
                usable.append(checkbox)
        except WebDriverException:
            pass

    return usable


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


def update_single_task(
    driver: WebDriver,
    task: TaskRow,
    target_status: str = "Closed",
    timeout: int = 20,
    log_func=None,
) -> TaskUpdateResult:
    result = TaskUpdateResult(
        ticket_number=task.ticket_number,
        ticket_url=task.ticket_url,
        tasks_url=task.tasks_url,
        task_id=task.task_id,
        edit_url=task.edit_url,
        description=task.description,
        assigned_to=task.assigned_to,
        previous_status=task.previous_status,
        target_status=target_status,
        action="update_to_closed",
    )

    if not task.edit_url:
        result.error = "No edit URL found for task."
        result.update_attempted = False
        result.update_success = False
        return result

    try:
        if log_func:
            log_func(f"    Opening task edit page for task {task.task_id}...")

        driver.switch_to.default_content()
        driver.get(task.edit_url)
        WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")

        driver.switch_to.default_content()
        if safe_find(driver, "task_edit_form") is None:
            try:
                iframe = wait_for(driver, "tasks_iframe", timeout=5)
                driver.switch_to.frame(iframe)
            except TimeoutException:
                pass

        wait_for(driver, "task_edit_form", timeout=timeout)

        before_checkboxes = visible_enabled_checkboxes(driver)
        result.checkboxes_visible_enabled_count = len(before_checkboxes)
        result.checkboxes_checked_before_count = sum(1 for cb in before_checkboxes if cb.is_selected())

        status_select = wait_for(driver, "task_status_select", timeout=timeout)
        choose_closed_status(status_select)

        for checkbox in visible_enabled_checkboxes(driver):
            if not checkbox.is_selected():
                checkbox.click()
                time.sleep(0.05)

        after_checkboxes = visible_enabled_checkboxes(driver)
        result.checkboxes_checked_after_count = sum(1 for cb in after_checkboxes if cb.is_selected())

        update_button = wait_for(driver, "task_update_button", timeout=timeout)
        result.update_attempted = True
        update_button.click()

        alert_text = handle_possible_alert(driver, log_func=log_func)
        if alert_text:
            result.error = f"Alert after Update: {alert_text}"

        time.sleep(1.0)

        driver.switch_to.default_content()
        driver.get(task.edit_url)
        WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")

        driver.switch_to.default_content()
        if safe_find(driver, "task_edit_form") is None:
            try:
                iframe = wait_for(driver, "tasks_iframe", timeout=5)
                driver.switch_to.frame(iframe)
            except TimeoutException:
                pass

        wait_for(driver, "task_edit_form", timeout=timeout)
        result.status_after_update = current_selected_status(driver)
        result.update_success = result.status_after_update.strip().lower() == "closed"

        if not result.update_success and not result.error:
            result.error = f"Status after update is '{result.status_after_update}', expected Closed."

    except Exception as exc:
        result.update_success = False
        result.error = f"{type(exc).__name__}: {exc}"

    return result


def save_results(results: List[TaskUpdateResult], output_path: Path) -> None:
    ordered_cols = [
        "ticket_number",
        "ticket_url",
        "tasks_url",
        "task_id",
        "edit_url",
        "description",
        "assigned_to",
        "previous_status",
        "target_status",
        "action",
        "update_attempted",
        "update_success",
        "checkboxes_visible_enabled_count",
        "checkboxes_checked_before_count",
        "checkboxes_checked_after_count",
        "status_after_update",
        "error",
    ]

    df = pd.DataFrame([asdict(r) for r in results])
    if df.empty:
        df = pd.DataFrame(columns=ordered_cols)
    else:
        df = df[ordered_cols]

    suffix = output_path.suffix.lower()
    if suffix == ".xlsx":
        df.to_excel(output_path, index=False)
    elif suffix == ".csv":
        df.to_csv(output_path, index=False)
    else:
        raise ValueError("Output must be .xlsx or .csv")


class TaskCloserGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Phase 2 - Ticket Task Closer")
        self.root.geometry("980x760")

        self.worker_thread: Optional[threading.Thread] = None
        self.stop_event = threading.Event()
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.progress_queue: queue.Queue[Tuple[int, int]] = queue.Queue()

        self.base_url_var = tk.StringVar(value="YOUR_TICKET_SYSTEM_URL/edit_bug.aspx?id=")
        self.ticket_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar(value=str(Path.cwd() / "ticket_task_closer_results.xlsx"))
        self.user_data_dir_var = tk.StringVar()
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.headless_var = tk.BooleanVar(value=False)
        self.pause_var = tk.DoubleVar(value=0.5)
        self.dry_run_var = tk.BooleanVar(value=False)

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

        ttk.Label(frame, text="Output audit file:").grid(row=2, column=0, sticky="w", **pad)
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
        ttk.Checkbutton(options_frame, text="Headless mode not recommended", variable=self.headless_var).pack(side="left")
        ttk.Checkbutton(options_frame, text="Dry run only - do not update", variable=self.dry_run_var).pack(side="left", padx=(20, 0))
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

        instructions = (
            "Behavior: For each ticket, opens Tasks/Time. Any task whose status is blank or not Closed is edited. "
            "The script sets Status to Closed, checks all visible/enabled checkboxes, clicks Update, and records the result."
        )
        ttk.Label(frame, text=instructions, foreground="gray", wraplength=900).grid(
            row=10, column=0, columnspan=3, sticky="w", padx=8, pady=(10, 5)
        )

        ttk.Label(frame, text="Live progress log:").grid(row=11, column=0, columnspan=3, sticky="w", padx=8, pady=(12, 5))
        self.log_text = tk.Text(frame, height=24, wrap="word")
        self.log_text.grid(row=12, column=0, columnspan=3, sticky="nsew", padx=8, pady=5)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=12, column=3, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(12, weight=1)

    def choose_ticket_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="Choose ticket list",
            filetypes=[("Ticket files", "*.txt *.csv *.xlsx *.xls"), ("All files", "*.*")],
        )
        if filename:
            self.ticket_file_var.set(filename)

    def choose_output_file(self) -> None:
        filename = filedialog.asksaveasfilename(
            title="Save audit results as",
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
        if not self.base_url_var.get().strip():
            messagebox.showerror("Missing base URL", "Enter the base ticket URL.")
            return
        if not self.ticket_file_var.get().strip():
            messagebox.showerror("Missing ticket list", "Choose a TXT, CSV, or XLSX ticket list file.")
            return
        if not self.output_file_var.get().strip():
            messagebox.showerror("Missing output file", "Choose where to save the audit results.")
            return

        self.stop_event.clear()
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.progress.configure(value=0, maximum=1)
        self.status_label.configure(text="Starting...")
        self.log_text.delete("1.0", "end")

        self.worker_thread = threading.Thread(target=self.run_task_closer, daemon=True)
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

    def run_task_closer(self) -> None:
        base_url = self.base_url_var.get().strip()
        tickets_path = Path(self.ticket_file_var.get().strip())
        output_path = Path(self.output_file_var.get().strip())
        headless = self.headless_var.get()
        dry_run = self.dry_run_var.get()
        pause_seconds = float(self.pause_var.get())
        user_data_dir = self.user_data_dir_var.get().strip()
        username = self.username_var.get()
        password = self.password_var.get()

        driver: Optional[WebDriver] = None
        results: List[TaskUpdateResult] = []

        try:
            ticket_numbers = read_ticket_numbers(tickets_path)
            if not ticket_numbers:
                self.log("No ticket numbers found in the selected file.")
                return

            self.progress_queue.put((0, len(ticket_numbers)))
            self.log(f"Loaded {len(ticket_numbers)} tickets.")
            if dry_run:
                self.log("DRY RUN is ON. No updates will be submitted.")
            self.log("Opening Chrome. Login will be attempted only if the login page is detected.")

            driver = make_driver(headless=headless, user_data_dir=user_data_dir)

            for index, ticket_number in enumerate(ticket_numbers, start=1):
                if self.stop_event.is_set():
                    self.log("Stopped before starting next ticket.")
                    break

                ticket_url = build_ticket_url(base_url, ticket_number)
                self.log(f"[{index}/{len(ticket_numbers)}] Processing ticket {ticket_number}...")

                _, _, ticket_error = open_ticket_and_login_if_needed(
                    driver=driver,
                    ticket_url=ticket_url,
                    username=username,
                    password=password,
                    log_func=self.log,
                )

                if ticket_error:
                    self.log(f"  Ticket error: {ticket_error}")
                    results.append(
                        TaskUpdateResult(
                            ticket_number=ticket_number,
                            ticket_url=ticket_url,
                            tasks_url="",
                            action="ticket_open_failed",
                            update_attempted=False,
                            update_success=False,
                            error=ticket_error,
                        )
                    )
                    save_results(results, output_path)
                    self.progress_queue.put((index, len(ticket_numbers)))
                    continue

                _, tasks_url = read_tasks_link(driver, ticket_url)
                if not tasks_url:
                    self.log("  Tasks/Time link not found.")
                    results.append(
                        TaskUpdateResult(
                            ticket_number=ticket_number,
                            ticket_url=ticket_url,
                            tasks_url="",
                            action="tasks_link_not_found",
                            update_attempted=False,
                            update_success=False,
                            error="Tasks/Time link not found.",
                        )
                    )
                    save_results(results, output_path)
                    self.progress_queue.put((index, len(ticket_numbers)))
                    continue

                self.log(f"  Opening Tasks/Time: {tasks_url}")
                tasks = open_tasks_page_and_get_tasks(driver, ticket_number, ticket_url, tasks_url)
                unfinished = [task for task in tasks if is_task_unfinished(task.previous_status)]

                self.log(f"  Found {len(tasks)} tasks. Unfinished/not closed: {len(unfinished)}.")

                if not unfinished:
                    results.append(
                        TaskUpdateResult(
                            ticket_number=ticket_number,
                            ticket_url=ticket_url,
                            tasks_url=tasks_url,
                            action="no_unfinished_tasks",
                            update_attempted=False,
                            update_success=True,
                        )
                    )
                    save_results(results, output_path)
                    self.progress_queue.put((index, len(ticket_numbers)))
                    continue

                for task in unfinished:
                    if self.stop_event.is_set():
                        self.log("  Stop requested. Finishing current ticket loop.")
                        break

                    self.log(
                        f"  Updating task {task.task_id}: status='{task.previous_status}', "
                        f"description='{task.description}'"
                    )

                    if dry_run:
                        dry_result = TaskUpdateResult(
                            ticket_number=task.ticket_number,
                            ticket_url=task.ticket_url,
                            tasks_url=task.tasks_url,
                            task_id=task.task_id,
                            edit_url=task.edit_url,
                            description=task.description,
                            assigned_to=task.assigned_to,
                            previous_status=task.previous_status,
                            action="dry_run_would_update_to_closed",
                            update_attempted=False,
                            update_success=None,
                            error="Dry run only. No update submitted.",
                        )
                        results.append(dry_result)
                        save_results(results, output_path)
                        continue

                    result = update_single_task(
                        driver=driver,
                        task=task,
                        target_status="Closed",
                        log_func=self.log,
                    )
                    results.append(result)
                    save_results(results, output_path)

                    if result.update_success:
                        self.log(f"    Task {task.task_id} updated successfully.")
                    else:
                        self.log(f"    Task {task.task_id} update failed: {result.error}")

                    if pause_seconds > 0:
                        time.sleep(pause_seconds)

                self.progress_queue.put((index, len(ticket_numbers)))

            self.log(f"Audit results saved to: {output_path}")
            self.log("Finished.")

        except Exception as exc:
            self.log(f"Fatal error: {type(exc).__name__}: {exc}")
            try:
                save_results(results, output_path)
            except Exception:
                pass
        finally:
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass


def cli_main() -> int:
    parser = argparse.ArgumentParser(description="Launch Phase 2 Ticket Task Closer GUI.")
    parser.parse_args()

    root = tk.Tk()
    TaskCloserGUI(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(cli_main())
