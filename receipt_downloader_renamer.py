import os
import re
import threading
import queue
import time
import random
from datetime import datetime
from typing import Dict, List

# --- Main GUI and Automation Libraries ---
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
from tkcalendar import DateEntry
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyPDF2 import PdfReader

class PLIDownloaderApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PLI Downloader & Renamer")
        self.root.geometry("600x650") # Increased height for new buttons
        self.root.configure(bg="#f0f2f5")

        self.login_url = "https://pli.indiapost.gov.in/CustomerPortal/PSLogin.action"
        self.driver = None
        
        self.download_dir = os.path.join(os.getcwd(), "receipts")
        os.makedirs(self.download_dir, exist_ok=True)

        self.log_queue = queue.Queue()
        self.create_widgets()
        self.process_log_queue()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        """Creates and arranges all the GUI elements."""
        style = ttk.Style()
        style.configure('TFrame', background='#f0f2f5')
        style.configure('TLabel', background='#f0f2f5', font=('Helvetica', 10))
        style.configure('TButton', font=('Helvetica', 10, 'bold'))

        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main_frame, text="PLI Downloader & Renamer", font=("Helvetica", 16, "bold")).pack(pady=(0, 10))

        instructions_frame = ttk.LabelFrame(main_frame, text="Instructions", padding="10")
        instructions_frame.pack(fill=tk.X, pady=10)
        instructions_text = (
            "1. Click 'Launch Chrome'.\n"
            "2. In the new Chrome window, log in and display the list of receipts.\n"
            "3. In this app, select your desired date range below.\n"
            "4. Click 'Download & Rename Receipts'."
        )
        ttk.Label(instructions_frame, text=instructions_text, justify=tk.LEFT).pack(anchor='w')

        date_frame = ttk.Frame(main_frame)
        date_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(date_frame, text="Download receipts from:").grid(row=0, column=0, padx=5, pady=5)
        self.start_date_entry = DateEntry(date_frame, date_pattern='dd/MM/yyyy', width=12, background='darkblue', foreground='white', borderwidth=2, toplevel_class=tk.Toplevel)
        self.start_date_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(date_frame, text="to:").grid(row=0, column=2, padx=5, pady=5)
        self.end_date_entry = DateEntry(date_frame, date_pattern='dd/MM/yyyy', width=12, background='darkblue', foreground='white', borderwidth=2, toplevel_class=tk.Toplevel)
        self.end_date_entry.grid(row=0, column=3, padx=5, pady=5)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.launch_button = ttk.Button(button_frame, text="1. Launch Chrome", command=self.launch_chrome_thread)
        self.launch_button.pack(side=tk.LEFT, expand=True, padx=5)

        self.download_button = ttk.Button(button_frame, text="2. Download & Rename Receipts", command=self.start_download_thread, state=tk.DISABLED)
        self.download_button.pack(side=tk.LEFT, expand=True, padx=5)
        
        # --- NEW BUTTONS ADDED HERE ---
        action_button_frame = ttk.Frame(main_frame)
        action_button_frame.pack(fill=tk.X, pady=10)

        self.reset_button = ttk.Button(action_button_frame, text="Reset", command=self.reset_app_state)
        self.reset_button.pack(side=tk.LEFT, expand=True, padx=5)
        
        self.stop_exit_button = ttk.Button(action_button_frame, text="Stop and Exit", command=self.on_closing)
        self.stop_exit_button.pack(side=tk.LEFT, expand=True, padx=5)
        # --- END OF NEW BUTTONS ---

        self.progress_bar = ttk.Progressbar(main_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(main_frame, text="Status Log:", font=('Helvetica', 12, 'bold')).pack(pady=(10, 5))
        self.status_log = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=10, font=('Courier New', 9))
        self.status_log.pack(fill=tk.BOTH, expand=True)

    def log_message(self, message: str):
        self.log_queue.put(message)

    def process_log_queue(self):
        try:
            while not self.log_queue.empty():
                message = self.log_queue.get_nowait()
                self.status_log.insert(tk.END, message + "\n")
                self.status_log.see(tk.END)
        finally:
            self.root.after(100, self.process_log_queue)

    def launch_chrome_thread(self):
        """Starts the browser launch in a thread to keep the GUI responsive."""
        self.launch_button.config(state=tk.DISABLED)
        self.progress_bar.start()
        thread = threading.Thread(target=self.launch_chrome, daemon=True)
        thread.start()

    def launch_chrome(self):
        """Launches a fully controlled Selenium browser instance."""
        self.log_message(" Launching controlled Chrome browser...")
        try:
            options = webdriver.ChromeOptions()
            prefs = {
                "download.default_directory": self.download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True,
            }
            options.add_experimental_option("prefs", prefs)
            
            self.driver = webdriver.Chrome(options=options)
            self.driver.get(self.login_url)
            
            self.log_message(" Browser launched. Please log in and navigate to the receipts page.")
            self.root.after(0, self.download_button.config, {'state': tk.NORMAL})
        except WebDriverException as e:
            self.log_message(f" ERROR: Could not launch Chrome. {e}")
            messagebox.showerror("Error", "Failed to launch Chrome. Please ensure it is installed and not blocked.")
        finally:
            self.root.after(0, self.progress_bar.stop)

    def start_download_thread(self):
        """Starts the main receipt downloading process in a separate thread."""
        if not self.driver or not self.is_driver_alive():
            messagebox.showerror("Error", "The browser is not running. Please launch it first.")
            return
        self.download_button.config(state=tk.DISABLED)
        self.progress_bar['value'] = 0
        thread = threading.Thread(target=self.download_and_rename_receipts, daemon=True)
        thread.start()

    def download_and_rename_receipts(self):
        """Gathers, sorts, downloads, and then renames receipts."""
        try:
            main_window_handle = self.driver.current_window_handle
            
            self.log_message(" Scanning page for receipts...")
            wait = WebDriverWait(self.driver, 10)
            robust_row_selector = (By.XPATH, "//a[text()='Download Receipt']/ancestor::tr")
            wait.until(EC.presence_of_element_located(robust_row_selector))
            table_rows = self.driver.find_elements(*robust_row_selector)
            
            if not table_rows:
                self.log_message(" No receipt rows found on the current page.")
                return

            self.log_message(f"Found {len(table_rows)} total receipts. Filtering and sorting...")
            
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date()
            
            receipts_to_process = []
            for row in table_rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) < 4: continue

                policy_num = cells[0].text
                premium_from_date_str = cells[3].text
                
                try:
                    receipt_date = datetime.strptime(premium_from_date_str, '%d/%m/%Y').date()
                    if start_date <= receipt_date <= end_date:
                        receipts_to_process.append({
                            "policy": policy_num,
                            "date_str": premium_from_date_str,
                            "date_obj": receipt_date
                        })
                except (ValueError, WebDriverException):
                    continue

            receipts_to_process.sort(key=lambda x: x['date_obj'])
            
            if not receipts_to_process:
                self.log_message(" No receipts match the selected date range.")
                return

            total_files = len(receipts_to_process)
            self.log_message(f"Found {total_files} receipts to process in the specified range.")
            self.root.after(0, self.progress_bar.config, {'maximum': total_files * 2})
            
            downloaded_files = []
            for i, receipt in enumerate(receipts_to_process):
                self.log_message(f"  ({i+1}/{total_files}) Downloading receipt for {receipt['policy']} ({receipt['date_str']})...")
                try:
                    all_links = self.driver.find_elements(By.LINK_TEXT, "Download Receipt")
                    link_to_click = None
                    for link in all_links:
                        row = link.find_element(By.XPATH, "./ancestor::tr")
                        cells = row.find_elements(By.TAG_NAME, "td")
                        if cells[0].text == receipt['policy'] and cells[3].text == receipt['date_str']:
                            link_to_click = link
                            break
                    
                    if not link_to_click:
                        self.log_message(f"     Could not re-find link for {receipt['policy']}. Skipping.")
                        continue

                    files_before = set(os.listdir(self.download_dir))
                    link_to_click.click()
                    
                    new_file_path = self.wait_for_download(files_before)
                    if new_file_path:
                        downloaded_files.append(new_file_path)
                        self.log_message(f"     Downloaded as: {os.path.basename(new_file_path)}")
                    else:
                        self.log_message(f"     Timed out waiting for file to download.")
                    
                    self.root.after(0, self.progress_bar.step)
                    
                    delay = random.uniform(3, 7)
                    self.log_message(f"    ... waiting for {delay:.1f} seconds.")
                    time.sleep(delay)

                except Exception as inner_e:
                    self.log_message(f"     Failed to process this receipt. Error: {inner_e}")
                    continue
            
            self.log_message("\n Starting renaming process...")
            for i, filepath in enumerate(downloaded_files):
                self.rename_pdf(filepath, i + 1, total_files)
                self.root.after(0, self.progress_bar.step)

            self.log_message(f"\n Process complete.")
        except TimeoutException:
            self.log_message(" No receipt rows found on the current page (Timed out waiting).")
        except Exception as e:
            self.log_message(f" An error occurred: {e}")
        finally:
            self.root.after(0, self.finalize_ui)

    def wait_for_download(self, files_before: set) -> str or None:
        """Waits for a new file to appear in the download dir."""
        for _ in range(20):
            time.sleep(1)
            files_after = set(os.listdir(self.download_dir))
            new_files = files_after - files_before
            if new_files:
                new_filename = new_files.pop()
                while new_filename.endswith('.crdownload'):
                    time.sleep(1)
                    files_after = set(os.listdir(self.download_dir))
                    new_files = files_after - files_before
                    if not new_files: return None
                    new_filename = new_files.pop()
                return os.path.join(self.download_dir, new_filename)
        return None

    def rename_pdf(self, filepath: str, count: int, total: int):
        """Reads a PDF, extracts info, and renames it."""
        self.log_message(f"  ({count}/{total}) Reading {os.path.basename(filepath)}...")
        for attempt in range(5):
            try:
                with open(filepath, 'rb') as f:
                    reader = PdfReader(f)
                    text = ""
                    for page in reader.pages:
                        text += page.extract_text().replace('\n', ' ')

                trans_num_match = re.search(r'([A-Z]{2}\d{8,})', text, re.IGNORECASE)
                date_match = re.search(r'(\d{2}/\d{2}/\d{4})\s+\d{2}/\d{2}/\d{4}', text, re.IGNORECASE)
                amount_match = re.search(r'Total Paid Amount\s*\*\s*:.*?(\d+)\.\d+', text, re.IGNORECASE | re.DOTALL)

                if trans_num_match and date_match and amount_match:
                    trans_num = trans_num_match.group(1)
                    date_str = date_match.group(1)
                    amount = amount_match.group(1)
                    
                    date_obj = datetime.strptime(date_str, '%d/%m/%Y')
                    formatted_date = date_obj.strftime('%d%m%Y')
                    
                    new_filename = f"{amount}_{trans_num}_{formatted_date}.pdf"
                    new_filepath = os.path.join(self.download_dir, new_filename)
                    
                else:
                    self.log_message(f"     Could not find all required info in PDF. Not renamed.")
                    return

                os.rename(filepath, new_filepath)
                self.log_message(f"     Renamed to: {new_filename}")
                return

            except FileNotFoundError:
                self.log_message(f"    ... file not found on attempt {attempt+1}, retrying...")
                time.sleep(1)
            except Exception as e:
                self.log_message(f"     Error reading or renaming file. Error: {e}")
                return
        
        self.log_message(f"     Failed to access {os.path.basename(filepath)} after multiple attempts.")

    def finalize_ui(self):
        """Resets the UI elements after an operation completes."""
        self.progress_bar['value'] = 0
        self.download_button.config(state=tk.NORMAL)

    # --- NEW METHOD ADDED HERE ---
    def reset_app_state(self):
        """Clears the log, resets the progress bar, and re-enables buttons."""
        self.status_log.delete('1.0', tk.END)
        self.progress_bar['value'] = 0
        self.log_message(" Application state has been reset.")
        self.launch_button.config(state=tk.NORMAL)
        if self.driver and self.is_driver_alive():
            self.download_button.config(state=tk.NORMAL)
        else:
            self.download_button.config(state=tk.DISABLED)
    # --- END OF NEW METHOD ---

    def is_driver_alive(self):
        """Checks if the Selenium browser window is still open."""
        try:
            _ = self.driver.title
            return True
        except Exception:
            return False

    def on_closing(self):
        """Handles cleanup when the application window is closed."""
        if self.driver and self.is_driver_alive():
            try:
                self.driver.quit()
            except Exception:
                pass
        self.root.destroy()

if __name__ == '__main__':
    root = tk.Tk()
    app = PLIDownloaderApp(root)
    root.mainloop()
