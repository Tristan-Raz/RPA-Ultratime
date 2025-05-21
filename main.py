import tkinter as tk
from tkinter import messagebox, filedialog
import pyautogui as ag
import time
from tkcalendar import Calendar
from datetime import datetime
import pandas as pd
import json
import pytesseract
from PIL import ImageGrab
import os
import sys
import re  # added for normalization


# Set the path to tesseract executable (adjust as needed)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
SETTINGS_FILE = "settings.json"


class GUI:
    def __init__(self):
        self.root = tk.Tk()
        # Initialize important variables
        self.last_names = []
        self.current_index = 0
        self.start_x = None
        self.start_y = None
        self.date1 = None
        self.date2 = None
        self.dates = [
            "Sunday", "Monday", "Tuesday", "Wednesday",
            "Thursday", "Friday", "Saturday"
        ]
        self.ocr_box = None  # To store OCR selection coordinates
        self.skipped_names = []  # To store names that were not found
        self.debug_mode = True  # Set to True to show OCR debug info                   #VITAL TO DEBUGGING
        self.root.bind('<Control-Shift-Q>', self.kill_process)                         #VITAL TO STOP!


        # Create GUI Window
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Set an adaptive window size (e.g., 30% of the screen width, 45% of the height)
        initial_width = int(screen_width * 0.45)
        initial_height = int(screen_height * 0.5)
        min_width = 600
        min_height = 700
        max_width = int(screen_width * 0.9)
        max_height = int(screen_height * 0.9)

        self.root.geometry(f"{initial_width}x{initial_height}")
        self.root.minsize(min_width, min_height)
        self.root.maxsize(max_width, max_height)
        self.root.title("PPD Sessions")
        self.root.resizable(True, True)

        # Configure row and column scaling
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=2)
        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=2)
        self.root.rowconfigure(2, weight=1)
        self.root.bind("<Configure>", self.on_resize)

        # Setup UI Components
        self.create_widgets()
        self.load_settings()

        self.max_scroll_attempts = 10  # Prevent infinite loops

        self.root.mainloop()

    ### === UI COMPONENTS === ###

    def kill_process(self, event=None):
        """Hotkey function to immediately terminate the process."""
        print("[DEBUG] Kill hotkey pressed. Terminating process.")
        # Optionally, add any cleanup code here.
        self.root.destroy()
        sys.exit(0)

    def on_resize(self, event):
        self.root.update_idletasks()

    def create_widgets(self):
        self.create_instructions()
        self.create_calendar()
        self.create_checkboxes()
        self.create_buttons()

    def create_instructions(self):
        instructtext = '''
        Welcome to the PPD Sessions Tool. To use:
        1. Select UltraView and ensure it is in FULLSCREEN mode.
        2. Select the Excel File containing names.
        3. Select the Dates to skip.
        4. Select Coordinates for "Select Employee Subgroup By:" and calibrate.
        5. The OCR will automatically read each name and process it:
           - If the name is found in the top (first) line, it scrolls down one step then back up, then presses Space.
           - If found in the second line, it scrolls down then presses Space.
        6. Names not found after scrolling are automatically logged.
        '''
        label = tk.Label(self.root, text=instructtext, justify="left", anchor="w")
        label.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="w")

    def create_calendar(self):
        calendar_frame = tk.Frame(self.root)
        calendar_frame.grid(row=1, column=1, sticky="nsew")
        self.cal = Calendar(calendar_frame, selectmode="day")
        self.cal.pack(expand=True, fill="both")
        tk.Button(calendar_frame, text="Select Date", command=self.select_date).pack(pady=5)

    def create_checkboxes(self):
        checkbox_frame = tk.Frame(self.root)
        checkbox_frame.grid(row=1, column=0, sticky="nw")
        self.check_vars = {}
        for date in self.dates:
            var = tk.IntVar(value=1 if date in ["Saturday", "Sunday"] else 0)
            self.check_vars[date] = var
            if date not in ["Saturday", "Sunday"]:
                tk.Checkbutton(checkbox_frame, text=date, variable=var).pack(anchor="w", padx=5, pady=2)

    def create_buttons(self):
        button_frame = tk.Frame(self.root)
        button_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky="ew")
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)
        self.root.columnconfigure(2, weight=1)
        button_width = 15
        tk.Button(button_frame, text="START PROCESS", command=self.start_process, width=button_width).grid(row=0,column=0,padx=5,pady=5,sticky="ew")
        tk.Button(button_frame, text="Show Dates", command=self.show_selected_dates, width=button_width).grid(row=0,column=1,padx=5,pady=5,sticky="ew")
        tk.Button(button_frame, text="Select File", command=self.select_file, width=button_width).grid(row=0, column=2,padx=5, pady=5,sticky="ew")
        tk.Button(button_frame, text="Reset", command=self.reset, width=button_width).grid(row=1, column=0, padx=5,pady=5, sticky="ew")
        tk.Button(button_frame, text="Save Settings", command=self.save_settings, width=button_width).grid(row=1,column=1,padx=5,pady=5,sticky="ew")
        tk.Button(button_frame, text="Load Settings", command=self.load_settings, width=button_width).grid(row=1,column=2,padx=5,pady=5,sticky="ew")
        tk.Button(self.root, text="Select Coordinate", command=self.capture_click).grid(row=3, column=1, pady=10,sticky="ew")
        tk.Button(button_frame, text="Set OCR Box", command=self.set_ocr_box, width=button_width).grid(row=2, column=0,padx=5, pady=5)
        tk.Button(button_frame, text="Test OCR Box", command=self.test_ocr_box, width=button_width).grid(row=2, column=1,padx=5, pady=5)

        # The "Automate" button has been removed as the OCR automation is automatically triggered.

    ### === FILE & DATA HANDLING === ###

    def save_settings(self):
        try:
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    settings = json.load(f)
            except (FileNotFoundError, json.JSONDecodeError):
                settings = {}

            settings.update({
                "last_file": self.last_file,
                "date1": self.date1,
                "date2": self.date2,
                "start_x": self.start_x,
                "start_y": self.start_y,
                "ocr_box": self.ocr_box,
                "checkboxes": {key: var.get() for key, var in self.check_vars.items()}
            })

            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4)
            messagebox.showinfo("Settings Saved", "Your settings have been saved successfully.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Error saving settings: {str(e)}")

    def load_settings(self):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                settings = json.load(f)

            self.last_file = settings.get("last_file", "")
            self.date1 = settings.get("date1", datetime.today().strftime("%m/%d/%Y"))
            self.date2 = settings.get("date2", self.date1)
            self.start_x = settings.get("start_x", None)
            self.start_y = settings.get("start_y", None)
            self.ocr_box = settings.get("ocr_box", None)

            for key, value in settings.get("checkboxes", {}).items():
                if key in self.check_vars:
                    self.check_vars[key].set(value)

            messagebox.showinfo("Settings Loaded", "Previous settings have been restored.")
        except (FileNotFoundError, json.JSONDecodeError):
            self.last_file = ""
            self.date1 = datetime.today().strftime("%m/%d/%Y")
            self.date2 = self.date1
            self.start_x = None
            self.start_y = None
            self.ocr_box = None

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.load_last_names(file_path)

    def load_last_names(self, file_path):
        try:
            df = pd.read_excel(file_path)
            if df.empty:
                messagebox.showwarning("Data Warning", "The selected file is empty.")
                return

            # Drop rows where column A is NaN
            df = df.dropna(subset=[df.columns[0]])

            # If the first row in column A contains "last name", skip it.
            if isinstance(df.iloc[0, 0], str) and "last name" in df.iloc[0, 0].lower():
                df = df.iloc[1:]

            # Create a list of tuples (last_name, first_name) from columns A and B.
            self.names = []
            for index, row in df.iterrows():
                last_name = str(row[df.columns[0]]).strip()
                first_name = str(row[df.columns[1]]).strip() if df.shape[1] > 1 else ""
                if last_name:  # only add rows with a non-empty last name
                    self.names.append((last_name, first_name))

            if not self.names:
                messagebox.showwarning("Data Warning", "No names found in the file.")
            else:
                messagebox.showinfo("File Loaded", "Names loaded successfully.")
                self.current_index = 0  # Reset index when a new file is loaded

        except FileNotFoundError:
            messagebox.showerror("File Error", "The selected file was not found.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def select_date(self):
        selected_date = self.cal.get_date()
        date_object = datetime.strptime(selected_date, "%m/%d/%y")
        formatted_date = date_object.strftime("%m/%d/%Y")

        if not self.date1:
            self.date1 = formatted_date
            self.date2 = formatted_date
            messagebox.showinfo("Selected Date",
                                f"First Date Selected: {self.date1}\nSecond Date (Auto-Set): {self.date2}")
        elif not self.date2 or self.date2 == self.date1:
            self.date2 = formatted_date
            messagebox.showinfo("Selected Date", f"Second Date Selected: {self.date2}")
        else:
            messagebox.showwarning("Selection Error", "Both dates have already been selected.")

    def show_selected_dates(self):
        if self.date1 and self.date2:
            messagebox.showinfo("Selected Dates", f"Date 1: {self.date1}, Date 2: {self.date2}")
        else:
            messagebox.showwarning("Date Selection", "Both dates have not been selected yet.")

    ### === PROCESSING LOGIC === ###

    def start_process(self):
        """Start the PPD processing sequence."""
        if not self.names:
            messagebox.showwarning("Process Error", "No names available to process.")
            return
        if self.start_x is None or self.start_y is None:
            messagebox.showwarning("Process Error", "Starting position not set.")
            return

        self.current_index = 0
        self.move_to_starting_position()
        self.enter_basic()
        self.get_ppd()

    def get_ppd(self):
        if not self.names:
            messagebox.showwarning("Error", "No names loaded!")
            return
        self.current_index = 0
        # Process first name immediately
        self.process_current_name()

    def move_to_starting_position(self):
        if self.start_x is not None and self.start_y is not None:
            ag.moveTo(self.start_x, self.start_y)
            ag.click()
        else:
            messagebox.showwarning("Error", "Starting position not set.")

    def enter_basic(self):
        # Navigation using Tab keys and data entry (adjust as needed)
        for _ in range(12):
            ag.press("tab")
            time.sleep(0.1)
        ag.typewrite("1")
        with ag.hold("shift"):
            ag.press("tab")
        ag.press("space")
        time.sleep(1)
        ag.typewrite("skill")
        time.sleep(3)
        ag.press("enter")
        with ag.hold("shift"):
            ag.press("tab")
            ag.press("tab")
        if self.date1:
            ag.typewrite(self.date1)
            time.sleep(1.5)
            ag.press('tab')
            time.sleep(1.5)
            if self.date2:
                ag.typewrite(self.date2)
                ag.press('tab')
            time.sleep(1)
        for _ in range(2):
            ag.press("tab")
        ag.press("space")
        for _ in range(6):
            ag.press("tab")
        ag.press("space")
        with ag.hold("shift"):
            for _ in range(6):
                ag.press("tab")
        for date in self.dates:
            if self.check_vars[date].get() == 1:
                ag.press("space")
                time.sleep(0.2)
            ag.press("tab")
            time.sleep(0.2)
        with ag.hold("shift"):
            for _ in range(19):
                ag.press("tab")
                time.sleep(0.1)
        ag.press("space")


    def process_current_name(self):
        """Types the last name (without spaces) and triggers OCR matching using the full name."""
        time.sleep(0.1)
        if self.current_index < len(self.names):
            last_name, first_name = self.names[self.current_index]
            # Remove spaces from the last name for typing to avoid auto-enter
            typed_last_name = last_name.replace(" ", "")
            # Use the full concatenated name for OCR matching (preserving spaces)
            full_name = f"{last_name} {first_name}".strip()
            time.sleep(.2)
            ag.typewrite(typed_last_name, interval=0.1)
            time.sleep(1.5)  # Allow UI to update before OCR scan
            self.automate_process(full_name)
        else:
            if self.skipped_names:
                messagebox.showinfo("Process Complete",
                                    f"All names have been processed.\nThe following names were skipped:\n{', '.join(self.skipped_names)}")
            else:
                messagebox.showinfo("Process Complete", "All names have been processed.")

    def automate_process(self, target_full_name):
        """Automates OCR-based name detection and action.
           Uses normalization to ignore special characters and case differences.
           Logs each OCR attempt with a screenshot.
           Tracks cumulative down presses to ensure that the offset is maintained
           across multiple OCR screenshots for the same name.
        """
        if not self.names or not self.ocr_box:
            messagebox.showwarning("Error", "No names loaded or OCR box not set.")
            return

        print(f"[DEBUG] Automating for: '{target_full_name}'")
        attempts = 0
        found = False
        tesseract_config = '--psm 6 --oem 3'
        cumulative_down = 0  # Total down key presses already applied for this name.
        max_found_index = -1  # Highest OCR line index where the target has been detected so far.

        # Normalization function: remove non-alphanumeric characters and convert to lowercase.
        def normalize_text(text):
            return re.sub(r'[^a-z0-9 ]', '', text.lower())

        normalized_target = normalize_text(target_full_name)

        while attempts < self.max_scroll_attempts:
            time.sleep(0.2)
            # Capture the OCR area.
            screenshot = ImageGrab.grab(bbox=self.ocr_box)
            screen_text = pytesseract.image_to_string(screenshot, config=tesseract_config).strip()
            # Split into non-empty lines.
            lines = [line.strip() for line in screen_text.splitlines() if line.strip()]
            print(f"[DEBUG] Attempt {attempts + 1}, OCR lines: {lines}")

            # Log the screenshot for this attempt.
            log_folder = "ocr_attempt_logs"
            if not os.path.exists(log_folder):
                os.makedirs(log_folder)
            log_filename = os.path.join(log_folder, f"attempt_{attempts + 1}.png")
            screenshot.save(log_filename)
            print(f"[DEBUG] Saved screenshot for attempt {attempts + 1} to {log_filename}")

            current_found_index = None
            for i, line in enumerate(lines):
                if normalized_target in normalize_text(line):
                    current_found_index = i
                    break

            if current_found_index is not None:
                # Update the max_found_index if this attempt shows the target further down.
                if current_found_index > max_found_index:
                    max_found_index = current_found_index
                print(f"[DEBUG] Detected '{target_full_name}' at OCR index {current_found_index}; "
                      f"max_found_index is now {max_found_index}.")
            else:
                print(f"[DEBUG] '{target_full_name}' not detected in this OCR attempt.")

            # Calculate additional down presses needed to reach the furthest detected index.
            additional_downs = max_found_index - cumulative_down if max_found_index >= 0 else 0
            if additional_downs > 0:
                print(
                    f"[DEBUG] Pressing down {additional_downs} times to align target (cumulative_down: {cumulative_down}).")
                for _ in range(additional_downs):
                    ag.press("down")
                    time.sleep(0.2)
                    cumulative_down += 1

            # Re-check OCR to see if the target is now in the first (selected) line.
            time.sleep(0.2)
            screenshot = ImageGrab.grab(bbox=self.ocr_box)
            screen_text = pytesseract.image_to_string(screenshot, config=tesseract_config).strip()
            lines = [line.strip() for line in screen_text.splitlines() if line.strip()]
            print(f"[DEBUG] After adjustment, OCR lines: {lines}")
            if lines and normalized_target in normalize_text(lines[0]):
                print(f"[DEBUG] '{target_full_name}' is now at the top (selected). Pressing space.")
                ag.press("space")
                found = True
                break
            else:
                # If still not selected, press down one more time as a fallback.
                print(f"[DEBUG] '{target_full_name}' still not at top. Pressing down once as fallback.")
                ag.press("down")
                time.sleep(1)
                cumulative_down += 1
                attempts += 1

        if not found:
            print(
                f"[DEBUG] '{target_full_name}' was not found after {self.max_scroll_attempts} attempts. Logging as skipped.")
            self.skipped_names.append(target_full_name)

        self.current_index += 1
        self.process_current_name()

    ### === SCREEN CAPTURE & OCR SETUP === ###

    def set_ocr_box(self):
        self.ocr_overlay = tk.Toplevel(self.root)
        self.ocr_overlay.overrideredirect(True)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.ocr_overlay.geometry(f"{screen_width}x{screen_height}+0+0")
        self.ocr_overlay.attributes("-alpha", 0.5)
        self.ocr_overlay.attributes("-topmost", True)
        self.ocr_overlay.config(bg='black')
        self.ocr_overlay.bind("<Escape>", lambda event: self.ocr_overlay.destroy())
        self.ocr_canvas = tk.Canvas(self.ocr_overlay, bg='black', highlightthickness=0)
        self.ocr_canvas.pack(fill="both", expand=True)
        self.ocr_canvas.bind("<Button-1>", self.ocr_start)
        self.ocr_canvas.bind("<B1-Motion>", self.ocr_move)
        self.ocr_canvas.bind("<ButtonRelease-1>", self.ocr_end)

    def ocr_start(self, event):
        self.start_x = event.x
        self.start_y = event.y
        self.rect = self.ocr_canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=2)

    def test_ocr_box(self):
        """Capture a test screenshot using the current OCR box coordinates with offsets,
        then save it to the 'ocr_attempt_logs' folder for inspection."""
        # Set offset values (adjust as needed for your display)
        offset_x = 22  # e.g., shift 2 pixels to the right
        offset_y = 0  # adjust y offset if necessary

        if not self.ocr_box:
            messagebox.showwarning("Test OCR Box", "OCR box not set. Please set it first.")
            return

        # Extract the OCR box coordinates (assumed to be absolute screen coordinates)
        x1, y1, x2, y2 = self.ocr_box
        # Apply offsets to the OCR box coordinates
        test_box = (x1 + offset_x, y1 + offset_y, x2 + offset_x, y2 + offset_y)

        # Capture a screenshot of the test area
        screenshot = ImageGrab.grab(bbox=test_box)

        # Ensure the debug folder exists
        debug_folder = "ocr_attempt_logs"
        if not os.path.exists(debug_folder):
            os.makedirs(debug_folder)

        # Save the screenshot with a fixed filename for testing
        filename = os.path.join(debug_folder, "test_ocr_box.png")
        screenshot.save(filename)
        messagebox.showinfo("Test OCR Box", f"Test screenshot saved to {filename}")


    def ocr_move(self, event):
        self.ocr_canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)

    def ocr_end(self, event):
        self.end_x = event.x
        self.end_y = event.y
        # Get the overlay's absolute position on screen
        overlay_x = self.ocr_overlay.winfo_rootx()
        overlay_y = self.ocr_overlay.winfo_rooty()

        # Adjust by a small offset to the right (tweak as needed)
        offset_x = 22  # Increase or decrease this value until the capture aligns correctly

        absolute_start_x = overlay_x + self.start_x + offset_x
        absolute_start_y = overlay_y + self.start_y
        absolute_end_x = overlay_x + self.end_x + offset_x
        absolute_end_y = overlay_y + self.end_y

        self.ocr_box = (absolute_start_x, absolute_start_y, absolute_end_x, absolute_end_y)
        self.ocr_overlay.destroy()
        messagebox.showinfo("OCR Box Set", f"OCR Box set to: {self.ocr_box}")

    def reset_ocr_box(self):
        self.ocr_box = None
        messagebox.showinfo("Reset", "OCR box reset.")

    def read_screen_text(self):
        if self.ocr_box:
            screenshot = ImageGrab.grab(bbox=self.ocr_box)
            text = pytesseract.image_to_string(screenshot, config='--psm 6')
            return text.strip()
        return ""

    ### === MOUSE & CLICK CAPTURE === ###

    def capture_click(self):
        self.click_window = tk.Toplevel(self.root)
        self.click_window.attributes('-fullscreen', True)
        self.click_window.attributes('-alpha', 0.5)
        self.click_window.title("Click to Set Starting Position")
        self.click_window.bind("<Button-1>", self.record_click)
        tk.Label(self.click_window, text="Click anywhere to set the starting position.", bg='white').pack(pady=20)

    def record_click(self, event):
        self.start_x = event.x_root
        self.start_y = event.y_root
        ag.moveTo(self.start_x, self.start_y)
        self.save_settings()
        messagebox.showinfo("Position Set", f"Starting position set to: ({self.start_x}, {self.start_y})")
        self.click_window.destroy()

    ### === RESET FUNCTION === ###

    def reset(self):
        self.last_names = []
        self.current_index = 0
        self.start_x = None
        self.start_y = None
        self.date1 = None
        self.date2 = None
        self.ocr_box = None
        self.skipped_names = []
        self.cal.selection_set(datetime.today())
        for key, var in self.check_vars.items():
            if key in ["Saturday", "Sunday"]:
                var.set(1)
            else:
                var.set(0)
        self.save_settings()
        messagebox.showinfo("Reset", "All fields, OCR box, and selections have been reset.")


if __name__ == "__main__":
    ag.PAUSE = 0  # Reduce pyautogui delays
    GUI()