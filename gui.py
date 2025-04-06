import calendar
import os
import sys
import tkinter as tk
import webbrowser
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from PIL import Image, ImageTk

# Import the main functionality from main.py
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from main import generate_calendar_html


class PlannerCalendarGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Planner Calendar Generator")
        self.root.geometry("700x600")
        self.root.resizable(True, True)

        # Icon
        icon_path = 'planner-calendar.ico' # Or .ico, .jpg, etc.

        # Check if the file exists before trying to load
        if os.path.exists(icon_path):
            try:
                # Open the image using Pillow
                img = Image.open(icon_path)
                # Convert it to a Tkinter PhotoImage object
                icon_image = ImageTk.PhotoImage(img)

                # Set the window icon using iconphoto
                self.root.iconphoto(False, icon_image)

            except Exception as e:
                print(f"Error loading icon with Pillow: {e}")
                print(f"Make sure '{icon_path}' exists, is a valid image file, and Pillow is installed.")
        else:
            print(f"Icon file not found: {icon_path}")

        # Variables
        self.excel_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar(value="planner_calendar.html")
        self.year = tk.StringVar()
        self.month = tk.StringVar(value="")
        self.no_wrap_text = tk.BooleanVar(value=False)
        self.color_saturation = tk.DoubleVar(value=0.7)
        self.color_lightness = tk.DoubleVar(value=0.85)
        self.color_by_label = tk.BooleanVar(value=False)
        self.color_by_bucket = tk.BooleanVar(value=False)
        self.prefix_labels = tk.BooleanVar(value=False)
        self.alternate_colors = tk.BooleanVar(value=False)
        
        # Set current year as default
        current_year = datetime.now().year
        self.year.set(str(current_year))
        
        # Create the main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create a notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Create tabs
        self.basic_tab = ttk.Frame(self.notebook)
        self.advanced_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.basic_tab, text="Basic Options")
        self.notebook.add(self.advanced_tab, text="Advanced Options")
        
        # Setup the UI components
        self._setup_basic_tab()
        self._setup_advanced_tab()
        
        # Add the generate button at the bottom
        generate_frame = ttk.Frame(main_frame)
        generate_frame.pack(fill=tk.X, pady=10)
        
        generate_btn = ttk.Button(
            generate_frame, 
            text="Generate Calendar", 
            command=self.generate_calendar,
            style="Generate.TButton"
        )
        generate_btn.pack(side=tk.RIGHT, padx=5)
        
        view_btn = ttk.Button(
            generate_frame, 
            text="Generate & View", 
            command=self.generate_and_view_calendar,
            style="Generate.TButton"
        )
        view_btn.pack(side=tk.RIGHT, padx=5)
        
        # Create custom styles
        self._setup_styles()
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(
            root, 
            textvariable=self.status_var, 
            relief=tk.SUNKEN, 
            anchor=tk.W
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _setup_styles(self):
        """Setup custom styles for the UI"""
        style = ttk.Style()
        style.configure("Generate.TButton", font=("Arial", 10, "bold"))
        style.configure("TLabel", padding=2)
        style.configure("Header.TLabel", font=("Arial", 11, "bold"))
        style.configure("TEntry", padding=5)
        
    def _setup_basic_tab(self):
        """Setup the basic options tab"""
        # File selection frame
        file_frame = ttk.LabelFrame(self.basic_tab, text="File Selection", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Excel file selection
        ttk.Label(file_frame, text="Excel File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        excel_entry = ttk.Entry(file_frame, textvariable=self.excel_file_path, width=50)
        excel_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        excel_btn = ttk.Button(file_frame, text="Browse...", command=self.browse_excel_file)
        excel_btn.grid(row=0, column=2, padx=5, pady=5)
        
        # Output file selection
        ttk.Label(file_frame, text="Output HTML:").grid(row=1, column=0, sticky=tk.W, pady=5)
        output_entry = ttk.Entry(file_frame, textvariable=self.output_file_path, width=50)
        output_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        output_btn = ttk.Button(file_frame, text="Browse...", command=self.browse_output_file)
        output_btn.grid(row=1, column=2, padx=5, pady=5)
        
        # Basic options frame
        options_frame = ttk.LabelFrame(self.basic_tab, text="Basic Options", padding=10)
        options_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Year selection
        ttk.Label(options_frame, text="Year:").grid(row=0, column=0, sticky=tk.W, pady=5)
        year_entry = ttk.Entry(options_frame, textvariable=self.year, width=10)
        year_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(options_frame, text="(Leave empty to use earliest year from data)").grid(
            row=0, column=2, sticky=tk.W, pady=5
        )
        
        # Month selection
        ttk.Label(options_frame, text="Month:").grid(row=1, column=0, sticky=tk.W, pady=5)
        month_values = [""] + list(calendar.month_name)[1:]
        month_dropdown = ttk.Combobox(options_frame, textvariable=self.month, values=month_values, width=15)
        month_dropdown.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(options_frame, text="(Leave empty for full year)").grid(
            row=1, column=2, sticky=tk.W, pady=5
        )
        
        # No wrap text option
        no_wrap_check = ttk.Checkbutton(
            options_frame, 
            text="Do not wrap task names", 
            variable=self.no_wrap_text
        )
        no_wrap_check.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Prefix labels option
        prefix_labels_check = ttk.Checkbutton(
            options_frame, 
            text="Prefix task names with their labels", 
            variable=self.prefix_labels
        )
        prefix_labels_check.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)

    def _setup_advanced_tab(self):
        """Setup the advanced options tab"""
        # Color options frame
        color_frame = ttk.LabelFrame(self.advanced_tab, text="Color Options", padding=10)
        color_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Color saturation
        ttk.Label(color_frame, text="Color Saturation (0.0-1.0):").grid(row=0, column=0, sticky=tk.W, pady=5)
        saturation_scale = ttk.Scale(
            color_frame, 
            from_=0.0, 
            to=1.0, 
            variable=self.color_saturation, 
            orient=tk.HORIZONTAL,
            length=200
        )
        saturation_scale.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        saturation_entry = ttk.Entry(color_frame, textvariable=self.color_saturation, width=5)
        saturation_entry.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        # Color lightness
        ttk.Label(color_frame, text="Color Lightness (0.0-1.0):").grid(row=1, column=0, sticky=tk.W, pady=5)
        lightness_scale = ttk.Scale(
            color_frame, 
            from_=0.0, 
            to=1.0, 
            variable=self.color_lightness, 
            orient=tk.HORIZONTAL,
            length=200
        )
        lightness_scale.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        lightness_entry = ttk.Entry(color_frame, textvariable=self.color_lightness, width=5)
        lightness_entry.grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        
        # Color grouping options
        color_group_frame = ttk.LabelFrame(self.advanced_tab, text="Color Grouping", padding=10)
        color_group_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Color by label option
        color_by_label_check = ttk.Checkbutton(
            color_group_frame, 
            text="Use task label for color generation", 
            variable=self.color_by_label,
            command=self.handle_color_by_change
        )
        color_by_label_check.grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # Color by bucket option
        color_by_bucket_check = ttk.Checkbutton(
            color_group_frame, 
            text="Use task bucket for color generation", 
            variable=self.color_by_bucket,
            command=self.handle_color_by_change
        )
        color_by_bucket_check.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # Alternate colors option
        alternate_colors_check = ttk.Checkbutton(
            color_group_frame, 
            text="Use alternate colors for tasks", 
            variable=self.alternate_colors
        )
        alternate_colors_check.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        # Help text
        help_frame = ttk.LabelFrame(self.advanced_tab, text="Help", padding=10)
        help_frame.pack(fill=tk.X, padx=10, pady=5)
        
        help_text = (
            "Color by Label: Tasks with the same label will have the same color.\n"
            "Color by Bucket: Tasks in the same bucket will have the same color.\n"
            "Alternate Colors: Uses a sequential color scheme instead of hash-based.\n"
            "Note: You cannot use both 'Color by Label' and 'Color by Bucket' at the same time."
        )
        
        help_label = ttk.Label(help_frame, text=help_text, wraplength=600, justify=tk.LEFT)
        help_label.pack(fill=tk.X, pady=5)

    def handle_color_by_change(self):
        """Handle mutual exclusivity of color by label and color by bucket"""
        if self.color_by_label.get() and self.color_by_bucket.get():
            if self.notebook.index(self.notebook.select()) == 1:  # If on advanced tab
                self.color_by_bucket.set(False)
            else:
                self.color_by_label.set(False)

    def browse_excel_file(self):
        """Open file dialog to select Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_file_path.set(file_path)
            # Set default output filename based on input filename
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            self.output_file_path.set(f"{base_name}_calendar.html")

    def browse_output_file(self):
        """Open file dialog to select output HTML file"""
        file_path = filedialog.asksaveasfilename(
            title="Save HTML File",
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
            initialfile=self.output_file_path.get()
        )
        if file_path:
            self.output_file_path.set(file_path)

    def validate_inputs(self):
        """Validate user inputs before generating calendar"""
        # Check if Excel file is provided and exists
        excel_file = self.excel_file_path.get()
        if not excel_file:
            messagebox.showerror("Error", "Please select an Excel file.")
            return False
        
        if not os.path.isfile(excel_file):
            messagebox.showerror("Error", f"Excel file not found: {excel_file}")
            return False
        
        # Validate year if provided
        year_str = self.year.get().strip()
        if year_str:
            try:
                year = int(year_str)
                if year < 1900 or year > 2100:
                    messagebox.showerror("Error", "Year must be between 1900 and 2100.")
                    return False
            except ValueError:
                messagebox.showerror("Error", "Year must be a valid integer.")
                return False
                
        # Validate month if provided
        month_str = self.month.get().strip()
        if month_str:
            # If it's a month name, it's valid as long as it's in the dropdown
            # which only contains valid month names
            pass
        
        # Validate color values
        saturation = self.color_saturation.get()
        lightness = self.color_lightness.get()
        
        if saturation < 0 or saturation > 1:
            messagebox.showerror("Error", "Color saturation must be between 0.0 and 1.0.")
            return False
            
        if lightness < 0 or lightness > 1:
            messagebox.showerror("Error", "Color lightness must be between 0.0 and 1.0.")
            return False
        
        # Check for mutual exclusivity of color options
        if self.color_by_label.get() and self.color_by_bucket.get():
            messagebox.showerror(
                "Error", 
                "You cannot use both 'Color by Label' and 'Color by Bucket' at the same time."
            )
            return False
            
        return True

    def generate_calendar(self):
        """Generate the calendar based on user inputs"""
        if not self.validate_inputs():
            return
        
        try:
            self.status_var.set("Reading Excel file...")
            self.root.update_idletasks()
            
            # Read Excel file
            try:
                tasks_df = pd.read_excel(
                    self.excel_file_path.get(), 
                    sheet_name="Tasks", 
                    engine="openpyxl"
                )
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read Excel file: {str(e)}")
                self.status_var.set("Ready")
                return
            
            # Check required columns
            required_columns = ["Task Name", "Start date", "Due date"]
            missing_cols = [col for col in required_columns if col not in tasks_df.columns]
            if missing_cols:
                messagebox.showerror(
                    "Error", 
                    f"Missing required columns in 'Tasks' sheet: {', '.join(missing_cols)}"
                )
                self.status_var.set("Ready")
                return
            
            # Generate task colors
            self.status_var.set("Generating task colors...")
            self.root.update_idletasks()
            
            unique_tasks = sorted(tasks_df["Task Name"].astype(str).unique())
            task_colors = {}
            
            import colorsys
            import hashlib
            
            # Constants
            HUE_STEP = 29  # From main.py
            
            for task_index, task_name in enumerate(unique_tasks):
                hue = None
                
                # Color by label logic
                if self.color_by_label.get():
                    matching_rows = tasks_df.loc[tasks_df["Task Name"] == task_name]
                    if not matching_rows.empty:
                        label = matching_rows["Labels"].iloc[0]
                        if pd.notna(label):
                            try:
                                hash_object = hashlib.md5(str(label).encode())
                                hash_digest = hash_object.hexdigest()
                                if self.alternate_colors.get():
                                    hue = (int(HUE_STEP * task_index) % 360) / 360.0
                                else:
                                    hue = (int(hash_digest, 16) % 360) / 360.0
                            except AttributeError:
                                pass
                
                # Color by bucket logic
                if self.color_by_bucket.get():
                    matching_rows = tasks_df.loc[tasks_df["Task Name"] == task_name]
                    if not matching_rows.empty:
                        bucket = matching_rows["Bucket Name"].iloc[0]
                        if pd.notna(bucket):
                            try:
                                hash_object = hashlib.md5(str(bucket).encode())
                                hash_digest = hash_object.hexdigest()
                                if self.alternate_colors.get():
                                    hue = (int(HUE_STEP * task_index) % 360) / 360.0
                                else:
                                    hue = (int(hash_digest, 16) % 360) / 360.0
                            except AttributeError:
                                pass
                
                # If hue wasn't set by label or bucket
                if hue is None:
                    hash_object = hashlib.md5(str(task_name).encode())
                    hash_digest = hash_object.hexdigest()
                    hue = (int(hash_digest, 16) % 360) / 360.0
                
                # Generate color
                rgb = colorsys.hls_to_rgb(hue, self.color_lightness.get(), self.color_saturation.get())
                hex_color = f"#{int(rgb[0] * 255):02x}{int(rgb[1] * 255):02x}{int(rgb[2] * 255):02x}"
                task_colors[task_name] = hex_color
            
            # Determine the year for the calendar
            target_year = None
            year_str = self.year.get().strip()
            if year_str:
                target_year = int(year_str)
            else:
                # Find the earliest year from valid start dates
                valid_start_dates = pd.to_datetime(tasks_df["Start date"], errors="coerce").dropna()
                if not valid_start_dates.empty:
                    target_year = valid_start_dates.min().year
                else:
                    target_year = datetime.now().year
            
            # Determine if a specific month is selected
            target_month = None
            month_str = self.month.get().strip()
            if month_str:
                target_month = month_str
            
            # Update status message
            if target_month:
                self.status_var.set(f"Generating calendar for {target_month} {target_year}...")
            else:
                self.status_var.set(f"Generating calendar for {target_year}...")
            self.root.update_idletasks()
            
            # Generate HTML
            html_content = generate_calendar_html(
                tasks_df, 
                self.no_wrap_text.get(), 
                target_year, 
                task_colors,
                self.prefix_labels.get(),
                target_month
            )
            
            # Write to file
            with open(self.output_file_path.get(), "w", encoding="utf-8") as f:
                f.write(html_content)
            
            self.status_var.set(f"Calendar generated successfully: {self.output_file_path.get()}")
            
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_var.set("Error occurred")
            return False

    def generate_and_view_calendar(self):
        """Generate the calendar and open it in the default web browser"""
        if self.generate_calendar():
            try:
                self.status_var.set("Opening calendar in web browser...")
                webbrowser.open(f"file://{os.path.abspath(self.output_file_path.get())}")
                self.status_var.set("Calendar opened in web browser")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open browser: {str(e)}")
                self.status_var.set("Failed to open browser")


def main():
    root = tk.Tk()
    PlannerCalendarGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
