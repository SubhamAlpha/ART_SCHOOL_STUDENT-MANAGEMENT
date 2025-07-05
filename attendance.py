import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os

# Path to the attendance sheet template image
ATTENDANCE_TEMPLATE = "attendance_sheet.jpg"  # Use your provided template

# Coordinates for each field (adjust as needed for your template)
# The y-coordinates are for the first row; each subsequent row is offset by ROW_HEIGHT
FIELD_COORDS = {
    "Roll No.": (120, 590),   # x, y for Roll No. cell (first row)
    "Name": (381, 590),      # x, y for Name cell (first row)
    "Year": (1163, 590),      # x, y for Year cell (first row) -- adjust as needed
}
ROW_HEIGHT = 106  # Vertical space between rows; adjust as per your template
MAX_ROWS_PER_SHEET = 26  # Number of students per sheet; adjust as per your template

# Font settings (adjust path/size as needed)
try:
    FONT = ImageFont.truetype("arial.ttf", 48)
except:
    FONT = ImageFont.load_default()

class AttendanceSheetApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Sheet Generator")
        self.root.geometry("500x180")
        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Select Student Data Excel/CSV File:").pack(anchor="w")
        entry_frame = ttk.Frame(frame)
        entry_frame.pack(fill="x", pady=(5, 0))
        self.file_entry = ttk.Entry(entry_frame, width=40)
        self.file_entry.pack(side="left", padx=(0, 5), fill="x", expand=True)
        browse_btn = ttk.Button(entry_frame, text="Browse", command=self.browse_file)
        browse_btn.pack(side="left")

        frame2 = ttk.Frame(self.root, padding=20)
        frame2.pack(fill="x")
        ttk.Button(frame2, text="Generate Attendance Sheets", command=self.generate_sheets).pack()

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel/CSV File",
            filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.csv")]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def generate_sheets(self):
        file_path = self.file_entry.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "Please select a valid Excel or CSV file.")
            return

        # Read required columns from Excel/CSV
        try:
            if file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {e}")
            return

        # Map columns (case-insensitive, robust to naming)
        col_map = {}
        for col in df.columns:
            col_lower = col.strip().lower()
            if "roll" in col_lower:
                col_map["Roll No."] = col
            elif "name" in col_lower and "guardian" not in col_lower and "father" not in col_lower:
                col_map["Name"] = col
            elif "year" in col_lower:
                col_map["Year"] = col

        required_fields = list(FIELD_COORDS.keys())
        missing = [k for k in required_fields if k not in col_map]
        if missing:
            messagebox.showerror("Error", f"Missing columns in file: {', '.join(missing)}")
            return

        # Output directory
        out_dir = os.path.join(os.getcwd(), "ATTENDANCE_SHEETS")
        os.makedirs(out_dir, exist_ok=True)

        total_students = len(df)
        sheet_count = 0
        for start_idx in range(0, total_students, MAX_ROWS_PER_SHEET):
            img = Image.open(ATTENDANCE_TEMPLATE).convert("RGB")
            draw = ImageDraw.Draw(img)
            end_idx = min(start_idx + MAX_ROWS_PER_SHEET, total_students)
            for row_num, idx in enumerate(range(start_idx, end_idx)):
                row = df.iloc[idx]
                y_offset = row_num * ROW_HEIGHT
                # Draw Roll No., Name, and Year in perfect alignment
                roll_value = "" if pd.isna(row[col_map["Roll No."]]) else str(row[col_map["Roll No."]])
                name_value = "" if pd.isna(row[col_map["Name"]]) else str(row[col_map["Name"]])
                year_value = "" if pd.isna(row[col_map["Year"]]) else str(row[col_map["Year"]])
                draw.text((FIELD_COORDS["Roll No."][0], FIELD_COORDS["Roll No."][1] + y_offset), roll_value, font=FONT, fill="black")
                draw.text((FIELD_COORDS["Name"][0], FIELD_COORDS["Name"][1] + y_offset), name_value, font=FONT, fill="black")
                draw.text((FIELD_COORDS["Year"][0], FIELD_COORDS["Year"][1] + y_offset), year_value, font=FONT, fill="black")
            # Save the sheet
            sheet_count += 1
            out_path = os.path.join(out_dir, f"attendance_sheet_{sheet_count}.jpg")
            img.save(out_path, "JPEG")

        messagebox.showinfo("Success", f"{sheet_count} attendance sheet(s) generated in {out_dir}")

def open_attendance_sheet_window():
    win = tk.Toplevel()
    AttendanceSheetApp(win)
    win.mainloop()
