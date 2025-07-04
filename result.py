import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os

# Path to the result sheet template image
RESULT_SHEET_TEMPLATE = "result_sheet.jpg"  # Use your provided template

# Coordinates for each field (adjust as needed for your template)
# The y-coordinates are for the first row; each subsequent row is offset by ROW_HEIGHT
FIELD_COORDS = {
    "Roll No.": (123, 603),
    "Name of the Candidate": (380, 590),
    "Name of the Guardian": (920, 590),
    "Year": (1420, 600),
}
ROW_HEIGHT = 108 # Vertical space between rows; adjust as per your template
MAX_ROWS_PER_SHEET = 26  # Number of students per sheet; adjust as per your template

# Font settings (adjust path/size as needed)
try:
    FONT = ImageFont.truetype("arial.ttf", 52)
except:
    FONT = ImageFont.load_default()

class ResultSheetApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Result Sheet Generator")
        self.root.geometry("500x180")
        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Select Student Data Excel File:").pack(anchor="w")
        entry_frame = ttk.Frame(frame)
        entry_frame.pack(fill="x", pady=(5, 0))
        self.file_entry = ttk.Entry(entry_frame, width=40)
        self.file_entry.pack(side="left", padx=(0, 5), fill="x", expand=True)
        browse_btn = ttk.Button(entry_frame, text="Browse", command=self.browse_file)
        browse_btn.pack(side="left")

        frame2 = ttk.Frame(self.root, padding=20)
        frame2.pack(fill="x")
        ttk.Button(frame2, text="Generate Result Sheets", command=self.generate_sheets).pack()

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def generate_sheets(self):
        excel_path = self.file_entry.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Error", "Please select a valid Excel file.")
            return

        # Read required columns from Excel
        try:
            df = pd.read_excel(excel_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")
            return

        # Map columns (case-insensitive, robust to naming)
        col_map = {}
        for col in df.columns:
            col_lower = col.strip().lower()
            if "roll" in col_lower:
                col_map["Roll No."] = col
            elif "name" in col_lower and "guardian" not in col_lower and "father" not in col_lower:
                col_map["Name of the Candidate"] = col
            elif "guardian" in col_lower or "father" in col_lower:
                col_map["Name of the Guardian"] = col
            elif "year" in col_lower:
                col_map["Year"] = col

        required_fields = list(FIELD_COORDS.keys())
        missing = [k for k in required_fields if k not in col_map]
        if missing:
            messagebox.showerror("Error", f"Missing columns in Excel: {', '.join(missing)}")
            return

        # Output directory
        out_dir = os.path.join(os.getcwd(), "RESULT_SHEETS")
        os.makedirs(out_dir, exist_ok=True)

        total_students = len(df)
        sheet_count = 0
        for start_idx in range(0, total_students, MAX_ROWS_PER_SHEET):
            img = Image.open(RESULT_SHEET_TEMPLATE).convert("RGB")
            draw = ImageDraw.Draw(img)
            end_idx = min(start_idx + MAX_ROWS_PER_SHEET, total_students)
            for row_num, idx in enumerate(range(start_idx, end_idx)):
                row = df.iloc[idx]
                y_offset = row_num * ROW_HEIGHT
                for field, (x, y) in FIELD_COORDS.items():
                    cell_value = row[col_map[field]]
                    value = "" if pd.isnull(cell_value) else str(cell_value) # type: ignore
                    draw.text((x, y + y_offset), value, font=FONT, fill="black")
            # Save the sheet
            sheet_count += 1
            out_path = os.path.join(out_dir, f"result_sheet_{sheet_count}.jpg")
            img.save(out_path, "JPEG")

        messagebox.showinfo("Success", f"{sheet_count} result sheet(s) generated in {out_dir}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ResultSheetApp(root)
    root.mainloop()
