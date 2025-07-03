import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import re

ADMIT_TEMPLATE = "admit.jpg"  # Path to your admit card template image

class AdmitCardGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Admit Card Generator")
        self.root.geometry("800x600")
        self.imported_data = None
        self.current_student_index = 0
        self.create_widgets()

    def create_widgets(self):
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=10, pady=5)
        import_btn = ttk.Button(top_frame, text="Import List", command=self.import_list)
        import_btn.pack(side="right", padx=5)

        admit_frame = ttk.LabelFrame(self.root, text="ADMIT DETAILS")
        admit_frame.pack(fill="x", padx=10, pady=5)
        self.admit_labels = [
            "Roll No.", "Name", "Examination for", "Year", "Subject", "Name of the Centre with Address"
        ]
        self.admit_entries = {}
        for i, label in enumerate(self.admit_labels):
            ttk.Label(admit_frame, text=f"{label}:").grid(row=i, column=0, padx=5, pady=5, sticky="e")
            entry = ttk.Entry(admit_frame)
            entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
            self.admit_entries[label] = entry

        exam_frame = ttk.LabelFrame(self.root, text="Annual Examination")
        exam_frame.pack(fill="x", padx=10, pady=5)
        self.exam_labels = [
            "Date", "Time (1st Part)", "Time (2nd Part)", "Place"
        ]
        self.exam_entries = {}
        for i, label in enumerate(self.exam_labels):
            ttk.Label(exam_frame, text=f"{label}:").grid(row=i, column=0, padx=5, pady=5, sticky="e")
            entry = ttk.Entry(exam_frame)
            entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
            self.exam_entries[label] = entry

        self.exam_entries["Date"].insert(0, "23.06.2024")
        self.exam_entries["Time (1st Part)"].insert(0, "8:00 AM to 9:30 AM")
        self.exam_entries["Time (2nd Part)"].insert(0, "X")

        nav_frame = ttk.Frame(self.root)
        nav_frame.pack(fill="x", padx=10, pady=5)
        prev_btn = ttk.Button(nav_frame, text="Previous", command=self.prev_student)
        prev_btn.pack(side="left", padx=5)
        next_btn = ttk.Button(nav_frame, text="Next", command=self.next_student)
        next_btn.pack(side="left", padx=5)
        generate_btn = ttk.Button(nav_frame, text="Generate Admit Card", command=self.generate_admit_card)
        generate_btn.pack(side="right", padx=5)
        generate_all_btn = ttk.Button(nav_frame, text="Generate All Admit Cards", command=self.generate_all_admit_cards)
        generate_all_btn.pack(side="right", padx=5)

    def import_list(self):
        filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        filename = filedialog.askopenfilename(title="Select file to import", filetypes=filetypes)
        if not filename:
            return
        try:
            if filename.lower().endswith('.csv'):
                df = pd.read_csv(filename)
            else:
                df = pd.read_excel(filename)
            self.imported_data = df
            self.current_student_index = 0
            self.load_student_data()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import file: {str(e)}")

    def load_student_data(self):
        if self.imported_data is None or self.current_student_index >= len(self.imported_data):
            return
        row = self.imported_data.iloc[self.current_student_index]
        for label in self.admit_labels:
            value = str(row[label]) if label in row else ""
            self.admit_entries[label].delete(0, tk.END)
            self.admit_entries[label].insert(0, value)

    def next_student(self):
        if self.imported_data is None:
            messagebox.showwarning("Warning", "No list imported!")
            return
        if self.current_student_index < len(self.imported_data) - 1:
            self.current_student_index += 1
            self.load_student_data()

    def prev_student(self):
        if self.imported_data is None:
            messagebox.showwarning("Warning", "No list imported!")
            return
        if self.current_student_index > 0:
            self.current_student_index -= 1
            self.load_student_data()

    def validate_fields(self, admit_data, exam_data):
        if not admit_data["Roll No."] or not re.match(r"^[\w/]+$", admit_data["Roll No."]):
            return False, "Invalid Roll No. (should be non-empty and alphanumeric)"
        if not admit_data["Name"] or not re.match(r"^[A-Za-z .'-]+$", admit_data["Name"]):
            return False, "Invalid Name (should contain only letters and spaces)"
        if not admit_data["Examination for"]:
            return False, "Examination for field cannot be empty"
        if not admit_data["Year"]:
            return False, "Year field cannot be empty"
        if not admit_data["Subject"]:
            return False, "Subject field cannot be empty"
        if not admit_data["Name of the Centre with Address"]:
            return False, "Centre/Address field cannot be empty"
        if not re.match(r"^\d{2}\.\d{2}\.\d{4}$", exam_data["Date"]):
            return False, "Date must be in DD.MM.YYYY format"
        if not exam_data["Time (1st Part)"]:
            return False, "Time (1st Part) cannot be empty"
        if not exam_data["Time (2nd Part)"]:
            return False, "Time (2nd Part) cannot be empty"
        if not exam_data["Place"]:
            return False, "Place field cannot be empty"
        return True, ""

    def generate_admit_card(self):
        admit_data = {label: self.admit_entries[label].get().strip() for label in self.admit_labels}
        exam_data = {label: self.exam_entries[label].get().strip() for label in self.exam_labels}
        valid, msg = self.validate_fields(admit_data, exam_data)
        if not valid:
            messagebox.showerror("Validation Error", msg)
            return
        self.save_admit_card_image(admit_data, exam_data)
        messagebox.showinfo("Success", "Admit card generated successfully as JPG.")

    def generate_all_admit_cards(self):
        if self.imported_data is None:
            messagebox.showwarning("Warning", "No list imported!")
            return
        exam_data = {label: self.exam_entries[label].get().strip() for label in self.exam_labels}
        for idx, row in self.imported_data.iterrows():
            admit_data = {label: str(row[label]) if label in row else "" for label in self.admit_labels}
            valid, msg = self.validate_fields(admit_data, exam_data)
            if not valid:
                messagebox.showerror("Validation Error", f"Row {str(idx)}: {msg}")
                continue
            self.save_admit_card_image(admit_data, exam_data)
        messagebox.showinfo("Success", "All admit cards generated as JPG.")

    def save_admit_card_image(self, admit_data, exam_data):
        if not os.path.exists(ADMIT_TEMPLATE):
            messagebox.showerror("Error", f"Template image '{ADMIT_TEMPLATE}' not found.")
            return
        img = Image.open(ADMIT_TEMPLATE).convert("RGB")
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype("arial.ttf", 22)
            font_bold = ImageFont.truetype("arialbd.ttf", 24)
        except:
            font = ImageFont.load_default()
            font_bold = font

        # --- Alignment based on your marked reference image ---
        # Roll No. (top right, red)
        roll_text = admit_data["Roll No."]
        roll_x, roll_y = 320, 91
        draw.text((roll_x, roll_y), roll_text, font=font_bold, fill="black")

        # Name (Sri/Sm./Km.) (left, black)
        name_text = admit_data["Name"]
        name_x, name_y = 299, 140
        draw.text((name_x, name_y), name_text, font=font, fill="black")

        # Examination for (left, black)
        exam_for_text = admit_data["Examination for"]
        exam_for_x, exam_for_y = 311, 187
        draw.text((exam_for_x, exam_for_y), exam_for_text, font=font, fill="black")

        # Year (left, black)
        year_text = admit_data["Year"]
        year_x, year_y = 170, 237
        draw.text((year_x, year_y), year_text, font=font, fill="black")

        # Subject (left, black)
        subject_text = admit_data["Subject"]
        subject_x, subject_y = 420, 239
        draw.text((subject_x, subject_y), subject_text, font=font, fill="black")

        # Name of the Centre with Address (left, black)
        centre_text = admit_data["Name of the Centre with Address"]
        centre_x, centre_y = 140, 323
        draw.text((centre_x, centre_y), centre_text, font=font, fill="black")

        # Date (right, black)
        date_text = exam_data["Date"]
        date_x, date_y = 760, 130
        draw.text((date_x, date_y), date_text, font=font, fill="black")

        # Time (1st Part) (right, black)
        time1_text = exam_data["Time (1st Part)"]
        time1_x, time1_y = 780, 209
        draw.text((time1_x, time1_y), time1_text, font=font, fill="black")

        # Time (2nd Part) (right, black)
        time2_text = exam_data["Time (2nd Part)"]
        time2_x, time2_y = 780, 260
        draw.text((time2_x, time2_y), time2_text, font=font, fill="black")

        # Place (right, black)
        place_text = exam_data["Place"]
        place_x, place_y = 720, 350
        draw.text((place_x, place_y), place_text, font=font, fill="black")

        out_dir = os.path.join(os.getcwd(), "ADMIT CARDS")
        if not os.path.exists(out_dir):
            os.makedirs(out_dir)
        roll = admit_data["Roll No."].replace("/", "_")
        name = admit_data["Name"].replace(" ", "_")
        out_path = os.path.join(out_dir, f"{roll}_{name}_admit_card.jpg")
        img.save(out_path, "JPEG")

def open_admit_card_window():
    root = tk.Toplevel()
    AdmitCardGenerator(root)
    root.mainloop()
