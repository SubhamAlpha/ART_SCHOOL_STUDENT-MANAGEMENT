import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os
from datetime import datetime

class AdmitCardGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Admit Card Generator")
        self.root.geometry("800x600")

        # Data holders
        self.imported_data = None
        self.current_student_index = 0

        # UI Setup
        self.create_widgets()

    def create_widgets(self):
        # Top frame for import button
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=10, pady=5)
        import_btn = ttk.Button(top_frame, text="Import List", command=self.import_list)
        import_btn.pack(side="right", padx=5)

        # Admit Details Frame
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

        # Annual Examination Frame
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

        # Set default values for examination
        self.exam_entries["Date"].insert(0, "23.06.2024")
        self.exam_entries["Time (1st Part)"].insert(0, "8:00 AM to 9:30 AM")
        self.exam_entries["Time (2nd Part)"].insert(0, "X")

        # Navigation and Generate buttons
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
            if label in row:
                self.admit_entries[label].delete(0, tk.END)
                self.admit_entries[label].insert(0, str(row[label]))
            else:
                self.admit_entries[label].delete(0, tk.END)

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

    def generate_admit_card(self):
        if self.imported_data is None and not all(self.admit_entries[label].get() for label in self.admit_labels):
            messagebox.showerror("Error", "Please import a list or fill all Admit Details!")
            return

        admit_data = {label: self.admit_entries[label].get() for label in self.admit_labels}
        exam_data = {label: self.exam_entries[label].get() for label in self.exam_labels}

        answer = messagebox.askyesno(
            "Save Location",
            "Default save location is: 'ADMIT CARDS'\n\nWould you like to choose a different location?"
        )
        admit_cards_dir = os.path.join(os.getcwd(), "ADMIT CARDS")
        if answer:
            directory = filedialog.askdirectory(title="Select Directory for ADMIT CARDS")
            if directory:
                admit_cards_dir = os.path.join(directory, "ADMIT CARDS")

        if not os.path.exists(admit_cards_dir):
            os.makedirs(admit_cards_dir)

        roll = admit_data["Roll No."].replace("/", "_")
        name = admit_data["Name"].replace(" ", "_")
        filename = f"{roll}_{name}_admit_card.pdf"
        filepath = os.path.join(admit_cards_dir, filename)

        try:
            c = canvas.Canvas(filepath, pagesize=A4)
            c.setFont("Helvetica-Bold", 16)
            c.drawString(100, 800, "ADMIT CARD")
            c.line(100, 790, 500, 790)
            c.setFont("Helvetica", 12)

            y_position = 750
            for label in self.admit_labels:
                c.drawString(100, y_position, f"{label}: {admit_data[label]}")
                y_position -= 30

            y_position -= 20
            c.setFont("Helvetica-Bold", 14)
            c.drawString(100, y_position, "Annual Examination")
            c.setFont("Helvetica", 12)
            y_position -= 30

            for label in self.exam_labels:
                c.drawString(100, y_position, f"{label}: {exam_data[label]}")
                y_position -= 30

            c.save()
            messagebox.showinfo("Success", f"Admit card generated as:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PDF: {str(e)}")

    def generate_all_admit_cards(self):
        if self.imported_data is None:
            messagebox.showwarning("Warning", "No list imported!")
            return

        answer = messagebox.askyesno(
            "Save Location",
            "Default save location is: 'ADMIT CARDS'\n\nWould you like to choose a different location?"
        )
        admit_cards_dir = os.path.join(os.getcwd(), "ADMIT CARDS")
        if answer:
            directory = filedialog.askdirectory(title="Select Directory for ADMIT CARDS")
            if directory:
                admit_cards_dir = os.path.join(directory, "ADMIT CARDS")

        if not os.path.exists(admit_cards_dir):
            os.makedirs(admit_cards_dir)

        exam_data = {label: self.exam_entries[label].get() for label in self.exam_labels}

        for i in range(len(self.imported_data)):
            row = self.imported_data.iloc[i]
            admit_data = {}
            for label in self.admit_labels:
                admit_data[label] = str(row[label]) if label in row else ""

            roll = admit_data["Roll No."].replace("/", "_")
            name = admit_data["Name"].replace(" ", "_")
            filename = f"{roll}_{name}_admit_card.pdf"
            filepath = os.path.join(admit_cards_dir, filename)

            try:
                c = canvas.Canvas(filepath, pagesize=A4)
                c.setFont("Helvetica-Bold", 16)
                c.drawString(100, 800, "ADMIT CARD")
                c.line(100, 790, 500, 790)
                c.setFont("Helvetica", 12)

                y_position = 750
                for label in self.admit_labels:
                    c.drawString(100, y_position, f"{label}: {admit_data[label]}")
                    y_position -= 30

                y_position -= 20
                c.setFont("Helvetica-Bold", 14)
                c.drawString(100, y_position, "Annual Examination")
                c.setFont("Helvetica", 12)
                y_position -= 30

                for label in self.exam_labels:
                    c.drawString(100, y_position, f"{label}: {exam_data[label]}")
                    y_position -= 30

                c.save()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to generate PDF for {name}: {str(e)}")
                continue

        messagebox.showinfo("Success", "All admit cards have been generated!")

def open_admit_card_window():
    root = tk.Toplevel()
    AdmitCardGenerator(root)
    root.mainloop()
