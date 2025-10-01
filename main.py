import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
import sys
from datetime import datetime

from certificate import CertificateGeneratorPopup
from Admitcard import open_admit_card_window
from attendance import open_attendance_sheet_window
import subprocess

EXCEL_FILE = 'student_data.xlsx'

class StudentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SILPITIRTHA SIKSHA NIKETAN")
        self.admit_cards_dir = os.path.join(os.getcwd(), "ADMIT CARDS")
        self.certificates_dir = os.path.join(os.getcwd(), "CERTIFICATES")
        self.ensure_admit_cards_dir()
        self.ensure_certificates_dir()
        self.columns = [
            'Roll No.', 'Name', 'Guardian Name', 'Address', 'Subject', 'Year',
            'Date of Birth', 'Sex', 'Phone Number'
        ]
        self.student_data = []
        self.sort_column = None
        self.sort_reverse = False
        self.column_filters = {col: "" for col in self.columns}
        self.create_menu()
        self.create_form()
        self.create_buttons()
        self.create_data_view()
        self.create_top_buttons()

    def create_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Admit Cards", command=self.open_admit_cards_folder)
        file_menu.add_command(label="Certificates", command=self.open_certificates_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Help", command=self.show_user_manual)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="User Manual", command=self.show_user_manual)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menubar)

    def create_top_buttons(self):
        top_btn_frame = ttk.Frame(self.root)
        top_btn_frame.place(relx=1.0, y=0, anchor="ne", x=-10)
        attendance_btn = ttk.Button(top_btn_frame, text="Generate Attendance Sheet", command=open_attendance_sheet_window)
        attendance_btn.pack(side="right", padx=5, pady=5)
        generate_results_btn = ttk.Button(top_btn_frame, text="Generate Results", command=self.run_marksheet_py)
        generate_results_btn.pack(side="right", padx=5, pady=5)

    def run_marksheet_py(self):
        try:
            if sys.platform.startswith('win'):
                subprocess.Popen(['python', 'result.py'], shell=True)
            else:
                subprocess.Popen(['python3', 'result.py'])
        except Exception as e:
            messagebox.showerror("Error", f"Could not run result.py: {e}")

    def open_admit_cards_folder(self):
        self.open_folder(self.admit_cards_dir)

    def open_certificates_folder(self):
        self.open_folder(self.certificates_dir)

    def open_folder(self, folder_path):
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            messagebox.showinfo("Folder Created", f"Created folder: {folder_path}")
        try:
            if sys.platform.startswith('win'):
                os.startfile(folder_path)
            elif sys.platform.startswith('darwin'):
                os.system(f'open "{folder_path}"')
            else:
                os.system(f'xdg-open "{folder_path}"')
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open folder: {str(e)}")

    def ensure_admit_cards_dir(self):
        if not os.path.exists(self.admit_cards_dir):
            os.makedirs(self.admit_cards_dir)

    def ensure_certificates_dir(self):
        if not os.path.exists(self.certificates_dir):
            os.makedirs(self.certificates_dir)

    def show_user_manual(self):
        manual = tk.Toplevel(self.root)
        manual.title("User Manual")
        manual.geometry("600x400")
        text = tk.Text(manual, wrap="word", padx=10, pady=10)
        text.pack(fill="both", expand=True)
        manual_text = """
Student Data Management - User Manual

1. Adding a Student
- Fill in all fields: Roll No., Name, Guardian Name, Address, Subject, Year,
Date of Birth, Sex, Phone Number.
- Click "Submit" to save the student to the database.

2. Generating Admit Card
- Click "Generate Admit Card" to open advanced admit card generation.

3. Generating Certificates
- Click "Generate Certificates" to create PDF certificates for all students.
- Certificates are saved in the CERTIFICATES folder.

4. Filtering and Sorting
- Click "Filter" to open the filter dialog.
- Set filters for any column (Name, Subject, Year, etc.).
- Click "Apply" to filter the table, or "Clear" to remove all filters.
- Click on any column header to sort by that column.

5. Exporting Records
- Click "Export Records" to save the student data to CSV or Excel.
- Only filtered/visible data will be exported.

6. File Menu Options
- Admit Cards: Opens the folder containing all admit cards
- Certificates: Opens the folder containing all certificates
- Help: Opens this user manual
- Exit: Closes the application

7. Help and Documentation
- Use the Help menu for this manual and information.
"""
        text.insert("1.0", manual_text)
        text.config(state="disabled")

    def show_about(self):
        messagebox.showinfo(
            "About",
            "Student Data Management\n\n"
            "A desktop application for managing student records, generating admit cards and certificates.\n"
            "Developed By Subham. Email: nanda.subham.001@gmail.com\n\n"
            "Version 1.0"
        )

    def validate_date(self, date_str):
        try:
            date = datetime.strptime(date_str, "%d-%m-%Y")
            if date > datetime.now():
                return False, "Date of Birth cannot be in the future."
            return True, ""
        except ValueError:
            return False, "Date of Birth must be in format DD-MM-YYYY."

    def validate_phone(self, phone_str):
        if not phone_str.isdigit():
            return False, "Phone Number must contain only digits."
        if len(phone_str) != 10:
            return False, "Phone Number must be 10 digits."
        return True, ""

    def create_form(self):
        form_frame = ttk.LabelFrame(self.root, text="Student Information")
        form_frame.pack(fill="x", padx=10, pady=5)
        labels = list(self.columns)
        self.entries = {}
        for i, label in enumerate(labels):
            ttk.Label(form_frame, text=f"{label}:").grid(row=i, column=0, padx=5, pady=5, sticky="e")
            if label == 'Subject':
                subject_dropdown = ttk.Combobox(form_frame, values=[
                    'Fine Arts', 'Dance', 'Hand Craft', 'Beautician',
                    'Recitations', 'Song', 'Musical Instruments'
                ])
                subject_dropdown.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                self.entries[label] = subject_dropdown
            elif label == 'Year':
                year_dropdown = ttk.Combobox(form_frame, values=[
                    'Pr-1', 'Pr-2', 'Pr', '1st', '2nd', '3rd', '4th', '5th', '6th', '7th'
                ])
                year_dropdown.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                self.entries[label] = year_dropdown
            elif label == 'Sex':
                sex_frame = ttk.Frame(form_frame)
                sex_frame.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                self.sex_var = tk.StringVar(value='Male')
                ttk.Radiobutton(sex_frame, text='Male', variable=self.sex_var, value='Male').pack(side='left', padx=5)
                ttk.Radiobutton(sex_frame, text='Female', variable=self.sex_var, value='Female').pack(side='left', padx=5)
                self.entries[label] = self.sex_var
            else:
                entry = ttk.Entry(form_frame)
                entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                self.entries[label] = entry

    def create_buttons(self):
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=10, pady=5)
        buttons = [
            ("Submit", self.submit_data, "Save the current student record"),
            ("Generate Admit Card", self.generate_admit_card, "Generate PDF admit card for selected student"),
            ("Generate Certificates", self.generate_certificates, "Open certificate generator popup"),
            ("Filter", self.show_filter_dialog, "Filter student records by any column"),
        ]
        for text, command, tooltip in buttons:
            btn = ttk.Button(btn_frame, text=text, command=command)
            btn.pack(side="left", padx=5)
            self.create_tooltip(btn, tooltip)
        export_btn = ttk.Button(btn_frame, text="Export Records", command=self.export_records)
        export_btn.pack(side="right", padx=5)
        self.create_tooltip(export_btn, "Export student records to CSV or Excel")

    def create_tooltip(self, widget, text):
        tooltip = tk.Toplevel(self.root)
        tooltip.withdraw()
        tooltip.overrideredirect(True)
        def show_tooltip(event):
            try:
                x, y, _, _ = widget.bbox("insert")
                x += widget.winfo_rootx() + 25
                y += widget.winfo_rooty() + 25
                tooltip.geometry(f"+{x}+{y}")
                tooltip_label.config(text=text)
                tooltip.deiconify()
            except:
                pass
        def hide_tooltip(event):
            tooltip.withdraw()
        tooltip_label = tk.Label(tooltip, text="", background="#ffffe0", relief="solid", borderwidth=1, padx=2, pady=2)
        tooltip_label.pack()
        widget.bind("<Enter>", show_tooltip)
        widget.bind("<Leave>", hide_tooltip)

    def create_data_view(self):
        data_frame = ttk.LabelFrame(self.root, text="Student Records")
        data_frame.pack(fill="both", expand=True, padx=10, pady=5)
        scroll_y = ttk.Scrollbar(data_frame, orient="vertical")
        scroll_x = ttk.Scrollbar(data_frame, orient="horizontal")
        self.tree = ttk.Treeview(data_frame, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)
        self.tree["columns"] = self.columns
        self.tree["show"] = "headings"
        for col in self.columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_tree(c))
            self.tree.column(col, width=120)
        self.tree.pack(fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.load_data()

    def load_data(self):
        if not os.path.exists(EXCEL_FILE):
            df = pd.DataFrame(columns=self.columns)
            df.to_excel(EXCEL_FILE, index=False)
        try:
            df = pd.read_excel(EXCEL_FILE)
            for col in self.columns:
                if col not in df.columns:
                    df[col] = ""
            df = df[self.columns]
            self.student_data = df.values.tolist()
            self.update_treeview()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {e}")
            self.student_data = []
            self.update_treeview()

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        filtered_data = []
        for row in self.student_data:
            match = True
            for i, col in enumerate(self.columns):
                filter_val = self.column_filters[col]
                if filter_val and str(row[i]).lower().find(filter_val.lower()) == -1:
                    match = False
                    break
            if match:
                filtered_data.append(row)
        if self.sort_column is not None:
            col_index = self.columns.index(self.sort_column)
            filtered_data.sort(key=lambda x: str(x[col_index]), reverse=self.sort_reverse)
        for row in filtered_data:
            self.tree.insert("", "end", values=row)

    def show_filter_dialog(self):
        filter_dialog = tk.Toplevel(self.root)
        filter_dialog.title("Filter Records")
        filter_dialog.geometry("400x350")
        filter_entries = {}
        for i, col in enumerate(self.columns):
            ttk.Label(filter_dialog, text=f"{col}:").grid(row=i, column=0, padx=5, pady=5, sticky="e")
            if col == 'Subject':
                entry = ttk.Combobox(filter_dialog, values=[''] + [
                    'Fine Arts', 'Dance', 'Hand Craft', 'Beautician',
                    'Recitations', 'Song', 'Musical Instruments'
                ])
                entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                entry.set(self.column_filters[col])
            elif col == 'Year':
                entry = ttk.Combobox(filter_dialog, values=[''] + [
                    'Pr-1', 'Pr-2', 'Pr', '1st', '2nd', '3rd', '4th', '5th', '6th', '7th'
                ])
                entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                entry.set(self.column_filters[col])
            elif col == 'Sex':
                entry = ttk.Combobox(filter_dialog, values=[''] + ['Male', 'Female'])
                entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                entry.set(self.column_filters[col])
            else:
                entry = ttk.Entry(filter_dialog)
                entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                entry.insert(0, self.column_filters[col])
            filter_entries[col] = entry
        def apply_filters():
            for col in self.columns:
                self.column_filters[col] = filter_entries[col].get()
            self.update_treeview()
            filter_dialog.destroy()
        def clear_filters():
            for col in self.columns:
                if col in ('Subject', 'Year', 'Sex'):
                    filter_entries[col].set('')
                else:
                    filter_entries[col].delete(0, tk.END)
            for col in self.columns:
                self.column_filters[col] = ""
            self.update_treeview()
        ttk.Button(filter_dialog, text="Apply", command=apply_filters).grid(
            row=len(self.columns), column=0, padx=5, pady=10, sticky="e"
        )
        ttk.Button(filter_dialog, text="Clear", command=clear_filters).grid(
            row=len(self.columns), column=1, padx=5, pady=10, sticky="w"
        )

    def sort_tree(self, col):
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = False
        self.update_treeview()

    def submit_data(self):
        data = {}
        for label, entry in self.entries.items():
            if label == 'Sex':
                data[label] = entry.get()
            elif isinstance(entry, (ttk.Combobox, ttk.Entry)):
                data[label] = entry.get()
            else:
                data[label] = entry.get()
        if not all(data.values()):
            messagebox.showerror("Error", "All fields are required!")
            return
        dob = data.get('Date of Birth', '')
        is_valid_dob, dob_error = self.validate_date(dob)
        if not is_valid_dob:
            messagebox.showerror("Error", dob_error)
            return
        phone = data.get('Phone Number', '')
        is_valid_phone, phone_error = self.validate_phone(phone)
        if not is_valid_phone:
            messagebox.showerror("Error", phone_error)
            return
        try:
            if os.path.exists(EXCEL_FILE):
                df_old = pd.read_excel(EXCEL_FILE)
                for col in self.columns:
                    if col not in df_old.columns:
                        df_old[col] = ""
                df_old = df_old[self.columns]
            else:
                df_old = pd.DataFrame(columns=self.columns)
            df_new = pd.DataFrame([data], columns=self.columns)
            df_all = pd.concat([df_old, df_new], ignore_index=True)
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                df_all.to_excel(writer, index=False)
            self.load_data()
            messagebox.showinfo("Success", "Data saved successfully!")
            for label, entry in self.entries.items():
                if label == 'Sex':
                    entry.set('Male')
                elif isinstance(entry, (ttk.Combobox, ttk.Entry)):
                    if isinstance(entry, ttk.Combobox):
                        entry.set('')
                    else:
                        entry.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {str(e)}")

    def generate_admit_card(self):
        open_admit_card_window()

    def export_records(self):
        filtered_rows = []
        for item in self.tree.get_children():
            filtered_rows.append(self.tree.item(item)['values'])
        if not filtered_rows:
            messagebox.showwarning("Warning", "No records to export!")
            return
        df = pd.DataFrame(filtered_rows, columns=self.columns)
        filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        filename = filedialog.asksaveasfilename(
            title="Save student records as...",
            filetypes=filetypes,
            defaultextension=".xlsx"
        )
        if not filename:
            return
        try:
            if filename.lower().endswith('.csv'):
                df.to_csv(filename, index=False)
                messagebox.showinfo("Success", f"Records exported to CSV:\n{filename}")
            else:
                df.to_excel(filename, index=False)
                messagebox.showinfo("Success", f"Records exported to Excel:\n{filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export records: {str(e)}")

    def generate_certificates(self):
        CertificateGeneratorPopup(self.root, template_image_path="Certificate.jpg")

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentApp(root)
    root.mainloop()
