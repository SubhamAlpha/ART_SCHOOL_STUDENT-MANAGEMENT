# ART_SCHOOL_STUDENT-MANAGEMENT

Art School Student Management System
An automated administrative toolkit built with Python to streamline student record management and document generation for educational institutions.

🚀 Overview
Managing student data manually can be time-consuming and prone to errors. This project automates the creation of professional, standardized documents (Admit Cards, Certificates, and Result Sheets) by integrating student data directly from Excel files into pre-designed graphic templates.

🛠️ Key Features
Admit Card Generator: Automatically populates student details onto examination admit card templates.

Certificate Automation: Generates personalized course completion or achievement certificates.

Result Management: Processes academic data to create structured result sheets.

Attendance Tracker: Generates formatted attendance logs for classroom management.

Excel Integration: Seamlessly reads from .xlsx files to handle bulk student data efficiently.

📂 Repository Structure
main.py: The central entry point for the management system.
also added a help section for description of the product.

Admitcard.py, certificate.py, result.py, attendance.py: Specialized modules for generating specific document types.

student_data.xlsx / students.xlsx: Excel templates used as the data source for automation.

.jpg Files: Graphic templates (e.g., AdmitCard.jpg, certificate.jpg) used as backgrounds for document generation.

ADMIT CARDS/ / CERTIFICATES/: Output directories where the generated personalized documents are saved.

🔧 Prerequisites
To run this project, you will need:

Python 3.x

The following Python libraries:

pandas (for Excel data handling)

Pillow (PIL) (for image processing and text overlay)

openpyxl (for reading/writing Excel files)

