import os
import pandas as pd
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.lib import colors

def create_certificate_template(filename, student_data):
    """
    Creates a professional PDF certificate for a student using built-in fonts.
    :param filename: Output PDF file path
    :param student_data: Dictionary with student info (Name, Subject, Year)
    """
    c = canvas.Canvas(filename, pagesize=landscape(A4))
    width, height = landscape(A4)

    # Background - cream/beige color like the reference
    c.setFillColor(colors.HexColor("#f5f2e8"))
    c.rect(0, 0, width, height, fill=1, stroke=0)

    # Ornate double border like the reference
    c.setStrokeColor(colors.HexColor("#d4af37"))  # Gold color
    c.setLineWidth(8)
    c.rect(0.8*cm, 0.8*cm, width-1.6*cm, height-1.6*cm, stroke=1, fill=0)
    
    c.setLineWidth(2)
    c.rect(1.5*cm, 1.5*cm, width-3*cm, height-3*cm, stroke=1, fill=0)

    # Decorative corners (using built-in symbols)
    c.setFont("Times-Roman", 20)
    c.setFillColor(colors.HexColor("#d4af37"))
    # Top corners
    c.drawString(2*cm, height-2.5*cm, "❋")
    c.drawRightString(width-2*cm, height-2.5*cm, "❋")
    # Bottom corners  
    c.drawString(2*cm, 2*cm, "❋")
    c.drawRightString(width-2*cm, 2*cm, "❋")

    # Title "Certificate of Completion" - using Times-Italic for script-like look
    c.setFont("Times-Italic", 36)
    c.setFillColor(colors.HexColor("#2c3e50"))
    c.drawCentredString(width/2, height-4.5*cm, "Certificate of Completion")

    # Decorative line under title
    c.setStrokeColor(colors.HexColor("#d4af37"))
    c.setLineWidth(1)
    c.line(width/2-8*cm, height-5*cm, width/2+8*cm, height-5*cm)

    # "PRESENTED TO" in small caps
    c.setFont("Helvetica", 14)
    c.setFillColor(colors.HexColor("#555555"))
    c.drawCentredString(width/2, height-6.5*cm, "PRESENTED TO")

    # Student Name - large and elegant
    c.setFont("Times-Italic", 32)
    c.setFillColor(colors.HexColor("#000000"))
    c.drawCentredString(width/2, height-8.5*cm, student_data['Name'])
    
    # Underline for name
    c.setStrokeColor(colors.HexColor("#333333"))
    c.setLineWidth(1)
    name_width = c.stringWidth(student_data['Name'], "Times-Italic", 32)
    c.line(width/2 - name_width/2, height-9*cm, width/2 + name_width/2, height-9*cm)

    # Description text
    c.setFont("Times-Roman", 16)
    c.setFillColor(colors.HexColor("#444444"))
    desc = f"for successfully completing the course in {student_data['Subject']} for the year {student_data['Year']}"
    c.drawCentredString(width/2, height-11*cm, desc)

    # Red wax seal (circle)
    c.setFillColor(colors.HexColor("#b71c1c"))
    c.circle(width/2, height-13*cm, 1.5*cm, fill=1, stroke=0)
    
    # Seal text
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.white)
    c.drawCentredString(width/2, height-13*cm, "SEAL")

    # Date and Signature sections
    c.setFont("Times-Roman", 12)
    c.setFillColor(colors.HexColor("#333333"))
    
    # Date section
    c.drawString(4*cm, 3*cm, "DATE")
    c.setLineWidth(1)
    c.setStrokeColor(colors.HexColor("#333333"))
    c.line(4*cm, 2.5*cm, 10*cm, 2.5*cm)
    
    # Signature section  
    c.drawRightString(width-4*cm, 3*cm, "SIGNATURE")
    c.line(width-10*cm, 2.5*cm, width-4*cm, 2.5*cm)

    # Additional decorative elements
    c.setFont("Times-Roman", 16)
    c.setFillColor(colors.HexColor("#d4af37"))
    c.drawCentredString(width/2, height-1.5*cm, "◆ ◆ ◆")

    c.save()

def generate_certificates_from_excel(excel_file, output_dir='CERTIFICATES'):
    """
    Generate certificates for all students in the Excel file.
    :param excel_file: Path to the Excel file with student data
    :param output_dir: Directory to save certificates (created if not exists)
    """
    # Read student data
    df = pd.read_excel(excel_file)
    
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    # Generate certificate for each student
    for _, row in df.iterrows():
        student_data = {
            'Name': str(row['Name']),
            'Subject': str(row['Subject']),
            'Year': str(row['Year'])
        }
        filename = os.path.join(output_dir, f"{student_data['Name'].replace(' ', '_')}_certificate.pdf")
        try:
            create_certificate_template(filename, student_data)
        except Exception as e:
            print(f"Error generating certificate for {student_data['Name']}: {e}")
