import getpass
import hashlib
import logging
import pandas as pd
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from fpdf import FPDF  # Library to generate PDF

# Define the employee data

df = pd.read_excel("employee.xlsx")  # Load employee data from an Excel file
# Create a DataFrame
df = pd.DataFrame
try:
    df = pd.read_excel("employee.xlsx")
except FileNotFoundError:
    print("Error: The file 'employee.xlsx' was not found.")
    exit()
except Exception as e:
    print(f"An error occurred while reading the Excel file: {e}")
    exit()
# Define the email sender credentials
USERNAME = "admin"  # Replace with the actual admin username
PASSWORD_HASH = hashlib.sha256("admin_password".encode()).hexdigest()  # Replace with the hashed admin password

def authenticate_user() -> bool:
    username = input("Enter admin username: ")
    password = getpass.getpass("Enter admin password: ")
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    if username == USERNAME and password_hash == PASSWORD_HASH:
        logging.info("Authentication successful.")
        return True
    logging.warning("Authentication failed.")
    return False

# Prompt the user for the sender's email and password
sender_email = input("Enter sender email address: ")
while True:
    sender_password = getpass.getpass("Enter sender email password: ")  # Use an App Password for Gmail
    confirm_password = getpass.getpass("Confirm sender email password: ")
    if sender_password == confirm_password:
        break
    else:
        print("Passwords do not match. Please try again.")

# Loop through each employee and send an email
for index, row in df.iterrows():
    employee_id = row['Employee ID']
    name = row['Name']
    email = row['Email']
    basic_salary = row['Basic Salary']
    allowances = row['Allowances']
    tax_deduction = row['Tax Deduction']
    nssa_deduction = row['NSSA Deduction']

    # Calculate net salary
    net_salary = basic_salary + allowances - tax_deduction - nssa_deduction

    # Generate a PDF for the payslip
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Payslip", ln=True, align='C')
    pdf.ln(10)
    pdf.cell(200, 10, txt=f"Employee ID: {employee_id}", ln=True)
    pdf.cell(200, 10, txt=f"Name: {name}", ln=True)
    pdf.cell(200, 10, txt=f"Email: {email}", ln=True)
    pdf.cell(200, 10, txt=f"Basic Salary: {basic_salary}", ln=True)
    pdf.cell(200, 10, txt=f"Allowances: {allowances}", ln=True)
    pdf.cell(200, 10, txt=f"Tax Deduction: {tax_deduction}", ln=True)
    pdf.cell(200, 10, txt=f"NSSA Deduction: {nssa_deduction}", ln=True)
    pdf.cell(200, 10, txt=f"Net Salary: {net_salary}", ln=True)

    # Save the PDF
    pdf_filename = f"Payslip_{employee_id}.pdf"
    pdf.output(pdf_filename)

    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = email
    msg['Subject'] = f"Payslip for {name}"

    # Add email body
    body = f"""
    Dear {name},

    Please find attached your payslip for this month.

    Best regards,
    SOTA NIGEL INNOVATIONS
    """
    msg.attach(MIMEText(body, 'plain'))

    # Attach the PDF
    with open(pdf_filename, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename={pdf_filename}',
        )
        msg.attach(part)

    # Send the email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, email, msg.as_string())
        server.quit()
        print(f"Email sent to {email} with payslip attached.")
    except Exception as e:
        print(f"Failed to send email to {email}. Error: {e}")

    # Print confirmation
    print(f"Payslip sent to {name} ({email})")
    print("------------------------")
