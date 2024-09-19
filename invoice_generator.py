import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Read the Excel file
excel_file = 'data/data.xlsx'
df = pd.read_excel(excel_file, header=None)

def extract_value(cell):
    if isinstance(cell, str):
        return cell.split(':')[-1].strip().replace("-", "", 1)
    return None

def get_invoice_no(current_month, current_year):
    month_to_invoice = {
        'April': '01', 'May': '02', 'June': '03', 'July': '04',
        'August': '05', 'September': '06', 'October': '07',
        'November': '08', 'December': '09', 'January': '10',
        'February': '11', 'March': '12'
    }
    invoice_no_prefix = month_to_invoice.get(current_month)
    if invoice_no_prefix:
        next_year_suffix = str(int(current_year[-2:]) + 1)
        return f'{invoice_no_prefix}/{current_year[-2:]}-{next_year_suffix}'
    return None

# Get the current month and year
now = datetime.now()
current_month = now.strftime("%B")
current_year = now.strftime("%Y")

# Generate invoice number
invoice_no = get_invoice_no(current_month, current_year)

# Get the current date in DD.MM.YYYY format
current_date = now.strftime("%d.%m.%Y")

# Extract specific data based on known structure
owner_name = df.iloc[1, 0] if isinstance(df.iloc[1, 0], str) else None
owner_address = df.iloc[2, 0] if isinstance(df.iloc[2, 0], str) else None
gstn = extract_value(df.iloc[3, 0])
pan = extract_value(df.iloc[3, 3])
client_details = df.iloc[5, 1] if isinstance(df.iloc[5, 1], str) else None
client_address = str(df.iloc[6, 1] + df.iloc[7, 1]) if isinstance(df.iloc[6, 1] + df.iloc[7, 1], str) else None
client_mobile = extract_value(df.iloc[8, 1]) if isinstance(df.iloc[8, 1], str) else None
client_gst = extract_value(df.iloc[9, 1]) if isinstance(df.iloc[9, 1], str) else None

# Update the item description
item_description = f'Rent for the month of {current_month} {current_year}'

# Extract item description and financial details
sac = df.iloc[11, 2] if isinstance(df.iloc[11, 2], (str, int, float)) else None
amount = df.iloc[11, 3] if isinstance(df.iloc[11, 3], (str, int, float)) else None
net_amount = df.iloc[12, 3] if isinstance(df.iloc[12, 3], (str, int, float)) else None
cgst = df.iloc[13, 3] if isinstance(df.iloc[13, 3], (str, int, float)) else None
sgst = df.iloc[14, 3] if isinstance(df.iloc[14, 3], (str, int, float)) else None
total = df.iloc[15, 3] if isinstance(df.iloc[15, 3], (str, int, float)) else None
amount_in_words = df.iloc[16, 0].split(':-')[-1].strip() if isinstance(df.iloc[16, 0], str) else None

invoice_data = {
    'invoice_no': invoice_no,
    'date': current_date,
    'owner_name': owner_name,
    'owner_address': owner_address,
    'gstn': gstn,
    'pan': pan,
    'client_details': client_details,
    'client_address': client_address,
    'client_mobile': client_mobile,
    'client_gst': client_gst,
    'item_description': item_description,
    'sac': sac,
    'amount': amount,
    'net_amount': net_amount,
    'cgst': cgst,
    'sgst': sgst,
    'total': total,
    'amount_in_words': amount_in_words,
}

# Update the relevant cells in the DataFrame
df.at[4, 3] = f'{invoice_no}'
df.at[5, 3] = f'{current_date}'
df.at[11, 1] = item_description

def generate_invoice(invoice_data, file_name):
    pdf = SimpleDocTemplate(file_name, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = []

    title_style = ParagraphStyle(
        'Title',
        parent=styles['Title'],
        fontName='Helvetica-Bold',
        fontSize=16,
        leading=20,
        alignment=1  # Center alignment
    )

    normal_style = ParagraphStyle(
        'Normal',
        parent=styles['Normal'],
        fontSize=10,
        leading=12
    )

    # Title
    elements.append(Paragraph('TAX INVOICE', title_style))
    elements.append(Spacer(1, 0.2 * inch))

    # Owner details
    owner_info = [
        invoice_data['owner_name'],
        invoice_data['owner_address'],
        f"GSTN: {invoice_data['gstn']}",
        f"PAN: {invoice_data['pan']}"
    ]
    for info in owner_info:
        elements.append(Paragraph(info, normal_style))

    elements.append(Spacer(1, 0.2 * inch))

    # Client details
    client_info = [
        'To,',
        invoice_data['client_details'],
        invoice_data['client_address'],
        f"Mobile: {invoice_data['client_mobile']}",
        f"GST: {invoice_data['client_gst']}"
    ]
    for info in client_info:
        elements.append(Paragraph(info, normal_style))

    elements.append(Spacer(1, 0.2 * inch))

    # Invoice details
    invoice_table_data = [
        ['Invoice No:', Paragraph(invoice_data['invoice_no'], normal_style), 'Date:', Paragraph(invoice_data['date'], normal_style)]
    ]
    invoice_table = Table(invoice_table_data, colWidths=[80, 200, 50, 100])
    invoice_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12)
    ]))
    elements.append(invoice_table)

    elements.append(Spacer(1, 0.2 * inch))

    # Item details
    item_table_data = [
        ['Sl No.', 'PARTICULARS', 'SAC', 'AMOUNT'],
        ['1', Paragraph(invoice_data['item_description'], normal_style), invoice_data['sac'], invoice_data['amount']],
        ['', '', 'Net Amount', invoice_data['net_amount']],
        ['', '', 'CGST @ 9%', invoice_data['cgst']],
        ['', '', 'SGST @ 9%', invoice_data['sgst']],
        ['', '', 'Grand Total', invoice_data['total']]
    ]
    item_table = Table(item_table_data, colWidths=[50, 300, 100, 100])
    item_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (1, 1), (-1, -1), 'Helvetica')
    ]))
    elements.append(item_table)

    elements.append(Spacer(1, 0.2 * inch))

    # Amount in words
    elements.append(Paragraph(f"Rupees: {invoice_data['amount_in_words']}", normal_style))
    elements.append(Spacer(1, 0.5 * inch))
    elements.append(Paragraph(invoice_data['owner_name'], normal_style))

    pdf.build(elements)

# Generate the PDF invoice
generate_invoice(invoice_data, 'Invoice.pdf')

print("Invoice generated and saved as Invoice.pdf")

def send_email(invoice_file, item_description):
    # Email details
    sender_email = "youremail@gmail.com" 
    receiver_emails = ['recipient1@example.com', 'recipient2@example.com']
    subject = "Invoice: " + item_description
    body = f"Hello,\n\nPlease find the attachment below\n\nThank you and regards,\nRahul Gowda"

    # Create a multipart message and set headers for the email
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = ", ".join(receiver_emails)
    message["Subject"] = subject

    # Attach the body with the msg instance
    message.attach(MIMEText(body, "plain"))

    # Attach the invoice file to the email
    with open(invoice_file, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {invoice_file}",
    )

    message.attach(part)

    # Connect to the SMTP server and send the email
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(sender_email, "yourpassword")
        text = message.as_string()
        server.sendmail(sender_email, receiver_emails, text)
        print("Invoice sent successfully")

# Send the invoice via email
send_email("Invoice.pdf", item_description=item_description)