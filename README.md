
# Invoice Generator

This Python project generates PDF invoices from an Excel file and sends them via email using Gmail's SMTP server. It reads data from an Excel file, processes it, and outputs a well-formatted invoice in PDF format, which can then be automatically emailed to specified recipients.

## Features

Reads data from an Excel file
\
Generates a PDF invoice with owner and client details
\
Includes tax details such as GST, CGST, and SGST
\
Automatically emails the generated invoice to specified recipients

## Prerequisites
Before you can run this project, you need to install the required dependencies. This project uses the following Python libraries:

pandas\
reportlab\
smtplib\
email\
MIMEBase\
datetime

Make sure to install them via pip:
```
pip install pandas reportlab
```
## Setup

1. Clone the Repository
First, clone the repository to your local machine:\
```
https://github.com/rahulgowdaa/invoice_generator.git
cd invoice_generator
```
2. Install Dependencies
Install the required Python libraries listed in the requirements.txt file:
```
pip install -r requirements.txt
```

3. Excel File Format
Ensure that the Excel file data.xlsx in the data/ directory follows this format:
```
Row 1-5: Owner and client details.
Row 6-16: Item description, amount details, and tax calculations (CGST, SGST, Total, etc.).
```
The project is hardcoded to look for specific rows and columns in the Excel file.

4. Update Email Credentials
Make sure you update your email credentials for sending the invoice via Gmail SMTP.

```
Replace the sender_email and password with your own:
sender_email = "youremail@gmail.com"
email_password = "yourpassword"
```

If you're using a Gmail account, you must enable less secure apps or use an App Password if you have 2FA enabled.

5. Generate the Invoice and Email
Run the invoice_generator.py file:
```
python invoice_generator.py
```
Once the invoice is generated, it will automatically be emailed to the recipients specified in the send_email function. You can customize the recipient list in the script:\
receiver_emails = ['recipient1@example.com', 'recipient2@example.com']

7. Customization
You can modify the project to suit your specific needs:

Excel Structure: Modify the hardcoded cell locations if your Excel file structure differs.\
SMTP Settings: Update the SMTP settings if you're using an email provider other than Gmail.

## Contributing
If you'd like to contribute, feel free to submit pull requests or issues!
