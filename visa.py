import os
import requests
import smtplib
from datetime import datetime, timedelta
from pyexcel_ods3 import get_data
from fpdf import FPDF
from PyPDF2 import PdfReader
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from bs4 import BeautifulSoup

def scrape_visa_decision_url():
    # URL of the webpage
    url = "https://www.ireland.ie/en/india/newdelhi/services/visas/processing-times-and-decisions/"
    
    # Headers to mimic a browser
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
    }
    
    # Send a GET request to the webpage with headers
    response = requests.get(url, headers=headers)
    
    # Check if the request was successful
    if response.status_code == 200:
        # Parse the HTML content
        soup = BeautifulSoup(response.content, "html.parser")
        
        # Find the specific hyperlink by text and attributes
        hyperlink_section = soup.find("a", href=True, string=lambda t: "Visa decisions made from 1 January 2025" in t if t else False)
        
        # Extract and return the full URL
        if hyperlink_section:
            return f"{url[:url.index('/', 8)]}{hyperlink_section['href']}"  # Ensures correct full URL
        else:
            return "Desired hyperlink not found."
    else:
        return f"Failed to fetch the webpage. Status code: {response.status_code}"

# Function to download a file
def download_file(url, filename):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        with open(filename, 'wb') as file:
            file.write(response.content)
        print(f"Downloaded file: {filename}")
        return filename
    else:
        print(f"Failed to download file: {response.status_code}")
        return None

# Function to convert ODS to PDF
def excel_to_pdf(ods_path, pdf_path):
    # Read data from .ods file
    data = get_data(ods_path)
    
    # Initialize PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    # Convert each sheet's data into the PDF
    for sheet_name, rows in data.items():
        pdf.cell(0, 10, txt=f"Sheet: {sheet_name}", ln=True, align="C")
        for row in rows:
            row_text = " | ".join(str(cell) for cell in row if cell is not None)
            pdf.cell(0, 10, txt=row_text, ln=True)
        pdf.ln(10)  # Add some space between sheets

    # Save the PDF
    pdf.output(pdf_path)
    print(f"Converted {ods_path} to {pdf_path}")

# Function to check if a string is in the PDF and retrieve the full row containing the string
def check_string_in_pdf(pdf_path, search_string):
    reader = PdfReader(pdf_path)
    found_rows = []
    for page in reader.pages:
        text = page.extract_text()
        for line in text.split("\n"):
            if search_string in line:
                found_rows.append(line)
    return found_rows

# Function to send email alert using Gmail
def send_email(subject, body, to_email, smtp_server, smtp_port, smtp_user, smtp_password):
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = to_email
    msg['Subject'] = subject

    # Add body to email
    msg.attach(MIMEText(body, 'plain'))

    # Send the email via Gmail's SMTP server
    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(smtp_user, smtp_password)
            server.sendmail(msg['From'], msg['To'], msg.as_string())
            print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Main logic
def main():
    visa_decision_url = scrape_visa_decision_url()
    print(f"Extracted URL: {visa_decision_url}")
    temp_file = "temp_file.ods"
    pdf_file = "output.pdf"
    search_string = ""  # Replace with the string to search for 

    if visa_decision_url.startswith("http"):  # Ensure a valid URL
        downloaded_file = download_file(visa_decision_url, temp_file)
        if downloaded_file:
            # Convert to PDF
            excel_to_pdf(temp_file, pdf_file)
            
            # Check if the string exists in the PDF and get the full row(s)
            found_rows = check_string_in_pdf(pdf_file, search_string)
            if found_rows:
                print(f"The string '{search_string}' was found in the PDF.")
                
                # Send email alert
                email_body = f"The string '{search_string}' was found in the PDF. Here are the rows:\n\n"
                email_body += "\n".join(found_rows)
                send_email(
                    subject="String Found in PDF",
                    body=email_body,
                    to_email="",  # Replace with recipient email
                    smtp_server="smtp.gmail.com",      # Gmail SMTP server
                    smtp_port=465,                     # SSL port
                    smtp_user="",  # Your Gmail email address
                    smtp_password=""  # Your Gmail App Password
                )
            else:
                print(f"The string '{search_string}' was NOT found in the PDF.")
        else:
            print("Failed to download the file.")
    else:
        print("Failed to extract a valid URL.")
    
    # Cleanup temporary files
    if os.path.exists(temp_file):
        os.remove(temp_file)
    if os.path.exists(pdf_file):
        os.remove(pdf_file)

if __name__ == "__main__":
    main()
