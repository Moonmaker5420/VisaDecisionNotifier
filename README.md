## VisaDecisionNotifier

VisaDecisionNotifier is a Python script that automates the process of checking ireland visa application decisions. It scrapes the official visa decision webpage, downloads the decision file, and searches for a specified application number. If the application number is found, the script captures a screenshot of the decision and sends an email notification via Gmail. If not found, it simply outputs "No string found." This tool helps applicants stay updated without manually checking the website.

## Update following fields :

search_string = ""  - Replace with the string to search for,
to_email=""         - recipient email,
smtp_user="",       - Your Gmail email address,
smtp_password=""    - Your Gmail App Password
