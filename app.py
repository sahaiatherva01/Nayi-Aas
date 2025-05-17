from flask import Flask, request, redirect
from datetime import datetime
import openpyxl
import os

app = Flask(__name__)
EXCEL_FILE = os.path.join(app.root_path, 'data.xlsx')

# Create Excel file with headers if it doesn't exist
if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Timestamp", "Full Name", "E-Mail", "Phone Number", "Area of Interest", "Message"])
    wb.save(EXCEL_FILE)

@app.route('/')
def home():
    try:
        with open('join.html', 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        return f"Could not load form: {e}"

@app.route('/submit-form', methods=['POST'])
def submit_form():
    try:
        name = request.form.get('Full Name')
        email = request.form.get('E-Mail')
        phone = request.form.get('Phone Number')
        interest = request.form.get('Area of Interest')
        message = request.form.get('Message')

        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            name,
            email,
            phone,
            interest,
            message
        ])
        wb.save(EXCEL_FILE)

        return redirect('/thank-you')
    except Exception as e:
        return f"An error occurred: {e}"

@app.route('/thank-you')
def thank_you():
    try:
        with open('thankyou.html', 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        return f"Could not load thank-you page: {e}"

if __name__ == '__main__':
    app.run(debug=True)
