from flask import Flask, render_template, request
import smtplib
import pandas as pd
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Folder to store uploaded files temporarily
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send', methods=['GET', 'POST'])
def send_email():
    if request.method == 'POST':
        sender_email = request.form['sender_email']
        password = request.form['password']   # app password
        subject = request.form['subject']
        message_body = request.form['message']
        excel_file = request.files['excel_file']
        attachments = request.files.getlist('attachments')

        # Save Excel file
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(excel_file.filename))
        excel_file.save(excel_path)

        # Read Excel
        df = pd.read_excel(excel_path)
        if 'email' not in df.columns:
            return "❌ Excel file must contain a column named 'email'"

        emails = df['email'].dropna().tolist()

        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(sender_email, password)
        except Exception as e:
            return f"❌ Login failed: {e}"

        sent_count = 0
        for mail in emails:
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = mail
            msg['Subject'] = subject

            msg.attach(MIMEText(message_body, 'plain'))

            # Attach all uploaded files
            for file in attachments:
                if file.filename == '':
                    continue
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
                file.save(file_path)
                with open(file_path, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{file.filename}"')
                msg.attach(part)

            try:
                server.send_message(msg)
                sent_count += 1
            except Exception as e:
                print(f"❌ Failed to send mail to {mail}: {e}")

        server.quit()
        return f"✅ Emails sent successfully to {sent_count} recipients!"

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
