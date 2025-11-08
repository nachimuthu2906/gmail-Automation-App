from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os

app = Flask(__name__)
app.secret_key = "supersecretkey"
app.config["UPLOAD_FOLDER"] = "uploads"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send_emails', methods=['POST'])
def send_emails():
    sender_email = request.form['sender_email']
    password = request.form['password']
    subject = request.form['subject']
    message_body = request.form['message']

    excel_file = request.files['excel_file']
    attachment_files = request.files.getlist('attachments')  # Multiple attachments

    if not excel_file:
        flash("Please upload an Excel file.")
        return redirect(url_for('index'))

    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])

    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
    excel_file.save(excel_path)

    # Save all attachments temporarily
    attachment_paths = []
    for file in attachment_files:
        if file and file.filename != "":
            path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(path)
            attachment_paths.append(path)

    try:
        df = pd.read_excel(excel_path)
        if "Email" not in df.columns:
            flash("Excel file must contain a column named 'Email'.")
            return redirect(url_for('index'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)

        sent_count = 0
        for _, row in df.iterrows():
            receiver = row["Email"]
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = receiver
            msg["Subject"] = subject

            # Attach HTML message body
            msg.attach(MIMEText(message_body, "html"))

            # Attach all uploaded files
            for path in attachment_paths:
                with open(path, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition", f"attachment; filename={os.path.basename(path)}"
                )
                msg.attach(part)

            server.send_message(msg)
            sent_count += 1

        server.quit()
        flash(f"✅ Successfully sent {sent_count} emails with {len(attachment_paths)} attachments each!")
    except Exception as e:
        flash(f"❌ Error: {e}")
    finally:
        # Clean up temporary files
        if os.path.exists(excel_path):
            os.remove(excel_path)
        for path in attachment_paths:
            if os.path.exists(path):
                os.remove(path)

    return redirect(url_for('index'))


if __name__ == '__main__':
    if not os.path.exists('uploads'):
        os.mkdir('uploads')
    app.run(debug=True)
