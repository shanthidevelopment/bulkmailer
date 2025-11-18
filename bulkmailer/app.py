from flask import Flask, render_template, request, render_template_string
import pandas as pd
import smtplib
from email.message import EmailMessage
import os
import time
from datetime import datetime
from tqdm import tqdm

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/send', methods=['POST'])
def send_bulk():
    try:
        start_time = datetime.now()

        sender = request.form['sender']
        password = request.form['password']
        subject = request.form['subject']
        body = request.form['body']

        excel_file = request.files['excel']
        excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
        excel_file.save(excel_path)

        # Multiple attachments
        attachments = []
        for file in request.files.getlist('attachment'):
            if file.filename:
                path = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(path)
                attachments.append(path)

        df = pd.read_excel(excel_path)
        if 'mailList' not in df.columns:
            return "Excel must contain a 'mailList' column."

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)

        results = []
        success_count, fail_count = 0, 0

        for _, row in tqdm(df.iterrows(), total=len(df)):
            recipient = row['mailList']
            if pd.isna(recipient):
                continue

            msg = EmailMessage()
            msg['From'] = sender
            msg['To'] = recipient
            msg['Subject'] = subject
            msg.add_alternative(body, subtype='html')

            for file_path in attachments:
                with open(file_path, 'rb') as f:
                    file_data = f.read()
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=os.path.basename(file_path))

            try:
                server.send_message(msg)
                results.append((recipient, "✅ Sent Successfully"))
                success_count += 1
            except Exception as e:
                results.append((recipient, f"❌ Failed - {e}"))
                fail_count += 1

        server.quit()
        end_time = datetime.now()

        total = success_count + fail_count
        duration = end_time - start_time

        # Summary HTML
        summary_html = f"""
        <html>
        <head>
        <title>Mail Summary | Shanthi IT Solution</title>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, sans-serif;
                background: #f4f6ff;
                padding: 40px;
            }}
            h2 {{
                color: #000851;
                text-align: center;
                margin-bottom: 20px;
            }}
            .summary {{
                background: white;
                border-radius: 15px;
                padding: 20px 30px;
                box-shadow: 0 10px 40px rgba(0,0,0,0.1);
                margin-bottom: 30px;
                line-height: 1.6;
            }}
            .summary p {{
                font-size: 16px;
                margin: 8px 0;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                background: white;
                border-radius: 10px;
                overflow: hidden;
            }}
            th, td {{
                padding: 12px;
                border-bottom: 1px solid #ddd;
                text-align: left;
            }}
            th {{
                background: #000851;
                color: white;
            }}
            .success {{
                color: green;
                font-weight: bold;
            }}
            .fail {{
                color: red;
                font-weight: bold;
            }}
            .footer {{
                text-align: center;
                margin-top: 30px;
                color: #555;
            }}
        </style>
        </head>
        <body>
            <div class="summary">
                <h2>📧 Bulk Mail Sending Summary</h2>
                <p><b>Start Time:</b> {start_time.strftime("%Y-%m-%d %H:%M:%S")}</p>
                <p><b>End Time:</b> {end_time.strftime("%Y-%m-%d %H:%M:%S")}</p>
                <p><b>Duration:</b> {str(duration).split('.')[0]}</p>
                <p><b>Total Emails:</b> {total}</p>
                <p><b>✅ Sent Successfully:</b> {success_count}</p>
                <p><b>❌ Failed:</b> {fail_count}</p>
            </div>

            <table>
                <tr>
                    <th>Recipient Email</th>
                    <th>Status</th>
                </tr>
        """

        for email, status in results:
            css_class = "success" if "✅" in status else "fail"
            summary_html += f"<tr><td>{email}</td><td class='{css_class}'>{status}</td></tr>"

        summary_html += """
            </table>
            <div class="footer">
                <p>© 2025 Shanthi IT Solution — All Rights Reserved</p>
            </div>
        </body>
        </html>
        """

        return render_template_string(summary_html)
    

    except Exception as e:
        return f"<h3>⚠️ Error: {str(e)}</h3>"

if __name__ == "__main__":
    app.run(debug=True)
