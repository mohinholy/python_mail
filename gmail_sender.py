import os
import base64
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

# Paths
EXCEL_FILE = r"C:\study_mat\resume_others\python_mail\unique_sent_emails.xlsx"
RESUME_FILE = r"C:\study_mat\resume_others\Mohinuddin_VLSI_Resume.pdf"

# Gmail API scopes
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

# =============================
# MODE SELECTION
# =============================
# Set mode = "apply"    --> first mail with resume
# Set mode = "followup" --> follow-up mail without resume
mode = "apply"   # change when needed

# =============================
# Email Templates (HTML)
# =============================

def get_application_mail():
    subject = "Application for VLSI/RTL/FPGA Position"
    body = """\
    <html>
      <body>
        <p>Respected Sir/Ma'am,</p>

        <p>I am <b>Mohinuddin Holy</b>. I am interested in a job related to RTL/FPGA or any VLSI-related position. 
        I have recently completed my <b>M.Tech in Information and Communication Technology</b> with a specialization 
        in <b>VLSI and Embedded Systems</b> from DAIICT Gandhinagar.</p>

        <p><u>My academic and project experience includes:</u></p>
        <ul>
          <li><b>SoC of Neural Networks on FPGA (Thesis)</b> – Implemented audio digit classification using Python and Verilog RTL as SoC in Vivado.</li>
          <li><b>8×8 SRAM Design using Cadence Virtuoso</b> – Designed SRAM cell array with decoder, pre-charge, and sense amplifier.</li>
          <li><b>FPGA-based IIR Filter Design</b> – Modelled in MATLAB and implemented using Verilog and FPGA.</li>
          <li><b>RFID-Based Wireless Communication System</b> – Built using FPGA and Arduino.</li>
        </ul>

        <p>I am enthusiastic about applying my skills in RTL design, FPGA development, and memory design.<br>
        I have attached my resume for your review, and I would be glad to discuss how I can contribute to your team.</p>

        <p>Looking forward to hearing from you.</p>

        <p>Best regards,<br>
        <b>Mohinuddin Holy</b><br>
        <b>+918735813414</b></p>
      </body>
    </html>
    """
    return subject, body, True  # attach resume

def get_followup_mail():
    subject = "Follow-up on My Application for VLSI/RTL/FPGA Position"
    body = """\
    <html>
      <body>
        <p>Respected Sir/Ma'am,</p>

        <p>I hope this message finds you well. I had applied earlier for a <b>VLSI/RTL/FPGA related position</b> 
        and wanted to kindly follow up regarding the status of my application.</p>

        <p>I remain very interested in contributing to your team and would be grateful for any update you could provide.</p>

        <p>Thank you for your time and consideration.</p>

        <p>Best regards,<br>
        <b>Mohinuddin Holy</b></p>
      </body>
    </html>
    """
    return subject, body, False  # no resume

# =============================
# Gmail API Authentication
# =============================
def gmail_service():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)

# =============================
# Create Email
# =============================
def create_message(to, subject, body, attach_resume=False):
    message = MIMEMultipart()
    message["to"] = to
    message["subject"] = subject

    # HTML body
    message.attach(MIMEText(body, "html"))

    # Attach resume only for application mail
    if attach_resume and os.path.exists(RESUME_FILE):
        with open(RESUME_FILE, "rb") as f:
            resume = MIMEApplication(f.read(), _subtype="pdf")
            resume.add_header("Content-Disposition", "attachment", filename=os.path.basename(RESUME_FILE))
            message.attach(resume)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {"raw": raw_message}

# =============================
# Send Emails
# =============================
def send_emails():
    service = gmail_service()

    # Load email list (column B → index 1, not 3)
    df = pd.read_excel(EXCEL_FILE)
    email_list = df.iloc[:, 1].dropna().tolist()

    # Select template
    if mode == "apply":
        subject, body, attach_resume = get_application_mail()
    else:
        subject, body, attach_resume = get_followup_mail()

    # Send mail to each address
    for email in email_list:
        try:
            msg = create_message(email, subject, body, attach_resume)
            service.users().messages().send(userId="me", body=msg).execute()
            print(f"✅ Sent to {email}")
        except Exception as e:
            print(f"❌ Failed for {email}: {e}")

if __name__ == "__main__":
    send_emails()
