import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from io import BytesIO

# Define the HTML email template
html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Collaboration with Intelligentsia Club</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f8f9fa;
            color: #2c3e50;
            line-height: 1.6;
        }
        .email-container {
            max-width: 700px;
            margin: auto;
            padding: 20px;
            background: #fff;
            border-radius: 10px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }
        .header {
            background: #3498db;
            padding: 10px;
            text-align: center;
            color: white;
            border-radius: 10px 10px 0 0;
        }
        .content {
            padding: 20px;
        }
        .footer {
            margin-top: 20px;
            text-align: center;
            font-size: 12px;
            color: gray;
        }
    </style>
</head>
<body>
    <div class="email-container">
        <div class="header">
            <h1>Collaboration with Intelligentsia Club</h1>
        </div>
        <div class="content">
            <p>Dear [Company] & Team,</p>
            <p>Warm greetings from Intelligentsia Club at KL University.</p>
            <p>It is a matter of great honor to reach out to your organization, <b>[Company]</b>. We at Intelligentsia Club are a student-led association that focuses on developing AI and machine learning through innovative projects and research.</p>
            <p>We believe that a collaboration between Intelligentsia Club and <b>[Company]</b> would result in remarkable advancements in AI research and innovation.</p>
            <p>We would love to schedule a meeting to discuss this further. Please let us know a convenient time.</p>
            <p>Best regards,<br>
            The Intelligentsia Club Team<br>
            KL University</p>
        </div>
        <div class="footer">
            <p>For more details, visit our website: <a href="https://intelligentsiaclub.netlify.app/">Intelligentsia Club</a></p>
        </div>
    </div>
</body>
</html>
"""


# Function to send a test email for validation
def send_test_email(outlook_user, outlook_password):
    try:
        msg = MIMEMultipart()
        msg["From"] = outlook_user
        msg["To"] = outlook_user
        msg["Subject"] = "Login Test"

        body = "This is a test email to validate your login credentials."
        msg.attach(MIMEText(body, "plain"))

        with smtplib.SMTP("smtp-mail.outlook.com", 587) as server:
            server.starttls()
            server.login(outlook_user, outlook_password)
            server.sendmail(outlook_user, outlook_user, msg.as_string())
        return True
    except Exception as e:
        return False


# Streamlit app
st.title("Collaboration Outreach Automation")

# Check if credentials are already stored in session_state
if "outlook_user" in st.session_state and "outlook_password" in st.session_state:
    st.success("Login validated successfully!")
    outlook_user = st.session_state.outlook_user
    outlook_password = st.session_state.outlook_password

    # Proceed with uploading CSV
    uploaded_file = st.file_uploader("Upload CSV file (with Company, Mail columns)",
                                     type=["csv"])

    if uploaded_file:
        data = pd.read_csv(uploaded_file)

        # Generate email content and display it
        emails = []
        for _, row in data.iterrows():
            email_body = html_template.replace("[Company]", row["Company"])
            emails.append({
                "To": row["Mail"],
                "Subject": f"Collaboration with {row['Company']}",
                "Body": email_body,
            })

        emails_df = pd.DataFrame(emails)
        st.write(emails_df)

        # Button to send generated emails
        if st.button("Send Emails"):
            for _, row in emails_df.iterrows():
                try:
                    msg = MIMEMultipart("alternative")
                    msg["Subject"] = row["Subject"]
                    msg["From"] = outlook_user
                    msg["To"] = row["To"]
                    msg.attach(MIMEText(row["Body"], "html"))

                    with smtplib.SMTP("smtp-mail.outlook.com", 587) as server:
                        server.starttls()
                        server.login(outlook_user, outlook_password)
                        server.sendmail(outlook_user, row["To"], msg.as_string())
                    st.success(f"Email sent successfully to {row['To']}")
                except Exception as e:
                    st.error(f"Failed to send email to {row['To']}. Error: {e}")
else:
    # Email and password input fields
    outlook_user = st.text_input("Outlook Email Address", "")
    outlook_password = st.text_input("Outlook Password", "", type="password")

    # Button to validate login credentials
    if outlook_user and outlook_password:
        if st.button("Validate Login"):
            st.info("Validating your login credentials...")
            if send_test_email(outlook_user, outlook_password):
                st.session_state.outlook_user = outlook_user  # Save email to session_state
                st.session_state.outlook_password = outlook_password  # Save password to session_state
                st.success("Login validated successfully!")
            else:
                st.error("Failed to validate login. Please check your credentials.")
    else:
        st.info("Please enter your Outlook email and password to proceed.")
