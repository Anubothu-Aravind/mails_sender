import pandas as pd
import win32com.client as win32

# Path to the CSV file containing company and email data
csv_path = r"C:\Users\91837\Downloads\KOGNITIV SPREADSHEETS FOR COLABS - Companies.csv"

# Load the CSV file
data = pd.read_csv(csv_path)

# Read the HTML template
html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Kognitiv Club</title>
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f8f9fa;
            color: #2c3e50;
            line-height: 1.8;
        }
        .email-container {
            max-width: 700px;
            margin: 40px auto;
            background: #ffffff;
            border-radius: 16px;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #2c3e50, #3498db);
            color: white;
            padding: -30px 10px;
            text-align: center;
            border-bottom: 4px solid rgba(255, 255, 255, 0.1);
        }
        .header h1 {
            margin: 0;
            font-size: 32px;
            font-weight: 700;
            letter-spacing: 0.5px;
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }
        .email-content {
            padding: 50px 40px;
            font-size: 16px;
            background: linear-gradient(to bottom, #ffffff, #f8f9fa);
        }
        .email-content p {
            margin-bottom: 24px;
            color: #34495e;
            line-height: 1.9;
        }
        .company-name {
            color: #2980b9;
            font-weight: 600;
            padding: 2px 6px;
            background: rgba(41, 128, 185, 0.1);
            border-radius: 4px;
            display: inline-block;
        }
        .signature {
            margin-top: 50px;
            padding-top: 30px;
            border-top: 3px solid #ecf0f1;
            font-size: 15px;
        }
        .signature-details {
            color: #7f8c8d;
            font-size: 14px;
            line-height: 1.8;
            margin-top: 15px;
        }
        .website-link {
            display: inline-block;
            margin-top: 25px;
            padding: 14px 28px;
            background: linear-gradient(135deg, #3498db, #2980b9);
            color: black;
            text-decoration: none;
            border-radius: 8px;
            transition: all 0.3s ease;
            font-weight: 600;
            box-shadow: 0 4px 12px rgba(52, 152, 219, 0.2);
        }
        .website-link:hover {
            background: linear-gradient(135deg, #2980b9, #2573a7);
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(52, 152, 219, 0.3);
        }
        .postscript {
            margin-top: 35px;
            font-size: 14px;
            color: #2c3e50;
            font-style: italic;
            padding: 20px;
            background: rgba(236, 240, 241, 0.7);
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        }
        .highlight {
            background: linear-gradient(120deg, rgba(52, 152, 219, 0.15) 0%, rgba(52, 152, 219, 0.15) 100%);
            padding: 3px 8px;
            border-radius: 4px;
            font-weight: 500;
        }
        .greeting {
            font-size: 18px;
            color: #2c3e50;
            margin-bottom: 30px;
            font-weight: 600;
        }
        .contact-info {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 8px;
            margin-top: 30px;
        }
        .contact-info p {
            margin: 8px 0;
        }
        .website-link {
            display: inline-block;
            margin-top: 15px;
            padding: 12px 24px;
            background: linear-gradient(135deg, #3498db, #2980b9);
            color: black;
            text-decoration: none;
            border-radius: 8px;
            transition: all 0.3s ease;
            font-weight: 600;
            box-shadow: 0 4px 12px rgba(52, 152, 219, 0.2);
        }
        .website-link:hover {
            background: linear-gradient(135deg, #2980b9, #2573a7);
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(52, 152, 219, 0.3);
        }
    </style>
</head>
<body>
    <div class="email-container">
        <div class="header">
            <h1>Kognitiv Club</h1>
        </div>
        <div class="email-content">
            <div class="greeting">Dear [Company] Team,</div>

            <p>Warm Greetings from Kognitiv Club at KL University.</p>

            <p>It is a matter of great honor to reach out to the pioneering organization, <span class="company-name">[Company]</span>, that boldly spearheads open research and innovation in the arena of Artificial Intelligence. We at Kognitiv Club are a student-led association that focuses on developing AI and machine learning through research, development, and meaningful projects.</p>

            <p>Our club has been able to execute projects on machine learning, NLP, and LLMs successfully. These projects include the <span class="highlight">sentiment analysis system for audio data</span> and the <span class="highlight">multi-lingual translation solution</span>. The initiatives are not only technical but also reflect the club's commitment to the challenges of real-world issues.</p>

            <p>We feel that this collaboration with <span class="company-name">[Company]</span> would be mutual, integrating your experience and leadership in the area of artificial intelligence with our enthusiasm and innovative thinking. Together, we could explore new frontiers in AI research, give practical exposure to students, and make a positive contribution to the broader AI community.</p>

            <p>We would really appreciate the opportunity to discuss possible areas for cooperation with your team. May we schedule a call or meeting at your earliest convenience to further discuss this? Your guidance and help will be really crucial in defining our way forward.</p>

            <div class="signature">
                <div>Best regards,</div>
                <div class="signature-details">
                    The Kognitiv Club Team<br>
                    KL University<br>
                    <a href="mailto:[Your Contact Email]" class="email-link">[Your Contact Email]</a>
                </div>
            </div>

            <div class="postscript">
                To learn more about us, please visit our website:
                <a href="https://intelligentsiaclub.netlify.app/" class="website-link">Visit Our Website</a>
            </div>
        </div>
    </div>
</body>
</html>
"""

# Initialize Outlook
outlook = win32.Dispatch('outlook.application')

# Loop through each row in the CSV
for _, row in data.iterrows():
    # Replace placeholders in the template
    html_content = html_template.replace("[Company]", row['Company'])
    html_content = html_content.replace("[Your Contact Email]", "your_email@example.com")  # Replace with your email

    # Create a new mail item
    mail = outlook.CreateItem(0)
    mail.Subject = f"Collaboration with {row['Company']}"
    mail.To = "2200032689@kluniversity.in"  # Fixed recipient email
    mail.HTMLBody = html_content

    # Uncomment the line below to display the email before sending
    # mail.Display()

    # Send the email
    mail.Send()

print("Emails sent successfully!")
