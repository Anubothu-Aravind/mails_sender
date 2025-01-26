import pandas as pd

# Load your CSV file
csv_path = r"C:\Users\91837\Downloads"  # Replace with the actual path to your CSV
data = pd.read_csv(csv_path)

# Template for the email
template = """
Subject: Collaboration with {Company}

Dear {Company} Team,

Warm Greetings from Intelligentsia Club at KL University.

It is a matter of great honor to reach out to the pioneering organization, {Company}, that boldly spearheads open research and innovation in the arena of Artificial Intelligence. We at Kognitiv Club are a student-led association that focuses on developing AI and machine learning through research, development, and meaningful projects.

Our club has been able to execute projects on machine learning, NLP, and LLMs successfully. These projects include the sentiment analysis system for audio data and the multi-lingual translation solution. The initiatives are not only technical but also reflect the club's commitment to the challenges of real-world issues.

We feel that this collaboration with {Company} would be mutual, integrating your experience and leadership in the area of artificial intelligence with our enthusiasm and innovative thinking. Together, we could explore new frontiers in AI research, give practical exposure to students, and make a positive contribution to the broader AI community.

We would really appreciate the opportunity to discuss possible areas for cooperation with your team. May we schedule a call or meeting at your earliest ease to further discuss this? Your guidance and help will be really crucial in defining our way forward.

Thank you for considering our request. We look forward eagerly to hearing from you and are anxious to work together towards achieving our mutual goals.

Best regards,  
The Intelligentsia Club Team  
KL University  
[Your Contact Email]  
"""

# Generate emails for each company
emails = []
for _, row in data.iterrows():
    company_name = row['Company']
    mail = row['Mail']
    email_body = template.format(Company=company_name)
    emails.append({"To": mail, "Subject": f"Collaboration with {company_name}", "Body": email_body})

# Save emails to an Excel file for review
output_path = r"C:\Users\91837\Desktop\Site\Generated_Emails.xlsx"
emails_df = pd.DataFrame(emails)
emails_df.to_excel(output_path, index=False)

print(f"Emails generated and saved to {output_path}")
