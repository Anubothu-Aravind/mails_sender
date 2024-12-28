
# Email Automation for Kognitiv Club Collaborations

This project streamlines the process of reaching out to organizations for potential collaborations by generating personalized emails from a list of companies and their details. The project consists of generating email drafts in Excel and sending them using Python scripts.

## Project Structure

```plaintext
├── Companies.csv                # Input CSV containing company data
├── Generated_Emails.xlsx        # Output Excel file with generated email drafts
├── LICENSE                      # License file for the project
├── content_genration_to_excel.py  # Script to generate email drafts and save them to Excel
└── mail.py                      # Script to send emails using Outlook
```

### Files Overview

#### `/Companies.csv`
Contains the company data in the following format:
```csv
S.No,Company,Mail,WebSites,LinkedIn,Draft
1,EleutherAI,info@eleuther.ai,https://www.eleuther.ai/,https://www.linkedin.com/company/eleutherai/,
```

#### `content_genration_to_excel.py`
- Loads company data from `Companies.csv`.
- Generates personalized email content using a predefined template.
- Saves the generated emails into an Excel file (`Generated_Emails.xlsx`) for review.

#### `mail.py`
- Loads the company data from `Companies.csv`.
- Sends personalized emails using Microsoft Outlook with an HTML-based email template.

### Features

1. **Automated Email Generation**: Uses pandas to read company details and format email drafts.
2. **Customizable HTML Templates**: Allows you to craft visually appealing emails with company-specific details.
3. **Seamless Integration with Outlook**: Automates sending emails via Outlook's COM object.
4. **Scalable Workflow**: Handles multiple companies from a single CSV file.

### Prerequisites

1. **Python Environment**: Ensure Python 3.x is installed along with the following libraries:
   - `pandas`
   - `openpyxl`
   - `pywin32` (for Outlook integration)

2. **Microsoft Outlook**: Must be installed and configured on your system to use the `mail.py` script.

### Setup Instructions

1. Clone the repository:
   ```bash
   git clone https://github.com/Anubothu-Aravind/mails_sender.git
   cd mails_sender
   ```

2. Install required dependencies:
   ```bash
   pip install pandas openpyxl pywin32
   ```

3. Update the file paths in the scripts (`content_genration_to_excel.py` and `mail.py`) to match your environment.

4. Replace placeholder values (like `[Your Contact Email]`) in the email template with your details.

### Usage

#### Generate Email Drafts
Run the following command to generate emails and save them to an Excel file:
```bash
python content_genration_to_excel.py
```

#### Send Emails
Use the `mail.py` script to send the generated emails:
```bash
python mail.py
```

### License

This project is licensed under the terms of the [MIT License](LICENSE).

### Contributing

Contributions are welcome! Please open an issue or submit a pull request with any improvements.

### Acknowledgements

Special thanks to the Kognitiv Club at KL University for their mission to foster collaboration and innovation in Artificial Intelligence.



### Contact  
For questions or suggestions, please email: [aanubothu@gmail.com]
