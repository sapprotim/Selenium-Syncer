AutoFetch AH is a Python-based automation tool that uses Selenium to collect user data from the AH CMS platform. It reads a list of users from an Excel file, navigates to each profile, extracts 16 specific data points, and saves the results efficiently. Designed to handle up to 1000 users per day, this tool boosts speed, accuracy, and eliminates manual work.

⚙️ Features
Automated login to AH CMS

Excel-driven user list processing

Extracts 16 key data points per user

Saves collected data to Excel

Logs out safely after each session

Handles 1000+ users daily with high reliability

📝 Requirements
Python 3.8+

Google Chrome (or compatible browser)

ChromeDriver matching your browser version

Install dependencies:

bash
Copy
Edit
pip install -r requirements.txt
🔐 Private Data Setup (Required before running)
⚠️ Note: This repository does not include sensitive/private files.
You must manually add the following files before running the script:

Contains your CMS login credentials:


user_list.xlsx – Excel file with the list of users to be processed.

Place both files in the project root directory.

🚀 Run the Script
bash
Copy
Edit
python autofetch_ah.py
📁 Output
The output Excel file with the extracted data will be saved in the project directory as output_data.xlsx.

📌 Notes
This tool is for internal use only.

Do not share sensitive files or credentials publicly.

Ensure CMS site structure doesn't change (may break selectors).
