Set-Content README.md -Value @'
# Selenium Automation Scripts

A collection of Selenium automation scripts.  
Currently includes a scraper for the ARIA Members Directory (https://aria.org.in/members-directory/)  
which exports members to Excel.

## Features
- Handles pagination
- Extracts: type, name, company, mobile_no, email, website
- Output: `aria_members.xlsx`

## Setup
```bash
python -m venv .venv
.\.venv\Scripts\activate   # Windows
pip install -r requirements.txt
python aria_members_final_v2.py
