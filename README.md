TMX Sanitizer Pro
TMX Sanitizer Pro is a Streamlit-based, human-in-the-loop tool for cleaning, analyzing, QA-checking, and exporting TMX files.
Core Features
Upload and analyze one or multiple TMX files
Apply selected safe auto-repairs
Normalize Unicode and language codes
Remove hidden characters and unwanted spacing
Flag broken or unbalanced tags without automatically repairing them
Flag target-equals-source segments
Flag German micro-QA issues
Upload a Do Not Translate / Brand Protection XLSX file
Resolve duplicates manually or apply a controlled merge policy
Edit segments directly in the interface
Export sanitized TMX files
Export XLSX QA reports
Do Not Translate / Brand Protection XLSX
Use this format:
Column A	Column B	Column C optional
Source protected term	Required target representation	Note
Example:
| Oracle | Oracle | Brand name |
| Azure | Azure | Product name |
Installation
```bash
pip install -r requirements.txt
streamlit run app.py
```
Deployment
For Streamlit Cloud, upload:
```text
app.py
requirements.txt
README.md
```
Then connect the GitHub repository to Streamlit Cloud.
Important Notice
Use at your own risk. This is an ongoing project and should be tested with sample TMX files before production use.
