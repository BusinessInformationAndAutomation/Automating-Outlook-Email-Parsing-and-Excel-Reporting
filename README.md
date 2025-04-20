# Automating-Outlook-Email-Parsing-and-Excel-Reporting
Summary: This script connects to Microsoft Outlook, filters emails by subject and date, extracts specific data using regex (in this case, total counts from automated check emails within the outlook emails), and writes that data into an Excel workbook intended to be consistnently updated with data. It also creates daily backup copies of the Excel file and updates worksheet titles based on dates.

Use Case Example: Used by analysts or IT support staff who receive automated check emails (e.g., health reports, departure logs) and want to automate storing and summarizing those metrics.This is intended to speed up daily checks within organizations that have an significant reliance on outlook email notifications.

This set of code does this the following:
- Connects to Outlook inbox
- Filters emails by subject and date
- Extracts key values from emails' body using regex
- Automatically populates an Excel workbook with values
- Copies and renames Excel files with date-based naming intending to remove manual portion of renaming
- Auto-generates new worksheets titled by date
