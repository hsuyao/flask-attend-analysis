# Attendance Analysis Web App

This is a Flask web application for analyzing attendance data from Excel files. It processes `.xls` and `.xlsx` files, converts `.xls` to `.xlsx` using `soffice`, generates attendance summaries, and displays the latest week's data with highlights for changes in attendance status.

## Features
- Upload an Excel file (`.xls` or `.xlsx`) with attendance data.
- Convert `.xls` files to `.xlsx` using `soffice` (LibreOffice command-line tool).
- Analyze attendance by week, district, and age group.
- View six-month trimmed mean statistics with age breakdown under the
  "主日" event.
- Highlight names in light green if they attended this week but not last week.
- Highlight names in light red if they attended last week but not this week.
- Sort highlighted names to the top of the list.
- Download the processed Excel file with new summary sheets.
- Change password after logging in.
- Reset password using username and registered email if you cannot log in.
- Event log tracks logins and admin actions (keeps last 100k entries).
- View event logs via the admin interface.

## Prerequisites
- Python 3.8 or higher
- LibreOffice installed (for `soffice` command to convert `.xls` to `.xlsx`)
- A Render account for deployment
- A GitHub account to host the repository

## Local Development
1. **Install LibreOffice**:
   - On Ubuntu/Debian:
     ```bash
     sudo apt update
     sudo apt install libreoffice
