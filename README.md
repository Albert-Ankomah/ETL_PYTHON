# 590 Winners Data Processing Script

This script automates the extraction and processing of lottery winner data from Outlook emails, transforms Excel files into CSV format, and manages database interactions.

## Features

- Extracts email attachments from Outlook based on specific criteria
- Processes Excel files into standardized CSV format
- Interacts with SQL Server database to check processing status
- Manages file movement between directories
- Handles special cases for Sundays

## Prerequisites

- Python 3.7+
- Required packages:
  ```bash
  pip install pyodbc pandas pywin32 pyyaml
Configuration
Create a config.yaml file with the following structure:

yaml
database:
  server: your_server_name
  database: your_database_name
  username: your_username
  password: your_password
tables:
  process: Keed.dbo.processStatus
  winners: Keed.dbo.winners
Set up the following directories:

Save folder: C:/Users/albert.boateng/PycharmProjects/590 winners

Destination folder: G:/DATA/590_winners

Usage
Run the script:

bash
python script_name.py
The script will:

Connect to SQL Server to check the last processed date

Download matching email attachments from Outlook

Convert Excel files to CSV format

Move processed files to the destination folder

Functions
connect_to_server()
Establishes connection to SQL Server

Retrieves the last processed date and checks record count

Returns the next processing date and count

download_attachments_from_outlook()
Searches Outlook inbox for emails with specific keywords

Downloads attachments matching the target date

Processes up to 500 most recent messages

create_csv_files()
Converts downloaded Excel files to CSV format

Adds processing date column

Cleans data by removing unnecessary columns

File Management Functions
delete_excel_files_in_folder(): Cleans up Excel files after conversion

move_files_loading_folder(): Moves processed files to destination

move_files_loading_folder_sunday(): Special handling for Sunday files

Email Processing Rules
The script looks for emails containing these keywords:

"Noon"

"Rush"

"National Weekly"

"WINNERS"

"LIST"

Database Schema Requirements
The script expects these tables:

processStatus table with columns:

date (datetime)

processname (varchar)

winners table with columns:

ppn_dt (datetime)

Error Handling
The script includes basic error handling for:

Database connection issues

Email processing errors

File operations

Maintenance
Regularly check the configuration file for updates

Monitor the save and destination folders for proper file movement

Review Outlook email filters periodically

Notes
The script handles Sunday files differently (expects only 1 file)

Ensure Outlook is running when executing the script

Database credentials are stored in config.yaml (keep this file secure)

text

This README provides:
1. Clear installation instructions
2. Configuration requirements
3. Detailed function documentation
4. Operational notes
5. Maintenance guidelines

You may want to customize the paths and database details to match your specific environment. The markdown format makes it easy to display on GitHub or other platforms.
