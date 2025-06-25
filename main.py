import os
from datetime import datetime, timedelta
# from fractions import Fraction
import pandas as pd
import win32com.client
import pyodbc
import shutil
import re
import yaml
from pathlib import Path

config_path = Path('config.yaml')

with open(config_path) as f:
    config = yaml.safe_load(f)

server_ = config['database']['server']
data_base = config['database']['database']
user = config['database']['username']
pass_word = config['database']['password']
process_status = config['tables']['process']
winners_table = config['tables']['winners']


def connect_to_server():
    server = server_
    database = data_base
    username = user
    password = pass_word

    # Create a connection string
    conn_str = f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={server};DATABASE={database};UID={username};PWD={password}'

    # Establish a connection
    print('Establishing Connection...')
    conn = pyodbc.connect(conn_str)
    if conn:
        print('Connection established successfully')
        cursor = conn.cursor()

        # Define your SQL query
        sql_query = f"""
        SELECT max(date)
        FROM {process_status}
        WHERE processname = ?
        """

        condition_value = '590 Collections and disbursement'
        cursor.execute(sql_query, (condition_value,))
        row = cursor.fetchall()
        print(row)
        row_string = row[0][0]
        new_row = datetime.strptime(row_string, "%Y-%m-%d")
        next_Date = new_row + timedelta(days=1)
        final_date = next_Date.strftime("%Y-%m-%d")
        # print(final_date)

        sql_query2 = f"""
        SELECT count(*)
        FROM {winners_table}
        WHERE ppn_dt = ?
        """

        date_check = next_Date
        cursor.execute(sql_query2, (date_check,))
        count = cursor.fetchall()
        check_count = count[0][0]
        print(count[0][0])

        # print(f'final_Date = {final_date}, count = {count}')
        return final_date, check_count

        #return final_date

    conn.close()


def create_csv_files(save_folder, rundate):
    path = save_folder
    files = os.listdir(path)
    for file in files:
        if file.endswith('xlsx'):
            filename = os.path.join(save_folder, file)
            print(filename)
            current_file = pd.read_excel(filename)
            print(current_file)
            current_file["DATE"] = rundate
            if 'DRAW ID' in current_file:
                current_file = current_file.drop(columns=['DRAW ID'])

            # current_file["DATE"] = current_file["DATE"].dt.strftime("%Y-%m-%d")
            # print(current_file["DATE"])
            # csv_filename = os.path.splitext(save_folder)[0] + '.csv'
            current_file.to_csv(filename + '.csv',
                             index=False,
                              header=True,
                              date_format='%Y-%m-%d')
            df = pd.DataFrame(current_file)
            print(df)


def save_attachments(email, save_folder):
    for attachment in email.Attachments:
        if attachment.FileName.endswith('.xlsx') or attachment.FileName.endswith('.xls'):
            attachment.SaveAsFile(os.path.join(save_folder, attachment.FileName))


def extract_date_from_filename(filename):

    # Define a regex pattern to capture 'Month Day, Year' format
    date_pattern = r"(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\s+\d{1,2},\s+\d{4}"
    match = re.search(date_pattern, filename, re.IGNORECASE)

    if match:
        date_str = match.group(0)
        try:
            # Convert the extracted date to a standard format
            date_obj = datetime.strptime(date_str, "%B %d, %Y")
            return date_obj.strftime("%Y-%m-%d")  # Return in "YYYY-MM-DD" format
        except ValueError:
            print(f"Error: Unable to parse date from filename '{filename}'")
            return None
    return None


def download_attachments_from_outlook(save_folder, next_date, count, max_messages=500):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    print(f"messages_count = {messages.count}")

    messages.Sort('[ReceivedTime]', True)

    processed_count = 0

    # for i in range(5):
    #     message = messages.GetNext()
    #     print("" + message.Subject, str(message.ReceivedTime))

    for message in messages:
        if processed_count >= max_messages:
            break

        try:
            if message.Attachments.Count > 0:
                if any(keyword in message.Subject for keyword in subject_keywords):
                    for attachment in message.Attachments:
                        # Extract the date from the attachment's filename
                        attachment_date = extract_date_from_filename(attachment.FileName)

                        # Save the attachment if the date matches
                        if attachment_date == next_date:
                            attachment.SaveAsFile(f"{save_folder}\\{attachment.FileName}")
                            print(f"Saved attachment '{attachment.FileName}' from email '{message.Subject}'")
                            break

            processed_count += 1

        except Exception as e:
            print(f"Error processing message '{message.Subject}': {e}")


    # for message in messages:
    #     for keyword in subject_keywords:
    #         if keyword in message.subject and message.Attachments.Count > 0:
    #             subject = message.Subject
    #             # print(subject)
    #             index = subject.find('LIST')
    #             file_date = subject[index + len("LIST"):].strip()
    #             # print(f"file_date = {file_date}")
    #             # file_date = (subject[-12:]).strip()
    #             date_conversion = datetime.strptime(file_date, "%B %d, %Y")
    #             date = date_conversion.strftime("%Y-%m-%d")
    #             # target_date = datetime.strptime(next_date, "%Y-%m-%d")
    #             # print(date)
    #             if date == next_date and count == 0:
    #                 save_attachments(message, save_folder)


def delete_excel_files_in_folder(save_folder):
    files = os.listdir(save_folder)

    # Filter out only Excel files and remove them
    for file in files:
        if file.endswith(".xlsx") or file.endswith(".xls"):
            os.remove(os.path.join(save_folder, file))


#move files from save folder to loading folder for data transfer
def move_files_loading_folder(save_folder, destination_folder):
    files = os.listdir(save_folder)
    if len(files) == 2:
        for file in files:
            source_path = os.path.join(save_folder, file)
            destination_path = os.path.join(destination_folder, file)
            shutil.move(source_path, destination_path)
            print("File moved successfully.")
    else:
        print("Number of files in the source folder is not equal to 2.")


def move_files_loading_folder_sunday(save_folder, destination_folder):
    files = os.listdir(save_folder)
    # print(f'length of files = {len(files)}')
    if len(files) == 1:
        for file in files:
            source_path = os.path.join(save_folder, file)
            destination_path = os.path.join(destination_folder, file)
            shutil.move(source_path, destination_path)
            print("File moved successfully.")
    else:
        print("No files available.")


# main
subject_keywords = ["Noon", "Rush", "National Weekly", "WINNERS", "LIST"]
save_folder = "C:/Users/albert.boateng/PycharmProjects/590 winners"
destination_folder = "G:/DATA/590_winners"


runDate = connect_to_server()
print(f"rundate = {runDate}")
final_date = runDate[0]
count_ = runDate[1]
print(f"final_date = {final_date}, count = {count_}")

date_object = datetime.strptime(final_date, '%Y-%m-%d')
day_name = date_object.strftime('%A')
print(f"day_name = {day_name}")
download_attachments_from_outlook(save_folder, final_date, count_)
create_csv_files(save_folder, final_date)
delete_excel_files_in_folder(save_folder)


if day_name != 'Sunday':
    move_files_loading_folder(save_folder, destination_folder)
else:
    move_files_loading_folder_sunday(save_folder, destination_folder)



