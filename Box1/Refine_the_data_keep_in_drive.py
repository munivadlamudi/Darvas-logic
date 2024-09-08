import os
import pandas as pd
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import pickle
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from datetime import datetime

print(os.getcwd())


# Google Drive authentication
def authenticate_drive(token_file, client_secrets_file):
    gauth = GoogleAuth()
    gauth.DEFAULT_SETTINGS['client_config_file'] = client_secrets_file
    if os.path.exists(token_file):
        with open(token_file, "rb") as token:
            gauth.LoadCredentialsFile(token_file)
    if gauth.credentials is None or gauth.access_token_expired:
        gauth.LocalWebserverAuth()
        gauth.SaveCredentialsFile(token_file)
    return GoogleDrive(gauth)

# Function to upload the refined output file to Google Drive
def upload_to_drive(drive, folder_name, file_path):
    folder_list = drive.ListFile({'q': f"title = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"}).GetList()
    if not folder_list:
        print(f"The '{folder_name}' folder was not found.")
        return
    folder_id = folder_list[0]['id']
    file_drive = drive.CreateFile({'title': os.path.basename(file_path), 'parents': [{'id': folder_id}]})
    file_drive.SetContentFile(file_path)
    file_drive.Upload()
    print(f'Uploaded file: {file_drive["title"]}')

#  Function to send an email with a message and DataFrame content
def send_email_with_report(df_filtered):
    SCOPES = ["https://www.googleapis.com/auth/gmail.send", "https://www.googleapis.com/auth/drive.file"]

    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)

        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    service = build("gmail", "v1", credentials=creds)

    sender = 'your-email@gmail.com'
    to = ['recipient-email@gmail.com']
    subject = f"Refined Report - {datetime.now().strftime('%Y-%m-%d')}"
    message_text = "Please find the attached refined report."

    # Filter the DataFrame to include necessary columns
    df_filtered = df_filtered[['Stock Symbol', 'Entry Date', 'Exit Date']]  # Adjust columns as needed

    # Convert DataFrame to HTML for email content
    df_html = df_filtered.to_html(index=False)

    message = MIMEMultipart()
    message["to"] = ', '.join(to)
    message["subject"] = subject
    msg = MIMEText(f"{message_text}\n\n{df_html}", "html")
    message.attach(msg)
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")

    try:
        sent_message = service.users().messages().send(userId="me", body={"raw": raw_message}).execute()
        print(f"Sent message with Message Id: {sent_message['id']}")
    except Exception as error:
        print(f"An error occurred: {error}")

# Main function to refine data and handle Google Drive upload + email
def main():
    # Step 1: Load the data from an Excel file
    file_path = 'Box1/Reports/Darvas_Box_Trade_Combined_Report.xlsx'
    df = pd.read_excel(file_path, sheet_name='Sheet1')

    # Step 2: Sort the data by 'Stock Symbol', 'Exit Date', and 'Entry Date' (oldest first)
    df_sorted = df.sort_values(by=['Stock Symbol', 'Exit Date', 'Entry Date'])

    # Step 3: Drop duplicates by keeping the oldest 'Entry Date' for each 'Stock Symbol' and 'Exit Date' combination
    df_filtered = df_sorted.drop_duplicates(subset=['Stock Symbol', 'Exit Date'], keep='first')

    # Step 4: Sort the filtered data by 'Stock Symbol' and 'Entry Date' in descending order
    df_filtered_sorted = df_filtered.sort_values(by=['Stock Symbol', 'Entry Date'], ascending=[True, False])

    # Step 5: Save the refined data to a new Excel file
    output_path = 'Box1/Reports/refined_output1.xlsx'
    df_filtered_sorted.to_excel(output_path, index=False)
    print("Data refinement and sorting are complete. Refined data saved to:", output_path)

    # Step 6: Authenticate with Google Drive and upload the file
    drive = authenticate_drive("drive_token.pickle", "client_secrets.json")
    upload_to_drive(drive, 'Market Tracker', output_path)  # Ensure 'Market Tracker' folder exists in your Drive

    # Step 7: Optionally, send an email with the report
    # send_email_with_report(df_filtered_sorted)

if __name__ == "__main__":
    main()
