from datetime import datetime
import os.path
import base64

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.send",
]

SAMPLE_SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SAMPLE_RANGE_NAME = "A1:H25"
SPREADSHEET_ID_FILE = "spreadsheet_id.txt"


def main():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)
        drive_service = build("drive", "v3", credentials=creds)
        gmail_service = build("gmail", "v1", credentials=creds)


        spreadsheet_id = get_or_create_spreadsheet_id(drive_service)

        sheet = service.spreadsheets()
        result = (
            sheet.values()
            .get(spreadsheetId=spreadsheet_id, range=SAMPLE_RANGE_NAME)
            .execute()
        )
        values = result.get("values", [])

        if not values:
            print("No data found.")
            return

        update_invoice_date(sheet, spreadsheet_id)
        update_invoice_number(sheet, spreadsheet_id, values)
        update_invoice_total(sheet, spreadsheet_id)

        download_pdf(drive_service, spreadsheet_id)
        send_email_with_attachment(gmail_service, f"spreadsheet_{spreadsheet_id}.pdf")
    except HttpError as err:
        print(err)

def get_or_create_spreadsheet_id(drive_service):
    """
    Retrieves or creates a new Google Spreadsheet ID.

    This function checks if a spreadsheet ID file exists. If it does, it reads the ID from the file.
    If the file does not exist, it uses a sample spreadsheet ID from an environment variable.
    Regardless of the source, it creates a copy of the spreadsheet and saves the new ID to the file.

    Args:
        drive_service: The Google Drive service instance used to interact with the Google Drive API.

    Returns:
        str: The new spreadsheet ID after creating a copy of the original spreadsheet.
    """
    if os.path.exists(SPREADSHEET_ID_FILE):
        with open(SPREADSHEET_ID_FILE, 'r') as file:
            spreadsheet_id = file.read().strip()
            print(f"Using existing spreadsheet ID from file: {spreadsheet_id}")
    else:
        spreadsheet_id = SAMPLE_SPREADSHEET_ID
        print(f"Using spreadsheet ID from environment variable: {spreadsheet_id}")

    # Always create a copy of the spreadsheet
    new_spreadsheet_id = create_copy_of_spreadsheet(drive_service, spreadsheet_id)
    with open(SPREADSHEET_ID_FILE, 'w') as file:
        file.write(new_spreadsheet_id)
    print(f"New spreadsheet ID saved: {new_spreadsheet_id}")
    return new_spreadsheet_id

def send_email_with_attachment(gmail_service, pdf_file_path):
    """
    Sends an email with a PDF attachment using the provided Gmail service.

    Args:
        gmail_service: The authenticated Gmail API service instance.
        pdf_file_path (str): The file path to the PDF file to be attached.

    Environment Variables:
        EMAIL_RECIPIENTS (str): The recipient email address. Defaults to 'default_recipient@example.com'.
        EMAIL_CC (str): The CC email address. Defaults to 'default_cc@example.com'.

    Returns:
        None

    Raises:
        googleapiclient.errors.HttpError: If an error occurs while sending the email.

    Example:
        gmail_service = build('gmail', 'v1', credentials=creds)
        send_email_with_attachment(gmail_service, '/path/to/invoice.pdf')
    """
    message = MIMEMultipart()
    recipient_email = os.getenv('EMAIL_RECIPIENTS', 'default_recipient@example.com')
    cc_email = os.getenv('EMAIL_CC', 'default_cc@example.com')
    message['to'] = recipient_email
    message['cc'] = cc_email
    message['subject'] = f'Invoice {datetime.now().strftime("%b %d, %Y")}'

    # Attach the PDF file
    with open(pdf_file_path, 'rb') as pdf_file:
        mime_base = MIMEBase('application', 'pdf')
        mime_base.set_payload(pdf_file.read())
        encoders.encode_base64(mime_base)
        mime_base.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_file_path))
        message.attach(mime_base)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    message_body = {
        'raw': raw_message
    }
    sent_message = gmail_service.users().messages().send(userId='me', body=message_body).execute()
    print(f"Email sent with ID: {sent_message['id']}")

def download_pdf(drive_service, spreadsheet_id):
    """
    Downloads a Google Spreadsheet as a PDF file.

    Args:
        drive_service: The Google Drive service instance used to interact with the Drive API.
        spreadsheet_id: The ID of the Google Spreadsheet to be downloaded.

    Returns:
        None

    Side Effects:
        Creates a PDF file named 'spreadsheet_<spreadsheet_id>.pdf' in the current working directory.
        Prints a message indicating the download is complete.
    """
    pdf_file = (
        drive_service.files()
        .export(fileId=spreadsheet_id, mimeType="application/pdf")
        .execute()
    )
    with open(f"spreadsheet_{spreadsheet_id}.pdf", "wb") as pdf:
        pdf.write(pdf_file)
    print("Spreadsheet downloaded as PDF.")


def create_copy_of_spreadsheet(drive_service, spreadsheet_id):
    """
    Creates a copy of a Google Spreadsheet.
    Args:
        drive_service (googleapiclient.discovery.Resource): The Google Drive service instance.
        spreadsheet_id (str): The ID of the spreadsheet to copy.
    Returns:
        str: The ID of the newly copied spreadsheet.
    Raises:
        googleapiclient.errors.HttpError: If an error occurs during the copy operation.
    """
    
    copy_title = f"Invoice {datetime.now().strftime('%b-%d-%Y')}"
    copy_sheet = {"name": copy_title}
    copied_file = (
        drive_service.files().copy(fileId=spreadsheet_id, body=copy_sheet).execute()
    )
    copied_file_id = copied_file.get("id")
    print(f"Spreadsheet copied with ID: {copied_file_id}")
    return copied_file_id


def update_invoice_total(sheet, spreadsheet_id):
    salary = os.getenv('SALARY', '0')
    update_cell(sheet, spreadsheet_id, "F19", salary)


def update_invoice_number(sheet, spreadsheet_id, values):
    row_12 = 11
    column_f = 5
    next_invoice_number = int(values[row_12][column_f]) + 1
    update_cell(sheet, spreadsheet_id, "F12", str(next_invoice_number))


def update_invoice_date(sheet, spreadsheet_id):
    current_date = datetime.now().strftime("%b %d, %Y")
    update_cell(sheet, spreadsheet_id, "C9", current_date)


def update_cell(sheet, spreadsheet_id: str, cell: str, value: str):
    """
    Updates a specific cell in a Google Sheets document with a new value.

    Args:
        sheet: The Google Sheets service instance.
        cell (str): The A1 notation of the cell to update.
        value (str): The new value to set in the specified cell.

    Returns:
        None
    """
    update_range = cell
    new_value = [[value]]
    body = {"values": new_value}
    sheet.values().update(
        spreadsheetId=spreadsheet_id,
        range=update_range,
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()


if __name__ == "__main__":
    main()
