# Google Sheets and Gmail Automation

This project automates the process of copying a Google Spreadsheet, updating specific cells, downloading it as a PDF, and sending it as an email attachment using the Google Sheets, Google Drive, and Gmail APIs.

## Prerequisites

- Python 3.x
- Google Cloud Project with the following APIs enabled:
  - Google Sheets API
  - Google Drive API
  - Gmail API
- OAuth 2.0 credentials (client ID and client secret) from the Google Cloud Console
- `credentials.json` file containing your OAuth 2.0 credentials

## Installation

1. Clone the repository:
   ```sh
   git clone https://github.com/Juan2418/facturainator
   cd facturainator
   ```

2. Create and activate a virtual environment:
   ```sh
   python3 -m venv venv
   source venv/bin/activate
   ```

3. Install the required packages:
   ```sh
   pip install -r requirements.txt
   ```

4. Create a `.env` file in the project directory and add the following environment variables:
   ```env
   SPREADSHEET_ID=your_spreadsheet_id
   EMAIL_RECIPIENTS=recipient1@example.com,recipient2@example.com
   EMAIL_CC=cc@example.com
   SALARY=0
   ```

5. Place your `credentials.json` file in the project directory.

## Usage

1. Run the script:
   ```sh
   python main.py
   ```

2. The script will:
   - Copy the specified Google Spreadsheet
   - Update specific cells in the copied spreadsheet
   - Download the copied spreadsheet as a PDF
   - Send an email with the PDF attached to the specified recipients

## Project Structure

```
.
├── .env
├── .gitignore
├── README.md
├── credentials.json
├── main.py
├── requirements.txt
└── spreadsheet_id.txt
```

## .gitignore

The [`.gitignore`](.gitignore) file includes common Python exclusions, as well as specific exclusions for Google API credentials, token files, and the [`spreadsheet_id.txt`](spreadsheet_id.txt) file. It also excludes the [`.env`](.env) file used for environment variables.
