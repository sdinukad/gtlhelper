# gtlhelper
Capture listings on pokemmo gtl using ocr and update to a google sheet
# Pokemmo GTL Helper

A simple desktop application to help track listings from Pokemmo's Global Trade Link (GTL) by capturing screen regions or clipboard images, performing OCR, and saving the data to Google Sheets and/or a local CSV file.

## Features

*   Capture specific GTL listing lines using an in-app region selection tool.
*   Process listing data from images copied to the clipboard (e.g., via Win+Shift+S).
*   Uses Tesseract OCR to extract text from images.
*   Parses OCR'd text to identify Item Name, Price, and Date.
*   Saves structured data to:
    *   Google Sheets (configurable)
    *   Local CSV file (`gtl_listings.csv`) (configurable)
*   Toggleable UI layout (horizontal/vertical).
*   Always-on-top window for convenience.

## Prerequisites

Before running this application, you will need:

1.  **Python 3.x:** Download from [python.org](https://www.python.org/). Make sure to check "Add Python to PATH" during installation.
2.  **Tesseract OCR Engine:**
    *   **Windows:** Download the installer from the [Tesseract at UB Mannheim page](https://github.com/UB-Mannheim/tesseract/wiki). During installation, ensure "English" language data is included. Note the installation path (e.g., `C:\Program Files\Tesseract-OCR`).
    *   **macOS:** Use Homebrew: `brew install tesseract`
    *   **Linux (Debian/Ubuntu):** `sudo apt update && sudo apt install tesseract-ocr`
3.  **Google Cloud Project & Google Sheets API:**
    *   You need a Google Cloud Platform (GCP) project.
    *   The Google Sheets API must be enabled for this project.
    *   You need to create Service Account credentials.

## Setup Instructions

1.  **Clone or Download this Repository:**
    ```bash
    # If using git
    git clone <repository_url>
    cd <repository_folder>
    ```
    Or download the files as a ZIP and extract them.

2.  **Install Python Dependencies:**
    Open a terminal or command prompt in the project folder and run:
    ```bash
    pip install -r requirements.txt
    ```
    (You will need to create a `requirements.txt` file. See below.)

3.  **Configure Tesseract Path (if needed):**
    Open the Python script (e.g., `gtl_helper_app.py`). Near the top, find the `TESSERACT_CMD_PATH` variable.
    *   **Windows Users:** If you installed Tesseract to a non-default location, update this path to point to your `tesseract.exe` (e.g., `r'C:\MyCustomPath\Tesseract-OCR\tesseract.exe'`). If Tesseract is correctly in your system PATH, the script might find it automatically.
    *   **macOS/Linux Users:** This usually doesn't need changing if Tesseract was installed via package managers.

4.  **Set up Google Sheets API Credentials:**

    a.  **Create a Service Account and Key:**
        1.  Go to the [Google Cloud Console](https://console.cloud.google.com/).
        2.  Select your GCP project.
        3.  Navigate to **"IAM & Admin" > "Service Accounts"**.
        4.  Click **"+ CREATE SERVICE ACCOUNT"**.
            *   Give it a name (e.g., "gtl-sheet-updater").
            *   Click "CREATE AND CONTINUE".
        5.  **Grant this service account access to project:**
            *   Select a role: **"Editor"** is a simple option that grants necessary permissions for Sheets. For more restricted access, you could create a custom role with `drive.file` and `spreadsheets` permissions.
            *   Click "CONTINUE".
        6.  Skip "Grant users access to this service account" (optional) and click "DONE".
        7.  Find the service account you just created in the list. Click on its email address.
        8.  Go to the **"KEYS"** tab.
        9.  Click **"ADD KEY" > "Create new key"**.
        10. Choose **JSON** as the key type and click **"CREATE"**.
        11. A JSON file will be downloaded. **Rename this file to `credentials.json`** and place it in the **same directory** as the Python script. **Treat this file like a password – keep it secure!**

    b.  **Get Your Google Sheet ID:**
        1.  Create a new Google Sheet or open the one you want to use.
        2.  Look at the URL in your browser's address bar. It will look something like this:
            `https://docs.google.com/spreadsheets/d/THIS_IS_THE_SPREADSHEET_ID/edit#gid=0`
        3.  Copy the long string of characters between `/d/` and `/edit`. This is your **Spreadsheet ID**.
        4.  Open the Python script and find the `SPREADSHEET_ID` variable. Paste your copied ID there:
            ```python
            SPREADSHEET_ID = 'YOUR_COPIED_SPREADSHEET_ID_HERE'
            ```

    c.  **Share the Google Sheet with the Service Account:**
        1.  In the Google Sheet, click the **"Share"** button (usually top right).
        2.  In the "Add people and groups" field, paste the **email address of the service account** you created (e.g., `gtl-sheet-updater@your-project-id.iam.gserviceaccount.com`). You can find this email on the Service Account details page in GCP.
        3.  Ensure it has **"Editor"** permission for the sheet.
        4.  Click "Send" (or "Share").

5.  **Configure Worksheet Name (Optional):**
    If your target sheet within the Google Spreadsheet is not named "Sheet1", update the `WORKSHEET_NAME` variable in the script:
    ```python
    WORKSHEET_NAME = 'YourActualSheetName'
    ```

## Running the Application

1.  Navigate to the project directory in your terminal/command prompt.
2.  Run the script:
    ```bash
    python gtl_helper_app.py 
    ```
    (Or whatever you named your main Python file).

## Usage

1.  **Capture Region & Preview:**
    *   Click "Capture Region & Preview".
    *   The main app window will minimize.
    *   A semi-transparent overlay will cover your screen. Click and drag to select the GTL listing line.
    *   Release the mouse. The app window will reappear, and the OCR'd data will be shown in the preview area.
2.  **Preview Clipboard:**
    *   Copy an image of a GTL listing to your clipboard (e.g., using Windows: `Win+Shift+S`, then select the area).
    *   Click "Preview Clipboard" in the app. The OCR'd data will be shown in the preview.
3.  **Save Previewed:**
    *   Once data is in the preview and looks correct, click "Save Previewed".
    *   The structured data will be saved to Google Sheets and/or the local `gtl_listings.csv` file, depending on the checkbox settings.
4.  **Settings:**
    *   Use the checkboxes to enable/disable saving to Google Sheets or CSV.
5.  **Toggle Layout:**
    *   Switch the app's button layout between horizontal and vertical for convenient placement.

## `requirements.txt` File

Create a file named `requirements.txt` in the same directory as your script with the following content:
Pillow>=9.0.0
pytesseract>=0.3.0
gspread>=5.0.0
google-auth>=2.0.0
google-auth-oauthlib>=0.4.0
google-auth-httplib2>=0.1.0
Users can then install these with `pip install -r requirements.txt`. (Note: `google-auth-oauthlib` and `google-auth-httplib2` are often pulled in as dependencies of `gspread` or `google-auth` but it's good to list them for clarity if specific versions were relevant during development).

## Troubleshooting

*   **Tesseract Not Found:** Ensure Tesseract is installed and the `TESSERACT_CMD_PATH` in the script is correct or Tesseract is in your system's PATH.
*   **Google Sheets Errors:**
    *   Double-check your `SPREADSHEET_ID`.
    *   Ensure `credentials.json` is correctly named and in the same folder.
    *   Verify the service account email has "Editor" access to your Google Sheet.
    *   Make sure the Google Sheets API is enabled in your GCP project.
*   **OCR Accuracy:**
    *   Ensure your screen captures are clear and tightly cropped around the relevant text.
    *   Experiment with the `scale_factor` and `threshold` values in the `preprocess_image` function if OCR results are consistently poor.
*   **NameError: name 'KNOWN_INPUT_TYPES' is not defined:** Ensure the `KNOWN_INPUT_TYPES = ["GTL", "GM"]` list is defined globally in the script.



4. Set up Google OAuth 2.0 Credentials (For User Authentication)
This application uses OAuth 2.0 to allow it to access your Google Sheets on your behalf. This means you will log in with your own Google account the first time you run the app, and the app will only access the sheets you intend for it to use. You will need to create OAuth 2.0 credentials within your own Google Cloud Project.
a. Create or Select a Google Cloud Platform (GCP) Project:
Go to the Google Cloud Console.
If you don't have a project, create a new one. If you do, select the project you want to use.
b. Enable Necessary APIs:
In your selected GCP project, navigate to "APIs & Services" > "Library".
Search for and enable the "Google Sheets API".
Search for and enable the "Google Drive API" (this is often required by the gspread library for discovering and accessing spreadsheets).
c. Configure the OAuth Consent Screen:
This screen is what users will see when asked to grant your application access.
Navigate to "APIs & Services" > "OAuth consent screen".
User Type: Choose "External" if you want anyone with a Google account to be able to use it (once your app is verified by Google if needed, or if you add them as test users). Choose "Internal" if this is only for users within your Google Workspace organization. For personal use or sharing with friends, "External" is common.
Click "CREATE".
App information:
App name: Enter a name for the application, e.g., "My GTL Helper" or "GTL OCR Tool".
User support email: Select your email address.
App logo: Optional.
Developer contact information: Enter your email address.
Click "SAVE AND CONTINUE".
Scopes:
Click "ADD OR REMOVE SCOPES".
Search for or manually find the Google Sheets API scope: https://www.googleapis.com/auth/spreadsheets (allows reading and writing to sheets).
Select it and click "UPDATE".
Click "SAVE AND CONTINUE".
Test users (Important for Initial Use):
While your app is in "Testing" publishing status (which it will be initially), only users added here can authorize the app.
Click "+ ADD USERS" and add your own Google email address(es) that you will use to test the application. If you plan to share with specific friends for testing, add their Google emails too.
Click "SAVE AND CONTINUE".
Review the summary and click "BACK TO DASHBOARD".
For now, you can leave the "Publishing status" as "Testing". If you later want anyone to use it without being on the test user list, you'd click "PUBLISH APP", which might trigger a Google verification process depending on the scopes and app usage.
d. Create OAuth 2.0 Client ID Credentials for a Desktop App:
Navigate to "APIs & Services" > "Credentials".
Click "+ CREATE CREDENTIALS" at the top.
Select "OAuth client ID".
Application type: Choose "Desktop app".
Name: Give it a name (e.g., "GTL Helper Desktop Client").
Click "CREATE".
A dialog box will appear showing your "Client ID" and "Client Secret". You don't need to copy these directly from here. Click "DOWNLOAD JSON" (or there might be a download icon).
The downloaded file will likely be named something like client_secret_[...long_string...].json.
Rename this downloaded file to exactly client_secret_desktop.json.
Place this client_secret_desktop.json file in the same directory as the Python script (e.g., gtlhelper.py). This file allows the application to identify itself to Google during the OAuth flow.
e. Get Your Target Google Sheet ID (User Choice):
The application will now allow you to set the target Google Sheet from within its settings.
When you run the app, after authenticating, go to the app's settings (expand "⚙️ Settings").
You will see a field for "Sheet URL/ID".
Create a new Google Sheet or open the one you want the app to write data to.
From the Google Sheet's URL in your browser (e.g., https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit), copy the SPREADSHEET_ID.
Paste this ID (or the full URL) into the entry field in the app and click "Set Sheet".
The app will save this ID for future sessions in a local app_settings.json file.
Important First Run:
The first time you run the application after setting up client_secret_desktop.json, your web browser will open.
You will be asked to log into your Google account (if not already logged in) and then to grant the "GTL Helper" (or whatever you named it on the consent screen) permission to access your Google Sheets.
After you grant permission, you'll be redirected, and a token.json file will be created in the script's directory. This file stores your authorization so you don't have to log in via the browser every time. Do not share your token.json file.
