# Outlook to PDF Exporter

A Python script that exports emails from a Microsoft Outlook account to a single PDF file and saves any attachments locally.

## Description

This command-line tool connects to the local Outlook desktop application, scans the Inbox and Sent Items folders, and finds all emails sent to or from a specified email address. It then compiles the body text of these emails into a single PDF and extracts all attachments into a separate folder for archival purposes.

## Features

- **Interactive**: Prompts the user for the target email address at runtime.
- **Automatic Naming**: Generates PDF and attachment folder names based on the target email.
- **Attachment Extraction**: Saves all attachments from matching emails into a dedicated folder.
- **Progress Bar**: Uses `tqdm` to display a live progress bar during the email scan.
- **Verbose Logging**: Provides real-time feedback on the script's progress.

## Prerequisites

- Windows Operating System
- Microsoft Outlook (Desktop version) installed and configured
- Python 3.6+

## Installation

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/rsgrizz/Outlook-to-PDF-Exporter.git](https://github.com/rsgrizz/Outlook-to-PDF-Exporter.git)
    cd Outlook-to-PDF-Exporter
    ```

2.  **Install the required libraries using pip:**
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  Make sure Microsoft Outlook is running.
2.  Run the script from your terminal:
    ```bash
    python export_script.py
    ```

3.  Follow the on-screen prompt to enter the email address you want to archive.
    ```
    Enter the target email address to search for: user@example.com
    ```
4.  The script will display its progress and, upon completion, you will find the generated PDF file and attachments folder in the same directory.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
