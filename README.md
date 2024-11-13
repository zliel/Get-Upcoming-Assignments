# Get-UpcomingAssignments.ps1

## Description

This script retrieves upcoming assignments from the Canvas LMS API and sends an email to the specified recipient with a table of the assignments attached. I found this script useful for keeping track of upcoming assignments in my courses, and I hope you find it useful as well.

## Prerequisites

- **PowerShell**: Ensure that you have PowerShell installed on your machine. You can download PowerShell from the [official website](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell).
- **Canvas API Token**: You will need a Canvas API token to authenticate with the Canvas API. You can generate an API token by following the instructions [here](https://community.canvaslms.com/t5/Admin-Guide/How-do-I-manage-API-access-tokens-as-an-admin/ta-p/89).
- **SMTP-Server**: You will need access to an SMTP server to send the Email. For my use case, I used Gmail's SMTP server. You can find more information on how to set up Gmail's SMTP server [here](https://support.google.com/a/answer/176600?hl=en).
- **Pandoc**: You will need Pandoc to put the Markdown table into a PDF. You can download Pandoc from the [official website](https://pandoc.org/installing.html).

## Configuration

After cloning the repository, and before running the script, you will need to configure the following variables in the `config.json` file:

```json
{
    "access_token": "YOUR_CANVAS_ACCESS_TOKEN_HERE",
    "semester_start": "YYYY-MM-DD",
    "From": "YOUR_EMAIL_HERE",
    "To": "RECIPIENT_EMAIL_HERE",
    "Attachment": "YOUR_OUTPUT_PDF_FILENAME_HERE",
    "Body": "Here's your next week's assignments!",
    "SMTPServer": "YOUR_SMTP_SERVER_HERE",
    "SMTPPort": "YOUR_SMTP_PORT_HERE",
    "SMTPPassword": "YOUR_SMTP_PASSWORD_HERE"
}
```

## Usage

To run the script, do the following:

1. Clone the repository:

    ```bash
    git clone https://github.com/zliel/Get-Upcoming-Assignments.git
    ```

2. Navigate to the directory:

    ```bash
    cd Get-UpcomingAssignments
    ```

3. Modify the `config.json` file with your Canvas API token, email addresses, and SMTP server information.

4. Run the script:

    ```bash
    .\Get-UpcomingAssignments.ps1
    ```

    - There's an optional paramter `-Cleanup` (or `-c`) that you can use to remove the generated markdown and PDF files after sending the email.
