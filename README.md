# Send emails with Outlook application

Example requires working Outlook installation on Windows system

## Environment variables

    - `OUTLOOK_ACCOUNT` (REQUIRED). This variable defines Outlook account name.
    - `EMAIL_RECIPIENT` (REQUIRED). This variable defines recipient of the email.
    - `EMAIL_ATTACHMENT` (OPTIONAL). This variable defines filepath to the email attachment.
    - `EMAIL_BODY` (OPTIONAL). This variable defines filepath to the email body (text file).

## Task steps

Task consist of three main level keywords:

- `Set Variables for the Task`
- `Open Outlook or Use already open Outlook`
- `Use New Email button to Send Email`

### Set Variables for the Task

Verifies that mandatory environment variables exist and sets necessary variables for the task.

### Open Outlook or Use already open Outlook

Opens Outlook application if it is not open or gets control of already open Outlook application window.

### Use New Email button to Send Email

Click `New Email` button on fill in information into opened dialog and send the email.

## Workarounds in use to improve automation effiency

- Clipboard is used in some places to paste text into input fields / textareas instead of sending keystrokes. The clipboard is cleared just in case in the `Task Teardown` step.
- Email body text is read from the `email_body.txt` or file defined by `EMAIL_BODY` environment variable if either exists. File content can include `<ATTACHMENT_TEXT>` text which will be replaced with the text information the filename of the attachment.
- Robot Framework IF/ELSE syntax is in use to make process more readable.
