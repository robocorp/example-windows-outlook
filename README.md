# Send emails with Outlook application

This example requires a working Outlook installation on a Windows system and a user
already configured and logged on.

## Environment variables setup

- `OUTLOOK_ACCOUNT` (REQUIRED): This variable incorporates the Outlook account name.
  (so the app window can be correctly identified and controlled)
- `EMAIL_RECIPIENT` (REQUIRED): To whom to send the e-mail. (can be the same address as
  above)
- `EMAIL_ATTACHMENT` (OPTIONAL): A file path to a file to attach. (can be relative)
- `EMAIL_BODY` (OPTIONAL): A path to a file containing the content of the e-mail body.
  (make sure that this contains an `<ATTACHMENT_TEXT>` placeholder inside, as this
  will be replaced by the file path you want to attach for confirmation purposes)

Our default [body](devdata/email_body.txt) contains:
```text
Greetings!

This e-mail has been sent by a Robocorp robot.
<ATTACHMENT_TEXT>
You can find my source at https://github.com/robocorp/example-windows-outlook

Best Regards,
Mark the Robot
```

## Subtasks

The main task (`Send email with Outlook application on Windows`) calls the following
keywords:

- `Set variables for this task`: Sets the final e-mail body content.
- `Open new Outlook or use the currently open one`: Ensures an active Outlook app open.
- `Press New Email button and send one`: Sends the e-mail with/out an attachment.

## Workarounds in use to improve automation efficiency

- Clipboard is used in some places to paste text into input fields / text areas instead
  of sending keystrokes (faster). The clipboard is cleared with the `Teardown Actions`
  keyword.
- A default e-mail body text is read from the [email_body.txt](devdata/email_body.txt)
  file in the absence of a custom body provided by the `EMAIL_BODY` environment
  variable. The file content should include an `<ATTACHMENT_TEXT>` text inside, which
  will be replaced with the file name of the attachment (if one is provided).

## Further reading

- [Desktop automation](https://robocorp.com/docs/development-guide/desktop)
- [Desktop robots in the Portal](https://robocorp.com/portal/collection/desktop-automation)
- [`RPA.Windows` library](https://robocorp.com/docs/libraries/rpa-framework/rpa-windows)
- [`RPA.Desktop` library](https://robocorp.com/docs/libraries/rpa-framework/rpa-desktop)
- [`RPA.FileSystem` library](https://robocorp.com/docs/libraries/rpa-framework/rpa-filesystem)
