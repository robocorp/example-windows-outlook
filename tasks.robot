*** Settings ***
Library    RPA.Desktop    WITH NAME    Desktop
Library    RPA.FileSystem
Library    RPA.Windows    WITH NAME    Windows
Library    OperatingSystem
Library    String

Suite Setup    Set Wait Time    1


*** Variables ***
${DEFAULT_EMAIL}    mika@beissi.onmicrosoft.com
${ACCOUNT_NAME}    %{OUTLOOK_ACCOUNT=${DEFAULT_EMAIL}}
${EMAIL_RECIPIENT}    %{EMAIL_RECIPIENT=${DEFAULT_EMAIL}}
${EMAIL_BODY}    %{EMAIL_BODY=devdata${/}email_body.txt}
${DEFAULT_MAIL_BODY}    Default message from the RPA process
${EMAIL_ATTACHMENT}    %{EMAIL_ATTACHMENT=${None}}
${ATTACHMENT_PLACEHOLDER}    <ATTACHMENT_TEXT>
${SUBJECT}    %{SUBJECT=Coming from Mark}

${LOCATOR_NEW_EMAIL}    name:"New Email" type:Button
${LOCATOR_NEW_MESSAGE}    subname:Message type:Window
${LOCATOR_EMAIL_TO}    name:To type:Edit
${LOCATOR_EMAIL_SUBJECT}    name:Subject type:Edit
${LOCATOR_EMAIL_BODY}    id:Body type:Edit
${LOCATOR_EMAIL_SEND}    name:Send type:Button
${LOCATOR_INSERT_FILE}    name:"File name:" type:Edit

${NEED_TO_CLOSE}    ${False}
${WAIT_TIME_BEFORE_CLOSE}    10s


*** Keywords ***
Is there a window with the provided title already open
    [Arguments]    ${expected_title}

    @{windows} =    List Windows
    FOR    ${window}    IN    @{windows}
        IF    "${expected_title}" in "${window}[title]"
            RETURN    ${True}
        END
    END

    RETURN    ${False}

Open new Outlook or use the currently open one
    [Documentation]    Opens the Outlook application if it is not open or gets into
    ...    control of an already open Outlook application window.

    ${isopen} =    Is there a window with the provided title already open
    ...    ${outlook_title}
    IF    not ${isopen}
        Windows Run    Outlook    wait_time=${5}
        Set Global Variable    ${NEED_TO_CLOSE}    ${True}
    END
    Control Window    desktop > subname:"${outlook_title}"

Pate text into element from clipboard
    [Arguments]    ${target}    ${text}

    Windows.Click    ${target}
    Desktop.Set Clipboard Value    ${text}
    Send Keys    ${target}    keys={LCTRL}v

Set variables for this task
    [Documentation]    Sets necessary variables for the task to work properly.

    Set Task Variable    ${outlook_title}    ${ACCOUNT_NAME} - Outlook
    ${email_body_exists} =    Does File Exist    ${EMAIL_BODY}
    IF    ${email_body_exists}
        ${email_body} =    Read File    ${EMAIL_BODY}
        Set Task Variable    ${email_body}    ${email_body}
    ELSE
        Set Task Variable    ${email_body}    ${DEFAULT_MAIL_BODY}
    END

Add an attachment if one was provided
    ${attachment_filepath} =    Absolute Path    ${EMAIL_ATTACHMENT}

    ${file_exists} =    Does File Exist    ${attachment_filepath}
    IF    ${file_exists}
        Send Keys    keys={LALT}NAFB
        Set Value    ${LOCATOR_INSERT_FILE}    ${attachment_filepath}    enter=${True}

        ${filename} =    Get File Name    ${attachment_filepath}
        ${email_body} =    Replace String    ${email_body}
        ...    ${ATTACHMENT_PLACEHOLDER}
        ...    \nThere is an attachment in this e-mail: \n\t${filename}\n
    ELSE
        ${email_body} =    Replace String    ${email_body}
        ...    ${ATTACHMENT_PLACEHOLDER}
        ...    ${EMPTY}
    END
    Set Task Variable    ${email_body}    ${email_body}

Press New Email button and send one
    [Documentation]    Clicks `New Email` button and fills in information into the
    ...    opened dialog, then finally sends the e-mail. (with/out an attachment)

    Windows.Click    ${LOCATOR_NEW_EMAIL}
    # All future actions are relative to this newly set anchor window. (the window for
    #  composing the message)
    Set Anchor    desktop > ${LOCATOR_NEW_MESSAGE}

    Send Keys    ${LOCATOR_EMAIL_TO}    keys=${EMAIL_RECIPIENT}
    Send Keys    ${LOCATOR_EMAIL_SUBJECT}    keys=${SUBJECT}
    Add an attachment if one was provided
    Pate text into element from clipboard    ${LOCATOR_EMAIL_BODY}    ${email_body}
    Windows.Click    ${LOCATOR_EMAIL_SEND}

    [Teardown]    Clear Anchor

Teardown Actions
    Clear Clipboard
    IF    ${NEED_TO_CLOSE}
        Log To Console    Outlook will be closed in ${WAIT_TIME_BEFORE_CLOSE}...
        Sleep    ${WAIT_TIME_BEFORE_CLOSE}
        Close Current Window
    END


*** Tasks ***
Send email with Outlook application on Windows
    Set variables for this task
    Open new Outlook or use the currently open one
    Press New Email button and send one
    Log    Task complete!

    [Teardown]    Teardown Actions
