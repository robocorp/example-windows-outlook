*** Settings ***
Library           RPA.Desktop.Windows
Library           RPA.Desktop    WITH NAME    Desktop
Library           RPA.FileSystem
Library           OperatingSystem
Library           String

*** Variables ***
${ACCOUNT_NAME}    mika@beissi.onmicrosoft.com
${DEFAULT_MAIL_SUBJECT}    Coming from Robot
${DEFAULT_MAIL_BODY}    Default message from the RPA process
${LOCATOR_NEW_EMAIL}    name:'New Email' and type:Button
${LOCATOR_EMAIL_TO}    name:To and type:Edit
${LOCATOR_EMAIL_SUBJECT}    name:Subject and type:Edit
${LOCATOR_EMAIL_BODY}    type:Document
${LOCATOR_EMAIL_SEND}    name:Send and type:Button
${LOCATOR_INSERT_FILE}    name:'File name:' and type:Edit
${SHORTCUT_INSERT_FILE}    %NAFB
${EMAIL_BODY_FILEPATH}    ${CURDIR}${/}email_body.txt
${NEED_TO_CLOSE}    ${FALSE}

*** Keywords ***
Input Encoded Text
    [Arguments]    ${text}    ${locator}=${NONE}
    ${text}=    Replace String    ${text}    ${SPACE}    {VK_SPACE}
    ${text}=    Replace String    ${text}    \n    {ENTER}
    IF    "${locator}" != "${NONE}"
        Type Into    ${locator}    ${text}
    ELSE
        Send Keys    ${text}{ENTER}
    END

*** Keywords ***
Is Window With Title Already Open
    [Arguments]    ${expected_title}
    ${windowlist}=    Get Window List
    FOR    ${window}    IN    @{windowlist}
        IF    "${expected_title}" in "${window}[title]"
            Return From Keyword    ${TRUE}
        END
    END
    [Return]    ${FALSE}

*** Keywords ***
Open Outlook or use already open Outlook
    ${isopen}=    Is Window With Title Already Open    ${outlook_title}
    IF    ${isopen}
        Open Dialog    ${outlook_title}    wildcard=True
    ELSE
        Open From Search    outlook    ${outlook_title}    wildcard=True    timeout=20
        Set Global Variable    ${NEED_TO_CLOSE}    ${TRUE}
    END

*** Keywords ***
Paste text from clipboard to element
    [Arguments]    ${text}    ${target}    ${method}=mouse
    Desktop.Set Clipboard Value    ${text}
    IF    "${method}" == "mouse"
        Mouse Click    ${target}
    ELSE IF    "${method}" == "keys"
        Send Keys    ${target}
    END
    Send Keys    ^v{ENTER}

*** Keywords ***
Set Variables for the Task
    Environment Variable Should Be Set    EMAIL_RECIPIENT
    ...    Environment variable 'EMAIL_RECIPIENT' needs to be set
    Environment Variable Should Be Set    OUTLOOK_ACCOUNT
    ...    Environment variable 'OUTLOOK_ACCOUNT' needs to be set
    Set Task Variable    ${email_recipient}    %{EMAIL_RECIPIENT}
    Set Task Variable    ${account_name}    %{OUTLOOK_ACCOUNT}
    Set Task Variable    ${outlook_title}    ${ACCOUNT_NAME} - Outlook
    ${email_body_exist}=    Does File Exist    %{EMAIL_BODY=${EMAIL_BODY_FILEPATH}}
    IF    ${email_body_exist}
        ${email_body}=    Read File    %{EMAIL_BODY=${EMAIL_BODY_FILEPATH}}
        Set Task Variable    ${email_body}    ${email_body}
    ELSE
        Set Task Variable    ${email_body}    ${DEFAULT_MAIL_BODY}
    END

*** Keywords ***
Add Attachment If It Has Been Given
    Set Task Variable    ${attachment_filepath}    %{EMAIL_ATTACHMENT=${NONE}}
    ${file_exists}=    Does File Exist    ${attachment_filepath}
    IF    ${file_exists}
        Paste text from clipboard to element    ${attachment_filepath}    ${SHORTCUT_INSERT_FILE}    method=keys
        ${filename}=    Get File Name    ${attachment_filepath}
        ${email_body}=    Replace String    ${email_body}
        ...    <ATTACHMENT_TEXT>
        ...    \nThere is attachment in the email: \n\t${filename}\n
    ELSE
        ${email_body}=    Replace String    ${email_body}
        ...    <ATTACHMENT_TEXT>
        ...    ${EMPTY}
    END
    Set Task Variable    ${email_body}    ${email_body}
    Sleep    5s

*** Keywords ***
Use New Email button to Send Email
    Mouse Click    ${LOCATOR_NEW_EMAIL}
    Refresh Window
    Open Dialog    Untitled    wildcard=True
    Add Attachment If It Has Been Given
    Input Encoded Text    ${email_recipient}    ${LOCATOR_EMAIL_TO}
    Input Encoded Text    %{SUBJECT=${DEFAULT_MAIL_SUBJECT}}    ${LOCATOR_EMAIL_SUBJECT}
    Paste text from clipboard to element    ${email_body}    ${LOCATOR_EMAIL_BODY}
    Mouse Click    ${LOCATOR_EMAIL_SEND}

*** Keywords ***
Teardown Actions
    Clear Clipboard
    IF    ${NEED_TO_CLOSE}
        Open Dialog    ${outlook_title}    wildcard=True
        Send Keys    %F%X
    END

*** Tasks ***
Sending Email From Outlook application
    [Teardown]    Teardown Actions
    Set Variables for the Task
    Open Outlook or use already open Outlook
    Use New Email button to Send Email
    Log    Done.
