*** Settings ***
Library           RPA.Desktop.Windows
Library           RPA.Desktop    WITH NAME    Desktop
Library           RPA.FileSystem
Library           String

*** Variables ***
${ACCOUNT_NAME}    mika@beissi.onmicrosoft.com
${DEFAULT_MAIL_RECIPIENT}    mika@robocorp.com
${DEFAULT_MAIL_SUBJECT}    Coming from Robot
${DEFAULT_MAIL_BODY}    Default message from the RPA process
${LOCATOR_NEW_EMAIL}    name:'New Email' and type:Button
${LOCATOR_EMAIL_TO}    name:To and type:Edit
${LOCATOR_EMAIL_SUBJECT}    name:Subject and type:Edit
${LOCATOR_EMAIL_BODY}    type:Document
${LOCATOR_EMAIL_SEND}    name:Send and type:Button
${LOCATOR_INSERT_FILE}    name:'File name:' and type:Edit
${SHORTCUT_INSERT_FILE}    %NAFB
${ATTACHMENT_FILEPATH}    ${CURDIR}${/}invoice.pdf
${EMAIL_BODY_FILEPATH}    ${CURDIR}${/}email_body.txt

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

*** Tasks ***
Sending Email From Outlook application
    ${outlook_title}=    Set Variable    ${ACCOUNT_NAME} - Outlook
    ${does_email_body_exist}=    Does File Exist    ${EMAIL_BODY_FILEPATH}
    ${email_body}=    Set Variable    %{BODY=${DEFAULT_MAIL_BODY}}
    IF    ${does_email_body_exist}
        ${email_body}=    Read File    ${EMAIL_BODY_FILEPATH}
    END
    ${isopen}=    Is Window With Title Already Open    ${outlook_title}
    IF    ${isopen}
        Open Dialog    ${outlook_title}    wildcard=True
    ELSE
        Open From Search    outlook    ${outlook_title}    wildcard=True    timeout=20
    END
    Mouse Click    ${LOCATOR_NEW_EMAIL}
    Refresh Window
    Open Dialog    Untitled    wildcard=True
    Input Encoded Text    %{RECIPIENT=${DEFAULT_MAIL_RECIPIENT}}    ${LOCATOR_EMAIL_TO}
    Input Encoded Text    %{SUBJECT=${DEFAULT_MAIL_SUBJECT}}    ${LOCATOR_EMAIL_SUBJECT}
    Desktop.Set Clipboard Value    ${email_body}
    Mouse Click    ${LOCATOR_EMAIL_BODY}
    Send Keys    ^v{ENTER}
    Send Keys    ${SHORTCUT_INSERT_FILE}
    Sleep    2s
    Desktop.Set Clipboard Value    ${ATTACHMENT_FILEPATH}
    Send Keys    ^v{ENTER}
    Send Keys    {ENTER}
    Mouse Click    ${LOCATOR_EMAIL_SEND}
    Log    Done.
