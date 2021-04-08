*** Settings ***
Library           RPA.Desktop.Windows
Library           RPA.Desktop    WITH NAME    Desktop
Library           String

*** Variables ***
${ACCOUNT_NAME}    mika@beissi.onmicrosoft.com
${LOCATOR_NEW_EMAIL}    name:'New Email' and type:Button
${LOCATOR_EMAIL_TO}    name:To and type:Edit
${LOCATOR_EMAIL_SUBJECT}    name:Subject and type:Edit
${LOCATOR_EMAIL_BODY}    type:Document
${LOCATOR_EMAIL_SEND}    name:Send and type:Button
${LOCATOR_INSERT_FILE}    name:'File name:' and type:Edit
${SHORTCUT_INSERT_FILE}    %NAFB
${ATTACHMENT_FILEPATH}    c:\\koodi\\testdata\\invoice.pdf

*** Keywords ***
Input Encoded Text
    [Arguments]    ${locator}    ${text}
    ${text}=    Replace String    ${text}    ${SPACE}    {VK_SPACE}
    ${text}=    Replace String    ${text}    \n    {ENTER}
    Type Into    ${locator}    ${text}

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
Minimal task
    ${outlook_title}=    Set Variable    ${ACCOUNT_NAME} - Outlook
    ${isopen}=    Is Window With Title Already Open    ${outlook_title}
    IF    ${isopen}
    Open Dialog    ${outlook_title}    wildcard=True
    ELSE
    Open From Search    outlook    ${outlook_title}    wildcard=True
    END
    Mouse Click    ${LOCATOR_NEW_EMAIL}
    Refresh Window
    Open Dialog    Untitled    wildcard=True
    Input Encoded Text    ${LOCATOR_EMAIL_TO}    mika@robocorp.com
    Input Encoded Text    ${LOCATOR_EMAIL_SUBJECT}    Coming from Robot
    Input Encoded Text    ${LOCATOR_EMAIL_BODY}    Message from the RPA process
    Send Keys    ${SHORTCUT_INSERT_FILE}
    Sleep    2s
    Desktop.Type Text    ${ATTACHMENT_FILEPATH}
    Desktop.Press Keys    enter
    #Mouse Click    ${LOCATOR_EMAIL_SEND}
    Log    Done.
