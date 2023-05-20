*** Settings ***
Library     RPA.Browser.Selenium    auto_close=${FALSE}
Library     RPA.HTTP
Library     RPA.Excel.Files
Library     RPA.PDF


*** Tasks ***
Insert The Sales Data For The Week And Export It As A Pdf
    [Documentation]    "Inserts the sales data for the week into the intranet and exports it as a PDF file."
    Open The Intranet Website
    Log In
    Download Excel File
    Fill Form Using Data From Excel
    Collect The Results
    Export The Table As A Pdf
    [Teardown]    Log Out And Close The Browser


*** Keywords ***
Open The Intranet Website
    [Documentation]    Opens the company intranet website.
    Open Available Browser    https://robotsparebinindustries.com/

Log In
    Input Text    alias:Username    maria
    Input Password    alias:Password    thoushallnotpass
    Click Element    alias:LoginButton
    Wait Until Page Contains Element    alias:Salesform

Fill And Submit Form For One Person
    [Arguments]    ${sales_rep}
    Input Text    alias:Firstname    ${sales_rep}[First Name]
    Input Text    alias:Lastname    ${sales_rep}[Last Name]
    Select From List By Value    alias:SalesTarget    ${sales_rep}[Sales Target]
    Input Text    alias:SalesResult    ${sales_rep}[Sales]
    Click Element    alias:SubmitSalesForm

Download Excel File
    [Documentation]    Downloads the Weekly Updated Excel File after Logging Into The Intranet.
    Download
    ...    https://robotsparebinindustries.com/SalesData.xlsx
    ...    %{ROBOT_ROOT}${/}SalesData.xlsx
    ...    overwrite=${TRUE}

Fill Form Using Data From Excel
    Open Workbook    %{ROBOT_ROOT}${/}SalesData.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=${TRUE}
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill And Submit Form For One Person    ${sales_rep}
    END
    Close Workbook

Collect The Results
    Capture Element Screenshot    alias:SalesSummary    ${OUTPUT_DIR}${/}sales_summary.png

Export The Table As A Pdf
    Wait Until Element Is Visible    alias:SalesResultTable
    ${sales_results_html}=    Get Element Attribute    alias:SalesResultTable    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales.pdf

Log Out And Close The Browser
    Click Element    alias:Logout
    Close Browser
