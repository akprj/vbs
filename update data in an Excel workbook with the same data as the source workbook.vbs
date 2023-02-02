Dim xlApp, xlBook1, xlBook2
Set xlApp = CreateObject("Excel.Application")
Set xlBook1 = xlApp.Workbooks.Open("C:\path\to\existing_workbook.xlsx")
Set xlBook2 = xlApp.Workbooks.Open("C:\path\to\new_workbook.xlsx")

' Copy data from first workbook to second workbook
xlBook1.Sheets(1).Range("A1:Z100").Copy xlBook2.Sheets(1).Range("A1")

' Save and close the workbooks
xlBook2.Save
xlBook2.Close
xlBook1.Close

' Quit the Excel application
xlApp.Quit
