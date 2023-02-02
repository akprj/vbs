Dim xlApp, xlBook1, xlBook2
Set xlApp = CreateObject("Excel.Application")
Set xlBook1 = xlApp.Workbooks.Open("C:\path\to\existing_workbook.xlsx")

' Check if the second workbook exists
If (Dir("C:\path\to\new_location\new_workbook.xlsx") <> "") Then
  Set xlBook2 = xlApp.Workbooks.Open("C:\path\to\new_location\new_workbook.xlsx")
Else
  Set xlBook2 = xlApp.Workbooks.Add
End If

' Copy data from first workbook to second workbook
xlBook1.Sheets(1).Range("A1:Z100").Copy xlBook2.Sheets(1).Range("A1")

' Save and close the workbooks
xlBook2.SaveAs "C:\path\to\new_location\new_workbook.xlsx"
xlBook2.Close
xlBook1.Close

' Quit the Excel application
xlApp.Quit
