Attribute VB_Name = "WorkbookUtil"
Option Explicit

Sub protectWorkbook(wb As Workbook, wbPassword As String)
    'wb : Workbook, the workbook to protect
    'wbPassword : String, the password String
    'Returns : Nothing, it protects the provided workbook
    'with the provided password string
    
    wb.Password = wbPassword
End Sub
