Attribute VB_Name = "WorkbookUtil"
Option Explicit

Sub protectWorkbook(wb As Workbook, wbPassword As String)
    wb.Password = wbPassword
End Sub
