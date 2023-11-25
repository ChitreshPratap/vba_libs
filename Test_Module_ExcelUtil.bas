Attribute VB_Name = "Test_Module_ExcelUtil"

Sub example_toColName()
    
    Dim resultColumn As String
    resultColumn = ExcelUtil.toColName(32)
    Debug.Print resultColumn
    
End Sub

Sub example_worksheetExists()

    Dim wb As Workbook
    Dim sheetExists As Boolean
    Set wb = ThisWorkbook
    sheetExists = ExcelUtil.worksheetExists(wb, "CheckSheet", True)
    Debug.Print sheetExists
    
End Sub

Sub example_getExcelLink1()

    Dim fso As New FileSystemObject
    Dim resultedLink As String
    Dim excelFileFullPath As String
    excelFileFullPath = ThisWorkbook.FullName
    resultedLink = ExcelUtil.getExcelLink1(fso, excelFileFullPath, "Link", "A3")
    Debug.Print resultedLink
    
End Sub

Sub example_getExcelLink2()
    
    Dim resultedLink As String
    Dim wb As Workbook
    Set wb = ThisWorkbook
    resultedLink = ExcelUtil.getExcelLink2(wb, "Link", "B3")
    Debug.Print resultedLink
    
End Sub

Sub example_getExcelLink3()
    
    Dim resultedLink As String
    Dim wb As Workbook
    Dim targetRange As Range
    Set wb = ThisWorkbook
    Set targetRange = wb.worksheets("Sheet1").Range("A3:D7")
    resultedLink = ExcelUtil.getExcelLink3(targetRange)
    Debug.Print resultedLink
    
End Sub

