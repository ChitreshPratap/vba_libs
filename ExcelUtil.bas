Attribute VB_Name = "ExcelUtil"

Function toColName(colNumber As Integer) As String
    'It returns the Alphabetical column name of the corresponding Integral column number
    toColName = Split(Cells(1, colNumber).EntireColumn.Address(0, 0), ":")(0)
    
End Function

Function getExcelLink1(fso As Object, excelFileName As String, sheetName As String, cellRange As String) As String
    'fso : FileSystemObject
    'It returns the excel link of the given workbook, sheetName and cell range.
    'Workbook not need to be open
    
    Dim fName As String
    Dim resultPathLink As String
    fName = Dir(excelFileName)
    resultPathLink = "'" & fso.GetParentFolderName(excelFileName) & "\[" & fso.getFileName(fName) & "]" & "'!" & cellRange
    getExcelLink1 = resultPathLink
    
End Function

Function getExcelLink2(wb As Workbook, sheetName As String, cellRange As String) As String
    
    'It returns the excel link of the given workbook, sheet name and cell range
    'Workbook must be open
    
    Dim fName As String
    Dim resultPathLink As String
    resultPathLink = "'[" & wb.Name & "]" & sheetName & "'!" & cellRange
    getExcelLink2 = resultPathLink
    
End Function

Function getExcelLink3(cellRange As Range) As String
    'It returns the excel link of the given cell range
    'Workbook must be open
    Dim fName As String
    Dim resultPathLink As String
    resultPathLink = "'[" & cellRange.Parent.Parent.Name & "]" & cellRange.Parent.Name & "'!" & cellRange.Address
    getExcelLink3 = resultPathLink
    
End Function

Function worksheetExists(wb As Workbook, sheetName As String, ifNotExistsRaiseError As Boolean) As Boolean
    'It checks whether given sheetName exists inside given workbook or not.
    'workbook must be open
    'It returns Boolean if : IfNotExistsRaiseError is False
    'it raise SheetNotFoundError if Sheet is not found and if IfNotExistsRaiseError is True
    
    Dim sht As Worksheet
    On Error Resume Next
    Set sht = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ifNotExistsRaiseError Then
        If sht Is Nothing Then
            Err.Raise vbObjectError + 1, "ExcelUtil.worksheetExists", "SheetNotFoundError : '" & sheetName & "' not found" _
            & vbNewLine & "File : '" & wb.FullName & "'"
        End If
    Else
        worksheetExists = Not sht Is Nothing
    End If
    
End Function

