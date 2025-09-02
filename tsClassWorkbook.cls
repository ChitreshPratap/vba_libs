VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tsClassWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim var_Workbook As Workbook
Dim var_className As String

Sub Class_Initialize()
    Set var_Workbook = Nothing
    var_className = "tsClassWorkbook"
End Sub

Property Get className()
    className = var_className
End Property


Property Set setWorkbook(wb As Workbook)
    Set var_Workbook = wb
End Property

Function getWorkbook() As Workbook
    If var_Workbook Is Nothing Then
        Err.Raise vbObjectError + 1001, ThisWorkbook.VBProject.Name & "." & Me.className & "." & "getWorkbook", "WorkbookNotSetError : Workbook object is required, Please set Workbook object using property 'setWorkbook'. Eg. : Me.setWorkbook=wb"
    End If
    Set getWorkbook = var_Workbook
End Function

Function getQueries() As Queries
    'Returns all Queries inside Workbook
    'If no query in workbook then returns --> Nothing
    
    Dim wb As Workbook
    Dim queriesResult As Queries
    Set wb = Me.getWorkbook
    Set queriesResult = ExcelUtil.getQueries(wb)
    getQueries = queriesResult
    
End Function

Function getQueryByName(queryName As String) As WorkbookQuery
    'Return the specified query from the workbook
    'If no query exists then raise error
    
    Dim tWb As Workbook
    Set tWb = Me.getWorkbook()
    Set getQueryByName = ExcelUtil.getQueryByName(tWb, queryName)
    
End Function

Function worksheetExists(sheetName As String, ifNotExistsRaiseError As Boolean) As Boolean
    'It checks whether given sheetName exists inside given workbook or not.
    'workbook must be open
    'It returns Boolean if : IfNotExistsRaiseError is False
    'it raise SheetNotFoundError if Sheet is not found and if IfNotExistsRaiseError is True
    Dim wb As Workbook
    Dim sheetExists As Boolean
    Set wb = Me.getWorkbook
    sheetExists = ExcelUtil.worksheetExists(wb, sheetName, ifNotExistsRaiseError)
    worksheetExists = sheetExists
    
End Function

Function getWorksheet(sheetName As String, ifNotExistsRaiseError As Boolean) As Worksheet
    Dim wb As Workbook
    Dim sheetReq As Worksheet
    Set wb = Me.getWorkbook
    Set sheetReq = ExcelUtil.getWorksheet(wb, sheetName, ifNotExistsRaiseError)
    getWorksheet = sheetReq
    
End Function


Sub deleteWorksheet(sheetNameToDelete As String)
    'It will delete the specified worksheets from the workbook.
    'It will do nothing if worksheet not exists
    'It not returns anything.
    Dim wb As Workbook
    Set wb = Me.getWorkbook()
    ExcelUtil.deleteWorksheet wb, sheetNameToDelete
End Sub

Function createWorksheet(sheetName As String, ifAlreadyExistsRaiseError As Boolean) As Worksheet
    'It will return the new created worksheet
    'It will create the new worksheet in the last and return that worksheet object
    'If worksheet already exists --> Then raise Error / return Nothing i.e., if worksheet exists it not do
    'anything with the existing worksheet.
    Dim wb As Workbook
    Dim reqSheet As Worksheet
    Set wb = Me.getWorkbook
    Set reqSheet = ExcelUtil.createWorksheet(wb, sheetName, ifAlreadyExistsRaiseError)
    createWorksheet = reqSheet
End Function

