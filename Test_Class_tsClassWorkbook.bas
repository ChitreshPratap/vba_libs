Attribute VB_Name = "Test_Class_tsClassWorkbook"


Sub testQueries()
    Dim wb As Workbook
    Set wb = Workbooks.Open("C:\Users\pc\Downloads\Example1.xlsx")
    Dim wbQueries As Queries
    
    Dim classWb As tsClassWorkbook
    Set classWb = New tsClassWorkbook
    Set classWb.setWorkbook = wb
    
    On Error GoTo handleError
    classWb.worksheetExists "Yes", False
    classWb.deleteWorksheet "No"
    classWb.createWorksheet "Dhost", False
    classWb.deleteWorksheet "Dhost"
    Set wbQueries = classWb.getQueries()
    
    Dim wbQuery As WorkbookQuery
    Set wbQuery = classWb.getQueryByName("table2")
    
    For Each k In wbQueries
        Debug.Print k.Name
    Next k
    Debug.Print wbQueries.Count
    Exit Sub
handleError:
    MsgBox Err.Source & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub
