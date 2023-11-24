VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim var_conn As ADODB.Connection
Dim var_excelFileName As String

Public Sub Class_Initialize()
    Set var_conn = New ADODB.Connection
End Sub

Public Property Let excelDBFileName(value As String)
    var_excelFileName = value
End Property

Public Property Get excelDBFileName() As String
    excelDBFileName = var_excelFileName
End Property

Public Property Get getOpenConnection() As ADODB.Connection
    
    var_conn.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & wbCopy_path & ";" & _
        "Extended Properties='Excel 12.0 Xml;HDR=Yes';"
    var_conn.Open
    Set getOpenConnection = var_conn
End Property

Public Sub closeConnection()
    var_conn.Close
    Set var_conn = Nothing
End Sub
    
