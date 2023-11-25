VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    Dim dbFileName As String
    dbFileName = Me.excelDBFileName
    If dbFileName = "" Then
        Err.Raise vbObjectError + 12, "ExcelConnection.getOpenConnection", "ExcelDBFile is blank, Please set excelDBFile"
    End If
    var_conn.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & dbFileName & ";" & _
        "Extended Properties='Excel 12.0 Xml;HDR=Yes';"
        On Error GoTo ConnectionError
    var_conn.Open
    Set getOpenConnection = var_conn
    Exit Property
ConnectionError:
    Err.Raise Err.Number + 6, "ExcelConnection.getOpenConnection", "ConnectionError : The connection failed to open. " & vbNewLine & Err.Description
End Property

Public Sub CloseConnection()
    On Error Resume Next
    var_conn.Close
    Set var_conn = Nothing
End Sub

