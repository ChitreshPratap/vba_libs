VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tsClsSAPUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim sapGUIAuto As Variant
Dim sapApplication As Variant
Dim sapConnection As Variant
Dim sapSession As Variant

Private Sub Class_Initialize()
    setVariableValues
End Sub

Sub setVariableValues()
    
    Set sapGUIAuto = GetObject("SAPGUI")
    Set sapApplication = sapGUIAuto.GetScriptingEngine
    Set sapConnection = sapApplication.Children(0)
    Set sapSession = sapConnection.Children(0)
    
End Sub

Function getSapSession() As Variant
    If sapSession Is Nothing Then
        setVariableValues
    End If
    Set getSapSession = sapSession
End Function

Private Sub Class_Terminate()
    
    Set sapGUIAuto = Nothing
    Set sapApplication = Nothing
    Set sapConnection = Nothing
    Set sapSession = Nothing

End Sub
