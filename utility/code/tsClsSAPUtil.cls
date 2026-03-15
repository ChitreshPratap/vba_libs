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

Sub destruct()
    Set sapGUIAuto = Nothing
    Set sapApplication = Nothing
    Set sapConnection = Nothing
    Set sapSession = Nothing

End Sub

Private Sub Class_Terminate()
    
    Set sapGUIAuto = Nothing
    Set sapApplication = Nothing
    Set sapConnection = Nothing
    Set sapSession = Nothing

End Sub

Sub saveWorkbook2(outputFileName As String)
            
    'Dim outputFileName as String
    Dim requiredName As String
    Dim requiredWB As Workbook
    Dim wbName As String
    requiredName = "Worksheet in ALVXXL*"
    
    For Each tWb In Workbooks
        If tWb.Name Like requiredName Then
            Set requiredWB = Workbooks.Add
            tWb.ActiveSheet.UsedRange.Copy
            requiredWB.Worksheets(1).Range("A1").PasteSpecial xlPasteValues
            requiredWB.Worksheets(1).Columns.AutoFit
            requiredWB.SaveAs outputFileName
            requiredWB.Close True
        End If
    Next tWb
        
End Sub


Sub saveWorkbook1()
    Dim xlWB As Workbook
    Dim xlApp As Excel.Application
    Set xlApp = GetObject(, "Excel.application")
    Set xlWB = xlApp.Workbooks.Item("Worksheet in Basis (1)")
    xlWB.SaveAs "Save.xlsx"
    
End Sub

Sub exampleFillingFBL3NGLRangeSingleGL(session As Variant)
    
    Dim glRangeList As Range
    Dim glSingleList As Range
    Dim glSingleList As Ranage
    Dim companyCodeList As Range
    Dim verticalPosition As Long
    Dim rowNumber As Long
    
    If Not glRangeList Is Nothing Then
        verticalPosition = 1
        rowNumber = 1
        
        For Each tempRangeGL In glRangeList
            If rowNumber = 1 Then
                session.FindById("wnd[1]/usr/tabsTAB/ctxtRSCEL-ILOW_I[1,0]").Text = CStr(tempRangeGL.value)
                session.FindById("wnd[1]/usr/tabsTAB/ctxtRSCEL-ILOW_I[2,0]").Text = CStr(tempRangeGL.Offset(0, 1).value)
            Else
                session.FindById("wnd[1]/usr/tabsTAB/ctxtRSCEL-ILOW_I[1," & CStr(1) & "]").Text = CStr(tempRangeGL.value)
                session.FindById("wnd[1]/usr/tabsTAB/ctxtRSCEL-ILOW_I[2," & CStr(1) & "]").Text = CStr(tempRangeGL.Offset(0, 1).value)
                session.FindById("wnd[1]/usr/tabsTAB").verticalScrollBar.Position = verticalPosition
                verticalPosition = verticalPosition + 1
            End If
            rowNumber = rowNumber + 1
        Next tempRangeGL
                        
    End If
    
    If Not glSingleList Is Nothing Then
        verticalPosition = 1
        rowNumber = 1
        
        For Each tempSingleGL In glSingleList
            
            If rowNumber = 1 Then
                session.FindById("wnd[1]/usr/tabsTAB/ctxtRSCEL-ILOW_I[1,0]").Text = CStr(tempSingleGL.value)
            Else
                session.FindById("wnd[1]/usr/tabsTAB/ctxtRSCEL-ILOW_I[1," & CStr(1) & "]").Text = CStr(tempSingleGL.value)
                session.FindById("wnd[1]/usr/tabsTAB").verticalScrollBar.Position = verticalPosition
                verticalPosition = verticalPosition + 1
            End If
            rowNumber = rowNumber + 1
        Next tempSingleGL
        
    End If
    
End Sub
