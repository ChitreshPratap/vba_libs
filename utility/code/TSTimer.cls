VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TSTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim var_className As String
Dim var_startTime As Date
Dim var_endTime As Date

Private Sub Class_Initialize()
    var_className = "TSTimer"
End Sub

Sub startTimer()
    var_startTime = Time
End Sub

Sub stopTimer()
    var_endTime = Time
End Sub

Public Property Get startTime() As Date
    If var_startTime = 0 Then
        Err.Raise vbObjectError + 1001, var_className & "_startTime : ", var_className & "_endTime" & vbNewLine & "ValueError : timer is not started, please call 'startTimer' to start timer."
    End If
    startTime = var_startTime
    
End Property

Public Property Get endTime() As Date
    If var_endTime = 0 Then
        Err.Raise vbObjectError + 1001, var_className & "_endTime : ", var_className & "_endTime" & vbNewLine & "ValueError : timer is not stopped, please call 'stopTimer' to stop timer."
    End If
    endTime = var_endTime
End Property

Public Function getTimeElapsed() As Date
    Dim st As Date
    Dim et As Date
    st = Me.startTime
    et = Me.endTime
    Dim tResult As Date
    tResult = et - st
    getTimeElapsed = tResult
End Function

Public Function getTimeElapsedFormatted(Optional tFormat As String = "HH:MM:SS") As String
    
    Dim tElapsed As Date
    tElapsed = getTimeElapsed()
    getTimeElapsedFormatted = Format(tElapsed, tFormat)
    
End Function

