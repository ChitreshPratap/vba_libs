VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBarLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim var_lblProgressBar As MSForms.Label
Dim var_maxProgressBarWidth As Double
Dim var_maxValue As Long

Private Sub Class_Initialize()
    var_maxValue = 0
End Sub

Public Property Let setVisible(isVisible As Boolean)
    var_lblProgressBar.visible = isVisible
End Property

Public Property Let maxValue(maxValue As Long)
    var_maxValue = maxValue
End Property

Public Property Get maxValue() As Long
    maxValue = var_maxValue
End Property

Public Property Let maxProgressBarWidth(maxBarWidth As Double)
    var_maxProgressBarWidth = maxBarWidth
End Property

Public Property Get maxProgressBarWidth() As Double
    maxProgressBarWidth = var_maxProgressBarWidth
End Property

Public Property Set labelProgressBar(lblProgressBar As MSForms.Label)
    Set var_lblProgressBar = lblProgressBar
End Property

Public Property Get labelProgressBar() As MSForms.Label
    Set labelProgressBar = var_lblProgressBar
End Property

Public Sub resetProgressbar()
    Me.maxValue = 0
End Sub

Public Sub updateProgress(currentValueToUpdate As Long)
    
    Dim maxValue As Long
    Dim currentValue As Long
    Dim progressBar As MSForms.Label
    Dim maxBarWidth As Double
    
    currentValue = currentValueToUpdate
    maxValue = Me.maxValue
    maxBarWidth = Me.maxProgressBarWidth
    
    Set progressBar = Me.labelProgressBar
    
    If maxValue = 0 Then
        Exit Sub
    End If
    
    If currentValueToUpdate > maxValue Then
        currentValue = maxValue
    End If
    
    If currentValueToUpdate < 0 Then
        currentValue = 0
    End If
    
    progressBar.Width = (maxBarWidth * currentValue) / maxValue
    If currentValue = maxValue Then
        progressBar.Caption = "Completed"
    Else
        progressBar.Caption = currentValue & "/" & maxValue
    End If
End Sub




