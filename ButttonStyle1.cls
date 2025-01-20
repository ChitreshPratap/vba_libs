VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButttonStyle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim var_mukhLabel As MSForms.Label
Dim var_originalFont As Long
Public WithEvents dynaLabelParent As MSForms.UserForm
Attribute dynaLabelParent.VB_VarHelpID = -1
Public WithEvents dynaLabel As MSForms.Label
Attribute dynaLabel.VB_VarHelpID = -1

Sub init(mukhLabel As Control)
    Set var_mukhLabel = mukhLabel
    Set dynaLabel = var_mukhLabel
    Set dynaLabelParent = var_mukhLabel.Parent
    var_originalFont = mukhLabel.Font.Size
End Sub

Private Sub styleHoverOn()
    With var_mukhLabel
        .Font.Size = var_originalFont + 1
    End With
End Sub

Private Sub styleHoverOff()
    With var_mukhLabel
        .Font.Size = var_originalFont
    End With
End Sub

Private Sub dynaLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    styleHoverOn
End Sub

Private Sub dynaLabelParent_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    styleHoverOff
End Sub

