VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxStyle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This style place a placeholder on a text box which will disappear when the textbox has color
' Place a underline at the bottom of the textbox
' No other configuration is required

Dim var_mukhTextBox As MSForms.TextBox
Dim var_lblCaption As MSForms.Label
Dim var_lblLine As MSForms.Label
Public WithEvents txtDynamic As MSForms.TextBox
Attribute txtDynamic.VB_VarHelpID = -1

Sub init(mukhTextBox As Control, Optional placeholder As String = "Enter Value")
    
    Set var_mukhTextBox = mukhTextBox
    Set var_lblCaption = mukhTextBox.Parent.Controls.Add("Forms.Label.1", "lblKey" & mukhTextBox.Name)
    Set var_lblLine = mukhTextBox.Parent.Controls.Add("Forms.Label.1", "lblLine" & mukhTextBox.Name)
    var_lblCaption.Caption = placeholder
    var_lblCaption.Font = var_mukhTextBox.Font
    var_lblCaption.Font.Size = var_mukhTextBox.Font.Size
    var_lblCaption.Font.Name = var_mukhTextBox.Font.Name
    Set txtDynamic = mukhTextBox

'    Set var_lblCaption = lblCaption
'    Set var_lblLine = lblLine
    
    var_mukhTextBox.BackStyle = fmBackStyleTransparent
    var_mukhTextBox.SpecialEffect = fmSpecialEffectFlat
    var_lblCaption.Top = var_mukhTextBox.Top + 1.5
    var_lblCaption.Height = var_mukhTextBox.Height - 1.5
    var_lblCaption.Width = var_mukhTextBox.Width - 7.5
    var_lblCaption.Left = var_mukhTextBox.Left + 7.5
    var_lblLine.Top = var_mukhTextBox.Top + var_mukhTextBox.Height - 1
    var_lblLine.Left = var_lblCaption.Left - 2
    var_lblLine.Width = var_lblCaption.Width + 2
    var_lblLine.Height = 1.3
    var_lblLine.BackColor = VarnahUtil.getFadeColor(var_mukhTextBox.ForeColor, -0.3)
    var_lblCaption.ForeColor = VarnahUtil.getFadeColor(var_mukhTextBox.ForeColor, 0.55)
    var_lblCaption.ZOrder (1)
    Me.styleOnTextChanged
    
End Sub

Sub styleOnTextChanged()
    If var_mukhTextBox.Text = "" Then
        var_lblCaption.Visible = True
    Else
        var_lblCaption.Visible = False
    End If
End Sub

Private Sub txtDynamic_Change()
    Me.styleOnTextChanged
End Sub

Public Sub txtDynamic_Enter()
    MsgBox "New txt dynamic"
End Sub

'Private Sub txtDynamic_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'    MsgBox "Insider Focus removes"
'End Sub


