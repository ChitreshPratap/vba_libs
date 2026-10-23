VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHeaderStyleDivider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IHeaderStyle

'==============================================================================
' Class   : CHeaderStyleDivider
' Purpose : Header style for CCollapsibleFrame - draws a horizontal line on
'           each side of the heading:   -------- Heading --------
'
' Usage   : Dim dividerStyle As CHeaderStyleDivider
'           Set dividerStyle = New CHeaderStyleDivider
'           dividerStyle.LineColor = RGB(120, 120, 120)
'           Set panel.HeaderStyle = dividerStyle
'
' How it fits the text:
'   - The container passes the heading area [TextLeft, TextRight], which
'     already excludes the icon (left or right side).
'   - The heading label is auto-sized to its real rendered width (font,
'     size, bold, italic, current expanded/collapsed text), then placed
'     by HeadingAlignment inside that area.
'   - Each line runs from the area edge to LineGap points from the text.
'     LeftLineWidth / RightLineWidth > 0 cap a line to that length (kept
'     next to the text). For "Left"/"Right" alignment they set the short
'     lead line (default 16 pt).
'   - A line shorter than 3 pt is hidden; text that is wider than the area
'     is clipped and both lines are hidden. Lines never cross the text.
'==============================================================================

Private Const DEFAULT_COLOR      As Long = &H787878     ' RGB(120, 120, 120)
Private Const DEFAULT_HEIGHT     As Single = 1
Private Const DEFAULT_GAP        As Single = 8
Private Const DEFAULT_LEAD       As Single = 16         ' lead line for Left/Right alignment
Private Const MIN_VISIBLE_LINE   As Single = 3
Private Const MAX_HEIGHT         As Single = 6
Private Const MIN_HEIGHT         As Single = 0.5
Private Const MAX_GAP            As Single = 50
Private Const TEXT_SAFETY        As Single = 1          ' extra width so text never clips
Private Const LEFT_SUFFIX        As String = "DividerLeft"
Private Const RIGHT_SUFFIX       As String = "DividerRight"
Private Const ERR_BAD_VALUE      As Long = vbObjectError + 611

'--- Settings -----------------------------------------------------------------
Private mColor          As Long
Private mLineH          As Single
Private mLeftW          As Single       ' 0 = auto
Private mRightW         As Single       ' 0 = auto
Private mGap            As Single
Private mAlign          As String       ' "Left" / "Center" / "Right"

'--- Last render context ------------------------------------------------------
Private mHasContext     As Boolean
Private mHost           As Object
Private mPrefix         As String
Private mTitle          As MSForms.Label
Private mBarTop         As Single
Private mBarHeight      As Single
Private mTextLeft       As Single
Private mTextRight      As Single
Private mHidden         As Boolean      ' header hidden by the panel

'--- Owned controls -----------------------------------------------------------
Private mLeftLine       As MSForms.Label
Private mRightLine      As MSForms.Label
Private mLeftName       As String
Private mRightName      As String

'==============================================================================
' LIFECYCLE
'==============================================================================

Private Sub Class_Initialize()
    mColor = DEFAULT_COLOR
    mLineH = DEFAULT_HEIGHT
    mLeftW = 0
    mRightW = 0
    mGap = DEFAULT_GAP
    mAlign = "Center"
End Sub

Private Sub Class_Terminate()
    ' References only; the form may already be unloading.
    Set mLeftLine = Nothing
    Set mRightLine = Nothing
    Set mTitle = Nothing
    Set mHost = Nothing
End Sub

'==============================================================================
' PUBLIC PROPERTIES
'==============================================================================

Public Property Get LineColor() As Long
    LineColor = mColor
End Property

Public Property Let LineColor(ByVal Value As Long)
    mColor = Value
    Draw
End Property

Public Property Get LineHeight() As Single
    LineHeight = mLineH
End Property

Public Property Let LineHeight(ByVal Value As Single)
    If Value < MIN_HEIGHT Or Value > MAX_HEIGHT Then
        Err.Raise ERR_BAD_VALUE, "CHeaderStyleDivider.LineHeight", _
                  "LineHeight must be between " & MIN_HEIGHT & " and " & MAX_HEIGHT & _
                  " points (got " & Value & ")."
    End If
    mLineH = Value
    Draw
End Property

' 0 = fill the available space.
Public Property Get LeftLineWidth() As Single
    LeftLineWidth = mLeftW
End Property

Public Property Let LeftLineWidth(ByVal Value As Single)
    If Value < 0 Then
        Err.Raise ERR_BAD_VALUE, "CHeaderStyleDivider.LeftLineWidth", _
                  "LeftLineWidth cannot be negative (got " & Value & "). Use 0 for automatic."
    End If
    mLeftW = Value
    Draw
End Property

' 0 = fill the available space.
Public Property Get RightLineWidth() As Single
    RightLineWidth = mRightW
End Property

Public Property Let RightLineWidth(ByVal Value As Single)
    If Value < 0 Then
        Err.Raise ERR_BAD_VALUE, "CHeaderStyleDivider.RightLineWidth", _
                  "RightLineWidth cannot be negative (got " & Value & "). Use 0 for automatic."
    End If
    mRightW = Value
    Draw
End Property

Public Property Get LineGap() As Single
    LineGap = mGap
End Property

Public Property Let LineGap(ByVal Value As Single)
    If Value < 0 Or Value > MAX_GAP Then
        Err.Raise ERR_BAD_VALUE, "CHeaderStyleDivider.LineGap", _
                  "LineGap must be between 0 and " & MAX_GAP & " points (got " & Value & ")."
    End If
    mGap = Value
    Draw
End Property

Public Property Get HeadingAlignment() As String
    HeadingAlignment = mAlign
End Property

Public Property Let HeadingAlignment(ByVal Value As String)
    Select Case LCase$(Trim$(Value))
        Case "left":   mAlign = "Left"
        Case "center": mAlign = "Center"
        Case "right":  mAlign = "Right"
        Case Else
            Err.Raise ERR_BAD_VALUE, "CHeaderStyleDivider.HeadingAlignment", _
                      "HeadingAlignment must be ""Left"", ""Center"" or ""Right"" (got """ & _
                      Value & """). Current value kept: """ & mAlign & """."
    End Select
    Draw
End Property

'==============================================================================
' IHeaderStyle
'==============================================================================

Private Sub IHeaderStyle_Render(ByVal Host As Object, _
                                ByVal NamePrefix As String, _
                                ByVal Title As MSForms.Label, _
                                ByVal BarLeft As Single, _
                                ByVal BarTop As Single, _
                                ByVal BarWidth As Single, _
                                ByVal BarHeight As Single, _
                                ByVal TextLeft As Single, _
                                ByVal TextRight As Single)

    ' Moved to a different panel? Drop the old decoration first.
    If mHasContext And StrComp(NamePrefix, mPrefix, vbTextCompare) <> 0 Then RemoveLines

    Set mHost = Host
    mPrefix = NamePrefix
    Set mTitle = Title
    mBarTop = BarTop
    mBarHeight = BarHeight
    mTextLeft = TextLeft
    mTextRight = TextRight
    mHasContext = True
    mHidden = False                 ' Render is only called for a shown header

    Draw
End Sub

Private Sub IHeaderStyle_Clear()
    RemoveLines
    Set mTitle = Nothing
    Set mHost = Nothing
    mHasContext = False
    mHidden = False
End Sub

Private Sub IHeaderStyle_SetVisible(ByVal Visible As Boolean)
    mHidden = Not Visible
    If mHost Is Nothing Then Exit Sub
    If Visible Then
        Draw                        ' re-decides each line (may stay hidden if too short)
    Else
        If HostHasControl(mLeftName) Then mHost.Controls(mLeftName).Visible = False
        If HostHasControl(mRightName) Then mHost.Controls(mRightName).Visible = False
    End If
End Sub

'==============================================================================
' DRAWING
'==============================================================================

Private Sub Draw()
    Dim avail As Single
    Dim textW As Single
    Dim textX As Single
    Dim lead As Single
    Dim leftEnd As Single, leftW As Single
    Dim rightStart As Single, rightW As Single
    Dim lineTop As Single

    If Not mHasContext Then Exit Sub
    If mHidden Then Exit Sub                 ' redrawn when shown again
    avail = mTextRight - mTextLeft
    If avail <= 0 Then Exit Sub

    EnsureLines

    ' 1. Real rendered width of the current heading
    mTitle.WordWrap = False
    mTitle.AutoSize = True
    mTitle.AutoSize = False
    textW = mTitle.Width + TEXT_SAFETY
    If textW > avail Then textW = avail

    ' 2. Place the heading inside [mTextLeft, mTextRight]
    Select Case mAlign
        Case "Left"
            lead = mLeftW
            If lead <= 0 Then lead = DEFAULT_LEAD
            textX = mTextLeft + lead + mGap
        Case "Right"
            lead = mRightW
            If lead <= 0 Then lead = DEFAULT_LEAD
            textX = mTextRight - lead - mGap - textW
        Case Else
            textX = mTextLeft + (avail - textW) / 2
    End Select
    If textX + textW > mTextRight Then textX = mTextRight - textW
    If textX < mTextLeft Then textX = mTextLeft

    mTitle.Left = textX
    mTitle.Width = textW

    ' 3. Lines: from the area edge to LineGap from the text (optionally capped)
    leftEnd = textX - mGap
    leftW = leftEnd - mTextLeft
    If mLeftW > 0 And leftW > mLeftW Then leftW = mLeftW

    rightStart = textX + textW + mGap
    rightW = mTextRight - rightStart
    If mRightW > 0 And rightW > mRightW Then rightW = mRightW

    lineTop = mBarTop + (mBarHeight - mLineH) / 2

    PlaceLine mLeftLine, mLeftName, leftEnd - leftW, lineTop, leftW
    PlaceLine mRightLine, mRightName, rightStart, lineTop, rightW
End Sub

Private Sub PlaceLine(ByVal lineLbl As MSForms.Label, ByVal lineName As String, _
                      ByVal x As Single, ByVal y As Single, ByVal w As Single)
    Dim ext As MSForms.Control

    Set ext = mHost.Controls(lineName)
    If w < MIN_VISIBLE_LINE Then
        If ext.Visible Then ext.Visible = False      ' not enough room: hide
        Exit Sub
    End If

    With lineLbl
        .BackColor = mColor
        .Left = x
        .Top = y
        .Width = w
        .Height = mLineH
    End With
    If Not ext.Visible Then ext.Visible = True
    ext.ZOrder fmZOrderFront                          ' above the bar
End Sub

Private Sub EnsureLines()
    Dim wantL As String
    Dim wantR As String

    wantL = mPrefix & LEFT_SUFFIX
    wantR = mPrefix & RIGHT_SUFFIX

    If Not mLeftLine Is Nothing Then
        If StrComp(mLeftName, wantL, vbTextCompare) = 0 Then Exit Sub
        RemoveLines
    End If

    mLeftName = wantL
    mRightName = wantR
    Set mLeftLine = GetOrAddLine(mLeftName)
    Set mRightLine = GetOrAddLine(mRightName)
End Sub

Private Function GetOrAddLine(ByVal ctlName As String) As MSForms.Label
    Dim lbl As MSForms.Label

    If HostHasControl(ctlName) Then
        Set lbl = mHost.Controls(ctlName)            ' reuse, never duplicate
    Else
        Set lbl = mHost.Controls.Add("Forms.Label.1", ctlName, True)
    End If

    With lbl
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With
    Set GetOrAddLine = lbl
End Function

Private Sub RemoveLines()
    Set mLeftLine = Nothing
    Set mRightLine = Nothing
    If Not mHost Is Nothing Then
        If HostHasControl(mLeftName) Then mHost.Controls.Remove mLeftName
        If HostHasControl(mRightName) Then mHost.Controls.Remove mRightName
    End If
    mLeftName = vbNullString
    mRightName = vbNullString
End Sub

Private Function HostHasControl(ByVal ctlName As String) As Boolean
    Dim ctl As MSForms.Control
    If mHost Is Nothing Or Len(ctlName) = 0 Then Exit Function
    For Each ctl In mHost.Controls
        If StrComp(ctl.Name, ctlName, vbTextCompare) = 0 Then
            HostHasControl = True
            Exit Function
        End If
    Next ctl
End Function
