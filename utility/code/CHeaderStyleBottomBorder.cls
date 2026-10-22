VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHeaderStyleBottomBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IHeaderStyle

'==============================================================================
' Class   : CHeaderStyleBottomBorder
' Purpose : Header style for CCollapsibleFrame - keeps the normal header and
'           draws a horizontal border along the bottom of the header bar.
'
' Usage   : Dim bottomStyle As CHeaderStyleBottomBorder
'           Set bottomStyle = New CHeaderStyleBottomBorder
'           bottomStyle.BorderColor = RGB(70, 130, 180)
'           Set panel.HeaderStyle = bottomStyle
'
' Notes   : - BorderWidth = 0 (default) means the full header width; larger
'             values than the header are clamped to it.
'           - The border is drawn INSIDE the bottom of the bar and is never
'             taller than the space below the heading text, so it cannot
'             cover the heading or the icon.
'           - Property changes redraw immediately once rendered.
'==============================================================================

Private Const DEFAULT_COLOR      As Long = &H555555     ' RGB(85, 85, 85)
Private Const DEFAULT_HEIGHT     As Single = 2
Private Const MAX_HEIGHT         As Single = 12
Private Const MIN_HEIGHT         As Single = 0.5
Private Const LINE_SUFFIX        As String = "BottomBorder"
Private Const ERR_BAD_VALUE      As Long = vbObjectError + 601

'--- Settings -----------------------------------------------------------------
Private mColor          As Long
Private mHeight         As Single
Private mWidth          As Single       ' 0 = full header width
Private mAlign          As String       ' "Left" / "Center" / "Right"

'--- Last render context (supplied by the container) --------------------------
Private mHasContext     As Boolean
Private mHost           As Object
Private mPrefix         As String
Private mTitle          As MSForms.Label
Private mBarLeft        As Single
Private mBarTop         As Single
Private mBarWidth       As Single
Private mBarHeight      As Single

'--- Owned control ------------------------------------------------------------
Private mLine           As MSForms.Label
Private mLineName       As String

'==============================================================================
' LIFECYCLE
'==============================================================================

Private Sub Class_Initialize()
    mColor = DEFAULT_COLOR
    mHeight = DEFAULT_HEIGHT
    mWidth = 0
    mAlign = "Center"
End Sub

Private Sub Class_Terminate()
    ' References only; the form may already be unloading.
    Set mLine = Nothing
    Set mTitle = Nothing
    Set mHost = Nothing
End Sub

'==============================================================================
' PUBLIC PROPERTIES
'==============================================================================

Public Property Get BorderColor() As Long
    BorderColor = mColor
End Property

Public Property Let BorderColor(ByVal Value As Long)
    mColor = Value
    Draw
End Property

Public Property Get BorderHeight() As Single
    BorderHeight = mHeight
End Property

Public Property Let BorderHeight(ByVal Value As Single)
    If Value < MIN_HEIGHT Or Value > MAX_HEIGHT Then
        Err.Raise ERR_BAD_VALUE, "CHeaderStyleBottomBorder.BorderHeight", _
                  "BorderHeight must be between " & MIN_HEIGHT & " and " & MAX_HEIGHT & _
                  " points (got " & Value & ")."
    End If
    mHeight = Value
    Draw
End Property

' 0 = full header width.
Public Property Get BorderWidth() As Single
    BorderWidth = mWidth
End Property

Public Property Let BorderWidth(ByVal Value As Single)
    If Value < 0 Then
        Err.Raise ERR_BAD_VALUE, "CHeaderStyleBottomBorder.BorderWidth", _
                  "BorderWidth cannot be negative (got " & Value & "). Use 0 for full width."
    End If
    mWidth = Value
    Draw
End Property

Public Property Get BorderAlignment() As String
    BorderAlignment = mAlign
End Property

Public Property Let BorderAlignment(ByVal Value As String)
    Select Case LCase$(Trim$(Value))
        Case "left":   mAlign = "Left"
        Case "center": mAlign = "Center"
        Case "right":  mAlign = "Right"
        Case Else
            Err.Raise ERR_BAD_VALUE, "CHeaderStyleBottomBorder.BorderAlignment", _
                      "BorderAlignment must be ""Left"", ""Center"" or ""Right"" (got """ & _
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
    If mHasContext And StrComp(NamePrefix, mPrefix, vbTextCompare) <> 0 Then RemoveLine

    Set mHost = Host
    mPrefix = NamePrefix
    Set mTitle = Title
    mBarLeft = BarLeft
    mBarTop = BarTop
    mBarWidth = BarWidth
    mBarHeight = BarHeight
    mHasContext = True

    Draw
End Sub

Private Sub IHeaderStyle_Clear()
    RemoveLine
    Set mTitle = Nothing
    Set mHost = Nothing
    mHasContext = False
End Sub

'==============================================================================
' DRAWING
'==============================================================================

Private Sub Draw()
    Dim w As Single
    Dim h As Single
    Dim x As Single
    Dim spaceBelowText As Single

    If Not mHasContext Then Exit Sub
    If mBarWidth <= 0 Or mBarHeight <= 0 Then Exit Sub

    EnsureLine

    ' Height: never reach up into the heading text
    h = mHeight
    spaceBelowText = (mBarTop + mBarHeight) - (mTitle.Top + mTitle.Height)
    If spaceBelowText < MIN_HEIGHT Then spaceBelowText = MIN_HEIGHT
    If h > spaceBelowText Then h = spaceBelowText

    ' Width: 0 or too wide = full width
    w = mWidth
    If w <= 0 Or w > mBarWidth Then w = mBarWidth

    Select Case mAlign
        Case "Left":  x = mBarLeft
        Case "Right": x = mBarLeft + mBarWidth - w
        Case Else:    x = mBarLeft + (mBarWidth - w) / 2
    End Select

    With mLine
        .BackColor = mColor
        .Left = x
        .Top = mBarTop + mBarHeight - h
        .Width = w
        .Height = h
    End With
    mHost.Controls(mLineName).ZOrder fmZOrderFront   ' above the bar
End Sub

Private Sub EnsureLine()
    Dim wanted As String

    wanted = mPrefix & LINE_SUFFIX
    If Not mLine Is Nothing Then
        If StrComp(mLineName, wanted, vbTextCompare) = 0 Then Exit Sub
        RemoveLine
    End If

    mLineName = wanted
    If HostHasControl(mLineName) Then
        Set mLine = mHost.Controls(mLineName)        ' reuse, never duplicate
    Else
        Set mLine = mHost.Controls.Add("Forms.Label.1", mLineName, True)
    End If

    With mLine
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Private Sub RemoveLine()
    If mLine Is Nothing Then Exit Sub
    Set mLine = Nothing
    If mHost Is Nothing Then Exit Sub
    If HostHasControl(mLineName) Then mHost.Controls.Remove mLineName
    mLineName = vbNullString
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
