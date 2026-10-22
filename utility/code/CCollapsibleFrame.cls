VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCollapsibleFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Class   : CCollapsibleFrame
' Purpose : Turns an EXISTING MSForms.Frame into a collapsible panel with a
'           dark-grey header (arrow + heading) and a light-grey content area.
'
' Rules   : - Frame Left/Top/Width/Height are never changed (verified after
'             every operation).
'           - Child controls are hidden/restored only via .Visible; their
'             positions and values are never touched.
'           - No API calls, no OnTime, no extra classes or modules.
'
' Usage   : Keep the instance in a FORM-LEVEL variable (events stop if the
'           object goes out of scope).
'
'           Private mPanel As CCollapsibleFrame
'           Set mPanel = New CCollapsibleFrame
'           mPanel.Initialize Me.Frame1, "Customer Information"
'==============================================================================

'--- Appearance ---------------------------------------------------------------
Private Const DEFAULT_HEADING       As String = "Details"
Private Const DEFAULT_HEADER_HEIGHT As Single = 24
Private Const MIN_HEADER_HEIGHT     As Single = 14

Private Const CLR_HEADER            As Long = &H555555      ' RGB(85, 85, 85)
Private Const CLR_CONTENT           As Long = &HF0F0F0      ' RGB(240, 240, 240)
Private Const CLR_TEXT              As Long = &HFFFFFF      ' White
Private Const CLR_FALLBACK_FORM     As Long = &H8000000F    ' System button face

Private Const ARROW_FONT            As String = "Marlett"   ' Standard Windows glyph font
Private Const GLYPH_EXPANDED        As String = "6"         ' Marlett: down triangle  (down)
Private Const GLYPH_COLLAPSED       As String = "4"         ' Marlett: right triangle (right)
Private Const TITLE_FONT            As String = "Segoe UI"
Private Const TITLE_FONT_SIZE       As Single = 10

Private Const ARROW_LEFT            As Single = 6
Private Const ARROW_WIDTH           As Single = 14
Private Const TITLE_GAP             As Single = 4
Private Const RIGHT_PAD             As Single = 8

'--- State --------------------------------------------------------------------
Private mFrame              As MSForms.Frame
Private WithEvents mBar     As MSForms.Label    ' header background (clickable)
Attribute mBar.VB_VarHelpID = -1
Private WithEvents mArrow   As MSForms.Label    ' arrow glyph (clickable)
Attribute mArrow.VB_VarHelpID = -1
Private WithEvents mTitle   As MSForms.Label    ' heading text (clickable)
Attribute mTitle.VB_VarHelpID = -1

Private mInitialized        As Boolean
Private mExpanded           As Boolean
Private mHeaderHeight       As Single
Private mPrefix             As String
Private mVisibility         As Collection       ' child Visible states, keyed by name
Private mCollapsedBackColor As Long

' Original geometry (never allowed to change)
Private mLeft               As Single
Private mTop                As Single
Private mWidth              As Single
Private mHeight             As Single

' Original appearance (restored by Detach)
Private mOrigCaption        As String
Private mOrigBorderStyle    As Long
Private mOrigSpecialEffect  As Long
Private mOrigBackColor      As Long

'==============================================================================
' PUBLIC API
'==============================================================================

Public Sub Initialize(ByVal FrameObject As MSForms.Frame, _
                      Optional ByVal HeadingText As String = vbNullString, _
                      Optional ByVal HeaderHeight As Single = DEFAULT_HEADER_HEIGHT)

    If FrameObject Is Nothing Then
        Err.Raise vbObjectError + 513, "CCollapsibleFrame.Initialize", _
                  "FrameObject is Nothing. Pass an existing Frame, e.g. Me.Frame1."
    End If

    ' Re-initialisation: undo previous styling first (no duplicate headers)
    If mInitialized Then Detach

    Set mFrame = FrameObject
    mPrefix = "ccf_" & mFrame.Name & "_"

    ' 1. Remember geometry and appearance
    With mFrame
        mLeft = .Left
        mTop = .Top
        mWidth = .Width
        mHeight = .Height
        mOrigCaption = .Caption
        mOrigBorderStyle = .BorderStyle
        mOrigSpecialEffect = .SpecialEffect
        mOrigBackColor = .BackColor
    End With

    If HeaderHeight < MIN_HEADER_HEIGHT Then HeaderHeight = DEFAULT_HEADER_HEIGHT
    If HeaderHeight > mHeight Then HeaderHeight = mHeight
    mHeaderHeight = HeaderHeight

    ' 2. Remove header controls left behind by another instance on this frame
    RemoveHeaderControls

    ' 3. Style the frame (flat, captionless, light-grey content)
    With mFrame
        .Caption = vbNullString
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleNone
        .BackColor = CLR_CONTENT
    End With
    mCollapsedBackColor = ParentBackColor()

    ' 4. Build the header
    BuildHeader HeadingText

    ' 5. Initial state
    Set mVisibility = New Collection
    mExpanded = True
    mInitialized = True
    UpdateArrow
    EnforceGeometry
End Sub

Public Sub Expand()
    Dim ctl As MSForms.Control
    Dim wasVisible As Boolean
    Dim found As Boolean

    If Not mInitialized Then Exit Sub
    If mExpanded Then Exit Sub

    mFrame.BackColor = CLR_CONTENT

    On Error Resume Next
    For Each ctl In mFrame.Controls
        If Not IsHeaderControl(ctl.Name) Then
            wasVisible = StoredVisible(ctl.Name, found)
            If found Then ctl.Visible = wasVisible
        End If
    Next ctl
    On Error GoTo 0

    Set mVisibility = New Collection
    mExpanded = True
    UpdateArrow
    EnforceGeometry
End Sub

Public Sub Collapse()
    Dim ctl As MSForms.Control

    If Not mInitialized Then Exit Sub
    If Not mExpanded Then Exit Sub

    Set mVisibility = New Collection

    On Error Resume Next
    For Each ctl In mFrame.Controls
        If Not IsHeaderControl(ctl.Name) Then
            mVisibility.Add ctl.Visible, ctl.Name   ' remember original state
            ctl.Visible = False
        End If
    Next ctl
    On Error GoTo 0

    ' Content area blends into the parent so only the header bar is seen
    mFrame.BackColor = mCollapsedBackColor
    mExpanded = False
    UpdateArrow
    EnforceGeometry
End Sub

Public Sub Toggle()
    If mExpanded Then
        Collapse
    Else
        Expand
    End If
End Sub

Public Property Get IsExpanded() As Boolean
    IsExpanded = mExpanded
End Property

Public Property Get IsInitialized() As Boolean
    IsInitialized = mInitialized
End Property

Public Property Get Heading() As String
    If Not mTitle Is Nothing Then Heading = mTitle.Caption
End Property

Public Property Let Heading(ByVal Value As String)
    If mTitle Is Nothing Then Exit Property
    If Len(Trim$(Value)) = 0 Then Value = DEFAULT_HEADING
    mTitle.Caption = Value
End Property

Public Property Get Frame() As MSForms.Frame
    Set Frame = mFrame
End Property

' Removes the header and restores the frame's original look and child visibility.
Public Sub Detach()
    If Not mInitialized Then Exit Sub

    Expand
    RemoveHeaderControls

    On Error Resume Next
    With mFrame
        .Caption = mOrigCaption
        .BorderStyle = mOrigBorderStyle
        .SpecialEffect = mOrigSpecialEffect
        .BackColor = mOrigBackColor
    End With
    On Error GoTo 0

    EnforceGeometry
    mInitialized = False
    Set mVisibility = Nothing
    Set mFrame = Nothing
End Sub

'==============================================================================
' EVENTS  (arrow, heading and bar all toggle the panel)
'==============================================================================

Private Sub mArrow_Click()
    Toggle
End Sub

Private Sub mArrow_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    Toggle      ' fast second click would otherwise be swallowed
End Sub

Private Sub mTitle_Click()
    Toggle
End Sub

Private Sub mTitle_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    Toggle
End Sub

Private Sub mBar_Click()
    Toggle
End Sub

Private Sub mBar_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    Toggle
End Sub

'==============================================================================
' PRIVATE HELPERS
'==============================================================================

Private Sub BuildHeader(ByVal HeadingText As String)
    Dim barWidth As Single

    barWidth = mFrame.InsideWidth
    If barWidth <= 0 Then barWidth = mWidth

    If Len(Trim$(HeadingText)) = 0 Then HeadingText = DEFAULT_HEADING

    ' Background bar
    Set mBar = mFrame.Controls.Add("Forms.Label.1", mPrefix & "Bar", True)
    With mBar
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BackColor = CLR_HEADER
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .Left = 0
        .Top = 0
        .Width = barWidth
        .Height = mHeaderHeight
    End With

    ' Arrow (top-left)
    Set mArrow = mFrame.Controls.Add("Forms.Label.1", mPrefix & "Arrow", True)
    With mArrow
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .ForeColor = CLR_TEXT
        .Font.Name = ARROW_FONT
        .Font.Size = TITLE_FONT_SIZE
        .Caption = GLYPH_EXPANDED
        .TextAlign = fmTextAlignCenter
        .Left = ARROW_LEFT
        .ControlTipText = "Expand / collapse"
    End With
    CenterVertically mArrow, ARROW_WIDTH

    ' Heading text
    Set mTitle = mFrame.Controls.Add("Forms.Label.1", mPrefix & "Title", True)
    With mTitle
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .ForeColor = CLR_TEXT
        .Font.Name = TITLE_FONT
        .Font.Size = TITLE_FONT_SIZE
        .Font.Bold = False
        .Caption = HeadingText
        .TextAlign = fmTextAlignLeft
        .Left = ARROW_LEFT + ARROW_WIDTH + TITLE_GAP
    End With
    CenterVertically mTitle, MaxSingle(10, barWidth - mTitle.Left - RIGHT_PAD)

    ' Header must sit above any existing content: bar first, then arrow/title
    mFrame.Controls(mPrefix & "Bar").ZOrder fmZOrderFront
    mFrame.Controls(mPrefix & "Arrow").ZOrder fmZOrderFront
    mFrame.Controls(mPrefix & "Title").ZOrder fmZOrderFront
End Sub

' Labels have no vertical alignment: size to the font, then centre in the bar.
Private Sub CenterVertically(ByVal lbl As MSForms.Label, ByVal fixedWidth As Single)
    lbl.WordWrap = False
    lbl.AutoSize = True
    lbl.AutoSize = False
    lbl.Width = fixedWidth
    lbl.Top = (mHeaderHeight - lbl.Height) / 2
    If lbl.Top < 0 Then lbl.Top = 0
End Sub

Private Sub UpdateArrow()
    If mArrow Is Nothing Then Exit Sub
    If mExpanded Then
        mArrow.Caption = GLYPH_EXPANDED
    Else
        mArrow.Caption = GLYPH_COLLAPSED
    End If
End Sub

Private Function IsHeaderControl(ByVal ctlName As String) As Boolean
    IsHeaderControl = (StrComp(Left$(ctlName, Len(mPrefix)), mPrefix, vbTextCompare) = 0)
End Function

Private Function StoredVisible(ByVal key As String, ByRef found As Boolean) As Boolean
    found = False
    If mVisibility Is Nothing Then Exit Function
    On Error Resume Next
    StoredVisible = mVisibility(key)
    found = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub RemoveHeaderControls()
    Set mBar = Nothing
    Set mArrow = Nothing
    Set mTitle = Nothing
    If mFrame Is Nothing Then Exit Sub
    RemoveIfExists mPrefix & "Title"
    RemoveIfExists mPrefix & "Arrow"
    RemoveIfExists mPrefix & "Bar"
End Sub

Private Sub RemoveIfExists(ByVal ctlName As String)
    Dim ctl As MSForms.Control
    For Each ctl In mFrame.Controls
        If StrComp(ctl.Name, ctlName, vbTextCompare) = 0 Then
            mFrame.Controls.Remove ctlName
            Exit For
        End If
    Next ctl
End Sub

' Colour used for the content area when collapsed, so it blends into the parent.
Private Function ParentBackColor() As Long
    Dim host As Object
    On Error GoTo Fallback
    Set host = mFrame              ' late-bound to reach the extender's Parent
    ParentBackColor = host.Parent.BackColor
    Exit Function
Fallback:
    ParentBackColor = CLR_FALLBACK_FORM
End Function

' Guarantee: the frame's footprint never changes.
Private Sub EnforceGeometry()
    If mFrame Is Nothing Then Exit Sub
    With mFrame
        If .Left <> mLeft Then .Left = mLeft
        If .Top <> mTop Then .Top = mTop
        If .Width <> mWidth Then .Width = mWidth
        If .Height <> mHeight Then .Height = mHeight
    End With
End Sub

Private Function MaxSingle(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then MaxSingle = a Else MaxSingle = b
End Function

Private Sub Class_Terminate()
    ' Release references only; the form may already be unloading.
    Set mBar = Nothing
    Set mArrow = Nothing
    Set mTitle = Nothing
    Set mVisibility = Nothing
    Set mFrame = Nothing
End Sub
