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
' Class   : CCollapsibleFrame   (v3)
'           v2   - custom colours, fonts, collapsible mode
'           v2.1 - step trace, tooltip fix
'           v3   - state-specific heading texts, icon position, custom icons
'
' Purpose : Turns an EXISTING MSForms.Frame into a collapsible panel with a
'           header bar (icon + heading) and a content area.
'
' Rules   : - Frame Left/Top/Width/Height are never changed (verified after
'             every operation).
'           - Child controls are hidden/restored only via .Visible; their
'             positions and values are never touched.
'           - No API calls, no OnTime, no extra classes or modules.
'
' Usage   : Keep the instance in a FORM-LEVEL variable (events stop if the
'           object goes out of scope). Every property can be set before or
'           after Initialize; after Initialize it applies immediately.
'
' Fallbacks (empty strings):
'           ExpandedText  = ""  -> base heading (Initialize HeadingText / Heading)
'           CollapsedText = ""  -> base heading
'           base heading  = ""  -> "Details"
'           ExpandedIcon  = ""  -> ChrW(&H25BC)  (down triangle)
'           CollapsedIcon = ""  -> ChrW(&H25B6)  (right triangle)
'
' Unicode : The VBA editor is ANSI - typing a symbol such as a triangle in code becomes "?".
'           Pass symbols with ChrW, e.g. panel.ExpandedIcon = ChrW(&H25B2).
'==============================================================================

'--- Defaults -----------------------------------------------------------------
Private Const DEFAULT_HEADING        As String = "Details"
Private Const DEFAULT_HEADER_HEIGHT  As Single = 24
Private Const MIN_HEADER_HEIGHT      As Single = 14

Private Const DEFAULT_HEADER_COLOR   As Long = &H555555     ' RGB(85, 85, 85)
Private Const DEFAULT_HEADER_FORE    As Long = &HFFFFFF     ' RGB(255, 255, 255)
Private Const DEFAULT_CONTENT_COLOR  As Long = &HF0F0F0     ' RGB(240, 240, 240)
Private Const DEFAULT_ARROW_COLOR    As Long = &HFFFFFF     ' RGB(255, 255, 255)
Private Const CLR_FALLBACK_FORM      As Long = &H8000000F   ' System button face

Private Const DEFAULT_FONT_NAME      As String = "Segoe UI"
Private Const DEFAULT_FONT_SIZE      As Single = 12
Private Const MIN_FONT_SIZE          As Single = 6
Private Const MAX_FONT_SIZE          As Single = 72

Private Const DEFAULT_ICON_FONT      As String = "Segoe UI Symbol"  ' has the triangle symbols
Private Const ICON_DOWN              As Long = &H25BC               ' down triangle
Private Const ICON_RIGHT             As Long = &H25B6               ' right triangle

Private Const POS_LEFT               As String = "Left"
Private Const POS_RIGHT              As String = "Right"

'--- Layout -------------------------------------------------------------------
Private Const ICON_EDGE_PAD          As Single = 6          ' icon to bar edge
Private Const MIN_ARROW_WIDTH        As Single = 14
Private Const ICON_HPAD              As Single = 4          ' extra width around glyph
Private Const TITLE_GAP              As Single = 4          ' icon to heading
Private Const TITLE_LEFT_NO_ARROW    As Single = 10         ' heading left edge w/o icon
Private Const RIGHT_PAD              As Single = 8          ' heading to right edge
Private Const HEADER_VPAD            As Single = 4          ' min space above/below text

'--- Error numbers ------------------------------------------------------------
Private Const ERR_NO_FRAME           As Long = vbObjectError + 513
Private Const ERR_BAD_FONT_SIZE      As Long = vbObjectError + 514
Private Const ERR_BAD_FONT_NAME      As Long = vbObjectError + 515
Private Const ERR_INIT_FAILED        As Long = vbObjectError + 516
Private Const ERR_BAD_ICON_POS       As Long = vbObjectError + 517

'--- Diagnostics: set True to print each Initialize step to the Immediate window
Private Const TRACE_STEPS            As Boolean = False

'--- Controls -----------------------------------------------------------------
Private mFrame              As MSForms.Frame
Private WithEvents mBar     As MSForms.Label    ' header background (clickable)
Attribute mBar.VB_VarHelpID = -1
Private WithEvents mArrow   As MSForms.Label    ' icon (clickable)
Attribute mArrow.VB_VarHelpID = -1
Private WithEvents mTitle   As MSForms.Label    ' heading text (clickable)
Attribute mTitle.VB_VarHelpID = -1

'--- State --------------------------------------------------------------------
Private mInitialized        As Boolean
Private mExpanded           As Boolean
Private mRequestedHeight    As Single           ' header height asked for
Private mHeaderHeight       As Single           ' header height actually used
Private mPrefix             As String
Private mVisibility         As Collection       ' child Visible states, keyed by name
Private mCollapsedBackColor As Long
Private mStep               As String           ' last step reached (diagnostics)

'--- Configurable style / behaviour -------------------------------------------
Private mHeaderColor        As Long
Private mHeaderFontColor    As Long
Private mContentColor       As Long
Private mArrowColor         As Long
Private mFontName           As String
Private mFontSize           As Single
Private mFontBold           As Boolean
Private mFontItalic         As Boolean
Private mCollapsible        As Boolean

Private mHeading            As String           ' base heading
Private mExpandedText       As String           ' "" = use base heading
Private mCollapsedText      As String           ' "" = use base heading
Private mIconPosition       As String           ' "Left" / "Right"
Private mExpandedIcon       As String           ' "" = default down triangle
Private mCollapsedIcon      As String           ' "" = default right triangle
Private mIconFontName       As String

'--- Original geometry (never allowed to change) ------------------------------
Private mLeft               As Single
Private mTop                As Single
Private mWidth              As Single
Private mHeight             As Single

'--- Original appearance (restored by Detach) ---------------------------------
Private mOrigCaption        As String
Private mOrigBorderStyle    As Long
Private mOrigSpecialEffect  As Long
Private mOrigBackColor      As Long

'==============================================================================
' LIFECYCLE
'==============================================================================

Private Sub Class_Initialize()
    mHeaderColor = DEFAULT_HEADER_COLOR
    mHeaderFontColor = DEFAULT_HEADER_FORE
    mContentColor = DEFAULT_CONTENT_COLOR
    mArrowColor = DEFAULT_ARROW_COLOR
    mFontName = DEFAULT_FONT_NAME
    mFontSize = DEFAULT_FONT_SIZE
    mFontBold = False
    mFontItalic = False
    mCollapsible = True
    mRequestedHeight = DEFAULT_HEADER_HEIGHT

    mHeading = vbNullString
    mExpandedText = vbNullString
    mCollapsedText = vbNullString
    mIconPosition = POS_LEFT
    mExpandedIcon = vbNullString        ' resolved to default by IconFor()
    mCollapsedIcon = vbNullString
    mIconFontName = DEFAULT_ICON_FONT
End Sub

Private Sub Class_Terminate()
    ' Release references only; the form may already be unloading.
    Set mBar = Nothing
    Set mArrow = Nothing
    Set mTitle = Nothing
    Set mVisibility = Nothing
    Set mFrame = Nothing
End Sub

'==============================================================================
' PUBLIC API - CORE
'==============================================================================

Public Sub Initialize(ByVal FrameObject As MSForms.Frame, _
                      Optional ByVal HeadingText As String = vbNullString, _
                      Optional ByVal HeaderHeight As Single = DEFAULT_HEADER_HEIGHT)

    If FrameObject Is Nothing Then
        Err.Raise ERR_NO_FRAME, "CCollapsibleFrame.Initialize", _
                  "FrameObject is Nothing. Pass an existing Frame, e.g. Me.Frame1."
    End If

    On Error GoTo Failed
    Trace "Initialize: start (" & FrameObject.Name & ")"

    ' Re-initialisation: undo previous styling first (no duplicate headers).
    ' All configured properties are kept.
    If mInitialized Then Detach

    Set mFrame = FrameObject
    mPrefix = "ccf_" & mFrame.Name & "_"

    ' A heading passed here becomes the base heading; an omitted one keeps
    ' any base heading already set through the Heading property.
    If Len(Trim$(HeadingText)) > 0 Then mHeading = HeadingText

    ' 1. Remember geometry and appearance
    Trace "Initialize: read frame properties"
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
    mRequestedHeight = HeaderHeight

    ' 2. Remove header controls left behind by another instance on this frame
    Trace "Initialize: remove old header controls"
    RemoveHeaderControls

    ' 3. Style the frame (flat, captionless)
    Trace "Initialize: style frame"
    With mFrame
        .Caption = vbNullString
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleNone
        .BackColor = mContentColor
    End With
    Trace "Initialize: read parent colour"
    mCollapsedBackColor = ParentBackColor()

    ' 4. Build the header and apply all styling
    Trace "Initialize: build header"
    BuildHeader

    ' 5. Initial state
    Set mVisibility = New Collection
    mExpanded = True
    mInitialized = True
    Trace "Initialize: apply colours"
    ApplyColors
    Trace "Initialize: apply fonts"
    ApplyTitleFont          ' also applies icon font + layout
    Trace "Initialize: update header state"
    UpdateHeaderState
    Trace "Initialize: layout header"
    LayoutHeader
    Trace "Initialize: enforce geometry"
    EnforceGeometry
    Trace "Initialize: done"
    Exit Sub

Failed:
    Err.Raise ERR_INIT_FAILED, "CCollapsibleFrame.Initialize", _
              "Initialize failed at step '" & mStep & "': " & _
              Err.Description & " (Err " & Err.Number & ")"
End Sub

Public Sub Expand()
    Dim ctl As MSForms.Control
    Dim wasVisible As Boolean
    Dim found As Boolean

    If Not mInitialized Then Exit Sub
    If mExpanded Then Exit Sub

    mFrame.BackColor = mContentColor

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
    UpdateHeaderState
    EnforceGeometry
End Sub

Public Sub Collapse()
    Dim ctl As MSForms.Control

    If Not mInitialized Then Exit Sub
    If Not mCollapsible Then Exit Sub       ' non-collapsible: never hide content
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
    UpdateHeaderState
    EnforceGeometry
End Sub

Public Sub Toggle()
    If Not mCollapsible Then Exit Sub
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

Public Property Get Frame() As MSForms.Frame
    Set Frame = mFrame
End Property

' Removes the header and restores the frame's original look and child visibility.
Public Sub Detach()
    If Not mInitialized Then Exit Sub

    ' Expand even when non-collapsible (Expand itself has no such guard)
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
' PUBLIC API - HEADING TEXT
'==============================================================================

' Get: the text currently displayed (or that would be displayed).
' Let: sets the BASE heading, used whenever ExpandedText/CollapsedText is "".
Public Property Get Heading() As String
    Heading = CurrentHeadingText()
End Property

Public Property Let Heading(ByVal Value As String)
    mHeading = Value
    UpdateTitleText
End Property

Public Property Get ExpandedText() As String
    ExpandedText = mExpandedText
End Property

Public Property Let ExpandedText(ByVal Value As String)
    mExpandedText = Value
    UpdateTitleText          ' visible change only if currently expanded
End Property

Public Property Get CollapsedText() As String
    CollapsedText = mCollapsedText
End Property

Public Property Let CollapsedText(ByVal Value As String)
    mCollapsedText = Value
    UpdateTitleText          ' visible change only if currently collapsed
End Property

'==============================================================================
' PUBLIC API - ICON
'==============================================================================

Public Property Get IconPosition() As String
    IconPosition = mIconPosition
End Property

Public Property Let IconPosition(ByVal Value As String)
    Select Case LCase$(Trim$(Value))
        Case "left":  mIconPosition = POS_LEFT
        Case "right": mIconPosition = POS_RIGHT
        Case Else
            Err.Raise ERR_BAD_ICON_POS, "CCollapsibleFrame.IconPosition", _
                      "IconPosition must be ""Left"" or ""Right"" (got """ & Value & _
                      """). Current value kept: """ & mIconPosition & """."
    End Select
    LayoutHeader
End Property

' Shown while expanded (click collapses). "" = default down triangle.
Public Property Get ExpandedIcon() As String
    ExpandedIcon = IconFor(True)
End Property

Public Property Let ExpandedIcon(ByVal Value As String)
    mExpandedIcon = Value
    UpdateArrow
    LayoutHeader             ' icon width may differ
End Property

' Shown while collapsed (click expands). "" = default right triangle.
Public Property Get CollapsedIcon() As String
    CollapsedIcon = IconFor(False)
End Property

Public Property Let CollapsedIcon(ByVal Value As String)
    mCollapsedIcon = Value
    UpdateArrow
    LayoutHeader
End Property

' Font used to draw the icon. Default "Segoe UI Symbol" (covers triangles, + and -).
' Set to "Marlett" with icons "6"/"4" to get the v1/v2 look.
Public Property Get IconFontName() As String
    IconFontName = mIconFontName
End Property

Public Property Let IconFontName(ByVal Value As String)
    Value = Trim$(Value)
    If Len(Value) = 0 Then
        Err.Raise ERR_BAD_FONT_NAME, "CCollapsibleFrame.IconFontName", _
                  "Icon font name cannot be empty."
    End If

    mIconFontName = Value
    If mArrow Is Nothing Then Exit Property  ' applied at Initialize

    If Not TrySetFontName(mArrow, Value) Then
        Debug.Print "CCollapsibleFrame: icon font '" & Value & "' could not be applied; kept '" & _
                    mArrow.Font.Name & "'."
        mIconFontName = mArrow.Font.Name
    End If
    LayoutHeader
End Property

'==============================================================================
' PUBLIC API - COLOURS  (Long values, e.g. RGB(45, 55, 72))
'==============================================================================

Public Property Get HeaderColor() As Long
    HeaderColor = mHeaderColor
End Property

Public Property Let HeaderColor(ByVal Value As Long)
    mHeaderColor = Value
    ApplyColors
End Property

Public Property Get HeaderFontColor() As Long
    HeaderFontColor = mHeaderFontColor
End Property

Public Property Let HeaderFontColor(ByVal Value As Long)
    mHeaderFontColor = Value
    ApplyColors
End Property

Public Property Get ContentColor() As Long
    ContentColor = mContentColor
End Property

Public Property Let ContentColor(ByVal Value As Long)
    mContentColor = Value
    ApplyColors      ' while collapsed, takes effect on the next Expand
End Property

Public Property Get ArrowColor() As Long
    ArrowColor = mArrowColor
End Property

Public Property Let ArrowColor(ByVal Value As Long)
    mArrowColor = Value
    ApplyColors
End Property

'==============================================================================
' PUBLIC API - HEADING FONT
'==============================================================================

Public Property Get HeaderFontName() As String
    HeaderFontName = mFontName
End Property

Public Property Let HeaderFontName(ByVal Value As String)
    Value = Trim$(Value)
    If Len(Value) = 0 Then
        Err.Raise ERR_BAD_FONT_NAME, "CCollapsibleFrame.HeaderFontName", _
                  "Font name cannot be empty."
    End If

    mFontName = Value
    If mTitle Is Nothing Then Exit Property   ' applied at Initialize

    If Not TrySetFontName(mTitle, Value) Then
        Debug.Print "CCollapsibleFrame: font '" & Value & "' could not be applied; kept '" & _
                    mTitle.Font.Name & "'."
        mFontName = mTitle.Font.Name
    End If
    LayoutHeader
End Property

Public Property Get HeaderFontSize() As Single
    HeaderFontSize = mFontSize
End Property

Public Property Let HeaderFontSize(ByVal Value As Single)
    If Value < MIN_FONT_SIZE Or Value > MAX_FONT_SIZE Then
        Err.Raise ERR_BAD_FONT_SIZE, "CCollapsibleFrame.HeaderFontSize", _
                  "Font size must be between " & MIN_FONT_SIZE & " and " & MAX_FONT_SIZE & _
                  " points (got " & Value & ")."
    End If
    mFontSize = Value
    ApplyTitleFont
End Property

Public Property Get HeaderFontBold() As Boolean
    HeaderFontBold = mFontBold
End Property

Public Property Let HeaderFontBold(ByVal Value As Boolean)
    mFontBold = Value
    ApplyTitleFont
End Property

Public Property Get HeaderFontItalic() As Boolean
    HeaderFontItalic = mFontItalic
End Property

Public Property Let HeaderFontItalic(ByVal Value As Boolean)
    mFontItalic = Value
    ApplyTitleFont
End Property

'==============================================================================
' PUBLIC API - BEHAVIOUR
'==============================================================================

Public Property Get Collapsible() As Boolean
    Collapsible = mCollapsible
End Property

Public Property Let Collapsible(ByVal Value As Boolean)
    mCollapsible = Value
    If Not mInitialized Then Exit Property   ' applied at Initialize

    If Not mCollapsible And Not mExpanded Then Expand
    UpdateHeaderState
    LayoutHeader                              ' shows/hides icon, shifts heading
End Property

'==============================================================================
' EVENTS  (icon, heading and bar all toggle; Toggle respects Collapsible.
'          Handlers never look at the displayed symbol.)
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
' PRIVATE HELPERS - BUILD & STYLE
'==============================================================================

Private Sub BuildHeader()
    ' Background bar
    Trace "BuildHeader: add bar"
    Set mBar = mFrame.Controls.Add("Forms.Label.1", mPrefix & "Bar", True)
    With mBar
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .Left = 0
        .Top = 0
    End With

    ' Icon
    Trace "BuildHeader: add icon"
    Set mArrow = mFrame.Controls.Add("Forms.Label.1", mPrefix & "Arrow", True)
    With mArrow
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .Caption = IconFor(True)
        .TextAlign = fmTextAlignCenter
        .WordWrap = False
        .ControlTipText = "Expand / collapse"
    End With

    ' Heading text
    Trace "BuildHeader: add title"
    Set mTitle = mFrame.Controls.Add("Forms.Label.1", mPrefix & "Title", True)
    With mTitle
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .Caption = HeadingFor(True)
        .TextAlign = fmTextAlignLeft
        .WordWrap = False
    End With

    ' Header must sit above any existing content: bar first, then icon/title
    Trace "BuildHeader: z-order"
    HeaderCtl("Bar").ZOrder fmZOrderFront
    HeaderCtl("Arrow").ZOrder fmZOrderFront
    HeaderCtl("Title").ZOrder fmZOrderFront
End Sub

Private Sub ApplyColors()
    If Not mBar Is Nothing Then mBar.BackColor = mHeaderColor
    If Not mTitle Is Nothing Then mTitle.ForeColor = mHeaderFontColor
    If Not mArrow Is Nothing Then mArrow.ForeColor = mArrowColor
    If mInitialized Then
        If mExpanded Then mFrame.BackColor = mContentColor
    End If
End Sub

' Applies heading font settings, the icon font, then re-lays out the header.
' Never touches caption text.
Private Sub ApplyTitleFont()
    If mTitle Is Nothing Then Exit Sub

    Trace "ApplyTitleFont: size/bold/italic"
    With mTitle.Font
        .Size = mFontSize
        .Bold = mFontBold
        .Italic = mFontItalic
    End With

    Trace "ApplyTitleFont: font name '" & mFontName & "'"
    If Not TrySetFontName(mTitle, mFontName) Then
        If Not TrySetFontName(mTitle, DEFAULT_FONT_NAME) Then
            Debug.Print "CCollapsibleFrame: could not apply font '" & mFontName & "'."
        End If
        mFontName = mTitle.Font.Name
    End If

    If Not mArrow Is Nothing Then
        Trace "ApplyTitleFont: icon font '" & mIconFontName & "'"
        If Not TrySetFontName(mArrow, mIconFontName) Then
            If Not TrySetFontName(mArrow, DEFAULT_ICON_FONT) Then
                Debug.Print "CCollapsibleFrame: could not apply icon font '" & mIconFontName & "'."
            End If
            mIconFontName = mArrow.Font.Name
        End If
        mArrow.Font.Size = mFontSize          ' icon scales with the heading
    End If

    LayoutHeader
End Sub

' Returns False (and leaves the previous font in place) if the name is rejected.
Private Function TrySetFontName(ByVal lbl As MSForms.Label, ByVal fontName As String) As Boolean
    Dim previous As String

    previous = lbl.Font.Name
    On Error GoTo Failed
    lbl.Font.Name = fontName
    TrySetFontName = True
    Exit Function

Failed:
    Err.Clear
    On Error Resume Next
    lbl.Font.Name = previous
    On Error GoTo 0
    TrySetFontName = False
End Function

' Sizes and positions bar, icon and heading. Grows the header if the font
' needs more room, so the heading always stays vertically centred.
Private Sub LayoutHeader()
    Dim barWidth As Single
    Dim iconWidth As Single
    Dim titleLeft As Single
    Dim titleRight As Single
    Dim h As Single
    Dim arrowCtl As MSForms.Control

    If mBar Is Nothing Or mTitle Is Nothing Or mArrow Is Nothing Then Exit Sub

    barWidth = mFrame.InsideWidth
    If barWidth <= 0 Then barWidth = mWidth

    ' Measure heading height and the wider of the two icons
    Trace "LayoutHeader: measure"
    FitToText mTitle
    iconWidth = MaxSingle(MIN_ARROW_WIDTH, WidestIcon() + ICON_HPAD)

    h = mRequestedHeight
    If mTitle.Height + 2 * HEADER_VPAD > h Then h = mTitle.Height + 2 * HEADER_VPAD
    If mArrow.Height + 2 * HEADER_VPAD > h Then h = mArrow.Height + 2 * HEADER_VPAD
    If h > mHeight Then h = mHeight
    mHeaderHeight = h

    ' Bar
    mBar.Left = 0
    mBar.Top = 0
    mBar.Width = barWidth
    mBar.Height = h

    ' Icon visibility
    Trace "LayoutHeader: icon visibility"
    Set arrowCtl = HeaderCtl("Arrow")
    If arrowCtl.Visible <> mCollapsible Then arrowCtl.Visible = mCollapsible

    ' Icon + heading positions
    mArrow.Width = iconWidth
    mArrow.Top = MaxSingle(0, (h - mArrow.Height) / 2)

    If Not mCollapsible Then
        titleLeft = TITLE_LEFT_NO_ARROW
        titleRight = barWidth - RIGHT_PAD
    ElseIf mIconPosition = POS_RIGHT Then
        mArrow.Left = MaxSingle(0, barWidth - ICON_EDGE_PAD - iconWidth)
        titleLeft = TITLE_LEFT_NO_ARROW
        titleRight = mArrow.Left - TITLE_GAP
    Else
        mArrow.Left = ICON_EDGE_PAD
        titleLeft = ICON_EDGE_PAD + iconWidth + TITLE_GAP
        titleRight = barWidth - RIGHT_PAD
    End If

    mTitle.Left = titleLeft
    mTitle.Width = MaxSingle(10, titleRight - titleLeft)
    mTitle.Top = MaxSingle(0, (h - mTitle.Height) / 2)
End Sub

' Natural width of the wider icon, so the icon box (and the heading next to
' it) does not shift when the state changes. Leaves the current icon shown
' and the label auto-sized to it (height is used for centring).
Private Function WidestIcon() As Single
    Dim w As Single

    mArrow.Caption = IconFor(True)
    FitToText mArrow
    w = mArrow.Width

    mArrow.Caption = IconFor(False)
    FitToText mArrow
    If mArrow.Width > w Then w = mArrow.Width

    mArrow.Caption = IconFor(mExpanded Or Not mInitialized)
    FitToText mArrow
    WidestIcon = w
End Function

' Labels have no vertical alignment: size to the font to learn the text height.
Private Sub FitToText(ByVal lbl As MSForms.Label)
    lbl.WordWrap = False
    lbl.AutoSize = True
    lbl.AutoSize = False
End Sub

'==============================================================================
' PRIVATE HELPERS - STATE TEXT / ICON
'==============================================================================

Private Sub UpdateHeaderState()
    UpdateArrow
    UpdateTitleText
End Sub

Private Sub UpdateArrow()
    If mArrow Is Nothing Then Exit Sub
    mArrow.Caption = IconFor(mExpanded)
End Sub

Private Sub UpdateTitleText()
    If mTitle Is Nothing Then Exit Sub
    mTitle.Caption = CurrentHeadingText()
End Sub

Private Function CurrentHeadingText() As String
    ' Before Initialize the panel counts as expanded.
    CurrentHeadingText = HeadingFor(mExpanded Or Not mInitialized)
End Function

Private Function HeadingFor(ByVal expandedState As Boolean) As String
    Dim s As String
    If expandedState Then s = mExpandedText Else s = mCollapsedText
    If Len(Trim$(s)) = 0 Then s = mHeading
    If Len(Trim$(s)) = 0 Then s = DEFAULT_HEADING
    HeadingFor = s
End Function

Private Function IconFor(ByVal expandedState As Boolean) As String
    If expandedState Then
        If Len(mExpandedIcon) > 0 Then IconFor = mExpandedIcon Else IconFor = ChrW$(ICON_DOWN)
    Else
        If Len(mCollapsedIcon) > 0 Then IconFor = mCollapsedIcon Else IconFor = ChrW$(ICON_RIGHT)
    End If
End Function

'==============================================================================
' PRIVATE HELPERS - HOUSEKEEPING
'==============================================================================

Private Sub Trace(ByVal stepName As String)
    mStep = stepName
    If TRACE_STEPS Then Debug.Print "CCollapsibleFrame | " & stepName
End Sub

' Extender view of a header control (Visible, ZOrder live here).
Private Function HeaderCtl(ByVal suffix As String) As MSForms.Control
    Set HeaderCtl = mFrame.Controls(mPrefix & suffix)
End Function

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
