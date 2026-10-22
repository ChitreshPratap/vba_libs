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
' Class   : CCollapsibleFrame   (v4L)
'           v2   - custom colours, fonts, collapsible mode
'           v2.1 - step trace, tooltip fix
'           v3   - state-specific heading texts, icon position, custom icons
'           v4L  - Label container mode + ContainerTransparent
'           v5   - pluggable header styles (HeaderStyle As IHeaderStyle)
'
' Header styles (optional):
'           Needs class modules IHeaderStyle (interface) plus any style such
'           as CHeaderStyleBottomBorder or CHeaderStyleDivider.
'               Set panel.HeaderStyle = myStyle      ' before or after Initialize
'               Set panel.HeaderStyle = Nothing      ' back to the default header
'           The container keeps all behaviour; a style only draws decoration
'           (and the Divider style positions the heading inside the text area
'           the container gives it). Style controls are named
'           "ccf_<container>_sty_*" so they are never treated as content.
'
' Purpose : Turns an existing container into a collapsible panel with a
'           header bar (icon + heading) and a content area.
'
' TWO CONTAINER MODES (chosen automatically by Initialize):
'
'   Frame mode  - Initialize Me.Frame1, ...
'                 The Frame hosts the content controls (unchanged behaviour).
'                 A Frame is ALWAYS opaque, so ContainerTransparent = True
'                 paints it with its parent's BackColor (looks transparent on
'                 a solid-colour form; a form Picture will NOT show through).
'
'   Label mode  - Initialize Me.Label1, ...
'                 The Label is only the panel's BACKGROUND. It cannot contain
'                 controls. Content controls sit on the same parent as the
'                 Label (the form, a Frame or a Page), placed inside the
'                 Label's rectangle. The class sends the Label to the back,
'                 puts the header on top, and treats every sibling control
'                 that lies fully inside the Label's rectangle as content.
'                 ContainerTransparent = True gives TRUE transparency
'                 (BackStyle = fmBackStyleTransparent): the form's colour and
'                 Picture show through.
'
' Rules   : - Container Left/Top/Width/Height are never changed (verified
'             after every operation).
'           - Content controls are hidden/restored only via .Visible; their
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
' Unicode : The VBA editor is ANSI - typing a symbol such as a triangle in
'           code becomes "?". Pass symbols with ChrW, e.g. ChrW(&H25B2).
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
Private Const INSIDE_TOLERANCE       As Single = 0.75       ' Label mode containment test
Private Const STYLE_TAG              As String = "sty_"     ' prefix for style-owned controls

'--- Error numbers ------------------------------------------------------------
Private Const ERR_NO_FRAME           As Long = vbObjectError + 513
Private Const ERR_BAD_FONT_SIZE      As Long = vbObjectError + 514
Private Const ERR_BAD_FONT_NAME      As Long = vbObjectError + 515
Private Const ERR_INIT_FAILED        As Long = vbObjectError + 516
Private Const ERR_BAD_ICON_POS       As Long = vbObjectError + 517
Private Const ERR_BAD_CONTAINER      As Long = vbObjectError + 518

'--- Diagnostics: set True to print each Initialize step to the Immediate window
Private Const TRACE_STEPS            As Boolean = False

'--- Container ----------------------------------------------------------------
Private mContainer          As Object           ' the Frame or Label passed in
Private mFrame              As MSForms.Frame    ' set in Frame mode only
Private mLabel              As MSForms.Label    ' set in Label mode only
Private mHost               As Object           ' owner of header controls
Private mHostName           As String
Private mIsLabelMode        As Boolean
Private mOffX               As Single           ' header offset inside mHost
Private mOffY               As Single

'--- Header style (optional strategy object) ----------------------------------
Private mStyle              As IHeaderStyle
Private mRegionLeft         As Single           ' heading area given to the style
Private mRegionRight        As Single

'--- Header controls ----------------------------------------------------------
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
Private mVisibility         As Collection       ' Visible state per hidden control
Private mHiddenNames        As Collection       ' names hidden by the last Collapse
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
Private mContainerTransparent As Boolean

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
Private mOrigBackStyle      As Long             ' Label mode only

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
    mContainerTransparent = False
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
    Set mHiddenNames = Nothing
    Set mFrame = Nothing
    Set mLabel = Nothing
    Set mHost = Nothing
    Set mContainer = Nothing
    Set mStyle = Nothing
End Sub

'==============================================================================
' PUBLIC API - CORE
'==============================================================================

' ContainerObject: an existing MSForms Frame (Frame mode, as before) or an
' MSForms Label (Label mode). Existing calls such as
'   panel.Initialize Me.Frame1, "Customer Information"
' work unchanged.
Public Sub Initialize(ByVal ContainerObject As Object, _
                      Optional ByVal HeadingText As String = vbNullString, _
                      Optional ByVal HeaderHeight As Single = DEFAULT_HEADER_HEIGHT)

    Dim useLabel As Boolean

    ' --- Validate the container type (before the trace handler) -------------
    If ContainerObject Is Nothing Then
        Err.Raise ERR_NO_FRAME, "CCollapsibleFrame.Initialize", _
                  "Container is Nothing. Pass an existing Frame or Label, e.g. Me.Frame1."
    End If

    If TypeOf ContainerObject Is MSForms.Frame Then
        useLabel = False
    ElseIf TypeOf ContainerObject Is MSForms.Label Then
        useLabel = True
    Else
        Err.Raise ERR_BAD_CONTAINER, "CCollapsibleFrame.Initialize", _
                  "Container must be an MSForms Frame or Label (got " & _
                  TypeName(ContainerObject) & ")."
    End If

    On Error GoTo Failed
    Trace "Initialize: start (" & ContainerObject.Name & ")"

    ' Re-initialisation: undo previous styling first (no duplicate headers).
    ' All configured properties are kept.
    If mInitialized Then Detach

    ' --- Bind the container --------------------------------------------------
    Set mContainer = ContainerObject
    mIsLabelMode = useLabel
    If mIsLabelMode Then
        Set mLabel = ContainerObject
        Set mFrame = Nothing
        Trace "Initialize: resolve Label parent"
        Set mHost = ContainerObject.Parent          ' form, Frame or Page
    Else
        Set mFrame = ContainerObject
        Set mLabel = Nothing
        Set mHost = mFrame
    End If
    mHostName = SafeName(mHost)
    mPrefix = "ccf_" & mContainer.Name & "_"

    ' A heading passed here becomes the base heading; an omitted one keeps
    ' any base heading already set through the Heading property.
    If Len(Trim$(HeadingText)) > 0 Then mHeading = HeadingText

    ' --- 1. Remember geometry and appearance ---------------------------------
    Trace "Initialize: read container properties"
    mLeft = mContainer.Left
    mTop = mContainer.Top
    mWidth = mContainer.Width
    mHeight = mContainer.Height

    If mIsLabelMode Then
        With mLabel
            mOrigCaption = .Caption
            mOrigBorderStyle = .BorderStyle
            mOrigSpecialEffect = .SpecialEffect
            mOrigBackColor = .BackColor
            mOrigBackStyle = .BackStyle
        End With
        mOffX = mLeft
        mOffY = mTop
    Else
        With mFrame
            mOrigCaption = .Caption
            mOrigBorderStyle = .BorderStyle
            mOrigSpecialEffect = .SpecialEffect
            mOrigBackColor = .BackColor
        End With
        mOffX = 0
        mOffY = 0
    End If

    If HeaderHeight < MIN_HEADER_HEIGHT Then HeaderHeight = DEFAULT_HEADER_HEIGHT
    mRequestedHeight = HeaderHeight

    ' --- 2. Remove header controls left behind by another instance ---------
    Trace "Initialize: remove old header controls"
    RemoveHeaderControls
    RemoveStyleLeftovers

    ' --- 3. Style the container (flat, captionless) --------------------------
    Trace "Initialize: style container"
    If mIsLabelMode Then
        With mLabel
            .Caption = vbNullString
            .SpecialEffect = fmSpecialEffectFlat
            .BorderStyle = fmBorderStyleNone
        End With
        ' Background Label goes behind every sibling, so it never covers content
        Trace "Initialize: send background Label to back"
        HostCtl(mContainer.Name).ZOrder fmZOrderBack
    Else
        With mFrame
            .Caption = vbNullString
            .SpecialEffect = fmSpecialEffectFlat
            .BorderStyle = fmBorderStyleNone
        End With
    End If

    ' --- 4. Build the header (always on top) ---------------------------------
    Trace "Initialize: build header"
    BuildHeader

    ' --- 5. Initial state ----------------------------------------------------
    Set mVisibility = New Collection
    Set mHiddenNames = New Collection
    mExpanded = True
    mInitialized = True
    Trace "Initialize: apply colours"
    ApplyColors             ' includes content background / transparency
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
    Dim nm As Variant

    If Not mInitialized Then Exit Sub
    If mExpanded Then Exit Sub

    ' Restore exactly the controls the last Collapse hid, to their old state.
    ' A control removed in the meantime is skipped.
    If Not mHiddenNames Is Nothing Then
        On Error Resume Next
        For Each nm In mHiddenNames
            mHost.Controls(CStr(nm)).Visible = CBool(mVisibility(CStr(nm)))
        Next nm
        On Error GoTo 0
    End If

    Set mVisibility = New Collection
    Set mHiddenNames = New Collection
    mExpanded = True
    ApplyContentBackground
    UpdateHeaderState
    EnforceGeometry
End Sub

Public Sub Collapse()
    Dim ctl As MSForms.Control

    If Not mInitialized Then Exit Sub
    If Not mCollapsible Then Exit Sub       ' non-collapsible: never hide content
    If Not mExpanded Then Exit Sub

    Set mVisibility = New Collection
    Set mHiddenNames = New Collection

    On Error Resume Next
    For Each ctl In mHost.Controls
        If IsContentControl(ctl) Then
            mVisibility.Add ctl.Visible, ctl.Name   ' remember original state
            If Err.Number = 0 Then mHiddenNames.Add ctl.Name
            Err.Clear
            ctl.Visible = False
        End If
    Next ctl
    On Error GoTo 0

    ' Content area blends away so only the header bar is seen
    mExpanded = False
    ApplyContentBackground
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

' Frame mode: the Frame. Label mode: Nothing (use Container).
Public Property Get Frame() As MSForms.Frame
    Set Frame = mFrame
End Property

' The Frame or Label passed to Initialize.
Public Property Get Container() As Object
    Set Container = mContainer
End Property

Public Property Get IsLabelMode() As Boolean
    IsLabelMode = mIsLabelMode
End Property

'==============================================================================
' PUBLIC API - HEADER STYLE
'==============================================================================

' Optional decoration strategy. Nothing = the default header.
' Use Set:  Set panel.HeaderStyle = myBottomBorderStyle
' One style instance should serve one panel at a time.
Public Property Get HeaderStyle() As IHeaderStyle
    Set HeaderStyle = mStyle
End Property

Public Property Set HeaderStyle(ByVal Value As IHeaderStyle)
    If Not mStyle Is Nothing Then
        If Not Value Is mStyle Then mStyle.Clear  ' remove old decoration only
    End If
    Set mStyle = Value
    LayoutHeader          ' restores default heading layout, then renders style
End Property

' Removes the header and restores the container's original look and the
' content controls' visibility.
Public Sub Detach()
    If Not mInitialized Then Exit Sub

    ' Expand even when non-collapsible (Expand itself has no such guard)
    Expand
    If Not mStyle Is Nothing Then mStyle.Clear   ' style keeps its settings
    RemoveHeaderControls

    On Error Resume Next
    If mIsLabelMode Then
        With mLabel
            .Caption = mOrigCaption
            .BorderStyle = mOrigBorderStyle
            .SpecialEffect = mOrigSpecialEffect
            .BackColor = mOrigBackColor
            .BackStyle = mOrigBackStyle
        End With
    Else
        With mFrame
            .Caption = mOrigCaption
            .BorderStyle = mOrigBorderStyle
            .SpecialEffect = mOrigSpecialEffect
            .BackColor = mOrigBackColor
        End With
    End If
    On Error GoTo 0

    EnforceGeometry
    mInitialized = False
    Set mVisibility = Nothing
    Set mHiddenNames = Nothing
    Set mFrame = Nothing
    Set mLabel = Nothing
    Set mHost = Nothing
    Set mContainer = Nothing
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
    ApplyColors      ' shown only when expanded and not transparent
End Property

Public Property Get ArrowColor() As Long
    ArrowColor = mArrowColor
End Property

Public Property Let ArrowColor(ByVal Value As Long)
    mArrowColor = Value
    ApplyColors
End Property

'==============================================================================
' PUBLIC API - TRANSPARENCY
'==============================================================================

' True  : content area is transparent.
'         Label mode - real transparency (BackStyle = Transparent).
'         Frame mode - Frame painted in its parent's BackColor (closest
'                      possible; a Frame cannot be transparent).
'         The header always stays opaque in HeaderColor.
' False : content area uses ContentColor (default).
Public Property Get ContainerTransparent() As Boolean
    ContainerTransparent = mContainerTransparent
End Property

Public Property Let ContainerTransparent(ByVal Value As Boolean)
    mContainerTransparent = Value
    ApplyContentBackground          ' no-op before Initialize; applied there
End Property

' Frame mode: re-reads the parent's colour after you change the form's or
' parent's BackColor. Label mode needs no refresh (true transparency).
Public Sub RefreshBackground()
    ApplyContentBackground
End Sub

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
    Set mBar = mHost.Controls.Add("Forms.Label.1", mPrefix & "Bar", True)
    With mBar
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .Left = mOffX
        .Top = mOffY
    End With

    ' Icon
    Trace "BuildHeader: add icon"
    Set mArrow = mHost.Controls.Add("Forms.Label.1", mPrefix & "Arrow", True)
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
    Set mTitle = mHost.Controls.Add("Forms.Label.1", mPrefix & "Title", True)
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
    HostCtl(mPrefix & "Bar").ZOrder fmZOrderFront
    HostCtl(mPrefix & "Arrow").ZOrder fmZOrderFront
    HostCtl(mPrefix & "Title").ZOrder fmZOrderFront
End Sub

Private Sub ApplyColors()
    If Not mBar Is Nothing Then mBar.BackColor = mHeaderColor
    If Not mTitle Is Nothing Then mTitle.ForeColor = mHeaderFontColor
    If Not mArrow Is Nothing Then mArrow.ForeColor = mArrowColor
    ApplyContentBackground
End Sub

' Single place that decides how the content area is painted.
'
'                     expanded & solid      expanded & transparent / collapsed
'   Label mode        opaque ContentColor   BackStyle = Transparent
'   Frame mode        ContentColor          parent's BackColor
Private Sub ApplyContentBackground()
    Dim target As Long

    If Not mInitialized Then Exit Sub
    If mContainer Is Nothing Then Exit Sub

    If mIsLabelMode Then
        If mContainerTransparent Or Not mExpanded Then
            If mLabel.BackStyle <> fmBackStyleTransparent Then _
                mLabel.BackStyle = fmBackStyleTransparent
        Else
            mLabel.BackColor = mContentColor
            If mLabel.BackStyle <> fmBackStyleOpaque Then _
                mLabel.BackStyle = fmBackStyleOpaque
        End If
    Else
        If mContainerTransparent Or Not mExpanded Then
            target = ParentBackColor()
        Else
            target = mContentColor
        End If
        If mFrame.BackColor <> target Then mFrame.BackColor = target
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
' All positions are relative to the container's top-left (mOffX, mOffY).
Private Sub LayoutHeader()
    Dim barWidth As Single
    Dim iconWidth As Single
    Dim titleLeft As Single
    Dim titleRight As Single
    Dim iconLeft As Single
    Dim h As Single
    Dim arrowCtl As MSForms.Control

    If mBar Is Nothing Or mTitle Is Nothing Or mArrow Is Nothing Then Exit Sub

    If mIsLabelMode Then
        barWidth = mWidth
    Else
        barWidth = mFrame.InsideWidth
        If barWidth <= 0 Then barWidth = mWidth
    End If

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
    mBar.Left = mOffX
    mBar.Top = mOffY
    mBar.Width = barWidth
    mBar.Height = h

    ' Icon visibility
    Trace "LayoutHeader: icon visibility"
    Set arrowCtl = HostCtl(mPrefix & "Arrow")
    If arrowCtl.Visible <> mCollapsible Then arrowCtl.Visible = mCollapsible

    ' Icon + heading positions (relative to the bar)
    If Not mCollapsible Then
        iconLeft = ICON_EDGE_PAD
        titleLeft = TITLE_LEFT_NO_ARROW
        titleRight = barWidth - RIGHT_PAD
    ElseIf mIconPosition = POS_RIGHT Then
        iconLeft = MaxSingle(0, barWidth - ICON_EDGE_PAD - iconWidth)
        titleLeft = TITLE_LEFT_NO_ARROW
        titleRight = iconLeft - TITLE_GAP
    Else
        iconLeft = ICON_EDGE_PAD
        titleLeft = ICON_EDGE_PAD + iconWidth + TITLE_GAP
        titleRight = barWidth - RIGHT_PAD
    End If

    mArrow.Width = iconWidth
    mArrow.Left = mOffX + iconLeft
    mArrow.Top = mOffY + MaxSingle(0, (h - mArrow.Height) / 2)

    mTitle.Left = mOffX + titleLeft
    mTitle.Width = MaxSingle(10, titleRight - titleLeft)
    mTitle.Top = mOffY + MaxSingle(0, (h - mTitle.Height) / 2)

    ' Heading area handed to a style (absolute host coordinates)
    mRegionLeft = mOffX + titleLeft
    mRegionRight = mOffX + MaxSingle(titleLeft + 10, titleRight)
    RenderStyle
End Sub

' Lets the active style (if any) decorate the header, then keeps the icon
' and heading in front of any decoration.
Private Sub RenderStyle()
    If mStyle Is Nothing Then Exit Sub
    If mBar Is Nothing Or mTitle Is Nothing Or mArrow Is Nothing Then Exit Sub

    Trace "RenderStyle"
    mStyle.Render mHost, mPrefix & STYLE_TAG, mTitle, _
                  mBar.Left, mBar.Top, mBar.Width, mBar.Height, _
                  mRegionLeft, mRegionRight

    HostCtl(mPrefix & "Arrow").ZOrder fmZOrderFront
    HostCtl(mPrefix & "Title").ZOrder fmZOrderFront
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
    RenderStyle             ' e.g. Divider re-fits lines to the new text width
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
' PRIVATE HELPERS - CONTENT DETECTION
'==============================================================================

' Frame mode : every control in the Frame except the header.
' Label mode : sibling controls (same parent as the Label) lying fully inside
'              the Label's rectangle, except the Label itself and the header.
Private Function IsContentControl(ByVal ctl As MSForms.Control) As Boolean
    If IsHeaderControl(ctl.Name) Then Exit Function

    If Not mIsLabelMode Then
        IsContentControl = True
        Exit Function
    End If

    If StrComp(ctl.Name, mContainer.Name, vbTextCompare) = 0 Then Exit Function
    If Not IsDirectChildOfHost(ctl) Then Exit Function
    IsContentControl = IsInsidePanel(ctl)
End Function

' UserForm.Controls also lists controls nested in Frames/MultiPages, whose
' Left/Top are relative to THEIR parent - exclude them from the geometry test.
' (Hiding their parent Frame hides them anyway.)
Private Function IsDirectChildOfHost(ByVal ctl As Object) As Boolean
    On Error GoTo NotChild
    IsDirectChildOfHost = (StrComp(ctl.Parent.Name, mHostName, vbTextCompare) = 0)
    Exit Function
NotChild:
    IsDirectChildOfHost = False
End Function

Private Function IsInsidePanel(ByVal ctl As MSForms.Control) As Boolean
    IsInsidePanel = (ctl.Left >= mLeft - INSIDE_TOLERANCE) And _
                    (ctl.Top >= mTop - INSIDE_TOLERANCE) And _
                    (ctl.Left + ctl.Width <= mLeft + mWidth + INSIDE_TOLERANCE) And _
                    (ctl.Top + ctl.Height <= mTop + mHeight + INSIDE_TOLERANCE)
End Function

'==============================================================================
' PRIVATE HELPERS - HOUSEKEEPING
'==============================================================================

Private Sub Trace(ByVal stepName As String)
    mStep = stepName
    If TRACE_STEPS Then Debug.Print "CCollapsibleFrame | " & stepName
End Sub

' Extender view of a control in the host (Visible, ZOrder live here).
Private Function HostCtl(ByVal ctlName As String) As MSForms.Control
    Set HostCtl = mHost.Controls(ctlName)
End Function

Private Function SafeName(ByVal obj As Object) As String
    On Error Resume Next
    SafeName = obj.Name
    On Error GoTo 0
End Function

Private Function IsHeaderControl(ByVal ctlName As String) As Boolean
    IsHeaderControl = (StrComp(Left$(ctlName, Len(mPrefix)), mPrefix, vbTextCompare) = 0)
End Function

Private Sub RemoveHeaderControls()
    Set mBar = Nothing
    Set mArrow = Nothing
    Set mTitle = Nothing
    If mHost Is Nothing Then Exit Sub
    RemoveIfExists mPrefix & "Title"
    RemoveIfExists mPrefix & "Arrow"
    RemoveIfExists mPrefix & "Bar"
End Sub

' Removes style controls left on this container by an earlier instance
' (names start with "ccf_<container>_sty_"). Collected first, then removed,
' so the Controls collection is not modified while it is being enumerated.
Private Sub RemoveStyleLeftovers()
    Dim ctl As MSForms.Control
    Dim names As Collection
    Dim nm As Variant
    Dim tag As String

    If mHost Is Nothing Then Exit Sub
    tag = mPrefix & STYLE_TAG
    Set names = New Collection
    For Each ctl In mHost.Controls
        If StrComp(Left$(ctl.Name, Len(tag)), tag, vbTextCompare) = 0 Then names.Add ctl.Name
    Next ctl
    For Each nm In names
        mHost.Controls.Remove CStr(nm)
    Next nm
End Sub

' Removes only this class's own header controls (matched by exact name).
Private Sub RemoveIfExists(ByVal ctlName As String)
    Dim ctl As MSForms.Control
    For Each ctl In mHost.Controls
        If StrComp(ctl.Name, ctlName, vbTextCompare) = 0 Then
            mHost.Controls.Remove ctlName
            Exit For
        End If
    Next ctl
End Sub

' Parent's current colour (Frame mode, transparent or collapsed).
' A MultiPage Page has no BackColor, so it falls back to the system face.
Private Function ParentBackColor() As Long
    On Error GoTo Fallback
    ParentBackColor = mContainer.Parent.BackColor
    Exit Function
Fallback:
    ParentBackColor = CLR_FALLBACK_FORM
End Function

' Guarantee: the container's footprint never changes.
Private Sub EnforceGeometry()
    If mContainer Is Nothing Then Exit Sub
    With mContainer
        If .Left <> mLeft Then .Left = mLeft
        If .Top <> mTop Then .Top = mTop
        If .Width <> mWidth Then .Width = mWidth
        If .Height <> mHeight Then .Height = mHeight
    End With
End Sub

Private Function MaxSingle(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then MaxSingle = a Else MaxSingle = b
End Function
