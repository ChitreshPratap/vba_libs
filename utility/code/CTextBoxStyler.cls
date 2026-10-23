VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTextBoxStyler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Class   : CTextBoxStyler  (v3)
' Purpose : Apply a clean, modern look to an EXISTING MSForms TextBox, with
'           independent left / top / right / bottom borders.
'
' Guarantees
'   - Never writes the TextBox's Height, Width, Top or Left.
'   - Sets AutoSize = False first, so font changes can never resize it.
'   - Checks the TextBox's geometry before and after every styling pass and
'     reports an error if it ever differs.
'
' How per-side borders work (MSForms has no native per-side border):
'   - Up to four thin Labels are created in the TextBox's own container
'     (UserForm, Frame or MultiPage Page) and placed just OUTSIDE its edges,
'     so they never cover the editable text area.
'   - Each helper is named  tbsBdr<L|T|R|B>_<TextBoxName>  and tagged
'     "CTextBoxStyler.Border". Existing helpers are reused (no duplicates);
'     a side switched off has its helper removed.
'   - Helpers are runtime controls: they disappear when the form unloads.
'
' Border geometry (all values in points, the unit of MSForms Left/Top/Width)
'   - <Side>BorderWidth     : thickness, perpendicular to the edge.
'   - <Side>BorderLength    : extent along the edge. 0 = full edge (default).
'                             Longer than the edge -> clipped to the edge.
'   - <Side>BorderStartFrom : where the segment is anchored on its edge.
'                             Top/Bottom : "Left" | "Right" | "Center"
'                             Left/Right : "Top"  | "Bottom" | "Center"
'   - A horizontal border also covers a corner when the vertical border on
'     that side is visible and reaches that corner, so corners join cleanly.
'
' Settings order
'   - Set properties before ApplyStyle -> used when ApplyStyle runs.
'   - Set properties after ApplyStyle  -> applied immediately to that same
'     TextBox (the class keeps a reference to it).
'   - ApplyStyle on another TextBox moves the reference to that box.
'   - Detach / ResetDefaults release the reference without changing the box.
'
' Requires: Microsoft Forms 2.0 Object Library (added automatically to any
'           project that contains a UserForm).
'===============================================================================
Option Explicit

Public Enum tbsBorderSide
    tbsLeft = 0
    tbsTop = 1
    tbsRight = 2
    tbsBottom = 3
End Enum

Private Const HELPER_TAG       As String = "CTextBoxStyler.Border"
Private Const HELPER_PREFIX    As String = "tbsBdr"
Private Const SIDE_LETTERS     As String = "LTRB"
Private Const MIN_BORDER_WIDTH As Single = 0.75
Private Const MAX_BORDER_WIDTH As Single = 20
Private Const MIN_FONT_SIZE    As Single = 6
Private Const MAX_FONT_SIZE    As Single = 72
Private Const EDGE_TOLERANCE   As Single = 0.01

' Border anchors (stored per side; shown to callers as Left/Top etc.)
Private Const ANCHOR_START  As Long = 0   ' Left for Top/Bottom, Top for Left/Right
Private Const ANCHOR_END    As Long = 1   ' Right for Top/Bottom, Bottom for Left/Right
Private Const ANCHOR_CENTER As Long = 2

Private Const ERR_FONT_NAME     As Long = vbObjectError + 513
Private Const ERR_FONT_SIZE     As Long = vbObjectError + 514
Private Const ERR_BORDER_WIDTH  As Long = vbObjectError + 515
Private Const ERR_BORDER_SIDE   As Long = vbObjectError + 516
Private Const ERR_NAME_CONFLICT As Long = vbObjectError + 517
Private Const ERR_NO_CONTAINER  As Long = vbObjectError + 518
Private Const ERR_BORDER_LENGTH As Long = vbObjectError + 519
Private Const ERR_BORDER_ANCHOR As Long = vbObjectError + 520

'--- Attached TextBox --------------------------------------------------------
Private m_Target As MSForms.TextBox   ' typed reference (styling)
Private m_Ctl    As Object            ' same object, late bound (Name/Parent/geometry)

'--- Visual settings ----------------------------------------------------------
Private m_BackColor         As Long
Private m_ReadOnlyBackColor As Long
Private m_ForeColor         As Long
Private m_BorderColor       As Long
Private m_FontName          As String
Private m_FontSize          As Single
Private m_FontBold          As Boolean
Private m_FontItalic        As Boolean
Private m_TextAlign         As fmTextAlign
Private m_BorderStyle       As fmBorderStyle
Private m_SpecialEffect     As fmSpecialEffect
Private m_Transparent       As Boolean

'--- Per-side borders (index = tbsBorderSide) ----------------------------------
Private m_BVisible(0 To 3) As Boolean
Private m_BColor(0 To 3)   As Long
Private m_BWidth(0 To 3)   As Single    ' thickness
Private m_BLength(0 To 3)  As Single    ' 0 = full edge
Private m_BAnchor(0 To 3)  As Long      ' ANCHOR_START / ANCHOR_END / ANCHOR_CENTER

'--- Optional state settings: Empty = leave the TextBox's current state alone -
Private m_PasswordChar As Variant
Private m_Enabled      As Variant
Private m_Locked       As Variant

'===============================================================================
' Lifecycle
'===============================================================================
Private Sub Class_Initialize()
    ResetDefaults
End Sub

Private Sub Class_Terminate()
    ' Styling and border helpers intentionally stay on the form after this
    ' object is released, so a local variable in UserForm_Initialize works.
    Set m_Target = Nothing
    Set m_Ctl = Nothing
End Sub

'-------------------------------------------------------------------------------
' Restores the built-in defaults and detaches from the current TextBox.
' The TextBox keeps its current look; call ApplyStyle again to restyle it.
'-------------------------------------------------------------------------------
Public Sub ResetDefaults()
    Dim i As Long

    Detach

    m_BackColor = RGB(255, 255, 255)          ' white
    m_ReadOnlyBackColor = RGB(243, 244, 246)  ' soft grey for Locked boxes
    m_ForeColor = RGB(33, 37, 41)             ' near-black text
    m_BorderColor = RGB(206, 212, 218)        ' light grey hairline border
    m_FontName = "Segoe UI"
    m_FontSize = 10
    m_FontBold = False
    m_FontItalic = False
    m_TextAlign = fmTextAlignLeft
    m_BorderStyle = fmBorderStyleSingle
    m_SpecialEffect = fmSpecialEffectFlat
    m_Transparent = False
    m_PasswordChar = Empty
    m_Enabled = Empty
    m_Locked = Empty

    For i = tbsLeft To tbsBottom
        m_BVisible(i) = False
        m_BColor(i) = m_BorderColor
        m_BWidth(i) = 1
        m_BLength(i) = 0
        m_BAnchor(i) = ANCHOR_START
    Next i
End Sub

'-------------------------------------------------------------------------------
' Stops live updates to the attached TextBox. Its styling and borders remain.
'-------------------------------------------------------------------------------
Public Sub Detach()
    Set m_Target = Nothing
    Set m_Ctl = Nothing
End Sub

'===============================================================================
' Main entry point
'-------------------------------------------------------------------------------
' TargetTextBox : any existing MSForms TextBox (e.g. Me.TextBox1).
' InitialText   : optional. When supplied, it is written to the TextBox's
'                 .Value property (the text inside the box). It does NOT set
'                 ControlTipText or any Label. When omitted, the existing
'                 .Value is left exactly as it is.
' Returns       : True on success, False if the object was invalid or styling
'                 failed (a clear message is shown).
'===============================================================================
Public Function ApplyStyle(ByVal TargetTextBox As Object, _
                           Optional ByVal InitialText As Variant) As Boolean

    If TargetTextBox Is Nothing Then
        ReportError "No TextBox was supplied (the object is Nothing)."
        Exit Function
    End If

    If Not TypeOf TargetTextBox Is MSForms.TextBox Then
        ReportError "Expected an MSForms TextBox, but received an object of type '" & _
                    TypeName(TargetTextBox) & "'."
        Exit Function
    End If

    Set m_Target = TargetTextBox
    Set m_Ctl = TargetTextBox

    If Not StyleTarget() Then Exit Function

    If Not IsMissing(InitialText) Then
        On Error GoTo TextFailed
        If IsNull(InitialText) Then
            m_Target.Value = vbNullString
        Else
            m_Target.Value = CStr(InitialText)
        End If
        On Error GoTo 0
    End If

    ApplyStyle = True
    Exit Function

TextFailed:
    ReportError "Could not set the text of '" & SafeName(m_Ctl) & "': " & Err.Description
End Function

'-------------------------------------------------------------------------------
' Re-applies all settings to the attached TextBox. Call this if you move or
' hide/show the TextBox at runtime, so the borders follow it.
'-------------------------------------------------------------------------------
Public Function Refresh() As Boolean
    If m_Target Is Nothing Then
        ReportError "Refresh called while no TextBox is attached."
        Exit Function
    End If
    Refresh = StyleTarget()
End Function

'-------------------------------------------------------------------------------
' Turns all four custom borders off and removes their helper Labels from the
' attached TextBox. The native border settings take effect again.
'-------------------------------------------------------------------------------
Public Sub RemoveBorders()
    Dim i As Long
    For i = tbsLeft To tbsBottom
        m_BVisible(i) = False
    Next i
    ApplyIfAttached
End Sub

'-------------------------------------------------------------------------------
' Convenience setters - one call, one restyle.
'-------------------------------------------------------------------------------
Public Sub SetBorder(ByVal Side As tbsBorderSide, ByVal Visible As Boolean, _
                     Optional ByVal Color As Variant, Optional ByVal Width As Variant, _
                     Optional ByVal Length As Variant, Optional ByVal StartFrom As Variant)
    Dim anchor As Long

    ' Validate everything first, so a bad argument changes nothing.
    ValidateSide Side
    If Not IsMissing(Width) Then ValidateBorderWidth CSng(Width)
    If Not IsMissing(Length) Then ValidateBorderLength CSng(Length)
    If Not IsMissing(StartFrom) Then anchor = ParseAnchor(Side, CStr(StartFrom))

    m_BVisible(Side) = Visible
    If Not IsMissing(Color) Then m_BColor(Side) = CLng(Color)
    If Not IsMissing(Width) Then m_BWidth(Side) = CSng(Width)
    If Not IsMissing(Length) Then m_BLength(Side) = CSng(Length)
    If Not IsMissing(StartFrom) Then m_BAnchor(Side) = anchor
    ApplyIfAttached
End Sub

Public Sub SetAllBorders(ByVal Visible As Boolean, _
                         Optional ByVal Color As Variant, Optional ByVal Width As Variant)
    Dim i As Long
    If Not IsMissing(Width) Then ValidateBorderWidth CSng(Width)

    For i = tbsLeft To tbsBottom
        m_BVisible(i) = Visible
        If Not IsMissing(Color) Then m_BColor(i) = CLng(Color)
        If Not IsMissing(Width) Then m_BWidth(i) = CSng(Width)
    Next i
    ApplyIfAttached
End Sub

'===============================================================================
' Core styling
'===============================================================================
Private Sub ApplyIfAttached()
    If Not m_Target Is Nothing Then StyleTarget
End Sub

'-------------------------------------------------------------------------------
' One full styling pass: visuals, borders, then a geometry check.
'-------------------------------------------------------------------------------
Private Function StyleTarget() As Boolean
    Dim l0 As Single, t0 As Single, w0 As Single, h0 As Single

    On Error GoTo StyleFailed

    l0 = m_Ctl.Left: t0 = m_Ctl.Top: w0 = m_Ctl.Width: h0 = m_Ctl.Height

    ApplyVisuals m_Target
    UpdateBorders

    If m_Ctl.Left <> l0 Or m_Ctl.Top <> t0 Or _
       m_Ctl.Width <> w0 Or m_Ctl.Height <> h0 Then
        ReportError "Geometry check failed on '" & SafeName(m_Ctl) & _
                    "': its size or position changed during styling."
        Exit Function
    End If

    StyleTarget = True
    Exit Function

StyleFailed:
    ReportError "Could not style TextBox '" & SafeName(m_Ctl) & "': " & Err.Description
End Function

'-------------------------------------------------------------------------------
' Visual properties of the TextBox itself. No geometry property is written.
'-------------------------------------------------------------------------------
Private Sub ApplyVisuals(ByVal tb As MSForms.TextBox)

    ' Guard: with AutoSize off, font changes cannot alter the control's size.
    tb.AutoSize = False

    ' Native border. Custom per-side borders replace it entirely.
    ' MSForms rule: a Single border only renders with a Flat effect, and a
    ' non-flat effect forces BorderStyle to None - so set them in order.
    If AnyBorderVisible() Then
        tb.SpecialEffect = fmSpecialEffectFlat
        tb.BorderStyle = fmBorderStyleNone
    ElseIf m_BorderStyle = fmBorderStyleSingle Then
        tb.SpecialEffect = fmSpecialEffectFlat
        tb.BorderStyle = fmBorderStyleSingle
        tb.BorderColor = m_BorderColor
    Else
        tb.BorderStyle = fmBorderStyleNone
        tb.SpecialEffect = m_SpecialEffect
    End If

    ' Text
    tb.ForeColor = m_ForeColor
    tb.TextAlign = m_TextAlign
    tb.SelectionMargin = False      ' removes the old-style left click margin

    With tb.Font
        .Name = m_FontName
        .Size = m_FontSize
        .Bold = m_FontBold
        .Italic = m_FontItalic
    End With

    ' Optional states - only applied when explicitly set on this instance
    If Not IsEmpty(m_PasswordChar) Then tb.PasswordChar = m_PasswordChar
    If Not IsEmpty(m_Enabled) Then tb.Enabled = m_Enabled
    If Not IsEmpty(m_Locked) Then tb.Locked = m_Locked

    ' Background
    If m_Transparent Then
        tb.BackStyle = fmBackStyleTransparent
        ' BackColor is not drawn while transparent. Matching it to the
        ' container means any opaque repaint (e.g. while editing) blends in.
        tb.BackColor = ContainerBackColor(m_BackColor)
    Else
        tb.BackStyle = fmBackStyleOpaque
        If tb.Locked Then
            tb.BackColor = m_ReadOnlyBackColor
        Else
            tb.BackColor = m_BackColor
        End If
    End If
End Sub

'-------------------------------------------------------------------------------
' Creates, positions, recolours or removes the four helper Labels.
' All coordinates come from the TextBox's own Left/Top/Width/Height in its
' container, so borders follow the TextBox wherever it is placed.
'
' Labels sit just outside the TextBox edges:
'   Left   : x = TextBox.Left - thickness,  along the TextBox height
'   Right  : x = TextBox.Left + Width,      along the TextBox height
'   Top    : y = TextBox.Top - thickness,   along the TextBox width
'   Bottom : y = TextBox.Top + Height,      along the TextBox width
' A Top/Bottom edge is widened by a vertical border's thickness when that
' vertical border is visible and reaches the same corner.
'-------------------------------------------------------------------------------
Private Sub UpdateBorders()
    Dim x As Single, y As Single, w As Single, h As Single
    Dim lw As Single, rw As Single, tw As Single, bw As Single
    Dim lStart As Single, lLen As Single
    Dim rStart As Single, rLen As Single
    Dim extL As Single, extR As Single
    Dim segStart As Single, segLen As Single

    x = m_Ctl.Left: y = m_Ctl.Top: w = m_Ctl.Width: h = m_Ctl.Height

    '--- Vertical borders: run along the TextBox height ------------------------
    If m_BVisible(tbsLeft) Then
        lw = m_BWidth(tbsLeft)
        PlaceSegment tbsLeft, y, h, lStart, lLen
        ShowBorder tbsLeft, x - lw, lStart, lw, lLen
    Else
        RemoveBorder tbsLeft
    End If

    If m_BVisible(tbsRight) Then
        rw = m_BWidth(tbsRight)
        PlaceSegment tbsRight, y, h, rStart, rLen
        ShowBorder tbsRight, x + w, rStart, rw, rLen
    Else
        RemoveBorder tbsRight
    End If

    '--- Top border ------------------------------------------------------------
    If m_BVisible(tbsTop) Then
        tw = m_BWidth(tbsTop)
        extL = 0: extR = 0
        If lw > 0 And lStart <= y + EDGE_TOLERANCE Then extL = lw
        If rw > 0 And rStart <= y + EDGE_TOLERANCE Then extR = rw
        PlaceSegment tbsTop, x - extL, w + extL + extR, segStart, segLen
        ShowBorder tbsTop, segStart, y - tw, segLen, tw
    Else
        RemoveBorder tbsTop
    End If

    '--- Bottom border ---------------------------------------------------------
    If m_BVisible(tbsBottom) Then
        bw = m_BWidth(tbsBottom)
        extL = 0: extR = 0
        If lw > 0 And lStart + lLen >= y + h - EDGE_TOLERANCE Then extL = lw
        If rw > 0 And rStart + rLen >= y + h - EDGE_TOLERANCE Then extR = rw
        PlaceSegment tbsBottom, x - extL, w + extL + extR, segStart, segLen
        ShowBorder tbsBottom, segStart, y + h, segLen, bw
    Else
        RemoveBorder tbsBottom
    End If
End Sub

'-------------------------------------------------------------------------------
' Works out where one border segment lies along its edge.
'   EdgeStart / EdgeLength : the available edge (points)
'   SegStart / SegLength   : returned position and rendered length
' Length 0 (or longer than the edge) renders the full edge.
'-------------------------------------------------------------------------------
Private Sub PlaceSegment(ByVal Side As Long, ByVal EdgeStart As Single, _
                         ByVal EdgeLength As Single, _
                         ByRef SegStart As Single, ByRef SegLength As Single)
    If EdgeLength < 0 Then EdgeLength = 0

    SegLength = m_BLength(Side)
    If SegLength <= 0 Or SegLength > EdgeLength Then SegLength = EdgeLength

    Select Case m_BAnchor(Side)
        Case ANCHOR_END
            SegStart = EdgeStart + EdgeLength - SegLength
        Case ANCHOR_CENTER
            SegStart = EdgeStart + (EdgeLength - SegLength) / 2
        Case Else
            SegStart = EdgeStart
    End Select
End Sub

Private Sub ShowBorder(ByVal side As Long, ByVal x As Single, ByVal y As Single, _
                       ByVal w As Single, ByVal h As Single)
    Dim lbl As MSForms.Label
    Set lbl = GetBorderLabel(side, True)
    lbl.Move x, y, w, h              ' moves the helper Label, never the TextBox
    lbl.BackColor = m_BColor(side)
    lbl.Visible = m_Ctl.Visible
End Sub

Private Sub RemoveBorder(ByVal side As Long)
    If GetBorderLabel(side, False) Is Nothing Then Exit Sub
    m_Ctl.Parent.Controls.Remove BorderName(side)
End Sub

'-------------------------------------------------------------------------------
' Returns this side's helper Label, creating it if requested.
' A control with the helper's name that this class did not create is never
' touched: creating raises a naming-conflict error, removing ignores it.
'-------------------------------------------------------------------------------
Private Function GetBorderLabel(ByVal side As Long, _
                                ByVal CreateIfMissing As Boolean) As MSForms.Label
    Dim nm As String
    Dim existing As Object
    Dim lbl As MSForms.Label

    nm = BorderName(side)
    Set existing = FindSibling(nm)

    If Not existing Is Nothing Then
        If TypeOf existing Is MSForms.Label Then
            If existing.Tag = HELPER_TAG Then
                Set GetBorderLabel = existing
                Exit Function
            End If
        End If
        If CreateIfMissing Then
            Err.Raise ERR_NAME_CONFLICT, "CTextBoxStyler", _
                "A control named '" & nm & "' already exists and was not created by CTextBoxStyler."
        End If
        Exit Function
    End If

    If Not CreateIfMissing Then Exit Function

    If Not IsSupportedContainer() Then
        Err.Raise ERR_NO_CONTAINER, "CTextBoxStyler", _
            "Per-side borders need the TextBox to sit on a UserForm, Frame or Page; " & _
            "its container is '" & TypeName(m_Ctl.Parent) & "'."
    End If

    Set lbl = m_Ctl.Parent.Controls.Add("Forms.Label.1", nm, True)

    With lbl
        .Tag = HELPER_TAG
        .Caption = vbNullString
        .AutoSize = False
        .WordWrap = False
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    Set GetBorderLabel = lbl
End Function

'===============================================================================
' Helpers
'===============================================================================
Private Function BorderName(ByVal side As Long) As String
    BorderName = HELPER_PREFIX & Mid$(SIDE_LETTERS, side + 1, 1) & "_" & m_Ctl.Name
End Function

'-------------------------------------------------------------------------------
' Looks a control up by name WITHOUT raising an error when it is missing,
' so it also works with "Break on All Errors" switched on in the VBA editor.
'-------------------------------------------------------------------------------
Private Function FindSibling(ByVal ControlName As String) As Object
    Dim ctl As Object
    If Not IsSupportedContainer() Then Exit Function
    For Each ctl In m_Ctl.Parent.Controls
        If StrComp(ctl.Name, ControlName, vbTextCompare) = 0 Then
            Set FindSibling = ctl
            Exit Function
        End If
    Next ctl
End Function

'-------------------------------------------------------------------------------
' True when the TextBox sits on a UserForm, Frame or MultiPage Page - the
' containers that support Controls.Add / Controls.Remove at runtime.
'-------------------------------------------------------------------------------
Private Function IsSupportedContainer() As Boolean
    Dim p As Object
    Set p = m_Ctl.Parent
    If p Is Nothing Then Exit Function
    IsSupportedContainer = (TypeOf p Is MSForms.UserForm) Or _
                           (TypeOf p Is MSForms.Frame) Or _
                           (TypeOf p Is MSForms.Page)
End Function

Private Function AnyBorderVisible() As Boolean
    Dim i As Long
    For i = tbsLeft To tbsBottom
        If m_BVisible(i) Then
            AnyBorderVisible = True
            Exit Function
        End If
    Next i
End Function

' UserForm and Frame have BackColor; a MultiPage Page does not, so use the
' fallback there instead of relying on error suppression.
Private Function ContainerBackColor(ByVal Fallback As Long) As Long
    Dim p As Object
    ContainerBackColor = Fallback
    Set p = m_Ctl.Parent
    If p Is Nothing Then Exit Function
    If (TypeOf p Is MSForms.UserForm) Or (TypeOf p Is MSForms.Frame) Then
        ContainerBackColor = p.BackColor
    End If
End Function

Private Sub ValidateBorderWidth(ByVal Value As Single)
    If Value < MIN_BORDER_WIDTH Or Value > MAX_BORDER_WIDTH Then
        Err.Raise ERR_BORDER_WIDTH, "CTextBoxStyler", _
            "Border width must be between " & MIN_BORDER_WIDTH & " and " & _
            MAX_BORDER_WIDTH & " points."
    End If
End Sub

Private Sub ValidateBorderLength(ByVal Value As Single)
    If Value < 0 Then
        Err.Raise ERR_BORDER_LENGTH, "CTextBoxStyler", _
            "Border length cannot be negative. Use 0 for the full edge."
    End If
End Sub

'-------------------------------------------------------------------------------
' Converts "Left" / "Right" / "Top" / "Bottom" / "Center" to an anchor code,
' checking the value is valid for that side. Case and spaces are ignored;
' "Centre" is also accepted.
'-------------------------------------------------------------------------------
Private Function ParseAnchor(ByVal Side As Long, ByVal Value As String) As Long
    Dim t As String
    t = LCase$(Trim$(Value))

    If t = "center" Or t = "centre" Then
        ParseAnchor = ANCHOR_CENTER
        Exit Function
    End If

    If Side = tbsTop Or Side = tbsBottom Then
        Select Case t
            Case "left":  ParseAnchor = ANCHOR_START: Exit Function
            Case "right": ParseAnchor = ANCHOR_END:   Exit Function
        End Select
    Else
        Select Case t
            Case "top":    ParseAnchor = ANCHOR_START: Exit Function
            Case "bottom": ParseAnchor = ANCHOR_END:   Exit Function
        End Select
    End If

    Err.Raise ERR_BORDER_ANCHOR, "CTextBoxStyler", _
        "Invalid StartFrom value '" & Value & "' for the " & SideName(Side) & _
        " border. Use " & AllowedAnchors(Side) & "."
End Function

Private Function AnchorName(ByVal Side As Long) As String
    Select Case m_BAnchor(Side)
        Case ANCHOR_CENTER
            AnchorName = "Center"
        Case ANCHOR_END
            If Side = tbsTop Or Side = tbsBottom Then AnchorName = "Right" Else AnchorName = "Bottom"
        Case Else
            If Side = tbsTop Or Side = tbsBottom Then AnchorName = "Left" Else AnchorName = "Top"
    End Select
End Function

Private Function AllowedAnchors(ByVal Side As Long) As String
    If Side = tbsTop Or Side = tbsBottom Then
        AllowedAnchors = """Left"", ""Right"" or ""Center"""
    Else
        AllowedAnchors = """Top"", ""Bottom"" or ""Center"""
    End If
End Function

Private Function SideName(ByVal Side As Long) As String
    SideName = Choose(Side + 1, "Left", "Top", "Right", "Bottom")
End Function

Private Sub LetBorderLength(ByVal Side As Long, ByVal Value As Single)
    ValidateBorderLength Value
    m_BLength(Side) = Value
    ApplyIfAttached
End Sub

Private Sub LetBorderStart(ByVal Side As Long, ByVal Value As String)
    m_BAnchor(Side) = ParseAnchor(Side, Value)
    ApplyIfAttached
End Sub

Private Sub ValidateSide(ByVal Side As Long)
    If Side < tbsLeft Or Side > tbsBottom Then
        Err.Raise ERR_BORDER_SIDE, "CTextBoxStyler", _
            "Border side must be tbsLeft, tbsTop, tbsRight or tbsBottom."
    End If
End Sub

Private Sub LetBorderVisible(ByVal Side As Long, ByVal Value As Boolean)
    m_BVisible(Side) = Value
    ApplyIfAttached
End Sub

Private Sub LetBorderColor(ByVal Side As Long, ByVal Value As Long)
    m_BColor(Side) = Value
    ApplyIfAttached
End Sub

Private Sub LetBorderWidth(ByVal Side As Long, ByVal Value As Single)
    ValidateBorderWidth Value
    m_BWidth(Side) = Value
    ApplyIfAttached
End Sub

Private Sub ReportError(ByVal Message As String)
    MsgBox Message, vbExclamation, "CTextBoxStyler"
End Sub

Private Function SafeName(ByVal obj As Object) As String
    On Error Resume Next
    SafeName = obj.Name
    If Len(SafeName) = 0 Then SafeName = TypeName(obj)
End Function

'===============================================================================
' Public properties - every Let applies immediately if a TextBox is attached
'===============================================================================
Public Property Get Target() As MSForms.TextBox
    Set Target = m_Target
End Property

Public Property Get IsAttached() As Boolean
    IsAttached = Not (m_Target Is Nothing)
End Property

'--- Background ----------------------------------------------------------------
Public Property Get BackColor() As Long
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal Value As Long)
    m_BackColor = Value
    ApplyIfAttached
End Property

Public Property Get ReadOnlyBackColor() As Long
    ReadOnlyBackColor = m_ReadOnlyBackColor
End Property
Public Property Let ReadOnlyBackColor(ByVal Value As Long)
    m_ReadOnlyBackColor = Value
    ApplyIfAttached
End Property

' Native MSForms transparency (BackStyle = fmBackStyleTransparent).
Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property
Public Property Let Transparent(ByVal Value As Boolean)
    m_Transparent = Value
    ApplyIfAttached
End Property

'--- Font ----------------------------------------------------------------------
Public Property Get FontColor() As Long
    FontColor = m_ForeColor
End Property
Public Property Let FontColor(ByVal Value As Long)
    m_ForeColor = Value
    ApplyIfAttached
End Property

' Kept for compatibility - same setting as FontColor.
Public Property Get ForeColor() As Long
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal Value As Long)
    m_ForeColor = Value
    ApplyIfAttached
End Property

Public Property Get FontName() As String
    FontName = m_FontName
End Property
Public Property Let FontName(ByVal Value As String)
    If Len(Trim$(Value)) = 0 Then
        Err.Raise ERR_FONT_NAME, "CTextBoxStyler", "FontName cannot be blank."
    End If
    m_FontName = Trim$(Value)
    ApplyIfAttached
End Property

Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property
Public Property Let FontSize(ByVal Value As Single)
    If Value < MIN_FONT_SIZE Or Value > MAX_FONT_SIZE Then
        Err.Raise ERR_FONT_SIZE, "CTextBoxStyler", _
            "FontSize must be between " & MIN_FONT_SIZE & " and " & MAX_FONT_SIZE & "."
    End If
    m_FontSize = Value
    ApplyIfAttached
End Property

Public Property Get FontBold() As Boolean
    FontBold = m_FontBold
End Property
Public Property Let FontBold(ByVal Value As Boolean)
    m_FontBold = Value
    ApplyIfAttached
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = m_FontItalic
End Property
Public Property Let FontItalic(ByVal Value As Boolean)
    m_FontItalic = Value
    ApplyIfAttached
End Property

Public Property Get TextAlign() As fmTextAlign
    TextAlign = m_TextAlign
End Property
Public Property Let TextAlign(ByVal Value As fmTextAlign)
    m_TextAlign = Value
    ApplyIfAttached
End Property

'--- Native (whole-box) border - used only when no per-side border is on -------
Public Property Get BorderColor() As Long
    BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal Value As Long)
    m_BorderColor = Value
    ApplyIfAttached
End Property

Public Property Get BorderStyle() As fmBorderStyle
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal Value As fmBorderStyle)
    m_BorderStyle = Value
    ApplyIfAttached
End Property

' Only used when BorderStyle = fmBorderStyleNone (MSForms limitation).
Public Property Get SpecialEffect() As fmSpecialEffect
    SpecialEffect = m_SpecialEffect
End Property
Public Property Let SpecialEffect(ByVal Value As fmSpecialEffect)
    m_SpecialEffect = Value
    ApplyIfAttached
End Property

'--- Left border -----------------------------------------------------------------
Public Property Get LeftBorderVisible() As Boolean
    LeftBorderVisible = m_BVisible(tbsLeft)
End Property
Public Property Let LeftBorderVisible(ByVal Value As Boolean)
    LetBorderVisible tbsLeft, Value
End Property

Public Property Get LeftBorderColor() As Long
    LeftBorderColor = m_BColor(tbsLeft)
End Property
Public Property Let LeftBorderColor(ByVal Value As Long)
    LetBorderColor tbsLeft, Value
End Property

Public Property Get LeftBorderWidth() As Single
    LeftBorderWidth = m_BWidth(tbsLeft)
End Property
Public Property Let LeftBorderWidth(ByVal Value As Single)
    LetBorderWidth tbsLeft, Value
End Property

' Length along the edge in points; 0 = full edge, longer = clipped to edge.
Public Property Get LeftBorderLength() As Single
    LeftBorderLength = m_BLength(tbsLeft)
End Property
Public Property Let LeftBorderLength(ByVal Value As Single)
    LetBorderLength tbsLeft, Value
End Property

' Anchor on the edge: "Top" | "Bottom" | "Center"
Public Property Get LeftBorderStartFrom() As String
    LeftBorderStartFrom = AnchorName(tbsLeft)
End Property
Public Property Let LeftBorderStartFrom(ByVal Value As String)
    LetBorderStart tbsLeft, Value
End Property

'--- Top border ------------------------------------------------------------------
Public Property Get TopBorderVisible() As Boolean
    TopBorderVisible = m_BVisible(tbsTop)
End Property
Public Property Let TopBorderVisible(ByVal Value As Boolean)
    LetBorderVisible tbsTop, Value
End Property

Public Property Get TopBorderColor() As Long
    TopBorderColor = m_BColor(tbsTop)
End Property
Public Property Let TopBorderColor(ByVal Value As Long)
    LetBorderColor tbsTop, Value
End Property

Public Property Get TopBorderWidth() As Single
    TopBorderWidth = m_BWidth(tbsTop)
End Property
Public Property Let TopBorderWidth(ByVal Value As Single)
    LetBorderWidth tbsTop, Value
End Property

' Length along the edge in points; 0 = full edge, longer = clipped to edge.
Public Property Get TopBorderLength() As Single
    TopBorderLength = m_BLength(tbsTop)
End Property
Public Property Let TopBorderLength(ByVal Value As Single)
    LetBorderLength tbsTop, Value
End Property

' Anchor on the edge: "Left" | "Right" | "Center"
Public Property Get TopBorderStartFrom() As String
    TopBorderStartFrom = AnchorName(tbsTop)
End Property
Public Property Let TopBorderStartFrom(ByVal Value As String)
    LetBorderStart tbsTop, Value
End Property

'--- Right border ----------------------------------------------------------------
Public Property Get RightBorderVisible() As Boolean
    RightBorderVisible = m_BVisible(tbsRight)
End Property
Public Property Let RightBorderVisible(ByVal Value As Boolean)
    LetBorderVisible tbsRight, Value
End Property

Public Property Get RightBorderColor() As Long
    RightBorderColor = m_BColor(tbsRight)
End Property
Public Property Let RightBorderColor(ByVal Value As Long)
    LetBorderColor tbsRight, Value
End Property

Public Property Get RightBorderWidth() As Single
    RightBorderWidth = m_BWidth(tbsRight)
End Property
Public Property Let RightBorderWidth(ByVal Value As Single)
    LetBorderWidth tbsRight, Value
End Property

' Length along the edge in points; 0 = full edge, longer = clipped to edge.
Public Property Get RightBorderLength() As Single
    RightBorderLength = m_BLength(tbsRight)
End Property
Public Property Let RightBorderLength(ByVal Value As Single)
    LetBorderLength tbsRight, Value
End Property

' Anchor on the edge: "Top" | "Bottom" | "Center"
Public Property Get RightBorderStartFrom() As String
    RightBorderStartFrom = AnchorName(tbsRight)
End Property
Public Property Let RightBorderStartFrom(ByVal Value As String)
    LetBorderStart tbsRight, Value
End Property

'--- Bottom border ---------------------------------------------------------------
Public Property Get BottomBorderVisible() As Boolean
    BottomBorderVisible = m_BVisible(tbsBottom)
End Property
Public Property Let BottomBorderVisible(ByVal Value As Boolean)
    LetBorderVisible tbsBottom, Value
End Property

Public Property Get BottomBorderColor() As Long
    BottomBorderColor = m_BColor(tbsBottom)
End Property
Public Property Let BottomBorderColor(ByVal Value As Long)
    LetBorderColor tbsBottom, Value
End Property

Public Property Get BottomBorderWidth() As Single
    BottomBorderWidth = m_BWidth(tbsBottom)
End Property
Public Property Let BottomBorderWidth(ByVal Value As Single)
    LetBorderWidth tbsBottom, Value
End Property

' Length along the edge in points; 0 = full edge, longer = clipped to edge.
Public Property Get BottomBorderLength() As Single
    BottomBorderLength = m_BLength(tbsBottom)
End Property
Public Property Let BottomBorderLength(ByVal Value As Single)
    LetBorderLength tbsBottom, Value
End Property

' Anchor on the edge: "Left" | "Right" | "Center"
Public Property Get BottomBorderStartFrom() As String
    BottomBorderStartFrom = AnchorName(tbsBottom)
End Property
Public Property Let BottomBorderStartFrom(ByVal Value As String)
    LetBorderStart tbsBottom, Value
End Property

'--- Optional states (Variant: Empty means "do not change") --------------------
' PasswordChar: pass "*" (or ChrW(9679) for a dot) to mask; "" to unmask.
Public Property Get PasswordChar() As Variant
    PasswordChar = m_PasswordChar
End Property
Public Property Let PasswordChar(ByVal Value As Variant)
    If IsEmpty(Value) Then
        m_PasswordChar = Empty
    Else
        m_PasswordChar = Left$(CStr(Value), 1)
    End If
    ApplyIfAttached
End Property

Public Property Get IsEnabled() As Variant
    IsEnabled = m_Enabled
End Property
Public Property Let IsEnabled(ByVal Value As Variant)
    If IsEmpty(Value) Then m_Enabled = Empty Else m_Enabled = CBool(Value)
    ApplyIfAttached
End Property

Public Property Get IsLocked() As Variant
    IsLocked = m_Locked
End Property
Public Property Let IsLocked(ByVal Value As Variant)
    If IsEmpty(Value) Then m_Locked = Empty Else m_Locked = CBool(Value)
    ApplyIfAttached
End Property
