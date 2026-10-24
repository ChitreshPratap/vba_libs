VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BorderUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Class   : BorderUtility
' Purpose : Draw fully configurable border segments around any MSForms control
'           (Label, Frame, TextBox, Image, ...) on an Excel VBA UserForm,
'           WITHOUT changing the target control in any way.
'
' Rendering
'   Every border segment is a separate, caption-less, opaque MSForms Label
'   ("helper") created at run time in the SAME container as the target
'   (UserForm, Frame or MultiPage page). Because target and helpers share one
'   coordinate system, positions are computed directly from the target's
'   Left/Top/Width/Height. The target itself is only ever READ.
'
' Units
'   * Length and Offset : LengthUnit property (default = centimetres).
'                         Switch with  .LengthUnit = bdrPoints
'                         The unit is applied when a value is passed in;
'                         values are stored internally in points.
'   * Thickness         : ALWAYS points (parameter names end in "Pt").
'   * 1 cm = 72 / 2.54 = 28.3465 pt. MSForms coordinates are points.
'
' Anchors (where a segment starts along its side)
'   Top / Bottom sides : "Left"  | "Center" | "Right"   (or Start/Center/End)
'   Left / Right sides : "Top"   | "Center" | "Bottom"  (or Start/Center/End)
'   * Start  : segment begins at the Left/Top edge and runs Length inward.
'   * Center : segment is centred on the side's midpoint.
'   * End    : segment ends exactly at the Right/Bottom edge.
'   Offset (optional, LengthUnit):
'   * Start/End anchors : positive = move INWARD (away from the anchor edge).
'   * Center anchor     : positive = move right (horizontal) / down (vertical).
'   An explicit offset is honoured even if it pushes the segment past a corner.
'
' Placement (instance-wide, .Placement)
'   bdrOutside  (default) segment sits just outside the edge  -> never covers
'                         the control's content.
'   bdrInside             segment sits on the inner side of the edge.
'   bdrStraddle           segment is centred on the edge line.
'   Frames / MultiPages are windowed controls that paint over sibling Labels,
'   so with AutoOutsideForWindowed = True (default) they always use Outside.
'
' Overflow (.OverflowPolicy)
'   bdrClamp  (default) a segment longer than its side is drawn at full side
'                        length.
'   bdrReject            AddSegment / SetSegmentLength raise an error if the
'                        length exceeds the side. (Rendering always clamps, so
'                        a target that shrinks later still cannot overflow.)
'
' Overlaps
'   Segments are separate Labels. On Refresh they are stacked in the order they
'   were added: the segment added LAST is drawn on top.
'
' Errors are raised as vbObjectError + 1000 + n, Source "BorderUtility".
'==============================================================================
Option Explicit

'------------------------------------------------------------------ enums ----
Public Enum bdrSide
    bdrLeft = 0
    bdrTop = 1
    bdrRight = 2
    bdrBottom = 3
End Enum

Public Enum bdrAnchor
    bdrStart = 0        ' Left (Top/Bottom sides) or Top (Left/Right sides)
    bdrCenter = 1
    bdrEnd = 2          ' Right (Top/Bottom sides) or Bottom (Left/Right sides)
End Enum

Public Enum bdrUnit
    bdrPoints = 0
    bdrCentimeters = 1
End Enum

Public Enum bdrPlacement
    bdrOutside = 0
    bdrInside = 1
    bdrStraddle = 2
End Enum

Public Enum bdrOverflow
    bdrClamp = 0
    bdrReject = 1
End Enum

'-------------------------------------------------------------- constants ----
Private Const CM_TO_PT As Double = 28.3464566929134
Private Const HELPER_ROOT As String = "zzBdr_"
Private Const MAX_THICKNESS_PT As Double = 144
Private Const ERR_BASE As Long = vbObjectError + 1000

'------------------------------------------------------------------ types ----
Private Type tSegment
    Id As String
    Side As Long
    LengthPt As Double          ' -1 = full side length
    Anchor As Long
    ThicknessPt As Double
    Color As Long
    OffsetPt As Double
    CornerExtend As Boolean     ' extend both ends to close corners (AddFullBorder)
    Visible As Boolean
    CtlName As String           ' name of the helper Label
End Type

'---------------------------------------------------------------- members ----
Private mTarget As Object
Private mParent As Object
Private mPrefix As String
Private mSegs() As tSegment
Private mCount As Long
Private mSerial As Long
Private mAutoId As Long
Private mSideVisible(0 To 3) As Boolean
Private mUnit As bdrUnit
Private mPlacement As bdrPlacement
Private mOverflow As bdrOverflow
Private mAutoRefresh As Boolean
Private mAutoOutsideForWindowed As Boolean

'============================================================== lifecycle ====
Private Sub Class_Initialize()
    mPrefix = HELPER_ROOT & Hex$(ObjPtr(Me)) & "_"
    mUnit = bdrCentimeters
    mPlacement = bdrOutside
    mOverflow = bdrClamp
    mAutoRefresh = True
    mAutoOutsideForWindowed = True
    ResetSides
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Dispose
End Sub

' Attach the utility to a control. Calling it again (same or another control)
' first removes every helper this instance created and clears all segments.
Public Sub Initialize(ByVal Target As Object)
    Dim p As Object

    If Target Is Nothing Then RaiseErr 1, "Initialize: Target is Nothing."
    If Not HasGeometry(Target) Then
        RaiseErr 2, "Initialize: Target does not expose Left/Top/Width/Height. " & _
                    "Pass an MSForms control (Label, Frame, TextBox, Image, ...)."
    End If

    On Error Resume Next
    Set p = Target.Parent
    On Error GoTo 0
    If p Is Nothing Then RaiseErr 3, "Initialize: Target has no parent container."
    If Not HasControls(p) Then RaiseErr 4, "Initialize: Target's parent has no Controls collection."

    ' Repeated initialisation: remove everything drawn by this instance.
    If Not mParent Is Nothing Then
        RemoveAllHelpers
        RemoveByPrefix mParent, mPrefix
    End If
    mCount = 0
    Erase mSegs
    ResetSides

    Set mTarget = Target
    Set mParent = p
    RemoveByPrefix mParent, mPrefix     ' defensive: stale helpers with our prefix
End Sub

' Remove all helpers and segments, and detach from the target.
Public Sub Dispose()
    On Error Resume Next
    Clear
    If Not mParent Is Nothing Then RemoveByPrefix mParent, mPrefix
    Set mTarget = Nothing
    Set mParent = Nothing
    ResetSides
End Sub

' Remove all segments and their helper controls (utility stays attached).
Public Sub Clear()
    RemoveAllHelpers
    mCount = 0
    Erase mSegs
End Sub

' Redraw every segment from the target's CURRENT geometry and restack them
' in insertion order. Call after moving/resizing the target or when
' AutoRefresh = False.
Public Sub Refresh()
    Dim i As Long
    EnsureInitialized
    For i = 1 To mCount
        RenderOne i, True
    Next i
End Sub

' Removes EVERY BorderUtility helper (all instances) from a container.
' Use only for emergency cleanup; live instances recreate theirs on Refresh.
Public Sub PurgeAllHelpers(ByVal Container As Object)
    If Container Is Nothing Then Exit Sub
    RemoveByPrefix Container, HELPER_ROOT
End Sub

'============================================================ add segments ===
' Add one independent segment. Returns its Id.
'   Side        : "Left" | "Top" | "Right" | "Bottom"  (or bdrSide value)
'   Length      : > 0, in LengthUnit
'   Anchor      : see header (or bdrAnchor value)
'   ThicknessPt : > 0, in points
'   Color       : RGB(), vb colour constant or system colour
'   Offset      : optional, in LengthUnit
'   Id          : optional unique name (case-insensitive); auto "SegN" if blank
Public Function AddSegment(ByVal Side As Variant, ByVal Length As Double, _
                           ByVal Anchor As Variant, ByVal ThicknessPt As Double, _
                           ByVal Color As Long, Optional ByVal Offset As Double = 0, _
                           Optional ByVal Id As String = vbNullString) As String
    Dim sd As Long, an As Long, lenPt As Double

    EnsureInitialized
    sd = ParseSide(Side)
    an = ParseAnchor(Anchor, sd)
    If Length <= 0 Then RaiseErr 20, "AddSegment: Length must be > 0. Use AddFullSide for a full-length border."
    ValidateThickness ThicknessPt, "AddSegment"
    ValidateColor Color, "AddSegment"
    lenPt = ToPt(Length)
    CheckOverflow sd, lenPt, "AddSegment"
    AddSegment = AddInternal(sd, lenPt, an, ThicknessPt, Color, ToPt(Offset), False, Id)
End Function

' Add a segment that always spans the whole side (follows target resizes).
Public Function AddFullSide(ByVal Side As Variant, ByVal ThicknessPt As Double, _
                            ByVal Color As Long, Optional ByVal Id As String = vbNullString) As String
    Dim sd As Long

    EnsureInitialized
    sd = ParseSide(Side)
    ValidateThickness ThicknessPt, "AddFullSide"
    ValidateColor Color, "AddFullSide"
    AddFullSide = AddInternal(sd, -1, bdrStart, ThicknessPt, Color, 0, False, Id)
End Function

' Add a closed rectangle: 4 full sides, corners filled.
' Ids: IdPrefix & "Left" / "Right" / "Top" / "Bottom"  (e.g. FullTop).
Public Sub AddFullBorder(ByVal ThicknessPt As Double, ByVal Color As Long, _
                         Optional ByVal IdPrefix As String = "Full")
    Dim nm As Variant

    EnsureInitialized
    ValidateThickness ThicknessPt, "AddFullBorder"
    ValidateColor Color, "AddFullBorder"
    For Each nm In Array("Left", "Right", "Top", "Bottom")
        If FindIndex(IdPrefix & nm) > 0 Then
            RaiseErr 22, "AddFullBorder: segment Id '" & IdPrefix & nm & "' already exists."
        End If
    Next nm

    AddInternal bdrLeft, -1, bdrStart, ThicknessPt, Color, 0, False, IdPrefix & "Left"
    AddInternal bdrRight, -1, bdrStart, ThicknessPt, Color, 0, False, IdPrefix & "Right"
    AddInternal bdrTop, -1, bdrStart, ThicknessPt, Color, 0, True, IdPrefix & "Top"
    AddInternal bdrBottom, -1, bdrStart, ThicknessPt, Color, 0, True, IdPrefix & "Bottom"
End Sub

'======================================================== modify segments ====
Public Sub SetSegmentColor(ByVal Id As String, ByVal Color As Long)
    Dim i As Long
    EnsureInitialized
    i = RequireIndex(Id, "SetSegmentColor")
    ValidateColor Color, "SetSegmentColor"
    mSegs(i).Color = Color
    If mAutoRefresh Then RenderOne i, False
End Sub

Public Sub SetSegmentThickness(ByVal Id As String, ByVal ThicknessPt As Double)
    Dim i As Long
    EnsureInitialized
    i = RequireIndex(Id, "SetSegmentThickness")
    ValidateThickness ThicknessPt, "SetSegmentThickness"
    mSegs(i).ThicknessPt = ThicknessPt
    If mAutoRefresh Then RenderOne i, False
End Sub

' Length in LengthUnit, > 0.
Public Sub SetSegmentLength(ByVal Id As String, ByVal Length As Double)
    Dim i As Long, lenPt As Double
    EnsureInitialized
    i = RequireIndex(Id, "SetSegmentLength")
    If Length <= 0 Then RaiseErr 20, "SetSegmentLength: Length must be > 0. Use SetSegmentFullLength for a full side."
    lenPt = ToPt(Length)
    CheckOverflow mSegs(i).Side, lenPt, "SetSegmentLength"
    mSegs(i).LengthPt = lenPt
    If mAutoRefresh Then RenderOne i, False
End Sub

Public Sub SetSegmentFullLength(ByVal Id As String)
    Dim i As Long
    EnsureInitialized
    i = RequireIndex(Id, "SetSegmentFullLength")
    mSegs(i).LengthPt = -1
    If mAutoRefresh Then RenderOne i, False
End Sub

' Change the anchor and, optionally, the offset (LengthUnit).
Public Sub SetSegmentAnchor(ByVal Id As String, ByVal Anchor As Variant, Optional ByVal Offset As Variant)
    Dim i As Long, an As Long
    EnsureInitialized
    i = RequireIndex(Id, "SetSegmentAnchor")
    an = ParseAnchor(Anchor, mSegs(i).Side)
    If Not IsMissing(Offset) Then
        If Not IsNumeric(Offset) Then RaiseErr 24, "SetSegmentAnchor: Offset must be numeric."
        mSegs(i).OffsetPt = ToPt(CDbl(Offset))
    End If
    mSegs(i).Anchor = an
    If mAutoRefresh Then RenderOne i, False
End Sub

Public Sub SetSegmentOffset(ByVal Id As String, ByVal Offset As Double)
    Dim i As Long
    EnsureInitialized
    i = RequireIndex(Id, "SetSegmentOffset")
    mSegs(i).OffsetPt = ToPt(Offset)
    If mAutoRefresh Then RenderOne i, False
End Sub

Public Sub SetSegmentVisible(ByVal Id As String, ByVal Visible As Boolean)
    Dim i As Long
    EnsureInitialized
    i = RequireIndex(Id, "SetSegmentVisible")
    mSegs(i).Visible = Visible
    If mAutoRefresh Then RenderOne i, False
End Sub

Public Sub RemoveSegment(ByVal Id As String)
    Dim i As Long, j As Long
    EnsureInitialized
    i = RequireIndex(Id, "RemoveSegment")
    RemoveHelper mSegs(i).CtlName
    For j = i To mCount - 1
        mSegs(j) = mSegs(j + 1)
    Next j
    mCount = mCount - 1
End Sub

'========================================================== side visibility ==
Public Property Get SideVisible(ByVal Side As Variant) As Boolean
    SideVisible = mSideVisible(ParseSide(Side))
End Property

Public Property Let SideVisible(ByVal Side As Variant, ByVal Value As Boolean)
    Dim sd As Long, i As Long
    sd = ParseSide(Side)
    mSideVisible(sd) = Value
    If mAutoRefresh And Not mTarget Is Nothing Then
        For i = 1 To mCount
            If mSegs(i).Side = sd Then RenderOne i, False
        Next i
    End If
End Property

Public Sub ShowSide(ByVal Side As Variant)
    SideVisible(Side) = True
End Sub

Public Sub HideSide(ByVal Side As Variant)
    SideVisible(Side) = False
End Sub

'=============================================================== settings ====
' Unit for Length and Offset arguments (thickness is always points).
' Changing it does not alter segments already added.
Public Property Get LengthUnit() As bdrUnit
    LengthUnit = mUnit
End Property

Public Property Let LengthUnit(ByVal Value As bdrUnit)
    If Value <> bdrPoints And Value <> bdrCentimeters Then RaiseErr 12, "LengthUnit: use bdrPoints or bdrCentimeters."
    mUnit = Value
End Property

Public Property Get Placement() As bdrPlacement
    Placement = mPlacement
End Property

Public Property Let Placement(ByVal Value As bdrPlacement)
    If Value < bdrOutside Or Value > bdrStraddle Then RaiseErr 13, "Placement: use bdrOutside, bdrInside or bdrStraddle."
    mPlacement = Value
    If mAutoRefresh And Not mTarget Is Nothing Then Refresh
End Property

' Placement actually used (may be forced to Outside for Frame/MultiPage).
Public Property Get ActivePlacement() As bdrPlacement
    ActivePlacement = EffectivePlacement()
End Property

Public Property Get OverflowPolicy() As bdrOverflow
    OverflowPolicy = mOverflow
End Property

Public Property Let OverflowPolicy(ByVal Value As bdrOverflow)
    If Value <> bdrClamp And Value <> bdrReject Then RaiseErr 14, "OverflowPolicy: use bdrClamp or bdrReject."
    mOverflow = Value
End Property

' True (default): every Add/Set/Show/Hide redraws immediately.
' False: nothing is drawn until Refresh is called.
Public Property Get AutoRefresh() As Boolean
    AutoRefresh = mAutoRefresh
End Property

Public Property Let AutoRefresh(ByVal Value As Boolean)
    mAutoRefresh = Value
End Property

Public Property Get AutoOutsideForWindowed() As Boolean
    AutoOutsideForWindowed = mAutoOutsideForWindowed
End Property

Public Property Let AutoOutsideForWindowed(ByVal Value As Boolean)
    mAutoOutsideForWindowed = Value
    If mAutoRefresh And Not mTarget Is Nothing Then Refresh
End Property

'============================================================== inspection ===
Public Property Get IsInitialized() As Boolean
    IsInitialized = Not mTarget Is Nothing
End Property

Public Property Get Target() As Object
    Set Target = mTarget
End Property

Public Property Get SegmentCount() As Long
    SegmentCount = mCount
End Property

' 1-based index -> Id
Public Property Get SegmentId(ByVal Index As Long) As String
    If Index < 1 Or Index > mCount Then RaiseErr 25, "SegmentId: index out of range."
    SegmentId = mSegs(Index).Id
End Property

Public Function SegmentExists(ByVal Id As String) As Boolean
    SegmentExists = FindIndex(Id) > 0
End Function

' Length of a side of the target, in LengthUnit.
Public Function SideLength(ByVal Side As Variant) As Double
    EnsureInitialized
    SideLength = FromPt(SideLengthPt(ParseSide(Side)))
End Function

' Multi-line text describing every segment and its computed rectangle (points).
Public Function DescribeSegments() As String
    Dim i As Long, s As String, ln As String
    Dim x As Double, y As Double, w As Double, h As Double

    EnsureInitialized
    s = "Target " & mTarget.Name & "  L=" & Fmt(mTarget.Left) & " T=" & Fmt(mTarget.Top) & _
        " W=" & Fmt(mTarget.Width) & " H=" & Fmt(mTarget.Height) & " (pt)" & vbCrLf
    For i = 1 To mCount
        With mSegs(i)
            If .LengthPt < 0 Then ln = "full" Else ln = Fmt(FromPt(.LengthPt)) & " " & UnitName()
            s = s & "  " & .Id & ": " & SideName(.Side) & ", len=" & ln & _
                ", anchor=" & AnchorName(.Side, .Anchor) & _
                ", offset=" & Fmt(FromPt(.OffsetPt)) & " " & UnitName() & _
                ", thick=" & Fmt(.ThicknessPt) & "pt, color=&H" & Hex$(.Color) & _
                ", visible=" & (.Visible And mSideVisible(.Side))
        End With
        If ComputeRect(i, x, y, w, h) Then
            s = s & "  -> L=" & Fmt(x) & " T=" & Fmt(y) & " W=" & Fmt(w) & " H=" & Fmt(h)
        Else
            s = s & "  -> (nothing to draw)"
        End If
        s = s & vbCrLf
    Next i
    DescribeSegments = s
End Function

'=============================================================== internals ===
Private Function AddInternal(ByVal sd As Long, ByVal lenPt As Double, ByVal an As Long, _
                             ByVal t As Double, ByVal clr As Long, ByVal offPt As Double, _
                             ByVal cornerExt As Boolean, ByVal Id As String) As String
    Id = Trim$(Id)
    If Len(Id) = 0 Then
        Do
            mAutoId = mAutoId + 1
            Id = "Seg" & mAutoId
        Loop While FindIndex(Id) > 0
    ElseIf FindIndex(Id) > 0 Then
        RaiseErr 22, "Duplicate segment Id '" & Id & "'."
    End If

    mSerial = mSerial + 1
    mCount = mCount + 1
    If mCount = 1 Then
        ReDim mSegs(1 To 8)
    ElseIf mCount > UBound(mSegs) Then
        ReDim Preserve mSegs(1 To UBound(mSegs) * 2)
    End If

    With mSegs(mCount)
        .Id = Id
        .Side = sd
        .LengthPt = lenPt
        .Anchor = an
        .ThicknessPt = t
        .Color = clr
        .OffsetPt = offPt
        .CornerExtend = cornerExt
        .Visible = True
        .CtlName = mPrefix & CStr(mSerial)
    End With

    If mAutoRefresh Then RenderOne mCount, True
    AddInternal = Id
End Function

' Create/update the helper Label of segment i.
Private Sub RenderOne(ByVal i As Long, ByVal BringToFront As Boolean)
    Dim c As Object, created As Boolean, showIt As Boolean
    Dim x As Double, y As Double, w As Double, h As Double

    On Error GoTo Fail
    showIt = mSegs(i).Visible And mSideVisible(mSegs(i).Side) And TargetVisible()
    If showIt Then showIt = ComputeRect(i, x, y, w, h)

    Set c = GetOrCreateHelper(mSegs(i).CtlName, created)
    If showIt Then
        c.BackColor = mSegs(i).Color
        c.Left = x
        c.Top = y
        c.Width = w
        c.Height = h
        If BringToFront Or created Then c.ZOrder 0      ' 0 = fmTop
        c.Visible = True
    Else
        c.Visible = False
    End If
    Exit Sub
Fail:
    RaiseErr 40, "Could not render segment '" & mSegs(i).Id & "': " & Err.Description
End Sub

' Geometry of segment i in the parent's coordinate system (points).
' Reads the target; never writes to it.
Private Function ComputeRect(ByVal i As Long, ByRef x As Double, ByRef y As Double, _
                             ByRef w As Double, ByRef h As Double) As Boolean
    Dim tL As Double, tT As Double, tW As Double, tH As Double
    Dim sideLen As Double, segLen As Double, s As Double, t As Double
    Dim k As Double, ext As Double, plc As Long

    On Error GoTo Fail
    tL = mTarget.Left: tT = mTarget.Top
    tW = mTarget.Width: tH = mTarget.Height
    t = mSegs(i).ThicknessPt
    plc = EffectivePlacement()

    If IsHorizontalSide(mSegs(i).Side) Then sideLen = tW Else sideLen = tH
    If sideLen <= 0 Then Exit Function

    ' length along the side (render-time clamp)
    If mSegs(i).LengthPt < 0 Then segLen = sideLen Else segLen = mSegs(i).LengthPt
    If segLen > sideLen Then segLen = sideLen

    ' start position along the side, relative to the side's start
    Select Case mSegs(i).Anchor
        Case bdrStart:  s = mSegs(i).OffsetPt
        Case bdrCenter: s = (sideLen - segLen) / 2 + mSegs(i).OffsetPt
        Case Else:      s = sideLen - segLen - mSegs(i).OffsetPt
    End Select

    ' shift across the edge according to placement
    Select Case plc
        Case bdrOutside:  k = t
        Case bdrStraddle: k = t / 2
        Case Else:        k = 0
    End Select

    ' close the corners for AddFullBorder top/bottom segments
    If mSegs(i).CornerExtend Then
        ext = k
        s = s - ext
        segLen = segLen + 2 * ext
    End If

    Select Case mSegs(i).Side
        Case bdrTop
            x = tL + s: w = segLen: h = t
            y = tT - k
        Case bdrBottom
            x = tL + s: w = segLen: h = t
            y = tT + tH - t + k
        Case bdrLeft
            y = tT + s: h = segLen: w = t
            x = tL - k
        Case Else   ' bdrRight
            y = tT + s: h = segLen: w = t
            x = tL + tW - t + k
    End Select

    ComputeRect = (w > 0 And h > 0)
    Exit Function
Fail:
    ComputeRect = False
End Function

Private Function GetOrCreateHelper(ByVal nm As String, ByRef created As Boolean) As Object
    Dim c As Object

    created = False
    On Error Resume Next
    Set c = mParent.Controls(nm)
    On Error GoTo 0

    If c Is Nothing Then
        Set c = mParent.Controls.Add("Forms.Label.1", nm, False)
        c.Caption = vbNullString
        c.AutoSize = False
        c.BackStyle = 1         ' fmBackStyleOpaque
        c.BorderStyle = 0       ' fmBorderStyleNone
        c.SpecialEffect = 0     ' fmSpecialEffectFlat
        created = True
    End If
    Set GetOrCreateHelper = c
End Function

Private Sub RemoveHelper(ByVal nm As String)
    If mParent Is Nothing Then Exit Sub
    On Error Resume Next
    mParent.Controls.Remove nm
End Sub

Private Sub RemoveAllHelpers()
    Dim i As Long
    For i = 1 To mCount
        RemoveHelper mSegs(i).CtlName
    Next i
End Sub

' Removes controls whose Name starts with prefix. A UserForm's Controls
' collection also lists nested controls, so each is removed via its own Parent.
Private Sub RemoveByPrefix(ByVal Container As Object, ByVal prefix As String)
    Dim i As Long, nm As String, c As Object
    On Error Resume Next
    For i = Container.Controls.count - 1 To 0 Step -1
        nm = vbNullString
        Set c = Nothing
        Set c = Container.Controls(i)
        If Not c Is Nothing Then
            nm = c.Name
            If Len(nm) >= Len(prefix) Then
                If Left$(nm, Len(prefix)) = prefix Then c.Parent.Controls.Remove nm
            End If
        End If
    Next i
End Sub

Private Function EffectivePlacement() As Long
    EffectivePlacement = mPlacement
    If mPlacement <> bdrOutside And mAutoOutsideForWindowed And Not mTarget Is Nothing Then
        Select Case TypeName(mTarget)
            Case "Frame", "MultiPage"
                EffectivePlacement = bdrOutside
        End Select
    End If
End Function

Private Function TargetVisible() As Boolean
    On Error Resume Next
    TargetVisible = CBool(mTarget.Visible)
End Function

Private Function SideLengthPt(ByVal sd As Long) As Double
    If IsHorizontalSide(sd) Then SideLengthPt = mTarget.Width Else SideLengthPt = mTarget.Height
End Function

Private Function IsHorizontalSide(ByVal sd As Long) As Boolean
    IsHorizontalSide = (sd = bdrTop Or sd = bdrBottom)
End Function

Private Sub CheckOverflow(ByVal sd As Long, ByVal lenPt As Double, ByVal ctx As String)
    Dim avail As Double
    If mOverflow <> bdrReject Then Exit Sub
    avail = SideLengthPt(sd)
    If lenPt > avail + 0.001 Then
        RaiseErr 21, ctx & ": length " & Fmt(FromPt(lenPt)) & " " & UnitName() & _
                     " exceeds the " & SideName(sd) & " side (" & Fmt(FromPt(avail)) & " " & UnitName() & ")."
    End If
End Sub

Private Sub ValidateThickness(ByVal t As Double, ByVal ctx As String)
    If t <= 0 Or t > MAX_THICKNESS_PT Then
        RaiseErr 30, ctx & ": thickness must be > 0 and <= " & MAX_THICKNESS_PT & " pt (got " & t & ")."
    End If
End Sub

Private Sub ValidateColor(ByVal Color As Long, ByVal ctx As String)
    If Color >= 0 And Color <= &HFFFFFF Then Exit Sub
    If Color >= &H80000000 And Color <= &H80000018 Then Exit Sub     ' system colours
    RaiseErr 31, ctx & ": invalid colour value " & Color & _
                 ". Use RGB(), a vb colour constant or a system colour (&H800000xx)."
End Sub

Private Function ParseSide(ByVal Side As Variant) As Long
    Dim s As String

    If IsEmpty(Side) Or IsNull(Side) Then RaiseErr 10, "Side is missing."
    If VarType(Side) <> vbString And IsNumeric(Side) Then
        If CDbl(Side) = 0 Or CDbl(Side) = 1 Or CDbl(Side) = 2 Or CDbl(Side) = 3 Then
            ParseSide = CLng(Side)
            Exit Function
        End If
        RaiseErr 10, "Unsupported side value: " & CStr(Side)
    End If

    s = LCase$(Trim$(CStr(Side)))
    Select Case s
        Case "left", "l":   ParseSide = bdrLeft
        Case "top", "t":    ParseSide = bdrTop
        Case "right", "r":  ParseSide = bdrRight
        Case "bottom", "b": ParseSide = bdrBottom
        Case Else
            RaiseErr 10, "Unsupported side '" & CStr(Side) & "'. Use Left, Top, Right or Bottom."
    End Select
End Function

Private Function ParseAnchor(ByVal Anchor As Variant, ByVal sd As Long) As Long
    Dim s As String, horiz As Boolean

    If IsEmpty(Anchor) Or IsNull(Anchor) Then RaiseErr 11, "Anchor is missing."
    horiz = IsHorizontalSide(sd)

    If VarType(Anchor) <> vbString And IsNumeric(Anchor) Then
        If CDbl(Anchor) = 0 Or CDbl(Anchor) = 1 Or CDbl(Anchor) = 2 Then
            ParseAnchor = CLng(Anchor)
            Exit Function
        End If
        RaiseErr 11, "Unsupported anchor value: " & CStr(Anchor)
    End If

    s = LCase$(Trim$(CStr(Anchor)))
    Select Case s
        Case "start":                              ParseAnchor = bdrStart
        Case "center", "centre", "middle", "mid":  ParseAnchor = bdrCenter
        Case "end":                                ParseAnchor = bdrEnd
        Case "left"
            If horiz Then ParseAnchor = bdrStart Else BadAnchor Anchor, sd
        Case "right"
            If horiz Then ParseAnchor = bdrEnd Else BadAnchor Anchor, sd
        Case "top"
            If Not horiz Then ParseAnchor = bdrStart Else BadAnchor Anchor, sd
        Case "bottom"
            If Not horiz Then ParseAnchor = bdrEnd Else BadAnchor Anchor, sd
        Case Else
            BadAnchor Anchor, sd
    End Select
End Function

Private Sub BadAnchor(ByVal Anchor As Variant, ByVal sd As Long)
    If IsHorizontalSide(sd) Then
        RaiseErr 11, "Unsupported anchor '" & CStr(Anchor) & "' for the " & SideName(sd) & _
                     " side. Use Left, Center or Right."
    Else
        RaiseErr 11, "Unsupported anchor '" & CStr(Anchor) & "' for the " & SideName(sd) & _
                     " side. Use Top, Center or Bottom."
    End If
End Sub

Private Function FindIndex(ByVal Id As String) As Long
    Dim i As Long
    For i = 1 To mCount
        If StrComp(mSegs(i).Id, Id, vbTextCompare) = 0 Then
            FindIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function RequireIndex(ByVal Id As String, ByVal ctx As String) As Long
    RequireIndex = FindIndex(Id)
    If RequireIndex = 0 Then RaiseErr 23, ctx & ": no segment with Id '" & Id & "'."
End Function

Private Function HasGeometry(ByVal o As Object) As Boolean
    Dim d As Double
    On Error GoTo Done
    d = o.Left + o.Top + o.Width + o.Height
    HasGeometry = True
Done:
End Function

Private Function HasControls(ByVal o As Object) As Boolean
    Dim n As Long
    On Error GoTo Done
    n = o.Controls.count
    HasControls = True
Done:
End Function

Private Sub ResetSides()
    Dim i As Long
    For i = 0 To 3
        mSideVisible(i) = True
    Next i
End Sub

Private Sub EnsureInitialized()
    If mTarget Is Nothing Then RaiseErr 5, "BorderUtility is not initialised. Call .Initialize <control> first."
End Sub

Private Function ToPt(ByVal v As Double) As Double
    If mUnit = bdrCentimeters Then ToPt = v * CM_TO_PT Else ToPt = v
End Function

Private Function FromPt(ByVal v As Double) As Double
    If mUnit = bdrCentimeters Then FromPt = v / CM_TO_PT Else FromPt = v
End Function

Private Function UnitName() As String
    If mUnit = bdrCentimeters Then UnitName = "cm" Else UnitName = "pt"
End Function

Private Function SideName(ByVal sd As Long) As String
    SideName = Choose(sd + 1, "Left", "Top", "Right", "Bottom")
End Function

Private Function AnchorName(ByVal sd As Long, ByVal an As Long) As String
    If IsHorizontalSide(sd) Then
        AnchorName = Choose(an + 1, "Left", "Center", "Right")
    Else
        AnchorName = Choose(an + 1, "Top", "Center", "Bottom")
    End If
End Function

Private Function Fmt(ByVal v As Double) As String
    Fmt = Format$(v, "0.00")
End Function

Private Sub RaiseErr(ByVal n As Long, ByVal msg As String)
    Err.Raise ERR_BASE + n, "BorderUtility", msg
End Sub
