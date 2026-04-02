VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HBoxLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: HBoxLayout
' Purpose: Arrange MSForms controls horizontally (like a Python HBox).
' Works with UserForm or Frame controls (any container that has .Controls, .Width, .InsideHeight, .InsideWidth).
' Uses MSForms.Control type for controls.

Option Explicit

Private mControls As Collection
Private mSpacing As Long
Private mPaddingLeft As Long
Private mPaddingRight As Long
Private mPaddingTop As Long
Private mPaddingBottom As Long
Private mVerticalAlign As String ' "top", "center", "bottom"
Private mStretchHeight As Boolean ' if True, stretch child controls to fit container height - paddings considered

' --- Constructor replacement (Initialize)
Private Sub Class_Initialize()
    Set mControls = New Collection
    mSpacing = 6
    mPaddingLeft = 6
    mPaddingRight = 6
    mPaddingTop = 6
    mPaddingBottom = 6
    mVerticalAlign = "center"
    mStretchHeight = False
End Sub

' --- Properties
Public Property Get Spacing() As Long
    Spacing = mSpacing
End Property
Public Property Let Spacing(val As Long)
    mSpacing = val
End Property

Public Property Get PaddingLeft() As Long
    PaddingLeft = mPaddingLeft
End Property
Public Property Let PaddingLeft(val As Long)
    mPaddingLeft = val
End Property

Public Property Get PaddingRight() As Long
    PaddingRight = mPaddingRight
End Property
Public Property Let PaddingRight(val As Long)
    mPaddingRight = val
End Property

Public Property Get PaddingTop() As Long
    PaddingTop = mPaddingTop
End Property
Public Property Let PaddingTop(val As Long)
    mPaddingTop = val
End Property

Public Property Get PaddingBottom() As Long
    PaddingBottom = mPaddingBottom
End Property
Public Property Let PaddingBottom(val As Long)
    mPaddingBottom = val
End Property

Public Property Get VerticalAlign() As String
    VerticalAlign = mVerticalAlign
End Property
Public Property Let VerticalAlign(val As String)
    val = LCase$(Trim$(val))
    If val = "top" Or val = "center" Or val = "bottom" Then
        mVerticalAlign = val
    Else
        Err.Raise vbObjectError + 1000, "HBoxLayout", "VerticalAlign must be 'top', 'center' or 'bottom'."
    End If
End Property

Public Property Get StretchHeight() As Boolean
    StretchHeight = mStretchHeight
End Property
Public Property Let StretchHeight(val As Boolean)
    mStretchHeight = val
End Property

' --- Methods to manage child controls
Public Sub Add(ctrl As Object)
    ' ctrl should be a MSForms.Control (e.g. CommandButton, TextBox). Caller must ensure control belongs to same container.
    mControls.Add ctrl
End Sub

Public Sub Remove(indexOrName As Variant)
    ' Accepts numeric index or control name
    If IsNumeric(indexOrName) Then
        mControls.Remove CLng(indexOrName)
    Else
        Dim i As Long
        For i = mControls.count To 1 Step -1
            If mControls(i).Name = CStr(indexOrName) Then
                mControls.Remove i
                Exit Sub
            End If
        Next i
    End If
End Sub

Public Sub Clear()
    Set mControls = Nothing
    Set mControls = New Collection
End Sub

Public Function count() As Long
    count = mControls.count
End Function

' --- Core: Arrange controls inside the container (UserForm or Frame)
Public Sub Arrange(container As Object)
    ' container must expose .Controls and .InsideWidth/.Width and .InsideHeight/.Height depend on host
    ' We'll handle: UserForm (use .Width and .Height) and Frame (.Width and .Height)
    Dim totalWidth As Long
    Dim i As Long
    Dim ctrl As MSForms.Control
    Dim x As Long
    Dim containerWidth As Long
    Dim containerHeight As Long
    Dim contentHeight As Long
    Dim availHeight As Long
    Dim ctrlHeight As Long
    Dim topPos As Long
    Dim useInsideWidth As Boolean
    
    ' Determine container dimensions.
    ' Many userforms/frames: use .Width and .Height as twips; sometimes frames expose InsideWidth/InsideHeight.
    On Error Resume Next
    containerWidth = container.InsideWidth
    If Err.Number <> 0 Then
        Err.Clear
        containerWidth = container.Width
    End If
    containerHeight = container.InsideHeight
    If Err.Number <> 0 Then
        Err.Clear
        containerHeight = container.Height
    End If
    On Error GoTo 0
    
    ' Remove invisible or nothing controls from layout calculation
    totalWidth = 0
    For i = 1 To mControls.count
        Set ctrl = mControls(i)
        If ctrl.visible Then
            totalWidth = totalWidth + ctrl.Width
        End If
    Next i
    ' Add spacing between visible controls
    Dim visibleCount As Long
    visibleCount = 0
    For i = 1 To mControls.count
        If mControls(i).visible Then visibleCount = visibleCount + 1
    Next i
    If visibleCount > 1 Then
        totalWidth = totalWidth + (visibleCount - 1) * mSpacing
    End If
    ' Add paddings
    totalWidth = totalWidth + mPaddingLeft + mPaddingRight
    
    ' Starting left position (respecting left padding)
    x = mPaddingLeft
    
    ' Available inner height: containerHeight minus paddings
    availHeight = containerHeight - (mPaddingTop + mPaddingBottom)
    If availHeight < 0 Then availHeight = 0
    
    ' Arrange each visible control
    For i = 1 To mControls.count
        Set ctrl = mControls(i)
        If ctrl.visible Then
            ' Determine height to place control at (stretch if requested)
            If mStretchHeight Then
                ctrlHeight = availHeight
                If ctrlHeight < 0 Then ctrlHeight = ctrl.Height
                ctrl.Height = ctrlHeight
            Else
                ctrlHeight = ctrl.Height
            End If
            
            Select Case mVerticalAlign
                Case "top"
                    topPos = mPaddingTop
                Case "center"
                    topPos = mPaddingTop + (availHeight - ctrlHeight) \ 2
                    If topPos < mPaddingTop Then topPos = mPaddingTop
                Case "bottom"
                    topPos = mPaddingTop + (availHeight - ctrlHeight)
                    If topPos < mPaddingTop Then topPos = mPaddingTop
                Case Else
                    topPos = mPaddingTop
            End Select
            
            ' Set position (Left & Top are in same units as Width/Height)
            ctrl.Left = x
            ctrl.Top = topPos
            ' Advance x
            x = x + ctrl.Width + mSpacing
        End If
    Next i
End Sub

' --- Helper: center whole group horizontally inside container
Public Sub CenterIn(container As Object)
    Dim groupWidth As Long
    Dim i As Long
    Dim ctrl As MSForms.Control
    Dim visibleCount As Long
    Dim containerWidth As Long
    
    On Error Resume Next
    containerWidth = container.InsideWidth
    If Err.Number <> 0 Then
        Err.Clear
        containerWidth = container.Width
    End If
    On Error GoTo 0
    
    groupWidth = 0
    visibleCount = 0
    For i = 1 To mControls.count
        Set ctrl = mControls(i)
        If ctrl.visible Then
            groupWidth = groupWidth + ctrl.Width
            visibleCount = visibleCount + 1
        End If
    Next i
    If visibleCount > 1 Then groupWidth = groupWidth + (visibleCount - 1) * mSpacing
    groupWidth = groupWidth + mPaddingLeft + mPaddingRight
    
    Dim startX As Long
    startX = (containerWidth - groupWidth) \ 2
    If startX < 0 Then startX = 0
    
    ' Temporarily set padding left to startX and re-arrange
    Dim oldPadLeft As Long
    oldPadLeft = mPaddingLeft
    mPaddingLeft = startX
    Arrange container
    mPaddingLeft = oldPadLeft
End Sub

