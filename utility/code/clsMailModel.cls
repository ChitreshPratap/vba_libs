VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMailModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "clsMailModel"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Class: clsMailModel
' Purpose: PURE DATA model for an email — To, CC, BCC (each
'          supporting multiple recipients), Subject, Body,
'          Attachments, etc. Holds no reference to
'          Outlook.Application and performs no Outlook actions.
'          Use modOutlookMailService to turn this into an
'          actual mail item (save/send/display/save-to-path).
'
' Error handling:
'   Single-item Add* methods (AddToRecipient, AddCCRecipient,
'   AddBCCRecipient, AddAttachment) RAISE an error on invalid
'   input (Err.Raise) instead of failing silently — wrap calls
'   in On Error to handle.

'   Batch Add* methods (AddToRecipients, AddCCRecipients,
'   AddBCCRecipients, AddAttachments) are lenient: they skip
'   invalid entries and return the count that succeeded, so one
'   bad address in a list doesn't abort the whole import. They
'   DO raise if the argument passed in isn't actually an array.
' ============================================================

' ---- Custom error numbers (vbObjectError + offset) ----
Const ERR_INVALID_EMAIL As Long = vbObjectError + 6001
Const ERR_ATTACHMENT_NOT_FOUND As Long = vbObjectError + 6002
Const ERR_NOT_AN_ARRAY As Long = vbObjectError + 6003
Const ERR_NO_RECIPIENT As Long = vbObjectError + 6004

Private mToRecipients As Collection  ' Collection of email Strings
Private mCCRecipients As Collection  ' Collection of email Strings
Private mBCCRecipients As Collection ' Collection of email Strings
Private mSubject As String
Private mBody As String
Private mHTMLBody As String
Private mIsHTML As Boolean
Private mImportance As Long          ' 0 = Low, 1 = Normal, 2 = High
Private mAttachments As Collection   ' Collection of file path Strings

Private Sub Class_Initialize()
    Set mToRecipients = New Collection
    Set mCCRecipients = New Collection
    Set mBCCRecipients = New Collection
    Set mAttachments = New Collection
    mImportance = 1 ' Normal
End Sub

Private Sub Class_Terminate()
    Set mToRecipients = Nothing
    Set mCCRecipients = Nothing
    Set mBCCRecipients = Nothing
    Set mAttachments = Nothing
End Sub

' ---- Recipients: To ----

' Add a single recipient. Raises ERR_INVALID_EMAIL if the address looks malformed.
Public Function AddToRecipient(ByVal EmailAddress As String) As Boolean
    AddToRecipient = AddRecipientToCollection(mToRecipients, EmailAddress)
End Function

' Add several recipients at once (array of email strings).
' Skips malformed addresses; returns the count successfully added.
Public Function AddToRecipients(ByVal EmailAddresses As Variant) As Long
    AddToRecipients = AddRecipientsToCollection(mToRecipients, EmailAddresses)
End Function

Public Sub ClearToRecipients()
    Set mToRecipients = New Collection
End Sub

Public Property Get ToRecipients() As Collection
    Set ToRecipients = mToRecipients
End Property

' Semicolon-joined string, ready for Outlook's .To property
Public Property Get ToRecipientsList() As String
    ToRecipientsList = JoinCollection(mToRecipients)
End Property

Public Property Get ToCount() As Long
    ToCount = mToRecipients.count
End Property

' ---- Recipients: CC ----

Public Function AddCCRecipient(ByVal EmailAddress As String) As Boolean
    AddCCRecipient = AddRecipientToCollection(mCCRecipients, EmailAddress)
End Function

Public Function AddCCRecipients(ByVal EmailAddresses As Variant) As Long
    AddCCRecipients = AddRecipientsToCollection(mCCRecipients, EmailAddresses)
End Function

Public Sub ClearCCRecipients()
    Set mCCRecipients = New Collection
End Sub

Public Property Get CCRecipients() As Collection
    Set CCRecipients = mCCRecipients
End Property

Public Property Get CCRecipientsList() As String
    CCRecipientsList = JoinCollection(mCCRecipients)
End Property

Public Property Get CCCount() As Long
    CCCount = mCCRecipients.count
End Property

' ---- Recipients: BCC ----

Public Function AddBCCRecipient(ByVal EmailAddress As String) As Boolean
    AddBCCRecipient = AddRecipientToCollection(mBCCRecipients, EmailAddress)
End Function

Public Function AddBCCRecipients(ByVal EmailAddresses As Variant) As Long
    AddBCCRecipients = AddRecipientsToCollection(mBCCRecipients, EmailAddresses)
End Function

Public Sub ClearBCCRecipients()
    Set mBCCRecipients = New Collection
End Sub

Public Property Get BCCRecipients() As Collection
    Set BCCRecipients = mBCCRecipients
End Property

Public Property Get BCCRecipientsList() As String
    BCCRecipientsList = JoinCollection(mBCCRecipients)
End Property

Public Property Get BCCCount() As Long
    BCCCount = mBCCRecipients.count
End Property

' ---- Subject / Body ----

Public Property Let Subject(ByVal Value As String)
    mSubject = Value
End Property
Public Property Get Subject() As String
    Subject = mSubject
End Property

' Plain text body
Public Property Let Body(ByVal Value As String)
    mBody = Value
End Property
Public Property Get Body() As String
    Body = mBody
End Property

' HTML body — setting this switches IsHTML to True automatically
Public Property Let HTMLBody(ByVal Value As String)
    mHTMLBody = Value
    mIsHTML = (Len(Value) > 0)
End Property
Public Property Get HTMLBody() As String
    HTMLBody = mHTMLBody
End Property

Public Property Get IsHTML() As Boolean
    IsHTML = mIsHTML
End Property

' ---- Misc ----

Public Property Let Importance(ByVal Value As Long)
    mImportance = Value
End Property
Public Property Get Importance() As Long
    Importance = mImportance
End Property

' ---- Attachments (stores paths only; no Outlook object touched here) ----

' Raises ERR_ATTACHMENT_NOT_FOUND if the file doesn't exist on disk.
Public Function AddAttachment(ByVal FilePath As String) As Boolean
    If Len(Dir(FilePath)) = 0 Then
        Err.Raise ERR_ATTACHMENT_NOT_FOUND, "clsMailModel.AddAttachment", _
            "Attachment file not found: '" & FilePath & "'"
    End If
    mAttachments.Add FilePath
    AddAttachment = True
End Function

' Skips files that don't exist; returns the count successfully added.
Public Function AddAttachments(ByVal FilePaths As Variant) As Long
    If Not IsArray(FilePaths) Then
        Err.Raise ERR_NOT_AN_ARRAY, "clsMailModel.AddAttachments", _
            "FilePaths must be an array of file path strings."
    End If

    Dim i As Long, successCount As Long
    For i = LBound(FilePaths) To UBound(FilePaths)
        On Error Resume Next
        Dim added As Boolean
        added = False
        added = AddAttachment(CStr(FilePaths(i)))
        On Error GoTo 0
        If added Then successCount = successCount + 1
    Next i
    AddAttachments = successCount
End Function

Public Sub ClearAttachments()
    Set mAttachments = New Collection
End Sub

Public Property Get Attachments() As Collection
    Set Attachments = mAttachments
End Property

Public Property Get AttachmentCount() As Long
    AttachmentCount = mAttachments.count
End Property

' ---- Validation ----

' Non-raising check — use when you just want a True/False answer.
Public Function IsValid() As Boolean
    IsValid = (mToRecipients.count > 0)
End Function

' Raising check — call this before building a real mail item so the
' caller gets a clear, specific error instead of an Outlook failure.
Public Sub ValidateOrRaise()
    If mToRecipients.count = 0 Then
        Err.Raise ERR_NO_RECIPIENT, "clsMailModel.ValidateOrRaise", _
            "At least one To recipient is required. Call AddToRecipient first."
    End If
End Sub

' ---- Private helpers ----

Private Function AddRecipientToCollection(ByVal TargetCollection As Collection, ByVal EmailAddress As String) As Boolean
    EmailAddress = Trim$(EmailAddress)
    If Not IsValidEmailFormat(EmailAddress) Then
        Err.Raise ERR_INVALID_EMAIL, "clsMailModel", "Invalid email address format: '" & EmailAddress & "'"
    End If
    TargetCollection.Add EmailAddress
    AddRecipientToCollection = True
End Function

Private Function AddRecipientsToCollection(ByVal TargetCollection As Collection, ByVal EmailAddresses As Variant) As Long
    If Not IsArray(EmailAddresses) Then
        Err.Raise ERR_NOT_AN_ARRAY, "clsMailModel", "EmailAddresses must be an array of email strings."
    End If

    Dim i As Long, successCount As Long
    For i = LBound(EmailAddresses) To UBound(EmailAddresses)
        On Error Resume Next
        Dim added As Boolean
        added = False
        added = AddRecipientToCollection(TargetCollection, CStr(EmailAddresses(i)))
        On Error GoTo 0
        If added Then successCount = successCount + 1
    Next i
    AddRecipientsToCollection = successCount
End Function

' Light validation — not a full RFC check, just catches empty/obviously bad input
Private Function IsValidEmailFormat(ByVal EmailAddress As String) As Boolean
    IsValidEmailFormat = (Len(EmailAddress) > 0) _
                          And (InStr(EmailAddress, "@") > 1) _
                          And (InStr(EmailAddress, ".") > InStr(EmailAddress, "@")) _
                          And (InStr(EmailAddress, " ") = 0)
End Function

Private Function JoinCollection(ByVal TargetCollection As Collection) As String
    Dim item As Variant
    Dim result As String
    For Each item In TargetCollection
        If Len(result) > 0 Then result = result & ";"
        result = result & CStr(item)
    Next item
    JoinCollection = result
End Function

