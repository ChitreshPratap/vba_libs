Attribute VB_Name = "OutlookUtility2"
'Attribute VB_Name = "modOutlookMailService"
Option Explicit

' ============================================================
' Module: modOutlookMailService
' Purpose: All Outlook automation logic. Takes a clsMailModel
'          (pure data) and turns it into a real Outlook
'          MailItem, then saves to Drafts / sends / displays /
'          saves to a specific file path.
'          The model itself never touches Outlook — this
'          module is the only place that does.
'
' Error handling:
'   Every Public function here raises an error (Err.Raise) on
'   failure instead of just returning False — wrap calls in
'   On Error to catch and inspect Err.Number / Err.Description.
'   Successful calls return True (or the MailItem, for
'   DisplayMail).
' ============================================================

' ---- Custom error numbers (vbObjectError + offset) ----
Public Const ERR_OUTLOOK_UNAVAILABLE As Long = vbObjectError + 7001
Public Const ERR_NULL_MODEL As Long = vbObjectError + 7002
Public Const ERR_BUILD_MAILITEM_FAILED As Long = vbObjectError + 7003
Public Const ERR_SAVE_FAILED As Long = vbObjectError + 7004
Public Const ERR_SEND_FAILED As Long = vbObjectError + 7005
Public Const ERR_DISPLAY_FAILED As Long = vbObjectError + 7006
Public Const ERR_INVALID_PATH As Long = vbObjectError + 7007
Public Const ERR_SAVETOPATH_FAILED As Long = vbObjectError + 7008

' File formats accepted by MailItem.SaveAs (matches Outlook's OlSaveAsType,
' redefined here so this module works without an Outlook object reference).
Public Enum MailSaveFormat
    msfTXT = 0
    msfRTF = 1
    msfTemplate = 2
    msfMSG = 3          ' default — standard Outlook .msg file
    msfDoc = 4
    msfHTML = 5
    msfVCard = 6
    msfVCal = 7
    msfICal = 8
    msfMSGUnicode = 9
End Enum

Private mOutApp As Object ' cached Outlook.Application, reused across calls

' Gets the running Outlook instance, or starts one if needed.
' Raises ERR_OUTLOOK_UNAVAILABLE if Outlook can't be reached at all.
Private Function GetOutlookApp() As Object
    If mOutApp Is Nothing Then
        On Error Resume Next
        Set mOutApp = GetObject(, "Outlook.Application")
        If mOutApp Is Nothing Then
            Set mOutApp = CreateObject("Outlook.Application")
        End If
        On Error GoTo 0

        If mOutApp Is Nothing Then
            Err.Raise ERR_OUTLOOK_UNAVAILABLE, "modOutlookMailService.GetOutlookApp", _
                "Could not start or connect to Outlook. Make sure Outlook is installed."
        End If
    End If
    Set GetOutlookApp = mOutApp
End Function

' Builds a live Outlook MailItem from a clsMailModel's data.
' Raises ERR_NULL_MODEL, clsMailModel's own validation error, or
' ERR_BUILD_MAILITEM_FAILED on failure.
Private Function BuildMailItem(ByVal Model As clsMailModel) As Object
    If Model Is Nothing Then
        Err.Raise ERR_NULL_MODEL, "modOutlookMailService.BuildMailItem", "Model argument cannot be Nothing."
    End If

    Model.ValidateOrRaise ' propagates clsMailModel's own error if invalid

    Dim OutApp As Object
    Dim OutMail As Object
    Dim att As Variant

    Set OutApp = GetOutlookApp()

    On Error GoTo ErrHandler
    Set OutMail = OutApp.CreateItem(0) ' 0 = olMailItem

    With OutMail
        .To = Model.ToRecipientsList
        .CC = Model.CCRecipientsList
        .BCC = Model.BCCRecipientsList
        .Subject = Model.Subject
        .Importance = Model.Importance

        If Model.IsHTML Then
            .HTMLBody = Model.HTMLBody
        Else
            .Body = Model.Body
        End If

        For Each att In Model.Attachments
            .Attachments.Add CStr(att)
        Next att
    End With

    Set BuildMailItem = OutMail
    Exit Function

ErrHandler:
    Err.Raise ERR_BUILD_MAILITEM_FAILED, "modOutlookMailService.BuildMailItem", _
        "Failed to build Outlook mail item: " & Err.Description
End Function

' Saves the mail to the Drafts folder instead of sending it.
Public Function SaveAsDraft(ByVal Model As clsMailModel) As Boolean
    Dim OutMail As Object
    Set OutMail = BuildMailItem(Model)

    On Error GoTo ErrHandler
    OutMail.Save
    SaveAsDraft = True
    Exit Function

ErrHandler:
    Err.Raise ERR_SAVE_FAILED, "modOutlookMailService.SaveAsDraft", _
        "Failed to save mail to Drafts: " & Err.Description
End Function

' Sends the mail immediately.
Public Function SendMail(ByVal Model As clsMailModel) As Boolean
    Dim OutMail As Object
    Set OutMail = BuildMailItem(Model)

    On Error GoTo ErrHandler
    OutMail.Send
    SendMail = True
    Exit Function

ErrHandler:
    Err.Raise ERR_SEND_FAILED, "modOutlookMailService.SendMail", _
        "Failed to send mail: " & Err.Description
End Function

' Opens the mail in an Outlook window for manual review/editing.
' Returns the MailItem in case the caller wants to keep working with it.
Public Function DisplayMail(ByVal Model As clsMailModel) As Object
    Dim OutMail As Object
    Set OutMail = BuildMailItem(Model)

    On Error GoTo ErrHandler
    OutMail.Display
    Set DisplayMail = OutMail
    Exit Function

ErrHandler:
    Err.Raise ERR_DISPLAY_FAILED, "modOutlookMailService.DisplayMail", _
        "Failed to display mail: " & Err.Description
End Function

' Saves the mail as a file at a specific path (e.g. "C:\Backup\Mail.msg")
' instead of putting it in Drafts. Default format is .msg.
' Raises ERR_INVALID_PATH if FilePath is blank, has no folder component,
' or the destination folder doesn't exist.
Public Function SaveToPath(ByVal Model As clsMailModel, ByVal FilePath As String, _
                            Optional ByVal SaveFormat As MailSaveFormat = msfMSG) As Boolean

    If Len(Trim$(FilePath)) = 0 Then
        Err.Raise ERR_INVALID_PATH, "modOutlookMailService.SaveToPath", "FilePath cannot be blank."
    End If

    Dim slashPos As Long
    slashPos = InStrRev(FilePath, "\")
    If slashPos = 0 Then
        Err.Raise ERR_INVALID_PATH, "modOutlookMailService.SaveToPath", _
            "FilePath must be a full path including a folder, e.g. 'C:\Backup\Mail.msg'. Got: '" & FilePath & "'"
    End If

    Dim FolderPath As String
    FolderPath = Left$(FilePath, slashPos)
    If Not FolderExists(FolderPath) Then
        Err.Raise ERR_INVALID_PATH, "modOutlookMailService.SaveToPath", _
            "Destination folder does not exist: '" & FolderPath & "'"
    End If

    Dim OutMail As Object
    Set OutMail = BuildMailItem(Model)

    On Error GoTo ErrHandler
    OutMail.SaveAs FilePath, CLng(SaveFormat)
    SaveToPath = True
    Exit Function

ErrHandler:
    Err.Raise ERR_SAVETOPATH_FAILED, "modOutlookMailService.SaveToPath", _
        "Failed to save mail to '" & FilePath & "': " & Err.Description
End Function

Private Function FolderExists(ByVal FolderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(Dir$(FolderPath, vbDirectory)) > 0)
    On Error GoTo 0
End Function

