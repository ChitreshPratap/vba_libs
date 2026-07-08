Attribute VB_Name = "Test_Module_OutlookUtility2"
Option Explicit

' ============================================================
' Module: modMailUsageExample
' Purpose: Shows how to use clsMailModel (data) together with
'          modOutlookMailService (Outlook actions), including
'          multiple To/CC/BCC recipients, saving to a specific
'          file path, and handling the errors both modules raise.
'
' Setup, once per project:
'   1. Import clsMailModel.cls          (Model  - data only)
'   2. Import modOutlookMailService.bas (Service - Outlook logic)
'   3. Import this module for reference, or write your own
'      calling code the same way.
' ============================================================

Sub Example_MultipleRecipients()
    On Error GoTo ErrHandler

    Dim Model As New clsMailModel

    ' Add recipients one at a time (raises if an address is malformed)
    Model.AddToRecipient "recipient1@example.com"
    Model.AddToRecipient "recipient2@example.com"
    Model.AddCCRecipient "cc1@example.com"

    ' Or add several at once via an array (invalid ones are skipped, not raised)
    Dim bccList(1 To 2) As String
    bccList(1) = "bcc1@example.com"
    bccList(2) = "bcc2@example.com"
    Model.AddBCCRecipients bccList

    Model.Subject = "Subject goes here"
    Model.Body = "Hello," & vbNewLine & vbNewLine & _
                 "This is the body of the email." & vbNewLine & vbNewLine & _
                 "Regards," & vbNewLine & "Abhilash"

    Model.AddAttachment "C:\Path\To\Your\Attachment.pdf"

    modOutlookMailService.SaveAsDraft Model

    MsgBox Model.ToCount & " To, " & Model.CCCount & " CC, " & _
           Model.BCCCount & " BCC recipient(s). Saved to Drafts.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Example_MultipleRecipients failed"
End Sub

Sub Example_MultipleAttachments_HTML()
    On Error GoTo ErrHandler

    Dim Model As New clsMailModel
    Dim toList(1 To 2) As String
    Dim files(1 To 2) As String

    toList(1) = "recipient1@example.com"
    toList(2) = "recipient2@example.com"
    files(1) = "C:\Path\To\File1.pdf"
    files(2) = "C:\Path\To\File2.xlsx"

    With Model
        .AddToRecipients toList
        .AddCCRecipient "cc@example.com"
        .Subject = "Report attached"
        .HTMLBody = "<p>Hi,</p><p>Please find the reports attached.</p><p>Regards,<br>Abhilash</p>"
        .AddAttachments files
    End With

    modOutlookMailService.SaveAsDraft Model

    MsgBox Model.AttachmentCount & " attachment(s) added. Saved to Drafts.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Example_MultipleAttachments_HTML failed"
End Sub

Sub Example_SaveToSpecificPath()
    On Error GoTo ErrHandler

    Dim Model As New clsMailModel
    Model.AddToRecipient "recipient@example.com"
    Model.Subject = "Saved to disk"
    Model.Body = "This mail is saved as a .msg file instead of going to Drafts."

    ' Saves as an Outlook .msg file at the given path (folder must already exist)
    modOutlookMailService.SaveToPath Model, "C:\MailBackup\SavedMail.msg"

    MsgBox "Mail saved to C:\MailBackup\SavedMail.msg", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Example_SaveToSpecificPath failed"
End Sub

Sub Example_DisplayForReviewThenSend()
    On Error GoTo ErrHandler

    Dim Model As New clsMailModel
    Dim OutMail As Object

    With Model
        .AddToRecipient "recipient@example.com"
        .Subject = "Review before sending"
        .Body = "Draft content here."
    End With

    ' Opens the mail window so a person can review/edit before sending
    Set OutMail = modOutlookMailService.DisplayMail(Model)

    ' To send immediately instead of reviewing, use:
    ' modOutlookMailService.SendMail Model
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Example_DisplayForReviewThenSend failed"
End Sub

' Shows how to react to specific error types by number.
Sub Example_HandlingSpecificErrors()
    On Error GoTo ErrHandler

    Dim Model As New clsMailModel
    Model.AddToRecipient "not-a-valid-address"   ' deliberately malformed
    modOutlookMailService.SaveAsDraft Model
    Exit Sub

ErrHandler:
    ' Note: constants declared Public in a class module (clsMailModel) are read
    ' through an instance, e.g. Model.ERR_INVALID_EMAIL — not ClassName.CONST —
    ' because the class has no default/predeclared instance.
    Select Case Err.Number
        Case Model.ERR_INVALID_EMAIL
            MsgBox "One of the addresses was invalid: " & Err.Description, vbExclamation
        Case modOutlookMailService.ERR_OUTLOOK_UNAVAILABLE
            MsgBox "Outlook isn't available on this machine.", vbCritical
        Case Else
            MsgBox "Unexpected error " & Err.Number & ": " & Err.Description, vbCritical
    End Select
End Sub

