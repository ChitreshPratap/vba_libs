Attribute VB_Name = "OutlookUtility"
Option Explicit

Function getOutlookAppObj(Optional raiseErrorIfNotInstalled As Boolean = True) As Object
    
    'raiseErrorIfNotInstalled: Boolean = True
    '       If True then it raise error if outlook is not installed in system otherwise returns Nothing
    'returns : outlook.Application
    
    Dim outlookApp As Object
    On Error Resume Next
    Set outlookApp = GetObject(Class:="Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    If (Err.Number <> 0) Or (outlookApp Is Nothing) Then
        If raiseErrorIfNotInstalled Then
            Err.Raise vbObjectError + 2001, "OutlookUtility_getOutlookAppObj", "Outlook is not available. Please ensure outlook is installed and running."
        Else
            Set outlookApp = Nothing
        End If
    End If
    Set getOutlookAppObj = outlookApp
    
End Function

Function getAccount(outlookObj As Object, mailBoxName As String, Optional raiseErrorIfMailBoxNotFound As Boolean = True) As Object
    'outlookObj : Outlook.Application
    'mailBoxName : String , the name of the account added in outlook
    'raiseErrorIfMailBoxNotFound : Boolean = True
    'returns : Outlook.Account
    
    Dim tOutlookApp As Object
    Dim tMailBoxName As String
    Dim resultAccount As Object
    Dim account As Object
    tMailBoxName = mailBoxName
    Set tOutlookApp = outlookObj
    
    For Each account In tOutlookApp.session.Accounts
        If InStr(LCase(account.DisplayName), LCase(tMailBoxName)) > 0 Then
            Set resultAccount = account
        End If
    Next account
    
    If resultAccount Is Nothing Then
        If raiseErrorIfMailBoxNotFound Then
            Err.Raise vbObjectError + 2002, "OutlookUtility_getAccount", "Account : '" & tMailBoxName & "' is not available in outlook. Please ensure specified account is logged-in in outlook."
        Else
            Set resultAccount = Nothing
        End If
        
    End If
    Set getAccount = resultAccount
    
End Function

Function getNewEmailItem(outlookApp As Object) As Object
    
    'outlookApp: Outlook.Application
    'returns: Outlook.mailItem
    Dim mailItem As Object
    Set mailItem = outlookApp.CreateItem(0)
    Set getNewEmailItem = mailItem
        
End Function

Function getEmailStyle() As String
    
    Dim mailStyle As String
    
    mailStyle = "<style>" & _
                "table{ border-collapse: collapse; width : 50%; font-family : Aptos;}" & _
                "th, td { padding : 2px; text-align : left; border : 1px solid #ddd;}" & _
                "th {font-size : 17px;}" & _
                "tr {font-size : 16px;}" & _
                "</style>"
    getEmailStyle = mailStyle

End Function
