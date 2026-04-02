Attribute VB_Name = "Test_Module_OutlookUtil"

Sub example_getAccount()
    
    Dim outlookApp As Object 'Outlook.Application
    Dim account As Object   'Outlook.Account
    
    Set outlookApp = OutlookUtility.getOutlookAppObj()
    Set account = OutlookUtility.getAccount(outlookApp, "chitreshpratapsingh20@gmail.com")
    Debug.Print account.CurrentUser

End Sub


Sub example_getOutlookAppObj()
    
    Dim outlookApp As Object 'Outlook.Application
    Dim account As Object 'Outlook.Account
    
    Set outlookApp = OutlookUtility.getOutlookAppObj()
    Set account = OutlookUtility.getAccount(outlookApp, "chitreshpratapsingh20@gmail.com")
    Debug.Print account.CurrentUser

End Sub
