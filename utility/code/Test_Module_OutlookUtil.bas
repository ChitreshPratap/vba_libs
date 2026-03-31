Attribute VB_Name = "Test_Module_OutlookUtil"

Sub example_getAccount()
    
    Dim outlookApp As Outlook.Application
    Dim account As Outlook.account
    
    Set outlookApp = OutlookUtility.getOutlookAppObj()
    Set account = OutlookUtility.getAccount(outlookApp, "chitreshpratapsingh20@gmail.com")
    Debug.Print account.CurrentUser

End Sub


Sub example_getOutlookAppObj()
    
    Dim outlookApp As Outlook.Application
    Dim account As Outlook.account
    
    Set outlookApp = OutlookUtility.getOutlookAppObj()
    Set account = OutlookUtility.getAccount(outlookApp, "chitreshpratapsingh20@gmail.com")
    Debug.Print account.CurrentUser

End Sub
