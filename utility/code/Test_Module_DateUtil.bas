Attribute VB_Name = "Test_Module_DateUtil"
Sub example_getQuarterNumber()
    
    Dim userDate As Date
    Dim resultQuarterNumber As Integer
    userDate = CDate("3/18/2023")
    resultQuarterNumber = DateUtil.getQuarterNumber(userDate)
    Debug.Print resultQuarterNumber
    
End Sub

Sub example_getFormattedString()
    Dim userDate As Date
    Dim strFormat As String
    Dim resultStrFormat As String
    strFormat = "Hello, Date : '%MM-DD-YYYY%' is day: %DD% , Month : %MM% , Year : %YY%"
    userDate = CDate("3/18/2023")
    
    resultStrFormat = DateUtil.getFormattedString(userDate, strFormat)
    Debug.Print resultStrFormat
    
End Sub

Sub example_getLastDateOfQuarter()

    Dim lastDateOfQuarter As Date
    lastDateOfQuarter = DateUtil.getLastDateOfQuarter(2023, 6)
    Debug.Print lastDateOfQuarter

End Sub
