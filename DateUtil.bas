Attribute VB_Name = "DateUtil"

Function getFormattedString(fDate As Date, stringToFormat As String) As String
    'It returns the formatted string and put the time values of given fDate in the formatted string
    'Date parts symbols must be enclosed inside % %
    'Example: It is %YYY% Year ==> Output : It is 2023 Year
    
    Dim formattedString As String
    Dim splittedString As Variant
    Dim resultString As String
    Dim tempString As String
    Dim i As Integer
    
    splittedString = Split(stringToFormat, "%")
    
    For i = LBound(splittedString) To UBound(splittedString)
        tempString = splittedString(i)
        If i Mod 2 <> 0 Then
            tempString = splittedString(i)
            tempString = Format(fDate, tempString)
        End If
        resultString = resultString & tempString
    Next i
    getFormattedString = resultString
End Function


Function getQuarterNumber(iDate As Date) As Integer
    'It returns quarter number of the given input date
    
    Dim resultQtr As Integer
    If Month(iDate) <= 3 Then
        resultQtr = 1
    ElseIf Month(iDate) <= 6 Then
        resultQtr = 2
    ElseIf Month(iDate) <= 9 Then
        resultQtr = 3
    ElseIf Month(iDate) <= 12 Then
        resultQtr = 3
    End If
    getQuarterNumber = resultQtr
    
End Function

Function getLastDateOfQuarter(iYear As Integer, iQuarterNumber As Integer) As Date
    'It returns last date of the given year and quarter number
    
    Dim resultDate As Integer
    Dim tempDate As Date
        
    If iQuarterNumber = 1 Then
        tempDate = CDate("3/26/" & CStr(iYear))
    ElseIf iQuarterNumber = 2 Then
        tempDate = CDate("6/26/" & CStr(iYear))
    ElseIf iQuarterNumber = 3 Then
        tempDate = CDate("9/26/" & CStr(iYear))
    ElseIf iQuarterNumber = 4 Then
        tempDate = CDate("12/26/" & CStr(iYear))
    Else
        Err.Raise Err.Number + 2, "DateUtil.getLastDateOfQuarter", "ValueError : Invalid Quarter number, it can be 1,2,3,4 but provided : " & CStr(iQuarterNumber)
    End If
    resultDate = WorksheetFunction.EoMonth(tempDate)
    getLastDateOfQuarter = resultDate
    
End Function

