Attribute VB_Name = "RangeUtil"


Function getStartCell(rRange As Range) As Range
    'Return : Range
    'rRange : Range, to find the start cell
    'It will return the start cell i.e., top-leftmost cell of the given range object
    Dim reqRange As Range
    Set reqRange = tRange.Cells(1, 1)
    getStartCell = reqRange
End Function

Function getEndCell(rRange As Range) As Range
    'It will return the start cell i.e., top-Rightmost cell of the given range object
    Dim reqRange As Range
    Set reqRange = tRange.Cells(rRange.Rows.count, rRange.Columns.count)
    Set getEndCell = reqRange
End Function

Function getUniqueValues(udRange As Range) As Dictionary
    'Return : Dictionary
    'udRange : Range object
    'It will return the unique values from the range object
    'The unique values will be the key of the dictionary
    
    Dim dictUniqueValues As Dictionary
    Set dictUniqueValues = New Dictionary
    
    Dim tempRange As Range
    
    For Each tempRange In udRange
        If Not dictUniqueValues.Exists(CStr(tempRange.value)) And tempRange <> "" Then
            dictUniqueValues.Add CStr(tempRange.value), Nothing
        End If
    Next tempRange
    Set getUniqueValues = dictUniqueValues
End Function

