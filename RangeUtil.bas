Attribute VB_Name = "RangeUtil"


Function getStartCell(rRange As Range) As Range
    'It will return the start cell i.e., top-left cell of the given range object
    Dim reqRange As Range
    Set reqRange = tRange.Cells(1, 1)
    getStartCell = reqRange
End Function

Function getEndCell(rRange As Range) As Range
    'It will return the start cell i.e., top-left cell of the given range object
    Dim reqRange As Range
    Set reqRange = tRange.Cells(rRange.Rows.Count, rRange.Columns.Count)
    Set getEndCell = reqRange
End Function

Function getUniqueValues(udRange As Range) As Dictionary
    
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

