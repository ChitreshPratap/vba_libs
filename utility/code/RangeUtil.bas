Attribute VB_Name = "RangeUtil"
Private Const moduleName As String = "RangeUtil"

Enum ERangeFormat
    rfText = 1
    rfNumber = 2
    rFGeneral = 3
End Enum

Sub convertToText(uRange As Range)
    'uRange : Range object to convert to text
    'Returns : Nothing
    'It convert the provided range object format to text
    'It reassign also each value convert into text
    'Errors:
    'If Range is nothing then --> Error
    
    Dim tRange As Range
    Dim tEachRange As Range
    On Error GoTo handleError
    Set tRange = uRange
    If uRange Is Nothing Then
        Err.Raise vbObjectError + 1006, moduleName & "_convertToText", "Input range is empty."
    End If
    With tRange
        .NumberFormat = "@"
        'Re-assigning converted TextValue
        For Each tEachRange In tRange.Cells
            If Not IsEmpty(tEachRange) Then
                tEachRange.value = CStr(tEachRange.value)
            End If
        Next tEachRange
    End With
    Exit Sub
handleError:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Function getStartCell(rRange As Range) As Range
    'rRange : Range, to find the start cell
    'Return : Range, the start cell as range object
    'It will return the start cell i.e., top-leftmost cell of the given range object
    Dim reqRange As Range
    Set reqRange = tRange.Cells(1, 1)
    getStartCell = reqRange
End Function

Function getEndCell(rRange As Range) As Range
    'rRange : Range, to find the end/last cell
    'Return : Range, the last/end cell as range object
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

