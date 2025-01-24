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



