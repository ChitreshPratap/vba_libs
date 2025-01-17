Attribute VB_Name = "FilterUtil"


Function getVisibleRowsCount(filterRange_Data As Range) As Long
    'It returns the number of visible rows in filter
    'It will include header in count
    
    Dim visibleRange As Range
    Dim visibleRowsCount As Long
    Dim ar As Variant
    visibleRowsCount = 0
    Set visibleRange = filterRange_Data.SpecialCells(xlCellTypeVisible)
    
    For Each ar In visibleRange.Areas
        visibleRowsCount = visibleRowsCount + ar.Rows.Count
    Next ar
    getVisibleRowsCount = visibleRowsCount
    
End Function


Sub copyPasteFilteredData(filterRange_Data As Range, targetRange As Range)

    Dim visibleRowsCount As Long
    
    visibleRowsCount = FilterUtil.getVisibleRowsCount(filterRange_Data)
    If visibleRowsCount > 1 Then
        filterRange_Data.SpecialCells(xlCellTypeVisible).Copy
        targetRange.PasteSpecial
        Application.CutCopyMode = False
    End If
    
End Sub


Sub FilterRecordsByCriteriaRange(filterRange_Data As Range, criteriaRange As Range, fieldNumber As Long)
    'It will filter the given filterRange_Data with the provided criteria Range
    'Filter range is the data range including headers
    'Criteria Range : The range object having the values to filter. It is 1D range object
    'Field Number it is the column number on which the filter needs to be implemented
    'Field number reference will be from the fieldRage_Data  i.e., FieldRange_Data column 1 is Field number 1

    Dim wsFilterData As Worksheet
    Set wsFilterData = filterRange_Data.Parent
    Dim criteriaArr() As Variant
    
    If criteriaRange.Cells.Count = 1 Then
        criteriaArr = Array(criteriaRange.value)
    Else
        criteriaArr = Application.Transpose(criteriaRange)
    End If
    
    If wsFilterData.AutoFilterMode Then
        wsFilterData.AutoFilterMode = False
    End If
    filterRange_Data.AutoFilter Field:=fieldNumber, Criteria1:=criteriaArr, Operator:=xlFilterValues
        
End Sub




Sub deleteVisibleData(filterRange_Data As Range, Optional deleteHeader As Boolean = False)

    Dim visibleRowsCount As Long
    visibleRowsCount = filterRange_Data.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count
    
    If visibleRowsCount > 1 Then
        If deleteHeader Then
            filterRange_Data.SpecialCells(xlCellTypeVisible).EntireRow.Delete
        Else
            filterRange_Data.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
    End If

End Sub

