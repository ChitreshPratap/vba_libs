Attribute VB_Name = "Test_Module_ArrayUtil"

Option Explicit


Sub example_to1DArray()
    Dim rngData As Range
    Dim data As Variant
    Dim tArray As Variant
    Dim t1Array As Variant
    Dim outputArray As Variant
    Set rngData = Range("D2:G16")
    data = ArrayUtil.convertRangeToArraySafe(rngData)

    tArray = ArrayUtil.getColumnsFromArray(data, Array(2))
    
    t1Array = ArrayUtil.to1DArray(tArray)

End Sub

Sub example_filterArrayNotLikeCI()
    
    Dim rngData As Range
    Dim data As Variant
    Dim outputArray As Variant
    Set rngData = Range("D2:G16")
    data = ArrayUtil.convertRangeToArraySafe(rngData)

    outputArray = ArrayUtil.filterArrayNotLikeCI(data, 1, Array("del*", "M*rut"))


End Sub


Sub example_filterArrayByPatternCI()
    
    Dim rngData As Range
    Dim data As Variant
    Dim outputArray As Variant
    Set rngData = Range("D2:G16")
    data = ArrayUtil.convertRangeToArraySafe(rngData)

    outputArray = ArrayUtil.filterArrayByPatternCI(data, 3, Array("delhi", "M*rut"))


End Sub


Sub example_filterArrayByValues()
    Dim rngData As Range
    Dim data As Variant
    Dim outputArray As Variant
    Set rngData = Range("D2:G16")
    data = ArrayUtil.convertRangeToArraySafe(rngData)

    outputArray = ArrayUtil.filterArrayByValues(data, 3, Array("delhi", "Meerut"))


End Sub


Sub example_visibleRangeToArray()
    Dim rngData As Range
    Dim data As Variant
    Dim outputArray As Variant
    Set rngData = Range("D2:G16")
    outputArray = ArrayUtil.visibleRangeToArray(rngData)


End Sub

Sub example_excludeRowsByIndex()
    Dim rngData As Range
    Dim data As Variant
    Dim outputArray As Variant
    Set rngData = Range("D2:G16")
    data = ArrayUtil.convertRangeToArraySafe(rngData)

    outputArray = ArrayUtil.excludeRowsByIndex(data, Array(1))


End Sub

Sub example_getUniqueRowsByColumns()
    
    Dim rngData As Range
    Dim data As Variant
    Dim outputArray As Variant
    Set rngData = Range("D2:G16")
    data = ArrayUtil.convertRangeToArraySafe(rngData)
    outputArray = ArrayUtil.getUniqueRowsByColumns(data, Array(1))
    outputArray = ArrayUtil.getColumnsFromArray(outputArray, Array(1))
    
    
End Sub

Sub example_getColumnsFromArray()
    
    Dim rngData As Range
    Dim data As Variant
    Dim outputArray As Variant
    Set rngData = Range("D2:G16")
    data = ArrayUtil.convertRangeToArraySafe(rngData)
    outputArray = ArrayUtil.getColumnsFromArray(data, Array(2))
    
    
End Sub

Sub example_getFilterArray()
    
    Dim rngData As Range
    Dim data As Variant
    Dim outputArray As Variant
    Set rngData = Range("D2:G16")
    data = ArrayUtil.convertRangeToArraySafe(rngData)
    outputArray = ArrayUtil.getFilterArray(data, 3, "Meerut")
    
    
End Sub


Sub example_convertRangeToArraySafe()
    
    Dim rngData As Range
    Dim data As Variant
    
    Set rngData = Range("D2:G16")
    data = ArrayUtil.convertRangeToArraySafe(rngData)
        
    Set rngData = Range("D2")
    data = ArrayUtil.convertRangeToArraySafe(rngData)
        
    Set rngData = Range("D2:D16")
    data = ArrayUtil.convertRangeToArraySafe(rngData)
        
    Set rngData = Range("D2:E3")
    data = ArrayUtil.convertRangeToArraySafe(rngData)
                
    
End Sub

Sub example_writeArrayToRangeSafe()

    Dim rngData As Range
    Dim data As Variant
    Dim outRange As Range
    
    Set rngData = Range("D2:D22")
    data = ArrayUtil.convertRangeToArraySafe(rngData)
    
    Set outRange = ArrayUtil.writeArrayToRangeSafe(Range("K2"), data)
    Debug.Print outRange.Address
End Sub

Sub test_extractRowAs1DArray()
    
    Dim ws As Worksheet
    Dim dRange As Range
    Dim arrVisibleRows As Variant
    Dim colHeaders As Variant
    Set ws = ThisWorkbook.Worksheets("Sheet5")
    Set dRange = ws.Range("D4:U22")
    
    arrVisibleRows = GetVisibleRowsAllColumns_AsArray(dRange)
    colHeaders = extractRowAs_1DArray(arrVisibleRows, 1)

End Sub

Sub test_extractColAs1DArray()
    
    Dim ws As Worksheet
    Dim dRange As Range
    Dim arrVisibleRows As Variant
    Dim colHeaders As Variant
    Set ws = ThisWorkbook.Worksheets("Sheet5")
    Set dRange = ws.Range("D4:U22")
    
    arrVisibleRows = GetVisibleRowsAllColumns_AsArray(dRange)
    colHeaders = extractColumnAs_1DArray(arrVisibleRows, 1)

End Sub

Sub test_getVisibleRowsAllColumns_AsArray()
    
    Dim ws As Worksheet
    Dim dRange As Range
    Set ws = ThisWorkbook.Worksheets("Sheet5")
    Set dRange = ws.Range("D4:U22")
    
    GetVisibleRowsAllColumns_AsArray dRange
    

End Sub

Sub testHJoin()
    
    Dim r As Range
    Dim arr As Variant
    Dim arr2, arr3, arrBlanks, arr4, arr5, arrIdent
    Dim n As Long, i As Long
    Set r = Range("D4:U22")
    
    arr = ArrayUtil.convertRangeToArraySafe(r)
    arr2 = ArrayUtil.convertRangeToArraySafe(Range("Z4:AB22"))
    arr3 = hJoin(arr, arr2)
    
    'Appending Blank column
    n = UBound(arr3, 1) - LBound(arr3, 1) + 1
    ReDim blanks(1 To n, 1 To 1)
    
    arr4 = hJoin(arr3, blanks)
    
    
    'Appending identity column
    ReDim arrIdent(1 To n, 1 To 1)
    For i = LBound(arr3, 1) To n
        arrIdent(i, 1) = i
    Next i
    
    arr5 = hJoin(arr4, arrIdent)
    
End Sub
