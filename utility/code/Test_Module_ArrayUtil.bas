Attribute VB_Name = "Test_Module_ArrayUtil"
Option Explicit

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


