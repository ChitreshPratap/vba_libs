Attribute VB_Name = "Test_Module_RangeUtil"



Sub example_uniqueValues()
    
    Dim allValues As Range
    Dim uniqueValues As Variant
    Dim k As Variant
    Set allValues = ThisWorkbook.Worksheets("Sheet1").Range("G7:G18")
    Set uniqueValues = RangeUtil.getUniqueValues(allValues)
    For Each k In uniqueValues
        Debug.Print k
    Next k
End Sub

