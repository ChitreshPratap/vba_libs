Attribute VB_Name = "Test_Module_FilterUtil"
'https://github.com/ricardocamisa?tab=repositories

Sub testcopyPasteFilterData()
    Dim filterDataRange As Range
    Set filterDataRange = ThisWorkbook.Worksheets("Sheet2").Range("F3:G17")
    
    FilterUtil.copyPasteFilteredData filterDataRange, ThisWorkbook.Worksheets("Sheet2").Range("K1")


End Sub

Sub testFilterRowCount()
    
    Dim filterDataRange As Range
    Dim filterRowsCount As Long
    Set filterDataRange = ThisWorkbook.Worksheets("Sheet2").Range("F3:G17")
    filterRowsCount = FilterUtil.getVisibleRowsCount(filterDataRange)
    Debug.Print filterRowsCount
    
End Sub


Sub testFilterByRange()
    Dim filterDataRange As Range
    Dim filterCriteria As Range
    
    Set filterCriteria = ThisWorkbook.Worksheets("Sheet1").Range("F2:F3")
    Set filterDataRange = ThisWorkbook.Worksheets("Sheet2").Range("F3:G17")
    FilterUtil.FilterRecordsByCriteriaRange filterDataRange, filterCriteria, 1


End Sub



Sub testDeleteVisibleRange()
    
    Dim filterDataRange As Range
    Set filterDataRange = ThisWorkbook.Worksheets("Sheet2").Range("A1:B20")
        
    FilterUtil.deleteVisibleData filterDataRange, False
    FilterUtil.deleteVisibleData filterDataRange, True

End Sub
