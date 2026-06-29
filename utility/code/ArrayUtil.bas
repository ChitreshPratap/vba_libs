Attribute VB_Name = "ArrayUtil"
Option Explicit

' Joins two arrays side by side into one 2D array.
' Each input may be 1D or 2D. A 1D array is treated as a single column.
' Both arrays must have the same number of rows.
' Result is always a 1-based 2D array.
Public Function hJoin(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long
    Dim lr1 As Long, lc1 As Long, lr2 As Long, lc2 As Long
    Dim dims1 As Integer, dims2 As Integer
    Dim result() As Variant
    Dim i As Long, j As Long

    ' --- Validate inputs are arrays ---
    If Not IsArray(arr1) Then Err.Raise vbObjectError + 500, "HJoin", "First argument is not an array."
    If Not IsArray(arr2) Then Err.Raise vbObjectError + 500, "HJoin", "Second argument is not an array."

    ' --- Determine number of dimensions ---
    dims1 = arrayDims(arr1)
    dims2 = arrayDims(arr2)

    If dims1 = 0 Then Err.Raise vbObjectError + 500, "HJoin", "First array is empty or not initialized."
    If dims2 = 0 Then Err.Raise vbObjectError + 500, "HJoin", "Second array is empty or not initialized."
    If dims1 > 2 Then Err.Raise vbObjectError + 500, "HJoin", "First array has more than 2 dimensions."
    If dims2 > 2 Then Err.Raise vbObjectError + 500, "HJoin", "Second array has more than 2 dimensions."

    ' --- Row/col bounds for array 1 ---
    If dims1 = 1 Then
        lr1 = LBound(arr1): r1 = UBound(arr1) - lr1 + 1
        lc1 = 0: c1 = 1                              ' 1D = one column
    Else
        lr1 = LBound(arr1, 1): r1 = UBound(arr1, 1) - lr1 + 1
        lc1 = LBound(arr1, 2): c1 = UBound(arr1, 2) - lc1 + 1
    End If

    ' --- Row/col bounds for array 2 ---
    If dims2 = 1 Then
        lr2 = LBound(arr2): r2 = UBound(arr2) - lr2 + 1
        lc2 = 0: c2 = 1
    Else
        lr2 = LBound(arr2, 1): r2 = UBound(arr2, 1) - lr2 + 1
        lc2 = LBound(arr2, 2): c2 = UBound(arr2, 2) - lc2 + 1
    End If

    ' --- Guard against empty / mismatched rows ---
    If r1 < 1 Then Err.Raise vbObjectError + 500, "HJoin", "First array has no rows."
    If r2 < 1 Then Err.Raise vbObjectError + 500, "HJoin", "Second array has no rows."
    
    If r1 <> r2 Then
        Err.Raise vbObjectError + 500, "HJoin", _
            "Row count mismatch: array1 has " & r1 & " rows, array2 has " & r2 & " rows."
    End If

    ' --- Build result ---
    ReDim result(1 To r1, 1 To c1 + c2)

    ' Copy array 1 into the left block
    For i = 1 To r1
        If dims1 = 1 Then
            result(i, 1) = arr1(lr1 + i - 1)
        Else
            For j = 1 To c1
                result(i, j) = arr1(lr1 + i - 1, lc1 + j - 1)
            Next j
        End If
    Next i

    ' Copy array 2 into the right block (offset by c1 columns)
    For i = 1 To r2
        If dims2 = 1 Then
            result(i, c1 + 1) = arr2(lr2 + i - 1)
        Else
            For j = 1 To c2
                result(i, c1 + j) = arr2(lr2 + i - 1, lc2 + j - 1)
            Next j
        End If
    Next i

    hJoin = result
End Function

' Returns the number of dimensions of an array (0 if uninitialized).
Private Function arrayDims(ByVal arr As Variant) As Integer
    Dim i As Integer, t As Long
    On Error GoTo Done
    If Not IsArray(arr) Then arrayDims = 0: Exit Function
    Do
        i = i + 1
        t = UBound(arr, i)      ' errors once i exceeds the real dimension count
    Loop
Done:
    arrayDims = i - 1
End Function




Function GetVisibleRowsAllColumns_AsArray(rng As Range) As Variant
    'It will return only visible rows and all columns of defined range
    
    Dim totalRows As Long
    Dim totalCols As Long
    Dim visibleRowCount As Long
    Dim r As Long, c As Long
    Dim outRow As Long
    Dim sourceData As Variant
    Dim outArr() As Variant

    totalRows = rng.Rows.count
    totalCols = rng.Columns.count

    ' Step 1: Count visible rows to size our final array
    visibleRowCount = 0
    For r = 1 To totalRows
        If Not rng.Rows(r).Hidden Then
            visibleRowCount = visibleRowCount + 1
        End If
    Next r

    ' Edge Case: If no rows are visible, return Empty
    If visibleRowCount = 0 Then
        GetVisibleRowsAllColumns_AsArray = Empty
        Exit Function
    End If

    ' Step 2: Dimension the output array (1-based to match standard Excel arrays)
    ReDim outArr(1 To visibleRowCount, 1 To totalCols)

    ' Step 3: Load the entire range into a memory array for maximum speed
    If totalRows = 1 And totalCols = 1 Then
        ReDim sourceData(1 To 1, 1 To 1)
        sourceData(1, 1) = rng.value
    Else
        sourceData = rng.value
    End If

    ' Step 4: Populate the output array
    ' We loop through rows checking visibility, but pull data from the fast memory array
    outRow = 1
    For r = 1 To totalRows
        If Not rng.Rows(r).Hidden Then
            ' If the row is visible, copy ALL columns into the new array
            For c = 1 To totalCols
                outArr(outRow, c) = sourceData(r, c)
            Next c
            outRow = outRow + 1 ' Move to the next slot in the output array
        End If
    Next r

    ' Step 5: Return the populated 2D array
    GetVisibleRowsAllColumns_AsArray = outArr
    
End Function


Function extractColumnAs_1DArray(arr As Variant, colIndex As Long) As Variant
    
    Dim i As Long
    Dim lowerRow As Long
    Dim upperRow As Long
    Dim result() As Variant
    
    ' 1. Check if the input is actually an array
    If Not IsArray(arr) Then
        extractColumnAs_1DArray = Empty
        Exit Function
    End If
    
    ' 2. Check if the requested column is within bounds
    If colIndex < LBound(arr, 2) Or colIndex > UBound(arr, 2) Then
        extractColumnAs_1DArray = Empty
        Exit Function
    End If
    
    ' 3. Get the bounds of the rows
    lowerRow = LBound(arr, 1)
    upperRow = UBound(arr, 1)
    
    ' 4. Dimension the 1D output array to hold all 100,000+ items
    ReDim result(lowerRow To upperRow)
    
    ' 5. Loop through the rows in memory (Lightning fast)
    For i = lowerRow To upperRow
        result(i) = arr(i, colIndex)
    Next i
    
    ' 6. Return the perfectly sized 1D array
    extractColumnAs_1DArray = result
End Function

Function extractRowAs_1DArray(arr As Variant, rowIndex As Long) As Variant
    Dim c As Long
    Dim lowerCol As Long
    Dim upperCol As Long
    Dim result() As Variant
    
    ' 1. Check if the input is actually an array
    If Not IsArray(arr) Then
        extractRowAs_1DArray = Empty
        Exit Function
    End If
    
    ' 2. Check if the requested row is within bounds
    If rowIndex < LBound(arr, 1) Or rowIndex > UBound(arr, 1) Then
        extractRowAs_1DArray = Empty
        Exit Function
    End If
    
    ' 3. Get the bounds of the columns (2nd dimension)
    lowerCol = LBound(arr, 2)
    upperCol = UBound(arr, 2)
    
    ' 4. Dimension the 1D output array to hold the row's columns
    ReDim result(lowerCol To upperCol)
    
    ' 5. Loop through the columns in memory (Executes instantly)
    For c = lowerCol To upperCol
        result(c) = arr(rowIndex, c)
    Next c
    
    ' 6. Return the safely extracted 1D array
    extractRowAs_1DArray = result
End Function
Function to1DArray(inputArray As Variant) As Variant
    
    On Error GoTo ErrorHandler
    
    Dim rowCount As Long
    Dim colCount As Long
    
    Dim i As Long
    
    'validating input Array
    If IsEmpty(inputArray) Then
        Err.Raise vbObjectError + 1000, "ArrayUtil_to1DArray", "Input array is empty."
    End If
    If Not IsArray(inputArray) Then
        Err.Raise vbObjectError + 1000, "ArrayUtil_to1DArray", "Input array is not an array."
    End If
    
    ' Get dimensions
    rowCount = UBound(inputArray, 1) - LBound(inputArray, 1) + 1
    colCount = UBound(inputArray, 2) - LBound(inputArray, 2) + 1
    
    ' Ensure only 1 column
    If colCount <> 1 Then
        Err.Raise vbObjectError + 1002, "Convert2DTo1D", "Array must have exactly one column."
    End If
    
    ' Resize result array
    ReDim result(1 To rowCount)
    
    ' Convert
    For i = 1 To rowCount
        result(i) = inputArray(i, 1)
    Next i
    
    to1DArray = result
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "Convert2DTo1D", "Error converting 2D to 1D: " & Err.Description
End Function


Function filterArrayByPatterns_getLikeNotLikePatterns(arr As Variant, filterCol As Long, patterns As Variant) As Collection

    Dim i As Long, j As Long, k As Long
    Dim rowCount As Long, colCount As Long
    Dim matchArr() As Variant, nonMatchArr() As Variant
    Dim matchCount As Long, nonMatchCount As Long
    Dim val As String, pat As String
    Dim isMatch As Boolean
    
    Dim result As New Collection
    
    On Error GoTo ErrorHandler
    
    ' Validation
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 22000, "ArrauUtil.filterArrayByPatterns_getLikeNotLikePatterns", "Input is not an array"
    End If
    
    If Not IsArray(patterns) Then
        Err.Raise vbObjectError + 22001, "ArrauUtil.filterArrayByPatterns_getLikeNotLikePatterns", "patterns must be an array"
    End If
    
    rowCount = UBound(arr, 1)
    colCount = UBound(arr, 2)
    
    If filterCol < 1 Or filterCol > colCount Then
        Err.Raise vbObjectError + 22002, "FilterArrayByPatternSplitCI", "Invalid column index"
    End If
    
    ' Optimize: convert patterns to lowercase once
    For k = LBound(patterns) To UBound(patterns)
        patterns(k) = LCase(CStr(patterns(k)))
    Next k
    
    ' First pass: count
    For i = 1 To rowCount
        
        val = LCase(CStr(arr(i, filterCol)))
        isMatch = False
        
        For k = LBound(patterns) To UBound(patterns)
            If val Like patterns(k) Then
                isMatch = True
                Exit For
            End If
        Next k
        
        If isMatch Then
            matchCount = matchCount + 1
        Else
            nonMatchCount = nonMatchCount + 1
        End If
        
    Next i
    
    ' Create arrays if needed
    If matchCount > 0 Then ReDim matchArr(1 To matchCount, 1 To colCount)
    If nonMatchCount > 0 Then ReDim nonMatchArr(1 To nonMatchCount, 1 To colCount)
    
    ' Second pass: populate
    Dim mRow As Long, nRow As Long
    mRow = 0: nRow = 0
    
    For i = 1 To rowCount
        
        val = LCase(CStr(arr(i, filterCol)))
        isMatch = False
        
        For k = LBound(patterns) To UBound(patterns)
            If val Like patterns(k) Then
                isMatch = True
                Exit For
            End If
        Next k
        
        If isMatch Then
            
            mRow = mRow + 1
            
            For j = 1 To colCount
                matchArr(mRow, j) = arr(i, j)
            Next j
            
        Else
            
            nRow = nRow + 1
            
            For j = 1 To colCount
                nonMatchArr(nRow, j) = arr(i, j)
            Next j
            
        End If
        
    Next i
    
    ' Add to collection
    result.Add matchArr, "MATCH"
    result.Add nonMatchArr, "NON_MATCH"
    
    Set filterArrayByPatterns_getLikeNotLikePatterns = result
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "ArrauUtil.filterArrayByPatterns_getLikeNotLikePatterns", Err.Description

End Function


Function filterArrayNotLikeCI(arr As Variant, filterCol As Long, patterns As Variant) As Variant

    Dim i As Long, j As Long, k As Long
    Dim rowCount As Long, colCount As Long
    Dim outArr() As Variant
    Dim outRow As Long
    Dim matchFound As Boolean
    Dim val As String, pat As String
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 21000, "ArrayUtil_FilterArrayNotLikeCI", "Input is not an array"
    End If
    
    If Not IsArray(patterns) Then
        Err.Raise vbObjectError + 21001, "ArrayUtil_FilterArrayNotLikeCI", "patterns must be an array"
    End If
    
    rowCount = UBound(arr, 1)
    colCount = UBound(arr, 2)
    
    If filterCol < 1 Or filterCol > colCount Then
        Err.Raise vbObjectError + 21002, "ArrayUtil_FilterArrayNotLikeCI", "Invalid column index"
    End If
    
    ' Pre-convert patterns to lowercase (optimization)
    For k = LBound(patterns) To UBound(patterns)
        patterns(k) = LCase(CStr(patterns(k)))
    Next k
    
    ' First pass: count rows to keep (NOT matching)
    Dim keepCount As Long
    keepCount = 0
    
    For i = 1 To rowCount
        
        val = LCase(CStr(arr(i, filterCol)))
        matchFound = False
        
        For k = LBound(patterns) To UBound(patterns)
            If val Like patterns(k) Then
                matchFound = True
                Exit For
            End If
        Next k
        
        ' Keep only if NOT matched
        If Not matchFound Then keepCount = keepCount + 1
        
    Next i
    
    ' If nothing to keep
    If keepCount = 0 Then
        filterArrayNotLikeCI = Empty
        Exit Function
    End If
    
    ' Create output array
    ReDim outArr(1 To keepCount, 1 To colCount)
    
    ' Second pass: copy rows
    outRow = 0
    
    For i = 1 To rowCount
        
        val = LCase(CStr(arr(i, filterCol)))
        matchFound = False
        
        For k = LBound(patterns) To UBound(patterns)
            If val Like patterns(k) Then
                matchFound = True
                Exit For
            End If
        Next k
        
        If Not matchFound Then
            
            outRow = outRow + 1
            
            For j = 1 To colCount
                outArr(outRow, j) = arr(i, j)
            Next j
            
        End If
        
    Next i
    
    filterArrayNotLikeCI = outArr
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "ArrayUtil_FilterArrayNotLikeCI", Err.Description

End Function

Function filterArrayByPatternCI(arr As Variant, filterCol As Long, patterns As Variant) As Variant
    'It filter the array based on a column
    'return the rows of the column which has same patterns provided
    'Case-Insensitive

    Dim i As Long, j As Long, k As Long
    Dim rowCount As Long, colCount As Long
    Dim outArr() As Variant
    Dim outRow As Long
    Dim matchFound As Boolean
    Dim val As String, pat As String
    
    On Error GoTo ErrorHandler
    
    ' Validation
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 20000, "ArrayUtil_FilterArrayByPatternCI", "Input is not an array"
    End If
    
    If Not IsArray(patterns) Then
        Err.Raise vbObjectError + 20001, "ArrayUtil_FilterArrayByPatternCI", "patterns must be an array"
    End If
    
    rowCount = UBound(arr, 1)
    colCount = UBound(arr, 2)
    
    ' First pass: count matches
    Dim matchCount As Long
    matchCount = 0
    
    For i = 1 To rowCount
        
        val = LCase(CStr(arr(i, filterCol)))
        matchFound = False
        
        For k = LBound(patterns) To UBound(patterns)
            pat = LCase(CStr(patterns(k)))
            
            If val Like pat Then
                matchFound = True
                Exit For
            End If
        Next k
        
        If matchFound Then matchCount = matchCount + 1
        
    Next i
    
    If matchCount = 0 Then
        filterArrayByPatternCI = Empty
        Exit Function
    End If
    
    ' Create output
    ReDim outArr(1 To matchCount, 1 To colCount)
    
    ' Second pass
    outRow = 0
    
    For i = 1 To rowCount
        
        val = LCase(CStr(arr(i, filterCol)))
        For k = LBound(patterns) To UBound(patterns)
            pat = LCase(CStr(patterns(k)))
            If val Like pat Then
                outRow = outRow + 1
                For j = 1 To colCount
                    outArr(outRow, j) = arr(i, j)
                Next j
                Exit For
            End If
        Next k
        
    Next i
    
    filterArrayByPatternCI = outArr
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "ArrayUtil_FilterArrayByPatternCI", Err.Description

End Function
Function filterArrayByValues(arr As Variant, filterCol As Long, filterValues As Variant) As Variant
    'It filter the array based on a column
    'return the rows of the column which has filter values exact match
    'Case-Sensitive
    
    Dim dict As Object
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    Dim outArr() As Variant
    Dim outRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 18000, "ArrayUtil_FilterArrayByValues", "Input data is not an array"
    End If
    
    If Not IsArray(filterValues) Then
        Err.Raise vbObjectError + 18001, "ArrayUtil_FilterArrayByValues", "filterValues must be an array"
    End If
    
    rowCount = UBound(arr, 1)
    colCount = UBound(arr, 2)
    
    If filterCol < 1 Or filterCol > colCount Then
        Err.Raise vbObjectError + 18002, "ArrayUtil_FilterArrayByValues", "Invalid column index"
    End If
    
    ' Store filter values in dictionary (fast lookup)
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = LBound(filterValues) To UBound(filterValues)
        dict(CStr(filterValues(i))) = True
    Next i
    
    ' First pass: count matching rows
    Dim matchCount As Long
    matchCount = 0
    
    For i = 1 To rowCount
    
        If dict.Exists(CStr(arr(i, filterCol))) Then
            matchCount = matchCount + 1
        End If
        
    Next i
    
    ' If no match ? return empty
    If matchCount = 0 Then
        filterArrayByValues = Empty
        Exit Function
    End If
    
    ' Create exact-sized output array
    ReDim outArr(1 To matchCount, 1 To colCount)
    
    ' Second pass: fill data
    outRow = 0
    
    For i = 1 To rowCount
        
        If dict.Exists(CStr(arr(i, filterCol))) Then
            
            outRow = outRow + 1
            
            For j = 1 To colCount
                outArr(outRow, j) = arr(i, j)
            Next j
            
        End If
        
    Next i
    
    filterArrayByValues = outArr
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "ArrayUtil_FilterArrayByValues", Err.Description

End Function

Function excludeRowsByIndex(arr As Variant, rowsToExclude As Variant) As Variant
    'It exclude the specified rows position from the array
    
    Dim dict As Object
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    Dim outArr() As Variant
    Dim outRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 17000, "ArrayUtil_ExcludeRowsByIndex", "Input data is not an array"
    End If
    
    If Not IsArray(rowsToExclude) Then
        Err.Raise vbObjectError + 17001, "ArrayUtil_ExcludeRowsByIndex", "rowsToExclude must be an array"
    End If
    
    rowCount = UBound(arr, 1)
    colCount = UBound(arr, 2)
    
    ' Use dictionary for O(1) lookup
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = LBound(rowsToExclude) To UBound(rowsToExclude)
        If IsNumeric(rowsToExclude(i)) Then
            If rowsToExclude(i) >= 1 And rowsToExclude(i) <= rowCount Then
                dict(rowsToExclude(i)) = True
            End If
        End If
    Next i
    
    ' If nothing to exclude ? return original array
    If dict.count = 0 Then
        excludeRowsByIndex = arr
        Exit Function
    End If
    
    ' Create output array (max possible size first)
    ReDim outArr(1 To rowCount - dict.count, 1 To colCount)
    
    outRow = 0
    
    ' Copy rows except excluded ones
    For i = 1 To rowCount
        
        If Not dict.Exists(i) Then
            
            outRow = outRow + 1
            
            For j = 1 To colCount
                outArr(outRow, j) = arr(i, j)
            Next j
            
        End If
        
    Next i
    
    ' Handle case: all rows excluded
    If outRow = 0 Then
        excludeRowsByIndex = Empty
        Exit Function
    End If
    
    ' Resize array to exact size (if needed)
    If outRow < UBound(outArr, 1) Then
        ReDim Preserve outArr(1 To outRow, 1 To colCount)
    End If
    
    excludeRowsByIndex = outArr
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "ArrayUtil_ExcludeRowsByIndex", Err.Description

End Function


Function visibleRangeToArray(rng As Range) As Variant
    'It returns the visible row from autofilter as an Array
    
    'It will work if the range has only hidden rows.
    'If columns are hidden this function will not work as expected.
    
    Dim visRng As Range, area As Range
    Dim arr As Variant, outArr() As Variant
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    Dim outRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If rng Is Nothing Then
        Err.Raise vbObjectError + 10000, "ArrayUtil_VisibleRangeToArray", "Input range is Nothing"
    End If
    
    ' Get visible cells only
    On Error Resume Next
    Set visRng = rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo ErrorHandler
    
    If visRng Is Nothing Then
        visibleRangeToArray = Empty
        Exit Function
    End If
    
    colCount = rng.Columns.count
    
    ' First pass: count total visible rows
    rowCount = 0
    For Each area In visRng.Areas
        rowCount = rowCount + area.Rows.count
    Next area
    
    If rowCount = 0 Then
        visibleRangeToArray = Empty
        Exit Function
    End If
    
    ' Create exact array
    ReDim outArr(1 To rowCount, 1 To colCount)
    
    ' Second pass: fill data
    outRow = 0
    
    For Each area In visRng.Areas
        
        arr = area.value
        
        ' Handle single cell area
        If area.Cells.count = 1 Then
            outRow = outRow + 1
            outArr(outRow, 1) = arr
        Else
            For i = 1 To UBound(arr, 1)
                outRow = outRow + 1
                
                For j = 1 To colCount
                    outArr(outRow, j) = arr(i, j)
                Next j
                
            Next i
        End If
        
    Next area
    
    visibleRangeToArray = outArr
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "ArrayUtil_VisibleRangeToArray", Err.Description

End Function



Function getUniqueRowsByColumns(arr As Variant, keyCols As Variant) As Variant
    'It returns the unique rows, it consider the specified columns to evaluate unique value
    
    Dim dict As Object
    Dim i As Long, j As Long, k As Long
    Dim key As String
    Dim rowCount As Long, colCount As Long
    Dim outArr() As Variant
    Dim outRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 11000, "ArrayUtil_GetUniqueRowsByColumns", "Input is not an array"
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    rowCount = UBound(arr, 1)
    colCount = UBound(arr, 2)
    
    ' First pass: build unique keys
    For i = 1 To rowCount
        
        key = ""
        
        For k = LBound(keyCols) To UBound(keyCols)
            key = key & "|" & arr(i, keyCols(k))
        Next k
        
        If Not dict.Exists(key) Then
            dict.Add key, i   ' store row index
        End If
        
    Next i
    
    ' Create output array
    ReDim outArr(1 To dict.count, 1 To colCount)
    
    ' Second pass: populate output
    outRow = 0
    
    Dim rowIndex As Variant
    
    For Each rowIndex In dict.Items
        
        outRow = outRow + 1
        
        For j = 1 To colCount
            outArr(outRow, j) = arr(rowIndex, j)
        Next j
        
    Next rowIndex
    
    getUniqueRowsByColumns = outArr
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "ArrayUtil_GetUniqueRowsByColumns", Err.Description

End Function


Function convertRangeToArraySafe(rng As Range) As Variant
    'It will convert the provided range object to array
    'It does not ignore hidden rows/column, it reads every row/column
    'no matter hidden or not
    
    Dim arr As Variant
    
    On Error GoTo ErrorHandler
    
    ' Check if range is nothing
    If rng Is Nothing Then
        Err.Raise vbObjectError + 1000, "ArrayUtil_convertRangeToArraySafe", "Provided range is Nothing"
    End If
    
    ' Check if range has no cells
    If rng.Cells.count = 0 Then
        Err.Raise vbObjectError + 1001, "ArrayUtil_convertRangeToArraySafe", "Input range has no cells"
    End If
    
    ' If single cell, convert to 2D array manually
    If rng.Cells.count = 1 Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = rng.value
        convertRangeToArraySafe = arr
        Exit Function
    End If
    
    ' Convert range to array
    arr = rng.value
    
    convertRangeToArraySafe = arr
    Exit Function

ErrorHandler:
    ' Raise error to calling procedure
    Err.Raise Err.Number, "ArrayUtil_convertRangeToArraySafe", Err.Description
    
End Function

Function writeArrayToRangeSafe(startCell As Range, arr As Variant) As Range
    'It will write the provided array at the start cell address.
    
    Dim numRows As Long
    Dim numCols As Long
    Dim ws As Worksheet
    Dim outputRange As Range
    
    On Error GoTo ErrorHandler
    
'    Validate inputs
'    If ws Is Nothing Then
'        Err.Raise vbObjectError + 2000, "WriteArrayToRangeSafe", "Worksheet is Nothing"
'    End If
    
    If startCell Is Nothing Then
        Err.Raise vbObjectError + 2001, "ArrayUtil_WriteArrayToRangeSafe", "Start cell is Nothing"
    End If
    
    If IsEmpty(arr) Then
        Err.Raise vbObjectError + 2002, "ArrayUtil_WriteArrayToRangeSafe", "Input array is Empty"
    End If
    
    ' Validate array dimensions
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 2003, "ArrayUtil_WriteArrayToRangeSafe", "Input is not an array"
    End If
    
    ' Get array size
    numRows = UBound(arr, 1) - LBound(arr, 1) + 1
    numCols = UBound(arr, 2) - LBound(arr, 2) + 1
    
    ' Write array to sheet
    Set ws = startCell.Parent
    Set outputRange = ws.Range(startCell.Address).Resize(numRows, numCols)
    outputRange.value = arr
    
    Set writeArrayToRangeSafe = outputRange
    Exit Function

ErrorHandler:
    ' Raise error to calling procedure
    Err.Raise Err.Number, "ArrayUtil_WriteArrayToRangeSafe", Err.Description

End Function



Function getFilterArray(arr As Variant, colIndex1 As Long, val1 As Variant, _
                     Optional colIndex2 As Long = 0, Optional val2 As Variant) As Variant

    'It will filter the array on column colIndex1 and ColIndex2
    'If colIndex2 is not provided then only based on colIndex1
    'It will filter the AND operator result on colIndex1 and ColIndex2
    
    Dim i As Long, j As Long
    Dim tempArr() As Variant, outArr() As Variant
    Dim rowCount As Long, colCount As Long
    Dim outRow As Long
    
    On Error GoTo ErrorHandler
    
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 4000, "ArrayUtil_getFilterArray", "Input is not an array"
    End If
    
    rowCount = UBound(arr, 1)
    colCount = UBound(arr, 2)
    
    ' Temporary array (max size)
    ReDim tempArr(1 To rowCount, 1 To colCount)
    
    outRow = 0
    
    ' Filtering
    For i = 1 To rowCount
        
        If colIndex2 = 0 Then
            If arr(i, colIndex1) = val1 Then
                outRow = outRow + 1
                
                For j = 1 To colCount
                    tempArr(outRow, j) = arr(i, j)
                Next j
            End If
        Else
            If arr(i, colIndex1) = val1 And arr(i, colIndex2) = val2 Then
                outRow = outRow + 1
                
                For j = 1 To colCount
                    tempArr(outRow, j) = arr(i, j)
                Next j
            End If
        End If
        
    Next i
    
    ' No match case
    If outRow = 0 Then
        getFilterArray = Empty
        Exit Function
    End If
    
    ' Create final array with exact size
    ReDim outArr(1 To outRow, 1 To colCount)
    
    For i = 1 To outRow
        For j = 1 To colCount
            outArr(i, j) = tempArr(i, j)
        Next j
    Next i
    
    getFilterArray = outArr
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "ArrayUtil_getFilterArray", Err.Description

End Function


Function getColumnsFromArray(arr As Variant, cols As Variant) As Variant
    'It will return the array with columns specified in same order and with equal number of rows
    'Cols is array of columns Eg. Array(3,4,1,2)

    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    Dim outArr() As Variant
    
    On Error GoTo ErrorHandler
    
    ' Validate input array
    If IsEmpty(arr) Then
        Err.Raise vbObjectError + 5000, "ArrayUtil_getColumnsFromArray", "Input is an empty array"
    End If
        
    ' Validate input array
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 5000, "ArrayUtil_getColumnsFromArray", "Input is not an array"
    End If
    
    ' Validate columns input (should be array)
    If Not IsArray(cols) Then
        Err.Raise vbObjectError + 5001, "ArrayUtil_getColumnsFromArray", "Columns parameter must be an array"
    End If
    
    rowCount = UBound(arr, 1)
    colCount = UBound(cols) - LBound(cols) + 1
    
    ' Initialize output array
    ReDim outArr(1 To rowCount, 1 To colCount)
    
    ' Extract required columns
    For i = 1 To rowCount
        For j = 1 To colCount
            
            If cols(LBound(cols) + j - 1) > UBound(arr, 2) Then
                Err.Raise vbObjectError + 5002, "ArrayUtil_getColumnsFromArray", "Column index out of bounds"
            End If
            
            outArr(i, j) = arr(i, cols(LBound(cols) + j - 1))
        
        Next j
    Next i
    
    getColumnsFromArray = outArr
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "ArrayUtil_getColumnsFromArray", Err.Description

End Function
