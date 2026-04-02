Attribute VB_Name = "ArrayUtil"
Option Explicit
Private Const moduleName As String = "ArrayUtil"

Function getArrayDimension(arr As Variant) As Variant
    'arr (Variant) - Input array (1D or 2D)
    
    'Returns : Array/Variant (1D array)
            '(1) = Row Count
            '(2) = Column Count
    ' It returns row count and column count of an array
    
    'Errors:
    'Raises error if input is not an array
            
    Dim result(1 To 2) As Long
    Dim rowCount As Long
    Dim columnCount As Long
    
    'Validate array
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 1000, moduleName & "_getDimension", "Input array is not an array."
    End If
    
    On Error GoTo OneDimensional
    ' Try 2D array
    rowCount = UBound(arr, 1) - LBound(arr, 1) + 1
    columnCount = UBound(arr, 2) - LBound(arr, 2) + 1

    GoTo Success
OneDimensional:
    ' Handle 1D array
    Err.Clear
    rowCount = UBound(arr) - LBound(arr) + 1
    columnCount = 1
Success:
    result(1) = rowCount
    result(2) = columnCount
    getArrayDimension = result
    
End Function

Function to1DArray(inputArray As Variant) As Variant
    
    'inputArray : Array() it is two dimension array with single column only
    'Returns : Array(), returns 1D array with one dimension,
    'If there are multiple column in inputArray then Raise Error
    'It converts 2D array with single column (input Array) to 1D array
    
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
    
    ' Get dimensions, totalRows and Colmns in array
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
    'arr : Array(), the source array in which need to remove specified rows
    'rowsToExclude : Array(), 1D array of rows which need to be excluded from the source array (arr)
    'Returns : Array(), the new array object which does not include the rowsToExclude
    'It exclude the specified rows position from the source array and return new array
        
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
    
    'rng : Range, entire range object in which only visible rows to be return as Array
    'Returns : Array,
    'It returns the visible row from autofilter as an Array
    
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
    'arr: Array, the source array, the array from which unique value need to extract
    'keyCols : Array, 1D array of columns which need to be considered to calculate unique values
    'Returns : Array, Entire array of unique value of specified key columns
    
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
    'rng : Range, the range to be converted to array
    'Returns : Array, output array
    
    'It will convert the provided range object to array
    
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
    'startCell : Range, the starting position to write the array
    'arr : Array, the array to write
    'Returns : Range, the range in which array is written
    
    'It will write the provided array at the start cell address and returns the range object
        
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
    'arr: Array, The source array from which the specified column needs to be extract
    'cols: Array (1D), array of column number which need to extract as seperate array, Eg. Array(3,4,1,2)
        'If Array(3) --> It will return the 2D array with only single column, 3rd column of source array
    'It will return the array(2D) with columns specified in same order and with equal number of rows
    'as the source array
    
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
