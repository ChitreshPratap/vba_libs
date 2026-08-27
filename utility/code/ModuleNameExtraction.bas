Attribute VB_Name = "ModuleNameExtraction"
'==================================================================================
' modCustomerNameExtract                                            version 2.0
'----------------------------------------------------------------------------------
' PURPOSE : Extract a Customer Name out of noisy Transaction Detail strings by
'           matching against a master list of customer names.
'

' PUBLIC API
'   ExtractCustomerNamesArray(TransactionDetails, CustomerNames, _
'                             [ReturnAllMatches], [MatchCounts])  As Variant
'       -> THE reusable engine. Pass any array / Range / scalar of transaction
'          details plus any array / Range / scalar of customer names.
'          Returns an array of the SAME SHAPE AND BOUNDS as TransactionDetails
'          holding the extracted name ("" where nothing matched).
'          Optional MatchCounts receives a same-shaped array of hit counts.
'
'   MatchArray(...)                 -> same, but against a pre-built index
'   BuildNameIndex(names, dN, dF)   -> build the lookup index once, reuse it
'   ExtractCustomerNames()          -> sheet-driven wrapper (Sheet1!A -> Sheet1!B)
'   ListUnmatchedTransactions()     -> dump the misses to a review sheet
'
' METHOD  : Word-window (n-gram) hashing.
'           A naive "loop every name inside every transaction" is 200,000 x 50,000
'           = 10,000,000,000 InStr calls and will never finish. Instead:
'             1. Normalise every master name         -> Dictionary (O(1) lookup)
'             2. Index the FIRST WORD of every name  -> Dictionary
'                (so a window only ever starts where a name could start)
'             3. Normalise each transaction, split into tokens, probe only the
'                windows beginning on a valid first word.
'           ~10-15 million probes total -> under a minute for 200k x 50k.
'
' RULES   : * Case / punctuation insensitive (both sides normalised).
'           * LONGEST match wins  ("HSR LIMITED" beats "HSR").
'           * Alphanumeric reference codes (FI1234CWG4567) act as barriers, so a
'             name is never matched across a code boundary.
'==================================================================================
Option Explicit
Option Base 0

'--------------------------- SHEET CONFIGURATION ----------------------------------
Private Const TXN_SHEET      As String = "Sheet1"   ' sheet holding transactions
Private Const TXN_COL        As String = "A"        ' transaction details column
Private Const OUT_COL        As String = "B"        ' customer name output column
Private Const DIAG_COL       As String = "C"        ' match-count diagnostics

Private Const NAME_SHEET     As String = "Sheet2"   ' sheet holding master names
Private Const NAME_COL       As String = "J"        ' master name column

Private Const HEADER_ROW     As Long = 1            ' data starts at HEADER_ROW + 1
Private Const CHUNK_ROWS     As Long = 50000        ' rows per read/write batch
Private Const WRITE_DIAG     As Boolean = True      ' write match count to DIAG_COL
Private Const RETURN_ALL     As Boolean = False     ' sub-level default

'--------------------------- ENGINE CONFIGURATION ---------------------------------
Private Const BREAK_ON_CODES As Boolean = True      ' treat ref-codes as barriers
Private Const CODE_MIN_LEN   As Long = 6            ' min length of a "code" token
Private Const MULTI_SEP      As String = " | "      ' separator when ReturnAllMatches
'----------------------------------------------------------------------------------


'==================================================================================
'  ####  THE REUSABLE FUNCTION  ####
'
'  ExtractCustomerNamesArray
'  -------------------------
'  IN  : TransactionDetails  Variant  - Range, 1D array, 2D array or single value
'        CustomerNames       Variant  - Range, 1D array, 2D array or single value
'        ReturnAllMatches    Boolean  - False (default) = longest match only
'                                       True            = every distinct match,
'                                                         joined by MULTI_SEP
'        MatchCounts         Variant  - (optional, ByRef out) same-shaped array of
'                                       how many distinct names each row matched
'  OUT : Variant array, same dimensions and bounds as TransactionDetails
'
'  EXAMPLES
'     Dim res As Variant
'     res = ExtractCustomerNamesArray(Sheet1.Range("A2:A200000"), _
'                                     Sheet2.Range("J2:J50000"))
'     Sheet1.Range("B2:B200000").Value2 = res
'
'     Dim cnt As Variant
'     res = ExtractCustomerNamesArray(arrTxn, arrNames, False, cnt)
'
'     ' also works straight from a worksheet cell (Excel 365 spills it):
'     ' =ExtractCustomerNamesArray(A2:A100, Sheet2!$J$2:$J$50000)
'==================================================================================
Public Function ExtractCustomerNamesArray(ByVal TransactionDetails As Variant, _
                                          ByVal CustomerNames As Variant, _
                                 Optional ByVal ReturnAllMatches As Boolean = False, _
                                 Optional ByRef MatchCounts As Variant) As Variant

    Dim dNames As Object, dFirst As Object

    BuildNameIndex CustomerNames, dNames, dFirst

    ExtractCustomerNamesArray = MatchArray(TransactionDetails, dNames, dFirst, _
                                           ReturnAllMatches, MatchCounts)

    Set dNames = Nothing
    Set dFirst = Nothing

End Function


'==================================================================================
' Worker: match a whole array against an ALREADY-BUILT index.
' Split out so a long run can build the 50k index once and reuse it per chunk.
'==================================================================================
Public Function MatchArray(ByVal TransactionDetails As Variant, _
                           ByVal dNames As Object, _
                           ByVal dFirst As Object, _
                  Optional ByVal ReturnAllMatches As Boolean = False, _
                  Optional ByRef MatchCounts As Variant) As Variant

    Dim vT As Variant, vOut As Variant, vCnt As Variant
    Dim nDims As Long, r As Long, c As Long, hits As Long

    vT = CoerceInput(TransactionDetails)
    nDims = ArrayDims(vT)

    Select Case nDims

        Case 2                                  '--- 2D (a Range is always 2D)
            ReDim vOut(LBound(vT, 1) To UBound(vT, 1), LBound(vT, 2) To UBound(vT, 2))
            ReDim vCnt(LBound(vT, 1) To UBound(vT, 1), LBound(vT, 2) To UBound(vT, 2))
            For r = LBound(vT, 1) To UBound(vT, 1)
                For c = LBound(vT, 2) To UBound(vT, 2)
                    vOut(r, c) = MatchCustomer(SafeStr(vT(r, c)), dNames, dFirst, _
                                               ReturnAllMatches, hits)
                    vCnt(r, c) = hits
                Next c
            Next r

        Case 1                                  '--- 1D
            ReDim vOut(LBound(vT) To UBound(vT))
            ReDim vCnt(LBound(vT) To UBound(vT))
            For r = LBound(vT) To UBound(vT)
                vOut(r) = MatchCustomer(SafeStr(vT(r)), dNames, dFirst, _
                                        ReturnAllMatches, hits)
                vCnt(r) = hits
            Next r

        Case Else                               '--- scalar -> 1x1 2D array
            ReDim vOut(1 To 1, 1 To 1)
            ReDim vCnt(1 To 1, 1 To 1)
            vOut(1, 1) = MatchCustomer(SafeStr(vT), dNames, dFirst, _
                                       ReturnAllMatches, hits)
            vCnt(1, 1) = hits

    End Select

    MatchCounts = vCnt
    MatchArray = vOut

End Function


'==================================================================================
' Build the two lookup dictionaries from any array / Range / scalar of names.
'   dNames : normalised full name -> original name (as typed in Sheet2!J)
'   dFirst : first word           -> max word count of any name starting with it
'==================================================================================
Public Sub BuildNameIndex(ByVal CustomerNames As Variant, _
                          ByRef dNames As Object, _
                          ByRef dFirst As Object)

    Dim vN As Variant
    Dim nDims As Long, r As Long, c As Long

    Set dNames = CreateObject("Scripting.Dictionary")
    Set dFirst = CreateObject("Scripting.Dictionary")
    dNames.CompareMode = vbBinaryCompare        ' both sides already UPPER-cased
    dFirst.CompareMode = vbBinaryCompare

    vN = CoerceInput(CustomerNames)
    
    nDims = ArrayDims(vN)

    Select Case nDims
        Case 2
            For r = LBound(vN, 1) To UBound(vN, 1)
                For c = LBound(vN, 2) To UBound(vN, 2)
                    AddNameToIndex SafeStr(vN(r, c)), dNames, dFirst
                Next c
            Next r
        Case 1
            For r = LBound(vN) To UBound(vN)
                AddNameToIndex SafeStr(vN(r)), dNames, dFirst
            Next r
        Case Else
            AddNameToIndex SafeStr(vN), dNames, dFirst
    End Select

End Sub


Private Sub AddNameToIndex(ByVal sRaw As String, _
                           ByVal dNames As Object, _
                           ByVal dFirst As Object)

    Dim sNorm As String, sFirst As String
    Dim parts() As String, wc As Long

    If Len(sRaw) = 0 Then Exit Sub

    sNorm = NormalizeText(sRaw)
    If Len(sNorm) = 0 Then Exit Sub

    If Not dNames.Exists(sNorm) Then dNames.Add sNorm, sRaw

    parts = Split(sNorm, " ")
    wc = UBound(parts) + 1
    sFirst = parts(0)

    If dFirst.Exists(sFirst) Then
        If dFirst(sFirst) < wc Then dFirst(sFirst) = wc
    Else
        dFirst.Add sFirst, wc
    End If

End Sub


'==================================================================================
' Find the customer name inside ONE transaction string.
'   hitCount returns how many DISTINCT master names were found.
'==================================================================================
Private Function MatchCustomer(ByVal sTxn As String, _
                               ByVal dNames As Object, _
                               ByVal dFirst As Object, _
                               ByVal ReturnAllMatches As Boolean, _
                               ByRef hitCount As Long) As String

    Dim sNorm As String
    Dim tok() As String, isCode() As Boolean
    Dim nTok As Long, j As Long, k As Long, maxW As Long
    Dim cand As String, best As String, bestLen As Long
    Dim allHits As String, skipTok As Boolean

    hitCount = 0
    MatchCustomer = vbNullString

    sNorm = NormalizeText(sTxn)
    If Len(sNorm) = 0 Then Exit Function

    tok = Split(sNorm, " ")
    nTok = UBound(tok)

    If BREAK_ON_CODES Then
        ReDim isCode(0 To nTok)
        For j = 0 To nTok
            isCode(j) = IsCodeToken(tok(j))
        Next j
    End If

    For j = 0 To nTok

        skipTok = False
        If BREAK_ON_CODES Then skipTok = isCode(j)

        If Not skipTok Then
            If dFirst.Exists(tok(j)) Then

                maxW = dFirst(tok(j))
                cand = tok(j)

                For k = 0 To maxW - 1

                    If k > 0 Then
                        If j + k > nTok Then Exit For
                        If BREAK_ON_CODES Then
                            If isCode(j + k) Then Exit For
                        End If
                        cand = cand & " " & tok(j + k)
                    End If

                    If dNames.Exists(cand) Then
                        If InStr(1, MULTI_SEP & allHits & MULTI_SEP, _
                                    MULTI_SEP & dNames(cand) & MULTI_SEP, _
                                    vbBinaryCompare) = 0 Then
                            hitCount = hitCount + 1
                            If Len(allHits) > 0 Then allHits = allHits & MULTI_SEP
                            allHits = allHits & dNames(cand)
                        End If
                        If Len(cand) > bestLen Then
                            bestLen = Len(cand)
                            best = dNames(cand)
                        End If
                    End If

                Next k
            End If
        End If
    Next j

    If ReturnAllMatches Then
        MatchCustomer = allHits
    Else
        MatchCustomer = best
    End If

End Function


'==================================================================================
' SHEET-DRIVEN WRAPPER : Sheet1!A -> Sheet1!B , master list Sheet2!J
'==================================================================================
Public Sub ExtractCustomerNames()

    Dim wsT As Worksheet, wsN As Worksheet
    Dim dNames As Object, dFirst As Object
    Dim lastT As Long, lastN As Long, nRows As Long, nMatched As Long
    Dim firstRow As Long, rowsInChunk As Long, i As Long
    Dim vTxn As Variant, vOut As Variant, vCnt As Variant
    Dim t0 As Double, calcMode As XlCalculation
    Dim sErr As String

    On Error GoTo ErrHandler
    t0 = Timer

    Set wsT = ThisWorkbook.Worksheets(TXN_SHEET)
    Set wsN = ThisWorkbook.Worksheets(NAME_SHEET)

    lastN = wsN.Cells(wsN.Rows.count, NAME_COL).End(xlUp).Row
    lastT = wsT.Cells(wsT.Rows.count, TXN_COL).End(xlUp).Row

    If lastN <= HEADER_ROW Then
        MsgBox "No customer names in " & NAME_SHEET & "!" & NAME_COL, vbExclamation
        Exit Sub
    End If
    If lastT <= HEADER_ROW Then
        MsgBox "No transactions in " & TXN_SHEET & "!" & TXN_COL, vbExclamation
        Exit Sub
    End If

    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Building customer name index ..."

    ' index built ONCE, reused for every chunk
    BuildNameIndex wsN.Cells(HEADER_ROW + 1, NAME_COL).Resize(lastN - HEADER_ROW, 1), _
                   dNames, dFirst

    nRows = lastT - HEADER_ROW
    wsT.Cells(HEADER_ROW + 1, OUT_COL).Resize(nRows, 1).ClearContents
    If WRITE_DIAG Then wsT.Cells(HEADER_ROW + 1, DIAG_COL).Resize(nRows, 1).ClearContents

    firstRow = HEADER_ROW + 1
    Do While firstRow <= lastT

        rowsInChunk = CLng(IIf(firstRow + CHUNK_ROWS - 1 > lastT, _
                               lastT - firstRow + 1, CHUNK_ROWS))

        vTxn = wsT.Cells(firstRow, TXN_COL).Resize(rowsInChunk, 1).Value2
        vOut = MatchArray(vTxn, dNames, dFirst, RETURN_ALL, vCnt)

        For i = LBound(vCnt, 1) To UBound(vCnt, 1)
            If vCnt(i, 1) > 0 Then nMatched = nMatched + 1
        Next i

        wsT.Cells(firstRow, OUT_COL).Resize(rowsInChunk, 1).Value2 = vOut
        If WRITE_DIAG Then _
            wsT.Cells(firstRow, DIAG_COL).Resize(rowsInChunk, 1).Value2 = vCnt

        Application.StatusBar = "Matching ... " & _
            Format$((firstRow - HEADER_ROW) / nRows, "0%") & "   (" & _
            Format$(firstRow - HEADER_ROW, "#,##0") & " of " & _
            Format$(nRows, "#,##0") & ")"

        firstRow = firstRow + rowsInChunk
    Loop

    If WRITE_DIAG Then wsT.Cells(HEADER_ROW, DIAG_COL).Value = "Match Count"

CleanUp:
    Application.StatusBar = False
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Len(sErr) > 0 Then
        MsgBox sErr, vbCritical, "Customer Name Extraction"
    Else
        MsgBox "Done." & vbCrLf & vbCrLf & _
               "Transactions processed : " & Format$(nRows, "#,##0") & vbCrLf & _
               "Matched                : " & Format$(nMatched, "#,##0") & "  (" & _
                                             Format$(nMatched / nRows, "0.0%") & ")" & vbCrLf & _
               "Unmatched              : " & Format$(nRows - nMatched, "#,##0") & vbCrLf & _
               "Master names indexed   : " & Format$(dNames.count, "#,##0") & vbCrLf & _
               "Elapsed                : " & Format$(Timer - t0, "0.0") & " sec", _
               vbInformation, "Customer Name Extraction"
    End If

    Set dNames = Nothing
    Set dFirst = Nothing
    Exit Sub

ErrHandler:
    sErr = "Error " & Err.Number & " in ExtractCustomerNames:" & vbCrLf & Err.Description
    Resume CleanUp
End Sub


'==================================================================================
' HELPERS
'==================================================================================

' Range -> its Value2 ; anything else passes straight through
Private Function CoerceInput(ByVal v As Variant) As Variant

    If IsObject(v) Then
        If TypeName(v) = "Range" Then
            CoerceInput = v.Value2
            Exit Function
        End If
    End If
    CoerceInput = v
End Function


' 0 = not an array (scalar) , 1 = 1D , 2 = 2D , ...
Private Function ArrayDims(ByRef v As Variant) As Long
    Dim i As Long, lb As Long
    If Not IsArray(v) Then Exit Function
    On Error GoTo Done
    Do
        i = i + 1
        lb = LBound(v, i)
    Loop
Done:
    ArrayDims = i - 1
End Function


' Errors / Null / Empty become "" instead of blowing up CStr
Private Function SafeStr(ByVal v As Variant) As String
    If IsObject(v) Then Exit Function
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    SafeStr = CStr(v)
End Function


'----------------------------------------------------------------------------------
' A "code token" = long enough AND contains both letters and digits.
'   FI1234CWG4567 -> True      CONFI23423FI8679 -> True
'   LIMITED       -> False     3M               -> False (too short, kept)
'----------------------------------------------------------------------------------
Private Function IsCodeToken(ByVal t As String) As Boolean

    Dim i As Long, ch As Long
    Dim hasDigit As Boolean, hasAlpha As Boolean

    If Len(t) < CODE_MIN_LEN Then Exit Function

    For i = 1 To Len(t)
        ch = AscW(Mid$(t, i, 1)) And &HFFFF&
        If ch >= 48 And ch <= 57 Then hasDigit = True Else hasAlpha = True
        If hasDigit And hasAlpha Then
            IsCodeToken = True
            Exit Function
        End If
    Next i

End Function


'==================================================================================
' OPTIONAL UTILITY : dump unmatched transactions to a review sheet
'==================================================================================
Public Sub ListUnmatchedTransactions()

    Dim wsT As Worksheet, wsU As Worksheet
    Dim lastT As Long, i As Long, n As Long
    Dim vTxn As Variant, vOut As Variant, res() As Variant

    On Error GoTo ErrHandler
    Set wsT = ThisWorkbook.Worksheets(TXN_SHEET)
    lastT = wsT.Cells(wsT.Rows.count, TXN_COL).End(xlUp).Row
    If lastT <= HEADER_ROW Then Exit Sub

    Application.ScreenUpdating = False

    vTxn = wsT.Cells(HEADER_ROW + 1, TXN_COL).Resize(lastT - HEADER_ROW, 1).Value2
    vOut = wsT.Cells(HEADER_ROW + 1, OUT_COL).Resize(lastT - HEADER_ROW, 1).Value2

    ReDim res(1 To UBound(vTxn, 1), 1 To 2)
    For i = 1 To UBound(vTxn, 1)
        If Len(SafeStr(vOut(i, 1))) = 0 Then
            n = n + 1
            res(n, 1) = HEADER_ROW + i
            res(n, 2) = vTxn(i, 1)
        End If
    Next i

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Unmatched").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler

    Set wsU = ThisWorkbook.Worksheets.Add(After:=wsT)
    wsU.Name = "Unmatched"
    wsU.Range("A1:B1").Value = Array("Source Row", "Transaction Details")
    If n > 0 Then wsU.Range("A2").Resize(n, 2).Value = res
    wsU.Columns("A:B").AutoFit

    Application.ScreenUpdating = True
    MsgBox Format$(n, "#,##0") & " unmatched transaction(s) listed.", vbInformation
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub


'==================================================================================
' QUICK SELF-TEST (run, then look at the Immediate window - Ctrl+G)
'==================================================================================
Public Sub Test_ExtractCustomerNamesArray()

    Dim txn(1 To 3) As Variant
    Dim nam(1 To 4) As Variant
    Dim res As Variant, cnt As Variant, i As Long

    txn(1) = "hsr hsr limited confi23423FI8679 FI1234CWG4567 digsCWG7665"
    txn(2) = "hsr confi23423FI8679 hdfc bank limited FI1234CWG4567 digsCWG7665"
    txn(3) = "zzz9999 nothing here at all"

    nam(1) = "HSR"
    nam(2) = "HSR Limited"
    nam(3) = "HDFC Bank Limited"
    nam(4) = "Digs"

    res = ExtractCustomerNamesArray(txn, nam, False, cnt)

    For i = LBound(res) To UBound(res)
        Debug.Print i & "  [" & res(i) & "]   hits=" & cnt(i)
    Next i
    '  1  [HSR Limited]        hits=2
    '  2  [HDFC Bank Limited]  hits=2
    '  3  []                   hits=0

End Sub


Sub test()
    Dim arrCustName As Variant
    Dim arrRawText As Variant
    Dim arrResult As Variant
    
    arrCustName = ArrayUtil.convertRangeToArraySafe(Sheet7.Range("A1:A8"))
    arrRawText = ArrayUtil.convertRangeToArraySafe(Sheet7.Range("F1:F21"))
    
    arrResult = ModuleNameExtraction.ExtractCustomerNamesArray(arrRawText, arrCustName, False, True)
    
    


End Sub
