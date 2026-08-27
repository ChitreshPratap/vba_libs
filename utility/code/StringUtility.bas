Attribute VB_Name = "StringUtility"
Option Explicit

'----------------------------------------------------------------------------------
' Normalise: UPPER-case, keep A-Z and 0-9 only, everything else -> single space.
'   "M/s. HDFC Bank  Ltd." -> "MS HDFC BANK LTD"
' Writes into a pre-allocated buffer via Mid$ (far faster than string &-building).
'----------------------------------------------------------------------------------
Private Function NormalizeText(ByVal s As String) As String

    Dim i As Long, n As Long, L As Long, ch As Long
    Dim buf As String, atSpace As Boolean

    L = Len(s)
    If L = 0 Then Exit Function

    buf = Space$(L)
    atSpace = True

    For i = 1 To L
        ch = AscW(Mid$(s, i, 1)) And &HFFFF&
        Select Case ch
            Case 48 To 57, 65 To 90                 ' 0-9 , A-Z
                n = n + 1
                Mid$(buf, n, 1) = ChrW$(ch)
                atSpace = False
            Case 97 To 122                          ' a-z -> A-Z
                n = n + 1
                Mid$(buf, n, 1) = ChrW$(ch - 32)
                atSpace = False
            Case Else                               ' separator
                If Not atSpace Then
                    n = n + 1
                    Mid$(buf, n, 1) = " "
                    atSpace = True
                End If
        End Select
    Next i

    If n > 0 Then
        If Mid$(buf, n, 1) = " " Then n = n - 1
    End If

    NormalizeText = Left$(buf, n)

End Function


