Attribute VB_Name = "PathUtility"
Option Explicit

Function getEnclosedDoubleQuote(uPath As String) As String
    'It returns the path enclosed in double quotes.
    'The double quotes are
    Dim tGivenString As String
    Dim tStartChar As String
    Dim tEndChar As String
    
    tStartChar = Left(uPath, 1)
    tEndChar = Right(uPath, 1)
    
    If tStartChar <> """" Then
        tGivenString = """" & uPath
    End If
    
    If tEndChar <> """" Then
        tGivenString = tGivenString & """"
    End If
    
    getEnclosedDoubleQuote = tGivenString
    
End Function


