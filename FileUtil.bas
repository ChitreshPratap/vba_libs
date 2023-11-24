Attribute VB_Name = "FileUtil"
'Requirement :
'Library : Microsoft Script Runtime

Sub createFolderPath(pathS As String)
'    It creates the given path if given path not exists. If provided path exists then do nothing.
'    String pathS : The path to create.
'    Returns : Nothing

    Dim fso As New FileSystemObject
    Dim subDirs As New Collection
    Dim i As Integer
    Dim subDr As Variant
    Dim subFName As String
    Dim existingPath As String
    Dim nPath As String
    
    nPath = pathS
    If fso.FolderExists(pathS) Then
        Exit Sub
    End If
    Do While nPath <> ""
        subFName = fso.GetBaseName(nPath)
        If fso.FolderExists(nPath) Then
            existingPath = nPath
        End If
        subDirs.Add subFName
        nPath = fso.GetParentFolderName(nPath)
    Loop
    subFName = existingPath
    For i = subDirs.Count To 1 Step -1
        subFName = subFName & "\" & subDirs.Item(i)
        fso.CreateFolder subFName
    Next i
    
End Sub


Function getFullFilePathByPattern(fullFilePathPattern As String, Optional ifNotExistsRaiseError As Boolean = True) As String
    
'    It returns the complete path of the provided path pattern.
'    String fullFilePathPattern : The path pattern to get the full file path.
'    Boolean ifNotExistsRaiseError :
'    Returns String : It returns the complete possible existing path of the given.

    Dim fso As New FileSystemObject
    Dim inputFolderPath As String
    Dim fullFileName As String
    Dim fileName As String
    
    inputFolderPath = fso.GetParentFolderName(fullFilePathPattern)
    fullFileName = fullFilePathPattern
    
    fileName = Dir(fullFileName)
    
    If fileName = "" Then
        
        If ifNotExistsRaiseError Then
            Err.Raise vbObjectError + 2, "FileUtil.getFullFilePathByPattern", "FileNotFoundError : " & vbNewLine & "File : '" & fullFileName & "' not found"
        Else
            getFullFilePathByPattern = fullFileName
        End If
    Else
        getFullFilePathByPattern = inputFolderPath & "\" & fileName
    End If

End Function
