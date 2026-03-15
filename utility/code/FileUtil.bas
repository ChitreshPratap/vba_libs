Attribute VB_Name = "FileUtil"
'Requirement :
'Library : Microsoft Script Runtime

Sub deleteAllFilesFromFolder(folderPath As String)
    
    Dim fso As FileSystemObject
    Dim fold As Folder
    Dim tFile As File
    Set fso = New FileSystemObject
        
    Dim fileCount As Long
    
    fold = fso.GetFolder(folderPath)
    fileCount = fold.Files.Count
    
    For Each tFile In fold.Files
        tFile.Delete
    Next tFile
    
    Set tFile = Nothing
    Set fold = Nothing
    Set fso = Nothing
    
End Sub

Function fileExists(filePath As String, ifNotExistsRaiseError As Boolean) As Boolean
    Dim fso As New FileSystemObject
    
    If fso.fileExists(filePath) Then
        fileExists = True
    Else
        If ifNotExistsRaiseError Then
            Err.Raise vbObjectError + 2, "FileUtil.fileExists", "FileNotFoundError : " & vbNewLine & "File : '" & filePath & "' not found"
        Else
            fileExists = False
        End If
    End If
End Function


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
            Exit Do
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


Function getFullFilePathsByPattern(fullFilePathPattern As String, Optional ifNotExistsRaiseError As Boolean = True, Optional ifMultipleFilesFoundRaiseError As Boolean = True) As Collection
    
'    It returns the collection of complete path of the provided file path pattern. The pattern is allowed in the file name only.
'    In the folder path patterns are not allowed. The folder in fullFilePathPattern must be without pattern.

'    String fullFilePathPattern : The file name pattern to get the full file path. The folder name must not include patterns otherwise : Error : Bad file name
'    Boolean ifNotExistsRaiseError :
'    Boolean ifMultipleFilesFoundRaiseError:
'    Returns Collection : It returns collection of the complete possible existing path of the given pattern.

    Dim fileCollection As Collection
    Set fileCollection = New Collection
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
            Set getFullFilePathsByPattern = fileCollection
        End If
    Else
        Do While fileName <> ""
            fileCollection.Add inputFolderPath & "\" & fileName
            fileName = Dir
        Loop
        Dim fileCount As Long
        fileCount = fileCollection.Count
        If fileCount > 1 Then
            If ifMultipleFilesFoundRaiseError Then
                Err.Raise vbObjectError + 2, "FileUtil.getFullFilePathsByPattern", "MultipleFileFoundWithPattern : " & vbNewLine & "File : '" & fullFileName & "' multiple files found"
            End If
        End If
        Set getFullFilePathsByPattern = fileCollection
    End If

End Function

Function getSelectedFolders(Optional titleDialogBox As String = "Select folder", Optional allowMultipleSelect As Boolean = False) As Variant
    Dim selectedFolderList As Variant
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = titleDialogBox
        .allowMultiSelect = allowMultipleSelect
        If .Show = -1 Then
            Set selectedFolderList = .SelectedItems
        End If
    End With
    Set getSelectedFolders = selectedFolderList
End Function

Function getSelectedFolder(Optional titleDialogBox As String = "Select folder", Optional allowMultipleSelect As Boolean = False) As String
    
    Dim selectedFolders As Variant
    Dim selectedFolder As String
    Set selectedFolders = getSelectedFolders(titleDialogBox, allowMultipleSelect)
    selectedFolder = selectedFolders(1)
    getSelectedFolder = selectedFolder
End Function


