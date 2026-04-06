Attribute VB_Name = "FileUtil"
'Requirement :
'Library : Microsoft Script Runtime


Private Sub TraverseFoldersCollection(ByVal folder As Object, _
                                      ByRef col As Collection)
            
    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        ' Add folder path
        col.Add subFolder.Path
        ' ?? Recursive call (SAFE)
        TraverseFoldersCollection subFolder, col
    Next subFolder

        
End Sub


Sub deleteAllFilesFromFolder(folderPath As String)
    
    Dim fso As FileSystemObject
    Dim fold As folder
    Dim tFile As File
    Set fso = New FileSystemObject
        
    Dim fileCount As Long
    
    Set fold = fso.GetFolder(folderPath)
    fileCount = fold.files.count
    
    For Each tFile In fold.files
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
    For i = subDirs.count To 1 Step -1
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
        fileCount = fileCollection.count
        If fileCount > 1 Then
            If ifMultipleFilesFoundRaiseError Then
                Err.Raise vbObjectError + 2, "FileUtil.getFullFilePathsByPattern", "MultipleFileFoundWithPattern : " & vbNewLine & "File : '" & fullFileName & "' multiple files found"
            End If
        End If
        Set getFullFilePathsByPattern = fileCollection
    End If

End Function

Function getSelectedFolder(Optional titleDialogBox As String = "Select folder", Optional ifNoFolderSelectionRaiseError As Boolean = False) As String
    'titleDialogBox : String, the title box of the popup box
    'ifNoFolderSelectionRaiseError : Boolean,
    '   If True, Raise error If popup box is cancelled or no any folder is selected
    '   if False, Return "", if no any folder is selected
    'Returns : String, The selected folder path
    
    'It shows the popup box to select the folder and
    'returns the selected folder path as String
    'Note:
    'This function not allow to select multiple folder.
    'It only allow to select only single folder
    
    Dim selectedFolderList As Variant
    Dim selectedFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = titleDialogBox
        If .Show = -1 Then
            Set selectedFolderList = .SelectedItems
        Else
            'Dialog is cancelled/Closed
            If ifNoFolderSelectionRaiseError Then
                Err.Raise vbObjectError + 1015, "ArrayUtil_getSelectedFolders", "No folder selected, please select folder : '" & titleDialogBox & "'"
            Else
                Set selectedFolderList = Nothing
            End If
        End If
    End With
    
    If selectedFolderList Is Nothing Then
        selectedFolder = ""
    Else
        selectedFolder = selectedFolderList(1)
    End If
    getSelectedFolder = selectedFolder
End Function

Function getFileNamesInsideFolder(folderPath As String, Optional filePatterns As Variant, _
                                Optional concateFolderPath As Boolean = False) As Scripting.Dictionary
    'It returns the dict
    'Key: 'count' , Value --> total files count
    'Key: 'items' , Value --> The 1D array of file names String
    'FilePatterns --> return the file name of specified pattern
    'if filePatterns not provided then --> search for all file names --> Array("*.*")
    'if no any folder found then --> count:0, items: empty array
    'concateFolderPath, if True then it concate folder path also with file name otherwise only file name
    'if folderPath ='' then Err
    
    Dim fileName As String
    Dim files As Collection
    Dim attr As Long
    Dim resultDict As Object
    Dim totalCount As Long
    Dim i As Long
    
    If IsMissing(filePatterns) Then
        filePatterns = Array("*.*")
    End If
    
    'Set resultDict = CreateObject("Scripting.Dictionary")
    Set resultDict = New Scripting.Dictionary
    
    'Non Empty Validate input
    If Trim(folderPath) = "" Then
        Err.Raise vbObjectError + 1000, "FileUtility_GetFileNames", "Folder path cannot be empty."
    End If
    
    ' Ensure folder path ends with "\"
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Check if folder exists
    On Error Resume Next
    attr = GetAttr(folderPath)
    If Err.Number <> 0 Or (attr And vbDirectory) = 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 1001, "FileUtility_GetFileNames", "Folder does not exist: " & folderPath
    End If
    On Error GoTo 0
    
    Dim fileColl As Dictionary
    Set fileColl = New Dictionary
'    totalCount = 0
    
    'Counting Files Count
    ' Get first file
    Dim ePattern As Variant
    For Each ePattern In filePatterns
        fileName = Dir(folderPath & ePattern)
            
        ' Loop through files
        Do While fileName <> ""
            If Not fileColl.Exists(fileName) Then
                    fileColl.Add fileName, folderPath & fileName
            End If
            fileName = Dir
        Loop
    Next ePattern
    
    totalCount = fileColl.count
    
    If totalCount = 0 Then
        resultDict.Add "count", 0
        resultDict.Add "items", Array()
        Set getFileNamesInsideFolder = resultDict
        Exit Function
    End If
    
    ' Allocate array
    ReDim arr(1 To totalCount)
    
    ' Convert to array
    ReDim arr(1 To totalCount)
    
    i = 1
    Dim ky As Variant
    For Each ky In fileColl.Keys
        If concateFolderPath Then
            arr(i) = fileColl(ky)
        Else
            arr(i) = ky
        End If
        i = i + 1
    Next ky
            
    Dim d As Dictionary
    resultDict.Add "count", totalCount
    resultDict.Add "items", arr
    Set getFileNamesInsideFolder = resultDict

End Function

Function getSubFoldersInsideFolder(folderPath As String, Optional includeSubfoldersRecursively As Boolean = False) As Scripting.Dictionary
    
    'It returns the dict
    'Key: 'count' , Value --> total subfolders count
    'Key: 'items' , Value --> The 1D array of sub-folders String
    'if no any folder found then --> count:0, items: empty array
    'if folderPath ='' then Err
    
    Dim resultDict As Scripting.Dictionary
    Set resultDict = New Scripting.Dictionary
    Dim col As Collection
    Dim arr() As String
    Dim i As Long
    Dim attr As Long
    Dim totalCount As Long
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    ' Validate input
    If Trim(folderPath) = "" Then
        Err.Raise vbObjectError + 5000, "ArrayUtil_getAllSubFoldersInsideFolder", "Folder path cannot be empty."
    End If
    
    ' Ensure trailing "\"
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Check folder exists
    On Error Resume Next
    attr = GetAttr(folderPath)
    If Err.Number <> 0 Or (attr And vbDirectory) = 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 5001, "ArrayUtil_getAllSubFoldersInsideFolders", "Folder does not exist."
    End If
    On Error GoTo 0
            
    Set col = New Collection
    
    If includeSubfoldersRecursively Then
        ' Start recursion
        TraverseFoldersCollection fso.GetFolder(folderPath), col
    Else
        Dim subFolder As Object
        Dim fold As folder
        Set fold = fso.GetFolder(folderPath)
        For Each subFolder In fold.SubFolders
            col.Add subFolder.Path
        Next subFolder
    End If
    
       
    ' If no subfolders ? return empty array
    If col.count = 0 Then
        resultDict.Add "count", 0
        resultDict.Add "items", Array()
        Set getAllSubFoldersInsideFolders = resultDict
        Exit Function
    End If
    
    ' Convert collection ? array
    ReDim arr(1 To col.count)
    
    For i = 1 To col.count
        arr(i) = col(i)
    Next i
    
    resultDict.Add "count", col.count
    resultDict.Add "items", arr
    Set getAllSubFoldersInsideFolder = resultDict
    
End Function


