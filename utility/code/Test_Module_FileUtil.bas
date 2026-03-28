Attribute VB_Name = "Test_Module_FileUtil"

Sub example_getSubFoldersInsideFolder()

    Dim resultDict As Object
    Dim t As Variant
    Dim files As Variant
    Dim folderPath As String
    folderPath = "C:\Users\DELL\OneDrive\Documents"
    
    'Return folder inside folder also
    Set resultDict = FileUtil.getSubFoldersInsideFolder(folderPath, True)
    
    'Return direct folders
   ' Set resultDict = FileUtil.getSubFoldersInsideFolder(folderPath, False)
    files = resultDict("items")

End Sub


Sub example_getFileNamesInsideFolder()

    Dim resultDict As Object
    Dim t As Variant
    Dim files As Variant
    Dim folderPath As String
    folderPath = "C:\Users\DELL\OneDrive\Documents"
    'Returns all files
    Set resultDict = FileUtil.getFileNamesInsideFolder(folderPath)
    'Returns only file names
    Set resultDict = FileUtil.getFileNamesInsideFolder(folderPath, Array("*.acc*", "*.pdf"))
    
    'Set resultDict = FileUtil.getFileNamesInsideFolder(folderPath, Array("*.acc*", "*.pdf"), True)
    files = resultDict("items")
    
End Sub


Sub example_getFullFilePathByPattern()
    
    Dim t As Variant
    Dim fileCollection As Collection
    Dim filePattern As String
    Dim resultedFilePath As String
    filePattern = "C:\Users\pc\Downloads\Dish\Daily*"
    Set fileCollection = FileUtil.getFullFilePathsByPattern(filePattern, False, True)
    For Each t In fileCollection
        Debug.Print t
    Next t


End Sub

Sub example_createFolderPath()
    
    Dim pathToCreate As String
    Dim resultedFilePath As String
    pathToCreate = "C:\Users\pc\Downloads\Dim\Dish\Bish\Topa"
    FileUtil.createFolderPath (pathToCreate)

End Sub

Sub example_getSelectedFolders()
    Dim fold As Variant
    Set fold = FileUtil.getSelectedFolders(True, "Select Math Folder", True)
    

End Sub

Sub example_getSelectedFolder()
    Dim fold As String
    fold = FileUtil.getSelectedFolder("Select Math Folder", True, True)
    

End Sub


