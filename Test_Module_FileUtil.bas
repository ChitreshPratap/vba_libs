Attribute VB_Name = "Test_Module_FileUtil"

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

