Attribute VB_Name = "Test_Module_FileUtil"

Sub example_getFullFilePathByPattern()
    
    Dim filePattern As String
    Dim resultedFilePath As String
    filePattern = "C:\Users\pc\Downloads\Dish\Manager Data*Nov.xlsx"
    resultedFilePath = FileUtil.getFullFilePathByPattern(filePattern, True)
    Debug.Print resultedFilePath

End Sub

Sub example_createFolderPath()
    
    Dim pathToCreate As String
    Dim resultedFilePath As String
    pathToCreate = "C:\Users\pc\Downloads\Dim\Dish\Bish\Topa"
    FileUtil.createFolderPath (pathToCreate)

End Sub

