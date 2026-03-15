Attribute VB_Name = "ShellUtility"
Option Explicit

Sub executeCommand(commandString As String)
    'It executes the command synchronously
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.Run commandString, 1, True
    
End Sub

Sub executeCommand2(commandString As String)
    'It executes the command asynchronously
    Shell commandString, vbNormalFocus '
End Sub


