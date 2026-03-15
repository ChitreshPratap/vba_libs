Attribute VB_Name = "Test_Class_TSTimer"

Sub exampleTimer()
    Dim tTimer As TSTimer
    Set tTimer = New TSTimer
    'tTimer.startTimer
    Application.Wait 5000
    tTimer.stopTimer
    tTimer.getTimeElapsedFormatted
End Sub
