Attribute VB_Name = "s_FireJavascriptEvent"
Sub FireJavascriptEvent()

'Fire Javascript Event:
Timecounter = 0
On Error Resume Next
Step!InputStepNumberHere:
    If SelectedElement Is Nothing And Timecounter < MaxTime Then
        Sleep 5
        Timecounter = Timecounter + 5
        Set SelectedElement = aExplorer.document.getElementById("!InputElementIDHere")
        GoTo Step!InputStepNumberHere
    ElseIf Timecounter >= MaxTime Then
        MsgBox ("Couldn't find your element @ " & "Step!InputStepNumberHere"): GoTo Troubleshooter
    ElseIf Not (SelectedElement Is Nothing) Then
        On Error GoTo 0: SelectedElement.FireEvent ("!InputVariableHere")
    End If
    Set SelectedElement = Nothing
        Do While aExplorer.Busy
            Application.Wait DateAdd("s", 1, Now)
        Loop
        
End Sub
