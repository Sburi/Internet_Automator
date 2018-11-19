Attribute VB_Name = "s_EnterText"
Sub EnterText()

''EnterText(!InputTitleHere): !InputHTMLHere
Timecounter = 0
On Error Resume Next
Step!InputStepNumberHere:
    If SelectedElement Is Nothing And Timecounter < MaxTime Then
        Sleep 5
        Timecounter = Timecounter + 5
        Set SelectedElement = aExplorer.document.!InputMostSpecificSelectorHere
        GoTo Step!InputStepNumberHere
    ElseIf Timecounter >= MaxTime Then
        MsgBox ("Couldn't find your element @ " & "Step!InputStepNumberHere"): On Error GoTo 0 'GoTo Troubleshooter
    ElseIf Not (SelectedElement Is Nothing) Then
        On Error GoTo 0
        Set TagnameElement = aExplorer.document.getElementsByTagName("!InputTagNameHere")
        For Each Ele In TagnameElement
            If !InputElementSelector(s)Here Then
                Ele.Focus: Ele.Value = !InputTextHere
                On Error Resume Next: Ele.FireEvent ("!InputJavascriptEventHere"): On Error GoTo 0
            End If
        Next Ele
    End If
    Set SelectedElement = Nothing
    Do While aExplorer.Busy
        Application.Wait DateAdd("s", 1, Now)
    Loop

End Sub
