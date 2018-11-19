Attribute VB_Name = "s_OptionSelection"
Sub OptionSelection()

''OptionSelection(!InputTitleHere): !InputHTMLHere
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
       Set TagorIDElement = aExplorer.document.!InputMostSpecificOptionSelectorHere
        For Each Ele In !InputTagorIDElementWithorWithoutOptionsHere
            !InputOptionforIDorOptionSelectorsHere1
            !InputOptionsforIDorOptionSelectorsHere2
                Ele.Focus
                On Error Resume Next: Ele.FireEvent("!InputJavascriptEventHere"): !InputSelectedIndex Ele.Selected = True: On Error GoTo 0
            End If
        Next Ele
    End If
    Do While aExplorer.Busy
        Application.Wait DateAdd("s", 1, Now)
    Loop

End Sub
