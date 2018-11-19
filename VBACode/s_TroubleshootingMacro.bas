Attribute VB_Name = "E_TroubleshootingMacro"
Sub TroubleshootingMacroSub()

If MsgBox("Did your code run correctly?", vbQuestion + vbYesNo) = vbNo Then
Issue = InputBox("Please enter the number corresponding to your issue:" & vbCrLf & vbCrLf & _
"1) No text was input, no buttons were clicked." & vbCrLf & _
"2) The code ran too quickly, it did not wait between entering text/clicking buttons." & vbCrLf & _
"3) Text was entered but a button was not clicked." & vbCrLf & _
"4) Text was not input." & vbCrLf & _
"5) Other.")
End If
If Issue = "1" Or Issue = "1)" Then
    MsgBox ("I've changed your ""Start Internet Explorer"" action by adding ""Medium"" to the Input Text column. This changes the security settings of Internet Explorer to match local internet settings. Compile the code, paste and run to see if the problem is solved.")
        For Each Cell In Range("B4:B600")
            If Cell.Value = "Start Internet Explorer" Then
                CurrentRow = Cell.Row - 2
                CurrentValue = WorksheetFunction.HLookup("Input                   Text", Range("B3:BZ600"), CurrentRow, 0)
                If CurrentValue = "" Then
                    IEColumn = WorksheetFunction.Match("Input                   Text", Range("B3:BZ3"), 0) + 1
                    Cells(Cell.Row, IEColumn).Value = "Medium"
                Else
                    IEColumn = WorksheetFunction.Match("Input                   Text", Range("B3:BZ3"), 0) + 1
                    Cells(Cell.Row, IEColumn).Value = ""
                End If
            End If
        Next Cell
    ElseIf Issue = "2" Or Issue = "2)" Then
    MsgBox ("This occurs when the code cannot determine whether or not a webpage has fully loaded. To correct this issue, insert a row after button and text actions and select the ""Make Explorer Wait De-looped"" option and input a variable for the number of seconds you want the code to wait.")
    ElseIf Issue = "3" Or Issue = "3)" Then
    MsgBox ("This often occurs because there is an associated javascript event that needs to be fired. Search the HTML of the button for something that starts with ""on"", for instance, ""onchange="". On the macro creation screen, insert a row after your button, then paste the HTML of the button and input the javascript into the variable name, for instance - onchange.")
    ElseIf Issue = "4" Or Issue = "4)" Then
    MsgBox ("This could be because you selected HTML code too far outside of your textbox. Try selecting and copying the code directly inside of the textbox, paste that HTML into the MacroCreator, then choose a new element selector.")
    ElseIf Issue = "5" Or Issue = "5)" Then
    MsgBox ("Please contact steven.buri@austincsi.com with a brief description of your issue and the associated workbook attached. I'll do my best to troubleshoot your issue. Thanks!")
End If


End Sub

