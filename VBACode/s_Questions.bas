Attribute VB_Name = "B_Questions"
Sub Questions()

Question = InputBox("Please enter the number below corresponding to your question:" & vbCrLf & _
"1) Can I get a brief overview?" & vbCrLf & _
"2) How can I add/insert/delete rows?" & vbCrLf & _
"3) How can I copy/paste a buttons html?" & vbCrLf & _
"4) I tried to run my code, it wouldn't let me press play." & vbCrLf & vbCrL & _
"5) Other.")

If Question = "1" Or Question = "1)" Then
    MsgBox ("Automating internet explorer requires 5 main functions. Starting Internet Explorer, navigating to web pages, entering text, clicking buttons, and choosing options (ex. dropdown selections). To use this worksheet, you'll need to choose the action in column B corrosponding to one of these main actions. Fill in the columns that get highlighted to the right to complete that code segment, each header has an info box explaining what to do if you need help." & vbCrLf & vbCrLf & "Every 3-4 actions, press the green compiler button and run your macro code to ensure it's operating as expected.")
ElseIf Question = "2" Or Question = "2)" Then
    MsgBox ("Never delete partial rows, never delete or insert columns. Doing so will break vital equations throughout the workbook and your code will not run. That said, to add/insert/delete rows please click the green lock button at the top left of the screen. Do not keep the worksheet unlocked.")
ElseIf Question = "3" Or Question = "3)" Then
    MsgBox ("Hover your mouse over the element you want to alter in your browser, right click --> Inspect element. You may need to do this twice. Once a new box pops up with code, right click the code --> Copy. Back on your worksheet, double click inside the Paste HTML column in your selected actions row and paste the code.")
ElseIf Question = "4" Or Question = "4)" Then
    MsgBox ("This could be for three main reasons. 1) You forgot to select all the necessary inputs for a code. 2) You selected the wrong HTML for the webpage element. 3) The HTML for the webpage element doesn't contain the necessary code to function. If it's option 3 please contact steven.buri@austincsi.com so I can work through the fix for you.")
ElseIf Question = "5" Or Question = "5)" Then
    MsgBox ("Please contact steven.buri@austincsi.com with your workbook attached and a description of the error you're running into.")
End If


End Sub
