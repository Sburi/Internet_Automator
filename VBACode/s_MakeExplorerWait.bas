Attribute VB_Name = "s_MakeExplorerWait"
Sub MakeExplorerWait()

''Make Explorer Wait
        Application.Wait DateAdd("s", !InputVariableHere, Now)
        Do While aExplorer.Busy
            Application.Wait DateAdd("s", 1, Now)
        Loop
        
End Sub
