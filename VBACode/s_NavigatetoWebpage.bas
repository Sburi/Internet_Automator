Attribute VB_Name = "s_NavigatetoWebpage"
Sub NavigatetoWebpage()

''Navigate to Webpage
    aExplorer.navigate ("!InputDesiredWebsite")
    Do While aExplorer.Busy
        Application.Wait DateAdd("s", 1, Now)
    Loop

End Sub
