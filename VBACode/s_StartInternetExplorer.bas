Attribute VB_Name = "s_StartInternetExplorer"
Sub StartInternetExplorer()

''REQUIRED REFERENCES
    'Ensure These References are Checked Under Tools (Above) --> References
    'Microsoft Internet Controls
    'Microsoft HTML Object Library

'SET INTERNET EXPLORER INSTANCE
    Dim aExplorer As InternetExplorer!InputOnlyTextHere
    Set aExplorer = New InternetExplorer!InputOnlyTextHere

'SET INTERNET EXPLORER VARIABLES
    With aExplorer
    .Visible = True
    .TheaterMode = True
    End With

'SET INTERNET EXPLORER NAVIGATION
    aExplorer.navigate ("!InputDesiredWebsite")
        Application.Wait DateAdd("s", 2, Now)

'SET BASE CODE VARIABLES
    MaxTime = 10000
    Dim SelectedElement As Object

End Sub
