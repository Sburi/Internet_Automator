Attribute VB_Name = "C_CompileMacro"
Sub CompileMacro()

'REQUIRED REFERENCES
    'Ensure These References are Checked Under Tools (Above) --> References
    'Microsoft Internet Controls
    'Microsoft HTML Object Library
    'Microsoft Forms 2.0 Object Library

CodeSnippetsColumn = WorksheetFunction.Match("Collects Final Code", Range("3:3"), 0)
CompiledMacro = Join(Application.Transpose(Range(Cells(4, CodeSnippetsColumn), Cells(500, CodeSnippetsColumn))), "")
CompiledMacro2 = "Declare Sub Sleep Lib ""kernel32"" (ByVal dwMilliseconds As Long)" & vbCrLf & vbCrLf & _
"Sub InternetModule() " & vbCrLf & vbCrLf & _
CompiledMacro & _
"Troubleshooter:" & vbCrLf & _
"Call TroubleshootingMacroSub" & vbCrLf & vbCrLf & _
"End Sub"

Dim DataObj As New MSForms.DataObject
DataObj.SetText CompiledMacro2
DataObj.PutInClipboard

    'Creating Help Message
    DirFile = Environ("USERPROFILE") & "\MacroCompiler.txt"
    If Len(Dir(DirFile)) = 0 Then
        If MsgBox("Press Alt-F11 to go to the VBA interface, then paste your text into A_Module1 and press play." & vbCrLf & vbCrLf & "Would you like to see this message again?", vbQuestion + vbYesNo, "Compiler Instructions") = vbYes Then
        'If Yes then do nothing
        Else
            Dim fso As Object
            Set fso = CreateObject("Scripting.FileSystemObject")
            Dim Fileout As Object
            Set Fileout = fso.CreateTextFile(DirFile, True, True)
                Fileout.Write "This file allows users to store preferences for the Internet Automator."
                Fileout.Close
        End If
    End If

End Sub




