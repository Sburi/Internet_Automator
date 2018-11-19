Attribute VB_Name = "F_Lock_Unlock_Workbook"
Sub Lock_UnlockButton()

    If Worksheets("MacroCreation").ProtectContents = True Then
        Worksheets("MacroCreation").Unprotect
        Worksheets("MacroCreation").Shapes("B_LockIndicator").Fill.ForeColor.RGB = RGB(255, 0, 0)
        MsgBox ("Warning, you've just unprotected the worksheet, click the lock button again to lock the worksheet." & vbCrLf & vbCrLf & _
        "Unlocking the worksheet allows you to insert/delete rows. However, you can accidentally delete essential equations in hidden columns of the worksheet. Do not keep the worksheet unlocked unless you're an experienced Excel user.")
    ElseIf Worksheets("MacroCreation").ProtectContents = False Then
        Worksheets("MacroCreation").Protect Password:="", _
            DrawingObjects:=True, _
            Contents:=True, _
            Scenarios:=True, _
            UserInterfaceOnly:=True, _
            AllowFormattingCells:=True, _
            AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, _
            AllowInsertingColumns:=False, _
            AllowInsertingRows:=False, _
            AllowInsertingHyperlinks:=True, _
            AllowDeletingColumns:=False, _
            AllowDeletingRows:=False, _
            AllowSorting:=False, _
            AllowFiltering:=False, _
            AllowUsingPivotTables:=False
            Worksheets("MacroCreation").Shapes("B_LockIndicator").Fill.ForeColor.RGB = RGB(0, 176, 80)
    End If

End Sub
