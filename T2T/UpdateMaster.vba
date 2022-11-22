Sub Master_update()

Application.ScreenUpdating = False

ThisWorkbook.Worksheets("MASTER").Activate

Call Populate_T2T

Application.Wait (Now + #12:00:01 AM#)

Call Update_Leave

MsgBox "Troops to Task has been updated."

End Sub

