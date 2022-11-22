Sub Save_DA6() 

' Make sure that the user wants to save this file
Dim answer
answer = MsgBox("You are about to save this DA6 to archive. " & _
                "Formulas and Macros will no longer work on the saved copy of this sheet. " & _
                "This sheet should only be saved once you are completely done with it, as there is no way to undo this action. " & vbNewLine & vbNewLine & _
                "Are you sure you want to save this sheet?", vbYesNoCancel)

If answer = vbYes Then
    
    ' Duplicate Current Sheet
    ActiveSheet.Copy after:=Worksheets(Sheets.Count - 2)
    sheetname = ActiveSheet.Range("F13").Text & " " & Year(ActiveSheet.Range("F14").Value)
    ActiveSheet.Name = sheetname
    
    ' Change all formulas to their displayed values to save memory and reduce file size
    Application.CutCopyMode = False
    ActiveSheet.Cells.Copy
    ActiveSheet.Cells.PasteSpecial Paste:=xlPasteValues
        
    ' Remove all Buttons and Conditional Formatting
    ActiveSheet.Buttons.Delete
    ActiveSheet.Cells.FormatConditions.Delete
    
    
    
    Dim choice, nextmonth
    nextmonth = UCase(Rows("13").Find(what:="*", after:=Rows("13").Cells(1, 8), LookIn:=xlValues).Value)
    
    choice = MsgBox("Would you like to generate a new DA6 for the next month? (" & nextmonth & ")", vbYesNo)
    
    If choice = vbYes Then
    
        ' Copy day counters over to new DA6
        NumCol1 = Rows("13").Find(what:="*", after:=Rows("13").Cells(1, 8), LookIn:=xlValues).column - 1
        bRow = Range("C200").End(xlUp).row
        tRow = 15
        
        ActiveSheet.Range(Cells(tRow, NumCol1), Cells(bRow, NumCol1)).Copy
        Sheets("DA6").Range("E15").PasteSpecial Paste:=xlPasteValues
        
        ' activate DA6 and clear everything
        Worksheets("DA6").Activate
        Worksheets("DA6").Range("F13").Value = nextmonth
        
        Call FullClear_DA6
    
    ElseIf choice = vbNo Then 
        ' do nothing
    End If
    
ElseIf answer = vbNo Or answer = vbCancel Then 
    ' do nothing
End If

End Sub

