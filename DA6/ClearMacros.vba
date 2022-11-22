Sub FullClear_DA6()

Dim Area As Range
Dim row As Range
Set Area = ActiveSheet.Range(ActiveSheet.Cells(15, 4), ActiveSheet.Cells(Rows.Count, 4).End(xlUp))

' Clear out the DA6 area
For Each row In Area.Rows
    For i = 6 To 74 Step 2
        If Not ActiveSheet.Cells(row.row, i) = "" Then
            ActiveSheet.Cells(row.row, i) = ""
        End If
    Next i
Next row

' clear out the selected duty days at the top
For j = 5 To 10 Step 1
    For i = 6 To 74 Step 1
        If Not ActiveSheet.Cells(j, i) = "" Then
            ActiveSheet.Cells(j, i) = ""
        End If
    Next i
Next j

End Sub



Sub Clear_Prompt_DA6()

' turn off screenupdating. for some reason this crashes excel
Application.ScreenUpdating = False

Dim Area1 As Range
Dim row1 As Range
Set Area1 = ActiveSheet.Range(ActiveSheet.Cells(15, 4), ActiveSheet.Cells(Rows.Count, 4).End(xlUp))

prompt = MsgBox("You are about to clear the previously assigned duties." & vbNewLine & "Would you like to clear the rest of the DA6 as well?", vbYesNoCancel, "Clear DA6")

If prompt = vbYes Then

    Call Clear_DA6
    MsgBox ("DA6 cleared.")
    
ElseIf prompt = vbNo Then

    ' Clear out the DA6 area
    For Each row1 In Area1.Rows
        For i = 6 To 74 Step 2
            If ActiveSheet.Cells(row1.row, i).Value = "#" Then
                ActiveSheet.Cells(row1.row, i).Value = ""
            End If
        Next i
    Next row1
    
    MsgBox ("DA6 cleared.")

ElseIf prompt = vbCancel Then 
    'do nothing
Else 
    ' do nothing
End If

' turn screenupdating back on to apply all changes at once. 
Application.ScreenUpdating = True

End Sub



Sub Clear_DA6()

Dim Area As Range
Dim row As Range
Set Area = ActiveSheet.Range(ActiveSheet.Cells(15, 4), ActiveSheet.Cells(Rows.Count, 4).End(xlUp))

' Clear out the DA6 area
For Each row In Area.Rows
    For i = 6 To 74 Step 2
        If Not ActiveSheet.Cells(row.row, i) = "" Then
            ActiveSheet.Cells(row.row, i) = ""
        End If
    Next i
Next row

End Sub
