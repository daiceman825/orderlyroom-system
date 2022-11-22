Sub adduser()

Dim test As Boolean
For i = 1 To ThisWorkbook.Worksheets.Count
    If ThisWorkbook.Sheets(i).Name = "Troops to Task" Then
        AddUserFormT2T.Show
    Else
    	MsgBox """Troops to Task"" sheet does not exist. There is no place to add entries!"
    End If
Next i

End Sub
