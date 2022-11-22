Private Function Find(aVal As String, bVal As String) As Long ' finds the row of the name in troops to task sheet using first(b) and last(a) name

  ' look in the 3rd(c) column
  Dim maxRow As Long
  maxRow = Range("C200").End(xlUp).row
  
  ' for each row in that column
  Dim row As Long
  For row = 3 To maxRow
    ' Get first and last names from Troops to Task sheet
    Dim a As String
    Dim b As String
    a = ThisWorkbook.Sheets("MASTER").Cells(row, 3).Text
    b = ThisWorkbook.Sheets("MASTER").Cells(row, 4).Text
    ' if the first and last name from Troops to task matches the first and last names from import sheet
    If a = aVal And b = bVal Then
      ' return the row number and exit the function
      Find = row
      Exit Function
    End If
  Next row
  ' if the first and last name is not in the Troops to task list, return a -1
  Find = -1
End Function



Sub Update_Leave()
    ' These variables are used in the for each loop
    Dim rng As Range
    Dim cell As Range
    Dim Tbl As ListObject
    Dim lastRow As Long
    Set Tbl = ThisWorkbook.Worksheets("hidden import sheet").ListObjects("Import_Data")
    lastRow = Tbl.ListColumns(1).Range.Rows.Count
    Set rng = Worksheets("hidden import sheet").Range("B2" & ":B" & lastRow)
    rng.ClearFormats
    
    ThisWorkbook.Worksheets("MASTER").Activate
    
    For Each cell In rng.Cells
     
         Dim addrcell As String
         addrcell = cell.Address(False, False, xlA1)
         
         ' Get dates: From(one) & To(two)
         date1 = Sheets("hidden import sheet").Range(addrcell).Offset(0, 1).Value
         date2 = Sheets("hidden import sheet").Range(addrcell).Offset(0, 2).Value
         Var_Col_one = ThisWorkbook.Sheets("MASTER").Range("2:2").Find(date1, , xlFormulas, xlWhole, xlByColumns, , True).Column
         Var_Col_two = ThisWorkbook.Sheets("MASTER").Range("2:2").Find(date2, , xlFormulas, xlWhole, xlByColumns, , True).Column
         
         ' Find the name in Troops to Task
         Dim fname As String, lname As String
         lname = UCase(Sheets("hidden import sheet").Range(addrcell).Text)
         fname = UCase(Sheets("hidden import sheet").Range(addrcell).Offset(0, -1).Text)
         t2trow = Find(lname, fname)
         
         ' if the row number is returned, continue
         If (t2trow > 0) Then
             
             ' Get type of leave
             xlString = Worksheets("hidden import sheet").Range(addrcell).Offset(0, 4).Text
         
             ' If TDY, mark as Y. If anything else, mark with first letter. ( L or P )
             If xlString = "TDY" Then
                 xlFirstChar = Right$(xlString, 1)
             Else:
                 xlFirstChar = Left$(xlString, 1)
             End If
         
             ' Mark days in all cells between Date1 to Date2
             Worksheets("MASTER").Range(Cells(t2trow, Var_Col_one), Cells(t2trow, Var_Col_two)).Value = xlFirstChar
         
         ' if not found (-1), highlight cell and move to next
         Else
             cell.Interior.Color = rgbCrimson
         End If
       
     Next cell

    'MsgBox ("Troops to Task has been successfully updated.")
End Sub
