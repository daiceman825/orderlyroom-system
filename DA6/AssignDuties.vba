Sub AssignDuties_DA6()

Dim DutyRng As Range
Set DutyRng = ThisWorkbook.Sheets("DA6").Range("F5:BS10")
Dim cell As Range, row As Range, column As Range

' iterating by columns, then by rows, in order to prevent multiple assignments on one person.
For Each column In DutyRng.Columns
    
    Dim CurrentCol As Integer
    CurrentCol = column.column

    For Each row In column.Rows
        Dim top As String, bottom As String
        Dim tRow As Integer, bRow As Integer, RowCount As Integer
            
        ' get rank requirements to determine what range of cells to use
        fullreq = Split(ActiveSheet.Cells(row.row, 5).Value, "-")
        top = fullreq(0)
        bottom = fullreq(1)
           
        ' get top and bottom rows to be used with CurrentCol to get the "working range"
        tRow = Range("C:C").Find(top, , xlValues, xlWhole, xlByRows, xlNext, True).row
        bRow = Range("C:C").Find(bottom, , xlValues, xlWhole, xlByRows, xlPrevious, True).row
        
        ' get number of rows for use in future loop
        RowCount = bRow - tRow + 1
        
        For Each cell In row.Cells
           If Not cell.Text = "" Then
                
                ' create two collections, one for the counters and one for the addresses of the empty cells
                ' i used this method so that i could find the largest counter and have the address at the same index number in the other collection
                ' ArrayOne.item(i) and ArrayTwo.item(i) give us corresponding Number(1) and Address(2) using the same index(i)
                ' this creates a way to get the address of the largest number in a range
                ' for example : if the largest number is 67 and its index in the number collection is 12, its accompanying address is the 12th item in the address collection
                Dim NumArr As Collection, AddArr As Collection
                Set NumArr = New Collection
                Set AddArr = New Collection
                
                ' iterate through "working range" and get all counter values and addresses of BLANK cells
                For i = 1 To RowCount Step 1
                
                    ' If blank or AI (alternate instructor) they are eligible for duty
                    If (Cells(tRow + i - 1, CurrentCol).Text = "" Or Cells(tRow + i - 1, CurrentCol).Text = "AI") And _
                    Not Cells(tRow + i - 1, CurrentCol + 2).Text = "PI" Then
                    ' PI (primary instructors) cannot have duty the day before instructing class
                   
                        AddArr.Add (Cells(tRow + i - 1, CurrentCol).Address)   ' add the address to the address collection
                        NumArr.Add (Cells(tRow + i - 1, CurrentCol + 1).Value) ' add the number  to the number  collection
                        
                   End If
                Next i
                
                ' loop through the number collection and find the largest value and get its index number
                Dim val As Integer
                val = 1
                For i = 2 To NumArr.Count Step 1
                    If NumArr.Item(i) > NumArr.Item(val) Then
                        val = i
                    End If
                Next i
                
                ' use index number to get the accompanying address from the address collection and put # there
                On Error Resume Next '<- handles unknown error with next line. if # is already in cell put it there anyway
                Range(AddArr.Item(val)).Value = "#"
                
                
           End If
        Next cell
    Next row
Next column

End Sub
