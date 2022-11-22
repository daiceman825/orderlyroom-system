Private Sub AddUser()

Dim fullname As String, rank As String
fullname = "SNUFFY, JAMES"
rank = "CPL"


ThisWorkbook.Sheets("DA6").Activate
Dim ws As Worksheet
Set ws = ThisWorkbook.ActiveSheet

Dim tRow As Integer, bRow As Integer


' create two collections
' one for putting the ranks in the "correct" order
' the other is for keeping track of the range where new names will be inserted
' example: CPT is the top rank on the list and thus the 1st item in the rankslist collection
' the corresponding range is the 1st item in the ranksrng collection, which allows us to reference them BOTH in ONE loop.
' for instance: rankslist(2) and ranksrng(2) will give us "1LT" and its "insert range" where we will insert the names of SMs that have this rank
Dim ranksrng As New Collection, rankslist As New Collection


' create list of approved ranks for DA6
' these ranks are in order of appearance on the DA6
''''''''''''''''''''''''''''
' reorder as needed        '
' Add more ranks as needed '
''''''''''''''''''''''''''''
With rankslist
    .Add "CPT"
    .Add "1LT"
    .Add "2LT"
    .Add "CW3"
    .Add "CW2"
    .Add "WO1"
    .Add "MSG"
    .Add "SFC"
    .Add "SSG"
    .Add "SGT"
    .Add "CPL"
    .Add "SPC"
    .Add "PFC"
    .Add "PV2"
    .Add "PVT"
End With


' missingrow is the default row if someone higher than CPT needs to be inserted
' this is the top row of the DA6 which makes it the "starting point"
missingrow = 15

' iterate through every rank on the list and get the range in which names of this rank can be inserted
For i = 1 To rankslist.Count Step 1

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' if at least one of the rank is currently on the sheet
    If (Not IsError(Application.Match(rankslist(i), ws.Range("C14", "C100"), 0))) Then
        
        ' get the top and bottom row (also works if only 1 row)
        tRow = ws.Range("C14", "C100").Find(rankslist(i), , xlValues, xlWhole, xlByRows, xlNext, True).row
        bRow = ws.Range("C14", "C100").Find(rankslist(i), , xlValues, xlWhole, xlByRows, xlPrevious, True).row
        
        ' add to range array to keep ranks with their related ranges.
        ' ie: CPT and its associated range are index one of both collections
        ranksrng.Add ws.Range("D" & tRow, "D" & bRow).Address
        
        ' set "missingrow" to the row number below the last row of this existing rank range.
        ' if the next rank does not exist on the page, it can be inserted at this row to keep the sheet in the correct order.
        missingrow = bRow + 1
      
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' if the rank does not exist on this sheet
    ElseIf IsError(Application.Match(rankslist(i), ws.Range("C14", "C100"), 0)) Then
    
        ' add a range to the range collection using the "missingrow" variable.
        ' this allows ranks that are not on this sheet to be inserted in the correct order
        ranksrng.Add ws.Range("D" & missingrow, "D" & missingrow).Address
        
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' if the rank of the SM being added equals the rank that the loop is currently on
    ' and this rank exists on this sheet
    If rankslist(i) = rank And (Not IsError(Application.Match(rank, ws.Range("C14", "C100"), 0))) Then

        ' this variable is used to track if the name has been inserted or not
        ' if removed, the macro will keep inputting the name because the cells are getting shifted down when the new line is inserted.
        Insert = False   '<- do not remove this variable
        
        ' iterate through the existing rank range to find where to insert the name alphabetically
        For Each cell In Range(ranksrng(i)).Cells
        
            ' strcomp compares strings and returns an integer
            x = StrComp(fullname, cell.Text)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '  1 is greater than, meaning fullname comes AFTER cell.text
            '  0 is equal to, meaning the two strings are the same
            ' -1 is less than, meaning fullname comes BEFORE cell.text  (this is how we know where to insert the name and keep alphebetical order)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If x < 0 And Insert = False Then
                
                outputaddr = cell.Address
                
                ' insert new row (shifting everything down) and add name, rank, and set counter to 0
                Range(outputaddr).EntireRow.Insert shift:=xlShiftDown
                Insert = True '<- mark the insert variable as true to exit the loop
                ws.Range(outputaddr).Value = fullname
                ws.Range(outputaddr).Offset(0, -1).Value = rank
                ws.Range(outputaddr).Offset(0, 1).Value = 0
                
                
                ' get current row for autofill
                outputrow = Range(outputaddr).row
                
                ' if being inserted on the top row, autofill from below. if not, autofill from above
                If outputrow <= 15 Then
                    ' autofill the days counter from below
                    ws.Range(Cells(outputrow + 1, 6), Cells(outputrow + 1, 70)).AutoFill Destination:=Range(Cells(outputrow + 1, 6), Cells(outputrow, 70))
                ElseIf outputrow > 15 Then
                    ' autofill the days counter from above
                    ws.Range(Cells(outputrow - 1, 6), Cells(outputrow - 1, 70)).AutoFill Destination:=Range(Cells(outputrow - 1, 6), Cells(outputrow, 70))
                End If
            End If
        Next cell
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' if their RANK is NOT currently on this sheet, insert them in the correct spot
    ElseIf rankslist(i) = rank And IsError(Application.Match(rank, ws.Range("C14", "C100"), 0)) Then
    
        
        ' insert new row and add name, rank, and set counter to 0
        ' the "insert" bool is not needed here as there is only one spot to insert if the rank does not exist on the sheet
        Range(ranksrng(i)).EntireRow.Insert shift:=xlShiftDown
        ws.Range(ranksrng(i)).Value = fullname
        ws.Range(ranksrng(i)).Offset(0, -1).Value = rank
        ws.Range(ranksrng(i)).Offset(0, 1).Value = 0
        
        
        ' get current row for autofill
        outputrow = Range(ranksrng(i)).row
        
        ' if being inserted on the top row, autofill from below. if not, autofill from above
        If outputrow <= 15 Then
            ' autofill the days counter from below
            ws.Range(Cells(outputrow + 1, 6), Cells(outputrow + 1, 70)).AutoFill Destination:=Range(Cells(outputrow + 1, 6), Cells(outputrow, 70))
        ElseIf outputrow > 15 Then
            ' autofill the days counter from above
            ws.Range(Cells(outputrow - 1, 6), Cells(outputrow - 1, 70)).AutoFill Destination:=Range(Cells(outputrow - 1, 6), Cells(outputrow, 70))
        End If
    
    End If
    
    
    
Next i

MsgBox rank & " " & fullname & " has been inserted."

End Sub
