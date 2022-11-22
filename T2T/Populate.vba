Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Private Function MyFind(aVal As String, bVal As Range) As Long ' finds name(aVal) using range(bVal)
  If (Not IsError(Application.Match(aVal, bVal, 0))) Then
        MyFind = Application.Match(aVal, bVal, 0)
    Else
        MyFind = -1 ' if not found, return -1
    End If
End Function



Sub Populate_T2T()

    On Error Resume Next
    Application.ScreenUpdating = False

    Dim MASTER As Worksheet, STAFF As Worksheet, SCHOOL As Worksheet, DETACHMENT As Worksheet, OFFICE As Worksheet
    ' indentation is for readability only
        Set MASTER = Workbooks("Master T2T").Worksheets("MASTER")
         Set STAFF = Workbooks("Master T2T").Worksheets("STAFF")
        Set SCHOOL = Workbooks("Master T2T").Worksheets("SCHOOL")
    Set DETACHMENT = Workbooks("Master T2T").Worksheets("DETACHMENT")
        Set OFFICE = Workbooks("Master T2T").Worksheets("OFFICE")
       
       
    Dim date1 As String, date2 As String, dateCOL3 As Long, update As Boolean
    
    ' Check if amount of names in list has changed using calculated (1,3) and static (1,4) variables
    If Cells(1, 3).Value = Cells(1, 4).Text Then ' if the list has not changed then continue as normal
    
        update = False
        date1 = Date - 5 ' Get all info from 5 days ago
        date2 = Date + 60 ' and up to 60 days from now
    
    ElseIf Cells(1, 3).Value <> Cells(1, 4).Text Then ' if the list has changed, then clear and re-populate
    
        MASTER.Range("E9", "ZZ250").ClearContents ' Clear the sheet
        Cells(1, 4).Value = Cells(1, 3).Value ' Update the list number value
    
        update = True
        date1 = Date - 40 ' get everything from the past 40 days
        date2 = Date + 60 ' and up to 60 days from now
        
    End If
    
    ' find column of date1 in the Master sheet, which is where the output starts
    dateCOL3 = MASTER.Range("2:2").Find(date1, , xlFormulas, xlWhole, xlByColumns, , True).Column
    
    Dim NameRange As Range, cell As Range
    Set NameRange = MASTER.Range("C3", Cells(150, 3).End(xlUp).Address)
    
    ' clear old comments and highlighted names
    NameRange.Cells.ClearComments
    NameRange.Cells.ClearFormats
    
    For Each cell In NameRange.Cells
        
        Dim name As String, organization As String, SearchRange As Range, DateRange As Range, rowvar As Long, dateCOL1 As Long, dateCOL2 As Long
        
        name = UCase(MASTER.Range(cell.Address).Value & ", " & MASTER.Range(cell.Address).Offset(0, 1).Value)
        organization = UCase(MASTER.Range(cell.Address).Offset(0, -2).Value)
        
        Select Case organization
        Case "BN", "COMPANY"
            organization = "STAFF"             ' If in STAFF
            Set DateRange = STAFF.Range("2:2") ' Look in staff sheet
            Set SearchRange = STAFF.Range("D:D")
        Case "SCHOOL"                            ' if in SCHOOL
            Set DateRange = SCHOOL.Range("2:2")  ' Look in SCHOOL sheet
            Set SearchRange = SCHOOL.Range("D:D")
        Case "EW"
            organization = "DETACHMENT"               ' if in DETACHMENT 
            Set DateRange = DETACHMENT.Range("2:2")   ' Look in DETACHMENT sheet
            Set SearchRange = DETACHMENT.Range("D:D")
        Case "OFFICE"
            Set DateRange = OFFICE.Range("2:2")   ' if in OFFICE
            Set SearchRange = OFFICE.Range("A:A") ' Look in OFFICE sheet
        Case "UNK", "SKIP"
            GoTo NextPerson                    ' skip everyone else
        End Select
        
        ' get row and column values from organization sheet
        rowvar = MyFind(name, SearchRange)
        
        ' find date 1 (Required variable for functionality)
        dateCOL1 = DateRange.Find(date1, , xlValues, xlWhole, xlByColumns, xlNext, True).Column
        
        ' check if the second date can be found, if not then dateCOL1 + 65
        If (Not IsError(DateRange.Find(date2, , xlValues, xlWhole, xlByColumns, xlPrevious, True))) Then
            dateCOL2 = DateRange.Find(date2, , xlValues, xlWhole, xlByColumns, xlPrevious, True).Column
        Else
            dateCOL2 = dateCOL1 + 65
        End If
        
        
        If rowvar > 0 Then ' if the person is on the sheet, continue
            
            Dim tlist As Collection
            Set tlist = New Collection
            
            Dim trange As Range, tcell As Range
            Worksheets(organization).Activate
            Set trange = ActiveSheet.Range(Cells(rowvar, dateCOL1), Cells(rowvar, dateCOL2))
            
            ' add all values in range (above) to an array
            For Each tcell In trange.Cells
                ' if nothing or if weekend, add null value to list
                If tcell.Value = "0" Or tcell.Value = "N" Or tcell.Value = "" Then
                    tlist.Add ""
                Else ' everything else, add as is
                    tlist.Add tcell.Value
                End If
            Next tcell
            
            ' put all array values onto master sheet
            For i = 1 To tlist.Count
            
                ' output to cell address using collection position(i) to shift output location
                ' subtracting 1 because collections start at 1 and we need to start at 0; if not, then output would be shifted right by one cell
                
                ' clear the cell of any previous data
                MASTER.Cells(cell.row, dateCOL3 + i - 1) = ""
                
                ' input new data
                MASTER.Cells(cell.row, dateCOL3 + i - 1) = tlist(i)
            Next i
        
        Else ' if person is not on the sheet, do nothing
        
        cell.Interior.Color = rgbCrimson
        cell.AddComment ("Please verify this individual's organization and update the database.")
        
        End If
        
        
NextPerson:
    Next cell
    
    ThisWorkbook.Worksheets("MASTER").Activate
    'MsgBox "All Troops to Task worksheets have been aggregated."

End Sub

