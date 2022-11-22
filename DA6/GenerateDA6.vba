Private Function Find(aVal As String, bVal As String) As Long 
' finds the row of the name in troops to task sheet using first(b) and last(a) name

  Dim maxRow As Long
  maxRow = Workbooks("Master T2T").Sheets("MASTER").Range("C250").End(xlUp).row

  Dim row As Long
  For row = 3 To maxRow ' Troops to task starts on 3rd row
    ' Get first and last names from Troops to Task sheet
    Dim a As String
    Dim b As String
    a = Workbooks("Master T2T").Sheets("MASTER").Cells(row, 3).Text
    b = Workbooks("Master T2T").Sheets("MASTER").Cells(row, 4).Text
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
Private Function DateFind(inDate As Variant, FindRange As Range) As Long
    If inDate = "" Then
        DateFind = -1
    ElseIf (Not IsError(Application.Match(CLng(CDate(inDate)), FindRange, 0))) Then
        DateFind = Application.Match(CLng(CDate(inDate)), FindRange, 0)
    Else
        DateFind = -1
    End If
End Function



Sub FillData_DA6()

Application.ScreenUpdating = False

Call Clear_DA6

' note: the master troops to task also needs to be open in excel when you run this macro.
Dim T2T As Worksheet
Set T2T = Workbooks("Master T2T").Sheets("MASTER")

' set shortcut to Current DA6 sheet
Dim DA6 As Worksheet
Set DA6 = ThisWorkbook.ActiveSheet

' declare range variable for foreach loop
Dim NameRng As Range
Dim cell As Range

' find the end of the list of names and get the row number
Set NameRng = DA6.Range(DA6.Cells(15, 4), DA6.Cells(Rows.Count, 4).End(xlUp))

' for each name on the DA6
For Each cell In NameRng.Cells
    
    ' get cell address
    Dim celladdr As String
    celladdr = cell.Address
    Dim cellrow As Integer
    cellrow = cell.row
    
    ' get first and last names
    Dim lastn As String, firstn As String
    
    fulln = Split(cell.Text, ", ")
    lastn = UCase(fulln(0))
    firstn = UCase(fulln(1))
    
    ' find what row the name is on in the Troops to Task workbook
    Dim t2trow As Long
    t2trow = Find(lastn, firstn)
    
    ' if the name cannot be found on the Troops to Task, skip it
    If t2trow > 0 Then
    
        For i = 6 To 74 ' in the DATE row
            If i Mod 2 = 0 Then ' skip every other column (the ones containing the counters)
            
                Dim dadate As String ' get date from DA6
                dadate = DA6.Cells(14, i).Value
            
                Dim t2tcol As Integer
                t2tcol = DateFind(dadate, T2T.Range("2:2"))
                
                If t2tcol > 0 Then
                
                    Dim tcell As String
                    tcell = T2T.Cells(t2trow, t2tcol).Value
                     
                    ' if nothing or if weekend, put nothing
                    If tcell = "" Or tcell = "N" Then
                        DA6.Cells(cellrow, i).Value = ""
                    ' if Leave, Pass, or TDY put an "A" to annotate ABSENT
                    ElseIf tcell = "P" Or tcell = "L" Or tcell = "Y" Then
                        DA6.Cells(cellrow, i).Value = "A"
                    ' if any other value is in the cell, put it on the DA6 as-is
                    Else
                        DA6.Cells(cellrow, i).Value = tcell
                    End If
                Else 'do nothing
                End If
                
            End If
        Next i ' next column
        
    Else ' do nothing
        'cell.Text.Color = rgbCrimson ' optional : Mark name RED
    End If
 
Next cell

DA6.Activate
MsgBox ("Your DA6 has been generated!")

End Sub 

