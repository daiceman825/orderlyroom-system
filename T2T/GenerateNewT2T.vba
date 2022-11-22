Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub Create_T2T()

' Check if the sheet already exists
For i = 1 To ThisWorkbook.Worksheets.Count
    If ThisWorkbook.Sheets(i).Name = "Troops to Task" Then
        ' if it does, then exit function and notify user
        MsgBox """Troops to Task"" worksheet already exists!" & vbNewLine & "Rename or Delete the existing Troops to Task worksheet to create a new one."
        GoTo ShtExists
    End If
Next i


Dim StartMonth As String, StartDate As Date, StringDate As String
StartMonth = UCase(Application.InputBox("Please enter the Month that you would like to start with.", "Enter Starting Month", "January"))

' Start at the first day of the Specified Month
StartDate = DateValue("1" & " " & StartMonth & " " & Year(Date))
' make into string
StringDate = StartDate


'######################################################
'# Generate Sheet

Sheets.Add.Name = "Troops to Task"

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Troops to Task")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Administrative headers
ws.Cells(2, 1) = "Platoon"
ws.Cells(2, 2) = "UIC"
ws.Cells(2, 3) = "Rank"
ws.Cells(2, 4) = "Name : Last, First"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' dates 367 days to account for Leap Year plus 1 day
For i = 5 To 373 Step 1 ' 1 to 367 plus 5 because starting column is E
    ws.Cells(2, i) = StartDate + i - 5
    ws.Cells(2, i).BorderAround ColorIndex:=1
Next i

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Input MONTH headers above respective areas on sheet
Dim merge1 As Integer, merge2 As Integer
For j = 1 To 12 Step 1 ' Input Months and Merge cells
    
    ' if the first day of the next month can be found, return the column number
    If (Not IsError(Range("D2:NZ2").Find(DateSerial(Year(StartDate), Month(StringDate) + j, 1), , xlFormulas, , xlByColumns, xlNext, True))) Then
        merge1 = Range("D2:NZ2").Find(DateSerial(Year(StartDate), Month(StringDate) + j - 1, 1), , xlFormulas, , xlByColumns, xlNext, True).Column
        merge2 = Range("D2:NZ2").Find(DateSerial(Year(StartDate), Month(StringDate) + j, 1), , xlFormulas, , xlByColumns, xlNext, True).Column - 1
    Else ' if not found, do nothing
    End If
    
    ' merge cells together and add month as text
    ws.Cells(1, merge1) = UCase(MonthName(Month(DateSerial(Year(StartDate), Month(StringDate) + j - 1, 1))))
    ws.Range(Cells(1, merge1), Cells(1, merge2)).Merge
    ws.Range(Cells(1, merge1), Cells(1, merge2)).HorizontalAlignment = xlCenter
    ws.Range(Cells(1, merge1), Cells(1, merge2)).BorderAround ColorIndex:=1, Weight:=xlThick

Next j


'###############################
'# Apply Conditional Formatting

' Apply Conditional Formatting to the cells using dynamic range
Dim t2tbookrng As Range, endrow As Integer, endcol As Integer
    
endrow = ThisWorkbook.Sheets("Troops to task").Cells(Rows.Count, 4).End(xlUp).Row
endcol = ThisWorkbook.Sheets("Troops to task").Cells(2, Columns.Count).End(xlToLeft).Column
    
Set t2tbookrng = ThisWorkbook.Sheets("Troops to Task").Range(Cells(3, 5), Cells(endrow, endcol))

Call ApplyCFRules(t2tbookrng)

'##############################
'# Input Sample Data
ws.Cells(3, 1).Value = "1st"
ws.Cells(3, 2).Value = "AA"
ws.Cells(3, 3).Value = "RNK"
ws.Cells(3, 4).Value = "DOE, JOHN"

' Format date row to show only DAY while containing the full date
ws.Range("2:2").NumberFormat = "d"
ws.Range("E2:NZ2").HorizontalAlignment = xlCenter
ws.Range("E2:NZ2").ColumnWidth = 5

' Resize Columns A:D
ws.Columns("A:D").AutoFit

' notify user when complete
MsgBox "Troops to Task successfully created! Sample data is provided."


ShtExists:
End Sub

