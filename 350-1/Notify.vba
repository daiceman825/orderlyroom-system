Sub SendMessage()

Dim prompt As VbMsgBoxResult
prompt = MsgBox("You are about to notify all Delinquent SM's on this sheet of their 350-1 Status. This document must be opened from Desktop and Outlook must be open for successful completion." & vbNewLine & vbNewLine & "Are you sure you want to continue?", _
                vbYesNoCancel, "Delinquency Notification")

If prompt = vbYes Then

    Dim LNrange As Range
    Dim cellone As Range
    
    ' set range to LN column, ending on the last name in the column
    Set LNrange = ThisWorkbook.Worksheets("Tracker").Range("D4", Range("D300").End(xlUp).Address)
    
    'Check whether outlook is open, if it is use get object, if not use create object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If olApp = Empty Then ' tell the user to open outlook if not open
        MsgBox "Please open and sign in to the Outlook Desktop Application BEFORE CLOSING THIS MESSAGE to continue."
        On Error Resume Next
        Set olApp = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If olApp = Empty Then GoTo TheEnd ' if the user STILL didnt open outlook, exit the application
        
    End If
    
    For Each cellone In LNrange.Cells
    
        Rank = ThisWorkbook.Sheets("Tracker").Range(cellone.Address).Offset(0, -1).Value
        email = ThisWorkbook.Sheets("Tracker").Range(cellone.Address).Offset(0, -2).Value
        If UCase(email) = "ARCHIVED" Or UCase(email) = "SKIP" Then GoTo NextName ' if person is archived or marked, skip them
    
        Dim CERTrange As Range, celltwo As Range, tempcol As Integer
        
        ' set range to include every cert that is tracked, ending at the last cert in the list
        tempcol = Range(Cells(3, 1), Cells(3, 100)).End(xlToRight).Column
        Set CERTrange = Range(Cells(cellone.Row, 7).Address, Cells(cellone.Row, tempcol).Address)
        CERTrange.NumberFormat = "0"
        
        For Each celltwo In CERTrange.Cells
        
            Dim URL As String
            ' cert names (table column headers) are on row 3
            cert = Cells(3, celltwo.Column).Text
            
            Select Case cert
            Case "GAT"
                URL = "GAT-url"
            Case "DC"
                URL = "DC-url"
            Case "AT1"
                cert = "Anti-Terrorism Lvl 1 Training"
                URL = "AT1-url"
            Case "EEO"
                URL = "EEO-url"
            Case "ASAP"
                cert = "Army Substance Abuse Program (ASAP)"
                URL = "ASAP-url"
            Case "TARP", "WNSF", "DAR", "ACE", "SHARP"
                GoTo NextIteration
                ' do nothing, these certs are face to face or not currently relevant. keeping in case they become relevant again.
            Case "Cyber Awareness"
                GoTo NextIteration ' do nothing, this cert is tracked by G1. keeping in case tracking becomes relevant again.
                ' URL = "CA-url"
            Case "MPWC"
                ' # No longer relevant (all ranks are required), but keeping in case it becomes relevant again. #
                ' if they are below E5, skip
                'Select Case Rank
                'Case "PVT", "PV2", "PFC", "SPC"
                '    GoTo NextIteration
                'Case Else ' continue as normal
                'End Select
                ' #
                
                cert = "Managing Personnel with Security Clearances"
                URL = "MPWC-url"
            
            End Select
                    
                    
            ' if more than 365 days old
            If celltwo.Value < Date - 365 Or celltwo.Value = "" Then
                expiry = DateDiff("d", celltwo.Value, Date - 365)
                
                'Prepare the mail object
                Set objmail = olApp.CreateItem(olMailItem)
                With objmail
                    .SentOnBehalfOfName = "email@email.mail"
                    .to = email
                    .Subject = cert & " has expired."
                    .body = "Hello, this is an automated message!" & vbNewLine & vbNewLine & _
                        "It has been identified that the following certification is " & expiry & " days past expiration:" & vbNewLine & vbNewLine & _
                        vbTab & cert & vbNewLine & vbNewLine & _
                        "You can re-acquire this certification through the following method:" & vbNewLine & vbNewLine & _
                        vbTab & URL & vbNewLine & vbNewLine & vbNewLine & vbNewLine & _
                        "Please e-mail the latest copy of your certification to the OPS mailbox at:" & vbNewLine & vbNewLine & _
                        vbTab & "email@email.mail" & vbNewLine & vbNewLine & vbNewLine & _
                        "Very Respectfully," & vbNewLine & vbNewLine & _
                        vbTab & "OPS"
                        
                    .display
                End With
                SendKeys "%s", True ' CTRL-S to send the email
                
                
            ' if within 15 days of expiring
            ElseIf celltwo.Value < Date - 365 + 15 And celltwo.Value <> "" Then
                expiry = DateDiff("d", Date - 365, celltwo.Value)
                
                'Prepare the mail object
                Set objmail = olApp.CreateItem(olMailItem)
                With objmail
                    .SentOnBehalfOfName = "email@email.mail"
                    .to = email
                    .Subject = cert & " expiring soon."
                    .body = "Hello, this is an automated message!" & vbNewLine & vbNewLine & vbNewLine & _
                        "It has been identified that the following certification is " & expiry & " days away from expiring." & vbNewLine & vbNewLine & vbNewLine & _
                        vbTab & cert & vbNewLine & vbNewLine & _
                        "You can re-acquire this certification through the following method:" & vbNewLine & vbNewLine & _
                        vbTab & URL & vbNewLine & vbNewLine & vbNewLine & vbNewLine & _
                        "Please e-mail the newest copy of your certification to the OPS mailbox at:" & vbNewLine & _
                        "email@email.mail" & vbNewLine & vbNewLine & vbNewLine & _
                        "Very Respectfully," & vbNewLine & vbNewLine & _
                        vbTab & "OPS"
                        
                    .display
                    
                End With
                SendKeys "%s", True ' CTRL-S to send the email
                
                    
            Else ' do nothing
            End If
            
NextIteration:
        Next celltwo
        CERTrange.NumberFormat = "mm/dd/yyyy;@"
    
        
NextName:
    Next cellone
     
     
MsgBox ("All personnel have been notified.")
     
     
Else ' do nothing
End If


TheEnd:
If olApp = Empty Then
    MsgBox "You have failed to open the Outlook Desktop Application in the allotted time." & vbNewLine & vbNewLine & "Exiting Macro."
End If
    
End Sub
