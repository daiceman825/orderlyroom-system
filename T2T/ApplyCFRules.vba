Sub ApplyCFRules(rng As Range)

With rng

    ' remove all conditional formatting
    .FormatConditions.Delete
    
    Application.CutCopyMode = False

    ' Copy from "Start Page"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This range can be changed as needed. However,
    ' ALL conditional formatting rules that need to be
    ' applied to the Troops to Task   MUST   be in the
    ' range that is referenced below
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ThisWorkbook.Sheets("HOLIDAYS").Range("J3:L150").Copy
    
    ' Paste Formatting rules to Rng
    ' Do not skip empty cells
    .PasteSpecial xlPasteFormats, , False
    
    Application.CutCopyMode = True
    
End With

End Sub

