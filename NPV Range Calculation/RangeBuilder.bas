'Simple script to copy data from a table to different cells, and copy back the computed value back to the table 
'See attached Excel File for practical usage to find a range of results in a NPV calculation 

Sub Button1_Click()

    'Define the workbook and worksheet
    Dim mainworkBook As Workbook
    Dim objNewWorksheet As Worksheet
    
    'Set the active worksheet
    Set mainworkBook = ActiveWorkbook
    Set objNewWorksheet = mainworkBook.Sheets("Financial projection")
    
    'Hardcoded.. surely there's a more elegant way to do this?
    For i = 13 To 21
    
    'Copy Incremental growth in market share for MM light from 2006 base of 0.25%
    objNewWorksheet.Range("K" & i).Copy
    objNewWorksheet.Range("B7").PasteSpecial xlPasteValues
    
    'MM Lager % sales loss from introduction of MM light
    objNewWorksheet.Range("L" & i).Copy
    objNewWorksheet.Range("B23").PasteSpecial xlPasteValues
    
    'Copy NPV 5 Years results back to table
    objNewWorksheet.Range("B32").Copy
    objNewWorksheet.Range("M" & i).PasteSpecial xlPasteValues
        
    Next i
    
End Sub
