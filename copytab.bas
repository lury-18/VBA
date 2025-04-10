Attribute VB_Name = "copytab"


Sub CopySheetWithoutLinks()

    Set twb = ThisWorkbook

    Set wb = Workbooks("book1.xlsb")
    ThisWorkbook.Sheets("Comments-new").Copy After:=wb.Sheets("Sheet1")
    
    

    
    For Each Cell In wb.Sheets("Comments-new").UsedRange
        If Cell.HasFormula Then
            If InStr(Cell.Formula, "[" & twb.Name & "]") > 0 Then
                Cell.Formula = Replace(Cell.Formula, "[" & twb.Name & "]", "")
            End If
        End If
    Next Cell

    
End Sub
     
