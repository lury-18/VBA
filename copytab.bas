Attribute VB_Name = "copytab"


Sub CopySheetWithoutLinks()

    Set twb = ThisWorkbook

    Set wb = Workbooks("AHQ PBI LCD TV DASHBOARD (Latest)_PPT automation7_SG_2024_WK11.xlsb")
    ThisWorkbook.Sheets("Comments-new").Copy After:=wb.Sheets("SELL in_thru Test")
    
    
    'With wb.Sheets("SELL in_thru Test").UsedRange
    '    .Replace What:="[" & twb.Name & "]", Replacement:="", LookAt:=xlPart
    'End With
    
    For Each Cell In wb.Sheets("Comments-new").UsedRange
        If Cell.HasFormula Then
            If InStr(Cell.Formula, "[" & twb.Name & "]") > 0 Then
                Cell.Formula = Replace(Cell.Formula, "[" & twb.Name & "]", "")
            End If
        End If
    Next Cell

    
End Sub
     