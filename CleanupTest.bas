Attribute VB_Name = "CleanupTest"

Sub cleantest()

    'Call Global_Variables

        
    Dim pptApp As Object    'PowerPoint.Application
    Dim pptpres As Object   'PowerPoint.Presentation
    Dim pptSlide As Object  'PowerPoint.Slide
    Dim i As Integer        'for looping
    Dim img As Object
    Dim country As String


    'create ppt object and open file
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then Set pptApp = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    
    If pptApp Is Nothing Then Set pptApp = CreateObject("PowerPoint.Application")

    ' Add the first slide
    Set pptpres = pptApp.Presentations.Add


    '======================================================================
    '=============Put the Original Code====================================
    '======================================================================
    Set pptSlide = pptpres.Slides.Add(pptpres.Slides.Count + 1, 12)
    lastRow2A = ThisWorkbook.Worksheets("Branch | Scorecard (to65)").Range("C:C").Find("65-69"" Total", LookIn:=xlValues, SearchDirection:=xlPrevious).Row
    
    Set rangeList2A = ThisWorkbook.Worksheets("Branch | Scorecard (to65)").Range("C4:IW" & lastRow2A).Offset(0, 0)
    
    rangeList2A.CopyPicture xlScreen, xlBitmap
    Application.Wait Now + TimeValue("00:00:01")
    
    
    Set img = pptSlide.Shapes.Paste
    With img
        .LockAspectRatio = msoFalse
        .Height = 420
        .Width = 960
        .Left = 0
        .Top = 10
    End With
          
    
    ' Create a new TEXTBOX for SINGLE CELL "COMMENTS!DV10:DV15" which has space row in excel below the 1ST image (distance 10)
    Set shape = pptSlide.Shapes.AddTextbox(1, 0, img.Top + img.Height + 10, 960, 210) ' 1 is msoTextOrientationHorizontal
    
    shape.TextFrame.TextRange.Text = FindLargestSOGap()

    With shape
        .Fill.ForeColor.RGB = RGB(254, 240, 240)
        .Line.ForeColor.RGB = RGB(0, 0, 255)
        .Line.Weight = 1 ' Adjust the thickness as needed
        shape.TextFrame.TextRange.Font.Color.RGB = RGB(0, 112, 192)
        shape.TextFrame.TextRange.Font.Name = "SST"
        shape.TextFrame.TextRange.Font.Size = 12
    End With
    '======================================================================
    '=============Put the Modified Code====================================
    '======================================================================
   
    '======================================================================
    '=============End of  Modified Code====================================
    '======================================================================
    Set pptSlide = Nothing
    Set pptpres = Nothing
    Set pptApp = Nothing
End Sub
