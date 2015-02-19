Attribute VB_Name = "CreatePowerPoint"
' TODO:
' X Logic to load design programatically

' Slide types
' >>1. Title slide
' 2. Section slides
'   Sector Focus:
'   Market Analysis
' 3. Single charts (resize in excel to 750.24 x 372.96)
'   Global Trends - AUM, Gross and Net
'   Global Trends - Gross)
'   Global Trends - Net Sales by Investment Type
'   Global Trends - Redemption Rates
'   Global Trends - Best Sectors Performance, Risk and Sales
'   Global Trends - Products and Players Performance, Risk and Sales (bottom)
'   Market Split - Local & Cross-border Net Sales
' 4. Double charts - Performance and Sales
' 5. Country-specific Top Bottom


Sub CreatePowerPoint()

 ' Add a reference to the Microsoft PowerPoint Library by:
    ' 1. Go to Tools in the VBA menu
    ' 2. Click on Reference
    ' 3. Scroll down to Microsoft PowerPoint X.0 Object Library, check the box, and press Okay
 
    ' First we declare the variables we will be using
    Dim oWorkbook           As Workbook
    Dim newPowerPoint       As PowerPoint.Application
    Dim themeLoc            As String
    Dim activeSlide         As PowerPoint.Slide
    Dim cht                 As Excel.ChartObject
    Dim pic                 As Picture
    Dim tmpRng              As Range
    Dim counter             As Integer
    Dim title               As String
     
     
    Set oWorkbook = ActiveWorkbook
    
     ' Look for existing instance
    On Error Resume Next
    Set newPowerPoint = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
     
    ' Create new PowerPoint
    If newPowerPoint Is Nothing Then
        Set newPowerPoint = New PowerPoint.Application
    End If
    
    ' Make a presentation in PowerPoint
    If newPowerPoint.Presentations.Count = 0 Then
        newPowerPoint.Presentations.Add
    End If
    
    ' Load design & set slide size
    themeLoc = "C:\Users\pgaraud\AppData\Roaming\Microsoft\Templates\Document Themes\SI.thmx"
    newPowerPoint.ActivePresentation.ApplyTheme themeLoc
    newPowerPoint.ActivePresentation.PageSetup.SlideSize = ppSlideSizeA4Paper
    
    ' Create slides
    
    ' Title slide
    
    title = "Global Trends - AUM, Gross and Net Sales"
    oWorkbook.Worksheets("Global Net vs Gross Sales").Activate
    ActiveSheet.ChartObjects(1).CopyPicture
    ActiveSheet.Pictures.Paste.Select
    Set pic = Selection
    NewSingleSlide newPowerPoint, pic, title
    pic.Delete
    
    title = "Global Trends - Gross Sales"
    
    
    title = "Global Trends - Net Sales by Investment Type"
    
    
    title = "Global Trends - Redemption Rates over Assets by Investment Type"
    
    
    title = "Global Trends - Best Sectors Performance, Risk and Sales"
    
    
    title = "Global Trends - Products and Player Performance, Risk and Sales (Bottom Categories)"
    
    
    
     
    AppActivate title:="Presentation1 - PowerPoint"
    Set activeSlide = Nothing
    Set newPowerPoint = Nothing
     
End Sub

' Create slide following formatting for a single chart
Function NewSingleSlide(ByRef PP As PowerPoint.Application, ByRef cht As Picture, ByRef title As String)

    Dim activeSlide     As PowerPoint.Slide
    Dim sHeight         As Long
    Dim sWidth          As Long
    
    ' Add and select new slide
    With PP.ActivePresentation
        .Slides.Add .Slides.Count + 1, ppLayoutBlank
        .Slides(.Slides.Count).CustomLayout = .Designs(1).SlideMaster.CustomLayouts(5)
        PP.ActiveWindow.View.GotoSlide .Slides.Count
        Set activeSlide = .Slides(.Slides.Count)
    End With
    
    ' Add chart/table
    cht.CopyPicture xlPrinter, xlPicture
    activeSlide.Shapes.PasteSpecial ppPasteMetafilePicture
    
    'Adjust the positioning of the Chart on Powerpoint Slide
    With PP.ActivePresentation.PageSetup
        sHeight = .slideHeight
        sWidth = .slideWidth
    End With
    
    With activeSlide.Shapes(2)
        .top = 0.15 * sHeight
        .left = 0.05 * sWidth
        .Width = 0.9 * sWidth
        .Height = 0.75 * sHeight
    End With
        
    ' Title
    activeSlide.Shapes(1).TextFrame.TextRange.text = title

End Function

Function CopyChartToPP(ByRef newPowerPoint As PowerPoint.Application, ByRef cht As ChartObject, Optional left As Long, _
                        Optional top As Long, Optional title As String)

            'Add a new slide where we will paste the chart
                Set activeSlide = newPowerPoint.ActivePresentation.Slides(newPowerPoint.ActivePresentation.Slides.Count)
                    
            'Copy the chart and paste it into the PowerPoint as a Metafile Picture
                cht.Select
                cht.CopyPicture xlPrinter, xlPicture
                activeSlide.Shapes.PasteSpecial DataType:=ppPasteMetafilePicture
                activeSlide.Shapes(activeSlide.Shapes.Count).Select
                
            'Adjust the positioning of the Chart on Powerpoint Slide
                If left = 0 Then left = 15
                If top = 0 Then top = 125
                newPowerPoint.ActiveWindow.Selection.ShapeRange.left = left
                newPowerPoint.ActiveWindow.Selection.ShapeRange.top = top
                
                activeSlide.Shapes(2).Width = 200
                activeSlide.Shapes(2).left = 505
                
            ' Title
                If title = "" And cht.Chart.HasTitle = True Then title = cht.Chart.ChartTitle.text
                activeSlide.Shapes(1).TextFrame.TextRange.text = title '& vbNewLine
                ' If want to insert commentary:
                ' activeSlide.Shapes(2).TextFrame.TextRange.InsertAfter (Range("J8").Value & vbNewLine)
                
            'Now let's change the font size of the callouts box
                activeSlide.Shapes(2).TextFrame.TextRange.Font.Size = 16

End Function

Function CopyPicToPP(ByRef newPowerPoint As PowerPoint.Application, ByRef pic As Picture, Optional left As Long, _
                        Optional top As Long, Optional title As String)

            'Add a new slide where we will paste the chart
                Set activeSlide = newPowerPoint.ActivePresentation.Slides(newPowerPoint.ActivePresentation.Slides.Count)
                    
            'Copy the chart and paste it into the PowerPoint as a Metafile Picture
                pic.Select
                pic.CopyPicture xlPrinter, xlPicture
                activeSlide.Shapes.PasteSpecial DataType:=ppPasteMetafilePicture
                activeSlide.Shapes(activeSlide.Shapes.Count).Select
                
            'Adjust the positioning of the Chart on Powerpoint Slide
                If left = 0 Then left = 15
                If top = 0 Then top = 125
                newPowerPoint.ActiveWindow.Selection.ShapeRange.left = left
                newPowerPoint.ActiveWindow.Selection.ShapeRange.top = top
                
                activeSlide.Shapes(2).Width = 200
                activeSlide.Shapes(2).left = 505
                
            ' Title
                activeSlide.Shapes(1).TextFrame.TextRange.text = title & vbNewLine
                ' If want to insert commentary:
                ' activeSlide.Shapes(2).TextFrame.TextRange.InsertAfter (Range("J8").Value & vbNewLine)
                
            'Now let's change the font size of the callouts box
                activeSlide.Shapes(2).TextFrame.TextRange.Font.Size = 16

End Function
