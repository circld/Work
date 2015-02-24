Attribute VB_Name = "CreatePowerPoint"
' TODO:
' Finish AddCountrySlide

' Slide types
' X 1. Title slide
' X 2. Section slides
'   X Market Analysis
' X 3. Single charts (resize in excel to 750.24 x 372.96)
'   X Global Trends - AUM, Gross and Net
'   X Global Trends - Gross)
'   ? Global Trends - Net Sales by Investment Type
'   X Global Trends - Redemption Rates
'   ? Global Trends - Best Sectors Performance, Risk and Sales
'   ? Global Trends - Products and Players Performance, Risk and Sales (bottom)
'   Market Split - Local & Cross-border Net Sales
' X 4. Double charts - Performance and Sales
' *5. Country-specific Top Bottom


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
    Dim chts()              As Excel.ChartObject
    Dim tbl                 As Range
    Dim tbls()              As Range
    Dim tmpRng              As Range
    Dim title               As String
    Dim i                   As Integer
     
     
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
    AddTitleSlide newPowerPoint
    
    ' Slide 2
    title = "Global Trends - AUM, Gross and Net Sales"
    oWorkbook.Worksheets("Global Net vs Gross Sales").Activate
    Set cht = ActiveSheet.ChartObjects(1)
    AddSingleSlide newPowerPoint, title, cht
    
    ' Slide 3
    title = "Global Trends - Gross Sales"
    oWorkbook.Worksheets("Global Gross Sales %").Activate
    Set cht = ActiveSheet.ChartObjects(1)
    AddSingleSlide newPowerPoint, title, cht
    
    ' Slide 4
    title = "Global Trends - Gross Sales by Investment Type"
    ' NOTE: don't currently have this report/chart
    ' oWorkbook.Worksheets().Activate
    ' Set cht = ActiveSheet.ChartObjects(1)
    
    ' Slide 5
    title = "Global Trends - Redemption Rates over Assets by Investment Type"
    oWorkbook.Worksheets("Redemption Rate Calculation").Activate
    Set cht = ActiveSheet.ChartObjects(1)
    AddSingleSlide newPowerPoint, title, cht
    
    ' Slide 6
    title = "Global Trends - Best Sectors Performance, Risk and Sales"
    ' NOTE: unclear which bubble chart this is; emailed 2/20
    ' oWorkbook.Worksheets().Activate
    ' Set cht = ActiveSheet.ChartObjects(1)
    
    ' Slide 7
    title = "Global Trends - Products and Player Performance, Risk and Sales (Bottom Categories)"
    ' NOTE: don't currently generate this chart
    
    ' Slide 8
    title = "Performance and Sales"
    oWorkbook.Worksheets("3Yr Euro TR Quartile").Activate
    ReDim chts(1 To ActiveSheet.ChartObjects.Count)
    For i = 1 To ActiveSheet.ChartObjects.Count
        Set chts(i) = ActiveSheet.ChartObjects(i)
    Next i
    AddDoubleSlide newPowerPoint, title, chts
    
    ' Slide 9 - Market Analysis
    AddSectionSlide newPowerPoint, "Market Analysis"
    
    ' Slide 10 - Local v Cross-border
    title = "Global Trends - Split–Local & Cross-border Net Sales"
    oWorkbook.Worksheets("Local vs Cross-border net sa").Activate
    Set cht = ActiveSheet.ChartObjects(1)
    AddSingleSlide newPowerPoint, title, cht
    
    ' Slide 11 - Italy
    ' Which args to pass to AddCountrySlide function?
    title = "Italy - Top & Bottom Products"
    ' AddCountrySlide
    
    AppActivate title:="Presentation1 - PowerPoint"
    Set activeSlide = Nothing
    Set newPowerPoint = Nothing
     
End Sub

' Create slide following formatting for a single chart
Function AddSingleSlide( _
    ByRef PP As PowerPoint.Application, _
    ByRef title As String, _
    Optional cht As ChartObject, _
    Optional tbl As Range _
    )

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
    If Not cht Is Nothing Then
        cht.Copy
    ElseIf Not tbl Is Nothing Then
        tbl.Copy
    Else
        Exit Function
    End If
    activeSlide.Shapes.Paste
    
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

' Create slide following formatting for a two charts
Function AddDoubleSlide( _
    ByRef PP As PowerPoint.Application, _
    ByRef title As String, _
    Optional chts As Variant, _
    Optional tbls As Variant _
    )

    Dim activeSlide     As PowerPoint.Slide
    Dim sHeight         As Long
    Dim sWidth          As Long
    Dim ub              As Long
    Dim i               As Integer
    
    
    ' Add and select new slide
    With PP.ActivePresentation
        .Slides.Add .Slides.Count + 1, ppLayoutBlank
        .Slides(.Slides.Count).CustomLayout = .Designs(1).SlideMaster.CustomLayouts(5)
        PP.ActiveWindow.View.GotoSlide .Slides.Count
        Set activeSlide = .Slides(.Slides.Count)
    End With
    
    ' Add chart/table
    ' nb. access chart property/methods via chts.Chart.<Property>
    If Not IsMissing(chts) Then
        For i = 1 To UBound(chts)
            chts(i).Copy
            activeSlide.Shapes.Paste
            Application.Wait (Now + TimeValue("0:00:02"))  ' crashes otherwise
        Next i
    ElseIf Not IsMissing(tbls) Then
        For i = 1 To UBound(tbls)
            tbls(i).Copy
            activeSlide.Shapes.Paste
        Next i
    Else
        Exit Function
    End If
    
    'Adjust the positioning of the Charts/Tables on Powerpoint Slide
    With PP.ActivePresentation.PageSetup
        sHeight = .slideHeight
        sWidth = .slideWidth
    End With
    
    For i = 2 To activeSlide.Shapes.Count
        With activeSlide.Shapes(i)
            .top = 0.15 * sHeight
            .left = 0.05 * sWidth + (i - 2) * sWidth / 2.1
            .Width = 0.9 * sWidth / 2.1
            .Height = 0.75 * sHeight
        End With
    Next i
        
    ' Title
    activeSlide.Shapes(1).TextFrame.TextRange.text = title

End Function

Function AddTitleSlide(ByRef PP As PowerPoint.Application)
    
    Dim dataDate            As Date
    Dim mth, yr, myr, txt   As String
    Dim txtRng              As TextRange
    Dim sHeight, sWidth     As Long
    
    
    ' Set appropriate month + year
    dataDate = DateAdd("m", -2, Now)
    mth = Format(dataDate, "mmm")
    yr = year(dataDate)
    myr = mth & " " & yr
    
    ' Add blank title slide
    With PP.ActivePresentation
        .Slides.Add .Slides.Count + 1, ppLayoutBlank
        .Slides(.Slides.Count).CustomLayout = .Designs(1).SlideMaster.CustomLayouts(1)
        PP.ActiveWindow.View.GotoSlide .Slides.Count
        Set activeSlide = .Slides(.Slides.Count)
    End With
    
    ' Set slide height/width
    With PP.ActivePresentation.PageSetup
        sHeight = .slideHeight
        sWidth = .slideWidth
    End With
    
    ' Title
    With activeSlide.Shapes(1).TextFrame.TextRange
        .text = Join(Array("Cross-Border Monthly Review", Chr(10), myr), "")
        .Font.Color.RGB = RGB(18, 74, 116)  ' Navy
        .Font.Size = 32
        Set txtRng = .Characters(InStr(.text, myr), Len(myr))
        txtRng.Font.Size = 20
    End With
    
    ' Blurb
    With activeSlide.Shapes(2)
    
        With .TextFrame.TextRange
            .text = Join(Array( _
                "Elisabetta Forelli, Senior Product and Data Manager", _
                " (eforelli@sionline.com)", Chr(10), Chr(10), _
                "Source for all charts: ", "Strategic Insight Simfund Global PRO" _
            ), "")
            .Font.Color.RGB = RGB(18, 74, 116)  ' Navy
            .Font.Size = 18
            .Font.Bold = msoTrue
            
            ' italicize email
            txt = "eforelli@sionline.com"
            Set txtRng = .Characters(InStr(.text, txt), Len(txt))
            txtRng.Font.Italic = msoTrue
            
            ' format disclaimer
            txt = "Source for all charts: Strategic Insight Simfund Global PRO"
            Set txtRng = .Characters(InStr(.text, txt), Len(txt))
            txtRng.Font.Bold = msoFalse
            txtRng.Font.Color.RGB = RGB(0, 0, 0)  ' Black
        End With
        
        .top = 0.5 * sHeight
        .left = activeSlide.Shapes(1).left
        .Height = 0.2 * sHeight
        .Width = 0.5 * sWidth
        
    End With
    
    ' Image
    With activeSlide.Shapes(3)
        .top = 0.3 * sHeight
        .left = 0.5 * sWidth
        .Height = 0.55 * sHeight
        .Width = 0.4 * sWidth
    End With
    
End Function

Sub AddSectionSlide(ByRef PP As PowerPoint.Application, ByRef title As String)

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
    
    ' Set slide height/width
    With PP.ActivePresentation.PageSetup
        sHeight = .slideHeight
        sWidth = .slideWidth
    End With
    
    ' Title
    With activeSlide.Shapes(1)
        .TextFrame.TextRange.text = title
        .TextFrame.TextRange.Font.Size = 28
        .TextFrame.WordWrap = msoFalse
        .TextFrame.AutoSize = ppAutoSizeShapeToFitText
        .left = (sWidth - .Width) / 2
        .top = (sHeight - .Height) / 2
    End With
    
End Sub

' Create slide with Country-specific formatting
Function AddCountrySlide( _
    ByRef PP As PowerPoint.Application, _
    ByRef title As String, _
    Optional chts As Variant, _
    Optional tbls As Variant _
    )

    Dim activeSlide     As PowerPoint.Slide
    Dim sHeight         As Long
    Dim sWidth          As Long
    Dim ub              As Long
    Dim i               As Integer
    
    ' TODO: edit code below to generate country-specific slide
    ' Add and select new slide
    With PP.ActivePresentation
        .Slides.Add .Slides.Count + 1, ppLayoutBlank
        .Slides(.Slides.Count).CustomLayout = .Designs(1).SlideMaster.CustomLayouts(5)
        PP.ActiveWindow.View.GotoSlide .Slides.Count
        Set activeSlide = .Slides(.Slides.Count)
    End With
    
    ' Add chart/table
    ' nb. access chart property/methods via chts.Chart.<Property>
    If Not IsMissing(chts) Then
        For i = 1 To UBound(chts)
            chts(i).Copy
            activeSlide.Shapes.Paste
            Application.Wait (Now + TimeValue("0:00:02"))
        Next i
    ElseIf Not IsMissing(tbls) Then
        For i = 1 To UBound(tbls)
            tbls(i).Copy
            activeSlide.Shapes.Paste
        Next i
    Else
        Exit Function
    End If
    
    'Adjust the positioning of the Charts/Tables on Powerpoint Slide
    With PP.ActivePresentation.PageSetup
        sHeight = .slideHeight
        sWidth = .slideWidth
    End With
    
    For i = 2 To activeSlide.Shapes.Count
        With activeSlide.Shapes(i)
            .top = 0.15 * sHeight
            .left = 0.05 * sWidth + (i - 2) * sWidth / 2.1
            .Width = 0.9 * sWidth / 2.1
            .Height = 0.75 * sHeight
        End With
    Next i
        
    ' Title
    activeSlide.Shapes(1).TextFrame.TextRange.text = title

End Function
