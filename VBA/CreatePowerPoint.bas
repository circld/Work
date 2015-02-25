Attribute VB_Name = "CreatePowerPoint"

Sub CreatePowerPoint()

 ' Add a reference to the Microsoft PowerPoint Library by:
    ' 1. Go to Tools in the VBA menu
    ' 2. Click on Reference
    ' 3. Scroll down to Microsoft PowerPoint X.0 Object Library, check the box, and press Okay
 
    ' Var declarations
    Dim FileName            As String
    Dim FPath                As String
    Dim oWorkbook           As Workbook
    Dim newPowerPoint       As PowerPoint.Application
    Dim themeLoc            As String
    Dim activeSlide         As PowerPoint.Slide
    Dim cht                 As Excel.ChartObject
    Dim chts()              As Excel.ChartObject
    Dim tbl                 As Range
    Dim tbls()              As Range
    Dim title               As String
    Dim i                   As Integer
    Dim CtryMap             As New Scripting.Dictionary
    Dim TablePage           As String
    Dim Ctry                As Variant
    Dim CtryName            As String
     
    
    ' Find Charts
    FPath = Workbooks("CBSOReportMacros.xlsb").Path & "\"
    FileName = Dir(FPath & "*.xlsm")
    Workbooks.Open FPath & FileName
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
    themeLoc = oWorkbook.Path & "\SI.thmx"
    newPowerPoint.ActivePresentation.ApplyTheme themeLoc
    newPowerPoint.ActivePresentation.PageSetup.SlideSize = ppSlideSizeA4Paper
    newPowerPoint.ActivePresentation.SaveAs _
        oWorkbook.Path & "\" & Format(Now, "yymmdd") & " CBSO Presentation.pptx"
    
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
    newPowerPoint.ActivePresentation.Save
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
    
    ' Country-specific slides
    ' Market Global TopBottom 5 Se
    ' Origination Markets Table
    Set CtryMap = MapChartsToCountries(oWorkbook.Name, "Market Global TopBottom 5 Se")
    TablePage = "Origination Markets Table"
    
    For Each Ctry In CtryMap.Keys
        CtryName = Ctry
        title = CtryName & " - Top & Bottom Products"
        ReDim chts(1 To 2)
        Set chts(1) = CtryMap(CtryName)(0)
        Set chts(2) = CtryMap(CtryName)(1)
        newPowerPoint.ActivePresentation.Save
        AddCountrySlide newPowerPoint, title, chts, _
            CountryTable(oWorkbook.Name, TablePage, CtryName)
    Next Ctry
    
    newPowerPoint.ActivePresentation.Save
    oWorkbook.Close False
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
            chts(i).CopyPicture
            activeSlide.Shapes.Paste
        Next i
    ElseIf Not IsMissing(tbls) Then
        For i = 1 To UBound(tbls)
            tbls(i).Copy
            activeSlide.Shapes.Paste.Select
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
            .SoftEdge.Type = msoSoftEdgeType3
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
    ByRef chts As Variant, _
    ByRef tbl As Range _
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
    
    ' Adjust chart titles
    chts(1).Chart.ChartTitle.text = "Top"
    chts(2).Chart.ChartTitle.text = "Bottom"
    
    ' Add charts & table
    ' Had to use CopyPicture to avoid unpredictable run-time errors (1004)
    ' The SoftEdge type is to hide the picture border
    chts(1).CopyPicture
    activeSlide.Shapes.Paste
    activeSlide.Shapes(activeSlide.Shapes.Count).SoftEdge.Type = msoSoftEdgeType3
    chts(2).CopyPicture
    activeSlide.Shapes.Paste
    activeSlide.Shapes(activeSlide.Shapes.Count).SoftEdge.Type = msoSoftEdgeType3
    tbl.Copy
    activeSlide.Shapes.Paste
    
    'Adjust the positioning of the Charts/Tables on Powerpoint Slide
    With PP.ActivePresentation.PageSetup
        sHeight = .slideHeight
        sWidth = .slideWidth
    End With
    
    For i = 2 To activeSlide.Shapes.Count
        With activeSlide.Shapes(i)
            If i < activeSlide.Shapes.Count Then
                .top = 0.15 * sHeight + (i - 2) * sHeight / 2.5
                .left = 0.05 * sWidth
                .Width = 0.9 * sWidth / 2.1
                .Height = 0.75 * sHeight / 2.1
            Else
                .top = 0.15 * sHeight
                .left = 0.05 * sWidth + sWidth / 2.1
                .Width = 0.9 * sWidth / 2.1
                .Height = 0.75 * sHeight
            End If
        End With
    Next i
        
    ' Title
    activeSlide.Shapes(1).TextFrame.TextRange.text = title

End Function

' Map TopBottom charts to name of country
Function MapChartsToCountries( _
    ByRef WbName As String, _
    ByRef TopBottomPage As String _
    ) As Scripting.Dictionary

    Dim oWorkbook   As Workbook
    Dim ChartMap    As New Scripting.Dictionary
    Dim i           As Integer
    
    Set oWorkbook = Workbooks(WbName)
    
    ' Ungroup charts
    UngroupCharts WbName, TopBottomPage
    
    ' Map TopBottom charts
    For i = 1 To oWorkbook.Worksheets(TopBottomPage).ChartObjects.Count / 2
        
        ' Map country name to chart array
        With oWorkbook.Worksheets(TopBottomPage)
            ChartMap.Add .ChartObjects(2 * i).Chart.ChartTitle.text, _
                Array(.ChartObjects(2 * i - 1), .ChartObjects(2 * i))
        End With
        
    Next i
    
    Set MapChartsToCountries = ChartMap  ' to return dictionary

End Function

' Highlights appropriate row in Origination Markets Table and returns range
Function CountryTable( _
    ByRef WbName As String, _
    ByRef TablePage As String, _
    ByRef CtryName As String _
    ) As Range

    Dim oWorkbook   As Workbook
    Dim oWorksheet  As Worksheet
    Dim tblRng      As Range
    Dim tblRow      As Variant

    Set oWorkbook = Workbooks(WbName)
    Set oWorksheet = oWorkbook.Worksheets(TablePage)
    Set tblRng = oWorksheet.Cells(1, 1).CurrentRegion
    
    ' Reset table formatting to unhighlighted state
    tblRng.Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    tblRng.BorderAround Weight:=xlMedium, Color:=RGB(0, 0, 0)
    tblRng.Rows(1).BorderAround Weight:=xlMedium, Color:=RGB(0, 0, 0)
    
    ' Highlight specific country row
    For Each tblRow In tblRng.Rows
        If tblRow.Cells(1, 1).Value = CtryName Then
            tblRow.Borders(xlEdgeTop).LineStyle = xlLineStyleNone
            tblRow.BorderAround 1, 3, Color:=RGB(255, 0, 0)
        End If
    Next tblRow
    
    Set CountryTable = tblRng

End Function

Sub UngroupCharts(ByRef WbName As String, ByRef TopBottomPage As String)
    
    Dim sh As Variant
    
    For Each sh In Workbooks(WbName).Worksheets(TopBottomPage).Shapes
        If sh.Type = msoGroup Then
            sh.Ungroup
        End If
    Next sh
    
End Sub
