Attribute VB_Name = "Module1"
' Module with all chart-producing procedures
Option Explicit


Sub GlobalNetGrossChart()

    Dim MyChart                 As Chart
    Dim MyRange                 As Range
    Dim StartCell, ChartLoc     As Range
    Dim ChartHeight, ChartWidth As Long
        
    Set MyRange = Range("A1:N5")
    
    Set MyChart = ActiveSheet.Shapes.AddChart(xlColumnClustered).Chart
    
    ' Chart location & size params
    ChartHeight = 30
    ChartWidth = 12
    Set StartCell = Cells(1, 1).Offset(6, 0)
    Set ChartLoc = Range(StartCell, StartCell.Offset(ChartHeight - 1, ChartWidth - 1))
    
    With MyChart
        
        ' Prepare data & presentation
        .SetSourceData Source:=MyRange
        .SeriesCollection(3).ChartType = xlLine
        .SeriesCollection(3).AxisGroup = xlSecondary
        .SeriesCollection(4).ChartType = xlLine
        .ChartGroups(1).GapWidth = 80
        
        ' Set chart location
        With .Parent
            .Top = ChartLoc.Top
            .Left = ChartLoc.Left
            .Height = ChartLoc.Height
            .Width = ChartLoc.Width
        End With
        
        ' Adjust axes (nb: values already in millions)
        With MyChart.Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.text = "€ Billions"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Italic = msoTrue
            .AxisTitle.Font.Bold = msoFalse
            .AxisTitle.Orientation = xlHorizontal
            .AxisTitle.Top = -5
            .AxisTitle.Left = 8
            .MajorGridlines.Delete
            .DisplayUnit = xlThousands
            .TickLabels.NumberFormat = "#,##0"
            .HasDisplayUnitLabel = False
            .TickLabels.Font.Size = 12
        End With
        
        With MyChart.Axes(xlValue, xlSecondary)
            .HasTitle = True
            .AxisTitle.text = "€ Trillions"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Italic = msoTrue
            .AxisTitle.Font.Bold = msoFalse
            .AxisTitle.Orientation = xlHorizontal
            .AxisTitle.Top = -5
            .AxisTitle.Left = MyChart.ChartArea.Width - 80
            .DisplayUnit = xlMillions
            .HasDisplayUnitLabel = False
            .TickLabels.Font.Size = 12
        End With
        
        MyChart.Axes(xlCategory).TickLabels.Font.Size = 12
        MyChart.Axes(xlCategory).TickLabels.Orientation = 15
        
        ' Colors
        ' Light Blue
        With MyChart.SeriesCollection(1).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.400000006
            .Transparency = 0
            .Solid
        End With
        
        ' Navy
        With MyChart.SeriesCollection(2).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText2
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
            .Transparency = 0
            .Solid
        End With

        .SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(152, 185, 84)
        .SeriesCollection(3).Format.Line.Weight = 2.25
        
        With MyChart.SeriesCollection(4).Format.Line
            .Visible = msoTrue
            .Weight = 2.25
            .DashStyle = msoLineSysDash
            .ForeColor.ObjectThemeColor = msoThemeColorText2
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
            .Transparency = 0
        End With

        ' Legend
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 12
        .Legend.Top = .Legend.Top - 15
        .Legend.Left = .PlotArea.Left + 10
        .Legend.Height = 30
        .Legend.Width = .PlotArea.Width - 20
        
        ' Correcting for squished shape (from incl axis titles)
        .PlotArea.Top = 12
        .PlotArea.Left = 0
        .PlotArea.Width = .ChartArea.Width - 10
        .PlotArea.Height = .ChartArea.Height - 20
        
    End With

    ActiveSheet.Shapes(1).Line.Visible = msoFalse
    Range("A1").Select
    
End Sub


Sub GlobalGrossPctChart()

    Dim MyChart As Chart
    Dim MyRange As Range
    Dim StartCell, ChartLoc     As Range
    Dim ChartHeight, ChartWidth As Long
    
    Set MyRange = Range("A1:N5")
    
    ' Chart location & size params
    ChartHeight = 30
    ChartWidth = 14
    Set StartCell = Cells(1, 1).Offset(6, 0)
    Set ChartLoc = Range(StartCell, StartCell.Offset(ChartHeight - 1, ChartWidth - 1))
    
    Set MyChart = ActiveSheet.Shapes.AddChart(xlColumnStacked100).Chart
    
    With MyChart
        
        ' Prepare data & presentation
        .SetSourceData Source:=MyRange
        .ChartGroups(1).GapWidth = 80
        
        ' Adjust axes (nb: values already in millions)
        With MyChart.Axes(xlValue, xlPrimary)
            .MajorGridlines.Delete
            .TickLabels.Font.Size = 12
        End With
        
        MyChart.Axes(xlCategory).TickLabels.Font.Size = 12
        MyChart.Axes(xlCategory).TickLabels.Orientation = 15
        
        ' Set Location
        With .Parent
            .Top = ChartLoc.Top
            .Left = ChartLoc.Left
            .Height = ChartLoc.Height
            .Width = ChartLoc.Width
        End With
        
        ' Colors
        ' Navy
        With MyChart.SeriesCollection(1).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText2
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
            .Transparency = 0
            .Solid
        End With
        
        ' Light Blue
        With MyChart.SeriesCollection(2).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.400000006
            .Transparency = 0
            .Solid
        End With
        
        ' Stripey
        With MyChart.SeriesCollection(3).Format.Fill
            .Visible = msoTrue
            .Patterned msoPatternDarkUpwardDiagonal
            .ForeColor.RGB = RGB(255, 255, 255)
            .BackColor.ObjectThemeColor = msoThemeColorBackground1
            .BackColor.TintAndShade = 0
            .BackColor.Brightness = -0.5
        End With
        
        ' Dot-ey
        With MyChart.SeriesCollection(4).Format.Fill
            .Visible = msoTrue
            .Patterned msoPattern50Percent
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.150000006
            .BackColor.ObjectThemeColor = msoThemeColorBackground1
            .BackColor.TintAndShade = 0
            .BackColor.Brightness = -0.5
        End With

        ' Legend
        .Legend.Position = xlLegendPositionTop
        .Legend.Font.Size = 12
        
    End With

    ActiveSheet.Shapes(1).Line.Visible = msoFalse
    Range("A1").Select
    
End Sub

Sub NetAvgPerfChart()

    Dim MyChart                 As Chart
    Dim Series1                 As Range
    Dim Series2                 As Range
    Dim i                       As Integer
    Dim StartCell, ChartLoc     As Range
    Dim ChartHeight, ChartWidth As Long
    
    Set Series1 = Range("A1:N3")
    Series1.Select
    
    Set Series2 = Range("O2:AA3")
    
    ' Chart location & size params
    ChartHeight = 24
    ChartWidth = 12
    Set StartCell = Cells(1, 1).Offset(6, 0)
    Set ChartLoc = Range(StartCell, StartCell.Offset(ChartHeight - 1, ChartWidth - 1))
    
    Set MyChart = ActiveSheet.Shapes.AddChart(xlColumnClustered).Chart
    
    With MyChart
        
        ' Prepare data & presentation
        .SetSourceData Source:=Series1
        ' Add ATR data (adding entire range not what we want)
        For i = 2 To 3
            With MyChart.SeriesCollection.NewSeries
                .Name = "ATR " & Range("A" & i)
                .Values = Range("O" & i & ":AA" & i)
                .XValues = Range("A2:N1")
            End With
        Next i
        
        .SeriesCollection(3).ChartType = xlLine
        .SeriesCollection(3).AxisGroup = xlSecondary
        .SeriesCollection(4).ChartType = xlLine
        .ChartGroups(1).GapWidth = 80
        .SeriesCollection(4).AxisGroup = xlSecondary
        
        ' Set Location
        With .Parent
            .Top = ChartLoc.Top
            .Left = ChartLoc.Left
            .Height = ChartLoc.Height
            .Width = ChartLoc.Width
        End With
        
        ' Adjust axes (nb: values already in millions)
        With MyChart.Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.text = "€ Billions"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Italic = msoTrue
            .AxisTitle.Font.Bold = msoFalse
            .AxisTitle.Orientation = xlHorizontal
            .AxisTitle.Top = -5
            .AxisTitle.Left = 8
            .MajorGridlines.Delete
            .DisplayUnit = xlThousands
            .TickLabels.NumberFormat = "#,##0"
            .HasDisplayUnitLabel = False
            .TickLabels.Font.Size = 12
        End With
        
        With MyChart.Axes(xlValue, xlSecondary)
            .HasTitle = True
            .AxisTitle.text = "%"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Italic = msoTrue
            .AxisTitle.Font.Bold = msoFalse
            .AxisTitle.Orientation = xlHorizontal
            .AxisTitle.Top = -5
            .AxisTitle.Left = MyChart.ChartArea.Width - 30
            .TickLabels.Font.Size = 12
            .TickLabels.NumberFormat = "#,##0.0"
        End With
        
        MyChart.Axes(xlCategory).TickLabelPosition = xlNone
        
        ' Colors
        ' Navy
        With MyChart.SeriesCollection(1).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText2
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
            .Transparency = 0
            .Solid
        End With
        
        ' Light Blue
        With MyChart.SeriesCollection(2).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.400000006
            .Transparency = 0
            .Solid
        End With

        .SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(152, 185, 84)
        .SeriesCollection(3).Format.Line.Weight = 2.25
        
        With MyChart.SeriesCollection(4).Format.Line
            .Visible = msoTrue
            .Weight = 2.25
            .DashStyle = msoLineSysDash
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.6000000238
            .Transparency = 0
        End With
        
        ' Legend
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 12
        .Legend.Top = .Legend.Top - 20
        .Legend.Left = .Legend.Left - 20
        .Legend.Height = 30
        .Legend.Width = 400
        
        ' Correcting for squished shape (from incl axis titles)
        .PlotArea.Top = 12
        .PlotArea.Left = 0
        .PlotArea.Width = 350
        .PlotArea.Height = 180
        
        ' Resize PlotArea
        With .PlotArea
            .Top = .Top + 5
            .Height = MyChart.ChartArea.Height - 25
            .Width = MyChart.ChartArea.Width - 10
        End With
        
    End With

    ActiveSheet.Shapes(1).Line.Visible = msoFalse
    Range("A1").Select
    
End Sub

Sub RedempCalcChart()

    Dim MySheet                 As Worksheet
    Dim MyChart                 As Chart
    Dim Series1                 As Range
    Dim i                       As Integer
    Dim StartCell, ChartLoc     As Range
    Dim ChartHeight, ChartWidth As Long
    
    Set Series1 = Range("A20:N22")
    
    Set MySheet = ActiveSheet
    Set MyChart = MySheet.Shapes.AddChart(xlColumnClustered).Chart
    
    ' Chart location & size params
    ChartHeight = 24
    ChartWidth = 12
    Set StartCell = Series1(1, 1).Offset(6, 0)
    Set ChartLoc = Range(StartCell, StartCell.Offset(ChartHeight - 1, ChartWidth - 1))
    
    With MyChart
        
        ' Prepare data & presentation
        .SetSourceData Source:=Series1
        .ChartGroups(1).GapWidth = 80
        .FullSeriesCollection(1).XValues = _
            "='" & MySheet.Name & "'!B1:N1"

        ' Set Location
        With .Parent
            .Top = ChartLoc.Top
            .Left = ChartLoc.Left
            .Height = ChartLoc.Height
            .Width = ChartLoc.Width
        End With

        ' Adjust axes (nb: values already in millions)
        With MyChart.Axes(xlValue, xlPrimary)
            .MajorGridlines.Delete
            .HasDisplayUnitLabel = False
            .TickLabels.Font.Size = 12
        End With
        
        MyChart.Axes(xlCategory).TickLabels.Font.Size = 11
        MyChart.Axes(xlCategory).TickLabels.Orientation = 25
        
        ' Colors
        ' Navy
        With MyChart.SeriesCollection(1).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText2
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
            .Transparency = 0
            .Solid
        End With
        
        ' Light Blue
        With MyChart.SeriesCollection(2).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.400000006
            .Transparency = 0
            .Solid
        End With
        
        ' Stripey
        With MyChart.SeriesCollection(3).Format.Fill
            .Visible = msoTrue
            .Patterned msoPatternLightUpwardDiagonal
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .BackColor.RGB = RGB(255, 255, 255)
        End With
        
        ' Legend
        .Legend.Position = xlLegendPositionTop
        .Legend.Left = MyChart.ChartArea.Left + 20
        .Legend.Font.Size = 12
        .Legend.Height = 30
        .Legend.Width = MyChart.ChartArea.Width - 50
        
    End With

    ActiveSheet.Shapes(1).Line.Visible = msoFalse
    Range("A1").Select

End Sub

Sub MTopBottomChart()
    
    Dim Countries()             As CMetaData
    Dim Area(2, 2), ChartLoc    As Range
    Dim MySheet                 As Worksheet
    Dim MyChart()               As Chart
    Dim i, j, Count, BlockCount As Long
    Dim LastRow                 As Long
    Dim ChartWidth, ChartHeight As Long
    Dim AuxTitle                As Shape
    
    Set MySheet = ActiveSheet
    Set Area(2, 1) = Cells(Rows.Count, 1).End(xlUp)
    Set Area(1, 1) = Cells(2, 1)  ' header row; blank row before data
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 2) = Cells(Area(2, 1).Row, Area(1, 2).Column)
    
    ' Set chart size
    ChartHeight = 12    ' in rows
    ChartWidth = 4      ' in cols
    
    BlockCount = Area(2, 1).Row
    Count = 0
    
    ' Find first row of each data block
    While BlockCount > 2
        Count = Count + 1
        BlockCount = Cells(BlockCount, 1).End(xlUp).End(xlUp).Row
    Wend
    
    ' To handle arbitrary number of countries
    ReDim Countries(1 To Count)
    ReDim MyChart(Count, 2)  ' 2 charts per country: top (1), bottom (2)
    
    ' Define Class CMetaData for each country
    BlockCount = Area(2, 1).Row  ' last row of last block
    For i = 1 To Count
        j = Count + 1 - i  ' counting down since starting @ bottom
        Set Countries(j) = New CMetaData
        Countries(j).LastRow = BlockCount
        Countries(j).LastCol = Area(1, 2).Column
        ' Move row counter to first row of block
        BlockCount = Cells(BlockCount, 1).End(xlUp).Row
        Countries(j).FirstRow = BlockCount
        Countries(j).FirstCol = Area(1, 1).Column
        Countries(j).Name = Cells(BlockCount, 1).Value
        ' Move to last row of next block
        BlockCount = Cells(BlockCount, 1).End(xlUp).Row
    Next i
    
    For i = 1 To Count
    
        For j = 1 To 2
        
            Set MyChart(i, j) = MySheet.Shapes.AddChart(xlBarStacked).Chart
            
            With MyChart(i, j)
                
                LastRow = Countries(i).DataRange.Rows.Count

                Set ChartLoc = _
                    Range(Countries(i).DataRange(1, 1).Offset(0, 8), _
                    Countries(i).DataRange(1, 1).Offset(ChartHeight, 8 + ChartWidth))
                
                With MyChart(i, j).Parent
                    .Height = ChartLoc.Height
                    .Width = ChartLoc.Width
                    .Top = ChartLoc.Top
                    .Left = ChartLoc.Left
                End With
                
                ' Source top 5
                If j = 1 Then
                    .SetSourceData Source:=Range( _
                        Countries(i).DataRange(1, 2).Address & ":" & _
                        Countries(i).DataRange(5, 2).Address & "," & _
                        Countries(i).DataRange(1, 4).Address & ":" & _
                        Countries(i).DataRange(5, 5).Address _
                        )

                Else
                ' Source bottom 5
                    .SetSourceData Source:=Range( _
                        Countries(i).DataRange(LastRow, 2).Address & ":" & _
                        Countries(i).DataRange(LastRow - 4, 2).Address & "," & _
                        Countries(i).DataRange(LastRow, 4).Address & ":" & _
                        Countries(i).DataRange(LastRow - 4, 5).Address _
                        )
                    MyChart(i, j).Parent.Top = _
                        ChartLoc.Offset(ChartHeight + 1, 0).Top
                End If
                
                ' Rename series
                .FullSeriesCollection(1).Name = Cells(Area(1, 1).Row, 4).Value
                .FullSeriesCollection(2).Name = Cells(Area(1, 1).Row, 5).Value
                
                ' Chart formatting bits & bobs
                .ChartGroups(1).GapWidth = 75
                .Axes(xlValue).MajorGridlines.Delete
                .Legend.Position = xlLegendPositionBottom
                .HasTitle = True
                .ChartTitle.text = Countries(i).Name
                
                Set AuxTitle = MyChart(i, j).Shapes.AddLabel(msoTextOrientationHorizontal, _
                    MyChart(i, j).ChartTitle.Left + 1, _
                    MyChart(i, j).ChartTitle.Top + 20, _
                    60, 19)
                AuxTitle.TextFrame2.TextRange.Characters.text = "€ Millions"
                
                .Axes(xlValue).TickLabels.NumberFormat = "#,##0_);[Red](#,##0)"
            
                ' Colors
                ' Local = Navy
                With .SeriesCollection(1).Format.Fill
                    .Visible = msoTrue
                    .ForeColor.ObjectThemeColor = msoThemeColorAccent1
                    .ForeColor.TintAndShade = 0
                    .ForeColor.Brightness = -0.5
                    .Transparency = 0
                    .Solid
                End With
                
                ' CB = Pale Green
                With .SeriesCollection(2).Format.Fill
                    .Visible = msoTrue
                    .ForeColor.ObjectThemeColor = msoThemeColorAccent3
                    .ForeColor.TintAndShade = 0
                    .ForeColor.Brightness = 0.400000006
                    .Transparency = 0
                    .Solid
                End With
            
            End With
        
        Next j
            
        ' Group together
        MySheet.Shapes.Range(Array(MyChart(i, 1).Parent.Name, _
            MyChart(i, 2).Parent.Name)).Group
        
    Next i
    
End Sub

Sub MktShareChart()

    Dim FundTypes(1 To 2)       As CMetaData
    Dim Area(2, 2), ChartLoc    As Range
    Dim MySheet                 As Worksheet
    Dim MyChart(2, 2)           As Chart
    Dim i, j, Count, BlockCount As Long
    Dim tmpCol                  As Long
    Dim ChartWidth, ChartHeight As Long
    Dim FNames(1 To 2)          As String
    Dim ManagerLabels(1 To 2)   As Range
    Dim TitlePre                As String
    Dim tmpCell                 As Range
    
    Set MySheet = ActiveSheet
    Set Area(2, 1) = Cells(Rows.Count, 1).End(xlUp)
    Set Area(1, 1) = Cells(1, 1)  ' header row; blank row before data
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 2) = Cells(Area(2, 1).Row, Area(1, 2).Column)
    FNames(1) = "Bond"
    FNames(2) = "Equity"
    ChartWidth = 8
    ChartHeight = 20
    TitlePre = " Market Share in Europe - "
        
    ' Initialize & define FundType CMetaData classes
    For i = 1 To 2
        Set FundTypes(i) = New CMetaData
        FundTypes(i).Name = FNames(i)
        
        With MySheet
            .Range(Area(1, 1), Area(2, 2)).AutoFilter _
                Field:=Area(1, 1).Column, Criteria1:=FNames(i)
            FundTypes(i).FirstRow = _
                .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Rows(1).Row
            FundTypes(i).LastRow = _
                .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1, 1).End(xlDown).Row
            FundTypes(i).FirstCol = Area(1, 1).Column
            FundTypes(i).LastCol = Area(1, 2).Column
            
            Set tmpCell = Cells(FundTypes(i).FirstRow, FundTypes(i).LastCol)
            Set ManagerLabels(i) = Range( _
                tmpCell.End(xlToRight), tmpCell.End(xlToRight).End(xlDown))
            
            .AutoFilterMode = False
        End With
        
    Next i
    
    ' Create charts
    For i = 1 To 2
    
        For j = 1 To 2
            
            Set MyChart(i, j) = MySheet.Shapes.AddChart(xlPie).Chart
            
            With MyChart(i, j)

                Set ChartLoc = _
                    Range(FundTypes(i).DataRange(1, 1).Offset( _
                        0, 8 + (ChartWidth + 1) * (j - 1)), _
                    FundTypes(i).DataRange(1, 1).Offset( _
                        ChartHeight, 8 + (ChartWidth) * j))
                
                With MyChart(i, j).Parent
                    .Height = ChartLoc.Height
                    .Width = ChartLoc.Width
                    .Top = ChartLoc.Top
                    .Left = ChartLoc.Left
                End With
                
                ' Set source
                If j = 1 Then
                    tmpCol = 3
                Else
                    tmpCol = Area(1, 2).Column
                End If
                
                .SetSourceData Source:=Range( _
                    FundTypes(i).DataRange(1, tmpCol).Address & ":" & _
                    FundTypes(i).DataRange(6, tmpCol).Address)
                
                ' Set labels
                .FullSeriesCollection(1).XValues = _
                    ManagerLabels(i)
                
                ' Labelling & title
                .HasTitle = True
                If j = 1 Then
                    .ChartTitle.text = FundTypes(i).Name & TitlePre & "3-Month Gross Sales"
                Else
                    .ChartTitle.text = FundTypes(i).Name & TitlePre & "Prior Year Same 3-Month Gross Sales"
                End If
                .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
                
                If i = 1 Then
                    ' Bond = Green
                    With .ChartTitle.Format.TextFrame2.TextRange.Font.Fill
                        .Visible = msoTrue
                        .ForeColor.ObjectThemeColor = msoThemeColorAccent3
                        .ForeColor.TintAndShade = 0
                        .ForeColor.Brightness = -0.25
                        .Transparency = 0
                        .Solid
                    End With
                Else
                    ' Equity = Navy
                    With .ChartTitle.Format.TextFrame2.TextRange.Font.Fill
                        .Visible = msoTrue
                        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
                        .ForeColor.TintAndShade = 0
                        .ForeColor.Brightness = -0.5
                        .Transparency = 0
                        .Solid
                    End With
                End If
                
                With .FullSeriesCollection(1)
                    .ApplyDataLabels
                    .DataLabels.ShowPercentage = True
                    .DataLabels.ShowValue = False
                End With
                
            End With
    
        Next j
    
    Next i

    Cells(1, 1).Select

End Sub

Sub BubbleChart(TitleText As String, xLabelText As String, yLabelText As String)

    Dim DataArea                As CMetaData
    Dim Area(2, 2)              As Range
    Dim ChartLoc                As Range
    Dim MySheet                 As Worksheet
    Dim MyChart                 As Chart
    Dim ChartWidth, ChartHeight As Long
    Dim ManagerLabels(1 To 2)   As Range
    Dim TextBoxText             As String
    Dim AxisLabels(1 To 2)      As String
    Dim r                       As Long
    Dim AxisType                As Variant
    Dim yLabel, Key             As Shape
    
    
    Set MySheet = ActiveSheet
    
    ' Define contents area
    Set Area(1, 1) = Cells(1, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight).Offset(0, -1)  ' ignore No Nulls column
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).Row, Area(1, 2).Column)
    AxisLabels(1) = "1-Year Return in Euro Weighted Average"  ' x-axis
    AxisLabels(2) = "YTD TR in Euro Weighted Average"  ' y-axis
    
    ' Instantiate & define DataArea
    Set DataArea = New CMetaData
    DataArea.FirstRow = Area(1, 1).Offset(1, 0).Row
    DataArea.LastRow = Area(2, 1).Row
    DataArea.FirstCol = Area(1, 1).Column
    DataArea.LastCol = Area(2, 2).Column
    
    ' Set chart params
    ChartWidth = 13
    ChartHeight = 29
    Set ChartLoc = Range(Cells(DataArea.FirstRow, DataArea.LastCol + 2), _
        Cells(DataArea.FirstRow + ChartHeight, DataArea.LastCol + 2 + ChartWidth))
    
    Set MyChart = MySheet.Shapes.AddChart(xlBubble).Chart
    
        With MyChart
        
            ' Set chart location
            With MyChart.Parent
            
                .Height = ChartLoc.Height
                .Width = ChartLoc.Width
                .Top = ChartLoc.Top
                .Left = ChartLoc.Left
            
            End With
            
            ' Set style
            .ClearToMatchStyle
            .ChartStyle = 269
            .ChartColor = 21

            For r = 1 To DataArea.DataRange.Rows.Count
                
                If r > .SeriesCollection.Count Then
                    .SeriesCollection.NewSeries
                End If
                
                .SeriesCollection(r).Name = DataArea.DataRange(r, 1)
                .SeriesCollection(r).XValues = DataArea.DataRange(r, 4)
                .SeriesCollection(r).Values = DataArea.DataRange(r, 3)
                .SeriesCollection(r).BubbleSizes = DataArea.DataRange(r, 2)
                
                ' Labels
                .SeriesCollection(r).Points(1).ApplyDataLabels
                .SeriesCollection(r).DataLabels.ShowSeriesName = True
                .SeriesCollection(r).DataLabels.ShowValue = False
                .SeriesCollection(r).DataLabels.Position = xlLabelPositionCenter
                
            Next r
            
            ' Chart formatting
            For Each AxisType In Array(xlCategory, xlValue)
            
                With MyChart.Axes(AxisType)
                    
                    .MajorGridlines.Delete
                    .TickLabels.Font.Size = 14
                    .TickLabels.NumberFormat = "0%"
                    
                    If AxisType = xlCategory Then
                        .HasTitle = True
                        .AxisTitle.Font.Size = 12
                        .AxisTitle.text = xLabelText
                        .AxisTitle.Font.Bold = False
                    End If
                                    
                End With
            
            Next AxisType
            
            .HasLegend = False
            .HasTitle = True
            .ChartTitle.text = " "
            
            ' Adjust for squished plot area
            Application.DisplayAlerts = False
            With .PlotArea
                .Top = MyChart.ChartArea.Top + 10
                .Left = MyChart.ChartArea.Left
                .Width = MyChart.ChartArea.Width
                .Height = MyChart.ChartArea.Height - 40
            End With
            Application.DisplayAlerts = True
            
            ' Add text boxes (ie yLabel & Key)
            Set yLabel = MyChart.Shapes.AddLabel(msoTextOrientationHorizontal, _
                MyChart.PlotArea.Left + 20, _
                MyChart.PlotArea.Top - 20, _
                180, 19)
            yLabel.TextFrame2.TextRange.Characters.text = yLabelText
            yLabel.TextFrame2.TextRange.Font.Size = 12
            Set Key = MyChart.Shapes.AddLabel(msoTextOrientationHorizontal, _
                MyChart.PlotArea.Left + MyChart.PlotArea.Width - 100, _
                MyChart.PlotArea.Top, _
                120, 20)
            Key.TextFrame2.TextRange.Characters.text = "Bubble Size = Net Sales in Euro"
            Key.TextFrame2.TextRange.Font.Size = 12
                
        End With
        
        ChartLoc.Cells(1, 1).Offset(-1, 1).Value = TitleText
End Sub

Sub LvCBChart()

    Dim DataArea                As CMetaData
    Dim Area(2, 2)              As Range
    Dim ChartLoc, StartCol      As Range
    Dim MySheet                 As Worksheet
    Dim MyChart                 As Chart
    Dim ChartWidth, ChartHeight As Long
    
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(1, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).Row, Area(1, 2).Column)
    ChartWidth = 8
    ChartHeight = 20
    
    Set DataArea = New CMetaData
    DataArea.FirstRow = Area(1, 1).Offset(1, 0).Row  ' ignore header row
    DataArea.LastRow = Area(2, 1).Row
    DataArea.FirstCol = Area(1, 1).Column
    DataArea.LastCol = Area(1, 2).Column
    
    Set StartCol = DataArea.DataRange(1, 1).End(xlToRight).Offset(0, 2)
    Set ChartLoc = Range(StartCol, StartCol.Offset(ChartHeight - 1, ChartWidth - 1))
    
    Set MyChart = ActiveSheet.Shapes.AddChart(xlColumnStacked).Chart
    
    With MyChart
        
        .SetSourceData Source:=DataArea.DataRange
        .SeriesCollection(1).Name = "Local Net Sales"
        .SeriesCollection(2).Name = "Cross-Border Net Sales"
        
        ' Location
        With .Parent
            .Height = ChartLoc.Height
            .Width = ChartLoc.Width
            .Top = ChartLoc.Top
            .Left = ChartLoc.Left
        End With
        
        ' Axes
        With .Axes(xlValue)
            .MajorGridlines.Delete
            .DisplayUnit = xlThousands
            .TickLabels.NumberFormat = "#,##0.0"
            .HasDisplayUnitLabel = False
            .TickLabels.Font.Size = 10
        End With
        
        ' Labelling
        .HasTitle = True
        With .ChartTitle
            .text = "€ Billions"
            .Format.TextFrame2.TextRange.Font.Size = 11
            .Format.TextFrame2.TextRange.Font.Italic = msoTrue
            .Format.TextFrame2.TextRange.Font.Bold = msoFalse
        End With
        
        .HasLegend = True
        With .Legend
            .Left = MyChart.ChartTitle.Left - 120
            .Top = MyChart.ChartTitle.Top + 30
            .Height = 20
            .Width = 180
        End With
        
        ' Resize PlotArea
        .PlotArea.Width = MyChart.ChartArea.Width - 20
        
        ' Colors
        ' Local = Navy
        With .SeriesCollection(1).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
            .Solid
        End With
        
        ' CB = Pale Green
        With .SeriesCollection(2).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent3
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.400000006
            .Transparency = 0
            .Solid
        End With
        
    End With

End Sub

Sub ManagerByCtryChart()

    Dim Area(2, 2), DataRange   As Range
    Dim ChartLoc, StartCell     As Range
    Dim MySheet                 As Worksheet
    Dim MyChart                 As Chart
    Dim ChartWidth, ChartHeight As Long

    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(2, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).Row, Area(1, 2).Column)
    Set DataRange = Range(Area(1, 1), Area(2, 2).Offset(0, -3))
    
    ' Chart params
    ChartWidth = 6
    ChartHeight = 20
    Set StartCell = Area(2, 1).Offset(2, 0)
    Set ChartLoc = Range(StartCell, StartCell.Offset(ChartHeight - 1, ChartWidth - 1))
    
    Set MyChart = MySheet.Shapes.AddChart(xlColumnStacked).Chart
    
    With MyChart
    
        .SetSourceData Source:=DataRange
        .SeriesCollection(2).PlotOrder = 1  ' Ensure Top 3 on top
        
        With .Parent
            .Height = ChartLoc.Height
            .Width = ChartLoc.Width
            .Top = ChartLoc.Top
            .Left = ChartLoc.Left
        End With
        
        ' Gap Width
        .ChartGroups(1).GapWidth = 150
        
        ' Axes
        With .Axes(xlValue)
            .MajorGridlines.Delete
            .DisplayUnit = xlThousands
            .TickLabels.NumberFormat = "#,##0.0"
            .HasDisplayUnitLabel = False
            .TickLabels.Font.Size = 10
            .HasTitle = True
            
            With .AxisTitle
                .Orientation = xlHorizontal
                .text = "€ Billions"
                .Font.Bold = msoFalse
                .Font.Italic = msoTrue
                .Top = 3
                .Left = 5
            End With
            
        End With
        
        ' Labelling
        .HasLegend = True
        With .Legend
            .Top = MyChart.PlotArea.Top
            .Left = MyChart.PlotArea.Left + 40
            .Width = MyChart.PlotArea.Width
            .Height = 20
        End With
        
        ' Resize/reorient plot area
        With .PlotArea
            .Top = 20
            .Left = 5
            .Width = MyChart.ChartArea.Width - 20
        End With
        
        ' Colors
        ' Top 3 = Navy
        With .SeriesCollection(2).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
            .Solid
        End With
        
        ' Middle = Gray stripes
        With .SeriesCollection(1).Format.Fill
            .Visible = msoTrue
            .Patterned msoPatternDarkUpwardDiagonal
            .BackColor.RGB = RGB(255, 255, 255)
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
        End With
        
        ' Bottom 3 = Pink
        With .SeriesCollection(3).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent2
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.8000000119
            .Transparency = 0
            .Solid
        End With
        
        
    End With

    Cells(1, 1).Select

End Sub

Sub GrossByRegionChart()

    Dim Area(2, 2), DataRange   As Range
    Dim ChartLoc, StartCell     As Range
    Dim MySheet                 As Worksheet
    Dim MyChart                 As Chart
    Dim ChartWidth, ChartHeight As Long
    
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(1, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).Row, Area(1, 2).Column)
    Set DataRange = Range(Area(1, 1), Area(2, 2))
    
    ' Chart params
    ChartWidth = 10
    ChartHeight = 20
    Set StartCell = Cells(Area(2, 1).Offset(7, 0).Row, 1)
    Set ChartLoc = Range(StartCell, _
        StartCell.Offset(ChartHeight - 1, ChartWidth - 1))
    
    Set MyChart = MySheet.Shapes.AddChart(xlColumnStacked100).Chart
    
    With MyChart
    
        .SetSourceData Source:=DataRange
        .PlotBy = Excel.XlRowCol.xlRows  ' plot by rows (default by cols, ie 3 series)
        .SeriesCollection(2).PlotOrder = 1
        
        With .Parent
            .Height = ChartLoc.Height
            .Width = ChartLoc.Width
            .Top = ChartLoc.Top
            .Left = ChartLoc.Left
        End With
        
        .ChartGroups(1).GapWidth = 80
        
        ' Axes & Legend
        .Axes(xlValue).MajorGridlines.Delete
        .Axes(xlValue).TickLabels.Font.Size = 12
        .Axes(xlCategory).TickLabels.Font.Size = 11
        
        With .Legend
            .Position = xlLegendPositionTop
            .Left = MyChart.PlotArea.Left
            .Width = MyChart.PlotArea.Width
            .Font.Size = 12
        End With
        
        ' Colors
        ' Europe = Navy
        With .SeriesCollection(1).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText2
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
            .Transparency = 0
            .Solid
        End With
        
        ' Asia Pacific = Light Blue
        With .SeriesCollection(2).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.400000006
            .Transparency = 0
            .Solid
        End With
        
        ' Rest of World = Dot-ey
        With .SeriesCollection(3).Format.Fill
            .Visible = msoTrue
            .Patterned msoPattern50Percent
            .BackColor.RGB = RGB(255, 255, 255)
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
        End With
        
    
    End With
    
    Cells(1, 1).Select

End Sub

Sub EquityCatSalesChart()

    Dim Area(2, 2), tmpRange    As Range
    Dim ChartLoc, StartCell     As Range
    Dim MySheet                 As Worksheet
    Dim MyChart                 As Chart
    Dim ChartWidth, ChartHeight As Long
    Dim NetSales, ATR           As CMetaData
    Dim Months, i               As Long
    Dim Headers                 As Range

    Set MySheet = ActiveSheet
    Set NetSales = New CMetaData
    Set ATR = New CMetaData
    Set tmpRange = Range(Cells(1, 1).End(xlToRight), Cells(1, 1).End(xlToRight).End(xlToRight))
    Months = tmpRange.Columns.Count - 1
    Set Area(1, 1) = Cells(2, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).Row, Area(1, 2).Column)
    
    ' Define Header range
    Set Headers = Range(Area(1, 1).Offset(0, 1), Area(1, 1).Offset(0, Months))
    Headers.Select
    
    ' Define DataRange
    NetSales.FirstRow = Area(1, 1).Offset(1, 0).Row
    NetSales.LastRow = Area(2, 1).Row
    NetSales.FirstCol = Area(1, 1).Offset(1, 1).Column
    NetSales.LastCol = Area(1, 1).Offset(1, Months).Column
    
    ATR.FirstRow = NetSales.FirstRow
    ATR.LastRow = NetSales.LastRow
    ATR.FirstCol = NetSales.LastCol + 1
    ATR.LastCol = Area(1, 2).Column
    
    ' Chart params
    ChartWidth = 12
    ChartHeight = 30
    Set StartCell = Cells(Area(2, 1).Offset(2, 0).Row, 1)
    Set ChartLoc = Range(StartCell, _
        StartCell.Offset(ChartHeight - 1, ChartWidth - 1))
    
    Set MyChart = MySheet.Shapes.AddChart(xlColumnClustered).Chart
    
    With MyChart
    
        .SetSourceData Source:=NetSales.DataRange.Rows(1)
        .SeriesCollection(1).XValues = Headers
        
        ' Define series (name & range)
        ' This could be refactored to be simpler;
        ' thought that order would affect legend layout (plot type overrides though)
        For i = 1 To 3
            If i <> 1 Then
                .SeriesCollection.NewSeries
                With .SeriesCollection(2 * i - 1)
                    .Name = Area(1, 1).Offset(i, 0).Value
                    .Values = NetSales.DataRange.Rows(i)
                    .AxisGroup = xlPrimary
                    .ChartType = xlColumnClustered
                End With
            Else
                .SeriesCollection(i).Name = Area(1, 1).Offset(i, 0).Value
            End If
            
            ' Even Series are ATR
            .SeriesCollection.NewSeries
            With .SeriesCollection(2 * i)
                .Name = "Avg TR of " & Area(1, 1).Offset(i, 0).Value
                .Values = ATR.DataRange.Rows(i)
                .AxisGroup = xlSecondary
                .ChartType = xlLine
            End With
        Next i

        ' Set chart location
        With .Parent
            .Width = ChartLoc.Width
            .Height = ChartLoc.Height
            .Top = ChartLoc.Top
            .Left = ChartLoc.Left
        End With
    
        ' Axes
        ' LHS y-axis
        With .Axes(xlValue, xlPrimary)
            .TickLabels.Font.Size = 12
            .MajorGridlines.Delete
            .DisplayUnit = xlThousands
            .TickLabels.NumberFormat = "#,##0"
            .HasDisplayUnitLabel = False
            .HasTitle = True
            
            With .AxisTitle
                .Orientation = xlHorizontal
                .text = "€ Billions"
                .Font.Size = 12
                .Font.Bold = msoFalse
                .Font.Italic = msoTrue
                .Top = 0
                .Left = 5
            End With
            
        End With
        
        ' RHS y-axis
        With .Axes(xlValue, xlSecondary)
            .TickLabels.Font.Size = 12
            .TickLabels.NumberFormat = "#,##0"
            .HasDisplayUnitLabel = False
            .HasTitle = True
            
            With .AxisTitle
                .Orientation = xlHorizontal
                .text = "%"
                .Font.Size = 12
                .Font.Bold = msoFalse
                .Font.Italic = msoTrue
                .Top = 0
                .Left = MyChart.ChartArea.Width - 25
            End With
            
        End With
        
        ' x-axis
        With .Axes(xlCategory)
            .TickLabels.Font.Size = 12
            .TickLabelPosition = xlTickLabelPositionLow
        End With
        
        ' Legend
        With .Legend
            .Top = MyChart.PlotArea.Top + 10
            .Left = MyChart.ChartArea.Left + 25
            .Height = 30
            .Width = MyChart.ChartArea.Width - 50
            .Font.Size = 10
        End With
    
        ' PlotArea adjustments
        .PlotArea.Top = 20
        .PlotArea.Left = .ChartArea.Left
        .PlotArea.Width = .ChartArea.Width
        
        ' Colors
        ' Equity Emerging Markets = Navy
        With .SeriesCollection(1).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText2
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
            .Transparency = 0
            .Solid
        End With
        
        ' Equity Europe = Light Blue
        With .SeriesCollection(3).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.400000006
            .Transparency = 0
            .Solid
        End With
        
        ' Equity North America = Gray stripey
        With .SeriesCollection(5).Format.Fill
            .Patterned msoPatternLightUpwardDiagonal
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.349999994
        .Transparency = 0
        End With
        
        ' Avg TR of Equity Emerging Market = Green line
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(152, 185, 84)
        .SeriesCollection(2).Format.Line.Weight = 2.25
        
        ' Avg TR of Equity Europe = Gray dotted line
        With .SeriesCollection(4).Format.Line
            .DashStyle = msoLineSysDash
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.6000000238
            .Transparency = 0.0099999905
        End With
        
        ' Avg TR of Equity North America = Navy line
        With .SeriesCollection(6).Format.Line
            .Visible = msoTrue
            .Weight = 2.25
            .ForeColor.ObjectThemeColor = msoThemeColorText2
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
            .Transparency = 0
        End With
        
    End With
    
    Cells(1, 1).Select
    
End Sub

Sub ManagerBubbleChart(Optional TitleText As String = "")

    Dim DataArea                As CMetaData
    Dim Area(2, 2)              As Range
    Dim ChartLoc                As Range
    Dim MySheet                 As Worksheet
    Dim MyChart                 As Chart
    Dim ChartWidth, ChartHeight As Long
    Dim ManagerLabels(1 To 2)   As Range
    Dim TextBoxText             As String
    Dim AxisLabels(1 To 2)      As String
    Dim r, i                    As Long
    Dim AxisType                As Variant
    Dim yLabel                  As Shape
    Dim BubbleLab(1 To 2)       As String
    
    
    Set MySheet = ActiveSheet
    BubbleLab(1) = "12-Month Net Sales in Euro Billions"
    BubbleLab(2) = "Asset-Weighted 1-Year Total Return"
    
    ' Define contents area
    Set Area(1, 1) = Cells(1, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight).Offset(0, -1)  ' ignore No Nulls column
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).Row, Area(1, 2).Column)
    AxisLabels(1) = "1-Year Return in Euro Weighted Average"  ' x-axis
    AxisLabels(2) = "YTD TR in Euro Weighted Average"  ' y-axis
    
    ' Instantiate & define DataArea
    Set DataArea = New CMetaData
    DataArea.FirstRow = Area(1, 1).Offset(1, 0).Row
    DataArea.LastRow = Area(2, 1).Row
    DataArea.FirstCol = Area(1, 1).Column
    DataArea.LastCol = Area(2, 2).Column
    
    ' Set chart params
    ChartWidth = 8
    ChartHeight = 20
    Set ChartLoc = Range(Cells(DataArea.FirstRow, DataArea.LastCol + 2), _
        Cells(DataArea.FirstRow + ChartHeight, DataArea.LastCol + 2 + ChartWidth))
    
    Set MyChart = MySheet.Shapes.AddChart(xlXYScatter).Chart
    
        With MyChart
        
            ' Set chart location
            With MyChart.Parent
            
                .Height = ChartLoc.Height
                .Width = ChartLoc.Width
                .Top = ChartLoc.Top
                .Left = ChartLoc.Left
            
            End With
            
            ' Set style
            .ClearToMatchStyle
            .ChartStyle = 269
            .ChartColor = 21

            For r = 1 To DataArea.DataRange.Rows.Count
                
                If r > .SeriesCollection.Count Then
                    .SeriesCollection.NewSeries
                End If
                
                .SeriesCollection(r).Name = DataArea.DataRange(r, 1)
                .SeriesCollection(r).XValues = DataArea.DataRange(r, 2)
                .SeriesCollection(r).Values = DataArea.DataRange(r, 3)
                
                ' Labels
                .SeriesCollection(r).Points(1).ApplyDataLabels
                .SeriesCollection(r).DataLabels.ShowSeriesName = True
                .SeriesCollection(r).DataLabels.ShowValue = False
                .SeriesCollection(r).DataLabels.Position = xlLabelPositionAbove
                
            Next r
            
            ' Axis, Legend, Gridlines
            For Each AxisType In Array(xlCategory, xlValue)
            
                With MyChart.Axes(AxisType)
                    
                    .TickLabels.Font.Size = 14
                    .TickLabels.NumberFormat = "0%"
                    ' Axis label colors
                    .TickLabels.Font.Color = RGB(146, 208, 80)
                    
                    If AxisType = xlCategory Then
                        .MajorGridlines.Delete
                        .HasTitle = True
                        .DisplayUnit = xlThousands
                        .HasDisplayUnitLabel = False
                        .TickLabels.NumberFormat = "#,##0"
                        .AxisTitle.Font.Size = 12
                        .AxisTitle.text = BubbleLab(1)
                        .AxisTitle.Font.Bold = False
                    Else
                        ' Horizontal Gridlines
                        With .MajorGridlines.Format.Line
                            .DashStyle = msoLineDashDot
                            .Visible = msoTrue
                            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
                            .ForeColor.TintAndShade = 0
                            .ForeColor.Brightness = -0.150000006
                            .Transparency = 0
                        End With
                    End If
                                    
                End With
            
            Next AxisType
            
            .HasLegend = False
                                    
            ' Adjust for squished plot area
            Application.DisplayAlerts = False
            With .PlotArea
                
                .Top = 20
                .Left = 0
                .Width = MyChart.ChartArea.Width
                .Height = MyChart.ChartArea.Height - 40
            End With
            Application.DisplayAlerts = True
            
            ' Add text boxes (ie yLabel)
            Set yLabel = MyChart.Shapes.AddLabel(msoTextOrientationHorizontal, _
                MyChart.PlotArea.Left + 20, _
                MyChart.PlotArea.Top - 20, _
                180, 19)
            yLabel.TextFrame2.TextRange.Characters.text = BubbleLab(2)
            yLabel.TextFrame2.TextRange.Font.Size = 12
            
            ' Colors
            .ChartColor = 13
            
            ' Vary point shapes
            For i = 1 To .SeriesCollection.Count
                
                With .SeriesCollection(i)
                    .MarkerStyle = i Mod 7 + 1
                    .MarkerSize = 8
                    .Format.Fill.Transparency = 0.5
                End With
            
            Next i
            
        End With
        
        ChartLoc.Cells(1, 1).Offset(-1, 1).Value = TitleText
        
        Cells(1, 1).Select

End Sub

Sub MSRegionChart(Optional kind As String = "countries")

    ' Builds MS Region chart using Net Sales
    
    Dim DataRange(1 To 2)       As CMetaData
    Dim Area(2, 2)              As Range
    Dim ChartLoc(1 To 2)        As Range
    Dim StartCell, Measure      As Range
    Dim MySheet                 As Worksheet
    Dim MyChart(1 To 2)         As Chart
    Dim ChartWidth, ChartHeight As Long
    Dim Months, tmpCell, i, j   As Long
    Dim Headers                 As Range
    Dim TitleBox                As Shape
    
    
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(2, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).Offset(0, 1).End(xlDown).Offset(0, -1)
    Set Area(2, 2) = Cells(Area(2, 1).Row, Area(1, 2).Column)
    
    ' Calculate number of periods & Europe start row
    Months = Range(Cells(1, 1).End(xlToRight), Cells(1, 1).End(xlToRight).End(xlToRight)).Columns.Count - 1
    tmpCell = Cells(Rows.Count, 1).End(xlUp).Row  ' Europe data start row
    
    ' Instantiate CMetaData objs
    For i = 1 To 2
        Set DataRange(i) = New CMetaData
    Next i
    
    ' Chart params
    ChartWidth = 12
    ChartHeight = 30
    
    For i = 1 To 2
        Set StartCell = Area(2, 1).Offset(2 + (i - 1) * ChartHeight, 0)
        Set ChartLoc(i) = Range(StartCell, StartCell.Offset(ChartHeight - 2, ChartWidth - 1))
    Next i
    
    ' Define data ranges
    Set Headers = Range(Area(1, 1).Offset(0, 2), Area(1, 1).Offset(0, 1 + Months))
    
    DataRange(1).FirstRow = Area(1, 1).Offset(1, 0).Row
    DataRange(1).LastRow = tmpCell - 1
    DataRange(1).FirstCol = Area(1, 1).Offset(0, 1).Column
    DataRange(1).LastCol = Area(1, 2).Column

    DataRange(2).FirstRow = tmpCell
    DataRange(2).LastRow = Area(2, 1).Row
    DataRange(2).FirstCol = Area(1, 1).Offset(0, 1).Column
    DataRange(2).LastCol = Area(1, 2).Column
    
    If kind = "countries" Then
        DataRange(1).Name = "Asia"
        DataRange(2).Name = "Europe"
    ElseIf kind = "cb v local" Then
        DataRange(1).Name = "Europe Cross-border"
        DataRange(2).Name = "Europe Local"
    End If
    
    ' Build Chart
    
    For i = 1 To 2
        
        Set Measure = Range(DataRange(i).DataRange(1, 1).Offset(0, Months + 1), _
            DataRange(i).DataRange(1, 1).Offset(DataRange(i).DataRange.Rows.Count - 1, DataRange(i).DataRange.Columns.Count - 1))
        Set MyChart(i) = MySheet.Shapes.AddChart(xlColumnStacked).Chart
        
        With MyChart(i)
        
            .SetSourceData Source:=Measure, PlotBy:=Excel.XlRowCol.xlRows
            
            For j = 1 To .SeriesCollection.Count
                With .SeriesCollection(j)
                    .Name = DataRange(i).DataRange(j, 1)
                    .XValues = Headers
                End With
            Next j
        
            .ChartGroups(1).GapWidth = 85
        
        ' Set chart location
        With .Parent
            .Height = ChartLoc(i).Height
            .Width = ChartLoc(i).Width
            .Top = ChartLoc(i).Top
            .Left = ChartLoc(i).Left
        End With
        
        ' Axes
        With .Axes(xlCategory)
            .TickLabels.Font.Size = 11
            .TickLabelPosition = xlTickLabelPositionLow
            .TickLabels.Orientation = 30
        End With
        
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.text = "€ Billions"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Italic = msoTrue
            .AxisTitle.Font.Bold = msoFalse
            .AxisTitle.Orientation = xlHorizontal
            .AxisTitle.Top = -5
            .AxisTitle.Left = 8
            .MajorGridlines.Delete
            .TickLabels.Font.Size = 11
            .DisplayUnit = xlThousands
            .HasDisplayUnitLabel = False
        End With
        
        ' Legend
        With .Legend
            .Font.Name = "Wingdings"
            .Font.Size = 11
            .Left = MyChart(i).ChartArea.Width - 350
            .Height = 30
            .Top = MyChart(i).ChartArea.Height - 85
            .Width = 300
        End With

        ' Title
        .HasTitle = True
        .ChartTitle.text = " "
        
        Set TitleBox = MyChart(i).Shapes.AddLabel(msoTextOrientationHorizontal, _
            MyChart(i).ChartTitle.Left - 78, _
            MyChart(i).ChartTitle.Top, _
            145, 18)
            
        With TitleBox
            .TextFrame2.TextRange.Characters.text = DataRange(i).Name
            .TextFrame2.TextRange.Font.Size = 11
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .Line.Visible = msoTrue
            .Line.ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .Line.ForeColor.Brightness = -0.25
            .TextFrame2.TextRange.ParagraphFormat.Alignment = _
                msoAlignCenter
            
        End With
        
        ' Color
        ' 5 stars = Green
        With .SeriesCollection(5).Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 176, 80)
            .Solid
        End With
        
        ' 4 stars = Navy Stipes
        With .SeriesCollection(4).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
            .BackColor.ObjectThemeColor = msoThemeColorBackground1
            .BackColor.TintAndShade = 0
            .BackColor.Brightness = 0
            .Patterned msoPatternDarkUpwardDiagonal
        End With
        
        ' 3 stars = Grey w/White Dots
        With .SeriesCollection(3).Format.Fill
            .Visible = msoTrue
            .Patterned msoPattern20Percent
            .ForeColor.RGB = RGB(255, 255, 255)
            .BackColor.ObjectThemeColor = msoThemeColorBackground1
            .BackColor.TintAndShade = 0
            .BackColor.Brightness = -0.15
        End With
        
        
        ' 2 stars = Sky Blue
        With .SeriesCollection(2).Format.Fill
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.Brightness = 0.400000006
        End With
        
        ' 1 star = Navy
        With .SeriesCollection(1).Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText2
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
            .Transparency = 0
            .Solid
        End With
        
        ' Resize PlotArea
        With .PlotArea
            .Top = 20
            .Height = MyChart(i).ChartArea.Height - 20
            .Left = MyChart(i).ChartArea.Left
            .Width = MyChart(i).ChartArea.Width - 10
        End With
            
        End With
    
    Next i
    
    Cells(1, 1).Select
    
End Sub
    
Sub MTopBottomTable()

    Dim Area(2, 2), DataArea        As Range
    Dim PasteArea, Header           As Range
    Dim MySheet                     As Worksheet
    Dim NumCountries, i             As Long
    
    
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(1, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).Offset(0, 1).End(xlDown).Offset(0, -1)
    Set Area(2, 2) = Cells(Area(2, 1).Row, Area(1, 2).Column)
    Set DataArea = Range(Area(1, 1).Offset(1, 1), Area(1, 1).Offset(4, 4))
    
    
    ' Create pretty pink boxes
    i = 0
    Do
        Set PasteArea = DataArea.Offset(1 + 2 * i, 5)
        Set Header = PasteArea.Rows(1).Offset(-1, 0)
        DataArea.Copy PasteArea
        
        With Header
            .Cells(1, 1).Value = DataArea(1, 1).Offset(0, -1).Value
            .Cells(1, 3).Value = "€m"
            .Cells(1, 4).Value = "Share"
            .Range(Cells(1, 3), Cells(1, 4)).HorizontalAlignment = xlRight
            .Font.Size = 8
            .Font.Bold = True
            .Interior.Pattern = xlSolid
            .Interior.Color = 16764159
            .BorderAround LineStyle:=xlContinuous
        End With
        
        With PasteArea
            .Font.Size = 8
            .Font.Bold = True
            .Interior.Pattern = xlSolid
            .Interior.Color = 16764159
            .BorderAround LineStyle:=xlContinuous
        End With
        
        ' Next Country
        i = i + 1
        Set DataArea = DataArea.Offset(4, 0)
        
    Loop Until DataArea(1, 1).Value = ""
    
    ' Final touches
    Range(Columns(PasteArea(1, 1).Column), Columns(PasteArea(1, 4).Column)).AutoFit
    
    Cells(1, 1).Select
    
End Sub

Sub EuroTRQuartileChart()
    
    Dim Area(2, 2), DataArea        As Range
    Dim ChartData                   As Range
    Dim MySheet                     As Worksheet
    Dim ChartLoc(1 To 2)            As Range
    Dim StartCell                   As Range
    Dim MyChart(1 To 2)             As Chart
    Dim ChartWidth, ChartHeight     As Long
    Dim Months, i, j                As Long
    
    
    Months = Range(Cells(1, 1).End(xlToRight), Cells(1, 1).End(xlToRight).End(xlToRight)).Columns.Count - 1
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(2, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).Row, Area(1, 2).Column)
    Set DataArea = Range(Area(1, 1), Area(2, 2))
    
    ' Set Chart location params
    ChartWidth = 14
    ChartHeight = 28
    For i = 1 To 2
        Set StartCell = Area(2, 1).Offset(2 + (i - 1) * (ChartHeight + 1), 0)
        Set ChartLoc(i) = Range(StartCell, StartCell.Offset(ChartHeight - 1, ChartWidth - 1))
    Next i
    
    For i = 1 To 2
    
        ' Create chart
        Set ChartData = Range(DataArea(1, 1).Offset(0, 1 + (i - 1) * Months), _
                DataArea(1, 1).Offset(4, i * Months))
        Set MyChart(i) = MySheet.Shapes.AddChart(xlColumnStacked).Chart
        
        With MyChart(i)
        
            .SetSourceData Source:=ChartData
            
            ' Add other series
            .SeriesCollection(1).Name = "1st Qt"
            .SeriesCollection(2).Name = "2nd Qt"
            .SeriesCollection(3).Name = "3rd Qt"
            .SeriesCollection(4).Name = "4th Qt"
            
            ' Reverse order so 1st Qt on top (when positive), etc
            .SeriesCollection(1).PlotOrder = 4
            .SeriesCollection(1).PlotOrder = 3
            .SeriesCollection(1).PlotOrder = 2
            
            .ChartGroups(1).GapWidth = 80
            
            ' Set chart location
            With .Parent
                .Top = ChartLoc(i).Top
                .Left = ChartLoc(i).Left
                .Height = ChartLoc(i).Height
                .Width = ChartLoc(i).Width
            End With
            
            ' Axes
            With .Axes(xlValue)
                .MajorGridlines.Delete
                .DisplayUnit = xlThousands
                .TickLabels.NumberFormat = "#,##0"
                .HasDisplayUnitLabel = False
                .HasTitle = True
                .AxisTitle.text = "€ Billions"
                .AxisTitle.Font.Size = 12
                .AxisTitle.Font.Italic = msoTrue
                .AxisTitle.Font.Bold = msoFalse
                .AxisTitle.Orientation = xlHorizontal
                .AxisTitle.Top = 0
                .AxisTitle.Left = 8
            End With
            
            With .Axes(xlCategory)
                .TickLabels.Font.Size = 12
                .TickLabelPosition = xlTickLabelPositionLow
                .TickLabels.Orientation = 30
            End With
            
            ' Legend
            With .Legend
                .Position = xlLegendPositionBottom
                .Top = .Top - 30
                .Left = .Left - 25
                .Width = 250
            End With
            
            ' Title
            .HasTitle = True
            With .ChartTitle
                .text = ChartData(1, 1).Offset(-1, 0).Value
                .Format.TextFrame2.TextRange.Font.Size = 12
                .Format.TextFrame2.TextRange.Font.Bold = False
            End With
            
            ' Color
            ' 1st Quartile = Navy
            With .SeriesCollection(4).Format.Fill
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorAccent1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = -0.5
                .Transparency = 0
                .Solid
            End With
            
            ' 2nd Quartile = Gray stripes
            With .SeriesCollection(3).Format.Fill
                .Visible = msoTrue
                .Patterned msoPatternDarkUpwardDiagonal
                .BackColor.RGB = RGB(255, 255, 255)
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = -0.25
            End With
            
            ' 3rd Quartile = Pink
            With .SeriesCollection(2).Format.Fill
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorAccent2
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = 0.8000000119
                .Transparency = 0
                .Solid
            End With
            
            ' 4th Quartile = Light Blue
            With .SeriesCollection(1).Format.Fill
                .Visible = msoTrue
                .Patterned msoPattern30Percent
                .ForeColor.ObjectThemeColor = msoThemeColorAccent1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = 0.400000006
                .BackColor.RGB = RGB(255, 255, 255)
                .Transparency = 0
            End With
            
            ' Resize PlotArea
            .PlotArea.Left = .ChartArea.Left
            .PlotArea.Width = .ChartArea.Width - 10
            .PlotArea.Top = 15
            .PlotArea.Height = .ChartArea.Height - 15
            
        End With
    
    Next i
    
    Cells(1, 1).Select
    
End Sub
