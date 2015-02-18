Attribute VB_Name = "CustomQuartileByTR"
Sub test()
    BuildCustomQuarterlyTR "Bond EUR Corporates", 15
End Sub

' Custom reporting (TRowe)
' finished 1/14/15

Sub BuildCustomQuarterlyTR(ByRef custText As String, ByRef Periods As Long, Optional ByVal Aggr As Long = 3)
    ' 1. Must be modified to work with different periods! (note that 1-Yr TR calculated using 4 quarters (hardcoded!))
    ' 2. For aggregating months to quarters, months must start and end on valid months (e.g. start on jan, end on march)
    ' 3. Note labelling assumes quarters!
    
    Dim TR, Flow        As Range
    Dim nRow, nCol      As Long
    Dim col, yr, qtr    As Integer
    Dim tmpRng          As Range
    
    
    Application.ScreenUpdating = False  ' To boost performance
    
    nRow = Cells(Rows.Count, 1).End(xlUp).row
    nCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Clear extraneous
    Range(Rows(nRow + 1), Rows(Rows.Count)).Delete
    Range(Cells(1, 1), Cells(Rows.Count, Columns.Count)).ClearFormats
    
    Set TR = Range(Cells(2, 2), Cells(nRow, 58))
    Set Flow = Range(Cells(2, 59), Cells(nRow, nCol))
    
    Rescale TR
    
    ' Aggregate Flow
    Flow.Select
    Aggregate Aggr, "SUM"
    
    ' Set labels (assumes quarters!)
    nCol = Range(Flow(1, 1), Flow(1, 1).End(xlToRight)).Columns.Count
    Flow.Rows(1).Offset(-1, 0).ClearContents
    TR.Rows(1).Offset(-1, 0).ClearContents
    col = 1
    Do While True
        Flow(1, col).Offset(-1, 0).Select
        For yr = 2011 To 2014
            
            For qtr = 1 To 4
            
                If col > nCol Then
                    Exit Do
                End If
                    
                Flow(1, col).Offset(-1, 0).Value = yr & "Q" & qtr
                TR(1, col).Offset(-1, 0).Value = Flow(1, col).Offset(-1, 0).Value & " 1-Yr Trailing TR"
                col = col + 1
                
            Next qtr
            
        Next yr
    Loop
    
    ' Aggregate TR
    TR.Select
    Aggregate Aggr, "GEOMEAN"

    ' Calculate 1-Yr Trailing
    Set TR = Range(Cells(2, 2), Cells(2, 2).End(xlToRight).End(xlDown))
    Set tmpRng = TR.Offset(TR.Rows.Count, 0)
    tmpRng(1, tmpRng.Columns.Count).Formula = "=IFERROR(GEOMEAN(" & TR(1, TR.Columns.Count).AddressLocal(False, False) & _
                                              ":" & TR(1, TR.Columns.Count).Offset(0, -3).Address(False, False) & ")," & _
                                              Chr(34) & Chr(34) & ")"
    tmpRng(1, tmpRng.Columns.Count).Copy tmpRng
    tmpRng.Copy
    TR.PasteSpecial xlPasteValues
    
    tmpRng.ClearContents
    Set tmpRng = Nothing
    
    ' Remove first four quarters (to match num obs of Flows)
    ' Check this code for future projects!
    Range(TR.Columns(1), TR.Columns(4)).Delete
    Range(Flow.Rows(1).Offset(-1, 0).Columns(1), Flow.Rows(1).Offset(-1, 0).Columns(4)).Delete xlToLeft
    
    ' Remove extra columns
    Set TR = Range(Cells(2, 2), Cells(2, 2).End(xlToRight).End(xlDown))
    Range(Columns(TR(1, TR.Columns.Count + 1).Column), Columns(Flow(1, 1).Column - 1)).Delete xlToLeft
    Set Flow = TR.Offset(0, TR.Columns.Count)

    CustomTRQuartiles custText:=custText, Yrs:=1, Periods:=Periods
    CustomTRQuartChart Periods

End Sub

Sub CustomTRQuartiles(ByRef custText As String, Yrs As Long, Optional Cutoff As Long, Optional ByVal Periods As Long = 13)

    ' This is a pre-processing step meant to be run prior to EuroTRQuartileData()
    ' for <Yrs>-year trailing ATR data
    ' Yrs: trailing period for TR geo avg to be calculated over
    ' Cutoff: max number of missing obs in TR trailing period that is acceptable

    Dim Area(2, 2), tmpRng, AvgTR   As Range
    Dim MySheet                     As Worksheet
    Dim i, tmpCol                   As Long
    Dim SumCB(), SumLocal()         As Double
    Dim CB, WAE, Quart              As Range
    Dim Headers                     As Range
    Dim F1, F2, L1, L2              As String
    
    
    ' Preliminaries
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(1, 1)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    
    ' Clear extraneous
    Range(Rows(Area(2, 1).row + 1), Rows(Rows.Count)).Delete
    Range(Cells(1, 1), Cells(Rows.Count, Columns.Count)).ClearFormats
    
    ' Specify ranges for Monthly WAE, CB & Local measures
    Set WAE = Range(Area(1, 1).Offset(1, 1), Area(2, 1).Offset(0, (Periods - 1) * Yrs + 1))
    Set CB = Range(Area(1, 1).Offset(1, WAE.Columns.Count + 1), _
        Area(2, 1).Offset(0, WAE.Columns.Count + Periods))
    
    ' Setup Start/Stop cells for Avg TR
    F1 = WAE(1, 1).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
    F2 = WAE(1, 2).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
    L1 = WAE(1, 1).Offset(0, (Periods - 1) * Yrs).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
    L2 = WAE(1, 1).Offset(0, (Periods - 1) * Yrs - 1).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
    
    ' Add (desc) Quartile at bottom (=QUARTILE.INC(IF(SUBTOTAL(...)...)...))
        ' must be array equation (ctrl-shift-enter)
    Set tmpRng = Range(WAE(1, 1).Offset(WAE.Rows.Count, 0), _
        WAE(1, 1).Offset(WAE.Rows.Count + 3, WAE.Columns.Count - 1))
    
    F1 = WAE.Columns(1).AddressLocal(RowAbsolute:=True, ColumnAbsolute:=False)
    For i = 1 To 4
        tmpRng(i, 1).Formula = "=QUARTILE.INC(" & F1 & ", " & 5 - i & ")"
    Next i
    
    tmpRng.Columns(1).Copy tmpRng
    
    ' Apply equation to assign quartile rank
    F1 = WAE(1, 1).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
    L1 = tmpRng.Columns(1).AddressLocal(RowAbsolute:=True, ColumnAbsolute:=False)
    Set AvgTR = Range(WAE(1, 1).Offset(0, WAE.Columns.Count + CB.Columns.Count), WAE(WAE.Rows.Count, WAE.Columns.Count).Offset(0, WAE.Columns.Count + CB.Columns.Count))
    
    AvgTR(1, 1).Formula = "=IFERROR(MATCH(" & F1 & ", " & L1 & ", -1), " & Chr(34) & Chr(34) & ")"
    
    Set Quart = Range(AvgTR.Columns(1), AvgTR.Columns(Periods))
    AvgTR(1, 1).Copy Quart
    Quart.Copy
    WAE.PasteSpecial Paste:=xlPasteValues
        
    ' Clean up
    Set WAE = Nothing
    Set tmpRng = Nothing
    Set AvgTR = Nothing
    Set CB = Nothing
    
    Range(Columns(Cells(1, 1).End(xlToRight).Column + 1), Columns(Columns.Count)).Clear
    Range(Rows(Cells(1, 1).End(xlDown).row + 1), Rows(Rows.Count)).Clear
    
    Range(Columns(Quart.Column + 3 * Periods), Columns(Columns.Count)).Clear
    Set Quart = Nothing

    CustomAggByQuartile Periods, custText

End Sub
Sub CustomAggByQuartile(ByRef Periods As Long, ByRef custText As String)

    Dim Quartiles(1 To 4)           As CMetaData
    Dim Area(2, 2), tmpRng          As Range
    Dim MySheet                     As Worksheet
    Dim CountryName, c              As Variant
    Dim i, j                        As Long
    Dim SumCB()                     As Double
    Dim Headers                     As Range
    
    
    ReDim SumCB(1 To Periods)
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(1, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    
    ' Clear extraneous
    With Range(Rows(Area(2, 1).row + 1), Rows(Rows.Count))
        .ClearContents
        .ClearFormats
        .UnMerge
        .AutoFit
    End With
    
    ' Instantiate CMetaData objects
    For j = 1 To 4
        Set Quartiles(j) = New CMetaData
    Next j

    ' Collect sums by TR quartile per period
    For j = 1 To 4  ' each quartile
    
        For i = 1 To Periods  ' each period
    
            With Range(Area(1, 1), Area(2, 2))
                .AutoFilter Field:=i + 1, Criteria1:=j
                
                ' Sum CB + Local net sales for quartile j in period i
                SumCB(i) = Application.WorksheetFunction.Subtotal(9, Columns(i + Periods + 1))
                
            End With
            
            MySheet.AutoFilterMode = False
            
        Next i
        
        ' Save CB & Local arrays in Quartiles
        Quartiles(j).Val1 = SumCB
        
    Next j
    
    ' Clear old data
    Range(Area(1, 1).Offset(1, 0), Area(2, 2)).ClearContents
    Range(Area(1, 1).Offset(0, Periods + 1), Area(1, 1).Offset(0, 2 * Periods)).Cut _
        Range(Area(1, 1).Offset(0, 1), Area(1, 1).Offset(0, 2 * Periods))
    
    ' Headers
    Set Headers = Rows(1)
    Headers.Select
    Call EditHeader
    Cells(1, 1).Value = "Quartile"
    With Rows(1)
        .ClearFormats
        .AutoFit
        .Font.Size = 8
    End With
    
    ' Insert new data from Quartiles
    Set tmpRng = Range(Cells(2, 1), Cells(5, 2 * Periods + 1))
    For j = 1 To 4
        tmpRng(j, 1).Value = j
        For i = 1 To Periods
            tmpRng(j, i + 1).Value = Quartiles(j).Val1(i)
        Next i
    Next j
    
    With tmpRng
        .Font.Size = 8
        .Rows.AutoFit
    End With
    Range(Columns(tmpRng(1, 1).Column), Columns(tmpRng.Columns.Count)).AutoFit
    
    ' Extra labels for clarity
    Set tmpRng = Range(Cells(1, 1), Cells(1, 1).End(xlToRight).End(xlDown))
    Range(Cells(2, 2), Cells(2, 2).Offset(3, 2 * Periods - 1)).NumberFormat = "#,##0.00"
    tmpRng.Columns.AutoFit
    tmpRng.Cut tmpRng.Offset(1, 0)
    Cells(1, 2).Value = custText & " Net Sales"
    With Rows(1)
        .Font.Size = 8
    End With
    
    Cells(1, 1).Select
    
End Sub


Sub CustomTRQuartChart(ByRef Periods As Long)

    Dim Area(2, 2), DataArea        As Range
    Dim ChartData                   As Range
    Dim MySheet                     As Worksheet
    Dim ChartLoc                    As Range
    Dim startCell                   As Range
    Dim MyChart                     As Chart
    Dim ChartWidth, ChartHeight     As Long
    Dim i, j                        As Long
    
    
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(2, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    Set DataArea = Range(Area(1, 1), Area(2, 2))
    
    ' Set Chart location params
    ChartWidth = 14
    ChartHeight = 28
    
    Set startCell = Area(2, 1).Offset(2, 0)
    Set ChartLoc = Range(startCell, startCell.Offset(ChartHeight - 1, ChartWidth - 1))

    
    ' Create chart
    Set ChartData = Range(DataArea(1, 1).Offset(0, 1), _
            DataArea(1, 1).Offset(4, Periods))
    Set MyChart = MySheet.Shapes.AddChart(xlColumnStacked).Chart
    
    With MyChart
    
        .SetSourceData Source:=ChartData, PlotBy:=xlRows
        
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
            .top = ChartLoc.top
            .left = ChartLoc.left
            .Height = ChartLoc.Height
            .Width = ChartLoc.Width
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
            .AxisTitle.top = 0
            .AxisTitle.left = 8
        End With
        
        With .Axes(xlCategory)
            .TickLabels.Font.Size = 12
            .TickLabelPosition = xlTickLabelPositionLow
            .TickLabels.Orientation = 30
        End With
        
        ' Legend
        With .Legend
            .Position = xlLegendPositionBottom
            .top = .top - 30
            .left = .left - 25
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
        .PlotArea.left = .ChartArea.left
        .PlotArea.Width = .ChartArea.Width - 10
        .PlotArea.top = 15
        .PlotArea.Height = .ChartArea.Height - 15
        
    End With
    
    Cells(1, 1).Select

End Sub


