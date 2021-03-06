Attribute VB_Name = "BuildReport"
' BuildReport containing the master procedures
' FormatData contains all data transformation procs
' CreateGraphs contains all charting procs

Function MatchSheetName(ThisName As String, ValidNames As Variant) As String

    ' If throwing errors, ensure that 'Tools/References/Microsoft VBScript
    ' Regular Expressions' is checked
    Dim Regex       As New VBScript_RegExp_55.RegExp
    Dim Matches
    Dim OneTrue     As Boolean
    
    For Each Valid In ValidNames
        Regex.Pattern = ".{2,3}" & Valid & "$"
        OneTrue = Regex.test(ThisName)
        MatchSheetName = Valid
        If OneTrue = True Then Exit For
    Next Valid
    
    If OneTrue = False Then
        MatchSheetName = ""
    End If
        
    
    
End Function


' Run this proc with Dash Report open (active worksheet insensitive)
Sub CreateReport()

    Dim FName       As String
    Dim FPath       As String
    Dim DateTime    As Date
    Dim mBefore     As Date
    Dim FileDate    As String
    Dim ReportDate  As String
    Dim MacroBook   As Workbook
    Dim DataBook    As Workbook
    Dim ValidSheets As Variant
    Dim iter        As Integer
    Dim BubbleLab(1 To 2) As String
    
    Application.ScreenUpdating = False  ' To boost performance
        
    Set MacroBook = Workbooks("CBSOReportMacros.xlsb")
    Workbooks.Open MacroBook.Path & "\" & MacroBook.Worksheets("Main").Cells(2, 3).Value
    Set DataBook = ActiveWorkbook
    
    
    ' Set Bubble chart standard labels (ex Global Bubbles)
    BubbleLab(1) = "Standard Deviation of Monthly Return in Euro (3 Year Weighted Average)"
    BubbleLab(2) = "Total Return in Euro 3-Year Weighted Average"
    
    ' Edit if changing reports (will require editing separate data
    ' transformation & charting procs in Modules 1 & 3)
    ValidSheets = Array("Global Net vs Gross Sales", "Global Gross Sales %", _
                        "Net Sales vs Avg. Performance", "Redemption Rate Calculation", _
                        "Morningstar Ratings", "Performance Global TopBottom", _
                        "Market Global TopBottom 5 M", "Market Global TopBottom 5 Y", "Market Share by Manager ", _
                        "Global Bubbles", "Local vs Cross-border net sal", _
                        "Manager Net Sales by Country", "Global Gross Sales % by Regi", _
                        "3 Equity Categories NetSales", "Global Bubbles Latest Qr", _
                        "Global Bubbles Prior Qr 3", "Global Bubbles Latest 12 Mth", _
                        "Manager Bubbles Latest 12 Mt", "Int'l&UK Bubbles Latest 12 M", _
                        "Manager Int'l&UK Latest 12 M", "Market Global Table TopBott", _
                        "Euro Net Flows By TR Quartil", "3Yr Euro TR Quartile", _
                        "Global Bubbles Int'l Net Flo", "Morningstar Europe", _
                        "Investment Type Gross Sales ", "Total Investment Type Gross ", _
                        "ETF vs Active YTD by categor", "Origination Markets Table", _
                        "Bubble Chart - Equity", "Bubble Chart - Bond", "1Yr Euro TR Quartile" _
                        )
                        
    ' Output file path & name (with dates)
    DateTime = Now
    mBefore = DateAdd("m", -1, Now)
    FileDate = Format(DateTime, "yymmdd")
    ReportDate = Format(mBefore, "mmmyyyy")
    
    FPath = ThisWorkbook.Path
    FName = FileDate & " " & "CBSOChartBook" & ReportDate
    
    ' Remove non-relevant sheets
    Application.DisplayAlerts = False
    For Each Item In DataBook.Sheets
        If MatchSheetName(Item.Name, ValidSheets) <> "" Then
            Item.Name = MatchSheetName(Item.Name, ValidSheets)
        Else
            DataBook.Worksheets(Item.Name).Delete
        End If
    Next Item
    Application.DisplayAlerts = True
    
    ' Global Net vs Gross
    DataBook.Worksheets("Global Net vs Gross Sales").Activate
    Call FormatData.GlobalNetGrossData
    Call CreateGraphs.GlobalNetGrossChart
    
    ' Global Gross Sales by Category
    DataBook.Worksheets("Global Gross Sales %").Activate
    Call FormatData.GlobalGrossCatData
    Call CreateGraphs.GlobalGrossCatChart
    
    ' Net Sales vs Avg. Performance
    DataBook.Worksheets("Net Sales vs Avg. Performance").Activate
    Call FormatData.NetAvgPerfData
    Call CreateGraphs.NetAvgPerfChart
    
    ' Redemption Rate Calculation
    DataBook.Worksheets("Redemption Rate Calculation").Activate
    Call FormatData.RedempCalcData
    Call CreateGraphs.RedempCalcChart
    
    ' MS Rating by Region
    DataBook.Worksheets("Morningstar Ratings").Activate
    Call FormatData.MSRegionData
    Call CreateGraphs.MSRegionChart
    
    ' Performance Global TopBottom Selling Cats
    DataBook.Worksheets("Performance Global TopBottom").Activate
    Call FormatData.PTopBottomData
    ' No chart to create, table only
    
    ' Market Global TopBottom Selling Cats (Month)
    DataBook.Worksheets("Market Global TopBottom 5 M").Activate
    Call FormatData.MTopBottomData
    Call CreateGraphs.MTopBottomChart
    
    ' Market Global TopBottom Selling Cats (Year)
    DataBook.Worksheets("Market Global TopBottom 5 Y").Activate
    Call FormatData.MTopBottomData
    Call CreateGraphs.MTopBottomChart
    
    ' Market Global TopBottom Table
    DataBook.Worksheets("Market Global Table TopBott").Activate
    Call FormatData.MTopBottomTableData
    Call CreateGraphs.MTopBottomTable
    
    ' Market Share by Manager
    DataBook.Worksheets("Market Share By Manager ").Activate
    Call FormatData.MktShareData
    Call CreateGraphs.MktShareChart
    
    ' Local vs Cross-border Net Sales
    DataBook.Worksheets("Local vs Cross-border net sal").Activate
    Call FormatData.LvCBData
    Call CreateGraphs.LvCBChart
    
    ' Manager Net Sales by Country
    DataBook.Worksheets("Manager Net Sales by Country").Activate
    Call FormatData.ManagerByCtryData
    Call CreateGraphs.ManagerByCtryChart
    
    ' Global Gross Sales % by Region
    DataBook.Worksheets("Global Gross Sales % by Regi").Activate
    Call FormatData.GrossByRegionData
    Call CreateGraphs.GrossByRegionChart
    
    ' 3 Equity Categories NetSales
    DataBook.Worksheets("3 Equity Categories NetSales").Activate
    Call FormatData.EquityCatSalesData
    Call CreateGraphs.EquityCatSalesChart
    
    ' Euro Net Flows By TR Quartile
    DataBook.Worksheets("Euro Net Flows By TR Quartil").Activate
    Call FormatData.EuroTRQuartileData
    Call CreateGraphs.EuroTRQuartileChart
    
    ' Euro Net Flows By 1-Yr TR Q
    DataBook.Worksheets("1Yr Euro TR Quartile").Activate
    Call FormatData.TrailTRData(1, 2)
    Call CreateGraphs.EuroTRQuartileChart
    
    ' Euro Net Flows By 3-Yr TR Q
    DataBook.Worksheets("3Yr Euro TR Quartile").Activate
    Call FormatData.TrailTRData(3, 5)
    Call CreateGraphs.EuroTRQuartileChart
        
    ' Manager Int'l & UK Latest 12 M
    DataBook.Worksheets("Manager Int'l&UK Latest 12 M").Activate
    Call FormatData.BubbleData(20)  ' Top 20
    Call CreateGraphs.ManagerBubbleChart("International & UK:" & vbCrLf & _
        "Top 20 Managers By Trailing 12-Month Net Flows")

    ' MS Europe CB v Local
    DataBook.Worksheets("Morningstar Europe").Activate
    Call FormatData.MSRegionData(1)
    Call CreateGraphs.MSRegionChart("cb v local")
    
    ' Investment Type by Country
    DataBook.Worksheets("Investment Type Gross Sales ").Activate
    Call FormatData.InvTypeGrossData
    Call CreateGraphs.InvTypeGrossChart
    
    ' Investment Type Total
    DataBook.Worksheets("Total Investment Type Gross ").Activate
    Call FormatData.InvTypeGrossData
    Call CreateGraphs.InvTypeGrossChart
    
    ' ETF v Active Bar
    DataBook.Worksheets("ETF vs Active YTD by categor").Activate
    Call FormatData.ActiveETFData
    Call CreateGraphs.ActiveETFChart
    
    ' Origination Markets Table
    DataBook.Worksheets("Origination Markets Table").Activate
    Call FormatData.MarketNetTblData
    Call CreateGraphs.MarketNetTblChart
    
    ' Bubble charts -------------------------------------------------------------------------------
    
    ' Global Bubbles
    DataBook.Worksheets("Global Bubbles").Activate
    Call FormatData.BubbleData  ' general bubble data proc
    Call CreateGraphs.BubbleChart("Asset-Weighted YTD and 1-Year Total Return vs 3-Month Trailing Net Flows", _
        "1-Year Return in Euro Weighted Average", _
        "YTD TR in Euro Weighted Average")
    
    ' Global Bubbles Latest Qr
    DataBook.Worksheets("Global Bubbles Latest Qr").Activate
    Call FormatData.BubbleData
    Call CreateGraphs.BubbleChart("Asset-Weighted 3-Year Total Return and Volatility vs. Latest Quarter Net Flows", _
        BubbleLab(1), BubbleLab(2))
    
    ' Global Bubbles Prior Year Qr
    DataBook.Worksheets("Global Bubbles Prior Qr 3").Activate
    Call FormatData.BubbleData  ' general bubble data proc
    Call CreateGraphs.BubbleChart("Asset-Weighted 3-Year Total Return and Volatility vs Net Flows This Quarter Last Year", _
        BubbleLab(1), BubbleLab(2))
    
    ' Global Bubbles Latest 12 Mth
    DataBook.Worksheets("Global Bubbles Latest 12 Mth").Activate
    Call FormatData.BubbleData
    Call CreateGraphs.BubbleChart("Asset-Weighted 3-Year Total Return and Volatility vs. 12 Month to Latest Month Net Flows", _
        BubbleLab(1), BubbleLab(2))
    
    ' Manager Bubbles Latest 12 Mth
    DataBook.Worksheets("Manager Bubbles Latest 12 Mt").Activate
    Call FormatData.BubbleData(20)  ' Top 20
    Call CreateGraphs.ManagerBubbleChart("Top 20 Manager By Trailing 12-Month Net Flows")
    
    ' Int'l & UK Bubbles Latest 12 Mth
    DataBook.Worksheets("Int'l&UK Bubbles Latest 12 M").Activate
    Call FormatData.BubbleData
    Call CreateGraphs.BubbleChart("International & UK:" & vbCrLf & _
        "Asset-Weighted 3-Year Total Return and Volatility vs. 12 Month to Latest Month Net Flows", _
        BubbleLab(1), BubbleLab(2))  ' vbCrLf = new line character
    ' Ensure title is in one merged cell
    With ActiveSheet
        .Range(Cells(1, Columns.Count).End(xlToLeft), Cells(1, Columns.Count).End(xlToLeft).Offset(0, 6)).Merge
    End With
        
    ' Int'l (GL) Bubbles
    DataBook.Worksheets("Global Bubbles Int'l Net Flo").Activate
    Call FormatData.BubbleData(10)
    Call CreateGraphs.BubbleChart("Asset-Weighted 3 Year International Total Return and Volatility vs YTD Net Sales", _
        BubbleLab(1), BubbleLab(2))
        
    ' Equity
    DataBook.Worksheets("Bubble Chart - Equity").Activate
    Call FormatData.BubbleData
    Call CreateGraphs.BubbleChart("Cross Border Performance", _
        "3 Yr TR % Standard Deviation", _
        "3 Yr TR %")
    
    ' Bond
    DataBook.Worksheets("Bubble Chart - Bond").Activate
    Call FormatData.BubbleData
    Call CreateGraphs.BubbleChart("Cross Border Performance", _
        "3 Yr TR % Standard Deviation", _
        "3 Yr TR %")
    

    ' Save file if it does not already exist
    If Dir(FPath & "\" & FName) <> "" Then
        MsgBox "File " & FPath & "\" & FName & " already exists"
    Else
        DataBook.SaveAs FileName:=FPath & "\" & FName & ".xlsm", _
            FileFormat:=52  ' to avoid compatibility problems, 2013 xlsm
    End If

End Sub
