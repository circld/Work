Attribute VB_Name = "Module2"
' Module2 containing the master procedures
' Module3 contains all data transformation procs
' Module1 contains all charting procs

Function MatchSheetName(ThisName As String, ValidNames As Variant) As String

    ' If throwing errors, ensure that 'Tools/References/Microsoft VBScript
    ' Regular Expressions' is checked
    Dim RegEx       As New VBScript_RegExp_55.RegExp
    Dim Matches
    Dim OneTrue     As Boolean
    
    For Each Valid In ValidNames
        RegEx.Pattern = ".{2,3}" & Valid & "$"
        OneTrue = RegEx.test(ThisName)
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
    Dim DataBook    As Workbook
    Dim ValidSheets As Variant
    Dim iter        As Integer
    Dim BubbleLab(1 To 2) As String
    
    Application.ScreenUpdating = False  ' To boost performance
        
    Set DataBook = ActiveWorkbook
    
    ' Set Bubble chart standard labels (ex Global Bubbles)
    BubbleLab(1) = "Standard Deviation of Monthly Return in Euro (3 Year Weighted Average)"
    BubbleLab(2) = "Total Return in Euro 3-Year Weighted Average"
    
    ' Edit if changing reports (will require editing separate data
    ' transformation & charting procs in Modules 1 & 3)
    ValidSheets = Array("Global Net vs Gross Sales", _
                        "Global Gross Sales %", _
                        "Net Sales vs Avg. Performance", _
                        "Redemption Rate Calculation", _
                        "Morningstar Ratings", _
                        "Performance Global TopBottom", _
                        "Market Global TopBottom 5 Se", _
                        "Market Share by Manager ", _
                        "Global Bubbles", _
                        "Local vs Cross-border net sa", _
                        "Manager Net Sales by Country", _
                        "Global Gross Sales % by Regi", _
                        "3 Equity Categories NetSales", _
                        "Global Bubbles Latest Qr", _
                        "Global Bubbles Prior Qr 3", _
                        "Global Bubbles Latest 12 Mth", _
                        "Manager Bubbles Latest 12 Mt", _
                        "Int'l&UK Bubbles Latest 12 M", _
                        "Manager Int'l&UK Latest 12 M", _
                        "Market Global Table TopBott", _
                        "Euro Net Flows By TR Quartil", _
                        "Global Bubbles Int'l Net Flo", _
                        "Morningstar Europe")
                        
    
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
    Call Module3.GlobalNetGrossData
    Call Module1.GlobalNetGrossChart
    
    ' Global Gross Sales %
    DataBook.Worksheets("Global Gross Sales %").Activate
    Call Module3.GlobalGrossPctData
    Call Module1.GlobalGrossPctChart
    
    ' Net Sales vs Avg. Performance
    DataBook.Worksheets("Net Sales vs Avg. Performance").Activate
    Call Module3.NetAvgPerfData
    Call Module1.NetAvgPerfChart
    
    ' Redemption Rate Calculation
    DataBook.Worksheets("Redemption Rate Calculation").Activate
    Call Module3.RedempCalcData
    Call Module1.RedempCalcChart
    
    ' MS Rating by Region
    DataBook.Worksheets("Morningstar Ratings").Activate
    Call Module3.MSRegionData
    Call Module1.MSRegionChart
    
    ' Performance Global TopBottom Selling Cats
    DataBook.Worksheets("Performance Global TopBottom").Activate
    Call Module3.PTopBottomData
    ' No chart to create, table only
    
    ' Market Global TopBottom Selling Cats
    DataBook.Worksheets("Market Global TopBottom 5 Se").Activate
    Call Module3.MTopBottomData
    Call Module1.MTopBottomChart
    
    ' Market Global TopBottom Table
    DataBook.Worksheets("Market Global Table TopBott").Activate
    Call Module3.MTopBottomTableData
    Call Module1.MTopBottomTable
    
    ' Market Share by Manager
    DataBook.Worksheets("Market Share By Manager ").Activate
    Call Module3.MktShareData
    Call Module1.MktShareChart
    
    ' Local vs Cross-border Net Sales
    DataBook.Worksheets("Local vs Cross-border net sa").Activate
    Call Module3.LvCBData
    Call Module1.LvCBChart
    
    ' Manager Net Sales by Country
    DataBook.Worksheets("Manager Net Sales by Country").Activate
    Call Module3.ManagerByCtryData
    Call Module1.ManagerByCtryChart
    
    ' Global Gross Sales % by Region
    DataBook.Worksheets("Global Gross Sales % by Regi").Activate
    Call Module3.GrossByRegionData
    Call Module1.GrossByRegionChart
    
    ' 3 Equity Categories NetSales
    DataBook.Worksheets("3 Equity Categories NetSales").Activate
    Call Module3.EquityCatSalesData
    Call Module1.EquityCatSalesChart
    
    ' Euro Net Flows By TR Quartile
    DataBook.Worksheets("Euro Net Flows By TR Quartil").Activate
    Call Module3.EuroTRQuartileData
    Call Module1.EuroTRQuartileChart
    
    ' Bubble charts
    
    ' Global Bubbles
    DataBook.Worksheets("Global Bubbles").Activate
    Call Module3.BubbleData  ' general bubble data proc
    Call Module1.BubbleChart("Asset-Weighted YTD and 1-Year Total Return vs 3-Month Trailing Net Flows", _
        "1-Year Return in Euro Weighted Average", _
        "YTD TR in Euro Weighted Average")
    
    ' Global Bubbles Latest Qr
    DataBook.Worksheets("Global Bubbles Latest Qr").Activate
    Call Module3.BubbleData
    Call Module1.BubbleChart("Asset-Weighted 3-Year Total Return and Volatility vs. Latest Quarter Net Flows", _
        BubbleLab(1), BubbleLab(2))
    
    ' Global Bubbles Prior Year Qr
    DataBook.Worksheets("Global Bubbles Prior Qr 3").Activate
    Call Module3.BubbleData  ' general bubble data proc
    Call Module1.BubbleChart("Asset-Weighted 3-Year Total Return and Volatility vs Net Flows This Quarter Last Year", _
        BubbleLab(1), BubbleLab(2))
    
    ' Global Bubbles Latest 12 Mth
    DataBook.Worksheets("Global Bubbles Latest 12 Mth").Activate
    Call Module3.BubbleData
    Call Module1.BubbleChart("Asset-Weighted 3-Year Total Return and Volatility vs. 12 Month to Latest Month Net Flows", _
        BubbleLab(1), BubbleLab(2))
    
    ' Manager Bubbles Latest 12 Mth
    DataBook.Worksheets("Manager Bubbles Latest 12 Mt").Activate
    Call Module3.BubbleData(20)  ' Top 20
    Call Module1.ManagerBubbleChart("Top 20 Manager By Trailing 12-Month Net Flows")
    
    ' Int'l & UK Bubbles Latest 12 Mth
    DataBook.Worksheets("Int'l&UK Bubbles Latest 12 M").Activate
    Call Module3.BubbleData
    Call Module1.BubbleChart("International & UK:" & vbCrLf & _
        "Asset-Weighted 3-Year Total Return and Volatility vs. 12 Month to Latest Month Net Flows", _
        BubbleLab(1), BubbleLab(2))  ' vbCrLf = new line character
    ' Ensure title is in one merged cell
    With ActiveSheet
        .Range(Cells(1, Columns.Count).End(xlToLeft), Cells(1, Columns.Count).End(xlToLeft).Offset(0, 6)).Merge
    End With
        
    ' Manager Int'l & UK Latest 12 M
    DataBook.Worksheets("Manager Int'l&UK Latest 12 M").Activate
    Call Module3.BubbleData(20)  ' Top 20
    Call Module1.ManagerBubbleChart("International & UK:" & vbCrLf & _
        "Top 20 Managers By Trailing 12-Month Net Flows")
    
    ' Int'l (GL) Bubbles
    DataBook.Worksheets("Global Bubbles Int'l Net Flo").Activate
    Call Module3.BubbleData(10)
    Call Module1.BubbleChart("Asset-Weighted 3 Year International Total Return and Volatility vs YTD Net Sales", _
        BubbleLab(1), BubbleLab(2))
        
    ' MS Europe CB v Local
    DataBook.Worksheets("Morningstar Europe").Activate
    Call Module3.MSRegionData(1)
    Call Module1.MSRegionChart("cb v local")

    ' Save file if it does not already exist
    If Dir(FPath & "\" & FName) <> "" Then
        MsgBox "File " & FPath & "\" & FName & " already exists"
    Else
        DataBook.SaveAs Filename:=FPath & "\" & FName & ".xlsm", _
            FileFormat:=52  ' to avoid compatibility problems, 2013 xlsm
    End If

End Sub
