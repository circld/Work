Attribute VB_Name = "FormatData"
' Requires (Tools/References)
'   Microsoft Scripting Runtime
'   Microsoft VBScript Regular Expressions

' Function for mapping country names to abbreviations
Function AbbrCtryName(ByVal NameRng As Range)

    Dim CountryNames    As New Scripting.Dictionary
    Dim country         As Variant

    ' Map country names to abbreviations
    CountryNames.Add "Italy", "ITA"
    CountryNames.Add "Switzerland", "SWI"
    CountryNames.Add "UK", "UK"
    CountryNames.Add "France", "FRA"
    CountryNames.Add "Hong Kong", "HK"
    CountryNames.Add "Germany", "GER"
    CountryNames.Add "Taiwan", "TWN"
    CountryNames.Add "Singapore", "SIN"
    CountryNames.Add "Spain", "SPA"
    CountryNames.Add "Luxembourg", "LUX"
    CountryNames.Add "Netherlands", "NETH"
    CountryNames.Add "Belgium", "BEL"
    CountryNames.Add "Benelux", "Benelux"
    
    For Each country In NameRng.Columns(1).Cells
        country.Value = CountryNames(country.Value)
    Next country

End Function

' Function for deleting qualifying rows
Function RemoveTrue(Area As Variant) As Long
    
    Dim MySheet As Worksheet
    Set MySheet = ActiveSheet
    
    With MySheet
        .Range(Area(1, 1), Area(2, 2).Offset(0, 1)).AutoFilter _
            Field:=Area(2, 2).Offset(0, 1).Column, _
            Criteria1:=1
        Application.DisplayAlerts = False
        .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Rows.Delete
        Application.DisplayAlerts = True
        .AutoFilterMode = False
    End With
    
End Function

' Rescale return percentage to decimal format
Sub Rescale(ByVal rng As Range)

    Dim MySheet As Worksheet
    Dim TmpNum  As Range
    
    Set MySheet = ActiveSheet
    Set TmpNum = Cells(1, Columns.Count)
    
    TmpNum.Value = 100
    TmpNum.Copy
    rng.PasteSpecial Operation:=xlPasteSpecialOperationDivide
    TmpNum.Value = 1
    TmpNum.Copy
    rng.PasteSpecial Operation:=xlPasteSpecialOperationAdd
    TmpNum.ClearContents
    
    ' To clear cells that were originally blank (now have value = 1)
    rng.Replace What:=1, Replacement:="", LookAt:=xlWhole

End Sub

' Module for data transformations

Sub EditHeader(Optional Prefix As String, Optional Suffix As String)

    Dim ActSheet    As Worksheet
    Dim HeaderRange As Range
    ' If throwing errors, ensure that Tools/References/Microsoft VBScript Regular Expressions is checked
    Dim Regex       As New VBScript_RegExp_55.RegExp
    Dim Matches
    
    Set ActSheet = ActiveSheet
    Set HeaderRange = Selection
    
    Regex.Pattern = ".*\[(.*)\].*"
    For Each Header In HeaderRange.Cells
        Set Matches = Regex.Execute(Header.Value)
        If Regex.Test(Header.Value) Then
            Header.Value = Join(Array(Prefix, Matches(0).SubMatches(0), Suffix), " ")
            Header.Value = LTrim(RTrim(Header.Value))
        End If
    Next Header
    
End Sub


' To sum columns, e.g. months to quarter Aggregate(NumCols:=3, func:="SUM")
Sub Aggregate(ByVal NumCols As Long, ByVal func As String)

    Dim DataRng                 As Range
    Dim Area(2, 2)              As Range
    Dim tmpRng                  As Range
    Dim step, col, reps, iter   As Long
    Dim startCell, endCell      As String
    Dim remainder               As Integer
    
    
    ' Define data area
    Set DataRng = Selection
    
    ' Set temporary data destination
    Set tmpRng = Range(Cells(DataRng.Rows(1).row, Columns.Count), _
                       Cells(DataRng.Rows(DataRng.Rows.Count).row, Columns.Count))
    
    ' Aggregate
    iter = 1
    remainder = DataRng.Columns.Count Mod NumCols
    reps = (DataRng.Columns.Count - remainder) / NumCols
    For step = 1 To reps * NumCols
        
        ' Aggregation function
        startCell = DataRng(1, step).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
        endCell = DataRng(1, step).Offset(0, NumCols - 1).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
        tmpRng(1, 1).Formula = "=" & func & "(" & startCell & ":" & endCell & ")"
        
        tmpRng(1, 1).Copy tmpRng
        tmpRng.Copy
        tmpRng.PasteSpecial xlPasteValues
        
        ' Clear columns just aggregated
        Range(DataRng.Columns(step), DataRng.Columns(step + NumCols - 1)).ClearContents
 
        ' Move data
        tmpRng.Copy DataRng.Columns(iter)
               
        step = step + NumCols - 1
        iter = iter + 1
        
    Next step

    ' Clear unaggregated data included in selection
    If remainder > 0 Then
        Range(DataRng.Columns(DataRng.Columns.Count + 1 - remainder), _
              DataRng.Columns(DataRng.Columns.Count)).ClearContents
    End If
    
    tmpRng.ClearContents
    DataRng.Select
    
End Sub


Sub GlobalNetGrossData()

    Dim Months      As Integer
    Dim MySheet     As Worksheet
    Dim LastRow, FirstRow, FirstCol, LastCol    As Long
    Dim Gross, Net, Assets, Headers, Final      As Range
    Dim Measures(1 To 3)                        As Range
    
    Set MySheet = ActiveSheet
    MySheet.Cells.UnMerge
    MySheet.Cells.ClearFormats
    MySheet.Activate
    ActiveWindow.FreezePanes = False
    Months = 13  ' Thirteen dates in series
    
    ' Define data area
    LastRow = Range("F" & Rows.Count).End(xlUp).row
    If Cells(LastRow, 1).Value = "" Then
        FirstCol = Cells(LastRow, 1).End(xlToRight).Column
    Else
        FirstCol = 3
    End If
    If Cells(1, 1).Value <> "" Then
        FirstRow = 2
    Else
        FirstRow = Cells(1, 1).End(xlDown).row + 1
    End If
    LastCol = FirstCol + Months - 1 ' Last Header column
    
    ' Edit HeaderNames
    Set Headers = Range(Cells(FirstRow - 1, FirstCol), _
        Cells(FirstRow - 1, FirstCol + Months - 1))
    Headers.Select
    Call EditHeader
    
    ' Copy chart data
    Range(Rows(LastRow + 1), Rows(Rows.Count)).Delete
    Set Net = Range(Cells(FirstRow, FirstCol), Cells(LastRow, LastCol))
    Set Gross = Range(Cells(FirstRow, FirstCol + Months), _
                    Cells(LastRow, FirstCol + 2 * Months - 1))
    Set Assets = Range(Cells(FirstRow, FirstCol + 2 * Months), _
                    Cells(LastRow, FirstCol + 3 * Months - 1))
    
    ' Define Measures range array
    Set Measures(1) = Net
    Set Measures(2) = Gross
    Set Measures(3) = Assets
    
    ' Define final output range
    Set Final = Range(Net.Columns(0).Rows(1), Net.Columns(Net.Columns.Count).Rows(4))
    
    For i = 1 To 3
        For col = 1 To Net.Columns.Count
            ' Measures(i).Columns(col).Select
            Final.Rows(i).Columns(col) = _
                Application.WorksheetFunction.Sum(Measures(i).Columns(col))
        Next col
    Next i
    
    Final.Rows(4) = _
        Application.WorksheetFunction.Average(Final.Rows(2))
        
    ' Swap Net & Gross (prior order necessary to sum correctly)
    Final.Rows(2).Cut
    Final.Rows(1).Insert Shift:=xlDown
    Set Final = Range(Final.Rows(0).Columns(1), _
        Final.Rows(Final.Rows.Count).Columns(Final.Columns.Count - 1))
    Final.Select
    
    ' Set row headers (nb. Final range changed by cut/insert operation)
    With Final
        .Cells(1, 0).Value = "Gross Sales"
        .Cells(2, 0).Value = "Net Sales"
        .Cells(3, 0).Value = "Total Assets (right axis)"
        .Cells(4, 0).Value = "Average Gross Sales"
    End With
    
    ' Move headers to correct position
    Cells(1, 1).Delete (xlToLeft)
    Cells(1, 1).Value = ""
    
    ' Clean up & format
    Range(Final.Rows(Final.Rows.Count + 1).Columns(0).Address, _
        Cells(Rows.Count, Columns.Count)).Delete
    Range(Cells(1, LastCol), Cells(Rows.Count, Columns.Count)).Delete
    Final.Cells.NumberFormat = "#,##0.00_);-#,##0.00"
    Final.Columns.AutoFit
    Columns(1).AutoFit
    
    Cells(1, 1).Select
    
End Sub

Sub GlobalGrossCatData()

    Dim Months      As Integer
    Dim MySheet     As Worksheet
    Dim FirstRow    As Long
    Dim LastRow     As Long
    Dim FirstCol    As Long
    Dim LastCol     As Long
    Dim Data1       As Range
    Dim Data2       As Range
    Dim Headers     As Range
    
    Set MySheet = ActiveSheet
    MySheet.Cells.UnMerge
    MySheet.Cells.ClearFormats
    MySheet.Activate
    ActiveWindow.FreezePanes = False
    Months = 13  ' Thirteen dates in series
    
    ' Define data area
    LastRow = Range("F" & Rows.Count).End(xlUp).row
    If Cells(LastRow, 1).Value = "" Then
        FirstCol = Cells(LastRow, 1).End(xlToRight).Column
    Else
        FirstCol = 2
    End If
    If Cells(1, 1).Value <> "" Then
        FirstRow = 2
    Else
        FirstRow = Cells(1, 1).End(xlDown).row + 1
    End If
    LastCol = FirstCol + Months - 1 ' Last Header column
    
    ' Edit HeaderNames
    Set Headers = Range(Cells(FirstRow - 1, FirstCol), _
        Cells(FirstRow - 1, FirstCol + Months - 1))
    Headers.Select
    Call EditHeader
    
    ' Copy chart data
    Set Data = Range(Rows(FirstRow + 3), Rows(FirstRow + 4))
    Data.Copy Range(Rows(FirstRow + 2), Rows(FirstRow + 3))
        
    ' Delete extraneous data & format
    Range(Rows(FirstRow + 4), Rows(Rows.Count)).Delete
    Range(Cells(FirstRow, FirstCol), Cells(FirstRow + 3, LastCol)).NumberFormat = "#,##0.00_);-#,##0.00"
    Range(Columns(1), Columns(Cells(1, 2).End(xlToRight).Column)).AutoFit
    
    Range("A1").Value = "€"
    Range("A1").Select
    
    
End Sub

Sub NetAvgPerfData()

    Dim Months      As Integer
    Dim MySheet     As Worksheet
    Dim FirstRow    As Long
    Dim LastRow     As Long
    Dim FirstCol    As Long
    Dim LastCol     As Long
    Dim Data         As Range
    Dim Headers     As Range
    
    Set MySheet = ActiveSheet
    MySheet.Cells.UnMerge
    MySheet.Cells.ClearFormats
    MySheet.Activate
    ActiveWindow.FreezePanes = False
    Months = 13  ' Thirteen dates in series
    
    ' Define data area
    LastRow = Range("F" & Rows.Count).End(xlUp).row
    If Cells(LastRow, 1).Value = "" Then
        FirstCol = Cells(LastRow, 1).End(xlToRight).Column
    Else
        FirstCol = 2
    End If
    If Cells(1, 1).Value <> "" Then
        FirstRow = 2
    Else
        FirstRow = Cells(1, 1).End(xlDown).row + 1
    End If
    LastCol = FirstCol + Months - 1 ' Last Header column
    
    ' Edit HeaderNames
    Set Headers = Range(Cells(FirstRow - 1, FirstCol), _
        Cells(FirstRow - 1, FirstCol + Months - 1))
    Union(Headers, Range(Cells(1, LastCol + 1), Cells(1, LastCol + Months))).Select
    Call EditHeader
    For Each c In Range(Cells(1, LastCol + 1), Cells(1, LastCol + Months).Cells)
        c.Value = c.Value & " ATR"
    Next c
    
    Range(Rows(LastRow + 1), Rows(Rows.Count)).Delete
        
    ' Delete extraneous data & format
    Range(Cells(FirstRow, FirstCol), Cells(FirstRow + 1, LastCol + Months)).Cells.NumberFormat = "#,##0.00_);-#,##0.00"
    Range(Columns(1), Columns(Cells(1, 2).End(xlToRight).Column)).AutoFit
    
    Range("A1").Value = "€"
    Range("A1").Select
    
    
End Sub

Sub RedempCalcData()

    Dim Months      As Integer
    Dim MySheet     As Worksheet
    Dim FirstRow, LastRow, FirstCol, LastCol   As Long
    Dim Redemp, Assets, Headers, Categories    As Range
    Dim i           As Long
    
    Set MySheet = ActiveSheet
    MySheet.Cells.UnMerge
    MySheet.Cells.ClearFormats
    MySheet.Activate
    ActiveWindow.FreezePanes = False
    Months = 13  ' Thirteen dates in series
    
    ' Define data area
    LastRow = Range("F" & Rows.Count).End(xlUp).row
    If Cells(LastRow, 1).Value = "" Then
        FirstCol = Cells(LastRow, 1).End(xlToRight).Column
    Else
        FirstCol = 2
    End If
    If Cells(1, 1).Value <> "" Then
        FirstRow = 2
    Else
        FirstRow = Cells(1, 1).End(xlDown).row + 1
    End If
    LastCol = FirstCol + Months - 1 ' Last Header column
    
    ' Edit HeaderNames
    Set Headers = Range(Cells(FirstRow - 1, FirstCol), _
        Cells(FirstRow - 1, FirstCol + Months - 1))
    Headers.Select
    Call EditHeader
    
    ' Copy chart data
    Range(Rows(LastRow + 1), Rows(Rows.Count)).Delete
    Set Categories = Range(Cells(FirstRow, 1), Cells(FirstRow, 1).End(xlDown))
    
    For i = 1 To 13
        Categories.Copy Range(Cells(i + 1, 1), Cells(i + 5, 1))
        Set Categories = Range(Cells(i + 1, 1), Cells(i + 5, 1))
        i = i + 4
    Next i

    ' Move assets to range below redemptions
    Set Assets = Range(Cells(FirstRow, LastCol + 1), _
        Cells(FirstRow + 2, 2 * LastCol - 1))
    Assets.Cut Range(Cells(FirstRow + 5, FirstCol), Cells(FirstRow + 7, LastCol))
    
    ' Label
    Range("A1").Value = "Redemptions"
    Range("A6").Value = "Assets"
    Range("A11").Value = "Redemptions rate"
    
    ' Add totals
    Cells(5, 2).Formula = "=SUM(R[-4]C[0]:R[-1]C[0])"
    For i = 1 To 2
        Cells(5, 2).Copy Range(Cells(5 * i, FirstCol), Cells(5 * i, LastCol))
    Next i
    
    ' Add Redemptions rate (Redemp / Assets)
    Range(Cells(12, FirstCol), Cells(15, LastCol)).Formula = "=R[-10]C[0] / R[-5]C[0]"
    
    ' Delete extraneous data & format
    Range(Cells(FirstRow - 1, LastCol + 1), Cells(Rows.Count, Columns.Count)).Delete
    Range(Cells(FirstRow, FirstCol), Cells(10, LastCol)).NumberFormat = "#,##0.00_);-#,##0.00"
    Range(Cells(12, FirstCol), Cells(15, LastCol)).NumberFormat = "0.00%;-0.00%"
    Range(Columns(1), Columns(Cells(1, 2).End(xlToRight).Column)).AutoFit
    
    Range("A1").Select
    
    
End Sub

Sub PTopBottomData()

    Dim MySheet     As Worksheet
    Dim FirstRow, LastRow, FirstCol, LastCol   As Long
    Dim top, Bottom, All   As Range
    Dim Final(1 To 2) As Range
    Dim total, Temp1, Temp2       As Long
    Dim Label(1 To 6) As String
    Dim Edge          As Variant
    
    Set MySheet = ActiveSheet
    MySheet.Cells.UnMerge
    MySheet.Cells.ClearFormats
    MySheet.Activate
    ActiveWindow.FreezePanes = False
    
    FirstRow = 2
    FirstCol = 2
    LastRow = Cells(1, 1).End(xlDown).row
    
    Range(Rows(LastRow + 1), Rows(Rows.Count)).Delete

    
    total = Application.WorksheetFunction.Sum(Range(Cells(FirstRow, FirstCol), Cells(LastRow, FirstCol)))
    
    ' Get rid of Net Sales
    Columns(FirstCol + 1).Delete
    
    LastCol = Cells(1, 1).End(xlToRight).Column
        
    ' Select Top 5 & Bottom 5
    Set top = Range(Cells(FirstRow, FirstCol - 1), Cells(FirstRow + 4, LastCol))
    Temp1 = top.Rows(top.Rows.Count).row
    Temp2 = top.Columns(top.Columns.Count).Column
    
    ' Set LastRow to last non-null Gross Sale value
    LastRow = Cells(FirstRow, FirstCol).End(xlDown).row
    Range(Rows(LastRow + 1), Rows(Rows.Count)).ClearContents
    Cells(FirstRow, FirstCol).Select
    Range(Rows(Temp1 + 1), Rows(Temp1 + 3)).ClearContents
    
    Range(Rows(Temp1 + 4), Rows(LastRow - 5)).Delete
    
    Set Bottom = Range(Cells(Temp1 + 4, FirstCol - 1), Cells(Temp1 + 8, Temp2))
    
    ' Format
    top.Cells.NumberFormat = "#,##0.00_);-#,##0.00"
    Bottom.Cells.NumberFormat = "#,##0.00_);-#,##0.00"
    
    ' Headers & formatting
    Label(1) = "Selling Categories"
    Label(2) = "Gross Sales"
    Label(3) = "Total Return"
    Label(4) = "€ Million"
    Label(5) = "12-mnth % in €"
    Label(6) = "All Categories"
    
    Rows(2).Insert Shift:=xlDown
    Set Final(1) = top
    Set Final(2) = Bottom
    
    For Each rng In Final
        With rng.Rows(-1)
            For i = 1 To 3
                .Columns(i) = Label(i)
            Next i
            .HorizontalAlignment = xlCenter
            With rng.Rows(-1).Font
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
            With rng.Rows(-1).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
                .PatternTintAndShade = 0
            End With
        End With
        
        With rng.Rows(0)
            For j = 2 To 3
                .Columns(j) = Label(j + 2)
            Next j
            .HorizontalAlignment = xlCenter
            .Font.Italic = True
        End With
        
        For c = 2 To 3
            With rng.Columns(c)
                If c = 2 Then
                    .Font.Bold = True
                End If
                .HorizontalAlignment = xlCenter
            End With
        Next c
    Next rng
    
    top.Rows(-1).Columns(1).Value = "Top " & top.Rows(-1).Columns(1).Value
    Bottom.Rows(-1).Columns(1).Value = "Bottom " & Bottom.Rows(-1).Columns(1).Value
    
    Range(Rows(1), Rows(Rows.Count)).AutoFit
    Range(Columns(1), Columns(Columns.Count)).AutoFit
    
    LastRow = Bottom.Rows(Bottom.Rows.Count + 2).row
    
    Set All = Range(Cells(LastRow, 1), Cells(LastRow, LastCol))
    All(1, 1).Value = Label(6)
    With All(1, 2)
        .Value = total
        .HorizontalAlignment = xlCenter
        .NumberFormat = "#,##0.00_);-#,##0.00"
    End With
    
    With All
        .Font.Bold = True
        .Font.Italic = True
        For Each Edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
            With All.Borders(Edge)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlThin
            End With
        Next Edge
    End With
    
    With Range(Cells(1, 1), Cells(LastRow, LastCol))
        .Borders(xlInsideVertical).LineStyle = xlNone
    End With
    With Range(Cells(1, 1), Cells(LastRow - 1, LastCol))
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    Cells(1, 1).Select
    
End Sub

Sub MTopBottomData()
    ' 1. find ranges for each country
    ' 2. Total net, local net, CB net
    ' 3. Order by Total net (w/in country)
    Dim Countries()                 As CMetaData
    Dim Area(2, 2)                  As Range
    Dim MySheet                     As Worksheet
    Dim CountryName                 As Variant
    Dim i, j, Count                 As Long
    Dim NumCtry                     As Long
    
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(1, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    
    ' Clear unnecessary rows & columns
    Range(Rows(Area(2, 1).row + 1), Rows(Rows.Count)).Delete
    Columns(3).Delete
    Columns(4).Delete
    
    ' Get rid of all null/zero rows
    Range(Cells(2, Area(1, 2).Column + 1), Cells(Area(2, 2).row, Area(2, 2).Column + 1)).Formula = _
        "=IF(SUM(RC[-3]:RC[-1]) = 0, 1, 0)"
    Call RemoveTrue(Area)
    
    Columns(Area(1, 2).Column + 1).Delete
    
    Set Area(1, 1) = Cells(1, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    
    ' Get country count & instantiate CMetaData items with names
    With Range(Area(1, 1), Area(2, 1))
        .AdvancedFilter Action:=xlFilterInPlace, Unique:=True
        NumCtry = .SpecialCells(xlCellTypeVisible).Cells.Count - 1
        ReDim Countries(1 To NumCtry)
        i = 1
        For Each CountryName In Range(Area(1, 1).Offset(1, 0), Area(2, 1)).SpecialCells(xlCellTypeVisible).Cells
            Set Countries(i) = New CMetaData
            Countries(i).Name = CountryName.Value
            Countries(i).FirstRow = CountryName.row
            i = i + 1
        Next CountryName
        MySheet.ShowAllData
    End With
    
    ' Cycle through markets & set range criteria
    For i = 1 To NumCtry
        
        MySheet.Range(Area(1, 1), Area(2, 2)).AutoFilter Field:=1, Criteria1:=Countries(i).Name
        
        ' Sort on Total Net Sales w/in country
        MySheet.AutoFilter.Sort.SortFields.Clear
        MySheet.AutoFilter.Sort.SortFields.Add Key:=Columns(3), SortOn:=xlSortOnValues, _
            Order:=xlDescending, DataOption:=xlSortNormal
            
        With MySheet.AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        ' Debugging Taiwan
        
        Count = MySheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
        Countries(i).LastRow = MySheet.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible)(Count, 1).row
        Countries(i).FirstCol = Area(1, 1).Column
        Countries(i).LastCol = Area(1, 2).Column
        
        ' Number formatting
        Range(Countries(i).DataRange.Columns(3), Countries(i).DataRange.Columns(5)).NumberFormat = _
            "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        
    Next i
    
    MySheet.AutoFilterMode = False
    
    ' Move into final locations & format/clear
    j = NumCtry - 1
    For i = 0 To NumCtry - 1
        Countries(NumCtry - i).DataRange.Cut Countries(NumCtry - i).DataRange.Offset(j, 0)
        j = j - 1
    Next i
    
    ' Headers
    Range(Columns(Area(1, 1).Column), Columns(Area(1, 2).Column)).AutoFit
    With Rows(1)
        .Cells.ClearFormats
        .Cells.Font.Size = 8
        .Cells.HorizontalAlignment = xlCenter
        .AutoFit
    End With
    
    Rows(1).Insert
    With Range(Cells(1, 3), Cells(1, 5))
        .Merge
        .Value = "Latest Month"
        .HorizontalAlignment = xlCenter
    End With
    Cells(2, 1).Value = "Market"
    Cells(2, 3).Value = "Total Net Sales"
    Cells(2, 4).Value = "Local Net Sales"
    Cells(2, 5).Value = "Cross-Border Net Sales"
    Cells(1, 1).Value = "In EUR"
    
    ' Add blank row before first country block
    Rows(Countries(1).FirstRow + 1).Insert  ' + 1 to account for added row
    
    Cells(1, 1).Select
    
End Sub

Sub MktShareData()
    ' Top 5 managers + cumulative (delete all sum(vals) = 0 rows)
    ' Only need 3-Mth & Sum(prior year months)
    ' Separate Bond & Equity data by empty row
    
    Dim Area(2, 2), tmpRow          As Range
    Dim MySheet                     As Worksheet
    Dim i, j, Count                 As Long
    Dim FType, Item                 As Variant
    Dim HLabels(1 To 4)             As String
    Dim AnonNames(1 To 6)           As String
    
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(1, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    HLabels(1) = "3 Mth"
    HLabels(2) = "Prior 12 Mth"
    HLabels(3) = "Prior 13 Mth"
    HLabels(4) = "Prior 14 Mth"
    
    ' Clear unnecessary rows & columns
    Range(Rows(Area(2, 1).row + 1), Rows(Rows.Count)).Delete
    
    ' Get rid of all null/zero rows
    Range(Cells(2, Area(1, 2).Column + 1), Cells(Area(2, 2).row, Area(2, 2).Column + 1)).Formula = _
        "=IF(SUM(RC[-3]:RC[-1]) = 0, 1, 0)"
    Call RemoveTrue(Area)
    
    ' Redefine Area params since deleting rows changed data range
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    
    Columns(Area(2, 2).Column + 1).Delete
    
    ' Create Prior Yr 3 Month column nb: no need to redefine Area last col
    Area(1, 2).Offset(0, 1).Value = "Sum of Prior Yr 3 Mth"
    Range(Area(1, 2).Offset(1, 1), Area(2, 2).Offset(0, 1)).Formula = _
        "=SUM(RC[-3]:RC[-1])"
    
    ' Add Other Manager rows to Bond & Equity
    For Each FType In Array("Bond", "Equity")
    
        Set Area(2, 1) = Area(1, 1).End(xlDown)
        Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
        
        ' Define anonymization by fund type
        If FType = "Bond" Then
            For i = 1 To 5
                AnonNames(i) = "Manager " & i
            Next i
        Else
            i = 1
            For Each Item In Array("A", "B", "C", "D", "E")
                AnonNames(i) = "Manager " & Item
                i = i + 1
            Next Item
        End If
        AnonNames(6) = "Other Managers"
            
        With MySheet
            .Range(Area(1, 1), Area(2, 2).Offset(0, 1)).AutoFilter _
                Field:=Area(1, 1).Column, _
                Criteria1:=FType
            
            Set tmpRow = Rows( _
                .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Rows(6).row _
                )
            tmpRow.Copy
            tmpRow.Insert Shift:=xlDown
            tmpRow.Offset(-1, 0).Cells(1, 2).Value = "Other Managers"
            Range(Cells(tmpRow.Offset(-1, 0).row, 3), _
                Cells(tmpRow.Offset(-1, 0).row, Area(1, 2).Column)).Formula = _
                "=SUM(R[1]C:R[" & .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Rows.Count - 6 & "]C)"
            
            Set tmpRow = Rows(.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Rows(1).row)
            If FType = "Equity" Then
                tmpRow.Insert Shift:=xlDown
            End If
            
            ' Anon Names
            For i = 1 To 6
                Cells(tmpRow.row - 1 + i, Area(1, 2).Offset(0, 3).Column).Value = _
                    AnonNames(i)
            Next i
            
            .AutoFilterMode = False
        End With
    Next FType
    
    ' Formatting
    Range(Area(1, 2), Area(2, 2)).Copy
    Range(Area(1, 2).Offset(0, 1), Area(2, 2).Offset(0, 1)).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    i = 0
    For i = 1 To 4
        Area(1, 1).Offset(0, 1 + i).Value = HLabels(i)
    Next i
    
    Range(Columns(Area(1, 1).Column), Columns(Area(2, 2).Column + 1)).AutoFit
    Rows(1).AutoFit
    
    Cells(1, 1).Select
    
End Sub


Sub BubbleData(Optional TopN As Long = 10)

    Dim Area(2, 2), tmpCell         As Range
    Dim MySheet                     As Worksheet
    Dim i, j, Count                 As Long
    Dim BlankFormula                As String
    
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(1, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    BlankFormula = "=AND("
    
    ' Clear extraneous Dash output
    Range(Rows(Area(2, 1).row + 1), Rows(Rows.Count)).Delete
    
    ' Test for blanks (generalized over cols != col(1)
    Count = Range(Area(1, 1).Offset(0, 1), Area(1, 2)).Columns.Count
    Area(1, 2).Offset(0, 1).Value = "No Nulls"
    For i = 1 To Count
        BlankFormula = BlankFormula & "ISBLANK(RC[-" & i & "]) = FALSE"
        If i <> Count Then
            BlankFormula = BlankFormula & ", "
        Else
            BlankFormula = BlankFormula & ")"
        End If
    Next i
    
    Range(Area(1, 2).Offset(1, 1), Area(2, 2).Offset(0, 1)).Formula = _
        BlankFormula
        
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    
    ' Remove any rows with blanks
    With MySheet
        .Range(Area(1, 1), Area(2, 2)).AutoFilter _
            Field:=Area(1, 2).Column, Criteria1:=False
        Application.DisplayAlerts = False
        .AutoFilter.Range.Offset(1, 0).Rows.Delete
        .AutoFilterMode = False
        Application.DisplayAlerts = True
    End With
    
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    
    ' Sort Measure Column
    Range(Area(1, 1), Area(2, 2)).Sort _
        Key1:=Range(Area(1, 1), Area(2, 2)).Columns(2), _
        Order1:=xlDescending, _
        Header:=xlYes
    
    ' Remove CB Net Sales records outside of top N
    Area(2, 2).Select
    Range(Area(1, 1).Offset(TopN + 1, 0), Area(2, 2)).Delete
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    
    ' Convert to WAE columns to pct
    Set tmpCell = Area(1, 2).Offset(0, 1)
    tmpCell.Value = 100
    tmpCell.Copy
    Range(Area(1, 1).Offset(1, 2), Area(2, 2).Offset(0, -1)).PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationDivide
    Selection.NumberFormat = "0.0%"
    tmpCell.ClearContents
        
    Range(Area(1, 2), Area(2, 2)).Copy
    Range(Area(1, 2).Offset(0, 1), Area(2, 2).Offset(0, 1)).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    ' Standardize formatting
    With Rows(1)
        .ClearFormats
        .AutoFit
        .Font.Size = 8
    End With
    
    Cells(1, 1).Select
    
End Sub

Sub LvCBData()

    Dim DataArea                    As CMetaData
    Dim Area(2, 2), tmpCell         As Range
    Dim MySheet                     As Worksheet
    Dim Count(1 To 2)               As Double
    Dim Item                        As Variant

    
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(1, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    
    Area(2, 1).Offset(1, 0).Value = "Benelux"
    Rows(Area(2, 1).Offset(2, 0).row).Insert
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    
    With Range(Rows(Area(2, 1).row + 1), Rows(Rows.Count))
        .ClearContents
        .ClearFormats
        .UnMerge
        .AutoFit
    End With
    
    Set DataArea = New CMetaData
    DataArea.FirstRow = Area(1, 1).Offset(1, 0).row  ' ignore header row
    DataArea.LastRow = Area(2, 1).row
    DataArea.FirstCol = Area(1, 1).Column
    DataArea.LastCol = Area(1, 2).Column

    Count(1) = 0
    Count(2) = 0
    For Each Item In DataArea.DataRange.Rows

        If Item.Cells(1, 1).Value = "Belgium" _
            Or Item.Cells(1, 1).Value = "Luxembourg" _
            Or Item.Cells(1, 1).Value = "Netherlands" _
            Then
            Count(1) = Count(1) + Item.Cells(1, 2).Value
            Count(2) = Count(2) + Item.Cells(1, 3).Value
            Item.ClearContents
        ElseIf Item.Cells(1, 1).Value = "Benelux" Then
            Item.Cells(1, 2) = Count(1)
            Item.Cells(1, 3) = Count(2)
        End If
        
    Next Item
    
    DataArea.DataRange.Sort Key1:=DataArea.DataRange.Columns(1), Header:=xlNo
    
    ' Sort by sum (descending)
    Count(1) = DataArea.DataRange.Columns.Count + 1
    Count(2) = DataArea.DataRange(1, 1).End(xlDown).row
    DataArea.LastCol = Count(1)
    DataArea.LastRow = Count(2)
    DataArea.DataRange.Columns(Count(1)).Formula = "=SUM(RC[-2]:RC[-1])"
    DataArea.DataRange.Sort Key1:=DataArea.DataRange.Columns(Count(1)), Header:=xlNo, Order1:=xlDescending
    DataArea.DataRange.Columns(Count(1)).ClearContents
    
    ' Abbreviate country names
    AbbrCtryName DataArea.DataRange.Columns(1)
    
    Cells(1, 1).Select

End Sub

Sub ManagerByCtryData()

    Dim Countries()                 As CMetaData
    Dim Area(2, 2), tmpCell, Filt   As Range
    Dim MySheet                     As Worksheet
    Dim Count                       As Long
    Dim Item                        As Variant
    Dim CtryNames, Vals             As Range
    Dim NumCtry, i, j, k            As Long
    
    Set MySheet = ActiveSheet
    Set Area(1, 1) = Cells(1, 1)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    
    With Range(Rows(Area(2, 1).row + 1), Rows(Rows.Count))
        .ClearContents
        .ClearFormats
        .UnMerge
        .AutoFit
    End With
    
    ' Grab unique country names & instantiate Countries metadata
    Set CtryNames = Range(Area(1, 2).Offset(1, 2), Area(2, 2).Offset(0, 2))
    Range(Area(1, 1).Offset(1, 0), Area(2, 1)).Copy CtryNames
    CtryNames.CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlNo
    Set CtryNames = Range(CtryNames.Cells(1, 1), CtryNames.Cells(1, 1).End(xlDown))
    NumCtry = CtryNames.Rows.Count
    ReDim Countries(1 To NumCtry)
    
    For i = 1 To NumCtry
        Set Countries(i) = New CMetaData
        Countries(i).Name = CtryNames(i, 1).Value
    Next i
    
    CtryNames.ClearContents
    Set Vals = Range(CtryNames(1, 1), CtryNames(3, 1).Offset(0, 1))

    With MySheet
    
        For i = 1 To NumCtry
            .Range(Area(1, 1), Area(2, 2)).AutoFilter _
                Field:=Area(1, 1).Column, _
                Criteria1:=Countries(i).Name
            
            Set Filt = Range(.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Address)
            Count = Filt.Range(Filt(1, 3), Filt(1, 3).End(xlDown)).Rows.Count

            ' Sum by country for top 3, middle, & bottom 3 managers for each measure
            For j = 1 To 2
                
                Vals.Cells(1, j).Formula = "=SUM(" & Range(Filt(1, j + 2), Filt(3, j + 2)).Address & ")"
                Vals.Cells(2, j).Formula = "=SUM(" & Range(Filt(4, j + 2), Filt(Count - 3, j + 2)).Address & ")"
                Vals.Cells(3, j).Formula = "=SUM(" & Range(Filt(Count - 2, j + 2), Filt(Count, j + 2)).Address & ")"
            
            Next j
            
            Countries(i).Val1 = Array(Vals(1, 1).Value, Vals(1, 2).Value)
            Countries(i).Val2 = Array(Vals(2, 1).Value, Vals(2, 2).Value)
            Countries(i).Val3 = Array(Vals(3, 1).Value, Vals(3, 2).Value)
            
            Vals.ClearContents
            
        Next i
        
        .AutoFilterMode = False
        
    End With
    
    ' Clear existing data in DataArea
    Range(Area(1, 1), Area(2, 2)).ClearContents
    
    ' Input data stored in Countries
    ' Headers
    Rows(1).Insert
    Rows(2).AutoFit
    Cells(2, 1).Value = "Origination Market"
    For i = 0 To 1
        j = 3 * i  ' column shift
        Cells(2, 2 + j).Value = "Top 3 Managers"
        Cells(2, 3 + j).Value = "Managers In Between"
        Cells(2, 4 + j).Value = "Bottom 3 Managers"
        Range(Cells(1, 2 + j), Cells(1, 4 + j)).Merge
        With Cells(1, 2 + j)
            .Value = "Latest Month by Cross-Border Net Sales"
            .HorizontalAlignment = xlCenter
        End With
    Next i
    For i = 1 To Cells(2, 1).End(xlToRight).Column
        Columns(i).AutoFit
    Next i
    
    ' Paste data
    For i = 1 To NumCtry
        j = 2 + i
        Cells(j, 1).Value = Countries(i).Name
        For k = 0 To 1
            Cells(j, 2 + 3 * k).Value = Countries(i).Val1(k)
            Cells(j, 3 + 3 * k).Value = Countries(i).Val2(k)
            Cells(j, 4 + 3 * k).Value = Countries(i).Val3(k)
        Next k
        Rows(i + 2).ClearFormats
        Rows(i + 2).Font.Size = 8
    Next i
    
    With Rows(2)
        .ClearFormats
        .Font.Size = 8
    End With
    Rows(1).Font.Size = 8
    
    ' Number format
    Range(Cells(3, 2), Cells(3, 2).End(xlDown).End(xlToRight)).NumberFormat = "#,##0.00"
    
    ' Abbreviate country names
    AbbrCtryName Range(Cells(3, 1), Cells(3, 1).End(xlDown))
    
    Cells(1, 1).Select
    
End Sub

Sub GrossByRegionData()

    Dim Area(2, 2)                  As Range
    Dim MySheet                     As Worksheet
    Dim Count                       As Long
    Dim Item                        As Variant
    Dim World                       As Range
    Dim Months                      As Long

    
    Months = 13
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

    ' Edit HeaderNames
    Set Headers = Range(Area(1, 1), Area(1, 2))
    Headers.Select
    Call EditHeader
    
    ' Separate "Rest of world" regions
    Rows(4).Insert
    Rows(4).Insert
    Cells(4, 1).Value = "Rest of World"
    Set World = Range(Cells(4, 2), Cells(4, Months + 1))
    World.Formula = "= SUM(R[2]C:R[5]C)"
    
    With Rows(1)
        .AutoFit
        .ClearFormats
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    Range(Columns(1), Columns(Area(2, 2).Column)).AutoFit
    
    Cells(1, 1).Value = "€"
    Cells(1, 1).Select
    
End Sub

Sub EquityCatSalesData()

    Dim Area(2, 2)                  As Range
    Dim MySheet                     As Worksheet
    Dim Count                       As Long
    Dim Item                        As Variant
    Dim World                       As Range
    Dim Months                      As Long


    Months = 13
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

    ' Edit HeaderNames
    Set Headers = Range(Area(1, 1), Area(1, 2))
    Headers.Select
    Call EditHeader
    
    With Range(Rows(Area(1, 1).row), Rows(Area(2, 2).row))
        .ClearFormats
        .AutoFit
        .Font.Size = 8
    End With
    Range(Columns(Area(1, 1).Column), Columns(Area(2, 2).Column)).AutoFit
    
    Rows(1).Insert
    Cells(1, 2).Value = "Cross Border Net Sales EUR"
    Cells(1, 2).Offset(0, Months).Value = "Avg Total Return"
    Rows(1).Font.Size = 8
    
    Cells(1, 1).Select

End Sub

Sub MSRegionData(Optional NumRegions As Long = 2)

    Dim Rating(2, 5)                As CMetaData
    Dim MySheet                     As Worksheet
    Dim Area(2, 2), Headers, tmp    As Range
    Dim Months                      As Long
    Dim region, mth, sales, stars   As Long
    Dim RegFilt(1 To 2), StarLabel  As String
    Dim tmpVals()                   As Double
    Dim DataRange, tmpRng           As Range
    Dim i, j, total                 As Long
    
    
    Months = 13
    RegFilt(1) = "Asia Pacific"
    RegFilt(2) = "Europe"
    ReDim tmpVals(1 To Months)
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
    
    ' Edit Headers
    Set Headers = Range(Area(1, 1), Area(1, 2))
    Headers.Select
    Call EditHeader
    
    ' Cleaning (clear formats + autofit)
    With Range(Rows(Area(1, 1).row), Rows(Area(2, 2).row))
        .ClearFormats
        .AutoFit
        .Font.Size = 8
    End With
    Range(Columns(Area(1, 1).Column), Columns(Area(2, 2).Column)).AutoFit
    
    ' Instantiate Star array
    For region = 1 To 2
        For stars = 1 To 5
            Set Rating(region, stars) = New CMetaData
        Next stars
    Next region
    
    ' Grab totals by region, rating, & type
    Set tmp = Area(1, 2).Offset(0, 1)  ' temp cell to store sum until assignment
    For region = 1 To 2
        
        For stars = 1 To 5
        
            For sales = 1 To 2
        
                For mth = 1 To Months
                    
                    Application.DisplayAlerts = False
                    
                    With MySheet.Range(Area(1, 1), Area(2, 2))
                        .AutoFilter Field:=Area(1, 1).Column, Criteria1:=RegFilt(region)
                        .AutoFilter Field:=Area(1, 1).Offset(0, mth + 1).Column, _
                            Criteria1:=stars
                    
                        Rating(region, stars).Name = RegFilt(region) & " " & stars & "-star"
                        
                        tmp.Formula = "=SUBTOTAL(9, " & Columns(2 + mth + sales * Months).Address & ")"
                        tmpVals(mth) = tmp.Value
                        
                        MySheet.AutoFilterMode = False
                        
                    End With
        
                    Application.DisplayAlerts = True
                        
                Next mth
                
                If sales = 1 Then
                    Rating(region, stars).Val1 = tmpVals  ' Gross
                Else
                    Rating(region, stars).Val2 = tmpVals  ' Net
                End If
            
            Next sales
        
        Next stars
    
    Next region
    
    tmp.ClearContents
    
    ' Clear existing data & input new data
    Range(Area(1, 1).Offset(1, 0), Area(2, 2)).ClearContents
    Area(1, 1).Offset(0, 1).Value = "Rating"
    Set DataRange = Range(Cells(1, 1).Offset(1, 0), _
        Cells(1, 1).Offset(10, 1 + 2 * Months))
    
    ' Add star labels
    For i = 1 To 10
        
        ' assign stars, region
        If i Mod 5 = 0 Then
            stars = 5
            region = i \ 5
        Else
            stars = i Mod 5
            region = i \ 5 + 1
        End If
        
        ' Populating DataRange w/data
        If stars = 1 Then DataRange(i, 1).Value = RegFilt(region)
        
        StarLabel = ""
        For j = 1 To stars
            StarLabel = StarLabel & "«"
        Next j
        
        DataRange(i, 2).Value = StarLabel
        DataRange(i, 2).Font.Name = "Wingdings"
        
        For mth = 1 To 13
            
            DataRange(i, 2 + mth).Value = Rating(region, stars).Val1(mth)
            DataRange(i, 2 + mth + Months).Value = Rating(region, stars).Val2(mth)
            
        Next mth
        
    Next i
    
    Range(Columns(DataRange.Columns.Count + 1), Columns(Columns.Count)).Delete
    DataRange.Columns.AutoFit
    DataRange.Rows.AutoFit
    DataRange.Offset(0, 2).Select
    DataRange.Offset(0, 2).NumberFormat = "#,##0.00"
    
    ' Add labels
    Rows(1).Insert
    If NumRegions = 2 Then
        Cells(1, 3).Value = "Gross Sales"
        Cells(1, 3 + Months).Value = "Net Sales"
    Else
        Range(Rows(3), Rows(7)).Delete (xlUp)
        Cells(1, 3).Value = "Dummy"
        Cells(1, 3 + Months).Value = "Dummy"
        Area(1, 1).Offset(1, 0).Value = "CrossBorder Net Sales"
        Area(1, 1).Offset(1 + 5).Value = "Local Net Sales"
        
        ' Move local vals down
        Set tmpRng = Range(Area(1, 1).Offset(1, 2 + Months), Area(1, 1).Offset(5, 2 + 2 * Months - 1))
        tmpRng.Cut tmpRng.Offset(5, 0)
        ' Move cb vals right
        Set tmpRng = Range(Area(1, 1).Offset(1, 2), Area(1, 1).Offset(5, 1 + Months))
        tmpRng.Cut tmpRng.Offset(0, Months)
        ' Move stars down
        Set tmpRng = Range(Area(1, 1).Offset(1, 1), Area(1, 1).Offset(5, 1))
        tmpRng.Copy tmpRng.Offset(5, 0)
    End If
    Rows(1).Font.Size = 8
    
    Cells(1, 1).Select
    
End Sub

Sub MTopBottomTableData()
    ' Consider adding regex to detect which columns to operate on (ie assign to ColCB, ColLoc)
    
    Dim Countries()                 As CMetaData
    Dim Area(2, 2), tmpRng          As Range
    Dim MySheet                     As Worksheet
    Dim CountryName, c              As Variant
    Dim i, j, NumCtry               As Long
    Dim tmpVals                     As Variant
    Dim tmpSum                      As Double
    Dim ColCB, ColLoc               As Long
    
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

    ' Get country count & instantiate CMetaData items with names
    With Range(Area(1, 1), Area(2, 1))
        .AdvancedFilter Action:=xlFilterInPlace, Unique:=True
        NumCtry = .SpecialCells(xlCellTypeVisible).Cells.Count - 1
        ReDim Countries(1 To NumCtry)
        i = 1
        For Each CountryName In Range(Area(1, 1).Offset(1, 0), Area(2, 1)).SpecialCells(xlCellTypeVisible).Cells
            Set Countries(i) = New CMetaData
            Countries(i).Name = CountryName.Value
            Countries(i).FirstRow = CountryName.row
            i = i + 1
        Next CountryName
        MySheet.ShowAllData
    End With
    
    ' Retrieve country-specific top & bottom CB & local managers net sales & share
    ColCB = 3
    ColLoc = 6
    
    For j = 1 To NumCtry
        With Range(Area(1, 1), Area(2, 2))
            
            .AutoFilter Field:=1, Criteria1:=Countries(j).Name
        
            ' Top CB Manager
            .AutoFilter Field:=ColCB, Criteria1:=">0"
            MySheet.AutoFilter.Sort.SortFields.Clear
            MySheet.AutoFilter.Sort.SortFields.Add Key:=Columns(ColCB), SortOn:=xlSortOnValues, _
                Order:=xlDescending, DataOption:=xlSortNormal
                
            With MySheet.AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            
            Set tmpRng = Range(Area(1, 1).Offset(1, 0), Area(2, 2)).SpecialCells(xlCellTypeVisible).Cells
            tmpSum = Application.WorksheetFunction.Subtotal(9, Columns(ColCB))
            Countries(j).Val1 = Array("Manager A", tmpRng(1, ColCB).Value, _
                tmpRng(1, ColCB).Value / tmpSum)
            
            ' Bottom CB Manager
            .AutoFilter Field:=ColCB  ' Clear autofilter applied in Top code
            MySheet.AutoFilter.Sort.SortFields.Clear
            MySheet.AutoFilter.Sort.SortFields.Add Key:=Columns(ColCB), SortOn:=xlSortOnValues, _
                Order:=xlAscending, DataOption:=xlSortNormal
                
            With MySheet.AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            
            Set tmpRng = Range(Area(1, 1).Offset(1, 0), Area(2, 2)).SpecialCells(xlCellTypeVisible).Cells
            Countries(j).Val2 = Array("Manager B", tmpRng(1, ColCB).Value, _
                tmpRng(1, ColCB).Value / tmpSum)
            
            ' Top Local Manager
            .AutoFilter Field:=ColLoc, Criteria1:=">0"
            MySheet.AutoFilter.Sort.SortFields.Clear
            MySheet.AutoFilter.Sort.SortFields.Add Key:=Columns(ColLoc), SortOn:=xlSortOnValues, _
                Order:=xlDescending, DataOption:=xlSortNormal
                
            With MySheet.AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            
            Set tmpRng = Range(Area(1, 1).Offset(1, 0), Area(2, 2)).SpecialCells(xlCellTypeVisible).Cells
            tmpSum = Application.WorksheetFunction.Subtotal(9, Columns(ColLoc))
            
            Countries(j).Val3 = Array(tmpRng(1, 2).Value, tmpRng(1, ColLoc).Value, _
                tmpRng(1, ColLoc).Value / tmpSum)
            
            ' Bottom Local Manager
            .AutoFilter Field:=ColLoc
            MySheet.AutoFilter.Sort.SortFields.Clear
            MySheet.AutoFilter.Sort.SortFields.Add Key:=Columns(ColLoc), SortOn:=xlSortOnValues, _
                Order:=xlAscending, DataOption:=xlSortNormal
                
            With MySheet.AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            
            Set tmpRng = Range(Area(1, 1).Offset(1, 0), Area(2, 2)).SpecialCells(xlCellTypeVisible).Cells
            Countries(j).Val4 = Array(tmpRng(1, 2).Value, tmpRng(1, ColLoc).Value, _
                tmpRng(1, ColLoc).Value / tmpSum)
                
        End With
    Next j
    
    MySheet.AutoFilterMode = False
    
    ' Clear existing data
    Range(Area(1, 1), Area(2, 2)).ClearContents
    
    ' Headers
    Cells(1, 1).Value = "Market"
    Cells(1, 2).Value = "Ranking"
    Cells(1, 3).Value = "Manager"
    Cells(1, 4).Value = "CB/Local Net Sales"
    Cells(1, 5).Value = "Share of Aggregate Net Sales"
    
    ' Input extracted data
    For i = 1 To NumCtry
    
        Set tmpRng = Range(Area(1, 1).Offset(4 * (i - 1) + 1, 0), Area(1, 2).Offset(4 * (i - 1) + 4, 0))
        
        tmpRng(1, 1).Value = Countries(i).Name
        tmpRng(1, 2).Value = "Top Cross-Border Manager"
        tmpRng(2, 2).Value = "Bottom Cross-Border Manager"
        tmpRng(3, 2).Value = "Top Local Manager"
        tmpRng(4, 2).Value = "Bottom Local Manager"
        
        tmpRng(1, 3).Value = Countries(i).Val1(0)
        tmpRng(1, 4).Value = Countries(i).Val1(1)
        tmpRng(1, 5).Value = Countries(i).Val1(2)
        
        tmpRng(2, 3).Value = Countries(i).Val2(0)
        tmpRng(2, 4).Value = Countries(i).Val2(1)
        tmpRng(2, 5).Value = Countries(i).Val2(2)
        
        tmpRng(3, 3).Value = Countries(i).Val3(0)
        tmpRng(3, 4).Value = Countries(i).Val3(1)
        tmpRng(3, 5).Value = Countries(i).Val3(2)
        
        tmpRng(4, 3).Value = Countries(i).Val4(0)
        tmpRng(4, 4).Value = Countries(i).Val4(1)
        tmpRng(4, 5).Value = Countries(i).Val4(2)
        
    Next i
    
    ' Formatting
    Set Area(1, 1) = Cells(2, 1)
    Set Area(2, 2) = Area(1, 1).End(xlToRight).End(xlDown)
    With Range(Area(1, 1), Area(2, 2))
        .Rows.AutoFit
        .Columns.AutoFit
        .Columns(4).NumberFormat = "#,##0"
    End With
    
    Range(Columns(Area(2, 2).Column + 1), Columns(Columns.Count)).ClearFormats
    Range(Rows(Area(2, 2).row + 1), Rows(Rows.Count)).ClearFormats
    
    ' Remove negative share pct
    For Each c In Range(Area(1, 1), Area(2, 2)).Columns(5).Cells
        If c.Value < 0 Then
            c.Value = " "
        End If
    Next c
    
    ' Number format
    Range(Area(1, 2).Offset(1, 0), Area(2, 2)).NumberFormat = "#%"
    
    Cells(1, 1).Select
    
End Sub

Sub EuroTRQuartileData()

    Dim Quartiles(1 To 4)           As CMetaData
    Dim Area(2, 2), tmpRng          As Range
    Dim MySheet                     As Worksheet
    Dim CountryName, c              As Variant
    Dim i, j, Months                As Long
    Dim SumCB(), SumLocal()         As Double
    Dim Headers                     As Range
    
    
    Months = 13
    ReDim SumCB(1 To Months)
    ReDim SumLocal(1 To Months)
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
    
        For i = 1 To Months  ' each period
    
            With Range(Area(1, 1), Area(2, 2))
                .AutoFilter Field:=i + 1, Criteria1:=j
                
                ' Sum CB + Local net sales for quartile j in period i
                SumCB(i) = Application.WorksheetFunction.Subtotal(9, Columns(i + Months + 1))
                SumLocal(i) = Application.WorksheetFunction.Subtotal(9, Columns(i + 2 * Months + 1))
                
            End With
            
            MySheet.AutoFilterMode = False
            
        Next i
        
        ' Save CB & Local arrays in Quartiles
        Quartiles(j).Val1 = SumCB
        Quartiles(j).Val2 = SumLocal
        
    Next j
    
    ' Clear old data
    Range(Area(1, 1).Offset(1, 0), Area(2, 2)).ClearContents
    Range(Area(1, 1).Offset(0, Months + 1), Area(1, 1).Offset(0, 3 * Months)).Cut _
        Range(Area(1, 1).Offset(0, 1), Area(1, 1).Offset(0, 2 * Months))
    
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
    Set tmpRng = Range(Cells(2, 1), Cells(5, 2 * Months + 1))
    For j = 1 To 4
        tmpRng(j, 1).Value = j
        For i = 1 To Months
            tmpRng(j, i + 1).Value = Quartiles(j).Val1(i)
            tmpRng(j, i + Months + 1).Value = Quartiles(j).Val2(i)
        Next i
    Next j
    
    With tmpRng
        .Font.Size = 8
        .Rows.AutoFit
    End With
    Range(Columns(tmpRng(1, 1).Column), Columns(tmpRng.Columns.Count)).AutoFit
    
    ' Extra labels for clarity
    Set tmpRng = Range(Cells(1, 1), Cells(1, 1).End(xlToRight).End(xlDown))
    Range(Cells(2, 2), Cells(2, 2).Offset(3, 2 * Months - 1)).NumberFormat = "#,##0.00"
    tmpRng.Columns.AutoFit
    tmpRng.Cut tmpRng.Offset(1, 0)
    Cells(1, 2).Value = "Cross Border Net Sales"
    Cells(1, 2).Offset(0, Months).Value = "Local Net Sales"
    With Rows(1)
        .Font.Size = 8
    End With
    
    Cells(1, 1).Select
    
End Sub


Sub TrailTRData(Yrs As Long, Optional Cutoff As Long)

    ' This is a pre-processing step meant to be run prior to EuroTRQuartileData()
    ' for <Yrs>-year trailing ATR data
    ' Yrs: trailing period for TR geo avg to be calculated over
    ' Cutoff: max number of missing obs in TR trailing period that is acceptable

    Dim Area(2, 2), tmpRng, AvgTR   As Range
    Dim MySheet                     As Worksheet
    Dim i, Months, tmpCol           As Long
    Dim SumCB(), SumLocal()         As Double
    Dim CB, Loc, WAE, Quart         As Range
    Dim Headers                     As Range
    Dim F1, F2, L1, L2              As String
    
    
    ' Preliminaries
    Set MySheet = ActiveSheet
    Months = 13
    Set Area(1, 1) = Cells(1, 1)
    Set Area(2, 1) = Area(1, 1).End(xlDown)
    Set Area(1, 2) = Area(1, 1).End(xlToRight)
    Set Area(2, 2) = Cells(Area(2, 1).row, Area(1, 2).Column)
    
    ' Clear extraneous
    Range(Rows(Area(2, 1).row + 1), Rows(Rows.Count)).Delete
    Range(Cells(1, 1), Cells(Rows.Count, Columns.Count)).ClearFormats
    
    ' Specify ranges for Monthly WAE, CB & Local measures
    Set WAE = Range(Area(1, 1).Offset(1, 1), Area(2, 1).Offset(0, (Months - 1) * (Yrs + 1) + 1))
    Set CB = Range(Area(1, 1).Offset(1, WAE.Columns.Count + 1), _
        Area(2, 1).Offset(0, WAE.Columns.Count + Months))
    Set Loc = Range(CB(1, 1).Offset(0, CB.Columns.Count), _
        CB(1, 1).Offset(CB.Rows.Count - 1, CB.Columns.Count - 1 + Months))
    ' First blank column
    Set tmpRng = Range(Loc(1, 1).Offset(0, Loc.Columns.Count), _
        Loc(1, 1).Offset(Loc.Rows.Count - 1, Loc.Columns.Count))
    
    ' Edit headers
    Set Headers = Rows(1)
    Headers.Select
    Call EditHeader
    
    ' Drop records with too few TR data points (ie, cannot calc 3yr trailing for any month)
    tmpRng(1, 1).Formula = "=IF(COUNT(" & _
        WAE.Rows(1).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) _
        & ") < (" & Yrs * Months & "), 1, 0)"
    tmpRng(1, 1).Copy tmpRng
    
    Call RemoveTrue(Area)
    tmpRng.ClearContents
        
    ' Rescale return percentages to decimal (for GEOMEAN call below)
    Call Rescale(WAE)
    
    ' Setup Start/Stop cells for Avg TR
    F1 = WAE(1, 1).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
    F2 = WAE(1, 2).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
    L1 = WAE(1, 1).Offset(0, (Months - 1) * Yrs - 1).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
    L2 = WAE(1, 1).Offset(0, (Months - 1) * Yrs - 2).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
    
    ' Calc 36-month trailing Avg TR (provided not missing 1st, last, & no more
    ' than Cutoff missing values in remaining 34 months)
    tmpRng(1, 1).Formula = "=IF(" & _
        "AND(NOT(ISBLANK(" & F1 & ")), NOT(ISBLANK(" & L1 & ")), " & _
            "COUNT(" & F2 & ":" & L2 & ") >= " & (Months - 1) * Yrs - 2 - Cutoff & "), " & _
        "GEOMEAN(" & F1 & ":" & L1 & "), " & Chr(34) & Chr(34) & ")"
    
    Set AvgTR = Range(tmpRng(1, 1), tmpRng(1, 1).Offset(tmpRng.Rows.Count - 1, Months))
    tmpRng(1, 1).Copy AvgTR
    
    ' Remove formula
    AvgTR.Copy
    AvgTR.PasteSpecial Paste:=xlPasteValues
    
    ' Add (desc) Quartile at bottom (=QUARTILE.INC(IF(SUBTOTAL(...)...)...))
        ' must be array equation (ctrl-shift-enter)
    Set tmpRng = Range(AvgTR(1, 1).Offset(AvgTR.Rows.Count, 0), _
        AvgTR(1, 1).Offset(AvgTR.Rows.Count + 3, AvgTR.Columns.Count - 1))
    
    F1 = AvgTR.Columns(1).AddressLocal(RowAbsolute:=True, ColumnAbsolute:=False)
    For i = 1 To 4
        tmpRng(i, 1).Formula = "=QUARTILE.INC(" & F1 & ", " & 5 - i & ")"
    Next i
    
    tmpRng.Columns(1).Copy tmpRng
    
    ' Apply equation to assign quartile rank
    F1 = AvgTR(1, 1).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
    L1 = tmpRng.Columns(1).AddressLocal(RowAbsolute:=True, ColumnAbsolute:=False)
    WAE(1, 1).Formula = "=IFERROR(MATCH(" & F1 & ", " & L1 & ", -1), " & Chr(34) & Chr(34) & ")"
    
    Set Quart = Range(WAE.Columns(1), WAE.Columns(Months))
    WAE(1, 1).Copy Quart
    Quart.Copy
    Quart.PasteSpecial Paste:=xlPasteValues
        
    ' Clean up
    Set WAE = Nothing
    Set tmpRng = Nothing
    Set AvgTR = Nothing
    
    Range(Columns(Cells(1, 1).End(xlToRight).Column + 1), Columns(Columns.Count)).Clear
    Range(Rows(Cells(1, 1).End(xlDown).row + 1), Rows(Rows.Count)).Clear
    
    CB.Offset(-1, 0).Rows(1).Copy Quart.Offset(-1, 0).Rows(1)
    CB.Offset(-1, 0).Rows(1).Copy Quart.Offset(-1, Months).Rows(1)
    CB.Offset(-1, 0).Rows(1).Copy Quart.Offset(-1, 2 * Months).Rows(1)
    CB.Copy Quart.Offset(0, Months)
    Loc.Copy Quart.Offset(0, 2 * Months)
    
    Set CB = Nothing
    Set Loc = Nothing

    Range(Columns(Quart.Column + 3 * Months), Columns(Columns.Count)).Clear
    Set Quart = Nothing

    Call EuroTRQuartileData
        
End Sub

Sub InvTypeGrossData()

    Dim DataRange As Range
    Dim Regex     As New VBScript_RegExp_55.RegExp
    Dim Matches
    
    Set DataRange = Range(Cells(1, 1), Cells(1, 1).End(xlDown).End(xlToRight))
    Regex.Pattern = ".*Market.*"
    Set Matches = Regex.Execute(DataRange(1, 1).Value)
    
    ' Clear extraneous
    With Range(Rows(DataRange.Rows.Count + 1), Rows(Rows.Count))
        .ClearContents
        .ClearFormats
        .Rows.AutoFit
    End With
    
    ' Rename headers
    DataRange.Rows(1).Select
    EditHeader
    
    ' If no market column, then add Total into first column
    If Not Regex.Test(DataRange(1, 1).Value) Then
        Columns(1).Insert
        Cells(1, 1).Value = "Origination Market (SI)"
        Range(Cells(2, 1), Cells(DataRange.Rows.Count, 1)).Value = "Total"
    End If
    
End Sub

Sub ActiveETFData()

    Dim MySheet         As Worksheet
    Dim OrigRng         As Range
    Dim FinalRng, Vis   As Range
    Dim tmpRng          As Range
    Dim Cats            As New Scripting.Dictionary
    Dim tmpArr(1 To 2)  As Double
    Dim Val             As Variant
    Dim rec             As Variant
    Dim Regex           As New VBScript_RegExp_55.RegExp
    Dim Match
    
    Set MySheet = ActiveSheet
    Set OrigRng = Range(Cells(1, 1), Cells(1, 1).End(xlDown).End(xlToRight))
    
    ' Clear extraneous
    With Range(Rows(OrigRng.Rows.Count + 1), Rows(Rows.Count))
        .ClearContents
        .ClearFormats
        .Rows.AutoFit
    End With
    
    ' Copy unique category names
    With OrigRng.Columns(1)
        .AdvancedFilter xlFilterInPlace, Unique:=True
        .SpecialCells(xlCellTypeVisible).Cells.Copy
        OrigRng.Columns(1).Offset(0, OrigRng.Columns.Count).PasteSpecial xlPasteAll
        .AdvancedFilter xlFilterInPlace
    End With
    
    Set FinalRng = Range(Selection(1, 1).Offset(1, 0), Selection.End(xlDown))
    
    ' Grab data for Fund Category
    For Each Val In FinalRng.Cells.Value
    
        With OrigRng
            .AutoFilter Field:=1, Criteria1:=Val, VisibleDropDown:=False
            Set Vis = OrigRng.Offset(1, 0).Resize(OrigRng.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
            For Each rec In Vis.Rows
                If rec.Cells(2).Value = "Yes" Then
                    tmpArr(1) = rec.Cells(3).Value
                Else
                    tmpArr(2) = rec.Cells(3).Value
                End If
            Next rec
        
        End With
        
        Cats.Add Val, tmpArr
        Erase tmpArr
        
    Next Val
    
    MySheet.AutoFilterMode = False
    
    ' Clear old data, adjust headers
    OrigRng.Offset(1, 0).Resize(OrigRng.Rows.Count - 1).ClearContents
    OrigRng(1, 3).Select
    EditHeader
    OrigRng(1, 2).Value = "ETF " & OrigRng(1, 3).Value
    OrigRng(1, 3).Value = "Non-ETF " & OrigRng(1, 3).Value
    OrigRng(1, 4).ClearContents
    FinalRng.Cut FinalRng.Offset(0, -3)
    Set FinalRng = Range(Cells(2, 1), Cells(2, 1).End(xlDown))
    Set tmpRng = Range(FinalRng.End(xlDown).Offset(1, 0), _
        FinalRng.End(xlDown).Offset(1, 2))
    tmpRng.Cells(1, 1).Value = "Mixed"
    
    ' Paste values & combinen mixed categories
    Regex.Pattern = "Mixed.*"
    
    For Each rec In Range(FinalRng, FinalRng.Offset(0, 2)).Rows
        Set Matches = Regex.Execute(rec.Cells(1, 1).Value)
        If Regex.Test(rec.Cells(1, 1).Value) Then
            tmpRng.Cells(1, 2).Value = tmpRng.Cells(1, 2).Value + _
                Cats(rec.Cells(1, 1).Value)(1)
            tmpRng.Cells(1, 3).Value = tmpRng.Cells(1, 3).Value + _
                Cats(rec.Cells(1, 1).Value)(2)
        Else
            rec.Cells(1, 2).Value = Cats(rec.Cells(1, 1).Value)(1)  ' Yes
            rec.Cells(1, 3).Value = Cats(rec.Cells(1, 1).Value)(2)  ' No
        End If
    Next rec
        
    Set Cats = Nothing
    
    ' Move Mixed Total up
    Range(Rows(tmpRng(1, 2).End(xlUp).row + 1), Rows(tmpRng(1, 2).row - 1)).Delete
    
    ' Clean up
    Set FinalRng = Range(Cells(1, 1), Cells(1, 1).End(xlDown).End(xlToRight))
    FinalRng.Columns.AutoFit
    Cells(1, 1).Select
    
End Sub

Sub MarketNetTblData()

    Dim tblRng          As Range
    Dim HeaderCell      As Variant
    Dim Regex           As New VBScript_RegExp_55.RegExp
    Dim Match           As VBScript_RegExp_55.MatchCollection
    Dim MonthText       As String
    
    Set tblRng = Cells(1, 1).CurrentRegion
    
    ' Clear extraneous
    With Range(Rows(tblRng.Rows.Count + 1), Rows(Rows.Count))
        .ClearContents
        .ClearFormats
        .AutoFit
    End With
    
    ' Edit headers
    For Each HeaderCell In tblRng.Rows(1).Cells
        HeaderCell.Value = Replace(HeaderCell.Value, " (SI)", "")
                
        ' Grab date
        Regex.Pattern = _
            "((?:Local|Cross Border) Net Sales EUR)\s\[(\d \w{2,3})-(\d{1,2})/(\d\d)\]"
        Set Match = Regex.Execute(HeaderCell.Value)
        If Regex.Test(HeaderCell.Value) Then
            MonthText = MonthName(Match(0).SubMatches(2), True) & " " & Match(0).SubMatches(3)
            HeaderCell.Value = Match(0).SubMatches(0) & " " & MonthText
            If Match(0).SubMatches(1) = "1 Yr" Then
                HeaderCell.Value = HeaderCell.Value + " Past Yr"
            End If
        End If
    Next HeaderCell
    
End Sub
