Attribute VB_Name = "CreatePDF"
Sub Main()

    Dim fileName As String
    Dim directory As String
    Dim zoomLevel As Integer
    Dim reportCount As Integer
    
    ' Improve performance
    Application.ScreenUpdating = False
    
    ' Set variables
    reportCount = 0
    With Application.FileDialog(msoFileDialogFolderPicker)
       .AllowMultiSelect = False
       .Show
       On Error Resume Next
       directory = .SelectedItems(1)
       Err.Clear
       On Error GoTo 0
    End With
    
    If directory = "" Then
        End
    End If
    directory = directory & "\"
    fileName = Dir(directory & "*_REPORT.xlsx")
    zoomLevel = Cells(5, 3).Value
    
    ' Iterate through matching files (when no matching file left, returns "")
    Do While Len(fileName) > 0
    
        CreatePDF directory & fileName
        reportCount = reportCount + 1
        fileName = Dir  ' get next match
    
    Loop
    
    MsgBox reportCount & " reports have been generated.", vbMsgBoxSetForeground
    
End Sub

Sub CreatePDF(fileName As String)

    ' Define variables & types
    Dim reportName As String
    Dim source As Workbook
    Dim dest As Workbook
    Dim methodologyText As Range
    Dim fundSummaryTable As Range
    Dim expenseGroupRankingsTable As Range
    Dim netTotalExpenseRatioTable As Range
    Dim destM, destFS, destGR, destExpense As Range
    Dim i, j, printRowCutoff, printColCutoff As Integer
    Dim tmpRng As Range
    Dim tbl As Variant
    
    ' Next line will need to change if naming convention changes
    reportName = Left(fileName, Len(fileName) - 14)
    
    ' Warn user to close report before proceeding
    If IsFileOpen(fileName) Then
        MsgBox ("Please close " & fileName & " before proceeding. Program exiting.")
        If True Then End  ' Exit program
    End If
    
    ' Set source & destination workbooks to variables
    Workbooks.Open fileName
    Set source = ActiveWorkbook
    Workbooks.add
    Set dest = ActiveWorkbook
    
    ' Hide gridlines in destination worksheet
    ActiveWindow.DisplayGridlines = False

    ' Assign reports to named ranges
    source.Worksheets("Methodology").Activate
    Set methodologyText = source.Worksheets("Methodology").Range( _
        Cells(2, 1), Cells(Cells(2, 3).End(xlDown).End(xlDown).row, Cells(2, 1).End(xlToRight).Column))
    Set fundSummaryTable = GetRange(1, 1, 0, 5)
    Set expenseGroupRankingsTable = GetRange(1, 1, 0, 7, 2, 3)
    
    ' Copy last column in each table & paste value to remove formulas
    For Each tbl In Array(fundSummaryTable, expenseGroupRankingsTable)
        i = tbl.Columns.Count
        Set tmpRng = ActiveSheet.Range(tbl.Cells(1, 2), tbl.Cells(1, i).End(xlDown))
        tmpRng.Select
        With tmpRng
            .Copy
            .PasteSpecial xlPasteValues
        End With
        
        ' Go through each cell in the first column of tbl and paste value if formula
        ' nb. formula will not have superscript
        For j = 1 To tbl.Rows.Count
            Set tmpRng = tbl(j, 1)
            If tmpRng.HasFormula Then
                tmpRng.Copy
                tmpRng.PasteSpecial xlPasteValues
            End If
        Next j
    Next tbl
    
    expenseGroupRankingsTable.Select
    source.Worksheets("Total Expense Ratio").Activate
    Set netTotalExpenseRatioTable = GetRange(2, 1, 0, 0, 1, 6)
    
    ' Copy to destination workbook
    
    ' Methodology text
    source.Worksheets("Methodology").Activate
    methodologyText.Copy
    dest.Worksheets(1).Activate
    Set destM = ActiveSheet.Range( _
        Cells(2, 1), Cells(2, 1).Offset( _
        methodologyText.Rows.Count - 1, methodologyText.Columns.Count - 1))
    
    With destM
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        .PasteSpecial xlPasteColumnWidths
        .Rows.RowHeight = 16.5
        .Rows(1).RowHeight = 17.25
    End With
    destM.Select
        

    ' Fund summary
    source.Worksheets("Methodology").Activate
    fundSummaryTable.Copy
    dest.Worksheets(1).Activate
    Set destFS = ActiveSheet.Range( _
        Cells(28, 1), Cells(28, 1).Offset( _
        fundSummaryTable.Rows.Count - 1, fundSummaryTable.Columns.Count - 1))
    
    With destFS
        .PasteSpecial xlPasteFormats
        .PasteSpecial xlPasteAll
        .Rows.RowHeight = 17
        .Rows(1).RowHeight = 17.25
    End With
    
    ' Group rankings
    source.Worksheets("Methodology").Activate
    expenseGroupRankingsTable.Select
    Selection.Copy
    dest.Worksheets(1).Activate
    Set destGR = ActiveSheet.Range( _
        Cells(43, 1), Cells(43, 1).Offset(expenseGroupRankingsTable.Rows.Count - 1, _
            expenseGroupRankingsTable.Columns.Count - 1))

    With destGR
        .PasteSpecial xlPasteFormats
        .PasteSpecial xlPasteAll
        .Rows.RowHeight = 17
        .Rows(1).RowHeight = 17.25
    End With
   
    ' Net total expense ratio
    dest.Activate
    dest.Worksheets.add After:=dest.Worksheets(1)
    ActiveWindow.DisplayGridlines = False
    
    source.Worksheets("Total Expense Ratio").Activate
    netTotalExpenseRatioTable.Copy
    
    dest.Worksheets(2).Activate
    Set destExpense = ActiveSheet.Range( _
        Cells(2, 1), Cells(2, 1).Offset(netTotalExpenseRatioTable.Rows.Count - 1, _
            netTotalExpenseRatioTable.Columns.Count - 1))
        
    ' No formulas to worry about-- xlPasteAll will include the superscripts properly
    With destExpense
        .PasteSpecial xlPasteFormats
        .PasteSpecial xlPasteColumnWidths
        .PasteSpecial xlPasteAll
        .Rows.RowHeight = 15
        .Rows(1).RowHeight = 17.25
    End With
    
    ' Print formatting (iterates over two worksheets in destination file)
    For i = 1 To 2
        If (i = 1) Then
            printRowCutoff = ThisWorkbook.Worksheets("Control Panel").Cells(6, 4).Value
            printColCutoff = destGR.Columns.Count
        Else
            printRowCutoff = ThisWorkbook.Worksheets("Control Panel").Cells(9, 4).Value
            printColCutoff = destExpense.Columns.Count
        End If
        dest.Worksheets(i).Activate
        
        ' Add Footer divider line
        With Range(Cells(printRowCutoff, 1), Cells(printRowCutoff, printColCutoff)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ThemeColor = 4
        End With
        
        With ActiveSheet.PageSetup
            .LeftFooter = _
                "&""Book Antiqua""&10 Strategic Insight, an Asset International Company"
            .RightFooter = "&""Book Antiqua""&10&P"
            .PrintArea = "$A$1:" & Cells(printRowCutoff, printColCutoff).Address( _
                RowAbsolute:=True, ColumnAbsolute:=True)
            .ScaleWithDocHeaderFooter = False  ' want consistent footer size across pages
            .AlignMarginsHeaderFooter = True
            .CenterHorizontally = True
            .CenterVertically = True
            ' Want to fit page to width & adjust footer line by row #
            .zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False  ' either 1 or False... good work MSFT
            .LeftMargin = 18
            .TopMargin = 18
            .RightMargin = 18
            .HeaderMargin = 18
            .FooterMargin = 18
            .BottomMargin = 50
        End With
    Next i
    
    ' Save as PDF
    ActiveWorkbook.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=reportName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    
    ' Clean up
    source.Close SaveChanges:=False
    dest.Close SaveChanges:=False

End Sub


Function GetRange( _
    Optional startRow As Integer = 1, _
    Optional startCol As Integer = 1, _
    Optional startRight As Integer = 1, _
    Optional startDown As Integer = 1, _
    Optional endRight As Integer = 1, _
    Optional endDown As Integer = 1 _
    ) As Range

    Dim startCell, endCell As Range
    Dim firstRow, lastRow, firstCol, lastCol As Integer
    Dim r, d As Integer
    
    Set startCell = ActiveSheet.Cells(startRow, startCol)
    
    ' Start cell
    For d = 1 To startDown
        Set startCell = startCell.End(xlDown)
    Next d
    firstRow = startCell.row
    
    For r = 1 To startRight
        Set startCell = startCell.End(xlToRight)
    Next r
    firstCol = startCell.Column
    
    ' End cell
    Set endCell = startCell
    
    For d = 1 To endDown
        Set endCell = endCell.End(xlDown)
    Next d
    lastRow = endCell.row
    
    Set endCell = startCell
    
    For r = 1 To endRight
        Set endCell = endCell.End(xlToRight)
    Next r
    lastCol = endCell.Column
    
    Set GetRange = Range(Cells(firstRow, firstCol), Cells(lastRow, lastCol))

End Function

Function IsFileOpen(fileName As String)

    ' Taken from:
    ' https://support.microsoft.com/en-us/kb/291295
    
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open fileName For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True

        ' Another error occurred.
        Case Else
            Error errnum
    End Select

End Function
