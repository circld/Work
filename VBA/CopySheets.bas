Attribute VB_Name = "CopySheets"
Sub CopySheets()

    Dim oBook               As Workbook
    Dim oSheet              As Worksheet
    Dim template            As String
    Dim StrFile             As String
    Dim i                   As Long
    
    
    directory = ThisWorkbook.Path & "\"
    StrFile = dir(directory & "*.xlsx")  ' get first match
    
    ' Performance
    Application.ScreenUpdating = False
    
    ' Open template file
    template = Cells(3, 2).Value
    Workbooks.Open directory & template
    
    Do While Len(StrFile) > 0
    
        If StrFile <> template Then
            Workbooks.Open directory & "\" & StrFile
            
            ' Call function to do all the work
            AddNewTab
            
            ActiveWorkbook.Close True
        End If
            
        StrFile = dir  ' get next match
    
    Loop
    
    ' Close template file
    Workbooks(template).Close False
    
End Sub

Sub AddNewTab()

    Dim oBook       As Workbook
    Dim oSheet      As Worksheet
    Dim srcBook     As Workbook
    Dim tempWb      As String
    Dim tempSht     As String
    Dim newShtName  As String
    Dim lnks        As Variant
    Dim lnk         As Variant
    
    
    Set oBook = ActiveWorkbook
    Set srcBook = Workbooks("CopyTabs.xlsb")
    
    ' Get template params
    srcBook.Activate
    With Worksheets(1)
        tempWb = .Cells(3, 2).Value
        tempSht = .Cells(4, 2).Value
        newShtName = .Cells(5, 2).Value
    End With
    
    ' Copy tab from template and paste as new tab in open file
    Workbooks(tempWb).Worksheets(tempSht).Activate
    ActiveSheet.Copy After:=oBook.Worksheets(oBook.Worksheets.Count)
    ActiveSheet.Name = newShtName
    
    ' Remove refs to other workbook
    lnks = oBook.LinkSources(xlExcelLinks)
    
    For Each lnk In lnks
        If lnk = Workbooks(tempWb).FullName Then
            oBook.ChangeLink Name:=lnk, NewName:=oBook.FullName
        End If
    Next lnk
    
End Sub
