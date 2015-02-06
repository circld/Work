Attribute VB_Name = "CreatePowerPoint"
Sub CreatePowerPoint()

 'Add a reference to the Microsoft PowerPoint Library by:
    '1. Go to Tools in the VBA menu
    '2. Click on Reference
    '3. Scroll down to Microsoft PowerPoint X.0 Object Library, check the box, and press Okay
 
    'First we declare the variables we will be using
        Dim newPowerPoint As PowerPoint.Application
        Dim activeSlide As PowerPoint.Slide
        Dim cht As Excel.ChartObject
        Dim pic As Picture
        Dim tmpRng As Range
        Dim counter As Integer
     
     'Look for existing instance
        On Error Resume Next
        Set newPowerPoint = GetObject(, "PowerPoint.Application")
        On Error GoTo 0
     
    'Let's create a new PowerPoint
        If newPowerPoint Is Nothing Then
            Set newPowerPoint = New PowerPoint.Application
        End If
    'Make a presentation in PowerPoint
        If newPowerPoint.Presentations.Count = 0 Then
            newPowerPoint.Presentations.Add
        End If
     
    'Show the PowerPoint
        newPowerPoint.Visible = True
    
    'Loop through each chart in the Excel worksheet and paste them into the PowerPoint
        For Each sht In ActiveWorkbook.Worksheets
            
            sht.Activate
            
            ' 2 charts per slide for Global TopBottom
            If sht.Name = "Market Global TopBottom 5 Se" Then
                i = 1
                For Each cht In ActiveSheet.ChartObjects
                    
                    If i Mod 2 = 1 Then
                        newPowerPoint.ActivePresentation.Slides.Add newPowerPoint.ActivePresentation.Slides.Count + 1, ppLayoutText
                        newPowerPoint.ActiveWindow.View.GotoSlide newPowerPoint.ActivePresentation.Slides.Count
                        CopyChartToPP newPowerPoint, cht
                    Else
                        newPowerPoint.ActiveWindow.View.GotoSlide newPowerPoint.ActivePresentation.Slides.Count
                        CopyChartToPP newPowerPoint, cht, 15 + cht.Width
                    End If
                    i = i + 1
                Next cht
                Set i = Nothing
            ' Copy picture object (table)
            ElseIf sht.Name = "Performance Global TopBottom" Then
                    Set tmpRng = Range(Cells(1, 1), Cells(17, 3))
                    tmpRng.Copy
                    tmpRng.Offset(0, 4).Select
                    ActiveSheet.Pictures.Paste.Select
                    Set pic = Selection
                    
                    newPowerPoint.ActivePresentation.Slides.Add newPowerPoint.ActivePresentation.Slides.Count + 1, ppLayoutText
                    newPowerPoint.ActiveWindow.View.GotoSlide newPowerPoint.ActivePresentation.Slides.Count
                    CopyPicToPP newPowerPoint, pic
                    
                    pic.Delete
            Else
                For Each cht In ActiveSheet.ChartObjects
                    
                    newPowerPoint.ActivePresentation.Slides.Add newPowerPoint.ActivePresentation.Slides.Count + 1, ppLayoutText
                    newPowerPoint.ActiveWindow.View.GotoSlide newPowerPoint.ActivePresentation.Slides.Count

                    CopyChartToPP newPowerPoint, cht
                    
                Next cht
            End If
        Next sht
     
    AppActivate title:="Presentation1 - PowerPoint"
    Set activeSlide = Nothing
    Set newPowerPoint = Nothing
     
End Sub

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
                activeSlide.Shapes(1).TextFrame.TextRange.text = title & vbNewLine
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
