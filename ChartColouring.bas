Attribute VB_Name = "ChartColouring"
Option Explicit

' ================================================================================================================================
' This code was written by Dan Golding. You are free to use it and
' edit is as you like so long as you leave this attribution.
' The latest version can be found here: https://github.com/DanGolding/Scatter-plot-with-color-grading-in-Excel
' ================================================================================================================================

Sub FetchDataFromTextFile(filename As String, sheetname As String)
'Read a text file containing a single line of space separated hex values and write them to column A of a worksheet
'Based on https://stackoverflow.com/a/16668538/1011724
    Dim line As Long
    Dim LineText As String
    Dim row As Integer
    Dim element As Integer
    Open filename For Input As #24
    row = 1
    While Not EOF(24)
        If line > 1 Then
            MsgBox "The text file should only have one line, did not read all the data!"
            Close #24
            Exit Sub
        Else
            Line Input #24, LineText
                Dim arr
                arr = Split(CStr(LineText), " ")
                For element = 1 To UBound(arr) + 1
                    Sheets(sheetname).Cells(element, row).Value = arr(element - 1)
                Next element
                line = line + 1
        End If
    Wend
    Close #24
End Sub

Sub MakeMap()
' Create the colour map on a new sheet as required by the chart colouring scripts. This requires a text file
' in the same directory as the worksheet that contains a single line of space separated hex values of the form
' #000000

    Dim filename As String
    Dim directory As String
    Dim fullFilename As String
    Dim cmapname As String

    'Name of the text file (without the extension)
    filename = "Colour Map (Sequential)"
    
    'the full path of the folder where the text file lives
    directory = "C:\Users\Bloggs\My Documents\" '< EDIT THIS
    
    'check directory ends in a backslash
    If Right(directory, 1) <> "\" Then directory = directory & "\"
    fullFilename = directory & filename & ".txt"
    
    If Dir(fullFilename) = "" Then MsgBox "Could not find the text file. Check the file name matches the name in the code"

    'Create a new sheet for the colour map
    Sheets.Add After:=ActiveSheet
    ActiveSheet.name = filename
    Dim sheetMap As Worksheet
    Set sheetMap = Sheets(name)
    
    'Read the colour map from the text file and store it in column A
    Call FetchDataFromTextFile(fullFilename & ".txt", name)
    
    'Convert the hex numbers to RGB values
    Dim lastRow As Integer
    lastRow = sheetMap.Range("A1").End(xlDown).row
    sheetMap.Range("B1:B" & lastRow).FormulaR1C1 = "=HEX2DEC(MID(RC1,2,2))"
    sheetMap.Range("C1:C" & lastRow).FormulaR1C1 = "=HEX2DEC(MID(RC1,4,2))"
    sheetMap.Range("D1:D" & lastRow).FormulaR1C1 = "=HEX2DEC(MID(RC1,6,2))"
    
    'Create an in situ visualisation of the colour map
    Dim row As Integer
    For row = 1 To lastRow
        Range("F" & row).Interior.Color = RGB(sheetMap.Range("B" & row).Value, sheetMap.Range("C" & row).Value, sheetMap.Range("D" & row).Value)
    Next row
    
End Sub

Function valueToMapPosition(datum As Variant, dataMin As Double, dataMax As Double, n As Integer) As Integer
' Normalise your data to fall in the range of the colour map for ease of lookup
    valueToMapPosition = CInt(((datum - dataMin) / (dataMax - dataMin)) * (n - 1)) + 1
End Function

Sub colourChartSequential()

' Colour a scatter chart according to sequential data (i.e. data that has no center such as dates)
    
    Dim sheetData As Worksheet
    Dim dataStartCol As String
    Dim dataStartRow As String
    Dim chartName As String
    ' Change the following four parameters to match your needs. The data according to which
    ' you wish to colour your chart should be in a single column starting in the cell specified
    ' by dataStartCol and dataStartRow
    Set sheetData = Worksheets("Divergent")
    dataStartCol = "A"
    dataStartRow = "2"
    chartName = "SequentialColour"

    ' sheetMap should be created by first runnin the MakeMap sub
    Dim sheetMap As Worksheet
    Set sheetMap = Worksheets("Colour Map (Sequential)")
    Dim n As Integer
    n = sheetMap.Range("A1").End(xlDown).row
    
    ' Read in the data and find its range
    Dim data As Variant
    Dim dataMin As Double
    Dim dataMax As Double
    Dim lastRow As Integer
    lastRow = sheetData.Range(dataStartCol & dataStartRow).End(xlDown).row
    data = sheetData.Range(dataStartCol & dataStartRow & ":" & dataStartCol & lastRow).Value2
    dataMin = Application.Min(data)
    dataMax = WorksheetFunction.Max(data)
    
    With sheetData.ChartObjects(chartName).Chart.FullSeriesCollection(1)
    
        Dim Count As Integer
        Dim colourRow As Integer
        Dim datum As Variant
        For Count = 1 To UBound(data)
             datum = data(Count, 1)
                colourRow = valueToMapPosition(datum, dataMin, dataMax, n)
                .Points(Count).Format.Fill.BackColor.RGB = RGB(sheetMap.Range("B" & colourRow).Value, sheetMap.Range("C" & colourRow).Value, sheetMap.Range("D" & colourRow).Value)
        Next Count
        
    End With

End Sub

Sub colourChartDivergent()
' Colour a scatter chart according to divergent data (i.e. data that has a center such as 0. In fact, this
' module currently assumes that the data is centered at 0. If this is not the case, alter the calculations
' of dataMin and dataMax below)
    
    Dim sheetData As Worksheet
    Dim dataStartCol As String
    Dim dataStartRow As String
    Dim chartName As String
    ' Change the following four parameters to match your needs. The data according to which
    ' you wish to colour your chart should be in a single column starting in the cell specified
    ' by dataStartCol and dataStartRow
    Set sheetData = Worksheets("Divergent")
    dataStartCol = "D"
    dataStartRow = "2"
    chartName = "DivergentColour"
    
    ' sheetMap should be created by first runnin the MakeMap sub
    Dim sheetMap As Worksheet
    Set sheetMap = Worksheets("Colour Map (Divergent)")
    Dim n As Integer
    n = sheetMap.Range("A1").End(xlDown).row
    
    ' Read in the data and find its range, center the range at 0
    Dim data As Variant
    Dim dataMin As Double
    Dim dataMax As Double
    Dim lastRow As Integer
    lastRow = sheetData.Range(dataStartCol & dataStartRow).End(xlDown).row
    data = sheetData.Range(dataStartCol & dataStartRow & ":" & dataStartCol & lastRow).Value2
    dataMin = WorksheetFunction.Min(data)
    dataMax = WorksheetFunction.Max(data)
    ' NB: This assumes the data is centered at 0, if it isn't then shift this min and max accordingly
    dataMax = WorksheetFunction.Max(dataMax, -dataMin)
    dataMin = -dataMax
    
    With sheetData.ChartObjects(chartName).Chart.FullSeriesCollection(1)
    
        Dim Count As Integer
        Dim colourRow As Integer
        Dim datum As Variant
        For Count = 1 To UBound(data)
            datum = data(Count, 1)
            colourRow = valueToMapPosition(datum, dataMin, dataMax, n)
            .Points(Count).Format.Fill.BackColor.RGB = RGB(sheetMap.Range("B" & colourRow).Value, sheetMap.Range("C" & colourRow).Value, sheetMap.Range("D" & colourRow).Value)
        Next Count
        
    End With

End Sub

Sub MakeColourBar()
    ' Create a new sheet with the colour bar on it. To use it, copy cells A1:D258 and paste them as a linked
    ' image. Resize the image keeping the aspect ratio constant and then resize the fontsize in column D
    ' of the colour bar to visually match the size of you chart's axis labels.
    ' NB: You need to put the min (Start) and max (End) values on the sheet yourself manually (using a
    ' formula if you want the colour bar to update dynamically). Also note that if you resize anything in
    ' the colour bar such as the volumn widths, you should create a new linked image otherwise the aspect
    ' ratio may become distorted.
    
    'Enter the parameters below i.e. the name of the sheet with the colour map and the name of the new sheet with your colour bar
    Dim name As String
    ' sheetMap should be first created by running the MakeMap sub
    Dim sheetMap As Worksheet
    name = "Colour Bar (Divergent)"
    Set sheetMap = Worksheets("Colour Map (Divergent)")
    
    'Create a new sheet for the colour map
    Sheets.Add After:=ActiveSheet
    ActiveSheet.name = name
    
    ' NB: Currently only supports n = 256 and n_ticks = 8
    Dim n As Integer ' The colour axis resolution i.e. the number of colour shown on the colour bar
    Dim n_ticks As Integer ' The number of tick marks shows
    n = 256
    n_ticks = 8
    
    ' Start, End and Step values for the colour bar tick marks
    Range("A" & n + 4).Value = "Start"
    Range("A" & n + 5).Value = "End"
    Range("A" & n + 6).Value = "Step"
    Range("D" & n + 6).FormulaR1C1 = "=(R[-1]C-R[-2]C)/" & n_ticks

    Dim row As Integer
    For row = 1 To n
        Range("B" & row + 1).Interior.Color = RGB(sheetMap.Range("B" & n - row + 1).Value, sheetMap.Range("C" & n - row + 1).Value, sheetMap.Range("D" & n - row + 1).Value)
    Next row
    
    ActiveWindow.DisplayGridlines = False
    Rows("2:257").RowHeight = 2
    Rows("1:1").RowHeight = 7.5 'This is for the tick mark labels
    Rows("258:258").RowHeight = 7.5 'This is for the tick mark labels
    Columns("B:B").ColumnWidth = 2.14
    
    With Range("B2:B257")
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
    End With
    
    Range("D1:D6").Merge
    Range("D1").Value = "=D261"
    Range("D253:D258").Merge
    Range("D253").Value = "=D260"
    'Merge rows for tick marks
    Dim mark As Integer
    For mark = 1 To 8
        Range("C" & (mark - 1) * (256 / 8) + 2 & ":C" & (mark) * (256 / 8) + 1).Merge
        Range("C" & (mark - 1) * (256 / 8) + 2).Borders(xlEdgeTop).Weight = xlMedium
        'Make the tick mark labels by merging the 10 cells in column D that center around each tick label
        If mark > 1 Then
            Range("D" & (mark - 1) * (256 / 8) + 2 - 5 & ":D" & (mark - 1) * (256 / 8) + 2 + 4).Merge
            Range("D" & (mark - 1) * (256 / 8) + 2 - 5).Value = "=D" & (mark) * (256 / 8) + 2 - 5 & " + D262"
        End If
    Next mark
    Range("C257").Borders(xlBottom).Weight = xlMedium
    
    Columns("C:C").ColumnWidth = 0.42
    Columns("D:D").VerticalAlignment = xlCenter
    Columns("D:D").HorizontalAlignment = xlLeft
    
End Sub



