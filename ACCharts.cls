VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Sub chartCopy(templateChart As Chart, Optional Copies As Double = 1, Optional LocRow As Double, Optional LocCol As Double, Optional TitRow As Double, _
Optional TitCol As Double, Optional useChTitleOffset As Boolean = True, Optional YRow As Double, Optional YCol As Double, _
Optional Xrow As Double, Optional XCol As Double)
    'Last modified: 2014029
    
    'Define Variables
    Dim SeriesNum As Integer, SeriesAmount As Integer, SeriesPaste As Integer, Quotdel As Integer
    Dim templateChartPosition As Range
    Dim TmplTitle() As Range, TmplX() As Range, TmplY() As Range, ErrRan() As Range
    Dim NewChart As Chart
    Dim NewSeriesX As Range, NewSeriesY As Range, NewSeriesTitle As Range
    Dim LastChartNum As Long, Copy As Long
    Dim Quot As String
    Dim isTitRange() As Boolean, xRange() As Boolean, ChtTitIsRange As Boolean
    Quot = """"
    ChtTitIsRange = True

    'Get number of series in chart and redim arrays accordingly
    SeriesAmount = templateChart.SeriesCollection().Count
    ReDim TmplTitle(SeriesAmount) As Range, TmplX(SeriesAmount) As Range, TmplY(SeriesAmount) As Range, ErrRan(SeriesAmount) As Range
    ReDim isTitRange(SeriesAmount) As Boolean, xRange(SeriesAmount)
    
       
    'Get chart position and align template chart to closest cell
    Set templateChartPosition = chartCell(templateChart.Parent)
    templateChart.Parent.Top = templateChartPosition.Top
    templateChart.Parent.Left = templateChartPosition.Left
    
    Application.DisplayStatusBar = True
    Application.StatusBar = "Copying Charts..."
    Application.ScreenUpdating = False
        
    'Test whether chart has title and if yes is it a reference or typed in. Returns ChtTitIsRange = true if a range
    On Error Resume Next
    If ActiveChart.ChartTitle.Formula = ActiveChart.ChartTitle.Text Or useChTitleOffset = False Then
        ChtTitIsRange = False
    End If
    chtTitle = ActiveChart.ChartTitle.Formula
    On Error GoTo 0
    
    'Get Series data for every series in chart
    For SeriesNum = 1 To SeriesAmount
        Dim CFml As String, SrsFml As String
        
        'Trim Formula
        SrsFml = Right(ActiveChart.SeriesCollection(SeriesNum).Formula, Len(ActiveChart.SeriesCollection(SeriesNum).Formula) - 8)
        CFml = Replace(Left(SrsFml, Len(SrsFml) - 3 - Int(SeriesNum / 10)), Quot, "")
        
        If InStr(1, SrsFml, Quot, 1) = 0 And Not Left(SrsFml, 1) = "," Then     'Is series Title a Range (or typed in)
            isTitRange(SeriesNum) = True 'Title is a range, assign title range to variable
            Set TmplTitle(SeriesNum) = SrsTitle(CFml)
        Else
            isTitRange(SeriesNum) = False 'typed in or no Series title, do nothing
        End If
        
        If InStr(1, SrsFml, "{", 1) = 0 Then 'Are x values a Range (or typed in), could also be y values but that should be fairly rare
            xRange(SeriesNum) = True 'Range
            Set TmplX(SeriesNum) = SrsX(CFml)
        Else
            xRange(SeriesNum) = False 'single number
        End If
            
        'Assign Y to range
        Set TmplY(SeriesNum) = SrsY(CFml)
        
    Next SeriesNum
           
    
    'create new charts
    For Copy = 1 To Copies
        Application.ScreenUpdating = True
        Application.StatusBar = "Copying Chart " & Copy & "of " & Copies
        Application.ScreenUpdating = False
        
        'Copy, activate and size the same size as original
        templateChart.ChartArea.Copy
        LastChartNum = ActiveSheet.ChartObjects().Count + 1
        ActiveSheet.Paste
        ActiveSheet.ChartObjects(LastChartNum).Activate
        Set NewChart = ActiveChart
        ActiveChart.Parent.Height = templateChart.Parent.Height
        ActiveChart.Parent.Width = templateChart.Parent.Width
        
        If ChtTitIsRange = True Then
            ActiveChart.ChartTitle.Formula = "=" & "'" & ActiveSheet.name & "'!" & Range(chtTitle).Offset(YRow * Copy, YCol * Copy).Address
        End If
        
        'chart positioning
        Dim locOffsetRow As Range
        Dim locOffsetcolumn As Range
        
        Set locOffsetRow = templateChartPosition.Offset(LocRow * Copy, 0)
        NewChart.Parent.Top = templateChartPosition.Top + Range(templateChartPosition, locOffsetRow).Height - locOffsetRow.Height 'Set row offset for copy location
        
        Set locOffsetcolumn = templateChartPosition.Offset(0, LocCol * Copy)
        NewChart.Parent.Left = templateChartPosition.Left + Range(templateChartPosition, locOffsetcolumn).Width - locOffsetRow.Width 'Set column offset for copy location 'set column offset for copy location
                
        'paste values in Series
        For SeriesNum = 1 To SeriesAmount
                
            With NewChart.SeriesCollection(SeriesNum)
                If isTitRange(SeriesNum) = True Then
                    Set NewSeriesTitle = TmplTitle(SeriesNum).Offset(TitRow * Copy, TitCol * Copy) 'Title range offset for copy
                    .name = NewSeriesTitle
                End If
                
                If xRange(SeriesNum) = True Then
                    Set NewSeriesX = TmplX(SeriesNum).Offset(Xrow * Copy, XCol * Copy) 'X range offset for copy
                    .XValues = NewSeriesX
                End If
                Set NewSeriesY = TmplY(SeriesNum).Offset(YRow * Copy, YCol * Copy) 'Y range offset for copy
                .Values = NewSeriesY
                                
            End With
        Next SeriesNum
    Next Copy
    
    Application.DisplayStatusBar = False
    Application.ScreenUpdating = True
    
End Sub

Function SrsTitle(Formula As String) As Range
    'Last modified: 20140216
    'returns Range for the Title of a
    Dim SeriesAllRanges() As String
    Dim Quot As String
    Quot = """"
    
    'Split off before first comma
    SeriesAllRanges() = Split(Formula, ",")
    If SeriesAllRanges(0) = "" Or Left(SeriesAllRanges(0), 1) = Quot Then 'if no title or typed in return nothing
        Exit Function
    End If
    Set SrsTitle = Range(SeriesAllRanges(0))

End Function

Function SrsX(Formula As String) As Range
    'Last modified: 20140214
    'returns x-value of a series as a Range
    'returns Range A0 if x value range is not defined

        Dim pos As Long, cmmct As Long, prntct As Long
        Dim cmmpos() As Long, prnto() As Long, prntc() As Long
        Dim SeriesAllRanges() As String
         
        
    'count commas and parenthesis in formula
        For pos = 1 To Len(Formula)
            If Mid(Formula, pos, 1) = "," Then
                cmmct = cmmct + 1
                ReDim Preserve cmmpos(cmmct)
                cmmpos(cmmct) = pos
            End If
            If Mid(Formula, pos, 2) = ",(" Then
                prntct = prntct + 1
                ReDim Preserve prnto(prntct)
                prnto(prntct) = pos
            End If
            If Mid(Formula, pos, 1) = ")" Then
                ReDim Preserve prntc(prntct)
                prntc(prntct) = pos
            End If
        Next pos


    'Extract X values
        If cmmct > 2 And prntct = 0 Then 'commas in series name, X and Y simple
            MsgBox ("This Macro doesn't work on series with commas in name (reference to cell with comma is ok though)")
        ElseIf cmmct = 2 Then       'no complex range
            SeriesAllRanges() = Split(Formula, ",")
            If Left(SeriesAllRanges(1), 1) = "{" Then 'if x is not a range but values input by hand return nothing. could also be y values but that should be rare
                Exit Function
            Else
                Set SrsX = Range(SeriesAllRanges(1))
            End If
            
        ElseIf prntct = 2 Then      'x and y complex, commas in title don't matter
            SeriesAllRanges() = Split(Formula, ",(")
            Set SrsX = Range("(" & SeriesAllRanges(1))
            
        ElseIf prntc(1) = Len(Formula) And prnto(1) > cmmpos(2) Then 'commas in series name, x simple, y complex
            MsgBox ("This Macro doesn't work on series with commas in name (reference to cell with comma is ok though)")
        ElseIf prntc(1) = Len(Formula) And prnto(1) = cmmpos(2) Then ' x simple, y complex
            SeriesAllRanges() = Split(Formula, ",(")
            If Left(SeriesAllRanges(1), 1) = "{" Then 'if x is not a range but values input return nothing. Could also be y values but that should be rare
                Exit Function
            Else
                Set SrsX = Range(Mid(Formula, cmmpos(1) + 1, cmmpos(2) - cmmpos(1) - 1))
            End If
            
        ElseIf cmmpos(1) = prnto(1) Then ' x complex, y simple
            Set SrsX = Range(Mid(Formula, cmmpos(1) + 1, prntc(1) - cmmpos(1)))
        ElseIf cmmpos(1) < prnto(1) Then 'commas in series name, x complex, y simple
            MsgBox ("This Macro doesn't work on series with commas in name (reference to cell with comma is ok though)")
            
        Else
            MsgBox ("Missed a case, something is wrong with the formula!")
        End If

End Function

Function SrsY(Formula As String) As Range

        Dim pos As Long, cmmct As Long, prntct As Long
        Dim cmmpos() As Long, prnto() As Long, prntc() As Long
        Dim SeriesAllRanges() As String
        

    'count commas and parenthesis in formula
        For pos = 1 To Len(Formula)
            If Mid(Formula, pos, 1) = "," Then
                cmmct = cmmct + 1
                ReDim Preserve cmmpos(cmmct)
                cmmpos(cmmct) = pos
            End If
            If Mid(Formula, pos, 2) = ",(" Then
                prntct = prntct + 1
                ReDim Preserve prnto(prntct)
                prnto(prntct) = pos
            End If
            If Mid(Formula, pos, 1) = ")" Then
                ReDim Preserve prntc(prntct)
                prntc(prntct) = pos
            End If
        Next pos
        
        
    'Extract X values
        If cmmct > 2 And prntct = 0 Then 'commas in series name, X and Y simple
            MsgBox ("This Macro doesn't work on series with commas in name (reference to cell with comma is ok though)")
        ElseIf cmmct = 2 Then       'no complex range
            SeriesAllRanges() = Split(Formula, ",")
            Set SrsY = Range(SeriesAllRanges(2))
            
        ElseIf prntct = 2 Then      'x and y complex, commas in title don't matter
            SeriesAllRanges() = Split(Formula, ",(")
            Set SrsY = Range("(" & SeriesAllRanges(2))
            
        ElseIf prntc(1) = Len(Formula) And prnto(1) > cmmpos(2) Then 'commas in series name, x simple, y complex
            MsgBox ("This Macro doesn't work on series with commas in name (reference to cell with comma is ok though)")
        ElseIf prntc(1) = Len(Formula) And prnto(1) = cmmpos(2) Then ' x simple, y complex
            Set SrsY = Range(Right(Formula, Len(Formula) - cmmpos(2)))
        
        ElseIf cmmpos(1) = prnto(1) Then ' x complex, y simple
            Set SrsY = Range(Right(Formula, Len(Formula) - prntc(1) - 1))
        ElseIf cmmpos(1) < prnto(1) Then 'commas in series name, x complex, y simple
            MsgBox ("This Macro doesn't work on series with commas in name (reference to cell with comma is ok though)")
            
        Else
            MsgBox ("Missed a case, something is wrong with the formula!")
        End If

End Function

Function chartPositionCell(templateChart As ChartObject) As Range
    'Last modified: 20130329
    'Returns left, top cell underlying the chart
    
    chartTop = templateChart.Top
    chartLeft = templateChart.Left
    For checkTop = 1 To 65000
        cellTop = Range("A" & checkTop).Top
        If (cellTop - chartTop) > 0.1 Then
            chartRow = checkTop - 1
            Exit For
        End If
    Next checkTop
    
    For checkLeft = 1 To 65000
        cellLeft = Cells(1, checkLeft).Left
        If (cellLeft - chartLeft) > 0 Then
            chartColumn = checkLeft - 1
            Exit For
        End If
    Next checkLeft
    
    Set chartPositionCell = Cells(chartRow, chartColumn)
End Function

Function addChart(chartAreaSelection As Range, Optional chtType As xlchartType = 65, _
Optional chtTitle As String, Optional yAxisTitle As String, Optional xAxisTitle As String) As Chart

    Dim newChtObj As ChartObject
    
    Set newChtObj = ActiveSheet.ChartObjects.Add(Top:=10, Left:=10, Height:=10, Width:=10)
    newChtObj.Activate
    
    With newChtObj
        .Left = chartAreaSelection.Left
        .Top = chartAreaSelection.Top
        .Width = chartAreaSelection.Width
        .Height = chartAreaSelection.Height
    End With
    
    Call addSeries(ActiveChart, Range("A4"), Range("A4")) 'One series has to be added to be able to add some elements, even if it's deleted later
    newChtObj.Activate
    
    With ActiveChart
        .ChartType = chtType
        
        If Not chtTitle = "" Then
            ActiveChart.SetElement (msoElementChartTitleAboveChart)
            On Error Resume Next
                dumpVariable = Range(chtTitle) ' check whether chtTitle describes a range
                If Err.Number = 0 Then
                    .ChartTitle.Text = "=" & chtTitle
                ElseIf Err.Number <> 0 Then
                    .ChartTitle.Text = chtTitle
                End If
            On Error GoTo 0
            .ChartTitle.Font.Size = 14
        End If
        
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue, xlPrimary).AxisTitle.Text = yAxisTitle
        .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = xAxisTitle
        .DisplayBlanksAs = xlInterpolated
    End With
    
    ActiveChart.SeriesCollection(1).Delete
    Set addChart = ActiveChart

End Function

Sub addSeries(cht As Chart, XValueRange As Range, YValueRange As Range, Optional seriesName As String, _
 Optional seriesColor As XlRgbColor = 19, Optional markerShape As XlMarkerStyle, Optional yError As Range)
   
    cht.SeriesCollection.NewSeries
    lastSeries = cht.SeriesCollection.Count
    
    With cht.SeriesCollection(lastSeries)
        .XValues = XValueRange
        .Values = YValueRange
        
        If Not yError Is Nothing Then
            .ErrorBar Direction:=xlY, Include:=xlBoth, Type:=xlCustom, MinusValues:=yError, Amount:=yError
        End If
        
        If seriesColor <> 19 Then
            On Error Resume Next
            .MarkerBackgroundColorIndex = seriesColor
            .MarkerForegroundColorIndex = seriesColor
            .Border.ColorIndex = seriesColor
            .Format.Fill.ForeColor.RGB = seriesColor
            On Error GoTo 0
        End If
        
        If markerShape <> 0 Then
            .MarkerStyle = markerShape
        End If
        
        If Not seriesName = "" Then 'check if seriesName can be translated into Range, if yes use value in that range
            On Error Resume Next
                .name = Range(seriesName).Value
                If Err.Number <> 0 Then
                    .name = seriesName
                Else
                    .name = "=" & "'" & Range(seriesName).Parent.name & "'!" & seriesName
                End If
            On Error GoTo 0
        End If
    End With
        
End Sub

Sub deleteAllSeries(ChartToClear As ChartObject)
    Dim ser As Series

    For Each ser In ChartToClear.Chart.SeriesCollection
        ser.Delete
    Next ser
End Sub

Function selectAll(ListToMark As ListBox)
    For a = 0 To ListToMark.ListCount - 1
        ListToMark.Selected(a) = True
    Next a
End Function









