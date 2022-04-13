Option Explicit
Public Const yScaleStr As String = "0.01,0.1,1,3,5,10,30,50,70,90,95,97,99,99.9,99.99"

Public Function PlotCumulativeChart(vSheetName As String)
    Dim nowSheet As Worksheet
    Dim i As Long
    Dim yScaleA
    Dim vChartInfo As chartInfo
    Dim nowRange As Range
    Dim nowChart As Chart
    Dim iCol As Long
    Dim nowAxis As Axis
    Dim nowSeries As Series
    Dim nowDL As DataLabel
    Dim xMax, xMin, yMax, yMin
    
    On Error Resume Next
    
    If IsExistSheet(vSheetName) Then
        Set nowSheet = Worksheets(vSheetName)
    Else
        Exit Function
    End If
    
    ' Get the OBC attribute of the chart
    '-----------------------------------------------------------------------
    Set nowRange = nowSheet.Range("A1:B1").CurrentRegion
    vChartInfo = getChartInfo(nowRange)
    
    ' Add Chart
    '-----------------------------------------------------------------------
    Set nowChart = nowSheet.ChartObjects.Add(30, 30, 450, 300).Chart
    nowChart.chartType = xlXYScatter
    
    ' Add Y Scale Series
    '-----------------------------------------------------------------------
    yScaleA = Split(yScaleStr, ",")
    iCol = 3
    With nowChart.SeriesCollection.NewSeries
        .XValues = "=" & nowSheet.Name & "!" & "R2C" & CStr(iCol) & ":R" & CStr(2 + UBound(yScaleA)) & "C" & CStr(iCol)
        .Values = "=" & nowSheet.Name & "!" & "R2C" & CStr(iCol + 1) & ":R" & CStr(2 + UBound(yScaleA)) & "C" & CStr(iCol + 1)
        .Name = "=" & nowSheet.Name & "!" & "R1C" & CStr(iCol)
    End With
    
    ' Add Data Series
    '-----------------------------------------------------------------------
    For iCol = 5 To nowSheet.UsedRange.Columns.Count Step 2
        If nowSheet.Cells(1, iCol) = "" Then Exit For
        With nowChart.SeriesCollection.NewSeries
            .XValues = "=" & nowSheet.Name & "!" & "R2C" & CStr(iCol) & ":R" & CStr(nowSheet.UsedRange.Rows.Count) & "C" & CStr(iCol)
            .Values = "=" & nowSheet.Name & "!" & "R2C" & CStr(iCol + 1) & ":R" & CStr(nowSheet.UsedRange.Rows.Count) & "C" & CStr(iCol + 1)
            .Name = "=" & nowSheet.Name & "!" & "R1C" & CStr(iCol)
        End With
    Next iCol
    
    ' Set X Axis
    '-----------------------------------------------------------------------
    Set nowAxis = nowChart.Axes(xlCategory)
    With nowAxis
        .MinimumScale = nowSheet.Cells(2, 3)
        .HasTitle = True
        .AxisTitle.Text = vChartInfo.xLabel
    End With
    
    ' Set Y Axis
    '-----------------------------------------------------------------------
    Set nowAxis = nowChart.Axes(xlValue)
    With nowAxis
        .MinimumScale = nowSheet.Cells(2, 4)
        .MaximumScale = nowSheet.Cells(2 + UBound(yScaleA), 4)
        .CrossesAt = nowSheet.Cells(2, 4)
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNone
        .MajorUnit = 20
        '.MajorGridlines.Border.LineStyle = xlNone
    End With
    
    ' Set X Axis
    '-----------------------------------------------------------------------
    Call getScaterMaxMin(nowChart, xMax, xMin, yMax, yMin)
    ' Set X Axis
    '-----------------------------------------------------------------------
    Call AxisScaleFit(nowChart.Axes(xlCategory), xMax, xMin, vChartInfo.xMax, vChartInfo.xMin, UCase(vChartInfo.XScale) = "LOG")
    ' Set X Axis
    '-----------------------------------------------------------------------
    Call SetAxisScaleType(nowChart, UCase(vChartInfo.XScale) = "LOG", False, vChartInfo)

    ' Set YScale series and datalabel
    '-----------------------------------------------------------------------
    For i = 0 To UBound(yScaleA)
        'nowSeries.Points(i + 1).DataLabel.Text = yScaleA(i)
        nowSheet.Cells(i + 2, 3) = nowChart.Axes(xlCategory).MinimumScale
    Next i
    Set nowSeries = nowChart.SeriesCollection("YSCALE")  '.ApplyDataLabels Type:=xlDataLabelsShowLabel
    nowSeries.ApplyDataLabels
    With nowSeries.DataLabels
        .Position = xlLabelPositionLeft
        '.HorizontalAlignment = xlRight
        .Font.Size = 8
    End With
    For i = 0 To UBound(yScaleA)
        nowSeries.Points(i + 1).DataLabel.Text = yScaleA(i)
        'nowSheet.Cells(i + 2, 3) = nowChart.Axes(xlCategory).MinimumScale
    Next i

    ' Add Target
    '----------------------------------------------------------------------
    Call plotCumulativeTarget(nowChart, vChartInfo)

    ' Set chart title
    '-----------------------------------------------------------------------
    nowChart.HasTitle = True
    nowChart.ChartTitle.Text = vChartInfo.ChartTitle
    nowChart.Legend.LegendEntries(1).Delete
    ' Set series style
    '-----------------------------------------------------------------------
    Call SetChartStyle(nowChart, "CUMULATIVE")
    DoEvents
    
    Set nowAxis = Nothing
    Set nowSheet = Nothing
End Function

Public Function plotCumulativeTarget(nowChart As Chart, vChartInfo As chartInfo)
    Dim nowSeries As Series
    Dim tempA
    Dim i As Integer
    Dim tmp As String
        
    If vChartInfo.vTargetXValueStr = "" Then Exit Function
    tempA = Split(vChartInfo.vTargetXValueStr, ",")
    tmp = "-10"
    For i = 1 To UBound(tempA)
        tmp = tmp & "," & "-10"
    Next i
    
    'For i = 0 To UBound(TempA)
        Set nowSeries = nowChart.SeriesCollection.NewSeries
        nowSeries.XValues = "={" & vChartInfo.vTargetXValueStr & "}"
        nowSeries.Values = "={" & tmp & "}"
        nowSeries.Name = "Target"
    'Next i
End Function
