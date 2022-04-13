'************************************************************
'*Title: PlotUniversalChart()
'*-----------------------------------------------------------
'*Notes: This program plot Universal chart.
'*
'*-----------------------------------------------------------
'*Include files:  OBC file and wat raw data
'*Output file: Universal chart
'*-----------------------------------------------------------
'*Date: 10/23/2004 kunshin chou - Initial code
'*2012/12/04    Rewrite for Excel 2010
'*************************************************************
Sub PlotUniversalChart(vSheetName As String)

Dim tempSheetName As String
Dim tmpChartName As String
Dim tmpParameter1 As String
Dim tmpParameter2 As String
Dim tmpCurChartName As String

Dim curChart As Chart

Dim sTitleInfo As String
Dim iNoteStart As Long
Dim iNoteEnd As Long

Dim SerialCount As Integer
Dim dataRows As Integer
Dim dataCols As Integer
      
Dim lLeft As Long
Dim lRight As Long
Dim lUpper As Long
Dim lBottom As Long


Dim curSerial As Long
Dim curXColpos As Long
Dim curYColpos As Long
Dim curSerialColpos As Long

Dim tmpStartRow As Long
Dim tmpEndRow As Long

Dim iSerialColStart As Long
Dim iSerialColEnd As Long
Dim iSerialRowStart As Long
Dim iSerialRowEnd As Long
Dim cSerialName As String

Dim tmpMax As Double
Dim tmpMin As Double
Dim curXMax As Double
Dim curXMin As Double
Dim curYMax As Double
Dim curYMin As Double
    
Dim XaxisMax As Double
Dim YaxisMax As Double
   
Dim vChartInfo As chartInfo
Dim vRange As Range
Dim ChartTitle As String
Dim xLabel As String
Dim yLabel As String

Dim vblnXLog As Boolean
Dim vblnYLog As Boolean

'------
Dim nowSheet As Worksheet, nowChart As Chart, nowSeries As Series, nowAxis As Axis
Dim i As Long
Dim j As Integer
Dim xRange As Range, yRange As Range
Dim varXMax, varXMin, varYMax, varYMin
Dim titleRange As Range

    On Error Resume Next

    Set nowSheet = Worksheets(vSheetName)
    nowSheet.Activate
    
    'Get Range
    Call GetRange(nowSheet.UsedRange, lLeft, lUpper, lRight, lBottom)
    iSerialColStart = 3
    iSerialColEnd = lRight
    iSerialRowStart = 3
    iSerialRowEnd = lBottom
    If iSerialColStart > iSerialColEnd Then Exit Sub
    
    
    ' Get the OBC attribute of the chart
    '------------------------------------------------------------------------------------------------------------------------------------
    Set vRange = nowSheet.Range(nowSheet.Cells(1, 1), nowSheet.Cells(1, 1).End(xlDown).Offset(0, 1))
    vChartInfo = getChartInfo(vRange)
    ChartTitle = vChartInfo.ChartTitle
    xLabel = vChartInfo.xLabel
    yLabel = vChartInfo.yLabel
    vblnXLog = IsKey(vChartInfo.XScale, "Log")
    vblnYLog = IsKey(vChartInfo.YScale, "Log")
        
    ' expand SS FF TT
    '------------------------------------------------------------------------------------------------------------------------------------
    Call expandCurveData(nowSheet, "SS", vChartInfo.vSS)
    Call expandCurveData(nowSheet, "FF", vChartInfo.vFF)
    Call expandCurveData(nowSheet, "TT", vChartInfo.vTT)
    Call expandCurveData(nowSheet, "GOLDENDIE", vChartInfo.vGoldendie)
    
    For iSerialRowEnd = iSerialRowEnd To 1 Step -1
        For i = 3 To nowSheet.Columns.Count
            If nowSheet.Cells(iSerialRowEnd, i) <> "" Then GoTo ExitForLabel
        Next i
    Next iSerialRowEnd
ExitForLabel:
    
    ' expand target
    '------------------------------------------------------------------------------------------------------------------------------------
    'Array chart don't append target
    If Left(UCase(nowSheet.Name), 8) <> "SCATTERA" Then _
        Call ExpandTargetData(vChartInfo.vTargetNameStr, vChartInfo.vTargetXValueStr, vChartInfo.vTargetYValueStr, iSerialRowStart, nowSheet.UsedRange.Columns.Count + 1)
    
    'Set nowChart = nowSheet.ChartObjects.Add(10, 10, 400, 300).Chart
    'For 2010 相容
    Set nowChart = myCreateChart(nowSheet, xlXYScatter, 10, 10, 400, 300)
    
    SerialCount = ((nowSheet.UsedRange.Columns.Count - iSerialColStart + 1) / 2)
    For curSerial = 0 To SerialCount - 1
        cSerialName = nowSheet.Cells(1, 3 + curSerial * 2)
        Set titleRange = nowSheet.Cells(1, 3 + curSerial * 2)
        Set xRange = nowSheet.Range(N2L(3 + curSerial * 2) & "3:" & N2L(3 + curSerial * 2) & CStr(iSerialRowEnd))
        Set yRange = nowSheet.Range(N2L(4 + curSerial * 2) & "3:" & N2L(4 + curSerial * 2) & CStr(iSerialRowEnd))
        Set nowSeries = nowChart.SeriesCollection.NewSeries
        With nowSeries
            .XValues = xRange
            .Values = yRange
            '.Name = cSerialName
            .Name = "=" & vSheetName & "!R" & titleRange.row & "C" & titleRange.Column
        End With
        If curSerial = 0 Then
            varXMax = WorksheetFunction.Max(xRange)
            varXMin = WorksheetFunction.Min(xRange)
            varYMax = WorksheetFunction.Max(yRange)
            varYMin = WorksheetFunction.Min(yRange)
        Else
            If varXMax < WorksheetFunction.Max(xRange) Then varXMax = WorksheetFunction.Max(xRange)
            If varXMin > WorksheetFunction.Min(xRange) Then varXMin = WorksheetFunction.Min(xRange)
            If varYMax < WorksheetFunction.Max(yRange) Then varYMax = WorksheetFunction.Max(yRange)
            If varYMin > WorksheetFunction.Min(yRange) Then varYMin = WorksheetFunction.Min(yRange)
        End If
    Next curSerial
    
    With nowChart
        .chartType = xlXYScatter
        .HasTitle = True
        .ChartTitle.Characters.Text = ChartTitle
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = xLabel
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = yLabel
    End With
    
    '------------------------------------------------------------------------------------------------------------------------------------
    ' End : Add the Series on the chart
    '------------------------------------------------------------------------------------------------------------------------------------

    ' plot Conner
    '------------------------------------------------------------------------------------------------------------------------------------
    curSerialColpos = iSerialColStart + curSerial * 2
    curXColpos = curSerialColpos
    curYColpos = curSerialColpos + 1
    Call ExpandConnerData(vChartInfo.vCornerXValueStr, iSerialRowStart, curXColpos)
    Call ExpandConnerData(vChartInfo.vCornerYValueStr, iSerialRowStart, curYColpos)

    Set vRange = Range(ActiveSheet.Cells(iSerialRowStart - 1, curXColpos), ActiveSheet.Cells(iSerialRowEnd, curYColpos))


    If vChartInfo.vCornerXValueStr <> "" And vChartInfo.vCornerYValueStr <> "" Then
        Call PlotConner(nowSheet.Name, vRange, nowChart)
    End If
    
    '------------------------------------------------------------------------------------------------------------------------------------
    Call SetChartStyle(nowChart)
    
    ' set scater's Max,Min
    '------------------------------------------------------------------------------------------------------------------------------------
    
    Dim CornerXMax As Double
    Dim CornerXMin As Double
    Dim CornerYMax As Double
    Dim CornerYMin As Double
    If nowSheet.Cells(1, nowSheet.UsedRange.Columns.Count - 1) = "Corner" Then
        CornerXMax = WorksheetFunction.Max(nowSheet.Range(Cells(3, nowSheet.UsedRange.Columns.Count - 1), Cells(7, nowSheet.UsedRange.Columns.Count - 1)))
        CornerXMin = WorksheetFunction.Min(nowSheet.Range(Cells(3, nowSheet.UsedRange.Columns.Count - 1), Cells(7, nowSheet.UsedRange.Columns.Count - 1)))
        CornerYMax = WorksheetFunction.Max(nowSheet.Range(Cells(3, nowSheet.UsedRange.Columns.Count - 0), Cells(7, nowSheet.UsedRange.Columns.Count - 0)))
        CornerYMin = WorksheetFunction.Min(nowSheet.Range(Cells(3, nowSheet.UsedRange.Columns.Count - 0), Cells(7, nowSheet.UsedRange.Columns.Count - 0)))
        If varXMax < CornerXMax Then varXMax = CornerXMax
        If varXMin > CornerXMin Then varXMin = CornerXMin
        If varYMax < CornerYMax Then varYMax = CornerYMax
        If varYMin > CornerYMin Then varYMin = CornerYMin
    End If

    If vChartInfo.xMax <> "" Then varXMax = CDbl(vChartInfo.xMax)
    If vChartInfo.xMin <> "" Then varXMin = CDbl(vChartInfo.xMin)
    If vChartInfo.yMax <> "" Then varYMax = CDbl(vChartInfo.yMax)
    If vChartInfo.yMin <> "" Then varYMin = CDbl(vChartInfo.yMin)
    
    Dim chtSetup As Chart
    Dim PlotSetupAxsValue As Axis, PlotSetupAxsCategory As Axis
    
    Set chtSetup = Worksheets("PlotSetup").ChartObjects(1).Chart
    Set PlotSetupAxsValue = chtSetup.Axes(xlValue)
    Set PlotSetupAxsCategory = chtSetup.Axes(xlCategory)
    
    'Call getScaterMaxMin(nowChart, varXMax, varXMin, varYMax, varYMin)

    If vChartInfo.xMax <> "" Then varXMax = CDbl(vChartInfo.xMax)
    If vChartInfo.xMin <> "" Then varXMin = CDbl(vChartInfo.xMin)
    If vChartInfo.yMax <> "" Then varYMax = CDbl(vChartInfo.yMax)
    If vChartInfo.yMin <> "" Then varYMin = CDbl(vChartInfo.yMin)
    
    
    'Debug.Print varXMax, varXMin, varYMax, varYMin
    
    'Debug
    'Stop
    
    ' set Scatter's XlCategory
    '------------------------------------------------------------------------------------------------------------------------------------
    'Call SetScatterXlCategory(ActiveChart, PlotSetupAxsCategory, varXMax, varXMin)
    ' set Scatter's XlValue
    '------------------------------------------------------------------------------------------------------------------------------------
    'Call SetScatterXlValue(ActiveChart, PlotSetupAxsValue, varYMax, varYMin)
    ' set Scatter
    '------------------------------------------------------------------------------------------------------------------------------------
'    Call SetAxisCrossesAt(nowChart)
    'Fit Scale
    '------------------------------------------------------------------------------------------------------------------------------------
    AxisScaleFit nowChart.Axes(xlCategory), varXMax, varXMin, vChartInfo.xMax, vChartInfo.xMin, vblnXLog
    AxisScaleFit nowChart.Axes(xlValue), varYMax, varYMin, vChartInfo.yMax, vChartInfo.yMin, vblnYLog
    ' set Scatter's ScaleType
    '------------------------------------------------------------------------------------------------------------------------------------
    Call SetAxisScaleType(nowChart, vblnXLog, vblnYLog, vChartInfo)
    
    ' Draw Trend Lines
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim noLineItems As String
    
    For i = chtSetup.SeriesCollection.Count To chtSetup.SeriesCollection.Count - 5 Step -1
        noLineItems = noLineItems & "&" & UCase(chtSetup.SeriesCollection(i).Name)
    Next i
    noLineItems = Mid(noLineItems, 2)
    
    If Not IsEmpty(vChartInfo.vSensitivity) Or UCase(vChartInfo.GTrendLines) = "YES" Then
        
        
        For j = 1 To nowChart.FullSeriesCollection.Count
            If Not IsKey(noLineItems, nowChart.FullSeriesCollection(j).Name) Then
                With nowChart.FullSeriesCollection(j)
                    .Trendlines.Add
                    .Trendlines(1).DisplayEquation = True
                    .Trendlines(1).DisplayRSquared = True
                    .Trendlines(1).Format.Line.Weight = 1.5
                    .Trendlines(1).Format.Line.ForeColor.RGB = Worksheets("PlotSetup").ChartObjects(1).Chart.SeriesCollection(j).Points(1).MarkerBackgroundColor
                    .Trendlines(1).DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Worksheets("PlotSetup").ChartObjects(1).Chart.SeriesCollection(j).Points(1).MarkerBackgroundColor
                    If InStr(UCase(vChartInfo.YScale), "LOG") And InStr(UCase(vChartInfo.XScale), "LINEAR") Then
                        .Trendlines(1).Type = xlExponential
                    ElseIf InStr(UCase(vChartInfo.YScale), "LINEAR") And InStr(UCase(vChartInfo.XScale), "LOG") Then
                        .Trendlines(1).Type = xlLogarithmic
                    End If
                End With
            End If
        Next j
    End If
    
    If UCase(vChartInfo.ChartExpression) = "MEDIAN" Then
        Call addThruTrendLine
    End If
    
Exit Sub

'Old Code
'*******************************************************************************************************************

   'if x or y value is empty by Dio
   'If IsEmptyValue(vSheetName) Then Exit Sub
   
    tempSheetName = vSheetName
    Sheets(tempSheetName).Select
    'tempSheetName = ActiveSheet.Name
'=====================
    ' Get the serial number of data set
    '------------------------------------------------------------------------------------------------------------------------------------
   
    dataCols = lRight
    dataRows = lLeft
    
Exit Sub
    

    curSerial = 0
        
    Do While curSerial <= SerialCount
            
            
            curSerialColpos = iSerialColStart + curSerial * 2
            curXColpos = curSerialColpos
            curYColpos = curSerialColpos + 1
            ' Count the naumber of Serial
            '------------------------------------------------
            curSerial = curSerial + 1
            
            tmpStartRow = iSerialRowStart
            tmpEndRow = iSerialRowEnd
            If ActiveSheet.Range(N2L(curSerialColpos) & CStr(iSerialRowEnd)) = "" Then _
                tmpEndRow = ActiveSheet.Range(N2L(curSerialColpos) & CStr(iSerialRowEnd)).End(xlUp).row  ' By Dio
            
            cSerialName = ActiveSheet.Cells(iSerialRowStart - 2, curSerialColpos).Value
           
            tmpParameter1 = ActiveSheet.Cells(iSerialRowStart - 1, curXColpos).Value
            tmpParameter2 = ActiveSheet.Cells(iSerialRowStart - 1, curYColpos).Value
            tmpCurChartName = tmpParameter1 & "Vs" & tmpParameter2
           
            'iChartNo = iChartNo + 1
            'iChartTotalNo = iChartTotalNo + 1
            '------------------------------------------------------------------------------------------------------------------------------------
            'New Series
            '------------------------------------------------------------------------------------------------------------------------------------
            ActiveChart.SeriesCollection.NewSeries
            'Add By Dio
            On Error Resume Next
            ' Setting the Serial name
            ActiveChart.SeriesCollection(curSerial).Name = "=" & tempSheetName & "!R" & iSerialColStart - 2 & "C" & curSerialColpos
            If Worksheets(tempSheetName).Cells(tmpStartRow, curSerialColpos) = "" Then
               ' Start : Setting the Serial YValue
               ActiveChart.SeriesCollection(curSerial).Values = "=" & tempSheetName & "!R" & tmpStartRow & "C" & curYColpos & ":R" & tmpEndRow & "C" & curYColpos
               ' Start : Setting the Serial XValue
               ActiveChart.SeriesCollection(curSerial).XValues = "=" & tempSheetName & "!R" & tmpStartRow & "C" & curSerialColpos & ":R" & tmpEndRow & "C" & curSerialColpos
            Else
               ' Start : Setting the Serial XValue
               ActiveChart.SeriesCollection(curSerial).XValues = "=" & tempSheetName & "!R" & tmpStartRow & "C" & curSerialColpos & ":R" & tmpEndRow & "C" & curSerialColpos
               ' Start : Setting the Serial YValue
               ActiveChart.SeriesCollection(curSerial).Values = "=" & tempSheetName & "!R" & tmpStartRow & "C" & curYColpos & ":R" & tmpEndRow & "C" & curYColpos
            End If
            On Error Resume Next
            ActiveChart.SeriesCollection(curSerial).Select
            With Selection.Border
                .Weight = xlThin
                .LineStyle = xlLineStyleNone
            End With
                  
            'For Symbol transpancy
            With Selection
               .MarkerBackgroundColorIndex = xlNone
               .MarkerForegroundColorIndex = xlAutomatic
               .Smooth = False
               .MarkerSize = 5
               .Shadow = False
            End With
            
            '------------------------------------------------------------------------------------------------------------------------------------
            ' Start : get max & min
            '------------------------------------------------------------------------------------------------------------------------------------
            If tmpEndRow >= tmpStartRow Then
                tmpMax = Application.WorksheetFunction.Max(ActiveSheet.Range(ActiveSheet.Cells(tmpStartRow, curXColpos), ActiveSheet.Cells(tmpEndRow, curXColpos)))
                tmpMin = Application.WorksheetFunction.Min(ActiveSheet.Range(ActiveSheet.Cells(tmpStartRow, curXColpos), ActiveSheet.Cells(tmpEndRow, curXColpos)))
                '            curMaxTmp = tmpMax
                '            curMinTmp = tmpMin
                curXMax = IIf(curXMax > tmpMax, curXMax, tmpMax)
                curXMin = IIf(curXMin < tmpMin, curXMin, tmpMin)
                          
                XaxisMax = IIf(XaxisMax > Application.WorksheetFunction.Max(ActiveSheet.Range(ActiveSheet.Cells(tmpStartRow, curXColpos), ActiveSheet.Cells(tmpEndRow, curXColpos))), YaxisMax, Application.WorksheetFunction.Max(ActiveSheet.Range(ActiveSheet.Cells(tmpStartRow, curXColpos), ActiveSheet.Cells(tmpEndRow, curXColpos))))
            End If

            If tmpEndRow >= tmpStartRow Then
                tmpMax = Application.WorksheetFunction.Max(Range(ActiveSheet.Cells(tmpStartRow, curYColpos), ActiveSheet.Cells(tmpEndRow, curYColpos)))
                tmpMin = Application.WorksheetFunction.Min(Range(ActiveSheet.Cells(tmpStartRow, curYColpos), ActiveSheet.Cells(tmpEndRow, curYColpos)))
                curYMax = IIf(curYMax > tmpMax, curYMax, tmpMax)
                curYMin = IIf(curYMin < tmpMin, curYMin, tmpMin)
                YaxisMax = IIf(YaxisMax > Application.WorksheetFunction.Max(ActiveSheet.Range(ActiveSheet.Cells(tmpStartRow, curYColpos), ActiveSheet.Cells(tmpEndRow, curYColpos))), YaxisMax, Application.WorksheetFunction.Max(ActiveSheet.Range(ActiveSheet.Cells(tmpStartRow, curYColpos), ActiveSheet.Cells(tmpEndRow, curYColpos))))
            End If
            On Error GoTo 0
   
    Loop
    
    Exit Sub
  


End Sub



'************************************************************
'*Title: PlotConner()
'*-----------------------------------------------------------
'*Notes: This program for plot the corner.
'*
'*-----------------------------------------------------------
'*Include files:  dataRange of Conner
'*Output file:
'*-----------------------------------------------------------
'*Date: 10/23/2004 kunshin chou - Initial code
'*************************************************************
Sub PlotConner(vSheetName As String, vdataRange As Range, vChartObj As Chart)

Dim lLeft As Long
Dim lRight As Long
Dim lUpper As Long
Dim lBottom As Long
Dim curSerial As Series
Dim i As Long
Dim rsrsTarget
Dim vSeriesCoolection As SeriesCollection


Call GetRange(vdataRange, lLeft, lUpper, lRight, lBottom)

        

        Set curSerial = vChartObj.SeriesCollection.NewSeries
        Set vSeriesCoolection = vChartObj.SeriesCollection
        
       Worksheets(vSheetName).Cells(1, lLeft) = "Corner"
       'curSerial.Name = "Corner"
       curSerial.Name = "=" & vSheetName & "!R" & "1" & "C" & lLeft
       curSerial.XValues = "=" & vSheetName & "!R" & lUpper + 1 & "C" & lLeft & ":R" & lBottom & "C" & lLeft
       curSerial.Values = "=" & vSheetName & "!R" & lUpper + 1 & "C" & lRight & ":R" & lBottom & "C" & lRight
               
       ' print the corner line
       curSerial.chartType = xlXYScatterLines
       
        If vSeriesCoolection.Count >= 6 Then
            With curSerial.Points(2)
                .Border.Weight = xlThin
                .Border.LineStyle = xlNone
                .MarkerBackgroundColorIndex = xlAutomatic
                .MarkerForegroundColorIndex = xlAutomatic
                .MarkerStyle = xlAutomatic
                .MarkerSize = 5
                .Shadow = False
            End With
        End If

End Sub

Public Function ExpandTargetData(sName As String, sXData As String, sYData As String, iRow As Long, iCol As Long) As Integer
    Dim iXTokenCnt As Long
    Dim iYTokenCnt As Long
    Dim i As Integer
    Dim vRow As Long
    Dim vCol As Long
    Dim TotalVal As Long
    
    vRow = iRow
    vCol = iCol
    
    iXTokenCnt = TokenCount(sXData, ";")
    iYTokenCnt = TokenCount(sYData, ";")
    
    If iXTokenCnt = 0 Then Exit Function
    If iYTokenCnt = 0 Then Exit Function
    If iXTokenCnt <> iYTokenCnt Then Exit Function
    
    TotalVal = iXTokenCnt
    
    For i = 1 To TotalVal
    
        If (i Mod 2) <> 0 Then ActiveSheet.Cells(vRow - 2, vCol) = xGetToken(sName, ",", CLng(i))
        
        Call ExpandConnerData(xGetToken(sXData, ";", CLng(i)), vRow, vCol)
        Call ExpandConnerData(xGetToken(sYData, ";", CLng(i)), vRow, vCol + 1)
        vCol = vCol + 1
    Next i
    
    ExpandTargetData = iTokenCnt
    
End Function


Public Function ExpandConnerData(sData As String, iRow As Long, iCol As Long) As Integer
    Dim iTokenCnt As Integer
    Dim i As Integer
    
    iTokenCnt = TokenCount(sData, ",")
    'TotalVal = iTokenCnt
    For i = 1 To iTokenCnt
        ActiveSheet.Cells(iRow + i - 1, iCol) = xGetToken(sData, ",", CLng(i))
    Next i
    ExpandConnerData = iTokenCnt
    
End Function

Sub expandCurveData(curSheet As Worksheet, vName As String, vData As Variant)

Dim i As Long
Dim curRow As Long
Dim col As Long
Dim row As Long

Dim lLeft As Long
Dim lRight As Long
Dim lUpper As Long
Dim lBottom As Long

Dim myArr As Variant
Dim myTmpArr()


If IsEmpty(vData) Then Exit Sub

row = UBound(vData, 1)
col = UBound(vData, 2)


ReDim myTmpArr(1 To row, 1 To col)
myArr = vData
curRow = 1
For i = 1 To row

    If myArr(i, 2) <> "" Then
        myTmpArr(curRow, 1) = myArr(i, 1)
        myTmpArr(curRow, 2) = myArr(i, 2)
        curRow = curRow + 1
    End If
    
Next i

If row <> 0 Then
        Call GetRange(curSheet.UsedRange, lLeft, lUpper, lRight, lBottom)
        
        iSerialColStart = 3
        iSerialColEnd = lRight
        iSerialRowStart = 3
        iSerialRowEnd = lBottom
            
        curSheet.Cells(1, iSerialColEnd + 1).Value = vName
        curSheet.Cells(1, iSerialColEnd + 2).Value = "'"
        curSheet.Range(curSheet.Cells(iSerialRowStart, iSerialColEnd + 1), curSheet.Cells(iSerialRowStart + row - 1, iSerialColEnd + 2)) = myTmpArr
        
End If






End Sub

Sub PlotCurve(curSheet As Worksheet, vChartObj As Chart, vName As String, vData As Variant)

    Dim col As Long
    Dim row As Long
    
    Dim lLeft As Long
    Dim lRight As Long
    Dim lUpper As Long
    Dim lBottom As Long
    
    Dim iSerialColStart As Long
    Dim iSerialColEnd As Long
    Dim iSerialRowStart As Long
    Dim iSerialRowEnd As Long
    Dim cSerialName As String
    Dim curSerie As Series
    
        row = UBound(vData, 1)
        col = UBound(vData, 2)
        
        Call GetRange(curSheet.UsedRange, lLeft, lUpper, lRight, lBottom)
        iSerialColStart = 3
        iSerialColEnd = lRight
        iSerialRowStart = 3
        iSerialRowEnd = lBottom
            
        curSheet.Cells(1, iSerialColEnd + 1).Value = vName
        curSheet.Cells(1, iSerialColEnd + 2).Value = "'"
        curSheet.Range(curSheet.Cells(iSerialRowStart, iSerialColEnd + 1), curSheet.Cells(iSerialRowStart + row - 1, iSerialColEnd + 2)) = vData
        curSheet.Range(curSheet.Cells(iSerialRowStart, iSerialColEnd + 1), curSheet.Cells(iSerialRowStart, iSerialColEnd + 2)).Select
        
        'New Series
        '------------------------------------------------------------------------------------------------------------------------------------
        Set curSerie = vChartObj.SeriesCollection.NewSeries
        
        ' Setting the Serial name
        curSerie.Name = curSheet.Cells(1, iSerialColEnd + 1)
        ' Start : Setting the Serial XValue
        curSerie.XValues = curSheet.Range(curSheet.Cells(iSerialRowStart, iSerialColEnd + 1), curSheet.Cells(iSerialRowStart + row - 1, iSerialColEnd + 1))
        ' Start : Setting the Serial YValue
        curSerie.Values = curSheet.Range(curSheet.Cells(iSerialRowStart, iSerialColEnd + 2), curSheet.Cells(iSerialRowStart + row - 1, iSerialColEnd + 2))
        
        'curSerie.Select
        With curSerie.Border
            .Weight = xlThin
            .LineStyle = xlLineStyleNone
        End With
              
        'For Symbol transpancy
        With curSerie
           .MarkerBackgroundColorIndex = xlNone
           .MarkerForegroundColorIndex = xlAutomatic
           .Smooth = False
           .MarkerSize = 5
           .Shadow = False
        End With

End Sub

Sub SetScatterXlCategory(ByRef chtXYScatter As Chart, PlotSetupAxsCategory As Axis, varXMax, varXMin)

        With chtXYScatter.Axes(xlCategory)
            
           
            .MajorUnitIsAuto = True
            .MinorUnitIsAuto = True
            .MaximumScaleIsAuto = True
            .MinimumScaleIsAuto = True
            .ReversePlotOrder = False
            .DisplayUnit = xlNone
            
            If PlotSetupAxsCategory.HasMajorGridlines Then
                .HasMajorGridlines = True
                .MajorGridlines.Border.ColorIndex = PlotSetupAxsCategory.MajorGridlines.Border.ColorIndex
                .MajorGridlines.Border.Weight = PlotSetupAxsCategory.MajorGridlines.Border.Weight
            Else
                .HasMajorGridlines = False
            End If
            
            If PlotSetupAxsCategory.HasMinorGridlines Then
                .HasMinorGridlines = True
                .MinorGridlines.Border.ColorIndex = PlotSetupAxsCategory.MinorGridlines.Border.ColorIndex
                .MinorGridlines.Border.Weight = PlotSetupAxsCategory.MinorGridlines.Border.Weight
            Else
                .HasMinorGridlines = False
            End If
            
            .TickLabels.Font.FontStyle = PlotSetupAxsCategory.TickLabels.Font.FontStyle
            .TickLabels.Font.Size = PlotSetupAxsCategory.TickLabels.Font.Size
            .TickLabels.AutoScaleFont = False
            If PlotSetupAxsCategory.HasTitle Then
                .HasTitle = True
                With .AxisTitle
                    '.Caption = PlotSetupAxsCategory.AxisTitle.Caption
                    .Font.FontStyle = PlotSetupAxsCategory.AxisTitle.Font.FontStyle
                    .Font.Size = PlotSetupAxsCategory.AxisTitle.Font.Size
                    .AutoScaleFont = False
                End With
            End If
        End With
        
        'AxisScaleFit chtXYScatter.Axes(xlCategory), varXMax, varXMin


End Sub

Sub SetScatterXlValue(chtXYScatter As Chart, PlotSetupAxsValue As Axis, varYMax, varYMin)

        With chtXYScatter.Axes(xlValue)
            
            .MajorUnitIsAuto = True
            .MinorUnitIsAuto = True
            .MaximumScaleIsAuto = True
            .MinimumScaleIsAuto = True
            .ReversePlotOrder = False
            .DisplayUnit = xlNone
            
            If PlotSetupAxsValue.HasMajorGridlines Then
                .HasMajorGridlines = True
                .MajorGridlines.Border.ColorIndex = PlotSetupAxsValue.MajorGridlines.Border.ColorIndex
                .MajorGridlines.Border.Weight = PlotSetupAxsValue.MajorGridlines.Border.Weight
            Else
                .HasMajorGridlines = False
            End If
            
            If PlotSetupAxsValue.HasMinorGridlines Then
                .HasMinorGridlines = True
                .MinorGridlines.Border.ColorIndex = PlotSetupAxsValue.MinorGridlines.Border.ColorIndex
                .MinorGridlines.Border.Weight = PlotSetupAxsValue.MinorGridlines.Border.Weight
            Else
                .HasMinorGridlines = False
            End If
            
            .TickLabels.Font.FontStyle = PlotSetupAxsValue.TickLabels.Font.FontStyle
            .TickLabels.Font.Size = PlotSetupAxsValue.TickLabels.Font.Size
            .TickLabels.AutoScaleFont = False
            
            If PlotSetupAxsValue.HasTitle Then
                .HasTitle = True
                With .AxisTitle
                    '.Caption = PlotSetupAxsValue.AxisTitle.Caption
                    .Font.FontStyle = PlotSetupAxsValue.AxisTitle.Font.FontStyle
                    .Font.Size = PlotSetupAxsValue.AxisTitle.Font.Size
                    .AutoScaleFont = False
                End With
                
            End If
            
        End With
        
        'AxisScaleFit chtXYScatter.Axes(xlValue), varYMax, varYMin

End Sub
Sub SetAxisScaleType(chtXYScatter As Chart, vblnXLog As Boolean, vblnYLog As Boolean, mChartInfo As chartInfo)
    Dim axsValue As Axis, axsCategory As Axis
    Dim varXMin, varYMin
    
    Set axsCategory = chtXYScatter.Axes(xlCategory)
    Set axsValue = chtXYScatter.Axes(xlValue)
    
    Call getScaterMaxMin(ActiveChart, varXMax, varXMin, varYMax, varYMin)
    '        AxisScaleFit axsCategory, varXMax, varXMin, vblnXLog
    '        AxisScaleFit axsValue, varYMax, varYMin, vblnYLog
    
    varXMin = axsCategory.MinimumScale
    varYMin = axsValue.MinimumScale
    
    '        varXMin = IIf(varXMin + axsCategory.MajorUnit = 0, 0.005, varXMin)
    '        varYMin = IIf(varYMin + axsValue.MajorUnit = 0, 0.005, varYMin)
    
    '        axsCategory.MinimumScale = varXMin
    '        axsValue.MinimumScale = varYMin
    '        If varXMin = 0 Then varXMin = varXMin + axsCategory.MajorUnit
    '        If varYMin = 0 Then varYMin = varYMin + axsValue.MajorUnit
    
    If vblnXLog And axsCategory.MinimumScale > 0 Then
        
    '                If varXMin > varYMin Then
    '                    axsCategory.Crosses = xlCustom
    '                    axsCategory.CrossesAt = varYMin
    '                    axsValue.Crosses = xlCustom
    '                    axsValue.CrossesAt = varYMin
    '                Else
    '                    axsCategory.Crosses = xlCustom
    '                    axsCategory.CrossesAt = varXMin
    '                    axsValue.Crosses = xlCustom
    '                    axsValue.CrossesAt = varYMin
    '                End If
        
        axsCategory.ScaleType = xlLogarithmic
        
    Else
        axsCategory.ScaleType = xlLinear
    End If
    
    If vblnYLog And axsValue.MinimumScale > 0 Then
    '                If varXMin > varYMin Then
    '                    axsCategory.Crosses = xlCustom
    '                    axsCategory.CrossesAt = varYMin
    '                    axsValue.Crosses = xlCustom
    '                    axsValue.CrossesAt = varYMin
    '                Else
    '                    axsCategory.Crosses = xlCustom
    '                    axsCategory.CrossesAt = varXMin
    '                    axsValue.Crosses = xlCustom
    '                    axsValue.CrossesAt = varYMin
    '                End If
        axsValue.ScaleType = xlLogarithmic
    Else
        axsValue.ScaleType = xlLinear
    End If
        
    ' Reverse order
    If IsKey(mChartInfo.XScale, "Reverse") Then
        axsCategory.Crosses = xlMaximum
        axsCategory.ReversePlotOrder = True
    Else
        axsCategory.Crosses = xlCustom
       axsCategory.CrossesAt = axsCategory.MinimumScale
    End If
    If IsKey(mChartInfo.YScale, "Reverse") Then
        axsValue.Crosses = xlMaximum
        axsValue.ReversePlotOrder = True
    Else
        axsValue.Crosses = xlCustom
       axsValue.CrossesAt = axsValue.MinimumScale
        'axsValue.Crosses = xlAxisCrossesMinimum
    End If
    
        
End Sub

Sub SetAxisCrossesAt(chtXYScatter As Chart)

        Dim axsValue As Axis, axsCategory As Axis
        Dim varXMin, varYMin
        
        Set axsCategory = chtXYScatter.Axes(xlCategory)
        Set axsValue = chtXYScatter.Axes(xlValue)
        
        varXMin = axsCategory.MinimumScale
        varYMin = axsValue.MinimumScale
        
        If varXMin > varYMin Then
            axsCategory.Crosses = xlCustom
            axsCategory.CrossesAt = varYMin
            axsValue.Crosses = xlCustom
            axsValue.CrossesAt = varYMin
        Else
            axsCategory.Crosses = xlCustom
            axsCategory.CrossesAt = varXMin
            axsValue.Crosses = xlCustom
            axsValue.CrossesAt = varYMin
            
        End If
        
        
End Sub



Sub getScaterMaxMin(ByRef chtXYScatter As Chart, ByRef varXMax, ByRef varXMin, ByRef varYMax, ByRef varYMin)

    Dim srsNew As Series
    ' Add by Dio
    On Error Resume Next
    For Each srsNew In chtXYScatter.SeriesCollection
        With srsNew
            Select Case UCase(srsNew.Name)
                Case "YSCALE"
                Case Else
                    If IsEmpty(varXMax) Then
                        varXMax = WorksheetFunction.Max(.XValues)
                        varXMin = WorksheetFunction.Min(.XValues)
                        varYMax = WorksheetFunction.Max(.Values)
                        varYMin = WorksheetFunction.Min(.Values)
                    Else
                        If varXMax < WorksheetFunction.Max(.XValues) Then varXMax = WorksheetFunction.Max(.XValues)
                        If varXMin > WorksheetFunction.Min(.XValues) Then varXMin = WorksheetFunction.Min(.XValues)
                        If varYMax < WorksheetFunction.Max(.Values) Then varYMax = WorksheetFunction.Max(.Values)
                        If varYMin > WorksheetFunction.Min(.Values) Then varYMin = WorksheetFunction.Min(.Values)
                    End If
            End Select
        End With
     Next

End Sub


Public Sub AxisScaleFit(ByRef rAxis As Axis, ByVal vvarMax, ByVal vvarMin, ByVal cMax, ByVal cMin, varLog As Boolean)

    Dim varMax, varMin
    Dim dblMajorUnit As Double
    Dim intRender As Integer
    Dim RoundNum As Integer
    Dim i As Integer
    
    With rAxis
'        Do
'            intRender = intRender + 1
'
'            varMax = .MaximumScale
'            dblMajorUnit = .MajorUnit
'
'
'            Do
'                varMax = varMax - dblMajorUnit
'                If varMax <= vvarMax Then
'                    .MaximumScale = varMax + dblMajorUnit
'                    Exit Do
'                End If
'            Loop
'
'            varMin = .MinimumScale
'            dblMajorUnit = .MajorUnit
'
'            Do
'                varMin = varMin + dblMajorUnit
'                If varMin >= vvarMin Then
'                    .MinimumScale = varMin - dblMajorUnit
'                    Exit Do
'                End If
'            Loop
'
'            If intRender = 5 Then Exit Do
'
'        Loop Until dblMajorUnit = .MajorUnit
        
        '.MaximumScale = IIf(.MaximumScale > vvarMax, vvarMax + .MajorUnit, vvarMax + .MajorUnit)
        '.MinimumScale = IIf(.MinimumScale > vvarMin, vvarMin - .MajorUnit, vvarMin - .MajorUnit)
        '.MaximumScale = (vvarMax / .MajorUnit + 1) * .MajorUnit
        '.MinimumScale = (vvarMin / .MajorUnit - 1) * .MajorUnit
        '.MinimumScale = IIf(.MinimumScale = 0, 0.005, .MinimumScale)
      'Debug.Print .MaximumScale, vvarMax
      'By Dio
      .MaximumScale = vvarMax
      .MinimumScale = vvarMin
      For i = 1 To 8
         If Int(vvarMax / .MajorUnit) = (vvarMax / .MajorUnit) Then
            .MaximumScale = (Int(vvarMax / .MajorUnit)) * .MajorUnit
         Else
            .MaximumScale = (Int(vvarMax / .MajorUnit) + 1) * .MajorUnit
         End If
         .MinimumScale = (Int(vvarMin / .MajorUnit)) * .MajorUnit
      Next i
      If varLog Then
         If vvarMin > 0 Then
            RoundNum = Int(Log10(vvarMin))
            .MinimumScale = 1 * 10 ^ RoundNum
         End If
         If cMin <> "" Then If cMin > 0 Then .MinimumScale = cMin
         If cMax <> "" Then If cMax > 0 Then .MaximumScale = cMax
      Else
         'With User define
         If cMin <> "" Then .MinimumScale = cMin
         If cMax <> "" Then .MaximumScale = cMax
      End If
      '.Crosses = xlAxisCrossesMinimum
      .Crosses = xlCustom
      .CrossesAt = .MinimumScale
      'Debug.Print .MinimumScale
    End With
    
End Sub

Private Sub SetPointsStyle(ByRef rsrsTarget As Series, ByVal vstrStyle, ByVal vintSeries As Integer, Optional ByVal cType As String)

On Error Resume Next

'xlMarkerStyleNone 無記號
'xlMarkerStyleAutomatic 自動設定記號
'xlMarkerStyleSquare 方形記號
'xlMarkerStyleDiamond 菱形記號
'xlMarkerStyleTriangle 三角形記號
'xlMarkerStyleX 帶 X 記號的方形記號
'xlMarkerStyleStar 帶星號的方形記號
'xlMarkerStyleDot 短橫條形記號
'xlMarkerStyleDash 長橫條形記號
'xlMarkerStyleCircle 圓形記號
'xlMarkerStylePlus 帶加號的方形記號
    
    Dim srsCol As SeriesCollection
    Dim srsSystem As Series
    Dim ynErrBar As Boolean
    
    Set srsCol = Worksheets("PlotSetup").ChartObjects(1).Chart.SeriesCollection
    ynErrBar = False
    
    If cType = "CUMULATIVE" Then vintSeries = vintSeries - 1
    
    With rsrsTarget
        Select Case UCase(vstrStyle)
        
        Case "TARGET"
            Set srsSystem = srsCol("TARGET")
            ynErrBar = True
        Case "CORNER"
            Set srsSystem = srsCol("CORNER")
         
        Case "SS"
            Set srsSystem = srsCol("SS")
            
        Case "FF"
            Set srsSystem = srsCol("FF")
            
        Case "TT"
            Set srsSystem = srsCol("TT")
        Case "USL"
            Set srsSystem = srsCol("USL")
            ynErrBar = True
        Case "LSL"
            Set srsSystem = srsCol("LSL")
            ynErrBar = True
        Case "GOLDEN"
            Set srsSystem = srsCol("GOLDENDIE")
            
        Case "MEDIAN"
            Set srsSystem = srsCol("MEDIAN")
        Case "YSCALE"
            .MarkerBackgroundColorIndex = 1
            .MarkerForegroundColorIndex = 1
            .MarkerStyle = xlDot
        Case Else
            If vintSeries Mod 25 = 0 Then
                Set srsSystem = srsCol(25)
            Else
                Set srsSystem = srsCol(vintSeries Mod 25)
                'Set srsSystem = srsCol(vintSeries)
            End If
            If InStr(UCase(vstrStyle), "MEDIAN") > 0 Then Set srsSystem = srsCol("MEDIAN-2")
        End Select
        
        If IsEmpty(srsSystem) Then Exit Sub

        .MarkerForegroundColorIndex = srsSystem.MarkerForegroundColorIndex
        .MarkerBackgroundColorIndex = srsSystem.MarkerBackgroundColorIndex
        .MarkerSize = srsSystem.MarkerSize
        .MarkerStyle = srsSystem.MarkerStyle 'xlMarkerStyleAutomatic
        
        .Border.LineStyle = srsSystem.Border.LineStyle
        If .Border.LineStyle <> xlNone Then
            .Border.ColorIndex = srsSystem.Border.ColorIndex
            .Border.Weight = srsSystem.Border.Weight
            If ynErrBar Then
               .ErrorBar Direction:=xlX, Include:=xlBoth, Type:=xlFixedValue, Amount:=1000
               .ErrorBars.Border.ColorIndex = srsSystem.Border.ColorIndex
               .ErrorBars.Border.LineStyle = srsSystem.Border.LineStyle
               .ErrorBars.Border.Weight = srsSystem.Border.Weight
               .ErrorBars.EndStyle = xlNoCap
            End If
        End If
        If cType = "CUMULATIVE" And UCase(vstrStyle) = "TARGET" Then
               .ErrorBar Direction:=xlY, Include:=xlBoth, Type:=xlFixedValue, Amount:=20
               .ErrorBars.Border.ColorIndex = 3
               .ErrorBars.Border.LineStyle = xlDashDot
               .ErrorBars.Border.Weight = xlMedium
               
               '.ErrorBars.Border.ColorIndex = srsSystem.Border.ColorIndex
               '.ErrorBars.Border.LineStyle = srsSystem.Border.LineStyle
               '.ErrorBars.Border.Weight = srsSystem.Border.Weight
               .ErrorBars.EndStyle = xlNoCap
        End If
    End With
    
End Sub

Public Sub SetAllSeriesStyle(vSheetName As String)


    Dim Worksheet As Object
    Dim Chtobj As ChartObject
    Dim rsrsTarget As Series
    
    Dim ChartIndex As Integer
    Dim MaxChartIndex As Integer
    Dim vintSeries As Integer
    Dim vstrStyle As String
    
    Dim i As Integer
    
    Set Worksheet = ActiveWorkbook.Worksheets(vSheetName)
    'Worksheet.Activate
    'Set srsCol = Worksheets("PlotSetup").ChartObjects(1).Chart.SeriesCollection
    
    'find the chart that have max Series on chart
    '----------------------------------------------------------

    For Each Chtobj In Worksheet.ChartObjects
        
        For Each rsrsTarget In Chtobj.Chart.SeriesCollection
        
            vintSeries = vintSeries + 1
            'vstrStyle = Mid(rsrsTarget.Name, 1, 6)
            vstrStyle = rsrsTarget.Name
            Call SetPointsStyle(rsrsTarget, vstrStyle, vintSeries)
        Next rsrsTarget
        
    Next Chtobj

End Sub

Public Sub SetChartStyle(vChart As Object, Optional ByVal cType As String)


    Dim Worksheet As Object
    Dim Chtobj As ChartObject
    Dim rsrsTarget As Series
    
    Dim ChartIndex As Integer
    Dim MaxChartIndex As Integer
    Dim vintSeries As Integer
    Dim vstrStyle As String
    
    Dim i As Integer
    
    'Add by Dio
    On Error Resume Next
    
    'Set Worksheet = ActiveWorkbook.Worksheets(vSheetName)
    'Worksheet.Activate
    'Set srsCol = Worksheets("PlotSetup").ChartObjects(1).Chart.SeriesCollection
    
    'find the chart that have max Series on chart
    '----------------------------------------------------------

        For Each rsrsTarget In vChart.SeriesCollection
            vintSeries = vintSeries + 1
            'vstrStyle = Mid(rsrsTarget.Name, 1, 6)
            vstrStyle = rsrsTarget.Name
            Call SetPointsStyle(rsrsTarget, vstrStyle, vintSeries, cType)
        Next rsrsTarget
        

End Sub
