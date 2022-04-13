Option Explicit

'************************************************************
'*Title: PlotBoxTrendChart()
'*-----------------------------------------------------------
'*Notes: This program plot Universal chart.
'*
'*-----------------------------------------------------------
'*Include files:  OBC file and wat raw data
'*Output file: Universal chart
'*-----------------------------------------------------------
'*Date: 10/23/2004 kunshin chou - Initial code
'*************************************************************
Sub PlotBoxTrendChart(vSheetName As String)

Dim curSheet As Worksheet
Dim nowChart As Chart
Dim BoxRange As Range
Dim dataRange As Range
Dim nameRange As Range

Dim lLeft As Long
Dim lRight As Long
Dim lUpper As Long
Dim lBottom As Long

Dim iSerialColStart As Long
Dim iSerialColEnd As Long
Dim iSerialRowStart As Long
Dim iSerialRowEnd As Long

Dim vChartInfo As chartInfo
Dim OBCRange As Range

Dim vSplitBy As String
Dim vSplitID As String
Dim vExtendBy As String
Dim vBoxGroupLotFlag As Boolean
Dim vBoxMaxMinFlag As Boolean
Dim vBoxWaferSeqFlag As Boolean
Dim vBoxSigma As String
    
Dim vCharttitle As String
Dim vXLabel As String
Dim vYLabel As String
Dim vblnXLog As Boolean
Dim vblnYLog As Boolean
Dim vYMax As String
Dim vYMin As String
Dim vGraphHi As String
Dim vGraphLo As String
Dim vGraphMax As String
Dim vGraphMin As String

Dim i As Long
Dim targetA
Dim iCol As Long, iRow As Long
Dim vTargetName As String

    On Error GoTo myEnd
    
    Set curSheet = Sheets(vSheetName)
    curSheet.Activate
    ' Get the serial number of data set
    '-----------------------------------------------------------------
    Call GetRange(curSheet.UsedRange, lLeft, lUpper, lRight, lBottom)
    
    iSerialColStart = 5
    iSerialColEnd = lRight
    iSerialRowStart = 4
    iSerialRowEnd = lBottom
    If iSerialColStart > iSerialColEnd Then Exit Sub
    ' Get the OBC attribute of the chart
    '-----------------------------------------------------------------
    Set OBCRange = curSheet.Range(curSheet.Cells(1, 1), curSheet.Cells(iSerialRowEnd, 2)).CurrentRegion
    vChartInfo = getChartInfo(OBCRange)
    
    vSplitBy = vChartInfo.SplitBy
    vSplitID = vChartInfo.SplitID
    vExtendBy = vChartInfo.aExtendBy
    vBoxGroupLotFlag = IIf(UCase(vChartInfo.BoxGroupLotYesNo) = "YES", True, False)
    vBoxMaxMinFlag = IIf(UCase(vChartInfo.BoxMaxMinYesNo) = "YES", True, False)
    vBoxWaferSeqFlag = IIf(UCase(vChartInfo.BoxWaferSeqYesNo) = "YES", True, False)
    vBoxSigma = vChartInfo.BoxSigma
    vCharttitle = vChartInfo.ChartTitle
    vXLabel = vChartInfo.xLabel
    vYLabel = vChartInfo.yLabel
    vblnXLog = IsKey(vChartInfo.XScale, "Log")
    vblnYLog = IsKey(vChartInfo.YScale, "Log")
    vYMax = vChartInfo.yMax
    vYMin = vChartInfo.yMin
    vGraphHi = vChartInfo.GHi
    vGraphLo = vChartInfo.GLo
    vGraphMax = vChartInfo.GMax
    vGraphMin = vChartInfo.GMin
    targetA = Split(vChartInfo.vTargetYValueStr, ",")
    vTargetName = vChartInfo.vTargetNameStr
    '-------------------plot Box-------------------------------------
    
    Set BoxRange = curSheet.Range(ActiveSheet.Cells(iSerialRowStart, iSerialColStart), ActiveSheet.Cells(iSerialRowStart + 3, iSerialColEnd))
    Set nowChart = myCreateChart(curSheet, xlStockOHLC, 10, 10, 400, 300)
    
    With nowChart
        .SetSourceData Source:=BoxRange, PlotBy:=xlRows
        .chartType = xlStockOHLC
        .HasLegend = False
        .SeriesCollection(1).Name = vGraphLo & "%"
        .SeriesCollection(2).Name = vGraphMax & "%"
        .SeriesCollection(3).Name = vGraphMin & "%"
        .SeriesCollection(4).Name = vGraphHi & "%"
    End With
    
    Set nameRange = curSheet.Range(ActiveSheet.Cells(iSerialRowStart - 1, iSerialColStart), ActiveSheet.Cells(iSerialRowStart - 1, iSerialColEnd))
    
    '------------------put Median on chart----------------------------
    Set dataRange = curSheet.Range(ActiveSheet.Cells(iSerialRowStart + 4, iSerialColStart), ActiveSheet.Cells(iSerialRowStart + 4, iSerialColEnd))
    Call appendMedian(nowChart, dataRange, nameRange, vBoxMaxMinFlag)
    
    '------------------put Max on chart-------------------------------
    Set dataRange = curSheet.Range(ActiveSheet.Cells(iSerialRowStart + 5, iSerialColStart), ActiveSheet.Cells(iSerialRowStart + 5, iSerialColEnd))
    Call appendMax(nowChart, dataRange, nameRange, vBoxMaxMinFlag)
    
    '------------------put Min on chart-------------------------------
    Set dataRange = curSheet.Range(ActiveSheet.Cells(iSerialRowStart + 6, iSerialColStart), ActiveSheet.Cells(iSerialRowStart + 6, iSerialColEnd))
    Call appendMin(nowChart, dataRange, nameRange, vBoxMaxMinFlag)
    
    '------------------put Target on chart-------------------------------
    'Set dataRange = curSheet.Range(ActiveSheet.Cells(iSerialRowStart + 7, iSerialColStart), ActiveSheet.Cells(iSerialRowStart + 7, iSerialColEnd))
    'Call appendTarget(ActiveChart, dataRange, nameRange, vBoxMaxMinFlag)
    'iCol = iSerialColStart
    iRow = iSerialRowStart + 7
    For i = 0 To UBound(targetA)
        Call appendTarget(nowChart, curSheet.Range(curSheet.Cells(iRow + i, iSerialColStart), curSheet.Cells(iRow + i, iSerialColEnd)), nameRange, vBoxMaxMinFlag, getCOL(vTargetName, ",", i + 1), i)
    Next i
    
    ' Setting the Title of chart
    '-----------------------------------------------------------------
    With nowChart
        .HasTitle = True
        .ChartTitle.Characters.Text = vCharttitle
    End With
    ' Setting the XLabel of chart
    '-----------------------------------------------------------------
    With nowChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = vXLabel
    End With
    ' Setting the YLable of chart
    '-----------------------------------------------------------------
    With nowChart
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = vYLabel
    End With

    ' set scater's Max,Min
    Dim varXMax, varXMin, varYMax, varYMin
    Dim chtSetup As Chart
    Dim PlotSetupAxsValue As Axis, PlotSetupAxsCategory As Axis
    
    Set chtSetup = Worksheets("PlotSetup").ChartObjects(2).Chart
    Set PlotSetupAxsValue = chtSetup.Axes(xlValue)
    Set PlotSetupAxsCategory = chtSetup.Axes(xlCategory)
    
    Call getBoxMaxMin(nowChart, varXMax, varXMin, varYMax, varYMin)

    If vChartInfo.xMax <> "" Then varXMax = CDbl(vChartInfo.xMax)
    If vChartInfo.xMin <> "" Then varXMin = CDbl(vChartInfo.xMin)
    If vChartInfo.yMax <> "" Then varYMax = CDbl(vChartInfo.yMax)
    If vChartInfo.yMin <> "" Then varYMin = CDbl(vChartInfo.yMin)
    
    ' adjsut Y log scale
    If vblnYLog Then
        Dim varYOrder As Integer
        varYOrder = 0
        varYMax = 10 ^ CInt(Log(varYMax) / Log(10))
        varYMin = 10 ^ CInt(Log(varYMin) / Log(10) - 1)
    End If
    
    ' set Scatter's XlCategory
    Call SetBoxXlCategory(nowChart, PlotSetupAxsCategory, nameRange, varXMax, varXMin)
    ' set Scatter's XlValue
    Call SetBoxXlValue(nowChart, PlotSetupAxsValue, varYMax, varYMin, vYMax, vYMin)
    '
    Call SetBoxAxisScaleType(nowChart, vblnXLog, vblnYLog, varYMax, varYMin)

    nowChart.SeriesCollection(2).ErrorBar Direction:=xlY, Include:=xlPlusValues, Type:=xlCustom, Amount:="={0}"
    nowChart.SeriesCollection(3).ErrorBar Direction:=xlY, Include:= _
        xlPlusValues, Type:=xlCustom, Amount:="={0}"
        
    ' Set up Chart's the parameter label
    '-------------------------------------------------------------------------------------------------------------------------
    Set nameRange = curSheet.Range(curSheet.Cells(1, iSerialColStart), curSheet.Cells(1, iSerialColEnd))
    
    Select Case UCase(vChartInfo.Label)
        Case "MEDIAN"
        
        Case Else
            If vChartInfo.XParameter(1) <> "" And Not vChartInfo.aGroupParams = "Yes" Then
                Call applyDataLabelWafer(nameRange, BoxRange, nowChart.SeriesCollection(5), vChartInfo)
            Else
                'Call applyDataLabel(nameRange, BoxRange, nowChart.SeriesCollection(5))
            End If
    End Select
    
    ' Set Series Style
    '---------------------
'    Call SetSeriesStyle(nowChart, "BOXTREND")
    
    On Error GoTo myEnd

    Exit Sub
myEnd:
End Sub

'************************************************************
'*Title: prepareBoxTrendData()
'*-----------------------------------------------------------
'*Notes: This program prepare data for BoxTrend Chart.
'*
'*-----------------------------------------------------------
'*Include files:
'*Output file:
'*-----------------------------------------------------------
'*Date: 10/23/2004 kunshin chou - Initial code
'*************************************************************
Sub prepareBoxTrendData(vSheetName As String)

Dim tempSheetName As String
Dim tmpChartName As String
Dim tmpParameter1 As String
Dim tmpParameter2 As String
Dim tmpCurChartName As String

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

Dim curSheet As Worksheet
Dim OBCRange As Range
Dim RawdataRange As Range

Dim vChartInfo As chartInfo
Dim ChartTitle As String
Dim xLabel As String
Dim yLabel As String
Dim vblnXLog As Boolean
Dim vblnYLog As Boolean
Dim vGraphHi As String
Dim vGraphLo As String
Dim vGraphMax As String
Dim vGraphMin As String
Dim vTargetYValue As String
Dim targetA
Dim iTarget As Integer
Dim iRow As Long, iCol As Long
Dim i As Long

    On Error Resume Next
    tempSheetName = vSheetName
    Set curSheet = Sheets(tempSheetName)
    curSheet.Select
    curSheet.UsedRange.Select
    
    ' Get the serial number of data set
    '-----------------------------------------------------------------
    Call GetRange(curSheet.UsedRange, lLeft, lUpper, lRight, lBottom)

    iSerialColStart = 5
    iSerialColEnd = lRight
    iSerialRowStart = 4
    iSerialRowEnd = lBottom
    
    If iSerialColStart > iSerialColEnd Then Exit Sub
    
    dataCols = lRight
    dataRows = lLeft
    SerialCount = lRight - iSerialColStart + 1

    ' Get the OBC attribute of the chart
    '-----------------------------------------------------------------
    Set OBCRange = curSheet.Range(curSheet.Cells(1, 1), curSheet.Cells(iSerialRowEnd, 2)).CurrentRegion
    vChartInfo = getChartInfo(OBCRange)
    
    ChartTitle = vChartInfo.ChartTitle
    xLabel = vChartInfo.xLabel
    yLabel = vChartInfo.yLabel
    vblnXLog = IsKey(vChartInfo.XScale, "Log")
    vblnYLog = IsKey(vChartInfo.YScale, "Log")
    vGraphHi = IIf(vChartInfo.GHi = "", "75", vChartInfo.GHi)
    vGraphLo = IIf(vChartInfo.GLo = "", "25", vChartInfo.GLo)
    vGraphMax = IIf(vChartInfo.GMax = "", "90", vChartInfo.GMax)
    vGraphMin = IIf(vChartInfo.GMin = "", "25", vChartInfo.GMin)
    targetA = Split(vChartInfo.vTargetYValueStr, ",")
    
    curSheet.Range(Cells(iSerialRowEnd + 1, iSerialColStart - 1), Cells(iSerialRowEnd + 1, iSerialColStart - 1)).Value = vGraphLo & "%"
    curSheet.Range(Cells(iSerialRowEnd + 2, iSerialColStart - 1), Cells(iSerialRowEnd + 2, iSerialColStart - 1)).Value = vGraphMax & "%"
    curSheet.Range(Cells(iSerialRowEnd + 3, iSerialColStart - 1), Cells(iSerialRowEnd + 3, iSerialColStart - 1)).Value = vGraphMin & "%"
    curSheet.Range(Cells(iSerialRowEnd + 4, iSerialColStart - 1), Cells(iSerialRowEnd + 4, iSerialColStart - 1)).Value = vGraphHi & "%"
    curSheet.Range(Cells(iSerialRowEnd + 5, iSerialColStart - 1), Cells(iSerialRowEnd + 5, iSerialColStart - 1)).Value = "Mid"
    curSheet.Range(Cells(iSerialRowEnd + 6, iSerialColStart - 1), Cells(iSerialRowEnd + 6, iSerialColStart - 1)).Value = "Max"
    curSheet.Range(Cells(iSerialRowEnd + 7, iSerialColStart - 1), Cells(iSerialRowEnd + 7, iSerialColStart - 1)).Value = "Min"
    iRow = iSerialRowEnd + 8
    iCol = iSerialColStart - 1
    curSheet.Cells(iRow, iCol) = "Target"
    iTarget = 0
    
    For curSerial = iSerialColStart To iSerialColEnd
        
        If Application.WorksheetFunction.countA(curSheet.Columns(curSerial)) <> 0 Then
            Set RawdataRange = curSheet.Range(curSheet.Cells(iSerialRowStart, curSerial), curSheet.Cells(iSerialRowEnd, curSerial))
            RawdataRange.Select
            curSheet.Range(Cells(iSerialRowEnd + 1, curSerial), Cells(iSerialRowEnd + 1, curSerial)).Select
            '---------------Low---------------
            curSheet.Range(Cells(iSerialRowEnd + 1, curSerial), Cells(iSerialRowEnd + 1, curSerial)).Value = Application.WorksheetFunction.Percentile(RawdataRange, CSng(vGraphLo) / 100)
            '---------------Max----------------
            curSheet.Range(Cells(iSerialRowEnd + 2, curSerial), Cells(iSerialRowEnd + 2, curSerial)).Value = Application.WorksheetFunction.Percentile(RawdataRange, CSng(vGraphMax) / 100)
            '---------------Min---------------
            curSheet.Range(Cells(iSerialRowEnd + 3, curSerial), Cells(iSerialRowEnd + 3, curSerial)).Value = Application.WorksheetFunction.Percentile(RawdataRange, CSng(vGraphMin) / 100)
            '---------------High--------------
            curSheet.Range(Cells(iSerialRowEnd + 4, curSerial), Cells(iSerialRowEnd + 4, curSerial)).Value = Application.WorksheetFunction.Percentile(RawdataRange, CSng(vGraphHi) / 100)
            '---------------Med---------------
            curSheet.Range(Cells(iSerialRowEnd + 5, curSerial), Cells(iSerialRowEnd + 5, curSerial)).Value = Application.WorksheetFunction.Quartile(RawdataRange, 2)
            '---------------Max---------------
            curSheet.Range(Cells(iSerialRowEnd + 6, curSerial), Cells(iSerialRowEnd + 6, curSerial)).Value = Application.WorksheetFunction.Quartile(RawdataRange, 4)
            '---------------Min---------------
            curSheet.Range(Cells(iSerialRowEnd + 7, curSerial), Cells(iSerialRowEnd + 7, curSerial)).Value = Application.WorksheetFunction.Quartile(RawdataRange, 0)
        Else
            iTarget = iTarget + 1
        End If
        
        '---------------TargetYValue---------------
        curSheet.Cells(iRow + i, curSerial) = targetA(iTarget)
    Next curSerial
    iRow = iRow + UBound(targetA)
    curSheet.Range(Cells(iSerialRowEnd + 1, iSerialColStart - 1), Cells(iRow, iSerialColEnd)).Cut
    curSheet.Range(Cells(iSerialRowStart, iSerialColStart - 1), Cells(iSerialRowStart, iSerialColStart - 1)).Insert Shift:=xlDown

End Sub


Sub applyDataLabel_Median(nameRange As Range, dataRange As Range, vSeries As Series)
    Dim i As Integer

    For i = 1 To dataRange.Columns.Count
        
    Next i

'    Dim i As Long
'    Dim vPoint As Point
'    Dim vName As String
'
'    On Error Resume Next
'    For i = 1 To dataRange.Columns.Count
'
'        If i = 1 Then
'            Set vPoint = vSeries.Points(1)
'            vName = nameRange.Cells(2, 1).Value
'            'If applyWafer Then vName = "#" & getCOL(nameRange.Cells(1, 1).Value, "#", 2)
'            If IsEmpty(vPoint) <> True Then
'
'                vPoint.ApplyDataLabels Type:= _
'                        xlDataLabelsShowLabel, AutoText:=True, LegendKey:=False
'
'                vPoint.DataLabel.Characters.Text = vName
'
'                vPoint.DataLabel.AutoScaleFont = False
'                With vPoint.DataLabel.Characters.Font
'                    .Name = "Times New Roman"
'                    '.FontStyle = "標準"
'                    .FontStyle = "粗體"
'                    .Size = 7
'                    .Strikethrough = False
'                    .Superscript = False
'                    .Subscript = False
'                    .OutlineFont = False
'                    .Shadow = False
'                    .Underline = xlUnderlineStyleNone
'                    .ColorIndex = xlAutomatic
'                End With
'                With vPoint.DataLabel
'                  .HorizontalAlignment = xlCenter
'                  .VerticalAlignment = xlCenter
'                  .Position = xlLabelPositionAbove '
'                  .Orientation = xlHorizontal
'                End With
'
'            End If
'
'        End If
'
'        If Application.WorksheetFunction.countA(dataRange.Columns(i)) = 0 Then
'            Set vPoint = vSeries.Points(i + 1)
'            vName = nameRange.Cells(2, i + 1).Value
'            'If applyWafer Then vName = "#" & getCOL(nameRange.Cells(1, i + 1).Value, "#", 2)
'            If IsEmpty(vPoint) <> True Then
'
'                vPoint.ApplyDataLabels Type:= _
'                        xlDataLabelsShowLabel, AutoText:=True, LegendKey:=False
'
'
'                vPoint.DataLabel.Characters.Text = vName
'
'                vPoint.DataLabel.AutoScaleFont = False
'                With vPoint.DataLabel.Characters.Font
'                    .Name = "Times New Roman"
'                    '.FontStyle = "標準"
'                    .FontStyle = "粗體"
'                    .Size = 7
'                    .Strikethrough = False
'                    .Superscript = False
'                    .Subscript = False
'                    .OutlineFont = False
'                    .Shadow = False
'                    .Underline = xlUnderlineStyleNone
'                    .ColorIndex = xlAutomatic
'                End With
'                With vPoint.DataLabel
'                  .HorizontalAlignment = xlCenter
'                  .VerticalAlignment = xlCenter
'                  .Position = xlLabelPositionAbove
'                  .Orientation = xlHorizontal
'                End With
'            End If
'
'        End If
'
'    Next i

End Sub

Sub applyDataLabelWafer(nameRange As Range, dataRange As Range, vSeries As Series, vChartInfo As chartInfo)  ', Optional applyWafer As Boolean = True)

Dim i As Long
Dim vPoint As Point
Dim vName As String
                    
    
'    ActiveChart.SeriesCollection(5).Select
'    ActiveChart.SeriesCollection(5).ApplyDataLabels
'    ActiveChart.SeriesCollection(5).DataLabels.Select
'    ActiveChart.SeriesCollection(5).Points(1).DataLabel.Select
'    ActiveChart.SeriesCollection(5).Points(1).DataLabel.Text = "ttt"
'    Selection.Format.TextFrame2.TextRange.Characters.Text = "ttt"
    
    
    On Error Resume Next
    
    'by dio
    vSeries.ApplyDataLabels xlDataLabelsShowLabel, False, True
    vSeries.DataLabels.Select
    For i = 1 To dataRange.Columns.Count
        
            Set vPoint = vSeries.Points(i)
            vName = "#" & getCOL(nameRange.Cells(1, i).Value, "#", 2)
            If UCase(vChartInfo.DataLabel) = "MEDIAN" Then vName = Application.WorksheetFunction.index(vSeries.Values, i)
            If IsEmpty(vPoint) <> True Then
            
                'vPoint.ApplyDataLabels Type:= _
                        xlDataLabelsShowLabel, AutoText:=True, LegendKey:=False
                'vPoint.ApplyDataLabels Type:=xlDataLabelsShowLabel, AutoText:=True, LegendKey:=False
                'vPoint.DataLabel.Characters.Text = vName
                '
                '.Select
                'ActiveChart.SeriesCollection(5).Points(1).DataLabel.Text = "ssss1"
                'vPoint.Select
                vPoint.DataLabel.Text = vName
                vPoint.DataLabel.Format.TextFrame2.TextRange.Characters.Text = vName
                
                vPoint.DataLabel.AutoScaleFont = False
                With vPoint.DataLabel.Characters.Font
                    .Name = "Times New Roman"
                    '.FontStyle = "標準"
                    .FontStyle = "粗體"
                    .Size = 8
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                    .Underline = xlUnderlineStyleNone
                    .ColorIndex = 3
                End With
                With vPoint.DataLabel
                  '.Left = .Left + 50
                  .HorizontalAlignment = xlCenter
                  .VerticalAlignment = xlCenter
                  '.Position = xlLabelPositionAbove '
                  .Position = xlLabelPositionBestFit
                  .Orientation = xlHorizontal
                End With
                
            End If
    Next i
End Sub


Sub appendMax(curChart As Chart, dataRange As Range, nameRange As Range, disable As Boolean)

    Dim curSeries As Series
    Dim curSeriesCollection As SeriesCollection
         
    If disable <> True Then
    
        Set curSeries = curChart.SeriesCollection.NewSeries
        curSeries.Values = dataRange
        curSeries.chartType = xlXYScatter
        curSeries.Name = "100%"
    
        With curSeries.Border
            .Weight = xlHairline
            .LineStyle = xlNone
        End With
        With curSeries
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = 1
            .MarkerStyle = xlCircle
            .Smooth = False
            .Shadow = False
        End With
    End If
    


        
End Sub
Sub appendMedian(curChart As Chart, dataRange As Range, nameRange As Range, disable As Boolean)

    Dim curSeries As Series
    Dim curSeriesCollection As SeriesCollection

    Set curSeries = curChart.SeriesCollection.NewSeries
    
    curSeries.Values = dataRange
    curSeries.chartType = xlXYScatter
     
        With curSeries.Border
            .Weight = xlThin
            .LineStyle = xlAutomatic
        End With
        With curSeries
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlAutomatic
            .MarkerStyle = xlAutomatic
            .Smooth = False
            .MarkerSize = 5
            .Shadow = False
            .Name = "Median"
        End With
        
End Sub

Sub appendMin(curChart As Chart, dataRange As Range, nameRange As Range, disable As Boolean)

    Dim curSeries As Series
    Dim curSeriesCollection As SeriesCollection

If disable <> True Then
    Set curSeries = curChart.SeriesCollection.NewSeries
    curSeries.Values = dataRange
    curSeries.chartType = xlXYScatter
    curSeries.Name = "0%"
    
        With curSeries.Border
            .Weight = xlHairline
            .LineStyle = xlNone
        End With
        With curSeries
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = 1
            .MarkerStyle = xlCircle
            .Smooth = False
            .Shadow = False
        End With
End If
        
End Sub

Sub appendTarget(curChart As Chart, dataRange As Range, nameRange As Range, disable As Boolean, Optional ByVal mName As String = "", Optional ByVal sn As Integer = 0)

   Dim curSeries As Series
   Dim curSeriesCollection As SeriesCollection
   Dim i As Integer
   Dim iColor As Integer

If Application.WorksheetFunction.countA(dataRange.Rows(1)) <> 0 Then

    Set curSeries = curChart.SeriesCollection.NewSeries
    curSeries.Values = dataRange
    curSeries.chartType = xlXYScatter

    mName = getCOL(mName, ":", 1)
    mName = IIf(mName <> "", mName, " ")

    With curSeries
        .MarkerBackgroundColorIndex = xlNone
        .MarkerForegroundColorIndex = xlNone
        .MarkerStyle = xlNone
        .Smooth = False
        .MarkerSize = 5
        .Shadow = False
        iColor = Val(getCOL(mName, ":", 2))
        If iColor = 0 Then iColor = 41
        .Name = "Target"
        If sn >= 0 Then
            .ApplyDataLabels Type:=xlDataLabelsShowLabel, AutoText:=True, LegendKey:=False
            For i = .Points.Count To 2 Step -1
                .Points(i).DataLabel.Delete
            Next i
            .Points(1).DataLabel.Left = 40
            .Points(1).DataLabel.Top = .Points(1).DataLabel.Top - 5
            .Points(1).DataLabel.Text = mName
            .Points(1).DataLabel.AutoScaleFont = False
            .Points(1).DataLabel.Font.Size = 8
            .Points(1).DataLabel.Font.ColorIndex = iColor
            .Points(1).DataLabel.Font.Bold = True
        End If
        .ErrorBar Direction:=xlX, Include:=xlBoth, Type:=xlFixedValue, Amount:=0.5
        .ErrorBars.Border.ColorIndex = iColor
        If iColor = 41 Then .ErrorBars.Border.LineStyle = xlDash
        .ErrorBars.Border.Weight = xlMedium
        .ErrorBars.EndStyle = xlNoCap
    End With
End If

        
End Sub

Sub SetBoxXlCategory(ByRef curChart As Chart, PlotSetupAxsCategory As Axis, nameRange As Range, varXMax, varXMin)

        With curChart.Axes(xlCategory)
            
            .CategoryNames = nameRange
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
                    .Font.FontStyle = PlotSetupAxsCategory.AxisTitle.Font.FontStyle
                    .Font.Size = PlotSetupAxsCategory.AxisTitle.Font.Size
                    .AutoScaleFont = False
                End With
            End If
        End With
        
        'AxisScaleFit curChart.Axes(xlCategory), varXMax, varXMin


End Sub

Sub SetBoxXlValue(curChart As Chart, PlotSetupAxsValue As Axis, varYMax, varYMin, vYMax, vYMin)
    
    Dim i As Integer
    
    With curChart.Axes(xlValue)
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
                .Font.FontStyle = PlotSetupAxsValue.AxisTitle.Font.FontStyle
                .Font.Size = PlotSetupAxsValue.AxisTitle.Font.Size
                .AutoScaleFont = False
            End With
            
        End If
                  
        If Len(vYMin) > 0 Then
            .MinimumScale = CDbl(vYMin)
        Else
            .MinimumScale = (Int(varYMin / .MajorUnit)) * .MajorUnit
        End If
        
        If Len(vYMax) > 0 Then
            .MaximumScale = CDbl(vYMax)
        Else
            .MaximumScale = (Int(varYMax / .MajorUnit) + 1) * .MajorUnit
        End If
        
        .Crosses = xlCustom
        .CrossesAt = .MinimumScale
        
        If Len(CStr(.MinimumScale)) > 7 Then .TickLabels.NumberFormat = "0.00E+00"
    End With
    With curChart.Axes(xlValue, xlSecondary)
        .MinimumScale = curChart.Axes(xlValue, xlPrimary).MinimumScale
        .MaximumScale = curChart.Axes(xlValue, xlPrimary).MaximumScale
        .ScaleType = curChart.Axes(xlValue, xlPrimary).ScaleType
        .TickLabelPosition = xlNone
        .MajorTickMark = xlNone
    End With
    
    With curChart.FullSeriesCollection(5)
        .Format.Line.Visible = msoFalse
        .MarkerStyle = 8
        .MarkerSize = 3
        .Format.Fill.Solid
        .Format.Fill.Visible = msoTrue
        .Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End With
    
End Sub

Sub SetBoxAxisScaleType(curChart As Chart, vblnXLog As Boolean, vblnYLog As Boolean, ByVal vvarMax, ByVal vvarMin)
    Dim axsValue As Axis, axsCategory As Axis, axsValueSub As Axis
    Dim varXMin, varYMin
    Dim RoundNum As Integer
   
    If vblnYLog = True Then
        Set axsValue = curChart.Axes(xlValue)
        Set axsValueSub = curChart.Axes(xlValue, xlSecondary)
        With axsValue
            .MaximumScale = vvarMax
            .MinimumScale = vvarMin
            If vvarMin > 0 Then
                RoundNum = Int(Log10(vvarMin))
                .MinimumScale = 1 * 10 ^ RoundNum
            End If
            .Crosses = xlCustom
            .CrossesAt = .MinimumScale
            If .MinimumScale > 0 Then .ScaleType = xlLogarithmic
            .CrossesAt = .MinimumScale
            If Len(CStr(.MinimumScale)) > 7 Then .TickLabels.NumberFormat = "0.00E+00"
        End With
        With axsValueSub
            .MaximumScale = vvarMax
            .MinimumScale = vvarMin
            If vvarMin > 0 Then
                RoundNum = Int(Log10(vvarMin))
                .MinimumScale = 1 * 10 ^ RoundNum
            End If
            .Crosses = xlCustom
            .CrossesAt = .MinimumScale
            If .MinimumScale > 0 Then .ScaleType = xlLogarithmic
            .CrossesAt = .MinimumScale
            If Len(CStr(.MinimumScale)) > 7 Then .TickLabels.NumberFormat = "0.00E+00"
        End With
    Else
        curChart.Axes(xlValue).ScaleType = xlLinear
        curChart.Axes(xlValue, xlSecondary).ScaleType = xlLinear
    End If
   
End Sub

Sub getBoxMaxMin(ByRef curChart As Chart, ByRef varXMax, ByRef varXMin, ByRef varYMax, ByRef varYMin)

    Dim srsNew As Series
        For Each srsNew In curChart.SeriesCollection
            With srsNew
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
            End With
        Next

End Sub


Public Function SetSeriesStyle(nowChart As Chart, mChartType As String)
   Dim j As Integer, m As Integer
   Dim nowSheet As Worksheet
   'Dim nowChart As Chart
   Dim nowLegend As Legend
   Dim nowSeries As Series
   Dim nowShape As Shape
   Dim nowAxis As Axis
   Dim nowDataLabels As DataLabels, TemplateDataLabel As DataLabel
   Dim nowPlotArea As PlotArea
   Dim boolTemp As Boolean
   Const NameList As String = "TARGET,CORNER,SS,FF,TT,GOLDEN,MEDIAN,USL,LSL"
   Dim tSeries As Series
   
   Dim TemplateChart As Chart
   Dim i As Integer
   Dim ynSeries As Boolean
   Dim tmpChart As Chart
   
   On Error Resume Next
   
   Set nowSheet = Worksheets("PlotSetup")
   For i = 1 To nowSheet.ChartObjects.Count
      Set tmpChart = nowSheet.ChartObjects(i).Chart
      If UCase(mChartType) = UCase(tmpChart.ChartTitle.Text) Then
         Set TemplateChart = tmpChart
         Exit For
      End If
   Next i
   
   If TemplateChart Is Nothing And UCase(mChartType) = "SCATTER" Then Set TemplateChart = Worksheets("PlotSetup").ChartObjects(1).Chart
   If TemplateChart Is Nothing And UCase(mChartType) = "BOXTREND" Then Set TemplateChart = Worksheets("PlotSetup").ChartObjects(2).Chart
   
   If TemplateChart Is Nothing Then Exit Function
   
   ynSeries = True

      '-----------------
      'Fit Series Style
      '-----------------
      If ynSeries Then
         boolTemp = Application.ScreenUpdating
         Application.ScreenUpdating = False
         For m = 1 To nowChart.SeriesCollection.Count
            Set nowSeries = nowChart.SeriesCollection(m)
            If IsKey(NameList, UCase(nowSeries.Name), ",") Then
               Set tSeries = TemplateChart.SeriesCollection(UCase(nowSeries.Name))
            Else
               Set tSeries = TemplateChart.SeriesCollection(m)
            End If
            If Not IsEmpty(tSeries) Then
               With tSeries
                  nowSeries.Border.ColorIndex = .Border.ColorIndex
                  nowSeries.Border.Weight = .Border.Weight
                  nowSeries.Border.LineStyle = .Border.LineStyle
                  nowSeries.MarkerBackgroundColorIndex = .MarkerBackgroundColorIndex
                  nowSeries.MarkerForegroundColorIndex = .MarkerForegroundColorIndex
                  nowSeries.MarkerStyle = .MarkerStyle
                  nowSeries.Smooth = .Smooth
                  nowSeries.MarkerSize = .MarkerSize
                  nowSeries.Shadow = .Shadow
               End With
            End If
            Set tSeries = Nothing
            Set nowSeries = Nothing
         Next m
         Application.ScreenUpdating = boolTemp
      End If
      If UCase(mChartType) = "BOXTREND" Then
            If Not TemplateChart.SeriesCollection(5).HasDataLabels Then _
               nowChart.SeriesCollection(5).HasDataLabels = TemplateChart.SeriesCollection(5).HasDataLabels
         nowChart.ChartGroups(1).UpBars.Interior.ColorIndex = 2
         nowChart.ChartGroups(1).UpBars.Interior.ColorIndex = TemplateChart.ChartGroups(1).UpBars.Interior.ColorIndex
      End If
   
   'release object
   Set nowSheet = Nothing
   Set nowChart = Nothing
   Set nowLegend = Nothing
   Set nowSeries = Nothing
   Set nowShape = Nothing
   Set nowAxis = Nothing
   Set nowDataLabels = Nothing
   Set TemplateDataLabel = Nothing
   Set nowPlotArea = Nothing
End Function
