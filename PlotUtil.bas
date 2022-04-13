Option Explicit

Type chartInfo

    ChartTitle As String            'Chart Title
    SplitBy As String               'Split By
    aExtendBy As String             'Extend By
    xLabel As String                'X Label
    yLabel As String                'Y Label
    XScale As String                'X Scale
    YScale As String                'Y Scale
    xMax As String                  'xMax
    xMin As String                  'xMin
    yMax As String                  'yMax
    yMin As String                  'yMin
    ChartExpression As String       'Chart Expression
    SplitID As String               'Split ID
    Groupby As String               'Group By
    Method As String                'Method
    GMax As String                  'Graph Max%
    GMin As String                  'Graph Min%
    GLo As String                   'Graph Lo%
    GHi As String                   'Graph Hi%
    MapType As String               'Map Type
    BoxSigma As String              'Sigma
    aGroupParams As String          'Group Params
    GDataFilter As String           'Data Filter
    GTrendLines As String           'Show TrendLines
    GaussFilterTimes As String      'Out of 3 Sigma Filter
    GaussIntervalValue As String    'Sigma Divide
    BoxMaxMinYesNo As String        'Disable Max Min
    BoxWaferSeqYesNo As String      'Wafer Seq
    BoxGroupLotYesNo As String      'Group Lot
    vTargetNameStr As String        'Target Name
    vTargetXValueStr As String      'Target xValue
    vTargetYValueStr As String      'Target yValue
    vCornerXValueStr As String      'Corner xValue
    vCornerYValueStr As String      'Corner yValue
    GroupWafer As String            'Group Wafer
    DataLabel As String             'Data Label
    LegendLabel As String           'Legend Label
    Label As String                 'Label
    XParameter As New Collection
    YParameter As New Collection
    vSS As Variant
    vFF As Variant
    vTT As Variant
    vGoldendie As Variant
    vSensitivity As Variant
    vGroup As Variant
    
    
End Type

Public Function myCreateChart(nowSheet As Worksheet, nowChartType As XlChartType, ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single)         'As Chart
    Dim nowChart As Chart
    Dim i As Integer
    
    If nowChartType = xlStockOHLC Then nowChartType = xlLine
    Set myCreateChart = nowSheet.ChartObjects.Add(L, T, W, H).Chart
End Function


Sub PlotAllChart()
    
    Dim i As Long
    Dim curSheet As Worksheet
    
    For i = 1 To Sheets.Count
        If InStr(UCase(Sheets(i).Name), "CHART") = 1 Then
            Call PlotUniversalChart(Sheets(i).Name)
        ElseIf InStr(UCase(Sheets(i).Name), "BOX") = 1 Then
            Call prepareBoxTrendData(Sheets(i).Name)
            Call PlotBoxTrendChart(Sheets(i).Name)
        End If
        Set curSheet = Sheets(i)
        Call adjustChartObject(curSheet)
    Next i

End Sub

Public Function GenChartHeader()
    
    Dim i As Long, j As Long
    Dim nowRange As Range
    Dim OBCRange As Range
    Dim nowSheet As Worksheet
    Dim SCount As Integer
    Dim BCount As Integer
    Dim CCount As Integer
    Dim ChartID As Integer
    Dim chartType As String
    Dim tmpStr As String
   
    'Delete Old Chart Sheets
    For i = Worksheets.Count To 1 Step -1
        tmpStr = UCase(Worksheets(i).Name)
        If Left(tmpStr, 7) = "SCATTER" And tmpStr <> "SCATTER" Then Worksheets(i).Delete
        If Left(tmpStr, 8) = "BOXTREND" And tmpStr <> "BOXTREND" Then Worksheets(i).Delete
        If Left(tmpStr, 10) = "CUMULATIVE" And tmpStr <> "CUMULATIVE" Then Worksheets(i).Delete
    Next i

    Set OBCRange = Worksheets("ChartType").UsedRange
    SCount = 0
    BCount = 0
    CCount = 0

    'AddSheet ("All_Chart")
    For i = 2 To OBCRange.Rows.Count
        If Trim(OBCRange.Cells(i, 1)) = "" Then Exit For
        If Trim(OBCRange.Cells(i, 2)) <> "" Then
            chartType = UCase(Trim(OBCRange.Cells(i, 1)))
            'If Trim(OBCRange.Cells(i, 1)) <> Trim(OBCRange.Cells(i - 1, 1)) Then SCount = 0
            If IsExistSheet(Trim(OBCRange.Cells(i, 2))) Then
                If UCase(Trim(OBCRange.Cells(i, 2))) = "SCATTER" Or UCase(Trim(OBCRange.Cells(i, 2))) = "BOXTREND" Or UCase(Trim(OBCRange.Cells(i, 2))) = "CUMULATIVE" Then
                    Worksheets(Trim(OBCRange.Cells(i, 2))).Name = "set_" & Trim(OBCRange.Cells(i, 2))
                    OBCRange.Cells(i, 2) = "set_" & Trim(OBCRange.Cells(i, 2))
                End If
                Set nowSheet = Worksheets(Trim(OBCRange.Cells(i, 2)))
                Set nowRange = nowSheet.UsedRange
                nowRange.ClearFormats
                For j = 1 To (nowRange.Columns.Count - 2) / 2
                    If Trim(nowRange.Cells(1, 2 * j + 1)) <> "" Then
                        If chartType = "SCATTER" Then SCount = SCount + 1: ChartID = SCount
                        If chartType = "BOXTREND" Then BCount = BCount + 1: ChartID = BCount
                        If chartType = "CUMULATIVE" Then CCount = CCount + 1: ChartID = CCount
                        ' 2022/1/15 for control genCharts by pinScatter
                        If Not IsExistSheet("!" & chartType & CStr(ChartID)) Then
                            AddSheet (chartType & CStr(ChartID))
                            nowRange.Range(N2L(2 * j + 1) & ":" & N2L(2 * j + 2)).Copy
                            Worksheets(chartType & CStr(ChartID)).Range("A1").PasteSpecial xlPasteValues
                            Worksheets(chartType & CStr(ChartID)).Cells.ClearFormats
                        End If
                    Else
                        Exit For
                    End If
                Next j
            End If
        End If
    Next i
End Function

'===================================
' GenChartSummary
'     Dio 2005/12/19 Modifyied
'===================================
Public Sub GenChartSummary()
    Dim i As Long, j As Long, n As Integer
    Dim nowRange As Range
    Dim ChartSheet As Worksheet
    Dim nowSheet As Worksheet
    Dim nowChart As Chart
    Dim ChartID As Integer
    Dim chartType As String
    Dim sheetName As String
    Dim nowRow As Long
    Dim nowShape As Shape
    Dim mHeight As Integer
    Dim nowpath As String
    Dim FSO As New FileSystemObject
    Dim cTop As Long, cLeft As Long
    Dim mChart As Chart
    Dim OBCRange As Range
    Dim iRow As Long
    Dim iOBC As Integer, iChart As Integer
    Dim ynExistHome As Boolean
    Dim ErrCount As Integer
    Dim firstSheet As String
      
    'Create worksheet
    '-------------------------------------
    AddSheet ("All_Chart")
    UnpinScatter
    Set nowSheet = Worksheets("All_Chart")
   
    If IsExistSheet("SCATTER1") Then
        nowSheet.Move before:=Worksheets("SCATTER1")
    ElseIf IsExistSheet("BOXTREND1") Then
        nowSheet.Move before:=Worksheets("BOXTREND1")
    ElseIf IsExistSheet("CUMULATIVE1") Then
        nowSheet.Move before:=Worksheets("CUMULATIVE1")
    Else
        nowSheet.Move After:=Worksheets("Data")
    End If
    'Table header
    '-------------------------------------
    nowSheet.Cells(1, 1) = "Chart Type"
    nowSheet.Cells(1, 2) = "Sheet Name"
    nowSheet.Cells(1, 3) = "Chart Title"
    nowSheet.Cells(1, 4) = "Source OBC"
    nowSheet.Cells(1, 5) = "Remark"
    iRow = 2
   
    If IsExistSheet("ChartType") Then
        Set OBCRange = Worksheets("ChartType").UsedRange
    Else
        Exit Sub
    End If

    cLeft = 200
    cTop = 5
    nowpath = ActiveWorkbook.Path
    If Right(nowpath, 1) <> "\" Then nowpath = nowpath & "\"
   
'    For i = 1 To ActiveWorkbook.Worksheets.Count
'        iOBC = i - 1
'        iChart = 1
        
'        If Left(Worksheets(i).Name, Len("SCATTER")) = "SCATTER" Then
'            ChartType = "SCATTER"
'        ElseIf Left(Worksheets(i).Name, Len("BOXTREND")) = "BOXTREND" Then
'            ChartType = "BOXTREND"
'        ElseIf Left(Worksheets(i).Name, Len("CUMULAITVE")) = "CUMULATIVE" Then
'            ChartType = "CUMULATIVE"
'        End If
'    Next i
    
    
    For i = 2 To OBCRange.Rows.Count
        iOBC = i - 1
        iChart = 1
        If Trim(OBCRange.Cells(i, 1)) = "" Then Exit For
        If Trim(OBCRange.Cells(i, 2)) <> "" And IsExistSheet(Trim(OBCRange.Cells(i, 2))) Then
            chartType = Trim(OBCRange.Cells(i, 1))
            If Trim(OBCRange.Cells(i, 1)) <> Trim(OBCRange.Cells(i - 1, 1)) Then ChartID = 0
            Set nowRange = Worksheets(Trim(OBCRange.Cells(i, 2))).UsedRange
            For j = 1 To (nowRange.Columns.Count - 2) / 2
                 If Trim(nowRange.Cells(1, 2 * j + 1)) <> "" Then
                    ChartID = ChartID + 1
                    nowSheet.Cells(iRow, 1) = chartType
                    'nowSheet.Cells(iRow, 2) = chartType & CStr(chartID)
                    'Add link Text -> Chart
                    '----------------------
                    nowSheet.Hyperlinks.Add anchor:=nowSheet.Range(N2L(2) & CStr(iRow)), Address:="", _
                                 SubAddress:=chartType & CStr(ChartID) & "!A1", TextToDisplay:=chartType & CStr(ChartID), ScreenTip:=chartType & CStr(ChartID)
                    nowSheet.Cells(iRow, 3) = nowRange.Range(N2L(2 * j + 1) & ":" & N2L(2 * j + 2)).Cells(1, 2)
                    nowSheet.Cells(iRow, 4) = Trim(OBCRange.Cells(i, 2)) '& ":" & CStr(iOBC) & "-" & CStr(iChart)
                    'remark
                    'If Worksheets(i).UsedRange.Columns.Count < 3 Then nowSheet.Cells(nowRow, 4) = "No Data"
                    'Add link Chart -> Text
                    '----------------------
                    Set ChartSheet = Worksheets(chartType & CStr(ChartID))
                    ChartSheet.Select
                    ChartSheet.Range("A1").Select
                    If ChartSheet.ChartObjects.Count > 0 Then
                        'chartSheet.ChartObjects(1).Activate
                        Set nowChart = ChartSheet.ChartObjects(1).Chart
                        mHeight = nowChart.ChartArea.Height - 33
                        'copy chart to sheet all_chart
                        '-----------------------------
                        ErrCount = 0
myResume:
                        ErrCount = ErrCount + 1
                        If ErrCount > 5 Then Debug.Print "Paste Error:", nowChart.Name: GoTo mypass
                        On Error GoTo myResume
                        'nowChart.Activate
                        'DoEvents
                        nowChart.ChartArea.Copy
                        DoEvents
                        'nowSheet.Select
                        'nowSheet.Range("A1").Select
                        nowSheet.Paste
                        'DoEvents
mypass:
                        Set nowShape = nowSheet.Shapes(nowSheet.Shapes.Count)
                        nowShape.Left = cLeft + (nowShape.width + cTop) * (iOBC - 1)
                        nowShape.Top = cTop + (nowShape.Height + cTop) * (iChart - 1)
                        Set mChart = nowSheet.ChartObjects(nowSheet.ChartObjects.Count).Chart
                        '位置固定 2009/12/15
                        mChart.Parent.Placement = xlFreeFloating
                        nowSheet.ChartObjects(nowSheet.ChartObjects.Count).Name = ChartSheet.Name
                        ynExistHome = False
                        For n = mChart.Shapes.Count To 1 Step -1
                            'Debug.Print mChart.Shapes(n).Name
                            If mChart.Shapes(n).Name = "home" Then mChart.Shapes(n).Delete: ynExistHome = True
                        Next n
                        'Set nowShape = mChart.Shapes.AddShape(msoShapeCurvedUpArrow, 2, 2, 20, 20)
                        'Set nowShape = mChart.Shapes.AddShape(msoShapeFlowchartAlternateProcess, 2, 2, 80, 20)
                        Set nowShape = mChart.Shapes.AddShape(msoShapeStripedRightArrow, 2, mHeight + 5, 40, 20)
                        nowShape.Fill.ForeColor.SchemeColor = 41
                        'nowShape.Select
                        'Selection.Characters.Text = "SCATTER1"
                        nowSheet.Hyperlinks.Add anchor:=nowShape, Address:="", _
                                             SubAddress:=ChartSheet.Name & "!" & "A1", ScreenTip:=ChartSheet.Name
                  
                        If Not ynExistHome Then
                            'use picture
                            If FSO.FileExists(nowpath & "home_icon.jpg") Then
                                Set nowShape = nowChart.Shapes.AddPicture(nowpath & "home_icon.jpg", msoFalse, msoTrue, 2, mHeight, 30, 30)
                                nowShape.Name = "home"
                            Else  'use text box
                                'chartSheet.Activate
                                Set nowShape = nowChart.Shapes.AddShape(msoShapeCurvedUpArrow, 2, mHeight, 20, 20)
                                nowShape.Name = "home"
                                'nowShape.Adjustments
                            End If
                            ChartSheet.Hyperlinks.Add anchor:=nowShape, Address:="", _
                                                SubAddress:=nowSheet.Name & "!" & N2L(2) & CStr(iRow), _
                                                ScreenTip:="Chart List"
                        End If
                  
                        ' Show corner information into sheet
                        '------------------------------------
                        On Error Resume Next
                        If mChart.Shapes.Count > 0 Then
                            If mChart.Shapes(1).Type = msoTextBox Then
                                nowSheet.Cells(iRow, 5) = mChart.Shapes(1).TextFrame.Characters.Text
                            End If
                        End If
                        On Error GoTo 0
                    End If
                    iRow = iRow + 1
                    iChart = iChart + 1
                Else
                    Exit For
                End If
            Next j
        End If
    Next i
   
    nowSheet.Activate
    DoEvents
    With nowSheet.UsedRange '.Range("A:C")
        '.Font.Name = "Arial"
        '.Font.Size = 12
        .Columns.AutoFit
        .Rows.AutoFit
    End With
   
    For i = 1 To nowSheet.ChartObjects.Count
        nowSheet.ChartObjects(i).Left = nowSheet.ChartObjects(i).Left + nowSheet.Columns("A:C").width - cLeft
    Next i
End Sub

Public Function GenScatter(waferList() As String, siteNum As Integer)
    Dim nowSheet As Worksheet
    Dim vChartInfo As chartInfo
    Dim i As Long, j As Long, n As Long, m As Long
    Dim iSheet As Integer, iParameter As Integer, iWafer As Integer, iSite As Integer, iSen As Integer, iGroup As Integer
    Dim SubChartType As String
    Dim nowRow As Long, nowCol As Long
    Dim SeriesName As String, SubSeriesName As String
    Dim vSpec As specInfo
    Dim ParaArray() As String
    Dim valueArray() As String
    Dim strFormula As String
    Dim LumpFunction As String
    Dim groupList
     
    For iSheet = 1 To Worksheets.Count
        If UCase(Left(Worksheets(iSheet).Name, 7)) = "SCATTER" Then
            Set nowSheet = Worksheets(iSheet)
            nowSheet.Activate
            
            vChartInfo = getChartInfo(nowSheet.Range("A1").CurrentRegion)
            
            If IsExistSheet("Grouping") And UCase(vChartInfo.SplitBy) <> "GROUP" Then
                vChartInfo.SplitBy = "Group"
                ReDim vChartInfo.vGroup(1 To Worksheets("Grouping").Cells(1, 1).CurrentRegion.Rows.Count - 1, 1 To 2)
                For i = 1 To Worksheets("Grouping").Cells(1, 1).CurrentRegion.Rows.Count - 1
                    vChartInfo.vGroup(i, 1) = Worksheets("Grouping").Cells(i + 1, 1).Value
                    vChartInfo.vGroup(i, 2) = Worksheets("Grouping").Cells(i + 1, 2).Value
                Next i
            End If
            
            SubChartType = "NonUniversal"
            For j = 1 To vChartInfo.XParameter.Count
                If Not IsNumeric(vChartInfo.XParameter(j)) Then SubChartType = "Universal"
            Next j
            If UCase(vChartInfo.aGroupParams) = "YES" Then SubChartType = "Universal"
            If UCase(vChartInfo.XParameter(1)) = "PARA" Then SubChartType = "Universal"
            If UCase(vChartInfo.SplitBy) = "GROUP" Then SubChartType = "Group"
            If Not IsEmpty(vChartInfo.vSensitivity) Then SubChartType = "Sensitivity"
            
            'Universal Chart
            '---------------
            If SubChartType = "Universal" Then
                For iWafer = 0 To UBound(waferList)
                    'With Split by to select wafers
                    If IsKey(vChartInfo.SplitBy, waferList(iWafer), ",") Or Trim(UCase(vChartInfo.SplitBy)) = "ALL" _
                        Or Trim(UCase(vChartInfo.SplitBy)) = "LOT" Or Trim(UCase(vChartInfo.SplitBy)) = "WAFER" Or Trim(UCase(vChartInfo.SplitBy)) = "SPLITID" Then
                        SeriesName = "#" & waferList(iWafer)
                        nowCol = nowSheet.UsedRange.Columns.Count + 1
                        nowRow = 3
                        '--------------------------
                        For iParameter = 1 To vChartInfo.YParameter.Count
                            If UCase(vChartInfo.aGroupParams) <> "YES" Then
                                SeriesName = vChartInfo.YParameter(iParameter) & "#" & waferList(iWafer)
                                nowCol = nowSheet.UsedRange.Columns.Count + 1
                                nowRow = 3
                            End If
                            SubSeriesName = vChartInfo.YParameter(iParameter)
                            GoSub mySub
                        Next iParameter
                    End If
                Next iWafer
                
            'Non-Universal Chart
            '-------------------
            ElseIf SubChartType = "NonUniversal" Then
                For iParameter = 1 To vChartInfo.YParameter.Count
                    SeriesName = vChartInfo.YParameter(iParameter)
                    If vChartInfo.LegendLabel = "X" Then SeriesName = vChartInfo.XParameter(iParameter)
                    nowCol = nowSheet.UsedRange.Columns.Count + 1
                    nowRow = 3
                    For iWafer = 0 To UBound(waferList)
                        SubSeriesName = "#" & waferList(iWafer)
                        If Trim(UCase(vChartInfo.SplitBy)) = "WAFER" Then
                            SeriesName = vChartInfo.YParameter(iParameter) & SubSeriesName
                            If vChartInfo.LegendLabel = "X" Then SeriesName = vChartInfo.XParameter(iParameter) & SubSeriesName
                            If iWafer > 0 Then
                                nowCol = nowSheet.UsedRange.Columns.Count + 1
                                nowRow = 3
                            End If
                        End If
                        GoSub mySub
                    Next iWafer
                Next iParameter
            'Group
            '-------------------
            ElseIf SubChartType = "Group" Then
                nowCol = nowSheet.UsedRange.Columns.Count + 1
                nowRow = 3
                For iGroup = LBound(vChartInfo.vGroup) To UBound(vChartInfo.vGroup)
                    SeriesName = vChartInfo.vGroup(iGroup, 2)
                    groupList = Split(vChartInfo.vGroup(iGroup, 1), ",")
                    For iWafer = 0 To UBound(waferList)
                        'With Split by to select wafers
                        For iSen = LBound(groupList) To UBound(groupList)
                            If waferList(iWafer) = groupList(iSen) Then
                                For iParameter = 1 To vChartInfo.YParameter.Count
                                    SubSeriesName = "#" & waferList(iWafer)
                                    GoSub mySub
                                Next iParameter
                            End If
                        Next iSen
                    Next iWafer
                    nowCol = nowSheet.UsedRange.Columns.Count + 1
                    nowRow = 3
                Next iGroup
            'Sensitivity Chart
            '-------------------
            Else
                For iParameter = 1 To vChartInfo.YParameter.Count
                    SeriesName = vChartInfo.YParameter(iParameter)
                    If vChartInfo.LegendLabel = "X" Then SeriesName = vChartInfo.XParameter(iParameter)
                    nowCol = nowSheet.UsedRange.Columns.Count + 1
                    nowRow = 3
                    For iWafer = 0 To UBound(waferList)
                        SubSeriesName = "#" & waferList(iWafer)
                        'With Split by to select wafers
                        If IsKey(vChartInfo.SplitBy, waferList(iWafer), ",") Or Trim(UCase(vChartInfo.SplitBy)) = "ALL" _
                            Or Trim(UCase(vChartInfo.SplitBy)) = "LOT" Or Trim(UCase(vChartInfo.SplitBy)) = "WAFER" Or Trim(UCase(vChartInfo.SplitBy)) = "SPLITID" Then
                            
                            GoSub mySub
                        End If
                    Next iWafer
                Next iParameter
            End If
            ' Add Median
            '-----------
            On Error Resume Next
            If InStr(UCase(vChartInfo.ChartExpression), "MEDIAN") > 0 Then
                If InStr(UCase(vChartInfo.SplitBy), "GROUP") > 0 And IsNumeric(vChartInfo.XParameter(0)) Then
                    nowCol = nowSheet.UsedRange.Columns.Count + 1
                    n = 3
                    For iWafer = 1 To UBound(vChartInfo.vGroup)
                        nowRow = 1
                        nowSheet.Cells(nowRow + 0, nowCol).Value = nowSheet.Cells(nowRow + 0, n).Value
                        nowSheet.Cells(nowRow + 1, nowCol).Value = nowSheet.Cells(nowRow + 1, n).Value
                        nowRow = 2
                        For iParameter = 1 To vChartInfo.YParameter.Count
                            Dim valArray
                            ReDim valArray(1 To siteNum * UBound(vChartInfo.vGroup))
                            m = 1
                            For j = 1 To vChartInfo.YParameter.Count * UBound(vChartInfo.vGroup) * siteNum
                                If Trim(nowSheet.Cells(nowRow + j, n).Value) = vChartInfo.XParameter(iParameter) Then
                                    valArray(m) = nowSheet.Cells(nowRow + j, n + 1).Value
                                    m = m + 1
                                End If
                            Next j
                            nowSheet.Cells(nowRow + 1, nowCol).Value = vChartInfo.XParameter(iParameter)
                            nowSheet.Cells(nowRow + 1, nowCol + 1).Value = myMedian(Join(valArray, ","))
                            nowRow = nowRow + 1
                        Next iParameter
                        nowCol = nowCol + 2
                        n = n + 2
                    Next iWafer
                Else
                    nowCol = nowSheet.UsedRange.Columns.Count + 1
                    n = nowCol
                    nowRow = 3
                    nowSheet.Cells(1, nowCol) = "MEDIAN"
                    j = 1
                    For iWafer = 0 To UBound(waferList)
                        For iParameter = 1 To vChartInfo.YParameter.Count
                            j = j + 2
                            If Not InStr(nowSheet.Range(N2L(j) & CStr(1)), "MEDIAN") > 0 Then
                                If SubChartType = "Universal" Then
                                    nowSheet.Cells(nowRow, nowCol) = Application.WorksheetFunction.Median(nowSheet.Range(N2L(j) & ":" & N2L(j)))
                                    nowSheet.Cells(nowRow, nowCol + 1) = Application.WorksheetFunction.Median(nowSheet.Range(N2L(j + 1) & ":" & N2L(j + 1)))
                                Else
                                    m = 3 + ((iParameter - 1) * (UBound(waferList) + 1) + iWafer) * 2
                                    nowSheet.Cells(nowRow, nowCol) = Application.WorksheetFunction.Median(nowSheet.Range(N2L(m) & ":" & N2L(m)))
                                    nowSheet.Cells(nowRow, nowCol + 1) = Application.WorksheetFunction.Median(nowSheet.Range(N2L(m + 1) & ":" & N2L(m + 1)))
                                End If
                            End If
                            nowRow = nowRow + 1
                            If UCase(vChartInfo.aGroupParams) = "YES" Then Exit For
                        Next iParameter
                        If UCase(vChartInfo.SplitBy) = "WAFER" Then
                            nowSheet.Cells(1, nowCol) = "MEDIAN" & "#" & waferList(iWafer)
                            nowCol = nowCol + 2
                            nowRow = 3
                        End If
                    Next iWafer
                End If
                If InStr(UCase(vChartInfo.ChartExpression), "RAWDATA") <= 0 Then
                      For j = n - 1 To 3 Step -1
                          nowSheet.Columns(j).Delete
                    
                          ' Add USL, LSL, Target - Ti Project
                          '------------------------------------
                          If UCase(Trim(vChartInfo.XParameter(1))) = "SITE" Then
                              vSpec = getSPECInfo(vChartInfo.YParameter(1))
                              nowCol = nowSheet.UsedRange.Columns.Count + 1
                              nowSheet.Cells(1, nowCol) = "LSL"
                              nowSheet.Cells(3, nowCol) = 0
                              nowSheet.Cells(3, nowCol + 1) = vSpec.mLow
                              nowSheet.Cells(4, nowCol) = siteNum + 1
                              nowSheet.Cells(4, nowCol + 1) = vSpec.mLow
                              nowCol = nowCol + 2
                              nowSheet.Cells(1, nowCol) = "Target"
                              nowSheet.Cells(3, nowCol) = 0
                              nowSheet.Cells(3, nowCol + 1) = vSpec.mTarget
                              nowSheet.Cells(4, nowCol) = siteNum + 1
                              nowSheet.Cells(4, nowCol + 1) = vSpec.mTarget
                              nowCol = nowCol + 2
                              nowSheet.Cells(1, nowCol) = "USL"
                              nowSheet.Cells(3, nowCol) = 0
                              nowSheet.Cells(3, nowCol + 1) = vSpec.mHigh
                              nowSheet.Cells(4, nowCol) = siteNum + 1
                              nowSheet.Cells(4, nowCol + 1) = vSpec.mHigh
                          End If
                      Next j
                  End If
            End If
        End If
    Next iSheet
Exit Function

mySub:
    If UCase(Trim(vChartInfo.XParameter(iParameter))) = "SITE" Then SeriesName = "#" & waferList(iWafer)
    SeriesName = Replace(SeriesName, "=", "")
    nowSheet.Cells(1, nowCol) = SeriesName
    'nowSheet.Cells(1, nowCol + 1) = "'"
    nowSheet.Cells(1, nowCol + 1) = "#" & waferList(iWafer)
    If nowSheet.Cells(2, nowCol) <> "" Then
        nowSheet.Cells(2, nowCol) = "'" & nowSheet.Cells(2, nowCol) & "," & SubSeriesName
    Else
        nowSheet.Cells(2, nowCol) = "'" & SubSeriesName
    End If
    'nowSheet.Cells(2, nowCol) = nowSheet.Cells(2, nowCol) & subSeriesName
    For iSite = 1 To siteNum
        'Debug.Print "Site " & CStr(iSite) & ":" & getValueByPara(WaferList(iWafer), vChartInfo.YParameter(iParameter), iSite)
        LumpFunction = ""
        'X Value
        '-------
        If Not IsEmpty(vChartInfo.vSensitivity) Then
            For iSen = 0 To UBound(vChartInfo.vSensitivity)
                If vChartInfo.vSensitivity(iSen + 1, 1) = waferList(iWafer) Then
                    nowSheet.Cells(nowRow, nowCol) = vChartInfo.vSensitivity(iSen + 1, 2)
                    Exit For
                End If
            Next iSen
        ElseIf Len(vChartInfo.XParameter(iParameter)) = 0 Then
            nowSheet.Cells(nowRow, nowCol) = waferList(iWafer)
        ElseIf UCase(Trim(vChartInfo.XParameter(iParameter))) = "SITE" Then
            nowSheet.Cells(nowRow, nowCol) = iSite
        ElseIf UCase(Trim(vChartInfo.XParameter(iParameter))) = "PARA" Then
            nowSheet.Cells(nowRow, nowCol) = iParameter
        ElseIf Not IsNumeric(vChartInfo.XParameter(iParameter)) Then
            'Non-Formula
            If Not Left(vChartInfo.XParameter(iParameter), 1) = "=" Then
                'With FACTOR
                vSpec = getSPECInfo(vChartInfo.XParameter(iParameter))
                nowSheet.Cells(nowRow, nowCol) = getValueByPara(waferList(iWafer), vChartInfo.XParameter(iParameter), iSite, vSpec)
                'If nowSheet.Cells(nowRow, nowCol) <> "" Then nowSheet.Cells(nowRow, nowCol) = vSPEC.mFAC * nowSheet.Cells(nowRow, nowCol)
                'With Filter
                If UCase(vChartInfo.GDataFilter) = "YES" Then
                    If (Not IsEmpty(vSpec.mHigh) And Val(nowSheet.Cells(nowRow, nowCol)) > Val(vSpec.mHigh)) Or _
                       (Not IsEmpty(vSpec.mLow) And Val(nowSheet.Cells(nowRow, nowCol)) < Val(vSpec.mLow)) Then
                        nowSheet.Cells(nowRow, nowCol) = ""
                    End If
                End If
                'Formula
            Else
                'Formula without filter
                strFormula = getCOL(vChartInfo.XParameter(iParameter), "=", 2)
                LumpFunction = FormulaParse(strFormula, ParaArray)
                If UBound(ParaArray) = 0 Then vSpec = getSPECInfo(ParaArray(0))
                If LumpFunction = "" Then
                    Call FormulaValue(ParaArray, valueArray, waferList(iWafer), iSite, vSpec)
                Else
                    Call FormulaValue(ParaArray, valueArray, waferList(iWafer), siteNum, vSpec, LumpFunction)
                End If
                nowSheet.Cells(nowRow, nowCol) = FormulaEval(strFormula, ParaArray, valueArray)
            End If
        Else
           nowSheet.Cells(nowRow, nowCol) = vChartInfo.XParameter(iParameter)
        End If
        'Y Value
        '-------
        'If Left(vChartInfo.YParameter(iParameter), 1) = "'" Then vChartInfo.YParameter(iParameter) = Mid(vChartInfo.YParameter(iParameter), 2)
        'Non-Formula
        If Not Left(vChartInfo.YParameter(iParameter), 1) = "=" Then
            'With FACTOR
            vSpec = getSPECInfo(vChartInfo.YParameter(iParameter))
            nowSheet.Cells(nowRow, nowCol + 1) = getValueByPara(waferList(iWafer), vChartInfo.YParameter(iParameter), iSite, vSpec)
            'If nowSheet.Cells(nowRow, nowCol + 1) <> "" Then nowSheet.Cells(nowRow, nowCol + 1) = vSPEC.mFAC * nowSheet.Cells(nowRow, nowCol + 1)
            'With Filter
            If UCase(vChartInfo.GDataFilter) = "YES" Then
                If (Not IsEmpty(vSpec.mHigh) And Val(nowSheet.Cells(nowRow, nowCol + 1)) > Val(vSpec.mHigh)) Or _
                   (Not IsEmpty(vSpec.mLow) And Val(nowSheet.Cells(nowRow, nowCol + 1)) < Val(vSpec.mLow)) Then
                    nowSheet.Cells(nowRow, nowCol + 1) = ""
                End If
            End If
        'Formula
        Else
            'Formula without filter
            strFormula = getCOL(vChartInfo.YParameter(iParameter), "=", 2)
            LumpFunction = FormulaParse(strFormula, ParaArray)
            If UBound(ParaArray) = 0 Then vSpec = getSPECInfo(ParaArray(0))
            If LumpFunction = "" Then
                Call FormulaValue(ParaArray, valueArray, waferList(iWafer), iSite, vSpec)
            Else
                Call FormulaValue(ParaArray, valueArray, waferList(iWafer), siteNum, vSpec, LumpFunction)
            End If
            nowSheet.Cells(nowRow, nowCol + 1) = FormulaEval(strFormula, ParaArray, valueArray)
        End If
        nowRow = nowRow + 1
        If LumpFunction <> "" Then Exit For
    Next iSite
Return

End Function

Public Function GenBoxTrend(waferList() As String, siteNum As Integer)
   
    Dim nowSheet As Worksheet
    Dim vChartInfo As chartInfo
    Dim i As Long, j As Long
    Dim iSheet As Integer, iParameter As Integer, iWafer As Integer, iSite As Integer, iGroup As Integer
    Dim SubChartType As String
    Dim nowRow As Long, nowCol As Long
    Dim SeriesName As String, SubSeriesName As String
    Dim vSpec As specInfo
    Dim ParaArray() As String
    Dim valueArray() As String
    Dim strFormula As String
    Dim LumpFunction As String
    Dim ynFirst As Boolean
    Dim CutCol As Integer, InsCol As Integer
    Dim groupList
   
    For iSheet = 1 To Worksheets.Count
        If UCase(Left(Worksheets(iSheet).Name, 8)) = "BOXTREND" Then
            Set nowSheet = Worksheets(iSheet)
            nowSheet.Activate
            vChartInfo = getChartInfo(nowSheet.Range("A1").CurrentRegion)
            
            If IsExistSheet("Grouping") And UCase(vChartInfo.SplitBy) <> "GROUP" Then
                vChartInfo.SplitBy = "Group"
                ReDim vChartInfo.vGroup(1 To Worksheets("Grouping").Cells(1, 1).CurrentRegion.Rows.Count - 1, 1 To 2)
                For i = 1 To Worksheets("Grouping").Cells(1, 1).CurrentRegion.Rows.Count - 1
                    vChartInfo.vGroup(i, 1) = Worksheets("Grouping").Cells(i + 1, 1).Value
                    vChartInfo.vGroup(i, 2) = Worksheets("Grouping").Cells(i + 1, 2).Value
                Next i
            End If
            
            ynFirst = True
            
            For iParameter = 1 To vChartInfo.YParameter.Count
                ynFirst = True
                nowCol = nowSheet.UsedRange.Columns.Count + 2
                If iParameter = 1 Then nowCol = nowCol + 1
                nowRow = 4
                If iParameter <> 1 And UCase(vChartInfo.GroupWafer) = "YES" Then nowCol = nowCol - 1
                
                If UCase(vChartInfo.SplitBy) = "GROUP" Then
                    For iGroup = LBound(vChartInfo.vGroup) To UBound(vChartInfo.vGroup)
                        groupList = Split(vChartInfo.vGroup(iGroup, 1), ",")
                        If Not ynFirst Then nowCol = nowSheet.UsedRange.Columns.Count + 1
                        If ynFirst And iParameter > 1 Then nowCol = nowSheet.UsedRange.Columns.Count + 2
                        nowRow = 4
                        For iWafer = 0 To UBound(waferList)
                            For i = 0 To UBound(groupList)
                                If waferList(iWafer) = groupList(i) Then
                                    SubSeriesName = vChartInfo.YParameter(iParameter)
                                    SeriesName = vChartInfo.YParameter(iParameter) & "#" & vChartInfo.vGroup(iGroup, 2)
                                    GoSub mySub
                                    ynFirst = False
                                End If
                            Next i
                        Next iWafer
                    Next iGroup
                Else
                    For iWafer = 0 To UBound(waferList)
                        If IsKey(vChartInfo.SplitBy, waferList(iWafer), ",") Or Trim(UCase(vChartInfo.SplitBy)) = "ALL" _
                            Or Trim(UCase(vChartInfo.SplitBy)) = "LOT" Or Trim(UCase(vChartInfo.SplitBy)) = "WAFER" Or Trim(UCase(vChartInfo.SplitBy)) = "SPLITID" Then
                            SubSeriesName = vChartInfo.YParameter(iParameter)
    
                            'Y/N group wafer
                            If Not UCase(vChartInfo.GroupWafer) = "YES" Then
                                SeriesName = vChartInfo.YParameter(iParameter) & "#" & waferList(iWafer)
                                If Not ynFirst Then nowCol = nowSheet.UsedRange.Columns.Count + 1
                                If ynFirst And iParameter > 1 Then nowCol = nowSheet.UsedRange.Columns.Count + 2
                                nowRow = 4
                            Else
                                SeriesName = vChartInfo.YParameter(iParameter)
                            End If
                            GoSub mySub
                            ynFirst = False
                        End If
                    Next iWafer
                End If
                
            Next iParameter
            If Trim(UCase(vChartInfo.SplitBy)) = "WAFER" Then
                For j = 5 To nowSheet.UsedRange.Columns.Count
                    If nowSheet.Cells(3, j) = "" And nowSheet.Cells(1, j) <> "" Then nowSheet.Cells(3, j) = nowSheet.Cells(3, j - 1)
                Next j
                For j = nowSheet.UsedRange.Columns.Count To 5 Step -1
                    If nowSheet.Cells(1, j) = "" And nowSheet.Cells(2, j) = "" Then nowSheet.Columns(j).Delete
                Next j
            
                For iWafer = 1 To UBound(waferList) + 1
                    For iParameter = 1 To vChartInfo.YParameter.Count
                        CutCol = 4 + iWafer + (iWafer - 1) * vChartInfo.YParameter.Count + (iParameter - 1) * ((UBound(waferList) + 1) + 1 - iWafer)
                        InsCol = 4 + iParameter + (iWafer - 1) * (vChartInfo.YParameter.Count + 1)
                        Debug.Print iWafer, iParameter, N2L(CutCol), N2L(InsCol)
                        If CutCol <> InsCol Then
                            nowSheet.Columns(CutCol).Cut
                            nowSheet.Columns(InsCol).Insert Shift:=xlToRight
                        End If
                    Next iParameter
                    If iWafer < UBound(waferList) + 1 Then nowSheet.Columns(4 + iWafer * (vChartInfo.YParameter.Count + 1)).Insert Shift:=xlToRight
                Next iWafer
            End If
        End If
    Next iSheet
Exit Function

mySub:
    SeriesName = Replace(SeriesName, "=", "")
    SeriesName = IIf(InStr(SeriesName, ":"), getCOL(SeriesName, ":", 2), SeriesName)
    nowSheet.Cells(1, nowCol) = SeriesName
    If nowSheet.Cells(2, nowCol) <> "" Then
        nowSheet.Cells(2, nowCol) = "'" & nowSheet.Cells(2, nowCol) & "," & SubSeriesName
    Else
        nowSheet.Cells(2, nowCol) = "'" & SubSeriesName
    End If
    '使用 X value
    If vChartInfo.XParameter(iParameter) <> "" And Not vChartInfo.aGroupParams = "Yes" Then
        If ynFirst Then nowSheet.Cells(3, nowCol) = vChartInfo.XParameter(iParameter)
    ElseIf UCase(vChartInfo.SplitBy) = "GROUP" Then
        nowSheet.Cells(3, nowCol) = vChartInfo.vGroup(iGroup, 2)
    Else
        nowSheet.Cells(3, nowCol) = waferList(iWafer)
    End If


    'Input: nowRow, nowCol, iWafer, iParameter
    For iSite = 1 To siteNum
        If Not Left(vChartInfo.YParameter(iParameter), 1) = "=" Then
            'With FACTOR
            vSpec = getSPECInfo(vChartInfo.YParameter(iParameter))
            nowSheet.Cells(nowRow, nowCol) = getValueByPara(waferList(iWafer), vChartInfo.YParameter(iParameter), iSite, vSpec)
            'With Filter
            If UCase(vChartInfo.GDataFilter) = "YES" Then
                If (Not IsEmpty(vSpec.mHigh) And Val(nowSheet.Cells(nowRow, nowCol)) > Val(vSpec.mHigh)) Or _
                (Not IsEmpty(vSpec.mLow) And Val(nowSheet.Cells(nowRow, nowCol)) < Val(vSpec.mLow)) Then
                nowSheet.Cells(nowRow, nowCol) = ""
                End If
            End If
        'Formula
        Else
            'Formula without filter
            vSpec = getSPECInfo(vChartInfo.YParameter(iParameter))
            strFormula = getCOL(vChartInfo.YParameter(iParameter), "=", 2)
            LumpFunction = FormulaParse(strFormula, ParaArray)
            If LumpFunction = "" Then
                Call FormulaValue(ParaArray, valueArray, waferList(iWafer), iSite, vSpec)
            Else
                Call FormulaValue(ParaArray, valueArray, waferList(iWafer), siteNum, vSpec, LumpFunction)
            End If
            nowSheet.Cells(nowRow, nowCol) = FormulaEval(strFormula, ParaArray, valueArray)
        End If
        nowRow = nowRow + 1
        If LumpFunction <> "" Then Exit For
    Next iSite
Return

End Function

Public Function GenCumulative(waferList() As String, siteNum As Integer)
    Dim vChartInfo As chartInfo
    Dim iSheet As Integer, iParameter As Integer, iWafer As Integer, iSite As Integer, iGroup As Integer
    Dim i As Long, j As Long
    Dim nowSheet As Worksheet
    Dim nowRow As Long, nowCol As Long
    Dim yScaleA
    Dim itemNum As Integer
    Dim xMax As Double, xMin As Double
    Dim vSpec As specInfo
    Dim titlePara As String
    Dim ParaArray() As String
    Dim valueArray() As String
    Dim strFormula As String
    Dim LumpFunction As String
    Dim groupList
    Dim ynFirst As Boolean

    For iSheet = 1 To Worksheets.Count
        If UCase(Left(Worksheets(iSheet).Name, 10)) = "CUMULATIVE" Then
            Set nowSheet = Worksheets(iSheet)
            nowSheet.Activate
            'Get chart informatiom
            '---------------------
            vChartInfo = getChartInfo(nowSheet.Range("A1").CurrentRegion)
            
            If IsExistSheet("Grouping") And UCase(vChartInfo.SplitBy) <> "GROUP" Then
                vChartInfo.SplitBy = "Group"
                ReDim vChartInfo.vGroup(1 To Worksheets("Grouping").Cells(1, 1).CurrentRegion.Rows.Count - 1, 1 To 2)
                For i = 1 To Worksheets("Grouping").Cells(1, 1).CurrentRegion.Rows.Count - 1
                    vChartInfo.vGroup(i, 1) = Worksheets("Grouping").Cells(i + 1, 1).Value
                    vChartInfo.vGroup(i, 2) = Worksheets("Grouping").Cells(i + 1, 2).Value
                Next i
            End If
            
            nowCol = 3
            yScaleA = Split(yScaleStr, ",")
            nowSheet.Cells(1, nowCol) = "YScale"
            For i = 0 To UBound(yScaleA)
                nowRow = i + 2
                nowSheet.Cells(nowRow, nowCol) = 0
                Select Case UCase(vChartInfo.Method)
                    Case "LINEAR"
                        nowSheet.Cells(nowRow, nowCol + 1) = yScaleA(i) / 100
                    Case "LOGNOR"
                        nowSheet.Cells(nowRow, nowCol + 1) = Application.WorksheetFunction.NormSInv(yScaleA(i) / 100)
                    Case "WEIBULL"
                        nowSheet.Cells(nowRow, nowCol + 1) = Application.WorksheetFunction.Ln(Application.WorksheetFunction.Ln(1 / (1 - yScaleA(i) / 100)))
                End Select
            Next i

            nowCol = 5
            nowRow = 1
            xMax = getValueByPara(waferList(0), vChartInfo.YParameter(1), 1, vSpec)
            xMin = xMax
            ' Split by wafer
            '----------------------------------------------------
            If IsKey(vChartInfo.SplitBy, "wafer") Then
                For iWafer = 0 To UBound(waferList)
                    If iWafer > 0 Then nowCol = nowCol + 2: nowRow = 1
                    For iParameter = 1 To vChartInfo.YParameter.Count
                        vSpec = getSPECInfo(vChartInfo.YParameter(iParameter))
                        If iParameter > 1 And UCase(vChartInfo.aGroupParams) = "NO" Then nowCol = nowCol + 2: nowRow = 1
                        If vChartInfo.XParameter(iParameter) = "" Then
                            nowSheet.Cells(1, nowCol) = nowSheet.Cells(1, nowCol) & "," & vChartInfo.YParameter(iParameter)
                        Else
                            nowSheet.Cells(1, nowCol) = nowSheet.Cells(1, nowCol) & "," & vChartInfo.XParameter(iParameter)
                        End If
                        For iSite = 1 To siteNum
                            nowRow = nowRow + 1
                            
                            'Non-Formula
                            If Not Left(vChartInfo.YParameter(iParameter), 1) = "=" Then
                                nowSheet.Cells(nowRow, nowCol) = getValueByPara(waferList(iWafer), vChartInfo.YParameter(iParameter), iSite, vSpec)
                            'Formula
                            Else
                                strFormula = getCOL(vChartInfo.YParameter(iParameter), "=", 2)
                                LumpFunction = FormulaParse(strFormula, ParaArray)
                                If LumpFunction = "" Then
                                    Call FormulaValue(ParaArray, valueArray, waferList(iWafer), iSite, vSpec)
                                Else
                                    Call FormulaValue(ParaArray, valueArray, waferList(iWafer), siteNum, vSpec, LumpFunction)
                                End If
                                nowSheet.Cells(nowRow, nowCol) = FormulaEval(strFormula, ParaArray, valueArray)
                            End If
                            
                            'With Filter
                            If UCase(vChartInfo.GDataFilter) = "YES" Then
                                If (Not IsEmpty(vSpec.mHigh) And Val(nowSheet.Cells(nowRow, nowCol)) > Val(vSpec.mHigh)) Or _
                                   (Not IsEmpty(vSpec.mLow) And Val(nowSheet.Cells(nowRow, nowCol)) < Val(vSpec.mLow)) Then
                                    nowSheet.Cells(nowRow, nowCol) = ""
                                End If
                            End If
                            If nowSheet.Cells(nowRow, nowCol) <> "" And nowSheet.Cells(nowRow, nowCol) < 9E+100 Then
                                If xMax < nowSheet.Cells(nowRow, nowCol) Then xMax = nowSheet.Cells(nowRow, nowCol)
                                If xMin > nowSheet.Cells(nowRow, nowCol) Then xMin = nowSheet.Cells(nowRow, nowCol)
                            End If
                        Next iSite
                        If Left(nowSheet.Cells(1, nowCol), 1) = "," Then nowSheet.Cells(1, nowCol) = "'" & Mid(nowSheet.Cells(1, nowCol), 2)
                        If UCase(vChartInfo.aGroupParams) = "NO" Then
                            nowSheet.Cells(1, nowCol) = "'" & nowSheet.Cells(1, nowCol) & "#" & waferList(iWafer)
                        Else
                            nowSheet.Cells(1, nowCol) = "#" & waferList(iWafer)
                        End If
                    Next iParameter
                Next iWafer
            ' Split by group
            '----------------------------------------------------
            ElseIf IsKey(vChartInfo.SplitBy, "group") Then
                For iGroup = LBound(vChartInfo.vGroup) To UBound(vChartInfo.vGroup)
                    groupList = Split(vChartInfo.vGroup(iGroup, 1), ",")
                    ynFirst = True
                    For i = 0 To UBound(groupList)
                        For iWafer = 0 To UBound(waferList)
                            If waferList(iWafer) = groupList(i) Then
                                For iParameter = 1 To vChartInfo.YParameter.Count
                                    vSpec = getSPECInfo(vChartInfo.YParameter(iParameter))
                                    If ynFirst Then
                                        If vChartInfo.XParameter(iParameter) = "" Then
                                            nowSheet.Cells(1, nowCol) = nowSheet.Cells(1, nowCol) & "," & vChartInfo.YParameter(iParameter)
                                        Else
                                            nowSheet.Cells(1, nowCol) = nowSheet.Cells(1, nowCol) & "," & vChartInfo.XParameter(iParameter)
                                        End If
                                    End If
                                    For iSite = 1 To siteNum
                                        nowRow = nowRow + 1
                                        
                                        'Non-Formula
                                        If Not Left(vChartInfo.YParameter(iParameter), 1) = "=" Then
                                            nowSheet.Cells(nowRow, nowCol) = getValueByPara(waferList(iWafer), vChartInfo.YParameter(iParameter), iSite, vSpec)
                                        'Formula
                                        Else
                                            strFormula = getCOL(vChartInfo.YParameter(iParameter), "=", 2)
                                            LumpFunction = FormulaParse(strFormula, ParaArray)
                                            If LumpFunction = "" Then
                                                Call FormulaValue(ParaArray, valueArray, waferList(iWafer), iSite, vSpec)
                                            Else
                                                Call FormulaValue(ParaArray, valueArray, waferList(iWafer), siteNum, vSpec, LumpFunction)
                                            End If
                                            nowSheet.Cells(nowRow, nowCol) = FormulaEval(strFormula, ParaArray, valueArray)
                                        End If
                                        
                                        'With Filter
                                        If UCase(vChartInfo.GDataFilter) = "YES" Then
                                            If (Not IsEmpty(vSpec.mHigh) And Val(nowSheet.Cells(nowRow, nowCol)) > Val(vSpec.mHigh)) Or _
                                               (Not IsEmpty(vSpec.mLow) And Val(nowSheet.Cells(nowRow, nowCol)) < Val(vSpec.mLow)) Then
                                                nowSheet.Cells(nowRow, nowCol) = ""
                                            End If
                                        End If
                                        If nowSheet.Cells(nowRow, nowCol) <> "" And nowSheet.Cells(nowRow, nowCol) < 9E+100 Then
                                            If xMax < nowSheet.Cells(nowRow, nowCol) Then xMax = nowSheet.Cells(nowRow, nowCol)
                                            If xMin > nowSheet.Cells(nowRow, nowCol) Then xMin = nowSheet.Cells(nowRow, nowCol)
                                        End If
                                    Next iSite
                                    If Left(nowSheet.Cells(1, nowCol), 1) = "," Then nowSheet.Cells(1, nowCol) = "'" & Mid(nowSheet.Cells(1, nowCol), 2)
                                Next iParameter
                                ynFirst = False
                            End If
                        Next iWafer
                    Next i
                    If UCase(vChartInfo.aGroupParams) = "YES" Then
                        nowSheet.Cells(1, nowCol) = vChartInfo.vGroup(iGroup, 2)
                    ElseIf nowSheet.Cells(1, nowCol) <> "" Then
                        nowSheet.Cells(1, nowCol) = vChartInfo.vGroup(iGroup, 2) & " of " & nowSheet.Cells(1, nowCol)
                    End If
                    nowCol = nowCol + 2
                    nowRow = 1
                Next iGroup
            ' Non split by wafer
            '----------------------------------------------------
            Else
                For iParameter = 1 To vChartInfo.YParameter.Count
                    vSpec = getSPECInfo(vChartInfo.YParameter(iParameter))
                    If iParameter > 1 And UCase(vChartInfo.aGroupParams) = "NO" Then nowCol = nowCol + 2: nowRow = 1: titlePara = ""
                    If vChartInfo.XParameter(iParameter) = "" Then
                        nowSheet.Cells(1, nowCol) = nowSheet.Cells(1, nowCol) & "," & vChartInfo.YParameter(iParameter)
                    Else
                        nowSheet.Cells(1, nowCol) = nowSheet.Cells(1, nowCol) & "," & vChartInfo.XParameter(iParameter)
                    End If
                    
                    For iWafer = 0 To UBound(waferList)
                        For iSite = 1 To siteNum
                            nowRow = nowRow + 1
                            
                            'Non-Formula
                            If Not Left(vChartInfo.YParameter(iParameter), 1) = "=" Then
                                nowSheet.Cells(nowRow, nowCol) = getValueByPara(waferList(iWafer), vChartInfo.YParameter(iParameter), iSite, vSpec)
                            'Formula
                            Else
                                strFormula = getCOL(vChartInfo.YParameter(iParameter), "=", 2)
                                LumpFunction = FormulaParse(strFormula, ParaArray)
                                If LumpFunction = "" Then
                                    Call FormulaValue(ParaArray, valueArray, waferList(iWafer), iSite, vSpec)
                                Else
                                    Call FormulaValue(ParaArray, valueArray, waferList(iWafer), siteNum, vSpec, LumpFunction)
                                End If
                                nowSheet.Cells(nowRow, nowCol) = FormulaEval(strFormula, ParaArray, valueArray)
                            End If
                            
                            'With Filter
                            If UCase(vChartInfo.GDataFilter) = "YES" Then
                               If (Not IsEmpty(vSpec.mHigh) And Val(nowSheet.Cells(nowRow, nowCol)) > Val(vSpec.mHigh)) Or _
                                  (Not IsEmpty(vSpec.mLow) And Val(nowSheet.Cells(nowRow, nowCol)) < Val(vSpec.mLow)) Then
                                  nowSheet.Cells(nowRow, nowCol) = ""
                               End If
                            End If
                            If nowSheet.Cells(nowRow, nowCol) <> "" Then
                                If xMax < nowSheet.Cells(nowRow, nowCol) Then xMax = nowSheet.Cells(nowRow, nowCol)
                                If xMin > nowSheet.Cells(nowRow, nowCol) Then xMin = nowSheet.Cells(nowRow, nowCol)
                            End If
                        Next iSite
                    Next iWafer
                    If Left(nowSheet.Cells(1, nowCol), 1) = "," Then nowSheet.Cells(1, nowCol) = "'" & Mid(nowSheet.Cells(1, nowCol), 2)
                Next iParameter
            End If
                        
            For nowCol = 5 To nowSheet.UsedRange.Columns.Count
                nowSheet.Range(N2L(nowCol) & "2" & ":" & N2L(nowCol) & CStr(nowSheet.UsedRange.Rows.Count)).Sort _
                    Key1:=nowSheet.Range(N2L(nowCol) & "2"), Order1:=xlAscending, header:=xlGuess, _
                    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
            Next nowCol
            
            For nowCol = 6 To nowSheet.UsedRange.Columns.Count + 1 Step 2
                itemNum = Application.WorksheetFunction.Count(nowSheet.Range(N2L(nowCol - 1) & ":" & N2L(nowCol - 1)))
                For nowRow = 1 To itemNum
                    Select Case UCase(vChartInfo.Method)
                        Case "LINEAR"
                            nowSheet.Cells(nowRow + 1, nowCol) = (nowRow - 0.5) / itemNum
                        Case "LOGNOR"
                            nowSheet.Cells(nowRow + 1, nowCol) = Application.WorksheetFunction.NormSInv((nowRow - 0.5) / itemNum)
                        Case "WEIBULL"
                            nowSheet.Cells(nowRow + 1, nowCol) = Application.WorksheetFunction.Ln(Application.WorksheetFunction.Ln(1 / (1 - (nowRow - 0.5) / itemNum)))
                    End Select
                Next nowRow
            Next nowCol
        End If
    Next iSheet

End Function



'************************************************************
'*Title: adjustChartObject()
'*-----------------------------------------------------------
'*Notes: This program adjust Chart Object
'*
'*-----------------------------------------------------------
'*Include files:  None
'*Output file:
'*-----------------------------------------------------------
'*Date: 10/23/2004 kunshin chou - Initial code
'*************************************************************
Sub adjustChartObject(curSheet As Worksheet)

    Dim chartObj As ChartObject
    Dim chtSetup As Chart
    
    Set chtSetup = Worksheets("PlotSetup").ChartObjects(1).Chart
    
    For Each chartObj In curSheet.ChartObjects
        chartObj.Placement = xlFreeFloating
        
        chartObj.Chart.ChartArea.Interior.ColorIndex = chtSetup.ChartArea.Interior.ColorIndex
        With chartObj.Chart.PlotArea.Interior
            .ColorIndex = chtSetup.PlotArea.Interior.ColorIndex
            .PatternColorIndex = chtSetup.PlotArea.Interior.PatternColorIndex
            .Pattern = chtSetup.PlotArea.Interior.Pattern
        End With
    Next chartObj


End Sub

'************************************************************
'*Title: GetChartInfo()
'*-----------------------------------------------------------
'*Notes: This program for get all Info of the chart
'*
'*-----------------------------------------------------------
'*Include files:  Range object
'*Output file: chartInfo
'*-----------------------------------------------------------
'*Date: 10/23/2004 kunshin chou - Initial code
'************************************************************
Public Function getChartInfo(vRange As Range) As chartInfo

Dim vChartInfo As chartInfo
Dim i As Long
Dim keyWord As String

Dim Start As Boolean
Dim keyChange As Boolean
Dim startRow As Long
Dim RowCount As Long
Dim endRow As Long
Dim tmp As Variant
Dim xValue As String
Dim yValue As String
Dim lStr As String
Dim rStr As String

Dim LTemp As String
Dim RTemp As String
Dim nowSheet As Worksheet

Set nowSheet = vRange.Worksheet

On Error Resume Next

If Len(Trim$(vRange.Cells(1, 1).Value) & "") > 0 Then
    For i = 1 To nowSheet.UsedRange.Rows.Count
        LTemp = UCase(Trim$(vRange.Cells(i, 1).Value))
        RTemp = UCase(Trim$(vRange.Cells(i, 2).Value))
        
        If LTemp = "CHART TITLE" Then vChartInfo.ChartTitle = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "SPLIT BY" Then vChartInfo.SplitBy = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "EXTEND BY" Then vChartInfo.aExtendBy = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "X LABEL" Then vChartInfo.xLabel = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "Y LABEL" Then vChartInfo.yLabel = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "X SCALE" Then vChartInfo.XScale = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "Y SCALE" Then vChartInfo.YScale = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "XMAX" Then vChartInfo.xMax = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "XMIN" Then vChartInfo.xMin = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "YMAX" Then vChartInfo.yMax = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "YMIN" Then vChartInfo.yMin = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "CHART EXPRESSION" Then vChartInfo.ChartExpression = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "SPLIT ID" Then vChartInfo.SplitID = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "GROUP BY" Then vChartInfo.Groupby = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "METHOD" Then vChartInfo.Method = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "GRAPH MAX%" Then vChartInfo.GMax = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "GRAPH MIN%" Then vChartInfo.GMin = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "GRAPH LO%" Then vChartInfo.GLo = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "GRAPH HI%" Then vChartInfo.GHi = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "MAP TYPE" Then vChartInfo.MapType = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "SIGMA" Then vChartInfo.BoxSigma = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "GROUP PARAMS" Then vChartInfo.aGroupParams = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "DATA FILTER" Then vChartInfo.GDataFilter = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "TRENDLINES" Then vChartInfo.GTrendLines = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "OUT OF 3 SIGMA FILTER" Then vChartInfo.GaussFilterTimes = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "SIGMA DIVIDE" Then vChartInfo.GaussIntervalValue = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "DISABLE MAX MIN" Then vChartInfo.BoxMaxMinYesNo = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "WAFER SEQ" Then vChartInfo.BoxWaferSeqYesNo = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "GROUP LOT" Then vChartInfo.BoxGroupLotYesNo = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "DATA LABEL" Then vChartInfo.DataLabel = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "LEGEND LABEL" Then vChartInfo.LegendLabel = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "TARGET NAME" Then vChartInfo.vTargetNameStr = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "TARGET XVALUE" Then vChartInfo.vTargetXValueStr = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "TARGET YVALUE" Then vChartInfo.vTargetYValueStr = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "CORNER XVALUE" Then vChartInfo.vCornerXValueStr = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "CORNER YVALUE" Then vChartInfo.vCornerYValueStr = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "GROUP WAFER" Then vChartInfo.GroupWafer = Trim$(vRange.Cells(i, 2).Value)
        If LTemp = "LABEL" Then vChartInfo.Label = Trim$(vRange.Cells(i, 2).Value)
        
        Select Case LTemp
          Case "TT", "SS", "FF", "GOLDENDIE", "SENSITIVITY", "GROUP"
             Start = False
        End Select
        
        If Start Then
            If lStr = "X" Or rStr = "Y" Then
                xValue = Trim$(Trim$(vRange.Cells(i, 1).Value))
                yValue = Trim$(Trim$(vRange.Cells(i, 2).Value))
            ElseIf lStr = "Y" Or rStr = "X" Then
                xValue = Trim$(Trim$(vRange.Cells(i, 2).Value))
                yValue = Trim$(Trim$(vRange.Cells(i, 1).Value))
            End If
            
            If Len(yValue) > 0 Then
                On Error Resume Next
                vChartInfo.YParameter.Add Trim$(yValue), CStr(i)
                On Error GoTo 0
            End If
            If Len(xValue) > -1 Then
                On Error Resume Next
                vChartInfo.XParameter.Add Trim$(xValue), CStr(i)
                On Error GoTo 0
            End If
        End If
        

        If (LTemp = "Y" And (RTemp = "X" Or RTemp = "")) Or (RTemp = "Y" And (LTemp = "X" Or LTemp = "")) Then
            Start = True
            lStr = LTemp
            rStr = RTemp
        End If

        If IsKeyWord(LTemp) = True Then
            If RowCount > 0 Then
                On Error Resume Next
                endRow = startRow + RowCount
                tmp = vRange.Range(nowSheet.Cells(startRow + 1, 1), nowSheet.Cells(endRow, 2)).Value
                If UCase(Trim$(keyWord)) = "SS" Then vChartInfo.vSS = tmp
                If UCase(Trim$(keyWord)) = "FF" Then vChartInfo.vFF = tmp
                If UCase(Trim$(keyWord)) = "TT" Then vChartInfo.vTT = tmp
                If UCase(Trim$(keyWord)) = "GOLDENDIE" Then vChartInfo.vGoldendie = tmp
                If UCase(Trim$(keyWord)) = "SENSITIVITY" Then vChartInfo.vSensitivity = tmp
                If UCase(Trim$(keyWord)) = "GROUP" Then vChartInfo.vGroup = tmp
                On Error GoTo 0
            End If
            keyWord = LTemp
            RowCount = 0
            startRow = i
        Else
            RowCount = RowCount + 1
        End If
    Next i
  
    If RowCount > 0 Then
        On Error Resume Next
        endRow = startRow + RowCount
        tmp = vRange.Range(nowSheet.Cells(startRow + 1, 1), nowSheet.Cells(endRow, 2)).Value
        If UCase(Trim$(keyWord)) = "SS" Then vChartInfo.vSS = tmp
        If UCase(Trim$(keyWord)) = "FF" Then vChartInfo.vFF = tmp
        If UCase(Trim$(keyWord)) = "TT" Then vChartInfo.vTT = tmp
        If UCase(Trim$(keyWord)) = "GOLDENDIE" Then vChartInfo.vGoldendie = tmp
        If UCase(Trim$(keyWord)) = "SENSITIVITY" Then vChartInfo.vSensitivity = tmp
        If UCase(Trim$(keyWord)) = "GROUP" Then vChartInfo.vGroup = tmp
        
        On Error GoTo 0
        RowCount = 0
    End If
    
End If
    
getChartInfo = vChartInfo

End Function

'************************************************************
'*Title: PlotRollOffCart()
'*-----------------------------------------------------------
'*Notes: This program plot RollOff Cart.
'*
'*-----------------------------------------------------------
'*Include files:  None
'*Output file:
'*-----------------------------------------------------------
'*Date: 10/23/2004 kunshin chou - Initial code
'*************************************************************
Sub PlotRollOffCart()

End Sub


'************************************************************
'*Title: IsKeyWord()
'*-----------------------------------------------------------
'*Notes:
'*
'*-----------------------------------------------------------
'*Include files:  None
'*Output file:
'*-----------------------------------------------------------
'*Date: 10/23/2004 kunshin chou - Initial code
'*************************************************************
Function IsKeyWord(vkeyword As String) As Boolean

    IsKeyWord = False
    
    If UCase(Trim$(vkeyword)) = "CHART TITLE" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "SPLIT BY" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "EXTEND BY" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "X LABEL" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "Y LABEL" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "X SCALE" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "Y SCALE" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "XMAX" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "XMIN" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "YMAX" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "YMIN" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "CHART EXPRESSION" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "SPLIT ID" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "GROUP BY" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "METHOD" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "GRAPH MAX%" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "GRAPH MIN%" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "GRAPH LO%" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "GRAPH HI%" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "MAP TYPE" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "SIGMA" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "GROUP PARAMS" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "DATA FILTER" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "OUT OF 3 SIGMA FILTER" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "SIGMA DIVIDE" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "DISABLE MAX MIN" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "WAFER SEQ" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "GROUP LOT" Then IsKeyWord = True
                
    If UCase(Trim$(vkeyword)) = "TARGET NAME" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "TARGET XVALUE" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "TARGET YVALUE" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "CORNER XVALUE" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "CORNER YVALUE" Then IsKeyWord = True

    If UCase(Trim$(vkeyword)) = "Y" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "CORNER" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "SS" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "FF" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "TT" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "GOLDENDIE" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "SENSITIVITY" Then IsKeyWord = True
    If UCase(Trim$(vkeyword)) = "GROUP" Then IsKeyWord = True
    
    
End Function
