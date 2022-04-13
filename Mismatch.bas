Option Explicit
Public ActiveItems() As String
Public PreScreen As Boolean
Public mHigh As Double
Public mLow As Double
Public MergeTSK As Boolean
Public filterTimes As Integer
Public ScreenPair As Boolean
Public ErrFlag As Boolean
Public Preshrink As Double

Sub RunMismatch()
    
    If Not IsExistSheet("Data") Then MsgBox ("Please load data before operation"): Exit Sub
    If Not IsExistSheet("Formula") Then MsgBox ("Please set formula before operation"): Exit Sub
    ErrFlag = False
    
    FrmOption.Show

    If ErrFlag = True Then Exit Sub
    Call SplitByType
    Call PreDataScreening
    Call MergeRepeatedTks("Global")
    Call Filterby3sigma("Global")
    Call MismatchingCalculation
    If ErrFlag = True Then Exit Sub
    
    Call MergeRepeatedTks("Δ")
    Call Filterby3sigma("Δ")
    
'    Call Speed
    Call PlotChart
'    Call Unspeed
    
    MsgBox "Finished"
End Sub

Sub SplitByType()
    Dim nSheet As Integer
    Dim iSheet As Integer, iWafer As Integer
    Dim mSheet As String
    Dim nowSheet As Worksheet
    Dim DataSheet As Worksheet
    Dim waferList() As String, siteNum As Integer
    Dim waferRange As Range
    Dim nowRow As Long
    Dim strRange As String
    Dim W As String, L As String
    Dim specRange As Range
    Dim waferNum As Integer, LotName As String, ProductID As String
    Dim iCol As Variant, iRow As Long
    Dim i As Long, j As Long
    Dim nowRange As Range
    Dim nowParameter As String
    Dim reValue As Variant
    Dim objPara As specInfo
    Dim nowCell As Range
    Dim formatStr As String
                
    nSheet = UBound(ActiveItems) + 1
    Set DataSheet = Worksheets("Data")
    siteNum = getSiteNum("Data")
    Call GetWaferList("Data", waferList)

    For iSheet = 1 To nSheet
        mSheet = ActiveItems(iSheet - 1) & "_Raw"
        Set nowSheet = AddSheet(mSheet)
        waferNum = UBound(waferList) + 1
        ProductID = Trim(Worksheets(dSheet).Cells(2, 2))
        LotName = Trim(Worksheets(dSheet).Cells(3, 2))
        If Left(LotName, 1) = ":" Then LotName = Mid(LotName, 2)
        If Left(ProductID, 1) = ":" Then ProductID = Mid(ProductID, 2)
        With nowSheet.Range("A1:M1")
            .Range("A1") = "PROD:"
            .Range("B1") = ProductID
            .Range("B1").Font.Color = RGB(255, 0, 0)
            .Range("C1") = "LOT:"
            .Range("D1") = LotName
            .Range("D1").Font.Color = RGB(255, 0, 0)
            .Range("E1") = "WAFERs:"
            .Range("F1") = waferNum
            .Range("F1").Font.Color = RGB(255, 0, 0)
            .Range("L1") = "SiteNum:"
            .Range("M1") = siteNum
            .Range("M1").Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    
        Set specRange = Worksheets(mSheet).UsedRange
        nowRow = nowSheet.UsedRange.Rows.Count + 1
        
        For iWafer = 0 To UBound(waferList)
            Set waferRange = DataSheet.Range("wafer_" & waferList(iWafer))
            For iRow = 1 To waferRange.Rows.Count
                If UCase(waferRange.Cells(iRow, 2).Value) = "PARAMETER" Then
                    specRange.Rows(nowRow) = _
                        Array("DEVICE", "ITEM", "UNIT", "Med", "Avg", "Stdev", "Med+3Stdev", "Med-3Stdev", "W", "L", "1/sqrt(WL)", "Δ" & ActiveItems(iSheet - 1), "Count")
                    Set reValue = getRangeByPara(waferList(iWafer), "Parameter", siteNum)
                    strRange = nowSheet.Rows(nowRow).Range(Cells(specRange.Columns.Count + 1), Cells(specRange.Columns.Count + siteNum)).Address(False, False)
                    Range(strRange).Value = reValue.Value
                    
                    With nowSheet.Rows(nowRow).Range(Cells(1, 1), Cells(specRange.Columns.Count + siteNum)).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent5
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    nowRow = nowRow + 1
                ElseIf UCase(getCOL(waferRange.Cells(iRow, 2), "_", 1)) = UCase(ActiveItems(iSheet - 1)) Then
                    nowParameter = waferRange.Cells(iRow, 2)
                    nowSheet.Cells(nowRow, 1) = getCOL(nowParameter, "_", 4)
                    nowSheet.Cells(nowRow, 2) = nowParameter
                    nowSheet.Cells(nowRow, 3) = waferRange.Cells(iRow, 3).Value
                    
                    Set reValue = getRangeByPara(waferList(iWafer), getCOL(nowParameter, ":", 1), siteNum)
                    strRange = nowSheet.Rows(nowRow).Range(Cells(specRange.Columns.Count + 1), Cells(specRange.Columns.Count + siteNum)).Address(False, False)
                    Range(strRange).Value = reValue.Value
                    
                    W = getCOL(nowParameter, "_", 2)
                    If Left(W, 1) = "p" Then W = "0" & W
                    W = Replace(W, "p", ".")
                    L = getCOL(nowParameter, "_", 3)
                    If Left(L, 1) = "p" Then L = "0" & L
                    L = Replace(L, "p", ".")
                    W = CStr(Val(W) * Preshrink)
                    L = CStr(Val(L) * Preshrink)
                                        
                    With nowSheet.Range("D" & CStr(nowRow))
                        .Cells(1, 1).FormulaLocal = "=Median(" & strRange & ")"
                        .Cells(1, 2).FormulaLocal = "=Average(" & strRange & ")"
                        .Cells(1, 3).FormulaLocal = "=Stdev(" & strRange & ")"
                        .Cells(1, 4).FormulaLocal = "=" & "RC[-3]+3*RC[-1]"
                        .Cells(1, 5).FormulaLocal = "=" & "RC[-4]-3*RC[-2]"
                        .Cells(1, 6).FormulaLocal = W
                        .Cells(1, 7).FormulaLocal = L
                        .Cells(1, 8).FormulaLocal = "=" & "1/sqrt(RC[-2]*RC[-1])"
                        .Cells(1, 9).FormulaLocal = "=" & "RC[-6]/RC[-1]"
                        .Cells(1, 10).FormulaLocal = "=Count(" & strRange & ")"
                    End With
                    nowRow = nowRow + 1
                End If
            Next iRow
            nowRow = nowRow + 1
        Next iWafer
        
        For Each iCol In Array(3, specRange.Columns.Count, specRange.Columns.Count + siteNum)
            Set nowRange = nowSheet.Columns(iCol).Range(Cells(2, 1), Cells(nowSheet.UsedRange.Rows.Count, 1))
                With nowRange.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
        Next iCol

        'Call MergeDevice(nowSheet.Name)
        Call RawdataFormatByUnit(mSheet)
        
        For Each iCol In Array(6, 9, 10, 13)
            If iCol = 6 Then
                formatStr = "0.0000"
            Else
                formatStr = "G/通用格式"
            End If
            For iRow = 3 To nowSheet.UsedRange.Rows.Count
                If IsNumeric(nowSheet.Cells(iRow, iCol)) Then nowSheet.Cells(iRow, iCol).NumberFormatLocal = formatStr
            Next iRow
        Next iCol
        
        nowSheet.Activate
        nowSheet.Cells.Select
        Selection.Font.Size = 10
        Selection.Font.Name = "Arial"
        ActiveWindow.Zoom = 75
        nowSheet.Cells.Select
        Selection.Columns.AutoFit
        Selection.Rows.AutoFit
        nowSheet.Range("A3").Select
        ActiveWindow.FreezePanes = True
        Set nowSheet = Nothing
        
        Call CopySheet(ActiveItems(iSheet - 1) & "_Raw", ActiveItems(iSheet - 1) & "_Global")
        Call CopySheet(ActiveItems(iSheet - 1) & "_Raw", "Δ" & ActiveItems(iSheet - 1))
        'Worksheets(ActiveItems(iSheet - 1) & "_Raw").Visible = False
    Next iSheet
End Sub

Sub PreDataScreening()
    Dim nSheet As Integer
    Dim iSheet As Integer
    Dim nowSheet As Worksheet
    Dim siteNum As Integer
    Dim nowRow As Long, nowCol As Long
    Dim nowRange As Range, NewRange As Range
    Dim nowItem As String
        
    If PreScreen = False Then Exit Sub
    
    nSheet = UBound(ActiveItems) + 1
    siteNum = getSiteNum("Data")
    
    For iSheet = 1 To nSheet
        Set nowSheet = Worksheets(ActiveItems(iSheet - 1) & "_Global")
        nowItem = ActiveItems(iSheet - 1)
        
        Set NewRange = nowSheet.Range("A1").CurrentRegion
        
        For nowRow = 2 To NewRange.Rows.Count
            If NewRange.Cells(nowRow, 2) <> "ITEM" And NewRange.Cells(nowRow, 2) <> "" Then
                For nowCol = 1 To siteNum
                    If NewRange.Cells(nowRow, 13 + nowCol) > 0 Then
                        If NewRange.Cells(nowRow, 13 + nowCol) > mHigh Then
                            With NewRange.Cells(nowRow, 13 + nowCol)
                                .Value = ""
                                .Interior.ColorIndex = 3
                            End With
                        End If
                        If NewRange.Cells(nowRow, 13 + nowCol) < mLow Then
                            With NewRange.Cells(nowRow, 13 + nowCol)
                                .Value = ""
                                .Interior.ColorIndex = 3
                            End With
                        End If
                    Else
                        If NewRange.Cells(nowRow, 13 + nowCol) * -1 > mHigh Then
                            With NewRange.Cells(nowRow, 13 + nowCol)
                                .Value = ""
                                .Interior.ColorIndex = 3
                            End With
                        End If
                        If NewRange.Cells(nowRow, 13 + nowCol) * -1 < mLow Then
                            With NewRange.Cells(nowRow, 13 + nowCol)
                                .Value = ""
                                .Interior.ColorIndex = 3
                            End With
                        End If
                    End If
                Next nowCol
            End If
        Next nowRow
    Next iSheet

End Sub
 
Sub MergeRepeatedTks(mType As String)
    Dim nSheet As Integer, nWafer As Integer
    Dim waferList() As String, siteNum As Integer
    Dim iSheet As Integer, iWafer As Integer
    Dim nowSheet As Worksheet
    Dim nowRange As Range
    Dim nRow As Long
    Dim cAddress As String
    Dim i As Long, j As Long, k As String, m As Long
        
    If MergeTSK = False Then Exit Sub

    nSheet = UBound(ActiveItems) + 1
    siteNum = getSiteNum("Data")
    Call GetWaferList("Data", waferList)
    nWafer = UBound(waferList) + 1

    For iSheet = 1 To nSheet
        If mType = "Global" Then
            Set nowSheet = Worksheets(ActiveItems(iSheet - 1) & "_Global")
        Else
            Set nowSheet = Worksheets("Δ" & ActiveItems(iSheet - 1))
        End If
        nRow = (nowSheet.UsedRange.Rows.Count - 1) / nWafer
        
        For iWafer = 0 To (nWafer - 1)
            Set nowRange = nowSheet.Range("A" & CStr(2 + nRow * iWafer) & ":" & N2L(siteNum + 13) & CStr(1 + nRow * (iWafer + 1)))
            
            For i = 2 To nRow
                If IsError(nowRange.Cells(i, 4).Value) = True Then nowRange.Cells(i, 4).Value = "Error"
                If Not nowRange.Cells(i, 4).Value = "" Or nowRange.Cells(i, 4).Value = "Error" Then
                    cAddress = ""
                    k = 0.9
                    m = 0
                    For j = 2 To nRow
                        If j >= 1 Then
                            If getCOL2(nowRange.Cells(j, 2), "_", 4) = getCOL2(nowRange.Cells(i, 2), "_", 4) Then
                                If Not j - k = 1 Then
                                    cAddress = cAddress & "," & nowRange.Range(N2L(13 + 1) & CStr(j) & ":" & N2L(13 + siteNum) & CStr(j)).Address(False, False)
                                    k = j
                                    m = m + 1
                                ElseIf j - k = 1 Then
                                    cAddress = getCOL2(cAddress, ":", m)
                                    cAddress = cAddress & ":" & nowRange.Range(N2L(13 + siteNum) & CStr(j)).Address(False, False)
                                    k = j
                                End If
                                nowRange.Cells(j, 4).Value = ""
                                nowRange.Cells(j, 4).Interior.ColorIndex = 6
                                nowRange.Cells(j, 5).Value = ""
                                nowRange.Cells(j, 5).Interior.ColorIndex = 6
                                nowRange.Cells(j, 6).Value = ""
                                nowRange.Cells(j, 6).Interior.ColorIndex = 6
                            End If
                        End If
                    Next j
                    cAddress = Mid(cAddress, 2)
                    nowRange.Cells(i, 4).FormulaLocal = "=MEDIAN(" & cAddress & ")"
                    nowRange.Cells(i, 4).Interior.ColorIndex = 6
                    nowRange.Cells(i, 5).FormulaLocal = "=AVERAGE(" & cAddress & ")"
                    nowRange.Cells(i, 5).Interior.ColorIndex = 6
                    nowRange.Cells(i, 6).FormulaLocal = "=STDEV(" & cAddress & ")"
                    nowRange.Cells(i, 6).Interior.ColorIndex = 6
                End If
            Next i
        Next iWafer
    Next iSheet
    
End Sub

Public Sub Filterby3sigma(mType As String)
    Dim nSheet As Integer, nWafer As Integer
    Dim iSheet As Integer, iWafer As Integer
    Dim waferList() As String, siteNum As Integer
    Dim nowSheet As Worksheet
    Dim nowRange As Range
    Dim nRow As Long
    Dim iSite As Range
    Dim n As Integer
    Dim vHigh As Variant, vLow As Variant
    Dim nowRow As Long
    Dim cAddress As String
    Dim ScreenRange As Range, mScreen As Range
        
    If filterTimes = 0 Then Exit Sub
        
    nSheet = UBound(ActiveItems) + 1
    siteNum = getSiteNum("Data")
    Call GetWaferList("Data", waferList)
    nWafer = UBound(waferList) + 1
        
    For iSheet = 1 To nSheet
        If mType = "Global" Then
            Set nowSheet = Worksheets(ActiveItems(iSheet - 1) & "_Global")
        Else
            Set nowSheet = Worksheets("Δ" & ActiveItems(iSheet - 1))
        End If
        
        nRow = (nowSheet.Cells(1, 1).CurrentRegion.Rows.Count - 1) / nWafer
            
        For iWafer = 0 To nWafer - 1
            Set nowRange = nowSheet.Range("D" & CStr(2 + (iWafer) * nRow) & ":" & N2L(13 + siteNum) & CStr(1 + (iWafer + 1) * nRow))
            For Each iSite In nowRange.Range("A2:" & N2L(10 + siteNum) & CStr(nRow))
                If IsError(iSite) Then
                    If iSite.Value = CVErr(xlErrDiv0) Or iSite.Value = CVErr(xlErrNum) Then iSite.Value = ""
                End If
            Next iSite
            
            For n = 1 To filterTimes
                For nowRow = 1 To nowRange.Rows.Count
                    If IsError(nowRange.Cells(nowRow, 1)) = True Then
                    ElseIf IsError(nowRange.Cells(nowRow, 3)) = True Then
                    ElseIf nowRange.Cells(nowRow, 1) = "Med" Then
                    ElseIf nowRange.Cells(nowRow, 1) = "" Then
                    Else
                        vHigh = nowRange.Cells(nowRow, 4)
                        vLow = nowRange.Cells(nowRow, 5)
                        cAddress = Replace(Mid(nowRange.Cells(nowRow, 1).FormulaLocal, 9), ")", "")
                        Set ScreenRange = nowSheet.Range(cAddress)
                        For Each mScreen In ScreenRange
                            If mScreen > vHigh Or mScreen < vLow Or mScreen = "" Then
                                With mScreen
                                    .Value = ""
                                    .Interior.ColorIndex = 7
                                End With
                                vHigh = nowRange.Cells(nowRow, 4)
                                vLow = nowRange.Cells(nowRow, 5)
                            End If
                        Next
                    End If
                    If nowRange.Cells(nowRow, 10) < siteNum Then nowRange.Cells(nowRow, 10).Interior.ColorIndex = 4
                Next nowRow
            Next n
        Next iWafer
    Next iSheet
        
End Sub
Sub MismatchingCalculation()
    Dim nSheet As Integer, nWafer As Integer
    Dim iSheet As Integer, iWafer As Integer
    Dim rawSheet As Worksheet, mmSheet As Worksheet
    Dim mmRange As Range, dataRange As Range
    Dim nRow As Long
    Dim nowRow As Long, nowCol As Long
    Dim waferList() As String, siteNum As Integer
    Dim i As Long, j As Long
    Dim P1 As String, P2 As String
    Dim strFormula As String
    Dim PA As String, PB As String
    
    nSheet = UBound(ActiveItems) + 1
    siteNum = getSiteNum("Data")
    Call GetWaferList("Data", waferList)
    nWafer = UBound(waferList) + 1
    
    For iSheet = 1 To nSheet
        Set rawSheet = Worksheets(ActiveItems(iSheet - 1) & "_Raw")
        Set mmSheet = Worksheets("Δ" & ActiveItems(iSheet - 1))
        nRow = (mmSheet.UsedRange.Rows.Count - 1) / nWafer
        strFormula = getValueByKey("Formula", ActiveItems(iSheet - 1), 2)
        If strFormula = "" Then MsgBox ("Please define formula for """ & ActiveItems(iSheet - 1)) & """": ErrFlag = True: Exit Sub
        For iWafer = 0 To nWafer - 1
            Set mmRange = mmSheet.Range("A" & CStr(2 + nRow * iWafer) & ":" & N2L(siteNum + 13) & CStr(1 + nRow * (iWafer + 1)))
            Set dataRange = rawSheet.Range("N" & CStr(2 + nRow * iWafer) & ":" & N2L(siteNum + 13) & CStr(1 + nRow * (iWafer + 1)))
            For nowRow = 2 To dataRange.Rows.Count
                If mmRange.Cells(nowRow, 2) = "" Then
                ElseIf CInt(getCOL(mmRange.Cells(nowRow, 2), "_", 6)) Mod 2 = 1 Then
                    mmRange.Cells(nowRow, 2).Value = "Δ" & getCOL2(mmRange.Cells(nowRow, 2).Value, "_", 5)
                    For nowCol = 1 To siteNum
                        P1 = dataRange.Worksheet.Name & "!" & dataRange.Cells(nowRow, nowCol).Address(False, False)
                        P2 = dataRange.Worksheet.Name & "!" & dataRange.Cells(nowRow + 1, nowCol).Address(False, False)
                        PA = dataRange.Cells(nowRow, nowCol).Value
                        PB = dataRange.Cells(nowRow + 1, nowCol).Value
                        If ScreenPair = True Then
                            If PA = "" Or PB = "" Then
                                mmRange.Cells(nowRow, nowCol + 13).Value = ""
                            Else
                                mmRange.Cells(nowRow, nowCol + 13).FormulaLocal = "=" & Replace(Replace(strFormula, "[P1]", P1), "[P2]", P2)
                            End If
                        Else
                            mmRange.Cells(nowRow, nowCol + 13).FormulaLocal = "=" & Replace(Replace(strFormula, "[P1]", P1), "[P2]", P2)
                        End If
                    Next nowCol
                End If
            Next nowRow
        Next iWafer
        
        For nowRow = mmSheet.UsedRange.Rows.Count To 2 Step -1
            If Not Left(mmSheet.Cells(nowRow, 2), 1) = "Δ" And Not mmSheet.Cells(nowRow, 2) = "ITEM" And Not mmSheet.Cells(nowRow, 1) = "" Then
                mmSheet.Rows(nowRow).Delete
            End If
        Next nowRow
    Next iSheet

End Sub

Sub PlotChart()
    Dim nSheet As Integer
    Dim iSheet As Integer
    
    nSheet = UBound(ActiveItems) + 1
    For iSheet = 1 To nSheet
        Call PlotDeviceMMChart("Δ" & ActiveItems(iSheet - 1))
    Next iSheet
       
End Sub

Public Function PlotDeviceMMChart(mSheet As String)

    Dim nowSheet As Worksheet, mmSheet As Worksheet
    Dim dataRange As Range, tmpRange As Range
    Dim nowChart As Chart, nowSeries As Series
    Dim waferList() As String, siteNum As Integer
    Dim nowWafer As String, MosType As String
    Dim nRow As Long
    Dim nowRow As Long, sRow As Long, eRow As Long
    Dim nWafer As Integer
    Dim iWafer As Integer
    Dim tmpStr As String
    Dim i As Integer
    Dim nChart As Integer
    Dim L As Single
    Dim T As Single
    Dim W As Single
    Dim H As Single
    
    siteNum = getSiteNum("Data")
    Call GetWaferList("Data", waferList)
    nWafer = UBound(waferList) + 1
    
    Set nowSheet = AddSheet("MM_" & Mid(mSheet, 2))
    Set mmSheet = Worksheets(mSheet)
    nRow = (mmSheet.Cells(1, 1).CurrentRegion.Rows.Count - 1) / nWafer
    
    Set dataRange = mmSheet.Range("A2:" & "M" & CStr(1 + nRow))
    sRow = 2
    nChart = 0
    Call ArrangeSheet(mSheet, , 2, nRow)
    MosType = Trim(dataRange.Cells(sRow, 1).Value)
    For nowRow = 2 To nRow
        If Not dataRange.Cells(nowRow, 1).Value = MosType Then
            eRow = nowRow - 1
            nChart = nChart + 1
            For iWafer = 0 To nWafer - 1
                GoSub PlotSub
            Next iWafer
            sRow = nowRow
            MosType = dataRange.Cells(sRow, 1).Value
        End If
    Next nowRow
    If eRow = 0 Or nowRow > nRow Then
        eRow = nRow
        nChart = nChart + 1
        For iWafer = 0 To nWafer - 1
            GoSub PlotSub
        Next iWafer
    End If
    
Exit Function
           
PlotSub:

    W = 400
    H = 300
    L = 10 + 400 * Int((nChart - 1) / 4)
    T = 10 + 300 * Int((nChart - 1) Mod 4)
    
    
    Set tmpRange = mmSheet.Range(((sRow + 1) + iWafer * nRow) & ":" & ((eRow + 1)) + iWafer * nRow)
    If iWafer = 0 Then Set nowChart = myCreateChart(nowSheet, xlXYScatter, L, T, W, H)
    Set nowSeries = nowChart.SeriesCollection.NewSeries
    nowSeries.XValues = tmpRange.Columns(11)    '1/sqrt(WL)
    nowSeries.Values = tmpRange.Columns(6)      'Stdev
    nowSeries.Name = "#" & waferList(iWafer)
    
    With nowChart
        .chartType = xlXYScatter
        .HasTitle = True
        .ChartTitle.Characters.Text = MosType & " MM Chart of " & mSheet
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "1/sqrt(WL)"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "σ" & mSheet
    End With
    
    Call SetChartStyle(nowChart)
    
    With nowChart.FullSeriesCollection(iWafer + 1)
        .Trendlines.Add
        .Trendlines(1).Intercept = 0
        .Trendlines(1).DisplayEquation = True
        .Trendlines(1).DisplayRSquared = True
        .Trendlines(1).Format.Line.Weight = 1.5
        .Trendlines(1).Format.Line.ForeColor.RGB = Worksheets("PlotSetup").ChartObjects(1).Chart.SeriesCollection(iWafer + 1).Points(1).MarkerBackgroundColor
        .Trendlines(1).DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Worksheets("PlotSetup").ChartObjects(1).Chart.SeriesCollection(iWafer + 1).Points(1).MarkerBackgroundColor
    End With

Return


    
End Function

Function ArrangeSheet(mSheet As String, Optional ByVal mCol As Integer = 1, Optional ByVal startRow As Long = 1, Optional Period As Long = 0, Optional ByVal Mode As String = 1)

    Dim nowSheet As Worksheet
    Dim nRow As Long
    Dim nowRow As Long
    Dim m As Integer
    
    Set nowSheet = Worksheets(mSheet)
    nRow = nowSheet.Cells(startRow, mCol).CurrentRegion.Rows.Count
    nowSheet.Activate
    
    If Period <> 0 Then
        For nowRow = nRow + 1 - Period To startRow Step (-1 * Period)
            Rows(nowRow).Insert Shift:=xlDown
            m = m + 1
            Rows(nowRow + 1).AutoFilter
            Worksheets(mSheet).AutoFilter.Sort.SortFields.Add key:=Range(N2L(mCol) & CStr(nowRow + 1)), _
                SortOn:=xlSortOnValues, _
                order:=Mode, _
                DataOption:=xlSortNormal
            With Worksheets(mSheet).AutoFilter.Sort
                .header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Rows(nowRow + 1).AutoFilter
        Next nowRow
        
        For nowRow = nRow + 1 - Period + m To startRow Step (-1 * (Period + 1))
            Rows(nowRow - 1).Delete
        Next nowRow
    Else
        Rows(startRow).AutoFilter
        Worksheets(mSheet).AutoFilter.Sort.SortFields.Add key:=Range(N2L(mCol) & CStr(nowRow + 1)), _
            SortOn:=xlSortOnValues, _
            order:=xlAscending, _
            DataOption:=xlSortNormal
        With Worksheets(mSheet).AutoFilter.Sort
            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        Rows(startRow).AutoFilter
    End If

End Function
