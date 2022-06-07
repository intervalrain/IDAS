Option Explicit

Public Sub LotRawData()
    Dim nowSheet As Worksheet
    Dim nowCol As Integer, nowRow As Integer
    Dim itemRange As Range
    Dim mSheet As String
    Dim tmpStr As String
    Dim i As Integer
   
    Set nowSheet = Worksheets("SPEC_list")
    For nowCol = 1 To nowSheet.UsedRange.Columns.Count
        mSheet = Trim(getCOL(nowSheet.Cells(1, nowCol), ":", 1))
        tmpStr = Trim(getCOL(nowSheet.Cells(1, nowCol), ":", 2))
        If Len(mSheet) > 23 Then mSheet = Left(mSheet, 23): nowSheet.Cells(1, nowCol).Value = mSheet & ":" & tmpStr
        If mSheet = "" Then Exit For
        For nowRow = nowSheet.UsedRange.Rows.Count To 1 Step -1
            If Trim(nowSheet.Cells(nowRow, nowCol)) <> "" Then Exit For
        Next nowRow
        If nowRow > 1 Then
            Set itemRange = nowSheet.Range(nowSheet.Cells(2, nowCol), nowSheet.Cells(nowRow, nowCol))
            Call LotRawDataSub(mSheet, itemRange)
        End If
    Next nowCol
    Set nowSheet = Nothing
End Sub

Public Function LotRawDataSub(mSheet As String, itemRange As Range)
   Dim waferList() As String
   Dim siteNum As Integer
   Dim subWafer() As String
   Dim maxWafer As Integer
   Dim i As Integer, j As Integer
   Dim SubRawName As String
   Dim tmpAry() As String
   Dim tmpStr As String
   Dim tmpList() As String
   
   Call GetWaferList(dSheet, waferList)
   siteNum = getSiteNum(dSheet)
   'grouping sequence adjust
   If IsExistSheet("Grouping") Then
        For i = 2 To Worksheets("Grouping").Cells(1, 1).CurrentRegion.Rows.Count
           tmpStr = tmpStr & "," & Worksheets("Grouping").Cells(i, 1).Value
        Next i
        tmpStr = Mid(tmpStr, 2)
        tmpAry = Split(tmpStr, ",")
        ReDim tmpList(0) As String
        For i = 0 To UBound(tmpAry)
           If IsInArray(tmpAry(i), waferList) Then
              tmpList(UBound(tmpList)) = tmpAry(i)
              ReDim Preserve tmpList(UBound(tmpList) + 1) As String
           End If
        Next i
        If UBound(tmpList) > 0 Then ReDim Preserve tmpList(UBound(tmpList) - 1) As String
        waferList = tmpList
   End If
   
   

   If (UBound(waferList) + 1) * siteNum > 1575 Then
      maxWafer = 1575 \ siteNum
      ReDim subWafer(0)
      For i = 0 To UBound(waferList)
         subWafer(UBound(subWafer)) = waferList(i)
         If UBound(subWafer) + 1 = maxWafer Then
            If i < maxWafer Then
               SubRawName = mSheet & "_Raw"
            Else
               SubRawName = mSheet & "_Raw_" & CStr(i \ maxWafer)
            End If
            Call RawList(SubRawName, subWafer, siteNum, itemRange)
            ReDim subWafer(0)
         Else
            ReDim Preserve subWafer(UBound(subWafer) + 1)
         End If
      Next i
      If subWafer(0) <> "" Then
         ReDim Preserve subWafer(UBound(subWafer) - 1)
         If i <= maxWafer Then
               SubRawName = mSheet & "_Raw"
            Else
               SubRawName = mSheet & "_Raw_" & CStr(i \ maxWafer)
            End If
            Call RawList(SubRawName, subWafer, siteNum, itemRange)
      End If
   Else
      Call RawList(mSheet & "_Raw", waferList, siteNum, itemRange)
   End If
End Function

Public Function GetWaferList(mSheet As String, ByRef mWaferlist() As String)
    Dim i As Long
    Dim tmpStr As String
    ReDim mWaferlist(0)
    On Error GoTo myError
    For i = 0 To UBound(WaferArray, 2)
        On Error GoTo 0
        If WaferArray(1, i) <> "NO" Then
            If mWaferlist(0) <> "" Then ReDim Preserve mWaferlist(UBound(mWaferlist) + 1)
            mWaferlist(UBound(mWaferlist)) = WaferArray(0, i)
        End If
    Next i
    

    Exit Function
myError:
    Call SetWaferRange
    Resume
End Function

Public Function RawList(mSheet As String, waferList() As String, siteNum As Integer, itemRange As Range)
    Dim specRange As Range
    Dim waferNum As Integer, LotName As String, ProductID As String
    Dim iCol As Long, iRow As Long
    Dim i As Long, j As Long, m As Long
    Dim nowRange As Range
    Dim nowParameter As String
    Dim reValue As Variant
    Dim nowSheet As Worksheet
    Dim paraIndex As Integer
    Dim nowRow As Integer
    Dim objPara As specInfo
    Dim nowCondition As FormatCondition
    Dim ParaArray() As String
    Dim valueArray() As String
    Dim strFormula As String
    Dim LumpFunction As String
    Dim iSite As Integer
    Dim exStr As String
    Dim nowCell As Range
    
    On Error Resume Next
    
    If IsExistSheet("SPEC_List") Then exStr = getExStr(getCOL(mSheet, "_Raw", 1))
       
    Set nowSheet = AddSheet(mSheet)
    nowSheet.Activate
    
    'Header
    With nowSheet.Range("A2:F2")
        .Cells(1, 1) = "DEVICE"
        .Cells(1, 2) = "ITEM"
        .Cells(1, 3) = "UNIT"
        .Cells(1, 4) = "SPEC Lo"
        .Cells(1, 5) = "TARGET"
        .Cells(1, 6) = "SPEC Hi"
    End With
    waferNum = UBound(waferList) + 1
    ProductID = Trim(Worksheets(dSheet).Cells(2, 2))
    LotName = Trim(Worksheets(dSheet).Cells(3, 2))
    If Left(LotName, 1) = ":" Then LotName = Mid(LotName, 2)
    If Left(ProductID, 1) = ":" Then ProductID = Mid(ProductID, 2)
    With nowSheet.Range("A1:F1")
        .Cells(1, 1) = "PROD:"
        .Cells(1, 2) = ProductID
        .Range("B1").Font.Color = RGB(255, 0, 0)
        .Cells(1, 3) = "LOT:"
        .Cells(1, 4) = LotName
        .Range("D1").Font.Color = RGB(255, 0, 0)
        .Cells(1, 5) = "WAFERs:"
        .Cells(1, 6) = waferNum
        .Range("F1").Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

        Set specRange = Worksheets(mSheet).UsedRange

    iRow = 1
    iCol = specRange.Columns.Count
    For i = 0 To UBound(waferList)
        For j = 1 To siteNum
            With nowSheet
                .Cells(iRow, iCol + i * siteNum + j) = waferList(i)
                .Cells(iRow + 1, iCol + i * siteNum + j) = CStr(j)
            End With
        Next j
    Next i
    
    'Data
    For paraIndex = 1 To itemRange.Rows.Count
        nowRow = paraIndex + 2
        nowParameter = Trim(itemRange.Cells(paraIndex, 1))
        If Left(nowParameter, 1) = "=" Then nowParameter = "'" & nowParameter
        objPara = getSPECInfo(nowParameter)
        nowSheet.Cells(nowRow, 1) = objPara.mDevice
        nowSheet.Cells(nowRow, 2) = objPara.mPara
        nowSheet.Cells(nowRow, 3) = objPara.mUnit
        nowSheet.Cells(nowRow, 4) = objPara.mLow
        nowSheet.Cells(nowRow, 5) = objPara.mTarget
        nowSheet.Cells(nowRow, 6) = objPara.mHigh

        For i = 0 To UBound(waferList)
            Set reValue = getRangeByPara(waferList(i), getCOL(nowParameter, ":", 1), siteNum)

            If Not reValue Is Nothing Then
                Set nowRange = reValue

                specRange.Range(Cells(nowRow, iCol + i * siteNum + 1), Cells(nowRow, iCol + (i + 1) * siteNum)).Value = nowRange.Value
                If objPara.mFAC <> 1 Then
                    For m = 1 To siteNum
                        If specRange.Cells(nowRow, iCol + i * siteNum + m) <> "" Then
                            specRange.Cells(nowRow, iCol + i * siteNum + m) = objPara.mFAC * specRange.Cells(nowRow, iCol + i * siteNum + m)
                        End If
                    Next m
                End If
                
                'Eliminate Out-Spec Data
                If IsKey(exStr, "WID") Then
                    If Not IsEmpty(objPara.mHigh) Then
                        For Each nowCell In nowRange
                            If nowCell.Value > Val(objPara.mHigh) Then nowCell.Value = ""
                        Next nowCell
                    End If
                    If Not IsEmpty(objPara.mLow) Then
                        For Each nowCell In nowRange
                            If nowCell.Value < Val(objPara.mLow) Then nowCell.Value = ""
                        Next nowCell
                    End If
                Else
                    'Highlight Out-Spec Data
                    If Not IsEmpty(objPara.mLow) Or Not IsEmpty(objPara.mHigh) Then
                        Set nowRange = specRange.Range(Cells(nowRow, iCol + i * siteNum + 1), Cells(nowRow, iCol + (i + 1) * siteNum))
                        With nowRange
                            .FormatConditions.Delete
                            If Not IsEmpty(objPara.mHigh) Then
                                Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & "F" & CStr(nowRow))
                                nowCondition.Font.ColorIndex = 3
                            End If
                            If Not IsEmpty(objPara.mLow) Then
                                Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & "D" & CStr(nowRow))
                                nowCondition.Font.ColorIndex = 4
                            End If
                        End With
                    End If
                End If
            'Derive Value by Calculating rawData
            ElseIf Left(nowParameter, 2) = "'=" Then
                strFormula = getCOL(nowParameter, "=", 2)
                LumpFunction = FormulaParse(strFormula, ParaArray)
                
                With specRange.Range(Cells(nowRow, iCol + i * siteNum + 1), Cells(nowRow, iCol + (i + 1) * siteNum))
                    For iSite = 1 To siteNum
                        If LumpFunction = "" Then
                           Call FormulaValue(ParaArray, valueArray, waferList(i), iSite, objPara)
                        Else
                           Call FormulaValue(ParaArray, valueArray, waferList(i), siteNum, objPara, LumpFunction)
                        End If
                        .Cells(1, iSite) = FormulaEval(strFormula, ParaArray, valueArray)
                    Next iSite
                End With
            End If
        Next i
    Next paraIndex

    'Set Wafer Split Line
    For i = 0 To UBound(waferList) + 1
        Set nowRange = nowSheet.Columns(iCol + i * siteNum)
            With nowRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
    Next i
   
    'Set Report Style
    Call RangeStyle(nowSheet, itemRange)
    Call MergeDevice(nowSheet.Name)
   
    'Set Number Format
    Call RawdataFormatByUnit(mSheet)
    Call HighLightSBG(mSheet)

    'Define Data Range from Raw-sheet
    Call SetRawWaferRange(mSheet)
   
    'Set Sheet Presentation Format
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

End Function

Public Sub LotSummary()
    Dim nowSheet As Worksheet
    Dim nowCol As Integer, nowRow As Integer
    Dim itemRange As Range
    Dim mSheet As String
    Dim tmpStr As String
    Dim i As Integer
       
    Set nowSheet = Worksheets("SPEC_list")
    For nowCol = 1 To nowSheet.UsedRange.Columns.Count
        mSheet = Trim(getCOL(nowSheet.Cells(1, nowCol), ":", 1))
        If Len(mSheet) > 23 Then mSheet = Left(mSheet, 23): nowSheet.Cells(1, nowCol).Value = mSheet & ":" & tmpStr
        If mSheet = "" Then Exit For
        For nowRow = nowSheet.UsedRange.Rows.Count To 1 Step -1
            If Trim(nowSheet.Cells(nowRow, nowCol)) <> "" Then Exit For
        Next nowRow
        If nowRow > 1 Then
            Set itemRange = nowSheet.Range(nowSheet.Cells(2, nowCol), nowSheet.Cells(nowRow, nowCol))
            Call LotSummarySub(mSheet, itemRange)
        End If
    Next nowCol
    Set nowSheet = Nothing
End Sub

Public Function LotSummarySub(mSheet As String, itemRange As Range)
    Dim specRange As Range
    Dim dataRange As Range
    Dim waferList() As String
    Dim iCol As Long, nowRow As Long, iRow As Long, nowCol As Long
    Dim outSpec As Integer
    Dim siteNum As Integer
    Dim ProductID As String
    Dim LotName As String
    Dim nowRange As Range
    Dim reValue As Variant
    Dim nowParameter As String
    Dim nowSheet As Worksheet
    Dim rawSheet As Worksheet
    Dim SumMode() As String
    Dim sItem
    Dim Headerarray
    Dim hcount As Integer
    Dim i As Integer, j As Integer, k As Integer, c As Integer
    Dim paraIndex As Long
    Dim objPara As specInfo
    Dim nowCondition As FormatCondition
    Dim ynL90G As Boolean
    Dim tmpStr As String
    Dim waferNum As Integer
    Dim exStr As String
    Dim mCPK As String
    Dim ynLargeSPEC As Boolean
    Dim tempA
    Dim tempRow() As Integer
    Dim tmpAry() As String
    Dim tmpList() As String
   
    On Error Resume Next
   
    If Worksheets(SSheet).UsedRange.Rows.Count > 1000 Then ynLargeSPEC = True
    Set nowSheet = AddSheet(mSheet & "_Summary")
    nowSheet.Activate
    
    Call GetWaferList(dSheet, waferList)
    siteNum = getSiteNum(dSheet)
   
    With nowSheet.Range("A2:F2")
        .Cells(1, 1) = "DEVICE"
        .Cells(1, 2) = "ITEM"
        .Cells(1, 3) = "UNIT"
        .Cells(1, 4) = "SPEC Lo"
        .Cells(1, 5) = "TARGET"
        .Cells(1, 6) = "SPEC Hi"
    End With
   
    waferNum = UBound(waferList) + 1
    ProductID = Trim(Worksheets(dSheet).Cells(2, 2))
    LotName = Trim(Worksheets(dSheet).Cells(3, 2))
    If Left(LotName, 1) = ":" Then LotName = Mid(LotName, 2)
    If Left(ProductID, 1) = ":" Then ProductID = Mid(ProductID, 2)
    With nowSheet.Range("A1:F1")
        .Cells(1, 1) = "PROD:"
        .Cells(1, 2) = ProductID
        .Range("B1").Font.Color = RGB(255, 0, 0)
        .Cells(1, 3) = "LOT:"
        .Cells(1, 4) = LotName
        .Range("D1").Font.Color = RGB(255, 0, 0)
        .Cells(1, 5) = "WAFERs:"
        .Cells(1, 6) = waferNum
        .Range("F1").Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    'Summary Presentation Mode Setting
    If IsExistSheet("SPEC_List") Then exStr = getExStr(mSheet)
    Headerarray = Array("Median", "Average", "Sigma", "Yield")
    SumMode = Split(exStr, "&")
    
    mCPK = ""
    If IsKey(exStr, "CPK133") Then mCPK = "1.33"
    If IsKey(exStr, "CPK150") Then mCPK = "1.5"
    If IsKey(exStr, "CPK167") Then mCPK = "1.67"
    If IsKey(exStr, "comp") Then c = 1
    
    For Each sItem In SumMode
        Call SummaryModeSetting(sItem, Headerarray, hcount)
    Next sItem
    
    If exStr = "" Then Call SummaryModeSetting(sItem, Headerarray, hcount)

    iCol = nowSheet.UsedRange.Columns.Count
    nowRow = 1
    Set specRange = nowSheet.UsedRange

    'Grouping
    If IsExistSheet("Grouping") Then
    
        For i = 2 To Worksheets("Grouping").Cells(1, 1).CurrentRegion.Rows.Count
            tmpStr = tmpStr & "," & Worksheets("Grouping").Cells(i, 1).Value
        Next i
        tmpStr = Mid(tmpStr, 2)
        tmpAry = Split(tmpStr, ",")
        ReDim tmpList(0) As String
        For i = 0 To UBound(tmpAry)
            If IsInArray(tmpAry(i), waferList) Then
                tmpList(UBound(tmpList)) = tmpAry(i)
                ReDim Preserve tmpList(UBound(tmpList) + 1) As String
            End If
        Next i
        If UBound(tmpList) > 0 Then ReDim Preserve tmpList(UBound(tmpList) - 1) As String
        waferList = tmpList
    
        Dim waferId() As String
        Dim groupId() As String
        ReDim waferId(Worksheets("Grouping").Cells(1, 1).CurrentRegion.Rows.Count - 2)
        ReDim groupId(Worksheets("Grouping").Cells(1, 1).CurrentRegion.Rows.Count - 2)
        For i = 2 To Worksheets("Grouping").Cells(1, 1).CurrentRegion.Rows.Count
            waferId(i - 2) = Worksheets("Grouping").Cells(i, 1).Value
            groupId(i - 2) = Worksheets("Grouping").Cells(i, 2).Value
        Next i
        For j = 0 To UBound(waferId)
            Call SetGrouping(mSheet & "_Raw", waferId(j), groupId(j))
        Next j
        waferList = groupId
        waferNum = UBound(waferList) + 1
    End If

    'Set Header
    For j = 0 To hcount - 1
        For i = 1 To UBound(waferList) + 1
            nowSheet.Cells(nowRow, iCol + j * waferNum + i) = Replace(waferList(i - 1), "PLUS", "+")
            nowSheet.Cells(nowRow + 1, iCol + j * waferNum + i) = Headerarray(j)
        Next i
    Next j
   
    For paraIndex = 1 To itemRange.Rows.Count
        nowRow = paraIndex + 2
        nowParameter = itemRange.Cells(paraIndex, 1)
        If Left(nowParameter, 1) = "=" Then nowParameter = "'" & nowParameter
        objPara = getSPECInfo(nowParameter)
            
        nowSheet.Cells(nowRow, 1) = objPara.mDevice
        nowSheet.Cells(nowRow, 2) = objPara.mPara
        nowSheet.Cells(nowRow, 3) = objPara.mUnit
        nowSheet.Cells(nowRow, 4) = objPara.mLow
        nowSheet.Cells(nowRow, 5) = objPara.mTarget
        nowSheet.Cells(nowRow, 6) = objPara.mHigh
            
        For i = 0 To UBound(waferList)
            If Not InStr(1, nowParameter, ";") > 1 Then
                Set reValue = getRawRangeByRow(mSheet, waferList(i), paraIndex + 2)
            Else
                If InStr(nowParameter, ":") Then nowParameter = getCOL(nowParameter, ":", 1)
                tempA = Split(nowParameter, ";")
                ReDim tempRow(UBound(tempA))
                ReDim reValue(UBound(tempA))
                For j = 0 To UBound(tempA)
                    For k = 1 To itemRange.Rows.Count
                        If tempA(j) = itemRange.Cells(k, 1).Value Then tempRow(j) = k: Exit For
                    Next k
                    reValue(j) = getRawRangeByRow(mSheet, waferList(i), tempRow(j) + 2)
                Next j
            End If
            If Not reValue Is Nothing Then
                If Not InStr(1, nowParameter, ";") > 1 Then
                    Set nowRange = reValue
                Else
                    Set nowRange = Nothing
                End If
                'Fill Median, Average, Sigma
                Dim bslRange As Range
                With nowSheet.Range(N2L(iCol + 1 + i) & CStr(nowRow) & ":" & N2L(iCol + waferNum * hcount) & CStr(nowRow))
                    If Not InStr(1, nowParameter, ";") > 1 Then
                        .Cells(1, 1 + (0 + c) * waferNum) = Application.WorksheetFunction.Median(nowRange)
                        .Cells(1, 1 + (1 + c) * waferNum) = Application.WorksheetFunction.Average(nowRange)
                        If nowRange.Columns.Count > 1 Then
                        .Cells(1, 1 + (2 + c) * waferNum) = Application.WorksheetFunction.StDev(nowRange)
                        End If
                    Else
                        .Cells(1, 1 + (0 + c) * waferNum) = Application.WorksheetFunction.Median(reValue)
                        .Cells(1, 1 + (1 + c) * waferNum) = Application.WorksheetFunction.Average(reValue)
                        If nowRange.Columns.Count > 1 Then
                        .Cells(1, 1 + (2 + c) * waferNum) = Application.WorksheetFunction.StDev(reValue)
                        End If
                    End If
                    
                    'Trend
                    Dim refTarget As Double
                    Dim strFormula As String
                    Dim LumpFunction As String
                    Dim ParaArray() As String
                    If InStr(nowParameter, "Trend") Then
                        If InStr(UCase(itemRange.Cells(paraIndex - 2)), "=LN") Then
                            strFormula = getCOL(itemRange.Cells(paraIndex - 2, 1), "=", 2)
                            LumpFunction = FormulaParse(strFormula, ParaArray)
                            objPara = getSPECInfo(ParaArray(0))
                            refTarget = Application.Evaluate(Replace(strFormula, ParaArray(0), objPara.mTarget))
                            .Cells(1, 1 + c * waferNum) = Application.WorksheetFunction.Forecast(refTarget, getRawRangeByRow(mSheet, waferList(i), paraIndex + 1), getRawRangeByRow(mSheet, waferList(i), paraIndex))
                        ElseIf InStr(UCase(itemRange.Cells(paraIndex - 1)), "=LN") Then
                            strFormula = getCOL(itemRange.Cells(paraIndex - 1, 1), "=", 2)
                            LumpFunction = FormulaParse(strFormula, ParaArray)
                            objPara = getSPECInfo(itemRange.Cells(paraIndex - 2))
                            refTarget = objPara.mTarget
                            .Cells(1, 1 + c * waferNum) = Exp(Application.WorksheetFunction.Forecast(refTarget, getRawRangeByRow(mSheet, waferList(i), paraIndex + 1), getRawRangeByRow(mSheet, waferList(i), paraIndex)))
                        Else
                            objPara = getSPECInfo(itemRange.Cells(paraIndex - 2))
                            refTarget = objPara.mTarget
                            .Cells(1, 1 + c * waferNum) = Application.WorksheetFunction.Forecast(refTarget, getRawRangeByRow(mSheet, waferList(i), paraIndex + 1), getRawRangeByRow(mSheet, waferList(i), paraIndex))
                        End If
                        
                        'R SQUARE
                        If IsKey(exStr, "R2") Then
                            .Cells(1, 1 + (Application.Match("R Square", Headerarray, 0) - 1) * waferNum).Value = WorksheetFunction.Correl(getRawRangeByRow(mSheet, waferList(i), paraIndex + 1), getRawRangeByRow(mSheet, waferList(i), paraIndex)) ^ 2
                            .Cells(1, 1 + (Application.Match("R Square", Headerarray, 0) - 1) * waferNum).NumberFormatLocal = "0.00"
                        End If
                    End If
                    
                    'Fill Shift to BSL(comp)
                    If i = 0 Then Set bslRange = .Cells(1, 1 + (0 + c) * waferNum)
                    If IsKey(exStr, "comp") And bslRange.Value <> "" And i <> 0 Then
                        Select Case getDiffCase(IIf(InStr(nowParameter, ":"), getCOL(nowParameter, ":", 2), nowParameter))
                            Case 1
                                .Cells(1, 1).Value = "=" & .Cells(1, 1 + waferNum).Address(False, False) & "-" & bslRange.Address(False, False)
                                .Cells(1, 1).NumberFormatLocal = "0.00"
                            Case 2
                                .Cells(1, 1).Value = "=" & .Cells(1, 1 + waferNum).Address(False, False) & "/" & bslRange.Address(False, False) & "-1"
                                .Cells(1, 1).NumberFormatLocal = "0.00%"
                            Case 3
                                .Cells(1, 1).Value = "=" & .Cells(1, 1 + waferNum).Address(False, False) & "/" & bslRange.Address(False, False)
                                .Cells(1, 1).NumberFormatLocal = "0.00x"
                            Case 4
                                .Cells(1, 1).Value = "=" & .Cells(1, 1 + waferNum).Address(False, False) & "-" & bslRange.Address(False, False)
                                .Cells(1, 1).NumberFormatLocal = nowSheet.Cells(nowRow, 5).NumberFormatLocal
                        End Select
                    End If
                    'Fill Yield
                    If .Cells(1, 1 + c * waferNum) <> "" Then
                        If InStr(UCase(nowParameter), "TREND") Then objPara = getSPECInfo(nowSheet.Cells(.Cells(1, 1).row, 2))
                        outSpec = 0
                        If Not IsEmpty(objPara.mLow) Or Not IsEmpty(objPara.mHigh) Then
                            If Not IsEmpty(objPara.mLow) Then _
                                outSpec = outSpec + Application.WorksheetFunction.CountIf(nowRange, "<" & objPara.mLow)
                            If Not IsEmpty(objPara.mHigh) Then _
                                outSpec = outSpec + Application.WorksheetFunction.CountIf(nowRange, ">" & objPara.mHigh)
                            .Cells(1, 1 + (3 + c) * waferNum) = Format((siteNum - outSpec) / siteNum, "0.00%")
                            If outSpec <> 0 Then .Cells(1, 1 + 3 * waferNum).Font.ColorIndex = 3
                            
                            'Highlight outSpec Value
                            With .Cells(1, 1 + c * waferNum)
                                .FormatConditions.Delete
                                If Not IsEmpty(objPara.mHigh) Then
                                    Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & "F" & "" & CStr(nowRow))
                                    nowCondition.Font.ColorIndex = 3
                                End If
                                If Not IsEmpty(objPara.mLow) Then
                                    Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & "D" & "" & CStr(nowRow))
                                    nowCondition.Font.ColorIndex = 4
                                End If
                            End With
                            With .Range(N2L(1 + (1 + c) * waferNum) & "1")
                                .FormatConditions.Delete
                                If Not IsEmpty(objPara.mHigh) Then
                                    Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & "F" & "" & CStr(nowRow))
                                    nowCondition.Font.ColorIndex = 3
                                End If
                                If Not IsEmpty(objPara.mLow) Then
                                    Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & "D" & "" & CStr(nowRow))
                                    nowCondition.Font.ColorIndex = 4
                                End If
                            End With
                        End If
                        
                        'Diff
                        Dim isTrend As Integer
                        If InStr(UCase(nowParameter), "TREND") Then isTrend = 1 Else isTrend = 0
                        If IsKey(exStr, "diff") Then
                            Select Case getDiffCase(IIf(InStr(nowParameter, ":"), getCOL(nowParameter, ":", 2), nowParameter))
                                Case 1
                                    .Cells(1, 1 + (Application.Match("Diff", Headerarray, 0) - 1) * waferNum).Formula = "=IF($E" & CStr(nowRow - isTrend) & "="""",""""," & N2L(iCol + i + 1 + c * waferNum) & CStr(nowRow) & "-$E" & CStr(nowRow - isTrend) & ")"
                                    .Cells(1, 1 + (Application.Match("Diff", Headerarray, 0) - 1) * waferNum).NumberFormatLocal = nowSheet.Cells(nowRow, 5).NumberFormatLocal
                                Case 2
                                    .Cells(1, 1 + (Application.Match("Diff", Headerarray, 0) - 1) * waferNum).Formula = "=IF($E" & CStr(nowRow - isTrend) & "="""",""""," & N2L(iCol + i + 1 + c * waferNum) & CStr(nowRow) & "/$E" & CStr(nowRow - isTrend) & "-1" & ")"
                                    .Cells(1, 1 + (Application.Match("Diff", Headerarray, 0) - 1) * waferNum).NumberFormatLocal = "0.00%"
                                Case 3
                                    .Cells(1, 1 + (Application.Match("Diff", Headerarray, 0) - 1) * waferNum).Formula = "=IF($E" & CStr(nowRow - isTrend) & "="""",""""," & N2L(iCol + i + 1 + c * waferNum) & CStr(nowRow) & "/$E" & CStr(nowRow - isTrend) & ")"
                                    .Cells(1, 1 + (Application.Match("Diff", Headerarray, 0) - 1) * waferNum).NumberFormatLocal = "0.000""x"""
                                Case 4
                                    .Cells(1, 1 + (Application.Match("Diff", Headerarray, 0) - 1) * waferNum).NumberFormatLocal = nowSheet.Cells(nowRow, 5).NumberFormatLoca
                            End Select
                        End If
                        
                        'WID
                        If IsKey(exStr, "WID") Then
                            .Cells(1, 1 + (Application.Match("3 Sigma/Median", Headerarray, 0) - 1) * waferNum).Formula = "=3*" & N2L(iCol + i + 1 + 2 * waferNum) & CStr(nowRow) & "/" & N2L(iCol + i + 1) & CStr(nowRow)
                        End If
                        
                        'OUTSPEC
                        If IsKey(exStr, "OUTSPEC") Then
                            .Cells(1, 1 + (Application.Match("OutSpec", Headerarray, 0) - 1) * waferNum) = outSpec
                        End If
                        
                        'ZTable
                        If IsKey(exStr, "ZTable") Then
                            If Not IsEmpty(objPara.mLow) And Not IsEmpty(objPara.mHigh) And Not IsEmpty(objPara.mTarget) Then
                                If UCase(Left(IIf(InStr(nowParameter, ":"), getCOL(nowParameter, ":", 2), nowParameter), 3)) = "IOF" Then
                                    .Cells(1, 1 + (Application.Match("Z Value", Headerarray, 0) - 1) * waferNum) = "=(LN(" & N2L(iCol + i + 1) & CStr(nowRow) & ")-LN($E" & CStr(nowRow) & "))/((" & "LN($F" & CStr(nowRow) & ")-LN($E" & CStr(nowRow) & "))/3)"
                                Else
                                    .Cells(1, 1 + (Application.Match("Z Value", Headerarray, 0) - 1) * waferNum) = "=(" & N2L(iCol + i + 1 + UBound(waferList) + 1) & CStr(nowRow) & "-" & "$E" & CStr(nowRow) & ")/(($F" & CStr(nowRow) & "-$D" & CStr(nowRow) & ")/6)"
                                End If
                                .Cells(1, 1 + (Application.Match("Z Value", Headerarray, 0) - 1) * waferNum).NumberFormat = "0.00"
                                
                                Select Case .Cells(1, 1 + (Application.Match("Z Value", Headerarray, 0) - 1) * waferNum).Value
                                    Case -1.5 To 1.5: .Cells(1, 1 + (Application.Match("Z Value", Headerarray, 0) - 1) * waferNum).Interior.ColorIndex = 4
                                    Case -3 To 3: .Cells(1, 1 + (Application.Match("Z Value", Headerarray, 0) - 1) * waferNum).Interior.ColorIndex = 6
                                    Case -4.5 To 4.5: .Cells(1, 1 + (Application.Match("Z Value", Headerarray, 0) - 1) * waferNum).Interior.ColorIndex = 45
                                    Case Else: .Cells(1, 1 + (Application.Match("Z Value", Headerarray, 0) - 1) * waferNum).Interior.ColorIndex = 3
                                End Select
                            End If
                        End If
                        
                        'Max
                        If IsKey(exStr, "Max") Then
                            If Not InStr(1, nowParameter, ";") > 1 Then
                                .Cells(1, 1 + (Application.Match("Max", Headerarray, 0) - 1) * waferNum) = Application.WorksheetFunction.Max(nowRange)
                                .Cells(1, 1 + (Application.Match("Min", Headerarray, 0) - 1) * waferNum) = Application.WorksheetFunction.Min(nowRange)
                            Else
                                .Cells(1, 1 + (Application.Match("Max", Headerarray, 0) - 1) * waferNum) = Application.WorksheetFunction.Max(reValue)
                                .Cells(1, 1 + (Application.Match("Min", Headerarray, 0) - 1) * waferNum) = Application.WorksheetFunction.Min(reValue)
                            End If
                        End If
                        
                        'Kvalue  (Xbar/(spec/2))
                        If IsKey(exStr, "kvalue") Then
                            If Not IsEmpty(objPara.mLow) And Not IsEmpty(objPara.mHigh) And Not IsEmpty(objPara.mTarget) Then
                                .Cells(1, 1 + (Application.Match("K Value", Headerarray, 0) - 1) * waferNum).Formula = "=IF((" & "$" & N2L(iCol + i + 1) & CStr(nowRow) & "-$E" & CStr(nowRow) & ")>=0," & _
                                                                                                                                          "(" & "$" & N2L(iCol + i + 1) & CStr(nowRow) & "-$E" & CStr(nowRow) & ")/(" & "$F" & CStr(nowRow) & "-$E" & CStr(nowRow) & ")," & _
                                                                                                                                          "(" & "$" & N2L(iCol + i + 1) & CStr(nowRow) & "-$E" & CStr(nowRow) & ")/(" & "$E" & CStr(nowRow) & "-$D" & CStr(nowRow) & "))"
                                .Cells(1, 1 + (Application.Match("K Value", Headerarray, 0) - 1) * waferNum).NumberFormatLocal = "0.00"
                                'Add Kvalue summary
                                If paraIndex = itemRange.Rows.Count Then
                                    .Cells(2, 1 + (Application.Match("K Value", Headerarray, 0) - 1) * waferNum).Formula = "=SUM(COUNTIF(" & N2L(iCol + i + 1 + (Application.Match("K Value", Headerarray, 0) - 1) * waferNum) & CStr(3) & ":" & N2L(iCol + i + 1 + (Application.Match("K Value", Headerarray, 0) - 1) * waferNum) & CStr(2 + paraIndex) & ",{"">=-1"","">1""})*{1,-1}) & ""/"" & COUNT(" & N2L(iCol + i + 1 + (Application.Match("K Value", Headerarray, 0) - 1) * waferNum) & CStr(3) & ":" & N2L(iCol + i + 1 + (Application.Match("K Value", Headerarray, 0) - 1) * waferNum) & CStr(2 + paraIndex) & ")"
                                End If
                            End If
                        End If
                                                          
                        'CPK, Ca, Cp
                        If mCPK <> "" Then
                            'CPK
                            With .Cells(1, 1 + (Application.Match("CPK", Headerarray, 0) - 1) * waferNum)
                                If Not IsEmpty(objPara.mLow) And Not IsEmpty(objPara.mHigh) Then
                                    If Not IsEmpty(objPara.mLow) And Not IsEmpty(objPara.mHigh) Then
                                        .Formula = "=MIN(($F" & CStr(nowRow) & "-$" & N2L(iCol + i + 1) & CStr(nowRow) & ")/($" & N2L(iCol + i + 1 + 2 * waferNum) & CStr(nowRow) & "*3),($" & N2L(iCol + i + 1) & CStr(nowRow) & "-$D" & CStr(nowRow) & ")/($" & N2L(iCol + i + 1 + 2 * waferNum) & CStr(nowRow) & "*3))"
                                    ElseIf Not IsEmpty(objPara.mLow) Then
                                        .Formula = "=($" & N2L(iCol + i + 1) & CStr(nowRow) & "-$D" & CStr(nowRow) & ")/($" & N2L(iCol + i + 1 + 2 * waferNum) & CStr(nowRow) & "*3)"
                                    Else
                                        .Formula = "=($F" & CStr(nowRow) & "-$" & N2L(iCol + i + 1) & CStr(nowRow) & ")/($" & N2L(iCol + i + 1 + 2 * waferNum) & CStr(nowRow) & "*3)"
                                    End If
                                    .NumberFormatLocal = "0.00"
                                    .FormatConditions.Delete
                                    Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="1")
                                    nowCondition.Interior.ColorIndex = 3
                                    Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=mCPK)
                                    nowCondition.Interior.ColorIndex = 27
                                End If
                                'CPK Summary
                                If paraIndex = itemRange.Rows.Count Then
                                    tmpStr = N2L(iCol + i + 1 + (Application.Match("CPK", Headerarray, 0) - 1) * waferNum) & CStr(3) & ":" & N2L(iCol + i + 1 + (Application.Match("CPK", Headerarray, 0) - 1) * waferNum) & CStr(2 + paraIndex)
                                    .Cells(2, 1).Formula = "=COUNTIF(" & tmpStr & ","">=1"") & ""/"" & COUNT(" & tmpStr & ")" & " & "", "" & " & "COUNTIF(" & tmpStr & ","">=" & mCPK & """) & ""/"" & COUNT(" & tmpStr & ")"
                                    .Cells(2, 1).HorizontalAlignment = xlRight
                                End If
                            End With
                            'Ca
                            With .Cells(1, 1 + (Application.Match("Ca", Headerarray, 0) - 1) * waferNum)
                                If Not IsEmpty(objPara.mLow) And Not IsEmpty(objPara.mHigh) And Not IsEmpty(objPara.mTarget) Then
                                    .Formula = "=ABS(" & "($" & N2L(iCol + i + 1) & CStr(nowRow) & "-$E" & CStr(nowRow) & ")*2/(" & "$F" & CStr(nowRow) & "-" & "$D" & CStr(nowRow) & "))"
                                    .NumberFormatLocal = "0.00"
                                    .FormatConditions.Delete
                                    Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="1")
                                    nowCondition.Interior.ColorIndex = 27
                                End If
                                'Ca Summary
                                If paraIndex = itemRange.Rows.Count Then
                                    .Cells(2, 1).Formula = "=COUNTIF(" & N2L(iCol + i + 1 + (Application.Match("Ca", Headerarray, 0) - 1) * waferNum) & CStr(3) & ":" & N2L(iCol + i + 1 + (Application.Match("Ca", Headerarray, 0) - 1) * waferNum) & CStr(2 + paraIndex) & ",""<=1"") & ""/"" & COUNT(" & N2L(iCol + i + 1 + (Application.Match("Ca", Headerarray, 0) - 1) * waferNum) & CStr(3) & ":" & N2L(iCol + i + 1 + (Application.Match("Ca", Headerarray, 0) - 1) * waferNum) & CStr(2 + paraIndex) & ")"
                                    .Cells(2, 1).HorizontalAlignment = xlRight
                                End If
                            End With
                            'Cp
                            With .Cells(1, 1 + (Application.Match("Cp", Headerarray, 0) - 1) * waferNum)
                                If Not IsEmpty(objPara.mLow) And Not IsEmpty(objPara.mHigh) Then
                                    .Formula = "=" & "(" & "$F" & CStr(nowRow) & "-" & "$D" & CStr(nowRow) & ")" & "/" & "($" & N2L(iCol + i + 1 + 2 * waferNum) & CStr(nowRow) & "*6)"
                                    .NumberFormatLocal = "0.00"
                                    .FormatConditions.Delete
                                    Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="1.1")
                                    nowCondition.Interior.ColorIndex = 27
                                End If
                                'Cp summary
                                If paraIndex = itemRange.Rows.Count Then
                                    .Cells(2, 1).Formula = "=COUNTIF(" & N2L(iCol + i + 1 + (Application.Match("Cp", Headerarray, 0) - 1) * waferNum) & CStr(3) & ":" & N2L(iCol + i + 1 + (Application.Match("Cp", Headerarray, 0) - 1) * waferNum) & CStr(2 + paraIndex) & ",""<=1"") & ""/"" & COUNT(" & N2L(iCol + i + 1 + (Application.Match("Cp", Headerarray, 0) - 1) * waferNum) & CStr(3) & ":" & N2L(iCol + i + 1 + (Application.Match("Cp", Headerarray, 0) - 1) * waferNum) & CStr(2 + paraIndex) & ")"
                                    .Cells(2, 1).HorizontalAlignment = xlRight
                                End If
                            End With
                        End If
                    End If
                End With
            End If
        Next i
    Next paraIndex
    
    'Summary Diff
    Call Diff_Summary(mSheet, waferList, Headerarray)
   
    'Set Header Split Line
    For i = 0 To hcount
        Set nowRange = nowSheet.Columns(iCol + i * waferNum)
        With nowRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
    Next i

    'Set ReportStyle
    Call RangeStyle(nowSheet, itemRange)
    Call MergeDevice(nowSheet.Name)
    
    'Set Number Format
    Call SummaryFormatByUnit(mSheet & "_Summary", hcount)
    Call HighLightSBG(mSheet & "_Summary")
    nowSheet.Activate
    nowSheet.Cells.Select
    Selection.Font.Size = 10
    Selection.Font.Name = "Arial"
    ActiveWindow.Zoom = 75
    Selection.Columns.AutoFit
    If IsKey(exStr, "COMP") Then nowSheet.Columns(7).EntireColumn.Hidden = True
    Selection.Rows.AutoFit
    nowSheet.Range("A3").Select
    ActiveWindow.FreezePanes = True
   
    'CPK, Ca, Cp Chart
    If mCPK <> "" Then
        Dim xChart As Long, yChart As Long
        xChart = nowSheet.UsedRange.Range("A:G").Columns.width + 5
        yChart = nowSheet.UsedRange.Rows.Height + 50
        nowCol = 3
        'CPK
        iRow = nowSheet.UsedRange.Rows.Count + 1
        nowRow = iRow
        With nowSheet.Range(N2L(nowCol) & CStr(nowRow))
            .Cells(1, 1) = "Cpk"
            .Cells(2, 1) = "Wafer"
            .Cells(2, 2) = "Cpk>=" & mCPK
            .Cells(2, 3) = "1<=Cpk<" & mCPK
            .Cells(2, 4) = "0<=Cpk<1"
            .Cells(2, 5) = "Cpk<0"
        End With
        nowRow = nowSheet.UsedRange.Rows.Count + 1
        With nowSheet.Range(N2L(nowCol) & CStr(nowRow))
            For i = 0 To UBound(waferList)
                tmpStr = N2L(iCol + i + 1 + (Application.Match("CPK", Headerarray, 0) - 1) * waferNum)
                .Range("A" & CStr(1 + i)).Value = "#" & waferList(i)
                .Range("B" & CStr(1 + i)).Formula = "=COUNTIF(" & tmpStr & CStr(3) & ":" & tmpStr & CStr(2 + paraIndex) & ","">=" & mCPK & """)"
                .Range("C" & CStr(1 + i)).Formula = "=COUNTIF(" & tmpStr & CStr(3) & ":" & tmpStr & CStr(2 + paraIndex) & ",""<" & mCPK & """)" & "-" & "COUNTIF(" & tmpStr & CStr(3) & ":" & tmpStr & CStr(2 + paraIndex) & ",""<1"")"
                .Range("D" & CStr(1 + i)).Formula = "=COUNTIF(" & tmpStr & CStr(3) & ":" & tmpStr & CStr(2 + paraIndex) & ",""<1"")" & "-" & "COUNTIF(" & tmpStr & CStr(3) & ":" & tmpStr & CStr(2 + paraIndex) & ",""<0"")"
                .Range("E" & CStr(1 + i)).Formula = "=COUNTIF(" & tmpStr & CStr(3) & ":" & tmpStr & CStr(2 + paraIndex) & ",""<0"")"
            Next i
        End With
        nowRow = nowSheet.UsedRange.Rows.Count
        Set nowRange = nowSheet.Range(N2L(nowCol) & CStr(nowRow)).CurrentRegion
        Call CPKPlot(nowSheet, nowRange, xChart, yChart)
        'Ca
        iRow = nowSheet.UsedRange.Rows.Count + 2
        nowRow = iRow
        With nowSheet.Range(N2L(nowCol) & CStr(nowRow))
            .Cells(1, 1) = "Ca"
            .Cells(2, 1) = "Wafer"
            .Cells(2, 2) = "Ca<=1"
            .Cells(2, 3) = "Ca>1"
        End With
        nowRow = nowSheet.UsedRange.Rows.Count + 1
        With nowSheet.Range(N2L(nowCol) & CStr(nowRow))
            For i = 0 To UBound(waferList)
                tmpStr = N2L(iCol + i + 1 + (Application.Match("Ca", Headerarray, 0) - 1) * waferNum)
                .Range("A" & CStr(1 + i)).Value = "#" & waferList(i)
                .Range("B" & CStr(1 + i)).Formula = "=COUNTIF(" & tmpStr & CStr(3) & ":" & tmpStr & CStr(2 + paraIndex) & ",""<=1"")"
                .Range("C" & CStr(1 + i)).Formula = "=COUNTIF(" & tmpStr & CStr(3) & ":" & tmpStr & CStr(2 + paraIndex) & ","">1"")"
            Next i
        End With
        nowRow = nowSheet.UsedRange.Rows.Count
        Set nowRange = nowSheet.Range(N2L(nowCol) & CStr(nowRow)).CurrentRegion
        Call CPKPlot(nowSheet, nowRange, xChart, yChart + 305)
        'Cp
        iRow = nowSheet.UsedRange.Rows.Count + 2
        nowRow = iRow
        With nowSheet.Range(N2L(nowCol) & CStr(nowRow))
            .Cells(1, 1) = "Cp"
            .Cells(2, 1) = "Wafer"
            .Cells(2, 2) = "Cp>=1.1"
            .Cells(2, 3) = "Cp<1.1"
        End With
        nowRow = nowSheet.UsedRange.Rows.Count + 1
        With nowSheet.Range(N2L(nowCol) & CStr(nowRow))
            For i = 0 To UBound(waferList)
                tmpStr = N2L(iCol + i + 1 + (Application.Match("Cp", Headerarray, 0) - 1) * waferNum)
                .Range("A" & CStr(1 + i)).Value = "#" & waferList(i)
                .Range("B" & CStr(1 + i)).Formula = "=COUNTIF(" & tmpStr & CStr(3) & ":" & tmpStr & CStr(2 + paraIndex) & ","">=1.1"")"
                .Range("C" & CStr(1 + i)).Formula = "=COUNTIF(" & tmpStr & CStr(3) & ":" & tmpStr & CStr(2 + paraIndex) & ",""<1.1"")"
            Next i
        End With
        nowRow = nowSheet.UsedRange.Rows.Count
        Set nowRange = nowSheet.Range(N2L(nowCol) & CStr(nowRow)).CurrentRegion
        Call CPKPlot(nowSheet, nowRange, xChart, yChart + 305 + 305)
    End If
   
End Function

Public Function Diff_Summary(mSheet As String, waferList() As String, Headerarray)
    
    Dim nowRow As Long, nowCol As Long, n As Long
    Dim nowSheet As Worksheet
    Dim waferNum As Integer
    Dim hcount As Integer
    Dim np As Integer
    
    Set nowSheet = Worksheets(mSheet & "_Summary")
    nowSheet.Activate
    waferNum = UBound(waferList) + 1
    hcount = UBound(Headerarray) + 1
    
    If nowSheet.Cells(2, 7) = "Shift to BSL" Then np = 1
    
    For nowRow = 1 To nowSheet.UsedRange.Rows.Count
        With nowSheet.Range("G1:" & N2L(6 + 2 * waferNum) & CStr(nowSheet.UsedRange.Rows.Count))
            If Left(nowSheet.Cells(nowRow, 2), 6) = "Diff.%" Then
                For nowCol = 1 To waferNum
                    If np = 1 Then .Cells(nowRow, nowCol).Formula = "=" & N2L(waferNum + 6 + nowCol) & CStr(nowRow) & "-" & N2L(waferNum + 6 + 1) & CStr(nowRow)
                    .Cells(nowRow, waferNum * np + nowCol).Formula = "=IF($E" & CStr(nowRow - 1) & "="""",""""," & N2L(6 + waferNum * np + nowCol) & CStr(nowRow - 1) & "/$E" & CStr(nowRow - 1) & "-1" & ")"
                    .Cells(nowRow, waferNum * np + nowCol).NumberFormatLocal = "0.00%"
                    .Cells(nowRow, waferNum * (np + 1) + nowCol).Formula = "=IF($E" & CStr(nowRow - 1) & "="""",""""," & N2L(6 + waferNum * np + nowCol + waferNum) & CStr(nowRow - 1) & "/$E" & CStr(nowRow - 1) & "-1" & ")"
                    .Cells(nowRow, waferNum * (np + 1) + nowCol).NumberFormatLocal = "0.00%"
                Next nowCol
            ElseIf Left(nowSheet.Cells(nowRow, 2), 5) = "Diff." Then
                Dim ratio
                ratio = 1
                If InStr(nowSheet.Cells(nowRow, 2), "*") Then
                    ratio = getCOL(nowSheet.Cells(nowRow, 2), "*", 2)
                    If IsNumeric(ratio) Then ratio = Val(ratio) Else ratio = 1
                    nowSheet.Cells(nowRow, 2) = getCOL(nowSheet.Cells(nowRow, 2), "*", 1)
                End If
                For nowCol = 1 To waferNum
                    If np = 1 Then .Cells(nowRow, nowCol).Formula = "=" & N2L(waferNum + 6 + nowCol) & CStr(nowRow) & "-" & N2L(waferNum + 6 + 1) & CStr(nowRow)
                    .Cells(nowRow, waferNum * np + nowCol).Formula = "=IF($E" & CStr(nowRow - 1) & "="""","""",(" & N2L(6 + waferNum * np + nowCol) & CStr(nowRow - 1) & "-$E" & CStr(nowRow - 1) & ")*" & ratio & ")"
                    .Cells(nowRow, waferNum * np + nowCol).NumberFormatLocal = "0.00"
                    .Cells(nowRow, waferNum * (np + 1) + nowCol).Formula = "=IF($E" & CStr(nowRow - 1) & "="""","""",(" & N2L(6 + waferNum * np + nowCol + waferNum) & CStr(nowRow - 1) & "-$E" & CStr(nowRow - 1) & ")*" & ratio & ")"
                    .Cells(nowRow, waferNum * (np + 1) + nowCol).NumberFormatLocal = "0.00"
                Next nowCol
            ElseIf Left(nowSheet.Cells(nowRow, 2), 5) = "Times" Then
                For nowCol = 1 To waferNum
                    If np = 1 Then .Cells(nowRow, nowCol).Formula = "=" & N2L(waferNum + 6 + nowCol) & CStr(nowRow) & "-" & N2L(waferNum + 6 + 1) & CStr(nowRow)
                    .Cells(nowRow, waferNum * np + nowCol).Formula = "=IF($E" & CStr(nowRow - 1) & "="""",""""," & N2L(waferNum * np + 6 + nowCol) & CStr(nowRow - 1) & "/$E" & CStr(nowRow - 1) & ")"
                    .Cells(nowRow, waferNum * np + nowCol).NumberFormatLocal = "0.000""x"""
                    .Cells(nowRow, waferNum * (np + 1) + nowCol).Formula = "=IF($E" & CStr(nowRow - 1) & "="""",""""," & N2L(waferNum * np + 6 + nowCol + waferNum) & CStr(nowRow - 1) & "/$E" & CStr(nowRow - 1) & ")"
                    .Cells(nowRow, waferNum * (np + 1) + nowCol).NumberFormatLocal = "0.000""x"""
                Next nowCol
            End If
        End With
    Next nowRow
    Set nowSheet = Nothing
    
End Function

Public Function CPKPlot(nowSheet As Worksheet, chartRange As Range, xChart As Long, yChart As Long)
    Dim mTitle As String
    Dim xLabel As String
    Dim yLabel As String
    Dim strRange As Range
    Dim nowChart As Chart
   
    mTitle = chartRange.Range("A1").Value
    xLabel = chartRange.Range("A2").Value
    yLabel = "Yield(%)"
    Set strRange = chartRange.Range("A2:" & N2L(chartRange.Columns.Count) & CStr(chartRange.Rows.Count))
    Set nowChart = nowSheet.ChartObjects.Add(xChart, yChart, 500, 300).Chart
    With nowChart
        .chartType = xlColumnStacked100
        .SetSourceData Source:=strRange, PlotBy:=xlColumns
        nowChart.HasTitle = True
        nowChart.ChartTitle.Text = mTitle
        nowChart.HasLegend = True
        nowChart.Legend.Position = xlLegendPositionTop
    End With
End Function

Public Function MergeDevice(mSheet As String)
   Dim nowSheet As Worksheet
   Const deviceColumn As Integer = 1
   Dim i As Long
   Dim sRow As Long
   Dim eRow As Long
   Dim oldDevice As String
   Dim YNAlert As Boolean
   
    Application.DisplayAlerts = False
    
    Set nowSheet = Worksheets(mSheet)
    sRow = 1
    For i = 2 To nowSheet.UsedRange.Rows.Count + 1
        If nowSheet.Cells(i, 1) <> oldDevice Or nowSheet.Cells(i, 1) = "" Then
            eRow = i - 1
            If eRow > sRow Then
                nowSheet.Range(Cells(sRow, 1), Cells(eRow, 1)).Merge
                nowSheet.Range(Cells(sRow, 1), Cells(eRow, 1)).VerticalAlignment = xlCenter
            End If
            oldDevice = nowSheet.Cells(i, 1)
            sRow = i
        End If
    Next i
    Set nowSheet = Nothing
    
    Application.DisplayAlerts = True

End Function

Public Function GetSiteList(mSheet As String, ByRef mSite() As String)
   Dim i As Long, j As Long
   Dim siteNum As Integer
   Dim tmpStr As String
   Dim tmpX As String, tmpY As String
   Dim aWafer() As String
   
   siteNum = getSiteNum(dSheet)
   Call GetWaferList(dSheet, aWafer)
   
   ReDim mSite(siteNum - 1)
   For i = 1 To siteNum
      tmpStr = getValueByPara(aWafer(0), "Parameter", i)
      tmpStr = getCOL(getCOL(tmpStr, "(", 2), ")", 1)
      tmpX = CInt(getCOL(tmpStr, ",", 1))
      tmpY = CInt(getCOL(tmpStr, ",", 2))
      mSite(i - 1) = CStr(i) & "(" & tmpX & "," & tmpY & ")"
   Next i
   
   
End Function

Public Function GetWaferArray(mSheet As String, ByRef mWaferlist() As String)
   Dim i As Long
   ReDim mWaferlist(1, 0)
   For i = 1 To GetRowRange(mSheet)
      If Trim(Worksheets(mSheet).Cells(i, 1)) = "TYPE_VECTOR" And mWaferlist(0, 0) <> "" Then Exit For
      If Worksheets(mSheet).Cells(i, 3) = "Unit" Then
         If mWaferlist(0, 0) <> "" Then ReDim Preserve mWaferlist(1, UBound(mWaferlist, 2) + 1)
         mWaferlist(0, UBound(mWaferlist, 2)) = getCOL(getCOL(Worksheets(mSheet).Cells(i, 4), "-", 1), "<", 2)
      End If
   Next i
End Function

Public Function GetRowRange(mSheet As String)
   Dim i As Long
   Dim nowSheet As Worksheet
   
   If Not IsExistSheet(mSheet) Then Exit Function
   Set nowSheet = Worksheets(mSheet)
   
   For i = nowSheet.UsedRange.Rows.Count To 1 Step -1
      If Application.WorksheetFunction.countA(nowSheet.Rows(i)) <> 0 Then
         GetRowRange = i
         Exit For
      End If
   Next i
End Function

Function ynRange(sheetName As String, RangeName As String)
   Dim mRange As Range
   
   On Error GoTo NoRange
   Set mRange = Worksheets(sheetName).Range(RangeName)
   ynRange = True
   Exit Function
NoRange:
   ynRange = False
End Function

Public Function DioPlotAllChart()
   Dim i As Long, j As Long
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim scatterChart As Chart
   Dim boxtrendChart As Chart
   
    Call DelSheet("All_Chart")
    For i = 1 To Worksheets.Count
        If Left(UCase(Sheets(i).Name), 7) = "SCATTER" Then
            If Worksheets(i).ChartObjects.Count >= 1 Then
                For j = Worksheets(i).ChartObjects.Count To 1 Step -1
                    Worksheets(i).ChartObjects(j).Delete
                Next j
            End If
            Call PlotUniversalChart(Sheets(i).Name)
        ElseIf Left(UCase(Sheets(i).Name), 8) = "BOXTREND" Then
            If Worksheets(i).ChartObjects.Count >= 1 Then
                For j = Worksheets(i).ChartObjects.Count To 1 Step -1
                    Worksheets(i).ChartObjects(j).Delete
                Next j
            End If
            Call prepareBoxTrendData(Sheets(i).Name)
            Call PlotBoxTrendChart(Sheets(i).Name)
        ElseIf Left(UCase(Sheets(i).Name), 10) = "CUMULATIVE" Then
            If Worksheets(i).ChartObjects.Count >= 1 Then
                For j = Worksheets(i).ChartObjects.Count To 1 Step -1
                    Worksheets(i).ChartObjects(j).Delete
                Next j
            End If
            Call PlotCumulativeChart(Sheets(i).Name)
        End If
        Call adjustChartObject(Sheets(i))
    Next i
End Function
Public Function getReceiver()
   Dim nowSheet As Worksheet
   Dim i As Long, j As Integer
   Dim TOlist As String
   Dim CClist As String
   Dim BCClist As String
   Dim tmpStr As String
   
   If Not IsExistSheet("RECEIVER") Then
      Err.Number = "1001"
      Err.Description = "No Worksheet[Receiver]"
      Exit Function
   End If
   Set nowSheet = Worksheets("RECEIVER")
   For i = 2 To nowSheet.Rows.Count
      If Trim(nowSheet.Cells(i, 1)) <> "" And Trim(nowSheet.Cells(i, 2)) <> "" Then
         tmpStr = Trim(nowSheet.Cells(i, 1))
         If Right(tmpStr, 1) = "," Then tmpStr = Left(tmpStr, Len(tmpStr) - 1)   '去除最後面的逗點
         For j = 1 To 5
            tmpStr = Replace(tmpStr, "  ", " ") '去除連續的空白
            tmpStr = Replace(tmpStr, ", ", ",") '去除逗點後的空白
            tmpStr = Replace(tmpStr, " ,", ",") '去除逗點前的空白
         Next j
         tmpStr = Replace(tmpStr, " ", "_") '空白=>底線
         '新舊部門處理
         tmpStr = UCase(tmpStr)
         tmpStr = Replace(tmpStr, "ATD_LOGIC-45NM_1", "ATD_Logic-Logic_1")
         tmpStr = Replace(tmpStr, "ATD_LOGIC-45NM_2", "ATD_Logic-Logic_2")
         
         tmpStr = Replace(tmpStr, ",", "@umc.com ") '逗點=>@umc.com空白
         If Not IsInvalidMail(tmpStr) Then
            Select Case UCase(Trim(nowSheet.Cells(i, 2)))
               Case "TO":  TOlist = TOlist & " " & tmpStr & "@umc.com"
               Case "CC":  CClist = CClist & " " & tmpStr & "@umc.com"
               Case "BCC": BCClist = BCClist & " " & tmpStr & "@umc.com"
            End Select
         End If
      End If
   Next i
   TOlist = Trim(TOlist)
   CClist = Trim(CClist)
   BCClist = Trim(BCClist)
   getReceiver = TOlist & "," & CClist & "," & BCClist
End Function

Public Sub Hide_Config()
   Dim i As Long
   
   For i = 1 To Worksheets.Count - 1
      If Worksheets(i).Name = "Data" Then Exit For
      Worksheets(i).Visible = 0
   Next i
End Sub

Public Sub Show_Config()
   Dim i As Long
   
   For i = 1 To Worksheets.Count - 1
      If Worksheets(i).Name = "Data" Then Exit For
      Worksheets(i).Visible = 1
   Next i
End Sub


Sub SetArrayRange()
   Dim nowSheet As Worksheet
   Dim i As Long, j As Long
   Dim nowRange As Range
   Dim nowWafer As String
   Dim temp
   Dim typeName As String
   Dim nowPara As String
   Dim sRow As Long, eRow As Long
   Dim sCol As Long, eCol As Long
   
   'Remove old Names
   '-----------------
   For i = ActiveWorkbook.Names.Count To 1 Step -1
      If InStr(ActiveWorkbook.Names(i).Name, "array") > 0 Then _
         ActiveWorkbook.Names(i).Delete
   Next i
   'ReDim WaferArray(0)
   
   typeName = "wafer_"
   Set nowSheet = Worksheets("Data")
   For i = 1 To nowSheet.UsedRange.Rows.Count
      If Trim(nowSheet.Cells(i, 1)) = "TYPE_SCALAR" Then typeName = "wafer_"
      If Trim(nowSheet.Cells(i, 1)) = "TYPE_VECTOR" Then typeName = "array_"
      If typeName = "array_" Then
         'wafer header
         If InStr(nowSheet.Cells(i, 1), "No") > 0 Then _
            nowWafer = getCOL(getCOL(nowSheet.Cells(i, 4), "<", 2), "-", 1)
         'parameter
         If IsNumeric(nowSheet.Cells(i, 1)) And nowSheet.Cells(i, 1) <> "" Then
            nowPara = Trim(nowSheet.Cells(i, 2))
            nowPara = Replace(nowPara, "+", "_")
            nowPara = Replace(nowPara, "-", "_")
            nowPara = Replace(nowPara, "*", "_")
            nowPara = Replace(nowPara, "/", "_")
            i = i + 1
            sRow = i
            j = 4
            sCol = j
            While nowSheet.Cells(i, j) <> ""
               j = j + 1
            Wend
            eCol = j - 1
            While nowSheet.Cells(i, 2) <> ""
               i = i + 1
            Wend
            eRow = i - 1
            Set nowRange = nowSheet.Range(N2L(sCol) & CStr(sRow) & ":" & N2L(eCol) & CStr(eRow))
            'Debug.Print sRow, eRow, N2L(sCol), N2L(eCol)
            'Debug.Print nowRange.Address
            nowSheet.Names.Add typeName & nowPara & "_" & nowWafer, nowRange
         End If
      End If
   Next i
End Sub

Public Sub GenArray()
   Dim nowSheet As Worksheet
   Dim i  As Long, j As Long, iWafer As Integer
   Dim chartNo As Integer
   Dim objChart As chartInfo
   Dim nowArray As String
   Dim siteStr As String
   Dim xRange As Range, yRange As Range
   Dim siteNum As Integer
   Dim waferList() As String
   Dim ynError As Boolean
   Dim siteArray As Variant, iSite As Variant
   Dim nowCol As Long, nowRow As Long
   Dim xPara As String
   Dim yPara As String
   Dim objParaX As specInfo
   Dim objParaY As specInfo
   Dim tempX As String, tempY As String
   Dim iPara As Integer
   Dim n As Long
   
   ' get wafer list and siteNum
   '---------------------------
   Call GetWaferList(dSheet, waferList)
   siteNum = getSiteNum(dSheet)
   
   Call SetArrayRange
   If Not IsExistSheet("array") Then MsgBox "worksheet array is not exist.": Exit Sub
   Set nowSheet = Worksheets("array")
   
   'Generate ScatterA chart header
   '------------------------------
   For i = 3 To nowSheet.UsedRange.Columns.Count Step 2
      If nowSheet.Cells(1, i) <> "" Then
         chartNo = chartNo + 1
         AddSheet ("ScatterA" & CStr(chartNo))
         nowSheet.Range(N2L(i) & ":" & N2L(i + 1)).Copy Worksheets("ScatterA" & CStr(chartNo)).Range("A1")
      End If
   Next i
   
   For i = 1 To chartNo
      Set nowSheet = Worksheets("ScatterA" & CStr(i))
      'Debug.Print nowSheet.Range("A1").CurrentRegion.Address
      objChart = getChartInfo(nowSheet.Range("A1").CurrentRegion)
      nowCol = 3
      For iPara = 1 To objChart.YParameter.Count
          xPara = objChart.XParameter(iPara)
          yPara = objChart.YParameter(iPara)
          objParaX = getSPECInfo(xPara)
          objParaY = getSPECInfo(yPara)
          siteStr = UCase(Trim(objChart.SplitBy))   'Site list => siteArray
          If siteStr = "ALL" Then
             ReDim siteArray(siteNum - 1)
             For j = 1 To siteNum
                siteArray(j - 1) = CStr(j)
             Next j
          Else
             siteArray = Split(siteStr, ",")
          End If
          For iWafer = 0 To UBound(waferList)
             nowRow = 2
             On Error Resume Next
             Set xRange = Worksheets("Data").Range("array_" & xPara & "_" & waferList(iWafer))
             Set yRange = Worksheets("Data").Range("array_" & yPara & "_" & waferList(iWafer))
             ynError = Err.Number <> 0
             On Error GoTo 0
             If Not ynError Then
                For Each iSite In siteArray
                   If InStr(UCase(objChart.ChartExpression), "ALL") > 0 Or InStr(UCase(objChart.ChartExpression), "RAWDATA") > 0 Then
                        For j = 1 To xRange.Rows.Count   'Array data
                           nowSheet.Cells(nowRow + j, nowCol) = xRange(j, CInt(iSite)) * objParaX.mFAC
                           nowSheet.Cells(nowRow + j, nowCol + 1) = yRange(j, CInt(iSite)) * objParaY.mFAC
                           'With Filter
                           If UCase(objChart.GDataFilter) = "YES" Then
                              'X Para
                              If (Not IsEmpty(objParaX.mHigh) And nowSheet.Cells(nowRow + j, nowCol) > objParaX.mHigh) Or _
                                 (Not IsEmpty(objParaX.mLow) And nowSheet.Cells(nowRow + j, nowCol) < Val(objParaX.mLow)) Then
                                 nowSheet.Cells(nowRow + j, nowCol) = ""
                              End If
                              'Y Para
                              If (Not IsEmpty(objParaY.mHigh) And nowSheet.Cells(nowRow + j, nowCol + 1) > objParaY.mHigh) Or _
                                 (Not IsEmpty(objParaY.mLow) And nowSheet.Cells(nowRow + j, nowCol + 1) < Val(objParaY.mLow)) Then
                                 nowSheet.Cells(nowRow + j, nowCol + 1) = ""
                              End If
                           End If
                        Next j
                        nowSheet.Cells(1, nowCol) = yPara & "@" & iSite & "#" & waferList(iWafer)
                        nowCol = nowCol + 2
                    End If
                    If InStr(UCase(objChart.ChartExpression), "ALL") > 0 Or InStr(UCase(objChart.ChartExpression), "AVERAGE") > 0 Then
                        nowSheet.Cells(nowRow + j, nowCol) = WorksheetFunction.Average(xRange.Columns(CInt(iSite)))
                        nowSheet.Cells(nowRow + j, nowCol + 1) = WorksheetFunction.Average(yRange.Columns(CInt(iSite)))
                        nowSheet.Cells(1, nowCol) = "Avg of " & yPara & "@" & iSite & "#" & waferList(iWafer)
                        nowCol = nowCol + 2
                    End If
                    If InStr(UCase(objChart.ChartExpression), "ALL") > 0 Or InStr(UCase(objChart.ChartExpression), "MEDIAN") > 0 Then
                        nowSheet.Cells(nowRow + j, nowCol) = WorksheetFunction.Median(xRange.Columns(CInt(iSite)))
                        nowSheet.Cells(nowRow + j, nowCol + 1) = WorksheetFunction.Median(yRange.Columns(CInt(iSite)))
                        nowSheet.Cells(1, nowCol) = "Med of " & yPara & "@" & iSite & "#" & waferList(iWafer)
                        nowCol = nowCol + 2
                    End If
                Next iSite
    '               nowCol = nowCol + 2
    '               For j = 2 To objChart.XParameter.Count 'Non-Array Data
    '                  tempX = objChart.XParameter(j)
    '                  tempY = objChart.YParameter(j)
    '                  nowSheet.Cells(nowRow + 1, nowCol) = getValueByPara(WaferList(iWafer), tempX, CInt(iSite))
    '                  nowSheet.Cells(nowRow + 1, nowCol + 1) = getValueByPara(WaferList(iWafer), tempY, CInt(iSite))
    '                  nowSheet.Cells(1, nowCol) = tempX & "@" & iSite & "#" & WaferList(iWafer)
    '                  nowCol = nowCol + 2
    '               Next j
               'nowCol = nowCol + 2
                
                'Target parameter
                '---------------------
                If objChart.vTargetXValueStr <> "" And objChart.vTargetYValueStr <> "" Then
                    tempX = objChart.vTargetXValueStr  '.XParameter(j)
                    tempY = objChart.vTargetYValueStr   '.YParameter(j)
                    n = 0
                    For Each iSite In siteArray
                      nowSheet.Cells(nowRow + 1 + n, nowCol) = getValueByPara(waferList(iWafer), tempX, CInt(iSite))
                      nowSheet.Cells(nowRow + 1 + n, nowCol + 1) = getValueByPara(waferList(iWafer), tempY, CInt(iSite))
                      nowSheet.Cells(1, nowCol) = objChart.vTargetNameStr 'tempX & "@" & iSite & "#" & WaferList(iWafer)
                      n = n + 1
                    Next iSite
                    nowCol = nowCol + 2
                End If
             End If
          Next iWafer
        Next iPara
   Next i
   
'               'With Filter
'            If UCase(vChartInfo.GDataFilter) = "YES" Then
'               If (Not IsEmpty(vSPEC.mHigh) And nowSheet.Cells(nowRow, nowCol) > vSPEC.mHigh) Or _
'                  (Not IsEmpty(vSPEC.mLow) And nowSheet.Cells(nowRow, nowCol) < Val(vSPEC.mLow)) Then
'                  nowSheet.Cells(nowRow, nowCol) = ""
'               End If
'            End If
            
   Call DioPlotAllChart
   'Call FitChart
   Call new_FitChart
End Sub
