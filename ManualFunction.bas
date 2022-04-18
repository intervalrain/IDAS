
' This is a macro collection from multiple authors
' which are demo, specific purposed, or customized functions to meet some engineers' request from other departments.


Dim ynSeries As Boolean

Type SAInfo
   Parameter As String
   MainType As String
   SubType As String
   width As Single
   Height As Single
   SAMIN As Single
   SAREF As Single
   SAMINvalue As Single
   SAREFvalue As Single
End Type
Option Explicit

Public Sub Model_AllMacro()
   Model_ChangeSeriesName
   Model_ChangeYAxisFormat
   Model_ModifyBoxTrend
End Sub

Public Function Manual_Vincent_WID()
   Dim nowSheet As Worksheet
   Dim rawSheet As Worksheet
   Dim widSheet As Worksheet
   Dim tmp As String
   Dim iRow As Long
   Dim i As Long, j As Long
   Dim waferNum As Integer
   Dim siteNum As Integer
   Dim waferList() As String
   Dim iWafer As Integer
   Dim rawRange As Range
   Dim medianA() As Single
   Dim sigmaA() As Single
   Dim widA() As Single
   Dim wiwA() As Single
   Dim n As Integer
   Dim maxWafer As Integer
   
   Set nowSheet = ActiveSheet
   If InStr(nowSheet.Name, "_Summary") <= 0 Then Exit Function
   tmp = getCOL(nowSheet.Name, "_Summary", 1)
   Set widSheet = AddSheet(tmp & "_WID")
   
   Call GetWaferList("Data", waferList)
   waferNum = UBound(waferList) + 1
   siteNum = getSiteNum("Data")
   ReDim medianA(siteNum)
   ReDim sigmaA(siteNum)
   ReDim widA(1 To siteNum)
   'Debug.Print "Num:", WaferNum, WaferList(0), SiteNum
   
   maxWafer = 245 \ siteNum
   iRow = 2
   For i = 3 To nowSheet.UsedRange.Rows.Count
      j = i + nowSheet.Cells(i, 1).MergeArea.Rows.Count - 1
      widSheet.Cells(iRow, 1) = nowSheet.Cells(i, 1)
      widSheet.Range(widSheet.Cells(iRow, 1), widSheet.Cells(iRow + 4, 1)).Merge
      widSheet.Range(widSheet.Cells(iRow, 1), widSheet.Cells(iRow + 4, 1)).VerticalAlignment = xlCenter
      widSheet.Cells(iRow, 2) = "WIW WID(Median)"
      widSheet.Cells(iRow + 1, 2) = "WIW WID(sigma)"
      widSheet.Cells(iRow + 2, 2) = "WIW WID(U%)"
      widSheet.Cells(iRow + 3, 2) = "WIW (U%)"
      widSheet.Cells(iRow + 4, 2) = "WID (U%)"
      For iWafer = 1 To waferNum
         If iWafer <= maxWafer Then
            Set rawSheet = Worksheets(tmp & "_Raw")
         Else
            Set rawSheet = Worksheets(tmp & "_Raw_" & CStr((iWafer - 1) \ maxWafer))
         End If
         
         Set rawRange = rawSheet.Range(N2L(7 + ((iWafer - 1) Mod maxWafer) * siteNum) & CStr(i) & ":" & N2L(7 + ((iWafer - 1) Mod maxWafer + 1) * siteNum - 1) & CStr(j))
         'Debug.Print rawRange.Address
         For n = 1 To siteNum
            medianA(n) = WorksheetFunction.Median(rawRange.Columns(n))
            sigmaA(n) = WorksheetFunction.StDev(rawRange.Columns(n))
            If medianA(n) = 0 Then medianA(n) = medianA(n) + 1E-23
            widA(n) = 3 * sigmaA(n) / medianA(n)
         Next n
         ReDim wiwA(1 To rawRange.Rows.Count)
         For n = 1 To rawRange.Rows.Count
            wiwA(n) = 3 * WorksheetFunction.StDev(rawRange.Rows(n)) / WorksheetFunction.Median(rawRange.Rows(n))
         Next n
         widSheet.Cells(1, 2 + iWafer) = "#" & waferList(iWafer - 1)
         widSheet.Cells(iRow, 2 + iWafer) = WorksheetFunction.Median(rawRange)
         widSheet.Cells(iRow + 1, 2 + iWafer) = WorksheetFunction.StDev(rawRange)
         widSheet.Cells(iRow + 2, 2 + iWafer).FormulaLocal = "=3*" & N2L(2 + iWafer) & CStr(iRow + 1) & "/" & N2L(2 + iWafer) & CStr(iRow)
         widSheet.Cells(iRow + 3, 2 + iWafer) = WorksheetFunction.Average(wiwA)
         widSheet.Cells(iRow + 4, 2 + iWafer) = WorksheetFunction.Average(widA)
         
         widSheet.Cells(iRow, 2 + iWafer).NumberFormat = "0.00"
         widSheet.Cells(iRow + 1, 2 + iWafer).NumberFormat = "0.00"
         widSheet.Cells(iRow + 2, 2 + iWafer).NumberFormat = "0.00%"
         widSheet.Cells(iRow + 3, 2 + iWafer).NumberFormat = "0.00%"
         widSheet.Cells(iRow + 4, 2 + iWafer).NumberFormat = "0.00%"
         
      Next iWafer
      i = j: iRow = iRow + 5
   Next i
   widSheet.Columns.AutoFit
   'nowSheet.Activate
End Function


Public Function Manual_Brief()
   Dim sheetName As String
   Dim nowSheet As Worksheet
   Dim fieldArray
   Dim i As Long, j As Long, n As Long
   Dim iSheet As Integer
   Dim tempA() As Long
   Dim exStr As String
   
   On Error Resume Next
   
   'sheetName = "Brief_Summary"
   fieldArray = Array("Median", "Average", "Sigma", "Yield")
   'If IsKey(exStr, "Max") Then fieldArray = Array("Median", "Average", "Sigma", "Yield", "Max", "Min", "Sigma%")
   
   '處理Diff ...
   '------------------------------------------------------------
   For iSheet = 1 To Worksheets.Count
      If Right(Worksheets(iSheet).Name, 8) = "_Summary" Then
         Set nowSheet = Worksheets(iSheet)
         '處理Diff and Diff%
         ReDim tempA(0)
         For i = 1 To nowSheet.UsedRange.Columns.Count
            If nowSheet.Cells(2, i) = fieldArray(0) Or nowSheet.Cells(2, i) = fieldArray(1) Then
               tempA(UBound(tempA)) = i
               ReDim Preserve tempA(UBound(tempA) + 1)
            End If
         Next i
         If UBound(tempA) > 0 Then ReDim Preserve tempA(UBound(tempA) - 1)
         For i = 1 To nowSheet.UsedRange.Rows.Count
            If Left(nowSheet.Cells(i, 2), 6) = "Diff.%" Then
               For j = 0 To UBound(tempA)
                  nowSheet.Range(Num2Letter(tempA(j)) & CStr(i)).Formula = "=IF($E" & CStr(i - 1) & "="""",""""," & Num2Letter(tempA(j)) & CStr(i - 1) & "/$E" & CStr(i - 1) & "-1" & ")"
                  nowSheet.Range(Num2Letter(tempA(j)) & CStr(i)).NumberFormat = "0.00%"
               Next j
            ElseIf Left(nowSheet.Cells(i, 2), 5) = "Diff." Then
               For j = 0 To UBound(tempA)
                  nowSheet.Range(Num2Letter(tempA(j)) & CStr(i)).Formula = "=IF($E" & CStr(i - 1) & "="""",""""," & Num2Letter(tempA(j)) & CStr(i - 1) & "-$E" & CStr(i - 1) & ")"
               Next j
            ElseIf Left(nowSheet.Cells(i, 2), 7) = "Time of" Then
               For j = 0 To UBound(tempA)
                  nowSheet.Range(Num2Letter(tempA(j)) & CStr(i)).Formula = "=IF($E" & CStr(i - 1) & "="""",""""," & Num2Letter(tempA(j)) & CStr(i - 1) & "/$E" & CStr(i - 1) & ")"
                  nowSheet.Range(Num2Letter(tempA(j)) & CStr(i)).NumberFormatLocal = "0.000""x"""
               Next j
            End If
         Next i
         Set nowSheet = Nothing
      End If
      If Right(Worksheets(iSheet).Name, 4) = "_Raw" Then
         Set nowSheet = Worksheets(iSheet)
         '處理Diff and Diff%
         For i = 1 To nowSheet.UsedRange.Columns.Count
            If nowSheet.Cells(2, i) = "1" Then
               n = i
               Exit For
            End If
         Next i

         For i = 1 To nowSheet.UsedRange.Rows.Count
            If Left(nowSheet.Cells(i, 2), 6) = "Diff.%" Then
               For j = n To nowSheet.UsedRange.Columns.Count
                  nowSheet.Range(Num2Letter(j) & CStr(i)).Formula = "=IF($E" & CStr(i - 1) & "="""",""""," & Num2Letter(j) & CStr(i - 1) & "/$E" & CStr(i - 1) & "-1" & ")"
                  nowSheet.Range(Num2Letter(j) & CStr(i)).NumberFormat = "0.00%"
               Next j
            ElseIf Left(nowSheet.Cells(i, 2), 5) = "Diff." Then
               For j = n To nowSheet.UsedRange.Columns.Count
                  nowSheet.Range(Num2Letter(j) & CStr(i)).Formula = "=IF($E" & CStr(i - 1) & "="""",""""," & Num2Letter(j) & CStr(i - 1) & "-$E" & CStr(i - 1) & ")"
               Next j
            ElseIf Left(nowSheet.Cells(i, 2), 4) = "Time" Then
               For j = n To nowSheet.UsedRange.Columns.Count
                  nowSheet.Range(Num2Letter(j) & CStr(i)).Formula = "=IF($E" & CStr(i - 1) & "="""",""""," & Num2Letter(j) & CStr(i - 1) & "/$E" & CStr(i - 1) & ")"
                  nowSheet.Range(Num2Letter(j) & CStr(i)).NumberFormatLocal = "0.000""x"""
               Next j
            End If
         Next i
         Set nowSheet = Nothing
      End If
      
   Next iSheet
   
   
   'Brief => 調整欄位順序
   '----------------------------------------------------------------------
   If IsExistSheet("SPEC_List") Then
      For iSheet = 1 To Worksheets("SPEC_List").UsedRange.Columns.Count
         If Trim(Worksheets("SPEC_List").Cells(1, iSheet)) = "" Then Exit For
         'If Trim(LCase(getCOL(Worksheets("SPEC_List").Cells(1, iSheet), ":", 2))) = "brief" Then
         If IsKey(getCOL(Worksheets("SPEC_List").Cells(1, iSheet), ":", 2), "brief") Then
            sheetName = getCOL(Worksheets("SPEC_List").Cells(1, iSheet), ":", 1) & "_Summary"
            If IsExistSheet(sheetName) Then
                fieldArray = Array("Median", "Average", "Sigma", "Yield")
                exStr = getCOL(Worksheets("SPEC_List").Cells(1, iSheet), ":", 2)
                If IsKey(exStr, "Max") Then fieldArray = Array("Median", "Average", "Sigma", "Yield", "Max", "Min", "Sigma%")
                If IsKey(exStr, "Diff") Then
                    ReDim Preserve fieldArray(UBound(fieldArray) + 1)
                    fieldArray(UBound(fieldArray)) = "Diff"
                    'fieldArray = Array("Median", "Average", "Sigma", "Yield", "Diff")
                End If
                'Else
                '  fieldArray = Array("Median", "Average", "Sigma", "Yield")
                'End If
               Set nowSheet = Worksheets(sheetName)
               '調換攔位順序
               For n = 0 To UBound(fieldArray)
                  ReDim tempA(0)
                  For i = 1 To nowSheet.UsedRange.Columns.Count
                     If nowSheet.Cells(2, i) = fieldArray(n) Then
                        tempA(UBound(tempA)) = i
                        ReDim Preserve tempA(UBound(tempA) + 1)
                     End If
                  Next i
                  If UBound(tempA) > 0 Then ReDim Preserve tempA(UBound(tempA) - 1)
                  nowSheet.Columns(tempA(0)).Borders(xlEdgeRight).LineStyle = xlNone
                  For i = 1 To UBound(tempA)
                     If tempA(i) > tempA(0) + i Then
                        nowSheet.Columns(tempA(i)).Cut
                        nowSheet.Columns(tempA(0) + i).Insert Shift:=xlToRight
                     End If
                     nowSheet.Columns(tempA(0) + i).Borders(xlEdgeRight).LineStyle = xlNone
                  Next i
                  With nowSheet.Columns(tempA(0) + i - 1).Borders(xlEdgeRight)
                     .LineStyle = xlContinuous
                     .Weight = xlMedium
                     .ColorIndex = xlAutomatic
                  End With
               Next n
               nowSheet.UsedRange.FormatConditions.Delete
               '加格式化條件
               For j = 6 To nowSheet.Columns.Count
                  If nowSheet.Cells(2, j) = "Median" Then
                    For i = 3 To nowSheet.Rows.Count
                        If nowSheet.Cells(i, 4) <> "" Then nowSheet.Cells(i, j).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & "D" & "" & CStr(3)).Font.ColorIndex = 4
                        If nowSheet.Cells(i, 6) <> "" Then nowSheet.Cells(i, j).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & "F" & "" & CStr(3)).Font.ColorIndex = 3
                    Next i
                  End If
               Next j
               Set nowSheet = Nothing
            End If
         End If
      Next iSheet
   End If
   
   
   
End Function

Sub Model_ChangeSeriesName()
   Dim i As Integer, j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim nowSeries As Series
   Dim xLabel As String
   Dim CateName As String
   Dim Unit As String
   
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "SCATTER") Then
         Set nowSheet = Worksheets(i)
         For j = 1 To nowSheet.ChartObjects.Count
            Set nowChart = nowSheet.ChartObjects(j).Chart
            If Not nowChart.HasAxis(xlCategory) Then Exit For
            xLabel = nowChart.Axes(xlCategory).AxisTitle.Text
            CateName = UCase(getCOL(xLabel, "(", 1))
            Unit = getCOL(getCOL(xLabel, "(", 2), ")", 1)
            If CateName <> "SA" And CateName <> "W" And CateName <> "L" Then Exit For
            If Unit = "" Then Exit For
            For m = 1 To nowChart.SeriesCollection.Count
               Set nowSeries = nowChart.SeriesCollection(m)
               If Application.WorksheetFunction.Max(nowSeries.XValues) = Application.WorksheetFunction.Min(nowSeries.XValues) Then
                  nowSeries.Name = CateName & "=" & CStr(nowSeries.XValues(1)) & "" & Unit
               End If
               'Debug.Print nowSeries.Name
            Next m
         Next j
      End If
   Next i
End Sub

Sub Model_ChangeYAxisFormat()
   Dim i As Integer, j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim nowAxis As Axis
   Dim yLabel As String
   Dim CateName As String
   Dim Unit As String
   
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "SCATTER") Then
         Set nowSheet = Worksheets(i)
         For j = 1 To nowSheet.ChartObjects.Count
            Set nowChart = nowSheet.ChartObjects(j).Chart
            If Not nowChart.HasAxis(xlValue) Then Exit For
            yLabel = nowChart.Axes(xlValue).AxisTitle.Text
            CateName = UCase(Left(yLabel, 3))
            Select Case CateName
               Case "IDS", "IDL", "IOF"
                  Set nowAxis = nowChart.Axes(xlValue)
                  nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
            End Select
         Next j
      End If
   Next i
End Sub

Sub Model_AddMedianTable()
   Dim i As Integer, j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim TableRange As Range
   Dim MedianRange As Range
   Dim xLabel As String
   Dim yLabel As String
   Dim nowSA As SAInfo
   
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "SCATTER") Then
         Set nowSheet = Worksheets(i)
         
         For j = 1 To nowSheet.UsedRange.Columns.Count
            If UCase(nowSheet.Cells(1, j)) = "MEDIAN" Then Exit For
         Next j
         If j < nowSheet.UsedRange.Columns.Count Then
            For m = 3 To nowSheet.UsedRange.Rows.Count
               If nowSheet.Cells(m, j) = "" Then Exit For
            Next m
            Set MedianRange = nowSheet.Range(Num2Letter(j) & CStr(1) & ":" & Num2Letter(j + 1) & CStr(m - 1))
            Set TableRange = nowSheet.Range(Num2Letter(nowSheet.UsedRange.Columns.Count + 1) & "1")
            Call getSAInfo(getFirstParameter(nowSheet.Name), MedianRange, nowSA)
            TableRange.Cells(1, 1) = "Med"
            TableRange.Range("A1").HorizontalAlignment = xlRight
            TableRange.Cells(2, 1) = "SA(um)"
            With TableRange.Range("A1:A2")
               nowSheet.Shapes.AddLine(.Left, .Top, .Left + .width, .Top + .Height).Select
               .Borders.LineStyle = xlContinuous
               .Borders(xlInsideHorizontal).LineStyle = xlNone
               .Interior.ColorIndex = 2
            End With
            With TableRange.Range("B1:C2")
               .Borders.LineStyle = xlContinuous
               .Borders(xlInsideHorizontal).LineStyle = xlNone
               .Interior.ColorIndex = 2
               .Cells(2, 1) = nowSA.MainType
               .Cells(2, 2) = nowSA.SubType
               .Columns.AutoFit
            End With

            For m = 3 To MedianRange.Rows.Count
               If MedianRange.Cells(m, 1) = "" Then Exit For
               With TableRange.Range("A" & CStr(m) & ":C" & CStr(m))
                  .Cells(1, 1) = MedianRange.Cells(m, 1)
                  .Cells(1, 2) = MedianRange.Cells(m, 2)
                  Select Case nowSA.MainType
                     Case "VTSN", "VTSP":  .Cells(1, 3) = (.Cells(1, 2) - Val(nowSA.SAREFvalue)) * 1000
                     Case "IDSN", "IDSP":  .Cells(1, 3) = (.Cells(1, 2) - Val(nowSA.SAREFvalue)) / Val(nowSA.SAREFvalue) * 100
                  End Select
                  .Borders.LineStyle = xlContinuous
                  .HorizontalAlignment = xlHAlignCenter
               End With
            Next m
         End If
      End If
   Next i
End Sub

Public Function getFirstParameter(sheetName As String)
   Dim nowSheet As Worksheet
   Dim i As Long
   
   Set nowSheet = Worksheets(sheetName)
   For i = 1 To nowSheet.UsedRange.Rows.Count
      If UCase(nowSheet.Cells(i, 1)) = "Y" Then getFirstParameter = nowSheet.Cells(i + 1, 1)
   Next i
End Function

Public Function getSAInfo(nowParameter As String, ByRef MedianRange As Range, ByRef nowSA As SAInfo)
   Dim i As Long
   
   nowSA.Parameter = nowParameter
   nowSA.MainType = UCase(getCOL(nowSA.Parameter, "_", 1))
   Select Case nowSA.MainType
      Case "VTSN":   nowSA.SubType = "VTSN_D(mV)"
      Case "VTSP":   nowSA.SubType = "VTSP_D(mV)"
      Case "IDSN":   nowSA.SubType = "IDSN(%)"
      Case "IDSP":   nowSA.SubType = "IDSP(%)"
   End Select
   nowSA.width = CSng(Replace(LCase(getCOL(nowSA.Parameter, "_", 2)), "p", "."))
   nowSA.Height = CSng(Replace(LCase(getCOL(nowSA.Parameter, "_", 3)), "p", "."))
   If nowSA.width <= 0.2 Then
      nowSA.SAMIN = 0.36
      nowSA.SAREF = 1.89
   Else
      nowSA.SAMIN = 0.32
      nowSA.SAREF = 1.76
   End If
   For i = 1 To MedianRange.Rows.Count
      If MedianRange.Cells(i, 1) = nowSA.SAMIN Then nowSA.SAMINvalue = MedianRange.Cells(i, 2)
      If MedianRange.Cells(i, 1) = nowSA.SAREF Then nowSA.SAREFvalue = MedianRange.Cells(i, 2)
   Next i
   
   'Debug.Print nowSA.Parameter, nowSA.Width, nowSA.Height, nowSA.SAMIN, nowSA.SAREF, nowSA.SAMINvalue, nowSA.SAREFvalue
   'Stop
End Function

Sub Model_ModifyBoxTrend()
   Dim i As Integer, j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim nowAxis As Axis
   Dim nowSeries As Series
   Dim nowPoint As Point
   Dim yLabel As String
   Dim CateName As String
   Dim Unit As String
   Dim tempA
   Dim waferNum As Integer
   Dim tmpStr As String
   Dim ArrayStr As String
   
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "BOXTREND") Then
         Set nowSheet = Worksheets(i)
         For j = 1 To nowSheet.ChartObjects.Count
            Set nowChart = nowSheet.ChartObjects(j).Chart
            Set nowAxis = nowChart.Axes(xlValue)
            nowAxis.MajorGridlines.Border.ColorIndex = 2
            Set nowAxis = nowChart.Axes(xlCategory)
            nowAxis.MajorGridlines.Border.ColorIndex = 2
            nowAxis.TickLabels.Font.Size = 12
            If InStr(nowAxis.AxisTitle.Text, "SA") <= 0 Then Exit For
            tempA = nowChart.SeriesCollection(1).XValues
            For m = 1 To UBound(tempA)
               If IsEmpty(tempA(m)) Then Exit For
            Next m
            waferNum = m - 1
            nowAxis.TickLabelSpacing = waferNum + 1
            nowAxis.TickMarkSpacing = waferNum + 1
            For m = 1 To UBound(tempA)
               If Not IsEmpty(tempA(m)) Then
                  If nowChart.SeriesCollection(5).Points(m).HasDataLabel Then
                     tmpStr = nowChart.SeriesCollection(5).Points(m).DataLabel.Text
                     tmpStr = getCOL(tmpStr, "_", 4)
                     tmpStr = Replace(tmpStr, "p", ".")
                     If Left(tmpStr, 1) = "." Then tmpStr = "0" & tmpStr
                  End If
               End If
            Next m
            Set nowSeries = nowChart.SeriesCollection(4)
            ArrayStr = ""
            For m = 1 To UBound(tempA)
               If Not IsEmpty(tempA(m)) Then
                  Set nowPoint = nowSeries.Points(m)
                  nowPoint.HasDataLabel = True
                  nowPoint.DataLabel.Text = CStr(tempA(m))
                  nowPoint.DataLabel.Font.Size = 8
                  nowPoint.DataLabel.Position = xlLabelPositionAbove
                  If nowChart.SeriesCollection(5).Points(m).HasDataLabel Then
                     tmpStr = nowChart.SeriesCollection(5).Points(m).DataLabel.Text
                     tmpStr = getCOL(tmpStr, "_", 4)
                     tmpStr = Replace(tmpStr, "p", ".")
                     If Left(tmpStr, 1) = "." Then tmpStr = "0" & tmpStr
                     ArrayStr = ArrayStr & "," & tmpStr
                  Else
                     ArrayStr = ArrayStr & "," & "0"
                  End If
                  nowChart.SeriesCollection(5).Points(m).HasDataLabel = False
               Else
                  ArrayStr = ArrayStr & "," & "0"
               End If
            Next m
            If Len(ArrayStr) > 1 Then ArrayStr = Mid(ArrayStr, 2)
            nowChart.SeriesCollection(1).XValues = Split(ArrayStr, ",")
            nowChart.SeriesCollection(5).Border.LineStyle = xlLineStyleNone
         Next j
      End If
   Next i
End Sub

Sub Manual_FilterBySSFF()
   Dim i As Integer, j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim ssRange As Range
   Dim ffRange As Range
   Dim ssBig As Boolean
   
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "SCATTER") Then
         Set nowSheet = Worksheets(i)
         Set ssRange = getRangeBySeriesName(nowSheet, "SS")
         If ssRange.Cells.Count > 1 Then
            Set ffRange = getRangeBySeriesName(nowSheet, "FF")
            If getValueByIndex(ssRange, ssRange.Cells(1, 1), 1, 2) > getValueByIndex(ffRange, ssRange.Cells(1, 1), 1, 2) Then
               ssBig = True
            Else
               ssBig = False
            End If
            For j = 3 To nowSheet.UsedRange.Columns.Count Step 2
               If nowSheet.Cells(2, j) = "" Then Exit For
               For m = 3 To nowSheet.UsedRange.Rows.Count
                  If getValueByIndex(ssRange, nowSheet.Cells(m, j), 1, 2) Then
                     If ssBig Then
                        If nowSheet.Cells(m, j + 1) > getValueByIndex(ssRange, nowSheet.Cells(m, j), 1, 2) Then
                           nowSheet.Cells(m, j + 1) = ""
                        End If
                        If nowSheet.Cells(m, j + 1) < getValueByIndex(ffRange, nowSheet.Cells(m, j), 1, 2) Then
                           nowSheet.Cells(m, j + 1) = ""
                        End If
                     Else
                        If nowSheet.Cells(m, j + 1) < getValueByIndex(ssRange, nowSheet.Cells(m, j), 1, 2) Then
                           nowSheet.Cells(m, j + 1) = ""
                        End If
                        If nowSheet.Cells(m, j + 1) > getValueByIndex(ffRange, nowSheet.Cells(m, j), 1, 2) Then
                           nowSheet.Cells(m, j + 1) = ""
                        End If
                     End If
                  End If
               Next m
            Next j
            'Debug.Print getValueByIndex(ssRange, 10, 1, 2)
            Call FitAxisScale(nowSheet)
         End If
      End If
   Next i
End Sub


Function getRangeBySeriesName(nowSheet As Worksheet, SeriesName As String) As Range
   Dim i As Long, j As Long
   
   Set getRangeBySeriesName = nowSheet.Range("A1")
   For i = 1 To nowSheet.UsedRange.Columns.Count
      If nowSheet.Cells(1, i) = SeriesName Then Exit For
   Next i
   If i > nowSheet.UsedRange.Columns.Count Then Exit Function
   For j = 3 To nowSheet.UsedRange.Rows.Count
      If nowSheet.Cells(j, i) = "" Then Exit For
   Next j
   If j = 3 Then Exit Function
   Set getRangeBySeriesName = nowSheet.Range(Num2Letter(i) & "3" & ":" & Num2Letter(i + 1) & CStr(j))
End Function

Function getValueByIndex(nowRange As Range, nowIndex, indexCol As Integer, valueCol As Integer)
   Dim i As Long
   
   getValueByIndex = False
   For i = 1 To nowRange.Rows.Count
      If nowRange.Cells(i, indexCol) = nowIndex Then
         getValueByIndex = nowRange.Cells(i, valueCol)
         Exit For
      End If
   Next i
End Function

Private Sub FitAxisScale(ByRef nowSheet As Worksheet)
   Dim yMin, yMax, xMin, xMax
   Dim vChartInfo As chartInfo
   Dim nowChart As Chart
   Dim nowAxis As Axis
   Dim ynLog As Boolean

   vChartInfo = getChartInfo(nowSheet.Range("A1").CurrentRegion)
   Set nowChart = nowSheet.ChartObjects(1).Chart
   Call getScaterMaxMin(nowChart, xMax, xMin, yMax, yMin)
   Set nowAxis = nowChart.Axes(xlValue)
   If IsKey(vChartInfo.YScale, "Log") Then
      ynLog = True
   Else
      ynLog = False
   End If
   Call AxisScaleFit(nowAxis, yMax, yMin, "", "", ynLog)
   'Public Sub AxisScaleFit(ByRef rAxis As Axis, ByVal vvarMax, ByVal vvarMin, ByVal cMax, ByVal cMin, varLog As Boolean)
End Sub

Private Sub Manual_FitLegend()
   Dim i As Integer, j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim nowLegend As Legend
   Dim nowSeries As Series
   Dim nowShape As Shape
   Dim nowAxis As Axis
  
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "SCATTER") Or InStr(UCase(Worksheets(i).Name), "BOXTREND") Then
         Set nowSheet = Worksheets(i)
         'nowSheet.Activate
         'nowSheet.Range("A1").Select
         For j = 1 To nowSheet.ChartObjects.Count
            'nowSheet.ChartObjects(j).Activate
            Set nowChart = nowSheet.ChartObjects(j).Chart
            'nowChart.ChartArea.Select
            '-----------------
            'Fit Chart Attribs
            '-----------------
            Set nowShape = nowSheet.Shapes(1)
            'Debug.Print nowShape.Name
            With nowShape
               .Top = 30
               .Left = 30
               .width = 450
               .Height = 300 + (j - 1) * 400
            End With
            DoEvents
            'nowChart.ChartArea.Width = 500
            '----------
            'Fit Legend
            '----------
            If nowChart.HasLegend Then
               Set nowLegend = nowChart.Legend
               'nowLegend.Select
               nowLegend.AutoScaleFont = False
               nowLegend.Font.Size = 9
               'nowLegend.Height = nowChart.ChartArea.Height - 40
               'Selection.Width = 50
               'nowLegend.Width = 500
               For m = 1 To nowChart.SeriesCollection.Count
                  Set nowSeries = nowChart.SeriesCollection(m)
                  'Debug.Print nowSeries.Name
                  On Error Resume Next
                  Select Case Len(nowSeries.Name)
                     Case Is > 25
                        nowLegend.LegendEntries(m).Font.Size = 8
                     Case Else
                        nowLegend.LegendEntries(m).Font.Size = 9
                  End Select
                  On Error GoTo 0
               Next m
               'nowLegend.Width = 150
               
               'nowLegend.Left = nowChart.ChartArea.Width - nowLegend.Width - 6 - 3
               'nowLegend.Top = (nowChart.ChartArea.Height - nowLegend.Height) / 2 + 10
               'Debug.Print "W:", nowShape.Width, nowLegend.Width
               'Debug.Print "H:", nowShape.Height, nowLegend.Height
               DoEvents
               nowLegend.Left = nowShape.width - nowLegend.width - 6 - 10
               nowLegend.Top = (nowShape.Height - nowLegend.Height) / 2 '+ 10
               'DoEvents
               'nowLegend.Width = nowShape.Width
               'DoEvents
               'Debug.Print "Legend:", nowLegend.Left, nowLegend.Top
               'Debug.Print nowChart.ChartArea.Width & ":" & nowLegend.Left & ":" & nowLegend.Width
            End If
            '----------
            'Fit Axis
            '----------
            If True Then 'nowChart.HasAxis(lCategory, xlPrimary) Then
               Set nowAxis = nowChart.Axes(xlCategory, xlPrimary)
               nowAxis.TickLabels.AutoScaleFont = True
               nowAxis.TickLabels.Font.Name = "Arial"
               nowAxis.TickLabels.Font.Size = 10
               If nowAxis.HasTitle Then nowAxis.AxisTitle.Font.Size = 12
               If InStr(UCase(nowSheet.Name), "BOXTREND") Then
                  nowAxis.TickLabels.Font.Size = 6
                  nowAxis.TickLabels.Orientation = xlTickLabelOrientationUpward
               End If
            End If
            If True Then 'nowChart.HasAxis(xlValue, xlPrimary) Then
               Set nowAxis = nowChart.Axes(xlValue, xlPrimary)
               nowAxis.TickLabels.AutoScaleFont = True
               nowAxis.TickLabels.Font.Name = "Arial"
               nowAxis.TickLabels.Font.Size = 10
               If nowAxis.HasTitle Then nowAxis.AxisTitle.Font.Size = 12
            End If
         Next j
      End If
   Next i
End Sub

Public Function new_FitChart()
    Dim i As Integer, j As Integer, n As Integer
    Dim nowChart As Chart
    Dim nowSheet As Worksheet
    Dim nowLegend As Legend
    Dim nowShape As Shape
    Dim vChartInfo As chartInfo
    Dim tmpStr As String
    Dim nowAxis As Axis

    For i = 1 To Worksheets.Count
        If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Or InStr(UCase(Worksheets(i).Name), "BOXTREND") > 0 Then
            Set nowSheet = Worksheets(i)
            For j = 1 To nowSheet.ChartObjects.Count
                Set nowChart = nowSheet.ChartObjects(j).Chart
                'nowChart.ClearToMatchStyle
                'nowChart.ChartStyle = 343
                '-----------------
                'Fit Chart Attribs
                '-----------------
                Set nowShape = nowSheet.Shapes(j)
                DoEvents
                With nowShape
                   .Top = 30 + (j - 1) * 400
                   .Left = 30
                   .width = 450
                   .Height = 300
                End With
                DoEvents
                '----------
                'Fit Legend
                '----------
                On Error Resume Next
                If nowChart.HasLegend Then
                   Set nowLegend = nowChart.Legend
                   nowLegend.Border.Color = 0
                   nowLegend.Border.ColorIndex = 1
                   nowLegend.Border.LineStyle = 1
                   nowLegend.Border.Weight = 2
                   
                   nowLegend.Fill.BackColor.SchemeColor = 2
                   nowLegend.Fill.ForeColor.SchemeColor = 19
                   nowLegend.Fill.Visible = msoTrue
                   
                   nowLegend.AutoScaleFont = False
                   nowLegend.Font.Size = 10
                   DoEvents
                   
                   nowLegend.Left = 60
                   nowLegend.Top = 30
                   nowLegend.Height = 30
                   nowLegend.width = 360
                   nowLegend.Interior.ColorIndex = 19
                End If
                On Error Resume Next
                vChartInfo = getChartInfo(nowSheet.Range("A1").CurrentRegion)
                If Trim(UCase(vChartInfo.XParameter(1))) = "SITE" Then
                   nowChart.Axes(xlCategory).TickLabels.Font.Size = 10
                   nowChart.Axes(xlCategory).MajorUnit = 1
                End If
                '----------
                'Add grid line
                If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Then
                    nowChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
                    nowChart.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
                    nowChart.SetElement (msoElementPrimaryCategoryGridLinesMinorMajor)
                End If
                
                
                '----------
                'Plot Area
                '----------
                nowChart.PlotArea.Left = 31.658
                nowChart.PlotArea.Top = 69
                nowChart.PlotArea.Format.Fill.BackColor.SchemeColor = 9
                
                nowChart.PlotArea.Border.ColorIndex = 1
                nowChart.PlotArea.LineStyle = 1
                nowChart.PlotArea.Weight = 2
                
                nowChart.PlotArea.Height = 203
                If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Then
                    nowChart.PlotArea.InsideHeight = 180
                    nowChart.PlotArea.InsideTop = 72
                ElseIf InStr(UCase(Worksheets(i).Name), "BOXTREND") > 0 Then
                    nowChart.PlotArea.InsideHeight = 222
                    nowChart.PlotArea.InsideTop = 30
                End If
                
                nowChart.PlotArea.InsideLeft = 60
                nowChart.PlotArea.InsideWidth = 360
                nowChart.PlotArea.Border.Color = 0
                nowChart.PlotArea.Border.ColorIndex = 1
                nowChart.PlotArea.Border.LineStyle = 1
                nowChart.PlotArea.Border.Weight = 2
                
                nowChart.PlotArea.Format.Line.Visible = msoTrue
                nowChart.PlotArea.Format.Line.Weight = 1.5

                '----------
                'Chart Title
                '----------
                nowChart.ChartTitle.Font.ColorIndex = 56
                nowChart.ChartTitle.Font.Size = 16
                nowChart.ChartTitle.Left = nowChart.PlotArea.InsideLeft + 180 - nowChart.ChartTitle.width / 2
                
                
                '----------
                'Axis Title
                '----------
                nowChart.Axes(xlValue).AxisTitle.Font.Size = 14
                nowChart.Axes(xlValue).AxisTitle.Font.ColorIndex = 56
                nowChart.Axes(xlValue).AxisTitle.Font.Bold = False
                nowChart.Axes(xlValue).AxisTitle.Top = 114
                nowChart.Axes(xlValue).AxisTitle.Left = 12
                nowChart.Axes(xlCategory).AxisTitle.Font.Size = 14
                nowChart.Axes(xlCategory).AxisTitle.Font.ColorIndex = 56
                nowChart.Axes(xlCategory).AxisTitle.Font.Bold = False
                nowChart.Axes(xlCategory).AxisTitle.Top = 272
                nowChart.Axes(xlCategory).AxisTitle.Left = 216
                
                '----------
                'Axis
                '----------
                nowChart.Axes(xlValue).Left = 32
                nowChart.Axes(xlValue).Top = 77
                nowChart.Axes(xlValue).TickLabels.Font.Size = 10
                nowChart.Axes(xlCategory).Left = 63
                nowChart.Axes(xlCategory).Top = 249
                nowChart.Axes(xlCategory).TickLabels.Font.Size = 10
            
                '----------
                'TrendLines
                '----------
                Dim nowLabel As DataLabel
                For n = 1 To nowChart.FullSeriesCollection.Count
                    If nowChart.FullSeriesCollection(n).Trendlines.Count > 0 Then
                        Set nowLabel = nowChart.FullSeriesCollection(n).Trendlines(1).DataLabel
                        nowLabel.Left = nowChart.PlotArea.InsideLeft + nowChart.PlotArea.InsideWidth - nowLabel.width
                        nowLabel.Top = nowChart.PlotArea.InsideTop + (n - 1) * nowLabel.Height
                    End If
                Next n

                
            
            If Trim(UCase(vChartInfo.XParameter(1))) = "PARA" Then
                nowChart.chartType = xlLineMarkers
                tmpStr = """" & vChartInfo.YParameter(1) & """"
                For n = 1 To vChartInfo.YParameter.Count
                    tmpStr = tmpStr & "," & """" & vChartInfo.YParameter(n) & """"
                Next n
                tmpStr = Replace(tmpStr, "=", "")
                'nowChart.SeriesCollection(1).XValues = "={""A"",""B""}"
                nowChart.SeriesCollection(1).XValues = "={" & tmpStr & "}"
                'nowChart.SeriesCollection(2).XValues = "={" & tmpStr & "}"
                nowChart.Axes(xlCategory).TickLabels.Font.Size = 8
                'nowChart.Axes(xlCategory).MajorUnit = 1
            End If
            
            
            On Error GoTo 0
         Next j
      End If
   Next i
End Function
Public Function FitSingleChart(nowSheet As Worksheet)
    
    Dim i As Integer, j As Integer, n As Integer
    Dim nowChart As Chart
    Dim nowShape As Shape
    Dim vChartInfo As chartInfo
    Dim tmpStr As String
    Dim nowLegend As Legend
    Dim nowAxis As Axis
    
    Set nowChart = nowSheet.ChartObjects(1).Chart
    Set nowShape = nowSheet.Shapes(1)
    
    '-----------------
    'Fit Chart Attribs
    '-----------------
    DoEvents
    With nowShape
       .Top = 30 + (j - 1) * 400
       .Left = 30
       .width = 450
       .Height = 300
    End With
    DoEvents
    '----------
    'Fit Legend
    '----------
    On Error Resume Next
    If nowChart.HasLegend Then
        Set nowLegend = nowChart.Legend
        nowLegend.Border.Color = 0
        nowLegend.Border.ColorIndex = 1
        nowLegend.Border.LineStyle = 1
        nowLegend.Border.Weight = 2
       
        nowLegend.Fill.BackColor.SchemeColor = 2
        nowLegend.Fill.ForeColor.SchemeColor = 19
        nowLegend.Fill.Visible = msoTrue
       
        nowLegend.AutoScaleFont = False
        nowLegend.Font.Size = 10
        DoEvents
       
        nowLegend.Left = 60
        nowLegend.Top = 30
        nowLegend.Height = 30
        nowLegend.width = 360
        nowLegend.Interior.ColorIndex = 19
    End If
    On Error Resume Next
    vChartInfo = getChartInfo(nowSheet.Range("A1").CurrentRegion)
    If Trim(UCase(vChartInfo.XParameter(1))) = "SITE" Then
       nowChart.Axes(xlCategory).TickLabels.Font.Size = 10
       nowChart.Axes(xlCategory).MajorUnit = 1
    End If
    '----------
    'Add grid line
    '----------
    If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Then
        nowChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
        nowChart.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
        nowChart.SetElement (msoElementPrimaryCategoryGridLinesMinorMajor)
    End If
    '----------
    'Plot Area
    '----------
    nowChart.PlotArea.Left = 31.658
    nowChart.PlotArea.Top = 69
    nowChart.PlotArea.Format.Fill.BackColor.SchemeColor = 9
    
    nowChart.PlotArea.Border.ColorIndex = 1
    nowChart.PlotArea.LineStyle = 1
    nowChart.PlotArea.Weight = 2
    
    nowChart.PlotArea.Height = 203
    If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Then
        nowChart.PlotArea.InsideHeight = 180
        nowChart.PlotArea.InsideTop = 72
    ElseIf InStr(UCase(Worksheets(i).Name), "BOXTREND") > 0 Then
        nowChart.PlotArea.InsideHeight = 222
        nowChart.PlotArea.InsideTop = 30
    End If
    
    nowChart.PlotArea.InsideLeft = 60
    nowChart.PlotArea.InsideWidth = 360
    nowChart.PlotArea.Border.Color = 0
    nowChart.PlotArea.Border.ColorIndex = 1
    nowChart.PlotArea.Border.LineStyle = 1
    nowChart.PlotArea.Border.Weight = 2
    
    nowChart.PlotArea.Format.Line.Visible = msoTrue
    nowChart.PlotArea.Format.Line.Weight = 1.5

    '----------
    'Chart Title
    '----------
    nowChart.ChartTitle.Font.ColorIndex = 56
    nowChart.ChartTitle.Font.Size = 16
    nowChart.ChartTitle.Left = nowChart.PlotArea.InsideLeft + 180 - nowChart.ChartTitle.width / 2
    
    '----------
    'Axis Title
    '----------
    nowChart.Axes(xlValue).AxisTitle.Font.Size = 14
    nowChart.Axes(xlValue).AxisTitle.Font.ColorIndex = 56
    nowChart.Axes(xlValue).AxisTitle.Font.Bold = False
    nowChart.Axes(xlValue).AxisTitle.Top = 114
    nowChart.Axes(xlValue).AxisTitle.Left = 12
    nowChart.Axes(xlCategory).AxisTitle.Font.Size = 14
    nowChart.Axes(xlCategory).AxisTitle.Font.ColorIndex = 56
    nowChart.Axes(xlCategory).AxisTitle.Font.Bold = False
    nowChart.Axes(xlCategory).AxisTitle.Top = 272
    nowChart.Axes(xlCategory).AxisTitle.Left = 216
    
    '----------
    'Axis
    '----------
    nowChart.Axes(xlValue).Left = 32
    nowChart.Axes(xlValue).Top = 77
    nowChart.Axes(xlValue).TickLabels.Font.Size = 10
    nowChart.Axes(xlCategory).Left = 63
    nowChart.Axes(xlCategory).Top = 249
    nowChart.Axes(xlCategory).TickLabels.Font.Size = 10

    '----------
    'TrendLines
    '----------
    Dim nowLabel As DataLabel
    For n = 1 To nowChart.FullSeriesCollection.Count
        If nowChart.FullSeriesCollection(n).Trendlines.Count > 0 Then
            Set nowLabel = nowChart.FullSeriesCollection(n).Trendlines(1).DataLabel
            nowLabel.Left = nowChart.PlotArea.InsideLeft + nowChart.PlotArea.InsideWidth - nowLabel.width
            nowLabel.Top = nowChart.PlotArea.InsideTop + (n - 1) * nowLabel.Height
        End If
    Next n

    If Trim(UCase(vChartInfo.XParameter(1))) = "PARA" Then
        nowChart.chartType = xlLineMarkers
        tmpStr = """" & vChartInfo.YParameter(1) & """"
        For n = 1 To vChartInfo.YParameter.Count
            tmpStr = tmpStr & "," & """" & vChartInfo.YParameter(n) & """"
        Next n
        tmpStr = Replace(tmpStr, "=", "")
        nowChart.SeriesCollection(1).XValues = "={" & tmpStr & "}"
        nowChart.Axes(xlCategory).TickLabels.Font.Size = 8
    End If
    
    On Error GoTo 0
   
End Function

Public Function old_FitChart()
   Dim i As Integer, j As Integer, n As Integer
   Dim nowChart As Chart
   Dim nowSheet As Worksheet
   Dim nowLegend As Legend
   Dim nowShape As Shape
   Dim vChartInfo As chartInfo
   Dim tmpStr As String
   Dim nowAxis As Axis
  
  
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Or InStr(UCase(Worksheets(i).Name), "BOXTREND") > 0 Then
         Set nowSheet = Worksheets(i)
         For j = 1 To nowSheet.ChartObjects.Count
            Set nowChart = nowSheet.ChartObjects(j).Chart
            '-----------------
            'Fit Chart Attribs
            '-----------------
            Set nowShape = nowSheet.Shapes(j)
            'Debug.Print nowShape.Name
            DoEvents
            With nowShape
               .Top = 30 + (j - 1) * 400
               .Left = 30
               .width = 450
               .Height = 300
            End With
            DoEvents
            '----------
            'Fit Legend
            '----------
            On Error Resume Next
            If nowChart.HasLegend Then
               Set nowLegend = nowChart.Legend
               nowLegend.AutoScaleFont = False
               nowLegend.Font.Size = 10
               DoEvents
               nowLegend.Left = nowShape.width - nowLegend.width - 6 - 10
               nowLegend.Top = (nowShape.Height - nowLegend.Height) / 2 '+ 10
               nowLegend.Interior.ColorIndex = 2
            End If
            On Error Resume Next
            vChartInfo = getChartInfo(nowSheet.Range("A1").CurrentRegion)
            If Trim(UCase(vChartInfo.XParameter(1))) = "SITE" Then
               nowChart.Axes(xlCategory).TickLabels.Font.Size = 10
               nowChart.Axes(xlCategory).MajorUnit = 1
            End If
            
'Avoid crash
'            If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Then
'                Set nowAxis = nowChart.Axes(xlCategory)
'                'If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
'                'Avoid crash
'                'nowAxis.TickLabels.AutoScaleFont = False
'                'nowAxis.TickLabels.Font.Size = 12
'                'Set nowAxis = Nothing
'                Set nowAxis = nowChart.Axes(xlValue)
'                'If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
'                'Avoid crash
'                'nowAxis.TickLabels.AutoScaleFont = False
'                'nowAxis.TickLabels.Font.Size = 12
'            Else
'                Set nowAxis = nowChart.Axes(xlCategory)
'                'If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
'                'Avoid crash
'                'nowAxis.TickLabels.AutoScaleFont = False
'                'nowAxis.TickLabels.Font.Size = 10
'                'Set nowAxis = Nothing
'                Set nowAxis = nowChart.Axes(xlValue)
'                'If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
'                'Avoid crash
'                'nowAxis.TickLabels.AutoScaleFont = False
'                'nowAxis.TickLabels.Font.Size = 10
'            End If
            
            
            
            If Trim(UCase(vChartInfo.XParameter(1))) = "PARA" Then
                nowChart.chartType = xlLineMarkers
                tmpStr = """" & vChartInfo.YParameter(1) & """"
                For n = 1 To vChartInfo.YParameter.Count
                    tmpStr = tmpStr & "," & """" & vChartInfo.YParameter(n) & """"
                Next n
                tmpStr = Replace(tmpStr, "=", "")
                'nowChart.SeriesCollection(1).XValues = "={""A"",""B""}"
                nowChart.SeriesCollection(1).XValues = "={" & tmpStr & "}"
                'nowChart.SeriesCollection(2).XValues = "={" & tmpStr & "}"
                nowChart.Axes(xlCategory).TickLabels.Font.Size = 8
                'nowChart.Axes(xlCategory).MajorUnit = 1
            End If
            
            
            On Error GoTo 0
         Next j
      End If
   Next i
End Function

Public Sub Manual_Device_FitChart() 
    Dim i As Long, j As Long
    Dim nowSheet As Worksheet
    Dim nowChart As Chart
    
    For i = 1 To Worksheets.Count
        If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Or InStr(UCase(Worksheets(i).Name), "BOXTREND") > 0 Then
            Set nowSheet = Worksheets(i)
            For j = 1 To nowSheet.ChartObjects.Count
                '------------------------------------------------------
                ' Device Chart Format
                '------------------------------------------------------
                Set nowChart = nowSheet.ChartObjects(j).Chart
                With nowChart
                    .HasTitle = False
                    If .chartType = xlXYScatter Then
                        With .Legend
                            .Top = 225
                            .Left = 183
                            .AutoScaleFont = False
                            .Font.Size = 12
                            .Font.FontStyle = "粗體"
                            .Font.Name = "Arial"
                            .Border.LineStyle = xlNone
                        End With
                    End If
                    '長和寬
                    .Parent.width = 391.2   '340
                    .Parent.Height = 292.8   '334
                    .PlotArea.Top = 0
                    .PlotArea.Left = 49
                    .PlotArea.width = 305   '280 - 21
                    .PlotArea.Height = 249  '293 - 30
                    '.ChartArea.Width = 384
                    '.ChartArea.Height = 284
                    
                    With .PlotArea.Border
                        .ColorIndex = 57
                        .Weight = xlMedium
                        .LineStyle = xlContinuous
                    End With
                    .ChartArea.Border.LineStyle = 0
                    
                    With .Axes(xlCategory)
                        .AxisTitle.AutoScaleFont = False
                        .AxisTitle.Font.FontStyle = "粗體"
                        .AxisTitle.Font.Size = 14
                        .AxisTitle.Font.Name = "Arial"
                        .AxisTitle.Top = 262
                        .TickLabels.AutoScaleFont = False
                        .TickLabels.Font.FontStyle = "粗體"
                        .TickLabels.Font.Size = 12
                        .TickLabels.Font.Name = "Arial"
                        .MajorTickMark = xlInside
                        .MinorTickMark = xlInside
                        .HasMajorGridlines = True
                        .MajorGridlines.Border.ColorIndex = 15
                    End With
                    With .Axes(xlValue)
                        .AxisTitle.AutoScaleFont = False
                        .AxisTitle.Font.FontStyle = "粗體"
                        .AxisTitle.Font.Size = 14
                        .AxisTitle.Font.Name = "Arial"
                        .AxisTitle.Left = 3
                        .TickLabels.AutoScaleFont = False
                        .TickLabels.Font.FontStyle = "粗體"
                        .TickLabels.Font.Size = 12
                        .TickLabels.Font.Name = "Arial"
                        .TickLabels.NumberFormatLocal = "0.E+00"
                        .HasMajorGridlines = True
                        .HasMinorGridlines = True
                        .MajorGridlines.Border.ColorIndex = 15
                        .MinorGridlines.Border.ColorIndex = 15
                        .MinorGridlines.Border.LineStyle = xlDot
                        .MajorTickMark = xlInside
                        .MinorTickMark = xlInside
                    End With
                
                End With
                '------------------------------------------------------
            Next j
        End If
    Next i

    Set nowChart = Nothing
    Set nowSheet = Nothing
    
    Call GenChartSummary
    
    
'        ActiveSheet.ChartObjects("圖表 1").Activate
'    ActiveChart.ChartArea.Select
'    ActiveSheet.Shapes("圖表 1").ScaleWidth 0.89, msoFalse, msoScaleFromTopLeft
'    ActiveSheet.Shapes("圖表 1").ScaleHeight 0.97, msoFalse, msoScaleFromTopLeft
'    ActiveChart.Axes(xlValue).AxisTitle.Select
'    ActiveChart.ChartArea.Select
'    Selection.AutoScaleFont = False
'    With Selection.Font
'        .Name = "新細明體"
'        .FontStyle = "粗體"
'        .Size = 14
'    End With
'    ActiveChart.Axes(xlCategory).Select
'    ActiveChart.Legend.Select
'    Selection.Left = 289
'    Selection.Top = 189
'    ActiveChart.PlotArea.Select
'    Selection.Width = 353
    
    
    
    
End Sub

Public Sub FitChart()
   Dim i As Integer, j As Integer
   Dim scatterChart As Chart
   Dim boxtrendChart As Chart
   Dim nowChart As Chart
   Dim nowSheet As Worksheet
   
   If IsExistSheet("PlotSetup") Then
      Set nowSheet = Worksheets("PlotSetup")
      For j = 1 To nowSheet.ChartObjects.Count
         Set nowChart = nowSheet.ChartObjects(j).Chart
         If nowChart.HasTitle Then
            If UCase(nowChart.ChartTitle.Text) = "SCATTER" Then Set scatterChart = nowChart
            If UCase(nowChart.ChartTitle.Text) = "BOXTREND" Then Set boxtrendChart = nowChart
         End If
      Next j
      If scatterChart Is Nothing And nowSheet.ChartObjects.Count >= 1 Then Set scatterChart = nowSheet.ChartObjects(1).Chart
      If boxtrendChart Is Nothing And nowSheet.ChartObjects.Count >= 2 Then Set boxtrendChart = nowSheet.ChartObjects(2).Chart
   Else
      Exit Sub
   End If
  
   ynSeries = (MsgBox("Fit series style?" & vbCrLf & vbCrLf & "The function must have some time to run.", vbOKCancel) = vbOK)
   For i = 1 To Worksheets.Count
      'Call AppendFile("c:\test.log", Worksheets(i).Name & vbCrLf)
      If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Then Call FitChartByType(Worksheets(i).Name, scatterChart, "SCATTER")
      If InStr(UCase(Worksheets(i).Name), "BOXTREND") > 0 Then Call FitChartByType(Worksheets(i).Name, boxtrendChart, "BOXTREND")
   Next i
    'Call FitChartByType("All_Chart", scatterChart, "SCATTER")
    Call GenChartSummary
End Sub
Public Function FitChartByType(mSheet As String, TemplateChart As Chart, mChartType As String)
   Dim j As Integer, m As Integer, i As Integer
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim nowLegend As Legend
   Dim nowSeries As Series
   Dim nowShape As Shape
   Dim nowAxis As Axis
   Dim nowDataLabels As DataLabels, TemplateDataLabel As DataLabel
   Dim nowPlotArea As PlotArea
   Dim boolTemp As Boolean
   Const NameList As String = "TARGET,CORNER,SS,FF,TT,GOLDEN,MEDIAN,USL,LSL"
   Dim tSeries As Series
   Dim tmp
   Dim YNMultiPara As Boolean
   
   On Error Resume Next
   If TemplateChart Is Nothing Then Exit Function
   
   Set nowSheet = Worksheets(mSheet)
   For j = 1 To nowSheet.ChartObjects.Count
      Set nowChart = nowSheet.ChartObjects(j).Chart
            '-----------------
            'Fit Chart Attribs
            '-----------------
            Set nowShape = nowSheet.Shapes(j)
            'Debug.Print nowShape.Name
            DoEvents
            If nowSheet.Name <> "All_Chart" Then
               With nowShape
                  .Top = 30 + (j - 1) * (TemplateChart.ChartArea.Height + 5)
                  .Left = 30
                  .width = TemplateChart.ChartArea.width
                  .Height = TemplateChart.ChartArea.Height
               End With
               DoEvents
            End If
            '----------
            'Fit Legend
            '----------
            nowChart.HasLegend = TemplateChart.HasLegend
            If nowChart.HasLegend Then
               Set nowLegend = nowChart.Legend
               'nowLegend.Position = TemplateChart.Legend.Position
               nowLegend.AutoScaleFont = False
               nowLegend.Font.Size = 8
               DoEvents
               nowLegend.Left = nowShape.width - nowLegend.width - 6 - 10
               nowLegend.Top = (nowShape.Height - nowLegend.Height) / 2 '+ 10
            End If
            DoEvents
            '-------------
            'Fit PlotArea
            '-------------
            Set nowPlotArea = nowChart.PlotArea
            nowPlotArea.width = nowLegend.Left - nowPlotArea.Left - 5
            
            Set nowPlotArea = Nothing
            Set nowLegend = Nothing
            Set nowShape = Nothing
            DoEvents
            '----------
            'Fit Axis
            '----------
            If mChartType <> "BOXTREND" Then
               Set nowAxis = nowChart.Axes(xlCategory)
               If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
               nowAxis.TickLabels.Font.Size = 10
               Set nowAxis = Nothing
            End If
            Set nowAxis = nowChart.Axes(xlValue)
            If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
            nowAxis.TickLabels.Font.Size = 10
            
            Set nowAxis = Nothing
            
      '----------
      'Fit Axis
      '----------
      If True Then ' X Axis
         Set nowAxis = nowChart.Axes(xlCategory, xlPrimary)
         With TemplateChart.Axes(xlCategory, xlPrimary)
            nowAxis.TickLabels.AutoScaleFont = True
            nowAxis.TickLabels.Font.Name = .TickLabels.Font.Name
            nowAxis.TickLabels.Font.Size = .TickLabels.Font.Size
            nowAxis.TickLabels.AutoScaleFont = False
            nowAxis.TickLabels.NumberFormatLocal = .TickLabels.NumberFormatLocal
            'If mChartType = "BOXTREND" Then nowAxis.TickLabels.NumberFormatLocal = "@"
            nowAxis.TickLabels.Orientation = .TickLabels.Orientation
            'End If
            nowAxis.HasTitle = .HasTitle
            If nowAxis.HasTitle Then
               nowAxis.AxisTitle.Font.Size = .AxisTitle.Font.Size
               nowAxis.AxisTitle.Font.Name = .AxisTitle.Font.Name
               nowAxis.AxisTitle.Font.ColorIndex = .AxisTitle.Font.ColorIndex
            End If
            nowAxis.HasMajorGridlines = .HasMajorGridlines
            If .HasMajorGridlines Then
                nowAxis.HasMajorGridlines = True
                nowAxis.MajorGridlines.Border.LineStyle = .MajorGridlines.Border.LineStyle
                nowAxis.MajorGridlines.Border.Weight = .MajorGridlines.Border.Weight
                nowAxis.MajorGridlines.Border.ColorIndex = .MajorGridlines.Border.ColorIndex
            End If
            nowAxis.HasMinorGridlines = .HasMinorGridlines
            If .HasMinorGridlines Then
                nowAxis.HasMinorGridlines = True
                nowAxis.MinorGridlines.Border.LineStyle = .MinorGridlines.Border.LineStyle
                nowAxis.MinorGridlines.Border.Weight = .MinorGridlines.Border.Weight
                nowAxis.MinorGridlines.Border.ColorIndex = .MinorGridlines.Border.ColorIndex
            End If
            
            
         End With
         Set nowAxis = Nothing
      End If
      If True Then ' Y Axis
         Set nowAxis = nowChart.Axes(xlValue, xlPrimary)
         With TemplateChart.Axes(xlValue, xlPrimary)
            nowAxis.TickLabels.AutoScaleFont = True
            nowAxis.TickLabels.Font.Name = .TickLabels.Font.Name
            nowAxis.TickLabels.Font.Size = .TickLabels.Font.Size
            nowAxis.TickLabels.AutoScaleFont = False
            nowAxis.TickLabels.NumberFormatLocal = .TickLabels.NumberFormatLocal
            'If mChartType = "BOXTREND" Then
            nowAxis.TickLabels.Orientation = .TickLabels.Orientation
            'End If
            nowAxis.HasTitle = .HasTitle
            If nowAxis.HasTitle Then
               nowAxis.AxisTitle.Font.Size = .AxisTitle.Font.Size
               nowAxis.AxisTitle.Font.Name = .AxisTitle.Font.Name
               nowAxis.AxisTitle.Font.ColorIndex = .AxisTitle.Font.ColorIndex
            End If
            nowAxis.HasMajorGridlines = .HasMajorGridlines
            If .HasMajorGridlines Then
                nowAxis.HasMajorGridlines = True
                nowAxis.MajorGridlines.Border.LineStyle = .MajorGridlines.Border.LineStyle
                nowAxis.MajorGridlines.Border.Weight = .MajorGridlines.Border.Weight
                nowAxis.MajorGridlines.Border.ColorIndex = .MajorGridlines.Border.ColorIndex
            End If
            nowAxis.HasMinorGridlines = .HasMinorGridlines
            If .HasMinorGridlines Then
                nowAxis.HasMinorGridlines = True
                nowAxis.MinorGridlines.Border.LineStyle = .MinorGridlines.Border.LineStyle
                nowAxis.MinorGridlines.Border.Weight = .MinorGridlines.Border.Weight
                nowAxis.MinorGridlines.Border.ColorIndex = .MinorGridlines.Border.ColorIndex
            End If
            
         End With
         Set nowAxis = Nothing
      End If
      DoEvents
      
      '-----------------
      'Fit Title
      '-----------------
      If nowChart.HasTitle Then
         With nowChart.ChartTitle
            .Font.Name = TemplateChart.ChartTitle.Font.Name
            .Font.Size = TemplateChart.ChartTitle.Font.Size
            .Font.ColorIndex = TemplateChart.ChartTitle.Font.ColorIndex
         End With
         DoEvents
      End If
      
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
               Set tSeries = TemplateChart.SeriesCollection("#" & CStr(m))
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
      If mChartType = "BOXTREND" Then
         ' Set DataLabel Style
         If nowChart.SeriesCollection(5).HasDataLabels And TemplateChart.SeriesCollection(5).HasDataLabels Then
            Set nowDataLabels = nowChart.SeriesCollection(5).DataLabels
            Set TemplateDataLabel = TemplateChart.SeriesCollection(5).Points(1).DataLabel
            nowDataLabels.AutoScaleFont = True
            nowDataLabels.Font.Name = TemplateDataLabel.Font.Name
            nowDataLabels.Font.FontStyle = TemplateDataLabel.Font.FontStyle
            nowDataLabels.Font.Size = TemplateDataLabel.Font.Size
            nowDataLabels.Font.Strikethrough = TemplateDataLabel.Font.Strikethrough
            nowDataLabels.Font.Superscript = TemplateDataLabel.Font.Superscript
            nowDataLabels.Font.Subscript = TemplateDataLabel.Font.Subscript
            nowDataLabels.Font.OutlineFont = TemplateDataLabel.Font.OutlineFont
            nowDataLabels.Font.Shadow = TemplateDataLabel.Font.Shadow
            nowDataLabels.Font.Underline = TemplateDataLabel.Font.Underline
            nowDataLabels.Font.ColorIndex = TemplateDataLabel.Font.ColorIndex
            nowDataLabels.Font.FontStyle = TemplateDataLabel.Font.FontStyle
            nowDataLabels.Font.Background = TemplateDataLabel.Font.Background
            YNMultiPara = False
            For Each tmp In nowChart.SeriesCollection(1).Values
                If tmp = "" Then YNMultiPara = True: Exit For
            Next tmp
            If Not YNMultiPara Then
                'nowDataLabels.Delete
                nowChart.SeriesCollection(5).ApplyDataLabels Type:=xlDataLabelsShowNone, LegendKey:=False
                nowChart.SeriesCollection(5).ApplyDataLabels Type:=xlDataLabelsShowValue, AutoText:=True, LegendKey:=False
                'nowDataLabels.Type = xlDataLabelsShowValue
            End If
         ElseIf nowChart.SeriesCollection(5).HasDataLabels And Not TemplateChart.SeriesCollection(5).HasDataLabels Then
            nowChart.SeriesCollection(5).HasDataLabels = TemplateChart.SeriesCollection(5).HasDataLabels
         End If
      End If
   Next j
   
   
   nowChart.Axes(xlCategory).TickLabels.AutoScaleFont = False
   nowChart.Axes(xlValue).TickLabels.AutoScaleFont = False
   nowChart.ChartTitle.AutoScaleFont = False
   
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

Public Function OLD_FitChartByType(mSheet As String, TemplateChart As Chart, mChartType As String)
   Dim j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim nowLegend As Legend
   Dim nowSeries As Series
   Dim nowShape As Shape
   Dim nowAxis As Axis
   Dim nowDataLabels As DataLabels, TemplateDataLabel As DataLabel
   Dim nowPlotArea As PlotArea
   
   On Error Resume Next
   If TemplateChart Is Nothing Then Exit Function
   
   Set nowSheet = Worksheets(mSheet)
   For j = 1 To nowSheet.ChartObjects.Count
      Set nowChart = nowSheet.ChartObjects(j).Chart
            '-----------------
            'Fit Chart Attribs
            '-----------------
            Set nowShape = nowSheet.Shapes(j)
            'Debug.Print nowShape.Name
            DoEvents
            If nowSheet.Name <> "All_Chart" Then
               With nowShape
                  .Top = 30 + (j - 1) * (300 + 100)
                  .Left = 30
                  .width = 450
                  .Height = 300
               End With
            DoEvents
            End If
            '----------
            'Fit Legend
            '----------
            If nowChart.HasLegend Then
               Set nowLegend = nowChart.Legend
               nowLegend.AutoScaleFont = False
               nowLegend.Font.Size = 8
               DoEvents
               nowLegend.Left = nowShape.width - nowLegend.width - 6 - 10
               nowLegend.Top = (nowShape.Height - nowLegend.Height) / 2 '+ 10
            End If
            DoEvents
            '-------------
            'Fit PlotArea
            '-------------
            Set nowPlotArea = nowChart.PlotArea
            nowPlotArea.width = nowLegend.Left - nowPlotArea.Left - 5
            
            Set nowPlotArea = Nothing
            Set nowLegend = Nothing
            Set nowShape = Nothing
            DoEvents
            '----------
            'Fit Axis
            '----------
            If mChartType <> "BOXTREND" Then
               Set nowAxis = nowChart.Axes(xlCategory)
               If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
               nowAxis.TickLabels.Font.Size = 10
            End If
            Set nowAxis = nowChart.Axes(xlValue)
            If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
            nowAxis.TickLabels.Font.Size = 10
            Set nowAxis = Nothing
            
'      '----------
'      'Fit Axis
'      '----------
'      If True Then ' X Axis
'         Set nowAxis = nowChart.Axes(xlCategory, xlPrimary)
'         With TemplateChart.Axes(xlCategory, xlPrimary)
'            nowAxis.TickLabels.AutoScaleFont = True
'            nowAxis.TickLabels.Font.Name = .TickLabels.Font.Name
'            nowAxis.TickLabels.Font.Size = .TickLabels.Font.Size
'            'If mChartType = "BOXTREND" Then
'            nowAxis.TickLabels.Orientation = .TickLabels.Orientation
'            'End If
'            nowAxis.HasTitle = .HasTitle
'            If nowAxis.HasTitle Then
'               nowAxis.AxisTitle.Font.Size = .AxisTitle.Font.Size
'               nowAxis.AxisTitle.Font.Name = .AxisTitle.Font.Name
'            End If
'         End With
'      End If
'      If True Then ' Y Axis
'         Set nowAxis = nowChart.Axes(xlValue, xlPrimary)
'         With TemplateChart.Axes(xlValue, xlPrimary)
'            nowAxis.TickLabels.AutoScaleFont = True
'            nowAxis.TickLabels.Font.Name = .TickLabels.Font.Name
'            nowAxis.TickLabels.Font.Size = .TickLabels.Font.Size
'            'If mChartType = "BOXTREND" Then
'            nowAxis.TickLabels.Orientation = .TickLabels.Orientation
'            'End If
'            nowAxis.HasTitle = .HasTitle
'            If nowAxis.HasTitle Then
'               nowAxis.AxisTitle.Font.Size = .AxisTitle.Font.Size
'               nowAxis.AxisTitle.Font.Name = .AxisTitle.Font.Name
'            End If
'         End With
'      End If

      '-----------------
      'Fit Series Style
      '-----------------
      If mChartType = "BOXTREND" Then
         For m = 1 To nowChart.SeriesCollection.Count
            Set nowSeries = nowChart.SeriesCollection(m)
            With TemplateChart.SeriesCollection(m)
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
         Next m
         ' Set DataLabel Style
         If nowChart.SeriesCollection(5).HasDataLabels And TemplateChart.SeriesCollection(5).HasDataLabels Then
            Set nowDataLabels = nowChart.SeriesCollection(5).DataLabels
            Set TemplateDataLabel = TemplateChart.SeriesCollection(5).Points(1).DataLabel
            nowDataLabels.AutoScaleFont = True
            nowDataLabels.Font.Name = TemplateDataLabel.Font.Name
            nowDataLabels.Font.FontStyle = TemplateDataLabel.Font.FontStyle
            nowDataLabels.Font.Size = TemplateDataLabel.Font.Size
            nowDataLabels.Font.Strikethrough = TemplateDataLabel.Font.Strikethrough
            nowDataLabels.Font.Superscript = TemplateDataLabel.Font.Superscript
            nowDataLabels.Font.Subscript = TemplateDataLabel.Font.Subscript
            nowDataLabels.Font.OutlineFont = TemplateDataLabel.Font.OutlineFont
            nowDataLabels.Font.Shadow = TemplateDataLabel.Font.Shadow
            nowDataLabels.Font.Underline = TemplateDataLabel.Font.Underline
            nowDataLabels.Font.ColorIndex = TemplateDataLabel.Font.ColorIndex
            nowDataLabels.Font.FontStyle = TemplateDataLabel.Font.FontStyle
            nowDataLabels.Font.Background = TemplateDataLabel.Font.Background
         ElseIf nowChart.SeriesCollection(5).HasDataLabels And Not TemplateChart.SeriesCollection(5).HasDataLabels Then
            nowChart.SeriesCollection(5).HasDataLabels = TemplateChart.SeriesCollection(5).HasDataLabels
         End If
      End If
   Next j
End Function


Public Sub Manual_SimpleSummary()
   Call SimpleSummary("SPEC_Simple", "Report_Simple")
End Sub

Private Function SimpleSummary(inSheet As String, outSheet As String)
   Dim specRange As Range
   Dim dataRange As Range
   Dim waferList() As String
   Dim iCol As Long, iRow As Long
   Dim outSpec As Integer
   Dim HColCount As Integer
   Dim HCol As Integer
   Dim siteNum As Integer
   Dim ProductID As String
   Dim mFactor
   Dim FactorSign1 As String, FactorSign2 As String
   Dim nowRange As Range
   Dim reValue As Variant
   Dim nowParameter As String
   Dim i As Long, j As Long
   
   Application.ScreenUpdating = False
   'SiteNum = Worksheets(TempDSheet).UsedRange.Columns.Count - 2
   siteNum = getSiteNum(dSheet)
   'ProductID = Trim(Worksheets(SRawData).Range("A2"))
   If ProductID = "A064A" Or ProductID = "A100A" Or ProductID = "A118" Then
      HColCount = 5
      HCol = 4
   Else
      HColCount = 4
      HCol = 3
   End If
   ' get Waferlist
   Call GetWaferList(dSheet, waferList)
   
   AddSheet (outSheet)
   'Worksheets(SSheet).UsedRange.Copy Worksheets(outsheet).Range("A1")
   'Worksheets(outsheet).Range("B:B").Delete Shift:=xlShiftToLeft
   'Worksheets(SRawData).Range("A1:" & "G" & CStr(Worksheets(SRawData).UsedRange.Rows.Count)).Copy Worksheets(outSheet).Range("A1")
   Worksheets(inSheet).UsedRange.Copy Worksheets(outSheet).Range("A1")
   'Worksheets(outSheet).Cells(2, 3) = UBound(WaferList) + 1
   Set specRange = Worksheets(outSheet).UsedRange

   
   ' fill wafer header
   Application.DisplayAlerts = False
   Worksheets(outSheet).Activate
   iRow = 2
   iCol = specRange.Columns.Count + 1
   For i = 0 To UBound(waferList)
      With Worksheets(outSheet)
         .Range(Num2Letter(iCol + i * HColCount) & CStr(iRow) & ":" & Num2Letter(iCol + i * HColCount + HCol) & CStr(iRow)).Select
         Selection.Merge
         Selection.Value = "#" & Trim(waferList(i))
         .Range(Num2Letter(iCol + i * HColCount) & CStr(iRow + 1) & ":" & Num2Letter(iCol + i * HColCount + HCol) & CStr(iRow + 1)).Select
         Selection.Cells(1, 1) = "Median"
         Selection.Cells(1, 2) = "Average"
         Selection.Cells(1, 3) = "3 Sigma"
         Selection.Cells(1, 4) = "Yield"
         If ProductID = "A064A" Or ProductID = "A100A" Or ProductID = "A118" Then Selection.Cells(1, 5) = "%"
      End With
      ' Add Split line
      Worksheets(outSheet).Range(Num2Letter(iCol + i * HColCount + HCol) & CStr(iRow) & ":" & Num2Letter(iCol + i * HColCount + HCol) & CStr(Worksheets(outSheet).UsedRange.Rows.Count)).Select
      With Selection.Borders(xlEdgeRight)
         .LineStyle = xlContinuous
         .Weight = xlMedium
         .ColorIndex = xlAutomatic
      End With
   Next i
   ' Add Split line
   Worksheets(outSheet).Range(Num2Letter(iCol) & CStr(iRow) & ":" & Num2Letter(iCol) & CStr(Worksheets(outSheet).UsedRange.Rows.Count)).Select
   With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .Weight = xlMedium
      .ColorIndex = xlAutomatic
   End With
   Worksheets(outSheet).Range(Num2Letter(iCol) & CStr(iRow) & ":" & Num2Letter(iCol + UBound(waferList) * HColCount + HCol) & CStr(iRow + 1)).Select
   'Selection.Font.Bold = True
   Selection.HorizontalAlignment = xlCenter
   Selection.Interior.ColorIndex = 8
   'Selection.Borders.LineStyle = xlContinuous
   Application.DisplayAlerts = True
   ' fill value
   Worksheets(outSheet).Activate
   iCol = specRange.Columns.Count + 1
   With Worksheets(outSheet)
      For i = 4 To .UsedRange.Rows.Count
         For j = 0 To UBound(waferList)
            .Range(Num2Letter(iCol + j * HColCount) & CStr(i) & ":" & Num2Letter(iCol + j * HColCount + HCol) & CStr(i)).Select
            nowParameter = Worksheets(outSheet).Cells(i, 2)
            Set reValue = getRangeByPara(waferList(j), nowParameter)
            If reValue Is Nothing Then Exit For
            Set nowRange = reValue
            Selection.Cells(1, 1) = Application.WorksheetFunction.Median(nowRange)
            Selection.Cells(1, 2) = Application.WorksheetFunction.Average(nowRange)
            Selection.Cells(1, 3) = Application.WorksheetFunction.StDev(nowRange) * 3
            If Worksheets(SSheet).Cells(i, 2) <> "" Then
               Selection.Cells(1, 1) = Selection.Cells(1, 1) * Worksheets(SSheet).Cells(i, 2)
               Selection.Cells(1, 2) = Selection.Cells(1, 2) * Worksheets(SSheet).Cells(i, 2)
               Selection.Cells(1, 3) = Abs(Selection.Cells(1, 3) * Worksheets(SSheet).Cells(i, 2))
            End If
            If ProductID = "A064A" Or ProductID = "A100A" Or ProductID = "A118" Then   ' A064A (%)
               If Trim(.Range(Num2Letter(5) & CStr(i))) <> "" Then _
                  Selection.Cells(1, 5) = Format((Selection.Cells(1, 1) - .Range(Num2Letter(5) & CStr(i))) / .Range(Num2Letter(5) & CStr(i)), "0.00%")
            End If
            'Selection.Cells(1, 4) = Selection.Cells(1, 3) * 3
            If Trim(.Range(Num2Letter(4) & CStr(i))) <> "" Or Trim(.Range(Num2Letter(6) & CStr(i))) <> "" Then
               Selection.Range("A1").Select  ' **Median
               Selection.FormatConditions.Delete
               Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & "F" & "" & CStr(i)
               Selection.FormatConditions(1).Font.ColorIndex = 3
               Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & "D" & "" & CStr(i)
               Selection.FormatConditions(2).Font.ColorIndex = 4
               'Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, _
               '   Formula1:="=$" & "D" & "$" & CStr(i), Formula2:="=$" & "F" & "$" & CStr(i)
               'Selection.FormatConditions(1).Font.ColorIndex = 3
               outSpec = 0
               If Worksheets(SSheet).Cells(i, 2) <> "" Then
                  mFactor = Worksheets(SSheet).Cells(i, 2)
               Else
                  mFactor = 1
               End If
               If mFactor >= 0 Then
                  FactorSign1 = "<"
                  FactorSign2 = ">"
               Else
                  FactorSign1 = ">"
                  FactorSign2 = "<"
               End If
               If Trim(.Range(Num2Letter(4) & CStr(i))) <> "" Then _
                  outSpec = outSpec + Application.WorksheetFunction.CountIf(nowRange, FactorSign1 & .Range(Num2Letter(4) & CStr(i)) / mFactor)
               If Trim(.Range(Num2Letter(6) & CStr(i))) <> "" Then _
                  outSpec = outSpec + Application.WorksheetFunction.CountIf(nowRange, FactorSign2 & .Range(Num2Letter(6) & CStr(i)) / mFactor)
               Selection.Range("D1") = Format((siteNum - outSpec) / siteNum, "0.00%")
               If outSpec <> 0 Then Selection.Range("D1").Font.ColorIndex = 3
            Else
               If Selection.Cells(1, 1) <> "" Then Selection.Cells(1, 4) = Format(1, "0.00%")
            End If
         Next j
      Next i
   End With
   
   'Sheet format
   'Call SummaryFormatByUnit(outSheet)
   Worksheets(outSheet).Activate
   Worksheets(outSheet).Cells.Select
   Selection.Font.Size = 10
   Selection.Font.Name = "Century Gothic"
   Worksheets(outSheet).Range("A1:" & Num2Letter(Worksheets(outSheet).UsedRange.Columns.Count) & CStr(3)).Select
   Selection.Font.Size = 12
   Worksheets(outSheet).Range("A4:" & "A" & CStr(Worksheets(outSheet).UsedRange.Rows.Count)).Select
   Selection.Font.Size = 12
   ActiveWindow.Zoom = 75
   Worksheets(outSheet).Cells.Select
   Selection.Columns.AutoFit
   Selection.Rows.AutoFit
   Worksheets(outSheet).Range("A4").Select
   ActiveWindow.FreezePanes = True
   Application.ScreenUpdating = True
   'MsgBox "sss"
End Function

Sub ExportChartAttrib()
    Dim nowSheet As Worksheet, sheetAllChart As Worksheet
    Const strHeader As String = "Chart,ChartTitle,XScale,XNumberFormat,XMax,XMin,Xmajor,Xminor,YScale,YNumberFormat,YMax,YMin,Ymajor,Yminor"
    Dim tempA
    Dim i As Integer, j As Integer
    Dim tmpStr As String
    Const sChar As String = "§"
    
    If IsExistSheet("All_Chart") Then
        Set sheetAllChart = Worksheets("All_Chart")
    Else
        Exit Sub
    End If
    Set nowSheet = AddSheet("ChartAttrib")
    tempA = Split(strHeader, ",")
    For i = 0 To UBound(tempA)
        nowSheet.Cells(1, i + 1) = tempA(i)
    Next i
    
    For i = 2 To sheetAllChart.UsedRange.Rows.Count
        If sheetAllChart.Cells(i, 1) = "" Then Exit For
        'nowSheet.Cells(i, 1) = sheetAllChart.Cells(i, 2).Value
        tmpStr = getChartAttrib(sheetAllChart.Cells(i, 2).Value)
        For j = 1 To UBound(tempA) + 1
            nowSheet.Cells(i, j) = "'" & getCOL(tmpStr, sChar, j)
        Next j
    Next i
    nowSheet.Columns.AutoFit
    Set nowSheet = Nothing
End Sub

Function getChartAttrib(mSheet As String)
    Dim tmpStr As String
    Dim reStr As String
    Dim nowChart As Chart, nowAxis As Axis
    Const sChar As String = "§"
    
    On Error Resume Next
    
    If Not IsExistSheet(mSheet) Then Exit Function
    If Worksheets(mSheet).ChartObjects.Count > 0 Then Set nowChart = Worksheets(mSheet).ChartObjects(1).Chart
    'Chart of Chart
    Set nowChart = Worksheets("All_Chart").ChartObjects(mSheet).Chart
    reStr = mSheet
    reStr = reStr & sChar & IIf(nowChart.HasTitle, nowChart.ChartTitle.Text, "")
    Set nowAxis = nowChart.Axes(xlCategory, xlPrimary)
    reStr = reStr & sChar & IIf(nowAxis.ScaleType = xlLinear, "Linear", "Log")
    reStr = reStr & sChar & nowAxis.TickLabels.NumberFormatLocal
    reStr = reStr & sChar & nowAxis.MaximumScale
    reStr = reStr & sChar & nowAxis.MinimumScale
    reStr = reStr & sChar & nowAxis.MajorUnit
    reStr = reStr & sChar & nowAxis.MinorUnit
     Set nowAxis = nowChart.Axes(xlValue, xlPrimary)
    reStr = reStr & sChar & IIf(nowAxis.ScaleType = xlLinear, "Linear", "Log")
    reStr = reStr & sChar & nowAxis.TickLabels.NumberFormatLocal
    reStr = reStr & sChar & nowAxis.MaximumScale
    reStr = reStr & sChar & nowAxis.MinimumScale
    reStr = reStr & sChar & nowAxis.MajorUnit
    reStr = reStr & sChar & nowAxis.MinorUnit
    getChartAttrib = reStr
End Function

Sub ImportChartAttrib()
    Dim nowSheet As Worksheet
    'Const strHeader As String = "Chart,ChartTitle,XScale,XMax,XMin,Xmajor,Xminor,YScale,YMax,YMin,Ymajor,Yminor"
    'Dim TempA
    Dim i As Integer, j As Integer
    Dim tmpStr As String
    Const sChar As String = "§"
    
    If IsExistSheet("ChartAttrib") Then
        Set nowSheet = Worksheets("ChartAttrib")
    Else
        Exit Sub
    End If

    
    For i = 2 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(i, 1) = "" Then Exit For
        tmpStr = nowSheet.Cells(i, 1)
        For j = 2 To nowSheet.UsedRange.Columns.Count
            tmpStr = tmpStr & sChar & nowSheet.Cells(i, j).Text
        Next j
        Call setChartAttrib(tmpStr)
    Next i
    
End Sub

Function setChartAttrib(tmpStr As String)
    Const sChar As String = "§"
    Const ChartID = 1
    Const ChartTitle = 2
    Const XScale = 3
    Const XNumberFormat = 4
    Const xMax = 5
    Const xMin = 6
    Const Xmajor = 7
    Const Xminor = 8
    Const YScale = 9
    Const YNumberFormat = 10
    Const yMax = 11
    Const yMin = 12
    Const Ymajor = 13
    Const Yminor = 14

'    Dim TmpStr As String
'    Dim reStr As String
    Dim mSheet As String
    Dim nowChart As Chart, nowAxis As Axis
    On Error Resume Next

    'For Excel XP
    If Application.Version <> "9.0" Then tmpStr = Replace(tmpStr, "G/通用格式", "")
    
    mSheet = getCOL(tmpStr, sChar, 1)
    If Not IsExistSheet(mSheet) Then Exit Function
    If Worksheets(mSheet).ChartObjects.Count > 0 Then Set nowChart = Worksheets(mSheet).ChartObjects(1).Chart
    'Chart of All_Chart
    Set nowChart = Worksheets("All_Chart").ChartObjects(mSheet).Chart
    If nowChart.HasTitle Then nowChart.ChartTitle.Text = getCOL(tmpStr, sChar, ChartTitle)

    'X-Axis
    '-------------------------------------------------
    Set nowAxis = nowChart.Axes(xlCategory, xlPrimary)
    nowAxis.TickLabels.NumberFormatLocal = getCOL(tmpStr, sChar, XNumberFormat)
    If getCOL(tmpStr, sChar, xMax) <> "" Then
        nowAxis.MaximumScale = Val(getCOL(tmpStr, sChar, xMax))
    Else
        nowAxis.MaximumScaleIsAuto = True
    End If
    If getCOL(tmpStr, sChar, xMin) <> "" Then
        nowAxis.MinimumScale = Val(getCOL(tmpStr, sChar, xMin))
    Else
        nowAxis.MinimumScaleIsAuto = True
    End If
    If getCOL(tmpStr, sChar, Xmajor) <> "" Then
        nowAxis.MajorUnit = Val(getCOL(tmpStr, sChar, Xmajor))
    Else
        nowAxis.MajorUnitIsAuto = True
    End If
    If getCOL(tmpStr, sChar, Xminor) <> "" Then
        nowAxis.MinorUnit = Val(getCOL(tmpStr, sChar, Xminor))
    Else
        nowAxis.MinorUnitIsAuto = True
    End If
    nowAxis.CrossesAt = xlCustom
    nowAxis.CrossesAt = nowAxis.MinimumScale
    nowAxis.ScaleType = IIf(UCase(getCOL(tmpStr, sChar, XScale)) = "LINEAR", xlLinear, xlLogarithmic)
    
    'X-Axis
    '-------------------------------------------------
    Set nowAxis = nowChart.Axes(xlValue, xlPrimary)
    nowAxis.TickLabels.NumberFormatLocal = getCOL(tmpStr, ",", YNumberFormat)
    If getCOL(tmpStr, sChar, yMax) <> "" Then
        nowAxis.MaximumScale = Val(getCOL(tmpStr, sChar, yMax))
    Else
        nowAxis.MaximumScaleIsAuto = True
    End If
    If getCOL(tmpStr, sChar, yMin) <> "" Then
        nowAxis.MinimumScale = Val(getCOL(tmpStr, sChar, yMin))
    Else
        nowAxis.MinimumScaleIsAuto = True
    End If
    
    If getCOL(tmpStr, sChar, Ymajor) <> "" Then
        nowAxis.MajorUnit = Val(getCOL(tmpStr, sChar, Ymajor))
    Else
        nowAxis.MajorUnitIsAuto = True
    End If
    If getCOL(tmpStr, sChar, Yminor) <> "" Then
        nowAxis.MinorUnit = Val(getCOL(tmpStr, sChar, Yminor))
    Else
        nowAxis.MinorUnitIsAuto = True
    End If
    nowAxis.CrossesAt = xlCustom
    nowAxis.CrossesAt = nowAxis.MinimumScale
    nowAxis.ScaleType = IIf(UCase(getCOL(tmpStr, sChar, YScale)) = "LINEAR", xlLinear, xlLogarithmic)
End Function

Sub Manual_RawDataReSorting()
    Dim nowSheet As Worksheet
    Dim i As Long
    Dim nowRange As Range
    
    Set nowSheet = Worksheets("Data")
    For i = 1 To nowSheet.Names.Count
        If InStr(nowSheet.Names(i).Name, "wafer") > 0 Then
            Set nowRange = nowSheet.Names(i).RefersToRange
            If nowRange.Cells(2, 1).Text <> "1" Then
                nowRange.Sort nowRange.Range("A1"), xlAscending, , , , , , xlYes
            Else
                nowRange.Sort nowRange.Range("B1"), xlAscending, , , , , , xlYes
            End If
        End If
    Next i
End Sub

Function removeTTFFSS()
    Dim nowSheet As Worksheet
    Dim i As Integer, j As Integer
    
    For j = 1 To Worksheets.Count
        If Left(Worksheets(j).Name, 7) = "SCATTER" Then
            Set nowSheet = Worksheets(j)
            For i = nowSheet.UsedRange.Columns.Count To 3 Step -1
                Select Case nowSheet.Cells(1, i)
                    Case "TT", "SS", "FF"
                        nowSheet.Columns(N2L(i) & ":" & N2L(i + 1)).Delete Shift:=xlToLeft
                End Select
            Next i
        End If
    Next j
End Function

Sub ReDraw()
   'Call GenChartHeader
   'Call GenScatter
   'Call GenBoxTrend
   'Call GenCumulative
   If Not IsExistSheet("PlotSetup") Then Exit Sub
   Call removeTTFFSS
   Call DioPlotAllChart
   'Call Manual_FitLegend
   Call new_FitChart
   Call CornerCount  'New function
   Call RawdataRange 'New function
   Call GenChartSummary
   'Application.StatusBar = "Finished!"
End Sub


Sub Manual_RemoveChartLink()
    Dim i As Integer, j As Integer
    Dim nowSheet As Worksheet
    Dim nowChart As Chart
    Dim nowShape As Shape
    Dim iSheet As Integer
    
    If Not IsExistSheet("All_Chart") Then Exit Sub
    
    Set nowSheet = Worksheets("All_Chart")
    
    For i = 1 To nowSheet.ChartObjects.Count
        Set nowChart = nowSheet.ChartObjects(i).Chart
        'Debug.Print nowChart.Name
        For j = nowChart.Shapes.Count To 1 Step -1
            If InStr(nowChart.Shapes(j).Hyperlink.ScreenTip, "SCATTER") > 0 Or _
               InStr(nowChart.Shapes(j).Hyperlink.ScreenTip, "BOXTREND") > 0 Or _
               InStr(nowChart.Shapes(j).Hyperlink.ScreenTip, "CUMULATIVE") > 0 Then _
               nowChart.Shapes(j).Delete
            'Debug.Print nowChart.Shapes(i).Hyperlink.ScreenTip
            'Debug.Print nowChart.Shapes(i).AlternativeText
        Next j
    Next i
    
    For iSheet = 1 To Worksheets.Count
        If InStr(UCase(Worksheets(iSheet).Name), "SCATTER") > 0 Or _
           InStr(UCase(Worksheets(iSheet).Name), "BOXTREND") > 0 Or _
           InStr(UCase(Worksheets(iSheet).Name), "CUMULATIVE") > 0 Then
    
            Set nowSheet = Worksheets(iSheet)
            Debug.Print nowSheet.Name
            For i = 1 To nowSheet.ChartObjects.Count
                Set nowChart = nowSheet.ChartObjects(i).Chart
                'Debug.Print nowChart.Name
                For j = nowChart.Shapes.Count To 1 Step -1
                    If InStr(nowChart.Shapes(j).Hyperlink.ScreenTip, "Chart List") > 0 Then _
                       nowChart.Shapes(j).Delete
                    'Debug.Print nowChart.Shapes(i).Hyperlink.ScreenTip
                    'Debug.Print nowChart.Shapes(i).AlternativeText
                Next j
            Next i
           
        End If
    Next iSheet
End Sub

Sub Manual_vincent_Ratio()
    Dim xYN As Boolean
    Dim yYN As Boolean
    Dim ynString As String
    Dim iRow As Long, iCol As Long
    Dim bRow As Long
    Dim nowSheet As Worksheet
    Dim i As Long
    
    ynString = InputBox("請輸入 X,Y 是否要取 Ratio:", , "0,1")
    xYN = IIf(getCOL(ynString, ",", 1) = "1", True, False)
    yYN = IIf(getCOL(ynString, ",", 2) = "1", True, False)
    
    Set nowSheet = ActiveSheet
    With nowSheet
        For i = 1 To nowSheet.UsedRange.Rows.Count
            If nowSheet.Cells(i, 1) = "Y" Then bRow = i + 1: Exit For
        Next i
        For iCol = 3 To nowSheet.UsedRange.Columns.Count Step 2
            For iRow = nowSheet.UsedRange.Rows.Count To bRow Step -1
                If nowSheet.Cells(iRow, iCol) <> "" And yYN Then
                   nowSheet.Cells(iRow, iCol) = "'=(" & CStr(nowSheet.Cells(iRow, iCol)) & "-" & _
                                               CStr(nowSheet.Cells(bRow, iCol)) & ")/" & _
                                               CStr(nowSheet.Cells(bRow, iCol))
                End If
                If nowSheet.Cells(iRow, iCol + 1) <> "" And xYN Then
                   nowSheet.Cells(iRow, iCol + 1) = "'=(" & CStr(nowSheet.Cells(iRow, iCol + 1)) & "-" & _
                                               CStr(nowSheet.Cells(bRow, iCol + 1)) & ")/" & _
                                               CStr(nowSheet.Cells(bRow, iCol + 1))
                End If
                
            Next iRow
        Next iCol
    End With
    
    MsgBox "ok"
End Sub

Sub CPK_Table()
    Dim i As Long, j As Long
    Dim setSheet As Worksheet, nowSheet As Worksheet
    Dim setRange As Range
    Dim Class1 As String, Class2 As String, notClass As String
    Dim waferList() As String
    Dim nItem As Integer, nPass As Integer
    Dim tmp As String
    
    'Get setRange
    If Not IsExistSheet("CPK_Option") Then Exit Sub
    Set setSheet = Worksheets("CPK_Option")
    For i = 1 To setSheet.UsedRange.Columns.Count
        If UCase(Trim(setSheet.Cells(1, i))) = "CPK_TABLE" Then
            'Debug.Print i
            'Debug.Print setSheet.Cells(1, i).End(xlDown).Address
            Set setRange = setSheet.Range(setSheet.Cells(2, i), setSheet.Cells(1, i).End(xlDown))
            'Debug.Print setRange.Address
            Exit For
        End If
    Next i
             
    Set nowSheet = AddSheet("CPK_Table")
    nowSheet.Cells(1, 1) = getCOL(Worksheets(dSheet).Range("B3"), ":", 2)
    nowSheet.Cells(1, 3) = "Lot"
    nowSheet.Cells(2, 1) = "Category"
    nowSheet.Cells(2, 2) = "S & M Item"
    nowSheet.Cells(2, 3) = "pass item"
    nowSheet.Cells(2, 4) = "pass %"
    nowSheet.Range("C1:D1").Merge
    With nowSheet.Range("A1:D2")
        .Interior.ColorIndex = 20
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    setRange.Copy nowSheet.Range("A3")
    
    Call GetWaferList(dSheet, waferList)
    
    For j = 3 To nowSheet.UsedRange.Rows.Count
        Class1 = "Class"
        Class2 = "Class2"
        notClass = "notClass"
        tmp = getCOL(getCOL(nowSheet.Cells(j, 1), "(", 2), ")", 1)
        If InStr(tmp, "&") > 0 Then
            Class1 = getCOL(tmp, "&", 1)
            Class2 = getCOL(tmp, "&", 2)
        Else
            Class1 = tmp
        End If
        If Class1 = "RS" Then notClass = "RS_M"
        For i = 0 To UBound(waferList)
            If j = 3 Then
                nowSheet.Cells(1, 5 + i * 2) = "#" & waferList(i)
                With nowSheet.Range(nowSheet.Cells(1, 5 + i * 2), nowSheet.Cells(1, 5 + i * 2 + 1))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .Interior.ColorIndex = 38
                    .Borders.LineStyle = xlContinuous
                End With
                nowSheet.Cells(2, 5 + i * 2) = "pass item"
                nowSheet.Cells(2, 5 + i * 2 + 1) = "pass %"
                With nowSheet.Range(nowSheet.Cells(2, 5 + i * 2), nowSheet.Cells(2, 5 + i * 2 + 1))
                    .HorizontalAlignment = xlCenter
                    .Interior.ColorIndex = 20
                    .Borders.LineStyle = xlContinuous
                End With
            End If
            tmp = getPassItemByClass(waferList(i), Class1, Class2, notClass)
            nItem = CInt(getCOL(tmp, ",", 1))
            nPass = CInt(getCOL(tmp, ",", 2))
            If i = 0 Then nowSheet.Cells(j, 2) = nItem
            nowSheet.Cells(j, 5 + i * 2) = nPass
            nowSheet.Cells(j, 5 + i * 2 + 1) = Format(nPass / nItem, "00.0%")
        Next i
    Next j
    nowSheet.Columns.AutoFit
End Sub

Function getPassItemByClass(ByVal mWafer As String, ByVal mClass As String, Optional mClass2 As String = "Class2", Optional notClass As String = "notClass")
    Dim nowSheet As Worksheet
    Dim iRow As Long
    Dim iCol As Long
    Dim nItem As Integer, nPass As Integer
    
    If Not IsExistSheet("All_Summary") Then Exit Function
    
    Set nowSheet = Worksheets("All_Summary")
    
    For iCol = 7 To nowSheet.UsedRange.Columns.Count
        If nowSheet.Cells(1, iCol) = mWafer And nowSheet.Cells(2, iCol) = "CPK" Then
            For iRow = 3 To nowSheet.UsedRange.Rows.Count
                If (Left(UCase(nowSheet.Cells(iRow, 2)), Len(mClass)) = UCase(mClass) _
                   Or Left(UCase(nowSheet.Cells(iRow, 2)), Len(mClass2)) = UCase(mClass2)) _
                   And Left(UCase(nowSheet.Cells(iRow, 2)), Len(notClass)) <> UCase(notClass) Then
                    'Debug.Print nowSheet.Cells(iRow, 2)
                    nItem = nItem + 1
                    If nowSheet.Cells(iRow, iCol) >= 1.33 Then nPass = nPass + 1
                End If
            Next iRow
        End If
    Next iCol
    
    getPassItemByClass = nItem & "," & nPass
End Function

' For 莊子健
Sub AddBoxSigmaPercentage()
    Dim nowSheet As Worksheet, nowChart As Chart, nowSeries As Series, nowRange As Range, nowAxis As Axis
    Dim bRow As Long, iCol As Long, iRow As Long, i As Long, j As Long
    Dim valueStr As String, mColl As New Collection
    Dim tempA As Variant
    
    
    For i = 1 To Worksheets.Count
        On Error GoTo myError
        If UCase(Left(Worksheets(i).Name, 8)) = "BOXTREND" Then
            Set nowSheet = Worksheets(i)
            iCol = nowSheet.UsedRange.Columns.Count
            iRow = nowSheet.Range("E5").CurrentRegion.Rows.Count
            bRow = nowSheet.Range("D4").End(xlDown).row + 1
            For j = 1 To mColl.Count: mColl.Remove 1: Next j
            For j = 5 To iCol
                If nowSheet.Cells(1, j) <> "" Then
                    Set nowRange = nowSheet.Range(N2L(j) & CStr(bRow) & ":" & N2L(j) & CStr(iRow))
                    mColl.Add Round(3 * WorksheetFunction.StDev(nowRange) / WorksheetFunction.Average(nowRange), 1)
                Else
                    mColl.Add 0, CStr(j)
                End If
            Next j
            valueStr = "{" & mColl(1)
            'valueStr = mColl(1)
            For j = 2 To mColl.Count: valueStr = valueStr & "," & mColl(j): Next j
            valueStr = valueStr & "}"
            'Debug.Print valueStr
            'TempA = Split(valueStr, ",")
            Set nowChart = nowSheet.ChartObjects(1).Chart
            Set nowSeries = nowChart.SeriesCollection.NewSeries
            nowSeries.Name = "Sigma%"
            nowSeries.chartType = xlXYScatterLines
            nowSeries.Values = valueStr
            nowSeries.MarkerStyle = xlMarkerStyleDiamond
            'nowSeries.MarkerBackgroundColorIndex = nowSeries.MarkerForegroundColor
            nowSeries.MarkerBackgroundColorIndex = 2
            nowSeries.MarkerForegroundColorIndex = 5
            nowSeries.Border.ColorIndex = 5
            nowSeries.Border.Weight = xlThick
            
            For j = 2 To nowSeries.Points.Count - 1
                If WorksheetFunction.index(nowSeries.Values, j) = 0 Then
                    nowSeries.Points(j).Border.LineStyle = xlNone
                    nowSeries.Points(j + 1).Border.LineStyle = xlNone
                    nowSeries.Points(j).MarkerStyle = xlNone
                End If
            Next j
        
            'nowSeries.MarkerSize = 10
            nowSeries.AxisGroup = xlSecondary
            'nowSeries.Values = valueStr
            Set nowAxis = nowChart.Axes(xlValue, xlSecondary)
            nowAxis.TickLabels.NumberFormatLocal = "0%"
            nowAxis.HasTitle = True
            nowAxis.AxisTitle.Characters.Text = "3 Sigma%"
            nowAxis.AxisTitle.Font.ColorIndex = 5
            nowAxis.TickLabels.Font.ColorIndex = 5
        End If
myError:
        
    Next i

End Sub

' For 林榮祥
Public Function Manual_MedianScore(ByVal mSheet As String)
    Dim nowSheet As Worksheet
    Dim iCol As Long, iRow As Long
    Dim i As Long
    Dim Sum As Integer, Pass As Integer
    
    Set nowSheet = Worksheets(mSheet)
    
    iRow = nowSheet.UsedRange.Rows.Count + 1
    For iCol = 7 To nowSheet.UsedRange.Columns.Count
        If nowSheet.Cells(2, iCol) = "Median" Then
            For i = 3 To nowSheet.UsedRange.Rows.Count
                If nowSheet.Cells(i, iCol) <> "" Then Sum = Sum + 1
                If nowSheet.Cells(i, iCol) > nowSheet.Cells(i, 4) And nowSheet.Cells(i, iCol) < nowSheet.Cells(i, 6) Then Pass = Pass + 1
            Next i
            nowSheet.Cells(iRow, iCol) = Pass / Sum
            nowSheet.Cells(iRow, iCol).NumberFormat = "0%"
        End If
    Next iCol
End Function

'For 蘇建中 2009/04/09
Public Function WAT_CP_Data()
    Dim aWafer() As String
    Dim aSite() As String
    Dim i As Long, j As Long
    Dim setupSheet As Worksheet
    Dim sheetName As String
   
    Call GetWaferList(dSheet, aWafer)
    Call GetSiteList(dSheet, aSite)
    If Not IsExistSheet("setupWaferMap") Then Exit Function
    Set setupSheet = Worksheets("setupWaferMap")
    
    'del old sheet
    For i = Worksheets.Count To 1 Step -1
        If Left(Worksheets(i).Name, 11) = "WAT_CP_Data" Then DelSheet (Worksheets(i).Name)
    Next i
    
    If (UBound(aWafer) + 1) * WorksheetFunction.countA(setupSheet.Range("A:A")) > 254 Then
        sheetName = "WAT_CP_Data_1"
    Else
        sheetName = "WAT_CP_Data"
    End If
    
    For i = 1 To setupSheet.UsedRange.Rows.Count
        If setupSheet.Cells(i, 1) = "" Then Exit For
        'Debug.Print setupSheet.Cells(i, 1)
        sheetName = WAT_CP_Data_sub(sheetName, aWafer, aSite, setupSheet.Cells(i, 1))
    Next i
    
    Debug.Print "WAT_CP Finished!"
End Function

'For 蘇建中
Public Function WAT_CP_Data_sub(mSheet As String, aWafer() As String, aSite() As String, nowPara As String)
    Dim nowSheet As Worksheet
    Dim tmp As String
    Dim i As Long, j As Long
    Dim iCol As Long
    
    If Not IsExistSheet(mSheet) Then AddSheet (mSheet)
    If UBound(aWafer) + 1 + Worksheets(mSheet).UsedRange.Columns.Count <= 256 Then
        Set nowSheet = AddSheet(mSheet, False)
    Else
        tmp = getCOL(mSheet, "_", 4)
        Set nowSheet = AddSheet("WAT_CP_Data" & CStr(CInt(tmp) + 1), False)
    End If
    
    If nowSheet.UsedRange.Rows.Count < 2 Then
        nowSheet.Cells(2, 1) = "Sequence"
        nowSheet.Cells(1, 2) = "Parameter"
        nowSheet.Cells(2, 2) = "Coordinate"
        For i = 0 To UBound(aSite)
            nowSheet.Cells(i + 3, 1) = i + 1
            nowSheet.Cells(i + 3, 2) = "(" & getCOL(aSite(i), "(", 2)
        Next i
    End If
    
    iCol = nowSheet.UsedRange.Columns.Count + 1
    For i = 0 To UBound(aWafer)
        nowSheet.Cells(1, iCol + i) = nowPara
        nowSheet.Cells(2, iCol + i) = aWafer(i)
        For j = 0 To UBound(aSite)
            nowSheet.Cells(j + 3, iCol + i) = getValueByPara(aWafer(i), nowPara, j + 1)
        Next j
    Next i
    
    WAT_CP_Data_sub = nowSheet.Name
End Function

Public Sub Manual_RotateNotch()
    Dim rType As String
    Dim iRow As Long, iCol As Long
    Dim nowSheet As Worksheet
    Dim tmp As String, x As String, y As String, siteStr As String
    Dim tmpStr As String
    
    rType = InputBox("1-逆時針旋轉90度", "Input Rotate", "1")
    If rType = "" Then Exit Sub
    
    Set nowSheet = Worksheets("Data")
    For iRow = 1 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(iRow, 1) = "No./DataType" Then Exit For
    Next iRow
    If iRow >= nowSheet.UsedRange.Rows.Count Then MsgBox "can't find wafer information": Exit Sub
    
    For iCol = 4 To nowSheet.UsedRange.Columns.Count
        If InStr(nowSheet.Cells(iRow, iCol), "(") <= 0 Then Exit For
        tmpStr = nowSheet.Cells(iRow, iCol)
        siteStr = getCOL(tmpStr, "(", 1)
        x = getCOL(getCOL(tmpStr, "(", 2), ",", 1)
        y = getCOL(getCOL(tmpStr, ")", 1), ",", 2)
        Select Case rType
            Case "1": tmp = x: x = y * -1: y = tmp
        End Select
        tmpStr = siteStr & "(" & x & "," & y & ")"
        'Debug.Print nowSheet.Cells(iRow, iCol), TmpStr
        nowSheet.Cells(iRow, iCol) = tmpStr
    Next iCol
    MsgBox "Finished"
End Sub

'--------------------------------------------------------
Dim ynSeries As Boolean

Type SAInfo
   Parameter As String
   MainType As String
   SubType As String
   width As Single
   Height As Single
   SAMIN As Single
   SAREF As Single
   SAMINvalue As Single
   SAREFvalue As Single
End Type
Option Explicit

Public Sub Model_AllMacro()
   Model_ChangeSeriesName
   Model_ChangeYAxisFormat
   Model_ModifyBoxTrend
End Sub

Public Function Manual_Vincent_WID()
   Dim nowSheet As Worksheet
   Dim rawSheet As Worksheet
   Dim widSheet As Worksheet
   Dim tmp As String
   Dim iRow As Long
   Dim i As Long, j As Long
   Dim waferNum As Integer
   Dim siteNum As Integer
   Dim waferList() As String
   Dim iWafer As Integer
   Dim rawRange As Range
   Dim medianA() As Single
   Dim sigmaA() As Single
   Dim widA() As Single
   Dim wiwA() As Single
   Dim n As Integer
   Dim maxWafer As Integer
   
   Set nowSheet = ActiveSheet
   If InStr(nowSheet.Name, "_Summary") <= 0 Then Exit Function
   tmp = getCOL(nowSheet.Name, "_Summary", 1)
   Set widSheet = AddSheet(tmp & "_WID")
   
   Call GetWaferList("Data", waferList)
   waferNum = UBound(waferList) + 1
   siteNum = getSiteNum("Data")
   ReDim medianA(siteNum)
   ReDim sigmaA(siteNum)
   ReDim widA(1 To siteNum)
   'Debug.Print "Num:", WaferNum, WaferList(0), SiteNum
   
   maxWafer = 245 \ siteNum
   iRow = 2
   For i = 3 To nowSheet.UsedRange.Rows.Count
      j = i + nowSheet.Cells(i, 1).MergeArea.Rows.Count - 1
      widSheet.Cells(iRow, 1) = nowSheet.Cells(i, 1)
      widSheet.Range(widSheet.Cells(iRow, 1), widSheet.Cells(iRow + 4, 1)).Merge
      widSheet.Range(widSheet.Cells(iRow, 1), widSheet.Cells(iRow + 4, 1)).VerticalAlignment = xlCenter
      widSheet.Cells(iRow, 2) = "WIW WID(Median)"
      widSheet.Cells(iRow + 1, 2) = "WIW WID(sigma)"
      widSheet.Cells(iRow + 2, 2) = "WIW WID(U%)"
      widSheet.Cells(iRow + 3, 2) = "WIW (U%)"
      widSheet.Cells(iRow + 4, 2) = "WID (U%)"
      For iWafer = 1 To waferNum
         If iWafer <= maxWafer Then
            Set rawSheet = Worksheets(tmp & "_Raw")
         Else
            Set rawSheet = Worksheets(tmp & "_Raw_" & CStr((iWafer - 1) \ maxWafer))
         End If
         
         Set rawRange = rawSheet.Range(N2L(7 + ((iWafer - 1) Mod maxWafer) * siteNum) & CStr(i) & ":" & N2L(7 + ((iWafer - 1) Mod maxWafer + 1) * siteNum - 1) & CStr(j))
         'Debug.Print rawRange.Address
         For n = 1 To siteNum
            medianA(n) = WorksheetFunction.Median(rawRange.Columns(n))
            sigmaA(n) = WorksheetFunction.StDev(rawRange.Columns(n))
            If medianA(n) = 0 Then medianA(n) = medianA(n) + 1E-23
            widA(n) = 3 * sigmaA(n) / medianA(n)
         Next n
         ReDim wiwA(1 To rawRange.Rows.Count)
         For n = 1 To rawRange.Rows.Count
            wiwA(n) = 3 * WorksheetFunction.StDev(rawRange.Rows(n)) / WorksheetFunction.Median(rawRange.Rows(n))
         Next n
         widSheet.Cells(1, 2 + iWafer) = "#" & waferList(iWafer - 1)
         widSheet.Cells(iRow, 2 + iWafer) = WorksheetFunction.Median(rawRange)
         widSheet.Cells(iRow + 1, 2 + iWafer) = WorksheetFunction.StDev(rawRange)
         widSheet.Cells(iRow + 2, 2 + iWafer).FormulaLocal = "=3*" & N2L(2 + iWafer) & CStr(iRow + 1) & "/" & N2L(2 + iWafer) & CStr(iRow)
         widSheet.Cells(iRow + 3, 2 + iWafer) = WorksheetFunction.Average(wiwA)
         widSheet.Cells(iRow + 4, 2 + iWafer) = WorksheetFunction.Average(widA)
         
         widSheet.Cells(iRow, 2 + iWafer).NumberFormat = "0.00"
         widSheet.Cells(iRow + 1, 2 + iWafer).NumberFormat = "0.00"
         widSheet.Cells(iRow + 2, 2 + iWafer).NumberFormat = "0.00%"
         widSheet.Cells(iRow + 3, 2 + iWafer).NumberFormat = "0.00%"
         widSheet.Cells(iRow + 4, 2 + iWafer).NumberFormat = "0.00%"
         
      Next iWafer
      i = j: iRow = iRow + 5
   Next i
   widSheet.Columns.AutoFit
   'nowSheet.Activate
End Function


Public Function Manual_Brief()
   Dim sheetName As String
   Dim nowSheet As Worksheet
   Dim fieldArray
   Dim i As Long, j As Long, n As Long
   Dim iSheet As Integer
   Dim tempA() As Long
   Dim exStr As String
   
   On Error Resume Next
   
   'sheetName = "Brief_Summary"
   fieldArray = Array("Median", "Average", "Sigma", "Yield")
   'If IsKey(exStr, "Max") Then fieldArray = Array("Median", "Average", "Sigma", "Yield", "Max", "Min", "Sigma%")
   
   '處理Diff ...
   '------------------------------------------------------------
   For iSheet = 1 To Worksheets.Count
      If Right(Worksheets(iSheet).Name, 8) = "_Summary" Then
         Set nowSheet = Worksheets(iSheet)
         '處理Diff and Diff%
         ReDim tempA(0)
         For i = 1 To nowSheet.UsedRange.Columns.Count
            If nowSheet.Cells(2, i) = fieldArray(0) Or nowSheet.Cells(2, i) = fieldArray(1) Then
               tempA(UBound(tempA)) = i
               ReDim Preserve tempA(UBound(tempA) + 1)
            End If
         Next i
         If UBound(tempA) > 0 Then ReDim Preserve tempA(UBound(tempA) - 1)
         For i = 1 To nowSheet.UsedRange.Rows.Count
            If Left(nowSheet.Cells(i, 2), 6) = "Diff.%" Then
               For j = 0 To UBound(tempA)
                  nowSheet.Range(Num2Letter(tempA(j)) & CStr(i)).Formula = "=IF($E" & CStr(i - 1) & "="""",""""," & Num2Letter(tempA(j)) & CStr(i - 1) & "/$E" & CStr(i - 1) & "-1" & ")"
                  nowSheet.Range(Num2Letter(tempA(j)) & CStr(i)).NumberFormat = "0.00%"
               Next j
            ElseIf Left(nowSheet.Cells(i, 2), 5) = "Diff." Then
               For j = 0 To UBound(tempA)
                  nowSheet.Range(Num2Letter(tempA(j)) & CStr(i)).Formula = "=IF($E" & CStr(i - 1) & "="""",""""," & Num2Letter(tempA(j)) & CStr(i - 1) & "-$E" & CStr(i - 1) & ")"
               Next j
            ElseIf Left(nowSheet.Cells(i, 2), 7) = "Time of" Then
               For j = 0 To UBound(tempA)
                  nowSheet.Range(Num2Letter(tempA(j)) & CStr(i)).Formula = "=IF($E" & CStr(i - 1) & "="""",""""," & Num2Letter(tempA(j)) & CStr(i - 1) & "/$E" & CStr(i - 1) & ")"
                  nowSheet.Range(Num2Letter(tempA(j)) & CStr(i)).NumberFormatLocal = "0.000""x"""
               Next j
            End If
         Next i
         Set nowSheet = Nothing
      End If
      If Right(Worksheets(iSheet).Name, 4) = "_Raw" Then
         Set nowSheet = Worksheets(iSheet)
         '處理Diff and Diff%
         For i = 1 To nowSheet.UsedRange.Columns.Count
            If nowSheet.Cells(2, i) = "1" Then
               n = i
               Exit For
            End If
         Next i

         For i = 1 To nowSheet.UsedRange.Rows.Count
            If Left(nowSheet.Cells(i, 2), 6) = "Diff.%" Then
               For j = n To nowSheet.UsedRange.Columns.Count
                  nowSheet.Range(Num2Letter(j) & CStr(i)).Formula = "=IF($E" & CStr(i - 1) & "="""",""""," & Num2Letter(j) & CStr(i - 1) & "/$E" & CStr(i - 1) & "-1" & ")"
                  nowSheet.Range(Num2Letter(j) & CStr(i)).NumberFormat = "0.00%"
               Next j
            ElseIf Left(nowSheet.Cells(i, 2), 5) = "Diff." Then
               For j = n To nowSheet.UsedRange.Columns.Count
                  nowSheet.Range(Num2Letter(j) & CStr(i)).Formula = "=IF($E" & CStr(i - 1) & "="""",""""," & Num2Letter(j) & CStr(i - 1) & "-$E" & CStr(i - 1) & ")"
               Next j
            ElseIf Left(nowSheet.Cells(i, 2), 4) = "Time" Then
               For j = n To nowSheet.UsedRange.Columns.Count
                  nowSheet.Range(Num2Letter(j) & CStr(i)).Formula = "=IF($E" & CStr(i - 1) & "="""",""""," & Num2Letter(j) & CStr(i - 1) & "/$E" & CStr(i - 1) & ")"
                  nowSheet.Range(Num2Letter(j) & CStr(i)).NumberFormatLocal = "0.000""x"""
               Next j
            End If
         Next i
         Set nowSheet = Nothing
      End If
      
   Next iSheet
   
   
   'Brief => 調整欄位順序
   '----------------------------------------------------------------------
   If IsExistSheet("SPEC_List") Then
      For iSheet = 1 To Worksheets("SPEC_List").UsedRange.Columns.Count
         If Trim(Worksheets("SPEC_List").Cells(1, iSheet)) = "" Then Exit For
         'If Trim(LCase(getCOL(Worksheets("SPEC_List").Cells(1, iSheet), ":", 2))) = "brief" Then
         If IsKey(getCOL(Worksheets("SPEC_List").Cells(1, iSheet), ":", 2), "brief") Then
            sheetName = getCOL(Worksheets("SPEC_List").Cells(1, iSheet), ":", 1) & "_Summary"
            If IsExistSheet(sheetName) Then
                fieldArray = Array("Median", "Average", "Sigma", "Yield")
                exStr = getCOL(Worksheets("SPEC_List").Cells(1, iSheet), ":", 2)
                If IsKey(exStr, "Max") Then fieldArray = Array("Median", "Average", "Sigma", "Yield", "Max", "Min", "Sigma%")
                If IsKey(exStr, "Diff") Then
                    ReDim Preserve fieldArray(UBound(fieldArray) + 1)
                    fieldArray(UBound(fieldArray)) = "Diff"
                    'fieldArray = Array("Median", "Average", "Sigma", "Yield", "Diff")
                End If
                'Else
                '  fieldArray = Array("Median", "Average", "Sigma", "Yield")
                'End If
               Set nowSheet = Worksheets(sheetName)
               '調換攔位順序
               For n = 0 To UBound(fieldArray)
                  ReDim tempA(0)
                  For i = 1 To nowSheet.UsedRange.Columns.Count
                     If nowSheet.Cells(2, i) = fieldArray(n) Then
                        tempA(UBound(tempA)) = i
                        ReDim Preserve tempA(UBound(tempA) + 1)
                     End If
                  Next i
                  If UBound(tempA) > 0 Then ReDim Preserve tempA(UBound(tempA) - 1)
                  nowSheet.Columns(tempA(0)).Borders(xlEdgeRight).LineStyle = xlNone
                  For i = 1 To UBound(tempA)
                     If tempA(i) > tempA(0) + i Then
                        nowSheet.Columns(tempA(i)).Cut
                        nowSheet.Columns(tempA(0) + i).Insert Shift:=xlToRight
                     End If
                     nowSheet.Columns(tempA(0) + i).Borders(xlEdgeRight).LineStyle = xlNone
                  Next i
                  With nowSheet.Columns(tempA(0) + i - 1).Borders(xlEdgeRight)
                     .LineStyle = xlContinuous
                     .Weight = xlMedium
                     .ColorIndex = xlAutomatic
                  End With
               Next n
               nowSheet.UsedRange.FormatConditions.Delete
               '加格式化條件
               For j = 6 To nowSheet.Columns.Count
                  If nowSheet.Cells(2, j) = "Median" Then
                    For i = 3 To nowSheet.Rows.Count
                        If nowSheet.Cells(i, 4) <> "" Then nowSheet.Cells(i, j).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & "D" & "" & CStr(3)).Font.ColorIndex = 4
                        If nowSheet.Cells(i, 6) <> "" Then nowSheet.Cells(i, j).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & "F" & "" & CStr(3)).Font.ColorIndex = 3
                    Next i
                  End If
               Next j
               Set nowSheet = Nothing
            End If
         End If
      Next iSheet
   End If
   
   
   
End Function

Sub Model_ChangeSeriesName()
   Dim i As Integer, j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim nowSeries As Series
   Dim xLabel As String
   Dim CateName As String
   Dim Unit As String
   
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "SCATTER") Then
         Set nowSheet = Worksheets(i)
         For j = 1 To nowSheet.ChartObjects.Count
            Set nowChart = nowSheet.ChartObjects(j).Chart
            If Not nowChart.HasAxis(xlCategory) Then Exit For
            xLabel = nowChart.Axes(xlCategory).AxisTitle.Text
            CateName = UCase(getCOL(xLabel, "(", 1))
            Unit = getCOL(getCOL(xLabel, "(", 2), ")", 1)
            If CateName <> "SA" And CateName <> "W" And CateName <> "L" Then Exit For
            If Unit = "" Then Exit For
            For m = 1 To nowChart.SeriesCollection.Count
               Set nowSeries = nowChart.SeriesCollection(m)
               If Application.WorksheetFunction.Max(nowSeries.XValues) = Application.WorksheetFunction.Min(nowSeries.XValues) Then
                  nowSeries.Name = CateName & "=" & CStr(nowSeries.XValues(1)) & "" & Unit
               End If
               'Debug.Print nowSeries.Name
            Next m
         Next j
      End If
   Next i
End Sub

Sub Model_ChangeYAxisFormat()
   Dim i As Integer, j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim nowAxis As Axis
   Dim yLabel As String
   Dim CateName As String
   Dim Unit As String
   
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "SCATTER") Then
         Set nowSheet = Worksheets(i)
         For j = 1 To nowSheet.ChartObjects.Count
            Set nowChart = nowSheet.ChartObjects(j).Chart
            If Not nowChart.HasAxis(xlValue) Then Exit For
            yLabel = nowChart.Axes(xlValue).AxisTitle.Text
            CateName = UCase(Left(yLabel, 3))
            Select Case CateName
               Case "IDS", "IDL", "IOF"
                  Set nowAxis = nowChart.Axes(xlValue)
                  nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
            End Select
         Next j
      End If
   Next i
End Sub

Sub Model_AddMedianTable()
   Dim i As Integer, j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim TableRange As Range
   Dim MedianRange As Range
   Dim xLabel As String
   Dim yLabel As String
   Dim nowSA As SAInfo
   
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "SCATTER") Then
         Set nowSheet = Worksheets(i)
         
         For j = 1 To nowSheet.UsedRange.Columns.Count
            If UCase(nowSheet.Cells(1, j)) = "MEDIAN" Then Exit For
         Next j
         If j < nowSheet.UsedRange.Columns.Count Then
            For m = 3 To nowSheet.UsedRange.Rows.Count
               If nowSheet.Cells(m, j) = "" Then Exit For
            Next m
            Set MedianRange = nowSheet.Range(Num2Letter(j) & CStr(1) & ":" & Num2Letter(j + 1) & CStr(m - 1))
            Set TableRange = nowSheet.Range(Num2Letter(nowSheet.UsedRange.Columns.Count + 1) & "1")
            Call getSAInfo(getFirstParameter(nowSheet.Name), MedianRange, nowSA)
            TableRange.Cells(1, 1) = "Med"
            TableRange.Range("A1").HorizontalAlignment = xlRight
            TableRange.Cells(2, 1) = "SA(um)"
            With TableRange.Range("A1:A2")
               nowSheet.Shapes.AddLine(.Left, .Top, .Left + .width, .Top + .Height).Select
               .Borders.LineStyle = xlContinuous
               .Borders(xlInsideHorizontal).LineStyle = xlNone
               .Interior.ColorIndex = 2
            End With
            With TableRange.Range("B1:C2")
               .Borders.LineStyle = xlContinuous
               .Borders(xlInsideHorizontal).LineStyle = xlNone
               .Interior.ColorIndex = 2
               .Cells(2, 1) = nowSA.MainType
               .Cells(2, 2) = nowSA.SubType
               .Columns.AutoFit
            End With

            For m = 3 To MedianRange.Rows.Count
               If MedianRange.Cells(m, 1) = "" Then Exit For
               With TableRange.Range("A" & CStr(m) & ":C" & CStr(m))
                  .Cells(1, 1) = MedianRange.Cells(m, 1)
                  .Cells(1, 2) = MedianRange.Cells(m, 2)
                  Select Case nowSA.MainType
                     Case "VTSN", "VTSP":  .Cells(1, 3) = (.Cells(1, 2) - Val(nowSA.SAREFvalue)) * 1000
                     Case "IDSN", "IDSP":  .Cells(1, 3) = (.Cells(1, 2) - Val(nowSA.SAREFvalue)) / Val(nowSA.SAREFvalue) * 100
                  End Select
                  .Borders.LineStyle = xlContinuous
                  .HorizontalAlignment = xlHAlignCenter
               End With
            Next m
         End If
      End If
   Next i
End Sub

Public Function getFirstParameter(sheetName As String)
   Dim nowSheet As Worksheet
   Dim i As Long
   
   Set nowSheet = Worksheets(sheetName)
   For i = 1 To nowSheet.UsedRange.Rows.Count
      If UCase(nowSheet.Cells(i, 1)) = "Y" Then getFirstParameter = nowSheet.Cells(i + 1, 1)
   Next i
End Function

Public Function getSAInfo(nowParameter As String, ByRef MedianRange As Range, ByRef nowSA As SAInfo)
   Dim i As Long
   
   nowSA.Parameter = nowParameter
   nowSA.MainType = UCase(getCOL(nowSA.Parameter, "_", 1))
   Select Case nowSA.MainType
      Case "VTSN":   nowSA.SubType = "VTSN_D(mV)"
      Case "VTSP":   nowSA.SubType = "VTSP_D(mV)"
      Case "IDSN":   nowSA.SubType = "IDSN(%)"
      Case "IDSP":   nowSA.SubType = "IDSP(%)"
   End Select
   nowSA.width = CSng(Replace(LCase(getCOL(nowSA.Parameter, "_", 2)), "p", "."))
   nowSA.Height = CSng(Replace(LCase(getCOL(nowSA.Parameter, "_", 3)), "p", "."))
   If nowSA.width <= 0.2 Then
      nowSA.SAMIN = 0.36
      nowSA.SAREF = 1.89
   Else
      nowSA.SAMIN = 0.32
      nowSA.SAREF = 1.76
   End If
   For i = 1 To MedianRange.Rows.Count
      If MedianRange.Cells(i, 1) = nowSA.SAMIN Then nowSA.SAMINvalue = MedianRange.Cells(i, 2)
      If MedianRange.Cells(i, 1) = nowSA.SAREF Then nowSA.SAREFvalue = MedianRange.Cells(i, 2)
   Next i
   
   'Debug.Print nowSA.Parameter, nowSA.Width, nowSA.Height, nowSA.SAMIN, nowSA.SAREF, nowSA.SAMINvalue, nowSA.SAREFvalue
   'Stop
End Function

Sub Model_ModifyBoxTrend()
   Dim i As Integer, j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim nowAxis As Axis
   Dim nowSeries As Series
   Dim nowPoint As Point
   Dim yLabel As String
   Dim CateName As String
   Dim Unit As String
   Dim tempA
   Dim waferNum As Integer
   Dim tmpStr As String
   Dim ArrayStr As String
   
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "BOXTREND") Then
         Set nowSheet = Worksheets(i)
         For j = 1 To nowSheet.ChartObjects.Count
            Set nowChart = nowSheet.ChartObjects(j).Chart
            Set nowAxis = nowChart.Axes(xlValue)
            nowAxis.MajorGridlines.Border.ColorIndex = 2
            Set nowAxis = nowChart.Axes(xlCategory)
            nowAxis.MajorGridlines.Border.ColorIndex = 2
            nowAxis.TickLabels.Font.Size = 12
            If InStr(nowAxis.AxisTitle.Text, "SA") <= 0 Then Exit For
            tempA = nowChart.SeriesCollection(1).XValues
            For m = 1 To UBound(tempA)
               If IsEmpty(tempA(m)) Then Exit For
            Next m
            waferNum = m - 1
            nowAxis.TickLabelSpacing = waferNum + 1
            nowAxis.TickMarkSpacing = waferNum + 1
            For m = 1 To UBound(tempA)
               If Not IsEmpty(tempA(m)) Then
                  If nowChart.SeriesCollection(5).Points(m).HasDataLabel Then
                     tmpStr = nowChart.SeriesCollection(5).Points(m).DataLabel.Text
                     tmpStr = getCOL(tmpStr, "_", 4)
                     tmpStr = Replace(tmpStr, "p", ".")
                     If Left(tmpStr, 1) = "." Then tmpStr = "0" & tmpStr
                  End If
               End If
            Next m
            Set nowSeries = nowChart.SeriesCollection(4)
            ArrayStr = ""
            For m = 1 To UBound(tempA)
               If Not IsEmpty(tempA(m)) Then
                  Set nowPoint = nowSeries.Points(m)
                  nowPoint.HasDataLabel = True
                  nowPoint.DataLabel.Text = CStr(tempA(m))
                  nowPoint.DataLabel.Font.Size = 8
                  nowPoint.DataLabel.Position = xlLabelPositionAbove
                  If nowChart.SeriesCollection(5).Points(m).HasDataLabel Then
                     tmpStr = nowChart.SeriesCollection(5).Points(m).DataLabel.Text
                     tmpStr = getCOL(tmpStr, "_", 4)
                     tmpStr = Replace(tmpStr, "p", ".")
                     If Left(tmpStr, 1) = "." Then tmpStr = "0" & tmpStr
                     ArrayStr = ArrayStr & "," & tmpStr
                  Else
                     ArrayStr = ArrayStr & "," & "0"
                  End If
                  nowChart.SeriesCollection(5).Points(m).HasDataLabel = False
               Else
                  ArrayStr = ArrayStr & "," & "0"
               End If
            Next m
            If Len(ArrayStr) > 1 Then ArrayStr = Mid(ArrayStr, 2)
            nowChart.SeriesCollection(1).XValues = Split(ArrayStr, ",")
            nowChart.SeriesCollection(5).Border.LineStyle = xlLineStyleNone
         Next j
      End If
   Next i
End Sub

Sub Manual_FilterBySSFF()
   Dim i As Integer, j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim ssRange As Range
   Dim ffRange As Range
   Dim ssBig As Boolean
   
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "SCATTER") Then
         Set nowSheet = Worksheets(i)
         Set ssRange = getRangeBySeriesName(nowSheet, "SS")
         If ssRange.Cells.Count > 1 Then
            Set ffRange = getRangeBySeriesName(nowSheet, "FF")
            If getValueByIndex(ssRange, ssRange.Cells(1, 1), 1, 2) > getValueByIndex(ffRange, ssRange.Cells(1, 1), 1, 2) Then
               ssBig = True
            Else
               ssBig = False
            End If
            For j = 3 To nowSheet.UsedRange.Columns.Count Step 2
               If nowSheet.Cells(2, j) = "" Then Exit For
               For m = 3 To nowSheet.UsedRange.Rows.Count
                  If getValueByIndex(ssRange, nowSheet.Cells(m, j), 1, 2) Then
                     If ssBig Then
                        If nowSheet.Cells(m, j + 1) > getValueByIndex(ssRange, nowSheet.Cells(m, j), 1, 2) Then
                           nowSheet.Cells(m, j + 1) = ""
                        End If
                        If nowSheet.Cells(m, j + 1) < getValueByIndex(ffRange, nowSheet.Cells(m, j), 1, 2) Then
                           nowSheet.Cells(m, j + 1) = ""
                        End If
                     Else
                        If nowSheet.Cells(m, j + 1) < getValueByIndex(ssRange, nowSheet.Cells(m, j), 1, 2) Then
                           nowSheet.Cells(m, j + 1) = ""
                        End If
                        If nowSheet.Cells(m, j + 1) > getValueByIndex(ffRange, nowSheet.Cells(m, j), 1, 2) Then
                           nowSheet.Cells(m, j + 1) = ""
                        End If
                     End If
                  End If
               Next m
            Next j
            'Debug.Print getValueByIndex(ssRange, 10, 1, 2)
            Call FitAxisScale(nowSheet)
         End If
      End If
   Next i
End Sub


Function getRangeBySeriesName(nowSheet As Worksheet, SeriesName As String) As Range
   Dim i As Long, j As Long
   
   Set getRangeBySeriesName = nowSheet.Range("A1")
   For i = 1 To nowSheet.UsedRange.Columns.Count
      If nowSheet.Cells(1, i) = SeriesName Then Exit For
   Next i
   If i > nowSheet.UsedRange.Columns.Count Then Exit Function
   For j = 3 To nowSheet.UsedRange.Rows.Count
      If nowSheet.Cells(j, i) = "" Then Exit For
   Next j
   If j = 3 Then Exit Function
   Set getRangeBySeriesName = nowSheet.Range(Num2Letter(i) & "3" & ":" & Num2Letter(i + 1) & CStr(j))
End Function

Function getValueByIndex(nowRange As Range, nowIndex, indexCol As Integer, valueCol As Integer)
   Dim i As Long
   
   getValueByIndex = False
   For i = 1 To nowRange.Rows.Count
      If nowRange.Cells(i, indexCol) = nowIndex Then
         getValueByIndex = nowRange.Cells(i, valueCol)
         Exit For
      End If
   Next i
End Function

Private Sub FitAxisScale(ByRef nowSheet As Worksheet)
   Dim yMin, yMax, xMin, xMax
   Dim vChartInfo As chartInfo
   Dim nowChart As Chart
   Dim nowAxis As Axis
   Dim ynLog As Boolean

   vChartInfo = getChartInfo(nowSheet.Range("A1").CurrentRegion)
   Set nowChart = nowSheet.ChartObjects(1).Chart
   Call getScaterMaxMin(nowChart, xMax, xMin, yMax, yMin)
   Set nowAxis = nowChart.Axes(xlValue)
   If IsKey(vChartInfo.YScale, "Log") Then
      ynLog = True
   Else
      ynLog = False
   End If
   Call AxisScaleFit(nowAxis, yMax, yMin, "", "", ynLog)
   'Public Sub AxisScaleFit(ByRef rAxis As Axis, ByVal vvarMax, ByVal vvarMin, ByVal cMax, ByVal cMin, varLog As Boolean)
End Sub

Private Sub Manual_FitLegend()
   Dim i As Integer, j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim nowLegend As Legend
   Dim nowSeries As Series
   Dim nowShape As Shape
   Dim nowAxis As Axis
  
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "SCATTER") Or InStr(UCase(Worksheets(i).Name), "BOXTREND") Then
         Set nowSheet = Worksheets(i)
         'nowSheet.Activate
         'nowSheet.Range("A1").Select
         For j = 1 To nowSheet.ChartObjects.Count
            'nowSheet.ChartObjects(j).Activate
            Set nowChart = nowSheet.ChartObjects(j).Chart
            'nowChart.ChartArea.Select
            '-----------------
            'Fit Chart Attribs
            '-----------------
            Set nowShape = nowSheet.Shapes(1)
            'Debug.Print nowShape.Name
            With nowShape
               .Top = 30
               .Left = 30
               .width = 450
               .Height = 300 + (j - 1) * 400
            End With
            DoEvents
            'nowChart.ChartArea.Width = 500
            '----------
            'Fit Legend
            '----------
            If nowChart.HasLegend Then
               Set nowLegend = nowChart.Legend
               'nowLegend.Select
               nowLegend.AutoScaleFont = False
               nowLegend.Font.Size = 9
               'nowLegend.Height = nowChart.ChartArea.Height - 40
               'Selection.Width = 50
               'nowLegend.Width = 500
               For m = 1 To nowChart.SeriesCollection.Count
                  Set nowSeries = nowChart.SeriesCollection(m)
                  'Debug.Print nowSeries.Name
                  On Error Resume Next
                  Select Case Len(nowSeries.Name)
                     Case Is > 25
                        nowLegend.LegendEntries(m).Font.Size = 8
                     Case Else
                        nowLegend.LegendEntries(m).Font.Size = 9
                  End Select
                  On Error GoTo 0
               Next m
               'nowLegend.Width = 150
               
               'nowLegend.Left = nowChart.ChartArea.Width - nowLegend.Width - 6 - 3
               'nowLegend.Top = (nowChart.ChartArea.Height - nowLegend.Height) / 2 + 10
               'Debug.Print "W:", nowShape.Width, nowLegend.Width
               'Debug.Print "H:", nowShape.Height, nowLegend.Height
               DoEvents
               nowLegend.Left = nowShape.width - nowLegend.width - 6 - 10
               nowLegend.Top = (nowShape.Height - nowLegend.Height) / 2 '+ 10
               'DoEvents
               'nowLegend.Width = nowShape.Width
               'DoEvents
               'Debug.Print "Legend:", nowLegend.Left, nowLegend.Top
               'Debug.Print nowChart.ChartArea.Width & ":" & nowLegend.Left & ":" & nowLegend.Width
            End If
            '----------
            'Fit Axis
            '----------
            If True Then 'nowChart.HasAxis(lCategory, xlPrimary) Then
               Set nowAxis = nowChart.Axes(xlCategory, xlPrimary)
               nowAxis.TickLabels.AutoScaleFont = True
               nowAxis.TickLabels.Font.Name = "Arial"
               nowAxis.TickLabels.Font.Size = 10
               If nowAxis.HasTitle Then nowAxis.AxisTitle.Font.Size = 12
               If InStr(UCase(nowSheet.Name), "BOXTREND") Then
                  nowAxis.TickLabels.Font.Size = 6
                  nowAxis.TickLabels.Orientation = xlTickLabelOrientationUpward
               End If
            End If
            If True Then 'nowChart.HasAxis(xlValue, xlPrimary) Then
               Set nowAxis = nowChart.Axes(xlValue, xlPrimary)
               nowAxis.TickLabels.AutoScaleFont = True
               nowAxis.TickLabels.Font.Name = "Arial"
               nowAxis.TickLabels.Font.Size = 10
               If nowAxis.HasTitle Then nowAxis.AxisTitle.Font.Size = 12
            End If
         Next j
      End If
   Next i
End Sub

Public Function new_FitChart()
    Dim i As Integer, j As Integer, n As Integer
    Dim nowChart As Chart
    Dim nowSheet As Worksheet
    Dim nowLegend As Legend
    Dim nowShape As Shape
    Dim vChartInfo As chartInfo
    Dim tmpStr As String
    Dim nowAxis As Axis

    For i = 1 To Worksheets.Count
        If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Or InStr(UCase(Worksheets(i).Name), "BOXTREND") > 0 Then
            Set nowSheet = Worksheets(i)
            For j = 1 To nowSheet.ChartObjects.Count
                Set nowChart = nowSheet.ChartObjects(j).Chart
                'nowChart.ClearToMatchStyle
                'nowChart.ChartStyle = 343
                '-----------------
                'Fit Chart Attribs
                '-----------------
                Set nowShape = nowSheet.Shapes(j)
                DoEvents
                With nowShape
                   .Top = 30 + (j - 1) * 400
                   .Left = 30
                   .width = 450
                   .Height = 300
                End With
                DoEvents
                '----------
                'Fit Legend
                '----------
                On Error Resume Next
                If nowChart.HasLegend Then
                   Set nowLegend = nowChart.Legend
                   nowLegend.Border.Color = 0
                   nowLegend.Border.ColorIndex = 1
                   nowLegend.Border.LineStyle = 1
                   nowLegend.Border.Weight = 2
                   
                   nowLegend.Fill.BackColor.SchemeColor = 2
                   nowLegend.Fill.ForeColor.SchemeColor = 19
                   nowLegend.Fill.Visible = msoTrue
                   
                   nowLegend.AutoScaleFont = False
                   nowLegend.Font.Size = 10
                   DoEvents
                   
                   nowLegend.Left = 60
                   nowLegend.Top = 30
                   nowLegend.Height = 30
                   nowLegend.width = 360
                   nowLegend.Interior.ColorIndex = 19
                End If
                On Error Resume Next
                vChartInfo = getChartInfo(nowSheet.Range("A1").CurrentRegion)
                If Trim(UCase(vChartInfo.XParameter(1))) = "SITE" Then
                   nowChart.Axes(xlCategory).TickLabels.Font.Size = 10
                   nowChart.Axes(xlCategory).MajorUnit = 1
                End If
                '----------
                'Add grid line
                If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Then
                    nowChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
                    nowChart.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
                    nowChart.SetElement (msoElementPrimaryCategoryGridLinesMinorMajor)
                End If
                
                
                '----------
                'Plot Area
                '----------
                nowChart.PlotArea.Left = 31.658
                nowChart.PlotArea.Top = 69
                nowChart.PlotArea.Format.Fill.BackColor.SchemeColor = 9
                
                nowChart.PlotArea.Border.ColorIndex = 1
                nowChart.PlotArea.LineStyle = 1
                nowChart.PlotArea.Weight = 2
                
                nowChart.PlotArea.Height = 203
                If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Then
                    nowChart.PlotArea.InsideHeight = 180
                    nowChart.PlotArea.InsideTop = 72
                ElseIf InStr(UCase(Worksheets(i).Name), "BOXTREND") > 0 Then
                    nowChart.PlotArea.InsideHeight = 222
                    nowChart.PlotArea.InsideTop = 30
                End If
                
                nowChart.PlotArea.InsideLeft = 60
                nowChart.PlotArea.InsideWidth = 360
                nowChart.PlotArea.Border.Color = 0
                nowChart.PlotArea.Border.ColorIndex = 1
                nowChart.PlotArea.Border.LineStyle = 1
                nowChart.PlotArea.Border.Weight = 2
                
                nowChart.PlotArea.Format.Line.Visible = msoTrue
                nowChart.PlotArea.Format.Line.Weight = 1.5

                '----------
                'Chart Title
                '----------
                nowChart.ChartTitle.Font.ColorIndex = 56
                nowChart.ChartTitle.Font.Size = 16
                nowChart.ChartTitle.Left = nowChart.PlotArea.InsideLeft + 180 - nowChart.ChartTitle.width / 2
                
                
                '----------
                'Axis Title
                '----------
                nowChart.Axes(xlValue).AxisTitle.Font.Size = 14
                nowChart.Axes(xlValue).AxisTitle.Font.ColorIndex = 56
                nowChart.Axes(xlValue).AxisTitle.Font.Bold = False
                nowChart.Axes(xlValue).AxisTitle.Top = 114
                nowChart.Axes(xlValue).AxisTitle.Left = 12
                nowChart.Axes(xlCategory).AxisTitle.Font.Size = 14
                nowChart.Axes(xlCategory).AxisTitle.Font.ColorIndex = 56
                nowChart.Axes(xlCategory).AxisTitle.Font.Bold = False
                nowChart.Axes(xlCategory).AxisTitle.Top = 272
                nowChart.Axes(xlCategory).AxisTitle.Left = 216
                
                '----------
                'Axis
                '----------
                nowChart.Axes(xlValue).Left = 32
                nowChart.Axes(xlValue).Top = 77
                nowChart.Axes(xlValue).TickLabels.Font.Size = 10
                nowChart.Axes(xlCategory).Left = 63
                nowChart.Axes(xlCategory).Top = 249
                nowChart.Axes(xlCategory).TickLabels.Font.Size = 10
            
                '----------
                'TrendLines
                '----------
                Dim nowLabel As DataLabel
                For n = 1 To nowChart.FullSeriesCollection.Count
                    If nowChart.FullSeriesCollection(n).Trendlines.Count > 0 Then
                        Set nowLabel = nowChart.FullSeriesCollection(n).Trendlines(1).DataLabel
                        nowLabel.Left = nowChart.PlotArea.InsideLeft + nowChart.PlotArea.InsideWidth - nowLabel.width
                        nowLabel.Top = nowChart.PlotArea.InsideTop + (n - 1) * nowLabel.Height
                    End If
                Next n

                
            
            If Trim(UCase(vChartInfo.XParameter(1))) = "PARA" Then
                nowChart.chartType = xlLineMarkers
                tmpStr = """" & vChartInfo.YParameter(1) & """"
                For n = 1 To vChartInfo.YParameter.Count
                    tmpStr = tmpStr & "," & """" & vChartInfo.YParameter(n) & """"
                Next n
                tmpStr = Replace(tmpStr, "=", "")
                'nowChart.SeriesCollection(1).XValues = "={""A"",""B""}"
                nowChart.SeriesCollection(1).XValues = "={" & tmpStr & "}"
                'nowChart.SeriesCollection(2).XValues = "={" & tmpStr & "}"
                nowChart.Axes(xlCategory).TickLabels.Font.Size = 8
                'nowChart.Axes(xlCategory).MajorUnit = 1
            End If
            
            
            On Error GoTo 0
         Next j
      End If
   Next i
End Function
Public Function FitSingleChart(nowSheet As Worksheet)
    
    Dim i As Integer, j As Integer, n As Integer
    Dim nowChart As Chart
    Dim nowShape As Shape
    Dim vChartInfo As chartInfo
    Dim tmpStr As String
    Dim nowLegend As Legend
    Dim nowAxis As Axis
    
    Set nowChart = nowSheet.ChartObjects(1).Chart
    Set nowShape = nowSheet.Shapes(1)
    
    '-----------------
    'Fit Chart Attribs
    '-----------------
    DoEvents
    With nowShape
       .Top = 30 + (j - 1) * 400
       .Left = 30
       .width = 450
       .Height = 300
    End With
    DoEvents
    '----------
    'Fit Legend
    '----------
    On Error Resume Next
    If nowChart.HasLegend Then
        Set nowLegend = nowChart.Legend
        nowLegend.Border.Color = 0
        nowLegend.Border.ColorIndex = 1
        nowLegend.Border.LineStyle = 1
        nowLegend.Border.Weight = 2
       
        nowLegend.Fill.BackColor.SchemeColor = 2
        nowLegend.Fill.ForeColor.SchemeColor = 19
        nowLegend.Fill.Visible = msoTrue
       
        nowLegend.AutoScaleFont = False
        nowLegend.Font.Size = 10
        DoEvents
       
        nowLegend.Left = 60
        nowLegend.Top = 30
        nowLegend.Height = 30
        nowLegend.width = 360
        nowLegend.Interior.ColorIndex = 19
    End If
    On Error Resume Next
    vChartInfo = getChartInfo(nowSheet.Range("A1").CurrentRegion)
    If Trim(UCase(vChartInfo.XParameter(1))) = "SITE" Then
       nowChart.Axes(xlCategory).TickLabels.Font.Size = 10
       nowChart.Axes(xlCategory).MajorUnit = 1
    End If
    '----------
    'Add grid line
    '----------
    If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Then
        nowChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
        nowChart.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
        nowChart.SetElement (msoElementPrimaryCategoryGridLinesMinorMajor)
    End If
    '----------
    'Plot Area
    '----------
    nowChart.PlotArea.Left = 31.658
    nowChart.PlotArea.Top = 69
    nowChart.PlotArea.Format.Fill.BackColor.SchemeColor = 9
    
    nowChart.PlotArea.Border.ColorIndex = 1
    nowChart.PlotArea.LineStyle = 1
    nowChart.PlotArea.Weight = 2
    
    nowChart.PlotArea.Height = 203
    If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Then
        nowChart.PlotArea.InsideHeight = 180
        nowChart.PlotArea.InsideTop = 72
    ElseIf InStr(UCase(Worksheets(i).Name), "BOXTREND") > 0 Then
        nowChart.PlotArea.InsideHeight = 222
        nowChart.PlotArea.InsideTop = 30
    End If
    
    nowChart.PlotArea.InsideLeft = 60
    nowChart.PlotArea.InsideWidth = 360
    nowChart.PlotArea.Border.Color = 0
    nowChart.PlotArea.Border.ColorIndex = 1
    nowChart.PlotArea.Border.LineStyle = 1
    nowChart.PlotArea.Border.Weight = 2
    
    nowChart.PlotArea.Format.Line.Visible = msoTrue
    nowChart.PlotArea.Format.Line.Weight = 1.5

    '----------
    'Chart Title
    '----------
    nowChart.ChartTitle.Font.ColorIndex = 56
    nowChart.ChartTitle.Font.Size = 16
    nowChart.ChartTitle.Left = nowChart.PlotArea.InsideLeft + 180 - nowChart.ChartTitle.width / 2
    
    '----------
    'Axis Title
    '----------
    nowChart.Axes(xlValue).AxisTitle.Font.Size = 14
    nowChart.Axes(xlValue).AxisTitle.Font.ColorIndex = 56
    nowChart.Axes(xlValue).AxisTitle.Font.Bold = False
    nowChart.Axes(xlValue).AxisTitle.Top = 114
    nowChart.Axes(xlValue).AxisTitle.Left = 12
    nowChart.Axes(xlCategory).AxisTitle.Font.Size = 14
    nowChart.Axes(xlCategory).AxisTitle.Font.ColorIndex = 56
    nowChart.Axes(xlCategory).AxisTitle.Font.Bold = False
    nowChart.Axes(xlCategory).AxisTitle.Top = 272
    nowChart.Axes(xlCategory).AxisTitle.Left = 216
    
    '----------
    'Axis
    '----------
    nowChart.Axes(xlValue).Left = 32
    nowChart.Axes(xlValue).Top = 77
    nowChart.Axes(xlValue).TickLabels.Font.Size = 10
    nowChart.Axes(xlCategory).Left = 63
    nowChart.Axes(xlCategory).Top = 249
    nowChart.Axes(xlCategory).TickLabels.Font.Size = 10

    '----------
    'TrendLines
    '----------
    Dim nowLabel As DataLabel
    For n = 1 To nowChart.FullSeriesCollection.Count
        If nowChart.FullSeriesCollection(n).Trendlines.Count > 0 Then
            Set nowLabel = nowChart.FullSeriesCollection(n).Trendlines(1).DataLabel
            nowLabel.Left = nowChart.PlotArea.InsideLeft + nowChart.PlotArea.InsideWidth - nowLabel.width
            nowLabel.Top = nowChart.PlotArea.InsideTop + (n - 1) * nowLabel.Height
        End If
    Next n

    If Trim(UCase(vChartInfo.XParameter(1))) = "PARA" Then
        nowChart.chartType = xlLineMarkers
        tmpStr = """" & vChartInfo.YParameter(1) & """"
        For n = 1 To vChartInfo.YParameter.Count
            tmpStr = tmpStr & "," & """" & vChartInfo.YParameter(n) & """"
        Next n
        tmpStr = Replace(tmpStr, "=", "")
        nowChart.SeriesCollection(1).XValues = "={" & tmpStr & "}"
        nowChart.Axes(xlCategory).TickLabels.Font.Size = 8
    End If
    
    On Error GoTo 0
   
End Function

Public Function old_FitChart()
   Dim i As Integer, j As Integer, n As Integer
   Dim nowChart As Chart
   Dim nowSheet As Worksheet
   Dim nowLegend As Legend
   Dim nowShape As Shape
   Dim vChartInfo As chartInfo
   Dim tmpStr As String
   Dim nowAxis As Axis
  
  
   For i = 1 To Worksheets.Count
      If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Or InStr(UCase(Worksheets(i).Name), "BOXTREND") > 0 Then
         Set nowSheet = Worksheets(i)
         For j = 1 To nowSheet.ChartObjects.Count
            Set nowChart = nowSheet.ChartObjects(j).Chart
            '-----------------
            'Fit Chart Attribs
            '-----------------
            Set nowShape = nowSheet.Shapes(j)
            'Debug.Print nowShape.Name
            DoEvents
            With nowShape
               .Top = 30 + (j - 1) * 400
               .Left = 30
               .width = 450
               .Height = 300
            End With
            DoEvents
            '----------
            'Fit Legend
            '----------
            On Error Resume Next
            If nowChart.HasLegend Then
               Set nowLegend = nowChart.Legend
               nowLegend.AutoScaleFont = False
               nowLegend.Font.Size = 10
               DoEvents
               nowLegend.Left = nowShape.width - nowLegend.width - 6 - 10
               nowLegend.Top = (nowShape.Height - nowLegend.Height) / 2 '+ 10
               nowLegend.Interior.ColorIndex = 2
            End If
            On Error Resume Next
            vChartInfo = getChartInfo(nowSheet.Range("A1").CurrentRegion)
            If Trim(UCase(vChartInfo.XParameter(1))) = "SITE" Then
               nowChart.Axes(xlCategory).TickLabels.Font.Size = 10
               nowChart.Axes(xlCategory).MajorUnit = 1
            End If
            
'Avoid crash
'            If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Then
'                Set nowAxis = nowChart.Axes(xlCategory)
'                'If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
'                'Avoid crash
'                'nowAxis.TickLabels.AutoScaleFont = False
'                'nowAxis.TickLabels.Font.Size = 12
'                'Set nowAxis = Nothing
'                Set nowAxis = nowChart.Axes(xlValue)
'                'If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
'                'Avoid crash
'                'nowAxis.TickLabels.AutoScaleFont = False
'                'nowAxis.TickLabels.Font.Size = 12
'            Else
'                Set nowAxis = nowChart.Axes(xlCategory)
'                'If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
'                'Avoid crash
'                'nowAxis.TickLabels.AutoScaleFont = False
'                'nowAxis.TickLabels.Font.Size = 10
'                'Set nowAxis = Nothing
'                Set nowAxis = nowChart.Axes(xlValue)
'                'If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
'                'Avoid crash
'                'nowAxis.TickLabels.AutoScaleFont = False
'                'nowAxis.TickLabels.Font.Size = 10
'            End If
            
            
            
            If Trim(UCase(vChartInfo.XParameter(1))) = "PARA" Then
                nowChart.chartType = xlLineMarkers
                tmpStr = """" & vChartInfo.YParameter(1) & """"
                For n = 1 To vChartInfo.YParameter.Count
                    tmpStr = tmpStr & "," & """" & vChartInfo.YParameter(n) & """"
                Next n
                tmpStr = Replace(tmpStr, "=", "")
                'nowChart.SeriesCollection(1).XValues = "={""A"",""B""}"
                nowChart.SeriesCollection(1).XValues = "={" & tmpStr & "}"
                'nowChart.SeriesCollection(2).XValues = "={" & tmpStr & "}"
                nowChart.Axes(xlCategory).TickLabels.Font.Size = 8
                'nowChart.Axes(xlCategory).MajorUnit = 1
            End If
            
            
            On Error GoTo 0
         Next j
      End If
   Next i
End Function

Public Sub Manual_Device_FitChart() '2010/06/10 陳宥先: 圖形格式特製化, 2012/11/22 泰慶:調整大小
    Dim i As Long, j As Long
    Dim nowSheet As Worksheet
    Dim nowChart As Chart
    
    For i = 1 To Worksheets.Count
        If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Or InStr(UCase(Worksheets(i).Name), "BOXTREND") > 0 Then
            Set nowSheet = Worksheets(i)
            For j = 1 To nowSheet.ChartObjects.Count
                '------------------------------------------------------
                ' Device Chart Format
                '------------------------------------------------------
                Set nowChart = nowSheet.ChartObjects(j).Chart
                With nowChart
                    .HasTitle = False
                    If .chartType = xlXYScatter Then
                        With .Legend
                            .Top = 225
                            .Left = 183
                            .AutoScaleFont = False
                            .Font.Size = 12
                            .Font.FontStyle = "粗體"
                            .Font.Name = "Arial"
                            .Border.LineStyle = xlNone
                        End With
                    End If
                    '長和寬
                    .Parent.width = 391.2   '340
                    .Parent.Height = 292.8   '334
                    .PlotArea.Top = 0
                    .PlotArea.Left = 49
                    .PlotArea.width = 305   '280 - 21
                    .PlotArea.Height = 249  '293 - 30
                    '.ChartArea.Width = 384
                    '.ChartArea.Height = 284
                    
                    With .PlotArea.Border
                        .ColorIndex = 57
                        .Weight = xlMedium
                        .LineStyle = xlContinuous
                    End With
                    .ChartArea.Border.LineStyle = 0
                    
                    With .Axes(xlCategory)
                        .AxisTitle.AutoScaleFont = False
                        .AxisTitle.Font.FontStyle = "粗體"
                        .AxisTitle.Font.Size = 14
                        .AxisTitle.Font.Name = "Arial"
                        .AxisTitle.Top = 262
                        .TickLabels.AutoScaleFont = False
                        .TickLabels.Font.FontStyle = "粗體"
                        .TickLabels.Font.Size = 12
                        .TickLabels.Font.Name = "Arial"
                        .MajorTickMark = xlInside
                        .MinorTickMark = xlInside
                        .HasMajorGridlines = True
                        .MajorGridlines.Border.ColorIndex = 15
                    End With
                    With .Axes(xlValue)
                        .AxisTitle.AutoScaleFont = False
                        .AxisTitle.Font.FontStyle = "粗體"
                        .AxisTitle.Font.Size = 14
                        .AxisTitle.Font.Name = "Arial"
                        .AxisTitle.Left = 3
                        .TickLabels.AutoScaleFont = False
                        .TickLabels.Font.FontStyle = "粗體"
                        .TickLabels.Font.Size = 12
                        .TickLabels.Font.Name = "Arial"
                        .TickLabels.NumberFormatLocal = "0.E+00"
                        .HasMajorGridlines = True
                        .HasMinorGridlines = True
                        .MajorGridlines.Border.ColorIndex = 15
                        .MinorGridlines.Border.ColorIndex = 15
                        .MinorGridlines.Border.LineStyle = xlDot
                        .MajorTickMark = xlInside
                        .MinorTickMark = xlInside
                    End With
                
                End With
                '------------------------------------------------------
            Next j
        End If
    Next i

    Set nowChart = Nothing
    Set nowSheet = Nothing
    
    Call GenChartSummary
    
    
'        ActiveSheet.ChartObjects("圖表 1").Activate
'    ActiveChart.ChartArea.Select
'    ActiveSheet.Shapes("圖表 1").ScaleWidth 0.89, msoFalse, msoScaleFromTopLeft
'    ActiveSheet.Shapes("圖表 1").ScaleHeight 0.97, msoFalse, msoScaleFromTopLeft
'    ActiveChart.Axes(xlValue).AxisTitle.Select
'    ActiveChart.ChartArea.Select
'    Selection.AutoScaleFont = False
'    With Selection.Font
'        .Name = "新細明體"
'        .FontStyle = "粗體"
'        .Size = 14
'    End With
'    ActiveChart.Axes(xlCategory).Select
'    ActiveChart.Legend.Select
'    Selection.Left = 289
'    Selection.Top = 189
'    ActiveChart.PlotArea.Select
'    Selection.Width = 353
    
    
    
    
End Sub

Public Sub FitChart()
   Dim i As Integer, j As Integer
   Dim scatterChart As Chart
   Dim boxtrendChart As Chart
   Dim nowChart As Chart
   Dim nowSheet As Worksheet
   
   If IsExistSheet("PlotSetup") Then
      Set nowSheet = Worksheets("PlotSetup")
      For j = 1 To nowSheet.ChartObjects.Count
         Set nowChart = nowSheet.ChartObjects(j).Chart
         If nowChart.HasTitle Then
            If UCase(nowChart.ChartTitle.Text) = "SCATTER" Then Set scatterChart = nowChart
            If UCase(nowChart.ChartTitle.Text) = "BOXTREND" Then Set boxtrendChart = nowChart
         End If
      Next j
      If scatterChart Is Nothing And nowSheet.ChartObjects.Count >= 1 Then Set scatterChart = nowSheet.ChartObjects(1).Chart
      If boxtrendChart Is Nothing And nowSheet.ChartObjects.Count >= 2 Then Set boxtrendChart = nowSheet.ChartObjects(2).Chart
   Else
      Exit Sub
   End If
  
   ynSeries = (MsgBox("Fit series style?" & vbCrLf & vbCrLf & "The function must have some time to run.", vbOKCancel) = vbOK)
   For i = 1 To Worksheets.Count
      'Call AppendFile("c:\test.log", Worksheets(i).Name & vbCrLf)
      If InStr(UCase(Worksheets(i).Name), "SCATTER") > 0 Then Call FitChartByType(Worksheets(i).Name, scatterChart, "SCATTER")
      If InStr(UCase(Worksheets(i).Name), "BOXTREND") > 0 Then Call FitChartByType(Worksheets(i).Name, boxtrendChart, "BOXTREND")
   Next i
    'Call FitChartByType("All_Chart", scatterChart, "SCATTER")
    Call GenChartSummary
End Sub
Public Function FitChartByType(mSheet As String, TemplateChart As Chart, mChartType As String)
   Dim j As Integer, m As Integer, i As Integer
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim nowLegend As Legend
   Dim nowSeries As Series
   Dim nowShape As Shape
   Dim nowAxis As Axis
   Dim nowDataLabels As DataLabels, TemplateDataLabel As DataLabel
   Dim nowPlotArea As PlotArea
   Dim boolTemp As Boolean
   Const NameList As String = "TARGET,CORNER,SS,FF,TT,GOLDEN,MEDIAN,USL,LSL"
   Dim tSeries As Series
   Dim tmp
   Dim YNMultiPara As Boolean
   
   On Error Resume Next
   If TemplateChart Is Nothing Then Exit Function
   
   Set nowSheet = Worksheets(mSheet)
   For j = 1 To nowSheet.ChartObjects.Count
      Set nowChart = nowSheet.ChartObjects(j).Chart
            '-----------------
            'Fit Chart Attribs
            '-----------------
            Set nowShape = nowSheet.Shapes(j)
            'Debug.Print nowShape.Name
            DoEvents
            If nowSheet.Name <> "All_Chart" Then
               With nowShape
                  .Top = 30 + (j - 1) * (TemplateChart.ChartArea.Height + 5)
                  .Left = 30
                  .width = TemplateChart.ChartArea.width
                  .Height = TemplateChart.ChartArea.Height
               End With
               DoEvents
            End If
            '----------
            'Fit Legend
            '----------
            nowChart.HasLegend = TemplateChart.HasLegend
            If nowChart.HasLegend Then
               Set nowLegend = nowChart.Legend
               'nowLegend.Position = TemplateChart.Legend.Position
               nowLegend.AutoScaleFont = False
               nowLegend.Font.Size = 8
               DoEvents
               nowLegend.Left = nowShape.width - nowLegend.width - 6 - 10
               nowLegend.Top = (nowShape.Height - nowLegend.Height) / 2 '+ 10
            End If
            DoEvents
            '-------------
            'Fit PlotArea
            '-------------
            Set nowPlotArea = nowChart.PlotArea
            nowPlotArea.width = nowLegend.Left - nowPlotArea.Left - 5
            
            Set nowPlotArea = Nothing
            Set nowLegend = Nothing
            Set nowShape = Nothing
            DoEvents
            '----------
            'Fit Axis
            '----------
            If mChartType <> "BOXTREND" Then
               Set nowAxis = nowChart.Axes(xlCategory)
               If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
               nowAxis.TickLabels.Font.Size = 10
               Set nowAxis = Nothing
            End If
            Set nowAxis = nowChart.Axes(xlValue)
            If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
            nowAxis.TickLabels.Font.Size = 10
            
            Set nowAxis = Nothing
            
      '----------
      'Fit Axis
      '----------
      If True Then ' X Axis
         Set nowAxis = nowChart.Axes(xlCategory, xlPrimary)
         With TemplateChart.Axes(xlCategory, xlPrimary)
            nowAxis.TickLabels.AutoScaleFont = True
            nowAxis.TickLabels.Font.Name = .TickLabels.Font.Name
            nowAxis.TickLabels.Font.Size = .TickLabels.Font.Size
            nowAxis.TickLabels.AutoScaleFont = False
            nowAxis.TickLabels.NumberFormatLocal = .TickLabels.NumberFormatLocal
            'If mChartType = "BOXTREND" Then nowAxis.TickLabels.NumberFormatLocal = "@"
            nowAxis.TickLabels.Orientation = .TickLabels.Orientation
            'End If
            nowAxis.HasTitle = .HasTitle
            If nowAxis.HasTitle Then
               nowAxis.AxisTitle.Font.Size = .AxisTitle.Font.Size
               nowAxis.AxisTitle.Font.Name = .AxisTitle.Font.Name
               nowAxis.AxisTitle.Font.ColorIndex = .AxisTitle.Font.ColorIndex
            End If
            nowAxis.HasMajorGridlines = .HasMajorGridlines
            If .HasMajorGridlines Then
                nowAxis.HasMajorGridlines = True
                nowAxis.MajorGridlines.Border.LineStyle = .MajorGridlines.Border.LineStyle
                nowAxis.MajorGridlines.Border.Weight = .MajorGridlines.Border.Weight
                nowAxis.MajorGridlines.Border.ColorIndex = .MajorGridlines.Border.ColorIndex
            End If
            nowAxis.HasMinorGridlines = .HasMinorGridlines
            If .HasMinorGridlines Then
                nowAxis.HasMinorGridlines = True
                nowAxis.MinorGridlines.Border.LineStyle = .MinorGridlines.Border.LineStyle
                nowAxis.MinorGridlines.Border.Weight = .MinorGridlines.Border.Weight
                nowAxis.MinorGridlines.Border.ColorIndex = .MinorGridlines.Border.ColorIndex
            End If
            
            
         End With
         Set nowAxis = Nothing
      End If
      If True Then ' Y Axis
         Set nowAxis = nowChart.Axes(xlValue, xlPrimary)
         With TemplateChart.Axes(xlValue, xlPrimary)
            nowAxis.TickLabels.AutoScaleFont = True
            nowAxis.TickLabels.Font.Name = .TickLabels.Font.Name
            nowAxis.TickLabels.Font.Size = .TickLabels.Font.Size
            nowAxis.TickLabels.AutoScaleFont = False
            nowAxis.TickLabels.NumberFormatLocal = .TickLabels.NumberFormatLocal
            'If mChartType = "BOXTREND" Then
            nowAxis.TickLabels.Orientation = .TickLabels.Orientation
            'End If
            nowAxis.HasTitle = .HasTitle
            If nowAxis.HasTitle Then
               nowAxis.AxisTitle.Font.Size = .AxisTitle.Font.Size
               nowAxis.AxisTitle.Font.Name = .AxisTitle.Font.Name
               nowAxis.AxisTitle.Font.ColorIndex = .AxisTitle.Font.ColorIndex
            End If
            nowAxis.HasMajorGridlines = .HasMajorGridlines
            If .HasMajorGridlines Then
                nowAxis.HasMajorGridlines = True
                nowAxis.MajorGridlines.Border.LineStyle = .MajorGridlines.Border.LineStyle
                nowAxis.MajorGridlines.Border.Weight = .MajorGridlines.Border.Weight
                nowAxis.MajorGridlines.Border.ColorIndex = .MajorGridlines.Border.ColorIndex
            End If
            nowAxis.HasMinorGridlines = .HasMinorGridlines
            If .HasMinorGridlines Then
                nowAxis.HasMinorGridlines = True
                nowAxis.MinorGridlines.Border.LineStyle = .MinorGridlines.Border.LineStyle
                nowAxis.MinorGridlines.Border.Weight = .MinorGridlines.Border.Weight
                nowAxis.MinorGridlines.Border.ColorIndex = .MinorGridlines.Border.ColorIndex
            End If
            
         End With
         Set nowAxis = Nothing
      End If
      DoEvents
      
      '-----------------
      'Fit Title
      '-----------------
      If nowChart.HasTitle Then
         With nowChart.ChartTitle
            .Font.Name = TemplateChart.ChartTitle.Font.Name
            .Font.Size = TemplateChart.ChartTitle.Font.Size
            .Font.ColorIndex = TemplateChart.ChartTitle.Font.ColorIndex
         End With
         DoEvents
      End If
      
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
               Set tSeries = TemplateChart.SeriesCollection("#" & CStr(m))
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
      If mChartType = "BOXTREND" Then
         ' Set DataLabel Style
         If nowChart.SeriesCollection(5).HasDataLabels And TemplateChart.SeriesCollection(5).HasDataLabels Then
            Set nowDataLabels = nowChart.SeriesCollection(5).DataLabels
            Set TemplateDataLabel = TemplateChart.SeriesCollection(5).Points(1).DataLabel
            nowDataLabels.AutoScaleFont = True
            nowDataLabels.Font.Name = TemplateDataLabel.Font.Name
            nowDataLabels.Font.FontStyle = TemplateDataLabel.Font.FontStyle
            nowDataLabels.Font.Size = TemplateDataLabel.Font.Size
            nowDataLabels.Font.Strikethrough = TemplateDataLabel.Font.Strikethrough
            nowDataLabels.Font.Superscript = TemplateDataLabel.Font.Superscript
            nowDataLabels.Font.Subscript = TemplateDataLabel.Font.Subscript
            nowDataLabels.Font.OutlineFont = TemplateDataLabel.Font.OutlineFont
            nowDataLabels.Font.Shadow = TemplateDataLabel.Font.Shadow
            nowDataLabels.Font.Underline = TemplateDataLabel.Font.Underline
            nowDataLabels.Font.ColorIndex = TemplateDataLabel.Font.ColorIndex
            nowDataLabels.Font.FontStyle = TemplateDataLabel.Font.FontStyle
            nowDataLabels.Font.Background = TemplateDataLabel.Font.Background
            YNMultiPara = False
            For Each tmp In nowChart.SeriesCollection(1).Values
                If tmp = "" Then YNMultiPara = True: Exit For
            Next tmp
            If Not YNMultiPara Then
                'nowDataLabels.Delete
                nowChart.SeriesCollection(5).ApplyDataLabels Type:=xlDataLabelsShowNone, LegendKey:=False
                nowChart.SeriesCollection(5).ApplyDataLabels Type:=xlDataLabelsShowValue, AutoText:=True, LegendKey:=False
                'nowDataLabels.Type = xlDataLabelsShowValue
            End If
         ElseIf nowChart.SeriesCollection(5).HasDataLabels And Not TemplateChart.SeriesCollection(5).HasDataLabels Then
            nowChart.SeriesCollection(5).HasDataLabels = TemplateChart.SeriesCollection(5).HasDataLabels
         End If
      End If
   Next j
   
   
   nowChart.Axes(xlCategory).TickLabels.AutoScaleFont = False
   nowChart.Axes(xlValue).TickLabels.AutoScaleFont = False
   nowChart.ChartTitle.AutoScaleFont = False
   
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

Public Function OLD_FitChartByType(mSheet As String, TemplateChart As Chart, mChartType As String)
   Dim j As Integer, m As Integer
   Dim nowSheet As Worksheet
   Dim nowChart As Chart
   Dim nowLegend As Legend
   Dim nowSeries As Series
   Dim nowShape As Shape
   Dim nowAxis As Axis
   Dim nowDataLabels As DataLabels, TemplateDataLabel As DataLabel
   Dim nowPlotArea As PlotArea
   
   On Error Resume Next
   If TemplateChart Is Nothing Then Exit Function
   
   Set nowSheet = Worksheets(mSheet)
   For j = 1 To nowSheet.ChartObjects.Count
      Set nowChart = nowSheet.ChartObjects(j).Chart
            '-----------------
            'Fit Chart Attribs
            '-----------------
            Set nowShape = nowSheet.Shapes(j)
            'Debug.Print nowShape.Name
            DoEvents
            If nowSheet.Name <> "All_Chart" Then
               With nowShape
                  .Top = 30 + (j - 1) * (300 + 100)
                  .Left = 30
                  .width = 450
                  .Height = 300
               End With
            DoEvents
            End If
            '----------
            'Fit Legend
            '----------
            If nowChart.HasLegend Then
               Set nowLegend = nowChart.Legend
               nowLegend.AutoScaleFont = False
               nowLegend.Font.Size = 8
               DoEvents
               nowLegend.Left = nowShape.width - nowLegend.width - 6 - 10
               nowLegend.Top = (nowShape.Height - nowLegend.Height) / 2 '+ 10
            End If
            DoEvents
            '-------------
            'Fit PlotArea
            '-------------
            Set nowPlotArea = nowChart.PlotArea
            nowPlotArea.width = nowLegend.Left - nowPlotArea.Left - 5
            
            Set nowPlotArea = Nothing
            Set nowLegend = Nothing
            Set nowShape = Nothing
            DoEvents
            '----------
            'Fit Axis
            '----------
            If mChartType <> "BOXTREND" Then
               Set nowAxis = nowChart.Axes(xlCategory)
               If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
               nowAxis.TickLabels.Font.Size = 10
            End If
            Set nowAxis = nowChart.Axes(xlValue)
            If nowAxis.ScaleType = xlScaleLogarithmic Then nowAxis.TickLabels.NumberFormatLocal = "0.0E+00"
            nowAxis.TickLabels.Font.Size = 10
            Set nowAxis = Nothing
            
'      '----------
'      'Fit Axis
'      '----------
'      If True Then ' X Axis
'         Set nowAxis = nowChart.Axes(xlCategory, xlPrimary)
'         With TemplateChart.Axes(xlCategory, xlPrimary)
'            nowAxis.TickLabels.AutoScaleFont = True
'            nowAxis.TickLabels.Font.Name = .TickLabels.Font.Name
'            nowAxis.TickLabels.Font.Size = .TickLabels.Font.Size
'            'If mChartType = "BOXTREND" Then
'            nowAxis.TickLabels.Orientation = .TickLabels.Orientation
'            'End If
'            nowAxis.HasTitle = .HasTitle
'            If nowAxis.HasTitle Then
'               nowAxis.AxisTitle.Font.Size = .AxisTitle.Font.Size
'               nowAxis.AxisTitle.Font.Name = .AxisTitle.Font.Name
'            End If
'         End With
'      End If
'      If True Then ' Y Axis
'         Set nowAxis = nowChart.Axes(xlValue, xlPrimary)
'         With TemplateChart.Axes(xlValue, xlPrimary)
'            nowAxis.TickLabels.AutoScaleFont = True
'            nowAxis.TickLabels.Font.Name = .TickLabels.Font.Name
'            nowAxis.TickLabels.Font.Size = .TickLabels.Font.Size
'            'If mChartType = "BOXTREND" Then
'            nowAxis.TickLabels.Orientation = .TickLabels.Orientation
'            'End If
'            nowAxis.HasTitle = .HasTitle
'            If nowAxis.HasTitle Then
'               nowAxis.AxisTitle.Font.Size = .AxisTitle.Font.Size
'               nowAxis.AxisTitle.Font.Name = .AxisTitle.Font.Name
'            End If
'         End With
'      End If

      '-----------------
      'Fit Series Style
      '-----------------
      If mChartType = "BOXTREND" Then
         For m = 1 To nowChart.SeriesCollection.Count
            Set nowSeries = nowChart.SeriesCollection(m)
            With TemplateChart.SeriesCollection(m)
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
         Next m
         ' Set DataLabel Style
         If nowChart.SeriesCollection(5).HasDataLabels And TemplateChart.SeriesCollection(5).HasDataLabels Then
            Set nowDataLabels = nowChart.SeriesCollection(5).DataLabels
            Set TemplateDataLabel = TemplateChart.SeriesCollection(5).Points(1).DataLabel
            nowDataLabels.AutoScaleFont = True
            nowDataLabels.Font.Name = TemplateDataLabel.Font.Name
            nowDataLabels.Font.FontStyle = TemplateDataLabel.Font.FontStyle
            nowDataLabels.Font.Size = TemplateDataLabel.Font.Size
            nowDataLabels.Font.Strikethrough = TemplateDataLabel.Font.Strikethrough
            nowDataLabels.Font.Superscript = TemplateDataLabel.Font.Superscript
            nowDataLabels.Font.Subscript = TemplateDataLabel.Font.Subscript
            nowDataLabels.Font.OutlineFont = TemplateDataLabel.Font.OutlineFont
            nowDataLabels.Font.Shadow = TemplateDataLabel.Font.Shadow
            nowDataLabels.Font.Underline = TemplateDataLabel.Font.Underline
            nowDataLabels.Font.ColorIndex = TemplateDataLabel.Font.ColorIndex
            nowDataLabels.Font.FontStyle = TemplateDataLabel.Font.FontStyle
            nowDataLabels.Font.Background = TemplateDataLabel.Font.Background
         ElseIf nowChart.SeriesCollection(5).HasDataLabels And Not TemplateChart.SeriesCollection(5).HasDataLabels Then
            nowChart.SeriesCollection(5).HasDataLabels = TemplateChart.SeriesCollection(5).HasDataLabels
         End If
      End If
   Next j
End Function


Public Sub Manual_SimpleSummary()
   Call SimpleSummary("SPEC_Simple", "Report_Simple")
End Sub

Private Function SimpleSummary(inSheet As String, outSheet As String)
   Dim specRange As Range
   Dim dataRange As Range
   Dim waferList() As String
   Dim iCol As Long, iRow As Long
   Dim outSpec As Integer
   Dim HColCount As Integer
   Dim HCol As Integer
   Dim siteNum As Integer
   Dim ProductID As String
   Dim mFactor
   Dim FactorSign1 As String, FactorSign2 As String
   Dim nowRange As Range
   Dim reValue As Variant
   Dim nowParameter As String
   Dim i As Long, j As Long
   
   Application.ScreenUpdating = False
   'SiteNum = Worksheets(TempDSheet).UsedRange.Columns.Count - 2
   siteNum = getSiteNum(dSheet)
   'ProductID = Trim(Worksheets(SRawData).Range("A2"))
   If ProductID = "A064A" Or ProductID = "A100A" Or ProductID = "A118" Then
      HColCount = 5
      HCol = 4
   Else
      HColCount = 4
      HCol = 3
   End If
   ' get Waferlist
   Call GetWaferList(dSheet, waferList)
   
   AddSheet (outSheet)
   'Worksheets(SSheet).UsedRange.Copy Worksheets(outsheet).Range("A1")
   'Worksheets(outsheet).Range("B:B").Delete Shift:=xlShiftToLeft
   'Worksheets(SRawData).Range("A1:" & "G" & CStr(Worksheets(SRawData).UsedRange.Rows.Count)).Copy Worksheets(outSheet).Range("A1")
   Worksheets(inSheet).UsedRange.Copy Worksheets(outSheet).Range("A1")
   'Worksheets(outSheet).Cells(2, 3) = UBound(WaferList) + 1
   Set specRange = Worksheets(outSheet).UsedRange

   
   ' fill wafer header
   Application.DisplayAlerts = False
   Worksheets(outSheet).Activate
   iRow = 2
   iCol = specRange.Columns.Count + 1
   For i = 0 To UBound(waferList)
      With Worksheets(outSheet)
         .Range(Num2Letter(iCol + i * HColCount) & CStr(iRow) & ":" & Num2Letter(iCol + i * HColCount + HCol) & CStr(iRow)).Select
         Selection.Merge
         Selection.Value = "#" & Trim(waferList(i))
         .Range(Num2Letter(iCol + i * HColCount) & CStr(iRow + 1) & ":" & Num2Letter(iCol + i * HColCount + HCol) & CStr(iRow + 1)).Select
         Selection.Cells(1, 1) = "Median"
         Selection.Cells(1, 2) = "Average"
         Selection.Cells(1, 3) = "3 Sigma"
         Selection.Cells(1, 4) = "Yield"
         If ProductID = "A064A" Or ProductID = "A100A" Or ProductID = "A118" Then Selection.Cells(1, 5) = "%"
      End With
      ' Add Split line
      Worksheets(outSheet).Range(Num2Letter(iCol + i * HColCount + HCol) & CStr(iRow) & ":" & Num2Letter(iCol + i * HColCount + HCol) & CStr(Worksheets(outSheet).UsedRange.Rows.Count)).Select
      With Selection.Borders(xlEdgeRight)
         .LineStyle = xlContinuous
         .Weight = xlMedium
         .ColorIndex = xlAutomatic
      End With
   Next i
   ' Add Split line
   Worksheets(outSheet).Range(Num2Letter(iCol) & CStr(iRow) & ":" & Num2Letter(iCol) & CStr(Worksheets(outSheet).UsedRange.Rows.Count)).Select
   With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .Weight = xlMedium
      .ColorIndex = xlAutomatic
   End With
   Worksheets(outSheet).Range(Num2Letter(iCol) & CStr(iRow) & ":" & Num2Letter(iCol + UBound(waferList) * HColCount + HCol) & CStr(iRow + 1)).Select
   'Selection.Font.Bold = True
   Selection.HorizontalAlignment = xlCenter
   Selection.Interior.ColorIndex = 8
   'Selection.Borders.LineStyle = xlContinuous
   Application.DisplayAlerts = True
   ' fill value
   Worksheets(outSheet).Activate
   iCol = specRange.Columns.Count + 1
   With Worksheets(outSheet)
      For i = 4 To .UsedRange.Rows.Count
         For j = 0 To UBound(waferList)
            .Range(Num2Letter(iCol + j * HColCount) & CStr(i) & ":" & Num2Letter(iCol + j * HColCount + HCol) & CStr(i)).Select
            nowParameter = Worksheets(outSheet).Cells(i, 2)
            Set reValue = getRangeByPara(waferList(j), nowParameter)
            If reValue Is Nothing Then Exit For
            Set nowRange = reValue
            Selection.Cells(1, 1) = Application.WorksheetFunction.Median(nowRange)
            Selection.Cells(1, 2) = Application.WorksheetFunction.Average(nowRange)
            Selection.Cells(1, 3) = Application.WorksheetFunction.StDev(nowRange) * 3
            If Worksheets(SSheet).Cells(i, 2) <> "" Then
               Selection.Cells(1, 1) = Selection.Cells(1, 1) * Worksheets(SSheet).Cells(i, 2)
               Selection.Cells(1, 2) = Selection.Cells(1, 2) * Worksheets(SSheet).Cells(i, 2)
               Selection.Cells(1, 3) = Abs(Selection.Cells(1, 3) * Worksheets(SSheet).Cells(i, 2))
            End If
            If ProductID = "A064A" Or ProductID = "A100A" Or ProductID = "A118" Then   ' A064A (%)
               If Trim(.Range(Num2Letter(5) & CStr(i))) <> "" Then _
                  Selection.Cells(1, 5) = Format((Selection.Cells(1, 1) - .Range(Num2Letter(5) & CStr(i))) / .Range(Num2Letter(5) & CStr(i)), "0.00%")
            End If
            'Selection.Cells(1, 4) = Selection.Cells(1, 3) * 3
            If Trim(.Range(Num2Letter(4) & CStr(i))) <> "" Or Trim(.Range(Num2Letter(6) & CStr(i))) <> "" Then
               Selection.Range("A1").Select  ' **Median
               Selection.FormatConditions.Delete
               Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & "F" & "" & CStr(i)
               Selection.FormatConditions(1).Font.ColorIndex = 3
               Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & "D" & "" & CStr(i)
               Selection.FormatConditions(2).Font.ColorIndex = 4
               'Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, _
               '   Formula1:="=$" & "D" & "$" & CStr(i), Formula2:="=$" & "F" & "$" & CStr(i)
               'Selection.FormatConditions(1).Font.ColorIndex = 3
               outSpec = 0
               If Worksheets(SSheet).Cells(i, 2) <> "" Then
                  mFactor = Worksheets(SSheet).Cells(i, 2)
               Else
                  mFactor = 1
               End If
               If mFactor >= 0 Then
                  FactorSign1 = "<"
                  FactorSign2 = ">"
               Else
                  FactorSign1 = ">"
                  FactorSign2 = "<"
               End If
               If Trim(.Range(Num2Letter(4) & CStr(i))) <> "" Then _
                  outSpec = outSpec + Application.WorksheetFunction.CountIf(nowRange, FactorSign1 & .Range(Num2Letter(4) & CStr(i)) / mFactor)
               If Trim(.Range(Num2Letter(6) & CStr(i))) <> "" Then _
                  outSpec = outSpec + Application.WorksheetFunction.CountIf(nowRange, FactorSign2 & .Range(Num2Letter(6) & CStr(i)) / mFactor)
               Selection.Range("D1") = Format((siteNum - outSpec) / siteNum, "0.00%")
               If outSpec <> 0 Then Selection.Range("D1").Font.ColorIndex = 3
            Else
               If Selection.Cells(1, 1) <> "" Then Selection.Cells(1, 4) = Format(1, "0.00%")
            End If
         Next j
      Next i
   End With
   
   'Sheet format
   'Call SummaryFormatByUnit(outSheet)
   Worksheets(outSheet).Activate
   Worksheets(outSheet).Cells.Select
   Selection.Font.Size = 10
   Selection.Font.Name = "Century Gothic"
   Worksheets(outSheet).Range("A1:" & Num2Letter(Worksheets(outSheet).UsedRange.Columns.Count) & CStr(3)).Select
   Selection.Font.Size = 12
   Worksheets(outSheet).Range("A4:" & "A" & CStr(Worksheets(outSheet).UsedRange.Rows.Count)).Select
   Selection.Font.Size = 12
   ActiveWindow.Zoom = 75
   Worksheets(outSheet).Cells.Select
   Selection.Columns.AutoFit
   Selection.Rows.AutoFit
   Worksheets(outSheet).Range("A4").Select
   ActiveWindow.FreezePanes = True
   Application.ScreenUpdating = True
   'MsgBox "sss"
End Function

Sub ExportChartAttrib()
    Dim nowSheet As Worksheet, sheetAllChart As Worksheet
    Const strHeader As String = "Chart,ChartTitle,XScale,XNumberFormat,XMax,XMin,Xmajor,Xminor,YScale,YNumberFormat,YMax,YMin,Ymajor,Yminor"
    Dim tempA
    Dim i As Integer, j As Integer
    Dim tmpStr As String
    Const sChar As String = "§"
    
    If IsExistSheet("All_Chart") Then
        Set sheetAllChart = Worksheets("All_Chart")
    Else
        Exit Sub
    End If
    Set nowSheet = AddSheet("ChartAttrib")
    tempA = Split(strHeader, ",")
    For i = 0 To UBound(tempA)
        nowSheet.Cells(1, i + 1) = tempA(i)
    Next i
    
    For i = 2 To sheetAllChart.UsedRange.Rows.Count
        If sheetAllChart.Cells(i, 1) = "" Then Exit For
        'nowSheet.Cells(i, 1) = sheetAllChart.Cells(i, 2).Value
        tmpStr = getChartAttrib(sheetAllChart.Cells(i, 2).Value)
        For j = 1 To UBound(tempA) + 1
            nowSheet.Cells(i, j) = "'" & getCOL(tmpStr, sChar, j)
        Next j
    Next i
    nowSheet.Columns.AutoFit
    Set nowSheet = Nothing
End Sub

Function getChartAttrib(mSheet As String)
    Dim tmpStr As String
    Dim reStr As String
    Dim nowChart As Chart, nowAxis As Axis
    Const sChar As String = "§"
    
    On Error Resume Next
    
    If Not IsExistSheet(mSheet) Then Exit Function
    If Worksheets(mSheet).ChartObjects.Count > 0 Then Set nowChart = Worksheets(mSheet).ChartObjects(1).Chart
    'Chart of Chart
    Set nowChart = Worksheets("All_Chart").ChartObjects(mSheet).Chart
    reStr = mSheet
    reStr = reStr & sChar & IIf(nowChart.HasTitle, nowChart.ChartTitle.Text, "")
    Set nowAxis = nowChart.Axes(xlCategory, xlPrimary)
    reStr = reStr & sChar & IIf(nowAxis.ScaleType = xlLinear, "Linear", "Log")
    reStr = reStr & sChar & nowAxis.TickLabels.NumberFormatLocal
    reStr = reStr & sChar & nowAxis.MaximumScale
    reStr = reStr & sChar & nowAxis.MinimumScale
    reStr = reStr & sChar & nowAxis.MajorUnit
    reStr = reStr & sChar & nowAxis.MinorUnit
     Set nowAxis = nowChart.Axes(xlValue, xlPrimary)
    reStr = reStr & sChar & IIf(nowAxis.ScaleType = xlLinear, "Linear", "Log")
    reStr = reStr & sChar & nowAxis.TickLabels.NumberFormatLocal
    reStr = reStr & sChar & nowAxis.MaximumScale
    reStr = reStr & sChar & nowAxis.MinimumScale
    reStr = reStr & sChar & nowAxis.MajorUnit
    reStr = reStr & sChar & nowAxis.MinorUnit
    getChartAttrib = reStr
End Function

Sub ImportChartAttrib()
    Dim nowSheet As Worksheet
    'Const strHeader As String = "Chart,ChartTitle,XScale,XMax,XMin,Xmajor,Xminor,YScale,YMax,YMin,Ymajor,Yminor"
    'Dim TempA
    Dim i As Integer, j As Integer
    Dim tmpStr As String
    Const sChar As String = "§"
    
    If IsExistSheet("ChartAttrib") Then
        Set nowSheet = Worksheets("ChartAttrib")
    Else
        Exit Sub
    End If

    
    For i = 2 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(i, 1) = "" Then Exit For
        tmpStr = nowSheet.Cells(i, 1)
        For j = 2 To nowSheet.UsedRange.Columns.Count
            tmpStr = tmpStr & sChar & nowSheet.Cells(i, j).Text
        Next j
        Call setChartAttrib(tmpStr)
    Next i
    
End Sub

Function setChartAttrib(tmpStr As String)
    Const sChar As String = "§"
    Const ChartID = 1
    Const ChartTitle = 2
    Const XScale = 3
    Const XNumberFormat = 4
    Const xMax = 5
    Const xMin = 6
    Const Xmajor = 7
    Const Xminor = 8
    Const YScale = 9
    Const YNumberFormat = 10
    Const yMax = 11
    Const yMin = 12
    Const Ymajor = 13
    Const Yminor = 14

'    Dim TmpStr As String
'    Dim reStr As String
    Dim mSheet As String
    Dim nowChart As Chart, nowAxis As Axis
    On Error Resume Next

    'For Excel XP
    If Application.Version <> "9.0" Then tmpStr = Replace(tmpStr, "G/通用格式", "")
    
    mSheet = getCOL(tmpStr, sChar, 1)
    If Not IsExistSheet(mSheet) Then Exit Function
    If Worksheets(mSheet).ChartObjects.Count > 0 Then Set nowChart = Worksheets(mSheet).ChartObjects(1).Chart
    'Chart of All_Chart
    Set nowChart = Worksheets("All_Chart").ChartObjects(mSheet).Chart
    If nowChart.HasTitle Then nowChart.ChartTitle.Text = getCOL(tmpStr, sChar, ChartTitle)

    'X-Axis
    '-------------------------------------------------
    Set nowAxis = nowChart.Axes(xlCategory, xlPrimary)
    nowAxis.TickLabels.NumberFormatLocal = getCOL(tmpStr, sChar, XNumberFormat)
    If getCOL(tmpStr, sChar, xMax) <> "" Then
        nowAxis.MaximumScale = Val(getCOL(tmpStr, sChar, xMax))
    Else
        nowAxis.MaximumScaleIsAuto = True
    End If
    If getCOL(tmpStr, sChar, xMin) <> "" Then
        nowAxis.MinimumScale = Val(getCOL(tmpStr, sChar, xMin))
    Else
        nowAxis.MinimumScaleIsAuto = True
    End If
    If getCOL(tmpStr, sChar, Xmajor) <> "" Then
        nowAxis.MajorUnit = Val(getCOL(tmpStr, sChar, Xmajor))
    Else
        nowAxis.MajorUnitIsAuto = True
    End If
    If getCOL(tmpStr, sChar, Xminor) <> "" Then
        nowAxis.MinorUnit = Val(getCOL(tmpStr, sChar, Xminor))
    Else
        nowAxis.MinorUnitIsAuto = True
    End If
    nowAxis.CrossesAt = xlCustom
    nowAxis.CrossesAt = nowAxis.MinimumScale
    nowAxis.ScaleType = IIf(UCase(getCOL(tmpStr, sChar, XScale)) = "LINEAR", xlLinear, xlLogarithmic)
    
    'X-Axis
    '-------------------------------------------------
    Set nowAxis = nowChart.Axes(xlValue, xlPrimary)
    nowAxis.TickLabels.NumberFormatLocal = getCOL(tmpStr, ",", YNumberFormat)
    If getCOL(tmpStr, sChar, yMax) <> "" Then
        nowAxis.MaximumScale = Val(getCOL(tmpStr, sChar, yMax))
    Else
        nowAxis.MaximumScaleIsAuto = True
    End If
    If getCOL(tmpStr, sChar, yMin) <> "" Then
        nowAxis.MinimumScale = Val(getCOL(tmpStr, sChar, yMin))
    Else
        nowAxis.MinimumScaleIsAuto = True
    End If
    
    If getCOL(tmpStr, sChar, Ymajor) <> "" Then
        nowAxis.MajorUnit = Val(getCOL(tmpStr, sChar, Ymajor))
    Else
        nowAxis.MajorUnitIsAuto = True
    End If
    If getCOL(tmpStr, sChar, Yminor) <> "" Then
        nowAxis.MinorUnit = Val(getCOL(tmpStr, sChar, Yminor))
    Else
        nowAxis.MinorUnitIsAuto = True
    End If
    nowAxis.CrossesAt = xlCustom
    nowAxis.CrossesAt = nowAxis.MinimumScale
    nowAxis.ScaleType = IIf(UCase(getCOL(tmpStr, sChar, YScale)) = "LINEAR", xlLinear, xlLogarithmic)
End Function

Sub Manual_RawDataReSorting()
    Dim nowSheet As Worksheet
    Dim i As Long
    Dim nowRange As Range
    
    Set nowSheet = Worksheets("Data")
    For i = 1 To nowSheet.Names.Count
        If InStr(nowSheet.Names(i).Name, "wafer") > 0 Then
            Set nowRange = nowSheet.Names(i).RefersToRange
            If nowRange.Cells(2, 1).Text <> "1" Then
                nowRange.Sort nowRange.Range("A1"), xlAscending, , , , , , xlYes
            Else
                nowRange.Sort nowRange.Range("B1"), xlAscending, , , , , , xlYes
            End If
        End If
    Next i
End Sub

Function removeTTFFSS()
    Dim nowSheet As Worksheet
    Dim i As Integer, j As Integer
    
    For j = 1 To Worksheets.Count
        If Left(Worksheets(j).Name, 7) = "SCATTER" Then
            Set nowSheet = Worksheets(j)
            For i = nowSheet.UsedRange.Columns.Count To 3 Step -1
                Select Case nowSheet.Cells(1, i)
                    Case "TT", "SS", "FF"
                        nowSheet.Columns(N2L(i) & ":" & N2L(i + 1)).Delete Shift:=xlToLeft
                End Select
            Next i
        End If
    Next j
End Function

Sub ReDraw()
   'Call GenChartHeader
   'Call GenScatter
   'Call GenBoxTrend
   'Call GenCumulative
   If Not IsExistSheet("PlotSetup") Then Exit Sub
   Call removeTTFFSS
   Call DioPlotAllChart
   'Call Manual_FitLegend
   Call new_FitChart
   Call CornerCount  'New function
   Call RawdataRange 'New function
   Call GenChartSummary
   'Application.StatusBar = "Finished!"
End Sub


Sub Manual_RemoveChartLink()
    Dim i As Integer, j As Integer
    Dim nowSheet As Worksheet
    Dim nowChart As Chart
    Dim nowShape As Shape
    Dim iSheet As Integer
    
    If Not IsExistSheet("All_Chart") Then Exit Sub
    
    Set nowSheet = Worksheets("All_Chart")
    
    For i = 1 To nowSheet.ChartObjects.Count
        Set nowChart = nowSheet.ChartObjects(i).Chart
        'Debug.Print nowChart.Name
        For j = nowChart.Shapes.Count To 1 Step -1
            If InStr(nowChart.Shapes(j).Hyperlink.ScreenTip, "SCATTER") > 0 Or _
               InStr(nowChart.Shapes(j).Hyperlink.ScreenTip, "BOXTREND") > 0 Or _
               InStr(nowChart.Shapes(j).Hyperlink.ScreenTip, "CUMULATIVE") > 0 Then _
               nowChart.Shapes(j).Delete
            'Debug.Print nowChart.Shapes(i).Hyperlink.ScreenTip
            'Debug.Print nowChart.Shapes(i).AlternativeText
        Next j
    Next i
    
    For iSheet = 1 To Worksheets.Count
        If InStr(UCase(Worksheets(iSheet).Name), "SCATTER") > 0 Or _
           InStr(UCase(Worksheets(iSheet).Name), "BOXTREND") > 0 Or _
           InStr(UCase(Worksheets(iSheet).Name), "CUMULATIVE") > 0 Then
    
            Set nowSheet = Worksheets(iSheet)
            Debug.Print nowSheet.Name
            For i = 1 To nowSheet.ChartObjects.Count
                Set nowChart = nowSheet.ChartObjects(i).Chart
                'Debug.Print nowChart.Name
                For j = nowChart.Shapes.Count To 1 Step -1
                    If InStr(nowChart.Shapes(j).Hyperlink.ScreenTip, "Chart List") > 0 Then _
                       nowChart.Shapes(j).Delete
                    'Debug.Print nowChart.Shapes(i).Hyperlink.ScreenTip
                    'Debug.Print nowChart.Shapes(i).AlternativeText
                Next j
            Next i
           
        End If
    Next iSheet
End Sub

Sub Manual_vincent_Ratio()
    Dim xYN As Boolean
    Dim yYN As Boolean
    Dim ynString As String
    Dim iRow As Long, iCol As Long
    Dim bRow As Long
    Dim nowSheet As Worksheet
    Dim i As Long
    
    ynString = InputBox("請輸入 X,Y 是否要取 Ratio:", , "0,1")
    xYN = IIf(getCOL(ynString, ",", 1) = "1", True, False)
    yYN = IIf(getCOL(ynString, ",", 2) = "1", True, False)
    
    Set nowSheet = ActiveSheet
    With nowSheet
        For i = 1 To nowSheet.UsedRange.Rows.Count
            If nowSheet.Cells(i, 1) = "Y" Then bRow = i + 1: Exit For
        Next i
        For iCol = 3 To nowSheet.UsedRange.Columns.Count Step 2
            For iRow = nowSheet.UsedRange.Rows.Count To bRow Step -1
                If nowSheet.Cells(iRow, iCol) <> "" And yYN Then
                   nowSheet.Cells(iRow, iCol) = "'=(" & CStr(nowSheet.Cells(iRow, iCol)) & "-" & _
                                               CStr(nowSheet.Cells(bRow, iCol)) & ")/" & _
                                               CStr(nowSheet.Cells(bRow, iCol))
                End If
                If nowSheet.Cells(iRow, iCol + 1) <> "" And xYN Then
                   nowSheet.Cells(iRow, iCol + 1) = "'=(" & CStr(nowSheet.Cells(iRow, iCol + 1)) & "-" & _
                                               CStr(nowSheet.Cells(bRow, iCol + 1)) & ")/" & _
                                               CStr(nowSheet.Cells(bRow, iCol + 1))
                End If
                
            Next iRow
        Next iCol
    End With
    
    MsgBox "ok"
End Sub

Sub CPK_Table()
    Dim i As Long, j As Long
    Dim setSheet As Worksheet, nowSheet As Worksheet
    Dim setRange As Range
    Dim Class1 As String, Class2 As String, notClass As String
    Dim waferList() As String
    Dim nItem As Integer, nPass As Integer
    Dim tmp As String
    
    'Get setRange
    If Not IsExistSheet("CPK_Option") Then Exit Sub
    Set setSheet = Worksheets("CPK_Option")
    For i = 1 To setSheet.UsedRange.Columns.Count
        If UCase(Trim(setSheet.Cells(1, i))) = "CPK_TABLE" Then
            'Debug.Print i
            'Debug.Print setSheet.Cells(1, i).End(xlDown).Address
            Set setRange = setSheet.Range(setSheet.Cells(2, i), setSheet.Cells(1, i).End(xlDown))
            'Debug.Print setRange.Address
            Exit For
        End If
    Next i
             
    Set nowSheet = AddSheet("CPK_Table")
    nowSheet.Cells(1, 1) = getCOL(Worksheets(dSheet).Range("B3"), ":", 2)
    nowSheet.Cells(1, 3) = "Lot"
    nowSheet.Cells(2, 1) = "Category"
    nowSheet.Cells(2, 2) = "S & M Item"
    nowSheet.Cells(2, 3) = "pass item"
    nowSheet.Cells(2, 4) = "pass %"
    nowSheet.Range("C1:D1").Merge
    With nowSheet.Range("A1:D2")
        .Interior.ColorIndex = 20
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    setRange.Copy nowSheet.Range("A3")
    
    Call GetWaferList(dSheet, waferList)
    
    For j = 3 To nowSheet.UsedRange.Rows.Count
        Class1 = "Class"
        Class2 = "Class2"
        notClass = "notClass"
        tmp = getCOL(getCOL(nowSheet.Cells(j, 1), "(", 2), ")", 1)
        If InStr(tmp, "&") > 0 Then
            Class1 = getCOL(tmp, "&", 1)
            Class2 = getCOL(tmp, "&", 2)
        Else
            Class1 = tmp
        End If
        If Class1 = "RS" Then notClass = "RS_M"
        For i = 0 To UBound(waferList)
            If j = 3 Then
                nowSheet.Cells(1, 5 + i * 2) = "#" & waferList(i)
                With nowSheet.Range(nowSheet.Cells(1, 5 + i * 2), nowSheet.Cells(1, 5 + i * 2 + 1))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .Interior.ColorIndex = 38
                    .Borders.LineStyle = xlContinuous
                End With
                nowSheet.Cells(2, 5 + i * 2) = "pass item"
                nowSheet.Cells(2, 5 + i * 2 + 1) = "pass %"
                With nowSheet.Range(nowSheet.Cells(2, 5 + i * 2), nowSheet.Cells(2, 5 + i * 2 + 1))
                    .HorizontalAlignment = xlCenter
                    .Interior.ColorIndex = 20
                    .Borders.LineStyle = xlContinuous
                End With
            End If
            tmp = getPassItemByClass(waferList(i), Class1, Class2, notClass)
            nItem = CInt(getCOL(tmp, ",", 1))
            nPass = CInt(getCOL(tmp, ",", 2))
            If i = 0 Then nowSheet.Cells(j, 2) = nItem
            nowSheet.Cells(j, 5 + i * 2) = nPass
            nowSheet.Cells(j, 5 + i * 2 + 1) = Format(nPass / nItem, "00.0%")
        Next i
    Next j
    nowSheet.Columns.AutoFit
End Sub

Function getPassItemByClass(ByVal mWafer As String, ByVal mClass As String, Optional mClass2 As String = "Class2", Optional notClass As String = "notClass")
    Dim nowSheet As Worksheet
    Dim iRow As Long
    Dim iCol As Long
    Dim nItem As Integer, nPass As Integer
    
    If Not IsExistSheet("All_Summary") Then Exit Function
    
    Set nowSheet = Worksheets("All_Summary")
    
    For iCol = 7 To nowSheet.UsedRange.Columns.Count
        If nowSheet.Cells(1, iCol) = mWafer And nowSheet.Cells(2, iCol) = "CPK" Then
            For iRow = 3 To nowSheet.UsedRange.Rows.Count
                If (Left(UCase(nowSheet.Cells(iRow, 2)), Len(mClass)) = UCase(mClass) _
                   Or Left(UCase(nowSheet.Cells(iRow, 2)), Len(mClass2)) = UCase(mClass2)) _
                   And Left(UCase(nowSheet.Cells(iRow, 2)), Len(notClass)) <> UCase(notClass) Then
                    'Debug.Print nowSheet.Cells(iRow, 2)
                    nItem = nItem + 1
                    If nowSheet.Cells(iRow, iCol) >= 1.33 Then nPass = nPass + 1
                End If
            Next iRow
        End If
    Next iCol
    
    getPassItemByClass = nItem & "," & nPass
End Function

' For 莊子健
Sub AddBoxSigmaPercentage()
    Dim nowSheet As Worksheet, nowChart As Chart, nowSeries As Series, nowRange As Range, nowAxis As Axis
    Dim bRow As Long, iCol As Long, iRow As Long, i As Long, j As Long
    Dim valueStr As String, mColl As New Collection
    Dim tempA As Variant
    
    
    For i = 1 To Worksheets.Count
        On Error GoTo myError
        If UCase(Left(Worksheets(i).Name, 8)) = "BOXTREND" Then
            Set nowSheet = Worksheets(i)
            iCol = nowSheet.UsedRange.Columns.Count
            iRow = nowSheet.Range("E5").CurrentRegion.Rows.Count
            bRow = nowSheet.Range("D4").End(xlDown).row + 1
            For j = 1 To mColl.Count: mColl.Remove 1: Next j
            For j = 5 To iCol
                If nowSheet.Cells(1, j) <> "" Then
                    Set nowRange = nowSheet.Range(N2L(j) & CStr(bRow) & ":" & N2L(j) & CStr(iRow))
                    mColl.Add Round(3 * WorksheetFunction.StDev(nowRange) / WorksheetFunction.Average(nowRange), 1)
                Else
                    mColl.Add 0, CStr(j)
                End If
            Next j
            valueStr = "{" & mColl(1)
            'valueStr = mColl(1)
            For j = 2 To mColl.Count: valueStr = valueStr & "," & mColl(j): Next j
            valueStr = valueStr & "}"
            'Debug.Print valueStr
            'TempA = Split(valueStr, ",")
            Set nowChart = nowSheet.ChartObjects(1).Chart
            Set nowSeries = nowChart.SeriesCollection.NewSeries
            nowSeries.Name = "Sigma%"
            nowSeries.chartType = xlXYScatterLines
            nowSeries.Values = valueStr
            nowSeries.MarkerStyle = xlMarkerStyleDiamond
            'nowSeries.MarkerBackgroundColorIndex = nowSeries.MarkerForegroundColor
            nowSeries.MarkerBackgroundColorIndex = 2
            nowSeries.MarkerForegroundColorIndex = 5
            nowSeries.Border.ColorIndex = 5
            nowSeries.Border.Weight = xlThick
            
            For j = 2 To nowSeries.Points.Count - 1
                If WorksheetFunction.index(nowSeries.Values, j) = 0 Then
                    nowSeries.Points(j).Border.LineStyle = xlNone
                    nowSeries.Points(j + 1).Border.LineStyle = xlNone
                    nowSeries.Points(j).MarkerStyle = xlNone
                End If
            Next j
        
            'nowSeries.MarkerSize = 10
            nowSeries.AxisGroup = xlSecondary
            'nowSeries.Values = valueStr
            Set nowAxis = nowChart.Axes(xlValue, xlSecondary)
            nowAxis.TickLabels.NumberFormatLocal = "0%"
            nowAxis.HasTitle = True
            nowAxis.AxisTitle.Characters.Text = "3 Sigma%"
            nowAxis.AxisTitle.Font.ColorIndex = 5
            nowAxis.TickLabels.Font.ColorIndex = 5
        End If
myError:
        
    Next i

End Sub

' For 林榮祥
Public Function Manual_MedianScore(ByVal mSheet As String)
    Dim nowSheet As Worksheet
    Dim iCol As Long, iRow As Long
    Dim i As Long
    Dim Sum As Integer, Pass As Integer
    
    Set nowSheet = Worksheets(mSheet)
    
    iRow = nowSheet.UsedRange.Rows.Count + 1
    For iCol = 7 To nowSheet.UsedRange.Columns.Count
        If nowSheet.Cells(2, iCol) = "Median" Then
            For i = 3 To nowSheet.UsedRange.Rows.Count
                If nowSheet.Cells(i, iCol) <> "" Then Sum = Sum + 1
                If nowSheet.Cells(i, iCol) > nowSheet.Cells(i, 4) And nowSheet.Cells(i, iCol) < nowSheet.Cells(i, 6) Then Pass = Pass + 1
            Next i
            nowSheet.Cells(iRow, iCol) = Pass / Sum
            nowSheet.Cells(iRow, iCol).NumberFormat = "0%"
        End If
    Next iCol
End Function

'For 蘇建中
Public Function WAT_CP_Data()
    Dim aWafer() As String
    Dim aSite() As String
    Dim i As Long, j As Long
    Dim setupSheet As Worksheet
    Dim sheetName As String
   
    Call GetWaferList(dSheet, aWafer)
    Call GetSiteList(dSheet, aSite)
    If Not IsExistSheet("setupWaferMap") Then Exit Function
    Set setupSheet = Worksheets("setupWaferMap")
    
    'del old sheet
    For i = Worksheets.Count To 1 Step -1
        If Left(Worksheets(i).Name, 11) = "WAT_CP_Data" Then DelSheet (Worksheets(i).Name)
    Next i
    
    If (UBound(aWafer) + 1) * WorksheetFunction.countA(setupSheet.Range("A:A")) > 254 Then
        sheetName = "WAT_CP_Data_1"
    Else
        sheetName = "WAT_CP_Data"
    End If
    
    For i = 1 To setupSheet.UsedRange.Rows.Count
        If setupSheet.Cells(i, 1) = "" Then Exit For
        'Debug.Print setupSheet.Cells(i, 1)
        sheetName = WAT_CP_Data_sub(sheetName, aWafer, aSite, setupSheet.Cells(i, 1))
    Next i
    
    Debug.Print "WAT_CP Finished!"
End Function

'For 蘇建中
Public Function WAT_CP_Data_sub(mSheet As String, aWafer() As String, aSite() As String, nowPara As String)
    Dim nowSheet As Worksheet
    Dim tmp As String
    Dim i As Long, j As Long
    Dim iCol As Long
    
    If Not IsExistSheet(mSheet) Then AddSheet (mSheet)
    If UBound(aWafer) + 1 + Worksheets(mSheet).UsedRange.Columns.Count <= 256 Then
        Set nowSheet = AddSheet(mSheet, False)
    Else
        tmp = getCOL(mSheet, "_", 4)
        Set nowSheet = AddSheet("WAT_CP_Data" & CStr(CInt(tmp) + 1), False)
    End If
    
    If nowSheet.UsedRange.Rows.Count < 2 Then
        nowSheet.Cells(2, 1) = "Sequence"
        nowSheet.Cells(1, 2) = "Parameter"
        nowSheet.Cells(2, 2) = "Coordinate"
        For i = 0 To UBound(aSite)
            nowSheet.Cells(i + 3, 1) = i + 1
            nowSheet.Cells(i + 3, 2) = "(" & getCOL(aSite(i), "(", 2)
        Next i
    End If
    
    iCol = nowSheet.UsedRange.Columns.Count + 1
    For i = 0 To UBound(aWafer)
        nowSheet.Cells(1, iCol + i) = nowPara
        nowSheet.Cells(2, iCol + i) = aWafer(i)
        For j = 0 To UBound(aSite)
            nowSheet.Cells(j + 3, iCol + i) = getValueByPara(aWafer(i), nowPara, j + 1)
        Next j
    Next i
    
    WAT_CP_Data_sub = nowSheet.Name
End Function

Public Sub Manual_RotateNotch()
    Dim rType As String
    Dim iRow As Long, iCol As Long
    Dim nowSheet As Worksheet
    Dim tmp As String, x As String, y As String, siteStr As String
    Dim tmpStr As String
    
    rType = InputBox("1-逆時針旋轉90度", "Input Rotate", "1")
    If rType = "" Then Exit Sub
    
    Set nowSheet = Worksheets("Data")
    For iRow = 1 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(iRow, 1) = "No./DataType" Then Exit For
    Next iRow
    If iRow >= nowSheet.UsedRange.Rows.Count Then MsgBox "can't find wafer information": Exit Sub
    
    For iCol = 4 To nowSheet.UsedRange.Columns.Count
        If InStr(nowSheet.Cells(iRow, iCol), "(") <= 0 Then Exit For
        tmpStr = nowSheet.Cells(iRow, iCol)
        siteStr = getCOL(tmpStr, "(", 1)
        x = getCOL(getCOL(tmpStr, "(", 2), ",", 1)
        y = getCOL(getCOL(tmpStr, ")", 1), ",", 2)
        Select Case rType
            Case "1": tmp = x: x = y * -1: y = tmp
        End Select
        tmpStr = siteStr & "(" & x & "," & y & ")"
        'Debug.Print nowSheet.Cells(iRow, iCol), TmpStr
        nowSheet.Cells(iRow, iCol) = tmpStr
    Next iCol
    MsgBox "Finished"
End Sub


'For 吳泰慶

Public Sub Manual_IonChart()
    Dim ChartSheet As Worksheet, nowSheet As Worksheet
    Dim i As Long, j As Long
    Dim nowChart As Chart
    Dim nowSeries As Series
    Dim nowTrend As Trendline
    Dim SeriesNum As Integer
    Dim TTx As Single, TTy As Single
    Dim tmp As String, tmpValue As Single
    Dim tmp2 As String
    
    If Not IsExistSheet("All_Chart") Then Exit Sub
    Set ChartSheet = Worksheets("All_Chart")
    Set nowSheet = AddSheet("Ion_Chart")
    
    SeriesNum = 0
    For i = 1 To ChartSheet.ChartObjects.Count
        'Debug.Print chartSheet.ChartObjects(i).Name
        Set nowChart = ChartSheet.ChartObjects(i).Chart
        'Chart Name
        nowSheet.Cells(1, i + 1) = ChartSheet.ChartObjects(i).Name
        'Get Trendline Equation String
        '-------------------------------
        For j = 1 To nowChart.SeriesCollection.Count
            Set nowSeries = nowChart.SeriesCollection(j)
            'Debug.Print nowSeries.Name
            If i = 1 Then nowSheet.Cells(j + 1, 1) = nowSeries.Name
            If nowSeries.Name = "TT" Or nowSeries.Name = "FF" Then
                nowSheet.Cells(j + 1, i + 1) = WorksheetFunction.index(nowSeries.XValues, 1) & "," & WorksheetFunction.index(nowSeries.Values, 1)
                If nowSeries.Name = "TT" Then
                    TTx = WorksheetFunction.index(nowSeries.XValues, 1)
                    TTy = WorksheetFunction.index(nowSeries.Values, 1)
                    'nowSheet.Cells(j + 1, i + 1) = TTx & "," & TTy
                End If
                If SeriesNum = 0 Then SeriesNum = j - 1
            Else
                '2010/06/17 RVN/HVN and Ion-Ioffd 用二項式, RVP/HVP用直線
                'If UCase(Right(getCOL(nowChart.ChartTitle.Characters.Text, "_", 1), 1)) = "N" Then
                If UCase(getCOL(nowChart.ChartTitle.Characters.Text, "_", 1)) = "RVN+HVN" And i Mod 3 = 1 Then
                    Set nowTrend = nowSeries.Trendlines.Add(Type:=xlPolynomial, order:=2 _
                                    , Forward:=0, Backward:=0, DisplayEquation:=True, DisplayRSquared:=False)
                Else
                    Set nowTrend = nowSeries.Trendlines.Add(Type:=xlLinear, Forward:=0 _
                                    , Backward:=0, DisplayEquation:=True, DisplayRSquared:=False)
                End If
                nowTrend.DataLabel.NumberFormatLocal = "0.0000000000E+00"
                nowSheet.Cells(j + 1, i + 1) = nowTrend.DataLabel.Caption
                nowTrend.Delete
                'Debug.Print nowChart.ChartTitle.Characters.Text, nowSeries.Name, nowSheet.Cells(j + 1, i + 1)
            End If
        Next j
        'Exit Sub
        ' Calc x value with TT Y value
        '---------------------------------
        For j = 1 To SeriesNum
'            tmp = getCOL(nowSheet.Cells(j + 1, i + 1), "=", 2)
'            tmp = "(" & CStr(TTy) & "-(" & getCOL(tmp, "x", 2) & "))/" & getCOL(tmp, "x", 1)
'            tmpValue = Application.Evaluate(tmp)
            tmpValue = getXfromEquation(nowSheet.Cells(j + 1, i + 1), TTy)
            If i Mod 3 = 1 Then
                nowSheet.Cells(j + 1, i + 1) = tmpValue / TTx - 1   'Ion 取差值
            Else
                nowSheet.Cells(j + 1, i + 1) = tmpValue
            End If
            'Debug.Print nowSheet.Cells(j + 1, i + 1), tmp, tmpValue
        Next j
    Next i
    
    'Generate Charts
    '-------------------
    For i = 1 To ChartSheet.ChartObjects.Count Step 3
        'Ion Vs Cgd
        '--------------------
        Set nowChart = nowSheet.ChartObjects.Add(10, 10 + (i \ 3) * 210, 300, 200).Chart
        nowChart.chartType = xlXYScatter
        For j = 1 To SeriesNum
            Set nowSeries = nowChart.SeriesCollection.NewSeries
            nowSeries.XValues = nowSheet.Cells(j + 1, i + 1 + 1)    'Cgd
            nowSeries.Values = nowSheet.Cells(j + 1, i + 1) 'Ion
            nowSeries.Name = nowSheet.Cells(j + 1, 1)
        Next j
        tmp = ChartSheet.ChartObjects(i).Chart.ChartTitle.Caption
        tmp = Replace(tmp, "Ion-Ioffs", "Ion-Cgd")
        nowChart.HasTitle = True
        nowChart.ChartTitle.Caption = tmp
        nowChart.Legend.AutoScaleFont = False
        nowChart.Legend.Font.Size = 8
        
        tmp2 = getCOL(getCOL(tmp, "_", 1), ",", 1)
        tmp = getCOL(getCOL(tmp, "(", 2), ")", 1)
        tmp = UCase(Mid(tmp, 1, 1)) & Mid(tmp, 2)
        With nowChart.Axes(xlValue)
            .Crosses = xlCustom
            .CrossesAt = .MinimumScale
            .HasTitle = True
            .AxisTitle.Characters.Text = tmp & "Ioffs_" & tmp2 & "_Ion(%)"
            .AxisTitle.AutoScaleFont = False
            .AxisTitle.Font.Size = 8
        End With
        With nowChart.Axes(xlCategory)
            .Crosses = xlCustom
            .CrossesAt = .MinimumScale
            .HasTitle = True
            .AxisTitle.Characters.Text = tmp & "Ioffs_" & tmp2 & "_Cgd(fF/um)"
            .AxisTitle.AutoScaleFont = False
            .AxisTitle.Font.Size = 8
        End With
    
        'Ion Vs Vts
        '--------------------
        Set nowChart = nowSheet.ChartObjects.Add(10 + 310, 10 + (i \ 3) * 210, 300, 200).Chart
        nowChart.chartType = xlXYScatter
        For j = 1 To SeriesNum
            Set nowSeries = nowChart.SeriesCollection.NewSeries
            nowSeries.XValues = nowSheet.Cells(j + 1, i + 1 + 2)    'Vts
            nowSeries.Values = nowSheet.Cells(j + 1, i + 1) 'Ion
            nowSeries.Name = nowSheet.Cells(j + 1, 1)
        Next j
        tmp = ChartSheet.ChartObjects(i).Chart.ChartTitle.Caption
        tmp = Replace(tmp, "Ion-Ioffs", "Ion-Vts")
        nowChart.HasTitle = True
        nowChart.ChartTitle.Caption = tmp
        nowChart.Legend.AutoScaleFont = False
        nowChart.Legend.Font.Size = 8
        
        tmp2 = getCOL(getCOL(tmp, "_", 1), ",", 1)
        tmp = getCOL(getCOL(tmp, "(", 2), ")", 1)
        tmp = UCase(Mid(tmp, 1, 1)) & Mid(tmp, 2)
        With nowChart.Axes(xlValue)
            .Crosses = xlCustom
            .CrossesAt = .MinimumScale
            .HasTitle = True
            .AxisTitle.Characters.Text = tmp & "Ioffs_" & tmp2 & "_Ion(%)"
            .AxisTitle.AutoScaleFont = False
            .AxisTitle.Font.Size = 8
        End With
        With nowChart.Axes(xlCategory)
            .Crosses = xlCustom
            .CrossesAt = .MinimumScale
            .HasTitle = True
            .AxisTitle.Characters.Text = tmp & "Ioffs_" & tmp2 & "_Vts(V)"
            .AxisTitle.AutoScaleFont = False
            .AxisTitle.Font.Size = 8
        End With
    
    Next i
    
    Set nowTrend = Nothing
    Set nowSeries = Nothing
    Set nowChart = Nothing
    Set nowChart = Nothing
    Set ChartSheet = Nothing
End Sub

Public Function getXfromEquation(ByVal strEq As String, ByVal y As Single)
    Dim a As Single, b As Single, c As Single
    Dim mType As Integer
    
    mType = 1
    If InStr(strEq, "x2") > 0 Then mType = 2
    
    Select Case mType
        Case 1: '直線方程式
            a = getCOL(getCOL(strEq, "=", 2), "x", 1)
            b = getCOL(strEq, "x", 2) - y
            getXfromEquation = -1 * b / a
        Case 2: '二項式方程式
            a = getCOL(getCOL(strEq, "=", 2), "x2", 1)
            b = getCOL(getCOL(strEq, "x2", 2), "x", 1)
            'If Left(b, 1) = "+" Then b = Mid(b, 2)
            c = getCOL(strEq, "x", 3) - y
            'If Left(c, 1) = "+" Then c = Mid(c, 2)
            If (b ^ 2 - 4 * a * c) < 0 Then Err.Raise 999, , "與Target Y 沒有交點"
            '    getXfromEquation = 1
            '    Exit Function
            'End If
            getXfromEquation = (-1 * b + Sqr(b ^ 2 - 4 * a * c)) / (2 * a)  '公式解
    End Select
    
End Function

Public Function Manual_L40LSI()
   Dim nowSheet As Worksheet
   Dim TargetSheet As Worksheet
   Dim i As Long, j As Long
   Dim iRow As Long, iCol As Long
   Dim waferColl As New Collection
   Const bCol = 6
   Const nRow = 3
   
   Set nowSheet = ActiveSheet
   Set TargetSheet = AddSheet("LSI_Temp")
   
   For j = 7 To nowSheet.Columns.Count
      If nowSheet.Cells(1, j) <> i And nowSheet.Cells(1, j) <> "" Then
         waferColl.Add nowSheet.Cells(1, j)
         i = nowSheet.Cells(1, j)
         'Debug.Print i
      End If
   Next j
   
   With TargetSheet
      .Range("A2") = "Split": .Range("A2:C2").Merge: .Range("A2:C2").HorizontalAlignment = xlCenter
      .Range("A3") = "Wafer#": .Range("A3:C3").Merge: .Range("A3:C3").HorizontalAlignment = xlCenter
      .Range("B4") = "Item"
      .Range("C4") = "Target"
      .Range("D4") = "Unit"
      For j = 1 To waferColl.Count
         .Cells(3, bCol + j - 1) = "#" & waferColl(j)
         .Range(N2L(bCol + j - 1) & "3:" & N2L(bCol + j - 1) & "4").Merge
         .Range(N2L(bCol + j - 1) & "3:" & N2L(bCol + j - 1) & "4").HorizontalAlignment = xlCenter
         .Range(N2L(bCol + j - 1) & "3:" & N2L(bCol + j - 1) & "4").VerticalAlignment = xlCenter
      Next j
   End With
   
   iRow = 5
   For i = 3 To nowSheet.UsedRange.Rows.Count
      With TargetSheet
         .Cells(iRow, 2) = nowSheet.Cells(i, 2)
         .Range(.Cells(iRow, 2), .Cells(iRow + nRow, 2)).Merge
         .Range(.Cells(iRow, 2), .Cells(iRow + nRow, 2)).VerticalAlignment = xlCenter
         .Cells(iRow, 3) = nowSheet.Cells(i, 5)
         .Range(.Cells(iRow, 3), .Cells(iRow + nRow, 3)).Merge
         .Range(.Cells(iRow, 3), .Cells(iRow + nRow, 3)).VerticalAlignment = xlCenter
         .Cells(iRow, 4) = nowSheet.Cells(i, 3)
         .Range(.Cells(iRow, 4), .Cells(iRow + nRow, 4)).Merge
         .Range(.Cells(iRow, 4), .Cells(iRow + nRow, 4)).VerticalAlignment = xlCenter
         .Cells(iRow, 5) = "Median"
         .Cells(iRow + 1, 5) = "Gap%"
         .Cells(iRow + 2, 5) = "U%"
         .Cells(iRow + 3, 5) = "3Sigma"
         For j = 1 To waferColl.Count
            .Cells(iRow, 5 + j) = nowSheet.Cells(i, 7 + (j - 1) * 4)
            .Cells(iRow + 1, 5 + j).Formula = "=" & N2L(5 + j) & CStr(iRow) & "/$C" & CStr(iRow) & "-1"
            .Cells(iRow + 2, 5 + j).Formula = "=" & N2L(5 + j) & CStr(iRow + 3) & "/" & N2L(5 + j) & CStr(iRow)
            .Cells(iRow + 3, 5 + j) = nowSheet.Cells(i, 7 + 2 + (j - 1) * 4) * 3
            
            .Cells(iRow + 1, 5 + j).NumberFormat = "0.0%"
            .Cells(iRow + 2, 5 + j).NumberFormat = "0.0%"
            If .Cells(iRow + 3, 5 + j) > 0.01 Then
               .Cells(iRow + 3, 5 + j).NumberFormat = "0.00"
            Else
               .Cells(iRow + 3, 5 + j).NumberFormatLocal = "0.00E+00"
            End If
         Next j
      End With
      iRow = iRow + nRow + 1
   Next i
   
   With TargetSheet
      
      .Columns.AutoFit
   End With
   'nowSheet.Activate
End Function

Public Function Manual_Vincent_AddFailTable()
    Dim nowSheet As Worksheet, setSheet As Worksheet, sumSheet As Worksheet
    Dim iSheet As Integer
    Dim iRow As Long, iCol As Long
    Dim bRow As Long
    'Dim p1_n As Integer, p2_n As Integer
    'Dim p1_p As Integer, p2_p As Integer
    'Dim p1_f1 As Integer, p2_f1 As Integer
    'Dim p1_f2 As Integer, p2_f2 As Integer
    'Dim p1_f3 As Integer, p2_f3 As Integer
    Dim PA(1 To 2, 1 To 7) As Integer
    Dim P As Integer
    Dim i As Integer, j As Integer
    Dim x As Integer, y As Integer
    Dim tmp As Variant
    Dim xRange As Range, yRange As Range
    Dim nowChart As Chart, nowSeries As Series
    Dim nowAxis As Axis
    Dim nowRange As Range
    
    Const cTar As Integer = 5, cSpecLo As Integer = 4, cSpecHi As Integer = 6
    Const cCri As Integer = 11, cPha As Integer = 10
    Const cF1 As Integer = 7, cF2 As Integer = 8, cF3 As Integer = 9
    
    Set sumSheet = AddSheet("Total summary")
    Set setSheet = Worksheets("SPEC_List")
    
    For iSheet = 1 To setSheet.UsedRange.Columns.Count
        If setSheet.Cells(1, iSheet) = "" Then Exit For
        If InStr(setSheet.Cells(1, iSheet), ":") < 1 And IsExistSheet(setSheet.Cells(1, iSheet) & "_Summary") Then
            Set nowSheet = Worksheets(setSheet.Cells(1, iSheet) & "_Summary")
            x = x + 1: y = 0
            Debug.Print x, nowSheet.Name
            
            'Remove non P1 and P2
            For iRow = nowSheet.UsedRange.Rows.Count To 3 Step -1
                If nowSheet.Cells(iRow, cPha) <> "P1" And nowSheet.Cells(iRow, cPha) <> "P2" Then nowSheet.Rows(iRow).Delete
            Next iRow
            
            For iCol = nowSheet.UsedRange.Columns.Count To 5 Step -1
                If nowSheet.Cells(2, iCol) = "K value" Then Exit For
                If nowSheet.Cells(2, iCol) = "Yield" Then
                    nowSheet.Columns(N2L(iCol + 1) & ":" & N2L(iCol + 1)).Insert Shift:=xlToRight
                    nowSheet.Cells(2, iCol + 1) = "K value"
                    With nowSheet.Columns(N2L(iCol + 1) & ":" & N2L(iCol + 1))
                        .Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
                        .Borders(xlEdgeRight).Weight = xlMedium
                        For iRow = 3 To nowSheet.UsedRange.Rows.Count
                            If nowSheet.Cells(iRow, cTar) <> "" And nowSheet.Cells(iRow, cF1) <> "" And nowSheet.Cells(iRow, iCol - 3) <> "" Then
                                With nowSheet.Cells(iRow, iCol + 1)
                                    .Cells(1, 1) = (nowSheet.Cells(iRow, iCol - 3) - nowSheet.Cells(iRow, 5)) / nowSheet.Cells(iRow, cF1)
                                    .NumberFormat = "0.00"
                                    .Font.ColorIndex = 1
                                    Select Case IIf(UCase(Left(nowSheet.Cells(iRow, 2), 3)) = "IOF", .Cells(1, 1), Abs(.Cells(1, 1)))
                                        Case Is < 0.5:   .Font.ColorIndex = 5
                                        Case Is > 1:   .Font.ColorIndex = 3
                                    End Select
                                End With
                            End If
                        Next iRow
                    End With
                End If
            Next iCol
            
            bRow = nowSheet.UsedRange.Rows.Count + 2
            For iCol = 1 To nowSheet.UsedRange.Columns.Count
                If nowSheet.Cells(2, iCol) = "Median" Then
                    For i = 1 To 2
                        For j = 1 To 7
                            PA(i, j) = 0
                        Next j
                    Next i
                    For iRow = 3 To nowSheet.UsedRange.Rows.Count
                        If nowSheet.Cells(iRow, 2) = "" Then Exit For
                        If nowSheet.Cells(iRow, cTar) <> "" And nowSheet.Cells(iRow, cCri) <> "" And nowSheet.Cells(iRow, cPha) <> "" And nowSheet.Cells(iRow, iCol) <> "" Then
                            If nowSheet.Cells(iRow, cPha) = "P1" Then P = 1
                            If nowSheet.Cells(iRow, cPha) = "P2" Then P = 2
                            PA(P, 1) = PA(P, 1) + 1
                            nowSheet.Cells(iRow, iCol).FormatConditions.Delete
                            Select Case IIf(UCase(Left(nowSheet.Cells(iRow, 2), 3)) = "IOF", nowSheet.Cells(iRow, iCol) - nowSheet.Cells(iRow, 5), Abs(nowSheet.Cells(iRow, iCol) - nowSheet.Cells(iRow, 5)))
                                Case Is < nowSheet.Cells(iRow, cF1)
                                    PA(P, 2) = PA(P, 2) + 1
                                    nowSheet.Cells(iRow, iCol).Font.ColorIndex = 1
                                    nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 0
                                Case Is < nowSheet.Cells(iRow, cF2)
                                    PA(P, 3) = PA(P, 3) + 1
                                    If nowSheet.Cells(iRow, iCol) - nowSheet.Cells(iRow, 5) > 0 Then
                                        nowSheet.Cells(iRow, iCol).Font.ColorIndex = 3
                                        'nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 26
                                    Else
                                        nowSheet.Cells(iRow, iCol).Font.ColorIndex = 4
                                        'nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 26
                                    End If
                                Case Is < nowSheet.Cells(iRow, cF3)
                                    PA(P, 4) = PA(P, 4) + 1
                                    If nowSheet.Cells(iRow, iCol) - nowSheet.Cells(iRow, 5) > 0 Then
                                        nowSheet.Cells(iRow, iCol).Font.ColorIndex = 1
                                        nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 26
                                    Else
                                        nowSheet.Cells(iRow, iCol).Font.ColorIndex = 1
                                        nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 35
                                    End If
                                Case Is > nowSheet.Cells(iRow, cF3)
                                    PA(P, 5) = PA(P, 5) + 1
                                    If nowSheet.Cells(iRow, iCol) - nowSheet.Cells(iRow, 5) > 0 Then
                                        nowSheet.Cells(iRow, iCol).Font.ColorIndex = 1
                                        nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 3
                                    Else
                                        nowSheet.Cells(iRow, iCol).Font.ColorIndex = 1
                                        nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 4
                                    End If
                            End Select
                            
                            ' K value
                            Select Case IIf(UCase(Left(nowSheet.Cells(iRow, 2), 3)) = "IOF", nowSheet.Cells(iRow, iCol + 4), Abs(nowSheet.Cells(iRow, iCol + 4)))
                                    Case Is <= 0.5: PA(P, 6) = PA(P, 6) + 1
                                    Case Is > 1: PA(P, 7) = PA(P, 7) + 1
                            End Select
                        End If
                    Next iRow
                    
                    i = 0
                    With nowSheet.Cells(bRow, iCol)
                        i = i + 1: .Cells(i, 1) = "Priority": .Cells(i, 2) = "1": .Cells(i, 3) = "2": .Cells(i, 4) = "Total"
                        i = i + 1: .Cells(i, 1) = "TEST Item": .Cells(i, 2) = PA(1, 1): .Cells(i, 3) = PA(2, 1): .Cells(i, 4) = PA(1, 1) + PA(2, 1)
                        i = i + 1: .Cells(i, 1) = "Pass Item": .Cells(i, 2) = PA(1, 2): .Cells(i, 3) = PA(2, 2): .Cells(i, 4) = PA(1, 2) + PA(2, 2)
                        i = i + 1: .Cells(i, 1) = "|K| < 0.5": .Cells(i, 2) = PA(1, 6): .Cells(i, 3) = PA(2, 6): .Cells(i, 4) = PA(1, 6) + PA(2, 6)
                        i = i + 1: .Cells(i, 1) = "Score_1": If PA(1, 1) <> 0 Then .Cells(i, 2) = Format(PA(1, 2) / PA(1, 1), "0%"): If PA(2, 1) <> 0 Then .Cells(i, 3) = Format(PA(2, 2) / PA(2, 1), "0%"): If (PA(1, 1) + PA(2, 1)) <> 0 Then .Cells(i, 4) = Format((PA(1, 2) + PA(2, 2)) / (PA(1, 1) + PA(2, 1)), "0%")
                        i = i + 1: .Cells(i, 1) = "Score_2": If PA(1, 1) <> 0 Then .Cells(i, 2) = Format((PA(1, 2) + PA(1, 3)) / PA(1, 1), "0%"): If PA(2, 1) <> 0 Then .Cells(i, 3) = Format((PA(2, 2) + PA(2, 3)) / PA(2, 1), "0%"): If (PA(1, 1) + PA(2, 1)) <> 0 Then .Cells(i, 4) = Format((PA(1, 2) + PA(2, 2) + PA(1, 3) + PA(2, 3)) / (PA(1, 1) + PA(2, 1)), "0%")
                        i = i + 1: .Cells(i, 1) = "Score_3": If PA(1, 1) <> 0 Then .Cells(i, 2) = Format((PA(1, 2) + PA(1, 3) + PA(1, 4)) / PA(1, 1), "0%"): If PA(2, 1) <> 0 Then .Cells(i, 3) = Format((PA(2, 2) + PA(2, 3) + PA(2, 4)) / PA(2, 1), "0%"): If (PA(1, 1) + PA(2, 1)) <> 0 Then .Cells(i, 4) = Format((PA(1, 2) + PA(2, 2) + PA(1, 3) + PA(2, 3) + PA(1, 4) + PA(2, 4)) / (PA(1, 1) + PA(2, 1)), "0%")
                        i = i + 1: .Cells(i, 1) = "|K| > 1": .Cells(i, 2) = PA(1, 7): .Cells(i, 3) = PA(2, 7): .Cells(i, 4) = PA(1, 7) + PA(2, 7)
                        'i = i + 1: .Cells(i, 1) = "Fail Item": .Cells(i, 2) = PA(1, 1) - PA(1, 2): .Cells(i, 3) = PA(2, 1) - PA(2, 2): .Cells(i, 4) = PA(1, 1) + PA(2, 1) - PA(1, 2) - PA(2, 2)
                        i = i + 1: .Cells(i, 1) = nowSheet.Cells(2, 7): .Cells(i, 2) = PA(1, 3): .Cells(i, 3) = PA(2, 3): .Cells(i, 4) = PA(1, 3) + PA(2, 3)
                        i = i + 1: .Cells(i, 1) = nowSheet.Cells(2, 8): .Cells(i, 2) = PA(1, 4): .Cells(i, 3) = PA(2, 4): .Cells(i, 4) = PA(1, 4) + PA(2, 4)
                        i = i + 1: .Cells(i, 1) = nowSheet.Cells(2, 9): .Cells(i, 2) = PA(1, 5): .Cells(i, 3) = PA(2, 5): .Cells(i, 4) = PA(1, 5) + PA(2, 5)
                        '調格式
                        .Range("A1:D1").Font.Bold = True
                        .Range("A1:D" & CStr(i)).Borders.Weight = xlThin
                        '.Range("A1:D10").Borders(xlEdgeRight).Weight = xlMedium
                        .Range("B1").Interior.ColorIndex = 3
                        .Range("C1").Interior.ColorIndex = 20
                        .Range("A3:D3").Interior.ColorIndex = 4
                        .Range("A4:D4").Interior.ColorIndex = 41
                        .Range("A8:D8").Interior.ColorIndex = 26
                    
                        y = y + 1
                        .Range("A1:D" & CStr(i)).Copy sumSheet.Range(N2L((x - 1) * 4 + 5) & CStr((y - 1) * (i + 5) + 6))
                        sumSheet.Range(N2L((x - 1) * 4 + 5) & CStr((y - 1) * (i + 5) + 6 - 1)) = setSheet.Cells(1, iSheet)
                        sumSheet.Range(N2L((x - 1) * 4 + 5) & CStr((y - 1) * (i + 5) + 6 - 1) & ":" & N2L((x - 1) * 4 + 5 + 3) & CStr((y - 1) * (i + 5) + 6 - 1)).Merge
                        sumSheet.Range(N2L((x - 1) * 4 + 5) & CStr((y - 1) * (i + 5) + 6 - 1) & ":" & N2L((x - 1) * 4 + 5 + 3) & CStr((y - 1) * (i + 5) + 6 - 1)).Interior.ColorIndex = 36
                        sumSheet.Range(N2L((x - 1) * 4 + 5) & CStr((y - 1) * (i + 5) + 6 - 1) & ":" & N2L((x - 1) * 4 + 5 + 3) & CStr((y - 1) * (i + 5) + 6 - 1)).HorizontalAlignment = xlHAlignCenter
                        sumSheet.Range(N2L((x - 1) * 4 + 5) & CStr((y - 1) * (i + 5) + 6 - 1) & ":" & N2L((x - 1) * 4 + 5 + 3) & CStr((y - 1) * (i + 5) + 6 - 1)).Borders.Weight = xlThin
                        If x = 1 Then
                            .Range("A1:D" & CStr(i)).Copy sumSheet.Range(N2L(1) & CStr((y - 1) * (i + 5) + 6))
                            sumSheet.Range(N2L(1) & CStr((y - 1) * (i + 5) + 6 - 1)) = "total summary"
                            sumSheet.Range(N2L(1) & CStr((y - 1) * (i + 5) + 6 - 1)).HorizontalAlignment = xlHAlignCenter
                            sumSheet.Range(N2L(1) & CStr((y - 1) * (i + 5) + 6 - 1)).Interior.ColorIndex = 36
                            sumSheet.Range(N2L(1) & CStr((y - 1) * (i + 5) + 6 - 1) & ":" & N2L(4) & CStr((y - 1) * (i + 5) + 6 - 1)).Merge
                            With sumSheet.Range(N2L(1) & CStr((y - 1) * (i + 5) + 6 - 4))
                                .Cells(1, 1) = "w/i target and criteria"
                                .Cells(1, 1).Interior.ColorIndex = 27
                                .Range("A1:D1").Merge
                                .Range("B2:D2").Merge
                                .Range("B3:D3").Merge
                                .Range("A1:D1").HorizontalAlignment = xlHAlignCenter
                                .Range("B2:D2").HorizontalAlignment = xlHAlignCenter
                                .Range("B3:D3").HorizontalAlignment = xlHAlignCenter
                                .Cells(2, 1) = "LOT ID"
                                .Cells(2, 2) = nowSheet.Cells(1, 4)
                                .Cells(3, 1) = "Wafer #"
                                .Cells(3, 2) = nowSheet.Cells(1, iCol)
                                .Range("A1:D4").Borders.Weight = xlThin
                            End With
                        End If
                    End With
                End If
            Next iCol
        End If
        
        'Add K chart
        Dim yCount As Integer   'For 2010 , 20130722
        For i = nowSheet.ChartObjects.Count To 1 Step -1
            nowSheet.ChartObjects(i).Delete
        Next i
        'Set nowChart = nowSheet.ChartObjects.Add(10, 100, 600, 400).Chart
        Set nowChart = myCreateChart(nowSheet, xlLineMarkers, 10, 10, 600, 300)
        
        nowChart.chartType = xlLineMarkers
        Set xRange = nowSheet.Range(nowSheet.Cells(3, 2), nowSheet.Cells(3, 2).End(xlDown))
        For iCol = 1 To nowSheet.UsedRange.Columns.Count
            If nowSheet.Cells(2, iCol) = "K value" Then
               'Set yRange = nowSheet.Range(nowSheet.Cells(3, iCol), nowSheet.Cells(3, iCol).End(xlDown))
               Set yRange = nowSheet.Range(nowSheet.Cells(3, iCol), nowSheet.Cells(3 + xRange.Rows.Count - 1, iCol))
               Set nowSeries = nowChart.SeriesCollection.NewSeries
               nowSeries.XValues = xRange
               nowSeries.Values = yRange
               If WorksheetFunction.Count(yRange) <> 0 Then
                 nowSeries.Name = "#" & nowSheet.Cells(1, iCol - 3)
                 nowSeries.Border.LineStyle = xlDot ' xlDash
                 If yCount = 0 Then yCount = WorksheetFunction.Count(yRange)
               End If
            End If
        Next iCol
        Set nowSeries = nowChart.SeriesCollection.NewSeries
        nowSeries.chartType = xlXYScatter
        'nowSeries.XValues = "{1,1}"
        nowSeries.XValues = "{" & yCount / 2 & "," & yCount / 2 & "}"
        nowSeries.Values = "{1,-1}"
        nowSeries.MarkerStyle = xlMarkerStyleDash 'xlMarkerStyleNone
        nowSeries.MarkerForegroundColorIndex = 3
        nowSeries.Name = "Criteria"
        With nowSeries
            .ErrorBar xlX, Include:=xlBoth, Type:=xlFixedValue, Amount:=yCount / 2
            .ErrorBars.Border.ColorIndex = 3
            .ErrorBars.Border.LineStyle = xlDash
            .ErrorBars.Border.Weight = xlMedium
            .ErrorBars.EndStyle = xlNoCap
        End With
        
        With nowChart
            .ChartArea.Interior.ColorIndex = 2
            .PlotArea.Interior.ColorIndex = 2
            .Legend.Interior.ColorIndex = 2
            .HasTitle = True
            .ChartTitle.Characters.Text = nowSheet.Name
            .ChartTitle.Characters.Font.Name = "Lucida Sans Unicode"
        End With
        Set nowAxis = nowChart.Axes(xlCategory)
        With nowAxis
            .TickLabels.Font.Size = 10
            .TickLabels.Font.Name = "Lucida Sans Unicode"
            '.TickLabels.Font.Bold = True
            '.CrossesAt = 1
            .TickLabelSpacing = 1
            .TickMarkSpacing = 1
            '.AxisBetweenCategories = True
            '.ReversePlotOrder = False
            .TickLabels.Orientation = xlUpward
        End With
        Set nowAxis = nowChart.Axes(xlValue)
        With nowAxis
            .HasTitle = True
            .AxisTitle.Characters.Text = "K-Value"
            .AxisTitle.Characters.Font.Name = "Lucida Sans Unicode"
            If .MaximumScale > Abs(.MinimumScale) Then .MinimumScale = -1 * .MaximumScale Else .MaximumScale = -1 * .MinimumScale
            If .MaximumScale > Abs(.MinimumScale) Then .MinimumScale = -1 * .MaximumScale Else .MaximumScale = -1 * .MinimumScale
            If .MaximumScale > Abs(.MinimumScale) Then .MinimumScale = -1 * .MaximumScale Else .MaximumScale = -1 * .MinimumScale
            If .MaximumScale > 10 Then
                .MaximumScale = 10
                .MinimumScale = -10
            End If
            .Crosses = xlCustom
            .CrossesAt = .MinimumScale
            '.ReversePlotOrder = False
            '.ScaleType = xlLinear
            '.DisplayUnit = xlNone
            .TickLabels.Font.Size = 12
            .TickLabels.Font.Name = "Lucida Sans Unicode"
            '.TickLabels.Font.Bold = True
            .TickLabels.NumberFormatLocal = "0_ "
        End With
        '---------------------------------------------------
    Next iSheet
    
    For iRow = 1 To sumSheet.UsedRange.Rows.Count
        If sumSheet.Cells(iRow, 1) = "Priority" Then
            With sumSheet.Cells(iRow + 1, 1 + 1)
                For x = 1 To 3
                    For y = 1 To 10
                        tmp = 0: j = 0
                        For i = x + 4 To sumSheet.UsedRange.Columns.Count Step 4
                            tmp = tmp + .Cells(y, i): j = j + 1
                        Next i
                        .Cells(y, x) = tmp
                    Next y
                    If .Cells(1, x) <> 0 Then
                        .Cells(4, x) = Format(.Cells(2, x) / .Cells(1, x), "0%")
                        .Cells(5, x) = Format((.Cells(2, x) + .Cells(8, x)) / .Cells(1, x), "0%")
                        .Cells(6, x) = Format((.Cells(2, x) + .Cells(8, x) + .Cells(9, x)) / .Cells(1, x), "0%")
                    End If
                Next x
            End With
        End If
    Next iRow
    
    'Stop
    
    i = 0
    iCol = sumSheet.UsedRange.Columns.Count + 2
    With sumSheet.Cells(1, iCol)
        For iRow = 2 To sumSheet.UsedRange.Rows.Count Step 16
            i = i + 1
            .Cells(1, 1 + i) = sumSheet.Cells(iRow + 2, 2)
            For j = 0 To 2
                .Cells(2 + j, 1) = "Score_" & CStr(j + 1)
                .Cells(2 + j, 1 + i) = sumSheet.Cells(iRow + 8 + j, 4)
                .Cells(2 + j, 1 + i).NumberFormatLocal = "0%"
            Next j
        Next iRow
    End With
    
    Set nowRange = sumSheet.Cells(1, iCol + 1).CurrentRegion
    'sumSheet.Cells(1, iCol).Select
    'Set nowChart = sumSheet.ChartObjects.Add(CSng(sumSheet.Range("A:" & N2L(iCol - 1)).Width), 100, 500, 300).Chart
    'For 2010 相容
    Set nowChart = myCreateChart(nowSheet, xlLineMarkers, CSng(sumSheet.Range("A:" & N2L(iCol - 1)).width), 100, 500, 300)

    With nowChart
        .chartType = xlLineMarkers
        .SetSourceData Source:=nowRange, PlotBy:=xlRows
        .HasTitle = True
        .ChartTitle.Characters.Text = "Lot Total Summary Score Chart"
        .PlotArea.Interior.ColorIndex = 2
        .ChartArea.Interior.ColorIndex = 2
        .Legend.Interior.ColorIndex = 2
    End With
    Set nowAxis = nowChart.Axes(xlCategory)
    With nowAxis
        .TickLabels.Font.Size = 10
        .TickLabels.Font.FontStyle = "Lucida Sans Unicode"
        .HasTitle = True
        .AxisTitle.Characters.Text = "Wafer numbers"
        '.TickLabels.Font.Bold = True
        '.CrossesAt = 1
        '.TickLabelSpacing = 1
        '.TickMarkSpacing = 1
        '.AxisBetweenCategories = True
        '.ReversePlotOrder = False
        '.TickLabels.Orientation = xlUpward
    End With
    Set nowAxis = nowChart.Axes(xlValue)
    With nowAxis
        .TickLabels.Font.Size = 10
        .TickLabels.Font.FontStyle = "Lucida Sans Unicode"
        .HasTitle = True
        .AxisTitle.Characters.Text = "Score"
    End With
    
    
    sumSheet.Activate
    ActiveWindow.Zoom = 70
    'Stop
End Function

Public Function Manual_Vincent_PerformanceFit() '2012/05/15 Vincent
    Dim nowSheet As Worksheet
    Dim tarSerise As String
    Dim lineType As String
    Dim xyType As Long
    Dim nowRange As Range
    Dim iRow As Long, iCol As Long, bRow As Long
    Dim i As Long, j As Long
    Dim xRange As Range, yRange As Range
    Dim reArray()
    Dim a, b
    Dim tempX(), tempY()
    Dim tmp As String
    
    Set nowSheet = ActiveSheet
    
    tarSerise = InputBox("Input series name to fit:", "Series Name", "Target")
    If tarSerise = "" Then Exit Function
    lineType = InputBox("Trendline type: (1-Linear, 2-exponent)", "Trendline Type", "1")
    If lineType = "" Then Exit Function
    xyType = MsgBox("Base on X ?", vbYesNo)
    
    iCol = nowSheet.UsedRange.Columns.Count + 2
    Set nowRange = nowSheet.Range(N2L(iCol) & "1")
    With nowRange
        .Cells(1, 2) = IIf(xyType = 6, "X", "Y")
        .Cells(1, 3) = IIf(xyType = 6, "Y", "X")
        .Cells(1, 4) = "%"
        .Cells(2, 1) = tarSerise
        bRow = 2
        For i = 3 To iCol - 2 Step 2
            If nowSheet.Cells(1, i) <> "" And UCase(nowSheet.Cells(1, i)) <> "TARGET" Then
                If nowSheet.Cells(1, i) = tarSerise Then
                    iRow = 2
                Else
                    bRow = bRow + 1
                    iRow = bRow
                End If
                If iRow > 2 Then
                    .Cells(iRow, 1) = nowSheet.Cells(1, i)
                    .Cells(iRow, 2).FormulaLocal = "=" & N2L(iCol + 1) & CStr(2)
                    .Cells(iRow, 4).FormulaLocal = "=" & "(" & N2L(iCol + 2) & CStr(iRow) & "-" & N2L(iCol + 2) & CStr(2) & ")/" & N2L(iCol + 2) & CStr(2)
                    .Cells(iRow, 4).NumberFormatLocal = "00.00%"
                End If
                Set xRange = nowSheet.Range(nowSheet.Cells(3, i), nowSheet.Cells(3, i).End(xlDown))
                Set yRange = nowSheet.Range(nowSheet.Cells(3, i + 1), nowSheet.Cells(3, i + 1).End(xlDown))
                'ReDim tempA(yRange.Rows.Count - 1)
                tempX = xRange.Value
                tempY = yRange.Value
                Debug.Print xRange.Address, yRange.Address
                Select Case lineType
                    Case "1":   'linear
                        reArray = WorksheetFunction.LinEst(yRange, xRange)
                        If xyType = vbYes Then 'Base on X
                            tmp = reArray(1) & "*" & N2L(iCol + 1) & CStr(iRow) & IIf(reArray(2) >= 0, "+", "") & reArray(2)
                            'reArray = WorksheetFunction.LinEst(yRange, xRange)
                        Else
                            tmp = "(" & N2L(iCol + 1) & CStr(iRow) & "-" & reArray(2) & ")/" & reArray(1)  'y=m*x+b => x=(y-b)/m
                            'reArray = WorksheetFunction.LinEst(xRange, yRange)
                        End If
                        '.Cells(iRow, 3).FormulaLocal = "=" & Round(reArray(1), 4) & "*" & N2L(iCol + 1) & CStr(iRow) & IIf(reArray(2) >= 0, "+", "") & Round(reArray(2), 4)
                        .Cells(iRow, 3).FormulaLocal = "=" & tmp
                    Case "2":   'exponent
                        For j = 1 To UBound(tempY): tempY(j, 1) = WorksheetFunction.Ln(tempY(j, 1)): Next j
                        a = WorksheetFunction.Slope(tempY, tempX)
                        b = Exp(WorksheetFunction.Intercept(tempY, tempX))
                        If xyType = vbYes Then 'Base on X
                            tmp = b & "*EXP(" & a & "*" & N2L(iCol + 1) & CStr(iRow) & ")"
                            'For j = 1 To UBound(tempY): tempY(j, 1) = WorksheetFunction.Ln(tempY(j, 1)): Next j
                            'a = WorksheetFunction.Slope(tempY, tempX)
                            'b = Exp(WorksheetFunction.Intercept(tempY, tempX))
                            'reArray = WorksheetFunction.LinEst(tempY, tempX)
                            'a = WorksheetFunction.index(WorksheetFunction.LinEst(WorksheetFunction.Ln(yRange), xRange), 1)
                            'b = WorksheetFunction.Exp(WorksheetFunction.index(WorksheetFunction.LinEst(WorksheetFunction.Ln(yRange), xRange), 2))
                        Else
                            tmp = "(Ln(" & N2L(iCol + 1) & CStr(iRow) & ") - Ln(" & b & "))/" & a
                            'For j = 1 To UBound(tempX): tempX(j, 1) = WorksheetFunction.Ln(tempX(j, 1)): Next j
                            'a = WorksheetFunction.Slope(tempX, tempY)
                            'b = Exp(WorksheetFunction.Intercept(tempX, tempY))
                        End If
                        '.Cells(iRow, 3).FormulaLocal = "=" & reArray(2) & "*EXP(" & reArray(1) & "*" & N2L(iCol + 1) & CStr(iRow) & ")"
                        .Cells(iRow, 3).FormulaLocal = "=" & tmp
                        '.Cells(iRow, 2) = b
                End Select
            End If
            If nowSheet.Cells(1, i) = "" Then i = 255
        Next i
    End With
    
    Set nowRange = nowRange.CurrentRegion
    nowRange.Borders.LineStyle = xlContinuous
End Function


Public Sub Manual_DieSelection()    'Vincent 2012/07/24
    Dim nowRange As Range
    Dim nowSheet As Worksheet
    Dim i As Integer
    Dim iRow As Integer
    Dim DataSheet As Worksheet
    Dim nowWafer As String
    Dim bRow As Integer, waferRow As Integer
    
    Set DataSheet = Worksheets("data")
    Set nowSheet = ActiveSheet
    'Set nowRange = nowSheet.Range("G3:" & N2L(nowSheet.UsedRange.Columns.Count) & CStr(nowSheet.UsedRange.Rows.Count))
    'Debug.Print nowRange.Address
    'Exit Sub
    With nowSheet.Range("G3:" & N2L(nowSheet.UsedRange.Columns.Count) & CStr(nowSheet.UsedRange.Rows.Count))
        .Select
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=$F3"
        .FormatConditions(1).Font.ColorIndex = 3
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=$D3"
        .FormatConditions(2).Font.ColorIndex = 4
    End With
    'Exit Sub
    bRow = 1
    iRow = nowSheet.UsedRange.Rows.Count + 1
    Set nowRange = nowSheet.Cells(nowSheet.UsedRange.Rows.Count + 6, 1)
    With nowRange
        .Cells(1, 1) = "wafer"
        .Cells(1, 2) = "die#"
        .Cells(1, 3) = "pass die"
    End With
    'get row of wafer
    For i = 1 To 20
        If DataSheet.Cells(i, 2) = "Parameter" Then waferRow = i: Exit For
    Next i
    
    For i = 7 To nowSheet.UsedRange.Columns.Count
        nowSheet.Cells(iRow, i).FormulaArray = "=SUM(IF(" & N2L(i) & "3:" & N2L(i) & CStr(iRow - 1) & ">F3:F" & CStr(iRow - 1) & ",1,0))"
        nowSheet.Cells(iRow + 1, i).FormulaArray = "=SUM(IF(" & N2L(i) & "3:" & N2L(i) & CStr(iRow - 1) & "<D3:D" & CStr(iRow - 1) & ",1,0))"
        nowSheet.Cells(iRow + 2, i).Formula = "=SUM(" & N2L(i) & CStr(iRow) & ":" & N2L(i) & CStr(iRow + 1) & ")"
        nowSheet.Cells(iRow + 3, i) = getCOL(DataSheet.Cells(waferRow, 3 + nowSheet.Cells(2, i)), ">", 2)
        If nowWafer <> nowSheet.Cells(1, i) Then
            bRow = bRow + 1
            nowWafer = nowSheet.Cells(1, i)
            nowRange.Cells(bRow, 1) = "#" & nowWafer
        End If
        If nowSheet.Cells(iRow + 2, i) = 0 Then
            nowRange.Cells(bRow, 3) = nowRange.Cells(bRow, 3) & nowSheet.Cells(iRow + 3, i) & " "
            nowRange.Cells(bRow, 2) = nowRange.Cells(bRow, 2) + 1
        End If
    Next i
     
    Call GenDieMap(nowSheet)
    
End Sub

Public Function GenDieMap(nowSheet As Worksheet)    '2013/02/25 for Vincent
    Dim bRow As Long, bCol As Long
    Dim i As Long, j As Long
    Dim nowRange As Range
    Dim xMin As Integer, xMax As Integer, yMin As Integer, yMax As Integer
    Dim iWafer As Integer
    Dim waferRange As Range
    Dim x As Integer, y As Integer
    Dim tmp As String
    
    Set nowRange = nowSheet.Cells(nowSheet.UsedRange.Rows.Count + 2, 1)
    
    For i = 1 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(i, 2) = "die#" Then bRow = i: Exit For
    Next i
    Set waferRange = nowSheet.Cells(bRow, 1).CurrentRegion
    bCol = 7
    
    For i = bCol To nowSheet.UsedRange.Columns.Count
        If nowSheet.Cells(1, i) <> nowSheet.Cells(1, i - 1) And i > bCol Then Exit For
        nowRange.Cells(i - bCol + 1, 1) = getCOL(getCOL(nowSheet.Cells(bRow - 2, i), ",", 1), "(", 2)
        nowRange.Cells(i - bCol + 1, 2) = getCOL(getCOL(nowSheet.Cells(bRow - 2, i), ",", 2), ")", 1)
    Next i
    
    Set nowRange = nowRange.CurrentRegion
    xMin = WorksheetFunction.Min(nowRange.Columns(1))
    xMax = WorksheetFunction.Max(nowRange.Columns(1))
    yMin = WorksheetFunction.Min(nowRange.Columns(2))
    yMax = WorksheetFunction.Max(nowRange.Columns(2))
    'Debug.Print xMin, xMax, yMin, yMax
    
    For iWafer = 1 To waferRange.Rows.Count - 1
        With nowSheet.Cells(bRow + (iWafer - 1) * (yMax - yMin + 3), bCol)
            .Cells(1, 1) = waferRange.Cells(1 + iWafer, 1)
            For x = 1 To xMax - xMin + 1: .Cells(1, 1 + x) = xMin - 1 + x: Next x
            For y = 1 To yMax - yMin + 1: .Cells(1 + y, 1) = yMax + 1 - y: Next y
            For i = 1 To nowRange.Rows.Count
                tmp = "(" & nowRange.Cells(i, 1) & "," & nowRange.Cells(i, 2) & ")"
                If IsKey(waferRange.Cells(1 + iWafer, 3), tmp, " ") Then tmp = "Pass" Else tmp = "Fail"
                .Cells(2 + yMax - nowRange.Cells(i, 2), 2 + nowRange.Cells(i, 1) - xMin) = tmp
                If tmp = "Pass" Then .Cells(2 + yMax - nowRange.Cells(i, 2), 2 + nowRange.Cells(i, 1) - xMin).Interior.ColorIndex = 6
            Next i
        End With
    Next iWafer
    
End Function

Public Sub FixChart()
    Dim nowChart As Chart
    Dim nowSheet As Worksheet
    
    'On Error GoTo myEnd
    
    Set nowSheet = ActiveSheet
    Set nowChart = ActiveChart
    
    If Left(nowSheet.Name, 8) = "BOXTREND" Then
        'Fit 副坐標軸
        With nowChart.Axes(xlValue, xlSecondary)
            .MinimumScale = nowChart.Axes(xlValue, xlPrimary).MinimumScale
            .MaximumScale = nowChart.Axes(xlValue, xlPrimary).MaximumScale
            '.TickLabelPosition = xlNone
            '.MajorTickMark = xlNone
        End With
    End If
    
    Exit Sub
myEnd:
End Sub

Public Sub reCountCorner()

    Dim tArray()
    Dim x As Double, y As Double
    Dim i As Integer, j As Integer
    Dim nowSheet As Worksheet
    Dim oInfo As chartInfo
    Dim iEnd As Long
    Dim inCount As Long, OutCount As Long
    Dim m As Long, n As Long
    Dim nowChart As Chart
    Dim nowShape As Shape
   
    On Error Resume Next
    Set nowSheet = ActiveSheet
    If nowSheet.ChartObjects.Count = 0 Then Exit Sub
    
    inCount = 0: OutCount = 0
    
    For iEnd = 0 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(iEnd + 1, 1) = "" And nowSheet.Cells(iEnd + 1, 2) = "" Then Exit For
    Next iEnd
    
    oInfo = getChartInfo(nowSheet.Range("A1:B" & CStr(iEnd)))
    
    If oInfo.vCornerXValueStr <> "" Then
        ReDim tArray(Len(oInfo.vCornerXValueStr) - Len(Replace(oInfo.vCornerXValueStr, ",", "")))
        For j = 0 To UBound(tArray)
            tArray(j) = Array(Val(getCOL(oInfo.vCornerXValueStr, ",", j + 1)), Val(getCOL(oInfo.vCornerYValueStr, ",", j + 1)))
        Next j
        
        Call CornerSeq(tArray)
        For m = 3 To nowSheet.UsedRange.Columns.Count Step 2
            If IsKey("target,median,tt,ff,ss", nowSheet.Cells(1, m), ",", True) Then Exit For
           
            For n = 3 To nowSheet.UsedRange.Rows.Count
                If nowSheet.Cells(n, m) <> "" And nowSheet.Cells(n, m + 1) <> "" Then
                    If ynInCorner(tArray, nowSheet.Cells(n, m), nowSheet.Cells(n, m + 1)) Then
                        inCount = inCount + 1
                    Else
                        OutCount = OutCount + 1
                    End If
                End If
            Next n
        Next m
    End If

    If nowSheet.ChartObjects.Count > 0 Then
        If inCount > 0 Or OutCount > 0 Then
            Set nowChart = nowSheet.ChartObjects(1).Chart
            For i = nowChart.Shapes.Count To 1 Step -1
                If nowChart.Shapes(i).Type = msoTextBox Then
                    nowChart.Shapes(i).Delete
                End If
            Next i
            Set nowShape = nowChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 100, 20)
            nowSheet.Activate

            nowChart.Parent.Activate
            nowShape.Select
            With Selection
                .Characters.Text = "In: " & CStr(inCount) & " Out: " & CStr(OutCount) & " = " & Format(inCount / (inCount + OutCount), "0.00%")
                .Characters.Font.Size = 12
                .Font.ColorIndex = 3
                .Font.Bold = True
                .AutoSize = True
            End With
            With nowShape
                .Top = nowChart.PlotArea.Top + 12
                .Left = nowChart.PlotArea.Left + 40
            End With
            
            nowChart.ChartArea.Select
        End If
    End If
End Sub


Public Sub PinScatter()

    Dim i As Integer
    
    For i = 1 To Worksheets.Count
        If Left(Worksheets(i).Name, Len("SCATTER")) = "SCATTER" Then
            Worksheets(i).Name = Replace(Worksheets(i).Name, "SCATTER", "!SCATTER")
        End If
        If Left(Worksheets(i).Name, Len("BOXTREND")) = "BOXTREND" Then
            Worksheets(i).Name = Replace(Worksheets(i).Name, "BOXTREND", "!BOXTREND")
        End If
        If Left(Worksheets(i).Name, Len("CUMULATIVE")) = "CUMULATIVE" Then
            Worksheets(i).Name = Replace(Worksheets(i).Name, "CUMULATIVE", "!CUMULATIVE")
        End If
    Next i
    
End Sub

Public Sub UnpinScatter()

    Dim i As Integer
    
    For i = 1 To Worksheets.Count
        If Left(Worksheets(i).Name, Len("!SCATTER")) = "!SCATTER" Then
            Worksheets(i).Name = Replace(Worksheets(i).Name, "!SCATTER", "SCATTER")
        End If
        If Left(Worksheets(i).Name, Len("!BOXTREND")) = "!BOXTREND" Then
            Worksheets(i).Name = Replace(Worksheets(i).Name, "!BOXTREND", "BOXTREND")
        End If
        If Left(Worksheets(i).Name, Len("!CUMULATIVE")) = "!CUMULATIVE" Then
            Worksheets(i).Name = Replace(Worksheets(i).Name, "!CUMULATIVE", "CUMULATIVE")
        End If
    Next i
    
End Sub



Public Sub UpdateSummaryTable()

    Dim nowSheet As Worksheet
    Dim mRange() As Range
    Dim nowRange As Range
    Dim origin As Range
    Dim nowRow As Integer, nowCol As Integer
    Dim iRow As Integer, iCol As Integer
    Dim i As Integer, j As Integer
    
    Call Speed
    
    Set nowSheet = ActiveWorkbook.ActiveSheet
    nowRow = 1: nowCol = 1
    nowSheet.Cells(nowRow, nowCol).Select

    nowSheet.Cells.Find(What:="\BLOCK", _
                        After:=ActiveCell, _
                        LookIn:=xlFormulas, _
                        LookAt:=xlPart, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False, _
                        MatchByte:=False, _
                        SearchFormat:=False).Select
    Set origin = ActiveCell
    ReDim mRange(0) As Range
    Set mRange(0) = origin.CurrentRegion
    Cells.FindNext(After:=ActiveCell).Activate
    
    Do While ActiveCell.row <> origin.row Or ActiveCell.Column <> origin.Column
        ReDim Preserve mRange(UBound(mRange) + 1) As Range
        Set mRange(UBound(mRange)) = ActiveCell.CurrentRegion
        Cells.FindNext(After:=ActiveCell).Activate
    Loop
    
    For i = LBound(mRange) To UBound(mRange)
        Set nowRange = mRange(i)
        GoSub mySub
    Next i
    nowSheet.Cells(1, 1).Select
                        
    Call Unspeed
Exit Sub

mySub:
    Dim srcSheet As Worksheet
    Dim header As String
    Dim srcHeader As Object
    Set srcHeader = CreateObject("Scripting.Dictionary")
    
    iRow = 1: iCol = 2
    Do
        If nowRange.Cells(1, iCol).Value <> "" Then
            If Not IsExistSheet(nowRange.Cells(1, iCol).Value) Then
                MsgBox ("Cannot find " & nowRange.Cells(1, iCol).Value & " worksheet.")
                Exit Sub
            End If
            Set srcSheet = ActiveWorkbook.Worksheets(nowRange.Cells(1, iCol).Value)
            srcHeader.RemoveAll
            For j = 1 To srcSheet.UsedRange.Columns.Count
                If UCase(srcSheet.Cells(2, j).Value) = UCase(nowRange.Cells(2, 1)) Then
                    srcHeader.Add UCase(CStr(srcSheet.Cells(1, j).Value)), j
                ElseIf j < 7 Then
                    srcHeader.Add UCase(CStr(srcSheet.Cells(2, j).Value)), j
                End If
            Next j
            
            header = UCase(CStr(nowRange.Cells(iRow + 1, iCol).Value))
            iRow = iRow + 2
            Do
                If nowRange.Cells(iRow, 1) <> "" And UCase(nowRange.Cells(iRow, 1)) <> "SKIP" Then
                    If srcHeader.Exists(header) Then
                        If srcSheet.Cells(nowRange.Cells(iRow, 1).Value, srcHeader(header)) <> "" Then
                            nowRange.Cells(iRow, iCol).Value = "='" & srcSheet.Name & "'!" & Cells(nowRange.Cells(iRow, 1).Value, srcHeader(header)).Address(False, False)
                            nowRange.Cells(iRow, iCol).NumberFormatLocal = srcSheet.Cells(nowRange.Cells(iRow, 1).Value, srcHeader(header)).NumberFormatLocal
                        Else
                            nowRange.Cells(iRow, iCol).ClearContents
                        End If
                    End If
                ElseIf UCase(nowRange.Cells(iRow, 1)) = "SKIP" Then
                
                Else
                    nowRange.Cells(iRow, iCol).ClearContents
                End If
                iRow = iRow + 1
            Loop Until iRow > nowRange.Rows.Count
            iRow = 1
        ElseIf Not srcSheet Is Nothing And nowRange.Cells(2, iCol).Value <> "" Then
            header = UCase(CStr(nowRange.Cells(iRow + 1, iCol).Value))
            iRow = iRow + 2
            Do
                If nowRange.Cells(iRow, 1) <> "" And UCase(nowRange.Cells(iRow, 1)) <> "SKIP" Then
                    If srcHeader.Exists(header) Then
                        If srcSheet.Cells(nowRange.Cells(iRow, 1).Value, srcHeader(header)) <> "" Then
                            nowRange.Cells(iRow, iCol).Value = "='" & srcSheet.Name & "'!" & Cells(nowRange.Cells(iRow, 1).Value, srcHeader(header)).Address(False, False)
                            nowRange.Cells(iRow, iCol).NumberFormatLocal = srcSheet.Cells(nowRange.Cells(iRow, 1).Value, srcHeader(header)).NumberFormatLocal
                        Else
                            nowRange.Cells(iRow, iCol).ClearContents
                        End If
                    End If
                ElseIf UCase(nowRange.Cells(iRow, 1)) = "SKIP" Then
                
                Else
                    nowRange.Cells(iRow, iCol).ClearContents
                End If
                iRow = iRow + 1
            Loop Until iRow > nowRange.Rows.Count
            iRow = 1
        End If
        iCol = iCol + 1
    Loop Until iCol > nowRange.Columns.Count
    Set nowRange = Nothing
    Set srcSheet = Nothing

Return

End Sub

Public Sub addThruTrendLine()

    Dim nowSheet As Worksheet
    Dim nowChart As Chart
    Dim nowSeries As Series
    
    Set nowSheet = ActiveSheet
    Set nowChart = nowSheet.ChartObjects(1).Chart
    
    Dim i As Integer
    
    For i = nowChart.SeriesCollection.Count To 1 Step -1
        Set nowSeries = nowChart.SeriesCollection(i)
        If nowSeries.Name = "SS" Or nowSeries.Name = "FF" Then
            nowSeries.Delete
        ElseIf nowSeries.Name = "TT" Then
        
        Else
            With nowSeries.Format.Line
                .ForeColor.ObjectThemeColor = msoThemeColorText1
                .Visible = msoTrue
                .ForeColor.RGB = nowSeries.MarkerBackgroundColor
                .Weight = 1.5
            End With
            nowSeries.MarkerForegroundColor = RGB(0, 0, 0)
            nowSeries.MarkerSize = 5
        End If
    Next i

End Sub

Public Sub add3Sigma()

    Dim nowSheet As Worksheet
    Dim nowChart As Chart
    Dim nowLabels As DataLabels
    Dim nowSeries As Series
    
    Set nowSheet = ActiveSheet
    Set nowChart = nowSheet.ChartObjects(1).Chart
    
    Dim threeSigma As Double
    Dim strFormat As String
    
    Dim i As Integer
    For i = nowChart.SeriesCollection.Count To 1 Step -1
        Set nowSeries = nowChart.SeriesCollection(i)
        If nowSeries.Name = "0%" Then Exit For
    Next i
    
    nowSeries.ApplyDataLabels
    Set nowLabels = nowSeries.DataLabels
    nowLabels.Position = xlLabelPositionBelow
    
    For i = 1 To nowLabels.Count
        threeSigma = 3 * WorksheetFunction.StDev(Range(N2L(4 + i) & 12 & ":" & N2L(4 + i) & WorksheetFunction.countA(Columns(4 + i))))
        If threeSigma > 1 Then
            strFormat = "0.00"
        ElseIf threeSigma > 0.01 Then
            strFormat = "0.000"
        ElseIf threeSigma > 0.001 Then
            strFormat = "0.0000"
        Else
            strFormat = "0.0E+00"
        End If
        
        nowLabels(i).Text = "3σ=" & Format(threeSigma, strFormat)
        nowLabels(i).Format.TextFrame2.TextRange.Font.Size = 14
        nowLabels(i).Format.TextFrame2.TextRange.Font.Bold = msoTrue
        nowLabels(i).Format.TextFrame2.TextRange.Font.Name = "Arial"
        nowLabels(i).Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Worksheets("PlotSetup").ChartObject(1).Chart.SeriesCollection(i).Format.Fill.BackColor.RGB
        
    Next i
    
'    Dim AxisX As Axis
'    Set AxisX = nowChart.Axes(xlCategory)
'
'    AxisX.TickLabels.Font.Size = 14
'    AxisX.TickLabels.Font.Bold = msoTrue
    
End Sub

Public Sub genSingleChart()
      
    Call Speed
    Dim waferList() As String
    Dim siteNum As Integer
   
    If Not IsExistSheet("PlotSetup") Then MsgBox "Please check PlotSetup sheet before the operation!!": Exit Sub
    If IsExistSheet("Grouping") Then
        If Not isGroupingSafe Then Exit Sub
    End If
        
    Dim i As Long, j As Long
    Dim nowCol As Integer
    Dim nowRange As Range
    Dim nowSheet As Worksheet
    Dim newSheet As Worksheet
    
    Dim chartType As String
    Dim tmpStr As String
            
    'Get position
    Set nowSheet = ActiveSheet
    If Trim(UCase(nowSheet.Cells(1, Selection.Column).Value)) = "CHART TITLE" Then
        nowCol = Selection.Column
    ElseIf Trim(UCase(nowSheet.Cells(1, Selection.Column - 1).Value)) = "CHART TITLE" Then
        nowCol = Selection.Column - 1
    Else
        Exit Sub
    End If
    
    'Get ChartType
    For i = 1 To nowSheet.Cells(1, nowCol).CurrentRegion.Rows.Count
        If Trim(UCase(Left(nowSheet.Cells(i, nowCol).Value, 5))) = "GRAPH" Then
            chartType = "BOXTREND"
            Exit For
        ElseIf Trim(UCase(Left(nowSheet.Cells(i, nowCol).Value, 6))) = "METHOD" Then
            chartType = "CUMULATIVE"
            Exit For
        ElseIf Trim(UCase(Left(nowSheet.Cells(i, nowCol).Value, 6))) = "CORNER" Then
            chartType = "SCATTER"
            Exit For
        End If
    Next i
    If chartType = "" Then Exit Sub
    
    Set nowRange = Range(nowSheet.Cells(1, nowCol), nowSheet.Cells(nowSheet.UsedRange.Rows.Count, nowCol + 1))
    nowRange.ClearFormats
    
    j = 1
    
    Do While IsExistSheet(chartType & "_" & N2L(j))
        j = j + 1
    Loop
    
    Call PinScatter
    
    Set newSheet = AddSheet(chartType & "_" & N2L(j), , nowSheet.Name)
    nowRange.Copy
    newSheet.Range("A1").PasteSpecial xlPasteValues
    newSheet.Cells.ClearFormats
    
    Call GetWaferList(dSheet, waferList)
    siteNum = getSiteNum(dSheet)
    
    Select Case chartType
        Case "SCATTER"
            Call GenScatter(waferList, siteNum)
            Call PlotUniversalChart(newSheet.Name)
        Case "BOXTREND"
            Call GenBoxTrend(waferList, siteNum)
            Call prepareBoxTrendData(newSheet.Name)
            Call PlotBoxTrendChart(newSheet.Name)
        Case "CUMULATIVE"
            Call GenCumulative(waferList, siteNum)
            Call PlotCumulativeChart(newSheet.Name)
    End Select
    Call adjustChartObject(newSheet)
    
    Call FitSingleChart(newSheet)
    Call reCountCorner
    Call RawdataRange
    Call UnpinScatter
    Call Unspeed
    
    
End Sub

Public Sub updateChartSetting()

    Dim nowSheet As Worksheet
    Dim chartType As String
    
    Set nowSheet = ActiveSheet
    
    Dim i As Integer, j As Integer
    
    If Not UCase(nowSheet.Cells(1, 1).Value) = "CHART TITLE" Then Exit Sub
    
    For i = 1 To nowSheet.UsedRange.Rows.Count
        
        If Left(UCase(nowSheet.Cells(i, 1).Value), 5) = "GRAPH" Then
            chartType = "Boxtrend"
            Exit For
        ElseIf Left(UCase(nowSheet.Cells(i, 1).Value), 6) = "METHOD" Then
            chartType = "AccumulativeChart"
            Exit For
        ElseIf Left(UCase(nowSheet.Cells(i, 1).Value), 6) = "CORNER" Then
            chartType = "UniversalCurve"
            Exit For
        End If
        
    Next i
    
    For i = 3 To nowSheet.Cells(1, 1).CurrentRegion.Columns.Count Step 2
        
        Select Case chartType
            Case "Boxtrend"
                nowSheet.Cells(21, i + 1) = Round(getSPEC(nowSheet.Cells(23, i).Value, "TT"), 3)
                
            Case "UniversalCurve"
                If IsNumeric(nowSheet.Cells(21, i).Value) Then
                    Dim cnt As Integer
                    Dim startRow As Integer
                    cnt = 0
                    For j = 21 To nowSheet.UsedRange.Rows.Count
                        If nowSheet.Cells(j, i).Value = "TT" Then startRow = j: Exit For
                        If Not nowSheet.Cells(j, i).Value = "" Then cnt = cnt + 1
                    Next j
                    If startRow = 0 then Exit For
                    For j = 1 To cnt
                        nowSheet.Cells(startRow + j, i + 1).Value = Round(getSPEC(nowSheet.Cells(20 + j, i + 1).Value, "TT"), 3)
                    Next j
                ElseIf Left(getCOL(nowSheet.Cells(21, i).Value, "_", 1), Len(getCOL(nowSheet.Cells(21, i).Value, "_", 1)) - 1) = Left(getCOL(nowSheet.Cells(21, i + 1).Value, "_", 1), Len(getCOL(nowSheet.Cells(21, i + 1).Value, "_", 1)) - 1) And _
                     Right(getCOL(nowSheet.Cells(21, i).Value, "_", 1), 1) <> Right(getCOL(nowSheet.Cells(21, i + 1).Value, "_", 1), 1) Then
                    nowSheet.Cells(16, i + 1).Value = Round(getSPEC(nowSheet.Cells(21, i).Value, "TT"), 3)
                    nowSheet.Cells(17, i + 1).Value = Round(getSPEC(nowSheet.Cells(21, i + 1).Value, "TT"), 3)
                    nowSheet.Cells(18, i + 1).Value = getCORNER(nowSheet.Cells(21, i).Value)
                    nowSheet.Cells(19, i + 1).Value = getCORNER(nowSheet.Cells(21, i + 1).Value)
                
                Else
                    Dim tmpStrX As String
                    Dim tmpStrY As String
                    For j = 21 To nowSheet.UsedRange.Rows.Count
                        If nowSheet.Cells(j, i + 1).Value = "" Then Exit For
                        If getSPEC(nowSheet.Cells(j, i).Value, "TT") <> 0 Then tmpStrX = tmpStrX & ", " & Round(getSPEC(nowSheet.Cells(j, i).Value, "TT"), 3)
                        If getSPEC(nowSheet.Cells(j, i + 1).Value, "TT") <> 0 Then tmpStrY = tmpStrY & ", " & Round(getSPEC(nowSheet.Cells(j, i + 1).Value, "TT"), 3)
                    Next j
                    tmpStrX = Mid(tmpStrX, 3)
                    tmpStrY = Mid(tmpStrY, 3)
                    nowSheet.Cells(16, i + 1).Value = tmpStrX
                    nowSheet.Cells(17, i + 1).Value = tmpStrY
                    tmpStrX = ""
                    tmpStrY = ""
                End If
                
            Case Else
                
        End Select
    Next i

End Sub

'For 吳泰慶

Public Sub Manual_IonChart()
    Dim ChartSheet As Worksheet, nowSheet As Worksheet
    Dim i As Long, j As Long
    Dim nowChart As Chart
    Dim nowSeries As Series
    Dim nowTrend As Trendline
    Dim SeriesNum As Integer
    Dim TTx As Single, TTy As Single
    Dim tmp As String, tmpValue As Single
    Dim tmp2 As String
    
    If Not IsExistSheet("All_Chart") Then Exit Sub
    Set ChartSheet = Worksheets("All_Chart")
    Set nowSheet = AddSheet("Ion_Chart")
    
    SeriesNum = 0
    For i = 1 To ChartSheet.ChartObjects.Count
        'Debug.Print chartSheet.ChartObjects(i).Name
        Set nowChart = ChartSheet.ChartObjects(i).Chart
        'Chart Name
        nowSheet.Cells(1, i + 1) = ChartSheet.ChartObjects(i).Name
        'Get Trendline Equation String
        '-------------------------------
        For j = 1 To nowChart.SeriesCollection.Count
            Set nowSeries = nowChart.SeriesCollection(j)
            'Debug.Print nowSeries.Name
            If i = 1 Then nowSheet.Cells(j + 1, 1) = nowSeries.Name
            If nowSeries.Name = "TT" Or nowSeries.Name = "FF" Then
                nowSheet.Cells(j + 1, i + 1) = WorksheetFunction.index(nowSeries.XValues, 1) & "," & WorksheetFunction.index(nowSeries.Values, 1)
                If nowSeries.Name = "TT" Then
                    TTx = WorksheetFunction.index(nowSeries.XValues, 1)
                    TTy = WorksheetFunction.index(nowSeries.Values, 1)
                    'nowSheet.Cells(j + 1, i + 1) = TTx & "," & TTy
                End If
                If SeriesNum = 0 Then SeriesNum = j - 1
            Else
                '2010/06/17 RVN/HVN and Ion-Ioffd 用二項式, RVP/HVP用直線
                'If UCase(Right(getCOL(nowChart.ChartTitle.Characters.Text, "_", 1), 1)) = "N" Then
                If UCase(getCOL(nowChart.ChartTitle.Characters.Text, "_", 1)) = "RVN+HVN" And i Mod 3 = 1 Then
                    Set nowTrend = nowSeries.Trendlines.Add(Type:=xlPolynomial, order:=2 _
                                    , Forward:=0, Backward:=0, DisplayEquation:=True, DisplayRSquared:=False)
                Else
                    Set nowTrend = nowSeries.Trendlines.Add(Type:=xlLinear, Forward:=0 _
                                    , Backward:=0, DisplayEquation:=True, DisplayRSquared:=False)
                End If
                nowTrend.DataLabel.NumberFormatLocal = "0.0000000000E+00"
                nowSheet.Cells(j + 1, i + 1) = nowTrend.DataLabel.Caption
                nowTrend.Delete
                'Debug.Print nowChart.ChartTitle.Characters.Text, nowSeries.Name, nowSheet.Cells(j + 1, i + 1)
            End If
        Next j
        'Exit Sub
        ' Calc x value with TT Y value
        '---------------------------------
        For j = 1 To SeriesNum
'            tmp = getCOL(nowSheet.Cells(j + 1, i + 1), "=", 2)
'            tmp = "(" & CStr(TTy) & "-(" & getCOL(tmp, "x", 2) & "))/" & getCOL(tmp, "x", 1)
'            tmpValue = Application.Evaluate(tmp)
            tmpValue = getXfromEquation(nowSheet.Cells(j + 1, i + 1), TTy)
            If i Mod 3 = 1 Then
                nowSheet.Cells(j + 1, i + 1) = tmpValue / TTx - 1   'Ion 取差值
            Else
                nowSheet.Cells(j + 1, i + 1) = tmpValue
            End If
            'Debug.Print nowSheet.Cells(j + 1, i + 1), tmp, tmpValue
        Next j
    Next i
    
    'Generate Charts
    '-------------------
    For i = 1 To ChartSheet.ChartObjects.Count Step 3
        'Ion Vs Cgd
        '--------------------
        Set nowChart = nowSheet.ChartObjects.Add(10, 10 + (i \ 3) * 210, 300, 200).Chart
        nowChart.chartType = xlXYScatter
        For j = 1 To SeriesNum
            Set nowSeries = nowChart.SeriesCollection.NewSeries
            nowSeries.XValues = nowSheet.Cells(j + 1, i + 1 + 1)    'Cgd
            nowSeries.Values = nowSheet.Cells(j + 1, i + 1) 'Ion
            nowSeries.Name = nowSheet.Cells(j + 1, 1)
        Next j
        tmp = ChartSheet.ChartObjects(i).Chart.ChartTitle.Caption
        tmp = Replace(tmp, "Ion-Ioffs", "Ion-Cgd")
        nowChart.HasTitle = True
        nowChart.ChartTitle.Caption = tmp
        nowChart.Legend.AutoScaleFont = False
        nowChart.Legend.Font.Size = 8
        
        tmp2 = getCOL(getCOL(tmp, "_", 1), ",", 1)
        tmp = getCOL(getCOL(tmp, "(", 2), ")", 1)
        tmp = UCase(Mid(tmp, 1, 1)) & Mid(tmp, 2)
        With nowChart.Axes(xlValue)
            .Crosses = xlCustom
            .CrossesAt = .MinimumScale
            .HasTitle = True
            .AxisTitle.Characters.Text = tmp & "Ioffs_" & tmp2 & "_Ion(%)"
            .AxisTitle.AutoScaleFont = False
            .AxisTitle.Font.Size = 8
        End With
        With nowChart.Axes(xlCategory)
            .Crosses = xlCustom
            .CrossesAt = .MinimumScale
            .HasTitle = True
            .AxisTitle.Characters.Text = tmp & "Ioffs_" & tmp2 & "_Cgd(fF/um)"
            .AxisTitle.AutoScaleFont = False
            .AxisTitle.Font.Size = 8
        End With
    
        'Ion Vs Vts
        '--------------------
        Set nowChart = nowSheet.ChartObjects.Add(10 + 310, 10 + (i \ 3) * 210, 300, 200).Chart
        nowChart.chartType = xlXYScatter
        For j = 1 To SeriesNum
            Set nowSeries = nowChart.SeriesCollection.NewSeries
            nowSeries.XValues = nowSheet.Cells(j + 1, i + 1 + 2)    'Vts
            nowSeries.Values = nowSheet.Cells(j + 1, i + 1) 'Ion
            nowSeries.Name = nowSheet.Cells(j + 1, 1)
        Next j
        tmp = ChartSheet.ChartObjects(i).Chart.ChartTitle.Caption
        tmp = Replace(tmp, "Ion-Ioffs", "Ion-Vts")
        nowChart.HasTitle = True
        nowChart.ChartTitle.Caption = tmp
        nowChart.Legend.AutoScaleFont = False
        nowChart.Legend.Font.Size = 8
        
        tmp2 = getCOL(getCOL(tmp, "_", 1), ",", 1)
        tmp = getCOL(getCOL(tmp, "(", 2), ")", 1)
        tmp = UCase(Mid(tmp, 1, 1)) & Mid(tmp, 2)
        With nowChart.Axes(xlValue)
            .Crosses = xlCustom
            .CrossesAt = .MinimumScale
            .HasTitle = True
            .AxisTitle.Characters.Text = tmp & "Ioffs_" & tmp2 & "_Ion(%)"
            .AxisTitle.AutoScaleFont = False
            .AxisTitle.Font.Size = 8
        End With
        With nowChart.Axes(xlCategory)
            .Crosses = xlCustom
            .CrossesAt = .MinimumScale
            .HasTitle = True
            .AxisTitle.Characters.Text = tmp & "Ioffs_" & tmp2 & "_Vts(V)"
            .AxisTitle.AutoScaleFont = False
            .AxisTitle.Font.Size = 8
        End With
    
    Next i
    
    Set nowTrend = Nothing
    Set nowSeries = Nothing
    Set nowChart = Nothing
    Set nowChart = Nothing
    Set ChartSheet = Nothing
End Sub

Public Function getXfromEquation(ByVal strEq As String, ByVal y As Single)
    Dim a As Single, b As Single, c As Single
    Dim mType As Integer
    
    mType = 1
    If InStr(strEq, "x2") > 0 Then mType = 2
    
    Select Case mType
        Case 1: '直線方程式
            a = getCOL(getCOL(strEq, "=", 2), "x", 1)
            b = getCOL(strEq, "x", 2) - y
            getXfromEquation = -1 * b / a
        Case 2: '二項式方程式
            a = getCOL(getCOL(strEq, "=", 2), "x2", 1)
            b = getCOL(getCOL(strEq, "x2", 2), "x", 1)
            'If Left(b, 1) = "+" Then b = Mid(b, 2)
            c = getCOL(strEq, "x", 3) - y
            'If Left(c, 1) = "+" Then c = Mid(c, 2)
            If (b ^ 2 - 4 * a * c) < 0 Then Err.Raise 999, , "與Target Y 沒有交點"
            '    getXfromEquation = 1
            '    Exit Function
            'End If
            getXfromEquation = (-1 * b + Sqr(b ^ 2 - 4 * a * c)) / (2 * a)  '公式解
    End Select
    
End Function

Public Function Manual_L40LSI()
   Dim nowSheet As Worksheet
   Dim TargetSheet As Worksheet
   Dim i As Long, j As Long
   Dim iRow As Long, iCol As Long
   Dim waferColl As New Collection
   Const bCol = 6
   Const nRow = 3
   
   Set nowSheet = ActiveSheet
   Set TargetSheet = AddSheet("LSI_Temp")
   
   For j = 7 To nowSheet.Columns.Count
      If nowSheet.Cells(1, j) <> i And nowSheet.Cells(1, j) <> "" Then
         waferColl.Add nowSheet.Cells(1, j)
         i = nowSheet.Cells(1, j)
         'Debug.Print i
      End If
   Next j
   
   With TargetSheet
      .Range("A2") = "Split": .Range("A2:C2").Merge: .Range("A2:C2").HorizontalAlignment = xlCenter
      .Range("A3") = "Wafer#": .Range("A3:C3").Merge: .Range("A3:C3").HorizontalAlignment = xlCenter
      .Range("B4") = "Item"
      .Range("C4") = "Target"
      .Range("D4") = "Unit"
      For j = 1 To waferColl.Count
         .Cells(3, bCol + j - 1) = "#" & waferColl(j)
         .Range(N2L(bCol + j - 1) & "3:" & N2L(bCol + j - 1) & "4").Merge
         .Range(N2L(bCol + j - 1) & "3:" & N2L(bCol + j - 1) & "4").HorizontalAlignment = xlCenter
         .Range(N2L(bCol + j - 1) & "3:" & N2L(bCol + j - 1) & "4").VerticalAlignment = xlCenter
      Next j
   End With
   
   iRow = 5
   For i = 3 To nowSheet.UsedRange.Rows.Count
      With TargetSheet
         .Cells(iRow, 2) = nowSheet.Cells(i, 2)
         .Range(.Cells(iRow, 2), .Cells(iRow + nRow, 2)).Merge
         .Range(.Cells(iRow, 2), .Cells(iRow + nRow, 2)).VerticalAlignment = xlCenter
         .Cells(iRow, 3) = nowSheet.Cells(i, 5)
         .Range(.Cells(iRow, 3), .Cells(iRow + nRow, 3)).Merge
         .Range(.Cells(iRow, 3), .Cells(iRow + nRow, 3)).VerticalAlignment = xlCenter
         .Cells(iRow, 4) = nowSheet.Cells(i, 3)
         .Range(.Cells(iRow, 4), .Cells(iRow + nRow, 4)).Merge
         .Range(.Cells(iRow, 4), .Cells(iRow + nRow, 4)).VerticalAlignment = xlCenter
         .Cells(iRow, 5) = "Median"
         .Cells(iRow + 1, 5) = "Gap%"
         .Cells(iRow + 2, 5) = "U%"
         .Cells(iRow + 3, 5) = "3Sigma"
         For j = 1 To waferColl.Count
            .Cells(iRow, 5 + j) = nowSheet.Cells(i, 7 + (j - 1) * 4)
            .Cells(iRow + 1, 5 + j).Formula = "=" & N2L(5 + j) & CStr(iRow) & "/$C" & CStr(iRow) & "-1"
            .Cells(iRow + 2, 5 + j).Formula = "=" & N2L(5 + j) & CStr(iRow + 3) & "/" & N2L(5 + j) & CStr(iRow)
            .Cells(iRow + 3, 5 + j) = nowSheet.Cells(i, 7 + 2 + (j - 1) * 4) * 3
            
            .Cells(iRow + 1, 5 + j).NumberFormat = "0.0%"
            .Cells(iRow + 2, 5 + j).NumberFormat = "0.0%"
            If .Cells(iRow + 3, 5 + j) > 0.01 Then
               .Cells(iRow + 3, 5 + j).NumberFormat = "0.00"
            Else
               .Cells(iRow + 3, 5 + j).NumberFormatLocal = "0.00E+00"
            End If
         Next j
      End With
      iRow = iRow + nRow + 1
   Next i
   
   With TargetSheet
      
      .Columns.AutoFit
   End With
   'nowSheet.Activate
End Function

Public Function Manual_Vincent_AddFailTable()
    Dim nowSheet As Worksheet, setSheet As Worksheet, sumSheet As Worksheet
    Dim iSheet As Integer
    Dim iRow As Long, iCol As Long
    Dim bRow As Long
    'Dim p1_n As Integer, p2_n As Integer
    'Dim p1_p As Integer, p2_p As Integer
    'Dim p1_f1 As Integer, p2_f1 As Integer
    'Dim p1_f2 As Integer, p2_f2 As Integer
    'Dim p1_f3 As Integer, p2_f3 As Integer
    Dim PA(1 To 2, 1 To 7) As Integer
    Dim P As Integer
    Dim i As Integer, j As Integer
    Dim x As Integer, y As Integer
    Dim tmp As Variant
    Dim xRange As Range, yRange As Range
    Dim nowChart As Chart, nowSeries As Series
    Dim nowAxis As Axis
    Dim nowRange As Range
    
    Const cTar As Integer = 5, cSpecLo As Integer = 4, cSpecHi As Integer = 6
    Const cCri As Integer = 11, cPha As Integer = 10
    Const cF1 As Integer = 7, cF2 As Integer = 8, cF3 As Integer = 9
    
    Set sumSheet = AddSheet("Total summary")
    Set setSheet = Worksheets("SPEC_List")
    
    For iSheet = 1 To setSheet.UsedRange.Columns.Count
        If setSheet.Cells(1, iSheet) = "" Then Exit For
        If InStr(setSheet.Cells(1, iSheet), ":") < 1 And IsExistSheet(setSheet.Cells(1, iSheet) & "_Summary") Then
            Set nowSheet = Worksheets(setSheet.Cells(1, iSheet) & "_Summary")
            x = x + 1: y = 0
            Debug.Print x, nowSheet.Name
            
            'Remove non P1 and P2
            For iRow = nowSheet.UsedRange.Rows.Count To 3 Step -1
                If nowSheet.Cells(iRow, cPha) <> "P1" And nowSheet.Cells(iRow, cPha) <> "P2" Then nowSheet.Rows(iRow).Delete
            Next iRow
            
            For iCol = nowSheet.UsedRange.Columns.Count To 5 Step -1
                If nowSheet.Cells(2, iCol) = "K value" Then Exit For
                If nowSheet.Cells(2, iCol) = "Yield" Then
                    nowSheet.Columns(N2L(iCol + 1) & ":" & N2L(iCol + 1)).Insert Shift:=xlToRight
                    nowSheet.Cells(2, iCol + 1) = "K value"
                    With nowSheet.Columns(N2L(iCol + 1) & ":" & N2L(iCol + 1))
                        .Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
                        .Borders(xlEdgeRight).Weight = xlMedium
                        For iRow = 3 To nowSheet.UsedRange.Rows.Count
                            If nowSheet.Cells(iRow, cTar) <> "" And nowSheet.Cells(iRow, cF1) <> "" And nowSheet.Cells(iRow, iCol - 3) <> "" Then
                                With nowSheet.Cells(iRow, iCol + 1)
                                    .Cells(1, 1) = (nowSheet.Cells(iRow, iCol - 3) - nowSheet.Cells(iRow, 5)) / nowSheet.Cells(iRow, cF1)
                                    .NumberFormat = "0.00"
                                    .Font.ColorIndex = 1
                                    Select Case IIf(UCase(Left(nowSheet.Cells(iRow, 2), 3)) = "IOF", .Cells(1, 1), Abs(.Cells(1, 1)))
                                        Case Is < 0.5:   .Font.ColorIndex = 5
                                        Case Is > 1:   .Font.ColorIndex = 3
                                    End Select
                                End With
                            End If
                        Next iRow
                    End With
                End If
            Next iCol
            
            bRow = nowSheet.UsedRange.Rows.Count + 2
            For iCol = 1 To nowSheet.UsedRange.Columns.Count
                If nowSheet.Cells(2, iCol) = "Median" Then
                    For i = 1 To 2
                        For j = 1 To 7
                            PA(i, j) = 0
                        Next j
                    Next i
                    For iRow = 3 To nowSheet.UsedRange.Rows.Count
                        If nowSheet.Cells(iRow, 2) = "" Then Exit For
                        If nowSheet.Cells(iRow, cTar) <> "" And nowSheet.Cells(iRow, cCri) <> "" And nowSheet.Cells(iRow, cPha) <> "" And nowSheet.Cells(iRow, iCol) <> "" Then
                            If nowSheet.Cells(iRow, cPha) = "P1" Then P = 1
                            If nowSheet.Cells(iRow, cPha) = "P2" Then P = 2
                            PA(P, 1) = PA(P, 1) + 1
                            nowSheet.Cells(iRow, iCol).FormatConditions.Delete
                            Select Case IIf(UCase(Left(nowSheet.Cells(iRow, 2), 3)) = "IOF", nowSheet.Cells(iRow, iCol) - nowSheet.Cells(iRow, 5), Abs(nowSheet.Cells(iRow, iCol) - nowSheet.Cells(iRow, 5)))
                                Case Is < nowSheet.Cells(iRow, cF1)
                                    PA(P, 2) = PA(P, 2) + 1
                                    nowSheet.Cells(iRow, iCol).Font.ColorIndex = 1
                                    nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 0
                                Case Is < nowSheet.Cells(iRow, cF2)
                                    PA(P, 3) = PA(P, 3) + 1
                                    If nowSheet.Cells(iRow, iCol) - nowSheet.Cells(iRow, 5) > 0 Then
                                        nowSheet.Cells(iRow, iCol).Font.ColorIndex = 3
                                        'nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 26
                                    Else
                                        nowSheet.Cells(iRow, iCol).Font.ColorIndex = 4
                                        'nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 26
                                    End If
                                Case Is < nowSheet.Cells(iRow, cF3)
                                    PA(P, 4) = PA(P, 4) + 1
                                    If nowSheet.Cells(iRow, iCol) - nowSheet.Cells(iRow, 5) > 0 Then
                                        nowSheet.Cells(iRow, iCol).Font.ColorIndex = 1
                                        nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 26
                                    Else
                                        nowSheet.Cells(iRow, iCol).Font.ColorIndex = 1
                                        nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 35
                                    End If
                                Case Is > nowSheet.Cells(iRow, cF3)
                                    PA(P, 5) = PA(P, 5) + 1
                                    If nowSheet.Cells(iRow, iCol) - nowSheet.Cells(iRow, 5) > 0 Then
                                        nowSheet.Cells(iRow, iCol).Font.ColorIndex = 1
                                        nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 3
                                    Else
                                        nowSheet.Cells(iRow, iCol).Font.ColorIndex = 1
                                        nowSheet.Cells(iRow, iCol).Interior.ColorIndex = 4
                                    End If
                            End Select
                            
                            ' K value
                            Select Case IIf(UCase(Left(nowSheet.Cells(iRow, 2), 3)) = "IOF", nowSheet.Cells(iRow, iCol + 4), Abs(nowSheet.Cells(iRow, iCol + 4)))
                                    Case Is <= 0.5: PA(P, 6) = PA(P, 6) + 1
                                    Case Is > 1: PA(P, 7) = PA(P, 7) + 1
                            End Select
                        End If
                    Next iRow
                    
                    i = 0
                    With nowSheet.Cells(bRow, iCol)
                        i = i + 1: .Cells(i, 1) = "Priority": .Cells(i, 2) = "1": .Cells(i, 3) = "2": .Cells(i, 4) = "Total"
                        i = i + 1: .Cells(i, 1) = "TEST Item": .Cells(i, 2) = PA(1, 1): .Cells(i, 3) = PA(2, 1): .Cells(i, 4) = PA(1, 1) + PA(2, 1)
                        i = i + 1: .Cells(i, 1) = "Pass Item": .Cells(i, 2) = PA(1, 2): .Cells(i, 3) = PA(2, 2): .Cells(i, 4) = PA(1, 2) + PA(2, 2)
                        i = i + 1: .Cells(i, 1) = "|K| < 0.5": .Cells(i, 2) = PA(1, 6): .Cells(i, 3) = PA(2, 6): .Cells(i, 4) = PA(1, 6) + PA(2, 6)
                        i = i + 1: .Cells(i, 1) = "Score_1": If PA(1, 1) <> 0 Then .Cells(i, 2) = Format(PA(1, 2) / PA(1, 1), "0%"): If PA(2, 1) <> 0 Then .Cells(i, 3) = Format(PA(2, 2) / PA(2, 1), "0%"): If (PA(1, 1) + PA(2, 1)) <> 0 Then .Cells(i, 4) = Format((PA(1, 2) + PA(2, 2)) / (PA(1, 1) + PA(2, 1)), "0%")
                        i = i + 1: .Cells(i, 1) = "Score_2": If PA(1, 1) <> 0 Then .Cells(i, 2) = Format((PA(1, 2) + PA(1, 3)) / PA(1, 1), "0%"): If PA(2, 1) <> 0 Then .Cells(i, 3) = Format((PA(2, 2) + PA(2, 3)) / PA(2, 1), "0%"): If (PA(1, 1) + PA(2, 1)) <> 0 Then .Cells(i, 4) = Format((PA(1, 2) + PA(2, 2) + PA(1, 3) + PA(2, 3)) / (PA(1, 1) + PA(2, 1)), "0%")
                        i = i + 1: .Cells(i, 1) = "Score_3": If PA(1, 1) <> 0 Then .Cells(i, 2) = Format((PA(1, 2) + PA(1, 3) + PA(1, 4)) / PA(1, 1), "0%"): If PA(2, 1) <> 0 Then .Cells(i, 3) = Format((PA(2, 2) + PA(2, 3) + PA(2, 4)) / PA(2, 1), "0%"): If (PA(1, 1) + PA(2, 1)) <> 0 Then .Cells(i, 4) = Format((PA(1, 2) + PA(2, 2) + PA(1, 3) + PA(2, 3) + PA(1, 4) + PA(2, 4)) / (PA(1, 1) + PA(2, 1)), "0%")
                        i = i + 1: .Cells(i, 1) = "|K| > 1": .Cells(i, 2) = PA(1, 7): .Cells(i, 3) = PA(2, 7): .Cells(i, 4) = PA(1, 7) + PA(2, 7)
                        'i = i + 1: .Cells(i, 1) = "Fail Item": .Cells(i, 2) = PA(1, 1) - PA(1, 2): .Cells(i, 3) = PA(2, 1) - PA(2, 2): .Cells(i, 4) = PA(1, 1) + PA(2, 1) - PA(1, 2) - PA(2, 2)
                        i = i + 1: .Cells(i, 1) = nowSheet.Cells(2, 7): .Cells(i, 2) = PA(1, 3): .Cells(i, 3) = PA(2, 3): .Cells(i, 4) = PA(1, 3) + PA(2, 3)
                        i = i + 1: .Cells(i, 1) = nowSheet.Cells(2, 8): .Cells(i, 2) = PA(1, 4): .Cells(i, 3) = PA(2, 4): .Cells(i, 4) = PA(1, 4) + PA(2, 4)
                        i = i + 1: .Cells(i, 1) = nowSheet.Cells(2, 9): .Cells(i, 2) = PA(1, 5): .Cells(i, 3) = PA(2, 5): .Cells(i, 4) = PA(1, 5) + PA(2, 5)
                        '調格式
                        .Range("A1:D1").Font.Bold = True
                        .Range("A1:D" & CStr(i)).Borders.Weight = xlThin
                        '.Range("A1:D10").Borders(xlEdgeRight).Weight = xlMedium
                        .Range("B1").Interior.ColorIndex = 3
                        .Range("C1").Interior.ColorIndex = 20
                        .Range("A3:D3").Interior.ColorIndex = 4
                        .Range("A4:D4").Interior.ColorIndex = 41
                        .Range("A8:D8").Interior.ColorIndex = 26
                    
                        y = y + 1
                        .Range("A1:D" & CStr(i)).Copy sumSheet.Range(N2L((x - 1) * 4 + 5) & CStr((y - 1) * (i + 5) + 6))
                        sumSheet.Range(N2L((x - 1) * 4 + 5) & CStr((y - 1) * (i + 5) + 6 - 1)) = setSheet.Cells(1, iSheet)
                        sumSheet.Range(N2L((x - 1) * 4 + 5) & CStr((y - 1) * (i + 5) + 6 - 1) & ":" & N2L((x - 1) * 4 + 5 + 3) & CStr((y - 1) * (i + 5) + 6 - 1)).Merge
                        sumSheet.Range(N2L((x - 1) * 4 + 5) & CStr((y - 1) * (i + 5) + 6 - 1) & ":" & N2L((x - 1) * 4 + 5 + 3) & CStr((y - 1) * (i + 5) + 6 - 1)).Interior.ColorIndex = 36
                        sumSheet.Range(N2L((x - 1) * 4 + 5) & CStr((y - 1) * (i + 5) + 6 - 1) & ":" & N2L((x - 1) * 4 + 5 + 3) & CStr((y - 1) * (i + 5) + 6 - 1)).HorizontalAlignment = xlHAlignCenter
                        sumSheet.Range(N2L((x - 1) * 4 + 5) & CStr((y - 1) * (i + 5) + 6 - 1) & ":" & N2L((x - 1) * 4 + 5 + 3) & CStr((y - 1) * (i + 5) + 6 - 1)).Borders.Weight = xlThin
                        If x = 1 Then
                            .Range("A1:D" & CStr(i)).Copy sumSheet.Range(N2L(1) & CStr((y - 1) * (i + 5) + 6))
                            sumSheet.Range(N2L(1) & CStr((y - 1) * (i + 5) + 6 - 1)) = "total summary"
                            sumSheet.Range(N2L(1) & CStr((y - 1) * (i + 5) + 6 - 1)).HorizontalAlignment = xlHAlignCenter
                            sumSheet.Range(N2L(1) & CStr((y - 1) * (i + 5) + 6 - 1)).Interior.ColorIndex = 36
                            sumSheet.Range(N2L(1) & CStr((y - 1) * (i + 5) + 6 - 1) & ":" & N2L(4) & CStr((y - 1) * (i + 5) + 6 - 1)).Merge
                            With sumSheet.Range(N2L(1) & CStr((y - 1) * (i + 5) + 6 - 4))
                                .Cells(1, 1) = "w/i target and criteria"
                                .Cells(1, 1).Interior.ColorIndex = 27
                                .Range("A1:D1").Merge
                                .Range("B2:D2").Merge
                                .Range("B3:D3").Merge
                                .Range("A1:D1").HorizontalAlignment = xlHAlignCenter
                                .Range("B2:D2").HorizontalAlignment = xlHAlignCenter
                                .Range("B3:D3").HorizontalAlignment = xlHAlignCenter
                                .Cells(2, 1) = "LOT ID"
                                .Cells(2, 2) = nowSheet.Cells(1, 4)
                                .Cells(3, 1) = "Wafer #"
                                .Cells(3, 2) = nowSheet.Cells(1, iCol)
                                .Range("A1:D4").Borders.Weight = xlThin
                            End With
                        End If
                    End With
                End If
            Next iCol
        End If
        
        'Add K chart
        Dim yCount As Integer   'For 2010 , 20130722
        For i = nowSheet.ChartObjects.Count To 1 Step -1
            nowSheet.ChartObjects(i).Delete
        Next i
        'Set nowChart = nowSheet.ChartObjects.Add(10, 100, 600, 400).Chart
        'For 2010 相容
        Set nowChart = myCreateChart(nowSheet, xlLineMarkers, 10, 10, 600, 300)
        
        nowChart.chartType = xlLineMarkers
        Set xRange = nowSheet.Range(nowSheet.Cells(3, 2), nowSheet.Cells(3, 2).End(xlDown))
        For iCol = 1 To nowSheet.UsedRange.Columns.Count
            If nowSheet.Cells(2, iCol) = "K value" Then
               'Set yRange = nowSheet.Range(nowSheet.Cells(3, iCol), nowSheet.Cells(3, iCol).End(xlDown))
               Set yRange = nowSheet.Range(nowSheet.Cells(3, iCol), nowSheet.Cells(3 + xRange.Rows.Count - 1, iCol))
               Set nowSeries = nowChart.SeriesCollection.NewSeries
               nowSeries.XValues = xRange
               nowSeries.Values = yRange
               If WorksheetFunction.Count(yRange) <> 0 Then
                 nowSeries.Name = "#" & nowSheet.Cells(1, iCol - 3)
                 nowSeries.Border.LineStyle = xlDot ' xlDash
                 If yCount = 0 Then yCount = WorksheetFunction.Count(yRange)
               End If
            End If
        Next iCol
        Set nowSeries = nowChart.SeriesCollection.NewSeries
        nowSeries.chartType = xlXYScatter
        'nowSeries.XValues = "{1,1}"
        nowSeries.XValues = "{" & yCount / 2 & "," & yCount / 2 & "}"
        nowSeries.Values = "{1,-1}"
        nowSeries.MarkerStyle = xlMarkerStyleDash 'xlMarkerStyleNone
        nowSeries.MarkerForegroundColorIndex = 3
        nowSeries.Name = "Criteria"
        With nowSeries
            .ErrorBar xlX, Include:=xlBoth, Type:=xlFixedValue, Amount:=yCount / 2
            .ErrorBars.Border.ColorIndex = 3
            .ErrorBars.Border.LineStyle = xlDash
            .ErrorBars.Border.Weight = xlMedium
            .ErrorBars.EndStyle = xlNoCap
        End With
        
        With nowChart
            .ChartArea.Interior.ColorIndex = 2
            .PlotArea.Interior.ColorIndex = 2
            .Legend.Interior.ColorIndex = 2
            .HasTitle = True
            .ChartTitle.Characters.Text = nowSheet.Name
            .ChartTitle.Characters.Font.Name = "Lucida Sans Unicode"
        End With
        Set nowAxis = nowChart.Axes(xlCategory)
        With nowAxis
            .TickLabels.Font.Size = 10
            .TickLabels.Font.Name = "Lucida Sans Unicode"
            '.TickLabels.Font.Bold = True
            '.CrossesAt = 1
            .TickLabelSpacing = 1
            .TickMarkSpacing = 1
            '.AxisBetweenCategories = True
            '.ReversePlotOrder = False
            .TickLabels.Orientation = xlUpward
        End With
        Set nowAxis = nowChart.Axes(xlValue)
        With nowAxis
            .HasTitle = True
            .AxisTitle.Characters.Text = "K-Value"
            .AxisTitle.Characters.Font.Name = "Lucida Sans Unicode"
            If .MaximumScale > Abs(.MinimumScale) Then .MinimumScale = -1 * .MaximumScale Else .MaximumScale = -1 * .MinimumScale
            If .MaximumScale > Abs(.MinimumScale) Then .MinimumScale = -1 * .MaximumScale Else .MaximumScale = -1 * .MinimumScale
            If .MaximumScale > Abs(.MinimumScale) Then .MinimumScale = -1 * .MaximumScale Else .MaximumScale = -1 * .MinimumScale
            If .MaximumScale > 10 Then
                .MaximumScale = 10
                .MinimumScale = -10
            End If
            .Crosses = xlCustom
            .CrossesAt = .MinimumScale
            '.ReversePlotOrder = False
            '.ScaleType = xlLinear
            '.DisplayUnit = xlNone
            .TickLabels.Font.Size = 12
            .TickLabels.Font.Name = "Lucida Sans Unicode"
            '.TickLabels.Font.Bold = True
            .TickLabels.NumberFormatLocal = "0_ "
        End With
        '---------------------------------------------------
    Next iSheet
    
    For iRow = 1 To sumSheet.UsedRange.Rows.Count
        If sumSheet.Cells(iRow, 1) = "Priority" Then
            With sumSheet.Cells(iRow + 1, 1 + 1)
                For x = 1 To 3
                    For y = 1 To 10
                        tmp = 0: j = 0
                        For i = x + 4 To sumSheet.UsedRange.Columns.Count Step 4
                            tmp = tmp + .Cells(y, i): j = j + 1
                        Next i
                        .Cells(y, x) = tmp
                    Next y
                    If .Cells(1, x) <> 0 Then
                        .Cells(4, x) = Format(.Cells(2, x) / .Cells(1, x), "0%")
                        .Cells(5, x) = Format((.Cells(2, x) + .Cells(8, x)) / .Cells(1, x), "0%")
                        .Cells(6, x) = Format((.Cells(2, x) + .Cells(8, x) + .Cells(9, x)) / .Cells(1, x), "0%")
                    End If
                Next x
            End With
        End If
    Next iRow
    
    'Stop
    
    i = 0
    iCol = sumSheet.UsedRange.Columns.Count + 2
    With sumSheet.Cells(1, iCol)
        For iRow = 2 To sumSheet.UsedRange.Rows.Count Step 16
            i = i + 1
            .Cells(1, 1 + i) = sumSheet.Cells(iRow + 2, 2)
            For j = 0 To 2
                .Cells(2 + j, 1) = "Score_" & CStr(j + 1)
                .Cells(2 + j, 1 + i) = sumSheet.Cells(iRow + 8 + j, 4)
                .Cells(2 + j, 1 + i).NumberFormatLocal = "0%"
            Next j
        Next iRow
    End With
    
    Set nowRange = sumSheet.Cells(1, iCol + 1).CurrentRegion
    'sumSheet.Cells(1, iCol).Select
    'Set nowChart = sumSheet.ChartObjects.Add(CSng(sumSheet.Range("A:" & N2L(iCol - 1)).Width), 100, 500, 300).Chart
    'For 2010 相容
    Set nowChart = myCreateChart(nowSheet, xlLineMarkers, CSng(sumSheet.Range("A:" & N2L(iCol - 1)).width), 100, 500, 300)

    With nowChart
        .chartType = xlLineMarkers
        .SetSourceData Source:=nowRange, PlotBy:=xlRows
        .HasTitle = True
        .ChartTitle.Characters.Text = "Lot Total Summary Score Chart"
        .PlotArea.Interior.ColorIndex = 2
        .ChartArea.Interior.ColorIndex = 2
        .Legend.Interior.ColorIndex = 2
    End With
    Set nowAxis = nowChart.Axes(xlCategory)
    With nowAxis
        .TickLabels.Font.Size = 10
        .TickLabels.Font.FontStyle = "Lucida Sans Unicode"
        .HasTitle = True
        .AxisTitle.Characters.Text = "Wafer numbers"
        '.TickLabels.Font.Bold = True
        '.CrossesAt = 1
        '.TickLabelSpacing = 1
        '.TickMarkSpacing = 1
        '.AxisBetweenCategories = True
        '.ReversePlotOrder = False
        '.TickLabels.Orientation = xlUpward
    End With
    Set nowAxis = nowChart.Axes(xlValue)
    With nowAxis
        .TickLabels.Font.Size = 10
        .TickLabels.Font.FontStyle = "Lucida Sans Unicode"
        .HasTitle = True
        .AxisTitle.Characters.Text = "Score"
    End With
    
    
    sumSheet.Activate
    ActiveWindow.Zoom = 70
    'Stop
End Function

Public Function Manual_Vincent_PerformanceFit() '2012/05/15 Vincent
    Dim nowSheet As Worksheet
    Dim tarSerise As String
    Dim lineType As String
    Dim xyType As Long
    Dim nowRange As Range
    Dim iRow As Long, iCol As Long, bRow As Long
    Dim i As Long, j As Long
    Dim xRange As Range, yRange As Range
    Dim reArray()
    Dim a, b
    Dim tempX(), tempY()
    Dim tmp As String
    
    Set nowSheet = ActiveSheet
    
    tarSerise = InputBox("Input series name to fit:", "Series Name", "Target")
    If tarSerise = "" Then Exit Function
    lineType = InputBox("Trendline type: (1-Linear, 2-exponent)", "Trendline Type", "1")
    If lineType = "" Then Exit Function
    xyType = MsgBox("Base on X ?", vbYesNo)
    
    iCol = nowSheet.UsedRange.Columns.Count + 2
    Set nowRange = nowSheet.Range(N2L(iCol) & "1")
    With nowRange
        .Cells(1, 2) = IIf(xyType = 6, "X", "Y")
        .Cells(1, 3) = IIf(xyType = 6, "Y", "X")
        .Cells(1, 4) = "%"
        .Cells(2, 1) = tarSerise
        bRow = 2
        For i = 3 To iCol - 2 Step 2
            If nowSheet.Cells(1, i) <> "" And UCase(nowSheet.Cells(1, i)) <> "TARGET" Then
                If nowSheet.Cells(1, i) = tarSerise Then
                    iRow = 2
                Else
                    bRow = bRow + 1
                    iRow = bRow
                End If
                If iRow > 2 Then
                    .Cells(iRow, 1) = nowSheet.Cells(1, i)
                    .Cells(iRow, 2).FormulaLocal = "=" & N2L(iCol + 1) & CStr(2)
                    .Cells(iRow, 4).FormulaLocal = "=" & "(" & N2L(iCol + 2) & CStr(iRow) & "-" & N2L(iCol + 2) & CStr(2) & ")/" & N2L(iCol + 2) & CStr(2)
                    .Cells(iRow, 4).NumberFormatLocal = "00.00%"
                End If
                Set xRange = nowSheet.Range(nowSheet.Cells(3, i), nowSheet.Cells(3, i).End(xlDown))
                Set yRange = nowSheet.Range(nowSheet.Cells(3, i + 1), nowSheet.Cells(3, i + 1).End(xlDown))
                'ReDim tempA(yRange.Rows.Count - 1)
                tempX = xRange.Value
                tempY = yRange.Value
                Debug.Print xRange.Address, yRange.Address
                Select Case lineType
                    Case "1":   'linear
                        reArray = WorksheetFunction.LinEst(yRange, xRange)
                        If xyType = vbYes Then 'Base on X
                            tmp = reArray(1) & "*" & N2L(iCol + 1) & CStr(iRow) & IIf(reArray(2) >= 0, "+", "") & reArray(2)
                            'reArray = WorksheetFunction.LinEst(yRange, xRange)
                        Else
                            tmp = "(" & N2L(iCol + 1) & CStr(iRow) & "-" & reArray(2) & ")/" & reArray(1)  'y=m*x+b => x=(y-b)/m
                            'reArray = WorksheetFunction.LinEst(xRange, yRange)
                        End If
                        '.Cells(iRow, 3).FormulaLocal = "=" & Round(reArray(1), 4) & "*" & N2L(iCol + 1) & CStr(iRow) & IIf(reArray(2) >= 0, "+", "") & Round(reArray(2), 4)
                        .Cells(iRow, 3).FormulaLocal = "=" & tmp
                    Case "2":   'exponent
                        For j = 1 To UBound(tempY): tempY(j, 1) = WorksheetFunction.Ln(tempY(j, 1)): Next j
                        a = WorksheetFunction.Slope(tempY, tempX)
                        b = Exp(WorksheetFunction.Intercept(tempY, tempX))
                        If xyType = vbYes Then 'Base on X
                            tmp = b & "*EXP(" & a & "*" & N2L(iCol + 1) & CStr(iRow) & ")"
                            'For j = 1 To UBound(tempY): tempY(j, 1) = WorksheetFunction.Ln(tempY(j, 1)): Next j
                            'a = WorksheetFunction.Slope(tempY, tempX)
                            'b = Exp(WorksheetFunction.Intercept(tempY, tempX))
                            'reArray = WorksheetFunction.LinEst(tempY, tempX)
                            'a = WorksheetFunction.index(WorksheetFunction.LinEst(WorksheetFunction.Ln(yRange), xRange), 1)
                            'b = WorksheetFunction.Exp(WorksheetFunction.index(WorksheetFunction.LinEst(WorksheetFunction.Ln(yRange), xRange), 2))
                        Else
                            tmp = "(Ln(" & N2L(iCol + 1) & CStr(iRow) & ") - Ln(" & b & "))/" & a
                            'For j = 1 To UBound(tempX): tempX(j, 1) = WorksheetFunction.Ln(tempX(j, 1)): Next j
                            'a = WorksheetFunction.Slope(tempX, tempY)
                            'b = Exp(WorksheetFunction.Intercept(tempX, tempY))
                        End If
                        '.Cells(iRow, 3).FormulaLocal = "=" & reArray(2) & "*EXP(" & reArray(1) & "*" & N2L(iCol + 1) & CStr(iRow) & ")"
                        .Cells(iRow, 3).FormulaLocal = "=" & tmp
                        '.Cells(iRow, 2) = b
                End Select
            End If
            If nowSheet.Cells(1, i) = "" Then i = 255
        Next i
    End With
    
    Set nowRange = nowRange.CurrentRegion
    nowRange.Borders.LineStyle = xlContinuous
End Function


Public Sub Manual_DieSelection()    'Vincent 2012/07/24
    Dim nowRange As Range
    Dim nowSheet As Worksheet
    Dim i As Integer
    Dim iRow As Integer
    Dim DataSheet As Worksheet
    Dim nowWafer As String
    Dim bRow As Integer, waferRow As Integer
    
    Set DataSheet = Worksheets("data")
    Set nowSheet = ActiveSheet
    'Set nowRange = nowSheet.Range("G3:" & N2L(nowSheet.UsedRange.Columns.Count) & CStr(nowSheet.UsedRange.Rows.Count))
    'Debug.Print nowRange.Address
    'Exit Sub
    With nowSheet.Range("G3:" & N2L(nowSheet.UsedRange.Columns.Count) & CStr(nowSheet.UsedRange.Rows.Count))
        .Select
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=$F3"
        .FormatConditions(1).Font.ColorIndex = 3
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=$D3"
        .FormatConditions(2).Font.ColorIndex = 4
    End With
    'Exit Sub
    bRow = 1
    iRow = nowSheet.UsedRange.Rows.Count + 1
    Set nowRange = nowSheet.Cells(nowSheet.UsedRange.Rows.Count + 6, 1)
    With nowRange
        .Cells(1, 1) = "wafer"
        .Cells(1, 2) = "die#"
        .Cells(1, 3) = "pass die"
    End With
    'get row of wafer
    For i = 1 To 20
        If DataSheet.Cells(i, 2) = "Parameter" Then waferRow = i: Exit For
    Next i
    
    For i = 7 To nowSheet.UsedRange.Columns.Count
        nowSheet.Cells(iRow, i).FormulaArray = "=SUM(IF(" & N2L(i) & "3:" & N2L(i) & CStr(iRow - 1) & ">F3:F" & CStr(iRow - 1) & ",1,0))"
        nowSheet.Cells(iRow + 1, i).FormulaArray = "=SUM(IF(" & N2L(i) & "3:" & N2L(i) & CStr(iRow - 1) & "<D3:D" & CStr(iRow - 1) & ",1,0))"
        nowSheet.Cells(iRow + 2, i).Formula = "=SUM(" & N2L(i) & CStr(iRow) & ":" & N2L(i) & CStr(iRow + 1) & ")"
        nowSheet.Cells(iRow + 3, i) = getCOL(DataSheet.Cells(waferRow, 3 + nowSheet.Cells(2, i)), ">", 2)
        If nowWafer <> nowSheet.Cells(1, i) Then
            bRow = bRow + 1
            nowWafer = nowSheet.Cells(1, i)
            nowRange.Cells(bRow, 1) = "#" & nowWafer
        End If
        If nowSheet.Cells(iRow + 2, i) = 0 Then
            nowRange.Cells(bRow, 3) = nowRange.Cells(bRow, 3) & nowSheet.Cells(iRow + 3, i) & " "
            nowRange.Cells(bRow, 2) = nowRange.Cells(bRow, 2) + 1
        End If
    Next i
     
    Call GenDieMap(nowSheet)
    
End Sub

Public Function GenDieMap(nowSheet As Worksheet)    '2013/02/25 for Vincent
    Dim bRow As Long, bCol As Long
    Dim i As Long, j As Long
    Dim nowRange As Range
    Dim xMin As Integer, xMax As Integer, yMin As Integer, yMax As Integer
    Dim iWafer As Integer
    Dim waferRange As Range
    Dim x As Integer, y As Integer
    Dim tmp As String
    
    Set nowRange = nowSheet.Cells(nowSheet.UsedRange.Rows.Count + 2, 1)
    
    For i = 1 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(i, 2) = "die#" Then bRow = i: Exit For
    Next i
    Set waferRange = nowSheet.Cells(bRow, 1).CurrentRegion
    bCol = 7
    
    For i = bCol To nowSheet.UsedRange.Columns.Count
        If nowSheet.Cells(1, i) <> nowSheet.Cells(1, i - 1) And i > bCol Then Exit For
        nowRange.Cells(i - bCol + 1, 1) = getCOL(getCOL(nowSheet.Cells(bRow - 2, i), ",", 1), "(", 2)
        nowRange.Cells(i - bCol + 1, 2) = getCOL(getCOL(nowSheet.Cells(bRow - 2, i), ",", 2), ")", 1)
    Next i
    
    Set nowRange = nowRange.CurrentRegion
    xMin = WorksheetFunction.Min(nowRange.Columns(1))
    xMax = WorksheetFunction.Max(nowRange.Columns(1))
    yMin = WorksheetFunction.Min(nowRange.Columns(2))
    yMax = WorksheetFunction.Max(nowRange.Columns(2))
    'Debug.Print xMin, xMax, yMin, yMax
    
    For iWafer = 1 To waferRange.Rows.Count - 1
        With nowSheet.Cells(bRow + (iWafer - 1) * (yMax - yMin + 3), bCol)
            .Cells(1, 1) = waferRange.Cells(1 + iWafer, 1)
            For x = 1 To xMax - xMin + 1: .Cells(1, 1 + x) = xMin - 1 + x: Next x
            For y = 1 To yMax - yMin + 1: .Cells(1 + y, 1) = yMax + 1 - y: Next y
            For i = 1 To nowRange.Rows.Count
                tmp = "(" & nowRange.Cells(i, 1) & "," & nowRange.Cells(i, 2) & ")"
                If IsKey(waferRange.Cells(1 + iWafer, 3), tmp, " ") Then tmp = "Pass" Else tmp = "Fail"
                .Cells(2 + yMax - nowRange.Cells(i, 2), 2 + nowRange.Cells(i, 1) - xMin) = tmp
                If tmp = "Pass" Then .Cells(2 + yMax - nowRange.Cells(i, 2), 2 + nowRange.Cells(i, 1) - xMin).Interior.ColorIndex = 6
            Next i
        End With
    Next iWafer
    
End Function

Public Sub FixChart()
    Dim nowChart As Chart
    Dim nowSheet As Worksheet
    
    'On Error GoTo myEnd
    
    Set nowSheet = ActiveSheet
    Set nowChart = ActiveChart
    
    If Left(nowSheet.Name, 8) = "BOXTREND" Then
        'Fit 副坐標軸
        With nowChart.Axes(xlValue, xlSecondary)
            .MinimumScale = nowChart.Axes(xlValue, xlPrimary).MinimumScale
            .MaximumScale = nowChart.Axes(xlValue, xlPrimary).MaximumScale
            '.TickLabelPosition = xlNone
            '.MajorTickMark = xlNone
        End With
    End If
    
    Exit Sub
myEnd:
End Sub

Public Sub reCountCorner()

    Dim tArray()
    Dim x As Double, y As Double
    Dim i As Integer, j As Integer
    Dim nowSheet As Worksheet
    Dim oInfo As chartInfo
    Dim iEnd As Long
    Dim inCount As Long, OutCount As Long
    Dim m As Long, n As Long
    Dim nowChart As Chart
    Dim nowShape As Shape
   
    On Error Resume Next
    Set nowSheet = ActiveSheet
    If nowSheet.ChartObjects.Count = 0 Then Exit Sub
    
    inCount = 0: OutCount = 0
    
    For iEnd = 0 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(iEnd + 1, 1) = "" And nowSheet.Cells(iEnd + 1, 2) = "" Then Exit For
    Next iEnd
    
    oInfo = getChartInfo(nowSheet.Range("A1:B" & CStr(iEnd)))
    
    If oInfo.vCornerXValueStr <> "" Then
        ReDim tArray(Len(oInfo.vCornerXValueStr) - Len(Replace(oInfo.vCornerXValueStr, ",", "")))
        For j = 0 To UBound(tArray)
            tArray(j) = Array(Val(getCOL(oInfo.vCornerXValueStr, ",", j + 1)), Val(getCOL(oInfo.vCornerYValueStr, ",", j + 1)))
        Next j
        
        Call CornerSeq(tArray)
        For m = 3 To nowSheet.UsedRange.Columns.Count Step 2
            If IsKey("target,median,tt,ff,ss", nowSheet.Cells(1, m), ",", True) Then Exit For
           
            For n = 3 To nowSheet.UsedRange.Rows.Count
                If nowSheet.Cells(n, m) <> "" And nowSheet.Cells(n, m + 1) <> "" Then
                    If ynInCorner(tArray, nowSheet.Cells(n, m), nowSheet.Cells(n, m + 1)) Then
                        inCount = inCount + 1
                    Else
                        OutCount = OutCount + 1
                    End If
                End If
            Next n
        Next m
    End If

    If nowSheet.ChartObjects.Count > 0 Then
        If inCount > 0 Or OutCount > 0 Then
            Set nowChart = nowSheet.ChartObjects(1).Chart
            For i = nowChart.Shapes.Count To 1 Step -1
                If nowChart.Shapes(i).Type = msoTextBox Then
                    nowChart.Shapes(i).Delete
                End If
            Next i
            Set nowShape = nowChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 100, 20)
            nowSheet.Activate

            nowChart.Parent.Activate
            nowShape.Select
            With Selection
                .Characters.Text = "In: " & CStr(inCount) & " Out: " & CStr(OutCount) & " = " & Format(inCount / (inCount + OutCount), "0.00%")
                .Characters.Font.Size = 12
                .Font.ColorIndex = 3
                .Font.Bold = True
                .AutoSize = True
            End With
            With nowShape
                .Top = nowChart.PlotArea.Top + 12
                .Left = nowChart.PlotArea.Left + 40
            End With
            
            nowChart.ChartArea.Select
        End If
    End If
End Sub


Public Sub PinScatter()

    Dim i As Integer
    
    For i = 1 To Worksheets.Count
        If Left(Worksheets(i).Name, Len("SCATTER")) = "SCATTER" Then
            Worksheets(i).Name = Replace(Worksheets(i).Name, "SCATTER", "!SCATTER")
        End If
        If Left(Worksheets(i).Name, Len("BOXTREND")) = "BOXTREND" Then
            Worksheets(i).Name = Replace(Worksheets(i).Name, "BOXTREND", "!BOXTREND")
        End If
        If Left(Worksheets(i).Name, Len("CUMULATIVE")) = "CUMULATIVE" Then
            Worksheets(i).Name = Replace(Worksheets(i).Name, "CUMULATIVE", "!CUMULATIVE")
        End If
    Next i
    
End Sub

Public Sub UnpinScatter()

    Dim i As Integer
    
    For i = 1 To Worksheets.Count
        If Left(Worksheets(i).Name, Len("!SCATTER")) = "!SCATTER" Then
            Worksheets(i).Name = Replace(Worksheets(i).Name, "!SCATTER", "SCATTER")
        End If
        If Left(Worksheets(i).Name, Len("!BOXTREND")) = "!BOXTREND" Then
            Worksheets(i).Name = Replace(Worksheets(i).Name, "!BOXTREND", "BOXTREND")
        End If
        If Left(Worksheets(i).Name, Len("!CUMULATIVE")) = "!CUMULATIVE" Then
            Worksheets(i).Name = Replace(Worksheets(i).Name, "!CUMULATIVE", "CUMULATIVE")
        End If
    Next i
    
End Sub



Public Sub UpdateSummaryTable()

    Dim nowSheet As Worksheet
    Dim mRange() As Range
    Dim nowRange As Range
    Dim origin As Range
    Dim nowRow As Integer, nowCol As Integer
    Dim iRow As Integer, iCol As Integer
    Dim i As Integer, j As Integer
    
    Call Speed
    
    Set nowSheet = ActiveWorkbook.ActiveSheet
    nowRow = 1: nowCol = 1
    nowSheet.Cells(nowRow, nowCol).Select

    nowSheet.Cells.Find(What:="\BLOCK", _
                        After:=ActiveCell, _
                        LookIn:=xlFormulas, _
                        LookAt:=xlPart, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False, _
                        MatchByte:=False, _
                        SearchFormat:=False).Select
    Set origin = ActiveCell
    ReDim mRange(0) As Range
    Set mRange(0) = origin.CurrentRegion
    Cells.FindNext(After:=ActiveCell).Activate
    
    Do While ActiveCell.row <> origin.row Or ActiveCell.Column <> origin.Column
        ReDim Preserve mRange(UBound(mRange) + 1) As Range
        Set mRange(UBound(mRange)) = ActiveCell.CurrentRegion
        Cells.FindNext(After:=ActiveCell).Activate
    Loop
    
    For i = LBound(mRange) To UBound(mRange)
        Set nowRange = mRange(i)
        GoSub mySub
    Next i
    nowSheet.Cells(1, 1).Select
                        
    Call Unspeed
Exit Sub

mySub:
    Dim srcSheet As Worksheet
    Dim header As String
    Dim srcHeader As Object
    Set srcHeader = CreateObject("Scripting.Dictionary")
    
    iRow = 1: iCol = 2
    Do
        If nowRange.Cells(1, iCol).Value <> "" Then
            If Not IsExistSheet(nowRange.Cells(1, iCol).Value) Then
                MsgBox ("Cannot find " & nowRange.Cells(1, iCol).Value & " worksheet.")
                Exit Sub
            End If
            Set srcSheet = ActiveWorkbook.Worksheets(nowRange.Cells(1, iCol).Value)
            srcHeader.RemoveAll
            For j = 1 To srcSheet.UsedRange.Columns.Count
                If UCase(srcSheet.Cells(2, j).Value) = UCase(nowRange.Cells(2, 1)) Then
                    srcHeader.Add UCase(CStr(srcSheet.Cells(1, j).Value)), j
                ElseIf j < 7 Then
                    srcHeader.Add UCase(CStr(srcSheet.Cells(2, j).Value)), j
                End If
            Next j
            
            header = UCase(CStr(nowRange.Cells(iRow + 1, iCol).Value))
            iRow = iRow + 2
            Do
                If nowRange.Cells(iRow, 1) <> "" And UCase(nowRange.Cells(iRow, 1)) <> "SKIP" Then
                    If srcHeader.Exists(header) Then
                        If srcSheet.Cells(nowRange.Cells(iRow, 1).Value, srcHeader(header)) <> "" Then
                            nowRange.Cells(iRow, iCol).Value = "='" & srcSheet.Name & "'!" & Cells(nowRange.Cells(iRow, 1).Value, srcHeader(header)).Address(False, False)
                            nowRange.Cells(iRow, iCol).NumberFormatLocal = srcSheet.Cells(nowRange.Cells(iRow, 1).Value, srcHeader(header)).NumberFormatLocal
                        Else
                            nowRange.Cells(iRow, iCol).ClearContents
                        End If
                    End If
                ElseIf UCase(nowRange.Cells(iRow, 1)) = "SKIP" Then
                
                Else
                    nowRange.Cells(iRow, iCol).ClearContents
                End If
                iRow = iRow + 1
            Loop Until iRow > nowRange.Rows.Count
            iRow = 1
        ElseIf Not srcSheet Is Nothing And nowRange.Cells(2, iCol).Value <> "" Then
            header = UCase(CStr(nowRange.Cells(iRow + 1, iCol).Value))
            iRow = iRow + 2
            Do
                If nowRange.Cells(iRow, 1) <> "" And UCase(nowRange.Cells(iRow, 1)) <> "SKIP" Then
                    If srcHeader.Exists(header) Then
                        If srcSheet.Cells(nowRange.Cells(iRow, 1).Value, srcHeader(header)) <> "" Then
                            nowRange.Cells(iRow, iCol).Value = "='" & srcSheet.Name & "'!" & Cells(nowRange.Cells(iRow, 1).Value, srcHeader(header)).Address(False, False)
                            nowRange.Cells(iRow, iCol).NumberFormatLocal = srcSheet.Cells(nowRange.Cells(iRow, 1).Value, srcHeader(header)).NumberFormatLocal
                        Else
                            nowRange.Cells(iRow, iCol).ClearContents
                        End If
                    End If
                ElseIf UCase(nowRange.Cells(iRow, 1)) = "SKIP" Then
                
                Else
                    nowRange.Cells(iRow, iCol).ClearContents
                End If
                iRow = iRow + 1
            Loop Until iRow > nowRange.Rows.Count
            iRow = 1
        End If
        iCol = iCol + 1
    Loop Until iCol > nowRange.Columns.Count
    Set nowRange = Nothing
    Set srcSheet = Nothing

Return

End Sub

Public Sub addThruTrendLine()

    Dim nowSheet As Worksheet
    Dim nowChart As Chart
    Dim nowSeries As Series
    
    Set nowSheet = ActiveSheet
    Set nowChart = nowSheet.ChartObjects(1).Chart
    
    Dim i As Integer
    
    For i = nowChart.SeriesCollection.Count To 1 Step -1
        Set nowSeries = nowChart.SeriesCollection(i)
        If nowSeries.Name = "SS" Or nowSeries.Name = "FF" Then
            nowSeries.Delete
        ElseIf nowSeries.Name = "TT" Then
        
        Else
            With nowSeries.Format.Line
                .ForeColor.ObjectThemeColor = msoThemeColorText1
                .Visible = msoTrue
                .ForeColor.RGB = nowSeries.MarkerBackgroundColor
                .Weight = 1.5
            End With
            nowSeries.MarkerForegroundColor = RGB(0, 0, 0)
            nowSeries.MarkerSize = 5
        End If
    Next i

End Sub

Public Sub add3Sigma()

    Dim nowSheet As Worksheet
    Dim nowChart As Chart
    Dim nowLabels As DataLabels
    Dim nowSeries As Series
    
    Set nowSheet = ActiveSheet
    Set nowChart = nowSheet.ChartObjects(1).Chart
    
    Dim threeSigma As Double
    Dim strFormat As String
    
    Dim i As Integer
    For i = nowChart.SeriesCollection.Count To 1 Step -1
        Set nowSeries = nowChart.SeriesCollection(i)
        If nowSeries.Name = "0%" Then Exit For
    Next i
    
    nowSeries.ApplyDataLabels
    Set nowLabels = nowSeries.DataLabels
    nowLabels.Position = xlLabelPositionBelow
    
    For i = 1 To nowLabels.Count
        threeSigma = 3 * WorksheetFunction.StDev(Range(N2L(4 + i) & 12 & ":" & N2L(4 + i) & WorksheetFunction.countA(Columns(4 + i))))
        If threeSigma > 1 Then
            strFormat = "0.00"
        ElseIf threeSigma > 0.01 Then
            strFormat = "0.000"
        ElseIf threeSigma > 0.001 Then
            strFormat = "0.0000"
        Else
            strFormat = "0.0E+00"
        End If
        
        nowLabels(i).Text = "3σ=" & Format(threeSigma, strFormat)
        nowLabels(i).Format.TextFrame2.TextRange.Font.Size = 14
        nowLabels(i).Format.TextFrame2.TextRange.Font.Bold = msoTrue
        nowLabels(i).Format.TextFrame2.TextRange.Font.Name = "Arial"
        nowLabels(i).Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Worksheets("PlotSetup").ChartObject(1).Chart.SeriesCollection(i).Format.Fill.BackColor.RGB
        
    Next i
    
'    Dim AxisX As Axis
'    Set AxisX = nowChart.Axes(xlCategory)
'
'    AxisX.TickLabels.Font.Size = 14
'    AxisX.TickLabels.Font.Bold = msoTrue
    
End Sub

Public Sub genSingleChart()
      
    Call Speed
    Dim waferList() As String
    Dim siteNum As Integer
   
    If Not IsExistSheet("PlotSetup") Then MsgBox "Please check PlotSetup sheet before the operation!!": Exit Sub
    If IsExistSheet("Grouping") Then
        If Not isGroupingSafe Then Exit Sub
    End If
        
    Dim i As Long, j As Long
    Dim nowCol As Integer
    Dim nowRange As Range
    Dim nowSheet As Worksheet
    Dim newSheet As Worksheet
    
    Dim chartType As String
    Dim tmpStr As String
            
    'Get position
    Set nowSheet = ActiveSheet
    If Trim(UCase(nowSheet.Cells(1, Selection.Column).Value)) = "CHART TITLE" Then
        nowCol = Selection.Column
    ElseIf Trim(UCase(nowSheet.Cells(1, Selection.Column - 1).Value)) = "CHART TITLE" Then
        nowCol = Selection.Column - 1
    Else
        Exit Sub
    End If
    
    'Get ChartType
    For i = 1 To nowSheet.Cells(1, nowCol).CurrentRegion.Rows.Count
        If Trim(UCase(Left(nowSheet.Cells(i, nowCol).Value, 5))) = "GRAPH" Then
            chartType = "BOXTREND"
            Exit For
        ElseIf Trim(UCase(Left(nowSheet.Cells(i, nowCol).Value, 6))) = "METHOD" Then
            chartType = "CUMULATIVE"
            Exit For
        ElseIf Trim(UCase(Left(nowSheet.Cells(i, nowCol).Value, 6))) = "CORNER" Then
            chartType = "SCATTER"
            Exit For
        End If
    Next i
    If chartType = "" Then Exit Sub
    
    Set nowRange = Range(nowSheet.Cells(1, nowCol), nowSheet.Cells(nowSheet.UsedRange.Rows.Count, nowCol + 1))
    nowRange.ClearFormats
    
    j = 1
    
    Do While IsExistSheet(chartType & "_" & N2L(j))
        j = j + 1
    Loop
    
    Call PinScatter
    
    Set newSheet = AddSheet(chartType & "_" & N2L(j), , nowSheet.Name)
    nowRange.Copy
    newSheet.Range("A1").PasteSpecial xlPasteValues
    newSheet.Cells.ClearFormats
    
    Call GetWaferList(dSheet, waferList)
    siteNum = getSiteNum(dSheet)
    
    Select Case chartType
        Case "SCATTER"
            Call GenScatter(waferList, siteNum)
            Call PlotUniversalChart(newSheet.Name)
        Case "BOXTREND"
            Call GenBoxTrend(waferList, siteNum)
            Call prepareBoxTrendData(newSheet.Name)
            Call PlotBoxTrendChart(newSheet.Name)
        Case "CUMULATIVE"
            Call GenCumulative(waferList, siteNum)
            Call PlotCumulativeChart(newSheet.Name)
    End Select
    Call adjustChartObject(newSheet)
    
    Call FitSingleChart(newSheet)
    Call reCountCorner
    Call RawdataRange
    Call UnpinScatter
    Call Unspeed
    
    
End Sub

Public Sub updateChartSetting()

    Dim nowSheet As Worksheet
    Dim chartType As String
    
    Set nowSheet = ActiveSheet
    
    Dim i As Integer, j As Integer
    
    If Not UCase(nowSheet.Cells(1, 1).Value) = "CHART TITLE" Then Exit Sub
    
    For i = 1 To nowSheet.UsedRange.Rows.Count
        
        If Left(UCase(nowSheet.Cells(i, 1).Value), 5) = "GRAPH" Then
            chartType = "Boxtrend"
            Exit For
        ElseIf Left(UCase(nowSheet.Cells(i, 1).Value), 6) = "METHOD" Then
            chartType = "AccumulativeChart"
            Exit For
        ElseIf Left(UCase(nowSheet.Cells(i, 1).Value), 6) = "CORNER" Then
            chartType = "UniversalCurve"
            Exit For
        End If
        
    Next i
    
    For i = 3 To nowSheet.Cells(1, 1).CurrentRegion.Columns.Count Step 2
        
        Select Case chartType
            Case "Boxtrend"
                nowSheet.Cells(21, i + 1) = Round(getSPEC(nowSheet.Cells(23, i).Value, "TT"), 3)
                
            Case "UniversalCurve"
                If IsNumeric(nowSheet.Cells(21, i).Value) Then
                    Dim cnt As Integer
                    Dim startRow As Integer
                    cnt = 0
                    For j = 21 To nowSheet.UsedRange.Rows.Count
                        If nowSheet.Cells(j, i).Value = "TT" Then startRow = j: Exit For
                        If Not nowSheet.Cells(j, i).Value = "" Then cnt = cnt + 1
                    Next j
                    For j = 1 To cnt
                        nowSheet.Cells(startRow + j, i + 1).Value = Round(getSPEC(nowSheet.Cells(20 + j, i + 1).Value, "TT"), 3)
                    Next j
                ElseIf Left(getCOL(nowSheet.Cells(21, i).Value, "_", 1), Len(getCOL(nowSheet.Cells(21, i).Value, "_", 1)) - 1) = Left(getCOL(nowSheet.Cells(21, i + 1).Value, "_", 1), Len(getCOL(nowSheet.Cells(21, i + 1).Value, "_", 1)) - 1) And _
                     Right(getCOL(nowSheet.Cells(21, i).Value, "_", 1), 1) <> Right(getCOL(nowSheet.Cells(21, i + 1).Value, "_", 1), 1) Then
                    nowSheet.Cells(16, i + 1).Value = Round(getSPEC(nowSheet.Cells(21, i).Value, "TT"), 3)
                    nowSheet.Cells(17, i + 1).Value = Round(getSPEC(nowSheet.Cells(21, i + 1).Value, "TT"), 3)
                    nowSheet.Cells(18, i + 1).Value = getCORNER(nowSheet.Cells(21, i).Value)
                    nowSheet.Cells(19, i + 1).Value = getCORNER(nowSheet.Cells(21, i + 1).Value)
                
                Else
                    Dim tmpStrX As String
                    Dim tmpStrY As String
                    For j = 21 To nowSheet.UsedRange.Rows.Count
                        If nowSheet.Cells(j, i + 1).Value = "" Then Exit For
                        If getSPEC(nowSheet.Cells(j, i).Value, "TT") <> 0 Then tmpStrX = tmpStrX & ", " & Round(getSPEC(nowSheet.Cells(j, i).Value, "TT"), 3)
                        If getSPEC(nowSheet.Cells(j, i + 1).Value, "TT") <> 0 Then tmpStrY = tmpStrY & ", " & Round(getSPEC(nowSheet.Cells(j, i + 1).Value, "TT"), 3)
                    Next j
                    tmpStrX = Mid(tmpStrX, 3)
                    tmpStrY = Mid(tmpStrY, 3)
                    nowSheet.Cells(16, i + 1).Value = tmpStrX
                    nowSheet.Cells(17, i + 1).Value = tmpStrY
                    tmpStrX = ""
                    tmpStrY = ""
                End If
                
            Case Else
                
        End Select
    Next i

End Sub
