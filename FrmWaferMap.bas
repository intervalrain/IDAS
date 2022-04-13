
Option Explicit

Private Sub CB_ReferToList_Click()
    If CB_ReferToList.Value = True Then
        ComboPara.Enabled = False
    Else
        ComboPara.Enabled = True
    End If

End Sub

Private Sub CmdRun_Click()
    Dim ParaRange As Range
    Dim item
    Dim i As Integer
    Set ParaRange = Worksheets("ChartType").Columns(4)
    Set ParaRange = ParaRange.Range("A2:A" & CStr(ParaRange.Cells(1, 1).CurrentRegion.Rows.Count))
    If CB_ReferToList.Value = True Then
        i = 1
        For Each item In ParaRange
            Call genMapSub(item.Text, "WaferMap" & CStr(i))
            i = i + 1
        Next item
    Else
        genMapSub (ComboPara.Text)
    End If
    Me.Hide

End Sub


Private Sub UserForm_Initialize()
    Dim mWafer() As String
    Dim ParaRange As Range
    Dim i As Integer
   
    Call GetWaferArray(dSheet, mWafer)
    Set ParaRange = Worksheets(dSheet).Range("wafer_" & mWafer(0, 0)).Columns(2)
    Set ParaRange = ParaRange.Range("A2:A" & CStr(ParaRange.Rows.Count))
    Names.Add "ParaList", ParaRange
    ComboPara.RowSource = "ParaList"
    ComboWafer.Clear
    For i = 0 To UBound(mWafer, 2)
        ComboWafer.AddItem mWafer(0, i)
    Next i
    ComboWafer.AddItem "ALL"
    ComboWafer.ListIndex = ComboWafer.ListCount - 1
    
    ComboSpec.Clear
    ComboSpec.AddItem "75%/25%"
    ComboSpec.AddItem "Med+1s/Med-1s"
    ComboSpec.AddItem "Med+3s/Med-3s"
    ComboSpec.AddItem "Spec Hi/Spec Lo"
    ComboSpec.ListIndex = 0
End Sub

Private Sub genMapSub(Para As String, Optional sheetName As String = "WaferMap")
    Dim waferList() As String
    Dim vSpec As specInfo
    Dim siteNum As Integer
    Dim nowSheet As Worksheet
    Dim i As Integer, j As Integer, m As Double, n As Integer, nowWafer As Integer
    Dim xMin As Integer, xMax As Integer, yMin As Integer, yMax As Integer
    Dim tmpX As Integer, tmpY As Integer
    Dim delOrNot As Boolean
    Dim tmpStr As String
    Dim formatStr As String
    Dim iCol As Long, iRow As Long
    Dim nowRange As Range, Range2 As Range
    Dim nowCondition As FormatCondition
    Dim nowShape As Shape
    Dim nowChart As Chart
    Dim vMax As Double, vMin As Double, vMedian As Double, vSigma As Double
    Dim mapRange As Range
    Dim objPara As specInfo
    
    
   
    If ComboWafer.Text = "ALL" Then
        Call GetWaferList(dSheet, waferList)
    Else
        ReDim waferList(0)
        waferList(0) = ComboWafer.Text
    End If

    If CB_ReferToList.Value = True Then
        delOrNot = True
    ElseIf IsExistSheet(sheetName) Then
        delOrNot = MsgBox("Are you sure to overlap the worksheet?", vbYesNo, "Hint:") = vbYes
    End If
    
    iCol = 3
    iRow = 1
    
    Set nowSheet = AddSheet(sheetName, delOrNot)
    
    nowSheet.Activate
    
    siteNum = getSiteNum(dSheet)
    vSpec = getSPECInfo(Para)
    
    If vSpec.mUnit <> "" Then
        Select Case vSpec.mUnit
            Case "a", "mV/V", "mV/dec", "ohm/sq", "um", "uA", "pA", "uA/Cell", "nA/Cell"
                formatStr = "0.00"
            Case "fF/um", "uA/um", "nA/um", "V", "fF/um^2", "fF/Cell"
                formatStr = "0.000"
            Case "A", "A/cm2", "A/um", "A/um^2", "A/Cell"
                formatStr = "0.000E+00"
            Case Else
                formatStr = "0.00"
        End Select
    Else
        formatStr = "0.000"
    End If
    
    For nowWafer = 0 To UBound(waferList)
        If nowSheet.UsedRange.Rows.Count = 1 Then
            Set mapRange = nowSheet.Range("1:14")
        Else
            Set mapRange = nowSheet.Range(nowSheet.UsedRange.Rows.Count + 2 & ":" & nowSheet.UsedRange.Rows.Count + 15)
        End If
        
        For i = 1 To siteNum
            tmpStr = getValueByPara(waferList(nowWafer), "parameter", i, vSpec)
            tmpStr = getCOL(getCOL(tmpStr, "(", 2), ")", 1)
            If tmpStr = "" Then MsgBox ("Could not derive site info."): Exit Sub
            tmpX = CInt(getCOL(tmpStr, ",", 1))
            tmpY = CInt(getCOL(tmpStr, ",", 2))
            If i = 1 Then
                xMin = tmpX: xMax = tmpX
                yMin = tmpY: yMax = tmpY
            Else
                If tmpX < xMin Then xMin = tmpX
                If tmpX > xMax Then xMax = tmpX
                If tmpY < yMin Then yMin = tmpY
                If tmpY > yMax Then yMax = tmpY
            End If
        Next i
        
        tmpX = WorksheetFunction.Max(Abs(xMin), xMax, Abs(yMin), yMax)
        xMin = tmpX * -1
        yMin = tmpX * -1
        xMax = tmpX * 1
        yMax = tmpX * 1
        
        xMin = xMin - 1: xMax = xMax + 1
        yMin = yMin - 1: yMax = yMax + 1
        For i = 1 To xMax - xMin + 1
            mapRange.Cells(iRow, iCol + i) = xMin + i - 1
        Next i
        mapRange.Range(Cells(iRow, iCol + 1), Cells(iRow, iCol + xMax - xMin + 1)).Interior.ColorIndex = 12
        mapRange.Range(Cells(iRow, iCol + 1), Cells(iRow, iCol + xMax - xMin + 1)).HorizontalAlignment = xlCenter
        For i = 1 To yMax - yMin + 1
            mapRange.Cells(iRow + i, iCol) = yMax + 1 - i
        Next i
        mapRange.Range(Cells(iRow + 1, iCol), Cells(iRow + yMax - yMin + 1, iCol)).Interior.ColorIndex = 12
        Set nowRange = mapRange.Range(Cells(iRow + 1, iCol + 1), Cells(iRow + yMax - yMin + 1, iCol + xMax - xMin + 1))
        For i = 1 To siteNum
            tmpStr = getValueByPara(waferList(nowWafer), "Parameter", i, vSpec)
            tmpStr = getCOL(getCOL(tmpStr, "(", 2), ")", 1)
            tmpX = CInt(getCOL(tmpStr, ",", 1))
            tmpY = CInt(getCOL(tmpStr, ",", 2))
            tmpStr = getValueByPara(waferList(nowWafer), Para, i, vSpec)
            nowRange.Cells(1 + (yMax - tmpY), 1 + (tmpX - xMin)) = tmpStr
            nowRange.Cells(1 + (yMax - tmpY), 1 + (tmpX - xMin)).NumberFormatLocal = formatStr
            With nowRange.Range(N2L(1 + (tmpX - xMin)) & CStr(1 + (yMax - tmpY)))
                .Interior.Color = RGB(255, 255, 200)
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Font.Color = RGB(255, 0, 0)
                .FormatConditions.Delete
                Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & mapRange.Cells(5, 2).Address)
                nowCondition.Interior.Color = RGB(255, 100, 100)
                nowCondition.Font.Color = RGB(255, 0, 0)
                Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="=" & mapRange.Cells(3, 2).Address, Formula2:="=" & mapRange.Cells(4, 2).Address)
                nowCondition.Interior.Color = RGB(255, 255, 200)
                nowCondition.Font.Color = RGB(0, 0, 255)
                Set nowCondition = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & mapRange.Cells(4, 2).Address)
                nowCondition.Interior.Color = RGB(100, 100, 255)
                nowCondition.Font.Color = RGB(0, 0, 255)
            End With
        Next i
        With mapRange
            .Cells(1, 1) = "Parameter"
            .Cells(1, 2) = Para
            .Cells(2, 1) = "Wafer"
            .Cells(2, 2) = waferList(nowWafer)
            .Cells(3, 1) = "Median"
            .Cells(3, 2).Formula = "=MEDIAN(" & N2L(nowRange.Column) & CStr(nowRange.row) & ":" & N2L(nowRange.Columns.Count + nowRange.Column - 1) & CStr(nowRange.Rows.Count + nowRange.row - 1) & ")"
            
        If ComboSpec.ListIndex = 0 Then
            .Cells(4, 1) = "25%"
            .Cells(4, 2).Formula = "=Quartile(" & N2L(nowRange.Column) & CStr(nowRange.row) & ":" & N2L(nowRange.Columns.Count + nowRange.Column - 1) & CStr(nowRange.Rows.Count + nowRange.row - 1) & ",1)"
            .Cells(5, 1) = "75%"
            .Cells(5, 2).Formula = "=Quartile(" & N2L(nowRange.Column) & CStr(nowRange.row) & ":" & N2L(nowRange.Columns.Count + nowRange.Column - 1) & CStr(nowRange.Rows.Count + nowRange.row - 1) & ",3)"
        ElseIf ComboSpec.ListIndex = 1 Then
            .Cells(4, 1) = "Med-s"
            .Cells(4, 2).Formula = "=" & N2L(mapRange.Column + 1) & CStr(mapRange.row + 2) & "-" & N2L(mapRange.Column + 1) & CStr(mapRange.row + 7)
            .Cells(5, 1) = "Med+s"
            .Cells(5, 2).Formula = "=" & N2L(mapRange.Column + 1) & CStr(mapRange.row + 2) & "+" & N2L(mapRange.Column + 1) & CStr(mapRange.row + 7)
        ElseIf ComboSpec.ListIndex = 2 Then
            objPara = getSPECInfo(Trim(Para))
            .Cells(4, 1) = "Med-3s"
            .Cells(4, 2).Formula = "=" & N2L(mapRange.Column + 1) & CStr(mapRange.row + 2) & "-3*" & N2L(mapRange.Column + 1) & CStr(mapRange.row + 7)
            .Cells(5, 1) = "Med+3s"
            .Cells(5, 2).Formula = "=" & N2L(mapRange.Column + 1) & CStr(mapRange.row + 2) & "+3*" & N2L(mapRange.Column + 1) & CStr(mapRange.row + 7)
        ElseIf ComboSpec.ListIndex = 3 Then
            objPara = getSPECInfo(Trim(Para))
            .Cells(4, 1) = "SPEC Lo"
            .Cells(4, 2).Formula = objPara.mLow
            .Cells(5, 1) = "SPEC Hi"
            .Cells(5, 2).Formula = objPara.mHigh
        End If
        
            .Cells(6, 1) = "Max"
            .Cells(6, 2) = "=MAX(" & N2L(nowRange.Column) & CStr(nowRange.row) & ":" & N2L(nowRange.Columns.Count + nowRange.Column - 1) & CStr(nowRange.Rows.Count + nowRange.row - 1) & ")"
            .Cells(7, 1) = "Min"
            .Cells(7, 2) = "=MIN(" & N2L(nowRange.Column) & CStr(nowRange.row) & ":" & N2L(nowRange.Columns.Count + nowRange.Column - 1) & CStr(nowRange.Rows.Count + nowRange.row - 1) & ")"
            .Cells(8, 1) = "Sigma"
            .Cells(8, 2) = "=STDEV(" & N2L(nowRange.Column) & CStr(nowRange.row) & ":" & N2L(nowRange.Columns.Count + nowRange.Column - 1) & CStr(nowRange.Rows.Count + nowRange.row - 1) & ")"
            
            
            vMedian = Application.WorksheetFunction.Median(nowRange)
            vSigma = Application.WorksheetFunction.StDev(nowRange)
            vMax = vMedian + 3 * vSigma
            vMin = vMedian - 3 * vSigma
        End With
        Set nowRange = mapRange.Range("B3:B8")
        nowRange.NumberFormatLocal = formatStr
           
        mapRange.Range(Cells(1, 1), Cells(8, 1)).Interior.ColorIndex = 6
        mapRange.Columns.AutoFit
        
        For i = 1 To nowRange.Rows.Count
            For j = 1 To nowRange.Columns.Count
                If nowRange.Cells(i, j) = "" Then
                    n = 0
                    If nowRange.Cells(i - 1, j) <> "" Then n = n + 1
                    If nowRange.Cells(i + 1, j) <> "" Then n = n + 1
                    If nowRange.Cells(i, j - 1) <> "" Then n = n + 1
                    If nowRange.Cells(i, j + 1) <> "" Then n = n + 1
                    If n >= 3 Then
                        nowRange.Range(N2L(j) & CStr(i)).Borders.Weight = xlThin
                        nowRange.Range(N2L(j) & CStr(i)).Interior.Color = RGB(255, 255, 200)
                    End If
                End If
            Next j
        Next i
    Next nowWafer
    Cells.Columns.AutoFit
    Cells.Rows.AutoFit
End Sub
