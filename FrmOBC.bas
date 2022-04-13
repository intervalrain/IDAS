

Private Sub UserForm_Initialize()
    Dim i As Integer
    Dim maxL As Integer
    Call getListPara
    For i = 0 To ListPara.ListCount - 1
        If Len(ListPara.List(i)) > maxL Then maxL = Len(ListPara.List(i))
    Next i
    ListPara.ColumnWidths = maxL * 5
    List_X.ColumnWidths = maxL * 5
    List_Y.ColumnWidths = maxL * 5
    
    ComboThru.AddItem "W"
    ComboThru.AddItem "L"
    
End Sub


Private Sub ToggleScatter_Click()
    If ToggleScatter = True Then
        ToggleScatter = True
        ToggleBox = False
        ToggleCum = False
        
        ComboThru.Visible = False
        CB_logX.Visible = True
        CB_logX = False
        CB_AddX.Visible = True
        CB_DelX.Visible = True
        CB_DelXAll.Visible = True
        List_X.Visible = True
        CB_ThruTrend.Visible = True
        Label2.Visible = True
        CB_Corner.Visible = True
        Call CB_DelXAll_Click
    ElseIf ToggleScatter = False And ToggleBox = False And ToggleCum = False Then
        ToggleScatter = True
    Else
        ComboThru.Visible = False
        CB_logX.Visible = False
        CB_logX = True
        CB_AddX.Visible = False
        CB_DelX.Visible = False
        CB_DelXAll.Visible = False
        Label2.Visible = False
        CB_Corner.Visible = False
        CB_Corner = False
        Call CB_DelXAll_Click
    End If
End Sub

Private Sub ToggleBox_Click()
    If ToggleBox = True Then
        ToggleScatter = False
        ToggleBox = True
        ToggleCum = False
        List_X.Visible = True
        
        CB_ThruTrend.Visible = True
        
    ElseIf ToggleScatter = False And ToggleBox = False And ToggleCum = False Then
        ToggleBox = True
    End If
End Sub

Private Sub ToggleCum_Click()
    If ToggleCum = True Then
        ToggleScatter = False
        ToggleBox = False
        ToggleCum = True
        
        ToggleLinear.Visible = True
        ToggleWeiBull.Visible = True
        ToggleLogNor.Visible = True
        
        CB_ThruTrend.Visible = False
        
        Label3 = "X"
    ElseIf ToggleScatter = False And ToggleBox = False And ToggleCum = False Then
        ToggleCum = True
    Else
        ToggleLinear.Visible = False
        ToggleWeiBull.Visible = False
        ToggleLogNor.Visible = False
        Label3 = "Y"
    End If
End Sub

Private Sub ToggleLinear_Click()
    If ToggleLinear = True Then
        ToggleLinear = True
        ToggleWeiBull = False
        ToggleLogNor = False
    ElseIf ToggleLinear = False And ToggleWeiBull = False And ToggleLogNor = False Then
        ToggleLinear = True
    End If
End Sub

Private Sub ToggleWeibull_Click()
    If ToggleWeiBull = True Then
        ToggleLinear = False
        ToggleWeiBull = True
        ToggleLogNor = False
    ElseIf ToggleLinear = False And ToggleWeiBull = False And ToggleLogNor = False Then
        ToggleWeiBull = True
    End If
End Sub

Private Sub ToggleLogNor_Click()
    If ToggleLogNor = True Then
        ToggleLinear = False
        ToggleWeiBull = False
        ToggleLogNor = True
    ElseIf ToggleLinear = False And ToggleWeiBull = False And ToggleLogNor = False Then
        ToggleLogNor = True
    End If
End Sub
Private Sub CB_ThruTrend_Click()
    If CB_ThruTrend = True Then
        ComboThru.Visible = True
        CB_logX.Visible = False
        CB_AddX.Visible = False
        CB_DelX.Visible = False
        CB_DelXAll.Visible = False
        Call genThruNum
    End If
            
    If CB_ThruTrend = False Then
        If ToggleScatter = True Then
            ComboThru.Visible = False
            CB_logX.Visible = True
            CB_AddX.Visible = True
            CB_DelX.Visible = True
            CB_DelXAll.Visible = True
            Call CB_DelXAll_Click
        ElseIf ToggleBox = True Then
            ComboThru.Visible = False
            CB_logX.Visible = False
            CB_AddX.Visible = False
            CB_DelX.Visible = False
            CB_DelXAll.Visible = False
        End If
    End If
    
End Sub

Private Sub List_X_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    Dim oClipBoard As MSForms.DataObject
    
    Set oClipBoard = New MSForms.DataObject
    For i = 0 To List_X.ListCount - 1
        If List_X.Selected(i) Then
            oClipBoard.SetText List_X.List(i)
            oClipBoard.PutInClipboard
        End If
    Next i
    Set oClipBoard = Nothing
End Sub

Private Sub List_Y_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    Dim oClipBoard As MSForms.DataObject
    
    Set oClipBoard = New MSForms.DataObject
    For i = 0 To List_Y.ListCount - 1
        If List_Y.Selected(i) Then
            oClipBoard.SetText List_Y.List(i)
            oClipBoard.PutInClipboard
        End If
    Next i
    Set oClipBoard = Nothing
End Sub


Private Sub ListPara_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    Dim oClipBoard As MSForms.DataObject
    
    Set oClipBoard = New MSForms.DataObject
    For i = 0 To ListPara.ListCount - 1
        If ListPara.Selected(i) Then
            oClipBoard.SetText ListPara.List(i)
            oClipBoard.PutInClipboard
        End If
    Next i
    Set oClipBoard = Nothing
End Sub

Private Sub TB_Keyword_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i As Integer
    If KeyCode = vbKeyReturn Then
        Call getListPara
        For i = 0 To ListPara.ListCount - 1
            ListPara.Selected(i) = True
        Next i
    End If
End Sub

Public Sub SelectTboxText(ByRef tBox As MSForms.TextBox)

    If LastEntered <> tBox.Name Then

        LastEntered = tBox.Name

        With tBox
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With

    End If

End Sub


Private Sub CB_AddX_Click()
    Dim tmpStr As String
    Dim i As Integer
   
    For i = 0 To ListPara.ListCount - 1
        If ListPara.Selected(i) Then
            tmpStr = ListPara.List(i)
            If CB_SetABS = True Then tmpStr = "ABS(" & tmpStr & ")"
            If CB_SetMED = True Then tmpStr = "MEDIAN(" & tmpStr & ")"
            If CB_SetABS = True Or CB_SetMED = True Then tmpStr = "'=" & tmpStr
            List_X.AddItem tmpStr
        End If
    Next i
End Sub


Private Sub CB_AddY_Click()
    Dim tmpStr As String
    Dim i As Integer
   
    For i = 0 To ListPara.ListCount - 1
        If ListPara.Selected(i) Then
            tmpStr = ListPara.List(i)
            If CB_SetABS = True Then tmpStr = "ABS(" & tmpStr & ")"
            If CB_SetMED = True Then tmpStr = "MEDIAN(" & tmpStr & ")"
            If CB_SetABS = True Or CB_SetMED = True Then tmpStr = "'=" & tmpStr
            List_Y.AddItem tmpStr
        End If
    Next i
    If CB_ThruTrend = True Then Call genThruNum
End Sub
Private Sub CB_DelX_Click()
    Dim i As Integer
   
    For i = List_X.ListCount - 1 To 0 Step -1
        If List_X.Selected(i) Then
            List_X.RemoveItem (i)
        End If
    Next i
End Sub

Private Sub CB_DelY_Click()
    Dim i As Integer
   
    For i = List_Y.ListCount - 1 To 0 Step -1
        If List_Y.Selected(i) Then
            List_Y.RemoveItem (i)
            If CB_ThruTrend = True Then List_X.RemoveItem (i)
        End If
    Next i
End Sub

Private Sub CB_DelXAll_Click()
    For i = List_X.ListCount - 1 To 0 Step -1
        List_X.RemoveItem (i)
    Next i
End Sub

Private Sub CB_DelYAll_Click()
    For i = List_Y.ListCount - 1 To 0 Step -1
        List_Y.RemoveItem (i)
    Next i
    If CB_ThruTrend = True Then List_X.Clear
End Sub

Private Sub CB_Add_Click()
    If ToggleScatter Then
        Call genUCurSheet(False)
    ElseIf ToggleBox Then
        Call genBoxSheet(False)
    ElseIf ToggleCum Then
        Call genCumSheet(False)
    End If
End Sub

Private Sub CB_Overwrite_Click()
    If ToggleScatter Then
        Call genUCurSheet(True)
    ElseIf ToggleBox Then
        Call genBoxSheet(True)
    ElseIf ToggleCum Then
        Call genCumSheet(True)
    End If
End Sub

Private Function getListPara()
    Dim mWafer() As String
    Dim ParaRange As Range
    Dim mRange As Range
   
    Call GetWaferArray(dSheet, mWafer)
    Set ParaRange = Worksheets(dSheet).Range("wafer_" & mWafer(0, 0)).Columns(2)
    Set ParaRange = ParaRange.Range("A2:A" & CStr(ParaRange.Rows.Count))
   
    ListPara.Clear
    For Each mRange In ParaRange
        If Not TB_Keyword = "" Then
            If InStr(TB_Keyword, "*") Then
                If UCase(mRange.Value) Like UCase(TB_Keyword) Then ListPara.AddItem mRange.Value
            Else
                If InStr(UCase(mRange.Value), UCase(TB_Keyword)) Then ListPara.AddItem mRange.Value
            End If
        Else
            ListPara.AddItem mRange.Value
        End If
    Next mRange

End Function

Private Function genThruNum()
    Dim i As Integer
    Dim L As Double
    Dim tmpStr
    Dim mCol As Integer
    
    If ComboThru = "" Or IsNumeric(ComboCol) = False Then Exit Function
    List_X.Clear
    If ComboThru = "W" Then
        mCol = ComboW
    ElseIf ComboThru = "L" Then
        mCol = ComboL
    ElseIf IsNumeric(ComboThru) Then
        mCol = ComboThru
    Else
        Exit Function
    End If
    
    For i = 0 To List_Y.ListCount - 1
        tmpStr = Replace(getCOL(List_Y.List(i), "_", mCol), "p", ".")
        If Not IsNumeric(tmpStr) Then Exit Function
        List_X.AddItem CDbl(tmpStr)
    Next i
End Function

Private Sub ComboThru_Change()
    Call genThruNum
End Sub

Private Sub ComboW_Change()
    Call genThruNum
End Sub

Private Sub ComboL_Change()
    Call genThruNum
End Sub

Private Function getLabel(Param As String)
    Dim tmpStr As String
    
    tmpStr = UCase(Param)
    
    If Left(tmpStr, 1) = "V" Then getLabel = Param & " (V)"
    If Left(tmpStr, 3) = "IOF" Then getLabel = Param & " (pA/um)"
    If Left(tmpStr, 1) = "I" Then getLabel = Param & " (uA/um)"
    If Left(tmpStr, 3) = "BVD" Then getLabel = Param & " (V)"
    If Left(tmpStr, 4) = "Rout" Then getLabel = Param & " (KOhm/um)"
    If Left(tmpStr, 1) = "R" Then getLabel = Param & " (Ohm)"
    If Left(tmpStr, 4) = "DIBL" Then getLabel = Param & " (V)"
    If Left(tmpStr, 3) = "SWS" Then getLabel = Param & " (mV/dec.)"
    If Left(tmpStr, 2) = "BF" Then getLabel = Param & " (mV)"
    If Left(tmpStr, 2) = "GM" Then getLabel = Param & " (uS/um)"
    If Left(tmpStr, 3) = "TOX" Then getLabel = Param & " (A)"
    If Left(tmpStr, 2) = "JG" Then getLabel = Param & " (A/cm2)"
    If Left(tmpStr, 1) = "C" Then getLabel = Param & " (fF/um)"
    
End Function

Private Function genUCurSheet(mType As Boolean)
    Dim nowSheet As Worksheet
    Dim nowCol As Integer
    Dim Title As String
    Dim LabelX As String
    Dim LabelY As String
    Dim TargetX As String
    Dim TargetY As String
    Dim CornerX As String
    Dim CornerY As String
    Dim width As Double
    Dim length As Double
    Dim widthp As String
    Dim lengthp As String
    Dim Device As String
    Dim i As Integer
    Dim SpecX As specInfo
    Dim SpecY As specInfo
    
    If TB_Sheetname = "" Then Exit Function
    Set nowSheet = AddSheet(TB_Sheetname, mType, "PlotSetup")
    
    On Error Resume Next
    LabelX = Replace(Replace(Replace(getCOL(List_X.List(0), "_", 1), "'=", ""), "MEDIAN(", ""), "ABS(", "")
    LabelY = Replace(Replace(Replace(getCOL(List_Y.List(0), "_", 1), "'=", ""), "MEDIAN(", ""), "ABS(", "")
    widthp = getCOL(List_Y.List(0), "_", ComboW)
    lengthp = getCOL(List_Y.List(0), "_", ComboL)
    Device = Replace(Replace(getCOL(List_Y.List(0), "_", ComboD), "(", ""), ")", "")
    width = CDbl(Replace(widthp, "p", "."))
    length = CDbl(Replace(lengthp, "p", "."))
    If TB_Title = "" Then
        If CB_ThruTrend Then
            Title = "Thru-" & ComboThru & " Trend " & "(" & IIf(ComboThru = "W", "L=" & length, "W=" & width) & ")"
        ElseIf CB_Corner Then
            Title = Left(LabelX, Len(LabelX) - 1) & " Corner of " & Device & "(W/L=" & width & "/" & length & ")"
        Else
            Title = Device & " " & LabelX & "-" & LabelY & " (W/L=" & width & "/" & length & ")"
        End If
    Else
        Title = TB_Title
        Title = Replace(Title, "[Device]", Device)
        Title = Replace(Title, "[Width]", width)
        Title = Replace(Title, "[Length]", length)
    End If
    
    For i = 0 To List_X.ListCount - 1
        SpecX = getSPECInfo(trimFunc(trimFunc(List_X.List(i), "ABS"), "MEDIUM"))
        If Not SpecX.mTarget = "" Then TargetX = TargetX & ", " & SpecX.mTarget
    Next i
    If Len(TargetX) > 1 Then TargetX = Mid(TargetX, 3)
    For i = 0 To List_Y.ListCount - 1
        SpecY = getSPECInfo(trimFunc(trimFunc(List_Y.List(i), "ABS"), "MEDIUM"))
        If Not SpecY.mTarget = "" Then TargetY = TargetY & ", " & SpecY.mTarget
    Next i
    
    If CB_Corner And IsExistSheet("Corner") And UBound(List_X.List) = 0 And UBound(List_Y.List) = 0 Then
        
        CornerX = getCORNER(List_X.List)
        CornerY = getCORNER(List_Y.List)
        
    End If
    If Len(TargetY) > 1 Then TargetY = Mid(TargetY, 3)
    
    '初始化
    If Not nowSheet.Cells(1, 1) = "Chart title" And Not nowSheet.Cells(2, 2) = "ALL/Lot/Wafer/SplitID" Then
        With nowSheet
            .Cells(1, 1) = "Chart title"
            .Cells(2, 1) = "Split by"
            .Cells(3, 1) = "X label"
            .Cells(4, 1) = "Y label"
            .Cells(5, 1) = "X scale"
            .Cells(6, 1) = "Y scale"
            .Cells(7, 1) = "XMin"
            .Cells(8, 1) = "XMax"
            .Cells(9, 1) = "YMin"
            .Cells(10, 1) = "YMax"
            .Cells(11, 1) = "Chart expression"
            .Cells(12, 1) = "Group Params"
            .Cells(13, 1) = "Data Filter"
            .Cells(14, 1) = "TrendLines"
            .Cells(15, 1) = "TARGET NAME"
            .Cells(16, 1) = "TARGET XVALUE"
            .Cells(17, 1) = "TARGET YVALUE"
            .Cells(18, 1) = "CORNER XVALUE"
            .Cells(19, 1) = "CORNER YVALUE"
            .Cells(20, 1) = "Y"
            
            .Cells(1, 2) = "VTON"
            .Cells(2, 2) = "ALL/Lot/Wafer/SplitID/Group"
            .Cells(3, 2) = "um"
            .Cells(4, 2) = "VTON"
            .Cells(5, 2) = "Linear/Log"
            .Cells(6, 2) = "Linear/Log"
            .Cells(7, 2) = "0.1"
            .Cells(8, 2) = "0.6"
            .Cells(9, 2) = "0.1"
            .Cells(10, 2) = "0.6"
            .Cells(11, 2) = "All+RawData+Average+Median"
            .Cells(12, 2) = "Yes/No"
            .Cells(13, 2) = "Yes/No"
            .Cells(14, 2) = "Yes/No"
            .Cells(15, 2) = ""
            .Cells(16, 2) = ""
            .Cells(17, 2) = ""
            .Cells(18, 2) = ""
            .Cells(19, 2) = ""
            .Cells(20, 2) = "X"
        End With
    End If
    
    nowCol = nowSheet.UsedRange.Columns.Count + 1
    With nowSheet
        .Cells(1, nowCol) = .Cells(1, 1)
        .Cells(2, nowCol) = .Cells(2, 1)
        .Cells(3, nowCol) = .Cells(3, 1)
        .Cells(4, nowCol) = .Cells(4, 1)
        .Cells(5, nowCol) = .Cells(5, 1)
        .Cells(6, nowCol) = .Cells(6, 1)
        .Cells(7, nowCol) = .Cells(7, 1)
        .Cells(8, nowCol) = .Cells(8, 1)
        .Cells(9, nowCol) = .Cells(9, 1)
        .Cells(10, nowCol) = .Cells(10, 1)
        .Cells(11, nowCol) = .Cells(11, 1)
        .Cells(12, nowCol) = .Cells(12, 1)
        .Cells(13, nowCol) = .Cells(13, 1)
        .Cells(14, nowCol) = .Cells(14, 1)
        .Cells(15, nowCol) = .Cells(15, 1)
        .Cells(16, nowCol) = .Cells(16, 1)
        .Cells(17, nowCol) = .Cells(17, 1)
        .Cells(18, nowCol) = .Cells(18, 1)
        .Cells(19, nowCol) = .Cells(19, 1)
        .Cells(20, nowCol) = "X"
        
        .Cells(1, nowCol + 1) = Title
        .Cells(2, nowCol + 1) = "ALL"
        If CB_ThruTrend Then
            .Cells(3, nowCol + 1) = "Length (um)"
        Else
            .Cells(3, nowCol + 1) = getLabel(LabelX)
        End If
        .Cells(4, nowCol + 1) = getLabel(LabelY)
        .Cells(5, nowCol + 1) = IIf(CB_logX = True, "Log", "Linear")
        .Cells(6, nowCol + 1) = IIf(CB_logY = True, "Log", "Linear")
        .Cells(7, nowCol + 1) = ""
        .Cells(8, nowCol + 1) = ""
        .Cells(9, nowCol + 1) = ""
        .Cells(10, nowCol + 1) = ""
        .Cells(11, nowCol + 1) = "All"
        .Cells(12, nowCol + 1) = "Yes"
        .Cells(13, nowCol + 1) = "No"
        .Cells(14, nowCol + 1) = "No"
        .Cells(15, nowCol + 1) = "Target"
        .Cells(16, nowCol + 1) = TargetX
        .Cells(17, nowCol + 1) = TargetY
        .Cells(18, nowCol + 1) = CornerX
        .Cells(19, nowCol + 1) = CornerY
        .Cells(20, nowCol + 1) = "Y"
        
        For i = 0 To List_X.ListCount - 1
            .Cells(21 + i, nowCol) = List_X.List(i)
        Next i
        For i = 0 To List_Y.ListCount - 1
            .Cells(21 + i, nowCol + 1) = List_Y.List(i)
        Next i
        
        If CB_ThruTrend Then
            For i = 0 To List_X.ListCount - 1
                .Cells(23 + List_X.ListCount * 1 + i, nowCol) = List_X.List(i)
                .Cells(24 + List_X.ListCount * 2 + i, nowCol) = List_X.List(i)
                .Cells(25 + List_X.ListCount * 3 + i, nowCol) = List_X.List(i)
                
                Dim mTT
                Dim mFF
                Dim mSS
                
                mTT = getSPEC(List_Y.List(i), "TT")
                mFF = getSPEC(List_Y.List(i), "FF")
                mSS = getSPEC(List_Y.List(i), "SS")
                
                .Cells(23 + List_X.ListCount * 1 + i, nowCol + 1) = IIf(mTT = 0, "", mTT)
                .Cells(24 + List_X.ListCount * 2 + i, nowCol + 1) = IIf(mFF = 0, "", mFF)
                .Cells(25 + List_X.ListCount * 3 + i, nowCol + 1) = IIf(mSS = 0, "", mSS)
            Next i
            .Cells(22 + List_X.ListCount * 1, nowCol) = "TT"
            .Cells(23 + List_X.ListCount * 2, nowCol) = "FF"
            .Cells(24 + List_X.ListCount * 3, nowCol) = "SS"
        End If
    End With
    
End Function

Private Function genBoxSheet(mType As Boolean)
    Dim nowSheet As Worksheet
    Dim nowCol As Integer
    Dim Title As String
    Dim LabelY As String
    Dim width As Double
    Dim length As Double
    Dim widthp As String
    Dim lengthp As String
    Dim Device As String
    Dim i As Integer
    Dim SpecY As specInfo
    
    If TB_Sheetname = "" Then Exit Function
    Set nowSheet = AddSheet(TB_Sheetname, mType, "PlotSetup")
    
    On Error Resume Next
    
    If CB_ThruTrend = True Then
        LabelX = "Rule(um)"
        LabelY = Replace(Replace(Replace(getCOL(List_Y.List(0), "_", 1), "'=", ""), "MEDIAN(", ""), "ABS(", "")
        widthp = getCOL(List_Y.List(0), "_", ComboW)
        lengthp = getCOL(List_Y.List(0), "_", ComboL)
        Device = Replace(Replace(getCOL(List_Y.List(0), "_", ComboD), "(", ""), ")", "")
        width = CDbl(Replace(widthp, "p", "."))
        length = CDbl(Replace(lengthp, "p", "."))
    Else
        LabelX = "Wafer"
        LabelY = Replace(Replace(Replace(getCOL(List_Y.List(0), "_", 1), "'=", ""), "MEDIAN(", ""), "ABS(", "")
        widthp = getCOL(List_Y.List(0), "_", ComboW)
        lengthp = getCOL(List_Y.List(0), "_", ComboL)
        Device = Replace(Replace(getCOL(List_Y.List(0), "_", ComboD), "(", ""), ")", "")
        width = CDbl(Replace(widthp, "p", "."))
        length = CDbl(Replace(lengthp, "p", "."))
    End If
            
    If TB_Title = "" Then
        Title = Device & " Boxtrend Chart"
    Else
        Title = TB_Title
        Title = Replace(Title, "[Device]", Device)
        Title = Replace(Title, "[Width]", width)
        Title = Replace(Title, "[Length]", length)
    End If
    
    For i = 0 To List_Y.ListCount - 1
        SpecY = getSPECInfo(trimFunc(trimFunc(List_Y.List(i), "ABS"), "MEDIUM"))
        If Not SpecY.mTarget = "" Then TargetY = TargetY & ", " & SpecY.mTarget
    Next i
    If Len(TargetY) > 1 Then TargetY = Mid(TargetY, 3)
    '初始化
    If Not nowSheet.Cells(1, 1) = "Chart title" And Not nowSheet.Cells(2, 2) = "ALL/Lot/Wafer/SplitID" Then
        With nowSheet
            .Cells(1, 1) = "Chart title"
            .Cells(2, 1) = "Split By"
            .Cells(3, 1) = "Split ID"
            .Cells(4, 1) = "X label"
            .Cells(5, 1) = "Y label"
            .Cells(6, 1) = "Y scale"
            .Cells(7, 1) = "YMin"
            .Cells(8, 1) = "Ymax"
            .Cells(9, 1) = "Graph Max%"
            .Cells(10, 1) = "Graph Hi%"
            .Cells(11, 1) = "Graph Lo%"
            .Cells(12, 1) = "Graph Min%"
            .Cells(13, 1) = "Extend By"
            .Cells(14, 1) = "Sigma"
            .Cells(15, 1) = "Data Filter"
            .Cells(16, 1) = "Disable Max Min"
            .Cells(17, 1) = "Wafer Seq"
            .Cells(18, 1) = "Group Lot"
            .Cells(19, 1) = "Group Wafer"
            .Cells(20, 1) = "Target Name"
            .Cells(21, 1) = "Target YValue"
            If CB_ThruTrend Then
                .Cells(22, 1) = "X"
                .Cells(22, 2) = "Y"
            Else
                .Cells(22, 1) = "Y"
            End If
        
            .Cells(1, 2) = "VTON"
            .Cells(2, 2) = "ALL/Lot/Wafer/SplitID/Group"
            .Cells(3, 2) = "ALL/{Split ID}"
            .Cells(4, 2) = "LOT/Wafer/SplitID"
            .Cells(5, 2) = "VTON"
            .Cells(6, 2) = "LOG"
            .Cells(7, 2) = "0.1"
            .Cells(8, 2) = "0.6"
            .Cells(9, 2) = "100"
            .Cells(10, 2) = "80"
            .Cells(11, 2) = "20"
            .Cells(12, 2) = "0"
            .Cells(13, 2) = "Lot/Params"
            .Cells(14, 2) = "Median/Average/None"
            .Cells(15, 2) = "Yes/No"
            .Cells(16, 2) = "Yes/No"
            .Cells(17, 2) = "Yes/No"
            .Cells(18, 2) = "Yes/No"
            .Cells(19, 2) = "Yes/No"
            .Cells(20, 2) = "Target,USL,LSL"
        End With
    End If
    
    nowCol = nowSheet.UsedRange.Columns.Count + 1
    With nowSheet
        .Cells(1, nowCol) = .Cells(1, 1)
        .Cells(2, nowCol) = .Cells(2, 1)
        .Cells(3, nowCol) = .Cells(3, 1)
        .Cells(4, nowCol) = .Cells(4, 1)
        .Cells(5, nowCol) = .Cells(5, 1)
        .Cells(6, nowCol) = .Cells(6, 1)
        .Cells(7, nowCol) = .Cells(7, 1)
        .Cells(8, nowCol) = .Cells(8, 1)
        .Cells(9, nowCol) = .Cells(9, 1)
        .Cells(10, nowCol) = .Cells(10, 1)
        .Cells(11, nowCol) = .Cells(11, 1)
        .Cells(12, nowCol) = .Cells(12, 1)
        .Cells(13, nowCol) = .Cells(13, 1)
        .Cells(14, nowCol) = .Cells(14, 1)
        .Cells(15, nowCol) = .Cells(15, 1)
        .Cells(16, nowCol) = .Cells(16, 1)
        .Cells(17, nowCol) = .Cells(17, 1)
        .Cells(18, nowCol) = .Cells(18, 1)
        .Cells(19, nowCol) = .Cells(19, 1)
        .Cells(20, nowCol) = .Cells(20, 1)
        .Cells(21, nowCol) = .Cells(21, 1)
        .Cells(22, nowCol) = "Y"
        
        .Cells(1, nowCol + 1) = Title
        .Cells(2, nowCol + 1) = "ALL"
        .Cells(3, nowCol + 1) = "ALL"
        .Cells(4, nowCol + 1) = LabelX
        .Cells(5, nowCol + 1) = getLabel(LabelY)
        .Cells(6, nowCol + 1) = IIf(CB_logY = True, "Log", "Linear")
        .Cells(7, nowCol + 1) = ""
        .Cells(8, nowCol + 1) = ""
        .Cells(9, nowCol + 1) = "100"
        .Cells(10, nowCol + 1) = "75"
        .Cells(11, nowCol + 1) = "25"
        .Cells(12, nowCol + 1) = "0"
        .Cells(13, nowCol + 1) = "Lot"
        .Cells(14, nowCol + 1) = "Median"
        .Cells(15, nowCol + 1) = "No"
        .Cells(16, nowCol + 1) = "No"
        .Cells(17, nowCol + 1) = "Yes"
        .Cells(18, nowCol + 1) = "Yes"
        .Cells(19, nowCol + 1) = "No"
        .Cells(20, nowCol + 1) = "Target"
        .Cells(21, nowCol + 1) = TargetY
        
        If CB_ThruTrend Then
            .Cells(22, nowCol) = "X"
            .Cells(22, nowcow + 1) = "Y"
            For i = 0 To List_X.ListCount - 1
                .Cells(23 + i, nowCol) = List_X.List(i)
            Next i
            For i = 0 To List_Y.ListCount - 1
                .Cells(23 + i, nowCol + 1) = List_Y.List(i)
            Next i
        Else
            For i = 0 To List_Y.ListCount - 1
                .Cells(23 + i, nowCol) = List_Y.List(i)
            Next i
        End If


    End With
    
End Function

Private Function genCumSheet(mType As Boolean)
    Dim nowSheet As Worksheet
    Dim nowCol As Integer
    Dim Title As String
    Dim LabelX As String
    Dim width As Double
    Dim length As Double
    Dim widthp As String
    Dim lengthp As String
    Dim Device As String
    Dim i As Integer
    Dim SpecY As specInfo
    
    If TB_Sheetname = "" Then Exit Function
    Set nowSheet = AddSheet(TB_Sheetname, mType, "PlotSetup")
    
    On Error Resume Next
    
    LabelY = Replace(Replace(Replace(getCOL(List_Y.List(0), "_", 1), "'=", ""), "MEDIAN(", ""), "ABS(", "")
    widthp = getCOL(List_Y.List(0), "_", ComboW)
    lengthp = getCOL(List_Y.List(0), "_", ComboL)
    Device = Replace(Replace(getCOL(List_Y.List(0), "_", ComboD), "(", ""), ")", "")
    width = CDbl(Replace(widthp, "p", "."))
    length = CDbl(Replace(lengthp, "p", "."))
    If TB_Title = "" Then
        Title = Device & " Accumulative Chart"
    Else
        Title = TB_Title
        Title = Replace(Title, "[Device]", Device)
        Title = Replace(Title, "[Width]", width)
        Title = Replace(Title, "[Length]", length)
    End If
    
    
    SpecY = getSPECInfo(List_Y.List(0))

    '初始化
    If Not nowSheet.Cells(1, 1) = "Chart title" And Not nowSheet.Cells(2, 2) = "ALL/Lot/Wafer/SplitID" Then
        With nowSheet
            .Cells(1, 1) = "Chart title"
            .Cells(2, 1) = "Split by"
            .Cells(3, 1) = "Split ID"
            .Cells(4, 1) = "X label"
            .Cells(5, 1) = "Method"
            .Cells(6, 1) = "Xmax"
            .Cells(7, 1) = "Xmin"
            .Cells(8, 1) = "X scale"
            .Cells(9, 1) = "Group Params"
            .Cells(10, 1) = "Data Filter"
            .Cells(11, 1) = "Y"
                    
            .Cells(1, 2) = "VTON"
            .Cells(2, 2) = "ALL/Lot/Wafer/SplitID/Group"
            .Cells(3, 2) = "ALL/{Split ID}"
            .Cells(4, 2) = getLabel(LabelX)
            .Cells(5, 2) = "Linear/WeiBull/LogNor"
            .Cells(6, 2) = "0.6"
            .Cells(7, 2) = "0.1"
            .Cells(8, 2) = "Linear/Log"
            .Cells(9, 2) = "Yes/No"
            .Cells(10, 2) = "Yes/No"
        End With
    End If
    
    nowCol = nowSheet.UsedRange.Columns.Count + 1
    With nowSheet
        .Cells(1, nowCol) = .Cells(1, 1)
        .Cells(2, nowCol) = .Cells(2, 1)
        .Cells(3, nowCol) = .Cells(3, 1)
        .Cells(4, nowCol) = .Cells(4, 1)
        .Cells(5, nowCol) = .Cells(5, 1)
        .Cells(6, nowCol) = .Cells(6, 1)
        .Cells(7, nowCol) = .Cells(7, 1)
        .Cells(8, nowCol) = .Cells(8, 1)
        .Cells(9, nowCol) = .Cells(9, 1)
        .Cells(10, nowCol) = .Cells(10, 1)
        .Cells(11, nowCol) = .Cells(11, 1)
        
        .Cells(1, nowCol + 1) = Title
        .Cells(2, nowCol + 1) = "ALL"
        .Cells(3, nowCol + 1) = "ALL"
        .Cells(4, nowCol + 1) = getLabel(LabelX)
        If ToggleLinear Then
            .Cells(5, nowCol + 1) = "Linear"
        ElseIf ToggleWeiBull Then
            .Cells(5, nowCol + 1) = "WeiBull"
        ElseIf ToggleWeiBull Then
            .Cells(5, nowCol + 1) = "LogNor"
        End If
        .Cells(6, nowCol + 1) = ""
        .Cells(7, nowCol + 1) = ""
        .Cells(8, nowCol + 1) = IIf(CB_logY = True, "Log", "Linear")
        .Cells(9, nowCol + 1) = "No"
        .Cells(10, nowCol + 1) = "No"

        For i = 0 To List_Y.ListCount - 1
            .Cells(12 + i, nowCol) = List_Y.List(i)
        Next i
    End With
    
End Function
