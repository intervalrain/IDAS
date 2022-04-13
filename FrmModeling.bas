
Option Explicit

Dim DataColl As New Collection

Private Sub CheckSA_Click()
    If Not CheckSA Then
        Label4.Visible = False
        ListSA.Visible = False
        CmdSASelAll.Visible = False
        CmdSASelNone.Visible = False
        Me.width = 260
    Else
        Label4.Visible = True
        ListSA.Visible = True
        CmdSASelAll.Visible = True
        CmdSASelNone.Visible = True
        Me.width = 350
    End If
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdLengthSelAll_Click()
    Dim i As Integer
    
    For i = 0 To ListLength.ListCount - 1
        ListLength.Selected(i) = True
    Next i
End Sub

Private Sub CmdLengthSelNone_Click()
    Dim i As Integer
    
    For i = 0 To ListLength.ListCount - 1
        ListLength.Selected(i) = False
    Next i
End Sub

Private Sub CmdRun_Click()
    Select Case ComboPlot.Text
        Case "Width": initSheet ("B1")
            If CheckSA.Value Then GenWidth Else GenWidth_NoSA
        Case "Length": initSheet ("B1")
            If CheckSA.Value Then GenLength Else GenLength_NoSA
        Case "SA": initSheet ("B1")
            If CheckSA.Value Then GenSA Else Exit Sub
    End Select
    Me.Hide
    
    Call GenCharts
    
    Unload Me
End Sub

Private Sub GenWidth_NoSA()
    Dim nowSheet As Worksheet
    Dim iWidth As Integer
    Dim iLength As Integer
    'Dim iSA As Integer
    Dim nowCol As Long
    Dim nowRow As Long
    Dim tmp As String
    Dim i As Integer
    
    Set nowSheet = Worksheets("B1")
    
    nowCol = 1
    For iLength = 0 To ListLength.ListCount - 1
        If ListLength.Selected(iLength) Then
        'For iSA = 0 To ListSA.ListCount - 1
        '    If ListSA.Selected(iSA) Then
            nowCol = nowCol + 2
            nowSheet.Cells(1, nowCol) = "Chart title"
            nowSheet.Cells(2, nowCol) = "Split By"
            nowSheet.Cells(3, nowCol) = "Split ID"
            nowSheet.Cells(4, nowCol) = "X label"
            nowSheet.Cells(5, nowCol) = "Y label"
            nowSheet.Cells(6, nowCol) = "Y scale"
            nowSheet.Cells(7, nowCol) = "yMax"
            nowSheet.Cells(8, nowCol) = "yMin"
            nowSheet.Cells(9, nowCol) = "Graph Max%"
            nowSheet.Cells(10, nowCol) = "Graph Hi%"
            nowSheet.Cells(11, nowCol) = "Graph Lo%"
            nowSheet.Cells(12, nowCol) = "Graph Min%"
            nowSheet.Cells(13, nowCol) = "Extend By"
            nowSheet.Cells(14, nowCol) = "Sigma"
            nowSheet.Cells(15, nowCol) = "Data Filter"
            nowSheet.Cells(16, nowCol) = "Disable Max Min"
            nowSheet.Cells(17, nowCol) = "Wafer Seq"
            nowSheet.Cells(18, nowCol) = "Group Lot"
            nowSheet.Cells(19, nowCol) = "Target yValue"
            
            nowSheet.Cells(1, nowCol + 1) = ComboDevice.Text & " L=" & ListLength.List(iLength) '& " SA=" & ListSA.List(iSA)
            nowSheet.Cells(2, nowCol + 1) = "ALL"
            nowSheet.Cells(3, nowCol + 1) = "ALL"
            nowSheet.Cells(4, nowCol + 1) = "Width"
            nowSheet.Cells(5, nowCol + 1) = ComboDevice.Text
            nowSheet.Cells(6, nowCol + 1) = "Linear"
            nowSheet.Cells(7, nowCol + 1) = ""
            nowSheet.Cells(8, nowCol + 1) = ""
            nowSheet.Cells(9, nowCol + 1) = "100"
            nowSheet.Cells(10, nowCol + 1) = "75"
            nowSheet.Cells(11, nowCol + 1) = "25"
            nowSheet.Cells(12, nowCol + 1) = "0"
            nowSheet.Cells(13, nowCol + 1) = "Lot"
            nowSheet.Cells(14, nowCol + 1) = "Median"
            nowSheet.Cells(15, nowCol + 1) = "No"
            nowSheet.Cells(16, nowCol + 1) = "No"
            nowSheet.Cells(17, nowCol + 1) = "Yes"
            nowSheet.Cells(18, nowCol + 1) = "No"
            
            nowSheet.Cells(20, nowCol) = "Y"
            
            nowRow = 20
            For iWidth = 0 To ListWidth.ListCount - 1
                tmp = ComboDevice.Text & "_" & ListWidth.List(iWidth) & "_" & ListLength.List(iLength) '& "_" & ListSA.List(iSA)
                For i = 1 To DataColl.Count
                    If Left(DataColl(i), Len(tmp)) = tmp Then
                        nowRow = nowRow + 1
                        nowSheet.Cells(nowRow, nowCol) = DataColl(i)
                        nowSheet.Cells(nowRow, nowCol + 1) = ListWidth.List(iWidth)
                    End If
                Next i
            Next iWidth
        '    End If
        'Next iSA
        End If
    Next iLength


    nowSheet.Cells.Font.Size = 8
    nowSheet.Columns.AutoFit
End Sub

Private Sub GenLength_NoSA()
    Dim nowSheet As Worksheet
    Dim iWidth As Integer
    Dim iLength As Integer
    'Dim iSA As Integer
    Dim nowCol As Long
    Dim nowRow As Long
    Dim tmp As String
    Dim i As Integer
    
    Set nowSheet = Worksheets("B1")
    
    nowCol = 1
    For iWidth = 0 To ListWidth.ListCount - 1
        If ListWidth.Selected(iWidth) Then
        'For iSA = 0 To ListSA.ListCount - 1
        '    If ListSA.Selected(iSA) Then
            nowCol = nowCol + 2
            nowSheet.Cells(1, nowCol) = "Chart title"
            nowSheet.Cells(2, nowCol) = "Split By"
            nowSheet.Cells(3, nowCol) = "Split ID"
            nowSheet.Cells(4, nowCol) = "X label"
            nowSheet.Cells(5, nowCol) = "Y label"
            nowSheet.Cells(6, nowCol) = "Y scale"
            nowSheet.Cells(7, nowCol) = "yMax"
            nowSheet.Cells(8, nowCol) = "yMin"
            nowSheet.Cells(9, nowCol) = "Graph Max%"
            nowSheet.Cells(10, nowCol) = "Graph Hi%"
            nowSheet.Cells(11, nowCol) = "Graph Lo%"
            nowSheet.Cells(12, nowCol) = "Graph Min%"
            nowSheet.Cells(13, nowCol) = "Extend By"
            nowSheet.Cells(14, nowCol) = "Sigma"
            nowSheet.Cells(15, nowCol) = "Data Filter"
            nowSheet.Cells(16, nowCol) = "Disable Max Min"
            nowSheet.Cells(17, nowCol) = "Wafer Seq"
            nowSheet.Cells(18, nowCol) = "Group Lot"
            nowSheet.Cells(19, nowCol) = "Target yValue"
            
            nowSheet.Cells(1, nowCol + 1) = ComboDevice.Text & " W=" & ListWidth.List(iWidth) '& " SA=" & ListSA.List(iSA)
            nowSheet.Cells(2, nowCol + 1) = "ALL"
            nowSheet.Cells(3, nowCol + 1) = "ALL"
            nowSheet.Cells(4, nowCol + 1) = "Length"
            nowSheet.Cells(5, nowCol + 1) = ComboDevice.Text
            nowSheet.Cells(6, nowCol + 1) = "Linear"
            nowSheet.Cells(7, nowCol + 1) = ""
            nowSheet.Cells(8, nowCol + 1) = ""
            nowSheet.Cells(9, nowCol + 1) = "100"
            nowSheet.Cells(10, nowCol + 1) = "75"
            nowSheet.Cells(11, nowCol + 1) = "25"
            nowSheet.Cells(12, nowCol + 1) = "0"
            nowSheet.Cells(13, nowCol + 1) = "Lot"
            nowSheet.Cells(14, nowCol + 1) = "Median"
            nowSheet.Cells(15, nowCol + 1) = "No"
            nowSheet.Cells(16, nowCol + 1) = "No"
            nowSheet.Cells(17, nowCol + 1) = "Yes"
            nowSheet.Cells(18, nowCol + 1) = "No"
            
            nowSheet.Cells(20, nowCol) = "Y"
            
            nowRow = 20
            For iLength = 0 To ListLength.ListCount - 1
                tmp = ComboDevice.Text & "_" & ListWidth.List(iWidth) & "_" & ListLength.List(iLength) '& "_" & ListSA.List(iSA)
                For i = 1 To DataColl.Count
                    If Left(DataColl(i), Len(tmp)) = tmp Then
                        nowRow = nowRow + 1
                        nowSheet.Cells(nowRow, nowCol) = DataColl(i)
                        nowSheet.Cells(nowRow, nowCol + 1) = ListLength.List(iLength)
                    End If
                Next i
            Next iLength
            End If
        'Next iSA
        'End If
    Next iWidth


    nowSheet.Cells.Font.Size = 8
    nowSheet.Columns.AutoFit
End Sub

Private Sub initSheet(mSheet As String)
    Dim nowSheet As Worksheet
    
    Set nowSheet = AddSheet("ChartType")
    nowSheet.Cells(1, 1) = "ChartType"
    nowSheet.Cells(1, 2) = "Worksheet"
    nowSheet.Cells(2, 1) = "BOXTREND"
    nowSheet.Cells(2, 2) = mSheet
    
    Set nowSheet = AddSheet(mSheet)
    
    nowSheet.Cells(1, 1) = "Chart title"
    nowSheet.Cells(1, 2) = "VTON"
    nowSheet.Cells(2, 1) = "Split By"
    nowSheet.Cells(2, 2) = "ALL/Lot/Wafer/SplitID"
    nowSheet.Cells(3, 1) = "Split ID"
    nowSheet.Cells(3, 2) = "ALL/{Split ID}"
    nowSheet.Cells(4, 1) = "X label"
    nowSheet.Cells(4, 2) = "LOT/Wafer/SplitID"
    nowSheet.Cells(5, 1) = "Y label"
    nowSheet.Cells(5, 2) = "VTON"
    nowSheet.Cells(6, 1) = "Y scale"
    nowSheet.Cells(6, 2) = "Log"
    nowSheet.Cells(7, 1) = "yMax"
    nowSheet.Cells(7, 2) = "0.6"
    nowSheet.Cells(8, 1) = "yMin"
    nowSheet.Cells(8, 2) = "0.1"
    nowSheet.Cells(9, 1) = "Graph Max%"
    nowSheet.Cells(9, 2) = "100"
    nowSheet.Cells(10, 1) = "Graph Hi%"
    nowSheet.Cells(10, 2) = "80"
    nowSheet.Cells(11, 1) = "Graph Lo%"
    nowSheet.Cells(11, 2) = "20"
    nowSheet.Cells(12, 1) = "Graph Min%"
    nowSheet.Cells(12, 2) = "0"
    nowSheet.Cells(13, 1) = "Extend By"
    nowSheet.Cells(13, 2) = "Lot/Params"
    nowSheet.Cells(14, 1) = "Sigma"
    nowSheet.Cells(14, 2) = "Median/Average/None"
    nowSheet.Cells(15, 1) = "Data Filter"
    nowSheet.Cells(15, 2) = "Yes/No"
    nowSheet.Cells(16, 1) = "Disable Max Min"
    nowSheet.Cells(16, 2) = "Yes/No"
    nowSheet.Cells(17, 1) = "Wafer Seq"
    nowSheet.Cells(17, 2) = "Yes/No"
    nowSheet.Cells(18, 1) = "Group Lot"
    nowSheet.Cells(18, 2) = "Yes/No"
    nowSheet.Cells(19, 1) = "Target yValue"
    nowSheet.Cells(20, 1) = "Y"
    nowSheet.Cells(21, 1) = "VTOP_10_10_HS_1p2V_MB1"
    nowSheet.Cells(22, 1) = "VTOP_10_1_HS_1p2V_MB1"
    nowSheet.Cells(23, 1) = "VTOP_10_p24_HS_1p2V_MB1"
End Sub

Private Sub GenWidth()
    Dim nowSheet As Worksheet
    Dim iWidth As Integer
    Dim iLength As Integer
    Dim iSA As Integer
    Dim nowCol As Long
    Dim nowRow As Long
    Dim tmp As String
    
    Set nowSheet = Worksheets("B1")
    
    nowCol = 1
    For iLength = 0 To ListLength.ListCount - 1
        If ListLength.Selected(iLength) Then
        For iSA = 0 To ListSA.ListCount - 1
            If ListSA.Selected(iSA) Then
            nowCol = nowCol + 2
            nowSheet.Cells(1, nowCol) = "Chart title"
            nowSheet.Cells(2, nowCol) = "Split By"
            nowSheet.Cells(3, nowCol) = "Split ID"
            nowSheet.Cells(4, nowCol) = "X label"
            nowSheet.Cells(5, nowCol) = "Y label"
            nowSheet.Cells(6, nowCol) = "Y scale"
            nowSheet.Cells(7, nowCol) = "yMax"
            nowSheet.Cells(8, nowCol) = "yMin"
            nowSheet.Cells(9, nowCol) = "Graph Max%"
            nowSheet.Cells(10, nowCol) = "Graph Hi%"
            nowSheet.Cells(11, nowCol) = "Graph Lo%"
            nowSheet.Cells(12, nowCol) = "Graph Min%"
            nowSheet.Cells(13, nowCol) = "Extend By"
            nowSheet.Cells(14, nowCol) = "Sigma"
            nowSheet.Cells(15, nowCol) = "Data Filter"
            nowSheet.Cells(16, nowCol) = "Disable Max Min"
            nowSheet.Cells(17, nowCol) = "Wafer Seq"
            nowSheet.Cells(18, nowCol) = "Group Lot"
            nowSheet.Cells(19, nowCol) = "Target yValue"
            
            nowSheet.Cells(1, nowCol + 1) = ComboDevice.Text & " L=" & ListLength.List(iLength) & " SA=" & ListSA.List(iSA)
            nowSheet.Cells(2, nowCol + 1) = "ALL"
            nowSheet.Cells(3, nowCol + 1) = "ALL"
            nowSheet.Cells(4, nowCol + 1) = "Width"
            nowSheet.Cells(5, nowCol + 1) = ComboDevice.Text
            nowSheet.Cells(6, nowCol + 1) = "Linear"
            nowSheet.Cells(7, nowCol + 1) = ""
            nowSheet.Cells(8, nowCol + 1) = ""
            nowSheet.Cells(9, nowCol + 1) = "100"
            nowSheet.Cells(10, nowCol + 1) = "75"
            nowSheet.Cells(11, nowCol + 1) = "25"
            nowSheet.Cells(12, nowCol + 1) = "0"
            nowSheet.Cells(13, nowCol + 1) = "Lot"
            nowSheet.Cells(14, nowCol + 1) = "Median"
            nowSheet.Cells(15, nowCol + 1) = "No"
            nowSheet.Cells(16, nowCol + 1) = "No"
            nowSheet.Cells(17, nowCol + 1) = "Yes"
            nowSheet.Cells(18, nowCol + 1) = "No"
            
            nowSheet.Cells(20, nowCol) = "Y"
            
            nowRow = 20
            For iWidth = 0 To ListWidth.ListCount - 1
                tmp = ComboDevice.Text & "_" & ListWidth.List(iWidth) & "_" & ListLength.List(iLength) & "_" & ListSA.List(iSA)
                If existInColl(DataColl, tmp) Then
                    nowRow = nowRow + 1
                    nowSheet.Cells(nowRow, nowCol) = DataColl(tmp)
                    nowSheet.Cells(nowRow, nowCol + 1) = ListWidth.List(iWidth)
                End If
            Next iWidth
            End If
        Next iSA
        End If
    Next iLength


    nowSheet.Cells.Font.Size = 8
    nowSheet.Columns.AutoFit
End Sub



Private Sub GenLength()
    Dim nowSheet As Worksheet
    Dim iWidth As Integer
    Dim iLength As Integer
    Dim iSA As Integer
    Dim nowCol As Long
    Dim nowRow As Long
    Dim tmp As String
    
    Set nowSheet = Worksheets("B1")
    
    nowCol = 1
    For iWidth = 0 To ListWidth.ListCount - 1
        If ListWidth.Selected(iWidth) Then
        For iSA = 0 To ListSA.ListCount - 1
            If ListSA.Selected(iSA) Then
            nowCol = nowCol + 2
            nowSheet.Cells(1, nowCol) = "Chart title"
            nowSheet.Cells(2, nowCol) = "Split By"
            nowSheet.Cells(3, nowCol) = "Split ID"
            nowSheet.Cells(4, nowCol) = "X label"
            nowSheet.Cells(5, nowCol) = "Y label"
            nowSheet.Cells(6, nowCol) = "Y scale"
            nowSheet.Cells(7, nowCol) = "yMax"
            nowSheet.Cells(8, nowCol) = "yMin"
            nowSheet.Cells(9, nowCol) = "Graph Max%"
            nowSheet.Cells(10, nowCol) = "Graph Hi%"
            nowSheet.Cells(11, nowCol) = "Graph Lo%"
            nowSheet.Cells(12, nowCol) = "Graph Min%"
            nowSheet.Cells(13, nowCol) = "Extend By"
            nowSheet.Cells(14, nowCol) = "Sigma"
            nowSheet.Cells(15, nowCol) = "Data Filter"
            nowSheet.Cells(16, nowCol) = "Disable Max Min"
            nowSheet.Cells(17, nowCol) = "Wafer Seq"
            nowSheet.Cells(18, nowCol) = "Group Lot"
            nowSheet.Cells(19, nowCol) = "Target yValue"
            
            nowSheet.Cells(1, nowCol + 1) = ComboDevice.Text & " W=" & ListWidth.List(iWidth) & " SA=" & ListSA.List(iSA)
            nowSheet.Cells(2, nowCol + 1) = "ALL"
            nowSheet.Cells(3, nowCol + 1) = "ALL"
            nowSheet.Cells(4, nowCol + 1) = "Length"
            nowSheet.Cells(5, nowCol + 1) = ComboDevice.Text
            nowSheet.Cells(6, nowCol + 1) = "Linear"
            nowSheet.Cells(7, nowCol + 1) = ""
            nowSheet.Cells(8, nowCol + 1) = ""
            nowSheet.Cells(9, nowCol + 1) = "100"
            nowSheet.Cells(10, nowCol + 1) = "75"
            nowSheet.Cells(11, nowCol + 1) = "25"
            nowSheet.Cells(12, nowCol + 1) = "0"
            nowSheet.Cells(13, nowCol + 1) = "Lot"
            nowSheet.Cells(14, nowCol + 1) = "Median"
            nowSheet.Cells(15, nowCol + 1) = "No"
            nowSheet.Cells(16, nowCol + 1) = "No"
            nowSheet.Cells(17, nowCol + 1) = "Yes"
            nowSheet.Cells(18, nowCol + 1) = "No"
            
            nowSheet.Cells(20, nowCol) = "Y"
            
            nowRow = 20
            For iLength = 0 To ListLength.ListCount - 1
                tmp = ComboDevice.Text & "_" & ListWidth.List(iWidth) & "_" & ListLength.List(iLength) & "_" & ListSA.List(iSA)
                If existInColl(DataColl, tmp) Then
                    nowRow = nowRow + 1
                    nowSheet.Cells(nowRow, nowCol) = DataColl(tmp)
                    nowSheet.Cells(nowRow, nowCol + 1) = ListLength.List(iLength)
                End If
            Next iLength
            End If
        Next iSA
        End If
    Next iWidth


    nowSheet.Cells.Font.Size = 8
    nowSheet.Columns.AutoFit
End Sub

Private Sub GenSA()
    Dim nowSheet As Worksheet
    Dim iWidth As Integer
    Dim iLength As Integer
    Dim iSA As Integer
    Dim nowCol As Long
    Dim nowRow As Long
    Dim tmp As String
    
    Set nowSheet = Worksheets("B1")
    
    nowCol = 1
    For iWidth = 0 To ListWidth.ListCount - 1
        If ListWidth.Selected(iWidth) Then
        For iLength = 0 To ListLength.ListCount - 1
            If ListLength.Selected(iLength) Then
            nowCol = nowCol + 2
            nowSheet.Cells(1, nowCol) = "Chart title"
            nowSheet.Cells(2, nowCol) = "Split By"
            nowSheet.Cells(3, nowCol) = "Split ID"
            nowSheet.Cells(4, nowCol) = "X label"
            nowSheet.Cells(5, nowCol) = "Y label"
            nowSheet.Cells(6, nowCol) = "Y scale"
            nowSheet.Cells(7, nowCol) = "yMax"
            nowSheet.Cells(8, nowCol) = "yMin"
            nowSheet.Cells(9, nowCol) = "Graph Max%"
            nowSheet.Cells(10, nowCol) = "Graph Hi%"
            nowSheet.Cells(11, nowCol) = "Graph Lo%"
            nowSheet.Cells(12, nowCol) = "Graph Min%"
            nowSheet.Cells(13, nowCol) = "Extend By"
            nowSheet.Cells(14, nowCol) = "Sigma"
            nowSheet.Cells(15, nowCol) = "Data Filter"
            nowSheet.Cells(16, nowCol) = "Disable Max Min"
            nowSheet.Cells(17, nowCol) = "Wafer Seq"
            nowSheet.Cells(18, nowCol) = "Group Lot"
            nowSheet.Cells(19, nowCol) = "Target yValue"
            
            nowSheet.Cells(1, nowCol + 1) = ComboDevice.Text & " W=" & ListWidth.List(iWidth) & " L=" & ListLength.List(iLength)
            nowSheet.Cells(2, nowCol + 1) = "ALL"
            nowSheet.Cells(3, nowCol + 1) = "ALL"
            nowSheet.Cells(4, nowCol + 1) = "SA"
            nowSheet.Cells(5, nowCol + 1) = ComboDevice.Text
            nowSheet.Cells(6, nowCol + 1) = "Linear"
            nowSheet.Cells(7, nowCol + 1) = ""
            nowSheet.Cells(8, nowCol + 1) = ""
            nowSheet.Cells(9, nowCol + 1) = "100"
            nowSheet.Cells(10, nowCol + 1) = "75"
            nowSheet.Cells(11, nowCol + 1) = "25"
            nowSheet.Cells(12, nowCol + 1) = "0"
            nowSheet.Cells(13, nowCol + 1) = "Lot"
            nowSheet.Cells(14, nowCol + 1) = "Median"
            nowSheet.Cells(15, nowCol + 1) = "No"
            nowSheet.Cells(16, nowCol + 1) = "No"
            nowSheet.Cells(17, nowCol + 1) = "Yes"
            nowSheet.Cells(18, nowCol + 1) = "No"
            
            nowSheet.Cells(20, nowCol) = "Y"
            
            nowRow = 20
            For iSA = 0 To ListSA.ListCount - 1
                tmp = ComboDevice.Text & "_" & ListWidth.List(iWidth) & "_" & ListLength.List(iLength) & "_" & ListSA.List(iSA)
                If existInColl(DataColl, tmp) Then
                    nowRow = nowRow + 1
                    nowSheet.Cells(nowRow, nowCol) = DataColl(tmp)
                    nowSheet.Cells(nowRow, nowCol + 1) = ListSA.List(iSA)
                End If
            Next iSA
            End If
        Next iLength
        End If
    Next iWidth


    nowSheet.Cells.Font.Size = 8
    nowSheet.Columns.AutoFit
End Sub

Public Function existInColl(tmpCollection As Collection, key As String) As Boolean
    Dim x As Variant
    On Error Resume Next
    x = tmpCollection.item(key)
    If IsEmpty(x) Or x = "" Then
        existInColl = False
    Else
        existInColl = True
    End If
    On Error GoTo 0
End Function

Private Sub CmdSASelAll_Click()
    Dim i As Integer
    
    For i = 0 To ListSA.ListCount - 1
        ListSA.Selected(i) = True
    Next i
End Sub

Private Sub CmdSASelNone_Click()
    Dim i As Integer
    
    For i = 0 To ListSA.ListCount - 1
        ListSA.Selected(i) = False
    Next i
End Sub

Private Sub CmdWidthSelAll_Click()
    Dim i As Integer
    
    For i = 0 To ListWidth.ListCount - 1
        ListWidth.Selected(i) = True
    Next i
End Sub

Private Sub CmdWidthSelNone_Click()
    Dim i As Integer
    
    For i = 0 To ListWidth.ListCount - 1
        ListWidth.Selected(i) = False
    Next i
End Sub

Private Sub ComboPlot_Click()
    ListWidth.Enabled = True
    ListLength.Enabled = True
    ListSA.Enabled = True
    CmdWidthSelAll.Enabled = True
    CmdWidthSelNone.Enabled = True
    CmdLengthSelAll.Enabled = True
    CmdLengthSelNone.Enabled = True
    CmdSASelAll.Enabled = True
    CmdSASelNone.Enabled = True
    Select Case ComboPlot.Text
        Case "Width":
            ListWidth.Enabled = False
            CmdWidthSelAll.Enabled = False
            CmdWidthSelNone.Enabled = False
        Case "Length":
            ListLength.Enabled = False
            CmdLengthSelAll.Enabled = False
            CmdLengthSelNone.Enabled = False
        Case "SA":
            ListSA.Enabled = False
            CmdSASelAll.Enabled = False
            CmdSASelNone.Enabled = False
    End Select
End Sub



Private Sub UserForm_Initialize()
    Dim DeviceColl As New Collection
    Dim WidthColl As New Collection
    Dim LengthColl As New Collection
    Dim SaColl As New Collection
    
    
    Dim nowSheet As Worksheet
    Dim i As Long
    
    If Not IsExistSheet("Data") Then MsgBox "Please load longfile first.": Exit Sub
    
    Set nowSheet = Worksheets("Data")
    
    On Error Resume Next
    For i = 11 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(i, 2) = "Parameter" And DataColl.Count > 0 Then Exit For
        If nowSheet.Cells(i, 2) <> "" And nowSheet.Cells(i, 2) <> "Parameter" Then
            DeviceColl.Add getCOL(nowSheet.Cells(i, 2), "_", 1), getCOL(nowSheet.Cells(i, 2), "_", 1)
            WidthColl.Add getCOL(nowSheet.Cells(i, 2), "_", 2), getCOL(nowSheet.Cells(i, 2), "_", 2)
            LengthColl.Add getCOL(nowSheet.Cells(i, 2), "_", 3), getCOL(nowSheet.Cells(i, 2), "_", 3)
            SaColl.Add getCOL(nowSheet.Cells(i, 2), "_", 4), getCOL(nowSheet.Cells(i, 2), "_", 4)
            'Err.Clear
            DataColl.Add nowSheet.Cells(i, 2), getCOL(nowSheet.Cells(i, 2), "_", 1) & "_" & getCOL(nowSheet.Cells(i, 2), "_", 2) & _
                "_" & getCOL(nowSheet.Cells(i, 2), "_", 3) & "_" & getCOL(nowSheet.Cells(i, 2), "_", 4)
        End If
        If Err.Number > 0 Then
            'Debug.Print "debug> " & nowSheet.Cells(i, 2) & ":" & getCOL(nowSheet.Cells(i, 2), "_", 1) & "_" & getCOL(nowSheet.Cells(i, 2), "_", 2) & _
                "_" & getCOL(nowSheet.Cells(i, 2), "_", 3) & "_" & getCOL(nowSheet.Cells(i, 2), "_", 4)
            'Stop
        End If
    Next i
    On Error GoTo 0
    
    For i = 1 To DeviceColl.Count
        ComboDevice.AddItem DeviceColl(i)
    Next i
    ComboDevice.ListIndex = 0
    For i = 1 To WidthColl.Count
        ListWidth.AddItem WidthColl(i)
    Next i
    For i = 1 To LengthColl.Count
        ListLength.AddItem LengthColl(i)
    Next i
    For i = 1 To SaColl.Count
        ListSA.AddItem SaColl(i)
    Next i
    
    ComboPlot.AddItem "Width"
    ComboPlot.AddItem "Length"
    ComboPlot.AddItem "SA"
    ComboPlot.ListIndex = 0
    
    If Not CheckSA Then
        Label4.Visible = False
        ListSA.Visible = False
        CmdSASelAll.Visible = False
        CmdSASelNone.Visible = False
        Me.width = 260
    End If
    
    'Stop
    'MsgBox DataColl.Count
End Sub
