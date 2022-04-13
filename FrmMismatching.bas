
Option Explicit

Private Sub CB_PreScreen_Change()
    If CB_PreScreen.Value = True Then
        TB_HiSpec.Visible = True
        TB_LoSpec.Visible = True
        Label3.Visible = True
        Label4.Visible = True
    Else
        TB_HiSpec.Visible = False
        TB_LoSpec.Visible = False
        Label3.Visible = False
        Label4.Visible = False
    End If
End Sub

Private Sub TB_Preshrink_Change()

End Sub

Private Sub UserForm_Activate()
    Call getItemFromData
    TB_FtTimes.Text = 6
    TB_HiSpec.Text = 1
    TB_LoSpec.Text = 0
End Sub

Private Function getItemFromData()
    Dim nowSheet As Worksheet
    Dim i As Long
    Dim collItem As New Collection
    Dim ynBody As Boolean
    
    If Not IsExistSheet("Data") Then MsgBox "Please load longfile before the operation.": Exit Function
    
    On Error Resume Next
    Set nowSheet = Worksheets("Data")
    For i = 1 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(i, 2) = "Parameter" And ynBody = True Then Exit For
        If nowSheet.Cells(i, 2) = "Parameter" Then ynBody = True
        
        If ynBody And nowSheet.Cells(i, 2) <> "Parameter" And nowSheet.Cells(i, 2) <> "" Then
            collItem.Add getCOL(nowSheet.Cells(i, 2), "_", 1), getCOL(nowSheet.Cells(i, 2), "_", 1)
        End If
    Next i
    On Error GoTo 0
    
    ListItemData.Clear
    For i = 1 To collItem.Count
        ListItemData.AddItem collItem(i)
    Next i

End Function

Private Sub CmdAdd_Click()
    Dim i As Integer
   
    For i = 0 To ListItemData.ListCount - 1
        If ListItemData.Selected(i) Then
            ListItem.AddItem ListItemData.List(i)
        End If
    Next i
    For i = ListItemData.ListCount - 1 To 0 Step -1
        If ListItemData.Selected(i) Then
            ListItemData.RemoveItem (i)
        End If
    Next i
End Sub

Private Sub CmdAddAll_Click()
    Dim i As Integer
   
    For i = 0 To ListItemData.ListCount - 1
        ListItem.AddItem ListItemData.List(i)
    Next i
    For i = ListItemData.ListCount - 1 To 0 Step -1
        ListItemData.RemoveItem (i)
    Next i
End Sub

Private Sub CmdRemove_Click()
    Dim i As Integer

    For i = 0 To ListItem.ListCount - 1
        If ListItem.Selected(i) Then
            ListItemData.AddItem ListItem.List(i)
        End If
    Next i
    For i = ListItem.ListCount - 1 To 0 Step -1
        If ListItem.Selected(i) Then
            ListItem.RemoveItem (i)
        End If
    Next i
End Sub

Private Sub CmdRemoveAll_Click()
    Dim i As Integer
   
    For i = 0 To ListItem.ListCount - 1
        ListItemData.AddItem ListItem.List(i)
    Next i
    For i = ListItem.ListCount - 1 To 0 Step -1
        ListItem.RemoveItem (i)
    Next i
End Sub

Private Sub CmdAddItem_Click()
    Dim i As Integer
    
    For i = 0 To ListItemData.ListCount - 1
        If ListItemData.Selected(i) Then ListItem.AddItem ListItemData.List(i)
    Next i
End Sub

Private Sub CmdRun_Click()
    Dim i As Integer
    Dim item
    If ListItem.ListCount = 0 Then
        Me.Hide: MsgBox ("Please select at least 1 item.")
        ErrFlag = True
        Exit Sub
    End If
    ReDim ActiveItems(ListItem.ListCount - 1) As String
    For i = 0 To ListItem.ListCount - 1
        ActiveItems(i) = ListItem.List(i)
    Next i
    
    If CB_PreScreen = False Then
        TB_HiSpec.Text = ""
        TB_LoSpec.Text = ""
    Else
        mHigh = TB_HiSpec
        mLow = TB_LoSpec
    End If
    
    PreScreen = CB_PreScreen
    MergeTSK = CB_MergeTSK
    filterTimes = TB_FtTimes
    ScreenPair = CB_ScreenPair
    Preshrink = TB_Preshrink
    Unload Me
End Sub

