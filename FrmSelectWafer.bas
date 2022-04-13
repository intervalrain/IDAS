
Private Function putWafer()
   Dim i As Long
   
   ListActive.Clear
   ListNonActive.Clear
   ListSequence.Clear
   On Error GoTo myError
   For i = 0 To UBound(WaferArray, 2)
      If WaferArray(1, i) <> "NO" Then
         ListActive.AddItem Trim(WaferArray(0, i))
      Else
         ListNonActive.AddItem Trim(WaferArray(0, i))
      End If
      ListSequence.AddItem Trim(WaferArray(0, i))
   Next i
   Exit Function
myError:
   If Not IsExistSheet("Data") Then MsgBox "Please load longfile before the operation.": Exit Function
   Call InitStep
   Resume
End Function

Private Sub CmdAdd_Click()
   Dim i As Integer
   
   For i = 0 To ListNonActive.ListCount - 1
      If ListNonActive.Selected(i) Then
         ListActive.AddItem ListNonActive.List(i)
      End If
   Next i
   For i = ListNonActive.ListCount - 1 To 0 Step -1
      If ListNonActive.Selected(i) Then
         ListNonActive.RemoveItem (i)
      End If
   Next i

End Sub

Private Sub CmdAddAll_Click()
   Dim i As Integer
   
   For i = 0 To ListNonActive.ListCount - 1
      ListActive.AddItem ListNonActive.List(i)
   Next i
   For i = ListNonActive.ListCount - 1 To 0 Step -1
      ListNonActive.RemoveItem (i)
   Next i

End Sub



Private Sub CmdLoad_Click()
    Dim mList() As String
    Dim i As Long, j As Long

    mList = GetHiddenOption("WaferSeq")
    If mList(0) = "" Then Exit Sub
    
    'Remove items which in setting
    
    For j = UBound(mList) To 0 Step -1
        For i = ListSequence.ListCount - 1 To 0 Step -1
            If ListSequence.List(i) = mList(j) Then
                ListSequence.RemoveItem (i)
                ListSequence.AddItem mList(j), 0
                Exit For
            End If
        Next i
    Next j

End Sub

Private Sub cmdOK_Click()
   Dim i As Long
   Dim j As Long
   Dim tmpStr As String
   Dim tmpAry
   'modify sequence
   For i = 0 To ListSequence.ListCount - 1
      WaferArray(0, i) = ListSequence.List(i)
   Next i
   'Active
   For i = 0 To ListActive.ListCount - 1
      For j = 0 To UBound(WaferArray, 2)
         If WaferArray(0, j) = ListActive.List(i) Then WaferArray(1, j) = ""
      Next j
   Next i
   'Non-Active
   For i = 0 To ListNonActive.ListCount - 1
      For j = 0 To UBound(WaferArray, 2)
         If WaferArray(0, j) = ListNonActive.List(i) Then WaferArray(1, j) = "NO"
      Next j
   Next i
   
   
   
   Unload Me
End Sub

Private Sub CmdRemove_Click()
   Dim i As Integer
   
   For i = 0 To ListActive.ListCount - 1
      If ListActive.Selected(i) Then
         ListNonActive.AddItem ListActive.List(i)
      End If
   Next i
   For i = ListActive.ListCount - 1 To 0 Step -1
      If ListActive.Selected(i) Then
         ListActive.RemoveItem (i)
      End If
   Next i
End Sub

Private Sub CmdRemoveAll_Click()
   Dim i As Integer
   
   For i = 0 To ListActive.ListCount - 1
      ListNonActive.AddItem ListActive.List(i)
   Next i
   For i = ListActive.ListCount - 1 To 0 Step -1
      ListActive.RemoveItem (i)
   Next i
End Sub

Private Sub CmdSave_Click()
    Dim mList() As String
    Dim i As Long
    
    ReDim mList(ListSequence.ListCount - 1)
    For i = 0 To ListSequence.ListCount - 1
        mList(i) = ListSequence.List(i)
    Next i
    Call SetHiddenOption("WaferSeq", mList)
End Sub

Private Sub CmdUp_Click()
   Dim i As Long
   Dim tmpStr As String
   Dim tmpYN As Boolean
   
   If ListSequence.Selected(0) Then Exit Sub
   For i = 1 To ListSequence.ListCount - 1
      If ListSequence.Selected(i) Then
         tmpStr = ListSequence.List(i - 1)
         ListSequence.List(i - 1) = ListSequence.List(i)
         ListSequence.List(i) = tmpStr
         tmpYN = ListSequence.Selected(i - 1)
         ListSequence.Selected(i - 1) = ListSequence.Selected(i)
         ListSequence.Selected(i) = tmpYN
      End If
   Next i
End Sub

Private Sub CmdDown_Click()
   Dim i As Long
   Dim tmpStr As String
   Dim tmpYN As Boolean
   
   If ListSequence.Selected(ListSequence.ListCount - 1) Then Exit Sub
   For i = ListSequence.ListCount - 2 To 0 Step -1
      If ListSequence.Selected(i) Then
         tmpStr = ListSequence.List(i + 1)
         ListSequence.List(i + 1) = ListSequence.List(i)
         ListSequence.List(i) = tmpStr
         tmpYN = ListSequence.Selected(i + 1)
         ListSequence.Selected(i + 1) = ListSequence.Selected(i)
         ListSequence.Selected(i) = tmpYN
      End If
   Next i
End Sub

Private Sub UserForm_Initialize()
   Call putWafer
End Sub
