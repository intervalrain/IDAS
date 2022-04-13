Option Explicit


Public Sub getSheet()

    Dim i As Long, j As Long, k As Long
    Dim xR As Integer, xC As Integer
    Dim nowSheet As Worksheet
    Dim mCount As Long
    Dim mType As String
    
    
    Set nowSheet = Worksheets("PPT")
    j = 12
    k = 1
    mType = nowSheet.Cells(7, 1).Text
    nowSheet.Range(Rows(13), Rows(nowSheet.UsedRange.Rows.Count)).ClearContents
    Range("E2").Value = "=TODAY()"
    mCount = 9
    
    If IsExistSheet("All_Chart") Then
        
        Select Case UCase(mType)
            Case "A"
                mCount = 1
            Case "B"
                mCount = 2
            Case "C"
                mCount = 4
            Case "D"
                mCount = 6
            Case "E"
                mCount = 9
            Case Else
                If IsNumeric(mType) Then
                    mCount = CInt(mType)
                    If mType < 1 Or mType > 9 Then
                        MsgBox ("Please enter a layout type between A to E or valid int between 1 to 9.")
                        Exit Sub
                    ElseIf mType = 1 Then
                        mType = "A"
                    ElseIf mType = 2 Then
                        mType = "B"
                    ElseIf mType <= 4 Then
                        mType = "C"
                    ElseIf mType <= 6 Then
                        mType = "D"
                    ElseIf mType <= 9 Then
                        mType = "E"
                    End If
                Else
                    Exit Sub
                End If
        End Select
            
        Dim cNum As Integer
        Dim cRow As Integer
        
        cNum = Worksheets("All_Chart").ChartObjects.Count
        If Not cNum Mod mCount = 0 Then
            cRow = Int(cNum / mCount) + 1
        Else
            cRow = cNum / mCount
        End If
            
        For xR = 1 To cRow
            nowSheet.Cells(j + xR, 1).Value = "YES"
            nowSheet.Cells(j + xR, 2).Value = Worksheets("All_Chart").Name
            nowSheet.Cells(j + xR, 3).Value = "=B" & j + xR
            nowSheet.Cells(j + xR, 4).Value = UCase(mType)
            nowSheet.Cells(j + xR, 5).Value = j + xR - 11
            For xC = 1 To mCount
                nowSheet.Cells(j + xR, 6 + xC).Value = Worksheets("All_Chart").ChartObjects(k).Chart.ChartTitle.Caption
                On Error Resume Next
                k = k + 1
            Next
        Next
        j = j + cRow
        k = 1
        
    End If
End Sub
