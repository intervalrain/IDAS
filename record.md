    Dim SpecDict As Object
    Set SpecDict = CreateObject("Scripting.Dictionary")
    Set DataDict = CreateObject("Scripting.Dictionary")
    
    Call setDict("SPEC", 3, SpecDict, Worksheets("SPEC").UsedRange, True)
    Call setDict("Data", 2, DataDict, Range(Worksheets("Data").Names(1)), True)

' 利用 HashTable 的概念對不同的 parameter 先做一次編碼
  一次迴圈 --> O(n) n = SPEC 的列數 or 量測的 parameter 數
Public Function setDict(ByVal sheetName As String, ByVal Target As Integer, ByRef Dict As Object, ByVal mRange As Range, Optional ByVal byRows As Boolean = True)
    
    Dim nowSheet As Worksheet
    If Not IsExistSheet(sheetName) Then Exit Function
    Set nowSheet = Worksheets(sheetName)
    
    Dim i As Long
    Dim n As Long

    On Error Resume Next
    
    If byRows = True Then
        For i = 1 To mRange.Rows.Count
            If Not Trim(mRange.Cells(i, Target).Value) = "" Then
                Dict.Add mRange.Cells(i, Target).Value, i
            End If
        Next i
    Else
        For i = 1 To mRange.Columns.Count
            If Not Trim(mRange.Cells(Target, i).Value) = "" Then
                Dict.Add mRange.Cells(Target, i).Value, i
            End If
        Next i
    End If
    
End Function

' 直接將 parameter 做 hash 後對 hashtable 做查值 --> 近乎 O(1)
Public Function getSPECByPara(ByVal nowPara As String, ByVal n As specColumn, Optional sheetName As String = "SPECTEMP")

    Dim reValue
    Dim nowRange As Range
    Dim TargetSheet As Worksheet
   
    If Left(nowPara, 1) = "'" Then nowPara = Mid(nowPara, 2)
   
    Set TargetSheet = Worksheets(sheetName)
    Set nowRange = TargetSheet.UsedRange
    On Error Resume Next
    reValue = TargetSheet.Cells(SpecDict(nowPara), n)
    If Not IsEmpty(reValue) Then
        If Trim(reValue) = "" Then Set reValue = Nothing
    End If
    getSPECByPara = reValue
End Function

' 直接將 parameter 做 hash 後對 hashtable 做查值 --> 近乎 O(1)
Public Function getRangeByPara(nowWafer As String, nowPara As String, Optional dieNum As Integer = 0)
    Dim nowRow As Long
    Dim nowRange As Range
   
    Set nowRange = Worksheets("Data").Range("wafer_" & nowWafer)
    Set getRangeByPara = Nothing
    
    If DataDict.Exists(nowPara) Then
        nowRow = DataDict(nowPara)
        Set getRangeByPara = nowRange.Range(N2L(4) & CStr(nowRow) & ":" & N2L(dieNum + 3) & CStr(nowRow))
    End If
    
End Function
Pros: Better speed
Cons: Could not revise raw data
