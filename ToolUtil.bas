    Option Explicit
    
    Private mlCalcStatus As Long
    Private mbInSpeed As Boolean
    
    Public Sub WriteTextFile(ByVal Filename As String, ByVal mStr As String, Optional ByVal Override As Boolean = False)
        Dim FS As New Scripting.FileSystemObject
        Dim f As TextStream
        Dim Iomode As Integer
        
        Iomode = 2                              'ForAppending
        If Override = True Then Iomode = 8      'ForWriting
        Set f = FS.OpenTextFile(Filename, 8, True)
        f.Write mStr
        f.Close
        Set FS = Nothing
        Set f = Nothing
    End Sub
    
    Public Function ReadTextFile(mFileName As String)
        Dim FSO As New Scripting.FileSystemObject
        Dim f As TextStream
        
        Set f = FSO.OpenTextFile(mFileName, 1)   'ForReading
        ReadTextFile = f.ReadAll
        f.Close
        Set FSO = Nothing
        Set f = Nothing
    End Function
    
    Public Function GetFileHeader(mFileName As String, mLine)
        Dim FileID As Long
        Dim temp As String, tempStr
        Dim i As Long
        
        FileID = FreeFile
        Open mFileName For Input As #FileID
             For i = 1 To mLine
                 If Not EOF(FileID) Then Line Input #FileID, temp: tempStr = tempStr & temp & vbCrLf
             Next i
        Close #FileID
        GetFileHeader = tempStr
    End Function
    
    Public Function getURL(mURL As String)
        On Error Resume Next
        
        Dim reTxt As String
        Dim xmlHttp As Object
        
        Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
        xmlHttp.Open "GET", mURL, False, "", ""
        xmlHttp.Send
        reTxt = Trim(xmlHttp.responseText)
        Set xmlHttp = Nothing
        
        getURL = reTxt
        
    End Function
    
    Public Function IsKey(mStr As String, mKey As String, Optional ByVal mSep As String = "&", Optional ByVal IsBoolean As Boolean = True)
        Dim tempA
        Dim i As Integer
       
        If IsBoolean Then
            IsKey = False
        Else
            IsKey = 0
        End If
        tempA = Split(mStr, mSep)
        For i = 0 To UBound(tempA)
            If Trim(UCase(tempA(i))) = Trim(UCase(mKey)) Then
                If IsBoolean Then
                    IsKey = True
                Else
                    IsKey = i + 1
                End If
                Exit For
            End If
        Next i
    End Function
    
    Public Sub Speed()
        On Error Resume Next
        Application.ScreenUpdating = False
        Application.DisplayStatusBar = False
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    End Sub
    
    Public Sub Unspeed()
        On Error Resume Next
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
    End Sub
    
    Public Function IsPrefix(mStr As String, mKey As String) As Boolean
        If UCase(Left(mStr, Len(mKey))) = UCase(mKey) Then
            IsPrefix = True
        Else
            IsPrefix = False
        End If
    End Function
    
    Public Function getCOL(ByVal mStr As String, ByVal mDel As String, ByVal mNum As Long)
        ' Get the exact column string
        Dim TempAry
        TempAry = Split(mStr, mDel)
        If mNum - 1 > UBound(TempAry) Or mNum < 1 Then
            getCOL = ""
        Else
            getCOL = Trim(TempAry(mNum - 1))
        End If
    End Function
    
    Public Function getCOL2(ByVal mStr As String, ByVal mDel As String, ByVal mNum As Long)
        ' Get the string before the indexed column
        Dim TempAry
        Dim i As Integer
       
        TempAry = Split(mStr, mDel)
        If mNum - 1 > UBound(TempAry) Or mNum < 1 Then
            getCOL2 = ""
        Else
            For i = 0 To mNum - 1
            getCOL2 = getCOL2 & mDel & Trim(TempAry(i))
            Next i
        End If
        getCOL2 = Mid(getCOL2, 2)
    End Function
    
    Public Function getCOL3(ByVal mStr As String, ByVal mDel As String, ByVal mStart As Long, Optional ByVal mEnd As Long = -1)
        ' Get the string between the indexed column(default the end is last of string)
        Dim TempAry
        Dim i As Integer
        
        TempAry = Split(mStr, mDel)
        If mEnd = -1 Then mEnd = UBound(TempAry) + 1
        
        If mStart - 1 > UBound(TempAry) Or mStart < 1 Then
            getCOL3 = ""
        Else
            For i = mStart - 1 To mEnd - 1
            getCOL3 = getCOL3 & mDel & Trim(TempAry(i))
            Next i
        End If
        getCOL3 = Mid(getCOL3, 2)
    End Function
    
    
    Public Function N2L(ByVal iNumber As Long)
        Dim iLetter As String
        Dim UpInt As Integer
        
        If iNumber = 0 Then N2L = "": Exit Function
        
        UpInt = (iNumber - 1) \ 26
        If UpInt > 0 Then iLetter = Chr(UpInt + 64)
        iLetter = iLetter & Chr(iNumber - UpInt * 26 + 64)
        N2L = iLetter
    End Function
    
    Public Function AddSheet(sheetName As String, Optional delOld As Boolean = True, Optional mSheet As String)
        Dim nowSheet As Worksheet
        Dim i As Integer
        If IsExistSheet(sheetName) Then
            If delOld = False Then
                Set AddSheet = Worksheets(sheetName)
                Exit Function
            ElseIf delOld = True Then
                Application.DisplayAlerts = False
                Application.Worksheets(sheetName).Delete
            End If
        End If
        On Error GoTo Err
        If mSheet = "" Then
            Set nowSheet = Worksheets.Add(, Worksheets(Worksheets.Count))
        Else
            Set nowSheet = Worksheets.Add(, Worksheets(mSheet))
        End If
        nowSheet.Name = sheetName
        Application.DisplayAlerts = True
        Set AddSheet = nowSheet
        Set nowSheet = Nothing
    Exit Function
Err:
        mSheet = Worksheets(Worksheets.Count).Name
        Resume
    End Function
    
    Public Function DelSheet(sheetName As String)
        Dim i As Integer
        Application.DisplayAlerts = False
        For i = 1 To Worksheets.Count
            If Worksheets(i).Name = sheetName Then Worksheets(i).Delete: Exit For
        Next
        Application.DisplayAlerts = True
    End Function
    
    Public Function IsExistSheet(sheetName As String)
        Dim i As Integer
       
        For i = 1 To Worksheets.Count
            If UCase(Worksheets(i).Name) = UCase(sheetName) Then IsExistSheet = True: Exit Function
        Next
        IsExistSheet = False
    End Function
    
    Public Function TrimCol(sheetName As String, numCol As Long)
        Dim i As Long
        Dim nowSheet As Worksheet
       
        Set nowSheet = Worksheets(sheetName)
        nowSheet.Cells.MergeCells = False
        For i = 1 To nowSheet.UsedRange.Rows.Count
            If Left(nowSheet.Cells(i, numCol), 1) = " " Or Right(nowSheet.Cells(i, numCol), 1) = " " Then _
            nowSheet.Cells(i, numCol) = Trim(nowSheet.Cells(i, numCol))
        Next
        Set nowSheet = Nothing
        TrimCol = True
    End Function
    
    Public Function TrimSheet(nowSheet As String)
        Dim i As Long
        Dim YNEmpty As Boolean
       
        With Worksheets(nowSheet)
            .AutoFilterMode = False
            YNEmpty = True
            Do While YNEmpty
                For i = 1 To .UsedRange.Columns.Count
                    If Trim(.Cells(.UsedRange.Rows.Count, i)) <> "" Then YNEmpty = False: Exit For
                Next i
                If YNEmpty Then .Rows(.UsedRange.Rows.Count).Delete
            Loop
            YNEmpty = True
            Do While YNEmpty
                For i = 1 To .UsedRange.Rows.Count
                    If Trim(.Cells(i, .UsedRange.Columns.Count)) <> "" Then YNEmpty = False: Exit For
                Next i
                If YNEmpty Then .Columns(.UsedRange.Columns.Count).Delete
            Loop
        End With
    End Function
    
    Public Function TrimSpecList()
        Dim nowSheet As Worksheet
        Dim i As Long, j As Long
        Dim tempA
    
        tempA = Array("\", "/", "?", "[", "]")
        Set nowSheet = Worksheets("SPEC_List")
        For i = 1 To nowSheet.UsedRange.Rows.Count
            For j = 1 To nowSheet.UsedRange.Columns.Count
                If Left(nowSheet.Cells(i, j), 1) = " " Or Right(nowSheet.Cells(i, j), 1) = " " Then _
                    nowSheet.UsedRange.Cells(i, j) = Trim(nowSheet.UsedRange.Cells(i, j))
            Next j
        Next i
        For i = 1 To nowSheet.UsedRange.Columns.Count
            For j = 0 To UBound(tempA)
                nowSheet.UsedRange.Cells(1, i) = Replace(nowSheet.UsedRange.Cells(1, i), tempA(j), "_")
            Next j
        Next i
        
        Set nowSheet = Nothing
        
    End Function
    
    Public Function getValue(SourceStr As String, SplitChar As String, ValueName As String, ValueChar As String)
        Dim tempA, tempB
        Dim i As Long
        tempA = Split(SourceStr, SplitChar)
        For i = 0 To UBound(tempA)
            If InStr(1, tempA(i), ValueChar) <> 0 Then
                tempB = Split(tempA(i), ValueChar)
                If UCase(Trim(tempB(0))) = UCase(ValueName) Then getValue = Trim(tempB(1)): Exit Function
            End If
        Next
        getValue = ""
    End Function
    Public Function getValueByKey(mSheet As String, item As String, mCol As Integer)
    
    Dim nowSheet As Worksheet
    Dim nowRow As Integer
    Dim tmpStr As String
    
    Set nowSheet = Worksheets(mSheet)
    nowRow = 1
    item = Trim(item)
    
    Do
        tmpStr = Trim(nowSheet.Cells(nowRow, 1).Value)
        If UCase(tmpStr) = UCase(item) Then
            getValueByKey = nowSheet.Cells(nowRow, mCol).Value: Exit Function
        End If
        nowRow = nowRow + 1
    Loop Until tmpStr = "" Or UCase(tmpStr) = "END"
        
    nowRow = 1
        
    Do
        tmpStr = Trim(nowSheet.Cells(nowRow, 1).Value)
        If UCase(getCOL(tmpStr, "*", 1)) = UCase(Left(item, Len(getCOL(tmpStr, "*", 1)))) Then
            getValueByKey = nowSheet.Cells(nowRow, mCol).Value: Exit Function
        End If
        nowRow = nowRow + 1
    Loop Until tmpStr = "" Or UCase(tmpStr) = "END"
    
    End Function
    Public Function myMedian(valueList As String)
        Dim MedianValue
        Dim tempA
        Dim i As Integer, j As Integer
        Dim temp
       
        Do
            valueList = Replace(valueList, ",,", ",")
        Loop Until InStr(valueList, ",,") = 0
    
        If Left(valueList, 1) = "," Then valueList = Mid(valueList, 2)
        If Right(valueList, 1) = "," Then valueList = Left(valueList, Len(valueList) - 1)
        If valueList = "" Then Exit Function: myMedian = 0
       
        tempA = Split(valueList, ",")
        For i = 0 To UBound(tempA) - 1
            For j = 0 To UBound(tempA) - 1
                If Val(tempA(j)) > Val(tempA(j + 1)) Then Call Swap(tempA(j), tempA(j + 1))
            Next j
        Next i
        If UBound(tempA) Mod 2 = 0 Then
            MedianValue = tempA(UBound(tempA) / 2)
        Else
            MedianValue = (Val(tempA(UBound(tempA) \ 2)) + Val(tempA(UBound(tempA) \ 2 + 1))) / 2
        End If
        myMedian = MedianValue
    End Function
    
    Public Function Log10(x)
        Log10 = Log(x) / Log(10)
    End Function
    
    Public Function Swap(ByRef mA, ByRef mB)
        Dim tmp
        tmp = mA
        mA = mB
        mB = tmp
    End Function
    
    Public Function ynInCorner(pArray(), ByVal x As Double, ByVal y As Double) As Boolean
        Dim i As Integer
        Dim d1 As Double, d2 As Double
        Dim dA() As Double
        Dim tmpValue As Double
        ReDim dA(UBound(pArray))
       
        ynInCorner = True
        For i = 0 To UBound(pArray)
            tmpValue = 0
            Select Case (pArray(i)(0) - x)
                Case 0:
                    Select Case (pArray(i)(1) - y)
                        Case 0:  ynInCorner = True: Exit Function
                        Case Is > 0: dA(i) = 90
                        Case Is < 0: dA(i) = 270
                    End Select
                Case Is < 0:
                    tmpValue = 180
            End Select
            If (pArray(i)(0) - x) <> 0 Then
                dA(i) = Atn((pArray(i)(1) - y) / (pArray(i)(0) - x)) * 180 / pi
            End If
            dA(i) = (dA(i) + 360 + tmpValue) Mod 360
        Next i
        For i = 0 To UBound(dA) - 1
            If (dA(i + 1) - dA(i) + 360) Mod 360 > 180 Then ynInCorner = False: Exit For
        Next i
    End Function
    
    Public Function FillSpace(ByVal SrcStr As String, ByVal sNum As Integer, Optional ynBefore As Boolean = False)
        If Len(SrcStr) < sNum Then
            If ynBefore Then
                FillSpace = Space(sNum - Len(SrcStr)) & SrcStr
            Else
                FillSpace = SrcStr & Space(sNum - Len(SrcStr))
            End If
        Else
            FillSpace = SrcStr
            End If
    End Function
    
    Public Function Coordinate(ByVal WaferMap As Range, ByVal siteNum As String, Optional ByVal waferNo = "", Optional order As String)
    
        Dim mRow As Integer
        Dim mCol As Integer
        Dim StrWf As String
        Dim x As String
        Dim y As String
        
        If order = "" Then order = siteNum
        If Not waferNo = "" Then StrWf = "<" & CStr(waferNo) & "-" & CStr(order) & ">"
        
        For mRow = 2 To WaferMap.Rows.Count
            For mCol = 2 To WaferMap.Columns.Count
                If WaferMap(mRow, mCol).Value = siteNum Then Exit For
            Next mCol
            If WaferMap(mRow, mCol).Value = siteNum Then Exit For
        Next mRow
        
        x = WaferMap.Cells(1, mCol).Value
        y = WaferMap.Cells(mRow, 1).Value
        
        Coordinate = StrWf & "(" & x & "," & y & ")"
    
    End Function
    
    Public Function CopySheet(inSheet As String, outSheet As String)
        
        Dim nowSheet As Worksheet, newSheet As Worksheet
              
        Set nowSheet = Worksheets(inSheet)
        Set newSheet = AddSheet(outSheet)
        
        nowSheet.Activate
        nowSheet.Cells.Copy
        newSheet.Activate
        newSheet.Paste
        ActiveWindow.Zoom = 75
        newSheet.Range("A3").Select
        ActiveWindow.FreezePanes = True
        
        Set nowSheet = Nothing
        Set newSheet = Nothing
        
    End Function
    
    Public Function AbsMax(ByVal mRange As Range)
        Dim Max(1) As Double
        Dim tmpAry
        Dim item
        Dim Sign As Single
        tmpAry = mRange
        For Each item In tmpAry
            If item = 0 Then
                Sign = 1
            Else
                Sign = item / (Abs(item))
            End If
            item = Abs(item)
            If item > Max(1) Then Max(1) = item: Max(0) = Sign
        Next item
        AbsMax = Application.Match(Max(0) * Max(1), tmpAry, 0)
        
    End Function
    
    Public Function countInStr(Optional Start, Optional String1, Optional String2, Optional Compare As VbCompareMethod = vbBinaryCompare)
        Dim length As Integer
        Dim i As Integer
        Dim mCount As Integer
        Dim tmpStr As String
        
        If IsError(String2) = True Then
            String2 = String1
            String1 = Start
            Start = 1
        End If
        length = Len(String1)
        For i = Start To length
            tmpStr = Mid(String1, i, Len(String2))
            If InStr(tmpStr, String2) Then mCount = mCount + 1
        Next i
        
        countInStr = mCount
    
    End Function
    
    Public Function sortByColumn(ByRef mCol As Integer, Optional isAscending As Boolean = True)
        
        Dim nowSheet As Worksheet
        Dim i As Integer
        Dim optAsc
        
        Set nowSheet = ActiveSheet
        
        On Error GoTo Err
        
        nowSheet.Rows(1).AutoFilter
        
        nowSheet.Sort.SortFields.Clear
        
        If isAscending Then
            optAsc = xlAscending
        Else
            optAsc = xlDescending
        End If
        nowSheet.Sort.SortFields.Add _
            key:=Columns(mCol), _
            SortOn:=xlSortOnValues, _
            order:=optAsc, _
            DataOption:=xlSortNormal
        
        With nowSheet.Sort
            .setRange nowSheet.Cells(1, 1).CurrentRegion
            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        nowSheet.Rows(1).AutoFilter
    Exit Function
    
Err:
        nowSheet.Rows(1).AutoFilter
        Resume
    
    End Function
    
    Public Function trimFunc(mStr As String, mFunc As String)
        
        Dim tmpStr
        tmpStr = Split(mStr, mFunc & "(")
        
        If UBound(tmpStr) <> 1 Then trimFunc = mStr: Exit Function
        
        mStr = tmpStr(1)
        mStr = getCOL(mStr, ")", 1)
        trimFunc = mStr
        
    End Function
    
    Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    
        Dim tmpStrToBeFound As String
        Dim tmpStr As String
        tmpStrToBeFound = "^" & stringToBeFound & "^"
        tmpStr = "^" & Join(arr, "^,^") & "^"
        IsInArray = IIf(InStr(tmpStr, tmpStrToBeFound) > 0, True, False)
        
    End Function
