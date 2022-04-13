Option Explicit

Public Function Load_SingleLongFile(Filename As String)

    Dim SubName As String

    If Not Filename = "" Then
        SubName = UCase(Mid(Filename, InStrRev(Filename, ".")))
        If getFileLine(Filename) > 1048576 Then
            Err.Raise 999, , "File line over 1048576!!!"
            Call Unspeed
            Exit Function
        End If
        Select Case SubName
            Case ".RPT"
                Call LoadLongFileRPT(Filename)
            Case ".LONG"
                Call LoadLongFileLONG(Filename)
        End Select
        Call InitStep
    End If
End Function
Public Function FileDialog(Optional FileType As String = "RPT,long")
    Dim TempAry
    Dim i As Long
    Dim FilterStr As String
    Dim tmpStr As String
    
    TempAry = Split(FileType, ",")
    For i = 0 To UBound(TempAry)
        FilterStr = FilterStr & "," & TempAry(i) & " File(*." & TempAry(i) & "),*." & TempAry(i)
    Next i
    FilterStr = FilterStr & "," & "All File(*.*),*.*"
    FilterStr = Mid(FilterStr, 2)
   
    tmpStr = Application.GetOpenFilename(FilterStr, 1, "Open File", "Open", False)
    If UCase(tmpStr) = "FALSE" Then
        FileDialog = ""
    Else
        FileDialog = tmpStr
    End If
End Function

Public Function getFileLine(mFile As String)
    Dim FSO As New FileSystemObject
    Dim f As TextStream
   
    Set f = FSO.OpenTextFile(mFile, ForReading)
    f.ReadAll
    getFileLine = f.Line
    Set f = Nothing
    Set FSO = Nothing
End Function


Public Function LoadLongFileRPT(ByVal mFileName As String)
    Dim nowSheet As Worksheet
    Dim i As Integer
    Dim siteNum As Integer
    Dim yn12i As Boolean
    
    Set nowSheet = AddSheet("Data")
    nowSheet.Activate
    With nowSheet.QueryTables.Add(Connection:="TEXT;" & mFileName, Destination:=Range("A1"))
        .Name = "IDAS"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = xlWindows
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1)
        .Refresh BackgroundQuery:=False
    End With

    If Trim(nowSheet.Cells(4, 1)) <> "" Then nowSheet.Rows(4).Delete: yn12i = True
    For i = nowSheet.UsedRange.Columns.Count To 6 Step -1
        If Trim(nowSheet.Cells(5, i).Value) = "" Then nowSheet.Columns(i).Delete
    Next i
    Call ConvertRPT
    Set nowSheet = Nothing
End Function

Public Function ConvertRPT()
    Dim mProduct_ID As String
    Dim mLot_ID As String
    Dim mTest_Plan_ID As String
    Dim mDateTime As String
    Dim i As Long
    Dim mRow As Long, mCol As Long
    Dim mWaferID As String
    Dim nowSheet As Worksheet
    Dim tmpStr As String
    
    Set nowSheet = Worksheets("Data")
    
    mProduct_ID = getValue(Cells(3, 1), "   ", "TYPE", "=")
    mLot_ID = getValue(Cells(3, 1), "   ", "LOT", "=")
    mTest_Plan_ID = getValue(Cells(3, 1), "   ", "Recipe", "=")
    mDateTime = getValue(Cells(3, 1), "   ", "DATE", "=")
    
    nowSheet.Activate
    nowSheet.Columns(1).Delete
    nowSheet.Rows(1).Select
    
    For i = 1 To 6
        Selection.Insert Shift:=xlDown
    Next
    
    nowSheet.Cells(1, 1) = "<Process_ID>"
    nowSheet.Cells(2, 1) = "<Product_ID>"
    nowSheet.Cells(3, 1) = "<Lot_ID>"
    nowSheet.Cells(4, 1) = "<Test_Plan_ID>"
    nowSheet.Cells(5, 1) = "<Limit_File>"
    nowSheet.Cells(6, 1) = "<Date/Time>"
    nowSheet.Cells(7, 1) = "( LONG REPORT )"
    nowSheet.Cells(8, 1) = "-------------"
    nowSheet.Cells(9, 1) = "TYPE_SCALAR"
    nowSheet.Cells(10, 1) = "-------------"
    
    nowSheet.Cells(1, 2) = ":" & "x"
    nowSheet.Cells(2, 2) = ":" & mProduct_ID
    nowSheet.Cells(3, 2) = ":" & mLot_ID
    nowSheet.Cells(4, 2) = ":" & mTest_Plan_ID
    nowSheet.Cells(5, 2) = ":" & "x"
    nowSheet.Cells(6, 2) = ":" & mDateTime
    nowSheet.Cells(7, 2) = ":" & "x"
    
    If Not IsExistSheet("CoorSetting") Then Call GenCoorSheet
    Worksheets("CoorSetting").Visible = xlSheetHidden
    
    mRow = 1: mCol = 1
    Do While Not (nowSheet.Cells(mRow, mCol) = "" And nowSheet.Cells(mRow + 1, mCol) = "")
        If Left(nowSheet.Cells(mRow, mCol), 9) = "*** WAFER" Then
            mWaferID = Trim(Mid(nowSheet.Cells(mRow, mCol), 10))
            nowSheet.Cells(mRow, mCol) = "No./DataType"
            nowSheet.Cells(mRow, 2) = "Parameter"
            nowSheet.Cells(mRow, 3) = "Unit"
            i = 4
            Do While Not nowSheet.Cells(mRow, i) = ""
                Cells(mRow, i) = Coordinate(Worksheets("CoorSetting").Cells(29, 1).CurrentRegion, "N" & nowSheet.Cells(mRow, i), mWaferID, i - 3)
                i = i + 1
            Loop
            nowSheet.Cells(mRow, i) = "W L"
            nowSheet.Cells(mRow, i + 1) = "RULE"
        End If
        mRow = mRow + 1
    Loop
    Cells.Select
    Selection.Columns.AutoFit
    Range("A1").Select
    Set nowSheet = Nothing
    
End Function


Public Function LoadLongFileLONG(ByVal mFileName As String)
    Dim nowSheet As Worksheet
    Dim i As Integer, ALen As Integer
    Dim ColA() As Integer
    Dim WidthA() As Integer
   
    Set nowSheet = AddSheet("Data")
   
    ALen = 250
    ReDim ColA(ALen)
    ReDim WidthA(ALen)
    For i = 0 To ALen
        ColA(i) = 1
    Next i
    WidthA(0) = 15
    WidthA(1) = getParaLengthOfLong(mFileName)
    WidthA(2) = 9
    For i = 3 To ALen
        WidthA(i) = 15
    Next i

    With nowSheet.QueryTables.Add(Connection:="TEXT;" & mFileName, Destination:=Range("A1"))
        .Name = "IDAS"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = xlWindows
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = ColA
        .TextFileFixedColumnWidths = WidthA
        .Refresh BackgroundQuery:=False
    End With
    Call FixLongHeader
    nowSheet.Columns.AutoFit
    Range("A1").Select
    Set nowSheet = Nothing
End Function

Private Function getParaLengthOfLong(mFileName As String)
    Dim tmp As String
    Dim i As Integer, m As Integer, n As Integer
        
    tmp = GetFileHeader(mFileName, 20)
    For i = 1 To 20
        If InStr(tmp, "Parameter") > 0 Then
            m = InStr(tmp, "Parameter")
            n = InStr(tmp, "Unit")
            getParaLengthOfLong = n - m
            Exit For
        End If
    Next i
End Function

Public Function FixLongHeader()
    Dim nowSheet As Worksheet
    Dim i As Long, iRow As Long
    Dim tmpStr As String, tempA As Variant
    Dim siteNum As Integer
    
    siteNum = getSiteNum("Data")
    Set nowSheet = Worksheets("Data")
    
    For iRow = 1 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(iRow, 1) = "No./DataType" Then
            If Left(nowSheet.Cells(iRow, siteNum + 4), 1) = "W" Then
                If Trim(nowSheet.Cells(iRow + 1, siteNum + 4)) = "" Then
                    Exit Function
                Else
                    nowSheet.Range(Columns(siteNum + 4), Columns(siteNum + 5)).Delete
                End If
            End If
            tmpStr = ""
            For i = 4 To nowSheet.UsedRange.Columns.Count
                tmpStr = tmpStr & nowSheet.Cells(iRow, i)
                nowSheet.Cells(iRow, i) = ""
            Next i
            tempA = Split(tmpStr, "<")
            For i = 1 To UBound(tempA)
                If InStr(tempA(i), "Rule") > 0 Then
                    nowSheet.Cells(iRow, 3 + i) = "<" & getCOL(tempA(i), ")", 1) & ")"
                Else
                    nowSheet.Cells(iRow, 3 + i) = "<" & tempA(i)
                End If
            Next i
            nowSheet.Cells(iRow, 3 + i) = "W L"
            nowSheet.Cells(iRow, 3 + i + 1) = "Rule"
        End If
    Next iRow
    Debug.Print "Fixed LongHeader"
End Function

Public Sub ExportFile()
    Dim Filename As String
    Dim SubName As String
    Dim temp As String
   
    temp = ActiveSheet.Cells(7, 1).Value
      
    If Not InStr(ActiveSheet.Cells(7, 1).Value, "LONG REPORT") > 0 Then Exit Sub
    Filename = Application.GetSaveAsFilename(ActiveSheet.Name, "ELab Longfile(*.long),*.long")
    If Filename <> "" And Filename <> "False" Then
        Call GenLongFile(Filename)
        MsgBox "Export longfile finished."
    End If
End Sub

Public Function GenLongFile(mFile As String)
    Dim nowSheet As Worksheet
    Dim tmp As String
    Dim i As Long, j As Long
    Dim ynHeader As Boolean
   
    Set nowSheet = ActiveSheet
    nowSheet.Columns.AutoFit
    DoEvents
    ynHeader = True
    For i = 1 To nowSheet.UsedRange.Rows.Count
        Application.StatusBar = "Rows... " & CStr(i) & "/" & nowSheet.UsedRange.Rows.Count & " = " & _
            Format(i / nowSheet.UsedRange.Rows.Count, "00%")
        With nowSheet
            If Left(.Cells(i, 1), 1) = """" Then .Cells(i, 1) = Replace(.Cells(i, 1), """", "")
            If InStr(.Cells(i, 1), "No.") > 0 Then ynHeader = False
            If ynHeader Then
                tmp = FillSpace(.Cells(i, 1), 15) & .Cells(i, 2)
            Else
                If .Cells(i, 1) <> "" Then
                    tmp = FillSpace(.Cells(i, 1), 15) & FillSpace(.Cells(i, 2), 50) & FillSpace(.Cells(i, 3), 9)
                    For j = 4 To .UsedRange.Columns.Count
                        If .Cells(i, j) = "" Then Exit For
                        If .Cells(i, j) = "W       L" Then
                            tmp = tmp & FillSpace("W", 8) & FillSpace("L", 8) & FillSpace("Rule", 8)
                            Exit For
                        End If
                        tmp = tmp & FillSpace(.Cells(i, j).Value, 15)
                    Next j
                Else
                    tmp = ""
                End If
            End If
        End With
        Call WriteTextFile(mFile, tmp & vbCrLf)
    Next i
    Application.StatusBar = False
End Function


Public Sub Load_SPECFile()
    Dim Filename As String
   
    Filename = FileDialog("xls,xlsm,xlsx")
   
    If Filename <> "" Then
        If Dir(Filename) = "" Then
            Err.Raise 999, , "BDF File DOES NOT exist!"
            Exit Sub
        End If
        Call LoadSpecFile(Filename)
        Call GenSPECTEMP
    End If
End Sub
Public Function LoadSpecFile(Filename As String)
    Dim nowBook As Workbook
    Dim SpecBook As Workbook
    Dim i As Long
    Dim oldSheetCount As Long
    Dim tmpSheet As Worksheet
    Dim SourceSheet As Worksheet
   
    Set nowBook = ActiveWorkbook
    nowBook.Worksheets.Add before:=nowBook.Worksheets(1)
    Application.DisplayAlerts = False
    For i = nowBook.Worksheets.Count To 2 Step -1
        nowBook.Worksheets(i).Delete
    Next i
    Set SpecBook = Workbooks.Open(Filename, 0, , , , , True)
   
    SpecBook.Worksheets.Copy After:=nowBook.Sheets(1)
   
    SpecBook.Close
    If nowBook.Worksheets.Count > 1 Then
        nowBook.Worksheets(1).Delete
    End If
   
    On Error Resume Next
    For i = nowBook.Styles.Count To 1 Step -1
        If Not nowBook.Styles(i).BuiltIn Then nowBook.Styles(i).Delete
    Next i
    Application.DisplayAlerts = True
End Function
Public Sub CheckFile()
    Dim i As Long, j As Long
    Dim WaferStr As String
    Dim Filename
    Dim FileID As Long
    Dim temp As String, tempStr As String
    Dim nowSheet As Worksheet
    
    Dim mProduct_ID As String
    Dim mLot_ID As String
    Dim mTester_ID As String
    Dim mTest_Plan_ID As String
    Dim mDateTime As String
    Dim mPreview(9) As String
    Dim mCount As Integer
    Dim mSiteCnt As Integer
    
    Filename = Application.GetOpenFilename("rpt File, *.rpt", 1, "Load rpt file", "Open", True)
    If VarType(Filename) = vbBoolean Then Exit Sub
    Set nowSheet = AddSheet("Check", , "Data")
    nowSheet.Cells(1, 1).Value = "Filename"
    nowSheet.Cells(1, 2).Value = "Shuttle"
    nowSheet.Cells(1, 3).Value = "Lot"
    nowSheet.Cells(1, 4).Value = "Tester_ID"
    nowSheet.Cells(1, 5).Value = "Recipe"
    nowSheet.Cells(1, 6).Value = "Date"
    nowSheet.Cells(1, 7).Value = "SiteNum"
    nowSheet.Cells(1, 8).Value = "WaferNum"
    nowSheet.Cells(1, 9).Value = "Wafer"
    nowSheet.Cells(1, 10).Value = "Preview"
    
    For i = 1 To UBound(Filename)
        
        FileID = FreeFile
        Open Filename(i) For Input As #FileID
        Do Until EOF(FileID)
            Line Input #FileID, temp
            If InStr(temp, "TYPE") Then
                mProduct_ID = getValue(temp, "   ", "TYPE", "=")
                mLot_ID = getValue(temp, "   ", "LOT", "=")
                mTester_ID = getValue(temp, "   ", "TESTER_ID", "=")
                mTest_Plan_ID = getValue(temp, "   ", "Recipe", "=")
                mDateTime = getValue(temp, "   ", "DATE", "=")
            ElseIf InStr(temp, "*** WAFER") Then
                tempStr = tempStr & ", #" & Trim(Mid(temp, 11, 3))
                mSiteCnt = (((Len(temp) - Len(Replace(temp, vbTab, ""))) / Len(vbTab)) - 4 + 1) / 2
                mCount = 1
            ElseIf mCount > 0 And Not Trim(getCOL(temp, vbTab, 3)) Like "*R*PC*" And Not temp Like "*TEM_offset*" And mCount < 11 Then
                mPreview(mCount - 1) = Trim(getCOL(temp, vbTab, 3))
                mCount = mCount + 1
            End If
        Loop
        nowSheet.Cells(i + 1, 1).Value = Mid(Filename(i), InStrRev(Filename(i), "\") + 1)
        nowSheet.Cells(i + 1, 2).Value = mProduct_ID
        nowSheet.Cells(i + 1, 3).Value = mLot_ID
        nowSheet.Cells(i + 1, 4).Value = mTester_ID
        nowSheet.Cells(i + 1, 5).Value = mTest_Plan_ID
        nowSheet.Cells(i + 1, 6).Value = mDateTime
        nowSheet.Cells(i + 1, 7).Value = mSiteCnt
        nowSheet.Cells(i + 1, 8).Value = Len(tempStr) - Len(Replace(tempStr, "#", ""))
        nowSheet.Cells(i + 1, 9).Value = Mid(tempStr, 3)
        tempStr = ""
        For j = 0 To 9
            tempStr = tempStr & mPreview(j) & Chr(10)
        Next j
        nowSheet.Columns(10).ColumnWidth = 100
        nowSheet.Cells(i + 1, 10) = Left(tempStr, Len(tempStr) - Len(Chr(10)))
        tempStr = ""
        Close #FileID
    Next i
    nowSheet.Cells.Font.Name = "Arial"
    nowSheet.Cells.Font.Size = 10
    ActiveWindow.Zoom = 75
    nowSheet.Cells.Columns.AutoFit
    nowSheet.Rows("2:" & nowSheet.UsedRange.Rows.Count).RowHeight = 12.75 * 3
    Set nowSheet = Nothing
    MsgBox "Finished"
End Sub

Public Function Load_MultiLongFiles(Filename)

    Dim FileID As Long
    Dim nowSheet As Worksheet
    Dim i As Long
    Dim nowRow As Long
    Dim tmpStr
    Dim header As String
    Dim recipe As String
    Dim flag As Boolean
    Dim Files() As String
    Dim Lot As String
    Dim LotColl As New Collection
    Dim temp As String
    Dim k
    Dim Coll
    
    ReDim Files(1, UBound(Filename) - 1) As String
    
    On Error Resume Next
    
    For i = 1 To UBound(Filename)
        FileID = FreeFile
        Open Filename(i) For Input As #FileID
        Do Until EOF(FileID)
            Line Input #FileID, temp
            If InStr(temp, "TYPE") Then
                Lot = getValue(temp, "   ", "LOT", "=")
                LotColl.Add Lot, Lot
            End If
            If InStr(temp, "*** WAFER") Then
                Files(0, i - 1) = Lot & CInt(Trim(Mid(temp, 11, 3)))
                Files(1, i - 1) = Filename(i)
                Close #FileID
                Exit Do
            End If
        Loop
        Close #FileID
    Next i
    
    Err.Clear
    On Error GoTo 0
    
    Set nowSheet = AddSheet("Data")
    nowSheet.Activate
    nowRow = 1
    
    For Each Coll In LotColl
    flag = True
        For i = 1 To 25
            For k = 0 To UBound(Filename) - 1
                If Files(0, k) = Coll & CStr(i) Then
                    FileID = FreeFile
                    Open Files(1, k) For Input As #FileID
                        Do While (Not EOF(FileID))
                            If nowRow = 1048576 Then
                                Err.Raise 999, , "File line over 1048576!!!"
                                Call Unspeed
                                Exit Function
                            End If
                            Line Input #FileID, tmpStr
                            If InStr(tmpStr, "*** WAFER") Then flag = True
                            If flag Then
                                nowSheet.Cells(nowRow, 1).Value = tmpStr
                                nowRow = nowRow + 1
                            End If
                        Loop
                    Close #FileID
                    flag = False
                End If
            Next k
        Next i
    Next Coll
        
    nowSheet.Columns(1).TextToColumns Destination:=Range("A1"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=True, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=False, _
        FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
                         Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), _
                         Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
                         Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), _
                         Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), _
                         Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array(26, 1), _
                         Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), _
                         Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), _
                         Array(49, 1), Array(50, 1), Array(51, 1), Array(52, 1), Array(53, 1), Array(54, 1), _
                         Array(55, 1), Array(56, 1), Array(57, 1), Array(58, 1), Array(59, 1), Array(60, 1), _
                         Array(61, 1), Array(62, 1), Array(63, 1), Array(64, 1), Array(65, 1), Array(66, 1), Array(67, 1)), _
        TrailingMinusNumbers:=True
        
    If Trim(nowSheet.Cells(4, 1)) <> "" Then nowSheet.Rows(4).Delete
    For i = nowSheet.UsedRange.Columns.Count To 6 Step -1
        If Trim(nowSheet.Cells(5, i).Value) = "" Then nowSheet.Columns(i).Delete
    Next i
    Call ConvertRPT2
    Call InitStep
        
    nowSheet.Activate
    Set nowSheet = Nothing
    
End Function

Public Function ConvertRPT2()
    Dim mProduct_ID As String
    Dim mLot_ID As String
    Dim mTest_Plan_ID As String
    Dim mDateTime As String
    Dim i As Long
    Dim mRow As Long, mCol As Long
    Dim mWaferID As String
    Dim nowSheet As Worksheet
    Dim tmpStr As String
    Dim nowRow As Long
    Dim SubName As String
    Dim lotCount As Integer
    
    Set nowSheet = Worksheets("Data")
    nowSheet.Activate
    nowRow = 1
    
    For nowRow = nowSheet.UsedRange.Rows.Count To 1 Step -1
        If InStr(Cells(nowRow, 1), "TYPE") Then
            mProduct_ID = getValue(Cells(nowRow, 1), "   ", "TYPE", "=")
            mLot_ID = getValue(Cells(nowRow, 1), "   ", "LOT", "=")
            mTest_Plan_ID = getValue(Cells(nowRow, 1), "   ", "Recipe", "=")
            mDateTime = getValue(Cells(nowRow, 1), "   ", "DATE", "=")
            For i = 1 To 6
                nowSheet.Rows(nowRow - 2).Insert Shift:=xlDown
            Next i
            nowSheet.Cells(nowRow - 2, 2) = "<Process_ID>"
            nowSheet.Cells(nowRow - 1, 2) = "<Product_ID>"
            nowSheet.Cells(nowRow - 0, 2) = "<Lot_ID>"
            nowSheet.Cells(nowRow + 1, 2) = "<Test_Plan_ID>"
            nowSheet.Cells(nowRow + 2, 2) = "<Limit_File>"
            nowSheet.Cells(nowRow + 3, 2) = "<Date/Time>"
            nowSheet.Cells(nowRow + 4, 2) = "( LONG REPORT )"
            nowSheet.Cells(nowRow + 5, 2) = "-------------"
            nowSheet.Cells(nowRow + 6, 2) = "TYPE_SCALAR"
            nowSheet.Cells(nowRow + 7, 2) = "-------------"
            
            nowSheet.Cells(nowRow - 2, 3) = ":" & "x"
            nowSheet.Cells(nowRow - 1, 3) = ":" & mProduct_ID
            nowSheet.Cells(nowRow - 0, 3) = ":" & mLot_ID
            nowSheet.Cells(nowRow + 1, 3) = ":" & mTest_Plan_ID
            nowSheet.Cells(nowRow + 2, 3) = ":" & "x"
            nowSheet.Cells(nowRow + 3, 3) = ":" & mDateTime
            nowSheet.Cells(nowRow + 4, 3) = ":" & "x"
        End If
    Next nowRow

    nowSheet.Columns(1).Delete
    If Not IsExistSheet("CoorSetting") Then Call GenCoorSheet
    Worksheets("CoorSetting").Visible = xlSheetHidden
    
    mRow = 1: mCol = 1
    Do While Not (nowSheet.Cells(mRow, mCol) = "" And nowSheet.Cells(mRow + 1, mCol) = "")
        If Left(nowSheet.Cells(mRow, mCol), 4) = "TYPE" Then
            If lotCount > 0 Then SubName = Chr(lotCount + 64)
            lotCount = lotCount + 1
        End If
        If Left(nowSheet.Cells(mRow, mCol), 9) = "*** WAFER" Then
            mWaferID = Trim(Mid(nowSheet.Cells(mRow, mCol), 10))
            nowSheet.Cells(mRow, mCol) = "No./DataType"
            nowSheet.Cells(mRow, 2) = "Parameter"
            nowSheet.Cells(mRow, 3) = "Unit"
            i = 4
            Do While Not nowSheet.Cells(mRow, i) = ""
                nowSheet.Cells(mRow, i) = Coordinate(Worksheets("CoorSetting").Cells(29, 1).CurrentRegion, "N" & nowSheet.Cells(mRow, i), mWaferID, i - 3)
                i = i + 1
            Loop
            nowSheet.Cells(mRow, i) = "W L"
            nowSheet.Cells(mRow, i + 1) = "RULE"
            If lotCount > 1 Then
                nowSheet.Rows(mRow).Replace What:=getCOL(nowSheet.Cells(mRow, 4), "-", 1), _
                                     Replacement:=getCOL(nowSheet.Cells(mRow, 4), "-", 1) & SubName, _
                                          LookAt:=xlPart, _
                                     SearchOrder:=xlByRows, _
                                       MatchCase:=False, _
                                    SearchFormat:=False, _
                                   ReplaceFormat:=False
            End If
        End If
        mRow = mRow + 1
    Loop
    nowSheet.Cells.Select
    Selection.Columns.AutoFit
    nowSheet.Range("A1").Select
    Set nowSheet = Nothing
    
End Function

Public Sub LoadlongFile()

    Dim Filename As String
    Dim SubName As String
    Filename = Application.GetOpenFilename("Long File(*.long),*.long" & ",All File(*.*),*.*", 1, "Open File", "Open", False)

    If Not Filename = "" Then
        SubName = UCase(Mid(Filename, InStrRev(Filename, ".")))
        If getFileLine(Filename) > 1048576 Then
            Err.Raise 999, , "File line over 1048576!!!"
            Call Unspeed
            Exit Sub
        End If
        Select Case SubName
            Case ".RPT"
                Call LoadLongFileRPT(Filename)
            Case ".LONG"
                Call LoadLongFileLONG(Filename)
        End Select
        Call InitStep
    End If
    
End Sub
Public Sub UEDA()

    If Not ActiveSheet.Cells(1, 1).Value = "PRODUCT" Then MsgBox ("Please move to the sheet you want to transform"): Exit Sub

    Call Speed

    Dim srcSheet As Worksheet
    Dim nowSheet As Worksheet
    Dim mRange As Range
    Dim header As Object
    
    Dim i As Long, j As Long
    Dim nowRow
    
    Set srcSheet = ActiveSheet
    Set nowSheet = AddSheet("TEMPUEDA")
    Set header = CreateObject("Scripting.Dictionary")
    
    nowSheet.Cells(1, 1) = "PRODUCT"
    nowSheet.Cells(1, 2) = "PROCESS"
    nowSheet.Cells(1, 3) = "LOT"
    nowSheet.Cells(1, 4) = "WAFER"
    nowSheet.Cells(1, 5) = "COMPONENTID"
    nowSheet.Cells(1, 6) = "PARAMETER"
    nowSheet.Cells(1, 7) = "MEASURE_TIME"
        nowSheet.Columns(7).NumberFormat = "yyyy-mm-dd hh:MM:ss"
    nowSheet.Cells(1, 8) = "SITE_SEQ"
    nowSheet.Cells(1, 9) = "SITE_VALUE"
    nowSheet.Cells(1, 10) = "TEST_PROG"
    nowSheet.Cells(1, 11) = "SHOT_X"
    nowSheet.Cells(1, 12) = "SHOT_Y"
    
    i = 2
    nowRow = 2
    Do
        header.RemoveAll
        Set mRange = srcSheet.Cells(i, 1).CurrentRegion
        For j = 1 To mRange.Columns.Count
            header.Add mRange.Cells(1, j).Value, j
        Next j
        
        On Error Resume Next
        
        For j = 1 To nowSheet.UsedRange.Columns.Count
            If header.Exists(nowSheet.Cells(1, j).Value) Then
                nowSheet.Range(nowSheet.Cells(nowRow, j), nowSheet.Cells(nowRow + mRange.Rows.Count - 2, j)).Value = srcSheet.Range(mRange.Cells(2, header(nowSheet.Cells(1, j).Value)), mRange.Cells(mRange.Rows.Count, header(nowSheet.Cells(1, j).Value))).Value
            ElseIf nowSheet.Cells(1, j).Value = "SITE_SEQ" Then
                nowSheet.Range(nowSheet.Cells(nowRow, j), nowSheet.Cells(nowRow + mRange.Rows.Count - 2, j)).Value = srcSheet.Range(mRange.Cells(2, header("SITE_NAME")), mRange.Cells(mRange.Rows.Count, header("SITE_NAME"))).Value
            ElseIf nowSheet.Cells(1, j).Value = "TEST_PROG" Then
                nowSheet.Range(nowSheet.Cells(nowRow, j), nowSheet.Cells(nowRow + mRange.Rows.Count - 2, j)).Value = srcSheet.Range(mRange.Cells(2, header("RECIPE")), mRange.Cells(mRange.Rows.Count, header("RECIPE"))).Value
            End If
        Next j
        
        nowRow = nowRow + mRange.Rows.Count - 1
        i = i + mRange.Rows.Count
    
    Loop Until i >= srcSheet.UsedRange.Rows.Count
    
    Call importUEDA
    
    DelSheet (nowSheet.Name)
    
    Call Unspeed

End Sub




Public Sub importUEDA()

    Dim nowSheet As Worksheet
    Dim newSheet As Worksheet
    Dim Test_Prog As Object

    Dim header As Object
    Dim site As Object
    Dim i As Long, j As Integer

    Set nowSheet = Worksheets("TEMPUEDA")
    Set header = CreateObject("Scripting.Dictionary")
    Set site = CreateObject("Scripting.Dictionary")
    Set Test_Prog = CreateObject("Scripting.Dictionary")
        
    For i = 1 To nowSheet.UsedRange.Columns.Count
        header.Add nowSheet.Cells(1, i).Value, i
    Next i
    
    header.Add "SITE", header.Count + 1
    nowSheet.Cells(1, header.Count) = "SITE"
    
    For i = 2 To nowSheet.UsedRange.Rows.Count
        If Not site.Exists(nowSheet.Cells(i, header("SHOT_X")).Value & "," & nowSheet.Cells(i, header("SHOT_Y")).Value) Then
            site.Add nowSheet.Cells(i, header("SHOT_X")).Value & "," & nowSheet.Cells(i, header("SHOT_Y")).Value, site.Count + 1
        End If
        nowSheet.Cells(i, header("SITE")).Value = site(nowSheet.Cells(i, header("SHOT_X")).Value & "," & nowSheet.Cells(i, header("SHOT_Y")).Value)
    Next i
    
    If header.Exists("MEASURE_TIME") Then Call sortByColumn(header("MEASURE_TIME"), False)
    If header.Exists("SITE") Then sortByColumn (header("SITE"))
    If header.Exists("PARAMETER") Then sortByColumn (header("PARAMETER"))
    If header.Exists("WAFER") Then sortByColumn (header("WAFER"))
    If header.Exists("LOT") Then sortByColumn (header("LOT"))
    
    nowSheet.Cells(1, 1).CurrentRegion.RemoveDuplicates Columns:=Array(header("PRODUCT"), header("PROCESS"), header("LOT"), header("WAFER"), header("PARAMETER"), header("SITE")), header:=xlYes

    
    Dim lotInfo()
    
    ReDim lotInfo(7, 0)     ' PROCESS, PRODUCT, LOT, TEST_PROG, LIMIT_TIME, MEASURE_TIME, START_ROW, END_ROW
    i = 2
    lotInfo(0, UBound(lotInfo, 2)) = nowSheet.Cells(i, header("PROCESS"))
    lotInfo(1, UBound(lotInfo, 2)) = nowSheet.Cells(i, header("PRODUCT"))
    lotInfo(2, UBound(lotInfo, 2)) = nowSheet.Cells(i, header("LOT"))
    lotInfo(3, UBound(lotInfo, 2)) = nowSheet.Cells(i, header("TEST_PROG"))
    lotInfo(4, UBound(lotInfo, 2)) = "x"
    lotInfo(5, UBound(lotInfo, 2)) = nowSheet.Cells(i, header("MEASURE_TIME"))
    lotInfo(6, UBound(lotInfo, 2)) = 2
    
    For i = 3 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(i, header("LOT")).Value <> nowSheet.Cells(i - 1, header("LOT")).Value Then
            lotInfo(7, UBound(lotInfo, 2)) = i - 1
            ReDim Preserve lotInfo(7, UBound(lotInfo, 2) + 1)
            If header.Exists("PROCESS") Then lotInfo(0, UBound(lotInfo, 2)) = nowSheet.Cells(i, header("PROCESS")) Else lotInfo(0, UBound(lotInfo, 2)) = "x"
            If header.Exists("PRODUCT") Then lotInfo(1, UBound(lotInfo, 2)) = nowSheet.Cells(i, header("PRODUCT")) Else lotInfo(1, UBound(lotInfo, 2)) = "x"
            If header.Exists("LOT") Then lotInfo(2, UBound(lotInfo, 2)) = nowSheet.Cells(i, header("LOT")) Else lotInfo(2, UBound(lotInfo, 2)) = "x"
            If header.Exists("TEST_PROG") Then lotInfo(3, UBound(lotInfo, 2)) = nowSheet.Cells(i, header("TEST_PROG")) Else lotInfo(3, UBound(lotInfo, 2)) = "x"
            If header.Exists("LIMIT_FILE") Then lotInfo(4, UBound(lotInfo, 2)) = nowSheet.Cells(i, header("LIMIT_FILE")) Else lotInfo(4, UBound(lotInfo, 2)) = "x"
            If header.Exists("MEASURE_TIME") Then lotInfo(5, UBound(lotInfo, 2)) = nowSheet.Cells(i, header("MEASURE_TIME")) Else lotInfo(5, UBound(lotInfo, 2)) = "x"
            lotInfo(6, UBound(lotInfo, 2)) = i
        End If
    Next i
    lotInfo(7, UBound(lotInfo, 2)) = i - 1
    'If UBound(lotInfo, 2) > 0 Then ReDim Preserve lotInfo(7, UBound(lotInfo, 2) - 1)
    i = 1
    
    Do While (IsExistSheet("UEDA" & i))
        i = i + 1
    Loop
    Set newSheet = AddSheet("UEDA" & i)
    
    Dim nowRow As Long
    Dim nowCol As Long
    Dim nowLot As Long
    
    For nowLot = 0 To UBound(lotInfo, 2)
        If nowRow = 0 Then nowRow = 1 Else nowRow = newSheet.UsedRange.Rows.Count + 2
        
        newSheet.Cells(nowRow + 0, 1) = "<Process_ID>"
        newSheet.Cells(nowRow + 0, 2) = ":" & lotInfo(0, nowLot)
        newSheet.Cells(nowRow + 1, 1) = "<Product_ID>"
        newSheet.Cells(nowRow + 1, 2) = ":" & lotInfo(1, nowLot)
        newSheet.Cells(nowRow + 2, 1) = "<Lot_ID>"
        newSheet.Cells(nowRow + 2, 2) = ":" & lotInfo(2, nowLot)
        newSheet.Cells(nowRow + 3, 1) = "<Test_Plan_ID>"
        newSheet.Cells(nowRow + 3, 2) = ":" & lotInfo(3, nowLot)
        newSheet.Cells(nowRow + 4, 1) = "<Limit_File>"
        newSheet.Cells(nowRow + 4, 2) = ":" & lotInfo(4, nowLot)
        newSheet.Cells(nowRow + 5, 1) = "<Date/Time>"
        newSheet.Cells(nowRow + 5, 2) = ":" & lotInfo(5, nowLot)
        newSheet.Cells(nowRow + 6, 1) = "( LONG REPORT )"
        newSheet.Cells(nowRow + 6, 2) = ":x"
        newSheet.Cells(nowRow + 7, 1) = "-------------"
        newSheet.Cells(nowRow + 8, 1) = "TYPE_SCALER"
        newSheet.Cells(nowRow + 9, 1) = "-------------"
        
        Dim isX As Boolean
        
        isX = Not (Left(getCOL(lotInfo(3, nowLot), "-", 2), 1) = "y")
               
        i = lotInfo(6, nowLot)
        nowRow = nowRow + 8
        nowCol = 3
        Do
            If nowSheet.Cells(i, header("WAFER")) <> nowSheet.Cells(i - 1, header("WAFER")) Then
                nowRow = nowRow + 2
                newSheet.Cells(nowRow, 1) = "No./DataType"
                newSheet.Cells(nowRow, 2) = "Parameter"
                newSheet.Cells(nowRow, 3) = "Unit"
                For j = 0 To site.Count - 1
                    newSheet.Cells(nowRow, 4 + j) = "<" & IIf(nowLot = 0, "", N2L(nowLot)) & Val(nowSheet.Cells(i, header("WAFER"))) & "-" & j + 1 & ">(" & site.Keys()(j) & ")"
                    nowCol = j + 4
                Next j
                
                newSheet.Cells(nowRow, nowCol + 1) = "W L"
                newSheet.Cells(nowRow, nowCol + 2) = "RULE"
                nowRow = nowRow + 1
            ElseIf nowSheet.Cells(i, header("PARAMETER")) <> nowSheet.Cells(i - 1, header("PARAMETER")) Then
                nowRow = nowRow + 1
            End If
            

            
            If IsNumeric(newSheet.Cells(nowRow - 1, 1)) Then newSheet.Cells(nowRow, 1) = newSheet.Cells(nowRow - 1, 1) + 1 Else newSheet.Cells(nowRow, 1) = 1
            newSheet.Cells(nowRow, 2) = nowSheet.Cells(i, header("PARAMETER"))
            
            nowCol = 3
            
            newSheet.Cells(nowRow, nowCol + nowSheet.Cells(i, header("SITE"))) = nowSheet.Cells(i, header("SITE_VALUE"))
            
            i = i + 1
               
        Loop Until i >= lotInfo(7, nowLot) + 1

    Next nowLot
    newSheet.Cells.EntireColumn.AutoFit
    
End Sub

Private Function CompareMap(Test_Prog1 As String, Test_Prog2 As String, Optional sheetName As String = "TEMPUEDA") As Boolean
    
    Dim nowSheet As Worksheet
    Dim header As Object
    Dim Coor1 As Object
    Dim Coor2 As Object
    Dim i As Long
    
    Set nowSheet = Worksheets(sheetName)
    Set header = CreateObject("Scripting.Dictionary")
    Set Coor1 = CreateObject("Scripting.Dictionary")
    Set Coor2 = CreateObject("Scripting.Dictionary")
    
    For i = 1 To nowSheet.UsedRange.Columns.Count
        header.Add nowSheet.Cells(1, i).Value, i
    Next i
    
    For i = 2 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(i, header("TEST_PROG")) = Test_Prog1 Then
            Do
                Coor1.Add nowSheet.Cells(i, header("SITE_SEQ")).Value, nowSheet.Cells(i, header("SHOT_X")).Value & "," & nowSheet.Cells(i, header("SHOT_Y")).Value
                i = i + 1
            Loop Until nowSheet.Cells(i, header("SITE_SEQ")) < nowSheet.Cells(i - 1, header("SITE_SEQ"))
            Exit For
        End If
    Next
    
    For i = 2 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(i, header("TEST_PROG")) = Test_Prog2 Then
            Do
                Coor2.Add nowSheet.Cells(i, header("SITE_SEQ")).Value, nowSheet.Cells(i, header("SHOT_X")).Value & "," & nowSheet.Cells(i, header("SHOT_Y")).Value
                i = i + 1
            Loop Until nowSheet.Cells(i, header("SITE_SEQ")) < nowSheet.Cells(i - 1, header("SITE_SEQ"))
            Exit For
        End If
    Next
    
    If Coor1.Count <> Coor2.Count Then Exit Function
    For i = 1 To Coor1.Count
        If Coor1(i) <> Coor2(i) Then Exit Function
    Next i
    CompareMap = True
    
End Function
