
Public Const DataPath = "C:\Loader\DataReport\"
Public Const SpecPath = "C:\Loader\DataReport\"

Public Type specInfo
    mPara As Variant
    mFAC As Variant
    mUnit As Variant
    mLow As Variant
    mTarget As Variant
    mHigh As Variant
    mDevice As Variant
    mModel As Variant
    mRound As Variant
    mDescription As Variant
End Type

Public Enum specColumn
    argPara = 1
    argFAC = 2
    argUnit = 3
    argLow = 4
    argTarget = 5
    argHigh = 6
    argDevice = 7
End Enum

Sub SetWaferRange()
    Dim nowSheet As Worksheet
    Dim i As Long
    Dim nowRange As Range
    Dim nowWafer As String
    Dim temp
    Dim typeName As String
       
    On Error Resume Next
   
    For i = ActiveWorkbook.Names.Count To 1 Step -1
        ActiveWorkbook.Names(i).Delete
    Next i
    ReDim WaferArray(1, 0)
    
    typeName = "wafer_"
    Set nowSheet = Worksheets("Data")
    For i = 1 To nowSheet.UsedRange.Rows.Count
        If Trim(nowSheet.Cells(i, 1)) = "TYPE_SCALAR" Then typeName = "wafer_"
        If Trim(nowSheet.Cells(i, 1)) = "TYPE_VECTOR" Then Exit Sub
        If InStr(nowSheet.Cells(i, 1), "No") > 0 Then
            nowWafer = getCOL(getCOL(nowSheet.Cells(i, 4), "<", 2), "-", 1)
            WaferArray(0, UBound(WaferArray, 2)) = nowWafer
            ReDim Preserve WaferArray(1, UBound(WaferArray, 2) + 1)
            temp = nowSheet.Cells(i - 1, 1)
            nowSheet.Cells(i - 1, 1) = ""
            Set nowRange = nowSheet.Cells(i, 1).CurrentRegion
            nowSheet.Names.Add typeName & nowWafer, nowRange
            nowSheet.Cells(i - 1, 1) = temp
        End If
    Next i
    ReDim Preserve WaferArray(1, UBound(WaferArray, 2) - 1)
    Set nowSheet = Nothing
End Sub

Public Function GenSPECTEMP()
    Dim iSheet As Worksheet
    Dim oSheet As Worksheet
    
    Set iSheet = Worksheets(SSheet)
    If Not UCase(iSheet.Cells(1, 1).Value) = "DEVICE" Then MsgBox "Sheet""(SPEC)"" has wrong format": Exit Function
    Set oSheet = AddSheet("SPECTEMP")
        
    iSheet.Cells(1, 1).CurrentRegion.Copy
    oSheet.Activate
    oSheet.Paste
    Columns(1).Select: Selection.Copy
    Columns(8).Select: Selection.PasteSpecial
    Columns(3).Select: Selection.Copy
    Columns(1).Select: Selection.PasteSpecial
    Columns(3).Delete Shift:=xlLeft

    oSheet.Visible = xlSheetHidden
    
    Set iSheet = Nothing
    Set oSheet = Nothing
    
End Function

Public Function SetRawWaferRange(mSheet As String)
    Dim nowSheet As Worksheet
    Dim sCol As Long, eCol As Long, iRow As Long, i As Long
    Dim oldWafer As String
    
    Set nowSheet = Worksheets(mSheet)
    iRow = nowSheet.UsedRange.Rows.Count
    
    For i = 7 To nowSheet.UsedRange.Columns.Count + 1
        If nowSheet.Cells(1, i) <> "" Or i = nowSheet.UsedRange.Columns.Count + 1 Then
            If oldWafer <> nowSheet.Cells(1, i) Then
                If oldWafer <> "" Then
                    eCol = i - 1
                    nowSheet.Names.Add "wafer_" & oldWafer, nowSheet.Range(Cells(1, sCol), Cells(iRow, eCol))
                End If
                oldWafer = nowSheet.Cells(1, i)
                sCol = i
            End If
        End If
    Next i
    
    Set nowSheet = Nothing
    
End Function

Public Function SetGrouping(mSheet As String, waferId As String, groupId As String)
    Dim GroupMember
    Dim i As Integer
    Dim nowSheet As Worksheet
    Dim nowRange As Range
    Set nowSheet = Worksheets(mSheet)
    
    groupId = Replace(groupId, "+", "PLUS")
    GroupMember = Split(waferId, ",")
    
    On Error Resume Next
    
    For i = 0 To UBound(GroupMember)
        If nowRange Is Nothing Then
            Set nowRange = nowSheet.Range("wafer_" & GroupMember(i))
        Else
            Set nowRange = Union(nowRange, nowSheet.Range("wafer_" & GroupMember(i)))
        End If
    Next i
    
    nowSheet.Names.Add "wafer_" & groupId, nowRange
    
    Set nowSheet = Nothing
    Set nowRange = Nothing
End Function

Public Function getSiteNum(sheetName As String, Optional waferList As String)
    Dim i As Long, j As Long
    Dim nowSheet As Worksheet
    Dim siteNum As Integer
    Dim mRange As Range
    
    Set nowSheet = Worksheets(sheetName)
    If waferList = "" Then
        Set mRange = nowSheet.UsedRange
    Else
        Set mRange = nowSheet.Range("wafer_" & waferList)
    End If
    
    For i = 1 To mRange.Rows.Count
       If InStr(mRange.Cells(i, 1), "No.") Then Exit For
    Next i
    For j = 1 To mRange.Columns.Count
       If InStr(mRange.Cells(i, j), "<") Then siteNum = siteNum + 1
    Next j
    getSiteNum = siteNum
End Function
Public Function getSPEC(ByVal nowPara As String, ByVal item As String)
    Dim n As Integer
    If InStr(nowPara, "(") Then
        Dim cnt As Integer
        cnt = Len(nowPara) - Len(Replace(nowPara, "(", ""))
        nowPara = getCOL(getCOL(nowPara, "(", cnt + 1), ")", 1)
    End If
    
    Select Case Trim(UCase(item))
        Case "DEVICE"
            n = 7
        Case "FAC"
            n = 2
        Case "ITEM"
            n = 1
        Case "PARAMETER"
            n = 1
        Case "PARA"
            n = 1
        Case "UNIT"
            n = 3
        Case "SPEC LO"
            n = 4
        Case "SS"
            n = 4
        Case "TARGET"
            n = 5
        Case "TT"
            n = 5
        Case "SPEC HI"
            n = 6
        Case "FF"
            n = 6
        Case Else
    End Select
    If n = 3 Then
        getSPEC = getCOL(getSPECByPara(nowPara, 3), ":", 1)
    ElseIf n = 4 Or n = 5 Or n = 6 Then
        getSPEC = CDbl(getSPECByPara(nowPara, n))
    Else
        getSPEC = getSPECByPara(nowPara, n)
    End If
End Function

Public Function getCORNER(ByVal nowPara As String)
    Dim CornerSheet As Worksheet
    Dim header As Object
    Dim FF
    Dim SS
    Dim SF
    Dim FS
    
    
    If Not IsExistSheet("Corner") Then Exit Function
    
    Set CornerSheet = Worksheets("Corner")
    Set header = CreateObject("Scripting.Dictionary")
    For i = 1 To Worksheets("Corner").UsedRange.Columns.Count
        If CornerSheet.Cells(1, i).Value = "" Then Exit For
        header.Add CornerSheet.Cells(1, i).Value, i
    Next i
    
    FF = getSPECByPara(trimFunc(trimFunc(nowPara, "ABS"), "MEDIUM"), header("FF"), "Corner")
    SS = getSPECByPara(trimFunc(trimFunc(nowPara, "ABS"), "MEDIUM"), header("SS"), "Corner")
    SF = getSPECByPara(trimFunc(trimFunc(nowPara, "ABS"), "MEDIUM"), header("SNFP"), "Corner")
    FS = getSPECByPara(trimFunc(trimFunc(nowPara, "ABS"), "MEDIUM"), header("FNSP"), "Corner")
    
    getCORNER = Join(Array(SS, SF, FF, FS, SS), ", ")
    
    If FF = Empty Or SF = Empty Or FS = Empty Or SS = Empty Then getCORNER = ""
    
End Function
Public Function getSPECInfo(ByVal nowPara As String, Optional sheetName As String = "SPECTEMP") As specInfo
    
    Dim vSpec As specInfo
    If InStr(nowPara, ":") Then nowPara = getCOL(nowPara, ":", 2)
    vSpec.mPara = getSPECByPara(nowPara, 1, sheetName)
    vSpec.mFAC = getSPECByPara(nowPara, 2, sheetName)
    vSpec.mUnit = getCOL(getSPECByPara(nowPara, 3, sheetName), ":", 1)
    vSpec.mRound = UCase(getCOL(getSPECByPara(nowPara, 3, sheetName), ":", 2))
    vSpec.mLow = getSPECByPara(nowPara, 4, sheetName)
    vSpec.mTarget = getSPECByPara(nowPara, 5, sheetName)
    vSpec.mHigh = getSPECByPara(nowPara, 6, sheetName)
    vSpec.mDevice = getSPECByPara(nowPara, 7, sheetName)
   
    If IsEmpty(vSpec.mFAC) Or Len(Trim(vSpec.mFAC)) = 0 Then vSpec.mFAC = 1
    If IsEmpty(vSpec.mPara) Then vSpec.mPara = nowPara
    getSPECInfo = vSpec
End Function
Public Function getSPECByPara(ByVal nowPara As String, ByVal n As specColumn, Optional sheetName As String = "SPECTEMP")
    Dim reValue
    Dim nowRange As Range
    Dim TargetSheet As Worksheet
   
    If Left(nowPara, 1) = "'" Then nowPara = Mid(nowPara, 2)
   
    Set TargetSheet = Worksheets(sheetName)
    Set nowRange = TargetSheet.UsedRange
    On Error Resume Next
    reValue = Trim(Application.WorksheetFunction.VLookup(nowPara, nowRange, n, False))
    If Not IsEmpty(reValue) Then
        If Trim(reValue) = "" Then Set reValue = Nothing
    End If
    getSPECByPara = reValue
End Function

Public Function getRangeByPara(nowWafer As String, nowPara As String, Optional dieNum As Integer = 0)
    Dim reValue
    Dim reRange As Range
    Dim nowRange As Range
   
    Set nowRange = Worksheets("Data").Range("wafer_" & nowWafer)
    Set getRangeByPara = Nothing
    Set reValue = nowRange.Find(What:=nowPara, LookAt:=xlWhole)
    If Not reValue Is Nothing Then
        If reValue.Column <> 2 Then Exit Function
        Set getRangeByPara = reValue.Range(N2L(3) & "1" & ":" & N2L(dieNum + 2) & "1")
    End If
    Set nowRange = Nothing
    Set reValue = Nothing
End Function

Public Function FormulaParse(ByRef sourceFormula As String, ByRef itemArray() As String) As String
    Dim i As Integer
    Dim tempA
    Dim j As Integer
    Dim DualFlag As Boolean
    Dim strFormula As String
    Dim TmpA As String, tmpB As String, tmp As String
   
    strFormula = sourceFormula
    strFormula = getCOL(strFormula, ":", 1)
    For i = 1 To Len(strFormula)
        Select Case Mid(strFormula, i, 1)
            Case "+", "-", "/", "*", "(", ")", ",": Mid(strFormula, i, 1) = vbTab
        End Select
    Next i
    tempA = Split(strFormula, vbTab)
    ReDim itemArray(0)
    For i = 0 To UBound(tempA)
        If tempA(i) <> "" And Not IsNumeric(tempA(i)) Then
            Select Case UCase(tempA(i))
                Case "ABS", "LOG", "LOG10", "SQRT", "TINORM", "LN", "EXP"
                
                Case "MEDIAN"
                    FormulaParse = "MEDIAN"

                Case Else
                    If Not Left(tempA(i), 1) = """" Then
                        DualFlag = False
                        For j = 0 To UBound(itemArray)
                            If tempA(i) = itemArray(j) Then DualFlag = True: Exit For
                        Next j
                        If Not DualFlag Then
                            itemArray(UBound(itemArray)) = tempA(i)
                            ReDim Preserve itemArray(UBound(itemArray) + 1)
                        End If
                    End If
            End Select
        End If
    Next i
    If UBound(itemArray) > 0 Then ReDim Preserve itemArray(UBound(itemArray) - 1)
    
End Function

Public Function RangeStyle(nowSheet As Worksheet, itemRange As Range)
    Dim nowRange As Range
    Dim i As Long, j As Long
    Dim specRange As Range
    Dim specSheet As Worksheet
  
    Set specSheet = Worksheets("SPEC")
    Set specRange = Worksheets("SPEC").Range("A1:G1")

    For i = 1 To specRange.Columns.Count - 1
        j = IIf(i > 1, i + 1, i)
        With nowSheet.Range(Cells(1, i), Cells(nowSheet.UsedRange.Rows.Count, i))
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeRight).LineStyle = specRange.Cells(1, j).Borders(xlEdgeRight).LineStyle
        End With
        With nowSheet.Cells(2, i)
            .Interior.ColorIndex = specRange.Cells(1, j).Interior.ColorIndex
            .Font.ColorIndex = specRange.Cells(1, j).Font.ColorIndex
            .Font.Bold = specRange.Cells(1, j).Font.Bold
            .HorizontalAlignment = specRange.Cells(1, j).HorizontalAlignment
        End With
        With nowSheet.Range(Cells(3, i), Cells(nowSheet.UsedRange.Rows.Count, i))
            .Interior.ColorIndex = specRange.Cells(2, j).Interior.ColorIndex
            .Font.ColorIndex = specRange.Cells(2, j).Font.ColorIndex
            .Font.Bold = specRange.Cells(2, j).Font.Bold
            .HorizontalAlignment = specRange.Cells(2, j).HorizontalAlignment
        End With
    Next i

    For i = 3 To nowSheet.UsedRange.Rows.Count
        
        nowSheet.Rows(i).Borders(xlEdgeBottom).Weight = itemRange.Rows(i - 2).Borders(xlEdgeBottom).Weight
        nowSheet.Rows(i).Borders(xlEdgeBottom).LineStyle = itemRange.Rows(i - 2).Borders(xlEdgeBottom).LineStyle
        
        With Range(Cells(i, 1), Cells(i, 3))
            .Interior.ColorIndex = itemRange.Rows(i - 2).Interior.ColorIndex
            .Font.ColorIndex = itemRange.Rows(i - 2).Font.ColorIndex
            .HorizontalAlignment = itemRange.Rows(i - 2).HorizontalAlignment
        End With
        With Range(Cells(i, 7), Cells(i, nowSheet.UsedRange.Columns.Count))
            .Interior.ColorIndex = itemRange.Rows(i - 2).Interior.ColorIndex
            .Font.ColorIndex = itemRange.Rows(i - 2).Font.ColorIndex
        End With
    Next i

End Function

Public Function RawdataFormatByUnit(mSheet As String)
    Dim i As Long
    Dim nowRange As Range
    Dim nowSheet As Worksheet
    Dim nowRow As Long
    Dim objSPEC As specInfo
    Dim nowParameter As String
    Dim formatStr As String
    Dim tmp As String
    
    Set nowSheet = Worksheets(mSheet)

    For nowRow = 3 To nowSheet.UsedRange.Rows.Count
        nowParameter = nowSheet.Cells(nowRow, 2)
        objSPEC = getSPECInfo(nowParameter)
        'Define Automatically
        If objSPEC.mRound = "" Then
            If objSPEC.mUnit <> "" Then
                Select Case objSPEC.mUnit
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
        'Define Manually
        Else
            If Left(objSPEC.mRound, 1) = "E" Then
                formatStr = "0." & String(CInt(Mid(objSPEC.mRound, 2)), "0") & "E+00"
            ElseIf Left(objSPEC.mRound, 1) = "%" Then
                If Len(objSPEC.mRound) = 1 Then
                    tmp = "0%"
                Else
                    tmp = "0." & String(CInt(Mid(objSPEC.mRound, 2)), "0") & "%"
                End If
                formatStr = tmp
            Else
                formatStr = "0." & String(CInt(objSPEC.mRound), "0")
            End If
        End If
        Set nowRange = nowSheet.Range(Cells(nowRow, 4), Cells(nowRow, nowSheet.UsedRange.Columns.Count))
        nowRange.NumberFormatLocal = formatStr
    Next nowRow
End Function
Public Sub temp()
    Call HighLightSBG(ActiveSheet.Name)
End Sub


Public Function HighLightSBG(mSheet As String)
    Dim nowSheet As Worksheet
    Dim nowRow As Long
    Dim nowCol As Long
    Dim staRow As Long
    Dim endRow As Long
    Dim keyStr As String
    Dim maxValue As Double
    
    Set nowSheet = Worksheets(mSheet)
    
    For nowRow = 3 To nowSheet.UsedRange.Rows.Count
        If Not nowSheet.Cells(nowRow, 2) = "" And nowSheet.Cells(nowRow, 2).Interior.Color = RGB(192, 192, 192) Then
            keyStr = Replace(nowSheet.Cells(nowRow, 2), getCOL(nowSheet.Cells(nowRow, 2), "_", 1), "")
            staRow = nowRow
            endRow = nowRow
        End If
        Do While InStr(nowSheet.Cells(nowRow, 2), keyStr) > 0 And nowSheet.Cells(nowRow, 2).Interior.Color = RGB(192, 192, 192)
            nowRow = nowRow + 1
            endRow = nowRow - 1
        Loop
        If endRow > staRow Then
            For nowCol = 7 To nowSheet.UsedRange.Columns.Count
                If nowSheet.Cells(staRow, nowCol) = "" Then Exit For
                nowSheet.Cells(staRow - 1 + AbsMax(Range(nowSheet.Cells(staRow, nowCol), nowSheet.Cells(endRow, nowCol))), nowCol).Font.Bold = True
            Next nowCol
            endRow = 0
            staRow = 0
            keyStr = ""
            nowRow = nowRow - 1
        End If
        
        
        
        'If Not nowParameter = "" Then keyHead = Left(g etCOL(nowParameter, "_", 1), 3)
        'If InStr(nowSheet.Cells(nowRow + 1, 2), keyStr) > 0 And InStr(nowSheet.Cells(nowRow + 2, 2), keyStr) > 0 And InStr(nowSheet.Cells(nowRow + 1, 2), keyHead) > 0 And InStr(nowSheet.Cells(nowRow + 2, 2), keyHead) > 0 Then
        '    If nowSheet.Cells(nowRow, 2) Like keyHead & "*" & keyStr And nowSheet.Cells(nowRow + 1, 2) Like keyHead & "*" & keyStr And nowSheet.Cells(nowRow + 2, 2) Like keyHead & "*" & keyStr And nowSheet.Cells(nowRow + 3, 2) Like keyHead & "*" & keyStr Then
        '    Else
        '        For nowCol = 7 To nowSheet.UsedRange.Columns.Count
        '            If nowSheet.Cells(nowRow, nowCol) = "" Then Exit For
        '            maxValue = Application.WorksheetFunction.Max(Abs(nowSheet.Cells(nowRow, nowCol).Value), Abs(nowSheet.Cells(nowRow + 1, nowCol).Value), Abs(nowSheet.Cells(nowRow + 2, nowCol).Value))
        '            If Abs(nowSheet.Cells(nowRow, nowCol).Value) = maxValue Then
        '                nowSheet.Cells(nowRow, nowCol).Font.Bold = True
        '            ElseIf Abs(nowSheet.Cells(nowRow + 1, nowCol).Value) = maxValue Then
        '                nowSheet.Cells(nowRow + 1, nowCol).Font.Bold = True
        '            ElseIf Abs(nowSheet.Cells(nowRow + 2, nowCol).Value) = maxValue Then
        '                nowSheet.Cells(nowRow + 2, nowCol).Font.Bold = True
        '            End If
        '        Next nowCol
        '        nowRow = nowRow + 2
        '    End If
        'End If
    Next nowRow
    Set nowSheet = Nothing
End Function

Public Function SummaryFormatByUnit(mSheet As String, hcount As Integer)
    Dim i As Long, j As Long
    Dim nowRange As Range
    Dim waferList() As String
    Dim HColCount  As Integer
    Dim HCol As Integer
    Dim ProductID As String
    Dim nowRow As Long
    Dim objSPEC As specInfo
    Dim nowParameter As String
    Dim formatStr As String
    Dim nowSheet As Worksheet
    Dim exStr As String
    Dim waferNum As Integer
    Dim waferId() As String
    Dim groupId() As String
    Dim tmp
   
    Set nowSheet = Worksheets(mSheet)
    
    If IsExistSheet("Grouping") Then
        waferNum = Worksheets("Grouping").Cells(1, 1).CurrentRegion.Rows.Count - 1
    Else
        Call GetWaferList(dSheet, waferList)
        waferNum = UBound(waferList) + 1
    End If
    
    For nowRow = 3 To nowSheet.UsedRange.Rows.Count
        nowParameter = nowSheet.Cells(nowRow, 2)
        If Left(UCase(nowSheet.Cells(nowRow, 2)), 5) = "TREND" Then nowParameter = nowSheet.Cells(nowRow - 1, 2)
        
        objSPEC = getSPECInfo(nowParameter)
        'Define Automatically
        If objSPEC.mRound = "" Then
            If objSPEC.mUnit <> "" Then
                Select Case objSPEC.mUnit
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
        Else
        'Define Manaully
            If Left(objSPEC.mRound, 1) = "E" Then
                formatStr = "0." & String(CInt(Mid(objSPEC.mRound, 2)), "0") & "E+00"
            ElseIf Left(objSPEC.mRound, 1) = "%" Then
                If Len(objSPEC.mRound) = 1 Then
                    tmp = "0%"
                Else
                    tmp = "0." & String(CInt(Mid(objSPEC.mRound, 2)), "0") & "%"
                End If
                formatStr = tmp
            Else
                formatStr = "0." & String(CInt(objSPEC.mRound), "0")
                If CInt(objSPEC.mRound) = 0 Then formatStr = "0"
            End If
        End If
        For j = 0 To hcount - 1
            tmp = nowSheet.Cells(2, 7 + j * waferNum).Value
            Set nowRange = nowSheet.Range(N2L(7 + j * waferNum) & CStr(nowRow) & ":" & N2L(6 + (j + 1) * waferNum) & CStr(nowRow))
            If tmp = "Median" Then
                If Not Left(nowSheet.Cells(nowRow, 2), 4) = "Diff" And Not Left(nowSheet.Cells(nowRow, 2), 5) = "Times" Then nowRange.NumberFormatLocal = formatStr
            ElseIf tmp = "Average" Or tmp = "Max" Or tmp = "Min" Then
                If Not Left(nowSheet.Cells(nowRow, 2), 4) = "Diff" And Not Left(nowSheet.Cells(nowRow, 2), 5) = "Times" Then nowRange.NumberFormatLocal = formatStr
            ElseIf tmp = "Sigma" Then
                nowRange.NumberFormatLocal = Replace(formatStr, ".0", ".00")
            ElseIf tmp = "Diff" Then
                If nowRange.NumberFormatLocal <> "0.00%" And _
                   nowRange.NumberFormatLocal <> "0.000""x""" Then _
                   nowRange.NumberFormatLocal = formatStr
            End If
        Next j
        'SPEC Lo, TARGET, SPEC Hi
        nowSheet.Range(N2L(4) & CStr(nowRow) & ":" & N2L(6) & CStr(nowRow)).NumberFormatLocal = formatStr
    Next nowRow

End Function

Public Function getRawRangeByRow(mSheet As String, nowWafer As String, nowRow As Long)
    Dim reValue
    Dim reRange As Range
    Dim i As Long, j As Long
    Dim nowRange As Range
    Dim nowSheet As Worksheet
    Dim nowName As String
    
    For i = 0 To 10
        If i = 0 Then
            nowName = mSheet & "_Raw"
        Else
            nowName = mSheet & "_Raw_" & CStr(i)
        End If
        If IsExistSheet(nowName) Then
            Set nowSheet = Worksheets(nowName)
            For j = 1 To nowSheet.Names.Count
                If getCOL(nowSheet.Names(j).Name, "!", 2) = "wafer_" & nowWafer Then
                    Set nowRange = Worksheets(nowName).Range("wafer_" & nowWafer)
                    Set getRawRangeByRow = nowRange.Rows(nowRow)
                    Exit Function
                End If
            Next j
        Else
            Exit Function
        End If
    Next i
End Function

Public Function getSPECFormatByPara(ByVal nowPara As String, ByVal n As specColumn)
   Dim reValue
   Dim nowRange As Range
   Dim TargetSheet As Worksheet
   Dim strFormat As String
   
   If Left(nowPara, 1) = "'" Then nowPara = Mid(nowPara, 2)
   
   Set TargetSheet = Worksheets("SPEC")
   Set nowRange = TargetSheet.UsedRange
   On Error Resume Next
   'reValue = Trim(Application.WorksheetFunction.VLookup(nowPara, nowRange, n, False))
   Set reValue = nowRange.Find(nowPara, LookIn:=xlValues, LookAt:=xlWhole)
   If Not IsEmpty(reValue) Then
      'If Trim(reValue) = "" Then Set reValue = Nothing
      Set reValue = reValue.Offset(0, n - 3)
      strFormat = reValue.NumberFormat
   End If
   getSPECFormatByPara = strFormat
End Function

Public Function SummaryModeSetting(ByVal header As String, ByRef Headerarray, ByRef ItemsCount As Integer)
    
    Dim ModeArray
    Dim i As Integer
    ModeArray = Array("comp", "diff", "r2", "outpsec", "ztable", "max", "min", "cpk133", "cpk150", "cpk167", "kvalue", "wid")
    
    If Not IsError(Application.Match(LCase(header), ModeArray, 0)) Then
        If InStr(LCase(header), "cpk") Then
            ReDim Preserve Headerarray(UBound(Headerarray) + 3)
            Headerarray(UBound(Headerarray) - 2) = "CPK"
            Headerarray(UBound(Headerarray) - 1) = "Ca"
            Headerarray(UBound(Headerarray) - 0) = "Cp"
        ElseIf InStr(LCase(header), "max") Or InStr(LCase(header), "min") Then
            If IsError(Application.Match("Max", Headerarray, 0)) Then
                ReDim Preserve Headerarray(UBound(Headerarray) + 2)
                Headerarray(UBound(Headerarray) - 1) = "Max"
                Headerarray(UBound(Headerarray) - 0) = "Min"
            End If
        Else
            ReDim Preserve Headerarray(UBound(Headerarray) + 1)
            If LCase(header) = "comp" Then
                For i = UBound(Headerarray) To LBound(Headerarray) + 1 Step -1
                    Headerarray(i) = Headerarray(i - 1)
                Next i
                Headerarray(0) = "Shift to BSL"
                
            End If
            If LCase(header) = "diff" Then Headerarray(UBound(Headerarray)) = "Diff"
            If LCase(header) = "r2" Then Headerarray(UBound(Headerarray)) = "R Square"
            If LCase(header) = "outspec" Then Headerarray(UBound(Headerarray)) = "OutSpec"
            If LCase(header) = "ztable" Then Headerarray(UBound(Headerarray)) = "Z Value"
            If LCase(header) = "kvalue" Then Headerarray(UBound(Headerarray)) = "K Value"
            If LCase(header) = "wid" Then Headerarray(UBound(Headerarray)) = "3 Sigma/Median"
        End If
    End If
    ItemsCount = UBound(Headerarray) + 1
    
End Function

Public Function getReceiver()
    Dim nowSheet As Worksheet
    Dim i As Long, j As Integer
    Dim TOlist As String
    Dim CClist As String
    Dim BCClist As String
    Dim tmpStr As String
   
    If Not IsExistSheet("RECEIVER") Then
        Err.Number = "1001"
        Err.Description = "No Worksheet[RECEIVER]"
        Exit Function
    End If
    Set nowSheet = Worksheets("RECEIVER")
    For i = 2 To nowSheet.Rows.Count
        If Trim(nowSheet.Cells(i, 1)) <> "" And Trim(nowSheet.Cells(i, 2)) <> "" Then
            tmpStr = Trim(nowSheet.Cells(i, 1))
            If Right(tmpStr, 1) = "," Then tmpStr = Left(tmpStr, Len(tmpStr) - 1)   '去除最後面的逗點
            For j = 1 To 5
                tmpStr = Replace(tmpStr, "  ", " ") '去除連續的空白
                tmpStr = Replace(tmpStr, ", ", ",") '去除逗點後的空白
                tmpStr = Replace(tmpStr, " ,", ",") '去除逗點前的空白
            Next j
            tmpStr = Replace(tmpStr, " ", "_") '空白=>底線
         
            tmpStr = UCase(tmpStr)
         
            tmpStr = Replace(tmpStr, ",", "@umc.com ") '逗點=>@umc.com空白
            Select Case UCase(Trim(nowSheet.Cells(i, 2)))
                Case "TO":  TOlist = TOlist & " " & tmpStr & "@umc.com"
                Case "CC":  CClist = CClist & " " & tmpStr & "@umc.com"
                Case "BCC": BCClist = BCClist & " " & tmpStr & "@umc.com"
            End Select
        End If
    Next i
    TOlist = Trim(TOlist)
    CClist = Trim(CClist)
    BCClist = Trim(BCClist)
    getReceiver = TOlist & "," & CClist & "," & BCClist
End Function

Public Function GenReport_Customer(srcSheet As String, nowSPECSheet As String, TarSheet As String)
   Dim srcRange As Range, itemRange As Range, TarRange As Range
   Dim bCol As Long, bRow As Long
   Dim i As Long
   Dim Temp2 As Long
   
   Application.ScreenUpdating = False
   ' trim unused cells
   Call TrimSheet(nowSPECSheet)
   ' remove space of ITEM
   Call TrimCol(nowSPECSheet, 3)
   ' use MyFilter
   AddSheet (tempSheet)
   With Worksheets(srcSheet)
      Set srcRange = .Range("B3:" & N2L(.UsedRange.Columns.Count) & CStr(.UsedRange.Rows.Count))
   End With
   Set itemRange = Worksheets(nowSPECSheet).Range("C4:C" & CStr(Worksheets(nowSPECSheet).UsedRange.Rows.Count))
   Set TarRange = Worksheets(tempSheet).Range("A1")
   Call MyFilter(srcRange, itemRange, TarRange)
   Set TarRange = Worksheets(tempSheet).UsedRange
   
'   ' 舊作法: 用進階篩選, 因會有些不可預期的錯誤, 改用 MyFilter
'   ' use Temp Sheet to Generate value range
'   AddSheet (TempSheet)
'   With Worksheets(SrcSheet)
'      .Range("B3:" & N2L(.UsedRange.Columns.Count) & CStr(.UsedRange.Rows.Count)).Copy Worksheets(TempSheet).Range("A1")
'   End With
'   ' assign fake field name
'   For i = 6 To Worksheets(TempSheet).UsedRange.Columns.Count
'      Worksheets(TempSheet).UsedRange.Cells(1, i) = "Value_" & CStr(i - 5)
'   Next i
'   ' use Advnce Filter to get value
'   bRow = 1
'   bCol = Worksheets(TempSheet).UsedRange.Columns.Count + 2
'   Set oldRange = Worksheets(TempSheet).UsedRange
'   Set itemRange = Worksheets(nowSPECSheet).Range("C3:C" & CStr(Worksheets(nowSPECSheet).UsedRange.Rows.Count))
'   Temp2 = itemRange.Rows.Count
'   On Error Resume Next
'   oldRange.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=itemRange, _
'               CopyToRange:=Range(N2L(bCol) & CStr(bRow)), Unique:=False
'   On Error GoTo 0
'   ' set value range
'   With Worksheets(TempSheet)
'      Set itemRange = .Range(N2L(bCol) & CStr(bRow) & ":" & N2L(.UsedRange.Columns.Count) & CStr(.UsedRange.Rows.Count))
'   End With
'   For i = 1 To itemRange.Rows.Count
'      If Trim(itemRange.Cells(i, 1)) = "" Then Exit For
'   Next i
'   If i < itemRange.Rows.Count Then Set itemRange = itemRange.Range("A1:" & N2L(itemRange.Columns.Count) & CStr(i - 1))
'   ' Remove no match parameter
'   If itemRange.Rows.Count > Temp2 Then
'      For i = 2 To itemRange.Rows.Count
'         If itemRange.Cells(i, 1) = "" Then Exit For
'         If Not CheckParameter(nowSPECSheet, itemRange.Cells(i, 1)) Then
'            Worksheets(TempSheet).Range(N2L(bCol) & CStr(i) & ":" & N2L(bCol + itemRange.Columns.Count - 1) & CStr(i)).Delete Shift:=xlUp
'            i = i - 1
'         End If
'      Next i
'      For j = 1 To itemRange.Rows.Count
'         If Trim(itemRange.Cells(j, 1)) = "" Then Exit For
'      Next j
'      If j < itemRange.Rows.Count Then Set itemRange = itemRange.Range("A1:" & N2L(itemRange.Columns.Count) & CStr(j - 1))
'   End If
'   ' Check parameter match ?
'   GenReport_Customer = True
'   If itemRange.Rows.Count <> Temp2 Then
'      Worksheets(TempSheet).Range("GG2") = "Error! Maybe have some SPEC_Customer parameter be not found in SPEC parameter"
'      Set itemRange = Worksheets(TempSheet).Range("GA1:GG2")
'      GenReport_Customer = False
'   End If
   
   ' to piece Sheet SummaryCutomer
   Call AddSheet(TarSheet)
   ' Copy SPEC
   With Worksheets(nowSPECSheet)
      .Range("A4:" & N2L(8) & CStr(.UsedRange.Rows.Count)).Copy Worksheets(TarSheet).Range("A4")
   End With
   ' ITEM to add remark
   With Worksheets(TarSheet).UsedRange
      For i = 2 To .Rows.Count
         If Trim(.Cells(i, 4)) <> "" Then
            Temp2 = Len(.Cells(i, 3))
            .Cells(i, 3) = .Cells(i, 3) & " " & Trim(.Cells(i, 4))
            .Cells(i, 3).Characters(Start:=Temp2 + 2, length:=Len(Trim(.Cells(i, 4)))).Font.ColorIndex = 5
         End If
      Next i
   End With
   ' Delete Column Factor and Column Remark
   Worksheets(TarSheet).Range("D:D").Delete Shift:=xlShiftToLeft
   Worksheets(TarSheet).Range("B:B").Delete Shift:=xlShiftToLeft
   ' Copy Header
   With Worksheets(srcSheet)
      .Range("A1:" & N2L(.UsedRange.Columns.Count) & CStr(3)).Copy Worksheets(TarSheet).Range("A1")
   End With
   With Worksheets(nowSPECSheet)
      .Range("F2:H3").Copy Worksheets(TarSheet).Range("D2:F3")
   End With
   ' Copy Value
   'itemRange.Range("F2:" & N2L(itemRange.Columns.Count) & CStr(itemRange.Rows.Count)).Copy Worksheets(TarSheet).Range("G4")
   TarRange.Range("F1:" & N2L(TarRange.Columns.Count) & CStr(TarRange.Rows.Count)).Copy Worksheets(TarSheet).Range("G4")
   ' delete target column
   Worksheets(TarSheet).Range("G:G").Delete Shift:=xlShiftToLeft
   ' Delete Temp Sheet
   DelSheet (tempSheet)
   
   'Sheet format
   Worksheets(TarSheet).Activate
   Worksheets(TarSheet).Cells.Select
   Selection.Font.Size = 10
   Selection.Font.Name = "Century Gothic"
   Worksheets(TarSheet).Range("A1:" & N2L(Worksheets(TarSheet).UsedRange.Columns.Count) & CStr(3)).Select
   Selection.Font.Size = 12
   Worksheets(TarSheet).Range("A4:" & "A" & CStr(Worksheets(TarSheet).UsedRange.Rows.Count)).Select
   Selection.Font.Size = 12
   ActiveWindow.Zoom = 75
   Worksheets(TarSheet).Cells.Select
   Selection.Columns.AutoFit
   ActiveWindow.Zoom = 75
   Worksheets(TarSheet).Range("A4").Select
   ActiveWindow.FreezePanes = True
   
   Application.ScreenUpdating = True
End Function

Public Function Condition_RawData_Customer(nowSheet As String)
   Dim waferList() As String
   Dim i As Long, j As Long
   
   Application.ScreenUpdating = False
   'get Waferlist
   Call GetWaferList(dSheet, waferList)
   
   Worksheets(nowSheet).Activate
   For i = 4 To Worksheets(nowSheet).UsedRange.Rows.Count
      With Worksheets(nowSheet)
         .Range(N2L(7) & CStr(i) & ":" & N2L(7 + 8 + UBound(waferList) * 9) & CStr(i)).Select
         Selection.FormatConditions.Delete
         If Trim(.Cells(i, 4)) <> "" Or Trim(.Cells(i, 6)) <> "" Then
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & "F" & "" & CStr(i)
            Selection.FormatConditions(1).Font.ColorIndex = 3
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & "D" & "" & CStr(i)
            Selection.FormatConditions(2).Font.ColorIndex = 4
'            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, _
'                  Formula1:="=$" & "D" & "$" & CStr(i), Formula2:="=$" & "F" & "$" & CStr(i)
'            Selection.FormatConditions(1).Font.ColorIndex = 3
         End If
      End With
   Next i
   Worksheets(nowSheet).Range("A1").Select
   Application.ScreenUpdating = True
End Function

Public Function Condition_Summary_Customer(nowSheet As String)
   Dim waferList() As String
   Dim i As Long, j As Long
   Const HColCount = 4
   Const HCol = 3
   
   Application.ScreenUpdating = False
   'get Waferlist
   Call GetWaferList(dSheet, waferList)
   
   Worksheets(nowSheet).Activate
   For i = 4 To Worksheets(nowSheet).UsedRange.Rows.Count
      For j = 0 To UBound(waferList)
         With Worksheets(nowSheet)
            .Range(N2L(7 + j * HColCount) & CStr(i)).Select
            Selection.FormatConditions.Delete
            If Trim(.Cells(i, 4)) <> "" Or Trim(.Cells(i, 6)) <> "" Then
               Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & "F" & "" & CStr(i)
               Selection.FormatConditions(1).Font.ColorIndex = 3
               Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & "D" & "" & CStr(i)
               Selection.FormatConditions(2).Font.ColorIndex = 4
'               Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, _
'                     Formula1:="=$" & "D" & "$" & CStr(i), Formula2:="=$" & "F" & "$" & CStr(i)
'               Selection.FormatConditions(1).Font.ColorIndex = 3
            End If
         End With
      Next j
   Next i
   Worksheets(nowSheet).Range("A1").Select
   Application.ScreenUpdating = True
End Function

'------------------------
' Generate flitered data
'------------------------
Public Function GenDataRange(SourceSheet As String, nowSPECSheet As String, OutputSheet As String)
   Dim bCol As Long, bRow As Long, iCol As Long, iRow As Long
   Dim oldRange As Range, itemRange As Range, dataRange As Range
   Dim nowTempSheet As String
   Dim i As Long, j As Long
   Dim waferList() As String
   
   ' Trim space of parameter in sheet Data
   Call TrimCol(SourceSheet, 2)
   ' get Waferlist
   Call GetWaferList(dSheet, waferList)
   
   nowTempSheet = "Temp"
   AddSheet (nowTempSheet)
   Worksheets(SourceSheet).UsedRange.Copy Worksheets(nowTempSheet).Range("A1")
   Do While Trim(Worksheets(nowTempSheet).Range("B1")) <> "Parameter"
      Worksheets(nowTempSheet).Range("1:1").Delete Shift:=xlShiftUp
   Loop
   Worksheets(nowTempSheet).Range("A:A").Delete Shift:=xlShiftToLeft
   Worksheets(nowTempSheet).Cells(1, 1) = "ITEM"
   ' Fill WaferNum
   j = 0
   For i = 1 To Worksheets(nowTempSheet).UsedRange.Rows.Count
      If Trim(Worksheets(nowTempSheet).Cells(i, 2)) = "Unit" Then
         j = j + 1
      Else
         Worksheets(nowTempSheet).Cells(i, 2) = waferList(j - 1)
      End If
      If Trim(Worksheets(nowTempSheet).Cells(i, 1)) = "" Then _
         Worksheets(nowTempSheet).Cells(i, 1) = "Fake_Para"
   Next i
   Set oldRange = Worksheets(nowTempSheet).UsedRange
   Set itemRange = Worksheets(nowSPECSheet).Range("C3:C" & CStr(Worksheets(nowSPECSheet).UsedRange.Rows.Count))
   bRow = 1
   bCol = Worksheets(nowTempSheet).UsedRange.Columns.Count + 2
   On Error Resume Next
   oldRange.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=itemRange, _
               CopyToRange:=Range(N2L(bCol) & CStr(bRow)), Unique:=False
   On Error GoTo 0
   If Range(N2L(bCol) & CStr(bRow)) = "" Then
      Dim tmpStr As String, nowChkFile As String
      tmpStr = Application.ActiveWorkbook.Name
      nowChkFile = Application.ActiveWorkbook.Path & "\" & Left(tmpStr, Len(tmpStr) - 4) & ".err"
      Shell "cmd /c echo error >" & nowChkFile
      GenDataRange = False
      GoTo myEnd
   End If
   i = 0
   Do While Worksheets(tempSheet).Cells(bRow, bCol + i) <> ""
      i = i + 1
   Loop
   iCol = bCol + i - 1
   i = 0
   Do While Worksheets(tempSheet).Cells(bRow + i, bCol) <> ""
      i = i + 1
   Loop
   iRow = bRow + i - 1
   Set dataRange = Worksheets(tempSheet).Range(N2L(bCol) & CStr(bRow) & ":" & N2L(iCol) & CStr(iRow))
   AddSheet (OutputSheet)
   dataRange.Copy Worksheets(OutputSheet).Range("A1")
   Worksheets(OutputSheet).Visible = False
myEnd:
   Set dataRange = Nothing
   Set oldRange = Nothing
   Set itemRange = Nothing
   DelSheet (nowTempSheet)
End Function



Public Function CheckParameter(nowSheet As String, nowPara As String)
   Dim i As Long, j As Long
   Dim yn As Boolean
   
   yn = False
   For i = 4 To Worksheets(nowSheet).UsedRange.Rows.Count
      If Trim(Worksheets(nowSheet).Cells(i, 3)) = Trim(nowPara) Then
         yn = True
         Exit For
      End If
   Next i
   CheckParameter = yn
End Function



Private Sub CheckParameter3()
   Dim i As Long, j As Long
   Dim yn As Boolean
   
   For i = 4 To Worksheets(SPECCustomer).UsedRange.Rows.Count
      yn = False
      For j = 4 To Worksheets(SSheet).UsedRange.Rows.Count
         If Trim(Worksheets(SPECCustomer).Cells(i, 3)) = Trim(Worksheets(SSheet).Cells(j, 3)) Then
            yn = True
            Exit For
         End If
      Next j
      If yn = False Then Debug.Print Trim(Worksheets(SPECCustomer).Cells(i, 3))
   Next i
End Sub

Private Sub CheckParamete2r()
   Dim i As Long, j As Long
   Dim cn As Integer
   Dim nowSheet As String
   nowSheet = SSheet ' SSheet
   
   For i = 4 To Worksheets(nowSheet).UsedRange.Rows.Count
      cn = 0
      For j = 4 To Worksheets(nowSheet).UsedRange.Rows.Count
         If Trim(Worksheets(nowSheet).Cells(i, 3)) = Trim(Worksheets(nowSheet).Cells(j, 3)) Then
            cn = cn + 1
         End If
      Next j
      If cn <> 1 Then Debug.Print Trim(Worksheets(nowSheet).Cells(i, 3)) & " : " & CStr(cn)
   Next i
End Sub

Public Function MyFilter(srcRange As Range, itemRange As Range, TarRange As Range)
   Dim i As Long, j As Long
   
   itemRange.Copy TarRange.Range("A1")
   For i = 1 To itemRange.Rows.Count
      For j = 1 To srcRange.Rows.Count
         If Trim(itemRange.Cells(i, 1)) = Trim(srcRange.Cells(j, 1)) Then
            srcRange.Range("B" & CStr(j) & ":" & N2L(srcRange.Columns.Count) & CStr(j)).Copy TarRange.Range("B" & CStr(i))
            Exit For
         End If
      Next j
   Next i
   
End Function

Private Sub kkk()
   Dim srcRange As Range, itemRange As Range, TarRange As Range
   Const TempSheet2 = "Temp2"
   
   AddSheet (tempSheet)
   AddSheet (TempSheet2)
   
   With Worksheets(TSheet)
      '.Range("B3:" & N2L(.UsedRange.Columns.Count) & CStr(.UsedRange.Rows.Count)).Copy Worksheets(TempSheet).Range("A1")
      Set srcRange = .Range("B3:" & N2L(.UsedRange.Columns.Count) & CStr(.UsedRange.Rows.Count))
   End With
   'Set srcRange = Worksheets(TempSheet).UsedRange
   Set itemRange = Worksheets(SPECCustomer).Range("C4:C" & CStr(Worksheets(SPECCustomer).UsedRange.Rows.Count))
   Set TarRange = Worksheets(TempSheet2).Range("A1")
   
   Call MyFilter(srcRange, itemRange, TarRange)
   
   'delsheet()
End Sub



Public Function initSheet(mProduct As String)
   Dim i As Integer
   
   Application.DisplayAlerts = False
   For i = Worksheets.Count To 1 Step -1
      If InStr(Worksheets(i).Name, "-") > 0 Then
         If getCOL(Worksheets(i).Name, "-", 1) = Trim(mProduct) Then
            Worksheets(i).Name = getCOL(Worksheets(i).Name, "-", 2)
         Else
            Worksheets(i).Delete
         End If
      End If
   Next
   Application.DisplayAlerts = True
End Function

'============================
' Set FAC Name 要改
'============================
Public Function SetFACName(nowSheet As String)

   Dim i As Long, j As Long
   Dim mRange As Range
   
   Set mRange = Worksheets(nowSheet).UsedRange

   For i = 4 To mRange.Rows.Count
      If mRange.Cells(i, 2) <> "" And mRange.Cells(i, 3) <> "" Then
         mRange.Range(N2L(2) & CStr(i)).Name = Replace(Replace(mRange.Cells(i, 3), "-", "_"), "/", "_") & "_" & "FAC"
      End If
   Next i
End Function

'============================
' Set Range Name
'============================
Public Function SetRangeName(nowSheet As String)
   Dim nowWafer As String
   Dim siteNum As Integer
   Dim i As Long, j As Long
   Dim mRange As Range
   
   Set mRange = Worksheets(nowSheet).UsedRange
   siteNum = mRange.Columns.Count - 5
   For i = 11 To mRange.Rows.Count
      If mRange.Cells(i, 2) = "Parameter" Then
         nowWafer = getCOL(getCOL(mRange.Cells(i, 4), "-", 1), "<", 2)
      Else
         If mRange.Cells(i, 2) <> "" And InStr(mRange.Cells(i, 2), "/") < 1 And InStr(mRange.Cells(i, 2), "(") < 1 Then _
            mRange.Range(N2L(4) & CStr(i) & ":" & N2L(4 + siteNum - 1) & CStr(i)).Name = Replace(Replace(mRange.Cells(i, 2), "-", "_"), "/", "_") & "_" & nowWafer
      End If
   Next i
End Function

Public Function FormulaValue(ByRef itemArray() As String, ByRef valueArray() As String, nowWafer As String, siteNum As Integer, vSpec As specInfo, Optional LumpFunction As String = "")
   Dim i As Integer, j As Integer
  
   ReDim valueArray(UBound(itemArray))
   If LumpFunction = "" Then
      'Normal formula
      For i = 0 To UBound(valueArray)
         valueArray(i) = CStr(getValueByPara(nowWafer, itemArray(i), siteNum, vSpec))
      Next i
   Else
      'Lump function formula
      For i = 0 To UBound(valueArray)
         For j = 1 To siteNum
            valueArray(i) = valueArray(i) & "," & CStr(getValueByPara(nowWafer, itemArray(i), j, vSpec))
         Next j
         If Len(valueArray(i)) > 1 Then valueArray(i) = Mid(valueArray(i), 2)
         Select Case UCase(LumpFunction)
            Case "MEDIAN"
               valueArray(i) = myMedian(valueArray(i))
         End Select
      Next i
   End If
End Function

Public Function getValueByPara(nowWafer As String, nowPara As String, ByVal siteNum As Integer, vSpec As specInfo)
   Dim reValue
   Dim nowRange As Range
   Dim TargetSheet As Worksheet
   Dim waferRange As Range
   
   If IsNumeric(nowPara) Then
      getValueByPara = Val(nowPara)
      Exit Function
   End If
   
   Set TargetSheet = Worksheets("Data")
   Set waferRange = TargetSheet.Range("wafer_" & nowWafer)
   Set nowRange = waferRange.Range("B1:" & N2L(waferRange.Columns.Count) & waferRange.Rows.Count)
   On Error Resume Next
   reValue = Application.WorksheetFunction.VLookup(nowPara, nowRange, 2 + siteNum, False)
   
   'With FACTOR
   If Not IsEmpty(reValue) Then
      vSpec = getSPECInfo(nowPara)
      reValue = reValue * vSpec.mFAC
   End If
   getValueByPara = reValue
End Function
Public Function FormulaEval(ByVal strFormula As String, ByRef itemArray() As String, ByRef valueArray() As String)
   Dim i As Integer
   
   strFormula = getCOL(strFormula, ":", 1)
   For i = 0 To UBound(itemArray)
      strFormula = Replace(strFormula, itemArray(i), Trim(valueArray(i)))
   Next i
   FormulaEval = Application.Evaluate(strFormula)
   If Not IsNumeric(FormulaEval) Then FormulaEval = ""
   'Debug.Print strFormula
   'Debug.Print Application.Evaluate(strFormula)
End Function



Public Function IsEmptyValue(mSheet As String) As Boolean
   Dim nowSheet As Worksheet
   Dim nowCol As Long, nowRow As Long
   Dim xCount As Long, yCount As Long
   
   IsEmptyValue = False
   Set nowSheet = Worksheets(mSheet)
   For nowCol = 3 To nowSheet.UsedRange.Columns.Count Step 2
      If nowSheet.Cells(1, nowCol) <> "" Then
         xCount = Application.WorksheetFunction.countA(nowSheet.Range(N2L(nowCol) & CStr(3) & ":" & N2L(nowCol) & CStr(nowSheet.UsedRange.Rows.Count)))
         yCount = Application.WorksheetFunction.countA(nowSheet.Range(N2L(nowCol + 1) & CStr(3) & ":" & N2L(nowCol + 1) & CStr(nowSheet.UsedRange.Rows.Count)))
         'Debug.Print xCount; yCount
         If xCount = 0 Or yCount = 0 Then
            IsEmptyValue = True
            Exit For
         End If
      End If
   Next nowCol
End Function

Public Function AddChartFormat(mChart As Chart, mName As String, Optional mDescription As String = "By AutoReport")
   On Error Resume Next
   Application.AddChartAutoFormat Chart:=mChart, Name:=mName, Description:=mDescription
End Function

Public Function RemoveChartFormat(mName As String)
   On Error Resume Next
   Application.DeleteChartAutoFormat Name:=mName
End Function

Public Function getExStr(mSheet As String)
    Dim nowSheet As Worksheet
    Dim i As Long
   
    Set nowSheet = Worksheets("SPEC_List")
    For i = 1 To nowSheet.UsedRange.Columns.Count
        If Left(UCase(getCOL(nowSheet.Cells(1, i), ":", 1)), 31) = UCase(mSheet) Then _
            getExStr = getCOL(nowSheet.Cells(1, i), ":", 2): Exit For
    Next i
End Function

Public Function getDiffCase(ByVal mStr As String) As Integer
    
    Dim mKey1() As String
    Dim mKey2() As String
    Dim item As Variant

    If IsPrefix(mStr, "VT") Or _
       IsPrefix(mStr, "BF") Or _
       IsPrefix(mStr, "BV") Or _
       IsPrefix(mStr, "TOX") Or _
       IsPrefix(mStr, "PVT") Or _
       IsPrefix(mStr, "DVT") Or _
       IsPrefix(mStr, "DIBL") Or _
       IsPrefix(mStr, "SNM") Or _
       IsPrefix(mStr, "WMS") Then
       getDiffCase = 1 '相減
    ElseIf IsPrefix(mStr, "IOF") Or _
       IsPrefix(mStr, "IL") Or _
       IsPrefix(mStr, "IB") Or _
       IsPrefix(mStr, "Isb") Or _
       IsPrefix(mStr, "Jg") Or _
       IsPrefix(mStr, "TREND.") Then
       getDiffCase = 3 '倍數 x
    Else
       getDiffCase = 2 '百分比 %
    End If
    
    mKey1 = Split("VHF,VSF,VCT,JGG,GDL,GD1,GD2,GD5,IOFNOM,IST,LK_NOM,IIOF,HPIOF,IOFCA,ISL,ILKCA,ILKHA,ILKMA,TXI,IGD", ",")
    mKey2 = Split("VTL,VTS,DVT,BVI,BVD,PTV", ",")
    For Each item In mKey1
        If InStr(1, mStr, item) > 0 Then getDiffCase = 1
    Next item
    For Each item In mKey2
        If InStr(1, mStr, item) > 0 Then getDiffCase = 4
    Next item
        
End Function

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
Public Function About()
   FrmAbout.Show
End Function

Public Function TiNorm(mPara As String, mValue, Optional LSL As Double = -100, Optional Target As Double = 0, Optional USL As Double = 100)
   Dim oSpec As specInfo
   
   On Error GoTo 0
   
   mPara = Replace(mPara, "%", "")
   oSpec = getSPECInfo(mPara)
   'If IsEmpty(oSpec) Then TiNorm = 999: Exit Function
   If IsEmpty(oSpec.mLow) Or IsEmpty(oSpec.mTarget) Or IsEmpty(oSpec.mHigh) Then
      TiNorm = 9999
   Else
      If mValue >= Val(oSpec.mTarget) Then
         TiNorm = (mValue - Val(oSpec.mTarget)) / (Val(oSpec.mHigh) - Val(oSpec.mTarget)) * (USL - Target) + Target
      Else
         TiNorm = Target - (Val(oSpec.mTarget) - mValue) / (Val(oSpec.mTarget) - Val(oSpec.mLow)) * (Target - LSL)
      End If
      'TiNorm = 300
   End If
   
   'Debug.Print mPara, oSpec.mTarget
   
End Function

Public Function SetPosition(mShape As Variant, ByVal mCount As Integer, ByVal mIndex As Integer)
   Dim H As Integer, W As Integer
   Dim wCount As Integer
   
   H = mShape.Parent.Master.Height
   W = mShape.Parent.Master.width
   
   wCount = Int(Sqr(mCount) + 0.99)
   mShape.width = (W - 40 - 10) / wCount
   If mCount = 1 Then mShape.width = mShape.width - 200
   mShape.Top = 50 + ((mIndex - 1) \ wCount) * (mShape.Height + 10 / wCount)
   mShape.Left = 20 + ((mIndex - 1) Mod wCount) * (mShape.width + 10 / wCount)

   '細部調整位置
   '-----------------
   Select Case mCount
      Case 1
         mShape.Left = (W - mShape.width) / 2
         mShape.Top = mShape.Top + 30
      Case 2
         'mShape.Top = (H - mShape.Height) / 2
         mShape.Top = mShape.Top + 30
      Case 3
         If mIndex = 3 Then mShape.Left = (W - mShape.width) / 2
   End Select
   '-----------------
   
'   Select Case mCount
'      Case 1
'         mShape.Width = W - 50
'         'mShape.Height = H - 100 - 50
'      Case 2, 3, 4
'         mShape.Width = (W - 50 - 10) / 2
'      Case 5, 6
'         mShape.Width = (W - 50 - 10) / 3
'   End Select
   
'   mShape.Top = 80 + ((mIndex - 1) \ 2) * (mShape.Height + 10)
'   mShape.Left = 25 + ((mIndex + 1) Mod 2) * (mShape.Width + 10)
End Function

Public Function CornerCount()
   'Dim tArray(1, 4) As Double
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
   For i = 1 To Worksheets.Count
      If UCase(Left(Worksheets(i).Name, 7)) = "SCATTER" Then
         Set nowSheet = Worksheets(i)
         inCount = 0: OutCount = 0
         For iEnd = 0 To nowSheet.UsedRange.Rows.Count
            If nowSheet.Cells(iEnd + 1, 1) = "" And nowSheet.Cells(iEnd + 1, 2) = "" Then Exit For
         Next iEnd
         oInfo = getChartInfo(nowSheet.Range("A1:B" & CStr(iEnd)))
         'Debug.Print oInfo.ChartTitle
         If oInfo.vCornerXValueStr <> "" Then
            ReDim tArray(Len(oInfo.vCornerXValueStr) - Len(Replace(oInfo.vCornerXValueStr, ",", "")))
            For j = 0 To UBound(tArray)
               tArray(j) = Array(Val(getCOL(oInfo.vCornerXValueStr, ",", j + 1)), Val(getCOL(oInfo.vCornerYValueStr, ",", j + 1)))
            Next j
            'If tArray(1)(1) > tArray(3)(1) Then Call Swap(tArray(1), tArray(3))
            Call CornerSeq(tArray)
            For m = 3 To nowSheet.UsedRange.Columns.Count Step 2
               If IsKey("target,median,tt,ff,ss", nowSheet.Cells(1, m), ",", True) Then Exit For
               'Debug.Print nowSheet.Cells(1, m)
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
         'Debug.Print "In:"; inCount, "Out:"; OutCount
         If nowSheet.ChartObjects.Count > 0 Then
            If inCount > 0 Or OutCount > 0 Then
               Set nowChart = nowSheet.ChartObjects(1).Chart
               Set nowShape = nowChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 100, 20)
               nowSheet.Activate
               'nowSheet.Range("A1").Select
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
'    With Selection.Characters(Start:=1, Length:=26).Font
'        .Name = "Times New Roman"
'        .FontStyle = "標準"
'        .Size = 12
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ColorIndex = xlAutomatic
'    End With
'    Selection.ShapeRange.Fill.Visible = msoFalse
'    Selection.ShapeRange.ScaleWidth 1.16, msoFalse, msoScaleFromBottomRight
'    Selection.ShapeRange.ScaleHeight 1.45, msoFalse, msoScaleFromBottomRight
'    Selection.Font.ColorIndex = 3
'    Selection.ShapeRange.Line.Visible = msoFalse
'    Selection.ShapeRange.IncrementLeft -54#
'    Selection.ShapeRange.IncrementTop -249.6
    
         End If
        
        'Worksheets(i).Range("A1").Select
      End If
   Next i
'   'Debug.Print "OK"
'   Exit Function
'   tArray = Array(Array(100, 100), Array(500, 120), Array(600, 500), Array(150, 400), Array(100, 100))
'   Randomize
'   Debug.Print Time
'   For i = 1 To 20000
'      X = Rnd * 700
'      Y = Rnd * 700
'      'Debug.Print "X="; x, "Y="; y, ;
'      'Debug.Print "In Range: "; ynInRange(tArray, x, y)
'   Next i
'   Debug.Print Time
End Function

Public Function CornerCountLevel()
   'Dim tArray(1, 4) As Double
   Dim tArray(), sArray()
   Dim x As Double, y As Double
   Dim i As Integer, j As Integer, n As Integer, P As Integer
   Dim nowSheet As Worksheet
   Dim oInfo As chartInfo
   Dim iEnd As Long
   Dim inCount As Long, OutCount As Long
   Dim m As Long
   Dim nowChart As Chart
   Dim nowShape As Shape
   Dim countA() As Long
   Const cornerLevel As Integer = 3
   Dim tmpStr As String
   Dim C1 As Integer, C2 As Integer
   'ReDim countA(cornerLevel, 2)
   Dim cornerRange As Range
   
   'On Error Resume Next
   For i = 1 To Worksheets.Count
      If Left(Worksheets(i).Name, 7) = "SCATTER" Then
         Set nowSheet = Worksheets(i)
         'inCount = 0: OutCount = 0
         ReDim countA(cornerLevel, 2)
         'Get Chart Infomation
         '---------------------
         For iEnd = 0 To nowSheet.UsedRange.Rows.Count
            If nowSheet.Cells(iEnd + 1, 1) = "" And nowSheet.Cells(iEnd + 1, 2) = "" Then Exit For
         Next iEnd
         oInfo = getChartInfo(nowSheet.Range("A1:B" & CStr(iEnd)))
         'Debug.Print oInfo.ChartTitle
         'Get Corner Count
         '-----------------
         If oInfo.vCornerXValueStr <> "" Then
            ReDim tArray(Len(oInfo.vCornerXValueStr) - Len(Replace(oInfo.vCornerXValueStr, ",", "")))
            For j = 0 To UBound(tArray)
               tArray(j) = Array(Val(getCOL(oInfo.vCornerXValueStr, ",", j + 1)), Val(getCOL(oInfo.vCornerYValueStr, ",", j + 1)))
            Next j
            'If tArray(1)(1) > tArray(3)(1) Then Call Swap(tArray(1), tArray(3))
            Call CornerSeq(tArray)
            
            '--------------------------------------------------------------------
            For P = 1 To nowSheet.UsedRange.Columns.Count
               If UCase(nowSheet.Cells(1, P)) = "CORNER" Then Set cornerRange = nowSheet.Range(N2L(P) & CStr(1))
            Next P
            For P = 0 To cornerLevel - 1
               'level2 corner count
               ReDim sArray(UBound(tArray))
               For j = 0 To UBound(tArray)
                  C1 = j Mod 4
                  C2 = (j + 2) Mod 4
                  sArray(j) = Array(Val(tArray(C1)(0)) + (Val(tArray(C2)(0)) - Val(tArray(C1)(0))) / (cornerLevel * 2) * P, _
                                    Val(tArray(C1)(1)) + (Val(tArray(C2)(1)) - Val(tArray(C1)(1))) / (cornerLevel * 2) * P)
               Next j
'               sArray(0) = Array(Val(tArray(0)(0) + (tArray(2)(0) - tArray(0)(0)) / 6), Val(tArray(0)(1) + (tArray(2)(1) - tArray(0)(1)) / 6))
'               sArray(1) = Array(Val(tArray(1)(0) + (tArray(3)(0) - tArray(1)(0)) / 6), Val(tArray(1)(1) + (tArray(3)(1) - tArray(1)(1)) / 6))
'               sArray(2) = Array(Val(tArray(2)(0) - (tArray(2)(0) - tArray(0)(0)) / 6), Val(tArray(2)(1) - (tArray(2)(1) - tArray(0)(1)) / 6))
'               sArray(3) = Array(Val(tArray(3)(0) - (tArray(3)(0) - tArray(1)(0)) / 6), Val(tArray(3)(1) - (tArray(3)(1) - tArray(1)(1)) / 6))
'               sArray(4) = Array(Val(tArray(0)(0) + (tArray(2)(0) - tArray(0)(0)) / 6), Val(tArray(0)(1) + (tArray(2)(1) - tArray(0)(1)) / 6))
               'If tArray(1)(1) > tArray(3)(1) Then Call Swap(tArray(1), tArray(3))
               'Call CornerSeq(sArray)
'               Debug.Print sArray(0)(0), sArray(0)(1)
'               Debug.Print sArray(1)(0), sArray(1)(1)
'               Debug.Print sArray(2)(0), sArray(2)(1)
'               Debug.Print sArray(3)(0), sArray(3)(1)
'               Debug.Print sArray(4)(0), sArray(4)(1)
               If P > 0 Then
                  For j = 0 To UBound(sArray)
                     cornerRange.Cells(3 + 6 * P + j, 1) = sArray(j)(0)
                     cornerRange.Cells(3 + 6 * P + j, 2) = sArray(j)(1)
                  Next j
               End If
               For m = 3 To nowSheet.UsedRange.Columns.Count Step 2
                  If IsKey("target,median,tt,ff,ss", nowSheet.Cells(1, m), ",", True) Then Exit For
                  'Debug.Print nowSheet.Cells(1, m)
                  For n = 3 To nowSheet.UsedRange.Rows.Count
                     If nowSheet.Cells(n, m) <> "" And nowSheet.Cells(n, m + 1) <> "" Then
                        If ynInCorner(sArray, nowSheet.Cells(n, m), nowSheet.Cells(n, m + 1)) Then
                           countA(P + 1, 1) = countA(P + 1, 1) + 1
                        Else
                           countA(P + 1, 2) = countA(P + 1, 2) + 1
                        End If
                     End If
                  Next n
               Next m
               Debug.Print "level" & CStr(P + 1) & " In:"; countA(P + 1, 1), "Out:"; countA(P + 1, 2)
            Next P
'--------------------------------------------------------------------
'            For m = 3 To nowSheet.UsedRange.Columns.Count Step 2
'               If IsKey("target,median,tt,ff,ss,corner", nowSheet.Cells(1, m), ",", True) Then Exit For
'               'Debug.Print nowSheet.Cells(1, m)
'               For n = 3 To nowSheet.UsedRange.Rows.Count
'                  If nowSheet.Cells(n, m) <> "" And nowSheet.Cells(n, m + 1) <> "" Then
'                     If ynInCorner(tArray, nowSheet.Cells(n, m), nowSheet.Cells(n, m + 1)) Then
'                        countA(1, 1) = countA(1, 1) + 1
'                     Else
'                        countA(1, 2) = countA(1, 2) + 1
'                     End If
'                  End If
'               Next n
'            Next m
'            Debug.Print "level1 In:"; countA(1, 1), "Out:"; countA(1, 2)
'
'            'level2 corner count
'            ReDim sArray(UBound(tArray))
'            'For j = 0 To UBound(tArray)
'            '   sArray(j) = Array(Val(tArray(j)(0)), Val(tArray(j)(1)))
'            'Next j
'            sArray(0) = Array(Val(tArray(0)(0) + (tArray(2)(0) - tArray(0)(0)) / 6), Val(tArray(0)(1) + (tArray(2)(1) - tArray(0)(1)) / 6))
'            sArray(1) = Array(Val(tArray(1)(0) + (tArray(3)(0) - tArray(1)(0)) / 6), Val(tArray(1)(1) + (tArray(3)(1) - tArray(1)(1)) / 6))
'            sArray(2) = Array(Val(tArray(2)(0) - (tArray(2)(0) - tArray(0)(0)) / 6), Val(tArray(2)(1) - (tArray(2)(1) - tArray(0)(1)) / 6))
'            sArray(3) = Array(Val(tArray(3)(0) - (tArray(3)(0) - tArray(1)(0)) / 6), Val(tArray(3)(1) - (tArray(3)(1) - tArray(1)(1)) / 6))
'            sArray(4) = Array(Val(tArray(0)(0) + (tArray(2)(0) - tArray(0)(0)) / 6), Val(tArray(0)(1) + (tArray(2)(1) - tArray(0)(1)) / 6))
'            'If tArray(1)(1) > tArray(3)(1) Then Call Swap(tArray(1), tArray(3))
'            'Call CornerSeq(sArray)
''            Debug.Print sArray(0)(0), sArray(0)(1)
''            Debug.Print sArray(1)(0), sArray(1)(1)
''            Debug.Print sArray(2)(0), sArray(2)(1)
''            Debug.Print sArray(3)(0), sArray(3)(1)
''            Debug.Print sArray(4)(0), sArray(4)(1)
'            For m = 3 To nowSheet.UsedRange.Columns.Count Step 2
'               If IsKey("target,median,tt,ff,ss", nowSheet.Cells(1, m), ",", True) Then Exit For
'               'Debug.Print nowSheet.Cells(1, m)
'               For n = 3 To nowSheet.UsedRange.Rows.Count
'                  If nowSheet.Cells(n, m) <> "" And nowSheet.Cells(n, m + 1) <> "" Then
'                     If ynInCorner(sArray, nowSheet.Cells(n, m), nowSheet.Cells(n, m + 1)) Then
'                        countA(2, 1) = countA(2, 1) + 1
'                     Else
'                        countA(2, 2) = countA(2, 2) + 1
'                     End If
'                  End If
'               Next n
'            Next m
'            Debug.Print "level2 In:"; countA(2, 1), "Out:"; countA(2, 2)
'
'            'level3 corner count
'            ReDim sArray(UBound(tArray))
'            'For j = 0 To UBound(tArray)
'            '   sArray(j) = Array(Val(tArray(j)(0)), Val(tArray(j)(1)))
'            'Next j
'            sArray(0) = Array(Val(tArray(0)(0) + (tArray(2)(0) - tArray(0)(0)) / 6 * 2), Val(tArray(0)(1) + (tArray(2)(1) - tArray(0)(1)) / 6 * 2))
'            sArray(1) = Array(Val(tArray(1)(0) + (tArray(3)(0) - tArray(1)(0)) / 6 * 2), Val(tArray(1)(1) + (tArray(3)(1) - tArray(1)(1)) / 6 * 2))
'            sArray(2) = Array(Val(tArray(2)(0) - (tArray(2)(0) - tArray(0)(0)) / 6 * 2), Val(tArray(2)(1) - (tArray(2)(1) - tArray(0)(1)) / 6 * 2))
'            sArray(3) = Array(Val(tArray(3)(0) - (tArray(3)(0) - tArray(1)(0)) / 6 * 2), Val(tArray(3)(1) - (tArray(3)(1) - tArray(1)(1)) / 6 * 2))
'            sArray(4) = Array(Val(tArray(0)(0) + (tArray(2)(0) - tArray(0)(0)) / 6 * 2), Val(tArray(0)(1) + (tArray(2)(1) - tArray(0)(1)) / 6 * 2))
'            'If tArray(1)(1) > tArray(3)(1) Then Call Swap(tArray(1), tArray(3))
'            'Call CornerSeq(sArray)
''            Debug.Print sArray(0)(0), sArray(0)(1)
''            Debug.Print sArray(1)(0), sArray(1)(1)
''            Debug.Print sArray(2)(0), sArray(2)(1)
''            Debug.Print sArray(3)(0), sArray(3)(1)
''            Debug.Print sArray(4)(0), sArray(4)(1)
'            For m = 3 To nowSheet.UsedRange.Columns.Count Step 2
'               If IsKey("target,median,tt,ff,ss", nowSheet.Cells(1, m), ",", True) Then Exit For
'               'Debug.Print nowSheet.Cells(1, m)
'               For n = 3 To nowSheet.UsedRange.Rows.Count
'                  If nowSheet.Cells(n, m) <> "" And nowSheet.Cells(n, m + 1) <> "" Then
'                     If ynInCorner(sArray, nowSheet.Cells(n, m), nowSheet.Cells(n, m + 1)) Then
'                        countA(3, 1) = countA(3, 1) + 1
'                     Else
'                        countA(3, 2) = countA(3, 2) + 1
'                     End If
'                  End If
'               Next n
'            Next m
'            Debug.Print "level3 In:"; countA(3, 1), "Out:"; countA(3, 2)
           
         End If
         
         tmpStr = "L1 In: " & countA(1, 1) & " Out: " & countA(1, 2) & " = " & Format(countA(1, 1) / (countA(1, 1) + countA(1, 2)), "0.00%") & vbCrLf & _
                  "L2 In: " & countA(2, 1) & " Out: " & countA(2, 2) & " = " & Format(countA(2, 1) / (countA(2, 1) + countA(2, 2)), "0.00%") & vbCrLf & _
                  "L3 In: " & countA(3, 1) & " Out: " & countA(3, 2) & " = " & Format(countA(3, 1) / (countA(3, 1) + countA(3, 2)), "0.00%") & vbCrLf
         
         If nowSheet.ChartObjects.Count > 0 Then
            If countA(1, 1) + countA(1, 2) > 0 Then
               Set nowChart = nowSheet.ChartObjects(1).Chart
               Set nowShape = nowChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 100, 20)
               nowSheet.Activate
               'nowSheet.Range("A1").Select
               nowChart.Parent.Activate
               nowShape.Select
               With Selection
                  '.Characters.Text = "In: " & CStr(countA(1, 1)) & " Out: " & CStr(countA(1, 2)) & " = " & Format(countA(1, 1) / (countA(1, 1) + countA(1, 2)), "0.00%")
                  .Characters.Text = tmpStr
                  .Characters.Font.Size = 10
                  .Font.ColorIndex = 5
                  .Font.Bold = True
                  .AutoSize = True
               End With
               With nowShape
                  
                  .Top = nowChart.PlotArea.Top + 12
                  .Left = nowChart.PlotArea.Left + 40
               End With
            End If
            For j = 8 To nowSheet.ChartObjects(1).Chart.SeriesCollection("Corner").Points.Count
               'nowSheet.ChartObjects(1).Chart.SeriesCollection("Corner").Name
               On Error Resume Next
               nowSheet.ChartObjects(1).Chart.SeriesCollection("Corner").Points(j).Border.LineStyle = xlDash
            Next j
'    With Selection.Characters(Start:=1, Length:=26).Font
'        .Name = "Times New Roman"
'        .FontStyle = "標準"
'        .Size = 12
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ColorIndex = xlAutomatic
'    End With
'    Selection.ShapeRange.Fill.Visible = msoFalse
'    Selection.ShapeRange.ScaleWidth 1.16, msoFalse, msoScaleFromBottomRight
'    Selection.ShapeRange.ScaleHeight 1.45, msoFalse, msoScaleFromBottomRight
'    Selection.Font.ColorIndex = 3
'    Selection.ShapeRange.Line.Visible = msoFalse
'    Selection.ShapeRange.IncrementLeft -54#
'    Selection.ShapeRange.IncrementTop -249.6
    
         End If
        
    
      End If
   Next i
   'Debug.Print "OK"
   Exit Function
   tArray = Array(Array(100, 100), Array(500, 120), Array(600, 500), Array(150, 400), Array(100, 100))
   Randomize
   Debug.Print Time
   For i = 1 To 20000
      x = Rnd * 700
      y = Rnd * 700
      'Debug.Print "X="; x, "Y="; y, ;
      'Debug.Print "In Range: "; ynInRange(tArray, x, y)
   Next i
   Debug.Print Time
End Function

Public Function CornerSeq(ByRef tArray)
   If (getAngle(tArray(0)(0), tArray(0)(1), tArray(2)(0), tArray(2)(1)) + 360 - _
      getAngle(tArray(0)(0), tArray(0)(1), tArray(1)(0), tArray(1)(1))) Mod 360 > 180 Then _
      Call Swap(tArray(1), tArray(3))
End Function

Public Function getAngle(x1, y1, x2, y2)
   Dim tmpValue As Integer
   Select Case (x2 - x1)
      Case 0:
         Select Case (y2 - y1)
            Case 0:  getAngle = -1
            Case Is > 0: getAngle = 90
            Case Is < 0: getAngle = 270
         End Select
         Exit Function
      Case Is < 0:
         tmpValue = 180
   End Select
   If (x2 - x1) <> 0 Then
      getAngle = Atn((y2 - y1) / (x2 - x1)) * 180 / pi
   End If
   getAngle = (getAngle + 360 + tmpValue) Mod 360
End Function

Public Sub Fu_Table()
   Dim nowSheet As Worksheet
   Dim i As Long, j As Long
   Dim iCol As Long, iRow As Long, iWafer As Integer
   Dim waferList() As String
   Dim reValue
   Dim lotID As String
   Dim nowRange As Range
   
   Set nowSheet = ActiveSheet
   iCol = 6
   lotID = Replace(Worksheets(dSheet).Cells(3, 2), ":", "")
   Call GetWaferList(dSheet, waferList)
   nowSheet.Activate
   For iWafer = 0 To UBound(waferList)
      For iRow = 1 To nowSheet.UsedRange.Rows.Count
         If UCase(CStr(nowSheet.Cells(iRow, 4))) = "TARGET" Then
            nowSheet.Cells(iRow, iCol + iWafer) = lotID & "#" & waferList(iWafer)
            If iWafer = UBound(waferList) Then _
               nowSheet.Range(N2L(iCol + iWafer + 1) & CStr(iRow) & ":" & N2L(nowSheet.UsedRange.Columns.Count) & CStr(iRow)).ClearContents
         Else
            Set reValue = getRangeByPara(waferList(iWafer), nowSheet.Cells(iRow, 2))
            If Not reValue Is Nothing Then
               Set nowRange = reValue
               nowSheet.Cells(iRow, iCol + iWafer) = Application.WorksheetFunction.Median(nowRange)
               If iWafer = UBound(waferList) Then _
                  nowSheet.Range(N2L(iCol + iWafer + 1) & CStr(iRow) & ":" & N2L(nowSheet.UsedRange.Columns.Count) & CStr(iRow)).ClearContents
            End If
         End If
      Next iRow
   Next iWafer

' Set reValue = getRangeByPara(WaferList(i), nowParameter)
'            If Not reValue Is Nothing Then
'               Set nowRange = reValue
   DoEvents
   MsgBox "Finish!!!"
End Sub

Public Function RawdataRange()
'   'Dim tArray(1, 4) As Double
'   Dim tArray()
   Dim x As Double, y As Double
   Dim i As Integer, j As Integer
   Dim nowSheet As Worksheet
   Dim oInfo As chartInfo
   Dim iEnd As Long
   Dim iSheet As Integer
'   Dim inCount As Long, OutCount As Long
'   Dim m As Long, n As Long
   Dim nowChart As Chart
   Dim nowShape As Shape
   Dim xMin, xMax, yMin, yMax, xMed, yMed
   Dim nowSeries As Series
   'Dim nowRange As Range
   Dim xData() As Variant
   Dim yData() As Variant
   
   On Error Resume Next
   For iSheet = 1 To Worksheets.Count
      If Left(Worksheets(iSheet).Name, 7) = "SCATTER" Then
         Set nowSheet = Worksheets(iSheet)
         For iEnd = 0 To nowSheet.UsedRange.Rows.Count
            If nowSheet.Cells(iEnd + 1, 1) = "" And nowSheet.Cells(iEnd + 1, 2) = "" Then Exit For
         Next iEnd
         oInfo = getChartInfo(nowSheet.Range("A1:B" & CStr(iEnd)))
         'Debug.Print oInfo.ChartTitle
         If IsKey(oInfo.ChartExpression, "range", "+") Then
            Debug.Print nowSheet.Name
            Set nowChart = nowSheet.ChartObjects(1).Chart
            ReDim xData(0): ReDim yData(0)
            For i = 1 To nowChart.SeriesCollection.Count
               'For j = 1 To nowChart.SeriesCollection(i).Points.Count
               'Debug.Print "X="; Application.WorksheetFunction.Max(nowChart.SeriesCollection(i).XValues)
               'Debug.Print "Y="; Application.WorksheetFunction.Max(nowChart.SeriesCollection(i).Values)
               'Next j
               If Not IsKey("target,median,corner,tt,ff,ss,range", nowChart.SeriesCollection(i).Name, ",", True) Then
                  For j = 1 To UBound(nowChart.SeriesCollection(i).XValues)
                     xData(UBound(xData)) = Application.WorksheetFunction.index(nowChart.SeriesCollection(i).XValues, j)
                     ReDim Preserve xData(UBound(xData) + 1)
                     yData(UBound(yData)) = Application.WorksheetFunction.index(nowChart.SeriesCollection(i).Values, j)
                     ReDim Preserve yData(UBound(yData) + 1)
                  Next j
                  
'                  If i = 1 Then
''                     Debug.Print typeName(nowChart.SeriesCollection(i).XValues)
''                     Debug.Print IsArray(nowChart.SeriesCollection(i).XValues)
''                     Debug.Print Application.WorksheetFunction.Small(nowChart.SeriesCollection(i).XValues, 50)
'                     xMax = Application.WorksheetFunction.Max(nowChart.SeriesCollection(i).XValues)
'                     xMin = Application.WorksheetFunction.Min(nowChart.SeriesCollection(i).XValues)
'                     yMax = Application.WorksheetFunction.Max(nowChart.SeriesCollection(i).Values)
'                     yMin = Application.WorksheetFunction.Min(nowChart.SeriesCollection(i).Values)
'                  Else
'                     If Application.WorksheetFunction.Max(nowChart.SeriesCollection(i).XValues) > xMax Then _
'                        xMax = Application.WorksheetFunction.Max(nowChart.SeriesCollection(i).XValues)
'                     If Application.WorksheetFunction.Min(nowChart.SeriesCollection(i).XValues) < xMin Then _
'                        xMin = Application.WorksheetFunction.Min(nowChart.SeriesCollection(i).XValues)
'                     If Application.WorksheetFunction.Max(nowChart.SeriesCollection(i).Values) > yMax Then _
'                        yMax = Application.WorksheetFunction.Max(nowChart.SeriesCollection(i).Values)
'                     If Application.WorksheetFunction.Min(nowChart.SeriesCollection(i).Values) < yMin Then _
'                        yMin = Application.WorksheetFunction.Min(nowChart.SeriesCollection(i).Values)
'                  End If
               End If
            Next i
            ReDim Preserve xData(UBound(xData) - 1)
            ReDim Preserve yData(UBound(yData) - 1)
            xMax = Application.WorksheetFunction.Max(xData)
            xMin = Application.WorksheetFunction.Min(xData)
            xMed = Application.WorksheetFunction.Median(xData)
            yMax = Application.WorksheetFunction.Max(yData)
            yMin = Application.WorksheetFunction.Min(yData)
            yMed = Application.WorksheetFunction.Median(yData)
            Debug.Print "xMax="; xMax, "xMin="; xMin, "xMed="; xMed, "yMax="; yMax, "yMin="; yMin, "yMed="; yMed
            
            'Text Label
            '--------------
            'Set nowChart = nowSheet.ChartObjects(1).Chart
            Set nowShape = nowChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 100, 20)
            nowSheet.Activate
            'nowSheet.Range("A1").Select
            nowChart.Parent.Activate
            nowShape.Select
            With Selection
               '.Characters.Text = "X=> Med:" & CStr(xMed) & " Max:" & CStr(xMax) & " Min:" & CStr(xMin) & " %:" & Format((xMin - xMed) / xMed, "0.00%") & "~" & Format((xMax - xMed) / xMed, "0.00%")
               .Characters.Text = "X Range%: " & Format((xMin - xMed) / xMed, "0.0%") & " ~ " & Format((xMax - xMed) / xMed, "0.0%") & vbCrLf & _
                                  "Y Range%: " & Format((yMin - yMed) / yMed, "0.0%") & " ~ " & Format((yMax - yMed) / yMed, "0.0%")
               .Characters.Font.Size = 10
               .Font.ColorIndex = 5
               .Font.Bold = True
               .AutoSize = True
            End With
            With nowShape
               .Top = nowChart.PlotArea.InsideTop + nowChart.PlotArea.InsideHeight - nowShape.Height
               .Left = nowChart.PlotArea.InsideLeft + nowChart.PlotArea.InsideWidth - nowShape.width
            End With
            For i = nowChart.SeriesCollection.Count To 1 Step -1
               If UCase(nowChart.SeriesCollection(i).Name) = "RANGE" Then nowChart.SeriesCollection(i).Delete
            Next i
            
            Set nowSeries = nowChart.SeriesCollection.NewSeries
            With nowSeries
               .Name = "Range"
               .XValues = "={" & CStr(xMin) & "," & CStr(xMax) & "," & CStr(xMax) & "," & CStr(xMin) & "," & CStr(xMin) & "}"
               .Values = "={" & CStr(yMin) & "," & CStr(yMin) & "," & CStr(yMax) & "," & CStr(yMax) & "," & CStr(yMin) & "}"
               .Border.ColorIndex = 5
               .Border.Weight = xlHairline
               .Border.LineStyle = xlContinuous
               .MarkerStyle = xlNone
            End With
         End If
      End If
   Next iSheet
   'Debug.Print "OK"
   Exit Function
End Function

Public Function IsInvalidMail(mStr As String)
    Const mList As String = "Jacob_Zeng,CRD_Logic-45NM_1-INT2,Juny_Ning"
    IsInvalidMail = IsKey(mList, mStr, ",")
End Function

Public Function SetHiddenOption(ByVal mItem As String, mList As Variant, Optional ByVal mSheet As String = "HiddenOption")
    Dim nowSheet As Worksheet
    Dim tempA As Variant
    Dim iCol As Long, i As Long, s As Long
    
    Set nowSheet = AddSheet(mSheet, False)
    nowSheet.Visible = xlSheetHidden
    
    Select Case LCase(typeName(mList))
        Case "string":
            tempA = Split(mList, ",")
            s = 0
        Case "variant()":
            tempA = mList
            s = 1
        Case "string()":
            tempA = mList
            s = 0
    End Select
    
    If nowSheet.Cells(1, 1) = "" Then
        iCol = 1
    Else
        For iCol = 1 To nowSheet.UsedRange.Columns.Count
            If nowSheet.Cells(1, iCol) = mItem Then
                nowSheet.Columns(iCol).Clear
                Exit For
            End If
            If nowSheet.Cells(1, iCol) = "" Then Exit For
        Next iCol
    End If
    
    nowSheet.Cells(1, iCol) = mItem
    For i = s To UBound(tempA)
        nowSheet.Cells(i + 2, iCol) = tempA(i)
    Next i
End Function

Public Function GetHiddenOption(ByVal mItem As String, Optional ByVal mSheet As String = "HiddenOption")
    Dim nowSheet As Worksheet
    Dim tempA() As String
    Dim iCol As Long, i As Long, j As Long
    
    Set nowSheet = Worksheets(mSheet)
    
    ReDim tempA(0)
    For iCol = 1 To nowSheet.UsedRange.Columns.Count
            If nowSheet.Cells(1, iCol) = mItem Then
                ReDim tempA(WorksheetFunction.countA(nowSheet.Columns(iCol)) - 2)
                'Debug.Print UBound(TempA)
                j = 0
                For i = 2 To nowSheet.Rows.Count
                    If nowSheet.Cells(i, iCol) <> "" Then
                        tempA(j) = nowSheet.Cells(i, iCol)
                        j = j + 1
                    End If
                Next i
                
                Exit For
            End If
    Next iCol

    GetHiddenOption = tempA
End Function

Public Function getHeaderCol(mStr As Variant, ByVal mKey As String)  '20100111 - get key in exstr column number
    Dim i As Integer
    
    getHeaderCol = -1
    For i = 0 To UBound(mStr)
        If UCase(mStr(i)) = UCase(mKey) Then getHeaderCol = i + 1: Exit For
    Next i
End Function

Public Sub getCoord(ByVal HttpURL As String)
    
    Dim nowSheet As Worksheet
    Dim dSheet As Worksheet
    Dim nowRange As Range
    Dim i As Integer, j As Integer
    Dim mRow As Long, mCol As Integer
    Dim mKey As String
    Dim myText
    Dim RowsCount As Integer
    Dim Text
    Dim ArrayMap
    Dim tempStr
    Dim mProduct_ID As String
    
    Set nowSheet = AddSheet("CoorSetting", False)
    Set dSheet = Worksheets("Data")
    
    nowSheet.Visible = xlSheetVisible
        
    mRow = 1: mCol = 1
    Do While Not (dSheet.Cells(mRow, mCol) = "" And dSheet.Cells(mRow + 1, mCol) = "")
        If dSheet.Cells(mRow, mCol) = "<Product_ID>" Then mProduct_ID = Mid(Trim(dSheet.Cells(mRow, mCol + 1)), 2)
        mRow = mRow + 1
    Loop
    
    tempStr = Split(HttpURL, "_")
    Select Case UCase(Replace(tempStr(UBound(tempStr)), ".waf", ""))
        Case "DEFAULT", "9P", "21P", "ALL"
            mKey = "ALL"
            GoSub mySub
        Case "ELAB"
            mKey = "ELAB"
            GoSub mySub
        Case Else
            On Error GoTo 0
            myText = getURL("http://p58rptdev01/WATWaferMap/WATMap.aspx?PRODUCT=" & Trim(mProduct_ID) & "&WAF=" & Trim(HttpURL))
            myText = Mid(myText, InStr(myText, "<td align"))
            myText = Left(myText, InStr(myText, "</table>") - 1)
            myText = Split(myText, "</tr><tr>")
            
            Text = myText(0)
            Do While InStr(Text, "<td")
                Text = Mid(Text, InStr(Text, "<td") + 4)
                RowsCount = RowsCount + 1
            Loop
            
            ReDim ArrayMap(UBound(myText))
            For i = 0 To UBound(myText)
                ArrayMap(i) = Split(myText(i), "</td>")
                For j = 0 To RowsCount
                    ArrayMap(i)(j) = Replace(IIf(Left(Replace(getCOL(ArrayMap(i)(j), ">", 2), "&nbsp;", ""), 1) = "N", "", Replace(getCOL(ArrayMap(i)(j), ">", 2), "&nbsp;", "")), "<br/", "")
                Next j
            Next
            
            nowSheet.Cells(1, 1).CurrentRegion.ClearContents
            
            Select Case Right(tempStr(0), 1)
                Case "Y"
                    For i = 0 To UBound(ArrayMap)
                        For j = 0 To UBound(ArrayMap(0)) - 1
                            nowSheet.Cells(i + 1, j + 1).Value = ArrayMap(i)(j)
                        Next j
                    Next i
                Case "X"
                    For i = 0 To UBound(ArrayMap)
                        For j = 0 To UBound(ArrayMap(0)) - 1
                            If i = 0 And j > 0 Then
                                nowSheet.Cells(j + 1, i + 1).Value = -1 * ArrayMap(i)(j)
                            Else
                                nowSheet.Cells(j + 1, i + 1).Value = ArrayMap(i)(j)
                            End If
                        Next j
                    Next i
            End Select
            
            Set nowRange = nowSheet.Cells(1, 1).CurrentRegion
    End Select
    
    'nowSheet.Visible = xlSheetHidden
    Call getCoorSub(nowRange)
Exit Sub

mySub:
    mRow = 1: mCol = 1
    Do While Not nowSheet.Cells(mRow, mCol) = mKey
        mRow = mRow + 1
    Loop
    Set nowRange = nowSheet.Cells(mRow, mCol).CurrentRegion
    Return
        
End Sub

Public Function getCoorSub(mRange As Range)

    Dim nowSheet As Worksheet
    Dim mRow As Long, mCol As Long
    Dim i As Integer
    
    Set nowSheet = Worksheets("Data")

    mRow = 1: mCol = 1
    Do While Not (nowSheet.Cells(mRow, mCol) = "" And nowSheet.Cells(mRow + 1, mCol) = "")
        If Left(nowSheet.Cells(mRow, mCol), 3) = "No." Then
            i = 4
            waferNum = getCOL(getCOL(nowSheet.Cells(mRow, i), "<", 2), "-", 1)
            Do While Not nowSheet.Cells(mRow, i) = "W L"
                nowSheet.Cells(mRow, i) = Coordinate(mRange, i - 3, waferNum)
                i = i + 1
            Loop
        End If
        mRow = mRow + 1
    Loop
End Function

Public Sub GenCoorSheet()

    Dim nowSheet As Worksheet
    
    Set nowSheet = AddSheet("CoorSetting")
    
    With nowSheet
        .Cells(15, 1) = "ALL"
        .Cells(16, 1) = "6"
        .Cells(17, 1) = "5"
        .Cells(18, 1) = "4"
        .Cells(19, 1) = "3"
        .Cells(20, 1) = "2"
        .Cells(21, 1) = "1"
        .Cells(22, 1) = "0"
        .Cells(23, 1) = "-1"
        .Cells(24, 1) = "-2"
        .Cells(25, 1) = "-3"
        .Cells(26, 1) = "-4"
        .Cells(27, 1) = "-5"
        .Cells(29, 1) = "NALL"
        .Cells(30, 1) = "6"
        .Cells(31, 1) = "5"
        .Cells(32, 1) = "4"
        .Cells(33, 1) = "3"
        .Cells(34, 1) = "2"
        .Cells(35, 1) = "1"
        .Cells(36, 1) = "0"
        .Cells(37, 1) = "-1"
        .Cells(38, 1) = "-2"
        .Cells(39, 1) = "-3"
        .Cells(40, 1) = "-4"
        .Cells(41, 1) = "-5"
        
        .Cells(15, 2) = "-5"
        .Cells(29, 2) = "-5"
        .Cells(33, 2) = "N834"
        .Cells(34, 2) = "N835"
        .Cells(35, 2) = "N836"
        .Cells(36, 2) = "N837"
        .Cells(37, 2) = "N838"
        .Cells(38, 2) = "N839"
        
        .Cells(15, 3) = "-4"
        .Cells(20, 3) = "24"
        .Cells(21, 3) = "65"
        .Cells(22, 3) = "20"
        .Cells(23, 3) = "66"
        .Cells(29, 3) = "-4"
        .Cells(31, 3) = "N828"
        .Cells(32, 3) = "N829"
        .Cells(33, 3) = "N830"
        .Cells(34, 3) = "N63"
        .Cells(35, 3) = "N64"
        .Cells(36, 3) = "N65"
        .Cells(37, 3) = "N66"
        .Cells(38, 3) = "N831"
        .Cells(39, 3) = "N832"
        .Cells(40, 3) = "N833"
        
        .Cells(15, 4) = "-3"
        .Cells(18, 4) = "19"
        .Cells(19, 4) = "62"
        .Cells(20, 4) = "40"
        .Cells(21, 4) = "63"
        .Cells(22, 4) = "1"
        .Cells(23, 4) = "64"
        .Cells(24, 4) = "41"
        .Cells(25, 4) = "21"
        .Cells(29, 4) = "-3"
        .Cells(30, 4) = "N824"
        .Cells(31, 4) = "N825"
        .Cells(32, 4) = "N55"
        .Cells(33, 4) = "N56"
        .Cells(34, 4) = "N57"
        .Cells(35, 4) = "N58"
        .Cells(36, 4) = "N59"
        .Cells(37, 4) = "N60"
        .Cells(38, 4) = "N61"
        .Cells(39, 4) = "N62"
        .Cells(40, 4) = "N826"
        .Cells(41, 4) = "N827"
        
        .Cells(15, 5) = "-2"
        .Cells(17, 5) = "38"
        .Cells(18, 5) = "7"
        .Cells(19, 5) = "26"
        .Cells(20, 5) = "59"
        .Cells(21, 5) = "39"
        .Cells(22, 5) = "2"
        .Cells(23, 5) = "27"
        .Cells(24, 5) = "60"
        .Cells(25, 5) = "8"
        .Cells(26, 5) = "61"
        .Cells(29, 5) = "-2"
        .Cells(30, 5) = "N822"
        .Cells(31, 5) = "N45"
        .Cells(32, 5) = "N46"
        .Cells(33, 5) = "N47"
        .Cells(34, 5) = "N48"
        .Cells(35, 5) = "N49"
        .Cells(36, 5) = "N50"
        .Cells(37, 5) = "N51"
        .Cells(38, 5) = "N52"
        .Cells(39, 5) = "N53"
        .Cells(40, 5) = "N54"
        .Cells(41, 5) = "N823"
        
        .Cells(15, 6) = "-1"
        .Cells(17, 6) = "32"
        .Cells(18, 6) = "57"
        .Cells(19, 6) = "33"
        .Cells(20, 6) = "12"
        .Cells(21, 6) = "34"
        .Cells(22, 6) = "35"
        .Cells(23, 6) = "36"
        .Cells(24, 6) = "13"
        .Cells(25, 6) = "58"
        .Cells(26, 6) = "37"
        .Cells(29, 6) = "-1"
        .Cells(30, 6) = "N820"
        .Cells(31, 6) = "N35"
        .Cells(32, 6) = "N36"
        .Cells(33, 6) = "N37"
        .Cells(34, 6) = "N38"
        .Cells(35, 6) = "N39"
        .Cells(36, 6) = "N40"
        .Cells(37, 6) = "N41"
        .Cells(38, 6) = "N42"
        .Cells(39, 6) = "N43"
        .Cells(40, 6) = "N44"
        .Cells(41, 6) = "N821"
        
        .Cells(15, 7) = "0"
        .Cells(17, 7) = "18"
        .Cells(18, 7) = "54"
        .Cells(19, 7) = "25"
        .Cells(20, 7) = "55"
        .Cells(21, 7) = "31"
        .Cells(22, 7) = "3"
        .Cells(23, 7) = "28"
        .Cells(24, 7) = "56"
        .Cells(25, 7) = "43"
        .Cells(26, 7) = "14"
        .Cells(29, 7) = "0"
        .Cells(30, 7) = "N818"
        .Cells(31, 7) = "N25"
        .Cells(32, 7) = "N26"
        .Cells(33, 7) = "N27"
        .Cells(34, 7) = "N28"
        .Cells(35, 7) = "N29"
        .Cells(36, 7) = "N30"
        .Cells(37, 7) = "N31"
        .Cells(38, 7) = "N32"
        .Cells(39, 7) = "N33"
        .Cells(40, 7) = "N34"
        .Cells(41, 7) = "N819"
        
        .Cells(15, 8) = "1"
        .Cells(17, 8) = "49"
        .Cells(18, 8) = "6"
        .Cells(19, 8) = "50"
        .Cells(20, 8) = "11"
        .Cells(21, 8) = "51"
        .Cells(22, 8) = "4"
        .Cells(23, 8) = "52"
        .Cells(24, 8) = "10"
        .Cells(25, 8) = "9"
        .Cells(26, 8) = "53"
        .Cells(29, 8) = "1"
        .Cells(30, 8) = "N816"
        .Cells(31, 8) = "N15"
        .Cells(32, 8) = "N16"
        .Cells(33, 8) = "N17"
        .Cells(34, 8) = "N18"
        .Cells(35, 8) = "N19"
        .Cells(36, 8) = "N20"
        .Cells(37, 8) = "N21"
        .Cells(38, 8) = "N22"
        .Cells(39, 8) = "N23"
        .Cells(40, 8) = "N24"
        .Cells(41, 8) = "N817"
        
        .Cells(15, 9) = "2"
        .Cells(18, 9) = "17"
        .Cells(19, 9) = "42"
        .Cells(20, 9) = "47"
        .Cells(21, 9) = "30"
        .Cells(22, 9) = "5"
        .Cells(23, 9) = "29"
        .Cells(24, 9) = "48"
        .Cells(25, 9) = "15"
        .Cells(29, 9) = "2"
        .Cells(30, 9) = "N812"
        .Cells(31, 9) = "N813"
        .Cells(32, 9) = "N7"
        .Cells(33, 9) = "N8"
        .Cells(34, 9) = "N9"
        .Cells(35, 9) = "N10"
        .Cells(36, 9) = "N11"
        .Cells(37, 9) = "N12"
        .Cells(38, 9) = "N13"
        .Cells(39, 9) = "N14"
        .Cells(40, 9) = "N814"
        .Cells(41, 9) = "N815"
        
        .Cells(15, 10) = "3"
        .Cells(19, 10) = "44"
        .Cells(20, 10) = "22"
        .Cells(21, 10) = "45"
        .Cells(22, 10) = "16"
        .Cells(23, 10) = "46"
        .Cells(24, 10) = "23"
        .Cells(29, 10) = "3"
        .Cells(31, 10) = "N808"
        .Cells(32, 10) = "N809"
        .Cells(33, 10) = "N1"
        .Cells(34, 10) = "N2"
        .Cells(35, 10) = "N3"
        .Cells(36, 10) = "N4"
        .Cells(37, 10) = "N5"
        .Cells(38, 10) = "N6"
        .Cells(39, 10) = "N810"
        .Cells(40, 10) = "N811"
        
        .Cells(15, 11) = "4"
        .Cells(29, 11) = "4"
        .Cells(32, 11) = "N800"
        .Cells(33, 11) = "N801"
        .Cells(34, 11) = "N802"
        .Cells(35, 11) = "N803"
        .Cells(36, 11) = "N804"
        .Cells(37, 11) = "N805"
        .Cells(38, 11) = "N806"
        .Cells(39, 11) = "N807"
                
        .Cells(43, 1) = "ELAB"
        .Cells(44, 1) = "4"
        .Cells(45, 1) = "3"
        .Cells(46, 1) = "2"
        .Cells(47, 1) = "1"
        .Cells(48, 1) = "0"
        .Cells(49, 1) = "-1"
        .Cells(50, 1) = "-2"
        .Cells(51, 1) = "-3"
        
        .Cells(43, 2) = "-5"
        
        
        .Cells(46, 2) = "54"
        .Cells(47, 2) = "35"
        .Cells(48, 2) = "34"
        .Cells(49, 2) = "15"
        .Cells(43, 3) = "-4"
        
        .Cells(45, 3) = "55"
        .Cells(46, 3) = "53"
        .Cells(47, 3) = "36"
        .Cells(48, 3) = "33"
        .Cells(49, 3) = "16"
        .Cells(50, 3) = "14"
        .Cells(43, 4) = "-3"
        
        .Cells(45, 4) = "56"
        .Cells(46, 4) = "52"
        .Cells(47, 4) = "37"
        .Cells(48, 4) = "32"
        .Cells(49, 4) = "17"
        .Cells(50, 4) = "13"
        .Cells(51, 4) = "1"
        .Cells(43, 5) = "-2"
        .Cells(44, 5) = "66"
        .Cells(45, 5) = "57"
        .Cells(46, 5) = "51"
        .Cells(47, 5) = "38"
        .Cells(48, 5) = "31"
        .Cells(49, 5) = "18"
        .Cells(50, 5) = "12"
        .Cells(51, 5) = "2"
        .Cells(43, 6) = "-1"
        .Cells(44, 6) = "65"
        .Cells(45, 6) = "58"
        .Cells(46, 6) = "50"
        .Cells(47, 6) = "39"
        .Cells(48, 6) = "30"
        .Cells(49, 6) = "19"
        .Cells(50, 6) = "11"
        .Cells(51, 6) = "3"
        .Cells(43, 7) = "0"
        .Cells(44, 7) = "64"
        .Cells(45, 7) = "59"
        .Cells(46, 7) = "49"
        .Cells(47, 7) = "40"
        .Cells(48, 7) = "29"
        .Cells(49, 7) = "20"
        .Cells(50, 7) = "10"
        .Cells(51, 7) = "4"
        .Cells(43, 8) = "1"
        .Cells(44, 8) = "63"
        .Cells(45, 8) = "60"
        .Cells(46, 8) = "48"
        .Cells(47, 8) = "41"
        .Cells(48, 8) = "28"
        .Cells(49, 8) = "21"
        .Cells(50, 8) = "9"
        .Cells(51, 8) = "5"
        .Cells(43, 9) = "2"
        
        .Cells(45, 9) = "61"
        .Cells(46, 9) = "47"
        .Cells(47, 9) = "42"
        .Cells(48, 9) = "27"
        .Cells(49, 9) = "22"
        .Cells(50, 9) = "8"
        .Cells(51, 9) = "6"
        .Cells(43, 10) = "3"
        
        .Cells(45, 10) = "62"
        .Cells(46, 10) = "46"
        .Cells(47, 10) = "43"
        .Cells(48, 10) = "26"
        .Cells(49, 10) = "23"
        .Cells(50, 10) = "7"
        .Cells(43, 11) = "4"
        
        
        .Cells(46, 11) = "45"
        .Cells(47, 11) = "44"
        .Cells(48, 11) = "25"
        .Cells(49, 11) = "24"
               
    End With

End Sub


Public Sub RowsToTable()

    Dim srcSheet As Worksheet
    Dim nowSheet As Worksheet
    Dim sel
    Dim i As Long, j As Long
    Dim tmpStr As String
    Dim nowRow As Integer
    Dim nowCol As Integer
    
    Set srcSheet = ActiveSheet
    If srcSheet.Cells(1, 1).CurrentRegion.Columns.Count > 1 Then Exit Sub
    
    On Error GoTo 0
    
    colNums = InputBox("Please input columns count: ", "Table Setting", 9)
    
    GoSub trans
    
Exit Sub
trans:
    
    Set nowSheet = AddSheet(srcSheet.Name & "_t", , srcSheet.Name)
    
    nowRow = 1: nowCol = 1
    
    For i = 1 To srcSheet.UsedRange.Rows.Count
        If srcSheet.Cells(i, 1) = "" Then
        Else
            If nowCol > colNums Then
                nowRow = nowRow + 1
                nowCol = 1
            End If
            nowSheet.Cells(nowRow, nowCol).Value = srcSheet.Cells(i, 1).Value
            nowCol = nowCol + 1
        End If
    Next i
    
    Dim RowStep As Integer
    
    For i = 2 To nowSheet.Cells.Rows.Count
        If nowSheet.Cells(i, 1) = srcSheet.Cells(1, 1) Then
            RowStep = i - 1
            Exit For
        End If
    Next i

    For i = nowSheet.UsedRange.Rows.Count + 1 To 2 Step -RowStep
        nowSheet.Rows(i).Select
        nowSheet.Rows(i).Delete
    Next i
    
Return

End Sub

Public Function isGroupingSafe()
    
    Dim nowSheet As Worksheet
    Dim dSheet As Worksheet
    Set nowSheet = Worksheets("Grouping")
    Set dSheet = Worksheets("Data")
    Dim waferListStr As String
    Dim waferList
    Dim rangeNames As String
    Dim rangeList
    
    Dim errMsg As String
    
    isGroupingSafe = True
    
    nowSheet.Cells.Interior.Pattern = xlNone
    
    Dim i As Integer

    For i = 2 To nowSheet.Cells(1, 1).CurrentRegion.Rows.Count
        waferListStr = waferListStr & "," & nowSheet.Cells(i, 1).Value
    Next i
    
    waferListStr = Mid(waferListStr, 2)
    waferList = Split(waferListStr, ",")

    ' Check wafer Id
    For i = 1 To dSheet.Names.Count
        rangeNames = rangeNames & "," & getCOL(dSheet.Names(i).Name, "wafer_", 2)
    Next i
    rangeNames = Mid(rangeNames, 2)
    rangeList = Split(rangeNames, ",")
    
    For i = LBound(waferList) To UBound(waferList)
        If Not IsNumeric(Application.Match(waferList(i), rangeList, 0)) Then
            errMsg = errMsg & "," & waferList(i)
        End If
    Next i
    If Not errMsg = "" Then errMsg = "wafer_" & Mid(errMsg, 2) & " 不存在!" & Chr(10)
    
    'Check Group Name
    Dim SignList()
    Dim Sign
    Dim flag As Boolean
    flag = False
    SignList = Array("`", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "=", "{", "}", "|", ";", ":", ",", ".", "<", ">", "\", "/", "?", "[", "]", "'", " ")
    For i = 2 To nowSheet.Cells(1, 1).CurrentRegion.Rows.Count
        For Each Sign In SignList
            If InStr(nowSheet.Cells(i, 2).Value, Sign) Then
                nowSheet.Cells(i, 2).Interior.Color = RGB(255, 0, 0)
                flag = True
            End If
        Next Sign
        For Each Sign In waferList
            If CStr(nowSheet.Cells(i, 2).Value) = Sign Then
                nowSheet.Cells(i, 2).Interior.Color = RGB(255, 0, 0)
                flag = True
            End If
        Next Sign
    Next i
    
    If flag Then errMsg = errMsg & Chr(10) & "此名稱的語法不正確。" & Chr(10) & _
                                             "請確認 GroupName: " & Chr(10) & _
                                             "        -以字母或底線 (_) 開頭" & Chr(10) & _
                                             "        -不包含空格或其他無效字元 (除了加號 (+) 字元)" & Chr(10) & _
                                             "        -不與 waferId 的命名衝突。"

    
    If Not errMsg = "" Then
        isGroupingSafe = False
        Err.Raise 1004, Err.Source, errMsg
    End If
    
    
End Function

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

