Option Explicit

'========================================
' IDAS(Integrated Data Analysis System)
'========================================
'
'TD3/SD/DT4 - Rain Hu
'

Public Const SSheet = "SPEC"
Public Const dSheet = "Data"
Public Const pi = 3.14159265358979

Public WaferArray() As String


Public Sub Load_LongFile()

    Dim Filename
    Filename = Application.GetOpenFilename("RPT File(*.RPT),*.RPT", 1, "Open File", "Open", True)
    If Not IsArray(Filename) Then Exit Sub
    If UBound(Filename) = 1 Then
        Call Load_SingleLongFile(CStr(Filename(1)))
    Else
        Call Load_MultiLongFiles(Filename)
    End If
    Worksheets("Data").Activate
    
End Sub


Public Sub InitStep()
    Dim nowSheet As Worksheet
    Set nowSheet = ActiveSheet
    If IsExistSheet(dSheet) Then
        Call TrimCol(dSheet, 2)                     'Trim Space for All Cells in "Parameter" Column
        Call SetWaferRange
    End If
    If IsExistSheet(SSheet) Then
        Call TrimCol(SSheet, 3)                     'Trim Space for All Celss in "Item" Column
        Call TrimSheet(SSheet)                      'Trim Empty Cells for Whole Sheet
        Call GenSPECTEMP                            'Generate a SEPCTEMP for loading
    End If
    If IsExistSheet("SPEC_List") Then
        Call TrimSpecList                           'Replace "\, /, ?, [, ]" as "_" to create worksheet
    End If
    nowSheet.Activate
    Set nowSheet = Nothing
End Sub

Public Sub Select_Wafer()
    If IsExistSheet("Data") Then FrmSelectWafer.Show
End Sub

Public Sub AutoRun()

    Call SummaryStep
    Call GenCharts

End Sub

Sub SummaryStep()
    
    If Not IsExistSheet("Data") Then MsgBox "Please load longfile before the operation.": Exit Sub
    If Not IsExistSheet("SPEC") Then MsgBox "Please check SEPC sheet before the operation!!": Exit Sub
    If Not IsExistSheet("SPEC_List") Then MsgBox "Please set SPEC_List before the operation!!": Exit Sub
    
    If IsExistSheet("Grouping") Then
        If Not isGroupingSafe Then Exit Sub
    End If
    
    Call Speed
    Call LotRawData
    Call LotSummary
    Call Unspeed
End Sub
Sub GenCharts()
   
    Dim waferList() As String
    Dim siteNum As Integer
   
    If Not IsExistSheet("ChartType") Then MsgBox "Please check ChartType sheet before the operation.": Exit Sub
    If Not IsExistSheet("PlotSetup") Then MsgBox "Please check PlotSetup sheet before the operation!!": Exit Sub
    If IsExistSheet("Grouping") Then
        If Not isGroupingSafe Then Exit Sub
    End If
    
    Call Speed
    
    Call GenChartHeader
    
    Call GetWaferList(dSheet, waferList)
    siteNum = getSiteNum(dSheet)
    
    Call GenScatter(waferList, siteNum)
    Call GenBoxTrend(waferList, siteNum)
    Call GenCumulative(waferList, siteNum)

    Call DioPlotAllChart

    Call new_FitChart
    Call CornerCount
    Call RawdataRange
    Call GenChartSummary
    
    Call Unspeed
   
End Sub

Public Sub GenPPT()

    Dim i As Long, j As Long, k As Long
    Dim nowSheet As Worksheet
    Dim SourceSheet As Worksheet
    Dim mPPT As New PowerPoint.Application
    Dim nowPPT As PowerPoint.Presentation
    Dim nowSlide As PowerPoint.Slide
    Dim nowShape As PowerPoint.Shape
    Dim x As Design
    Dim CopyPage As Long, mCount As Long
    Dim nChart As New Collection
    Dim LType As String
    Dim mFile
    '******************set parameter******************
    Const pTitle As Integer = 1     'row 1
    Const pBlank As Integer = 12    'row 12
    Const pOrder As Integer = 5     'col 5
    Const ppTitle As Integer = 3    'col 3
    Const pContent As Integer = 6   'col 6
    Const pType As Integer = 4      'col 4
    '******************pre-check******************
    If Not IsExistSheet("PPT") Then
        MsgBox "Cannot access worksheet ""PPT"""
        Exit Sub
    End If
    '*********************************************
    Set nowSheet = Worksheets("PPT")
    CopyPage = nowSheet.Cells(9, 1).Value
'    mPPT.Visible = True
    '******************pre-check2******************
    On Error GoTo ErrHandler
    Set nowPPT = mPPT.Presentations.Open(Application.ThisWorkbook.Path & "\" & "PPT File.pptx")
    mPPT.Visible = True
    On Error GoTo 0
    '******************set ppLayoutTitle******************
    With nowPPT.Slides(1)
        For i = 1 To nowSheet.Cells(pTitle, 1).CurrentRegion.Columns.Count
            If nowSheet.Cells(pTitle, i) = "Date" Then
                .Shapes(i).TextFrame.TextRange.Text = Date
            ElseIf Not nowSheet.Cells(pTitle + 1, i) = "" Then
                .Shapes(i).TextFrame.TextRange.Text = nowSheet.Cells(pTitle + 1, i)
            End If
        Next i
    End With
    On Error Resume Next
    '******************by case******************
    For j = 1 To nowSheet.Cells(pBlank, 1).CurrentRegion.Rows.Count - 1
        DoEvents
        If Not UCase(nowSheet.Cells(j + pBlank, 1)) = "" Then
            LType = nowSheet.Cells(j + pBlank, pType).Text
            Set SourceSheet = Worksheets(nowSheet.Cells(j + pBlank, ppTitle - 1).Text)
            '******************CopyPage******************
            nowPPT.Slides(CopyPage).Copy
            Set x = nowPPT.Slides(CopyPage).Design
            nowPPT.Slides.Paste.Design = x
            Set nowSlide = nowPPT.Slides(CopyPage)
            '******************set text******************
            With nowPPT.Slides(nowPPT.Slides.Count)
                If nowSheet.Cells(j + pBlank, ppTitle) = "" Then
                    .Shapes(1).TextFrame.TextRange.Text = nowSheet.Cells(j + pBlank, ppTitle - 1)
                Else
                    .Shapes(1).TextFrame.TextRange.Text = nowSheet.Cells(j + pBlank, ppTitle)
                End If
                .Shapes(2).TextFrame.TextRange.Text = nowSheet.Cells(j + pBlank, pContent)
                '******************paste chart******************
                Dim W As Integer
                Dim H As Integer
                Dim cNum As Integer

                cNum = nowSlide.Shapes.Count
                
                For i = 1 To SourceSheet.ChartObjects.Count         'object
                    For k = 0 To 8                                  'site
                        If SourceSheet.ChartObjects(i).Chart.ChartTitle.Caption = nowSheet.Cells(j + pBlank, pContent + k + 1).Value Then
                            SourceSheet.ChartObjects(i).CopyPicture
                            .Shapes.PasteSpecial
                            cNum = cNum + 1
                            '******************Set position******************
                            'W = .Shapes(cNum).Parent.Master.width
                            'H = .Shapes(cNum).Parent.Master.Height
                            W = .Master.width
                            H = .Master.Height
                                                    
                            Select Case LType
                                Case "A" '1
                                   .Shapes(cNum).width = W * 0.75
                                   .Shapes(cNum).Top = H * 0.23
                                   .Shapes(cNum).Left = (W - .Shapes(cNum).width) / 2
                                     If k = 1 Then Exit For
                                Case "B" '2
                                   .Shapes(cNum).width = W * 0.5
                                   .Shapes(cNum).Top = H * 0.46
                                   .Shapes(cNum).Left = W / 2 * k
                                     If k = 2 Then Exit For
                                Case "C" '4
                                   .Shapes(cNum).Height = H * 0.373
                                   .Shapes(cNum).Top = 0.177 * H + 0.373 * H * Int(k / 2)
                                   .Shapes(cNum).Left = W / 2 - .Shapes(cNum).width * ((k + 1) Mod 2)
                                     If k = 4 Then Exit For
                                Case "D" '6
                                   .Shapes(cNum).width = W * 0.333
                                   .Shapes(cNum).Top = 0.3 * H + .Shapes(cNum).Height * Int(k / 3)
                                   .Shapes(cNum).Left = .Shapes(cNum).width * (k Mod 3)
                                     If k = 6 Then Exit For
                                Case "E" '9
                                   .Shapes(cNum).width = W * 0.333
                                   .Shapes(cNum).Top = H - .Shapes(cNum).Height * Int((11 - k) / 3)
                                   .Shapes(cNum).Left = .Shapes(cNum).width * (k Mod 3)
                             End Select
                        End If
                        'Exit For
                    Next k
                Next i
            End With
        End If
    Next j
    nowPPT.Slides(CopyPage).Delete
    
    '******************save file & finish******************
    On Error GoTo 0
    If nowSheet.Cells(4, 2).Value <> "" Then nowPPT.SaveAs (Application.ThisWorkbook.Path & "\" & nowSheet.Cells(4, 2))
    
    Set nowSheet = Nothing
    Set SourceSheet = Nothing
    Set nowSlide = Nothing
    Set nowPPT = Nothing
    Set nowShape = Nothing
    Set mPPT = Nothing
    Exit Sub
    '******************************************************
ErrHandler:
    MsgBox "Cannot access ""PPT file.pptx. Please select PPT sample file manually."""
    mFile = Application.GetOpenFilename("pptx File, *.pptx", 1, "Load PPT sample file", "Open", False)
    If mFile = False Then Exit Sub
    Set nowPPT = mPPT.Presentations.Open(mFile)
    Resume Next

End Sub

Public Function AutoMacro(specFile As String, rawFile As String, Optional autoMode As String = "")
    Dim tmpStr As String
    Dim nowLongFile As String
    Dim SubName As String
      
    ThisWorkbook.UpdateRemoteReferences = False
      
    ExMode = autoMode
   
    Call LoadSpecFile(specFile)
    Call TrimCol(SSheet, 3)
    Call TrimSpecList
    Call GenSPECTEMP
   
    If getFileLine(Filename) > 1048576 Then
        AutoMacro = "Error:999:File line over 1048576!!!"
        Exit Function
    End If
   
    SubName = UCase(Mid(rawFile, InStrRev(rawFile, ".")))
    If SubName = ".LONG" Then
        Call LoadLongFileLONG(rawFile)
    Else
        Call LoadLongFileRPT(rawFile)
    End If
   
    If Worksheets("Data").Cells(16, 1) = "" Then AutoMacro = "Error:998:No Data !!!": Exit Function
    If Worksheets("Data").UsedRange.Rows.Count <= 1048576 Then
        Call InitStep
        Call SummaryStep
        Call GenCharts
    End If
    AutoMacro = getReceiver()
    If IsExistSheet("RECEIVER") Then Worksheets("RECEIVER").Activate
   
End Function


Sub GenWaferMap()
    FrmWaferMap.Show
End Sub

Sub GenWaferMapManual()
    FrmWaferMapManual.Show
End Sub

Sub Modeling_Boxtrend()
    FrmModeling.Show
End Sub

Sub FrmManualFunction()
    FrmManual.Show
End Sub

