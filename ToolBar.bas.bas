Option Explicit

Private Sub Workbook_Open()

    Dim CB As CommandBar
    Dim CBB As CommandBarButton
    Dim i As Integer
    
    Application.ScreenUpdating = False

'Add Bar

    Set CB = Nothing
    
    On Error Resume Next
        Application.CommandBars("IDAS").Delete
        Application.CommandBars("Autoreport").Delete
        Application.CommandBars("menu").Delete
        Application.CommandBars("DRCS").Delete
    On Error GoTo 0
    
    Set CB = Application.CommandBars.Add(Name:="IDAS", Temporary:=True)
    With CB
        .Visible = True
        .Position = msoBarTop
    End With
   
'Add Button
   
    Set CBB = CB.Controls.Add(Type:=msoControlButton)
    With CBB
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "Load Data"
        .FaceId = 23
        .OnAction = "Load_LongFile"
    End With
    
    Set CBB = CB.Controls.Add(Type:=msoControlButton)
    With CBB
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "Initial"
        .FaceId = 602
        .OnAction = "initStep"
    End With
    
    Set CBB = CB.Controls.Add(Type:=msoControlButton)
    With CBB
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "Select Wafer"
        .FaceId = 98
        .OnAction = "Select_Wafer"
    End With
    
    Set CBB = CB.Controls.Add(Type:=msoControlButton)
    With CBB
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "Summary Table"
        .FaceId = 107
        .OnAction = "SummaryStep"
    End With
    
    Set CBB = CB.Controls.Add(Type:=msoControlButton)
    With CBB
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "Generate Charts"
        .FaceId = 430
        .OnAction = "GenCharts"
    End With
    
    Set CBB = CB.Controls.Add(Type:=msoControlButton)
    With CBB
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "Manual Functions"
        .FaceId = 278
        .OnAction = "FrmManualFunction"
    End With

    Application.ScreenUpdating = True

End Sub

Private Sub Workbook_Deactivate()

On Error Resume Next
    
    Application.CommandBars("IDAS").Visible = False

On Error GoTo 0

End Sub

Private Sub Workbook_activate()

On Error Resume Next
    
    Application.CommandBars("IDAS").Visible = True

On Error GoTo 0

End Sub




