Public IsPinned As Boolean

Option Explicit

Private Sub CMB_Check_Click()
    Me.Hide
    Call CheckFile
End Sub

Private Sub CMB_ExpFile_Click()
    Me.Hide
    Call ExportFile
End Sub

Private Sub CMB_genPPT_Click()
    Me.Hide
    Call GenPPT
End Sub

Private Sub CMB_GenSingleChart_Click()
    Me.Hide
    Call genSingleChart
End Sub

Private Sub CMB_GetCoor_Click()
    If Not IsExistSheet("Data") Then MsgBox "Please load longfile before the operation.": Me.Hide: Exit Sub
    Me.Hide
    FrmWaferMapSetting.Show
End Sub

Private Sub CMB_GenWaferMap_Click()
    Me.Hide
    FrmWaferMap.Show
End Sub

Private Sub CMB_Hint_Click()
    Me.Hide
    FrmHint.Show
End Sub

Private Sub CMB_GenMismatch_Click()
    Me.Hide
    Call RunMismatch
End Sub

Private Sub CMB_ImportUEDA_Click()
    Me.Hide
    Call UEDA

End Sub

Private Sub CMB_LoadLongFile_Click()
    Me.Hide
    Call LoadlongFile
End Sub

Private Sub CMB_OBC_Click()
    Me.Hide
    FrmOBC.Show
End Sub


Private Sub CMB_PlotMMChart_Click()
    Me.Hide
    FrmOption.Show
    Call PlotChart
End Sub

Private Sub CMB_R2T_Click()
    Me.Hide
    Call RowsToTable
End Sub

Private Sub CMB_ReCntCorner_Click()
    Me.Hide
    Call reCountCorner
End Sub

Private Sub CMB_Redraw_Click()
    Me.Hide
    Call GenChartSummary
End Sub

Private Sub CMB_SpecFile_Click()
    Me.Hide
    Call Load_SPECFile
End Sub

Private Sub CMB_PinScatter_Click()
    Me.Hide
    If IsPinned Then
        IsPinned = False
        CMB_PinScatter.Caption = "Pin Scatter"
        Call UnpinScatter
    Else
        IsPinned = True
        CMB_PinScatter.Caption = "Unpin Scatter"
        Call PinScatter
    End If
End Sub

Private Sub UserForm_Initialize()

    If IsExistSheet("!SCATTER1") Or IsExistSheet("!BOXTREND1") Or IsExistSheet("!CUMULATIVE1") Then
        IsPinned = True
        CMB_PinScatter.Caption = "Unpin Scatter"
    Else
        IsPinned = False
        CMB_PinScatter.Caption = "Pin Scatter"
    End If

End Sub
