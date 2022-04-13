Private Sub CmdRun_Click()
    Call getCoord(TextBoxWaferMap.Value)
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    TextBoxWaferMap.Value = "Default"
End Sub
