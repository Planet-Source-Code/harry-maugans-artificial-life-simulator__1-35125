Attribute VB_Name = "modGrid"
Public Const NUMWIDTH = 100
Public Const NUMHEIGHT = 100

Public GridWidth As Integer
Public GridHeight As Integer



Public Sub InitGrid()
    Dim X, Y As Long
    
    GridWidth = frmMain.ScaleWidth / NUMWIDTH
    GridHeight = frmMain.ScaleHeight / NUMHEIGHT
    
    frmLoading.PB1.Max = frmMain.ScaleHeight / 100
    frmMain.ForeColor = &HC0C0C0
    For Y = 0 To frmMain.ScaleHeight
        For X = 0 To frmMain.ScaleWidth
            If X Mod GridWidth = 0 Then
                frmMain.PSet (X, Y)
            ElseIf Y Mod GridHeight = 0 Then
                frmMain.PSet (X, Y)
            End If

        Next X
        frmLoading.PB1.Value = Y / 100
        DoEvents
    Next Y
End Sub
