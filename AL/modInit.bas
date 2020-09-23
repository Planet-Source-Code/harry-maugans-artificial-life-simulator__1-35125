Attribute VB_Name = "modInit"

Sub Main()
    frmLoading.Show
    frmLoading.AddStatus "Loading cellular enviroment"
    DoEvents
    Load frmMain
    frmMain.Show
    frmLoading.AddStatus "Cellular observing interface"
    DoEvents
    Load frmStatus
    frmStatus.Top = frmMain.Top + frmMain.Height - frmStatus.Height
    frmStatus.Left = frmMain.Left + frmMain.Width
    frmStatus.Show
    frmLoading.AddStatus "Initalizing cellular grid"
    DoEvents
    modGrid.InitGrid
    frmLoading.AddStatus "Setting the cellular size ratio"
    SetCellSize
    frmLoading.AddStatus "Initalizing the cell attributes"
    CellHealth = 64
    frmLoading.AddStatus "Creating cellular molds"
    MakeCells
    frmLoading.AddStatus "Giving birth to Adam and Eve"
    StartAE
    
    Unload frmLoading
End Sub

Public Function StartAE()
    ReDim PltCells(0 To (NUMWIDTH * NUMHEIGHT - 1)) As New clsCell
    ReDim AnmCells(0 To (NUMWIDTH * NUMHEIGHT - 1)) As New clsCell
    ReDim PltAlive(0 To (NUMWIDTH * NUMHEIGHT - 1)) As Boolean
    ReDim AnmAlive(0 To (NUMWIDTH * NUMHEIGHT - 1)) As Boolean
    ReDim ByCoords(0 To NUMWIDTH - 1, 0 To NUMHEIGHT - 1) As Boolean
    
    AnmAlive(0) = True
    AnmCells(0).Birth True, (NUMWIDTH / 2) - 1, NUMHEIGHT / 2, NextAnm, NextCell
    PltAlive(0) = True
    PltCells(0).Birth False, NUMWIDTH / 2, NUMHEIGHT / 2, NextPlt, NextCell
End Function

Public Function SetCellSize()
    frmMain.shpCell(0).Height = modGrid.GridHeight
    frmMain.shpCell(0).Width = modGrid.GridWidth
End Function

Private Sub MakeCells()
    Dim i As Long
    
    frmLoading.PB1.Value = 0
    frmLoading.PB1.Max = (modGrid.NUMHEIGHT * modGrid.NUMWIDTH) / 100
    For i = 1 To modGrid.NUMHEIGHT * modGrid.NUMWIDTH
        Load frmMain.shpCell(i)
        frmMain.shpCell(i).Visible = False
        frmLoading.PB1.Value = i / 100
        DoEvents
    Next i
End Sub
