Attribute VB_Name = "modVars"
Public Days As Long
Public SimHours As Integer
Public SimMinutes As Byte
Public SimSeconds As Byte

Public TotPltCells As Long
Public TotAnmCells As Long
Public TotPltBirths As Long
Public TotPltDeaths As Long
Public TotAnmBirths As Long
Public TotAnmDeaths As Long

Public PltCells() As New clsCell
Public AnmCells() As New clsCell
Public PltAlive() As Boolean
Public AnmAlive() As Boolean
Public ByCoords() As Boolean

Public CellHealth As Byte

Public Sub UpdateVars()
    frmStatus.txtAnmBirths.Text = TotAnmBirths
    frmStatus.txtAnmDeaths.Text = TotAnmDeaths
    frmStatus.txtAnmLiv.Text = TotAnmCells
    frmStatus.txtDays.Text = Days
    frmStatus.txtPltBirths.Text = TotPltBirths
    frmStatus.txtPltDeaths.Text = TotPltDeaths
    frmStatus.txtPltLiv.Text = TotPltCells
    frmStatus.txtSimTime.Text = SimHours & ":" & Format(SimMinutes, "00") & ":" & Format(SimSeconds, "00")
    frmStatus.txtTotBirths.Text = TotAnmBirths + TotPltBirths
    frmStatus.txtTotDeaths.Text = TotAnmDeaths + TotPltDeaths
    frmStatus.txtTotLivCells.Text = TotAnmCells + TotPltCells
    frmStatus.Refresh
End Sub

Public Function NextAnm() As Long
    Dim i As Long
    
    For i = 0 To UBound(AnmAlive)
        If AnmAlive(i) = False Then
            NextAnm = i
            Exit Function
        End If
    Next i
End Function

Public Function NextPlt() As Long
    Dim i As Long
    
    For i = 0 To UBound(PltAlive)
        If PltAlive(i) = False Then
            NextPlt = i
            Exit Function
        End If
    Next i
End Function

Public Function NextCell() As Long
    Dim i As Long
    
    For i = 0 To modGrid.NUMHEIGHT * modGrid.NUMWIDTH
        If frmMain.shpCell(i).Visible = False Then
            NextCell = i
            Exit Function
        End If
    Next i
End Function

