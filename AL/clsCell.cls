VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''
' User Defines             '
''''''''''''''''''''''''''''
Private Type Kingdoms
    Animalia As Boolean
    Plantae As Boolean
End Type
Private Type GridPlace
    X As Integer
    Y As Integer
    AbsX As Long
    AbsY As Long
End Type

''''''''''''''''''''''''''''
' Public Attributes        '
''''''''''''''''''''''''''''
'Public Alive As Boolean


''''''''''''''''''''''''''''
' Cell Attributes          '
''''''''''''''''''''''''''''
Private Type xThis
    Health As Byte
    Age As Byte
    Kingdom As Kingdoms
    Placement As GridPlace
    CellNum As Long
    SpeciesNum As Long
End Type

Dim This As xThis


Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    '
End Sub

Public Sub Birth(Species As Boolean, X As Integer, Y As Integer, SpeciesNumber As Long, CellNumber As Long)  'species, true is anm, false is plt
    This.Health = CellHealth
    This.Age = 0
    This.CellNum = CellNumber
    This.SpeciesNum = SpeciesNumber
    This.Kingdom.Animalia = Species
    This.Kingdom.Plantae = Not Species
    If This.Kingdom.Animalia = True Then
        TotAnmBirths = TotAnmBirths + 1
    Else
        TotPltBirths = TotPltBirths + 1
    End If
    TotAnmCells = TotAnmCells + Abs(CInt(This.Kingdom.Animalia))
    TotPltCells = TotPltCells + Abs(CInt(This.Kingdom.Plantae))
    This.Placement.X = X
    This.Placement.Y = Y
    This.Placement.AbsX = X * modGrid.GridWidth
    This.Placement.AbsY = Y * modGrid.GridHeight
    ByCoords(X, Y) = True
    ShowMe
    UpdateVars
End Sub

Private Sub ShowMe()
    frmMain.shpCell(This.CellNum).Left = This.Placement.AbsX
    frmMain.shpCell(This.CellNum).Top = This.Placement.AbsY
    
    If This.Kingdom.Animalia = True Then
        frmMain.shpCell(This.CellNum).FillColor = vbRed
    Else
        frmMain.shpCell(This.CellNum).FillColor = vbGreen
    End If
    
    frmMain.shpCell(This.CellNum).Visible = True
End Sub

Public Sub Live()
    This.Health = This.Health / 2
    If This.Health <= 1 Then
        Die
    End If
    
    This.Age = This.Age + 1
    If This.Age Mod 2 = 0 Then
        Reproduce
    End If
End Sub

Private Sub Reproduce()
    Dim TempNum As Long
    Dim SData
    
    SData = Split(GetClosestCoords, ",")
    If This.Kingdom.Animalia = True Then
        TempNum = NextAnm
        AnmAlive(TempNum) = True
        AnmCells(TempNum).Birth True, CInt(SData(0)), CInt(SData(1)), TempNum, NextCell
    Else
        TempNum = NextPlt
        PltAlive(TempNum) = True
        PltCells(TempNum).Birth False, CInt(SData(0)), CInt(SData(1)), TempNum, NextCell
    End If
    DoEvents
End Sub

Private Sub Die()
    frmMain.shpCell(This.CellNum).Visible = False
    If This.Kingdom.Animalia = True Then
        AnmAlive(This.SpeciesNum) = False
        TotAnmDeaths = TotAnmDeaths + 1
        TotAnmCells = TotAnmCells - 1
    Else
        PltAlive(This.SpeciesNum) = False
        TotPltDeaths = TotPltDeaths + 1
        TotPltCells = TotPltCells - 1
    End If
    DoEvents
End Sub

Private Function GetClosestCoords() As String
    Dim X1 As Long, Y1 As Long
    Dim CX As Long, CY As Long 'closest x & y
    Dim CNX As Long, CNY As Long 'closest algorithm output number
    
    CX = 51
    CY = 51
    For Y1 = 0 To modGrid.NUMHEIGHT - 1
        For X1 = 0 To modGrid.NUMWIDTH - 1
            If ByCoords(X1, Y1) = False Then
                If Abs(X1 - This.Placement.X) < CX Then
                    If Abs(Y1 - This.Placement.Y) < CY Then
                        CX = Abs(X1 - This.Placement.X)
                        CNX = X1
                        CY = Abs(Y1 - This.Placement.Y)
                        CNY = Y1
                    End If
                End If
            End If
        Next X1
    Next Y1
    
    GetClosestCoords = CNX & "," & CNY
End Function
