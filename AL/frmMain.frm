VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALS"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10500
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   700
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   Begin VB.Shape shpCell 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_GotFocus()
    If frmLoading.Visible = True Then
        frmLoading.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Me.Top = Screen.Height / 2 - (Me.Height / 2)
    Me.Left = Screen.Width / 2 - (Me.Width / 2)
    Me.Left = Me.Left - (Me.Left / 2)
    Me.Top = Me.Top - (Me.Top / 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmStatus
    End
End Sub
