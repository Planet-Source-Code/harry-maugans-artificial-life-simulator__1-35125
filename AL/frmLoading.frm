VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmLoading 
   BorderStyle     =   0  'None
   Caption         =   "loading"
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBlah 
      Height          =   285
      Left            =   44400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   150
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   1590
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtLoading 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Loading... Please wait"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   120
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    AddStatus "Initalizing Program..."
End Sub

Public Sub AddStatus(Text As String)
    If Me.Visible = True Then
        txtLoading.Text = txtLoading.Text & vbCrLf & Text
        frmLoading.SetFocus
        txtBlah.SetFocus
    Else
        txtLoading.Text = txtLoading.Text & Text
    End If
    txtLoading.SelStart = Len(txtLoading.Text)
    txtLoading.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    For i = 0 To 5000
        DoEvents
    Next i
End Sub

Private Sub txtLoading_GotFocus()
    txtBlah.SetFocus
End Sub
