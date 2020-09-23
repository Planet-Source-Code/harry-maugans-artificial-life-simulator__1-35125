VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   4575
   Begin VB.Frame Frame2 
      Caption         =   "Cells"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   4335
      Begin VB.TextBox txtAnmDeaths 
         Height          =   285
         Left            =   2400
         TabIndex        =   23
         Text            =   "0"
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox txtPltDeaths 
         Height          =   285
         Left            =   2400
         TabIndex        =   21
         Text            =   "0"
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtAnmBirths 
         Height          =   285
         Left            =   2400
         TabIndex        =   19
         Text            =   "0"
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtPltBirths 
         Height          =   285
         Left            =   2400
         TabIndex        =   17
         Text            =   "0"
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtTotDeaths 
         Height          =   285
         Left            =   2400
         TabIndex        =   15
         Text            =   "0"
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtTotBirths 
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Text            =   "0"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtAnmLiv 
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Text            =   "0"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtPltLiv 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Text            =   "0"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtTotLivCells 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Text            =   "0"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Total Animal cell deaths:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   4200
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "Total Plant cell deaths:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Total Animal cell births:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Total Plant cell births:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Total number of deaths:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Total number of births:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Number of living Animal Cells:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Number of living Plant Cells:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Total number of living Cells:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdAIDS 
         Caption         =   "Inflict sterility upon any new offspring"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox txtLog 
         Height          =   285
         Left            =   2280
         TabIndex        =   28
         Text            =   "textsim.txt"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Timer tmrLife 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   720
         Top             =   3000
      End
      Begin VB.Timer tmrSim 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   3000
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause Simulation"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CommandButton cmdBegin 
         Caption         =   "Begin Simulation"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtSimTime 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "00:00:00"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtDays 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0"
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Logging filename (log/):"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Artificial Life Simulation coded by: Harry Maugans"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2520
         Width           =   4095
      End
      Begin VB.Label Label11 
         Caption         =   "Simulation running time:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Number of days elapsed:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAIDS_Click()
    CellHealth = 0
    cmdAIDS.Enabled = False
End Sub

Private Sub cmdBegin_Click()
    cmdBegin.Enabled = False
    cmdPause.Enabled = True
    cmdPause.SetFocus
    
    tmrSim.Enabled = True
    tmrLife.Enabled = True
End Sub

Private Sub cmdPause_Click()
    If tmrSim.Enabled = True Then
        tmrSim.Enabled = False
        tmrLife.Enabled = False
        cmdPause.Caption = "Resume Simulation"
    Else
        tmrSim.Enabled = True
        tmrLife.Enabled = True
        cmdPause.Caption = "Pause Simulation"
    End If
End Sub

Private Sub Form_GotFocus()
    If frmLoading.Visible = True Then
        frmLoading.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmMain
    End
End Sub

Private Sub tmrLife_Timer()
    Dim i As Long
    
    If tmrLife.Interval = 100 Then
        'Make cells day go by
        For i = 0 To UBound(PltCells)
            If PltAlive(i) = True Then
                PltCells(i).Live
            End If
            If AnmAlive(i) = True Then
                AnmCells(i).Live
            End If
        Next i
        
        'do var stuff
        Days = Days + 1
        UpdateVars
    End If
End Sub

Private Sub tmrSim_Timer()
    If tmrSim.Interval = 1000 Then
        SimSeconds = SimSeconds + 1
        If SimSeconds >= 60 Then
            SimSeconds = 0
            SimMinutes = SimMinutes + 1
            LogIt
            If SimMinutes >= 60 Then
                SimMinutes = 0
                SimHours = SimHours + 1
            End If
        End If
    End If
End Sub

Private Sub LogIt()
    Open App.Path & "\Logs\" & txtLog.Text For Append As #1
        Print #1, "Logged At: " & Time & " on " & Date
        Print #1, "Simulation running time: " & Me.txtSimTime.Text
        Print #1, "Number of days elapsed: " & Me.txtDays.Text
        Print #1, "Total number of living cells: " & Me.txtTotLivCells.Text
        Print #1, "Total number of living plant cells: " & Me.txtPltLiv.Text
        Print #1, "Total number of living animal cells: " & Me.txtAnmLiv.Text
        Print #1, "Total number of births: " & Me.txtTotBirths.Text
        Print #1, "Total number of plant cell births: " & Me.txtPltBirths.Text
        Print #1, "Total number of animal cell births: " & Me.txtAnmBirths.Text
        Print #1, "Total number of deaths: " & Me.txtTotDeaths.Text
        Print #1, "Total number of plant cell deaths: " & Me.txtPltDeaths.Text
        Print #1, "Total number of animal cell deaths: " & Me.txtAnmDeaths.Text
        Print #1, "----------------------------------------------------" & vbCrLf
    Close #1
End Sub
