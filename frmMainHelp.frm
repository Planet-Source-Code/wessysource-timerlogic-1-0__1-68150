VERSION 5.00
Begin VB.Form frmMainHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   Icon            =   "frmMainHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   240
      ScaleHeight     =   3075
      ScaleWidth      =   5355
      TabIndex        =   2
      Top             =   1320
      Width           =   5415
      Begin VB.Image ig 
         Height          =   480
         Left            =   0
         Picture         =   "frmMainHelp.frx":0442
         Top             =   960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image iy 
         Height          =   480
         Left            =   0
         Picture         =   "frmMainHelp.frx":1084
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image iw 
         Height          =   480
         Left            =   0
         Picture         =   "frmMainHelp.frx":1CC6
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About TimerLogic 1.0"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   7
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Image i 
         Height          =   480
         Index           =   4
         Left            =   360
         Picture         =   "frmMainHelp.frx":2908
         Top             =   2400
         Width           =   480
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Macros"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   6
         Top             =   2040
         Width           =   525
      End
      Begin VB.Image i 
         Height          =   480
         Index           =   3
         Left            =   360
         Picture         =   "frmMainHelp.frx":354A
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personalize tasks"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   5
         Top             =   1560
         Width           =   1230
      End
      Begin VB.Image i 
         Height          =   480
         Index           =   2
         Left            =   360
         Picture         =   "frmMainHelp.frx":418C
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preset tasks"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   1080
         Width           =   870
      End
      Begin VB.Image i 
         Height          =   480
         Index           =   1
         Left            =   360
         Picture         =   "frmMainHelp.frx":4DCE
         Top             =   960
         Width           =   480
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What is the utility? What can it do?"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   2475
      End
      Begin VB.Image i 
         Height          =   480
         Index           =   0
         Left            =   360
         Picture         =   "frmMainHelp.frx":5A10
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   1755
         Left            =   1200
         Picture         =   "frmMainHelp.frx":6652
         Top             =   720
         Width           =   3855
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‡"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   72
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1515
      Left            =   4440
      TabIndex        =   8
      Top             =   0
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Help index:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TimerLogic 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   2550
   End
End
Attribute VB_Name = "frmMainHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim a
    For a = 0 To 4
        l(a).MousePointer = vbCustom
        l(a).MouseIcon = LoadPicture(App.Path & "/finger.cur")
    Next a
End Sub

Private Sub l_Click(Index As Integer)
    l(Index).ForeColor = &H404080
    i(Index).Picture = ig.Picture
    Wait 1
    Select Case Index
        Case 0: frmHelp0.Show
        Case 1: frmHelp1.Show
        Case 2: frmHelp2.Show
        Case 3: frmHelp3.Show
        Case 4: frmAbout.Show
    End Select
End Sub

Private Sub l_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim a
    For a = 0 To 4
        If a = Index Then
            l(a).FontUnderline = True
            i(a).Picture = iy
        Else
            l(a).FontUnderline = False
            i(a).Picture = iw
        End If
    Next a
End Sub

