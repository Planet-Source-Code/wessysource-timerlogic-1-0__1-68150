VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   Picture         =   "frmsplsh.frx":0000
   ScaleHeight     =   2820
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   3600
      Top             =   1200
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   3480
      Top             =   2040
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3240
      ScaleHeight     =   105
      ScaleWidth      =   1425
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
      Begin VB.Shape sh 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   90
         Left            =   0
         Top             =   15
         Width           =   15
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6120
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   6480
      Top             =   1440
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
      Height          =   1515
      Left            =   6960
      TabIndex        =   2
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ª"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   72
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   210
      Left            =   3600
      TabIndex        =   0
      Top             =   1320
      Width           =   795
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.ForeColor = RGB(153, 204, 255)
    Label2.ForeColor = RGB(68, 90, 101)
    Label3.ForeColor = RGB(68, 90, 101)
    sh.BackColor = Label3.ForeColor
End Sub

Private Sub Timer1_Timer()
    frmmain.Show
    SetFocus
    Timer1.Enabled = False
    Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
    frmmain.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Timer3_Timer()
    sh.Width = sh.Width + 10
    If sh.Width >= Picture1.Width / 2 Then
        Timer4.Enabled = True
        Timer3.Enabled = False
    End If
End Sub

Private Sub Timer4_Timer()
    Timer4.Interval = 1
    sh.Width = sh.Width + 5
End Sub
