VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Runnning task..."
   ClientHeight    =   525
   ClientLeft      =   12075
   ClientTop       =   10440
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   2520
      Top             =   120
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   1830
      TabIndex        =   3
      Top             =   270
      Width           =   480
   End
   Begin VB.Label lblWarn 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   1830
      TabIndex        =   2
      Top             =   30
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Start time:"
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
      Left            =   870
      TabIndex        =   1
      Top             =   270
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Task name:"
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
      Left            =   750
      TabIndex        =   0
      Top             =   30
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   270
      Picture         =   "frmAlert.frx":0000
      Top             =   150
      Width           =   240
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    MakeAlwaysOnTop Me, True
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
