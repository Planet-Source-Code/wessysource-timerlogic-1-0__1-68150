VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1200
      Picture         =   "frmSearch.frx":0152
      TabIndex        =   19
      Top             =   840
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   4065
      TabIndex        =   7
      Top             =   1440
      Width           =   4095
      Begin VB.Label lItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   600
         TabIndex        =   18
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label la 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   720
         TabIndex        =   16
         Top             =   840
         Width           =   45
      End
      Begin VB.Label lTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   720
         TabIndex        =   15
         Top             =   600
         Width           =   45
      End
      Begin VB.Label lTask 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   720
         TabIndex        =   14
         Top             =   360
         Width           =   45
      End
      Begin VB.Label lTaskName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1200
         TabIndex        =   13
         Top             =   120
         Width           =   45
      End
      Begin VB.Label la 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label la 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label la 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Task:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label la 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblNoResults 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No results"
         Height          =   195
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      Picture         =   "frmSearch.frx":0D94
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      Picture         =   "frmSearch.frx":19D6
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   3240
      Picture         =   "frmSearch.frx":2618
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmSearch.frx":325A
      Left            =   120
      List            =   "frmSearch.frx":326A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Equals to:"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Search by:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSearch_Click()
    Select Case Combo1.Text
        Case "Task name"
            Search frmmain.LV, "Task name", txtSearch
        Case "Task"
            Search frmmain.LV, "Task", txtSearch
        Case "Time"
            Search frmmain.LV, "Time", txtSearch
        Case "Date"
            Search frmmain.LV, "Date", txtSearch
    End Select
End Sub

Private Sub Command1_Click()
    Dim a
    For a = 0 To 4
        la(a).Visible = False
    Next a
    lTaskName = ""
    lTask = ""
    lTime = ""
    lDate = ""
    lItem = ""
    txtSearch = ""
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
