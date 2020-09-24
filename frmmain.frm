VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "TimerLogic 1.0"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form2"
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer timShowRunAllAnimation 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6240
      Top             =   5520
   End
   Begin VB.Timer timerEnd 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4920
      Top             =   3600
   End
   Begin VB.Timer RunAllTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2400
      Top             =   600
   End
   Begin MSComctlLib.ImageList IL2 
      Left            =   7320
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1F94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3838
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":448A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":50DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5D2E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "List operations"
      Height          =   1425
      Left            =   6360
      TabIndex        =   25
      Top             =   8925
      Width           =   4425
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   1140
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2011
         ButtonWidth     =   3440
         ButtonHeight    =   1005
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "IL2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Edit selected     "
               Object.ToolTipText     =   "Edit selected"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Sort by                "
               Object.ToolTipText     =   "Sort by task name"
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A"
                     Text            =   "Task name"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "B"
                     Text            =   "Time"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "C"
                     Text            =   "Date"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Save task          "
               ImageIndex      =   4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Search                  "
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.Timer MainTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   600
   End
   Begin VB.Timer timToShowIcons 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8160
      Top             =   7440
   End
   Begin VB.Timer timToEnableStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9840
      Top             =   6960
   End
   Begin MSComctlLib.ListView LV2 
      Height          =   1335
      Left            =   10800
      TabIndex        =   24
      Top             =   9000
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2355
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ILsmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Event"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Runned at"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   3043
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Priority"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   1860
      ScaleHeight     =   1275
      ScaleWidth      =   4395
      TabIndex        =   11
      Top             =   9000
      Width           =   4455
      Begin VB.PictureBox picPrior 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   855
         ScaleHeight     =   195
         ScaleWidth      =   3225
         TabIndex        =   16
         Top             =   540
         Width           =   3255
         Begin VB.Label lPriority 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 %"
            Height          =   195
            Left            =   1440
            TabIndex        =   18
            Top             =   0
            Width           =   255
         End
         Begin VB.Shape shPriority 
            BackColor       =   &H8000000D&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   255
            Left            =   0
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.Label lItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3840
         TabIndex        =   28
         Top             =   1050
         Width           =   45
      End
      Begin VB.Label Label1 
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
         Index           =   5
         Left            =   3360
         TabIndex        =   27
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label lDisable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disable task"
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3360
         MouseIcon       =   "frmmain.frx":6980
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   60
         Width           =   870
      End
      Begin VB.Label lDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   840
         TabIndex        =   22
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label lTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   840
         TabIndex        =   21
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label1 
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
         Index           =   4
         Left            =   270
         TabIndex        =   20
         Top             =   1050
         Width           =   480
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   285
         TabIndex        =   19
         Top             =   810
         Width           =   480
      End
      Begin VB.Label lstatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   855
         TabIndex        =   17
         Top             =   300
         Width           =   45
      End
      Begin VB.Label lTask 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   840
         TabIndex        =   15
         Top             =   60
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Priority:"
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
         Left            =   135
         TabIndex        =   14
         Top             =   540
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   135
         TabIndex        =   13
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   255
         TabIndex        =   12
         Top             =   60
         Width           =   495
      End
   End
   Begin VB.Timer timToShowBar 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8760
      Top             =   4680
   End
   Begin VB.Timer timToCloseTray 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8520
      Top             =   5880
   End
   Begin VB.Timer timRunningSelected 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7080
      Top             =   5280
   End
   Begin MSComctlLib.ImageList ILsmall 
      Left            =   4320
      Top             =   10320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6AD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7276
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":73D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":772A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7D96
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ILAAA 
      Left            =   3360
      Top             =   10080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":80E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":853A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":898C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9130
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9482
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9774
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A16A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A8BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":AC0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":AF60
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B2B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B704
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":BB56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   10710
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   776
            TextSave        =   "14:46"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1765
            MinWidth        =   1765
            TextSave        =   "17/07/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Picture         =   "frmmain.frx":BFA8
            Text            =   "Timer status: disabled"
            TextSave        =   "Timer status: disabled"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19404
            MinWidth        =   19404
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   776
            MinWidth        =   776
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   9735
      Left            =   0
      ScaleHeight     =   9675
      ScaleWidth      =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1860
      Begin VB.CommandButton Command2 
         Caption         =   "Preferences"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   5520
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tasks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1815
      End
      Begin VB.Image i 
         Height          =   450
         Index           =   5
         Left            =   0
         MouseIcon       =   "frmmain.frx":C4AA
         MousePointer    =   99  'Custom
         Picture         =   "frmmain.frx":C5FC
         Top             =   7920
         Width           =   1320
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show clock"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   390
         TabIndex        =   29
         Top             =   8520
         Width           =   840
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   10
         Top             =   7200
         Width           =   540
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Load list..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   9
         Top             =   5040
         Width           =   720
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Macros"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   8
         Top             =   3720
         Width           =   525
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Applications && folders"
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Presets"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   1320
         Width           =   525
      End
      Begin VB.Image i 
         Height          =   480
         Index           =   4
         Left            =   600
         MouseIcon       =   "frmmain.frx":E52E
         Picture         =   "frmmain.frx":E838
         Top             =   6600
         Width           =   480
      End
      Begin VB.Image i 
         Height          =   480
         Index           =   3
         Left            =   600
         MouseIcon       =   "frmmain.frx":EC7A
         Picture         =   "frmmain.frx":EF84
         Top             =   4440
         Width           =   480
      End
      Begin VB.Image i 
         Height          =   480
         Index           =   2
         Left            =   600
         MouseIcon       =   "frmmain.frx":F3C6
         Picture         =   "frmmain.frx":F6D0
         Top             =   3120
         Width           =   480
      End
      Begin VB.Image i 
         Height          =   480
         Index           =   1
         Left            =   600
         MouseIcon       =   "frmmain.frx":F9DA
         Picture         =   "frmmain.frx":FCE4
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image i 
         Height          =   480
         Index           =   0
         Left            =   600
         MouseIcon       =   "frmmain.frx":FE36
         Picture         =   "frmmain.frx":10140
         Top             =   720
         Width           =   480
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1005
      ButtonWidth     =   2170
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ILAAA"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save list"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Erase database"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Timer"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Run selected"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Run all"
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Delete task"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Delete from"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Delete list"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Hide"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   16
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView LV 
      Height          =   8295
      Left            =   1860
      TabIndex        =   4
      Top             =   600
      Width           =   13410
      _ExtentX        =   23654
      _ExtentY        =   14631
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ILsmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Task name"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Task"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Start time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Start date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Priority"
         Object.Width           =   1296
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Type"
         Object.Width           =   1296
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Comments / Parameters"
         Object.Width           =   6033
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPop"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show main"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu slkdfssss 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListView1_Click()
    MsgBox ListView1.ColumnHeaders(4).Width & " " & ListView1.ColumnHeaders(5).Width & " " & ListView1.ColumnHeaders(6).Width & " " & ListView1.ColumnHeaders(8).Width

End Sub

Private Sub chkLow_Click()
    chkNormal.Value = vbUnchecked
    chkHigh.Value = vbUnchecked
End Sub
Private Sub chkNormal_Click()

    chkLow.Value = vbUnchecked
    chkHigh.Value = vbUnchecked
End Sub
Private Sub chkHigh_Click()

    chkNormal.Value = vbUnchecked
    chkLow.Value = vbUnchecked
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkOnce_Click()
    If chkOnce.Value = vbChecked Then
        chkOnce.Caption = "Personalize"
        optDay.Enabled = True
        optMonth.Enabled = True
        optYear.Enabled = True
        
    Else
        chkOnce.Caption = "Only once"
        optDay.Enabled = False
        optMonth.Enabled = False
        optYear.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    For b = 0 To 5
        i(b).MousePointer = vbCustom
    Next b
    shpriority.BackColor = RGB(49, 106, 197)
    On Error Resume Next
    Set bbd = OpenDatabase(App.Path & "\DataBaseList.mdb")
    Set tbl = bbd.OpenRecordset("tblSettings")
    
    tbl.MoveLast
    timRunAllInterval = tbl("Interval 1")
    timWaitRun = tbl("Interval 2")
    If tbl("Prompt") = True Then
        showPrompt = True
    Else
        showPrompt = False
    End If
    If tbl("Sound") = True Then
        playSound = True
    Else
        playSound = False
    End If
    If tbl("GridLines") = True Then
        showGrid = True
        LV.GridLines = True
    Else
        showGrid = False
        LV.GridLines = False
    End If
    Debug.Print tbl("GridLines")
    If tbl("Full select") = True Then
        fullSelect = True
        LV.FullRowSelect = True
    Else
        fullSelect = False
        LV.FullRowSelect = False
    End If
    If tbl("HoverSelection") = True Then
        Hover = True
        LV.HoverSelection = True
    Else
        Hover = False
        LV.HoverSelection = False
    End If
    If tbl("HotTracking") = True Then
        hTracking = True
        LV.HotTracking = True
    Else
        hTracking = False
        LV.HotTracking = False
    End If
    DoEvents
    bbd.Close

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Result As Long
 Dim msg As Long
'the value of X will vary depending upon the scalemode setting
  If Me.ScaleMode = vbPixels Then
   msg = x
  Else
   msg = x / Screen.TwipsPerPixelX
  End If
  Select Case msg
   Case WM_LBUTTONUP        '514 restore form window
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
   Case WM_LBUTTONDBLCLK    '515 restore form window
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
   Case WM_RBUTTONUP        '517 display popup menu
    Result = SetForegroundWindow(Me.hwnd)
    Me.PopupMenu Me.mnuPop
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   frmConfirmation.Show
   Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Shell_NotifyIcon NIM_DELETE, nid

End Sub

Private Sub i_Click(Index As Integer)
Dim Itm As ListItem
    Select Case Index
        Case 0
            Task = "Impl"
            frmWait.Show 1
        Case 1
            Task = "Pzl"
            frmWait.Show 1
        Case 2
            Task = "Ovr"
            frmWait.Show
        Case 3
            Task = "Ld"
            frmWait.Show 1
        Case 4
            frmOptions.Show 1
        Case 5
            Dim dblReturn As Double
            dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)

    End Select
End Sub

Private Sub i_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim i
    For i = 0 To 5
        If i = Index Then
            l(i).FontUnderline = True
        Else
            l(i).FontUnderline = False
        End If
    Next
    With StatusBar1
        Select Case Index
            Case 0: .Panels(4).Text = "Presets: you can select an a preset task, a task programmed."
            Case 1: .Panels(4).Text = "Applications & folders: you can program that an exe file, folder path or a web page, opens automaticly."
            Case 2: .Panels(4).Text = "Macros: in this task, you can record mouse events and movements."
            Case 3: .Panels(4).Text = "Load list: this load an a task list from data base program. Previously you must save the list."
            Case 4: .Panels(4).Text = "Options: you can modify the program configuration."
            Case 5: .Panels(4).Text = "Show clock: show the system clock."
        End Select
    End With
End Sub

Private Sub lDisable_Click()
    Select Case lDisable.Caption
        Case "Disable task"
            LV.selectedItem.SubItems(5) = "Disabled"
            LV.selectedItem.SmallIcon = 4
            lDisable = "Enable task"
        Case "Enable task"
            LV.selectedItem.SubItems(5) = "Waiting"
            Select Case LV.selectedItem.SubItems(7)
                Case "Impl"
                LV.selectedItem.SmallIcon = 2
                Case "Pzl"
                LV.selectedItem.SmallIcon = 3
                Case "Ovr"
                LV.selectedItem.SmallIcon = 5
            End Select
            
            lDisable = "Disable task"
    End Select
End Sub

Private Sub lDisable_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lDisable.ForeColor = vbBlue
    lDisable.FontUnderline = True
End Sub









Private Sub LV_DblClick()
    If LV.ListItems.Count = 0 Then Exit Sub
    Select Case LV.selectedItem.SubItems(7)
        Case "Impl"
            editImpl = True
            Edit "Impl"
                   
        Case "Pzl"
            editPzl = True
            Edit "Pzl"
        Case "Ovr"
            editOvr = True
            Edit "Ovr"
    End Select
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lTask = LV.selectedItem.SubItems(1)
    lstatus = LV.selectedItem.SubItems(5)
    lTime = LV.selectedItem.SubItems(3)
    lDate = LV.selectedItem.SubItems(4)
    lItem = LV.selectedItem.Index
    lDisable.Enabled = True
    If LV.selectedItem.SubItems(5) = "Disabled" Then
        lDisable = "Enable task"
    Else
        lDisable = "Disable task"
    End If
    lPriority = Val(LV.selectedItem.SubItems(6)) & " %"
    shpriority.Width = ((picPrior.Width) * Val(lPriority)) / 100
    If shpriority.Width >= (picPrior.Width / 2) + lPriority.Width Then
        lPriority.ForeColor = vbWhite
    Else
        lPriority.ForeColor = vbBlack
    End If
    Toolbar2.Buttons(1).Enabled = True
End Sub
Private Sub MainTimer_Timer()
    Dim k%, tType$
    For k% = 1 To LV.ListItems.Count
        If (LV.ListItems.Item(k%).SubItems(3) = Time) And (LV.ListItems.Item(k%).SubItems(4) = Date) Then
            If (LV.ListItems.Item(k%).SubItems(5) = "Disabled") Then Exit Sub
            tType$ = LV.ListItems.Item(k%).SubItems(7)
            Run tType$, LV, k%, False, playSound
            If frmmain.WindowState = vbMinimized Then
                Load frmAlert
                frmAlert.lblWarn = LV.ListItems.Item(k%).SubItems(1)
                frmAlert.lblTime = LV.ListItems.Item(k%).SubItems(3)
                frmAlert.Show 1
            End If
        End If
    Next
End Sub

Private Sub optDay_Click()
Dim withoutParam As String

If Right(LV.selectedItem.SubItems(8), 1) = "D" Or Right(LV.selectedItem.SubItems(8), 1) = "M" Or _
Right(LV.selectedItem.SubItems(8), 1) = "Y" Then
    withoutParam = Left(LV.selectedItem.SubItems(8), Len(LV.selectedItem.SubItems(8)) - 3)
    LV.selectedItem.SubItems(8) = ""
    LV.selectedItem.SubItems(8) = withoutParam & ", D"
Else
    LV.selectedItem.SubItems(8) = LV.selectedItem.SubItems(8) & ", D"
End If
End Sub

Private Sub optMonth_Click()
Dim withoutParam As String

If Right(LV.selectedItem.SubItems(8), 1) = "D" Or Right(LV.selectedItem.SubItems(8), 1) = "M" Or _
Right(LV.selectedItem.SubItems(8), 1) = "Y" Then
    withoutParam = Left(LV.selectedItem.SubItems(8), Len(LV.selectedItem.SubItems(8)) - 3)
    LV.selectedItem.SubItems(8) = ""
    LV.selectedItem.SubItems(8) = withoutParam & ", M"
Else
    LV.selectedItem.SubItems(8) = LV.selectedItem.SubItems(8) & ", M"
End If
End Sub

Private Sub optYear_Click()
Dim withoutParam As String

If Right(LV.selectedItem.SubItems(8), 1) = "D" Or Right(LV.selectedItem.SubItems(8), 1) = "M" Or _
Right(LV.selectedItem.SubItems(8), 1) = "Y" Then
    withoutParam = Left(LV.selectedItem.SubItems(8), Len(LV.selectedItem.SubItems(8)) - 3)
    LV.selectedItem.SubItems(8) = ""
    LV.selectedItem.SubItems(8) = withoutParam & ", Y"
Else
    LV.selectedItem.SubItems(8) = LV.selectedItem.SubItems(8) & ", Y"
End If
End Sub

Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lDisable.ForeColor = vbBlack
    lDisable.FontUnderline = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i
    For i = 0 To 5
        l(i).FontUnderline = False
    Next
End Sub

Private Sub RunAllTimer_Timer()
RunAllTimer.Interval = timRunAllInterval
Static c%
c% = c% + 1
If c% > 0 Then
    RunAll LV.ListItems.Item(c).SubItems(7), LV, c%
End If
End Sub

Private Sub timerEnd_Timer()
     frmConfirmation.EndApp
End Sub

Private Sub timRunningSelected_Timer()
    If Toolbar1.Buttons(4).Value = tbrPressed Then
        
        StatusBar1.Panels(3).Picture = LoadPicture(App.Path & "/timenabled.bmp")
        StatusBar1.Panels(3).Text = "Timer status: enabled"
    Else
        StatusBar1.Panels(3).Picture = LoadPicture(App.Path & "/timdisabled.bmp")
        StatusBar1.Panels(3).Text = "Timer status: disabled"
    End If
    timRunningSelected.Enabled = False
End Sub

Private Sub timShowRunAllAnimation_Timer()
    Static i
    i = i + 1
    With StatusBar1.Panels(3)
        Select Case i
            Case 1
                .Picture = LoadPicture(App.Path & "/rundown_0.bmp")
            Case 2
                .Picture = LoadPicture(App.Path & "/rundown_1.bmp")
            Case 3
                .Picture = LoadPicture(App.Path & "/rundown_2.bmp")
            Case 4
                .Picture = LoadPicture(App.Path & "/rundown_3.bmp")
            Case 5
                .Picture = LoadPicture(App.Path & "/rundown_4.bmp")
            Case 6
                .Picture = LoadPicture(App.Path & "/rundown_5.bmp")
                i = 1
        End Select
    End With
End Sub

Private Sub timToCloseTray_Timer()
    Static i
    i = i + 1
    StatusBar1.Panels(5).Text = i
    If IsAutomatic = False Then
        If i = Val(Mid(LV.selectedItem.SubItems(8), 7)) Then
            close_door
            i = 0
            StatusBar1.Panels(5).Text = ""
            timToCloseTray.Enabled = False
        End If
    Else
        If i = Val(Mid(LV.ListItems.Item(yourIndex).SubItems(8), 7)) Then
            close_door
            i = 0
            StatusBar1.Panels(5).Text = ""
            timToCloseTray.Enabled = False
        End If
    End If

End Sub

Private Sub timToEnableStart_Timer()
    Static i
    i = i + 1
    StatusBar1.Panels(5) = i
    If IsAutomatic = False Then
        If i = Val(Mid(LV.selectedItem.SubItems(8), 11)) Then
            EnableStartButton True
            i = 0
            StatusBar1.Panels(5) = ""
            timToEnableStart.Enabled = False
        End If
    Else
        If i = Val(Mid(LV.ListItems.Item(yourIndex).SubItems(8), 11)) Then
            EnableStartButton True
            i = 0
            StatusBar1.Panels(5) = ""
            timToEnableStart.Enabled = False
        End If
    End If
End Sub

Private Sub timToShowBar_Timer()
    Static i
    i = i + 1
    StatusBar1.Panels(5) = i
    If IsAutomatic = False Then
        If i = Val(Mid(LV.selectedItem.SubItems(8), 7)) Then
            show_taskbar
            i = 0
            StatusBar1.Panels(5) = ""
            timToShowBar.Enabled = False
        End If
    Else
        If i = Val(Mid(LV.ListItems.Item(yourIndex).SubItems(8), 7)) Then
            show_taskbar
            i = 0
            StatusBar1.Panels(5) = ""
            timToShowBar.Enabled = False
        End If
    End If
End Sub

Private Sub timToShowIcons_Timer()
    Static i
    Dim hwnda As Long
    i = i + 1
    StatusBar1.Panels(5) = i
    If IsAutomatic = False Then
        If i = Val(Mid(LV.selectedItem.SubItems(8), 15)) Then
            hwnda = FindWindowEx(0&, 0&, "Progman", vbNullString)
            ShowWindow hwnda, 5
            i = 0
            StatusBar1.Panels(5) = ""
            timToShowIcons.Enabled = False
        End If
    Else
        If i = Val(Mid(LV.ListItems.Item(yourIndex).SubItems(8), 15)) Then
            hwnda = FindWindowEx(0&, 0&, "Progman", vbNullString)
            ShowWindow hwnda, 5
            i = 0
            StatusBar1.Panels(5) = ""
            timToShowIcons.Enabled = False
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1
            Set bbd = OpenDatabase(App.Path & "\DataBaseList.mdb")
            Set tbl = bbd.OpenRecordset("ListOfTasks")
            
            MousePointer = vbHourglass
            Dim i
            For i = 1 To LV.ListItems.Count
                tbl.AddNew
                
                tbl("TaskName") = LV.ListItems.Item(i).SubItems(1)
                tbl("Task") = LV.ListItems.Item(i).SubItems(2)
                tbl("Time") = LV.ListItems.Item(i).SubItems(3)
                tbl("Date") = LV.ListItems.Item(i).SubItems(4)
                tbl("Status") = LV.ListItems.Item(i).SubItems(5)
                tbl("Priority") = LV.ListItems.Item(i).SubItems(6)
                tbl("Type") = LV.ListItems.Item(i).SubItems(7)
                tbl("Comments") = LV.ListItems.Item(i).SubItems(8)
                tbl.Update
            Next i
                

                bbd.Close
                
                DoEvents
                Toolbar1.Buttons(2).Enabled = True
                MsgBox "Successfully saved at " & vbCrLf & vbCrLf & "''" & (App.Path & "\DateBaseList.mdb''"), vbInformation + vbOKOnly, App.Title
                
                MousePointer = vbDefault
        Case 2
        On Error GoTo ShowDisplay
        If MsgBox("Do you want to remove database? Tasks saved will be not available.", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        Else
            Set bbd = OpenDatabase(App.Path & "\DataBaseList.mdb")
            Set tbl = bbd.OpenRecordset("ListOfTasks")
       
            MousePointer = vbHourglass
            tbl.MoveFirst
            Do Until tbl.BOF
            tbl.Delete
            tbl.MoveNext
            Loop '

                bbd.Close
                
                DoEvents
                Toolbar1.Buttons(2).Enabled = False
                Exit Sub
    End If
ShowDisplay:
                If Err.Number = 3021 Then
                    MsgBox "Data base content deleted successfully.", vbInformation, App.Title
                    MousePointer = vbDefault
                    Toolbar1.Buttons(2).Enabled = False
                End If
        Case 4
            If showPrompt = True Then
                MsgBox "The task list is empty.", vbExclamation, App.Title:
                Toolbar1.Buttons(4).Value = tbrUnpressed
                Exit Sub
            Else
                If Toolbar1.Buttons(4).Value = tbrPressed Then
                    StatusBar1.Panels(3).Picture = LoadPicture(App.Path & "/timenabled.bmp")
                    StatusBar1.Panels(3).Text = "Timer status: enabled"
                    MainTimer.Enabled = True
                Else
                    StatusBar1.Panels(3).Picture = LoadPicture(App.Path & "/timdisabled.bmp")
                    StatusBar1.Panels(3).Text = "Timer status: disabled"
                    MainTimer.Enabled = False
                End If
            End If
        Case 5
            IsAutomatic = False
            If LV.ListItems.Count = 0 Then MsgBox "You can't run. The task list is empty.", vbExclamation, App.Title: Exit Sub
            StatusBar1.Panels(3).Picture = LoadPicture(App.Path & "/runningsel.bmp")
            StatusBar1.Panels(3).Text = "Running selected"
            timRunningSelected.Enabled = True
            Select Case LV.selectedItem.SubItems(2)
                Case "Shut down"
                    Select Case LV.selectedItem.SubItems(8)
                        Case "No Prompt"
                        'exitwindows code
                        Case "Prompt"
                            If MsgBox("Do you want to shut down computer?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                               cExitWindows.ExitWindows WE_SHUTDOWN
                            Else
                            End If
                        Case Else
                            
                            secBeforeShutDown = Val(Mid(LV.selectedItem.SubItems(8), 7))
                            frmDelayShutDown.Show 1
                    End Select
                Case "Reboot"
                    Select Case LV.selectedItem.SubItems(8)
                        Case "No Prompt"
                        Case "Prompt"
                            If MsgBox("Do you want to reboot your computer?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                                cExitWindows.ExitWindows WE_REBOOT
                            Else
                                MsgBox "NO"
                            End If
                        Case Else
                            
                            secBeforeShutDown = Val(Mid(LV.selectedItem.SubItems(8), 7))
                            frmDelayShutDown.Show 1
                    End Select
                Case "Log Off"
                    Select Case LV.selectedItem.SubItems(8)
                        Case "No Prompt"
                        Case "Prompt"
                            If MsgBox("Do you want to log off your session?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                                cExitWindows.ExitWindows WE_LOGOFF
                            Else
                                MsgBox "NO"
                            End If
                        Case Else
                            
                            secBeforeShutDown = Val(Mid(LV.selectedItem.SubItems(8), 7))
                            frmDelayShutDown.Show 1
                    End Select
                Case "Open/close cd tray"
                open_door
                If Len(LV.selectedItem.SubItems(8)) > 0 Then
                    timToCloseTray.Enabled = True
                End If
                
                Case "Screensaver"
                    Dim lResult As Long
                    lResult = SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
                    
                Case "Hide/show status bar"
                hide_taskbar
                If Len(LV.selectedItem.SubItems(8)) > 0 Then
                    timToShowBar.Enabled = True
                End If
                
                Case "Set cursor pos"
                    Dim posX As Double, posY As Double
                            
                    posX = Mid(LV.selectedItem.SubItems(8), 5, InStr(5, LV.selectedItem.SubItems(8), ",") - 5)
                    posY = Mid(LV.selectedItem.SubItems(8), InStr(5, LV.selectedItem.SubItems(8), ",") + 6)
                    
                    SetCursorPos posX, posY
                Case "Beep"
                    Dim Freq, Dur
                    Freq = Mid(LV.selectedItem.SubItems(8), 7, InStr(7, LV.selectedItem.SubItems(8), ",") - 7)
                    Dur = Mid(LV.selectedItem.SubItems(8), InStr(7, LV.selectedItem.SubItems(8), ",") + 7)
                    
                    Beep Freq, Dur
                Case "Show window"
                    Dim Window_Handle As Long
                    Dim subItemWindowTitle
                    subItemWindowTitle = LV.selectedItem.SubItems(8)
                    Window_Handle = FindWindow(vbNullString, subItemWindowTitle)
                    If Window_Handle Then
                        ShowWindow Window_Handle, 3
                    Else: MsgBox "Window title not found opened.", vbCritical, App.Title
                    End If
                Case "Disable start button"
                    EnableStartButton False
                    If Len(LV.selectedItem.SubItems(8)) > 0 Then
                        timToEnableStart.Enabled = True
                    End If
                Case "Windows Update"
                    ShellExecute hwnd, "open", "wupdmgr", "", "", 1
                Case "Hide/show icons"
                    Dim hwnda As Long
                    hwnda = FindWindowEx(0&, 0&, "Progman", vbNullString)
                    ShowWindow hwnda, 0
                    If Len(LV.selectedItem.SubItems(8)) > 0 Then
                        timToShowIcons.Enabled = True
                    End If
                Case "Kill file"
                    If MsgBox("This is only simulation. Do you want to delete this file now really?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                        MsgBox "The file " & LV.selectedItem.SubItems(8) & " should be removed from the path." & vbCrLf & "Now, the file is in his path yet.", vbInformation, App.Title: Exit Sub
                    Else
                        Kill LV.selectedItem.SubItems(8)
                    End If
                Case "Create file"
                    Dim inFileName, outPath
                    inFileName = Mid(LV.selectedItem.SubItems(8), 11, InStr(11, LV.selectedItem.SubItems(8), "Path:") - 13)
                    outPath = Mid(LV.selectedItem.SubItems(8), InStr(11, LV.selectedItem.SubItems(8), "Path:") + 6)
                    
                    Open outPath & "/" & inFileName For Output As 1#
                    Close 1#
                Case "Remove trash content"
                    If MsgBox("This is only simulation. Do you want to remove recycle bin content now?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                        MsgBox "The recycle bin content should have to be removed. Nothing is modified.", vbInformation, App.Title: Exit Sub
                    Else
                        Call EmptyRec(7)
                    End If
                Case "Print text"
                    Printer.Print "Auto print at: " & LV.selectedItem.SubItems(3) & " of " & LV.selectedItem.SubItems(4)
                    Printer.Print LV.selectedItem.SubItems(8)
                    Printer.EndDoc
                Case "Macro"
                    isFromMain = True
                    Load frmOvr
                    FN = LV.selectedItem.SubItems(8)
                    LoadFile FN
                    With frmOvr
                        .AutomaticTimer.Enabled = True
                        .Show
                    End With
            End Select
        Case 6
            If LV.ListItems.Count = 0 Then MsgBox "The task list is empty.", vbExclamation, App.Title: Toolbar1.Buttons(6).Value = tbrUnpressed: Exit Sub
            RunAllTimer.Enabled = True
            timShowRunAllAnimation.Enabled = True
            StatusBar1.Panels(3).Text = "Running down"
        Case 8
            If LV.ListItems.Count = 0 Then MsgBox "The task list is empty.", vbExclamation, App.Title: Exit Sub
            frmPrint.Show 1
        Case 10
            If LV.ListItems.Count = 0 Then MsgBox "The task list is empty.", vbExclamation, App.Title: Exit Sub
            If MsgBox("Do you want to remove from this list: " & LV.selectedItem.SubItems(1) & "?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                LV.ListItems.Remove LV.selectedItem.Index
            End If
        Case 11
            On Error Resume Next
            If LV.ListItems.Count = 0 Then MsgBox "The task list is empty.", vbExclamation, App.Title: Exit Sub
            If MsgBox("Do you want to remove from: " & LV.selectedItem.SubItems(1) & " to " & LV.ListItems.Item(LV.ListItems.Count).SubItems(1) & "?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            Else
                Dim l, a, b
                a = LV.selectedItem.Index
                b = LV.ListItems.Count
                For l = a To b
                    LV.ListItems.Remove l
                Next
            End If
        Case 12
            If LV.ListItems.Count = 0 Then MsgBox "The task list is empty.", vbExclamation, App.Title: Exit Sub
            If MsgBox("Do you want to remove all this list? This will not affect data base.", vbQuestion + vbYesNo, App.Title) = vbYes Then
                LV.ListItems.Clear
                Toolbar1.Buttons(10).Enabled = False
                Toolbar1.Buttons(11).Enabled = False
                Toolbar1.Buttons(12).Enabled = False
                
                Toolbar2.Buttons(1).Enabled = False
                Toolbar2.Buttons(2).Enabled = False
                Toolbar2.Buttons(3).Enabled = False
                Toolbar2.Buttons(4).Enabled = False
                Me.i(3).Enabled = True
                Me.l(3).Enabled = True
            End If
        Case 14
            frmMainHelp.Show
        Case 15
            Me.Show
            Me.Refresh
            With nid
                .cbSize = Len(nid)
                .hwnd = Me.hwnd
                .uId = vbNull
                .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
                .uCallBackMessage = WM_MOUSEMOVE
                .hIcon = Me.Icon
                .szTip = "This is a Sample Tool Tip" & vbNullChar
            End With
            Shell_NotifyIcon NIM_ADD, nid
            App.TaskVisible = False
           ' Hide
        Case 16
            Form_QueryUnload 1, 0
    End Select
                
                
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            Select Case LV.selectedItem.SubItems(7)
                Case "Impl"
                    editImpl = True
                    Edit "Impl"
                    
                Case "Pzl"
                    editPzl = True
                    Edit "Pzl"
                Case "Ovr"
                    editOvr = True
                    Edit "Ovr"
            End Select
        Case 3
            Set bbd = OpenDatabase(App.Path & "\DataBaseList.mdb")
            Set tbl = bbd.OpenRecordset("ListOfTasks")
            
            MousePointer = vbHourglass
            tbl.AddNew
                
            tbl("TaskName") = LV.selectedItem.SubItems(1)
            tbl("Task") = LV.selectedItem.SubItems(2)
            tbl("Time") = LV.selectedItem.SubItems(3)
            tbl("Date") = LV.selectedItem.SubItems(4)
            tbl("Status") = LV.selectedItem.SubItems(5)
            tbl("Priority") = LV.selectedItem.SubItems(6)
            tbl("Type") = LV.selectedItem.SubItems(7)
            tbl("Comments") = LV.selectedItem.SubItems(8)
            tbl.Update

            bbd.Close
            DoEvents
            Toolbar1.Buttons(2).Enabled = True
            MousePointer = vbDefault
            MsgBox "Task: ''" & LV.selectedItem.SubItems(1) & "'' added to database.", vbInformation, App.Title
            DoEvents
        Case 4
            If LV.ListItems.Count = 0 Then MsgBox "The list is empty.", vbExclamation, App.Title: Exit Sub

            frmSearch.Show
    End Select
End Sub

Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "A"
            Toolbar2.Buttons(2).Caption = "Sort by task name"
            Toolbar2.Buttons(2).Image = 2
            LV.SortKey = 1
            LV.Sorted = True
        Case "B"
            Toolbar2.Buttons(2).Caption = "Sort by time        "
            Toolbar2.Buttons(2).Image = 6
            LV.SortKey = 3
            LV.Sorted = True
        Case "C"
            Toolbar2.Buttons(2).Caption = "Sort by date        "
            LV.SortKey = 4
            LV.Sorted = True
            Toolbar2.Buttons(2).Image = 7
    End Select
End Sub
