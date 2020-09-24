VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHelp1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help - Preset tasks"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   Icon            =   "frmHelp1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic2 
      BackColor       =   &H00C0FFFF&
      Height          =   10455
      Left            =   0
      ScaleHeight     =   10395
      ScaleWidth      =   15195
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   15255
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5760
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp1.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp1.frx":0794
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp1.frx":0A86
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp1.frx":0DD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp1.frx":112A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   6000
         Top             =   3000
         Width           =   135
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   5880
         Top             =   360
         Width           =   135
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   360
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You must activate ""Timer"" option in main window to run tasks automaticly."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7440
         TabIndex        =   19
         Top             =   5280
         Width           =   7320
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remember:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   6240
         TabIndex        =   18
         Top             =   5280
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If all information are correctly, in our main list will be added your preset task.                        "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6240
         TabIndex        =   17
         Top             =   4680
         Width           =   9270
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp1.frx":147C
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6240
         TabIndex        =   16
         Top             =   2880
         Width           =   8655
      End
      Begin VB.Image Image9 
         Height          =   495
         Left            =   6240
         Picture         =   "frmHelp1.frx":1575
         Top             =   3960
         Width           =   3000
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp1.frx":630F
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6240
         TabIndex        =   15
         Top             =   1320
         Width           =   8655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp1.frx":643E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6240
         TabIndex        =   14
         Top             =   240
         Width           =   8655
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp1.frx":651F
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   600
         TabIndex        =   13
         Top             =   8040
         Width           =   4815
      End
      Begin VB.Image Image8 
         Height          =   1695
         Left            =   600
         Picture         =   "frmHelp1.frx":6683
         Top             =   6240
         Width           =   3555
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp1.frx":1A10D
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   600
         TabIndex        =   12
         Top             =   4680
         Width           =   4815
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp1.frx":1A1D6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   4815
      End
      Begin VB.Image Image7 
         Height          =   2895
         Left            =   600
         Picture         =   "frmHelp1.frx":1A272
         Top             =   1680
         Width           =   3495
      End
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H00C0FFFF&
      Height          =   10455
      Left            =   0
      ScaleHeight     =   10395
      ScaleWidth      =   15195
      TabIndex        =   0
      Top             =   600
      Width           =   15255
      Begin MSComctlLib.ImageList IL 
         Left            =   5760
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp1.frx":3B270
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp1.frx":3B5C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp1.frx":3B8B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp1.frx":3BC06
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp1.frx":3BF58
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   7680
         Top             =   5880
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   7680
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   7680
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "If you enter a invalid time format or number, an alert prompt will show to warn you."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7920
         TabIndex        =   9
         Top             =   7200
         Width           =   5055
      End
      Begin VB.Image Image5 
         Height          =   330
         Left            =   7920
         Picture         =   "frmHelp1.frx":3C2AA
         Top             =   6720
         Width           =   2955
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp1.frx":3F5CC
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7920
         TabIndex        =   8
         Top             =   5760
         Width           =   5055
      End
      Begin VB.Image Image4 
         Height          =   2310
         Left            =   7920
         Picture         =   "frmHelp1.frx":3F665
         Top             =   3240
         Width           =   2565
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Secondly you should select your date task in the calendar control. Select your month and next, your day:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7920
         TabIndex        =   7
         Top             =   2280
         Width           =   4815
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   7920
         Picture         =   "frmHelp1.frx":52D0F
         Top             =   1680
         Width           =   2940
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Firstly you must name your task. For this you type in the textbox:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7920
         TabIndex        =   6
         Top             =   960
         Width           =   4815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "How program your preset task?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7680
         TabIndex        =   5
         Top             =   360
         Width           =   4755
      End
      Begin VB.Image Image2 
         Height          =   5700
         Left            =   240
         Picture         =   "frmHelp1.frx":566BD
         Top             =   2760
         Width           =   6960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preset task window:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmHelp1.frx":D793F
         Top             =   120
         Width           =   480
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preset tasks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label l2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp1.frx":D7D81
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   4815
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   1005
      ButtonWidth     =   1535
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "IL"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Main menu"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Back"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHelp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    pic1.BackColor = RGB(255, 255, 217)
    pic2.BackColor = pic1.BackColor
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            frmMainHelp.Show
        Case 3
            pic2.Visible = False
            pic1.Visible = True
            Toolbar1.Buttons(4).Enabled = True
            Toolbar1.Buttons(3).Enabled = False
        Case 4
            pic2.Visible = True
            pic1.Visible = False
            Toolbar1.Buttons(4).Enabled = False
            Toolbar1.Buttons(3).Enabled = True
        Case 5
            Unload Me
        
    End Select
End Sub
