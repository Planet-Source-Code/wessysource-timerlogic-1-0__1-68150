VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHelp2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help - Personalize tasks"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   Icon            =   "frmHelp2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic3 
      BackColor       =   &H00C0FFFF&
      Height          =   10455
      Left            =   0
      ScaleHeight     =   10395
      ScaleWidth      =   15195
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   15255
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   3240
         Top             =   1680
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
               Picture         =   "frmHelp2.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp2.frx":0794
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp2.frx":0A86
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp2.frx":0DD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp2.frx":112A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label21 
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
         Left            =   360
         TabIndex        =   27
         Top             =   6240
         Width           =   1155
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "You must activate ""Timer"" option in main  window to run tasks automaticly."
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
         Height          =   615
         Left            =   1560
         TabIndex        =   26
         Top             =   6240
         Width           =   4920
      End
      Begin VB.Image Image13 
         Height          =   3540
         Left            =   360
         Picture         =   "frmHelp2.frx":147C
         Top             =   2160
         Width           =   5295
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp2.frx":3E5EE
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
         Left            =   360
         TabIndex        =   25
         Top             =   960
         Width           =   6015
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Tab"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   1185
      End
   End
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
         Left            =   3240
         Top             =   1680
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
               Picture         =   "frmHelp2.frx":3E697
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp2.frx":3E9E9
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp2.frx":3ECDB
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp2.frx":3F02D
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp2.frx":3F37F
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image Image12 
         Height          =   315
         Left            =   11280
         Picture         =   "frmHelp2.frx":3F6D1
         Top             =   3120
         Width           =   960
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp2.frx":406D3
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   11280
         TabIndex        =   22
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Content all the information that you must program to open any Windows folder."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5760
         TabIndex        =   21
         Top             =   6000
         Width           =   4815
      End
      Begin VB.Image Image11 
         Height          =   3555
         Left            =   5760
         Picture         =   "frmHelp2.frx":407D9
         Top             =   6720
         Width           =   5280
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folder tab"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5760
         TabIndex        =   20
         Top             =   5400
         Width           =   1290
      End
      Begin VB.Image Image10 
         Height          =   3540
         Left            =   5760
         Picture         =   "frmHelp2.frx":7D9BB
         Top             =   1560
         Width           =   5190
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "This tab content all the information that you must program to run an application automaticly."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5760
         TabIndex        =   19
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application tab"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5760
         TabIndex        =   18
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "In this window we can find three different tabs to program an application, program to open a folder or program to open a web page."
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
         Left            =   240
         TabIndex        =   17
         Top             =   7440
         Width           =   4815
      End
      Begin VB.Image Image9 
         Height          =   345
         Left            =   240
         Picture         =   "frmHelp2.frx":B98BD
         Top             =   8640
         Width           =   5415
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Window Tabs"
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
         TabIndex        =   16
         Top             =   6960
         Width           =   2040
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "This is the way to program an application task. Now we explain the types of personalize tasks."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   15
         Top             =   6120
         Width           =   4815
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   240
         Top             =   5640
         Width           =   135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Select priority a value between 0 and 100 inclusives."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   14
         Top             =   5520
         Width           =   4815
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   240
         Top             =   3840
         Width           =   135
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "If you want, you can write your own comments to reffer this task in main list window."
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
         Left            =   480
         TabIndex        =   13
         Top             =   3720
         Width           =   4815
      End
      Begin VB.Image Image8 
         Height          =   750
         Left            =   480
         Picture         =   "frmHelp2.frx":BFA63
         Top             =   4560
         Width           =   2265
      End
      Begin VB.Image Image7 
         Height          =   645
         Left            =   480
         Picture         =   "frmHelp2.frx":C53B5
         Top             =   2760
         Width           =   2040
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "When you have selected your file, in the ""Exe path"" box have assignated the path of the file"
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
         Left            =   480
         TabIndex        =   12
         Top             =   1920
         Width           =   4815
      End
      Begin VB.Image Image6 
         Height          =   300
         Left            =   480
         Picture         =   "frmHelp2.frx":C987F
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Next, if you want to open an application you must select EXE file with ""Browse"" button:"
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
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   4815
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   240
         Top             =   840
         Width           =   135
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
         Left            =   3240
         Top             =   1680
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
               Picture         =   "frmHelp2.frx":CA5E1
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp2.frx":CA933
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp2.frx":CAC25
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp2.frx":CAF77
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp2.frx":CB2C9
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp2.frx":CB61B
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
         Left            =   6960
         TabIndex        =   9
         Top             =   5520
         Width           =   5055
      End
      Begin VB.Image Image5 
         Height          =   330
         Left            =   6960
         Picture         =   "frmHelp2.frx":CB6B4
         Top             =   6480
         Width           =   2955
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
         Left            =   6960
         TabIndex        =   8
         Top             =   6960
         Width           =   5055
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   6720
         Top             =   5640
         Width           =   135
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
         Left            =   6960
         TabIndex        =   7
         Top             =   720
         Width           =   4815
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   6960
         Picture         =   "frmHelp2.frx":CE9D6
         Top             =   1440
         Width           =   2940
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
         Left            =   6960
         TabIndex        =   6
         Top             =   2040
         Width           =   4815
      End
      Begin VB.Image Image4 
         Height          =   2310
         Left            =   6960
         Picture         =   "frmHelp2.frx":D2384
         Top             =   3000
         Width           =   2565
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   6720
         Top             =   840
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   6720
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "How program your personalize task?"
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
         Left            =   6720
         TabIndex        =   5
         Top             =   240
         Width           =   5535
      End
      Begin VB.Image Image2 
         Height          =   4710
         Left            =   240
         Picture         =   "frmHelp2.frx":E5A2E
         Top             =   2640
         Width           =   5700
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmHelp2.frx":13D0B8
         Top             =   120
         Width           =   480
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Applications && folders"
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
         TabIndex        =   3
         Top             =   240
         Width           =   3330
      End
      Begin VB.Label l2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp2.frx":13D20A
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
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label l3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personalize tasks window"
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
         TabIndex        =   1
         Top             =   2040
         Width           =   3870
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
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
Attribute VB_Name = "frmHelp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    pic1.BackColor = RGB(255, 255, 217)
    pic2.BackColor = RGB(255, 255, 217)
    pic3.BackColor = RGB(255, 255, 217)
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Label26_Click()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Static h
    Select Case Button.Index
        Case 1
            frmMainHelp.Show
        Case 3
            
            
            
            Select Case h
                Case 2
                    pic2.Visible = True
                    pic3.Visible = False
                    Toolbar1.Buttons(3).Enabled = True
                    h = h - 1
                Case 1
                    pic1.Visible = True
                    pic2.Visible = False
                    Toolbar1.Buttons(3).Enabled = False
                    Toolbar1.Buttons(4).Enabled = True
                    h = h - 1
            End Select
            Debug.Print h
        Case 4
            h = h + 1

            Select Case h
                Case 1
                    pic2.Visible = True
                    Toolbar1.Buttons(3).Enabled = True
                Case 2
                    pic3.Visible = True
                    Toolbar1.Buttons(4).Enabled = False

            End Select
                                            Debug.Print h
        Case 5
            Unload Me
        
    End Select

End Sub
