VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHelp0 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help - What is the utility? What can it do?"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   Icon            =   "frmHelp0.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5595
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   600
      Width           =   6255
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
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp0.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp0.frx":0794
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp0.frx":0A86
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   240
         Top             =   4800
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   240
         Top             =   3840
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   240
         Top             =   3120
         Width           =   135
      End
      Begin VB.Label l43 
         BackStyle       =   0  'Transparent
         Caption         =   "Run a macro. It can play your mouse movements and events, recorded previously."
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
         Left            =   480
         TabIndex        =   8
         Top             =   4680
         Width           =   5070
      End
      Begin VB.Label l42 
         BackStyle       =   0  'Transparent
         Caption         =   "Run a personalize task, for example open an user aplication entering exe path, open a folder or an a web page."
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
         Left            =   480
         TabIndex        =   7
         Top             =   3720
         Width           =   5070
      End
      Begin VB.Label l41 
         BackStyle       =   0  'Transparent
         Caption         =   "Run a preset (or programmed) task, for example shut down your pc, reboot and log Off your session"
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
         Left            =   480
         TabIndex        =   6
         Top             =   3000
         Width           =   5070
      End
      Begin VB.Label l4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "It can do the follows:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   2070
      End
      Begin VB.Label l3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What can it do?"
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
         Width           =   2385
      End
      Begin VB.Label l2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp0.frx":0DD8
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
         TabIndex        =   3
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What is the utility?"
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
         TabIndex        =   2
         Top             =   120
         Width           =   2790
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   1005
      ButtonWidth     =   1535
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "IL"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Main menu"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHelp0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Picture1.BackColor = RGB(255, 255, 217)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            
            frmMainHelp.Show
            
        Case 2
            'With Printer
            '    .FontSize = 20
            '    .Print vbNewLine
            '    .Print "TIMERLOGIC 1.0 - HELP SYSTEM"
            '    .Print vbNewLine
            '    .FontName = l1.FontName
            '    .FontBold = True
            '    .FontSize = l1.FontSize
            '    .Print l1.Caption
            '    .Print vbNewLine
            '    .FontName = l2.FontName
            '    .FontBold = False
            '    .FontSize = l2.FontSize
            '    .Print "TimerLogic 1.0 is an automatism application that " & vbCrLf & "you can program your " & vbCrLf & "own tasks on time and date and " & vbCrLf & "this run automaticly without your presence."
            '    .Print vbNewLine
            '    .FontName = l3.FontName
            '    .FontBold = True
            '    .FontSize = l3.FontSize
            '    .Print l3.Caption
            '    .Print vbNewLine
            '    .FontName = l4.FontName
            '    .FontBold = False
            '    .FontSize = l4.FontSize
            '    .Print l4.Caption
            '    .Print vbNewLine
            '    .Print " ." & l41.Caption & vbCrLf
            '    .Print " ." & l42.Caption & vbCrLf
            '    .Print " ." & l43.Caption
            '    .EndDoc
            'End With
        Case 3
            Unload Me
    End Select
End Sub
