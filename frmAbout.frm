VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help - About TimerLogic 1.0"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic1 
      BackColor       =   &H00C0FFFF&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      Begin MSComctlLib.ImageList IL 
         Left            =   5280
         Top             =   480
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
               Picture         =   "frmAbout.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAbout.frx":0794
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAbout.frx":0A86
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author:   Wessource"
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
         TabIndex        =   7
         Top             =   1920
         Width           =   2025
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "wes_borland_fvck84@hotmail.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   1200
         MouseIcon       =   "frmAbout.frx":0DD8
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1560
         Width           =   3360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact:"
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
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (C) 2006. All rights reserved."
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
         TabIndex        =   4
         Top             =   1200
         Width           =   3810
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmAbout.frx":0F2A
         Top             =   120
         Width           =   480
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About TimerLogic 1.0"
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
         Top             =   240
         Width           =   3240
      End
      Begin VB.Label l2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build 1.0.0"
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
         TabIndex        =   1
         Top             =   840
         Width           =   1020
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
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
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    pic1.BackColor = RGB(255, 255, 217)
End Sub

Private Sub Label3_Click()
    ShellExecute hwnd, "open", "www.hotmail.com", "", "", 1
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
