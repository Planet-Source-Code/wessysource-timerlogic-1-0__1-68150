VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHelp3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help - Macros"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12660
   Icon            =   "frmHelp3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   12660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic1 
      BackColor       =   &H00C0FFFF&
      Height          =   10455
      Left            =   0
      ScaleHeight     =   10395
      ScaleWidth      =   12555
      TabIndex        =   0
      Top             =   600
      Width           =   12615
      Begin MSComctlLib.ImageList IL 
         Left            =   3840
         Top             =   2640
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
               Picture         =   "frmHelp3.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp3.frx":0794
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHelp3.frx":0A86
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
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
         Height          =   270
         Left            =   1560
         TabIndex        =   24
         Top             =   9960
         Width           =   7395
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
         Left            =   240
         TabIndex        =   23
         Top             =   9960
         Width           =   1155
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If you press ESC while recording this stops."
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
         Left            =   7680
         TabIndex        =   22
         Top             =   9600
         Width           =   4215
      End
      Begin VB.Image Image9 
         Height          =   210
         Left            =   7680
         Picture         =   "frmHelp3.frx":0DD8
         Top             =   9360
         Width           =   1980
      End
      Begin VB.Image Image8 
         Height          =   285
         Left            =   7680
         Picture         =   "frmHelp3.frx":23C2
         Top             =   9000
         Width           =   3510
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp3.frx":5844
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   7680
         TabIndex        =   21
         Top             =   7080
         Width           =   4350
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
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
         TabIndex        =   20
         Top             =   6600
         Width           =   870
      End
      Begin VB.Image Image7 
         Height          =   195
         Left            =   10440
         Picture         =   "frmHelp3.frx":594A
         Top             =   5880
         Width           =   240
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Click in APPLY button      to add your macro in the main task list window."
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
         Left            =   8280
         TabIndex        =   19
         Top             =   5880
         Width           =   4215
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   8040
         Top             =   6000
         Width           =   135
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Next you must introduce the task name in the box, the start time to play the macro and the date."
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
         Left            =   8280
         TabIndex        =   18
         Top             =   4920
         Width           =   4215
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   8040
         Top             =   5040
         Width           =   135
      End
      Begin VB.Image Image6 
         Height          =   270
         Left            =   9000
         Picture         =   "frmHelp3.frx":5BFC
         Top             =   4200
         Width           =   330
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp3.frx":6106
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
         Left            =   8280
         TabIndex        =   17
         Top             =   3600
         Width           =   4215
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   8040
         Top             =   3720
         Width           =   135
      End
      Begin VB.Image Image5 
         Height          =   225
         Left            =   9000
         Picture         =   "frmHelp3.frx":61A9
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp3.frx":64F7
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
         Left            =   8280
         TabIndex        =   16
         Top             =   2040
         Width           =   4215
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   8040
         Top             =   2160
         Width           =   135
      End
      Begin VB.Image Image4 
         Height          =   270
         Left            =   10440
         Picture         =   "frmHelp3.frx":65A6
         Top             =   960
         Width           =   270
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Press the REC button      and it will show a window that ask you how many seconds have your macro. Enter a numeric value."
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
         Left            =   8280
         TabIndex        =   15
         Top             =   960
         Width           =   4215
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   8040
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To record a new macro follow the next steps:"
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
         Left            =   7680
         TabIndex        =   14
         Top             =   600
         Width           =   4485
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "How to record and program?"
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
         TabIndex        =   13
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apply: add to main task list."
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
         Left            =   3840
         TabIndex        =   12
         Top             =   7920
         Width           =   2715
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New: create a new macro"
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
         Left            =   3840
         TabIndex        =   11
         Top             =   7440
         Width           =   2535
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open: select your valid txt coordinates file to play after."
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
         Left            =   480
         TabIndex        =   10
         Top             =   9360
         Width           =   5445
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save: Save the coordinates file."
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
         Left            =   480
         TabIndex        =   9
         Top             =   8400
         Width           =   3090
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save As: TimerLogic show a save window to save TXT coordinates file."
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
         Left            =   480
         TabIndex        =   8
         Top             =   8880
         Width           =   6975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Play: play the file recorded."
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
         Left            =   480
         TabIndex        =   7
         Top             =   7920
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rec: start to record"
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
         Left            =   480
         TabIndex        =   6
         Top             =   7440
         Width           =   1920
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   3600
         Top             =   8040
         Width           =   135
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   240
         Top             =   9480
         Width           =   135
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   3600
         Top             =   7560
         Width           =   135
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   240
         Top             =   9000
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   240
         Top             =   8520
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   240
         Top             =   8040
         Width           =   135
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   240
         Top             =   7560
         Width           =   135
      End
      Begin VB.Image Image3 
         Height          =   285
         Left            =   240
         Picture         =   "frmHelp3.frx":69D8
         Top             =   6960
         Width           =   5970
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "We can operate all with in top toolbar."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   6480
         Width           =   4815
      End
      Begin VB.Image Image2 
         Height          =   2925
         Left            =   240
         Picture         =   "frmHelp3.frx":C2DE
         Top             =   3360
         Width           =   6075
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmHelp3.frx":46160
         Top             =   120
         Width           =   480
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Macros"
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
         Left            =   840
         TabIndex        =   3
         Top             =   120
         Width           =   1110
      End
      Begin VB.Label l2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp3.frx":4646A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label l3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Macros task window:"
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
         Top             =   2760
         Width           =   3150
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12660
      _ExtentX        =   22331
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
Attribute VB_Name = "frmHelp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()

End Sub

Private Sub Form_Load()
    pic1.BackColor = RGB(255, 255, 217)
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
