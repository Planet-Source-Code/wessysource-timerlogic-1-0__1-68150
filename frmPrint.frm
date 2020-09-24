VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   315
      Left            =   3720
      TabIndex        =   7
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2760
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdHelp1 
      Caption         =   "Help"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Additional"
      Height          =   1695
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   1695
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmPrint.frx":0000
         Left            =   240
         List            =   "frmPrint.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cmbCopies 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Text            =   "1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Quality:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Copies:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Orientation"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Horizontal"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   960
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Vertical"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   240
         Picture         =   "frmPrint.frx":0024
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmPrint.frx":0766
         Top             =   840
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCopies_Change()
    If IsNumeric(cmbCopies) = False Then MsgBox "You must introduce a numeric value.", vbCritical, App.Title: Exit Sub
    If Val(cmbCopies) > 100 Or Val(cmbCopies) < 1 Then MsgBox "You must introduce a value between 1 and 100 inclusive.", vbCritical, App.Title: Exit Sub
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
    Printer.Copies = Val(cmbCopies)
    If Option1.Value = True Then
        Printer.Orientation = vbHorizontal
    Else
        Printer.Orientation = vbVertical
    End If
    If Combo1.List(Combo1.ListIndex) = "High" Then Printer.PrintQuality = 3
    If Combo1.List(Combo1.ListIndex) = "Normal" Then Printer.PrintQuality = 2
    If Combo1.List(Combo1.ListIndex) = "Low" Then Printer.PrintQuality = 1
    MousePointer = vbHourglass
    Wait 3
    Printer.FontBold = True
    Dim h
    
    Printer.Print "     TASK LIST DATED IN " & Date & " - " & App.Title & vbNewLine & vbNewLine
    With frmmain.LV
        For h = 1 To .ListItems.Count
            Printer.FontBold = True
            Printer.Print ("    * Task name: ")
            Printer.FontBold = False
            Printer.Print "     " & .ListItems(h).SubItems(1) 'task name
            Printer.FontBold = True
            Printer.Print ("    * Task: ")
            Printer.FontBold = False
            Printer.Print "     " & .ListItems(h).SubItems(2) ' task
            Printer.FontBold = True
            Printer.Print ("    * Start to: ")
            Printer.FontBold = False
            Printer.Print "     " & .ListItems(h).SubItems(3)
            Printer.FontBold = True
            Printer.Print ("    * In: ")
            Printer.FontBold = False
            Printer.Print "     " & .ListItems(h).SubItems(4)
            Printer.FontBold = True
            Printer.Print ("    * Status: ")
            Printer.FontBold = False
            Printer.Print "     " & .ListItems(h).SubItems(5)
            Printer.FontBold = True
            Printer.Print ("    * Priority: ")
            Printer.FontBold = False
            Printer.Print "     " & .ListItems(h).SubItems(6)
            Printer.FontBold = True
            Printer.Print ("    * Task type: ")
            Printer.FontBold = False
            Printer.Print "     " & .ListItems(h).SubItems(7)
            Printer.FontBold = True
            Printer.Print ("    * Comments / parameters: ")
            Printer.FontBold = False
            Printer.Print "     " & (.ListItems(h).SubItems(8))
            Printer.Print ("    " & String(100, "-"))
        Next
    End With
    Printer.EndDoc
    MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_Load()
    Option1.BackColor = BackColor
    Option2.BackColor = BackColor
    Dim i
    For i = 1 To 100
        cmbCopies.AddItem i
    Next i
End Sub
