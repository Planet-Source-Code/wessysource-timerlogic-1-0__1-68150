VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options panel"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7245
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Main list"
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   4935
      Begin VB.CheckBox chkGrid 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Grid lines"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkFull 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Full select"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkHot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "HotTracking"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkHover 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "HoverSelection"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Grid line: if you check, main list will have a grid lines to guide best the user vision."
         Height          =   375
         Left            =   1800
         TabIndex        =   24
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label10 
         Caption         =   "Full select: if you check, when click, is selected all row. Opposite only small icon."
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "HoverSelection: you can select items without click, only moving mouse on top."
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label8 
         Caption         =   "HotTracking: if you check is activated the complete pursuit between elements."
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Main timer"
      Height          =   2055
      Left            =   5160
      TabIndex        =   10
      Top             =   120
      Width           =   1935
      Begin VB.CheckBox chkSound 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Play sound"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alert without content"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "If list is empty TimerLogic show an alert prompt.    When run play sound."
         Height          =   675
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1755
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Run"
      Height          =   2055
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtWaitRun 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Wait:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Milliseconds"
         Height          =   195
         Left            =   1170
         TabIndex        =   8
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "This is the wait time before run a selected task."
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Run all"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtRunAllInterval 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "5000"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "This is the timer interval between the running of tasks"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Milliseconds"
         Height          =   195
         Left            =   1170
         TabIndex        =   3
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sequency interval:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    timRunAllInterval = Val(txtRunAllInterval)
    timWaitRun = Val(txtWaitRun)
    Set bbd = OpenDatabase(App.Path & "\DataBaseList.mdb")
    Set tbl = bbd.OpenRecordset("tblSettings")
       
    MousePointer = vbHourglass
                
    
    If Check1 Then
        showPrompt = True
    Else
        showPrompt = False
    End If
    If chkSound Then
        playSound = True
    Else
        playSound = False
    End If
    If chkHot Then
        hTracking = True
        frmmain.LV.HotTracking = True
    Else
        hTracking = False
        frmmain.LV.HotTracking = False
    End If
    If chkHover Then
        Hover = True
        frmmain.LV.HoverSelection = True
    Else
        Hover = False
        frmmain.LV.HoverSelection = False
    End If
    If chkFull Then
        fullSelect = True
        frmmain.LV.FullRowSelect = True
    Else
        fullSelect = False
        frmmain.LV.FullRowSelect = False
    End If
    If chkGrid Then
        showGrid = True
        frmmain.LV.GridLines = True
    Else
        showGrid = False
        frmmain.LV.GridLines = False
    End If
    tbl.AddNew
                
    tbl("Interval 1") = timRunAllInterval
    tbl("Interval 2") = timWaitRun
    tbl("Prompt") = showPrompt
    tbl("Sound") = playSound
    tbl("GridLines") = showGrid
    tbl("Full select") = fullSelect
    tbl("HoverSelection") = Hover
    tbl("HotTracking") = hTracking
    tbl.Update

    bbd.Close

    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Check1.BackColor = BackColor
    chkSound.BackColor = BackColor
    chkHot.BackColor = BackColor
    chkHover.BackColor = BackColor
    chkFull.BackColor = BackColor
    chkGrid.BackColor = BackColor
    txtRunAllInterval = timRunAllInterval
    txtWaitRun = timWaitRun
    If showPrompt = True Then
        Check1.Value = vbChecked
    Else
        Check1.Value = vbUnchecked
    End If
    If playSound = True Then
        chkSound.Value = vbChecked
    Else
        chkSound.Value = vbUnchecked
    End If
    If showGrid = True Then
        chkGrid.Value = vbChecked
    Else
        chkGrid.Value = vbUnchecked
    End If
    If fullSelect = True Then
        chkFull.Value = vbChecked
    Else
        chkFull.Value = vbUnchecked
    End If
    If Hover = True Then
        chkHover.Value = vbChecked
    Else
        chkHover.Value = vbUnchecked
    End If
    If hTracking = True Then
        chkHot.Value = vbChecked
    Else
        chkHot.Value = vbUnchecked
    End If
End Sub
