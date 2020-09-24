VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImpl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preset tasks"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   Icon            =   "frmImpl.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNow 
      Caption         =   "Now"
      Height          =   255
      Left            =   2520
      TabIndex        =   50
      Top             =   3360
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CDiag 
      Left            =   3240
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2880
      Top             =   2280
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parameters"
      Height          =   2025
      Left            =   3120
      TabIndex        =   11
      Top             =   3240
      Width           =   3735
      Begin VB.CheckBox chkDisabled 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Start disabled"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtPrinter 
         Appearance      =   0  'Flat
         Height          =   1035
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Text            =   "frmImpl.frx":0442
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txtPathNewFileToCreate 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   960
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox txtFileToCreateName 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2040
         TabIndex        =   44
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtPathFileToKill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton cmdSelFileToKill 
         Caption         =   "Select file"
         Height          =   375
         Left            =   960
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtDelayIcons 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDelayStartBut 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   37
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkDelayStartBut 
         Appearance      =   0  'Flat
         Caption         =   "Enable in                sec."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Timer timEnable 
         Enabled         =   0   'False
         Left            =   2760
         Top             =   480
      End
      Begin VB.TextBox txtWinTit 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton cmdTestBeep 
         Caption         =   "Test"
         Height          =   375
         Left            =   2520
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtFreq 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1140
         MaxLength       =   4
         TabIndex        =   31
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtDur 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1140
         MaxLength       =   5
         TabIndex        =   30
         Text            =   "0"
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox picPrior 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         ScaleHeight     =   225
         ScaleWidth      =   2145
         TabIndex        =   29
         Top             =   1320
         Width           =   2175
         Begin VB.Shape shpriority 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00800000&
            BorderStyle     =   0  'Transparent
            Height          =   255
            Left            =   0
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.TextBox txtPosY 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   27
         Text            =   "0000.0"
         Top             =   825
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPosX 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   26
         Text            =   "0000.0"
         Top             =   465
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtDelaySBar 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDelayTray 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSecsBeforeShutDown 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   255
         Left            =   1260
         TabIndex        =   18
         Top             =   930
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton optAlert 
         Appearance      =   0  'Flat
         Caption         =   "Alert before              sec."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton optNoPrompt 
         Appearance      =   0  'Flat
         Caption         =   "Without prompt"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtPriority 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   720
         MaxLength       =   3
         TabIndex        =   14
         Top             =   1320
         Width           =   495
      End
      Begin VB.CheckBox chkDelaySBar 
         Appearance      =   0  'Flat
         Caption         =   "Delay             sec. before show"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox chkDelayTray 
         Appearance      =   0  'Flat
         Caption         =   "Delay             sec. before close"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.OptionButton optPrompt 
         Appearance      =   0  'Flat
         Caption         =   "With prompt"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkDelayIcons 
         Appearance      =   0  'Flat
         Caption         =   "Show in               sec."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblChars 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3240
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lPathNewFileToCreate 
         AutoSize        =   -1  'True
         Caption         =   "Path of place the file:"
         Height          =   195
         Left            =   360
         TabIndex        =   45
         Top             =   720
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label lFileToCreateName 
         AutoSize        =   -1  'True
         Caption         =   "Name of file to create:"
         Height          =   195
         Left            =   360
         TabIndex        =   43
         Top             =   360
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label lWinTit 
         AutoSize        =   -1  'True
         Caption         =   "Window title:"
         Height          =   195
         Left            =   480
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lFreq 
         AutoSize        =   -1  'True
         Caption         =   "Frequency:                  Hz."
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   495
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label lDur 
         AutoSize        =   -1  'True
         Caption         =   "Duration:                   millisec."
         Height          =   195
         Left            =   360
         TabIndex        =   32
         Top             =   855
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Label lblPosActual 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         Height          =   195
         Left            =   2040
         TabIndex        =   28
         Top             =   675
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label ly 
         AutoSize        =   -1  'True
         Caption         =   "Position in Y:"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lx 
         AutoSize        =   -1  'True
         Caption         =   "Position in X:"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Priority:              %"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1260
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tasks"
      Height          =   3225
      Left            =   3105
      TabIndex        =   10
      Top             =   30
      Width           =   3735
      Begin MSComctlLib.ListView TL 
         Height          =   2895
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   5106
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Task"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Priority"
            Object.Width           =   1164
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Param"
            Object.Width           =   2523
         EndProperty
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calendar"
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2895
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2310
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         StartOfWeek     =   52887554
         CurrentDate     =   38891
      End
   End
   Begin VB.TextBox txtTaskName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Start date:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   915
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
      Left            =   120
      TabIndex        =   4
      Top             =   3360
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "frmImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TItem As ListItem

Private Sub chkDelayIcons_Click()
    If chkDelayIcons.Value = vbChecked Then
        txtDelayIcons.BackColor = vbWhite
        txtDelayIcons.Enabled = True
    Else
        txtDelayIcons.BackColor = BackColor
        txtDelayIcons.Enabled = False
    End If
End Sub

Private Sub chkDelaySBar_Click()
    If chkDelaySBar.Value = vbChecked Then
        txtDelaySBar.BackColor = vbWhite
        txtDelaySBar.Enabled = True
    Else
        txtDelaySBar.BackColor = BackColor
        txtDelaySBar.Enabled = False
    End If
End Sub

Private Sub chkDelayStartBut_Click()
    If chkDelayStartBut.Value = vbChecked Then
        txtDelayStartBut.BackColor = vbWhite
        txtDelayStartBut.Enabled = True
    Else
        txtDelayStartBut.BackColor = BackColor
        txtDelayStartBut.Enabled = False
    End If
End Sub

Private Sub chkDelayTray_Click()
    If chkDelayTray.Value = vbChecked Then
        txtDelayTray.BackColor = vbWhite
        txtDelayTray.Enabled = True
    Else
        txtDelayTray.BackColor = BackColor
        txtDelayTray.Enabled = False
    End If
End Sub

Private Sub cmdApply_Click()
    Dim LVTask As ListItem
    If Len(txtTaskName) = 0 Then MsgBox "You must introduce a task name.", vbExclamation, App.Title: txtTaskName.SetFocus: Exit Sub
    If Len(txtTime) = 0 Then
        MsgBox "You must introduce a start time.", vbExclamation, App.Title: txtTime.SetFocus: Exit Sub
    Else
        If Len(txtDate) = 0 Then
            MsgBox "You must introduce a start date.", vbExclamation, App.Title: txtDate.SetFocus: Exit Sub
        Else
            If Len(txtPriority) = 0 Then
                MsgBox "You must introduce a value as priority.", vbExclamation, App.Title: txtPriority.SetFocus: Exit Sub
            Else
                If Len(txtTime) < 8 Then
                    If txtTime = "Now" Or txtTime = "now" Then
                        GoTo Ok
                    Else
                        MsgBox "You must introduce a valid time.", vbExclamation, App.Title: txtTime.SetFocus: Exit Sub
                    End If
                End If
            End If
        End If
    End If
    
Ok: 'add to list
    If Not editImpl Then
        With frmmain
        Select Case TL.selectedItem.Text
            Case "Shut down"
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Reboot"
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Log Off"
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Open/close cd tray"
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Screensaver"
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Hide/show status bar"
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Set cursor pos"
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Beep"
                If chkDisabled.Value = vbUnchecked Then
                    If Val(txtFreq) = 0 Or Val(txtDur) = 0 Then MsgBox "You must introduce a valid frequency and duration.", vbExclamation, App.Title: txtFreq.SetFocus: Exit Sub
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    If Val(txtFreq) = 0 Or Val(txtDur) = 0 Then MsgBox "You must introduce a valid frequency and duration.", vbExclamation, App.Title: txtFreq.SetFocus: Exit Sub
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Show window"
                If Len(txtWinTit) = 0 Then MsgBox "You must enter a window title.", vbCritical, App.Title: Exit Sub
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Disable start button"
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Windows Update"
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Hide/show icons"
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Kill file"
                If Len(txtPathFileToKill) = 0 Then MsgBox "You must select the file.", vbCritical, App.Title: Exit Sub
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Create file"
                If Len(txtFileToCreateName) = 0 Then MsgBox "You must name your file.", vbCritical, App.Title: Exit Sub
                If Len(txtPathNewFileToCreate) = 0 Then MsgBox "You must introduce the file path.", vbCritical, App.Title: Exit Sub
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Remove trash content"
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    LVTask.SubItems(3) = txtTime
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Print text"
                If Len(txtPrinter) = 0 Then MsgBox "You must introduce your text in the box.", vbCritical, App.Title: Exit Sub
                If chkDisabled.Value = vbUnchecked Then
                    Set LVTask = .LV.ListItems.Add(, , "", , 2)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    If Left(txtTime, 1) = "0" Then
                        LVTask.SubItems(3) = Mid(txtTime, 2)
                    Else
                        LVTask.SubItems(3) = txtTime
                    End If
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Waiting"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    Set LVTask = .LV.ListItems.Add(, , "", , 4)
                    LVTask.SubItems(1) = txtTaskName
                    LVTask.SubItems(2) = TL.selectedItem.Text
                    If Left(txtTime, 1) = "0" Then
                        LVTask.SubItems(3) = Mid(txtTime, 2)
                    Else
                        LVTask.SubItems(3) = txtTime
                    End If
                    LVTask.SubItems(4) = txtDate
                    LVTask.SubItems(5) = "Disabled"
                    LVTask.SubItems(6) = txtPriority & " %"
                    LVTask.SubItems(7) = "Impl"
                    LVTask.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
        End Select
        .Toolbar1.Buttons(10).Enabled = True
        .Toolbar1.Buttons(11).Enabled = True
        .Toolbar1.Buttons(12).Enabled = True
                    
        .Toolbar2.Buttons(1).Enabled = False
        .Toolbar2.Buttons(2).Enabled = True
        .Toolbar2.Buttons(3).Enabled = True
        .Toolbar2.Buttons(4).Enabled = True
    End With
Else
    With frmmain
        Select Case TL.selectedItem.Text
            Case "Shut down"
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Reboot"
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Log Off"
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Open/close cd tray"
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Screensaver"
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Hide/show status bar"
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Set cursor pos"
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Beep"
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Show window"
                If Len(txtWinTit) = 0 Then MsgBox "You must enter a window title.", vbCritical, App.Title: Exit Sub
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Disable start button"
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Windows Update"
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Hide/show icons"
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Kill file"
                If Len(txtPathFileToKill) = 0 Then MsgBox "You must select the file.", vbCritical, App.Title: Exit Sub
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Create file"
                If Len(txtFileToCreateName) = 0 Then MsgBox "You must name your file.", vbCritical, App.Title: Exit Sub
                If Len(txtPathNewFileToCreate) = 0 Then MsgBox "You must introduce the file path.", vbCritical, App.Title: Exit Sub
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Remove trash content"
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
            Case "Print text"
                If Len(txtPrinter) = 0 Then MsgBox "You must introduce your text in the box.", vbCritical, App.Title: Exit Sub
                If chkDisabled.Value = vbUnchecked Then
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 2
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Waiting"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                Else
                    .LV.ListItems.Item(.LV.selectedItem.Index).SmallIcon = 4
                    .LV.selectedItem.SubItems(1) = txtTaskName
                    .LV.selectedItem.SubItems(2) = TL.selectedItem.Text
                    .LV.selectedItem.SubItems(3) = txtTime
                    .LV.selectedItem.SubItems(4) = txtDate
                    .LV.selectedItem.SubItems(5) = "Disabled"
                    .LV.selectedItem.SubItems(6) = txtPriority & " %"
                    .LV.selectedItem.SubItems(7) = "Impl"
                    .LV.selectedItem.SubItems(8) = TL.selectedItem.SubItems(2)
                End If
        End Select
        .Toolbar1.Buttons(10).Enabled = True
        .Toolbar1.Buttons(11).Enabled = True
        .Toolbar1.Buttons(12).Enabled = True
                    
        .Toolbar2.Buttons(1).Enabled = False
        .Toolbar2.Buttons(2).Enabled = True
        .Toolbar2.Buttons(3).Enabled = True
        .Toolbar2.Buttons(4).Enabled = True
    End With
End If
editImpl = False
Unload Me


End Sub

Private Sub cmdCancel_Click()
    editImpl = False
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Wait 1
    frmHelp1.Show 1
End Sub

Private Sub cmdNow_Click()
    txtTime = Time
End Sub

Private Sub cmdSelFileToKill_Click()
    CDiag.DefaultExt = "*.*"
    CDiag.DialogTitle = "Select any file..."
    CDiag.Filter = "All files (*.*)|*.*"
    CDiag.ShowOpen
    txtPathFileToKill = CDiag.FileName
End Sub

Private Sub cmdTestBeep_Click()
    If Val(txtDur) = 0 Or Val(txtFreq) = 0 Then MsgBox "You must introduce a value > 0.", vbExclamation, App.Title: Exit Sub
    cmdTestBeep.Enabled = False
    Beep Val(txtFreq), Val(txtDur)
    timEnable.Interval = txtDur
    timEnable.Enabled = True
End Sub

Private Sub Form_Load()
    Set TItem = TL.ListItems.Add(1, , "Shut down")
    Set TItem = TL.ListItems.Add(2, , "Reboot")
    Set TItem = TL.ListItems.Add(3, , "Log Off")
    Set TItem = TL.ListItems.Add(4, , "Open/close cd tray")
    Set TItem = TL.ListItems.Add(5, , "Screensaver")
    Set TItem = TL.ListItems.Add(6, , "Hide/show status bar")
    Set TItem = TL.ListItems.Add(7, , "Set cursor pos")
    Set TItem = TL.ListItems.Add(8, , "Beep")
    Set TItem = TL.ListItems.Add(9, , "Show window")
    Set TItem = TL.ListItems.Add(10, , "Disable start button")
    Set TItem = TL.ListItems.Add(11, , "Windows Update")
    Set TItem = TL.ListItems.Add(12, , "Hide/show icons")
    Set TItem = TL.ListItems.Add(13, , "Kill file")
    Set TItem = TL.ListItems.Add(14, , "Create file")
    Set TItem = TL.ListItems.Add(15, , "Remove trash content")
    Set TItem = TL.ListItems.Add(16, , "Print text")
    txtPathFileToKill.BackColor = BackColor
    picPrior.BackColor = BackColor
    chkDisabled.BackColor = BackColor
    shpriority.BackColor = RGB(1, 97, 234)
    
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
editImpl = False
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    If MonthView1.Day < 10 Then
        If MonthView1.Month < mvwOctober Then
            txtDate = "0" & MonthView1.Day & "/0" & MonthView1.Month & "/" & MonthView1.Year
        Else
            txtDate = "0" & MonthView1.Day & "/" & MonthView1.Month & "/" & MonthView1.Year
        End If
    Else
        If MonthView1.Month < mvwOctober Then
            txtDate = MonthView1.Day & "/0" & MonthView1.Month & "/" & MonthView1.Year
        Else
            txtDate = MonthView1.Day & "/" & MonthView1.Month & "/" & MonthView1.Year
        End If
    End If
End Sub

Private Sub optAlert_Click()
txtSecsBeforeShutDown.Enabled = True
txtSecsBeforeShutDown.BackColor = vbWhite
End Sub

Private Sub optNoPrompt_Click()
txtSecsBeforeShutDown.Enabled = False
txtSecsBeforeShutDown.BackColor = BackColor
TL.selectedItem.SubItems(2) = "No Prompt"
End Sub

Private Sub optPrompt_Click()
txtSecsBeforeShutDown.Enabled = False
txtSecsBeforeShutDown.BackColor = BackColor
TL.selectedItem.SubItems(2) = "Prompt"
End Sub

Private Sub Text1_Change()

End Sub

Private Sub timEnable_Timer()
    cmdTestBeep.Enabled = True
    timEnable.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Dim Result As Long
    
    Result = Module1.GetCursorPos(PosAPI)
    
    If Result <> 0 Then
        lblPosActual = "Actual = X:" & PosAPI.X & ", Y:" & PosAPI.Y
    Else
        lblPosActual = "-1"
    End If
End Sub

Private Sub TL_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdApply.Enabled = True
    Select Case TL.selectedItem.Text
        Case "Shut down"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = True
            optPrompt.Visible = True
            optAlert.Visible = True
            Me.txtSecsBeforeShutDown.Visible = True
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Reboot"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = True
            optPrompt.Visible = True
            optAlert.Visible = True
            Me.txtSecsBeforeShutDown.Visible = True
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Log Off"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = True
            optPrompt.Visible = True
            optAlert.Visible = True
            Me.txtSecsBeforeShutDown.Visible = True
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Open/close cd tray"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            chkDelayTray.Visible = True
            txtDelayTray.Visible = True
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Screensaver"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Hide/show status bar"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelaySBar.Visible = True
            txtDelaySBar.Visible = True
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            
        Case "Set cursor pos"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = True
            lx.Visible = True
            ly.Visible = True
            txtPosX.Visible = True
            txtPosY.Visible = True
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Beep"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = True
            lFreq.Visible = True
            lDur.Visible = True
            txtFreq.Visible = True
            txtDur.Visible = True
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Show window"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = True
            txtWinTit.Visible = True
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Disable start button"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = True
            txtDelayStartBut.Visible = True
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Windows Update"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Hide/show icons"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = True
            txtDelayIcons.Visible = True
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Kill file"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = True
            txtPathFileToKill.Visible = True
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Create file"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = True
            txtFileToCreateName.Visible = True
            lPathNewFileToCreate.Visible = True
            txtPathNewFileToCreate.Visible = True
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Remove trash content"
            lblChars.Visible = False
            txtPrinter.Visible = False
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
        Case "Print text"
            lblChars.Visible = True
            txtPrinter.Visible = True
            lFileToCreateName.Visible = False
            txtFileToCreateName.Visible = False
            lPathNewFileToCreate.Visible = False
            txtPathNewFileToCreate.Visible = False
            cmdSelFileToKill.Visible = False
            txtPathFileToKill.Visible = False
            chkDelayIcons.Visible = False
            txtDelayIcons.Visible = False
            chkDelayStartBut.Visible = False
            txtDelayStartBut.Visible = False
            lWinTit.Visible = False
            txtWinTit.Visible = False
            cmdTestBeep.Visible = False
            lFreq.Visible = False
            lDur.Visible = False
            txtFreq.Visible = False
            txtDur.Visible = False
            lblPosActual.Visible = False
            lx.Visible = False
            ly.Visible = False
            txtPosX.Visible = False
            txtPosY.Visible = False
            chkDelayTray.Visible = False
            txtDelayTray.Visible = False
            optNoPrompt.Visible = False
            optPrompt.Visible = False
            optAlert.Visible = False
            Me.txtSecsBeforeShutDown.Visible = False
            chkDelaySBar.Visible = False
            txtDelaySBar.Visible = False
    End Select
End Sub

Private Sub txtDelayIcons_Change()
    If Len(txtDelayIcons) = 0 Then Exit Sub
    If IsNumeric(txtDelayIcons) = False Then MsgBox "Invalid value as seconds.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(2) = "Show icons in " & txtDelayIcons & " sec."
End Sub

Private Sub txtDelaySBar_Change()
    If Len(txtDelaySBar) = 0 Then Exit Sub
    If IsNumeric(txtDelaySBar) = False Then MsgBox "Invalid value as seconds.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(2) = "Delay " & txtDelaySBar & " sec."
End Sub

Private Sub txtDelayStartBut_Change()
    If Len(txtDelayStartBut) = 0 Then Exit Sub
    If IsNumeric(txtDelayStartBut) = False Then MsgBox "Invalid value as seconds.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(2) = "Enable in " & txtDelayStartBut & " sec."
End Sub

Private Sub txtDelayTray_Change()
    If Len(txtDelayTray) = 0 Then Exit Sub
    If IsNumeric(txtDelayTray) = False Then MsgBox "Invalid value as seconds.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(2) = "Delay " & txtDelayTray & " sec."

End Sub

Private Sub txtFileToCreateName_Change()
    If Len(txtFileToCreateName) = 0 Then Exit Sub
    If IsNumeric(txtFileToCreateName) = True Then MsgBox "Invalid filename.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(2) = "FileName: " & txtFileToCreateName & ", Path: " & txtPathNewFileToCreate
End Sub

Private Sub txtFreq_Change()
    If Len(txtFreq) = 0 Then Exit Sub
    If IsNumeric(txtFreq) = False Then MsgBox "Invalid value as frequency.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(2) = "Freq: " & txtFreq & ", Dur: " & txtDur
End Sub
Private Sub txtDur_Change()
    If Len(txtDur) = 0 Then Exit Sub
    If IsNumeric(txtDur) = False Then MsgBox "Invalid value as duration.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(2) = "Freq: " & txtFreq & ", Dur: " & txtDur
End Sub

Private Sub txtPathFileToKill_Change()
    TL.selectedItem.SubItems(2) = CDiag.FileName
End Sub

Private Sub txtPathNewFileToCreate_Change()
    If Len(txtPathNewFileToCreate) = 0 Then Exit Sub
    If IsNumeric(txtPathNewFileToCreate) = True Then MsgBox "Invalid path.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(2) = "FileName: " & txtFileToCreateName & ", Path: " & txtPathNewFileToCreate
End Sub

Private Sub txtPosX_Change()
    If Len(txtPosX) = 0 Then Exit Sub
    If IsNumeric(txtPosX) = False Then MsgBox "Invalid value as mouse position.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(2) = "X = " & txtPosX & ", Y = " & txtPosY
End Sub
Private Sub txtPosY_Change()
    If Len(txtPosY) = 0 Then Exit Sub
    If IsNumeric(txtPosY) = False Then MsgBox "Invalid value as mouse position.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(2) = "X = " & txtPosX & ", Y = " & txtPosY
End Sub

Private Sub txtPrinter_Change()
    lblChars = Len(txtPrinter)
    TL.selectedItem.SubItems(2) = txtPrinter
End Sub

Private Sub txtPrinter_Click()
    If Left(txtPrinter, 5) = "Enter" Then
        txtPrinter = ""
    Else
        Exit Sub
    End If

End Sub

Private Sub txtPriority_Change()
    
    If Val(txtPriority) < 0 Or Val(txtPriority) > 100 Then MsgBox "A priority number too longer.", vbCritical, App.Title: Exit Sub
    If Len(txtPriority) = 0 Then Exit Sub
    If IsNumeric(txtPriority) = False Then MsgBox "Invalid value as priority.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(1) = txtPriority & " %"

   shpriority.Width = ((Val(txtPriority)) * picPrior.Width) / 100
End Sub

Private Sub txtSecsBeforeShutDown_Change()
    If Len(txtSecsBeforeShutDown) = 0 Then Exit Sub
    If IsNumeric(txtSecsBeforeShutDown) = False Then MsgBox "Invalid value as seconds.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(2) = "Delay " & txtSecsBeforeShutDown & " sec."
End Sub

Private Sub txtWinTit_Change()
    If Len(txtWinTit) = 0 Then Exit Sub
    If IsNumeric(txtWinTit) = True Then MsgBox "Invalid window title.", vbCritical, App.Title: Exit Sub
    TL.selectedItem.SubItems(2) = txtWinTit

End Sub
