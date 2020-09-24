VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOvr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Macros"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   Icon            =   "frmOvr.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMinimize 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Minimize main window"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3720
      TabIndex        =   10
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtTaskName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ILAAA 
      Left            =   3000
      Top             =   1440
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
            Picture         =   "frmOvr.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOvr.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOvr.frx":09AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOvr.frx":0E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOvr.frx":1252
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOvr.frx":16A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOvr.frx":1AF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ILAAA"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrRecord 
      Left            =   0
      Top             =   2880
   End
   Begin VB.Timer tmrPlay 
      Left            =   360
      Top             =   2880
   End
   Begin VB.CheckBox chkHide 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hide window while recording (recommended)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.Timer AutomaticTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5040
      Top             =   2355
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   960
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.txt"
      Filter          =   "Mouse Recorder Text Files|*.txt"
      InitDir         =   "C:\"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3000
      TabIndex        =   8
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3000
      TabIndex        =   6
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Task name:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3060
      TabIndex        =   4
      Top             =   600
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   2760
      X2              =   2760
      Y1              =   480
      Y2              =   1680
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "To stop recording or playing, press ESC key"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   3090
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1515
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6015
   End
End
Attribute VB_Name = "frmOvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub AutomaticTimer_Timer()
    cmdPlay_Click
    AutomaticTimer.Enabled = False
End Sub

Private Sub Form_Initialize()

    InitCommonControls  'XP Style Support

End Sub

Private Sub Form_Load()
    chkHide.BackColor = BackColor
    chkMinimize.BackColor = BackColor
    lblInfo.BackColor = BackColor
    txtTaskName.BackColor = BackColor
    txtTime.BackColor = BackColor
    txtDate.BackColor = BackColor
    FreshForm           'Initiate Variables & Controls

End Sub

Private Sub cmdPlay_Click()
    On Error Resume Next
    'Check to see whether current screen resolution matches the recorded file resolution
    If RES <> CurrentResolution() Then
        If MsgBox("Your current screen resolution does not match the resolution of the file to be played back. Are you sure you want to Continue ?", vbCritical Or vbYesNo, "Resolution does not match") = vbNo Then Exit Sub
    End If
    
    i = 1                                    'Initiate PlayBack
    If j <= 0 Then Exit Sub                  'Abort PlayBack if nothing to play
    UpdateControls False, False, False, True 'Update Controls
    Toolbar1.Buttons(2).Enabled = False
    If HW Then Me.Hide                       'Should the Window Hide while PlayBack ?

End Sub

Private Sub mnuAbout_Click()

    frmAbout.Show vbModal

End Sub

Private Sub mnuFileExit_Click()

    Unload Me

End Sub

Private Sub mnuFileNew_Click()

    FreshForm

End Sub

Private Sub mnuFileOpen_Click()

    ComDlg.ShowOpen             'Show file Open Dialog
    FN = ComDlg.FileName        'Get the Chosen File Name
    If FN = "" Then Exit Sub    'Make Sure User have selected a file
    LoadFile FN                 'Now Load the file
    Toolbar1.Buttons(1).Enabled = True

End Sub

Private Sub mnuFileSave_Click()

    If FN = "" Then mnuFileSaveAs_Click: Exit Sub
    SaveFile FN

End Sub

Private Sub mnuFileSaveAs_Click()

    ComDlg.ShowSave             'Show Save As Dialog
    FN = ComDlg.FileName        'Get the Chosen File Name
    If FN = "" Then Exit Sub    'Make Sure User have selected a file
    SaveFile FN                 'Save to the selected file

End Sub

Private Sub tmrRecord_Timer()

    Record

End Sub

Private Sub tmrPlay_Timer()

    Play

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            On Error GoTo Error

    'Calculate total number of Samples to be Recorded
            Samples = SPS * Val(InputBox("Number of seconds to record :", "Don't change resolution while recording."))
    
            If Samples <= 0 Then Exit Sub            'Abort Recording if nothing to record
            i = 0                                    'Initiate Samples Counter
            ReDim Cursor(Samples)                    'Resize Cursor State Array
            UpdateControls False, False, True, False 'Update Controls
            UnSaved = True                           'Script is not saved yet
            frmOvr.chkMinimize.Enabled = False
            Toolbar1.Buttons(13).Enabled = False
            HW = CBool(chkHide.Value)                'Save the Option for applying at PlayBack
            MW = CBool(chkMinimize.Value)
            If HW Then Me.Hide                       'Should the Window Hide while Recording ?
            If MW Then frmmain.WindowState = vbMinimized
            Exit Sub

Error:
            MsgBox "Time too longer.", vbCritical, "Error"
            Exit Sub

        Case 3
            If RES <> CurrentResolution() Then
                If MsgBox("Your current screen resolution does not match the resolution of the file to be played back. Are you sure you want to Continue ?", vbCritical Or vbYesNo, "Resolution does not match") = vbNo Then Exit Sub
            End If
                
            i = 1                                    'Initiate PlayBack
            If j <= 0 Then Exit Sub                  'Abort PlayBack if nothing to play
            UpdateControls False, False, False, True 'Update Controls
            frmOvr.chkMinimize.Enabled = False
            Toolbar1.Buttons(13).Enabled = False
            If HW Then Me.Hide                       'Should the Window Hide while PlayBack ?

        Case 5
            If FN = "" Then Call SaveAsFile: Exit Sub
            SaveFile FN
        Case 7
            ComDlg.ShowSave             'Show Save As Dialog
            FN = ComDlg.FileName        'Get the Chosen File Name
            If FN = "" Then Exit Sub    'Make Sure User have selected a file
            SaveFile FN
        Case 9
            ComDlg.ShowOpen             'Show file Open Dialog
            FN = ComDlg.FileName        'Get the Chosen File Name
            If FN = "" Then Exit Sub    'Make Sure User have selected a file
            LoadFile FN                 'Now Load the file
            Toolbar1.Buttons(13).Enabled = True
        Case 11
            FreshForm
        Case 13
            Dim lstOvr As ListItem
            If Len(txtTaskName) = 0 Then MsgBox "You must introduce a task name.", vbExclamation, App.Title: txtTaskName.SetFocus: Exit Sub
            If Len(txtTime) = 0 Then MsgBox "You must introduce a start time.", vbExclamation, App.Title: txtTime.SetFocus: Exit Sub
            If Len(txtDate) = 0 Then MsgBox "You must introduce a start date.", vbExclamation, App.Title: txtDate.SetFocus: Exit Sub
            With frmmain.LV
                Set lstOvr = .ListItems.Add(, , , , 5)
                lstOvr.SubItems(1) = txtTaskName
                lstOvr.SubItems(2) = "Macro"
                lstOvr.SubItems(3) = txtTime
                lstOvr.SubItems(4) = txtDate
                lstOvr.SubItems(5) = "Waiting"
                lstOvr.SubItems(6) = "-"
                lstOvr.SubItems(7) = "Ovr"
                lstOvr.SubItems(8) = ComDlg.FileName
            End With
            
    End Select
        
End Sub
Private Sub SaveAsFile()
    ComDlg.ShowSave             'Show Save As Dialog
    FN = ComDlg.FileName        'Get the Chosen File Name
    If FN = "" Then Exit Sub    'Make Sure User have selected a file
    SaveFile FN
End Sub
