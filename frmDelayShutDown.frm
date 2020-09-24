VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmDelayShutDown 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alert prompt"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   Icon            =   "frmDelayShutDown.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timDelay 
      Interval        =   1000
      Left            =   2400
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   5400
      Picture         =   "frmDelayShutDown.frx":0442
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Your computer will be shuted down in secs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1110
      TabIndex        =   1
      Top             =   165
      Width           =   3705
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
      URL             =   "C:\Documents and Settings\HEAD\Mis documentos\VB\TaskList 2.0\Waiting to shutdown.AVI"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   10
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   -1  'True
      windowlessVideo =   -1  'True
      enabled         =   0   'False
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   873
   End
End
Attribute VB_Name = "frmDelayShutDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DisableCloseWindow Me.hwnd
    Label1 = "Your computer will be shutted down in " & secBeforeShutDown & " sec."
End Sub

Private Sub timDelay_Timer()
    Select Case IsAutomatic
        Case False
        
            Select Case frmmain.LV.selectedItem.SubItems(2)
                Case "Shut down"
                    secBeforeShutDown = secBeforeShutDown - 1
                    Label1 = "Your computer will be shutted down in " & secBeforeShutDown & " sec."
                    If secBeforeShutDown = 0 Then cExitWindows.ExitWindows WE_SHUTDOWN
                Case "Reboot"
                    secBeforeShutDown = secBeforeShutDown - 1
                    Label1 = "Your computer will be reboot in " & secBeforeShutDown & " sec."
                    If secBeforeShutDown = 0 Then cExitWindows.ExitWindows WE_REBOOT
                Case "Log Off"
                    secBeforeShutDown = secBeforeShutDown - 1
                    Label1 = "Your computer will log off session in " & secBeforeShutDown & " sec."
                    If secBeforeShutDown = 0 Then cExitWindows.ExitWindows WE_LOGOFF
            End Select
        Case True
            Select Case frmmain.LV.ListItems.Item(yourIndex).SubItems(2)
                Case "Shut down"
                    secBeforeShutDown = secBeforeShutDown - 1
                    Label1 = "Your computer will be shutted down in " & secBeforeShutDown & " sec."
                    If secBeforeShutDown = 0 Then cExitWindows.ExitWindows WE_SHUTDOWN
                Case "Reboot"
                    secBeforeShutDown = secBeforeShutDown - 1
                    Label1 = "Your computer will be reboot in " & secBeforeShutDown & " sec."
                    If secBeforeShutDown = 0 Then cExitWindows.ExitWindows WE_REBOOT
                Case "Log Off"
                    secBeforeShutDown = secBeforeShutDown - 1
                    Label1 = "Your computer will log off session in " & secBeforeShutDown & " sec."
                    If secBeforeShutDown = 0 Then cExitWindows.ExitWindows WE_LOGOFF
            End Select
    End Select
End Sub
