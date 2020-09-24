Attribute VB_Name = "MouseEvents"

'API functions
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

'Mouse Event API constants
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type eMouseState
    Pos As POINTAPI
    LButton As Boolean
    MButton As Boolean
    RButton As Boolean
End Type


Option Explicit
Public Const SPS As Long = 50 'Recorded Samples Per Second

Public Cursor() As eMouseState
Public i As Long, j As Long, Samples As Long
Public pLB As Boolean, pMB As Boolean, pRB As Boolean
Public PlayOnly As Boolean, UnSaved As Boolean, Esc As Boolean
Public FN As String, RES As String, RT As String, HW As Boolean, MW As Boolean
Public Aux1 As String, Aux2 As String, Aux3 As String, Aux4 As String, Aux5 As String, Aux6 As String

Public Sub FreshForm()

    'Initialize Variables & Controls
    UnSaved = False: PlayOnly = False
    UpdateControls True, False, False, False
    frmOvr.Toolbar1.Buttons(13).Enabled = False
    UpdateInfo "", "", "", 0, True
    
    Debug.Print CurrentResolution
    'Setting Timers intervals to meet SPS requirement
    frmOvr.tmrRecord.Interval = 1000 / SPS
    frmOvr.tmrPlay.Interval = 1000 / SPS

End Sub

Public Sub UpdateInfo(ByVal iFN As String, ByVal iRES As String, ByVal iRT As String, ByVal iDU As Long, ByVal iHW As Boolean)

    Dim Info As String
    
    If iFN = "" Then iFN = "-"
    If iRES = "" Then iRES = "-"
    If iRT = "" Then iRT = "-"
    
    Info = "Macro   :  " + Left(iFN, 35) + vbCrLf
    Aux1 = Info
    Info = Info + "Recorded in :  " + iRT + vbCrLf
    Aux2 = Info
    Info = Info + "Resolution  :  " + iRES + vbCrLf
    Aux3 = Info
    Info = Info + "Duration    :  " + CStr(iDU) + " Sec" + vbCrLf
    Aux4 = Info
    Info = Info + "Hide Window :  " + CStr(iHW) + vbCrLf
    Aux5 = Info
    Info = Info + "File saved  :  " + CStr(Not UnSaved)
    Aux6 = Info
    
    frmOvr.lblInfo = Info

End Sub

Public Sub UpdateControls(cmdR As Boolean, cmdP As Boolean, tmrR As Boolean, tmrP As Boolean)

    With frmOvr
        
        If PlayOnly Then
            .Toolbar1.Buttons(1).Enabled = False
            .chkHide.Enabled = False
        Else
            .Toolbar1.Buttons(1).Enabled = cmdR
            .chkHide.Enabled = cmdR
        End If
        .Toolbar1.Buttons(3).Enabled = cmdP
        .tmrRecord.Enabled = tmrR
        .tmrPlay.Enabled = tmrP
        
    End With

End Sub

Public Function CurrentResolution() As String

    CurrentResolution = CStr(Screen.Width / Screen.TwipsPerPixelX) + " x " + CStr(Screen.Height / Screen.TwipsPerPixelY)

End Function

Public Sub SaveFile(ByVal FileName As String)
On Error GoTo Error

    Dim Count As Long, DateNow As String
    DateNow = CStr(Date)
    
    'Now Save Recorded Mouse Script
    Open FileName For Output Access Write Lock Write As #1
        Write #1, DateNow, HW, RES, j
        For Count = 1 To j
            Write #1, Cursor(Count).Pos.X, Cursor(Count).Pos.Y, Cursor(Count).LButton, Cursor(Count).MButton, Cursor(Count).RButton
            DoEvents
        Next
    Close #1
    
    UnSaved = False                                   'Now Mouse Recorder Script is Saved
    UpdateInfo FileName, RES, DateNow, j / SPS, HW    'Ubdate Info Label
    Exit Sub

Error:
Close #1
frmOvr.ComDlg.FileName = "": FN = ""
MsgBox "Error, Cannot Open file for Save!", vbCritical, "Error"
End Sub

Public Sub LoadFile(ByVal FileName As String)
On Error GoTo Error

    Dim Count As Long
    
    'Now Load Recorded Mouse Script
    Open FN For Input Access Read Lock Write As #1
        Input #1, RT, HW, RES, j
        If j <= 0 Then GoTo Error
        ReDim Cursor(j)
        For Count = 1 To j
            Input #1, Cursor(Count).Pos.X, Cursor(Count).Pos.Y, Cursor(Count).LButton, Cursor(Count).MButton, Cursor(Count).RButton
            DoEvents
        Next
    Close #1
    
    UpdateInfo FN, RES, RT, j / SPS, HW
    
    PlayOnly = True
    UnSaved = False
    If HW Then frmOvr.chkHide = 1 Else frmOvr.chkHide = 0
    UpdateControls False, True, False, False
    frmOvr.Toolbar1.Buttons(2).Enabled = True
    Exit Sub

Error:
Close #1
frmOvr.ComDlg.FileName = "": FN = ""
frmOvr.Toolbar1.Buttons(2).Enabled = True
'MsgBox "Error, Cannot Open file for PlayBack!", vbCritical, "Error"
End Sub

Public Sub Record()
    On Error Resume Next
    MouseEvents.GetCursorPos Cursor(i).Pos                                'Record Cursor Position
    Esc = CBool(GetAsyncKeyState(vbKeyEscape))                  'Monitor the Esc Key in Case User want to skip
    Cursor(i).LButton = CBool(GetAsyncKeyState(vbLeftButton))   'Left Button State
    Cursor(i).MButton = CBool(GetAsyncKeyState(vbMiddleButton)) 'Middle Button State
    Cursor(i).RButton = CBool(GetAsyncKeyState(vbRightButton))  'Right Button State
    
    'Prepare for next Position if Not finished yet Else Stop Recording
    If (i < Samples) And (Not Esc) Then
        i = i + 1
    Else
        j = i - 1
        UpdateControls True, True, False, False
        MsgBox "Record has finished.", vbInformation, "Succesfully"
        frmOvr.Toolbar1.Buttons(13).Enabled = True
        frmOvr.chkMinimize.Enabled = True
        RES = CurrentResolution
        UpdateInfo FN, RES, CStr(Date), j / SPS, HW
        If MW Then frmmain.WindowState = vbMaximized
        frmOvr.Show
    End If

End Sub

Public Sub Play()

    'Position Cursor where it should be
    SetCursorPos Cursor(i).Pos.X, Cursor(i).Pos.Y
    
    'ReGenerate Left Mouse Button Events
    If (Not pLB) And (Cursor(i).LButton) Then mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    If (pLB) And (Not Cursor(i).LButton) Then mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
    'ReGenerate Middle Mouse Button Events
    If (Not pMB) And (Cursor(i).MButton) Then mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
    If (pMB) And (Not Cursor(i).MButton) Then mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
    
    'ReGenerate Right Mouse Button Events
    If (Not pRB) And (Cursor(i).RButton) Then mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    If (pRB) And (Not Cursor(i).RButton) Then mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    
    'Monitor the Esc Key in Case User wants to skip
    Esc = CBool(GetAsyncKeyState(vbKeyEscape))
    
    'Prepare for next Position if Not finished yet Else Stop PlayBack
    If (i < j) And (Not Esc) Then
        pLB = Cursor(i).LButton     'Save previous LMB state
        pMB = Cursor(i).MButton     'Save previous MMB state
        pRB = Cursor(i).RButton     'Save previous RMB state
        i = i + 1                   'Next Sample
    Else
        UpdateControls True, True, False, False
        frmOvr.AutomaticTimer.Enabled = False
        If isFromMain = False Then
            MsgBox "The playing has finished!", vbInformation, "Succesfully"
            frmOvr.chkMinimize.Enabled = True
            frmOvr.Toolbar1.Buttons(13).Enabled = True
            If MW Then frmmain.WindowState = vbMaximized
            frmOvr.Show
        Else
            frmmain.Show
            Unload frmOvr
        End If
    End If

End Sub


