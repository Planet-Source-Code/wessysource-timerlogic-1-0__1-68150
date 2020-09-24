Attribute VB_Name = "Module1"
Global listitm As ListItem
Global lst As ListItem
Global timObject As Object

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public PosAPI As POINTAPI

Public Const MF_BYPOSITION = &H400&

Public Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" _
    (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
Public Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public WhatTASK As String, bindKey As Variant
Public InputFreq, InputDurt, InputMsg
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function Sleep Lib "kernel32" (ByVal dbMilliseconds As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Const NV_CLOSEMSGBOX As Long = &H5000&
Public yourIndex As Integer

Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public indexSelected As Integer, selIndexColor As Integer
Public isFromMain As Boolean
Public val_txt_delay As Integer, val_txt_delay2 As Integer
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Const SHERB_NOCONFIRMATION As Long = &H1    '1
Const SHERB_NOPROGRESSUI As Long = &H2      '2
Const SHERB_NOSOUND As Long = &H4           '4


Global bbd As Database
Global tbl As Recordset
Public SQL As String
Global timRunAllInterval As Integer
Global timWaitRun As Integer
Global showPrompt As Boolean
Global playSound As Boolean
Global hTracking As Boolean
Global Hover As Boolean
Global fullSelect As Boolean
Global showGrid As Boolean
Public editImpl As Boolean
Public editPzl As Boolean
Public editOvr As Boolean
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
Public cExitWindows As New clsExitWindows
Public EditAsApp As Boolean, EditAsFold As Boolean, EditAsImpl As Boolean, EditAsOvr As Boolean
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_SYSCOMMAND = &H112&
Public Const SC_SCREENSAVE = &HF140&

Public Task As String
Public secBeforeShutDown As Integer
Public IsAutomatic As Boolean

Public Sub DisableCloseWindow(lhWnd As Long)
    Dim hSystemMenu As Long
    hSystemMenu = GetSystemMenu(lhWnd, 0)
    Call RemoveMenu(hSystemMenu, 6, MF_BYPOSITION)
    Call RemoveMenu(hSystemMenu, 5, MF_BYPOSITION)
End Sub
Public Sub open_door()
    mciSendString "set cdaudio door open", 0, 0, 0
End Sub


Public Sub close_door()
    mciSendString "set cdaudio door closed", 0, 0, 0
End Sub
Public Sub Wait(nSeconds As Long)
    Sleep nSeconds * 1000
End Sub
Public Sub hide_taskbar()
rtn = FindWindow("Shell_traywnd", "") 'get the Window
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW) 'hide the Tasbar
End Sub
Public Sub show_taskbar()
'show th taskbar
rtn = FindWindow("Shell_traywnd", "") 'get the Window
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar
End Sub
Public Sub EnableStartButton(Optional Enabled As Boolean = True)
    'this will enable/disable any window wit
    '     h a little modifaction
    Dim lhWnd As Long 'declare variables
    'find start button hWnd
    lhWnd& = FindWindowEx(FindWindow("Shell_TrayWnd", ""), 0&, "Button", vbNullString)
    'call the enablewindow api and do the wh
    '     at needs to be done
    Call EnableWindow(lhWnd&, CLng(Enabled))
End Sub
Public Sub EmptyRec(para As Long)
    Dim nRet As Long
    nRet = SHEmptyRecycleBin(frmmain.hwnd, vbNullString, para)
End Sub

Public Sub Run(TypeOfTask As String, taskList As ListView, taskListItm As Integer, isAll As Boolean, Sound As Boolean)
Dim lstEvent As ListItem
IsAutomatic = True
If isAll = False Then
    Select Case TypeOfTask
        Case "Impl"
            Select Case taskList.ListItems.Item(taskListItm).SubItems(2) 'task
                Case "Shut down"
                    Select Case Left(taskList.ListItems.Item(taskListItm).SubItems(8), 3)
                        Case "Pro" 'with prompt
                            If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                            
                            If MsgBox("Do you want to shut down computer?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                                cExitWindows.ExitWindows WE_SHUTDOWN
                            Else
                            End If
                            Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                            lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                            lstEvent.SubItems(2) = "Runned"
                            lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                        Case "No " 'with no prompt
                            If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                            Wait 2
                            cExitWindows.ExitWindows WE_SHUTDOWN
                            Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                            lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                            lstEvent.SubItems(2) = "Runned"
                            lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                        Case "Del" 'delay
                            If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                            Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                            lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                            lstEvent.SubItems(2) = "Runned"
                            lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                            yourIndex = taskListItm
                            secBeforeShutDown = Val(Mid(taskList.ListItems.Item(taskListItm).SubItems(8), 7))
                            frmDelayShutDown.Show 1
                    End Select
                Case "Reboot"
                    Select Case Left(taskList.ListItems.Item(taskListItm).SubItems(8), 3)
                        Case "Pro" 'with prompt
                            If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                            If MsgBox("Do you want to reboot your computer?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                            cExitWindows.ExitWindows WE_REBOOT
                            Else
                            End If
                            Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                            lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                            lstEvent.SubItems(2) = "Runned"
                            lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                        Case "No " 'with no prompt
                            If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                            Wait 2
                            cExitWindows.ExitWindows WE_REBOOT
                            Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                            lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                            lstEvent.SubItems(2) = "Runned"
                            lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                        Case "Del" 'delay
                            If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                            Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                            lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                            lstEvent.SubItems(2) = "Runned"
                            lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                            yourIndex = taskListItm
                            secBeforeShutDown = Val(Mid(taskList.ListItems.Item(taskListItm).SubItems(8), 7))
                            frmDelayShutDown.Show 1
                    End Select
                Case "Log Off"
                    Select Case Left(taskList.ListItems.Item(taskListItm).SubItems(8), 3)
                        Case "Pro" 'with prompt
                            If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                            If MsgBox("Do you want to shut down computer?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                            cExitWindows.ExitWindows WE_LOGOFF
                            Else
                            End If
                            Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                            lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                            lstEvent.SubItems(2) = "Runned"
                            lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                        Case "No " 'with no prompt
                            If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                            Wait 2
                            cExitWindows.ExitWindows WE_LOGOFF
                            Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                            lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                            lstEvent.SubItems(2) = "Runned"
                            lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                        Case "Del" 'delay
                            If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                            Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                            lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                            lstEvent.SubItems(2) = "Runned"
                            lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                            yourIndex = taskListItm
                            secBeforeShutDown = Val(Mid(taskList.ListItems.Item(taskListItm).SubItems(8), 7))
                            frmDelayShutDown.Show 1
                    End Select
                Case "Open/close cd tray"
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    open_door
                    yourIndex = taskListItm
                    If Len(taskList.ListItems.Item(taskListItm).SubItems(8)) > 0 Then
                        frmmain.timToCloseTray.Enabled = True
                    End If
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "Screensaver"
                    Dim lResult As Long
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    lResult = SendMessage(frmmain.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "Hide/show status bar"
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    hide_taskbar
                    yourIndex = taskListItm
                    If Len(taskList.ListItems.Item(taskListItm).SubItems(8)) > 0 Then
                        frmmain.timToShowBar.Enabled = True
                    End If
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "Set cursor pos"
                    Dim posX As Double, posY As Double
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                            
                    posX = Mid(taskList.ListItems.Item(taskListItm).SubItems(8), 5, InStr(5, taskList.ListItems.Item(taskListItm).SubItems(8), ",") - 5)
                    posY = Mid(taskList.ListItems.Item(taskListItm).SubItems(8), InStr(5, taskList.ListItems.Item(taskListItm).SubItems(8), ",") + 6)
                    
                    SetCursorPos posX, posY
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "Beep"
                    Dim Freq, Dur
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    Freq = Mid(taskList.ListItems.Item(taskListItm).SubItems(8), 7, InStr(7, taskList.ListItems.Item(taskListItm).SubItems(8), ",") - 7)
                    Dur = Mid(taskList.ListItems.Item(taskListItm).SubItems(8), InStr(7, taskList.ListItems.Item(taskListItm).SubItems(8), ",") + 7)
                    
                    Beep Freq, Dur
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "Show window"
                    Dim Window_Handle As Long
                    Dim subItemWindowTitle
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    subItemWindowTitle = taskList.ListItems.Item(taskListItm).SubItems(8)
                    Window_Handle = FindWindow(vbNullString, subItemWindowTitle)
                    If Window_Handle Then
                        ShowWindow Window_Handle, 3
                    Else: MsgBox "Window title not found opened.", vbCritical, App.Title
                    End If
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "Disable start button"
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    yourIndex = taskListItm
                    EnableStartButton False
                    If Len(taskList.ListItems.Item(taskListItm).SubItems(8)) > 0 Then
                        frmmain.timToEnableStart.Enabled = True
                    End If
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "Windows Update"
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    ShellExecute hwnd, "open", "wupdmgr", "", "", 1
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "Hide/show icons"
                    Dim hwnda As Long
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    yourIndex = taskListItm
                    hwnda = FindWindowEx(0&, 0&, "Progman", vbNullString)
                    ShowWindow hwnda, 0
                    If Len(taskList.ListItems.Item(taskListItm).SubItems(8)) > 0 Then
                        frmmain.timToShowIcons.Enabled = True
                    End If
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "Kill file"
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    Kill taskList.ListItems.Item(taskListItm).SubItems(8)
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "Create file"
                    Dim inFileName, outPath
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    inFileName = Mid(taskList.ListItems.Item(taskListItm).SubItems(8), 11, InStr(11, taskList.ListItems.Item(taskListItm).SubItems(8), "Path:") - 13)
                    outPath = Mid(taskList.ListItems.Item(taskListItm).SubItems(8), InStr(11, taskList.ListItems.Item(taskListItm).SubItems(8), "Path:") + 6)
                    
                    Open outPath & "/" & inFileName For Output As 1#
                    Close 1#
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "Remove trash content"
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    Call EmptyRec(7)
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "Print text"
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    Printer.FontItalic = True
                    Printer.Print "AUTO PRINT AT: " & taskList.ListItems.Item(taskListItm).SubItems(3) & " of " & taskList.ListItems.Item(taskListItm).SubItems(4) & vbCrLf
                    Printer.FontItalic = False
                    Printer.Print taskList.ListItems.Item(taskListItm).SubItems(8)
                    Printer.EndDoc
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
            End Select
        Case "Pzl"
            Select Case Left(taskList.ListItems.Item(taskListItm).SubItems(2), 1)
                Case "R"
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    ShellExecute frmmain.hwnd, "open", Mid$(taskList.ListItems.Item(taskListItm).SubItems(2), 6), "", "", 1
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned"
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
                Case "O"
                    If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
                    If Mid$(taskList.ListItems.Item(taskListItm).SubItems(2), 5, 1) = "F" Then
                        ShellExecute frmmain.hwnd, "open", Mid$(taskList.ListItems.Item(taskListItm).SubItems(2), 11), "", "", 1
                    Else
                        ShellExecute frmmain.hwnd, "open", Mid$(taskList.ListItems.Item(taskListItm).SubItems(2), 10), "", "", 1
                    End If
                    Set lstEvent = frmmain.LV2.ListItems.Add(, , taskList.ListItems.Item(taskListItm).SubItems(1), , 7)
                    lstEvent.SubItems(1) = taskList.ListItems.Item(taskListItm).SubItems(3)
                    lstEvent.SubItems(2) = "Runned" 'status
                    lstEvent.SubItems(3) = taskList.ListItems.Item(taskListItm).SubItems(6)
            End Select
        Case "Ovr"
            If Sound Then mciExecute "Play C:\WINDOWS\Media\sound.wav"
            isFromMain = True
            Load frmOvr
            FN = taskList.ListItems.Item(taskListItm).SubItems(8)
            LoadFile FN
            With frmOvr
                .AutomaticTimer.Enabled = True
                .Show 1
            End With
    End Select
    taskList.ListItems(taskListItm).SubItems(5) = "Runned"
    taskList.ListItems(taskListItm).SmallIcon = 6
Else
    Select Case TypeOfTask
        Case "Impl"
            Select Case taskList.ListItems.Item(taskListItm).SubItems(2) 'task
                Case "Shut down"
                    Select Case Left(taskList.ListItems.Item(taskListItm).SubItems(8), 3)
                        Case "Pro" 'with prompt
                            If MsgBox("Do you want to shut down computer?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                                cExitWindows.ExitWindows WE_SHUTDOWN
                            Else
                            End If
                        Case "No " 'with no prompt
                            Wait 2
                            cExitWindows.ExitWindows WE_SHUTDOWN
                        Case "Del" 'delay
                            yourIndex = taskListItm
                            secBeforeShutDown = Val(Mid(taskList.ListItems.Item(taskListItm).SubItems(8), 7))
                            frmDelayShutDown.Show 1
                    End Select
                Case "Reboot"
                    Select Case Left(taskList.ListItems.Item(taskListItm).SubItems(8), 3)
                        Case "Pro" 'with prompt
                            If MsgBox("Do you want to reboot your computer?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                                cExitWindows.ExitWindows WE_REBOOT
                            Else
                            End If
                        Case "No " 'with no prompt
                            Wait 2
                            cExitWindows.ExitWindows WE_REBOOT
                        Case "Del" 'delay
                            yourIndex = taskListItm
                            secBeforeShutDown = Val(Mid(taskList.ListItems.Item(taskListItm).SubItems(8), 7))
                            frmDelayShutDown.Show 1
                    End Select
                Case "Log Off"
                    Select Case Left(taskList.ListItems.Item(taskListItm).SubItems(8), 3)
                        Case "Pro" 'with prompt
                            If MsgBox("Do you want to shut down computer?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                                cExitWindows.ExitWindows WE_LOGOFF
                            Else
                            End If
                        Case "No " 'with no prompt
                            Wait 2
                            cExitWindows.ExitWindows WE_LOGOFF
                        Case "Del" 'delay
                            yourIndex = taskListItm
                            secBeforeShutDown = Val(Mid(taskList.ListItems.Item(taskListItm).SubItems(8), 7))
                            frmDelayShutDown.Show 1
                    End Select
                Case "Open/close cd tray"
                    open_door
                    yourIndex = taskListItm
                    If Len(taskList.ListItems.Item(taskListItm).SubItems(8)) > 0 Then
                        frmmain.timToCloseTray.Enabled = True
                    End If
                Case "Screensaver"
                    Dim lResult2 As Long
                    lResult2 = SendMessage(frmmain.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
                Case "Hide/show status bar"
                    hide_taskbar
                    yourIndex = taskListItm
                    If Len(taskList.ListItems.Item(taskListItm).SubItems(8)) > 0 Then
                        frmmain.timToShowBar.Enabled = True
                    End If
                Case "Set cursor pos"
                    Dim posX2 As Double, posY2 As Double
                            
                    posX2 = Mid(taskList.ListItems.Item(taskListItm).SubItems(8), 5, InStr(5, taskList.ListItems.Item(taskListItm).SubItems(8), ",") - 5)
                    posY2 = Mid(taskList.ListItems.Item(taskListItm).SubItems(8), InStr(5, taskList.ListItems.Item(taskListItm).SubItems(8), ",") + 6)
                    
                    SetCursorPos posX2, posY2
                Case "Beep"
                    Dim Freq2, Dur2
                    Freq2 = Mid(taskList.ListItems.Item(taskListItm).SubItems(8), 7, InStr(7, taskList.ListItems.Item(taskListItm).SubItems(8), ",") - 7)
                    Dur2 = Mid(taskList.ListItems.Item(taskListItm).SubItems(8), InStr(7, taskList.ListItems.Item(taskListItm).SubItems(8), ",") + 7)
                    
                    Beep Freq2, Dur2
                Case "Show window"
                    Dim Window_Handle2 As Long
                    Dim subItemWindowTitle2
                    subItemWindowTitle2 = taskList.ListItems.Item(taskListItm).SubItems(8)
                    Window_Handle2 = FindWindow(vbNullString, subItemWindowTitle2)
                    If Window_Handle2 Then
                        ShowWindow Window_Handle2, 3
                    Else: MsgBox "Window title not found opened.", vbCritical, App.Title
                    End If
                Case "Disable start button"
                    yourIndex = taskListItm
                    EnableStartButton False
                    If Len(taskList.ListItems.Item(taskListItm).SubItems(8)) > 0 Then
                        frmmain.timToEnableStart.Enabled = True
                    End If
                Case "Windows Update"
                    ShellExecute hwnd, "open", "wupdmgr", "", "", 1
                Case "Hide/show icons"
                    Dim hwnda2 As Long
                    yourIndex = taskListItm
                    hwnda2 = FindWindowEx(0&, 0&, "Progman", vbNullString)
                    ShowWindow hwnda2, 0
                    If Len(taskList.ListItems.Item(taskListItm).SubItems(8)) > 0 Then
                        frmmain.timToShowIcons.Enabled = True
                    End If
                Case "Kill file"
                    Kill taskList.ListItems.Item(taskListItm).SubItems(8)
                Case "Create file"
                    Dim inFileName2, outPath2
                    inFileName2 = Mid(taskList.ListItems.Item(taskListItm).SubItems(8), 11, InStr(11, taskList.ListItems.Item(taskListItm).SubItems(8), "Path:") - 13)
                    outPath2 = Mid(taskList.ListItems.Item(taskListItm).SubItems(8), InStr(11, taskList.ListItems.Item(taskListItm).SubItems(8), "Path:") + 6)
                    
                    Open outPath2 & "/" & inFileName2 For Output As 1#
                    Close 1#
                Case "Remove trash content"
                    Call EmptyRec(7)
                Case "Print text"
                    Printer.FontItalic = True
                    Printer.Print "AUTO PRINT AT: " & taskList.ListItems.Item(taskListItm).SubItems(3) & " of " & taskList.ListItems.Item(taskListItm).SubItems(4) & vbCrLf
                    Printer.FontItalic = False
                    Printer.Print taskList.ListItems.Item(taskListItm).SubItems(8)
                    Printer.EndDoc
            End Select
        Case "Pzl"
            Select Case Left(taskList.ListItems.Item(taskListItm).SubItems(2), 1)
                Case "R"
                    ShellExecute frmmain.hwnd, "open", Mid$(taskList.ListItems.Item(taskListItm).SubItems(2), 6), "", "", 1
                Case "O"
                    If Mid$(taskList.ListItems.Item(taskListItm).SubItems(2), 5, 1) = "F" Then
                        ShellExecute frmmain.hwnd, "open", Mid$(taskList.ListItems.Item(taskListItm).SubItems(2), 11), "", "", 1
                    Else
                        ShellExecute frmmain.hwnd, "open", Mid$(taskList.ListItems.Item(taskListItm).SubItems(2), 10), "", "", 1
                    End If
            End Select
        Case "Ovr"
            isFromMain = True
            Load frmOvr
            FN = taskList.ListItems.Item(taskListItm).SubItems(8)
            LoadFile FN
            With frmOvr
                .AutomaticTimer.Enabled = True
                .Show
            End With
    End Select
    taskList.ListItems.Item(taskListItm).Selected = True
End If
End Sub
Public Sub RunAll(TypeTask As String, taskList As ListView, indexCounter As Integer)
    TypeTask = taskList.ListItems.Item(indexCounter).SubItems(7)
    Run TypeTask, taskList, indexCounter, True, False
    taskList.ListItems.Item(indexCounter).Selected = True
    If indexCounter >= taskList.ListItems.Count Then
        frmmain.Toolbar1.Buttons(6).Value = tbrUnpressed
        If frmmain.Toolbar1.Buttons(4).Value = tbrPressed Then
            
            frmmain.StatusBar1.Panels(3).Picture = LoadPicture(App.Path & "/timenabled.bmp")
            frmmain.StatusBar1.Panels(3).Text = "Timer status: enabled"
        Else
            frmmain.StatusBar1.Panels(3).Picture = LoadPicture(App.Path & "/timdisabled.bmp")
            frmmain.StatusBar1.Panels(3).Text = "Timer status: disabled"
        End If
        frmmain.timShowRunAllAnimation.Enabled = False
        
    frmmain.RunAllTimer.Enabled = False
    End If
End Sub
Public Sub Search(taskList As ListView, SearchBy As String, textToSearch As String)
Dim i
    Select Case SearchBy
        Case "Task name"
            For i = 1 To taskList.ListItems.Count
                If taskList.ListItems.Item(i).SubItems(1) = textToSearch Then
                    With frmSearch
                        .la(0).Visible = True
                        .la(1).Visible = True
                        .la(2).Visible = True
                        .la(3).Visible = True
                        .la(4).Visible = True
                        .lblNoResults.Visible = False
                        .lTaskName = taskList.ListItems.Item(i).SubItems(1)
                        .lTask = taskList.ListItems.Item(i).SubItems(2)
                        .lTime = taskList.ListItems.Item(i).SubItems(3)
                        .lDate = taskList.ListItems.Item(i).SubItems(4)
                        .lItem = taskList.ListItems.Item(i).Index
                    End With
                End If
            Next
        Case "Task"
            For i = 1 To taskList.ListItems.Count
                If taskList.ListItems.Item(i).SubItems(2) = textToSearch Then
                    With frmSearch
                        .la(0).Visible = True
                        .la(1).Visible = True
                        .la(2).Visible = True
                        .la(3).Visible = True
                        .la(4).Visible = True
                        .lblNoResults.Visible = False
                        .lTaskName = taskList.ListItems.Item(i).SubItems(1)
                        .lTask = taskList.ListItems.Item(i).SubItems(2)
                        .lTime = taskList.ListItems.Item(i).SubItems(3)
                        .lDate = taskList.ListItems.Item(i).SubItems(4)
                        .lItem = taskList.ListItems.Item(i).Index
                    End With
                End If
            Next
        Case "Time"
            For i = 1 To taskList.ListItems.Count
                If taskList.ListItems.Item(i).SubItems(3) = textToSearch Then
                    With frmSearch
                        .la(0).Visible = True
                        .la(1).Visible = True
                        .la(2).Visible = True
                        .la(3).Visible = True
                        .la(4).Visible = True
                        .lblNoResults.Visible = False
                        .lTaskName = taskList.ListItems.Item(i).SubItems(1)
                        .lTask = taskList.ListItems.Item(i).SubItems(2)
                        .lTime = taskList.ListItems.Item(i).SubItems(3)
                        .lDate = taskList.ListItems.Item(i).SubItems(4)
                        .lItem = taskList.ListItems.Item(i).Index
                    End With
                End If
            Next
        Case "Date"
            For i = 1 To taskList.ListItems.Count
                If taskList.ListItems.Item(i).SubItems(4) = textToSearch Then
                    With frmSearch
                        .la(0).Visible = True
                        .la(1).Visible = True
                        .la(2).Visible = True
                        .la(3).Visible = True
                        .la(4).Visible = True
                        .lblNoResults.Visible = False
                        .lTaskName = taskList.ListItems.Item(i).SubItems(1)
                        .lTask = taskList.ListItems.Item(i).SubItems(2)
                        .lTime = taskList.ListItems.Item(i).SubItems(3)
                        .lDate = taskList.ListItems.Item(i).SubItems(4)
                        .lItem = taskList.ListItems.Item(i).Index
                    End With
                End If
            Next
    End Select
End Sub
Public Sub Edit(TypeTask As String)
    With frmmain
        Select Case TypeTask
            Case "Impl"
                Load frmImpl
                frmImpl.txtTaskName = .LV.selectedItem.SubItems(1)
                frmImpl.txtTime = .LV.selectedItem.SubItems(3)
                frmImpl.txtDate = .LV.selectedItem.SubItems(4)
                frmImpl.txtPriority = .LV.selectedItem.SubItems(6)
                If .LV.selectedItem.SubItems(5) = "Disabled" Then
                    frmImpl.chkDisabled.Value = vbChecked
                Else
                    frmImpl.chkDisabled.Value = vbUnchecked
                End If
                frmImpl.Caption = "Preset tasks - Edit mode"
                frmImpl.Show 1
            Case "Pzl"
                Load frmPzl
                Select Case Left(.LV.selectedItem.SubItems(2), 1)
                    Case "R"
                        frmPzl.txtTaskName1 = .LV.selectedItem.SubItems(1)
                        frmPzl.txtTime1 = .LV.selectedItem.SubItems(3)
                        frmPzl.txtDate1 = .LV.selectedItem.SubItems(4)
                        frmPzl.txtPriority1 = .LV.selectedItem.SubItems(6)
                        If .LV.selectedItem.SubItems(5) = "Disabled" Then
                            frmPzl.chkDisabled1.Value = vbChecked
                        Else
                            frmPzl.chkDisabled1.Value = vbUnchecked
                        End If
                        frmPzl.txtExePath = Mid(.LV.selectedItem.SubItems(2), 6)
                        frmPzl.txtComments1 = .LV.selectedItem.SubItems(8)
                        frmPzl.Caption = "Applications & folders - Edit mode"
                        frmPzl.sTab.Tab = 0
                        frmPzl.Show 1
                    Case "O"
                        If Mid$(.LV.selectedItem.SubItems(2), 5, 1) = "F" Then
                            frmPzl.txtTaskName2 = .LV.selectedItem.SubItems(1)
                            frmPzl.txtTime2 = .LV.selectedItem.SubItems(3)
                            frmPzl.txtDate2 = .LV.selectedItem.SubItems(4)
                            frmPzl.txtPriority2 = .LV.selectedItem.SubItems(6)
                            If .LV.selectedItem.SubItems(5) = "Disabled" Then
                                frmPzl.chkDisabled2.Value = vbChecked
                            Else
                                frmPzl.chkDisabled2.Value = vbUnchecked
                            End If
                            frmPzl.txtFolderPath = Mid(.LV.selectedItem.SubItems(2), 11)
                            frmPzl.txtComments2 = .LV.selectedItem.SubItems(8)
                            frmPzl.Caption = "Applications & folders - Edit mode"
                            frmPzl.sTab.Tab = 1
                            frmPzl.Show 1
                        Else
                            frmPzl.txtTaskName3 = .LV.selectedItem.SubItems(1)
                            frmPzl.txtTime3 = .LV.selectedItem.SubItems(3)
                            frmPzl.txtDate3 = .LV.selectedItem.SubItems(4)
                            frmPzl.txtPriority3 = .LV.selectedItem.SubItems(6)
                            If .LV.selectedItem.SubItems(5) = "Disabled" Then
                                frmPzl.chkDisabled3.Value = vbChecked
                            Else
                                frmPzl.chkDisabled3.Value = vbUnchecked
                            End If
                            frmPzl.txtWebPage = Mid(.LV.selectedItem.SubItems(2), 10)
                            frmPzl.txtComments3 = .LV.selectedItem.SubItems(8)
                            frmPzl.Caption = "Applications & folders - Edit mode"
                            frmPzl.sTab.Tab = 2
                            frmPzl.Show 1
                        End If
                End Select
            Case "Ovr"
        End Select
    End With
End Sub
