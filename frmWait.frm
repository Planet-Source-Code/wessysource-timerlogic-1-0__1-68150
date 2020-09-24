VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wait"
   ClientHeight    =   540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2385
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   2385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   600
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmWait.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please wait..."
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   165
      Width           =   945
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Select Case Task
        Case "Impl"
            frmImpl.Show 1
            Unload frmWait
        Case "Pzl"
            frmPzl.Show 1
            Unload Me
        Case "Ovr"
            frmOvr.Show 1
            Unload Me
        Case "Ld"
            On Error GoTo showError
            MousePointer = vbHourglass
            Set bbd = OpenDatabase(App.Path & "/DataBaseList.mdb")
            SQL = "SELECT * FROM ListOfTasks"
            Set tbl = bbd.OpenRecordset(SQL)
            Sleep 1000
            tbl.MoveFirst
            
            Do Until tbl.EOF
                
                If tbl("Type") = "Impl" Then
                    If tbl("Status") = "Runned" Then
                        Set Itm = frmmain.LV.ListItems.Add(, , , , 6)
                        Itm.SubItems(1) = tbl("TaskName")
                        Itm.SubItems(2) = tbl("Task")
                        Itm.SubItems(3) = tbl("Time")
                        Itm.SubItems(4) = tbl("Date")
                        Itm.SubItems(5) = tbl("Status")
                        Itm.SubItems(6) = tbl("Priority")
                        Itm.SubItems(7) = tbl("Type")
                        Itm.SubItems(8) = tbl("Comments")
                    ElseIf tbl("Status") = "Disabled" Then
                        Set Itm = frmmain.LV.ListItems.Add(, , , , 4)
                        Itm.SubItems(1) = tbl("TaskName")
                        Itm.SubItems(2) = tbl("Task")
                        Itm.SubItems(3) = tbl("Time")
                        Itm.SubItems(4) = tbl("Date")
                        Itm.SubItems(5) = tbl("Status")
                        Itm.SubItems(6) = tbl("Priority")
                        Itm.SubItems(7) = tbl("Type")
                        Itm.SubItems(8) = tbl("Comments")
                    Else
                        Set Itm = frmmain.LV.ListItems.Add(, , , , 2)
                        Itm.SubItems(1) = tbl("TaskName")
                        Itm.SubItems(2) = tbl("Task")
                        Itm.SubItems(3) = tbl("Time")
                        Itm.SubItems(4) = tbl("Date")
                        Itm.SubItems(5) = tbl("Status")
                        Itm.SubItems(6) = tbl("Priority")
                        Itm.SubItems(7) = tbl("Type")
                        Itm.SubItems(8) = tbl("Comments")
                    End If
                    GoTo letsMove
                ElseIf tbl("Type") = "Pzl" Then
                    If tbl("Status") = "Runned" Then
                        Set Itm = frmmain.LV.ListItems.Add(, , , , 6)
                        Itm.SubItems(1) = tbl("TaskName")
                        Itm.SubItems(2) = tbl("Task")
                        Itm.SubItems(3) = tbl("Time")
                        Itm.SubItems(4) = tbl("Date")
                        Itm.SubItems(5) = tbl("Status")
                        Itm.SubItems(6) = tbl("Priority")
                        Itm.SubItems(7) = tbl("Type")
                        Itm.SubItems(8) = tbl("Comments")
                    ElseIf tbl("Status") = "Disabled" Then
                        Set Itm = frmmain.LV.ListItems.Add(, , , , 4)
                        Itm.SubItems(1) = tbl("TaskName")
                        Itm.SubItems(2) = tbl("Task")
                        Itm.SubItems(3) = tbl("Time")
                        Itm.SubItems(4) = tbl("Date")
                        Itm.SubItems(5) = tbl("Status")
                        Itm.SubItems(6) = tbl("Priority")
                        Itm.SubItems(7) = tbl("Type")
                        Itm.SubItems(8) = tbl("Comments")
                    Else
                        Set Itm = frmmain.LV.ListItems.Add(, , , , 3)
                        Itm.SubItems(1) = tbl("TaskName")
                        Itm.SubItems(2) = tbl("Task")
                        Itm.SubItems(3) = tbl("Time")
                        Itm.SubItems(4) = tbl("Date")
                        Itm.SubItems(5) = tbl("Status")
                        Itm.SubItems(6) = tbl("Priority")
                        Itm.SubItems(7) = tbl("Type")
                        Itm.SubItems(8) = tbl("Comments")
                    End If
                    GoTo letsMove
                ElseIf tbl("Type") = "Ovr" Then
                    If tbl("Status") = "Runned" Then
                        Set Itm = frmmain.LV.ListItems.Add(, , , , 6)
                        Itm.SubItems(1) = tbl("TaskName")
                        Itm.SubItems(2) = tbl("Task")
                        Itm.SubItems(3) = tbl("Time")
                        Itm.SubItems(4) = tbl("Date")
                        Itm.SubItems(5) = tbl("Status")
                        Itm.SubItems(6) = tbl("Priority")
                        Itm.SubItems(7) = tbl("Type")
                        Itm.SubItems(8) = tbl("Comments")
                    ElseIf tbl("Status") = "Disabled" Then
                        Set Itm = frmmain.LV.ListItems.Add(, , , , 4)
                        Itm.SubItems(1) = tbl("TaskName")
                        Itm.SubItems(2) = tbl("Task")
                        Itm.SubItems(3) = tbl("Time")
                        Itm.SubItems(4) = tbl("Date")
                        Itm.SubItems(5) = tbl("Status")
                        Itm.SubItems(6) = tbl("Priority")
                        Itm.SubItems(7) = tbl("Type")
                        Itm.SubItems(8) = tbl("Comments")
                    Else
                        Set Itm = frmmain.LV.ListItems.Add(, , , , 5)
                        Itm.SubItems(1) = tbl("TaskName")
                        Itm.SubItems(2) = tbl("Task")
                        Itm.SubItems(3) = tbl("Time")
                        Itm.SubItems(4) = tbl("Date")
                        Itm.SubItems(5) = tbl("Status")
                        Itm.SubItems(6) = tbl("Priority")
                        Itm.SubItems(7) = tbl("Type")
                        Itm.SubItems(8) = tbl("Comments")
                    End If
                End If
letsMove:       tbl.MoveNext
        Loop
            MousePointer = vbDefault
            frmmain.Toolbar1.Buttons(10).Enabled = True
            frmmain.Toolbar1.Buttons(11).Enabled = True
            frmmain.Toolbar1.Buttons(12).Enabled = True
            
            frmmain.Toolbar2.Buttons(1).Enabled = False
            frmmain.Toolbar2.Buttons(2).Enabled = True
            frmmain.Toolbar2.Buttons(3).Enabled = True
            frmmain.Toolbar2.Buttons(4).Enabled = True
            frmmain.i(3).Enabled = False
            frmmain.l(3).Enabled = False
            Unload Me
            Exit Sub
showError:
            If Err.Number = 3021 Then
                MsgBox "Data base is empty.", vbInformation, App.Title
                MousePointer = vbDefault
            End If
    End Select
    Unload Me
End Sub
