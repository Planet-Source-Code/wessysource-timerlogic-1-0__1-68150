VERSION 5.00
Begin VB.Form frmConfirmation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exit"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5220
   Icon            =   "frmConfirmation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "No"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save list"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Y&es"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Do you want to quit without saving task list?"
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
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3795
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmConfirmation.frx":0442
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
Set bbd = OpenDatabase(App.Path & "\DataBaseList.mdb")
Set tbl = bbd.OpenRecordset("ListOfTasks")
            
            MousePointer = vbHourglass
            frmmain.Show
            Dim i
            For i = 1 To frmmain.LV.ListItems.Count
                tbl.AddNew
                
                tbl("TaskName") = frmmain.LV.ListItems.Item(i).SubItems(1)
                tbl("Task") = frmmain.LV.ListItems.Item(i).SubItems(2)
                tbl("Time") = frmmain.LV.ListItems.Item(i).SubItems(3)
                tbl("Date") = frmmain.LV.ListItems.Item(i).SubItems(4)
                tbl("Status") = frmmain.LV.ListItems.Item(i).SubItems(5)
                tbl("Priority") = frmmain.LV.ListItems.Item(i).SubItems(6)
                tbl("Type") = frmmain.LV.ListItems.Item(i).SubItems(7)
                tbl("Comments") = frmmain.LV.ListItems.Item(i).SubItems(8)
                tbl.Update
            Next i
                

                bbd.Close
                
                DoEvents
              
                MousePointer = vbDefault
                frmmain.timerEnd.Enabled = True
End Sub

Private Sub Command3_Click()
frmmain.Show
    Unload Me
    
End Sub
Public Sub EndApp()
    End
End Sub
