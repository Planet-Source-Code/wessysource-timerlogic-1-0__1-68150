VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPzl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Applications & folders"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "frmPzl.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDiag 
      Left            =   2640
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab sTab 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7011
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "  Applications"
      TabPicture(0)   =   "frmPzl.frx":0152
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtTaskName1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "MonthView1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtDate1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtTime1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtExePath"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdBrowse"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtComments1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdApply1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdCancel1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdHelp1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtPriority1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "picPrior1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "chkDisabled1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdNow1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "  Folders"
      TabPicture(1)   =   "frmPzl.frx":016E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdNow2"
      Tab(1).Control(1)=   "chkDisabled2"
      Tab(1).Control(2)=   "cmdPaste"
      Tab(1).Control(3)=   "txtTaskName2"
      Tab(1).Control(4)=   "txtDate2"
      Tab(1).Control(5)=   "txtTime2"
      Tab(1).Control(6)=   "txtFolderPath"
      Tab(1).Control(7)=   "txtComments2"
      Tab(1).Control(8)=   "cmdApply2"
      Tab(1).Control(9)=   "cmdCancel2"
      Tab(1).Control(10)=   "cmdHelp2"
      Tab(1).Control(11)=   "txtPriority2"
      Tab(1).Control(12)=   "picPrior2"
      Tab(1).Control(13)=   "MonthView2"
      Tab(1).Control(14)=   "Label12"
      Tab(1).Control(15)=   "Label11"
      Tab(1).Control(16)=   "Label10"
      Tab(1).Control(17)=   "Image2"
      Tab(1).Control(18)=   "Label9"
      Tab(1).Control(19)=   "Label8"
      Tab(1).Control(20)=   "Label7"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "  Web"
      TabPicture(2)   =   "frmPzl.frx":018A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(1)=   "Label14"
      Tab(2).Control(2)=   "Label15"
      Tab(2).Control(3)=   "Label16"
      Tab(2).Control(4)=   "Label17"
      Tab(2).Control(5)=   "Label18"
      Tab(2).Control(6)=   "Image3"
      Tab(2).Control(7)=   "MonthView3"
      Tab(2).Control(8)=   "txtTaskName3"
      Tab(2).Control(9)=   "txtDate3"
      Tab(2).Control(10)=   "txtTime3"
      Tab(2).Control(11)=   "txtWebPage"
      Tab(2).Control(12)=   "txtComments3"
      Tab(2).Control(13)=   "cmdApply3"
      Tab(2).Control(14)=   "cmdCancel3"
      Tab(2).Control(15)=   "cmdHelp3"
      Tab(2).Control(16)=   "txtPriority3"
      Tab(2).Control(17)=   "picPrior3"
      Tab(2).Control(18)=   "cmdPaste3"
      Tab(2).Control(19)=   "chkDisabled3"
      Tab(2).Control(20)=   "cmdNow3"
      Tab(2).ControlCount=   21
      Begin VB.CommandButton cmdNow3 
         Caption         =   "Now"
         Height          =   255
         Left            =   -70410
         TabIndex        =   60
         Top             =   240
         Width           =   465
      End
      Begin VB.CommandButton cmdNow2 
         Caption         =   "Now"
         Height          =   255
         Left            =   -70410
         TabIndex        =   59
         Top             =   240
         Width           =   465
      End
      Begin VB.CommandButton cmdNow1 
         Caption         =   "Now"
         Height          =   255
         Left            =   4590
         TabIndex        =   58
         Top             =   240
         Width           =   465
      End
      Begin VB.CheckBox chkDisabled3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Start disabled"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74880
         TabIndex        =   57
         Top             =   3420
         Width           =   1455
      End
      Begin VB.CheckBox chkDisabled2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Start disabled"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74880
         TabIndex        =   56
         Top             =   3420
         Width           =   1455
      End
      Begin VB.CheckBox chkDisabled1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Start disabled"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   3420
         Width           =   1455
      End
      Begin VB.CommandButton cmdPaste3 
         Caption         =   "Paste"
         Height          =   315
         Left            =   -70680
         TabIndex        =   54
         Top             =   1200
         Width           =   735
      End
      Begin VB.PictureBox picPrior3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -72600
         ScaleHeight     =   255
         ScaleWidth      =   705
         TabIndex        =   46
         Top             =   3075
         Width           =   735
         Begin VB.Shape shPriority3 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00800000&
            BorderStyle     =   0  'Transparent
            Height          =   285
            Left            =   0
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.TextBox txtPriority3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   45
         Top             =   3075
         Width           =   495
      End
      Begin VB.CommandButton cmdHelp3 
         Caption         =   "Help"
         Height          =   315
         Left            =   -74880
         TabIndex        =   44
         Top             =   3045
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel3 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   -71760
         TabIndex        =   43
         Top             =   3045
         Width           =   855
      End
      Begin VB.CommandButton cmdApply3 
         Caption         =   "Apply"
         Height          =   315
         Left            =   -70800
         TabIndex        =   42
         Top             =   3045
         Width           =   855
      End
      Begin VB.TextBox txtComments3 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   -72000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtWebPage 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   -72000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtTime3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71160
         TabIndex        =   39
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtDate3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71160
         TabIndex        =   38
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtTaskName3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73800
         TabIndex        =   37
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   -70920
         TabIndex        =   36
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtTaskName2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73800
         TabIndex        =   29
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtDate2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71160
         TabIndex        =   27
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtTime2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71160
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtFolderPath 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   -72000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtComments2 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   -72000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton cmdApply2 
         Caption         =   "Apply"
         Height          =   315
         Left            =   -70800
         TabIndex        =   23
         Top             =   3045
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   -71760
         TabIndex        =   22
         Top             =   3045
         Width           =   855
      End
      Begin VB.CommandButton cmdHelp2 
         Caption         =   "Help"
         Height          =   315
         Left            =   -74880
         TabIndex        =   21
         Top             =   3045
         Width           =   855
      End
      Begin VB.TextBox txtPriority2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   20
         Top             =   3075
         Width           =   495
      End
      Begin VB.PictureBox picPrior2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -72600
         ScaleHeight     =   255
         ScaleWidth      =   705
         TabIndex        =   19
         Top             =   3075
         Width           =   735
         Begin VB.Shape shPriority2 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00800000&
            BorderStyle     =   0  'Transparent
            Height          =   285
            Left            =   0
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.PictureBox picPrior1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2400
         ScaleHeight     =   255
         ScaleWidth      =   705
         TabIndex        =   18
         Top             =   3075
         Width           =   735
         Begin VB.Shape shpriority1 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00800000&
            BorderStyle     =   0  'Transparent
            Height          =   285
            Left            =   0
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.TextBox txtPriority1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   16
         Top             =   3075
         Width           =   495
      End
      Begin VB.CommandButton cmdHelp1 
         Caption         =   "Help"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   3045
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel1 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   3240
         TabIndex        =   14
         Top             =   3045
         Width           =   855
      End
      Begin VB.CommandButton cmdApply1 
         Caption         =   "Apply"
         Height          =   315
         Left            =   4200
         TabIndex        =   13
         Top             =   3045
         Width           =   855
      End
      Begin VB.TextBox txtComments1 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   315
         Left            =   4200
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtExePath 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtTime1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtDate1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2310
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         StartOfWeek     =   20905986
         CurrentDate     =   38894
      End
      Begin VB.TextBox txtTaskName1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
      Begin MSComCtl2.MonthView MonthView2 
         Height          =   2310
         Left            =   -74880
         TabIndex        =   28
         Top             =   600
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         StartOfWeek     =   20905986
         CurrentDate     =   38894
      End
      Begin MSComCtl2.MonthView MonthView3 
         Height          =   2310
         Left            =   -74880
         TabIndex        =   47
         Top             =   600
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         StartOfWeek     =   20905986
         CurrentDate     =   38894
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   -71970
         Picture         =   "frmPzl.frx":01A6
         Top             =   1245
         Width           =   240
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Priority:              %"
         Height          =   195
         Left            =   -73920
         TabIndex        =   53
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Comments:"
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
         Left            =   -72000
         TabIndex        =   52
         Top             =   2280
         Width           =   930
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Web page:"
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
         Left            =   -71640
         TabIndex        =   51
         Top             =   1260
         Width           =   945
      End
      Begin VB.Label Label15 
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
         Left            =   -72120
         TabIndex        =   50
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label14 
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
         Left            =   -72120
         TabIndex        =   49
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label13 
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
         Left            =   -74880
         TabIndex        =   48
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label12 
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
         Left            =   -74880
         TabIndex        =   35
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label11 
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
         Left            =   -72120
         TabIndex        =   34
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label10 
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
         Left            =   -72120
         TabIndex        =   33
         Top             =   240
         Width           =   885
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   -72000
         Picture         =   "frmPzl.frx":04E8
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Folder:"
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
         Left            =   -71610
         TabIndex        =   32
         Top             =   1260
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Comments:"
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
         Left            =   -72000
         TabIndex        =   31
         Top             =   2280
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Priority:              %"
         Height          =   195
         Left            =   -73920
         TabIndex        =   30
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Priority:              %"
         Height          =   195
         Left            =   1080
         TabIndex        =   17
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Comments:"
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
         Left            =   3000
         TabIndex        =   11
         Top             =   2280
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Exe path:"
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
         Left            =   3375
         TabIndex        =   8
         Top             =   1245
         Width           =   825
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   3000
         Picture         =   "frmPzl.frx":092A
         Top             =   1200
         Width           =   255
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
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   885
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
         Left            =   2880
         TabIndex        =   6
         Top             =   600
         Width           =   915
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
         TabIndex        =   2
         Top             =   120
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmPzl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApply1_Click()
Dim ExeItem As ListItem
If Len(txtTaskName1) = 0 Then MsgBox "You must introduce a task name.", vbExclamation, App.Title: txtTaskName1.SetFocus: Exit Sub
    If Len(txtTime1) = 0 Then
        MsgBox "You must introduce a start time.", vbExclamation, App.Title: txtTime1.SetFocus: Exit Sub
    Else
        If Len(txtDate1) = 0 Then
            MsgBox "You must introduce a start date.", vbExclamation, App.Title: txtDate1.SetFocus: Exit Sub
        Else
            If Len(txtPriority1) = 0 Then
                MsgBox "You must introduce a value as priority.", vbExclamation, App.Title: txtPriority1.SetFocus: Exit Sub
            Else
                If Len(txtExePath) = 0 Then
                    MsgBox "You must select an EXE file to run.", vbExclamation, App.Title: txtExePath.SetFocus: Exit Sub
                Else
                    If Len(txtTime1) < 8 Then
                        If txtTime = "Now" Or txtTime = "now" Then
                            GoTo Ok
                        Else
                            MsgBox "You must introduce a valid time.", vbExclamation, App.Title: txtTime1.SetFocus: Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
    
Ok: 'add to list
If Not editPzl Then
    With frmmain.LV
        Set ExeItem = .ListItems.Add(, , "", , 3)
        ExeItem.SubItems(1) = txtTaskName1
        ExeItem.SubItems(2) = "Run: " & txtExePath
        ExeItem.SubItems(3) = txtTime1
        ExeItem.SubItems(4) = txtDate1
        ExeItem.SubItems(5) = "Waiting"
        ExeItem.SubItems(6) = txtPriority1 & " %"
        ExeItem.SubItems(7) = "Pzl"
        ExeItem.SubItems(8) = txtComments1
    End With
    With frmmain
        .Toolbar1.Buttons(10).Enabled = True
        .Toolbar1.Buttons(11).Enabled = True
        .Toolbar1.Buttons(12).Enabled = True
                    
        .Toolbar2.Buttons(1).Enabled = False
        .Toolbar2.Buttons(2).Enabled = True
        .Toolbar2.Buttons(3).Enabled = True
        .Toolbar2.Buttons(4).Enabled = True
    End With
Else
    With frmmain.LV
        .ListItems.Item(.selectedItem.Index).SmallIcon = 3
        .selectedItem.SubItems(1) = txtTaskName1
        .selectedItem.SubItems(2) = "Run: " & txtExePath
        .selectedItem.SubItems(3) = txtTime1
        .selectedItem.SubItems(4) = txtDate1
        .selectedItem.SubItems(5) = "Waiting"
        .selectedItem.SubItems(6) = txtPriority1 & " %"
        .selectedItem.SubItems(7) = "Pzl"
        .selectedItem.SubItems(8) = txtComments1
    End With
    With frmmain
        .Toolbar1.Buttons(10).Enabled = True
        .Toolbar1.Buttons(11).Enabled = True
        .Toolbar1.Buttons(12).Enabled = True
                    
        .Toolbar2.Buttons(1).Enabled = False
        .Toolbar2.Buttons(2).Enabled = True
        .Toolbar2.Buttons(3).Enabled = True
        .Toolbar2.Buttons(4).Enabled = True
    End With
End If
editPzl = False
Unload Me
End Sub

Private Sub cmdApply2_Click()
Dim FoldItem As ListItem
If Len(txtTaskName2) = 0 Then MsgBox "You must introduce a task name.", vbExclamation, App.Title: txtTaskName2.SetFocus: Exit Sub
    If Len(txtTime2) = 0 Then
        MsgBox "You must introduce a start time.", vbExclamation, App.Title: txtTime2.SetFocus: Exit Sub
    Else
        If Len(txtDate2) = 0 Then
            MsgBox "You must introduce a start date.", vbExclamation, App.Title: txtDate2.SetFocus: Exit Sub
        Else
            If Len(txtPriority2) = 0 Then
                MsgBox "You must introduce a value as priority.", vbExclamation, App.Title: txtPriority2.SetFocus: Exit Sub
            Else
                If Len(txtFolderPath) = 0 Then
                    MsgBox "You must select an EXE file to run.", vbExclamation, App.Title: txtFolderPath.SetFocus: Exit Sub
                Else
                    If Right(txtFolderPath, 3) = "exe" Then
                        MsgBox "You don't select an EXE file here.", vbExclamation, App.Title: txtFolderPath.SetFocus: Exit Sub
                    Else
                        If Len(txtTime2) < 8 Then
                            If txtTime = "Now" Or txtTime = "now" Then
                                GoTo Ok
                            Else
                                MsgBox "You must introduce a valid time.", vbExclamation, App.Title: txtTime2.SetFocus: Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
Ok: 'add to list
If Not editPzl Then
    With frmmain.LV
        Set FoldItem = .ListItems.Add(, , "", , 3)
        FoldItem.SubItems(1) = txtTaskName2
        FoldItem.SubItems(2) = "OpenFold: " & txtFolderPath
        FoldItem.SubItems(3) = txtTime2
        FoldItem.SubItems(4) = txtDate2
        FoldItem.SubItems(5) = "Waiting"
        FoldItem.SubItems(6) = txtPriority2 & " %"
        FoldItem.SubItems(7) = "Pzl"
        FoldItem.SubItems(8) = txtComments2
    End With
    With frmmain
        .Toolbar1.Buttons(10).Enabled = True
        .Toolbar1.Buttons(11).Enabled = True
        .Toolbar1.Buttons(12).Enabled = True
                    
        .Toolbar2.Buttons(1).Enabled = False
        .Toolbar2.Buttons(2).Enabled = True
        .Toolbar2.Buttons(3).Enabled = True
        .Toolbar2.Buttons(4).Enabled = True
    End With
Else
    With frmmain.LV
          .ListItems.Item(.selectedItem.Index).SmallIcon = 3
          .selectedItem.SubItems(1) = txtTaskName2
          .selectedItem.SubItems(2) = "OpenFold: " & txtFolderPath
          .selectedItem.SubItems(3) = txtTime2
          .selectedItem.SubItems(4) = txtDate2
          .selectedItem.SubItems(5) = "Waiting"
          .selectedItem.SubItems(6) = txtPriority2 & " %"
          .selectedItem.SubItems(7) = "Pzl"
          .selectedItem.SubItems(8) = txtComments2
      End With
      With frmmain
          .Toolbar1.Buttons(10).Enabled = True
          .Toolbar1.Buttons(11).Enabled = True
          .Toolbar1.Buttons(12).Enabled = True
                      
          .Toolbar2.Buttons(1).Enabled = False
          .Toolbar2.Buttons(2).Enabled = True
          .Toolbar2.Buttons(3).Enabled = True
          .Toolbar2.Buttons(4).Enabled = True
      End With
End If
editPzl = False
Unload Me

End Sub

Private Sub cmdApply3_Click()
Dim WebItem As ListItem
If Len(txtTaskName3) = 0 Then MsgBox "You must introduce a task name.", vbExclamation, App.Title: txtTaskName3.SetFocus: Exit Sub
    If Len(txtTime3) = 0 Then
        MsgBox "You must introduce a start time.", vbExclamation, App.Title: txtTime3.SetFocus: Exit Sub
    Else
        If Len(txtDate3) = 0 Then
            MsgBox "You must introduce a start date.", vbExclamation, App.Title: txtDate3.SetFocus: Exit Sub
        Else
            If Len(txtPriority3) = 0 Then
                MsgBox "You must introduce a value as priority.", vbExclamation, App.Title: txtPriority3.SetFocus: Exit Sub
            Else
                If Len(txtWebPage) = 0 Then
                    MsgBox "You must introduce the web page.", vbExclamation, App.Title: txtWebPage.SetFocus: Exit Sub
                Else
                    If Left(txtWebPage, 3) <> "www" Then
                        MsgBox "You must introduce a valid web page.", vbExclamation, App.Title: txtWebPage.SetFocus: Exit Sub
                    Else
                        If Len(txtTime3) < 8 Then
                            If txtTime = "Now" Or txtTime = "now" Then
                                GoTo Ok
                            Else
                                MsgBox "You must introduce a valid time.", vbExclamation, App.Title: txtTime3.SetFocus: Exit Sub
                            End If
                        End If
                    End If
                 End If
            End If
        End If
    End If
    
Ok: 'add to list
If Not editPzl Then
    With frmmain.LV
        Set WebItem = .ListItems.Add(, , "", , 3)
        WebItem.SubItems(1) = txtTaskName3
        WebItem.SubItems(2) = "OpenWeb: " & txtWebPage
        If Left(txtTime3, 1) = "0" Then
            WebItem.SubItems(3) = Mid(txtTime3, 2)
        Else
            WebItem.SubItems(3) = txtTime3
        End If
        WebItem.SubItems(4) = txtDate3
        WebItem.SubItems(5) = "Waiting"
        WebItem.SubItems(6) = txtPriority3 & " %"
        WebItem.SubItems(7) = "Pzl"
        WebItem.SubItems(8) = txtComments3
    End With
    With frmmain
        .Toolbar1.Buttons(10).Enabled = True
        .Toolbar1.Buttons(11).Enabled = True
        .Toolbar1.Buttons(12).Enabled = True
                    
        .Toolbar2.Buttons(1).Enabled = False
        .Toolbar2.Buttons(2).Enabled = True
        .Toolbar2.Buttons(3).Enabled = True
        .Toolbar2.Buttons(4).Enabled = True
    End With
Else
  With frmmain.LV
        .ListItems.Item(.selectedItem.Index).SmallIcon = 3
        .selectedItem.SubItems(1) = txtTaskName3
        .selectedItem.SubItems(2) = "OpenWeb: " & txtWebPage
        .selectedItem.SubItems(3) = txtTime3
        .selectedItem.SubItems(4) = txtDate3
        .selectedItem.SubItems(5) = "Waiting"
        .selectedItem.SubItems(6) = txtPriority3 & " %"
        .selectedItem.SubItems(7) = "Pzl"
        .selectedItem.SubItems(8) = txtComments3
    End With
    With frmmain
        .Toolbar1.Buttons(10).Enabled = True
        .Toolbar1.Buttons(11).Enabled = True
        .Toolbar1.Buttons(12).Enabled = True
                    
        .Toolbar2.Buttons(1).Enabled = False
        .Toolbar2.Buttons(2).Enabled = True
        .Toolbar2.Buttons(3).Enabled = True
        .Toolbar2.Buttons(4).Enabled = True
    End With
End If
editPzl = False
Unload Me

End Sub

Private Sub cmdBrowse_Click()
    CDiag.DefaultExt = "*.exe"
    CDiag.DialogTitle = "Select an EXE file..."
    CDiag.Filter = "Applications (*.exe)|*.exe"
    CDiag.ShowOpen
    txtExePath = CDiag.FileName
End Sub

Private Sub cmdCancel1_Click()
    editPzl = False
    Unload Me
End Sub

Private Sub cmdCancel2_Click()
    editPzl = False
    Unload Me
End Sub

Private Sub cmdCancel3_Click()
    editPzl = False
    Unload Me
End Sub

Private Sub cmdHelp1_Click()
    Wait 1
    frmHelp2.Show 1
End Sub

Private Sub cmdHelp2_Click()
    Wait 1
    frmHelp2.Show 1
End Sub

Private Sub cmdHelp3_Click()
    Wait 1
    frmHelp2.Show 1
End Sub
Private Sub cmdNow1_Click()
    txtTime1 = Time
End Sub
Private Sub cmdNow2_Click()
    txtTime2 = Time
End Sub
Private Sub cmdNow3_Click()
    txtTime3 = Time
End Sub
Private Sub cmdPaste_Click()
    txtFolderPath = Clipboard.GetText
    
End Sub

Private Sub cmdPaste3_Click()
    txtWebPage = Clipboard.GetText
End Sub

Private Sub Form_Load()
    sTab.TabPicture(0) = LoadPicture(App.Path & "/exepic.bmp")
    sTab.TabPicture(1) = LoadPicture(App.Path & "/foldpic.bmp")
    sTab.TabPicture(2) = LoadPicture(App.Path & "/web.bmp")
    picPrior1.BackColor = BackColor
    picPrior2.BackColor = BackColor
    picPrior3.BackColor = BackColor
    chkDisabled1.BackColor = BackColor
    chkDisabled2.BackColor = BackColor
    chkDisabled3.BackColor = BackColor
    shpriority1.BackColor = RGB(1, 97, 234)
    shPriority2.BackColor = shpriority1.BackColor
    shPriority3.BackColor = shPriority2.BackColor
    
    If Clipboard.GetText = "" Then
        cmdPaste.Enabled = False
        cmdPaste3.Enabled = False
    Else
        cmdPaste.Enabled = True
        cmdPaste3.Enabled = True
    End If
    
End Sub

Private Sub txtPriority_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    editPzl = False
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    If MonthView1.Day < 10 Then
        If MonthView1.Month < mvwOctober Then
            txtDate1 = "0" & MonthView1.Day & "/0" & MonthView1.Month & "/" & MonthView1.Year
        Else
            txtDate1 = "0" & MonthView1.Day & "/" & MonthView1.Month & "/" & MonthView1.Year
        End If
    Else
        If MonthView1.Month < mvwOctober Then
            txtDate1 = MonthView1.Day & "/0" & MonthView1.Month & "/" & MonthView1.Year
        Else
            txtDate1 = MonthView1.Day & "/" & MonthView1.Month & "/" & MonthView1.Year
        End If
    End If
End Sub

Private Sub MonthView2_DateClick(ByVal DateClicked As Date)
    If MonthView2.Day < 10 Then
        If MonthView2.Month < mvwOctober Then
            txtDate2 = "0" & MonthView2.Day & "/0" & MonthView2.Month & "/" & MonthView2.Year
        Else
            txtDate2 = "0" & MonthView2.Day & "/" & MonthView2.Month & "/" & MonthView2.Year
        End If
    Else
        If MonthView2.Month < mvwOctober Then
            txtDate2 = MonthView2.Day & "/0" & MonthView2.Month & "/" & MonthView2.Year
        Else
            txtDate2 = MonthView2.Day & "/" & MonthView2.Month & "/" & MonthView2.Year
        End If
    End If
End Sub
Private Sub MonthView3_DateClick(ByVal DateClicked As Date)
    If MonthView3.Day < 10 Then
        If MonthView3.Month < mvwOctober Then
            txtDate3 = "0" & MonthView3.Day & "/0" & MonthView3.Month & "/" & MonthView3.Year
        Else
            txtDate3 = "0" & MonthView3.Day & "/" & MonthView3.Month & "/" & MonthView3.Year
        End If
    Else
        If MonthView3.Month < mvwOctober Then
            txtDate3 = MonthView3.Day & "/0" & MonthView3.Month & "/" & MonthView3.Year
        Else
            txtDate3 = MonthView3.Day & "/" & MonthView3.Month & "/" & MonthView3.Year
        End If
    End If
End Sub


Private Sub txtExePath_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtPriority1_Change()
    If Val(txtPriority1) < 0 Or Val(txtPriority1) > 100 Then MsgBox "A priority number too longer.", vbCritical, App.Title: Exit Sub
    If Len(txtPriority1) = 0 Then Exit Sub
    If IsNumeric(txtPriority1) = False Then MsgBox "Invalid value as priority.", vbCritical, App.Title: Exit Sub
    shpriority1.Width = (Val(txtPriority1) * picPrior1.Width) / 100
End Sub

Private Sub txtPriority2_Change()
    If Val(txtPriority2) < 0 Or Val(txtPriority2) > 100 Then MsgBox "A priority number too longer.", vbCritical, App.Title: Exit Sub
    If Len(txtPriority2) = 0 Then Exit Sub
    If IsNumeric(txtPriority2) = False Then MsgBox "Invalid value as priority.", vbCritical, App.Title: Exit Sub
    shPriority2.Width = (Val(txtPriority2) * picPrior2.Width) / 100
End Sub

Private Sub txtPriority3_Change()
    If Val(txtPriority3) < 0 Or Val(txtPriority3) > 100 Then MsgBox "A priority number too longer.", vbCritical, App.Title: Exit Sub
    If Len(txtPriority3) = 0 Then Exit Sub
    If IsNumeric(txtPriority3) = False Then MsgBox "Invalid value as priority.", vbCritical, App.Title: Exit Sub
    shPriority3.Width = (Val(txtPriority3) * picPrior3.Width) / 100

End Sub

