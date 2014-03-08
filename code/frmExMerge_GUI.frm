VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmExMerge_GUI 
   Caption         =   "AIExBK (Ascii Exchange Backup)"
   ClientHeight    =   8130
   ClientLeft      =   240
   ClientTop       =   435
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   10935
   Begin TabDlg.SSTab SSTab3 
      Height          =   7695
      Left            =   540
      TabIndex        =   0
      Top             =   240
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Intro"
      TabPicture(0)   =   "frmExMerge_GUI.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fIntro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "List4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Database"
      TabPicture(1)   =   "frmExMerge_GUI.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Thresholds"
      TabPicture(2)   =   "frmExMerge_GUI.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Session"
      TabPicture(3)   =   "frmExMerge_GUI.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "ExMerge"
      TabPicture(4)   =   "frmExMerge_GUI.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Content"
      TabPicture(5)   =   "frmExMerge_GUI.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Process"
      TabPicture(6)   =   "frmExMerge_GUI.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Tab 7"
      TabPicture(7)   =   "frmExMerge_GUI.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "List3"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Label16"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).ControlCount=   2
      Begin VB.ListBox List4 
         Height          =   5715
         Left            =   7200
         TabIndex        =   27
         Top             =   1140
         Width           =   2655
      End
      Begin VB.ListBox List3 
         Height          =   5325
         ItemData        =   "frmExMerge_GUI.frx":00E0
         Left            =   -74760
         List            =   "frmExMerge_GUI.frx":00F0
         TabIndex        =   25
         Top             =   780
         Width           =   9435
      End
      Begin VB.Frame fIntro 
         Height          =   6195
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   6735
         Begin VB.ListBox List2 
            Height          =   2205
            ItemData        =   "frmExMerge_GUI.frx":0284
            Left            =   1440
            List            =   "frmExMerge_GUI.frx":02A6
            TabIndex        =   24
            Top             =   2460
            Width           =   4935
         End
         Begin VB.ListBox List1 
            Height          =   1230
            ItemData        =   "frmExMerge_GUI.frx":03AE
            Left            =   1440
            List            =   "frmExMerge_GUI.frx":03C4
            TabIndex        =   22
            Top             =   1140
            Width           =   4935
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Next -->"
            Height          =   435
            Left            =   5400
            TabIndex        =   17
            Top             =   5580
            Width           =   1035
         End
         Begin VB.Label Label14 
            Caption         =   "Future options:"
            Height          =   255
            Left            =   180
            TabIndex        =   23
            Top             =   2460
            Width           =   1155
         End
         Begin VB.Label Label13 
            Caption         =   "Options include:"
            Height          =   255
            Left            =   180
            TabIndex        =   21
            Top             =   1140
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "AIExBK is a flexible application that allows you backup many aspects of Exchange."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   20
            Top             =   300
            Width           =   6135
         End
         Begin VB.Label Label8 
            Caption         =   "Choose an option and click 'Next' to continue."
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   5760
            Width           =   3435
         End
         Begin VB.Label Label7 
            Caption         =   $"frmExMerge_GUI.frx":046A
            Height          =   495
            Left            =   240
            TabIndex        =   18
            Top             =   4860
            Width           =   6135
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6195
         Left            =   -74520
         TabIndex        =   1
         Top             =   780
         Width           =   6735
         Begin VB.OptionButton Option4 
            Caption         =   "Backup Individual Mailboxes"
            Height          =   195
            Left            =   180
            TabIndex        =   6
            Top             =   2100
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Backup Exchange Database (mail)"
            Enabled         =   0   'False
            Height          =   195
            Left            =   180
            TabIndex        =   5
            Top             =   2580
            Width           =   3015
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Next -->"
            Height          =   435
            Left            =   5400
            TabIndex        =   4
            Top             =   5400
            Width           =   1035
         End
         Begin VB.TextBox Text2 
            Height          =   315
            Left            =   3240
            MaxLength       =   190
            TabIndex        =   3
            Top             =   2280
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Height          =   1695
            Left            =   3240
            MaxLength       =   1950
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   2
            Top             =   3000
            Width           =   3135
         End
         Begin VB.Label Label11 
            Caption         =   $"frmExMerge_GUI.frx":050C
            Height          =   495
            Left            =   180
            TabIndex        =   15
            Top             =   300
            Width           =   6135
         End
         Begin VB.Label Label10 
            Caption         =   "Choose an option and click 'Next' to continue."
            Height          =   255
            Left            =   180
            TabIndex        =   14
            Top             =   1620
            Width           =   3435
         End
         Begin VB.Label Label9 
            Caption         =   "c:\program files\exchsrvr\bin\exmerge -F c:\adminexmerge.ini -B -D"
            Height          =   255
            Left            =   180
            TabIndex        =   13
            Top             =   5220
            Visible         =   0   'False
            Width           =   4935
         End
         Begin VB.Label Label6 
            Caption         =   "Exmerge.log in c:\program files\exchsrvr\bin\"
            Height          =   255
            Left            =   180
            TabIndex        =   12
            Top             =   5460
            Visible         =   0   'False
            Width           =   3435
         End
         Begin VB.Label Label5 
            Caption         =   "Note: Only Groups/Users that have been selected from the list in step 2 will be backed up per session."
            Height          =   495
            Left            =   180
            TabIndex        =   11
            Top             =   960
            Width           =   5715
         End
         Begin VB.Label Label4 
            Caption         =   "Backup Session Name:"
            Height          =   255
            Left            =   3240
            TabIndex        =   10
            Top             =   2040
            Width           =   1875
         End
         Begin VB.Label Label3 
            Caption         =   "Backup Session Description:"
            Height          =   255
            Left            =   3240
            TabIndex        =   9
            Top             =   2760
            Width           =   2295
         End
         Begin VB.Label Label2 
            Caption         =   "Hour:"
            Height          =   255
            Left            =   180
            TabIndex        =   8
            Top             =   5880
            Width           =   435
         End
         Begin VB.Label Label1 
            Height          =   255
            Left            =   600
            TabIndex        =   7
            Top             =   5880
            Width           =   435
         End
      End
      Begin VB.Label Label15 
         Caption         =   "sequence:"
         Height          =   195
         Left            =   7200
         TabIndex        =   28
         Top             =   900
         Width           =   2175
      End
      Begin VB.Label Label16 
         Caption         =   "Instructions:"
         Height          =   195
         Left            =   -74700
         TabIndex        =   26
         Top             =   540
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmExMerge_GUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Init
End Sub

Sub Init()
    Me.Width = 12000
    Me.Height = 8600
    SSTab3.Width = 10185
    SSTab3.Height = 7600
    SSTab3.Top = 100
    SSTab3.Left = 100
End Sub

Sub sequence()
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
    List4.AddItem
End Sub
