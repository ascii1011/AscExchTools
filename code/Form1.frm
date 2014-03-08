VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6675
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   11774
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabMaxWidth     =   1764
      TabCaption(0)   =   "Intro"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fIntro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Step 1"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fMailboxFile"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fAdminExFile"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Step 2"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fObjectsToBackup"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fManageGroups"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Step 3"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fProcess"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Sessions"
      TabPicture(4)   =   "Form1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fResults"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Process"
      TabPicture(5)   =   "Form1.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fProcessErrors"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Frame fManageGroups 
         Caption         =   "Manage Groups"
         Height          =   6195
         Left            =   -71760
         TabIndex        =   109
         Top             =   360
         Width           =   3615
         Begin VB.ListBox List1 
            Height          =   2400
            Left            =   120
            TabIndex        =   118
            Top             =   2640
            Width           =   2355
         End
         Begin VB.TextBox Text5 
            Height          =   315
            Left            =   120
            TabIndex        =   117
            Top             =   1680
            Width           =   2535
         End
         Begin VB.CommandButton Command3 
            Caption         =   "<-- Add"
            Height          =   315
            Left            =   60
            TabIndex        =   116
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Cleared List"
            Height          =   315
            Left            =   120
            TabIndex        =   115
            Top             =   5400
            Width           =   1155
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Remove -->"
            Height          =   315
            Left            =   1500
            TabIndex        =   114
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            TabIndex        =   113
            Top             =   540
            Width           =   2355
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Create"
            Height          =   315
            Left            =   120
            TabIndex        =   112
            Top             =   900
            Width           =   1155
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Delete"
            Height          =   315
            Left            =   1500
            TabIndex        =   111
            Top             =   900
            Width           =   1155
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Next -->"
            Height          =   315
            Left            =   1500
            TabIndex        =   110
            Top             =   5400
            Width           =   1035
         End
         Begin VB.Label Label39 
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   5760
            TabIndex        =   124
            Top             =   5580
            Width           =   2415
         End
         Begin VB.Label Label5 
            Caption         =   "Users (/CN=):"
            Height          =   195
            Left            =   120
            TabIndex        =   123
            Top             =   2400
            Width           =   1035
         End
         Begin VB.Label Label6 
            Caption         =   "AD User Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label21 
            Height          =   195
            Left            =   720
            TabIndex        =   121
            Top             =   5160
            Width           =   1755
         End
         Begin VB.Label Label29 
            Caption         =   "Groups:"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label31 
            Caption         =   "Note:"
            Height          =   195
            Left            =   180
            TabIndex        =   119
            Top             =   5160
            Width           =   435
         End
      End
      Begin VB.Frame fAdminExFile 
         Caption         =   "Create AdminEx.ini File"
         Height          =   4155
         Left            =   -74880
         TabIndex        =   86
         Top             =   2400
         Width           =   5295
         Begin VB.CheckBox Check16 
            Caption         =   "Backup only Current."
            Height          =   195
            Left            =   3000
            TabIndex        =   102
            Top             =   2880
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Next -->"
            Height          =   435
            Left            =   4020
            TabIndex        =   101
            Top             =   3600
            Width           =   1035
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Set Test Variables"
            Height          =   375
            Left            =   120
            TabIndex        =   100
            Top             =   3600
            Width           =   1515
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Set for Modcon002"
            Height          =   375
            Left            =   120
            TabIndex        =   99
            Top             =   3180
            Width           =   1515
         End
         Begin VB.CheckBox Check7 
            Caption         =   "ZipFiles after backup"
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   2880
            Value           =   1  'Checked
            Width           =   1875
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Mail > 1Gig,break down semi-annually."
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            TabIndex        =   97
            Top             =   2220
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.TextBox Text6 
            Height          =   315
            Left            =   840
            TabIndex        =   96
            Text            =   "d:\backup\exchange"
            Top             =   300
            Width           =   2295
         End
         Begin VB.CommandButton Command6 
            Caption         =   "..."
            Height          =   315
            Left            =   4620
            TabIndex        =   95
            Top             =   300
            Width           =   255
         End
         Begin VB.TextBox Text7 
            Height          =   315
            Left            =   1140
            TabIndex        =   94
            Text            =   "modcon002"
            Top             =   780
            Width           =   3435
         End
         Begin VB.TextBox Text8 
            Height          =   315
            Left            =   1140
            TabIndex        =   93
            Text            =   "d:\backup\exchange\pst\"
            Top             =   1140
            Width           =   3435
         End
         Begin VB.CommandButton Command7 
            Caption         =   "..."
            Height          =   315
            Left            =   4620
            TabIndex        =   92
            Top             =   1200
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Saving Folder Rules"
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   2280
            Width           =   1755
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Extracting Data from Dumpster"
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   2580
            Width           =   2475
         End
         Begin VB.TextBox Text15 
            Height          =   315
            Left            =   3420
            TabIndex        =   89
            Text            =   "AdminEx.ini"
            Top             =   300
            Width           =   1155
         End
         Begin VB.TextBox Text18 
            Height          =   315
            Left            =   1140
            TabIndex        =   88
            Text            =   "d:\backup\exchange\AdminEx.log"
            Top             =   1860
            Width           =   3435
         End
         Begin VB.TextBox Text19 
            Height          =   315
            Left            =   1140
            TabIndex        =   87
            Text            =   "d:\backup\exchange\archive\"
            Top             =   1500
            Width           =   3435
         End
         Begin VB.Label Label9 
            Caption         =   "Filename:"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label10 
            Caption         =   "Server Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Destination:"
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label26 
            Caption         =   "\"
            Height          =   255
            Left            =   3240
            TabIndex        =   105
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label27 
            Caption         =   "Log File Path:"
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label28 
            Caption         =   "Archive Path:"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   1560
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1995
         Left            =   -69600
         TabIndex        =   74
         Top             =   2580
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CheckBox Check3 
            Caption         =   "For this session only backup certain date range."
            Height          =   195
            Left            =   180
            TabIndex        =   83
            Top             =   480
            Width           =   3855
         End
         Begin VB.TextBox Text12 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2520
            TabIndex        =   82
            ToolTipText     =   "Ex. '4:00 PM'"
            Top             =   1380
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   3660
            Picture         =   "Form1.frx":00A8
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   81
            Top             =   1380
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2040
            Picture         =   "Form1.frx":0232
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   80
            Top             =   1380
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.TextBox Text11 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1020
            TabIndex        =   79
            Top             =   1380
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox Text10 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2520
            TabIndex        =   78
            ToolTipText     =   "Ex. '4:00 PM'"
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   3660
            Picture         =   "Form1.frx":0674
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   77
            Top             =   960
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2040
            Picture         =   "Form1.frx":07FE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   76
            Top             =   960
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.TextBox Text9 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1020
            TabIndex        =   75
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "End Date:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   180
            TabIndex        =   85
            Top             =   1440
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label Label12 
            Caption         =   "Start Date:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   180
            TabIndex        =   84
            Top             =   1020
            Visible         =   0   'False
            Width           =   795
         End
      End
      Begin VB.Frame fObjectsToBackup 
         Caption         =   "Objects to backup"
         Height          =   6195
         Left            =   -74880
         TabIndex        =   65
         Top             =   360
         Width           =   3075
         Begin VB.OptionButton Option5 
            Caption         =   "Groups"
            Height          =   195
            Left            =   180
            TabIndex        =   68
            Top             =   480
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Users"
            Height          =   195
            Left            =   1860
            TabIndex        =   67
            Top             =   480
            Width           =   915
         End
         Begin VB.ListBox List5 
            Height          =   4560
            Left            =   180
            Style           =   1  'Checkbox
            TabIndex        =   66
            Top             =   1200
            Width           =   2715
         End
         Begin VB.Label Label36 
            Height          =   195
            Left            =   2160
            TabIndex        =   73
            Top             =   5880
            Width           =   735
         End
         Begin VB.Label Label35 
            Caption         =   "Selected:"
            Height          =   195
            Left            =   1320
            TabIndex        =   72
            Top             =   5880
            Width           =   675
         End
         Begin VB.Label Label34 
            Caption         =   "Count:"
            Height          =   195
            Left            =   180
            TabIndex        =   71
            Top             =   5880
            Width           =   495
         End
         Begin VB.Label Label33 
            Height          =   195
            Left            =   720
            TabIndex        =   70
            Top             =   5880
            Width           =   555
         End
         Begin VB.Label Label30 
            Caption         =   "Check Off Objects to Backup:"
            Height          =   255
            Left            =   180
            TabIndex        =   69
            Top             =   960
            Width           =   2235
         End
      End
      Begin VB.Frame fProcess 
         Height          =   6195
         Left            =   -74880
         TabIndex        =   41
         Top             =   360
         Width           =   6735
         Begin VB.Frame Frame1 
            Enabled         =   0   'False
            Height          =   2535
            Left            =   180
            TabIndex        =   49
            Top             =   1740
            Width           =   1995
            Begin VB.CheckBox Check9 
               Caption         =   "Mailboxes created."
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   56
               Top             =   240
               Width           =   1635
            End
            Begin VB.CheckBox Check10 
               Caption         =   "Settings Loaded."
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   55
               Top             =   540
               Width           =   1635
            End
            Begin VB.CheckBox Check11 
               Caption         =   "Starting Backup."
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   54
               Top             =   840
               Width           =   1755
            End
            Begin VB.CheckBox Check12 
               Caption         =   "Backup Completed."
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   53
               Top             =   1140
               Width           =   1755
            End
            Begin VB.CheckBox Check13 
               Caption         =   "Compiling Details."
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   52
               Top             =   1440
               Width           =   1755
            End
            Begin VB.CheckBox Check14 
               Caption         =   "No Errors Found."
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   51
               Top             =   1740
               Width           =   1755
            End
            Begin VB.CheckBox Check15 
               Caption         =   "Data Compressed."
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   50
               Top             =   2040
               Width           =   1755
            End
         End
         Begin VB.ListBox List6 
            Height          =   1425
            ItemData        =   "Form1.frx":0C40
            Left            =   120
            List            =   "Form1.frx":0C42
            TabIndex        =   48
            Top             =   4620
            Width           =   2115
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Start Over"
            Height          =   255
            Left            =   5280
            TabIndex        =   47
            Top             =   1320
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox Text20 
            Height          =   315
            Left            =   180
            TabIndex        =   46
            Top             =   1140
            Width           =   2835
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Kill after zip."
            Height          =   195
            Left            =   3120
            TabIndex        =   45
            Top             =   1080
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Move After zip"
            Height          =   195
            Left            =   3120
            TabIndex        =   44
            Top             =   900
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Use Project1.exe to zip"
            Height          =   195
            Left            =   3120
            TabIndex        =   43
            Top             =   720
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Process"
            Height          =   255
            Left            =   5280
            TabIndex        =   42
            Top             =   780
            Width           =   1035
         End
         Begin ComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   180
            TabIndex        =   57
            Top             =   420
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   450
            _Version        =   327682
            Appearance      =   1
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   4335
            Left            =   2340
            TabIndex        =   58
            Top             =   1680
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   7646
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Progress"
            TabPicture(0)   =   "Form1.frx":0C44
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "List4"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Report Details:"
            TabPicture(1)   =   "Form1.frx":0C60
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Text16"
            Tab(1).ControlCount=   1
            Begin VB.TextBox Text16 
               Height          =   3795
               Left            =   -74880
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   60
               Top             =   420
               Width           =   4035
            End
            Begin VB.ListBox List4 
               Height          =   3765
               Left            =   120
               TabIndex        =   59
               Top             =   420
               Width           =   4035
            End
         End
         Begin ComctlLib.ProgressBar ProgressBar2 
            Height          =   195
            Left            =   180
            TabIndex        =   61
            Top             =   1620
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   344
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Label Label40 
            Caption         =   "Users:"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   4380
            Width           =   1935
         End
         Begin VB.Label Label32 
            Caption         =   "Current Session Name:"
            Height          =   255
            Left            =   180
            TabIndex        =   63
            Top             =   900
            Width           =   2655
         End
         Begin VB.Label Label20 
            Height          =   195
            Left            =   180
            TabIndex        =   62
            Top             =   180
            Width           =   6015
         End
      End
      Begin VB.Frame fProcessErrors 
         Height          =   6195
         Left            =   -74880
         TabIndex        =   36
         Top             =   360
         Width           =   6735
         Begin VB.ListBox List3 
            Height          =   2205
            Left            =   120
            TabIndex        =   38
            Top             =   420
            Width           =   5475
         End
         Begin VB.TextBox Text14 
            Height          =   3075
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   37
            Top             =   2940
            Width           =   5475
         End
         Begin VB.Label Label19 
            Caption         =   "Process Entries:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   180
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "Process Details:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   2700
            Width           =   1815
         End
      End
      Begin VB.Frame fResults 
         Height          =   6195
         Left            =   -74880
         TabIndex        =   28
         Top             =   360
         Width           =   6735
         Begin VB.TextBox Text13 
            Height          =   1635
            Left            =   3360
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   30
            Top             =   4440
            Width           =   3135
         End
         Begin VB.ListBox List2 
            Height          =   1620
            Left            =   120
            TabIndex        =   29
            Top             =   4440
            Width           =   3195
         End
         Begin TrueDBGrid70.TDBGrid TDBGrid1 
            Height          =   3435
            Left            =   120
            TabIndex        =   31
            Top             =   420
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   6059
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Session"
            Columns(0).DataField=   ""
            Columns(0).DataWidth=   100
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Start Date"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Name"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Ending Stage"
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0).DividerColor=   12307669
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1958"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1879"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2381"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2302"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2910"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2831"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=1931"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1852"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DataMode        =   2
            DefColWidth     =   0
            EditDropDown    =   0   'False
            HeadLines       =   1
            FootLines       =   1
            TabAction       =   1
            WrapCellPointer =   -1  'True
            MultipleLines   =   0
            CellTipsWidth   =   0
            MultiSelect     =   2
            DeadAreaBackColor=   12307669
            ScrollTrack     =   -1  'True
            RowDividerColor =   12307669
            RowSubDividerColor=   12307669
            DirectionAfterEnter=   1
            MaxRows         =   250000
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=44,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=27,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=28,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=43,.parent=17"
            _StyleDefs(52)  =   "Named:id=33:Normal"
            _StyleDefs(53)  =   ":id=33,.parent=0"
            _StyleDefs(54)  =   "Named:id=34:Heading"
            _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   ":id=34,.wraptext=-1"
            _StyleDefs(57)  =   "Named:id=35:Footing"
            _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   "Named:id=36:Selected"
            _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=37:Caption"
            _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(63)  =   "Named:id=38:HighlightRow"
            _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=39:EvenRow"
            _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(67)  =   "Named:id=40:OddRow"
            _StyleDefs(68)  =   ":id=40,.parent=33"
            _StyleDefs(69)  =   "Named:id=41:RecordSelector"
            _StyleDefs(70)  =   ":id=41,.parent=34"
            _StyleDefs(71)  =   "Named:id=42:FilterBar"
            _StyleDefs(72)  =   ":id=42,.parent=33"
            _StyleDefs(73)  =   "Named:id=25:payment"
            _StyleDefs(74)  =   ":id=25,.parent=33,.fgcolor=&HFF&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(75)  =   ":id=25,.strikethrough=0,.charset=0"
            _StyleDefs(76)  =   ":id=25,.fontname=MS Sans Serif"
            _StyleDefs(77)  =   "Named:id=26:Balance"
            _StyleDefs(78)  =   ":id=26,.parent=25,.fgcolor=&HC0C0C0&,.borderColor=&H80000007&,.bold=-1"
            _StyleDefs(79)  =   ":id=26,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(80)  =   ":id=26,.fontname=MS Sans Serif"
         End
         Begin VB.Label Label23 
            Caption         =   "Session Count:"
            Height          =   255
            Left            =   4080
            TabIndex        =   35
            Top             =   3900
            Width           =   1155
         End
         Begin VB.Label lblSessionCount 
            Height          =   195
            Left            =   5280
            TabIndex        =   34
            Top             =   3900
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Session Log Details:"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   4200
            Width           =   2115
         End
         Begin VB.Label Label16 
            Caption         =   "Session Log Entries:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   180
            Width           =   1995
         End
      End
      Begin VB.Frame fIntro 
         Height          =   6195
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   6735
         Begin VB.TextBox Text22 
            Height          =   1695
            Left            =   3240
            MaxLength       =   1950
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   18
            Top             =   3000
            Width           =   3135
         End
         Begin VB.TextBox Text21 
            Height          =   315
            Left            =   3240
            MaxLength       =   190
            TabIndex        =   17
            Top             =   2280
            Width           =   3135
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Next -->"
            Height          =   435
            Left            =   5400
            TabIndex        =   16
            Top             =   5400
            Width           =   1035
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Backup Exchange Database (mail)"
            Enabled         =   0   'False
            Height          =   195
            Left            =   180
            TabIndex        =   15
            Top             =   2580
            Width           =   3015
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Backup Individual Mailboxes"
            Height          =   195
            Left            =   180
            TabIndex        =   14
            Top             =   2100
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.Label Label41 
            Height          =   255
            Left            =   600
            TabIndex        =   27
            Top             =   5880
            Width           =   435
         End
         Begin VB.Label Label24 
            Caption         =   "Hour:"
            Height          =   255
            Left            =   180
            TabIndex        =   26
            Top             =   5880
            Width           =   435
         End
         Begin VB.Label Label38 
            Caption         =   "Backup Session Description:"
            Height          =   255
            Left            =   3240
            TabIndex        =   25
            Top             =   2760
            Width           =   2295
         End
         Begin VB.Label Label37 
            Caption         =   "Backup Session Name:"
            Height          =   255
            Left            =   3240
            TabIndex        =   24
            Top             =   2040
            Width           =   1875
         End
         Begin VB.Label Label22 
            Caption         =   "Note: Only Groups/Users that have been selected from the list in step 2 will be backed up per session."
            Height          =   495
            Left            =   180
            TabIndex        =   23
            Top             =   960
            Width           =   5715
         End
         Begin VB.Label Label15 
            Caption         =   "Exmerge.log in c:\program files\exchsrvr\bin\"
            Height          =   255
            Left            =   180
            TabIndex        =   22
            Top             =   5460
            Visible         =   0   'False
            Width           =   3435
         End
         Begin VB.Label Label14 
            Caption         =   "c:\program files\exchsrvr\bin\exmerge -F c:\adminexmerge.ini -B -D"
            Height          =   255
            Left            =   180
            TabIndex        =   21
            Top             =   5220
            Visible         =   0   'False
            Width           =   4935
         End
         Begin VB.Label Label8 
            Caption         =   "Choose an option and click 'Next' to continue."
            Height          =   255
            Left            =   180
            TabIndex        =   20
            Top             =   1620
            Width           =   3435
         End
         Begin VB.Label Label7 
            Caption         =   $"Form1.frx":0C7C
            Height          =   495
            Left            =   180
            TabIndex        =   19
            Top             =   300
            Width           =   6135
         End
      End
      Begin VB.Frame fMailboxFile 
         Caption         =   "Create Template"
         Height          =   1995
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   5355
         Begin VB.TextBox Text17 
            Height          =   315
            Left            =   3660
            TabIndex        =   7
            Text            =   "mailboxes.txt"
            Top             =   300
            Width           =   1275
         End
         Begin VB.TextBox Text4 
            Height          =   315
            Left            =   1560
            TabIndex        =   6
            Text            =   "Recipients"
            Top             =   1560
            Width           =   3675
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   1560
            TabIndex        =   5
            Text            =   "FIRST ADMINISTRATIVE GROUP"
            Top             =   1140
            Width           =   3675
         End
         Begin VB.TextBox Text2 
            Height          =   315
            Left            =   1560
            TabIndex        =   4
            Text            =   "Modern Consumer"
            Top             =   720
            Width           =   3675
         End
         Begin VB.CommandButton Command4 
            Caption         =   "..."
            Height          =   315
            Left            =   4980
            TabIndex        =   3
            Top             =   300
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   840
            TabIndex        =   2
            Text            =   "d:\backup\exchange"
            Top             =   300
            Width           =   2535
         End
         Begin VB.Label Label25 
            Caption         =   "\"
            Height          =   255
            Left            =   3480
            TabIndex        =   12
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label4 
            Caption         =   "CN1 (/CN=):"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1620
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Group (/OU):"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Organization (/O=):"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   780
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Filename:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
