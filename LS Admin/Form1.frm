VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Login Script Administrator"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Compname 
      Height          =   285
      Left            =   6240
      TabIndex        =   166
      Top             =   6600
      Width           =   3135
   End
   Begin VB.TextBox DomainName 
      Height          =   285
      Left            =   6240
      TabIndex        =   165
      Top             =   6960
      Width           =   3135
   End
   Begin VB.TextBox username 
      Height          =   285
      Left            =   1560
      TabIndex        =   164
      Top             =   6600
      Width           =   3135
   End
   Begin VB.TextBox DomainController 
      Height          =   285
      Left            =   1560
      TabIndex        =   163
      Top             =   6960
      Width           =   3135
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11245
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "User / Group Settings"
      TabPicture(0)   =   "Form1.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Machine Settings"
      TabPicture(1)   =   "Form1.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SSTab3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Global Settings"
      TabPicture(2)   =   "Form1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label30"
      Tab(2).Control(1)=   "Label37"
      Tab(2).Control(2)=   "Label38"
      Tab(2).Control(3)=   "Label39"
      Tab(2).Control(4)=   "Label40"
      Tab(2).Control(5)=   "Command2"
      Tab(2).Control(6)=   "Exchange"
      Tab(2).Control(7)=   "PDC"
      Tab(2).Control(8)=   "Command3"
      Tab(2).Control(9)=   "BDC1"
      Tab(2).Control(10)=   "Command4"
      Tab(2).Control(11)=   "BDC2"
      Tab(2).Control(12)=   "Command5"
      Tab(2).Control(13)=   "BDC3"
      Tab(2).Control(14)=   "Command6"
      Tab(2).Control(15)=   "Command16"
      Tab(2).ControlCount=   16
      Begin VB.CommandButton Command16 
         Caption         =   "Save Global Defaults"
         Height          =   255
         Left            =   -71040
         TabIndex        =   16
         Top             =   4800
         Width           =   2655
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -67920
         TabIndex        =   12
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox BDC3 
         Height          =   285
         Left            =   -72720
         TabIndex        =   11
         Top             =   3720
         Width           =   4695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -67920
         TabIndex        =   9
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox BDC2 
         Height          =   285
         Left            =   -72720
         TabIndex        =   8
         Top             =   3120
         Width           =   4695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -67920
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox BDC1 
         Height          =   285
         Left            =   -72720
         TabIndex        =   5
         Top             =   2520
         Width           =   4695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -67920
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox PDC 
         Height          =   285
         Left            =   -72720
         TabIndex        =   2
         Top             =   1920
         Width           =   4695
      End
      Begin VB.TextBox Exchange 
         Height          =   285
         Left            =   -72720
         TabIndex        =   14
         Top             =   4320
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -67920
         TabIndex        =   15
         Top             =   4320
         Width           =   1455
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   147
         Top             =   600
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   9551
         _Version        =   393216
         Tabs            =   10
         Tab             =   8
         TabsPerRow      =   4
         TabHeight       =   397
         WordWrap        =   0   'False
         OLEDropMode     =   1
         TabCaption(0)   =   "Windows NT Profiles"
         TabPicture(0)   =   "Form1.frx":0496
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lpse"
         Tab(0).Control(1)=   "Scrollts"
         Tab(0).Control(2)=   "ts"
         Tab(0).Control(3)=   "doc1"
         Tab(0).Control(4)=   "doc"
         Tab(0).Control(5)=   "Scrollms"
         Tab(0).Control(6)=   "tm"
         Tab(0).Control(7)=   "tfdb"
         Tab(0).Control(8)=   "cpdo"
         Tab(0).Control(9)=   "sndpo"
         Tab(0).Control(10)=   "snct"
         Tab(0).Control(11)=   "adsnc"
         Tab(0).Control(12)=   "dccorp"
         Tab(0).Control(13)=   "Frame2"
         Tab(0).Control(14)=   "Label19"
         Tab(0).Control(15)=   "Label18"
         Tab(0).Control(16)=   "Label17"
         Tab(0).Control(17)=   "Label16"
         Tab(0).ControlCount=   18
         TabCaption(1)   =   "Internet Explorer"
         TabPicture(1)   =   "Form1.frx":04B2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label5"
         Tab(1).Control(1)=   "Label6"
         Tab(1).Control(2)=   "stp"
         Tab(1).Control(3)=   "sep"
         Tab(1).Control(4)=   "dicw"
         Tab(1).Control(5)=   "Frame1"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Microsoft Office"
         TabPicture(2)   =   "Form1.frx":04CE
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label45"
         Tab(2).Control(1)=   "md"
         Tab(2).Control(2)=   "Command17"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Microsoft Outlook"
         TabPicture(3)   =   "Form1.frx":04EA
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label31"
         Tab(3).Control(1)=   "Label29"
         Tab(3).Control(2)=   "Label28"
         Tab(3).Control(3)=   "Frame7"
         Tab(3).Control(4)=   "Frame6"
         Tab(3).Control(5)=   "ptewoe"
         Tab(3).Control(6)=   "Command9"
         Tab(3).Control(7)=   "Exchange2"
         Tab(3).Control(8)=   "mn"
         Tab(3).Control(9)=   "pn"
         Tab(3).Control(10)=   "oepwsn"
         Tab(3).ControlCount=   11
         TabCaption(4)   =   "Drive Mappings"
         TabPicture(4)   =   "Form1.frx":0506
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Command21"
         Tab(4).Control(1)=   "HomeRoot"
         Tab(4).Control(2)=   "HomeDrive"
         Tab(4).Control(3)=   "Command8"
         Tab(4).Control(4)=   "Command7"
         Tab(4).Control(5)=   "Drive"
         Tab(4).Control(6)=   "Path"
         Tab(4).Control(7)=   "DM"
         Tab(4).Control(8)=   "DelMap"
         Tab(4).Control(9)=   "Label48"
         Tab(4).Control(10)=   "Label47"
         Tab(4).Control(11)=   "Label22"
         Tab(4).Control(12)=   "Label23"
         Tab(4).ControlCount=   13
         TabCaption(5)   =   "System"
         TabPicture(5)   =   "Form1.frx":0522
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "dret"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Control Panel"
         TabPicture(6)   =   "Form1.frx":053E
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Display"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "Shell"
         TabPicture(7)   =   "Form1.frx":055A
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "rrcfsm"
         Tab(7).Control(1)=   "rsdcfsm"
         Tab(7).Control(2)=   "haiod"
         Tab(7).Control(3)=   "neninn"
         Tab(7).Control(4)=   "hnn"
         Tab(7).Control(5)=   "hdimc"
         Tab(7).Control(6)=   "rfcfsm"
         Tab(7).Control(7)=   "rtfsosm"
         Tab(7).Control(8)=   "rffsosm"
         Tab(7).Control(9)=   "rmnd"
         Tab(7).ControlCount=   10
         TabCaption(8)   =   "General"
         TabPicture(8)   =   "Form1.frx":0576
         Tab(8).ControlEnabled=   -1  'True
         Tab(8).Control(0)=   "Frame5"
         Tab(8).Control(0).Enabled=   0   'False
         Tab(8).Control(1)=   "Frame4"
         Tab(8).Control(1).Enabled=   0   'False
         Tab(8).ControlCount=   2
         TabCaption(9)   =   "Run"
         TabPicture(9)   =   "Form1.frx":0592
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "Label46"
         Tab(9).Control(1)=   "rp"
         Tab(9).Control(2)=   "Command18"
         Tab(9).Control(3)=   "Command19"
         Tab(9).Control(4)=   "run"
         Tab(9).Control(5)=   "Command20"
         Tab(9).ControlCount=   6
         Begin VB.CommandButton Command21 
            Caption         =   "Browse"
            Height          =   255
            Left            =   -67680
            TabIndex        =   174
            Top             =   4320
            Width           =   1335
         End
         Begin VB.TextBox HomeRoot 
            Height          =   285
            Left            =   -71640
            TabIndex        =   173
            Top             =   4320
            Width           =   3735
         End
         Begin VB.ComboBox HomeDrive 
            Height          =   315
            ItemData        =   "Form1.frx":05AE
            Left            =   -72600
            List            =   "Form1.frx":05F4
            TabIndex        =   172
            Top             =   4320
            Width           =   735
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Delete Selected Item"
            Height          =   255
            Left            =   -71880
            TabIndex        =   56
            Top             =   3600
            Width           =   3375
         End
         Begin VB.ListBox run 
            Height          =   1035
            ItemData        =   "Form1.frx":0650
            Left            =   -73440
            List            =   "Form1.frx":0652
            Sorted          =   -1  'True
            TabIndex        =   55
            Top             =   2280
            Width           =   6375
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Browse"
            Height          =   255
            Left            =   -68160
            TabIndex        =   54
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Add"
            Height          =   255
            Left            =   -73440
            TabIndex        =   52
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox rp 
            Height          =   285
            Left            =   -72480
            TabIndex        =   53
            Top             =   1560
            Width           =   4095
         End
         Begin VB.CheckBox lpse 
            Caption         =   "Limit Profile size enabled"
            Height          =   255
            Left            =   -74880
            TabIndex        =   57
            Top             =   840
            Width           =   4335
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Browse"
            Height          =   255
            Left            =   -67800
            TabIndex        =   98
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox md 
            Height          =   285
            Left            =   -71880
            TabIndex        =   97
            Top             =   2280
            Width           =   3975
         End
         Begin VB.Frame Frame4 
            Caption         =   "Load Settings From"
            Height          =   1815
            Left            =   360
            TabIndex        =   156
            Top             =   960
            Width           =   8775
            Begin VB.CommandButton Command12 
               Caption         =   "Load User / Group Template"
               Height          =   375
               Left            =   4680
               TabIndex        =   46
               Top             =   840
               Width           =   3375
            End
            Begin VB.ComboBox LoadGroup 
               Height          =   315
               Left            =   240
               Sorted          =   -1  'True
               TabIndex        =   45
               Top             =   1320
               Width           =   3135
            End
            Begin VB.ComboBox LoadUser 
               Height          =   315
               Left            =   240
               Sorted          =   -1  'True
               TabIndex        =   43
               Top             =   600
               Width           =   3135
            End
            Begin VB.Label Label25 
               Caption         =   "Group"
               Height          =   255
               Left            =   1440
               TabIndex        =   44
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label24 
               Caption         =   "Username"
               Height          =   255
               Left            =   1320
               TabIndex        =   42
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Save Settings To"
            Height          =   1920
            Left            =   360
            TabIndex        =   155
            Top             =   2895
            Width           =   8775
            Begin VB.CommandButton Command13 
               Caption         =   "Save User / Group Template"
               Height          =   375
               Left            =   4680
               TabIndex        =   51
               Top             =   840
               Width           =   3375
            End
            Begin VB.ComboBox SaveGroup 
               Height          =   315
               Left            =   240
               Sorted          =   -1  'True
               TabIndex        =   50
               Top             =   1440
               Width           =   3135
            End
            Begin VB.ComboBox SaveUser 
               Height          =   315
               Left            =   240
               Sorted          =   -1  'True
               TabIndex        =   48
               Top             =   600
               Width           =   3135
            End
            Begin VB.Label Label27 
               Caption         =   "Username"
               Height          =   255
               Left            =   1320
               TabIndex        =   47
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label26 
               Caption         =   "Group"
               Height          =   255
               Left            =   1440
               TabIndex        =   49
               Top             =   1200
               Width           =   615
            End
         End
         Begin VB.CheckBox rmnd 
            Caption         =   "Remove the ""Map Network Drive"" and ""Disconnect Network Drive"" options"
            Height          =   255
            Left            =   -72000
            TabIndex        =   146
            Top             =   4455
            Width           =   5775
         End
         Begin VB.CheckBox rffsosm 
            Caption         =   "Remove folders from Settings on Start menu"
            Height          =   255
            Left            =   -72000
            TabIndex        =   138
            Top             =   1575
            Width           =   3735
         End
         Begin VB.CheckBox rtfsosm 
            Caption         =   "Remove Taskbar from Settings on Start menu"
            Height          =   255
            Left            =   -72000
            TabIndex        =   139
            Top             =   1935
            Width           =   3735
         End
         Begin VB.CheckBox rfcfsm 
            Caption         =   "Remove Find command from Start menu"
            Height          =   255
            Left            =   -72000
            TabIndex        =   140
            Top             =   2295
            Width           =   3495
         End
         Begin VB.CheckBox hdimc 
            Caption         =   "Hide drives in My Computer"
            Height          =   255
            Left            =   -72000
            TabIndex        =   141
            Top             =   2655
            Width           =   3495
         End
         Begin VB.CheckBox hnn 
            Caption         =   "Hide Network Neighborhood"
            Height          =   255
            Left            =   -72000
            TabIndex        =   142
            Top             =   3015
            Width           =   3615
         End
         Begin VB.CheckBox neninn 
            Caption         =   "No Entire Network in Network Neighborhood"
            Height          =   255
            Left            =   -72000
            TabIndex        =   143
            Top             =   3375
            Width           =   3615
         End
         Begin VB.CheckBox haiod 
            Caption         =   "Hide all items on desktop"
            Height          =   255
            Left            =   -72000
            TabIndex        =   144
            Top             =   3735
            Width           =   2175
         End
         Begin VB.CheckBox rsdcfsm 
            Caption         =   "Remove Shut Down command from Start menu"
            Height          =   255
            Left            =   -72000
            TabIndex        =   145
            Top             =   4095
            Width           =   3735
         End
         Begin VB.CheckBox rrcfsm 
            Caption         =   "Remove Run command from Start menu"
            Height          =   255
            Left            =   -72000
            TabIndex        =   137
            Top             =   1200
            Width           =   3735
         End
         Begin VB.CheckBox dret 
            Caption         =   "Disable Registry editing tools"
            Height          =   255
            Left            =   -71520
            TabIndex        =   131
            Top             =   2535
            Width           =   3975
         End
         Begin VB.VScrollBar Scrollts 
            Enabled         =   0   'False
            Height          =   255
            Left            =   -67200
            Max             =   0
            Min             =   120
            TabIndex        =   154
            Top             =   4575
            Value           =   30
            Width           =   255
         End
         Begin VB.TextBox ts 
            Enabled         =   0   'False
            Height          =   285
            Left            =   -68160
            TabIndex        =   84
            Text            =   "30"
            Top             =   4575
            Width           =   855
         End
         Begin VB.ComboBox doc1 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Form1.frx":0654
            Left            =   -68400
            List            =   "Form1.frx":065E
            TabIndex        =   81
            Top             =   3735
            Width           =   2175
         End
         Begin VB.ComboBox doc 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Form1.frx":0687
            Left            =   -68400
            List            =   "Form1.frx":0691
            TabIndex        =   78
            Top             =   2895
            Width           =   2175
         End
         Begin VB.VScrollBar Scrollms 
            Enabled         =   0   'False
            Height          =   255
            Left            =   -67200
            Max             =   0
            Min             =   10000
            TabIndex        =   153
            Top             =   2055
            Value           =   2000
            Width           =   255
         End
         Begin VB.TextBox tm 
            Enabled         =   0   'False
            Height          =   285
            Left            =   -68160
            TabIndex        =   75
            Text            =   "2000"
            Top             =   2055
            Width           =   855
         End
         Begin VB.CheckBox tfdb 
            Caption         =   "Timeout for dialog boxes"
            Height          =   255
            Left            =   -70320
            TabIndex        =   82
            Top             =   4215
            Width           =   4455
         End
         Begin VB.CheckBox cpdo 
            Caption         =   "Choose profile default operation"
            Height          =   255
            Left            =   -70320
            TabIndex        =   79
            Top             =   3375
            Width           =   4455
         End
         Begin VB.CheckBox sndpo 
            Caption         =   "Slow network default profile operation"
            Height          =   255
            Left            =   -70320
            TabIndex        =   76
            Top             =   2535
            Width           =   4575
         End
         Begin VB.CheckBox snct 
            Caption         =   "Slow network connection timeout"
            Height          =   255
            Left            =   -70320
            TabIndex        =   73
            Top             =   1695
            Width           =   4575
         End
         Begin VB.CheckBox adsnc 
            Caption         =   "Automatically detect slow network connections"
            Height          =   255
            Left            =   -70320
            TabIndex        =   72
            Top             =   1335
            Width           =   4335
         End
         Begin VB.CheckBox dccorp 
            Caption         =   "Delete cached copies of roaming profiles"
            Height          =   255
            Left            =   -70320
            TabIndex        =   71
            Top             =   975
            Width           =   4215
         End
         Begin VB.Frame Frame2 
            Caption         =   "Limit Profile Size"
            Enabled         =   0   'False
            Height          =   4215
            Left            =   -74880
            TabIndex        =   152
            Top             =   1080
            Width           =   4335
            Begin VB.TextBox cm 
               Height          =   285
               Left            =   1440
               TabIndex        =   59
               Text            =   $"Form1.frx":06BA
               Top             =   360
               Width           =   2775
            End
            Begin VB.TextBox mps 
               Height          =   285
               Left            =   1680
               TabIndex        =   61
               Text            =   "30000"
               Top             =   840
               Width           =   855
            End
            Begin VB.VScrollBar ScrollKB 
               Height          =   255
               LargeChange     =   100
               Left            =   2640
               Max             =   0
               Min             =   30000
               SmallChange     =   100
               TabIndex        =   62
               Top             =   840
               Value           =   30000
               Width           =   255
            End
            Begin VB.CheckBox irifl 
               Caption         =   "Include registry in file list"
               Height          =   255
               Left            =   120
               TabIndex        =   63
               Top             =   1200
               Width           =   3735
            End
            Begin VB.CheckBox nuwpssie 
               Caption         =   "Notify user when profile storage space is exceeded."
               Height          =   195
               Left            =   120
               TabIndex        =   64
               Top             =   1560
               Width           =   3975
            End
            Begin VB.TextBox ruexm 
               Height          =   285
               Left            =   2280
               TabIndex        =   66
               Text            =   "15"
               Top             =   1920
               Width           =   735
            End
            Begin VB.VScrollBar ScrollMin 
               Height          =   255
               Left            =   3120
               Max             =   0
               Min             =   60
               TabIndex        =   67
               Top             =   1920
               Value           =   15
               Width           =   255
            End
            Begin VB.TextBox pdirectory 
               Height          =   285
               Left            =   120
               TabIndex        =   69
               Text            =   "Temporary Internet Files;Temp"
               Top             =   3000
               Width           =   4095
            End
            Begin VB.Label Label11 
               Caption         =   "Custom Message"
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label12 
               Caption         =   "Max Profile size (KB)"
               Height          =   255
               Left            =   120
               TabIndex        =   60
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label13 
               Caption         =   "Remind user every X minutes:"
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   1920
               Width           =   2175
            End
            Begin VB.Label Label14 
               Caption         =   "Prevent the following directories from roaming with the profile:"
               Height          =   375
               Left            =   120
               TabIndex        =   68
               Top             =   2520
               Width           =   4095
            End
            Begin VB.Label Label15 
               Caption         =   "You can enter multiple name, semi-colon seperated, all relative to the root of the user's profile"
               Height          =   495
               Left            =   120
               TabIndex        =   70
               Top             =   3480
               Width           =   4095
            End
         End
         Begin VB.Frame Display 
            Caption         =   "Display"
            Height          =   4095
            Left            =   -74760
            TabIndex        =   151
            Top             =   960
            Width           =   9015
            Begin VB.VScrollBar Scrollss 
               Height          =   255
               Left            =   5760
               Max             =   1
               Min             =   30
               TabIndex        =   184
               Top             =   1920
               Value           =   5
               Width           =   255
            End
            Begin VB.TextBox SSTime 
               Height          =   285
               Left            =   5280
               TabIndex        =   183
               Text            =   "5"
               Top             =   1920
               Width           =   375
            End
            Begin VB.CheckBox SSPassword 
               Caption         =   "Password Protected"
               Height          =   255
               Left            =   3840
               TabIndex        =   179
               Top             =   2280
               Width           =   3135
            End
            Begin VB.CommandButton Command22 
               Caption         =   "Browse"
               Height          =   255
               Left            =   6960
               TabIndex        =   178
               Top             =   1440
               Width           =   975
            End
            Begin VB.TextBox SSFile 
               Height          =   285
               Left            =   3840
               TabIndex        =   177
               Top             =   1440
               Width           =   3015
            End
            Begin VB.CheckBox datdi 
               Caption         =   "Deny access to display icon"
               Height          =   255
               Left            =   600
               TabIndex        =   132
               Top             =   1200
               Width           =   2415
            End
            Begin VB.CheckBox hbt 
               Caption         =   "Hide Background tab"
               Height          =   255
               Left            =   600
               TabIndex        =   133
               Top             =   1560
               Width           =   2295
            End
            Begin VB.CheckBox hsst 
               Caption         =   "Hide Screen Saver tab"
               Height          =   255
               Left            =   600
               TabIndex        =   134
               Top             =   1920
               Width           =   2295
            End
            Begin VB.CheckBox hat 
               Caption         =   "Hide Appearance tab"
               Height          =   255
               Left            =   600
               TabIndex        =   135
               Top             =   2280
               Width           =   2415
            End
            Begin VB.CheckBox hst 
               Caption         =   "Hide Settings tab"
               Height          =   255
               Left            =   600
               TabIndex        =   136
               Top             =   2640
               Width           =   2415
            End
            Begin VB.Label Label51 
               Caption         =   "Timeout (Minutes)"
               Height          =   255
               Left            =   3840
               TabIndex        =   182
               Top             =   1920
               Width           =   1455
            End
            Begin VB.Label Label50 
               Caption         =   "Screen Save Properties"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4560
               TabIndex        =   181
               Top             =   840
               Width           =   2175
            End
            Begin VB.Label Label49 
               Caption         =   "Path"
               Height          =   255
               Left            =   3840
               TabIndex        =   180
               Top             =   1200
               Width           =   855
            End
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Add"
            Height          =   255
            Left            =   -73920
            TabIndex        =   123
            Top             =   1455
            Width           =   1095
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Browse"
            Height          =   255
            Left            =   -67680
            TabIndex        =   126
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ComboBox Drive 
            Height          =   315
            ItemData        =   "Form1.frx":074F
            Left            =   -72600
            List            =   "Form1.frx":0795
            TabIndex        =   124
            Top             =   1455
            Width           =   735
         End
         Begin VB.TextBox Path 
            Height          =   285
            Left            =   -71640
            TabIndex        =   125
            Top             =   1455
            Width           =   3735
         End
         Begin VB.Frame Frame1 
            Caption         =   "Proxy Settings"
            Height          =   1815
            Left            =   -74880
            TabIndex        =   150
            Top             =   3135
            Width           =   4215
            Begin VB.CheckBox ep 
               Caption         =   "Enable Proxy"
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   1440
               Width           =   2535
            End
            Begin VB.CheckBox bpsfla 
               Caption         =   "Bypass Proxy server for local addresses"
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   1080
               Width           =   3255
            End
            Begin VB.TextBox psp 
               Height          =   285
               Left            =   3240
               TabIndex        =   93
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox psa 
               Height          =   285
               Left            =   120
               TabIndex        =   91
               Top             =   600
               Width           =   2895
            End
            Begin VB.Label Label8 
               Caption         =   "Port i.e. (80)"
               Height          =   255
               Left            =   3240
               TabIndex        =   92
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label7 
               Caption         =   "Proxy Server Address i.e. (http://10.0.0.1)"
               Height          =   255
               Left            =   120
               TabIndex        =   90
               Top             =   360
               Width           =   3015
            End
         End
         Begin VB.CheckBox dicw 
            Caption         =   "Disable Internet Connection Wizard"
            Height          =   375
            Left            =   -74880
            TabIndex        =   89
            Top             =   2535
            Width           =   5535
         End
         Begin VB.TextBox sep 
            Height          =   285
            Left            =   -74880
            TabIndex        =   88
            Top             =   2160
            Width           =   4215
         End
         Begin VB.TextBox stp 
            Height          =   285
            Left            =   -74880
            TabIndex        =   86
            Top             =   1575
            Width           =   4215
         End
         Begin VB.ListBox DM 
            Columns         =   2
            Height          =   840
            ItemData        =   "Form1.frx":07F1
            Left            =   -72600
            List            =   "Form1.frx":07F3
            Sorted          =   -1  'True
            TabIndex        =   129
            Top             =   2280
            Width           =   4695
         End
         Begin VB.CommandButton DelMap 
            Caption         =   "Delete selected Map"
            Height          =   255
            Left            =   -71400
            TabIndex        =   130
            Top             =   3240
            Width           =   2415
         End
         Begin VB.CheckBox oepwsn 
            Caption         =   "&Overwrite existing profile with same name"
            Height          =   255
            Left            =   -74760
            TabIndex        =   99
            Top             =   855
            Width           =   3615
         End
         Begin VB.TextBox pn 
            Height          =   285
            Left            =   -73680
            TabIndex        =   101
            Top             =   1215
            Width           =   3375
         End
         Begin VB.TextBox mn 
            Height          =   285
            Left            =   -73680
            TabIndex        =   103
            Top             =   1575
            Width           =   3375
         End
         Begin VB.TextBox Exchange2 
            Height          =   285
            Left            =   -73440
            TabIndex        =   105
            Top             =   1935
            Width           =   1935
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Browse"
            Height          =   255
            Left            =   -71400
            TabIndex        =   106
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CheckBox ptewoe 
            Caption         =   "Prompt to empty &Wastebasket on Exit"
            Height          =   255
            Left            =   -74760
            TabIndex        =   107
            Top             =   2535
            Width           =   3855
         End
         Begin VB.Frame Frame6 
            Caption         =   "PAB & &Personal Folder Settings"
            Height          =   2055
            Left            =   -70080
            TabIndex        =   149
            Top             =   855
            Width           =   4215
            Begin VB.TextBox PFPath 
               Height          =   285
               Left            =   600
               TabIndex        =   109
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton Command10 
               Caption         =   "Browse"
               Height          =   255
               Left            =   3240
               TabIndex        =   110
               Top             =   240
               Width           =   855
            End
            Begin VB.CheckBox epab 
               Caption         =   "Enable Personal Address Book"
               Height          =   255
               Left            =   120
               TabIndex        =   111
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox PABFile 
               Height          =   285
               Left            =   1320
               TabIndex        =   113
               Top             =   960
               Width           =   2775
            End
            Begin VB.CheckBox epf 
               Caption         =   "Enable Personal Folders"
               Height          =   255
               Left            =   120
               TabIndex        =   114
               Top             =   1320
               Width           =   3615
            End
            Begin VB.TextBox PSTFile 
               Height          =   285
               Left            =   1320
               TabIndex        =   116
               Top             =   1680
               Width           =   2775
            End
            Begin VB.Label Label32 
               Caption         =   "Folder"
               Height          =   255
               Left            =   120
               TabIndex        =   108
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label33 
               Caption         =   "PAB File Name"
               Height          =   255
               Left            =   120
               TabIndex        =   112
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label34 
               Caption         =   "PST File Name"
               Height          =   255
               Left            =   120
               TabIndex        =   115
               Top             =   1680
               Width           =   1335
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "&Offline Access Settings"
            Height          =   1815
            Left            =   -70080
            TabIndex        =   148
            Top             =   3135
            Width           =   4215
            Begin VB.TextBox OfflinePath 
               Height          =   285
               Left            =   600
               TabIndex        =   118
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton Command11 
               Caption         =   "Browse"
               Height          =   255
               Left            =   3240
               TabIndex        =   119
               Top             =   240
               Width           =   855
            End
            Begin VB.CheckBox eof1 
               Caption         =   "Enable Offline Folders"
               Height          =   255
               Left            =   120
               TabIndex        =   120
               Top             =   840
               Width           =   2895
            End
            Begin VB.TextBox OSTFile 
               Height          =   285
               Left            =   1320
               TabIndex        =   122
               Top             =   1200
               Width           =   2775
            End
            Begin VB.Label Label35 
               Caption         =   "Folder"
               Height          =   255
               Left            =   120
               TabIndex        =   117
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label36 
               Caption         =   "OST File Name"
               Height          =   255
               Left            =   120
               TabIndex        =   121
               Top             =   1200
               Width           =   1215
            End
         End
         Begin VB.Label Label48 
            Caption         =   "Home Drive Root"
            Height          =   255
            Left            =   -70560
            TabIndex        =   176
            Top             =   4080
            Width           =   2655
         End
         Begin VB.Label Label47 
            Caption         =   "Home Drive"
            Height          =   255
            Left            =   -72600
            TabIndex        =   175
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label Label46 
            Caption         =   "Path"
            Height          =   255
            Left            =   -72480
            TabIndex        =   171
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label45 
            Caption         =   "Path to 'My Documents'"
            Height          =   255
            Left            =   -71880
            TabIndex        =   96
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label Label19 
            Caption         =   "Time (seconds)"
            Enabled         =   0   'False
            Height          =   255
            Left            =   -69600
            TabIndex        =   83
            Top             =   4575
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "Default option"
            Enabled         =   0   'False
            Height          =   255
            Left            =   -69600
            TabIndex        =   80
            Top             =   3735
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "Default option"
            Enabled         =   0   'False
            Height          =   255
            Left            =   -69600
            TabIndex        =   77
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "Time (milliseconds)"
            Enabled         =   0   'False
            Height          =   255
            Left            =   -69600
            TabIndex        =   74
            Top             =   2055
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Search Page"
            Height          =   255
            Left            =   -74880
            TabIndex        =   87
            Top             =   1935
            Width           =   4215
         End
         Begin VB.Label Label5 
            Caption         =   "Start Page"
            Height          =   255
            Left            =   -74880
            TabIndex        =   85
            Top             =   1320
            Width           =   4215
         End
         Begin VB.Label Label22 
            Caption         =   "Drive"
            Height          =   255
            Left            =   -72600
            TabIndex        =   127
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label23 
            Caption         =   "Path"
            Height          =   255
            Left            =   -71880
            TabIndex        =   128
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label28 
            Caption         =   "P&rofile Name"
            Height          =   255
            Left            =   -74760
            TabIndex        =   100
            Top             =   1215
            Width           =   975
         End
         Begin VB.Label Label29 
            Caption         =   "&Mailbox Name"
            Height          =   255
            Left            =   -74760
            TabIndex        =   102
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label31 
            Caption         =   "&Exchange Server"
            Height          =   255
            Left            =   -74760
            TabIndex        =   104
            Top             =   1935
            Width           =   1335
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   5175
         Left            =   120
         TabIndex        =   157
         Top             =   840
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   9128
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   397
         WordWrap        =   0   'False
         OLEDropMode     =   1
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "Form1.frx":07F5
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame8"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame9"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Windows NT Network"
         TabPicture(1)   =   "Form1.frx":0811
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chdss"
         Tab(1).Control(1)=   "chdsw"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Windows NT System"
         TabPicture(2)   =   "Form1.frx":082D
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "dl"
         Tab(2).Control(1)=   "swtal"
         Tab(2).Control(2)=   "dcp"
         Tab(2).Control(3)=   "dlw"
         Tab(2).Control(4)=   "dtm"
         Tab(2).Control(5)=   "rlss"
         Tab(2).Control(6)=   "esfadb"
         Tab(2).Control(7)=   "dndllou"
         Tab(2).Control(8)=   "Frame3"
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "Desktop"
         TabPicture(3)   =   "Form1.frx":0849
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label3"
         Tab(3).Control(1)=   "Label4"
         Tab(3).Control(2)=   "tw"
         Tab(3).Control(3)=   "Command1"
         Tab(3).Control(4)=   "wn"
         Tab(3).ControlCount=   5
         Begin VB.Frame Frame9 
            Caption         =   "Save Settings To"
            Height          =   2055
            Left            =   360
            TabIndex        =   160
            Top             =   2445
            Width           =   8775
            Begin VB.CommandButton Command15 
               Caption         =   "Save Machine Template"
               Height          =   375
               Left            =   4440
               TabIndex        =   22
               Top             =   960
               Width           =   3375
            End
            Begin VB.ComboBox SaveMachine 
               Height          =   315
               Left            =   240
               Sorted          =   -1  'True
               TabIndex        =   21
               Top             =   1080
               Width           =   3135
            End
            Begin VB.Label Label2 
               Caption         =   "Machine"
               Height          =   255
               Left            =   1320
               TabIndex        =   20
               Top             =   840
               Width           =   735
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Load Settings From"
            Height          =   1815
            Left            =   360
            TabIndex        =   159
            Top             =   525
            Width           =   8775
            Begin VB.CommandButton Command14 
               Caption         =   "Load Machine Template"
               Height          =   375
               Left            =   4440
               TabIndex        =   19
               Top             =   840
               Width           =   3375
            End
            Begin VB.ComboBox LoadMachine 
               Height          =   315
               Left            =   240
               Sorted          =   -1  'True
               TabIndex        =   18
               Top             =   840
               Width           =   3135
            End
            Begin VB.Label Label1 
               Caption         =   "Machine"
               Height          =   255
               Left            =   1320
               TabIndex        =   17
               Top             =   600
               Width           =   735
            End
         End
         Begin VB.CheckBox chdss 
            Caption         =   "Create hidden drive shares (server)"
            Height          =   255
            Left            =   -71640
            TabIndex        =   24
            Top             =   2565
            Width           =   4095
         End
         Begin VB.CheckBox chdsw 
            Caption         =   "Create hidden drive shares (workstation)"
            Height          =   255
            Left            =   -71640
            TabIndex        =   23
            Top             =   2085
            Width           =   4095
         End
         Begin VB.CheckBox dl 
            Caption         =   "Disable Logoff"
            Height          =   255
            Left            =   -74520
            TabIndex        =   26
            Top             =   885
            Width           =   3855
         End
         Begin VB.CheckBox swtal 
            Caption         =   "Show welcome tips at logon"
            Height          =   255
            Left            =   -70320
            TabIndex        =   31
            Top             =   1245
            Width           =   3855
         End
         Begin VB.CheckBox dcp 
            Caption         =   "Disable Change Password"
            Height          =   255
            Left            =   -70320
            TabIndex        =   30
            Top             =   885
            Width           =   3855
         End
         Begin VB.CheckBox dlw 
            Caption         =   "Disable Lock Workstation"
            Height          =   255
            Left            =   -70320
            TabIndex        =   29
            Top             =   525
            Width           =   3855
         End
         Begin VB.CheckBox dtm 
            Caption         =   "Disable Task Manager"
            Height          =   255
            Left            =   -74520
            TabIndex        =   27
            Top             =   1245
            Width           =   3855
         End
         Begin VB.CheckBox rlss 
            Caption         =   "Run logon scripts synchronously."
            Height          =   255
            Left            =   -74520
            TabIndex        =   25
            Top             =   525
            Width           =   3855
         End
         Begin VB.CheckBox esfadb 
            Caption         =   "Enable shutdown from Authentication dialog box"
            Height          =   255
            Left            =   -74520
            TabIndex        =   28
            Top             =   1605
            Width           =   3735
         End
         Begin VB.CheckBox dndllou 
            Caption         =   "Do not display last logged on username"
            Height          =   255
            Left            =   -70320
            TabIndex        =   32
            Top             =   1605
            Width           =   3615
         End
         Begin VB.Frame Frame3 
            Caption         =   "Logon Banner"
            Height          =   1575
            Left            =   -74520
            TabIndex        =   158
            Top             =   2085
            Width           =   8775
            Begin VB.TextBox lbCaption 
               Height          =   285
               Left            =   120
               TabIndex        =   34
               Top             =   480
               Width           =   8535
            End
            Begin VB.TextBox lbText 
               Height          =   285
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   36
               Top             =   1200
               Width           =   8535
            End
            Begin VB.Label Label20 
               Caption         =   "Caption"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   240
               Width           =   8535
            End
            Begin VB.Label Label21 
               Caption         =   "Text"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   960
               Width           =   8535
            End
         End
         Begin VB.TextBox wn 
            Height          =   285
            Left            =   -72480
            TabIndex        =   38
            Top             =   1605
            Width           =   4215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Browse"
            Height          =   255
            Left            =   -68160
            TabIndex        =   39
            Top             =   1605
            Width           =   1215
         End
         Begin VB.CheckBox tw 
            Caption         =   "Tile Wallpaper"
            Height          =   255
            Left            =   -72480
            TabIndex        =   41
            Top             =   2565
            Width           =   3615
         End
         Begin VB.Label Label4 
            Caption         =   "Specify location and name (e.g. c:\winnt\winnt256.bmp)"
            Height          =   255
            Left            =   -72480
            TabIndex        =   40
            Top             =   2040
            Width           =   3975
         End
         Begin VB.Label Label3 
            Caption         =   "Wallpaper Name"
            Height          =   255
            Left            =   -72480
            TabIndex        =   37
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Wallpaper Name"
            Height          =   255
            Left            =   -72480
            TabIndex        =   162
            Top             =   1365
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Specify location and name (e.g. c:\winnt\winnt256.bmp)"
            Height          =   255
            Left            =   -72480
            TabIndex        =   161
            Top             =   2085
            Width           =   3975
         End
      End
      Begin VB.Label Label40 
         Caption         =   "Backup Domain Controller"
         Height          =   255
         Left            =   -72720
         TabIndex        =   10
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label Label39 
         Caption         =   "Backup Domain Controller"
         Height          =   255
         Left            =   -72720
         TabIndex        =   7
         Top             =   2880
         Width           =   6015
      End
      Begin VB.Label Label38 
         Caption         =   "Backup Domain Controller"
         Height          =   255
         Left            =   -72720
         TabIndex        =   4
         Top             =   2280
         Width           =   4935
      End
      Begin VB.Label Label37 
         Caption         =   "Primary Domain Controller"
         Height          =   255
         Left            =   -72720
         TabIndex        =   1
         Top             =   1680
         Width           =   5175
      End
      Begin VB.Label Label30 
         Caption         =   "Exchange Server"
         Height          =   255
         Left            =   -72720
         TabIndex        =   13
         Top             =   4080
         Width           =   3255
      End
   End
   Begin VB.Label Label44 
      Alignment       =   1  'Right Justify
      Caption         =   "Domain"
      Height          =   255
      Left            =   4920
      TabIndex        =   170
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      Caption         =   "Computer Name"
      Height          =   255
      Left            =   4920
      TabIndex        =   169
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label42 
      Alignment       =   1  'Right Justify
      Caption         =   "PDC"
      Height          =   255
      Left            =   120
      TabIndex        =   168
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   167
      Top             =   6600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
wn.Text = BrowseForFolder("")
End Sub

Private Sub Command10_Click()
PFPath.Text = BrowseForFolder("")
End Sub

Private Sub Command11_Click()
OfflinePath.Text = BrowseForFolder("")
End Sub

Private Sub Command12_Click()
LoadUser.Text = ""
SaveUser.Text = ""
LoadGroup.Text = ""
SaveGroup.Text = ""

run.Clear
DM.Clear
On Error Resume Next
Open DomainController.Text & "\admin$\system32\repl\import\scripts\Policies\GBLUser.ini" For Input As #1
Input #1, Temp
lpse.Value = Temp
Line Input #1, Temp
1
cm.Text = Temp
Input #1, Temp
If Abs(Temp / 1) <> Abs(Temp) Then GoTo 1
mps.Text = Temp
Input #1, Temp
irifl.Value = Temp
Input #1, Temp
nuwpssie.Value = Temp
Input #1, Temp
ruexm.Text = Temp
Input #1, Temp
pdirectory.Text = Temp
Input #1, Temp
dccorp.Value = Temp
Input #1, Temp
adsnc.Value = Temp
Input #1, Temp
snct.Value = Temp
Input #1, Temp
tm.Text = Temp
Input #1, Temp
sndpo.Value = Temp
Input #1, Temp
doc.Text = Temp
Input #1, Temp
cpdo.Value = Temp
Input #1, Temp
doc1.Text = Temp
Input #1, Temp
tfdb.Value = Temp
Input #1, Temp
ts.Text = Temp
Input #1, Temp
stp.Text = Temp
Input #1, Temp
sep.Text = Temp
Input #1, Temp
dicw.Value = Temp
Input #1, Temp
psa.Text = Temp
Input #1, Temp
psp.Text = Temp
Input #1, Temp
bpsfla.Value = Temp
Input #1, Temp
ep.Value = Temp
Input #1, Temp
md.Text = Temp
Input #1, Temp
oepwsn.Value = Temp
Input #1, Temp
pn.Text = Temp
Input #1, Temp
mn.Text = Temp
Input #1, Temp
Exchange2.Text = Temp
Input #1, Temp
ptewoe.Value = Temp
Input #1, Temp
PFPath.Text = Temp
Input #1, Temp
epab.Value = Temp
Input #1, Temp
PABFile.Text = Temp
Input #1, Temp
epf.Value = Temp
Input #1, Temp
PSTFile.Text = Temp
Input #1, Temp
OfflinePath.Text = Temp
Input #1, Temp
eof1.Value = Temp
Input #1, Temp
OSTFile.Text = Temp
Input #1, Temp
dret.Value = Temp
Input #1, Temp
datdi.Value = Temp
Input #1, Temp
hbt.Value = Temp
Input #1, Temp
hsst.Value = Temp
Input #1, Temp
hat.Value = Temp
Input #1, Temp
hst.Value = Temp
Input #1, Temp
rrcfsm.Value = Temp
Input #1, Temp
rffsosm.Value = Temp
Input #1, Temp
rtfsosm.Value = Temp
Input #1, Temp
rfcfsm.Value = Temp
Input #1, Temp
hdimc.Value = Temp
Input #1, Temp
hnn.Value = Temp
Input #1, Temp
neninn.Value = Temp
Input #1, Temp
haiod.Value = Temp
Input #1, Temp
rsdcfsm.Value = Temp
Input #1, Temp
rmnd.Value = Temp
Input #1, Temp
HomeDrive.Text = Temp
Input #1, Temp
HomeRoot.Text = Temp
Input #1, Temp
SSFile.Text = Temp
Input #1, Temp
Temp = Temp / 60
SSTime.Text = Temp
Scrollss.Value = Temp
Input #1, Temp

SSPassword.Value = Temp
Input #1, Temp
1000
Input #1, Temp
If Temp = "[Mappings]" Then GoTo 1001
run.AddItem Temp
GoTo 1000
1001
Do While Not EOF(1)
Input #1, Temp
DM.AddItem Temp
Loop
Close
MsgBox "User Template Loaded."
Close
End Sub

Private Sub Command13_Click()
SSTime1 = Abs(SSTime.Text) * 60
Open PDC.Text & "\admin$\system32\repl\import\scripts\Policies\GBLUser.ini" For Output As #1
Print #1, lpse.Value
Print #1, cm.Text
Print #1, mps.Text
Print #1, irifl.Value
Print #1, nuwpssie.Value
Print #1, ruexm.Text
Print #1, pdirectory.Text
Print #1, dccorp.Value
Print #1, adsnc.Value
Print #1, snct.Value
Print #1, tm.Text
Print #1, sndpo.Value
Print #1, doc.Text
Print #1, cpdo.Value
Print #1, doc1.Text
Print #1, tfdb.Value
Print #1, ts.Text
Print #1, stp.Text
Print #1, sep.Text
Print #1, dicw.Value
Print #1, psa.Text
Print #1, psp.Text
Print #1, bpsfla.Value
Print #1, ep.Value
Print #1, md.Text
Print #1, oepwsn.Value
Print #1, pn.Text
Print #1, mn.Text
Print #1, Exchange2.Text
Print #1, ptewoe.Value
Print #1, PFPath.Text
Print #1, epab.Value
Print #1, PABFile.Text
Print #1, epf.Value
Print #1, PSTFile.Text
Print #1, OfflinePath.Text
Print #1, eof1.Value
Print #1, OSTFile.Text
Print #1, dret.Value
Print #1, datdi.Value
Print #1, hbt.Value
Print #1, hsst.Value
Print #1, hat.Value
Print #1, hst.Value
Print #1, rrcfsm.Value
Print #1, rffsosm.Value
Print #1, rtfsosm.Value
Print #1, rfcfsm.Value
Print #1, hdimc.Value
Print #1, hnn.Value
Print #1, neninn.Value
Print #1, haiod.Value
Print #1, rsdcfsm.Value
Print #1, rmnd.Value
Print #1, HomeDrive.Text
Print #1, HomeRoot.Text
Print #1, SSFile.Text
Print #1, SSTime1
Print #1, SSPassword.Value
Print #1, "[RUN]"
For X = 0 To run.ListCount - 1
Print #1, run.List(X)
Next X
Print #1, "[Mappings]"
For X = 0 To DM.ListCount - 1
Print #1, DM.List(X)
Next X
Close

If Not BDC1.Text = "" Then
Open BDC1.Text & "\admin$\system32\repl\import\scripts\Policies\GBLUser.ini" For Output As #1
Print #1, lpse.Value
Print #1, cm.Text
Print #1, mps.Text
Print #1, irifl.Value
Print #1, nuwpssie.Value
Print #1, ruexm.Text
Print #1, pdirectory.Text
Print #1, dccorp.Value
Print #1, adsnc.Value
Print #1, snct.Value
Print #1, tm.Text
Print #1, sndpo.Value
Print #1, doc.Text
Print #1, cpdo.Value
Print #1, doc1.Text
Print #1, tfdb.Value
Print #1, ts.Text
Print #1, stp.Text
Print #1, sep.Text
Print #1, dicw.Value
Print #1, psa.Text
Print #1, psp.Text
Print #1, bpsfla.Value
Print #1, ep.Value
Print #1, md.Text
Print #1, oepwsn.Value
Print #1, pn.Text
Print #1, mn.Text
Print #1, Exchange2.Text
Print #1, ptewoe.Value
Print #1, PFPath.Text
Print #1, epab.Value
Print #1, PABFile.Text
Print #1, epf.Value
Print #1, PSTFile.Text
Print #1, OfflinePath.Text
Print #1, eof1.Value
Print #1, OSTFile.Text
Print #1, dret.Value
Print #1, datdi.Value
Print #1, hbt.Value
Print #1, hsst.Value
Print #1, hat.Value
Print #1, hst.Value
Print #1, rrcfsm.Value
Print #1, rffsosm.Value
Print #1, rtfsosm.Value
Print #1, rfcfsm.Value
Print #1, hdimc.Value
Print #1, hnn.Value
Print #1, neninn.Value
Print #1, haiod.Value
Print #1, rsdcfsm.Value
Print #1, rmnd.Value
Print #1, HomeDrive.Text
Print #1, HomeRoot.Text
Print #1, SSFile.Text
Print #1, SSTime1
Print #1, SSPassword.Value
Print #1, "[RUN]"
For X = 0 To run.ListCount - 1
Print #1, run.List(X)
Next X
Print #1, "[Mappings]"
For X = 0 To DM.ListCount - 1
Print #1, DM.List(X)
Next X
Close
End If

If Not BDC2.Text = "" Then
Open BDC2.Text & "\admin$\system32\repl\import\scripts\Policies\GBLUser.ini" For Output As #1
Print #1, lpse.Value
Print #1, cm.Text
Print #1, mps.Text
Print #1, irifl.Value
Print #1, nuwpssie.Value
Print #1, ruexm.Text
Print #1, pdirectory.Text
Print #1, dccorp.Value
Print #1, adsnc.Value
Print #1, snct.Value
Print #1, tm.Text
Print #1, sndpo.Value
Print #1, doc.Text
Print #1, cpdo.Value
Print #1, doc1.Text
Print #1, tfdb.Value
Print #1, ts.Text
Print #1, stp.Text
Print #1, sep.Text
Print #1, dicw.Value
Print #1, psa.Text
Print #1, psp.Text
Print #1, bpsfla.Value
Print #1, ep.Value
Print #1, md.Text
Print #1, oepwsn.Value
Print #1, pn.Text
Print #1, mn.Text
Print #1, Exchange2.Text
Print #1, ptewoe.Value
Print #1, PFPath.Text
Print #1, epab.Value
Print #1, PABFile.Text
Print #1, epf.Value
Print #1, PSTFile.Text
Print #1, OfflinePath.Text
Print #1, eof1.Value
Print #1, OSTFile.Text
Print #1, dret.Value
Print #1, datdi.Value
Print #1, hbt.Value
Print #1, hsst.Value
Print #1, hat.Value
Print #1, hst.Value
Print #1, rrcfsm.Value
Print #1, rffsosm.Value
Print #1, rtfsosm.Value
Print #1, rfcfsm.Value
Print #1, hdimc.Value
Print #1, hnn.Value
Print #1, neninn.Value
Print #1, haiod.Value
Print #1, rsdcfsm.Value
Print #1, rmnd.Value
Print #1, HomeDrive.Text
Print #1, HomeRoot.Text
Print #1, SSFile.Text
Print #1, SSTime1
Print #1, SSPassword.Value
Print #1, "[RUN]"
For X = 0 To run.ListCount - 1
Print #1, run.List(X)
Next X
Print #1, "[Mappings]"
For X = 0 To DM.ListCount - 1
Print #1, DM.List(X)
Next X
Close
End If

If Not BDC3.Text = "" Then
Open BDC3.Text & "\admin$\system32\repl\import\scripts\Policies\GBLUser.ini" For Output As #1
Print #1, lpse.Value
Print #1, cm.Text
Print #1, mps.Text
Print #1, irifl.Value
Print #1, nuwpssie.Value
Print #1, ruexm.Text
Print #1, pdirectory.Text
Print #1, dccorp.Value
Print #1, adsnc.Value
Print #1, snct.Value
Print #1, tm.Text
Print #1, sndpo.Value
Print #1, doc.Text
Print #1, cpdo.Value
Print #1, doc1.Text
Print #1, tfdb.Value
Print #1, ts.Text
Print #1, stp.Text
Print #1, sep.Text
Print #1, dicw.Value
Print #1, psa.Text
Print #1, psp.Text
Print #1, bpsfla.Value
Print #1, ep.Value
Print #1, md.Text
Print #1, oepwsn.Value
Print #1, pn.Text
Print #1, mn.Text
Print #1, Exchange2.Text
Print #1, ptewoe.Value
Print #1, PFPath.Text
Print #1, epab.Value
Print #1, PABFile.Text
Print #1, epf.Value
Print #1, PSTFile.Text
Print #1, OfflinePath.Text
Print #1, eof1.Value
Print #1, OSTFile.Text
Print #1, dret.Value
Print #1, datdi.Value
Print #1, hbt.Value
Print #1, hsst.Value
Print #1, hat.Value
Print #1, hst.Value
Print #1, rrcfsm.Value
Print #1, rffsosm.Value
Print #1, rtfsosm.Value
Print #1, rfcfsm.Value
Print #1, hdimc.Value
Print #1, hnn.Value
Print #1, neninn.Value
Print #1, haiod.Value
Print #1, rsdcfsm.Value
Print #1, rmnd.Value
Print #1, HomeDrive.Text
Print #1, HomeRoot.Text
Print #1, SSFile.Text
Print #1, SSTime1
Print #1, SSPassword.Value
Print #1, "[RUN]"
For X = 0 To run.ListCount - 1
Print #1, run.List(X)
Next X
Print #1, "[Mappings]"
For X = 0 To DM.ListCount - 1
Print #1, DM.List(X)
Next X
Close
End If
LoadUser.Text = ""
SaveUser.Text = ""
LoadGroup.Text = ""
SaveGroup.Text = ""
MsgBox "User Template Saved."
End Sub

Private Sub Command14_Click()
On Error GoTo 1203
Open DomainController.Text & "\admin$\system32\repl\import\scripts\Policies\GBLMachine.ini" For Input As #1
Input #1, Temp
chdsw.Value = Temp
Input #1, Temp
chdss.Value = Temp
Input #1, Temp
rlss.Value = Temp
Input #1, Temp
dl.Value = Temp
Input #1, Temp
dtm.Value = Temp
Input #1, Temp
esfadb.Value = Temp
Input #1, Temp
dlw.Value = Temp
Input #1, Temp
dcp.Value = Temp
Input #1, Temp
swtal.Value = Temp
Input #1, Temp
dndllou.Value = Temp
Line Input #1, Temp
lbCaption.Text = Temp
Line Input #1, Temp
lbText.Text = Temp
Input #1, Temp
wn.Text = Temp
Input #1, Temp
tw.Value = Temp
Close
MsgBox "Machine Template loaded."
GoTo 1202
1203 MsgBox "No Machine Template Created."
1202
End Sub

Private Sub Command15_Click()
Open PDC.Text & "\admin$\system32\repl\import\scripts\Policies\GBLMachine.ini" For Output As #1
Print #1, chdsw.Value
Print #1, chdss.Value
Print #1, rlss.Value
Print #1, dl.Value
Print #1, dtm.Value
Print #1, esfadb.Value
Print #1, dlw.Value
Print #1, dcp.Value
Print #1, swtal.Value
Print #1, dndllou.Value
Print #1, lbCaption.Text
Print #1, lbText.Text
Print #1, wn.Text
Print #1, tw.Value
Close

If Not BDC1.Text = "" Then
    Open BDC1.Text & "\admin$\system32\repl\import\scripts\Policies\GBLMachine.ini" For Output As #1
Print #1, chdsw.Value
Print #1, chdss.Value
Print #1, rlss.Value
Print #1, dl.Value
Print #1, dtm.Value
Print #1, esfadb.Value
Print #1, dlw.Value
Print #1, dcp.Value
Print #1, swtal.Value
Print #1, dndllou.Value
Print #1, lbCaption.Text
Print #1, lbText.Text
Print #1, wn.Text
Print #1, tw.Value
Close
End If

If Not BDC2.Text = "" Then
    Open BDC2.Text & "\admin$\system32\repl\import\scripts\Policies\GBLMachine.ini" For Output As #1
Print #1, chdsw.Value
Print #1, chdss.Value
Print #1, rlss.Value
Print #1, dl.Value
Print #1, dtm.Value
Print #1, esfadb.Value
Print #1, dlw.Value
Print #1, dcp.Value
Print #1, swtal.Value
Print #1, dndllou.Value
Print #1, lbCaption.Text
Print #1, lbText.Text
Print #1, wn.Text
Print #1, tw.Value
Close
End If

If Not BDC3.Text = "" Then
    Open BDC3.Text & "\admin$\system32\repl\import\scripts\Policies\GBLMachine.ini" For Output As #1
Print #1, chdsw.Value
Print #1, chdss.Value
Print #1, rlss.Value
Print #1, dl.Value
Print #1, dtm.Value
Print #1, esfadb.Value
Print #1, dlw.Value
Print #1, dcp.Value
Print #1, swtal.Value
Print #1, dndllou.Value
Print #1, lbCaption.Text
Print #1, lbText.Text
Print #1, wn.Text
Print #1, tw.Value
Close
End If


MsgBox "Machine Template Saved."

End Sub

Private Sub Command16_Click()
Open PDC.Text & "\Admin$\system32\repl\import\scripts\policies\Global.ini" For Output As #1
Print #1, PDC.Text
Print #1, BDC1.Text
Print #1, BDC2.Text
Print #1, BDC3.Text
Print #1, Exchange.Text
Close

If BDC1.Text <> "" Then
Open BDC1.Text & "\Admin$\system32\repl\import\scripts\policies\Global.ini" For Output As #1
Print #1, PDC.Text
Print #1, BDC1.Text
Print #1, BDC2.Text
Print #1, BDC3.Text
Print #1, Exchange.Text
Close
End If

If BDC2.Text <> "" Then
Open BDC1.Text & "\Admin$\system32\repl\import\scripts\policies\Global.ini" For Output As #1
Print #1, PDC.Text
Print #1, BDC1.Text
Print #1, BDC2.Text
Print #1, BDC3.Text
Print #1, Exchange.Text
Close
End If

If BDC2.Text <> "" Then
Open BDC1.Text & "\Admin$\system32\repl\import\scripts\policies\Global.ini" For Output As #1
Print #1, PDC.Text
Print #1, BDC1.Text
Print #1, BDC2.Text
Print #1, BDC3.Text
Print #1, Exchange.Text
Close
End If

MsgBox "Global Defaults Saved."
End Sub

Private Sub Command17_Click()
md.Text = BrowseForFolder("")
End Sub

Private Sub Command18_Click()
run.AddItem rp.Text
rp.Text = ""

End Sub

Private Sub Command19_Click()
rp.Text = BrowseForFolder("")
End Sub

Private Sub Command2_Click()
Exchange.Text = BrowseForFolder("")
End Sub

Private Sub Command20_Click()
On Error Resume Next
run.RemoveItem (run.ListIndex)

End Sub

Private Sub Command21_Click()
HomeRoot.Text = BrowseForFolder("")
End Sub

Private Sub Command22_Click()
SSFile = BrowseForFolder("")
End Sub

Private Sub Command3_Click()
PDC.Text = BrowseForFolder("")
End Sub

Private Sub Command4_Click()
BDC1.Text = BrowseForFolder("")
End Sub

Private Sub Command5_Click()
BDC2.Text = BrowseForFolder("")
End Sub

Private Sub Command6_Click()
BDC3.Text = BrowseForFolder("")
End Sub

Private Sub Command7_Click()
Path.Text = BrowseForFolder("")
End Sub

Private Sub Command8_Click()
Mapping = Drive.Text & Chr(9) & Path.Text
DM.AddItem (Mapping)
Drive.Text = ""
Path.Text = ""

End Sub

Private Sub Command9_Click()
PDC.Text = BrowseForFolder("")
End Sub

Private Sub cpdo_Click()
If cpdo.Value = 1 Then
    Label18.Enabled = True
    doc1.Enabled = True
End If
If cpdo.Value = 0 Then
    Label18.Enabled = False
    doc1.Enabled = False
End If
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub DelMap_Click()
DM.RemoveItem (DM.ListIndex)



End Sub

Private Sub Form_Load()
App.TaskVisible = False

SSTab1.Tab = "2"
SSTab2.Tab = "8"
SSTab3.Tab = "0"


Set AU = CreateObject("Persits.ASPUser")
username.Text = AU.GetUserName
DomainName.Text = AU.GetDomainName
Compname.Text = AU.GetComputerName
DomainController.Text = AU.DomainController
AU.Server = DomainController.Text

If AU.Users(username.Text).Groups("Domain Admins") = "" Then
    MsgBox "You do not have suffecient permissions to run this program. You must be in the 'Domain Admins' group."
    End
    End If
   
On Error Resume Next
Open DomainController.Text & "\netlogon\policies\Global.ini" For Input As #1
Input #1, Temp
PDC.Text = Temp
Input #1, Temp
BDC1.Text = Temp
Input #1, Temp
BDC2.Text = Temp
Input #1, Temp
BDC3.Text = Temp
Input #1, Temp
Exchange.Text = Temp
Close

For Each SRV In AU.Servers
    LoadMachine.AddItem (SRV)
    SaveMachine.AddItem (SRV)
    Next

For Each User In AU.Users
    LoadUser.AddItem (User.Name)
    SaveUser.AddItem (User.Name)
Next

For Each Group In AU.Groups
LoadGroup.AddItem (Group.Name)
SaveGroup.AddItem (Group.Name)
Next
For Each Group In AU.LocalGroups
LoadGroup.AddItem (Group.Name)
SaveGroup.AddItem (Group.Name)
Next

End Sub


Private Sub Scroll_Change()
mps.Text = ScrollKB.Value

End Sub

Private Sub LoadGroup_Click()

run.Clear
DM.Clear

On Error GoTo 10001
Open DomainController.Text & "\admin$\system32\repl\import\scripts\Policies\" & LoadGroup & ".ini" For Input As #1
Input #1, Temp
lpse.Value = Temp
Line Input #1, Temp
1
cm.Text = Temp
Input #1, Temp
If Abs(Temp / 1) <> Abs(Temp) Then GoTo 1
mps.Text = Temp
Input #1, Temp
irifl.Value = Temp
Input #1, Temp
nuwpssie.Value = Temp
Input #1, Temp
ruexm.Text = Temp
Input #1, Temp
pdirectory.Text = Temp
Input #1, Temp
dccorp.Value = Temp
Input #1, Temp
adsnc.Value = Temp
Input #1, Temp
snct.Value = Temp
Input #1, Temp
tm.Text = Temp
Input #1, Temp
sndpo.Value = Temp
Input #1, Temp
doc.Text = Temp
Input #1, Temp
cpdo.Value = Temp
Input #1, Temp
doc1.Text = Temp
Input #1, Temp
tfdb.Value = Temp
Input #1, Temp
ts.Text = Temp
Input #1, Temp
stp.Text = Temp
Input #1, Temp
sep.Text = Temp
Input #1, Temp
dicw.Value = Temp
Input #1, Temp
psa.Text = Temp
Input #1, Temp
psp.Text = Temp
Input #1, Temp
bpsfla.Value = Temp
Input #1, Temp
ep.Value = Temp
Input #1, Temp
md.Text = Temp
Input #1, Temp
oepwsn.Value = Temp
Input #1, Temp
pn.Text = Temp
Input #1, Temp
mn.Text = Temp
Input #1, Temp
Exchange2.Text = Temp
Input #1, Temp
ptewoe.Value = Temp
Input #1, Temp
PFPath.Text = Temp
Input #1, Temp
epab.Value = Temp
Input #1, Temp
PABFile.Text = Temp
Input #1, Temp
epf.Value = Temp
Input #1, Temp
PSTFile.Text = Temp
Input #1, Temp
OfflinePath.Text = Temp
Input #1, Temp
eof1.Value = Temp
Input #1, Temp
OSTFile.Text = Temp
Input #1, Temp
dret.Value = Temp
Input #1, Temp
datdi.Value = Temp
Input #1, Temp
hbt.Value = Temp
Input #1, Temp
hsst.Value = Temp
Input #1, Temp
hat.Value = Temp
Input #1, Temp
hst.Value = Temp
Input #1, Temp
rrcfsm.Value = Temp
Input #1, Temp
rffsosm.Value = Temp
Input #1, Temp
rtfsosm.Value = Temp
Input #1, Temp
rfcfsm.Value = Temp
Input #1, Temp
hdimc.Value = Temp
Input #1, Temp
hnn.Value = Temp
Input #1, Temp
neninn.Value = Temp
Input #1, Temp
haiod.Value = Temp
Input #1, Temp
rsdcfsm.Value = Temp
Input #1, Temp
rmnd.Value = Temp
Input #1, Temp
HomeDrive.Text = Temp
Input #1, Temp
HomeRoot.Text = Temp
Input #1, Temp
SSFile.Text = Temp
Input #1, Temp
Temp = Temp / 60
Scrollss.Value = Temp
SSTime.Text = Temp
Input #1, Temp
SSPassword.Value = Temp
Input #1, Temp
1000
Input #1, Temp
If Temp = "[Mappings]" Then GoTo 1001
run.AddItem Temp
GoTo 1000
1001
Do While Not EOF(1)
Input #1, Temp
DM.AddItem Temp
Loop
Close
MsgBox "Group '" & LoadGroup & "' Loaded."
GoTo 10002
10001
MsgBox "No Template for group '" & LoadGroup & "'."
10002
Close
LoadUser.Text = ""
SaveUser.Text = ""
LoadGroup.Text = ""
SaveGroup.Text = ""
End Sub

Private Sub LoadMachine_Click()
On Error GoTo 1201
If LoadMachine.Text = "" Then GoTo 1201
Open DomainController.Text & "\admin$\system32\repl\import\scripts\Policies\" & LoadMachine.Text & ".ini" For Input As #1
Input #1, Temp
chdsw.Value = Temp
Input #1, Temp
chdss.Value = Temp
Input #1, Temp
rlss.Value = Temp
Input #1, Temp
dl.Value = Temp
Input #1, Temp
dtm.Value = Temp
Input #1, Temp
esfadb.Value = Temp
Input #1, Temp
dlw.Value = Temp
Input #1, Temp
dcp.Value = Temp
Input #1, Temp
swtal.Value = Temp
Input #1, Temp
dndllou.Value = Temp
Input #1, Temp
lbCaption.Text = Temp
Input #1, Temp
lbText.Text = Temp
Input #1, Temp
wn.Text = Temp
Input #1, Temp
tw.Value = Temp
Close
MsgBox "Machine '" & LoadMachine.Text & "' loaded."
1201
End Sub

Private Sub lps_Click()
If lps.Value = "0" Then
    Frame2.Enabled = False
End If
If lps.Value = "1" Then
    Frame2.Enabled = True
End If

End Sub

Private Sub LoadUser_Click()

run.Clear
DM.Clear

On Error GoTo 10001
Open DomainController.Text & "\admin$\system32\repl\import\scripts\Policies\" & LoadUser & ".ini" For Input As #1
Input #1, Temp
lpse.Value = Temp
Line Input #1, Temp
1
cm.Text = Temp
Input #1, Temp
If Abs(Temp / 1) <> Abs(Temp) Then GoTo 1
mps.Text = Temp
Input #1, Temp
irifl.Value = Temp
Input #1, Temp
nuwpssie.Value = Temp
Input #1, Temp
ruexm.Text = Temp
Input #1, Temp
pdirectory.Text = Temp
Input #1, Temp
dccorp.Value = Temp
Input #1, Temp
adsnc.Value = Temp
Input #1, Temp
snct.Value = Temp
Input #1, Temp
tm.Text = Temp
Input #1, Temp
sndpo.Value = Temp
Input #1, Temp
doc.Text = Temp
Input #1, Temp
cpdo.Value = Temp
Input #1, Temp
doc1.Text = Temp
Input #1, Temp
tfdb.Value = Temp
Input #1, Temp
ts.Text = Temp
Input #1, Temp
stp.Text = Temp
Input #1, Temp
sep.Text = Temp
Input #1, Temp
dicw.Value = Temp
Input #1, Temp
psa.Text = Temp
Input #1, Temp
psp.Text = Temp
Input #1, Temp
bpsfla.Value = Temp
Input #1, Temp
ep.Value = Temp
Input #1, Temp
md.Text = Temp
Input #1, Temp
oepwsn.Value = Temp
Input #1, Temp
pn.Text = Temp
Input #1, Temp
mn.Text = Temp
Input #1, Temp
Exchange2.Text = Temp
Input #1, Temp
ptewoe.Value = Temp
Input #1, Temp
PFPath.Text = Temp
Input #1, Temp
epab.Value = Temp
Input #1, Temp
PABFile.Text = Temp
Input #1, Temp
epf.Value = Temp
Input #1, Temp
PSTFile.Text = Temp
Input #1, Temp
OfflinePath.Text = Temp
Input #1, Temp
eof1.Value = Temp
Input #1, Temp
OSTFile.Text = Temp
Input #1, Temp
dret.Value = Temp
Input #1, Temp
datdi.Value = Temp
Input #1, Temp
hbt.Value = Temp
Input #1, Temp
hsst.Value = Temp
Input #1, Temp
hat.Value = Temp
Input #1, Temp
hst.Value = Temp
Input #1, Temp
rrcfsm.Value = Temp
Input #1, Temp
rffsosm.Value = Temp
Input #1, Temp
rtfsosm.Value = Temp
Input #1, Temp
rfcfsm.Value = Temp
Input #1, Temp
hdimc.Value = Temp
Input #1, Temp
hnn.Value = Temp
Input #1, Temp
neninn.Value = Temp
Input #1, Temp
haiod.Value = Temp
Input #1, Temp
rsdcfsm.Value = Temp
Input #1, Temp
rmnd.Value = Temp
Input #1, Temp
HomeDrive.Text = Temp
Input #1, Temp
HomeRoot.Text = Temp
Input #1, Temp
SSFile.Text = Temp
Input #1, Temp
Temp = Temp / 60
Scrollss.Value = Temp
SSTime.Text = Temp
Input #1, Temp
SSPassword.Value = Temp
Input #1, Temp
1000
Input #1, Temp
If Temp = "[Mappings]" Then GoTo 1001
run.AddItem Temp
GoTo 1000
1001
Do While Not EOF(1)
Input #1, Temp
DM.AddItem Temp
Loop
Close
MsgBox "User '" & LoadUser & "' Loaded."
GoTo 10002
10001
MsgBox "No Template for user '" & LoadUser & "'."
10002
Close
LoadUser.Text = ""
SaveUser.Text = ""
LoadGroup.Text = ""
SaveGroup.Text = ""
End Sub

Private Sub lpse_Click()
If lpse.Value = "0" Then
    Frame2.Enabled = False
End If
If lpse.Value = "1" Then
    Frame2.Enabled = True
End If

End Sub

Private Sub SaveGroup_click()
If SaveGroup.Text <> "" Then
SSTime1 = Abs(SSTime.Text) * 60
Open PDC.Text & "\admin$\system32\repl\import\scripts\Policies\" & SaveGroup & ".ini" For Output As #1
Print #1, lpse.Value
Print #1, cm.Text
Print #1, mps.Text
Print #1, irifl.Value
Print #1, nuwpssie.Value
Print #1, ruexm.Text
Print #1, pdirectory.Text
Print #1, dccorp.Value
Print #1, adsnc.Value
Print #1, snct.Value
Print #1, tm.Text
Print #1, sndpo.Value
Print #1, doc.Text
Print #1, cpdo.Value
Print #1, doc1.Text
Print #1, tfdb.Value
Print #1, ts.Text
Print #1, stp.Text
Print #1, sep.Text
Print #1, dicw.Value
Print #1, psa.Text
Print #1, psp.Text
Print #1, bpsfla.Value
Print #1, ep.Value
Print #1, md.Text
Print #1, oepwsn.Value
Print #1, pn.Text
Print #1, mn.Text
Print #1, Exchange2.Text
Print #1, ptewoe.Value
Print #1, PFPath.Text
Print #1, epab.Value
Print #1, PABFile.Text
Print #1, epf.Value
Print #1, PSTFile.Text
Print #1, OfflinePath.Text
Print #1, eof1.Value
Print #1, OSTFile.Text
Print #1, dret.Value
Print #1, datdi.Value
Print #1, hbt.Value
Print #1, hsst.Value
Print #1, hat.Value
Print #1, hst.Value
Print #1, rrcfsm.Value
Print #1, rffsosm.Value
Print #1, rtfsosm.Value
Print #1, rfcfsm.Value
Print #1, hdimc.Value
Print #1, hnn.Value
Print #1, neninn.Value
Print #1, haiod.Value
Print #1, rsdcfsm.Value
Print #1, rmnd.Value
Print #1, HomeDrive.Text
Print #1, HomeRoot.Text
Print #1, SSFile.Text
Print #1, SSTime1
Print #1, SSPassword.Value
Print #1, "[RUN]"
For X = 0 To run.ListCount - 1
Print #1, run.List(X)
Next X
Print #1, "[Mappings]"
For X = 0 To DM.ListCount - 1
Print #1, DM.List(X)
Next X
Close

If Not BDC1.Text = "" Then
Open BDC1.Text & "\admin$\system32\repl\import\scripts\Policies\" & SaveGroup & ".ini" For Output As #1
Print #1, lpse.Value
Print #1, cm.Text
Print #1, mps.Text
Print #1, irifl.Value
Print #1, nuwpssie.Value
Print #1, ruexm.Text
Print #1, pdirectory.Text
Print #1, dccorp.Value
Print #1, adsnc.Value
Print #1, snct.Value
Print #1, tm.Text
Print #1, sndpo.Value
Print #1, doc.Text
Print #1, cpdo.Value
Print #1, doc1.Text
Print #1, tfdb.Value
Print #1, ts.Text
Print #1, stp.Text
Print #1, sep.Text
Print #1, dicw.Value
Print #1, psa.Text
Print #1, psp.Text
Print #1, bpsfla.Value
Print #1, ep.Value
Print #1, md.Text
Print #1, oepwsn.Value
Print #1, pn.Text
Print #1, mn.Text
Print #1, Exchange2.Text
Print #1, ptewoe.Value
Print #1, PFPath.Text
Print #1, epab.Value
Print #1, PABFile.Text
Print #1, epf.Value
Print #1, PSTFile.Text
Print #1, OfflinePath.Text
Print #1, eof1.Value
Print #1, OSTFile.Text
Print #1, dret.Value
Print #1, datdi.Value
Print #1, hbt.Value
Print #1, hsst.Value
Print #1, hat.Value
Print #1, hst.Value
Print #1, rrcfsm.Value
Print #1, rffsosm.Value
Print #1, rtfsosm.Value
Print #1, rfcfsm.Value
Print #1, hdimc.Value
Print #1, hnn.Value
Print #1, neninn.Value
Print #1, haiod.Value
Print #1, rsdcfsm.Value
Print #1, rmnd.Value
Print #1, HomeDrive.Text
Print #1, HomeRoot.Text
Print #1, SSFile.Text
Print #1, SSTime1
Print #1, SSPassword.Value
Print #1, "[RUN]"
For X = 0 To run.ListCount - 1
Print #1, run.List(X)
Next X
Print #1, "[Mappings]"
For X = 0 To DM.ListCount - 1
Print #1, DM.List(X)
Next X
Close
End If

If Not BDC2.Text = "" Then
Open BDC2.Text & "\admin$\system32\repl\import\scripts\Policies\" & SaveGroup & ".ini" For Output As #1
Print #1, lpse.Value
Print #1, cm.Text
Print #1, mps.Text
Print #1, irifl.Value
Print #1, nuwpssie.Value
Print #1, ruexm.Text
Print #1, pdirectory.Text
Print #1, dccorp.Value
Print #1, adsnc.Value
Print #1, snct.Value
Print #1, tm.Text
Print #1, sndpo.Value
Print #1, doc.Text
Print #1, cpdo.Value
Print #1, doc1.Text
Print #1, tfdb.Value
Print #1, ts.Text
Print #1, stp.Text
Print #1, sep.Text
Print #1, dicw.Value
Print #1, psa.Text
Print #1, psp.Text
Print #1, bpsfla.Value
Print #1, ep.Value
Print #1, md.Text
Print #1, oepwsn.Value
Print #1, pn.Text
Print #1, mn.Text
Print #1, Exchange2.Text
Print #1, ptewoe.Value
Print #1, PFPath.Text
Print #1, epab.Value
Print #1, PABFile.Text
Print #1, epf.Value
Print #1, PSTFile.Text
Print #1, OfflinePath.Text
Print #1, eof1.Value
Print #1, OSTFile.Text
Print #1, dret.Value
Print #1, datdi.Value
Print #1, hbt.Value
Print #1, hsst.Value
Print #1, hat.Value
Print #1, hst.Value
Print #1, rrcfsm.Value
Print #1, rffsosm.Value
Print #1, rtfsosm.Value
Print #1, rfcfsm.Value
Print #1, hdimc.Value
Print #1, hnn.Value
Print #1, neninn.Value
Print #1, haiod.Value
Print #1, rsdcfsm.Value
Print #1, rmnd.Value
Print #1, HomeDrive.Text
Print #1, HomeRoot.Text
Print #1, SSFile.Text
Print #1, SSTime1
Print #1, SSPassword.Value
Print #1, "[RUN]"
For X = 0 To run.ListCount - 1
Print #1, run.List(X)
Next X
Print #1, "[Mappings]"
For X = 0 To DM.ListCount - 1
Print #1, DM.List(X)
Next X
Close
End If

If Not BDC3.Text = "" Then
Open BDC3.Text & "\admin$\system32\repl\import\scripts\Policies\" & SaveGroup & ".ini" For Output As #1
Print #1, lpse.Value
Print #1, cm.Text
Print #1, mps.Text
Print #1, irifl.Value
Print #1, nuwpssie.Value
Print #1, ruexm.Text
Print #1, pdirectory.Text
Print #1, dccorp.Value
Print #1, adsnc.Value
Print #1, snct.Value
Print #1, tm.Text
Print #1, sndpo.Value
Print #1, doc.Text
Print #1, cpdo.Value
Print #1, doc1.Text
Print #1, tfdb.Value
Print #1, ts.Text
Print #1, stp.Text
Print #1, sep.Text
Print #1, dicw.Value
Print #1, psa.Text
Print #1, psp.Text
Print #1, bpsfla.Value
Print #1, ep.Value
Print #1, md.Text
Print #1, oepwsn.Value
Print #1, pn.Text
Print #1, mn.Text
Print #1, Exchange2.Text
Print #1, ptewoe.Value
Print #1, PFPath.Text
Print #1, epab.Value
Print #1, PABFile.Text
Print #1, epf.Value
Print #1, PSTFile.Text
Print #1, OfflinePath.Text
Print #1, eof1.Value
Print #1, OSTFile.Text
Print #1, dret.Value
Print #1, datdi.Value
Print #1, hbt.Value
Print #1, hsst.Value
Print #1, hat.Value
Print #1, hst.Value
Print #1, rrcfsm.Value
Print #1, rffsosm.Value
Print #1, rtfsosm.Value
Print #1, rfcfsm.Value
Print #1, hdimc.Value
Print #1, hnn.Value
Print #1, neninn.Value
Print #1, haiod.Value
Print #1, rsdcfsm.Value
Print #1, rmnd.Value
Print #1, HomeDrive.Text
Print #1, HomeRoot.Text
Print #1, SSFile.Text
Print #1, SSTime.Text
Print #1, SSPassword.Value
Print #1, "[RUN]"
For X = 0 To run.ListCount - 1
Print #1, run.List(X)
Next X
Print #1, "[Mappings]"
For X = 0 To DM.ListCount - 1
Print #1, DM.List(X)
Next X
Close
End If
MsgBox "Settings for '" & SaveGroup.Text & "' saved."
LoadUser.Text = ""
SaveUser.Text = ""
LoadGroup.Text = ""
SaveGroup.Text = ""
End If
End Sub

Private Sub SaveMachine_Click()
If SaveMachine.Text <> "" Then
Open PDC.Text & "\admin$\system32\repl\import\scripts\Policies\" & SaveMachine & ".ini" For Output As #1
Print #1, chdsw.Value
Print #1, chdss.Value
Print #1, rlss.Value
Print #1, dl.Value
Print #1, dtm.Value
Print #1, esfadb.Value
Print #1, dlw.Value
Print #1, dcp.Value
Print #1, swtal.Value
Print #1, dndllou.Value
Print #1, lbCaption.Text
Print #1, lbText.Text
Print #1, wn.Text
Print #1, tw.Value
Close

If Not BDC1.Text = "" Then
    Open BDC1.Text & "\admin$\system32\repl\import\scripts\Policies\" & SaveMachine & ".ini" For Output As #1
Print #1, chdsw.Value
Print #1, chdss.Value
Print #1, rlss.Value
Print #1, dl.Value
Print #1, dtm.Value
Print #1, esfadb.Value
Print #1, dlw.Value
Print #1, dcp.Value
Print #1, swtal.Value
Print #1, dndllou.Value
Print #1, lbCaption.Text
Print #1, lbText.Text
Print #1, wn.Text
Print #1, tw.Value
Close
End If

If Not BDC2.Text = "" Then
    Open BDC2.Text & "\admin$\system32\repl\import\scripts\Policies\" & SaveMachine & ".ini" For Output As #1
Print #1, chdsw.Value
Print #1, chdss.Value
Print #1, rlss.Value
Print #1, dl.Value
Print #1, dtm.Value
Print #1, esfadb.Value
Print #1, dlw.Value
Print #1, dcp.Value
Print #1, swtal.Value
Print #1, dndllou.Value
Print #1, lbCaption.Text
Print #1, lbText.Text
Print #1, wn.Text
Print #1, tw.Value
Close
End If

If Not BDC3.Text = "" Then
    Open BDC3.Text & "\admin$\system32\repl\import\scripts\Policies\" & SaveMachine & ".ini" For Output As #1
Print #1, chdsw.Value
Print #1, chdss.Value
Print #1, rlss.Value
Print #1, dl.Value
Print #1, dtm.Value
Print #1, esfadb.Value
Print #1, dlw.Value
Print #1, dcp.Value
Print #1, swtal.Value
Print #1, dndllou.Value
Print #1, lbCaption.Text
Print #1, lbText.Text
Print #1, wn.Text
Print #1, tw.Value
Close
End If


MsgBox "Settings for '" & SaveMachine & "' saved."
End If

End Sub

Private Sub SaveUser_Click()
SSTime1 = Abs(SSTime.Text) * 60
If SaveUser.Text <> "" Then
Open PDC.Text & "\admin$\system32\repl\import\scripts\Policies\" & SaveUser & ".ini" For Output As #1
Print #1, lpse.Value
Print #1, cm.Text
Print #1, mps.Text
Print #1, irifl.Value
Print #1, nuwpssie.Value
Print #1, ruexm.Text
Print #1, pdirectory.Text
Print #1, dccorp.Value
Print #1, adsnc.Value
Print #1, snct.Value
Print #1, tm.Text
Print #1, sndpo.Value
Print #1, doc.Text
Print #1, cpdo.Value
Print #1, doc1.Text
Print #1, tfdb.Value
Print #1, ts.Text
Print #1, stp.Text
Print #1, sep.Text
Print #1, dicw.Value
Print #1, psa.Text
Print #1, psp.Text
Print #1, bpsfla.Value
Print #1, ep.Value
Print #1, md.Text
Print #1, oepwsn.Value
Print #1, pn.Text
Print #1, mn.Text
Print #1, Exchange2.Text
Print #1, ptewoe.Value
Print #1, PFPath.Text
Print #1, epab.Value
Print #1, PABFile.Text
Print #1, epf.Value
Print #1, PSTFile.Text
Print #1, OfflinePath.Text
Print #1, eof1.Value
Print #1, OSTFile.Text
Print #1, dret.Value
Print #1, datdi.Value
Print #1, hbt.Value
Print #1, hsst.Value
Print #1, hat.Value
Print #1, hst.Value
Print #1, rrcfsm.Value
Print #1, rffsosm.Value
Print #1, rtfsosm.Value
Print #1, rfcfsm.Value
Print #1, hdimc.Value
Print #1, hnn.Value
Print #1, neninn.Value
Print #1, haiod.Value
Print #1, rsdcfsm.Value
Print #1, rmnd.Value
Print #1, HomeDrive.Text
Print #1, HomeRoot.Text
Print #1, SSFile.Text
Print #1, SSTime1
Print #1, SSPassword.Value
Print #1, "[RUN]"
For X = 0 To run.ListCount - 1
Print #1, run.List(X)
Next X
Print #1, "[Mappings]"
For X = 0 To DM.ListCount - 1
Print #1, DM.List(X)
Next X
Close

If Not BDC1.Text = "" Then
Open BDC1.Text & "\admin$\system32\repl\import\scripts\Policies\" & SaveUser & ".ini" For Output As #1
Print #1, lpse.Value
Print #1, cm.Text
Print #1, mps.Text
Print #1, irifl.Value
Print #1, nuwpssie.Value
Print #1, ruexm.Text
Print #1, pdirectory.Text
Print #1, dccorp.Value
Print #1, adsnc.Value
Print #1, snct.Value
Print #1, tm.Text
Print #1, sndpo.Value
Print #1, doc.Text
Print #1, cpdo.Value
Print #1, doc1.Text
Print #1, tfdb.Value
Print #1, ts.Text
Print #1, stp.Text
Print #1, sep.Text
Print #1, dicw.Value
Print #1, psa.Text
Print #1, psp.Text
Print #1, bpsfla.Value
Print #1, ep.Value
Print #1, md.Text
Print #1, oepwsn.Value
Print #1, pn.Text
Print #1, mn.Text
Print #1, Exchange2.Text
Print #1, ptewoe.Value
Print #1, PFPath.Text
Print #1, epab.Value
Print #1, PABFile.Text
Print #1, epf.Value
Print #1, PSTFile.Text
Print #1, OfflinePath.Text
Print #1, eof1.Value
Print #1, OSTFile.Text
Print #1, dret.Value
Print #1, datdi.Value
Print #1, hbt.Value
Print #1, hsst.Value
Print #1, hat.Value
Print #1, hst.Value
Print #1, rrcfsm.Value
Print #1, rffsosm.Value
Print #1, rtfsosm.Value
Print #1, rfcfsm.Value
Print #1, hdimc.Value
Print #1, hnn.Value
Print #1, neninn.Value
Print #1, haiod.Value
Print #1, rsdcfsm.Value
Print #1, rmnd.Value
Print #1, HomeDrive.Text
Print #1, HomeRoot.Text
Print #1, SSFile.Text
Print #1, SSTime1
Print #1, SSPassword.Value
Print #1, "[RUN]"
For X = 0 To run.ListCount - 1
Print #1, run.List(X)
Next X
Print #1, "[Mappings]"
For X = 0 To DM.ListCount - 1
Print #1, DM.List(X)
Next X
Close
End If

If Not BDC2.Text = "" Then
Open BDC2.Text & "\admin$\system32\repl\import\scripts\Policies\" & SaveUser & ".ini" For Output As #1
Print #1, lpse.Value
Print #1, cm.Text
Print #1, mps.Text
Print #1, irifl.Value
Print #1, nuwpssie.Value
Print #1, ruexm.Text
Print #1, pdirectory.Text
Print #1, dccorp.Value
Print #1, adsnc.Value
Print #1, snct.Value
Print #1, tm.Text
Print #1, sndpo.Value
Print #1, doc.Text
Print #1, cpdo.Value
Print #1, doc1.Text
Print #1, tfdb.Value
Print #1, ts.Text
Print #1, stp.Text
Print #1, sep.Text
Print #1, dicw.Value
Print #1, psa.Text
Print #1, psp.Text
Print #1, bpsfla.Value
Print #1, ep.Value
Print #1, md.Text
Print #1, oepwsn.Value
Print #1, pn.Text
Print #1, mn.Text
Print #1, Exchange2.Text
Print #1, ptewoe.Value
Print #1, PFPath.Text
Print #1, epab.Value
Print #1, PABFile.Text
Print #1, epf.Value
Print #1, PSTFile.Text
Print #1, OfflinePath.Text
Print #1, eof1.Value
Print #1, OSTFile.Text
Print #1, dret.Value
Print #1, datdi.Value
Print #1, hbt.Value
Print #1, hsst.Value
Print #1, hat.Value
Print #1, hst.Value
Print #1, rrcfsm.Value
Print #1, rffsosm.Value
Print #1, rtfsosm.Value
Print #1, rfcfsm.Value
Print #1, hdimc.Value
Print #1, hnn.Value
Print #1, neninn.Value
Print #1, haiod.Value
Print #1, rsdcfsm.Value
Print #1, rmnd.Value
Print #1, HomeDrive.Text
Print #1, HomeRoot.Text
Print #1, SSFile.Text
Print #1, SSTime1
Print #1, SSPassword.Value
Print #1, "[RUN]"
For X = 0 To run.ListCount - 1
Print #1, run.List(X)
Next X
Print #1, "[Mappings]"
For X = 0 To DM.ListCount - 1
Print #1, DM.List(X)
Next X
Close
End If

If Not BDC3.Text = "" Then
Open BDC3.Text & "\admin$\system32\repl\import\scripts\Policies\" & SaveUser & ".ini" For Output As #1
Print #1, lpse.Value
Print #1, cm.Text
Print #1, mps.Text
Print #1, irifl.Value
Print #1, nuwpssie.Value
Print #1, ruexm.Text
Print #1, pdirectory.Text
Print #1, dccorp.Value
Print #1, adsnc.Value
Print #1, snct.Value
Print #1, tm.Text
Print #1, sndpo.Value
Print #1, doc.Text
Print #1, cpdo.Value
Print #1, doc1.Text
Print #1, tfdb.Value
Print #1, ts.Text
Print #1, stp.Text
Print #1, sep.Text
Print #1, dicw.Value
Print #1, psa.Text
Print #1, psp.Text
Print #1, bpsfla.Value
Print #1, ep.Value
Print #1, md.Text
Print #1, oepwsn.Value
Print #1, pn.Text
Print #1, mn.Text
Print #1, Exchange2.Text
Print #1, ptewoe.Value
Print #1, PFPath.Text
Print #1, epab.Value
Print #1, PABFile.Text
Print #1, epf.Value
Print #1, PSTFile.Text
Print #1, OfflinePath.Text
Print #1, eof1.Value
Print #1, OSTFile.Text
Print #1, dret.Value
Print #1, datdi.Value
Print #1, hbt.Value
Print #1, hsst.Value
Print #1, hat.Value
Print #1, hst.Value
Print #1, rrcfsm.Value
Print #1, rffsosm.Value
Print #1, rtfsosm.Value
Print #1, rfcfsm.Value
Print #1, hdimc.Value
Print #1, hnn.Value
Print #1, neninn.Value
Print #1, haiod.Value
Print #1, rsdcfsm.Value
Print #1, rmnd.Value
Print #1, HomeDrive.Text
Print #1, HomeRoot.Text
Print #1, SSFile.Text
Print #1, SSTime1
Print #1, SSPassword.Value
Print #1, "[RUN]"
For X = 0 To run.ListCount - 1
Print #1, run.List(X)
Next X
Print #1, "[Mappings]"
For X = 0 To DM.ListCount - 1
Print #1, DM.List(X)
Next X
Close
End If
MsgBox "Settings for '" & SaveUser.Text & "' saved."
LoadUser.Text = ""
SaveUser.Text = ""
LoadGroup.Text = ""
SaveGroup.Text = ""
End If
End Sub

Private Sub ScrollKB_Change()
mps.Text = ScrollKB.Value
End Sub

Private Sub ScrollMin_Change()
ruexm.Text = ScrollMin.Value

End Sub

Private Sub Scrollms_Change()
tm.Text = ScrollMin.Value

End Sub

Private Sub Scrollss_Change()
SSTime.Text = Scrollss.Value
End Sub

Private Sub snct_Click()
If snct.Value = 1 Then
    Label16.Enabled = True
    tm.Enabled = True
    Scrollms.Enabled = True
    
End If
If snct.Value = 0 Then
    Label16.Enabled = False
    tm.Enabled = False
    Scrollms.Enabled = False
    
End If
End Sub

Private Sub sndpo_Click()
If sndpo.Value = 1 Then
    Label17.Enabled = True
    doc.Enabled = True
End If
If sndpo.Value = 0 Then
    Label17.Enabled = False
    doc.Enabled = False
End If
End Sub

Private Sub tfdb_Click()
If tfdb.Value = 1 Then
    Label19.Enabled = True
    ts.Enabled = True
    Scrollts.Enabled = True
    
End If
If tfdb.Value = 0 Then
    Label19.Enabled = False
    ts.Enabled = False
    Scrollts.Enabled = False
    
End If
End Sub

Private Sub VScroll1_Change()

End Sub
