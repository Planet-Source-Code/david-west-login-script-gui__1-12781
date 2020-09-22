VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Main 
   Caption         =   "Login Script Admin"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   13
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   397
      WordWrap        =   0   'False
      OLEDropMode     =   1
      TabCaption(0)   =   "Global Settings"
      TabPicture(0)   =   "Main.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(4)=   "Label30"
      Tab(0).Control(5)=   "Command6"
      Tab(0).Control(6)=   "BDC3"
      Tab(0).Control(7)=   "Command5"
      Tab(0).Control(8)=   "BDC2"
      Tab(0).Control(9)=   "Command4"
      Tab(0).Control(10)=   "BDC1"
      Tab(0).Control(11)=   "Command3"
      Tab(0).Control(12)=   "PDC"
      Tab(0).Control(13)=   "Exchange"
      Tab(0).Control(14)=   "Command2"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Internet Explorer"
      TabPicture(1)   =   "Main.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Check1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Microsoft Office"
      TabPicture(2)   =   "Main.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Microsoft Outlook"
      TabPicture(3)   =   "Main.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(1)=   "Frame6"
      Tab(3).Control(2)=   "ptewoe"
      Tab(3).Control(3)=   "Command9"
      Tab(3).Control(4)=   "Exchange2"
      Tab(3).Control(5)=   "mn"
      Tab(3).Control(6)=   "pn"
      Tab(3).Control(7)=   "oepwsn"
      Tab(3).Control(8)=   "Label31"
      Tab(3).Control(9)=   "Label29"
      Tab(3).Control(10)=   "Label28"
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "Drive Mappings"
      TabPicture(4)   =   "Main.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label22"
      Tab(4).Control(1)=   "Label23"
      Tab(4).Control(2)=   "Command8"
      Tab(4).Control(3)=   "Command7"
      Tab(4).Control(4)=   "Drive"
      Tab(4).Control(5)=   "Path"
      Tab(4).Control(6)=   "DM"
      Tab(4).Control(7)=   "DelMap"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Windows NT Network"
      TabPicture(5)   =   "Main.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "chdsw"
      Tab(5).Control(1)=   "chdss"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Control Panel"
      TabPicture(6)   =   "Main.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Display"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Desktop"
      TabPicture(7)   =   "Main.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label10"
      Tab(7).Control(1)=   "Label9"
      Tab(7).Control(2)=   "tw"
      Tab(7).Control(3)=   "Command1"
      Tab(7).Control(4)=   "wn"
      Tab(7).ControlCount=   5
      TabCaption(8)   =   "Shell"
      TabPicture(8)   =   "Main.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "rrcfsm"
      Tab(8).Control(1)=   "rsdcfsm"
      Tab(8).Control(2)=   "haiod"
      Tab(8).Control(3)=   "neninn"
      Tab(8).Control(4)=   "hnn"
      Tab(8).Control(5)=   "hdimc"
      Tab(8).Control(6)=   "rfcfsm"
      Tab(8).Control(7)=   "rtfsosm"
      Tab(8).Control(8)=   "rffsosm"
      Tab(8).Control(9)=   "rmnd"
      Tab(8).ControlCount=   10
      TabCaption(9)   =   "System"
      TabPicture(9)   =   "Main.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "dret"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Windows NT System"
      TabPicture(10)  =   "Main.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Check5"
      Tab(10).Control(1)=   "Check6"
      Tab(10).Control(2)=   "Check7"
      Tab(10).Control(3)=   "Check8"
      Tab(10).Control(4)=   "Check9"
      Tab(10).Control(5)=   "Check4"
      Tab(10).Control(6)=   "Check10"
      Tab(10).Control(7)=   "Check11"
      Tab(10).Control(8)=   "Frame3"
      Tab(10).ControlCount=   9
      TabCaption(11)  =   "Windows NT Profiles"
      TabPicture(11)  =   "Main.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Scrollts"
      Tab(11).Control(1)=   "ts"
      Tab(11).Control(2)=   "doc1"
      Tab(11).Control(3)=   "doc"
      Tab(11).Control(4)=   "Scrollms"
      Tab(11).Control(5)=   "tm"
      Tab(11).Control(6)=   "tfdb"
      Tab(11).Control(7)=   "cpdo"
      Tab(11).Control(8)=   "sndpo"
      Tab(11).Control(9)=   "snct"
      Tab(11).Control(10)=   "adsnc"
      Tab(11).Control(11)=   "dccorp"
      Tab(11).Control(12)=   "Frame2"
      Tab(11).Control(13)=   "Label19"
      Tab(11).Control(14)=   "Label18"
      Tab(11).Control(15)=   "Label17"
      Tab(11).Control(16)=   "Label16"
      Tab(11).ControlCount=   17
      TabCaption(12)  =   "General"
      TabPicture(12)  =   "Main.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Frame4"
      Tab(12).Control(0).Enabled=   0   'False
      Tab(12).Control(1)=   "Frame5"
      Tab(12).Control(1).Enabled=   0   'False
      Tab(12).ControlCount=   2
      Begin VB.Frame Frame7 
         Caption         =   "&Offline Access Settings"
         Height          =   1815
         Left            =   -70080
         TabIndex        =   136
         Top             =   3360
         Width           =   4215
         Begin VB.TextBox OSTFile 
            Height          =   285
            Left            =   1320
            TabIndex        =   142
            Top             =   1200
            Width           =   2775
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Enable Offlince Folders"
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   840
            Width           =   2895
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Browse"
            Height          =   255
            Left            =   3240
            TabIndex        =   139
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox OfflinePath 
            Height          =   285
            Left            =   600
            TabIndex        =   138
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label36 
            Caption         =   "OST File Name"
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label35 
            Caption         =   "Folder"
            Height          =   255
            Left            =   120
            TabIndex        =   137
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "PAB & &Personal Folder Settings"
         Height          =   2055
         Left            =   -70080
         TabIndex        =   126
         Top             =   1080
         Width           =   4215
         Begin VB.TextBox PSTFile 
            Height          =   285
            Left            =   1320
            TabIndex        =   135
            Top             =   1680
            Width           =   2775
         End
         Begin VB.CheckBox epf 
            Caption         =   "Enable Personal Folders"
            Height          =   255
            Left            =   120
            TabIndex        =   133
            Top             =   1320
            Width           =   3615
         End
         Begin VB.TextBox PABFile 
            Height          =   285
            Left            =   1320
            TabIndex        =   132
            Top             =   960
            Width           =   2775
         End
         Begin VB.CheckBox epab 
            Caption         =   "Enable Personal Address Book"
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   600
            Width           =   3495
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Browse"
            Height          =   255
            Left            =   3240
            TabIndex        =   129
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox PFPath 
            Height          =   285
            Left            =   600
            TabIndex        =   128
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label34 
            Caption         =   "PST File Name"
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label33 
            Caption         =   "PAB File Name"
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label32 
            Caption         =   "Folder"
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CheckBox ptewoe 
         Caption         =   "Prompt to empty &Wastebasket on Exit"
         Height          =   255
         Left            =   -74760
         TabIndex        =   125
         Top             =   2760
         Width           =   3855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -71400
         TabIndex        =   124
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Exchange2 
         Height          =   285
         Left            =   -73440
         TabIndex        =   123
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -68520
         TabIndex        =   120
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox Exchange 
         Height          =   285
         Left            =   -73320
         TabIndex        =   118
         Top             =   4200
         Width           =   4695
      End
      Begin VB.TextBox mn 
         Height          =   285
         Left            =   -73680
         TabIndex        =   119
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox pn 
         Height          =   285
         Left            =   -73680
         TabIndex        =   116
         Top             =   1440
         Width           =   3375
      End
      Begin VB.CheckBox oepwsn 
         Caption         =   "&Overwrite existing profile with same name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   114
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Frame Frame5 
         Caption         =   "Save Settings To"
         Height          =   2055
         Left            =   -74640
         TabIndex        =   104
         Top             =   3120
         Width           =   8775
         Begin VB.TextBox DomainName 
            Height          =   285
            Left            =   240
            TabIndex        =   113
            Top             =   1680
            Width           =   3255
         End
         Begin VB.TextBox Compname 
            Height          =   285
            Left            =   600
            TabIndex        =   112
            Top             =   240
            Width           =   3135
         End
         Begin VB.ComboBox SaveGroup 
            Height          =   315
            Left            =   4800
            Sorted          =   -1  'True
            TabIndex        =   109
            Top             =   840
            Width           =   3135
         End
         Begin VB.ComboBox SaveUser 
            Height          =   315
            Left            =   600
            Sorted          =   -1  'True
            TabIndex        =   106
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label Label27 
            Caption         =   "Username"
            Height          =   255
            Left            =   1560
            TabIndex        =   108
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label26 
            Caption         =   "Group"
            Height          =   255
            Left            =   6240
            TabIndex        =   107
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Load Settings From"
         Height          =   1815
         Left            =   -74640
         TabIndex        =   100
         Top             =   1200
         Width           =   8775
         Begin VB.TextBox DomainController 
            Height          =   285
            Left            =   5160
            TabIndex        =   111
            Top             =   1440
            Width           =   3135
         End
         Begin VB.TextBox username 
            Height          =   285
            Left            =   360
            TabIndex        =   110
            Top             =   0
            Width           =   2775
         End
         Begin VB.ComboBox LoadGroup 
            Height          =   315
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   105
            Top             =   1320
            Width           =   3135
         End
         Begin VB.ComboBox LoadUser 
            Height          =   315
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   101
            Top             =   600
            Width           =   3135
         End
         Begin VB.Label Label25 
            Caption         =   "Group"
            Height          =   255
            Left            =   1440
            TabIndex        =   103
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label24 
            Caption         =   "Username"
            Height          =   255
            Left            =   1320
            TabIndex        =   102
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.CheckBox chdss 
         Caption         =   "Create hidden drive shares (server)"
         Height          =   255
         Left            =   -72000
         TabIndex        =   99
         Top             =   3120
         Width           =   4095
      End
      Begin VB.CheckBox chdsw 
         Caption         =   "Create hidden drive shares (workstation)"
         Height          =   255
         Left            =   -72000
         TabIndex        =   98
         Top             =   2640
         Width           =   4095
      End
      Begin VB.CommandButton DelMap 
         Caption         =   "Delete selected Map"
         Height          =   255
         Left            =   -71400
         TabIndex        =   95
         Top             =   3960
         Width           =   2415
      End
      Begin VB.ListBox DM 
         Columns         =   2
         Height          =   840
         ItemData        =   "Main.frx":016C
         Left            =   -72600
         List            =   "Main.frx":016E
         Sorted          =   -1  'True
         TabIndex        =   94
         Top             =   2880
         Width           =   4695
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Disable Logoff"
         Height          =   255
         Left            =   -74640
         TabIndex        =   93
         Top             =   1560
         Width           =   3855
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Show welcome tips at logon"
         Height          =   255
         Left            =   -70440
         TabIndex        =   92
         Top             =   1920
         Width           =   3855
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Disable Change Password"
         Height          =   255
         Left            =   -70440
         TabIndex        =   91
         Top             =   1560
         Width           =   3855
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Disable Lock Workstation"
         Height          =   255
         Left            =   -70440
         TabIndex        =   90
         Top             =   1200
         Width           =   3855
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Disable Task Manager"
         Height          =   255
         Left            =   -74640
         TabIndex        =   89
         Top             =   1920
         Width           =   3855
      End
      Begin VB.TextBox PDC 
         Height          =   285
         Left            =   -73320
         TabIndex        =   76
         Top             =   1800
         Width           =   4695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -68520
         TabIndex        =   75
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox BDC1 
         Height          =   285
         Left            =   -73320
         TabIndex        =   74
         Top             =   2400
         Width           =   4695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -68520
         TabIndex        =   73
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox BDC2 
         Height          =   285
         Left            =   -73320
         TabIndex        =   72
         Top             =   3000
         Width           =   4695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -68520
         TabIndex        =   71
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox BDC3 
         Height          =   285
         Left            =   -73320
         TabIndex        =   70
         Top             =   3600
         Width           =   4695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -68520
         TabIndex        =   69
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   68
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   67
         Top             =   2400
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Disable Internet Connection Wizard"
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   2760
         Width           =   5535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Proxy Settings"
         Height          =   1815
         Left            =   120
         TabIndex        =   59
         Top             =   3360
         Width           =   4215
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   63
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   3240
            TabIndex        =   62
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Bypass Proxy server for local addresses"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   1080
            Width           =   3255
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Enable Proxy"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label7 
            Caption         =   "Proxy Server Address i.e. (http://10.0.0.1)"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label8 
            Caption         =   "Port i.e. (80)"
            Height          =   255
            Left            =   3240
            TabIndex        =   64
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox Path 
         Height          =   285
         Left            =   -71760
         TabIndex        =   58
         Top             =   1680
         Width           =   3735
      End
      Begin VB.ComboBox Drive 
         Height          =   315
         ItemData        =   "Main.frx":0170
         Left            =   -72600
         List            =   "Main.frx":01B6
         TabIndex        =   57
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -67920
         TabIndex        =   56
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add"
         Height          =   255
         Left            =   -73920
         TabIndex        =   55
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Frame Display 
         Caption         =   "Display"
         Height          =   4095
         Left            =   -74760
         TabIndex        =   49
         Top             =   1080
         Width           =   9015
         Begin VB.CheckBox hst 
            Caption         =   "Hide Settings tab"
            Height          =   255
            Left            =   3360
            TabIndex        =   54
            Top             =   2640
            Width           =   2415
         End
         Begin VB.CheckBox hat 
            Caption         =   "Hide Appearance tab"
            Height          =   255
            Left            =   3360
            TabIndex        =   53
            Top             =   2280
            Width           =   2415
         End
         Begin VB.CheckBox hsst 
            Caption         =   "Hide Screen Saver tab"
            Height          =   255
            Left            =   3360
            TabIndex        =   52
            Top             =   1920
            Width           =   2295
         End
         Begin VB.CheckBox hbt 
            Caption         =   "Hide Background tab"
            Height          =   255
            Left            =   3360
            TabIndex        =   51
            Top             =   1560
            Width           =   2295
         End
         Begin VB.CheckBox datdi 
            Caption         =   "Deny access to display icon"
            Height          =   255
            Left            =   3360
            TabIndex        =   50
            Top             =   1200
            Width           =   3855
         End
      End
      Begin VB.TextBox wn 
         Height          =   285
         Left            =   -72600
         TabIndex        =   48
         Top             =   1800
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   255
         Left            =   -68280
         TabIndex        =   47
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox tw 
         Caption         =   "Tile Wallpaper"
         Height          =   255
         Left            =   -72600
         TabIndex        =   46
         Top             =   2760
         Width           =   3615
      End
      Begin VB.Frame Frame2 
         Caption         =   "Limit Profile Size"
         Height          =   4215
         Left            =   -74880
         TabIndex        =   32
         Top             =   1080
         Width           =   4335
         Begin VB.TextBox cm 
            Height          =   285
            Left            =   1440
            TabIndex        =   40
            Text            =   $"Main.frx":0212
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox mps 
            Height          =   285
            Left            =   1680
            TabIndex        =   39
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
            TabIndex        =   38
            Top             =   840
            Width           =   255
         End
         Begin VB.CheckBox irifl 
            Caption         =   "Include registry in file list"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   3735
         End
         Begin VB.CheckBox nuwpssie 
            Caption         =   "Notify user when profile storage space is exceeded."
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   1560
            Width           =   3975
         End
         Begin VB.TextBox ruexm 
            Height          =   285
            Left            =   2280
            TabIndex        =   35
            Text            =   "15"
            Top             =   1920
            Width           =   735
         End
         Begin VB.VScrollBar ScrollMin 
            Height          =   255
            Left            =   3120
            Max             =   0
            Min             =   60
            TabIndex        =   34
            Top             =   1920
            Value           =   15
            Width           =   255
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Text            =   "Temporary Internet Files;Temp"
            Top             =   3000
            Width           =   4095
         End
         Begin VB.Label Label11 
            Caption         =   "Custom Message"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Max Profile size (KB)"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label13 
            Caption         =   "Remind user every X minutes:"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label14 
            Caption         =   "Prevent the following directories from roaming with the profile:"
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   2520
            Width           =   4095
         End
         Begin VB.Label Label15 
            Caption         =   "You can enter multiple name, semi-colon seperated, all relative to the root of the user's profile"
            Height          =   495
            Left            =   120
            TabIndex        =   41
            Top             =   3480
            Width           =   4095
         End
      End
      Begin VB.CheckBox dccorp 
         Caption         =   "Delete cached copies of roaming profiles"
         Height          =   255
         Left            =   -70320
         TabIndex        =   31
         Top             =   1200
         Width           =   4215
      End
      Begin VB.CheckBox adsnc 
         Caption         =   "Automatically detect slow network connections"
         Height          =   255
         Left            =   -70320
         TabIndex        =   30
         Top             =   1560
         Width           =   4335
      End
      Begin VB.CheckBox snct 
         Caption         =   "Slow network connection timeout"
         Height          =   255
         Left            =   -70320
         TabIndex        =   29
         Top             =   1920
         Width           =   4575
      End
      Begin VB.CheckBox sndpo 
         Caption         =   "Slow network default profile operation"
         Height          =   255
         Left            =   -70320
         TabIndex        =   28
         Top             =   2760
         Width           =   4575
      End
      Begin VB.CheckBox cpdo 
         Caption         =   "Choose profile default operation"
         Height          =   255
         Left            =   -70320
         TabIndex        =   27
         Top             =   3600
         Width           =   4455
      End
      Begin VB.CheckBox tfdb 
         Caption         =   "Timeout for dialog boxes"
         Height          =   255
         Left            =   -70320
         TabIndex        =   26
         Top             =   4440
         Width           =   4455
      End
      Begin VB.TextBox tm 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68160
         TabIndex        =   25
         Text            =   "2000"
         Top             =   2280
         Width           =   855
      End
      Begin VB.VScrollBar Scrollms 
         Enabled         =   0   'False
         Height          =   255
         Left            =   -67200
         Max             =   0
         Min             =   10000
         TabIndex        =   24
         Top             =   2280
         Value           =   2000
         Width           =   255
      End
      Begin VB.ComboBox doc 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Main.frx":02A7
         Left            =   -68400
         List            =   "Main.frx":02B1
         TabIndex        =   23
         Top             =   3120
         Width           =   2175
      End
      Begin VB.ComboBox doc1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Main.frx":02DA
         Left            =   -68400
         List            =   "Main.frx":02E4
         TabIndex        =   22
         Top             =   3960
         Width           =   2175
      End
      Begin VB.TextBox ts 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68160
         TabIndex        =   21
         Text            =   "30"
         Top             =   4800
         Width           =   855
      End
      Begin VB.VScrollBar Scrollts 
         Enabled         =   0   'False
         Height          =   255
         Left            =   -67200
         Max             =   0
         Min             =   120
         TabIndex        =   20
         Top             =   4800
         Value           =   30
         Width           =   255
      End
      Begin VB.CheckBox rrcfsm 
         Caption         =   "Remove Run command from Start menu"
         Height          =   255
         Left            =   -72480
         TabIndex        =   19
         Top             =   1440
         Width           =   3735
      End
      Begin VB.CheckBox rsdcfsm 
         Caption         =   "Remove Shut Down command from Start menu"
         Height          =   255
         Left            =   -72480
         TabIndex        =   18
         Top             =   4320
         Width           =   3735
      End
      Begin VB.CheckBox haiod 
         Caption         =   "Hide all items on desktop"
         Height          =   255
         Left            =   -72480
         TabIndex        =   17
         Top             =   3960
         Width           =   2175
      End
      Begin VB.CheckBox neninn 
         Caption         =   "No Entire Network in Network Neighborhood"
         Height          =   255
         Left            =   -72480
         TabIndex        =   16
         Top             =   3600
         Width           =   3615
      End
      Begin VB.CheckBox hnn 
         Caption         =   "Hide Network Neighborhood"
         Height          =   255
         Left            =   -72480
         TabIndex        =   15
         Top             =   3240
         Width           =   3615
      End
      Begin VB.CheckBox hdimc 
         Caption         =   "Hide drives in My Computer"
         Height          =   255
         Left            =   -72480
         TabIndex        =   14
         Top             =   2880
         Width           =   3495
      End
      Begin VB.CheckBox rfcfsm 
         Caption         =   "Remove Find command from Start menu"
         Height          =   255
         Left            =   -72480
         TabIndex        =   13
         Top             =   2520
         Width           =   3495
      End
      Begin VB.CheckBox rtfsosm 
         Caption         =   "Remove Taskbar from Settings on Start menu"
         Height          =   255
         Left            =   -72480
         TabIndex        =   12
         Top             =   2160
         Width           =   3735
      End
      Begin VB.CheckBox rffsosm 
         Caption         =   "Remove folders from Settings on Start menu"
         Height          =   255
         Left            =   -72480
         TabIndex        =   11
         Top             =   1800
         Width           =   3735
      End
      Begin VB.CheckBox dret 
         Caption         =   "Disable Registry editing tools"
         Height          =   255
         Left            =   -71760
         TabIndex        =   10
         Top             =   2640
         Width           =   3975
      End
      Begin VB.CheckBox rmnd 
         Caption         =   "Remove the ""Map Network Drive"" and ""Disconnect Network Drive"" options"
         Height          =   255
         Left            =   -72480
         TabIndex        =   9
         Top             =   4680
         Width           =   5775
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Run logon scripts synchronously."
         Height          =   255
         Left            =   -74640
         TabIndex        =   8
         Top             =   1200
         Width           =   3855
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Enable shutdown from Authentication dialog box"
         Height          =   255
         Left            =   -74640
         TabIndex        =   7
         Top             =   2280
         Width           =   3735
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Do not display last logged on user name"
         Height          =   255
         Left            =   -70440
         TabIndex        =   6
         Top             =   2280
         Width           =   3615
      End
      Begin VB.Frame Frame3 
         Caption         =   "Logon Banner"
         Height          =   2415
         Left            =   -74640
         TabIndex        =   1
         Top             =   2760
         Width           =   8775
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   8535
         End
         Begin VB.TextBox Text7 
            Height          =   1095
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   1200
            Width           =   8535
         End
         Begin VB.Label Label20 
            Caption         =   "Caption"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   8535
         End
         Begin VB.Label Label21 
            Caption         =   "Text"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   960
            Width           =   8535
         End
      End
      Begin VB.Label Label31 
         Caption         =   "&Exchange Server"
         Height          =   255
         Left            =   -74760
         TabIndex        =   122
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label30 
         Caption         =   "Exchange Server"
         Height          =   255
         Left            =   -73320
         TabIndex        =   121
         Top             =   3960
         Width           =   3255
      End
      Begin VB.Label Label29 
         Caption         =   "&Mailbox Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   117
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "P&rofile Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   115
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "Path"
         Height          =   255
         Left            =   -71880
         TabIndex        =   97
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Drive"
         Height          =   255
         Left            =   -72600
         TabIndex        =   96
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Primary Domain Controller"
         Height          =   255
         Left            =   -73320
         TabIndex        =   88
         Top             =   1560
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "Backup Domain Controller"
         Height          =   255
         Left            =   -73320
         TabIndex        =   87
         Top             =   2160
         Width           =   4935
      End
      Begin VB.Label Label3 
         Caption         =   "Backup Domain Controller"
         Height          =   255
         Left            =   -73320
         TabIndex        =   86
         Top             =   2760
         Width           =   6015
      End
      Begin VB.Label Label4 
         Caption         =   "Backup Domain Controller"
         Height          =   255
         Left            =   -73320
         TabIndex        =   85
         Top             =   3360
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "Start Page"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "Search Page"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Label Label9 
         Caption         =   "Wallpaper Name"
         Height          =   255
         Left            =   -72600
         TabIndex        =   82
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Specify location and name (e.g. c:\winnt\winnt256.bmp)"
         Height          =   255
         Left            =   -72600
         TabIndex        =   81
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Label Label16 
         Caption         =   "Time (milliseconds)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -69600
         TabIndex        =   80
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Default option"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -69600
         TabIndex        =   79
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Default option"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -69600
         TabIndex        =   78
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Time (seconds)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -69600
         TabIndex        =   77
         Top             =   4800
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Main"
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

Private Sub Command2_Click()
Exchange.Text = BrowseForFolder("")
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


SSTab1.Tab = "12"
Set AU = CreateObject("Persits.ASPUser")
username.Text = AU.GetUserName
DomainName.Text = GetDomainName
Compname.Text = AU.GetComputerName
DomainController.Text = AU.DomainController
If AU.Users(username.Text).Groups("Domain Admins") = "" Then
    MsgBox "You do not have suffecient permissions to run this program. You must be in the 'Domain Admins' group."
    End
    End If
    
    
For Each SRV In AU.Servers
    LoadMachine.AddItem (SRV)
    Next

For Each User In AU.Users
    LoadUser.AddItem (User.Name)
    SaveUser.AddItem (User.Name)
Next
AU.Server = DomainController.Text
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

Private Sub ScrollMin_Change()
ruexm.Text = ScrollMin.Value

End Sub

Private Sub Scrollms_Change()
tm.Text = Scrollms.Value
End Sub

Private Sub Scrollts_Change()
ts.Text = Scrollts.Value
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
