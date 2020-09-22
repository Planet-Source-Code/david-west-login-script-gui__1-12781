VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Login Script"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Application Running"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   6255
      Begin VB.Label App 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   6015
      End
   End
   Begin MSComctlLib.ProgressBar Status 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6495
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2055
      Left            =   0
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label per 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Main.Show
App.Caption = "Loading Machine Settings"
per.Caption = "14" & "%"
Status.Value = "14.2"
Main.Refresh

On Error Resume Next
'Dim r As Long
'Dim s As Long
'Dim t As String
't = String(2049, Chr(0))
's = 2084
'r = GetEnvironmentVariable("LogonServer", t, s)
'   logonserver = Left(t, r)
'r = GetEnvironmentVariable("ComputerName", t, s)
'   Computername = Left(t, r)
'r = GetEnvironmentVariable("UserName", t, s)
'   username = Left(t, r)
 Set AU = CreateObject("Persits.aspuser")
username = AU.GetUserName
DomainName = AU.GetDomainName
Computername = AU.GetComputerName
logonserver = AU.DomainController

Path = logonserver & "\Netlogon\Policies\"
Path2 = logonserver & "\Netlogon\"
FiletoCopy = Path & "ASPUser.dll"
FileCopy FiletoCopy, "C:\Winnt\System32\ASPUser.dll"
Program = Shell("Regsvr32 /s C:\Winnt\System32\ASPUser.dll", 0)
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Persits Software\ASPUser\RegKey"
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Persits Software\ASPUser\RegKey", "", "88517-52554-08877", REG_SZ
Set User = AU.Users(AU.GetUserName)
Dim Groups(500)
u = 1
For Each Group In User.Groups
Groups(u) = Group
u = u + 1
Next
Open Path & "Global.ini" For Input As #1
Do While Not EOF(1)
Line Input #1, ExchangeServer
Loop
Close
Temp = boolExists(Path & AU.GetComputerName & ".ini")
If Temp = "False" Then GoTo 1
Open Path & AU.GetComputerName & ".ini" For Input As #1
Line Input #1, chdsw
Line Input #1, chdss
Line Input #1, rlss
Line Input #1, DL
Line Input #1, DTM
Line Input #1, esfadb
Line Input #1, DLW
Line Input #1, DCP
Line Input #1, swtal
Line Input #1, dndllou
Line Input #1, Caption1
Line Input #1, Text1
Line Input #1, wn
Line Input #1, tw
Close
GoTo 2
1
Open Path & "GBLMachine.ini" For Input As #1
Line Input #1, chdsw
Line Input #1, chdss
Line Input #1, rlss
Line Input #1, DL
Line Input #1, DTM
Line Input #1, esfadb
Line Input #1, DLW
Line Input #1, DCP
Line Input #1, swtal
Line Input #1, dndllou
Line Input #1, Caption1
Line Input #1, Text1
Line Input #1, wn
Line Input #1, tw
Close
2
App.Caption = "Changing Machine Settings"
per.Caption = "28" & "%"
Status.Value = "28.4"
Main.Refresh
SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\LanManServer\Parameters", "AutoShareWks", chdsw, REG_DWORD
SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\LanManServer\Parameters", "AutoShareServer", chdss, REG_DWORD
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "RunLogonScriptSync", rlss, REG_SZ
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableLockWorkstation", DLW, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", DTM, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableChangePassword", DCP, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableLogoff", DL, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Tips", "Show", swtal, REG_DWORD
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "ShutDownWithoutlogon", esfadb, REG_SZ
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "DontDisplayLastUserName", dndllou, REG_SZ
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "LegalNoticeCaption", Caption1, REG_SZ
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "LegalNoticeText", Text1, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoChangingWallpaper", "0", REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "WallPaper", wn, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", tw, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoChangingWallpaper", "1", REG_DWORD
App.Caption = "Loading User Settings"
per.Caption = "42" & "%"
Status.Value = "42.6"
Main.Refresh
Temp = boolExists(Path & AU.GetUserName & ".ini")
If Temp = "False" Then GoTo 1111
Open Path & AU.GetUserName & ".ini" For Input As #1
Line Input #1, lpse
Line Input #1, cm
Line Input #1, mps
Line Input #1, irifl
Line Input #1, nuwpssie
Line Input #1, ruexm
Line Input #1, pdirectory
Line Input #1, dccorp
Line Input #1, adsnc
Line Input #1, snct
Line Input #1, tm
Line Input #1, sndpo
Line Input #1, doc
Line Input #1, cpdo
Line Input #1, doc1
Line Input #1, tfdb
Line Input #1, ts
Line Input #1, stp
Line Input #1, sep
Line Input #1, dicw
Line Input #1, psa
Line Input #1, psp
Line Input #1, bpsfla
Line Input #1, ep
Line Input #1, md
Line Input #1, oepwsn
Line Input #1, pn
Line Input #1, MN
Line Input #1, Exchange2
Line Input #1, ptewoe
Line Input #1, PfPath
Line Input #1, epab
Line Input #1, PABFile
Line Input #1, epf
Line Input #1, PSTFile
Line Input #1, OfflinePath
Line Input #1, eof1
Line Input #1, OSTFile
Line Input #1, dret
Line Input #1, datdi
Line Input #1, hbt
Line Input #1, hsst
Line Input #1, hat
Line Input #1, hst
Line Input #1, rrcfsm
Line Input #1, rffsosm
Line Input #1, rtfsosm
Line Input #1, rfcfsm
Line Input #1, hdimc
Line Input #1, hnn
Line Input #1, neninn
Line Input #1, haiod
Line Input #1, rsdcfsm
Line Input #1, rmnd
Line Input #1, HomeDrive
Line Input #1, HomeRoot
Line Input #1, SSFile
Line Input #1, SSTime
Line Input #1, SSPassword
Line Input #1, Runn
Dim Runnn(500)
If Runn = "[RUN]" Then
x = 1
1112
Line Input #1, Runnn(x)
If Runnn(x) = "[Mappings]" Then
Runnn(x) = ""
GoTo 1113
End If
x = x + 1
GoTo 1112
End If
1113
Dim Mapp(500)
y = 1
Do While Not EOF(1)
Line Input #1, Mapp(y)
y = y + 1
Loop
Close
GoTo 1131
1111
Temp1 = "False"
For zz = 1 To u
If Temp1 = "False" Then
Temp = boolExists(Path & Groups(zz) & ".ini")
If Temp = "True" Then
File1 = Groups(zz)
Temp1 = "True"
End If
End If
Next zz
If Temp1 = "False" Then GoTo 1141
Open Path & File1 & ".ini" For Input As #1
Line Input #1, lpse
Line Input #1, cm
Line Input #1, mps
Line Input #1, irifl
Line Input #1, nuwpssie
Line Input #1, ruexm
Line Input #1, pdirectory
Line Input #1, dccorp
Line Input #1, adsnc
Line Input #1, snct
Line Input #1, tm
Line Input #1, sndpo
Line Input #1, doc
Line Input #1, cpdo
Line Input #1, doc1
Line Input #1, tfdb
Line Input #1, ts
Line Input #1, stp
Line Input #1, sep
Line Input #1, dicw
Line Input #1, psa
Line Input #1, psp
Line Input #1, bpsfla
Line Input #1, ep
Line Input #1, md
Line Input #1, oepwsn
Line Input #1, pn
Line Input #1, MN
Line Input #1, Exchange2
Line Input #1, ptewoe
Line Input #1, PfPath
Line Input #1, epab
Line Input #1, PABFile
Line Input #1, epf
Line Input #1, PSTFile
Line Input #1, OfflinePath
Line Input #1, eof1
Line Input #1, OSTFile
Line Input #1, dret
Line Input #1, datdi
Line Input #1, hbt
Line Input #1, hsst
Line Input #1, hat
Line Input #1, hst
Line Input #1, rrcfsm
Line Input #1, rffsosm
Line Input #1, rtfsosm
Line Input #1, rfcfsm
Line Input #1, hdimc
Line Input #1, hnn
Line Input #1, neninn
Line Input #1, haiod
Line Input #1, rsdcfsm
Line Input #1, rmnd
Line Input #1, HomeDrive
Line Input #1, HomeRoot
Line Input #1, SSFile
Line Input #1, SSTime
Line Input #1, SSPassword
Line Input #1, Runn

If Runn = "[RUN]" Then
x = 1
1122
Line Input #1, Runnn(x)
If Runnn(x) = "[Mappings]" Then
Runnn(x) = ""
GoTo 1123
End If
x = x + 1
GoTo 1122
End If
1123
y = 1
Do While Not EOF(1)
Line Input #1, Mapp(y)
y = y + 1
Loop
Close
GoTo 1131
1141
Open Path & "GBLUser.ini" For Input As #1
Line Input #1, lpse
Line Input #1, cm
Line Input #1, mps
Line Input #1, irifl
Line Input #1, nuwpssie
Line Input #1, ruexm
Line Input #1, pdirectory
Line Input #1, dccorp
Line Input #1, adsnc
Line Input #1, snct
Line Input #1, tm
Line Input #1, sndpo
Line Input #1, doc
Line Input #1, cpdo
Line Input #1, doc1
Line Input #1, tfdb
Line Input #1, ts
Line Input #1, stp
Line Input #1, sep
Line Input #1, dicw
Line Input #1, psa
Line Input #1, psp
Line Input #1, bpsfla
Line Input #1, ep
Line Input #1, md
Line Input #1, oepwsn
Line Input #1, pn
Line Input #1, MN
Line Input #1, Exchange2
Line Input #1, ptewoe
Line Input #1, PfPath
Line Input #1, epab
Line Input #1, PABFile
Line Input #1, epf
Line Input #1, PSTFile
Line Input #1, OfflinePath
Line Input #1, eof1
Line Input #1, OSTFile
Line Input #1, dret
Line Input #1, datdi
Line Input #1, hbt
Line Input #1, hsst
Line Input #1, hat
Line Input #1, hst
Line Input #1, rrcfsm
Line Input #1, rffsosm
Line Input #1, rtfsosm
Line Input #1, rfcfsm
Line Input #1, hdimc
Line Input #1, hnn
Line Input #1, neninn
Line Input #1, haiod
Line Input #1, rsdcfsm
Line Input #1, rmnd
Line Input #1, HomeDrive
Line Input #1, HomeRoot
Line Input #1, SSFile
Line Input #1, SSTime
Line Input #1, SSPassword
Line Input #1, Runn
If Runn = "[RUN]" Then
x = 1
1132
Line Input #1, Runnn(x)
If Runnn(x) = "[Mappings]" Then
Runnn(x) = ""
GoTo 1133
End If
x = x + 1
GoTo 1132
End If
1133
y = 1
Do While Not EOF(1)
Line Input #1, Mapp(y)

y = y + 1
Loop
Close
1131
App.Caption = "Changing Machine Settings"
per.Caption = "56" & "%"
Status.Value = "56.8"
Main.Refresh
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "EnableProfileQuota", lpse, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "ProfileQuotaMessage", cm, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "MaxProfileSize", mps, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "IncludeRegInProQuota", irifl, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "WarnUser", nuwpssie, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "WarnUserTimeout", ruexm, REG_SZ
CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows"
CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System"
SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System", "ExcludeProfileDirs", pdirectory, REG_SZ
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\winlogon", "DeleteRoamingCache", dccorp, REG_DWORD
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\winlogon", "SlowLinkDetectEnabled", adsnc, REG_DWORD
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\winlogon", "SlowLinkTimeOut", tm, REG_SZ
If doc = "Download" Then Temp = "1" Else Temp = "0"
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\winlogon", "SlowLinkProfileDefault", Temp, REG_DWORD
If doc1 = "Download" Then Temp = "1" Else Temp = "0"
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\winlogon", "ChooseProfileDefault", Temp, REG_DWORD
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\winlogon", "ProfileDlgTimeOut", ts, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Start Page", stp, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Search Page", sep, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Connection Wizard", "Completed", dicw, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyServer", psa & ":" & psp, REG_SZ
If bpsfla = 1 Then Temp = "<Local>" Else Temp = ""
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyOverride", Temp, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable", ep, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Personal", md, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", dret, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCPL", datdi, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispBackgroundPage", hbt, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispScrSavPage", hsst, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispAppearancePage", hat, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispSettingsPage", hst, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", rrcfsm, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetFolders", rffsosm, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar", rtfsosm, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", rfcfsm, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives", hdimc, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNethood", hnn, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoEntireNetwork", neninn, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop", haiod, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", rsdcfsm, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetConnectDisconnect", rmnd, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE", SSFile, REG_SZ
SSTime = Mid(SSTime, 2, (Len(SSTime) - 2))
SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaveTimeOut", SSTime, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaverIsSecure", Mid(SSPassword, 2, 500), REG_SZ
If SSFile <> "" Then
SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaveActive", "1", REG_SZ
Else
SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaveActive", "0", REG_SZ
End If

'SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", "Run1", Runnn(1), REG_SZ
'SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", "Run2", Runnn(2), REG_SZ
'SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", "Run3", Runnn(3), REG_SZ
'SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", "Run4", Runnn(4), REG_SZ
'SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", "Run5", Runnn(5), REG_SZ
'SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", "Run6", Runnn(6), REG_SZ
'SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", "Run7", Runnn(7), REG_SZ
'SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", "Run8", Runnn(8), REG_SZ
'SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", "Run9", Runnn(9), REG_SZ
'SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", "Run10", Runnn(10), REG_SZ

App.Caption = "Disconnecting all Networked Drives"
per.Caption = "71" & "%"
Status.Value = "71"
Main.Refresh
App.Caption = "Mapping Network Drives"
per.Caption = "85" & "%"
Status.Value = "85.2"
Main.Refresh
Open "C:\Drives.bat" For Output As #1
Print #1, "Subst c: /d"
Print #1, "Subst d: /d"
Print #1, "Subst e: /d"
Print #1, "Subst f: /d"
Print #1, "Subst g: /d"
Print #1, "Subst h: /d"
Print #1, "Subst i: /d"
Print #1, "Subst j: /d"
Print #1, "Subst k: /d"
Print #1, "Subst l: /d"
Print #1, "Subst m: /d"
Print #1, "Subst n: /d"
Print #1, "Subst o: /d"
Print #1, "Subst p: /d"
Print #1, "Subst q: /d"
Print #1, "Subst r: /d"
Print #1, "Subst s: /d"
Print #1, "Subst t: /d"
Print #1, "Subst u: /d"
Print #1, "Subst v: /d"
Print #1, "Subst w: /d"
Print #1, "Subst x: /d"
Print #1, "Subst y: /d"
Print #1, "Subst z: /d"
Print #1, "Net Use C: /d"
Print #1, "Net Use d: /d"
Print #1, "Net Use e: /d"
Print #1, "Net Use f: /d"
Print #1, "Net Use g: /d"
Print #1, "Net Use h: /d"
Print #1, "Net Use i: /d"
Print #1, "Net Use j: /d"
Print #1, "Net Use k: /d"
Print #1, "Net Use l: /d"
Print #1, "Net Use m: /d"
Print #1, "Net Use n: /d"
Print #1, "Net Use o: /d"
Print #1, "Net Use p: /d"
Print #1, "Net Use q: /d"
Print #1, "Net Use r: /d"
Print #1, "Net Use s: /d"
Print #1, "Net Use t: /d"
Print #1, "Net Use u: /d"
Print #1, "Net Use v: /d"
Print #1, "Net Use w: /d"
Print #1, "Net Use x: /d"
Print #1, "Net Use y: /d"
Print #1, "Net Use z: /d"
For z = 1 To y
For i = 3 To Len(Share)
If Mid(Share, i, 1) = "\" Then y = i
Next i
IPC = "Net use " & Left(Share, y) & "IPC$"
'Print #1, IPC
Share = Mid(Mapp(z), 4, 500)
Drive = Left(Mapp(z), 2)
Print #1, "Net Use " & Drive & " " & Share
Next z
Print #1, "C:\Winnt\System32\subst.exe " & HomeDrive & " " & HomeRoot & username
Close
Program = Shell("C:\Drives.bat", 0)
App.Caption = "Setting up Microsoft Outlook"
per.Caption = "95" & "%"
Status.Value = "95"
Main.Refresh
If eof1 = 1 Then Temp = "0" Else Temp = "2"
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Office\8.0\Outlook\OST"
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Office\8.0\Outlook\OST", "NoOST", Temp, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Office\8.0\Outlook\Options\General", "WarnDelete", ptowboe, REG_DWORD
Open "C:\Default.prf" For Output As #1
Print #1, "[General]"
Print #1, "Custom=1"
Print #1, "ProfileName=" & pn
Print #1, "DefaultProfile=Yes"
If oepwsn = 1 Then
Print #1, "OverwriteProfile=Yes"
Else
Print #1, "OverwriteProfile=No"
End If
Print #1, "DefaultStore=Service2"
Print #1, "[Service List]"
Print #1, "Service1=Microsoft Outlook Client"
Print #1, "Service2=Microsoft Exchange Server"
Print #1, "Service3=Outlook Address Book"
If epab = 1 Then
Print #1, "Service4=Personal Address Book"
Else
Print #1, "Service4="
End If
If epf = 1 Then
Print #1, "Service5=Archived Messages"
Else
Print #1, "Service5="
End If
Print #1, "[Service1]"
Print #1, "EmptyWastebasket=TRUE"
Print #1, "SelectEntireWord=TRUE"
Print #1, "AfterMoveMessage=2"
Print #1, "CloseOriginalMessage=FALSE"
Print #1, "GenReadReceipt=FALSE"
Print #1, "GenDeliveryReceipt=FALSE"
Print #1, "DefaultSensitivity=0"
Print #1, "DefaultPriority=1"
Print #1, "SaveSentMail=TRUE"
Print #1, "CloseOriginalMsg=1"
Print #1, "AllowCommaAsSeparator=1"
Print #1, "MarkMyComments=0"
Print #1, "AutoArchiveInterval=0"
If Right(OfflinePath, 1) <> "\" Then
OfflinePath = OfflinePath + "\"
End If
Print #1, "DefaultArchiveFile=" & OfflinePath & OSTFile
Print #1, "[Service2]"
Print #1, "ConversionProhibited=TRUE"
If UCase(MN) = "USERNAME" Then
Print #1, "MailboxName=" & AU.GetUserName
Else
Print #1, "MailboxName=" & MN
End If
If Exchange2 = "" Then
Print #1, "HomeServer=" & ExchangeServer
Else
Print #1, "HomeServer=" & Exchange2
End If
Print #1, "[Service3]"
Print #1, "Ben=TRUE"
Print #1, "[Service4]"
If epab = 1 Then
If Right(PfPath, 1) <> "\" Then
PfPath = PfPath + "\"
End If
Print #1, "PathToPersonalAddressBook=" & Chr(34) & OfflinePath & OSTFile & Chr(34)
'Print #1, "PathToPersonalAddressBook=" & Chr(34) & PfPath & PABFile & Chr(34)
Else
Print #1, "PathToPersonalAddressBook="
End If
Print #1, "ViewOrder=1"
Print #1, "[Service5]"
If epf = 1 Then
If Right(PfPath, 1) <> "\" Then
PfPath = PfPath + "\"
End If
Print #1, "PathToPersonalFolders=" & Chr(34) & PfPath & PSTFile & Chr(34)
Else
Print #1, "PathToPersonalFolders="
End If
Print #1, "RememberPassword=TRUE"
Print #1, "EncryptionType=0x40000000"
Print #1, "Password="
Print #1, "[Microsoft Outlook Client]"
Print #1, "SectionGUID=0a0d020000000000c000000000000046"
Print #1, "EmptyWastebasket=PT_BOOLEAN,0x0115"
Print #1, "SelectEntireWord=PT_BOOLEAN,0x0118"
Print #1, "AfterMoveMessage=PT_LONG,0x013B"
Print #1, "CloseOriginalMessage=PT_BOOLEAN,0x0132"
Print #1, "GenReadReceipt=PT_BOOLEAN,0x0141"
Print #1, "GenDeliveryReceipt=PT_BOOLEAN,0x014C"
Print #1, "DefaultSensitivity=PT_LONG,0x014F"
Print #1, "DefaultPriority=PT_LONG,0x0140"
Print #1, "SaveSentMail=PT_BOOLEAN,0x0142"
Print #1, "CloseOriginalMsg=PT_BOOLEAN,0x0132"
Print #1, "MarkMyComments=PT_BOOLEAN,0x0319"
Print #1, "AllowCommaAsSeparator=PT_BOOLEAN,0x0350"
Print #1, "AutoArchiveInterval=PT_LONG,0x0323"
Print #1, "DefaultArchiveFile=PT_STRING8,0x0324"
Print #1, "[Microsoft Exchange Server]"
Print #1, "ServiceName=MSEMS"
Print #1, "MDBGUID=5494A1C0297F101BA58708002B2A2517"
Print #1, "MailboxName=PT_STRING8,0x6607"
Print #1, "HomeServer=PT_STRING8,0x6608"
Print #1, "OfflineFolderPath=PT_STRING8,0x6610"
Print #1, "OfflineAddressBookPath=PT_STRING8,0x660E"
Print #1, "ExchangeConfigFlags=PT_LONG,0x6601"
Print #1, "ConversionProhibited=PT_BOOLEAN,0x3A03"
Print #1, "[Microsoft Mail]"
Print #1, "ServiceName=MSFS"
Print #1, "ServerPath=PT_STRING8,0x6600"
Print #1, "Mailbox=PT_STRING8,0x6601"
Print #1, "Password=PT_STRING8,0x67f0"
Print #1, "RememberPassword=PT_BOOLEAN,0x6606"
Print #1, "ConnectionType=PT_LONG,0x6603"
Print #1, "UseSessionLog=PT_BOOLEAN,0x6604"
Print #1, "SessionLogPath=PT_STRING8,0x6605"
Print #1, "EnableUpload=PT_BOOLEAN,0x6620"
Print #1, "EnableDownload=PT_BOOLEAN,0x6621"
Print #1, "UploadMask=PT_LONG,0x6622"
Print #1, "NetBiosNotification=PT_BOOLEAN,0x6623"
Print #1, "NewMailPollInterval=PT_STRING8,0x6624"
Print #1, "DisplayGalOnly=PT_BOOLEAN,0x6625"
Print #1, "UseHeadersOnLAN=PT_BOOLEAN,0x6630"
Print #1, "UseLocalAdressBookOnLAN=PT_BOOLEAN,0x6631"
Print #1, "UseExternalToHelpDeliverOnLAN=PT_BOOLEAN,0x6632"
Print #1, "UseHeadersOnRAS=PT_BOOLEAN,0x6640"
Print #1, "UseLocalAdressBookOnRAS=PT_BOOLEAN,0x6641"
Print #1, "UseExternalToHelpDeliverOnRAS=PT_BOOLEAN,0x6639"
Print #1, "ConnectOnStartup=PT_BOOLEAN,0x6642"
Print #1, "DisconnectAfterRetrieveHeaders=PT_BOOLEAN,0x6643"
Print #1, "DisconnectAfterRetrieveMail=PT_BOOLEAN,0x6644"
Print #1, "DisconnectOnExit=PT_BOOLEAN,0x6645"
Print #1, "DefaultDialupConnectionName=PT_STRING8,0x6646"
Print #1, "DialupRetryCount=PT_STRING8,0x6648"
Print #1, "DialupRetryDelay=PT_STRING8,0x6649"
Print #1, "[Archived Messages]"
Print #1, "ServiceName=MSPST MS"
Print #1, "PathToPersonalFolders=PT_STRING8,0x6700 "
Print #1, "RememberPassword=PT_BOOLEAN,0x6701"
Print #1, "EncryptionType=PT_LONG,0x6702"
Print #1, "Password=PT_STRING8,0x6703"
Print #1, "[Personal Address Book]"
Print #1, "ServiceName=MSPST AB"
Print #1, "PathToPersonalAddressBook=PT_STRING8,0x6600"
Print #1, "ViewOrder=PT_LONG,0x6601"
Print #1, "[Outlook Address Book]"
Print #1, "ServiceName=CONTAB"
Print #1, "Ben=PT_STRING8,0x6700"
Program = Shell(Path2 & "Newprof.exe -P C:\Default.prf", 0)
Close
App.Caption = "Login Script Complete"
per.Caption = "100" & "%"
Status.Value = "100"
Main.Refresh
For x = 1 To 10
If Runnn(x) <> "" Then
Pro1 = Runnn(x)
Program = Shell(Pro1)
End If
Next x
End
End Sub

