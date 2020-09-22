Attribute VB_Name = "Module1"
Option Explicit

' ==========================================================
' = Get Time Zone Info                                     =
' ==========================================================

Public Const TIME_ZONE_ID_UNKNOWN = 0
Public Const TIME_ZONE_ID_STANDARD = 1
Public Const TIME_ZONE_ID_DAYLIGHT = 2

Public Declare Function GetTimeZoneInformation Lib "kernel32" _
   (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
   
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

' ==========================================================
' = Get Windows Information                                =
' ==========================================================

Public Const MAX_COMPUTERNAME_LENGTH = 31
Public Const MAX_PATH = 260
Public Const UNLEN = 256

Public Declare Function GetComputerName Lib "kernel32" _
   Alias "GetComputerNameA" (ByVal lpBuffer As String, _
   nSize As Long) As Long
   
Public Declare Function GetSystemDirectory Lib "kernel32" _
   Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
   ByVal nSize As Long) As Long
   
Public Declare Function GetUserName Lib "advapi32.dll" _
   Alias "GetUserNameA" (ByVal lpBuffer As String, _
   nSize As Long) As Long

Public Declare Function GetWindowsDirectory Lib "kernel32" _
   Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
   ByVal nSize As Long) As Long

' ==========================================================
' = Get Environment Information                            =
' ==========================================================

Public Declare Function GetEnvironmentVariable Lib "kernel32" _
   Alias "GetEnvironmentVariableA" (ByVal lpName As String, _
   ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function SetEnvironmentVariable Lib "kernel32" _
   Alias "SetEnvironmentVariableA" (ByVal lpName As String, _
   ByVal lpValue As String) As Long

' ==========================================================
' = Get Process Information                                =
' ==========================================================

'This Type is needed for the following routines, but was
'previously defined in this program
'
'Public Type SYSTEMTIME
'        wYear As Integer
'        wMonth As Integer
'        wDayOfWeek As Integer
'        wDay As Integer
'        wHour As Integer
'        wMinute As Integer
'        wSecond As Integer
'        wMilliseconds As Integer
'End Type

Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Declare Function FileTimeToSystemTime Lib "kernel32" _
   (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Public Declare Function GetCurrentProcess Lib "kernel32" _
   () As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" _
   () As Long
   
Public Declare Function GetPriorityClass Lib "kernel32" _
   (ByVal hProcess As Long) As Long
   
Public Declare Function GetProcessTimes Lib "kernel32" _
   (ByVal hProcess As Long, lpCreationTime As FILETIME, _
   lpExitTime As FILETIME, lpKernelTime As FILETIME, _
   lpUserTime As FILETIME) As Long
   
Public Declare Function GetProcessWorkingSetSize Lib "kernel32" _
   (ByVal hProcess As Long, lpMinimumWorkingSetSize As Long, _
   lpMaximumWorkingSetSize As Long) As Long

' ==========================================================
' = Get Sleep Information                                  =
' ==========================================================

Public Declare Sub Sleep Lib "kernel32" _
   (ByVal dwMilliseconds As Long)

' ==========================================================
' = Get System Color Information                           =
' ==========================================================

Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8

Public Declare Function GetSysColor Lib "user32" _
   (ByVal nIndex As Long) As Long

Public Declare Function SetSysColors Lib "user32" _
   (ByVal nChanges As Long, lpSysColor As Long, _
   lpColorValues As Long) As Long

' ==========================================================
' = Process                                                =
' ==========================================================

Public Const CREATE_NEW_CONSOLE = &H10
Public Const CREATE_NEW_PROCESS_GROUP = &H200
Public Const CREATE_SUSPENDED = &H4
Public Const DEBUG_PROCESS = &H1
Public Const DEBUG_ONLY_THIS_PROCESS = &H2
Public Const DETACHED_PROCESS = &H8
Public Const INFINITE = &HFFFF

Public Const SYNCHRONIZE = &H100000

Public Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Public Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Public Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long

Public Declare Function CreateProcess Lib "kernel32" _
   Alias "CreateProcessA" _
   (ByVal lpApplicationName As String, _
   ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, _
   ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, _
   ByVal lpCurrentDriectory As String, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long

Public Declare Function OpenProcess Lib "kernel32" _
   (ByVal dwDesiredAccess As Long, _
   ByVal bInheritHandle As Long, _
   ByVal dwProcessId As Long) As Long

Public Declare Function WaitForInputIdle Lib "user32" _
   (ByVal hProcess As Long, _
   ByVal dwMilliseconds As Long) As Long
   
Public Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, _
   ByVal dwMilliseconds As Long) As Long

' ==========================================================
' = Run the program associated with a file                 =
' ==========================================================

' Error conditions
Public Const ERROR_BAD_FORMAT = 11&
Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_PATH_NOT_FOUND = 3&
Public Const SE_ERR_ACCESSDENIED = 5
Public Const SE_ERR_ASSOCINCOMPLETE = 27
Public Const SE_ERR_DDEBUSY = 30
Public Const SE_ERR_DDEFAIL = 29
Public Const SE_ERR_DDETIMEOUT = 28
Public Const SE_ERR_DLLNOTFOUND = 32
Public Const SE_ERR_FNF = 2
Public Const SE_ERR_NOASSOC = 31
Public Const SE_ERR_OOM = 8
Public Const SE_ERR_PNF = 3
Public Const SE_ERR_SHARE = 26

' ShowCmd values
Public Const SW_HIDE = 0
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Declare Function ShellExecute Lib "shell32.dll" _
   Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, _
   ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
   
' ==========================================================
' = Get the icon associated with a program                 =
' ==========================================================
   
Public Declare Function DestroyIcon Lib "user32" _
   (ByVal hIcon As Long) As Long

Public Declare Function DrawIcon Lib "user32" _
   (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
   ByVal hIcon As Long) As Long

Public Declare Function ExtractIcon Lib "shell32.dll" _
   Alias "ExtractIconA" _
   (ByVal hInst As Long, ByVal lpszExeFileName As String, _
   ByVal nIconIndex As Long) As Long
   
Public Declare Function FindExecutable Lib "shell32.dll" _
   Alias "FindExecutableA" _
   (ByVal lpFile As String, ByVal lpDirectory As String, _
   ByVal lpResult As String) As Long
   
' ==========================================================
' = Send Message                                           =
' ==========================================================

Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_LIMITTEXT = &H141
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SETCURSEL = &H14E
Public Const CB_SETEDITSEL = &H142
Public Const CB_SETEXTENDEDUI = &H155
Public Const CB_SETITEMDATA = &H151
Public Const CB_SETITEMHEIGHT = &H153
Public Const CB_SETLOCALE = &H159
Public Const CB_SHOWDROPDOWN = &H14F

Public Const EM_CANUNDO = &HC6
Public Const EM_EMPTYUNDOBUFFER = &HCD
Public Const EM_FMTLINES = &HC8
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_GETHANDLE = &HBD
Public Const EM_GETLINE = &HC4
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETWORDBREAKPROC = &HD1
Public Const EM_LIMITTEXT = &HC5
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINESCROLL = &HB6
Public Const EM_SETREADONLY = &HCF
Public Const EM_SETSEL = &HB1
Public Const EM_SETTABSTOPS = &HCB
Public Const EM_SETWORDBREAKPROC = &HD0
Public Const EM_UNDO = &HC7


Public Const LB_FINDSTRING = &H18F
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_GETTOPINDEX = &H18E
Public Const LB_SETTABSTOPS = &H192
Public Const LB_SETTOPINDEX = &H197

Public Const WM_DDE_FIRST = &H3E0
Public Const WM_USER = &H400
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_COMMNOTIFY = &H44
Public Const WM_COPY = &H301
Public Const WM_CUT = &H300
Public Const WM_GETFONT = &H31
Public Const WM_GETHOTKEY = &H33
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_HOTKEY = &H312
Public Const WM_PAINT = &HF
Public Const WM_PASTE = &H302
Public Const WM_SIZE = &H5
Public Const WM_SYSCOMMAND = &H112
Public Const WM_UNDO = &H304

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
   lParam As Any) As Long
   
' ==========================================================
' = Shutdown Windows                                       =
' ==========================================================

Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

Public Declare Function ExitWindowsEx Lib "user32" _
   (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

' ==========================================================
' = Support Functions                                      =
' ==========================================================

Public Function FileTimeToString(f As FILETIME) As String

Dim r As Long
Dim s As SYSTEMTIME
Dim t As String

r = FileTimeToSystemTime(f, s)
If r <> 0 Then
   t = FormatNumber(s.wYear, 0, , , vbFalse)
   t = t & "-" & FormatNumber(s.wMonth, 0)
   t = t & "-" & FormatNumber(s.wDay, 0)
   If t = "1601-1-1" Then
      t = ""
   End If
   t = t & " " & FormatNumber(s.wHour, 0)
   t = t & ":" & FormatNumber(s.wMinute, 0)
   t = t & ":" & FormatNumber(s.wSecond, 0)
   t = t & ":" & FormatNumber(s.wMilliseconds, 0)

Else
   t = "Error in converting value"

End If

FileTimeToString = t

End Function

Public Function RunProg(prog As String) As Long
    
Dim p As PROCESS_INFORMATION
Dim s As STARTUPINFO
Dim r As Long
    
s.cb = Len(s)
s.dwFlags = 0
s.lpDesktop = vbNullString
s.lpReserved = vbNullString
s.lpTitle = vbNullString

r = CreateProcess(prog, vbNullString, 0, 0, True, _
   NORMAL_PRIORITY_CLASS, 0, vbNullString, s, p)
If r <> 0 Then
   r = CloseHandle(p.hThread)
   If r <> 0 Then
      r = WaitForInputIdle(p.hProcess, INFINITE)
      r = CloseHandle(p.hProcess)
   End If
End If

RunProg = p.dwProcessId

End Function

Public Sub WaitProg(Pid As Long)

Dim r As Long
Dim pHandle As Long

pHandle = OpenProcess(SYNCHRONIZE, 0, Pid)

r = WaitForSingleObject(pHandle, INFINITE)

r = CloseHandle(pHandle)

End Sub

Public Function StartMeUp(f As String)

Dim i As Integer
Dim d As String

i = InStrRev(f, "\")
If i > 0 Then
   d = Left(f, i - 1)

Else
   d = App.Path

End If

StartMeUp = ShellExecute(Main.hWnd, "open", f, vbNullString, d, SW_SHOWNORMAL)

End Function

Public Function SendAsyncProc(hWnd As Long, uMsg As Long, dwData As Long, lResult As Long) As Long

MsgBox "uMsg = " & FormatNumber(uMsg, 0) & _
   " dwData = " & FormatNumber(dwData, 0) & _
   " lResult = " & FormatNumber(lResult, 0), "SendAsyncProc"
   
SendAsyncProc = 0

End Function



