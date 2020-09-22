Attribute VB_Name = "Settings"
Option Explicit
Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259
Global Const KEY_ALL_ACCESS = &H3F
Global Const REG_OPTION_NON_VOLATILE = 0



Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)

Public Declare Function Connect Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszstrPassword As String, ByVal lpszLocalName As String) As Long
' Example | Program = Connect("\\ServerName\ShareName","Password","LocalDrive:")
Public Declare Function MapDisconnect Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long

Public Type SHITEMID    'Browse Dialog
   cb             As Long
   abID           As Byte
End Type

Public Type ITEMIDLIST  'Browse Dialog
   mkid           As SHITEMID
End Type

Public Type BROWSEINFO  'Browse Dialog
   hOwner         As Long
   pidlRoot       As Long
   pszDisplayName As String
   lpszTitle      As String
   ulFlags        As Long
   lpfn           As Long
   lParam         As Long
   iImage         As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_BROWSEINCLUDEFILES = &H4000
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long


' Example | Program = Disconnect ("Drivename:",0)
Sub Main()
    'Examples of each function:
    'CreateNewKey HKEY_CURRENT_USER, "TestKey\SubKey1\SubKey2"
    'SetKeyValue HKEY_CURRENT_USER, "TestKey\SubKey1", "Test", "Testing, Testing", REG_SZ
    'MsgBox QueryValue(HKEY_CURRENT_USER, "TestKey\SubKey1", "Test")
    'DeleteKey HKEY_CURRENT_USER, "TestKey\SubKey1\SubKey2"
    'DeleteValue HKEY_CURRENT_USER, "TestKey\SubKey1", "Test"
End Sub


Public Function DeleteKey(lPredefinedKey As Long, sKeyName As String)
' Description:
'   This Function will Delete a key
'
' Syntax:
'   DeleteKey Location, KeyName
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is name of the key you wish to delete, it may include subkeys (example "Key1\SubKey1")


    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hKey As Long         'handle of open key
    
    'open the specified key
    
    'lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)
    'RegCloseKey (hKey)
End Function

Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
' Description:
'   This Function will delete a value
'
' Syntax:
'   DeleteValue Location, KeyName, ValueName
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is the name of the key that the value you wish to delete is in
'   , it may include subkeys (example "Key1\SubKey1")
'
'   ValueName is the name of value you wish to delete

       Dim lRetVal As Long         'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       'open the specified key

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = RegDeleteValue(hKey, sValueName)
       RegCloseKey (hKey)
End Function

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String

    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select

End Function





Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError



    ' Determine the size and type of data to be read

    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If

        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:

    QueryValueEx = lrc
    Exit Function

QueryValueExError:

    Resume QueryValueExExit

End Function
Public Function CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
' Description:
'   This Function will create a new key
'
' Syntax:
'   QueryValue Location, KeyName
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is name of the key you wish to create, it may include subkeys (example "Key1\SubKey1")

    
    
    Dim hNewKey As Long         'handle to the new key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function
    
    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Function



Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
' Description:
'   This Function will set the data field of a value
'
' Syntax:
'   QueryValue Location, KeyName, ValueName, ValueSetting, ValueType
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is the key that the value is under (example: "Key1\SubKey1")
'
'   ValueName is the name of the value you want create, or set the value of (example: "ValueTest")
'
'   ValueSetting is what you want the value to equal
'
'   ValueType must equal either REG_SZ (a string) Or REG_DWORD (an integer)

       Dim lRetVal As Long         'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       'open the specified key

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hKey)

End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
' Description:
'   This Function will return the data field of a value
'
' Syntax:
'   Variable = QueryValue(Location, KeyName, ValueName)
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is the key that the value is under (example: "Software\Microsoft\Windows\CurrentVersion\Explorer")
'
'   ValueName is the name of the value you want to access (example: "link")

       Dim lRetVal As Long         'result of the API functions
       Dim hKey As Long         'handle of opened key
       Dim vValue As Variant      'setting of queried value


       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       'MsgBox vValue
       QueryValue = vValue
       RegCloseKey (hKey)
End Function






'PURPOSE:   Displays a dialog that allows the user to select a directory from a tree
'ARGUMENTS:
'   IN szPrompt:        Prompt to display
'RETURNS:   The path selected, or "" if cancel was selected
Public Function BrowseForFolder(szPrompt As String) As String
   Dim bi As BROWSEINFO
   Dim pidl As Long
   Dim nRet As Long
   Dim szPath As String
   
   szPath = Space$(512)
   
   'Fill struct
   bi.hOwner = 0&
   bi.pidlRoot = 0&
   bi.lpszTitle = szPrompt
   bi.ulFlags = BIF_BROWSEINCLUDEFILES
   
   'Display the dialog and get the selected path
   pidl& = SHBrowseForFolder(bi)
   SHGetPathFromIDList ByVal pidl&, ByVal szPath
   
   'Return value
   BrowseForFolder = Trim$(szPath)
End Function

'You can also use those options In combination Or instead of BIF_RETURNONLYDIRS:
'BIF_BROWSEFORCOMPUTER  Only Return computers. If the user selects anything other than a computer, the OK button Is grayed.
'BIF_BROWSEFORPRINTER  Only Return printers. If the user selects anything other than a printer, the OK button Is grayed.
'BIF_BROWSEINCLUDEFILES  The browse dialog will display files As well As folders.
'BIF_DONTGOBELOWDOMAIN  Do Not include network folders below the domain level In the tree view control.
'BIF_EDITBOX  Version 4.71. The browse dialog includes an edit control In which the user can Type the Name of an Item.
'BIF_RETURNFSANCESTORS  Only Return file system ancestors. If the user selects anything other than a file system ancestor, the OK button Is grayed.
'BIF_RETURNONLYFSDIRS  Only Return file system directories. If the user selects folders that are Not part of the file system, the OK button Is grayed.
'BIF_STATUSTEXT  Include a status area In the dialog box. The callback Function can Set the status Text by sending messages To the dialog box.
'BIF_VALIDATE  Version 4.71. If the user types an invalid Name into the edit box, the browse dialog will Call the application's BrowseCallbackProc with the 'BFFM_VALIDATEFAILED message. This flag is ignored if BIF_EDITBOX is not specified.



Public Function boolExists(strFile As String)
    boolExists = False
    On Error Resume Next
    boolExists = (FileLen(strFile) = FileLen(strFile))
End Function
