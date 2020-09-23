Attribute VB_Name = "mod_Registry"
Option Explicit
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Project:      General Functions
' Program:      Registry Functions
' Author:       V.A. van den Braken
' Version:      1.1
' Date:         30-07-1997, 02-08-1997
' Copyright:    Copyright © 1997 Deltec BV, Naarden
' Description:  Functions to access/modify/write the Windows Registry
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
    Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, _
    phkResult As Long) As Long
'
Private Declare Function RegEnumValue Lib "advapi32.dll" _
    Alias "RegEnumValueA" _
    (ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpValueName As String, _
    lpcbValueName As Long, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Byte, _
    lpcbData As Long) As Long
'
Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long
'
Private Const ERROR_SUCCESS = 0
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_EVENT = &H1
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
'
' FILETIME structure for use with RegEnumKeyEx
'
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
'
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
'Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'
Enum HKEYS
    vHKEY_CLASSES_ROOT = &H80000000
    vHKEY_CURRENT_USER = &H80000001
    vHKEY_LOCAL_MACHINE = &H80000002
    vHKEY_USERS = &H80000003
    vHKEY_PERFORMcANCE_DATA = &H80000004
    vHKEY_CURRENT_CONFIG = &H80000005
    vHKEY_DYN_DATA = &H80000006
End Enum
'
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const REG_OPTION_NON_VOLATILE As Long = 0       ' Key is preserved when system is rebooted
Private Const REG_SZ As Long = 1                        ' Unicode null terminated string
'


Public Function REGDeleteSetting( _
    ByVal regHKEY As HKEYS, _
    ByVal sSection As String, _
    Optional ByVal sKey As String) As Boolean
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    ' Purpose: Delete Section/Key from Registry
    '
    '  REGDeleteSetting vHKEY_USERS,"Section"
    '    Deletes "HKEY_USER\Section\"
    '    from the registry and all Key and Values under the section
    '
    '  REGDeleteSetting vHKEY_USERS,"Section1\Section2"
    '    same but deletes all from "HKEY_USERS\Section1\Section2"
    '
    '  REGDeleteSetting vHKEY_USERS,"Section",Key"
    '    Deletes "HKEY_USER\Section\Key"
    '    from the registry and Values under the key
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    Dim lReturn As Long
    Dim hKey As Long
    '
    If Len(sKey) Then
        lReturn = RegOpenKeyEx(regHKEY, REGSubKey(sSection), 0&, KEY_ALL_ACCESS, hKey)
        '
        If lReturn = ERROR_SUCCESS Then
            If sKey = "*" Then sKey = vbNullString
            lReturn = RegDeleteValue(hKey, sKey)
        End If
    Else
        lReturn = RegOpenKeyEx(regHKEY, REGSubKey(), 0&, KEY_ALL_ACCESS, hKey)
        '
        If lReturn = ERROR_SUCCESS Then
            lReturn = RegDeleteKey(hKey, sSection)
        End If
    End If
    '
    REGDeleteSetting = (lReturn = ERROR_SUCCESS)
    '
End Function

Public Function REGGetSetting( _
    ByVal regHKEY As HKEYS, _
    ByVal sSection As String, _
    ByVal sKey As String, _
    Optional ByVal sDefault As String) As String
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    ' Purpose: Gets Values from the registry
    '
    ' REGGetSetting vHKEY_CURRENT_USER,"Section","Key","DefaultStringWhenEmpty"
    '   Gets Value from "HKEY_CURRENT_USER\Section\Key"
    '   When empty it returns the omitted default("DefaultStringWhenEmpty")
    '   or an empty string when no default is specified
    '
    ' REGGetSetting vHKEY_CURRENT_USER,"Section1\Section2","Key","DefaultStringWhenEmpty"
    '   same but gets value from "HKEY_CURRENT_USER\Section1\Section2\Key"
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    Dim lReturn As Long
    Dim hKey As Long
    Dim lType As Long
    Dim lBytes As Long
    Dim sBuffer As String
    '
    REGGetSetting = sDefault
    '
    lReturn = RegOpenKeyEx(regHKEY, REGSubKey(sSection), 0&, KEY_ALL_ACCESS, hKey)
    If lReturn = 5 Then '
        lReturn = RegOpenKeyEx(regHKEY, REGSubKey(sSection), 0&, KEY_EXECUTE, hKey)
    End If
    '
    If lReturn = ERROR_SUCCESS Then
        If sKey = "*" Then
            sKey = vbNullString
        End If
        lReturn = RegQueryValueEx(hKey, sKey, 0&, lType, ByVal sBuffer, lBytes)
        If lReturn = ERROR_SUCCESS Then
            If lBytes > 0 Then
                sBuffer = Space$(lBytes)
                lReturn = RegQueryValueEx(hKey, sKey, 0&, lType, ByVal sBuffer, Len(sBuffer))
                If lReturn = ERROR_SUCCESS Then
                    REGGetSetting = Left$(sBuffer, lBytes - 1)
                End If
            End If
        End If
    End If
    '
End Function

Public Function REGSaveSetting(ByVal regHKEY As HKEYS, ByVal sSection As String, ByVal sKey As String, ByVal sValue As String) As Boolean
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    ' Purpose: Save Value to the registry
    '
    ' REGSaveSetting vHKEY_CURRENT_USER, "Section", "Key", "Test"
    '   Saves the value "Test" to "HKEY_CURRENT_USER\Section\Key"
    '   And will create the The Sections if they do not exist
    '
    ' REGSaveSetting vHKEY_CURRENT_USER, "Section1\Section2", "Key", "Test"
    '   same but save "to HKEY_CURRENT_USER\Section1\Section2\Key"
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    Dim lRet As Long
    Dim hKey As Long
    Dim lResult As Long
    lRet = RegCreateKeyEx(regHKEY, REGSubKey(sSection), 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, lResult)
    If lRet = ERROR_SUCCESS Then
        If sKey = "*" Then sKey = vbNullString
        lRet = RegSetValueEx(hKey, sKey, 0&, REG_SZ, ByVal sValue, Len(sValue))
        Call RegCloseKey(hKey)
    End If
    REGSaveSetting = (lRet = ERROR_SUCCESS)
End Function

Private Function REGSubKey(Optional ByVal sSection As String) As String
    If Left$(sSection, 1) = "\" Then
        sSection = Mid$(sSection, 2)
    End If
    If Right$(sSection, 1) = "\" Then
        sSection = Mid$(sSection, 1, Len(sSection) - 1)
    End If
    REGSubKey = sSection
End Function

Public Function ListSubValue(PredefinedKey As HKEYS, KeyName As String, Index As Long) As String
On Error GoTo Err
    '
    Dim rc As Long
    Dim hKey As Long
    Dim lpName As String
    Dim lpcbName As Long
    Dim lpReserved As Long
    Dim lpftLastWriteTime As FILETIME
    Dim i As Integer
    '
    ' Make sure there is no backslash preceding the branch
    '
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1)
    End If
    '
    ' Attempt to open the registry
    '
    rc = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    '
    If rc = ERROR_SUCCESS Then
        'Allocate buffers for lpName
        lpcbName = 255
        lpName = String$(lpcbName, Chr(0))
        '
        rc = RegEnumValue(hKey, Index, lpName, lpcbName, 0, ByVal 0&, ByVal 0&, ByVal 0&)
        '
        If rc = ERROR_SUCCESS Then
            ListSubValue = Left(lpName, lpcbName)
        Else
            ListSubValue = ""
        End If
        '
        RegCloseKey hKey
    End If
    '
    Exit Function
    '
Err:
    ListSubValue = ""
End Function
