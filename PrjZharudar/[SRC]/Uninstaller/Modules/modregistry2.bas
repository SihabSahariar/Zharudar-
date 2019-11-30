Attribute VB_Name = "modregistry2"
'---------------------------------------------------------------------------------------
' Nama File  : modregistry2.bas
' Tanggal    : 8/29/2005 22:27
' Programmer : Rusman Indradi (rusman@olivault.com)
' Lokasi     : Bogor, INDONESIA
' Catatan    : Rusman Indradi ekeur stres Gw Teh euY... untuk sapa yach program ini..
'              ok deCh untuk Temen gw saudara gw yayang GW CroTZ selalu.... :)
'              tHanKz tO Rizki Priatna, Abby, Ronny, pon-pon, Maryam thaNk's for
'              yOur support Euy..... Hapy CodinG and dont forGEt me Ok....
'              unTuk mAryam And pon-pon kapan Ceng-Ceng lg euY......
'
' Website    : wwww.olivault.com
' Contact HP : ?
' E-mail     : intouch@olivault.com
'
'                                  Roes Love Maryam
'
'Note       : This Code Source is destined to You which wish to learn
'             programming.by using is Visual Basic 6.0. If You use this code source,
'             expect that remain to mention the name of me in part of Your About
'             application( Credit Title) as well as in part of Your place code source
'             using it ( IDEA of VB6). Usage of code source for the purpose of is
'             commercial / profit, HAVE TO PERMIT OF its OWNER.
'             Trespasser- an of this thing can be ensnared by penalization
'             related to misdemeanour of Copyrights and [Code/Law] Rights Of Intellectual.
'---------------------------------------------------------------------------------------
Option Explicit

Public Const REG_NONE = 0                       ' No value type
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Public Const REG_BINARY = 3                     ' Free form binary
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Public Const REG_LINK = 6                       ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Public Const REG_RESOURCE_LIST = 8              ' Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10

Public Enum nKeyTypes
    nKeyTypes_STRING = 1
    nKeyTypes_DWORD = 4
End Enum

Public Enum hKeyNames
    HKEY_CLASSES_ROOT = &H80000000


    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

Public Const ERROR_SUCCESS = 0&
Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259

Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = &H20000
Private Const STANDARD_RIGHTS_WRITE = &H20000
Private Const STANDARD_RIGHTS_EXECUTE = &H20000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Public Const KEY_ALL_ACCESS = &H3F

Public Const REG_OPTION_NON_VOLATILE = 0

'INI File Functions
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

'Registry Functions
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

'To show the system settings cahnged restart dialogue
Public Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long

'Initialize common controls to use the new windows XP themes
Public Declare Sub InitCommonControls Lib "comctl32" ()

'Environment Functions
Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Registry Functions
Public Sub EnumRegKeys(ByRef returnName As Collection, Optional ByRef returnSubs As Collection, Optional hKeyName As String = "HKEY_CURRENT_USER", Optional keyName As String = "SOFTWARE", Optional ByVal checkForSubs As Boolean = False)
    Dim lRetVal As Long      'result of the API functions
    Dim lngResult2 As Long      'result of the API functions
    Dim hKey2 As Long
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant    'setting of queried value
    Dim lngKeyHandle As Long
    Dim lngResult As Long
    Dim lngCurIdx As Long
    Dim strValue As String
    Dim lngValueLen As Long
    Dim lngData As Long
    Dim lngDataLen As Long
    Dim strResult As String
    Dim lKeyName As Long
    Dim SubLevel As Boolean

    Set returnName = New Collection
    Set returnSubs = New Collection
    
    keyName = CompileKeyString(keyName)
    
    lKeyName = resolveHkeyLong(hKeyName)
    
    Do
        lRetVal = RegOpenKeyEx(lKeyName, keyName, 0, KEY_READ, hKey)
        lngValueLen = 2000
        strValue = String(lngValueLen, 0)
        lngDataLen = 2000
        lngResult = RegEnumKey(hKey, lngCurIdx, ByVal strValue, lngValueLen)
        lngCurIdx = lngCurIdx + 1
        RegCloseKey (hKey)
        
        If lngResult = ERROR_SUCCESS Then
            strResult = Left(strValue, lngValueLen)
            If InStr(1, strResult, Chr(0) & Chr(0) & Chr(0) & Chr(0), vbTextCompare) <> 0 Then
                strResult = Mid(strResult, 1, InStr(1, strResult, Chr(0) & Chr(0) & Chr(0) & Chr(0), vbTextCompare) - 1)
            Else
                strResult = strResult
            End If
            If checkForSubs = True Then
                If keyName = "" Then
                    lngResult2 = RegOpenKeyEx(lKeyName, strResult, 0, KEY_READ, hKey2)
                Else
                    lngResult2 = RegOpenKeyEx(lKeyName, keyName & "\" & strResult, 0, KEY_READ, hKey2)
                End If
                strValue = String(lngValueLen, 0)
                lngResult2 = RegEnumKey(hKey2, 0, ByVal strValue, lngValueLen)
                RegCloseKey (hKey2)
                If lngResult2 = ERROR_SUCCESS Then
                    SubLevel = True
                Else
                    SubLevel = False
                End If
                returnSubs.Add SubLevel
            End If
            returnName.Add strResult
        End If
    Loop While lngResult = ERROR_SUCCESS
End Sub

Public Sub EnumRegValues(ByRef returnName As Collection, Optional ByRef returnData As Collection, Optional ByRef returnType As Collection, Optional hKeyName As String = "HKEY_CURRENT_USER", Optional keyName As String = "SOFTWARE")
    Dim lRetVal As Long      'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim hKey2 As Long
    Dim vValue As Variant    'setting of queried value
    Dim Count As Integer
    Dim lngKeyHandle As Long
    Dim lngResult As Long
    Dim lngCurIdx As Long
    Dim strValue As String
    Dim lngValueLen As Long
    Dim lngData As Long
    Dim lngDataLen As Long
    Dim strResult As String
    Dim lKeyName As Long
    Dim retData As String
    Dim retType As Long

    lKeyName = resolveHkeyLong(hKeyName)

    Set returnName = New Collection
    Set returnData = New Collection
    Set returnType = New Collection
    
    keyName = CompileKeyString(keyName)
    
    lRetVal = RegOpenKeyEx(lKeyName, keyName, 0, KEY_READ, hKey)

    Do
        lngValueLen = 2000
        strValue = String(lngValueLen, 0)
        lngDataLen = 2000
        lngResult = RegEnumValue(hKey, lngCurIdx, ByVal strValue, lngValueLen, 0&, REG_DWORD, ByVal lngData, lngDataLen)
        lngCurIdx = lngCurIdx + 1
        If lngResult = ERROR_SUCCESS Then
            strResult = Left(strValue, lngValueLen)
            Call returnName.Add(strResult)
            Call RegOpenKeyEx(lKeyName, keyName, 0, KEY_ALL_ACCESS, hKey2)
            Call QueryValueEx(hKey2, strResult, retData, retType)
            Call RegCloseKey(hKey2)
            Call returnData.Add(retData)
            Call returnType.Add(retType)
        End If
    Loop While lngResult = ERROR_SUCCESS

    RegCloseKey (hKey)
End Sub

Public Function GetSetting(AppName As String, SECTION As String, KEY As String, Optional default As String, Optional hKeyName As hKeyNames = HKEY_LOCAL_MACHINE, Optional AppNameHeader As String = "SOFTWARE", Optional openReadWrite As Boolean = False)
    Dim lRetVal As Long      'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant    'setting of queried value
    Dim keyString As String
    
        On Error GoTo e_Trap
        
        keyString = CompileKeyString(AppNameHeader, AppName, SECTION)
    
        If openReadWrite = True Then
            lRetVal = RegOpenKeyEx(hKeyName, keyString, 0, KEY_ALL_ACCESS, hKey)
        Else
            lRetVal = RegOpenKeyEx(hKeyName, keyString, 0, KEY_READ, hKey)
        End If
        lRetVal = QueryValueEx(hKey, KEY, vValue)
        If IsEmpty(vValue) Or Trim(vValue) = "" Then
            vValue = default
        End If
        GetSetting = vValue
        RegCloseKey (hKey)
        Exit Function
e_Trap:
        vValue = default
        Exit Function
End Function

Public Function SaveSetting(AppName As String, SECTION As String, KEY As String, Setting As String, Optional hKeyName As hKeyNames = HKEY_LOCAL_MACHINE, Optional AppNameHeader As String = "SOFTWARE", Optional KeyType As nKeyTypes = nKeyTypes_STRING) As Boolean
    Dim lRetVal As Long       'result of the SetValueEx function
    Dim hKey As Long          'handle of open key
    Dim keyString As String
    
        On Error GoTo e_Trap
        
        keyString = CompileKeyString(AppNameHeader, AppName, SECTION)
    
        lRetVal = RegCreateKeyEx(hKeyName, keyString, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = SetValueEx(hKey, KEY, KeyType, Setting)
        RegCloseKey (hKey)
        SaveSetting = True
        Exit Function
e_Trap:
        SaveSetting = False
        Exit Function
End Function
Public Function DeleteSetting(AppName As String, Optional SECTION As String, Optional KEY As String, Optional hKeyName As hKeyNames = HKEY_LOCAL_MACHINE, Optional AppNameHeader As String = "SOFTWARE", Optional recurseSubs As Boolean = True) As Boolean
    Dim hNewKey As Long       'handle to the new key
    Dim lRetVal As Long       'result of the SetValueEx function
    Dim hKey As Long          'handle of open key
    Dim keyString As String
    Dim returnName As Collection
    Dim returnSubs As Collection
    Dim Count As Integer
    
        On Error GoTo e_Trap
        
        keyString = CompileKeyString(AppNameHeader, AppName, SECTION)
        
        If KEY <> "" Then
            lRetVal = RegCreateKeyEx(hKeyName, keyString, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
            lRetVal = RegDeleteValue(hKey, KEY)
            RegCloseKey (hKey)
        Else
            lRetVal = RegDeleteKey(hKeyName, keyString)
            If lRetVal = ERROR_CANTWRITE Then
                Call EnumRegKeys(returnName, returnSubs, resolveHkeyString(hKeyName), keyString)
                For Count = 1 To returnName.Count
                    Call DeleteSetting(keyString & "\" & returnName(Count), "", "", hKeyName, "")
                Next Count
                lRetVal = RegDeleteKey(hKeyName, keyString)
            End If
        End If
        If lRetVal = ERROR_SUCCESS Then
            DeleteSetting = True
        Else
            DeleteSetting = False
        End If
        Exit Function
e_Trap:
        DeleteSetting = False
        Exit Function
End Function

Public Function AssociateFileType(extension As String, Optional useNotepadToEdit As Boolean = True) As Boolean
    Dim lRetVal As Long       'result of the SetValueEx function
    Dim hKey As Long          'handle of open key
    Dim appPath As String
    
        On Error GoTo e_Trap
              
        If Mid(App.Path, Len(App.Path) - 1, 1) = "\" Then
            appPath = App.Path & App.EXEName & ".exe"
        Else
            appPath = App.Path & "\" & App.EXEName & ".exe"
        End If
        
        lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, App.Title, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = SetValueEx(hKey, "", REG_SZ, App.Title & " App")
        RegCloseKey (hKey)
    
        lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, "." & LCase(extension), 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = SetValueEx(hKey, "", REG_SZ, App.Title)
        RegCloseKey (hKey)
        
        lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, App.Title & "\shell\open\command", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = SetValueEx(hKey, "", REG_SZ, appPath & " %1")
        RegCloseKey (hKey)
        
        If useNotepadToEdit = True Then
            lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, App.Title & "\shell\edit\command", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
            lRetVal = SetValueEx(hKey, "", REG_SZ, "notepad.exe %1")
            RegCloseKey (hKey)
        End If
        
        lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, App.Title & "\DefaultIcon", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = SetValueEx(hKey, "", REG_SZ, appPath)
        RegCloseKey (hKey)
    
        AssociateFileType = True
        Exit Function
e_Trap:
        AssociateFileType = False
        Exit Function
End Function

Public Sub CreateRunOnStartup(Optional commandLine As String, Optional hKeyName As hKeyNames = HKEY_LOCAL_MACHINE)
    If commandLine <> "" Then
        commandLine = " " & commandLine
    End If
    Call SaveSetting("CurrentVersion", "Run", App.Title, App.Path & "\" & App.EXEName & ".exe" & commandLine, hKeyName, "Software\Microsoft\Windows")
End Sub

Public Sub DeleteRunOnStartup(Optional hKeyName As hKeyNames = HKEY_LOCAL_MACHINE)
    Call DeleteSetting("CurrentVersion", "Run", App.Title, hKeyName, "Software\Microsoft\Windows")
End Sub




' INI Functions
Public Function GetIniInt(SECTION As String, KEY As String, IniLocation As String, Optional default As Long) As Long
    GetIniInt = GetPrivateProfileInt(SECTION, KEY, default, IniLocation)
End Function

Public Function GetIniString(SECTION As String, KEY As String, IniLocation As String, Optional default As String) As String
    Dim ReturnValue As String * 128
    Dim i, sLet
    Dim iLen As Long
    Dim Length As Long
            Length = GetPrivateProfileString(SECTION, KEY, default, ReturnValue, 128, IniLocation)
            i = InStr(1, Trim(ReturnValue), Chr(0))
            iLen = Len(Trim(ReturnValue))
            GetIniString = CStr(Left(Trim(ReturnValue), (i - 1)))
End Function

Public Function SaveIniString(SECTION As String, KEY As String, Setting As String, IniLocation As String) As Long
    SaveIniString = WritePrivateProfileString(SECTION, KEY, Setting, IniLocation)
End Function

' Environment Function
'Public Property Get Environ(variableName As String) As String
'    Environ = GetSetting("Session Manager", "Environment", variableName, "", HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control")
'End Property
'Public Property Let Environ(variableName As String, Setting As String)
'    Call SaveSetting("Session Manager", "Environment", variableName, Setting, HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control")
'    Call SetEnvironmentVariable(variableName, Setting)
'End Property
Public Sub VerifyPath(pathString As String)
    Dim CurrentPath As String
    
        pathString = Trim(pathString)
        If pathString = "" Then Exit Sub
        
        CurrentPath = Environ("PATH")
        If Mid(pathString, 1, 1) = ";" Then
            pathString = Mid(pathString, 2)
        End If
        If Mid(pathString, Len(pathString), 1) = ";" Then
            pathString = Mid(pathString, 1, Len(pathString) - 1)
        End If
        If InStr(1, UCase(CurrentPath), UCase(pathString), vbTextCompare) = 0 Then
            If Mid(CurrentPath, Len(CurrentPath), 1) = ";" Then
                Environ("PATH") = CurrentPath & pathString
            Else
                Environ("PATH") = CurrentPath & ";" & pathString
            End If
        End If
End Sub

Public Function resolveHkeyLong(hKeyName As String) As hKeyNames
    Select Case UCase(hKeyName)
        Case "HKEY_CURRENT_USER"
            resolveHkeyLong = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE"
            resolveHkeyLong = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            resolveHkeyLong = HKEY_USERS
        Case "HKEY_CLASSES_ROOT"
            resolveHkeyLong = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_CONFIG"
            resolveHkeyLong = HKEY_CURRENT_CONFIG
    End Select
End Function
Public Function resolveHkeyString(hKeyName As hKeyNames) As String
    Select Case UCase(hKeyName)
        Case HKEY_CURRENT_USER
            resolveHkeyString = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE
            resolveHkeyString = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS
            resolveHkeyString = "HKEY_USERS"
        Case HKEY_CLASSES_ROOT
            resolveHkeyString = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_CONFIG
            resolveHkeyString = "HKEY_CURRENT_CONFIG"
    End Select
End Function

' Private Functions
Private Function CompileKeyString(Optional AppNameHeader As String, Optional AppName As String, Optional SECTION As String) As String
    If AppNameHeader <> "" Then
        CompileKeyString = AppNameHeader
    End If
    If AppName <> "" Then
        If CompileKeyString <> "" Then
            CompileKeyString = CompileKeyString & "\"
        End If
        CompileKeyString = CompileKeyString & AppName
    End If
    If SECTION <> "" Then
        If CompileKeyString <> "" Then
            CompileKeyString = CompileKeyString & "\"
        End If
        CompileKeyString = CompileKeyString & SECTION
    End If
    Do While InStr(1, CompileKeyString, "\\", vbTextCompare) <> 0
        If InStr(1, CompileKeyString, "\\", vbTextCompare) <> 0 Then
            CompileKeyString = Mid(CompileKeyString, 1, InStr(1, CompileKeyString, "\\", vbTextCompare) - 1) & Mid(CompileKeyString, InStr(1, CompileKeyString, "\\", vbTextCompare) + 1)
        End If
    Loop

    Do While InStr(1, CompileKeyString, "/", vbTextCompare) <> 0
        If InStr(1, CompileKeyString, "/", vbTextCompare) <> 0 Then
            CompileKeyString = Mid(CompileKeyString, 1, InStr(1, CompileKeyString, "/", vbTextCompare) - 1) & "\" & Mid(CompileKeyString, InStr(1, CompileKeyString, "/", vbTextCompare) + 1)
        End If
    Loop

End Function
Private Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ, REG_EXPAND_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD, REG_DWORD_BIG_ENDIAN
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select
End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValuename As String, vValue As Variant, Optional datatype As Long) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    Dim Count As Integer
    Dim Holder As String
    Dim NewVal As String

    On Error GoTo QueryValueExError
    vValue = ""
    
    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValuename, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    datatype = lType
    Select Case lType
        ' For strings
        Case REG_SZ, REG_EXPAND_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValuename, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        ' For DWORDS
        Case REG_DWORD, REG_DWORD_BIG_ENDIAN:
            lrc = RegQueryValueExLong(lhKey, szValuename, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case REG_BINARY
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValuename, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                Holder = Left$(sValue, cch - 1)
                vValue = ""
                For Count = 1 To Len(Holder)
                    NewVal = Format(Hex(Asc(Mid(Holder, Count, 1))), "00")
                    If Len(NewVal) = 1 Then
                        NewVal = "0" & NewVal
                    End If
                    vValue = vValue & NewVal & " "
                Next Count
                vValue = Trim(vValue)
            Else
                vValue = Empty
            End If
            
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



