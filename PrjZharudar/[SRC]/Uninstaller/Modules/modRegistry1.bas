Attribute VB_Name = "modRegistry1"
'---------------------------------------------------------------------------------------
' Nama File  : modRegistry1.bas
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

Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
End Type
 
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long



Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&

Const REG_NONE = 0&
Const REG_SZ = 1&
Const REG_EXPAND_SZ = 2&
Const REG_BINARY = 3&
Const REG_DWORD = 4&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ

Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte
Dim DisplayErrorMsg As Boolean
Dim i As Integer
Dim GetErrorMsg As String
Function SetDWORDValue(SubKey As String, Entry As String, Value As Long)

    Call ParseKey(SubKey, MainKeyHandle)
    
    If MainKeyHandle Then
        rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
        If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
            rtn = RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4) 'write the value
            If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
                If DisplayErrorMsg = True Then 'if the user want errors displayed
                    MsgBox ErrorMsg(rtn)        'display the error
                End If
            End If
            rtn = RegCloseKey(hKey) 'close the key
        Else 'if there was an error opening the key
            If DisplayErrorMsg = True Then 'if the user want errors displayed
                MsgBox ErrorMsg(rtn) 'display the error
            End If
        End If
    End If

End Function

Function GetDWORDValue(SubKey As String, Entry As String)

    Call ParseKey(SubKey, MainKeyHandle)
    
    If MainKeyHandle Then
       rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
       If rtn = ERROR_SUCCESS Then 'if the key could be opened then
          rtn = RegQueryValueExA(hKey, Entry, 0, REG_DWORD, lBuffer, 4) 'get the value from the registry
          If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
             rtn = RegCloseKey(hKey)  'close the key
             GetDWORDValue = lBuffer  'return the value
          Else                        'otherwise, if the value couldnt be retreived
             GetDWORDValue = "Error"  'return Error to the user
             If DisplayErrorMsg = True Then 'if the user wants errors displayed
                MsgBox ErrorMsg(rtn)        'tell the user what was wrong
             End If
          End If
       Else 'otherwise, if the key couldnt be opened
          GetDWORDValue = "Error"        'return Error to the user
          If DisplayErrorMsg = True Then 'if the user wants errors displayed
             MsgBox ErrorMsg(rtn)        'tell the user what was wrong
          End If
       End If
    End If

End Function

Function SetBinaryValue(SubKey As String, Entry As String, Value As String)

    Call ParseKey(SubKey, MainKeyHandle)
    
    If MainKeyHandle Then
       rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
       If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
          lDataSize = Len(Value)
          ReDim ByteArray(lDataSize)
          For i = 1 To lDataSize
          ByteArray(i) = Asc(Mid$(Value, i, 1))
          Next
          rtn = RegSetValueExB(hKey, Entry, 0, REG_BINARY, ByteArray(1), lDataSize) 'write the value
          If Not rtn = ERROR_SUCCESS Then   'if the was an error writting the value
             If DisplayErrorMsg = True Then 'if the user want errors displayed
                MsgBox ErrorMsg(rtn)        'display the error
             End If
          End If
          rtn = RegCloseKey(hKey) 'close the key
       Else 'if there was an error opening the key
          If DisplayErrorMsg = True Then 'if the user wants errors displayed
             MsgBox ErrorMsg(rtn) 'display the error
          End If
       End If
    End If

End Function

Function GetBinaryValue(SubKey As String, Entry As String)

    Call ParseKey(SubKey, MainKeyHandle)
    
    If MainKeyHandle Then
       rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
       If rtn = ERROR_SUCCESS Then 'if the key could be opened
          lBufferSize = 1
          rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, 0, lBufferSize) 'get the value from the registry
          sBuffer = Space(lBufferSize)
          rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize) 'get the value from the registry
          If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
             rtn = RegCloseKey(hKey)  'close the key
             GetBinaryValue = sBuffer 'return the value to the user
          Else                        'otherwise, if the value couldnt be retreived
             GetBinaryValue = "Error" 'return Error to the user
             If DisplayErrorMsg = True Then 'if the user wants to errors displayed
                MsgBox ErrorMsg(rtn)  'display the error to the user
             End If
          End If
       Else 'otherwise, if the key couldnt be opened
          GetBinaryValue = "Error" 'return Error to the user
          If DisplayErrorMsg = True Then 'if the user wants to errors displayed
             MsgBox ErrorMsg(rtn)  'display the error to the user
          End If
       End If
    End If

End Function


Function GetMainKeyHandle(MainKeyName As String) As Long

    Const HKEY_CLASSES_ROOT = &H80000000
    Const HKEY_CURRENT_USER = &H80000001
    Const HKEY_LOCAL_MACHINE = &H80000002
    Const HKEY_USERS = &H80000003
    Const HKEY_PERFORMANCE_DATA = &H80000004
    Const HKEY_CURRENT_CONFIG = &H80000005
    Const HKEY_DYN_DATA = &H80000006
       
    Select Case MainKeyName
           Case "HKEY_CLASSES_ROOT"
                GetMainKeyHandle = HKEY_CLASSES_ROOT
           Case "HKEY_CURRENT_USER"
                GetMainKeyHandle = HKEY_CURRENT_USER
           Case "HKEY_LOCAL_MACHINE"
                GetMainKeyHandle = HKEY_LOCAL_MACHINE
           Case "HKEY_USERS"
                GetMainKeyHandle = HKEY_USERS
           Case "HKEY_PERFORMANCE_DATA"
                GetMainKeyHandle = HKEY_PERFORMANCE_DATA
           Case "HKEY_CURRENT_CONFIG"
                GetMainKeyHandle = HKEY_CURRENT_CONFIG
           Case "HKEY_DYN_DATA"
                GetMainKeyHandle = HKEY_DYN_DATA
    End Select

End Function

Function ErrorMsg(lErrorCode As Long) As String
    
    'If an error does accurr, and the user wants error messages displayed, then
    'display one of the following error messages
    
    Select Case lErrorCode
           Case 1009, 1015
                GetErrorMsg = "The Registry Database is corrupt!"
           Case 2, 1010
                GetErrorMsg = "Bad Key Name"
           Case 1011
                GetErrorMsg = "Can't Open Key"
           Case 4, 1012
                GetErrorMsg = "Can't Read Key"
           Case 5
                GetErrorMsg = "Access to this key is denied"
           Case 1013
                GetErrorMsg = "Can't Write Key"
           Case 8, 14
                GetErrorMsg = "Out of memory"
           Case 87
                GetErrorMsg = "Invalid Parameter"
           Case 234
                GetErrorMsg = "There is more data than the buffer has been allocated to hold."
           Case Else
                GetErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
    End Select

End Function

Function SetStringValue(SubKey As String, Entry As String, Value As String)

    Call ParseKey(SubKey, MainKeyHandle)
    
    If MainKeyHandle Then
       rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
       If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
          rtn = RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value)) 'write the value
          If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
             If DisplayErrorMsg = True Then 'if the user wants errors displayed
                MsgBox ErrorMsg(rtn)        'display the error
             End If
          End If
          rtn = RegCloseKey(hKey) 'close the key
       Else 'if there was an error opening the key
          If DisplayErrorMsg = True Then 'if the user wants errors displayed
             MsgBox ErrorMsg(rtn)        'display the error
          End If
       End If
    End If

End Function

Function GetStringValue(SubKey As String, Entry As String)

    Call ParseKey(SubKey, MainKeyHandle)
    
    If MainKeyHandle Then
       rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
       If rtn = ERROR_SUCCESS Then 'if the key could be opened then
          sBuffer = Space(255)     'make a buffer
          lBufferSize = Len(sBuffer)
          rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize) 'get the value from the registry
          If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
             rtn = RegCloseKey(hKey)  'close the key
             sBuffer = Trim(sBuffer)
             GetStringValue = Left(sBuffer, Len(sBuffer) - 1) 'return the value to the user
          Else                        'otherwise, if the value couldnt be retreived
             GetStringValue = "Error" 'return Error to the user
             If DisplayErrorMsg = True Then 'if the user wants errors displayed then
                MsgBox ErrorMsg(rtn)  'tell the user what was wrong
             End If
          End If
       Else 'otherwise, if the key couldnt be opened
          GetStringValue = "Error"       'return Error to the user
          If DisplayErrorMsg = True Then 'if the user wants errors displayed then
             MsgBox ErrorMsg(rtn)        'tell the user what was wrong
          End If
       End If
    End If

End Function

Private Sub ParseKey(keyName As String, Keyhandle As Long)
    
    rtn = InStr(keyName, "\") 'return if "\" is contained in the Keyname
    
    If Left(keyName, 5) <> "HKEY_" Or Right(keyName, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
       MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + keyName 'display error to the user
       Exit Sub 'exit the procedure
    ElseIf rtn = 0 Then 'if the Keyname contains no "\"
       Keyhandle = GetMainKeyHandle(keyName)
       keyName = "" 'leave Keyname blank
    Else 'otherwise, Keyname contains "\"
       Keyhandle = GetMainKeyHandle(Left(keyName, rtn - 1)) 'seperate the Keyname
       keyName = Right(keyName, Len(keyName) - rtn)
    End If

End Sub
Function CreateKey(SubKey As String)

    Call ParseKey(SubKey, MainKeyHandle)
    
    If MainKeyHandle Then
       rtn = RegCreateKey(MainKeyHandle, SubKey, hKey) 'create the key
       If rtn = ERROR_SUCCESS Then 'if the key was created then
          rtn = RegCloseKey(hKey)  'close the key
       End If
    End If

End Function

Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
'penambahan registry key delete value
'untuk string
Dim keyhand As Long
Dim rtn
    rtn = RegOpenKey(hKey, strPath, keyhand)
    rtn = RegDeleteValue(keyhand, strValue)
    rtn = RegCloseKey(keyhand)
    
End Function

Function AddSlash(lzPath As String) As String
    If Right(lzPath, 1) = "\" Then AddSlash = lzPath Else AddSlash = lzPath & "\"
    
End Function

Private Function RemoveNulls(lzString As String) As String
Dim XPos As Integer
    XPos = InStr(lzString, vbNullChar)
    If XPos > 0 Then
        lzString = Left(lzString, Len(lzString) - 1)
        RemoveNulls = lzString
    End If
    
End Function

Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call errlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
End Function
Function GetValues(hKey As Long, Value_type As Long, lzPath As String, strValue As String) As String
Dim Value As String
Dim StrLen As Long
On Error Resume Next
    ' The Key you want to open
    If RegOpenKeyEx(hKey, lzPath, 0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    'Get the subkey's Value
    StrLen = 256
    Value = Space(StrLen)
    
    If RegQueryValueEx(hKey, strValue, 0&, Value_type, ByVal Value, StrLen) <> ERROR_SUCCESS Then
       Exit Function
    Else
        ' Remove all trailing null character
        Value = Left(Value, StrLen - 1)
        GetValues = Value
    End If
    
    ' Close the key.
    If RegCloseKey(hKey) <> ERROR_SUCCESS Then
        MsgBox "Error closing key." ' Show error if it happens
    End If
    
End Function
