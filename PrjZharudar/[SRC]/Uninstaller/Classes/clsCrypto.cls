VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsCrypto.cls   Oct 2000
' By: Alex Rohr   arohr@ub2b.com
' Sedikit di Modify by : Rusman Indradi
'___________________________________________________________________________
'  Microsoft Windows
'  Copyright 1992 - 1998 Microsoft Corporation.
'
'  File:       wincrypt.h
'
'  Contents:   Cryptographic API Prototypes and Definitions
'
'----------------------------------------------------------------------------
Option Explicit
Private Const ALG_CLASS_DATA_ENCRYPT           As Integer = 24576
Private Const ALG_CLASS_HASH                   As Long = 32768
' Algorithm types
Private Const ALG_TYPE_ANY                     As Integer = 0
Private Const ALG_TYPE_BLOCK                   As Integer = 1536
' RC2 sub-ids
Private Const ALG_SID_RC2                      As Integer = 2
Private Const ALG_SID_MD5                      As Integer = 3
' algorithm identifier definitions
Private Const CALG_MD5                         As Double = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_RC2                         As Double = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK) Or ALG_SID_RC2)
' dwFlags definitions for CryptAcquireContext
Private Const CRYPT_NEWKEYSET                  As Long = &H8
Private Const HP_HASHVAL                       As Long = &H2  '- Hash value
Private Const HP_HASHSIZE                      As Long = &H4  '- Hash value size
' CryptSetProvParam
'
Private Const PROV_RSA_FULL                    As Integer = 1
' Provider friendly names
'
Private Const MS_DEF_PROV                      As String = "Microsoft Base Cryptographic Provider v1.0"
'WinCrypt API Declarations
' Private property buffers
Private m_sPassword                            As String      ' Password used to create encryption key
Private m_sInBuffer                            As String      ' Used as an input buffer for all
' data to be encrypted or decrypted
Private m_sOutBuffer                           As String      ' Used as an output buffer for all
' data that has been encrypted or decrypted
Private m_sErrorMsg                            As String      ' Error message string
' Private class-level variables
Private m_lHCryptProv                          As Long        ' Handle for the cryptographic
' service provider (CSP)
Private m_lHSessionKey                         As Long        ' Session key for encrypting and
' decrypting data
Private m_sHashBuffer                          As String      ' Used as a buffer for all the hash value
Private m_iHashLength                          As Integer     ' Used to hold the hash length
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As Long, _
                                                                                               ByVal pszContainer As String, _
                                                                                               ByVal pszProvider As String, _
                                                                                               ByVal dwProvType As Long, _
                                                                                               ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, _
                                                                  ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, _
                                                             ByVal Algid As Long, _
                                                             ByVal hBaseData As Long, _
                                                             ByVal dwFlags As Long, _
                                                             phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, _
                                                           ByVal hHash As Long, _
                                                           ByVal Final As Long, _
                                                           ByVal dwFlags As Long, _
                                                           ByVal pbData As String, _
                                                           pdwDataLen As Long, _
                                                           ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, _
                                                           ByVal hHash As Long, _
                                                           ByVal Final As Long, _
                                                           ByVal dwFlags As Long, _
                                                           ByVal pbData As String, _
                                                           pdwDataLen As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, _
                                                              ByVal Algid As Long, _
                                                              ByVal hKey As Long, _
                                                              ByVal dwFlags As Long, _
                                                              phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, _
                                                            ByVal pbData As String, _
                                                            ByVal dwDataLen As Long, _
                                                            ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, _
                                                                ByVal dwParam As Long, _
                                                                ByVal pbData As String, _
                                                                pdwDataLen As Long, _
                                                                ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashDWParam Lib "advapi32.dll" Alias "CryptGetHashParam" (ByVal hHash As Long, _
                                                                                            ByVal dwParam As Long, _
                                                                                            pbData As Long, _
                                                                                            pdwDataLen As Long, _
                                                                                            ByVal dwFlags As Long) As Long

Private Sub Class_Initialize()

    If Not InitUser Then
        m_sErrorMsg = "Unable to initialize CryptoAPI."
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
    End If

End Sub

Private Sub Class_Terminate()

Dim lResult As Long

    If (m_lHCryptProv <> 0) Then
        lResult = CryptReleaseContext(m_lHCryptProv, 0)
    End If

End Sub

Public Function DecryptFileData() As Boolean

Dim strInputData     As String
Dim strHash          As String
Dim strHashLength    As String
Dim lHashLength      As Long

    DecryptFileData = False
    ' decrypt file data
    If Not DecryptMessageData Then
        Exit Function
    End If
    ' extract original hash length
    strHashLength = Left$(m_sOutBuffer, 2)
    lHashLength = CLng(strHashLength)
    ' extract origianl hash value
    strHash = Mid$(m_sOutBuffer, 3, lHashLength)
    ' extract original file contents
    strInputData = Mid$(m_sOutBuffer, (3 + lHashLength))
    ' copy origianl file contents to the input buffer
    m_sInBuffer = strInputData
    ' hash origianl file contents
    If Not HashFile Then
        Exit Function
    End If
    ' compare tht original hash value to the new hash value
    If (strHash = m_sHashBuffer) Then
        ' file was un-altered
        m_sOutBuffer = strInputData
    Else 'NOT (STRHASH...
        ' hash values don't match
        MsgBox "File has been corrupted since original encryption!", vbInformation, "File Error"
        m_sOutBuffer = ""
        Exit Function
    End If
    DecryptFileData = True

End Function

Public Function DecryptMessageData() As Boolean

Dim sCryptBuffer As String
Dim lCryptBufLen As Long

    On Error Resume Next
        DecryptMessageData = False
        ' Prepare sCryptBuffer for CryptDecrypt
        lCryptBufLen = Len(m_sInBuffer)
        sCryptBuffer = String$(lCryptBufLen, vbNullChar)
        LSet sCryptBuffer = m_sInBuffer
        ' Decrypt data
        If Not CBool(CryptDecrypt(m_lHSessionKey, 0, 1, 0, sCryptBuffer, lCryptBufLen)) Then
            m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptDecrypt!"
            MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
            Exit Function
        End If
        ' Apply decrypted string from sCryptBuffer to private buffer for OutBuffer property
        m_sOutBuffer = Mid$(sCryptBuffer, 1, lCryptBufLen)
        DecryptMessageData = True
    On Error GoTo 0

End Function

Public Function DestroySessionKey() As Boolean

    DestroySessionKey = False
    ' Destroy the session key
    If Not CBool(CryptDestroyKey(m_lHSessionKey)) Then
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptDestroyKey!"
        MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
        Exit Function
    End If
    DestroySessionKey = True

End Function

Public Function EncryptFileData() As Boolean

Dim strInputData As String

    EncryptFileData = False
    ' concatenate length of hash, hash value, and origianl file
    strInputData = Trim$(Str$(m_iHashLength)) + m_sHashBuffer + m_sInBuffer
    ' copy new strin to in buffer
    m_sInBuffer = strInputData
    ' call encryptmessagedata
    EncryptFileData = EncryptMessageData

End Function

Public Function EncryptMessageData() As Boolean

Dim lResult      As Long
Dim sCryptBuffer As String
Dim lCryptLength As Long
Dim lCryptBufLen As Long

    On Error Resume Next
        EncryptMessageData = False
        ' Determine the size of the buffer needed for encrypting the data in the InBuffer property
        lCryptLength = Len(m_sInBuffer)
        lResult = CryptEncrypt(m_lHSessionKey, 0, 1, 0, vbNullString, lCryptLength, lCryptBufLen)
        ' Prepare a string buffer for the CryptEncrypt function
        lCryptBufLen = lCryptLength * 2
        lCryptLength = Len(m_sInBuffer)
        sCryptBuffer = String$(lCryptBufLen, vbNullChar)
        ' Copy in the contents of the InBuffer property
        LSet sCryptBuffer = m_sInBuffer
        ' Encrypt data
        If Not CBool(CryptEncrypt(m_lHSessionKey, 0, 1, 0, sCryptBuffer, lCryptLength, lCryptBufLen)) Then
            m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptEncrypt!"
            MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
            Exit Function
        End If
        ' Copy the encrypted data to the OutBuffer property
        m_sOutBuffer = Mid$(sCryptBuffer, 1, lCryptLength)
        EncryptMessageData = True
    On Error GoTo 0

End Function

Public Property Get ErrorMsg() As String

    ErrorMsg = m_sErrorMsg

End Property

Public Function GeneratePasswordKey() As Boolean

Dim lHash As Long

    On Error Resume Next
        GeneratePasswordKey = False
        ' Create an empty hash object.
        If Not CBool(CryptCreateHash(m_lHCryptProv, CALG_MD5, 0, 0, lHash)) Then
            m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptCreateHash!"
            MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
            Exit Function
        End If
        ' Hash the password string.
        If Not CBool(CryptHashData(lHash, m_sPassword, Len(m_sPassword), 0)) Then
            m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptHashData!"
            MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
            Exit Function
        End If
        ' Create a derived block cipher session key.
        If Not CBool(CryptDeriveKey(m_lHCryptProv, CALG_RC2, lHash, 0, m_lHSessionKey)) Then
            m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptDeriveKey!"
            MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
            Exit Function
        End If
        ' Destroy the hash object
        If Not CBool(CryptDestroyHash(lHash)) Then
            m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptDestroyHash!"
            MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
            Exit Function
        End If
        GeneratePasswordKey = True
    On Error GoTo 0

End Function

Public Function HashFile() As Boolean

Dim lHash          As Long
Dim sCryptBuffer   As String
Dim lCryptBufLen   As Long
Dim lCryptHashSize As Long

    On Error Resume Next
        HashFile = False
        ' Create an empty hash object.
        If Not CBool(CryptCreateHash(m_lHCryptProv, CALG_MD5, 0, 0, lHash)) Then
            m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptCreateHash!"
            MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
            Exit Function
        End If
        ' Hash the password string.
        If Not CBool(CryptHashData(lHash, m_sInBuffer, Len(m_sInBuffer), 0)) Then
            m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptHashData!"
            MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
            Exit Function
        End If
        ' Extract the hash value size.
        lCryptBufLen = 4
        If Not CBool(CryptGetHashDWParam(lHash, HP_HASHSIZE, lCryptHashSize, lCryptBufLen, 0)) Then
            m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptGetHashParam!"
            MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
            Exit Function
        End If
        ' Prepare sCryptBuffer for CryptDecrypt
        sCryptBuffer = String$(lCryptHashSize, vbNullChar)
        lCryptBufLen = lCryptHashSize
        ' Extract the hash value.
        If Not CBool(CryptGetHashParam(lHash, HP_HASHVAL, sCryptBuffer, lCryptBufLen, 0)) Then
            m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptGetHashParam!"
            MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
            Exit Function
        End If
        ' Destroy the Hash object
        If Not CBool(CryptDestroyHash(lHash)) Then
            m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptDestroyHash!"
            MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
            Exit Function
        End If
        ' Apply hash string from sCryptBuffer to private buffer for HashBuffer variables
        m_sHashBuffer = Mid$(sCryptBuffer, 1, lCryptBufLen)
        m_iHashLength = lCryptBufLen
        HashFile = True
    On Error GoTo 0

End Function

Public Property Get InBuffer() As String

    InBuffer = m_sInBuffer

End Property

Public Property Let InBuffer(vNewValue As String)

    m_sInBuffer = vNewValue

End Property

Private Function InitUser() As Boolean

Dim sContainer As String
Dim sProvider  As String

    On Error Resume Next
        InitUser = False
        ' Prepare string buffers
        sContainer = vbNullChar
        sProvider = MS_DEF_PROV & vbNullChar
        ' Attempt to acquire a handle to the default key container.
        If Not CBool(CryptAcquireContext(m_lHCryptProv, sContainer, sProvider, PROV_RSA_FULL, 0)) Then
            ' Create default key container.
            If Not CBool(CryptAcquireContext(m_lHCryptProv, sContainer, sProvider, PROV_RSA_FULL, CRYPT_NEWKEYSET)) Then
                m_sErrorMsg = "Error creating key container - " & CStr(Err.LastDllError)
                MsgBox m_sErrorMsg, vbOKOnly, "WinTweaker Encryption"
                Exit Function
            End If
        End If
        InitUser = True
    On Error GoTo 0

End Function

Public Property Get OutBuffer() As String

    OutBuffer = m_sOutBuffer

End Property

Public Property Get Password() As String

    Password = m_sPassword

End Property

Public Property Let Password(ByVal sNewValue As String)

    m_sPassword = sNewValue

End Property

