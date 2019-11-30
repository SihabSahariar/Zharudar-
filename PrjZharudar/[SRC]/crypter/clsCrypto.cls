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
Option Explicit
Private Const ALG_CLASS_ANY = 0
Private Const ALG_CLASS_SIGNATURE = 8192
Private Const ALG_CLASS_MSG_ENCRYPT = 16384
Private Const ALG_CLASS_DATA_ENCRYPT = 24576
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_CLASS_KEY_EXCHANGE = 40960
Private Const ALG_TYPE_ANY = 0
Private Const ALG_TYPE_DSS = 512
Private Const ALG_TYPE_RSA = 1024
Private Const ALG_TYPE_BLOCK = 1536
Private Const ALG_TYPE_STREAM = 2048
Private Const ALG_TYPE_DH = 2560
Private Const ALG_TYPE_SECURECHANNEL = 3072
Private Const ALG_SID_RC2 = 2
Private Const ALG_SID_RC4 = 1
Private Const ALG_SID_SEAL = 2
Private Const ALG_SID_DH_SANDF = 1
Private Const ALG_SID_DH_EPHEM = 2
Private Const ALG_SID_AGREED_KEY_ANY = 3
Private Const ALG_SID_KEA = 4
Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA = 4
Private Const ALG_SID_SHA1 = 4
Private Const ALG_SID_MAC = 5
Private Const ALG_SID_RIPEMD = 6
Private Const ALG_SID_RIPEMD160 = 7
Private Const ALG_SID_SSL3SHAMD5 = 8
Private Const ALG_SID_HMAC = 9
Private Const CALG_MD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_RC2 = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK) Or ALG_SID_RC2)
Private Const CALG_SSL3_SHAMD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SSL3SHAMD5)
Private Const CRYPT_VERIFYCONTEXT = &HF0000000
Private Const CRYPT_NEWKEYSET = &H8
Private Const CRYPT_DELETEKEYSET = &H10
Private Const CRYPT_MACHINE_KEYSET = &H20
Private Const HP_ALGID = &H1
Private Const HP_HASHVAL = &H2
Private Const HP_HASHSIZE = &H4
Private Const HP_HMAC_INFO = &H5
Private Const CRYPT_EXPORTABLE = &H1
Private Const CRYPT_USER_PROTECTED = &H2
Private Const CRYPT_CREATE_SALT = &H4
Private Const CRYPT_UPDATE_KEY = &H8
Private Const CRYPT_NO_SALT = &H10
Private Const CRYPT_PREGEN = &H40
Private Const CRYPT_RECIPIENT = &H10
Private Const CRYPT_INITIATOR = &H40
Private Const CRYPT_ONLINE = &H80
Private Const CRYPT_SF = &H100
Private Const CRYPT_CREATE_IV = &H200
Private Const CRYPT_KEK = &H400
Private Const CRYPT_DATA_KEY = &H800
Private Const CRYPT_SERVER = &H400
Private Const KEY_LENGTH_MASK = &HFFFF0000
Private Const CRYPT_MACHINE_DEFAULT = &H1
Private Const CRYPT_USER_DEFAULT = &H2
Private Const CRYPT_DELETE_DEFAULT = &H4
Private Const CRYPT_FAILED = 0
Private Const CRYPT_SUCCEED = -1
Private Const PROV_RSA_FULL = 1
Private Const PROV_RSA_SIG = 2
Private Const PROV_DSS = 3
Private Const PROV_FORTEZZA = 4
Private Const PROV_MS_EXCHANGE = 5
Private Const PROV_SSL = 6
Private Const PROV_RSA_SCHANNEL = 12
Private Const PROV_DSS_DH = 13
Private Const PROV_EC_ECDSA_SIG = 14
Private Const PROV_EC_ECNRA_SIG = 15
Private Const PROV_EC_ECDSA_FULL = 16
Private Const PROV_EC_ECNRA_FULL = 17
Private Const PROV_SPYRUS_LYNKS = 20
Private Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Private Const MS_ENHANCED_PROV = "Microsoft Enhanced Cryptographic Provider v1.0"
Private Const MS_DEF_RSA_SIG_PROV = "Microsoft RSA Signature Cryptographic Provider"
Private Const MS_DEF_RSA_SCHANNEL_PROV = "Microsoft Base RSA SChannel Cryptographic Provider"
Private Const MS_ENHANCED_RSA_SCHANNEL_PROV = "Microsoft Enhanced RSA SChannel Cryptographic Provider"
Private Const MS_DEF_DSS_PROV = "Microsoft Base DSS Cryptographic Provider"
Private Const MS_DEF_DSS_DH_PROV = "Microsoft Base DSS and Diffie-Hellman Cryptographic Provider"
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    phProv As Long, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hBaseData As Long, _
    ByVal dwFlags As Long, _
    phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" ( _
    ByVal hKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    ByVal pbData As String, _
    pdwDataLen As Long, _
    ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    ByVal pbData As String, _
    pdwDataLen As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" ( _
    ByVal hHash As Long, _
    ByVal pbData As String, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" ( _
    ByVal hHash As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" ( _
    ByVal hHash As Long, ByVal dwParam As Long, _
    ByVal pbData As String, pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashDWParam Lib "advapi32.dll" _
    Alias "CryptGetHashParam" (ByVal hHash As Long, _
    ByVal dwParam As Long, pbData As Long, _
    pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Const ERROR_INVALID_HANDLE As Long = 6
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const NTE_BAD_KEY As Long = &H80090003
Private Const NTE_BAD_UID As Long = &H80090001
Private Const NTE_NO_KEY As Long = &H8009000D
Private Const NTE_BAD_SIGNATURE As Long = &H80090006
Private m_sPassword As String
Private m_sInBuffer As String
Private m_sOutBuffer As String
Private m_sErrorMsg As String
Private m_lHCryptProv As Long
Private m_lHSessionKey As Long
Private m_sHashBuffer As String
Private m_iHashLength As Integer
Public Property Get InBuffer() As String
    InBuffer = m_sInBuffer
End Property
Public Property Let InBuffer(vNewValue As String)
    m_sInBuffer = vNewValue
End Property
Public Property Get OutBuffer() As String
    OutBuffer = m_sOutBuffer
End Property
Public Property Get ErrorMsg() As String
    ErrorMsg = m_sErrorMsg
End Property
Private Sub Class_Initialize()
    If Not InitUser Then
        m_sErrorMsg = "Unable to initialize CryptoAPI."
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
    End If
End Sub
Private Function InitUser() As Boolean
    On Error Resume Next
    InitUser = False
    
    Dim lDataSize As Long
    
    Dim lResult As Long
    Dim sResult As String
    Dim sContainer As String
    Dim sProvider As String
    sContainer = vbNullChar
    sProvider = MS_DEF_PROV & vbNullChar
        
    If Not CBool(CryptAcquireContext(m_lHCryptProv, sContainer, sProvider, PROV_RSA_FULL, 0)) Then
    
        If Not CBool(CryptAcquireContext(m_lHCryptProv, sContainer, sProvider, PROV_RSA_FULL, CRYPT_NEWKEYSET)) Then
            m_sErrorMsg = "Error creating key container - " & CStr(Err.LastDllError)
            MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
            
            Exit Function
        End If
        
    End If

    InitUser = True

    Exit Function
    
End Function

Private Sub Class_Terminate()
    Dim lResult As Long
    If (m_lHCryptProv <> 0) Then lResult = CryptReleaseContext(m_lHCryptProv, 0)
End Sub

Public Function GeneratePasswordKey() As Boolean
    On Error Resume Next
    Dim lHash As Long
    Dim lResult As Long
    
    GeneratePasswordKey = False
    

    If Not CBool(CryptCreateHash(m_lHCryptProv, CALG_MD5, 0, 0, lHash)) Then
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptCreateHash!"
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
        Exit Function
    End If
    

    If Not CBool(CryptHashData(lHash, m_sPassword, Len(m_sPassword), 0)) Then
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptHashData!"
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
        Exit Function
    End If
    

    If Not CBool(CryptDeriveKey(m_lHCryptProv, CALG_RC2, lHash, 0, m_lHSessionKey)) Then
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptDeriveKey!"
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
        Exit Function
    End If
    

    If Not CBool(CryptDestroyHash(lHash)) Then
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptDestroyHash!"
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
        
        Exit Function
    End If
    
    GeneratePasswordKey = True
End Function

Public Property Get Password() As String
    Password = m_sPassword
End Property

Public Property Let Password(ByVal sNewValue As String)
    m_sPassword = sNewValue
End Property

Public Function EncryptMessageData()
On Error Resume Next
    EncryptMessageData = False

    Dim lDataSize As Long
    Dim lResult As Long
    Dim sCryptBuffer As String
    Dim lCryptLength As Long
    Dim lCryptBufLen As Long

    lCryptLength = Len(m_sInBuffer)
    lResult = CryptEncrypt(m_lHSessionKey, 0, 1, 0, vbNullString, lCryptLength, lCryptBufLen)
    
    lCryptBufLen = lCryptLength * 2
    lCryptLength = Len(m_sInBuffer)
    sCryptBuffer = String(lCryptBufLen, vbNullChar)
    
    LSet sCryptBuffer = m_sInBuffer

    If Not CBool(CryptEncrypt(m_lHSessionKey, 0, 1, 0, sCryptBuffer, lCryptLength, lCryptBufLen)) Then
        
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptEncrypt!"
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
        
        Exit Function
    End If

    m_sOutBuffer = Mid$(sCryptBuffer, 1, lCryptLength)

    EncryptMessageData = True
End Function

Public Function DestroySessionKey() As Boolean
    Dim lResult As Long
    
    DestroySessionKey = False
    
    If Not CBool(CryptDestroyKey(m_lHSessionKey)) Then
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptDestroyKey!"
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
        
        Exit Function
    End If
    
    
    DestroySessionKey = True
End Function

Public Function DecryptMessageData() As Boolean
On Error Resume Next
    DecryptMessageData = False

    Dim lDataSize As Long
    Dim lResult As Long
    Dim sCryptBuffer As String
    Dim lCryptLength As Long
    Dim lCryptBufLen As Long

    lCryptBufLen = Len(m_sInBuffer)
    sCryptBuffer = String(lCryptBufLen, vbNullChar)
    LSet sCryptBuffer = m_sInBuffer
    
    If Not CBool(CryptDecrypt(m_lHSessionKey, 0, 1, 0, sCryptBuffer, lCryptBufLen)) Then
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptDecrypt!"
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
        
        Exit Function
    End If
    
    m_sOutBuffer = Mid$(sCryptBuffer, 1, lCryptBufLen)
    
    DecryptMessageData = True
End Function

Public Function HashFile() As Boolean
    On Error Resume Next
    Dim lHash As Long
    Dim lResult As Long
    Dim sCryptBuffer As String
    Dim lCryptBufLen As Long
    Dim lCryptHashSize As Long

    HashFile = False
    
    If Not CBool(CryptCreateHash(m_lHCryptProv, CALG_MD5, 0, 0, lHash)) Then
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptCreateHash!"
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
        Exit Function
    End If
    
    If Not CBool(CryptHashData(lHash, m_sInBuffer, Len(m_sInBuffer), 0)) Then
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptHashData!"
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
        Exit Function
    End If
    
    lCryptBufLen = 4
    If Not CBool(CryptGetHashDWParam(lHash, HP_HASHSIZE, lCryptHashSize, lCryptBufLen, 0)) Then
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptGetHashParam!"
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
        Exit Function
    End If
    
    sCryptBuffer = String(lCryptHashSize, vbNullChar)
    lCryptBufLen = lCryptHashSize

    If Not CBool(CryptGetHashParam(lHash, HP_HASHVAL, sCryptBuffer, lCryptBufLen, 0)) Then
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptGetHashParam!"
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
        Exit Function
    End If
    
    If Not CBool(CryptDestroyHash(lHash)) Then
        m_sErrorMsg = "Error " & CStr(Err.LastDllError) & " during CryptDestroyHash!"
        MsgBox m_sErrorMsg, vbOKOnly, "VB Crypto"
        Exit Function
    End If
    
    m_sHashBuffer = Mid$(sCryptBuffer, 1, lCryptBufLen)
    m_iHashLength = lCryptBufLen
    
    HashFile = True
End Function

Public Function EncryptFileData() As Boolean
    EncryptFileData = False
    Dim strInputData As String
    strInputData = Trim(Str(m_iHashLength)) + m_sHashBuffer + m_sInBuffer
    m_sInBuffer = strInputData
    EncryptFileData = EncryptMessageData
End Function

Public Function DecryptFileData() As Boolean
    DecryptFileData = False
    Dim strInputData     As String
    Dim strHash          As String
    Dim strHashLength    As String
    Dim lHashLength      As Long
    
    If Not DecryptMessageData Then Exit Function
    strHashLength = Left(m_sOutBuffer, 2)
    lHashLength = CLng(strHashLength)
    strHash = Mid(m_sOutBuffer, 3, lHashLength)
    strInputData = Mid(m_sOutBuffer, (3 + lHashLength))
    m_sInBuffer = strInputData
    If Not HashFile Then Exit Function
    If (strHash = m_sHashBuffer) Then
        m_sOutBuffer = strInputData
    Else
        MsgBox "File has been corrupted since original encryption!", vbInformation, "File Error"
        m_sOutBuffer = ""
        Exit Function
    End If
    DecryptFileData = True
End Function
