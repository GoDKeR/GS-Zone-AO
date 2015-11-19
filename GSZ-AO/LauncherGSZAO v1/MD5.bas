Attribute VB_Name = "MD5"
'**************************************************************
' MD5.mod - Computes MD5 hashes of files using Windows CryptoApi (Advapi)
'
' Developed by Marco (Marco Vanotti - marco@vanotti.com.ar)
'**************************************************************

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

Public Const PROV_RSA_FULL          As Long = 1
Public Const CRYPT_VERIFYCONTEXT    As Long = &HF0000000

Public Const HP_HASHVAL             As Long = 2
Public Const HP_HASHSIZE            As Long = 4

Public Const ALG_TYPE_ANY           As Long = 0
Public Const ALG_CLASS_HASH         As Long = 32768
Public Const ALG_SID_MD5            As Long = 3

Public Const CALG_MD5               As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5

Public Const BUFSIZE                As Long = 1024
Public Const MD5LEN                 As Long = 16

Private Const MS_DEFAULT_PROVIDER   As String = _
    "Microsoft Base Cryptographic Provider v1.0"
    
Private Const CHUNK                 As Long = 16384

Private Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextA" ( _
    ByRef phProv As Long, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long
    
Private Declare Function CryptCreateHash Lib "advapi32" ( _
    ByVal hProv As Long, _
    ByVal algid As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    ByRef phHash As Long) As Long
    
Private Declare Function CryptDestroyHash Lib "advapi32" ( _
    ByVal hHash As Long) As Long

Private Declare Function CryptGetHashParam Lib "advapi32" ( _
    ByVal hHash As Long, _
    ByVal dwParam As Long, _
    ByRef pbData As Any, _
    ByRef pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptHashData Lib "advapi32" ( _
    ByVal hHash As Long, _
    ByRef pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptReleaseContext Lib "advapi32" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long

    
Private dwStatus As Long
Private hProvider As Long
Private hHash As Long
   
Public Function MD5File(ByVal sFile As String) As String
    Dim rgbHash(MD5LEN) As Byte
    Dim bTmp() As Byte
    Dim cbHash As Long
    Dim lFileLen As Long
    Dim lRemainder As Long
    Dim lChunks As Long
    Dim nF As Integer
    Dim i As Long

    If Not (CryptAcquireContext(hProvider, vbNullString, MS_DEFAULT_PROVIDER, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0) Then
        dwStatus = GetLastError()
        MD5File = "Error en CryptAcquireContext, info: " & dwStatus
        Exit Function
    End If
    
    If Not (CryptCreateHash(hProvider, CALG_MD5, 0, 0, hHash) <> 0) Then
        dwStatus = GetLastError()
        MD5File = "Error en CryptCreateHash, info: " & dwStatus
        Call CryptReleaseContext(hProvider, 0)
        Exit Function
    End If

    If Not FileExists(sFile) Then
        MD5File = "FILE NOT FOUND"
        Call CryptReleaseContext(hProvider, 0)
        Call CryptDestroyHash(hHash)
        Exit Function
    End If

    nF = FreeFile()
    lFileLen = FileLen(sFile)
    lChunks = lFileLen \ CHUNK          'TODO: Test for overflow with big files
    lRemainder = lFileLen Mod CHUNK     'TODO: Test for overflow with big files
    Open sFile For Binary Access Read As #nF

    
        ReDim bTmp(CHUNK - 1) As Byte
        
        For i = 1 To lChunks
            Get #nF, , bTmp
            If Not (CryptHashData(hHash, bTmp(0), UBound(bTmp) + 1, 0&) <> 0) Then
                dwStatus = GetLastError()
                MD5File = "Error en CryptHashData, info: " & dwStatus
                Call CryptReleaseContext(hProvider, 0)
                Call CryptDestroyHash(hHash)
                Close #nF
                Exit Function
            End If
        Next i
        
        If (lRemainder > 0) Then
            ReDim bTmp(lRemainder - 1) As Byte
            Get #nF, , bTmp

            If Not (CryptHashData(hHash, bTmp(0), UBound(bTmp) + 1, 0&) <> 0) Then
                dwStatus = GetLastError()
                MD5File = "Error en CryptHashData, info: " & dwStatus
                Call CryptReleaseContext(hProvider, 0)
                Call CryptDestroyHash(hHash)
                Close #nF
                Exit Function
            End If
        End If

        
    Close #nF
    cbHash = MD5LEN
    
    If (CryptGetHashParam(hHash, HP_HASHVAL, rgbHash(0), cbHash, 0&)) Then
        For i = 0 To cbHash - 1
            MD5File = MD5File & Right$("0" & Hex$(rgbHash(i)), 2)
        Next i
    Else
        dwStatus = GetLastError()
        MD5File = "Falló CryptGetHashParam, info: " & dwStatus
    End If
    
    
    Call CryptDestroyHash(hHash)
    Call CryptReleaseContext(hProvider, 0)
End Function




