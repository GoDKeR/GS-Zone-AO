Attribute VB_Name = "modDirectorio"
'World Grid Maker
'Copyright (C) 2012 GS-Zone
'
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
'You can contact me at:
'info@gs-zone.org

Option Explicit

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Const WM_USER = &H400
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BFFM_INITIALIZED = &H1
Private Const BFFM_SETSELECTIONA = (WM_USER + 102)
Private Const MAX_PATH = 260

Private Declare Function SHGetIDListFromPath Lib "Shell32" Alias "#162" (ByVal pszPath As String) As Long
Private Declare Function SHGetFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ppidl As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long



Public Function ExplorarDirectorio(ByVal hWnd As Long, ByVal Mensaje As String, ByVal PathInit As String) As String
On Error Resume Next

    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    
    ExplorarDirectorio = vbNullString
    
    lpIDList = SHGetFolderLocation(hWnd, 6, SHGetIDListFromPath(PathInit), 0, tBrowseInfo.pIDLRoot)

    With tBrowseInfo
        .hWndOwner = hWnd
        .lpfnCallback = adr(AddressOf BrowseCallbackProc)
        .lpszTitle = lstrcat(Mensaje, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_NEWDIALOGSTYLE
        .lParam = SHGetIDListFromPath(StrConv(PathInit, vbUnicode))
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        ExplorarDirectorio = sBuffer & "\"
    Else
        ExplorarDirectorio = vbNullString
        Exit Function
    End If
    
End Function

Private Function adr(n As Long) As Long
    adr = n
End Function
 
Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
  If uMsg = BFFM_INITIALIZED Then
      Call SendMessage(hWnd, BFFM_SETSELECTIONA, False, ByVal lpData)
  End If
End Function

Public Function ParsePath(strFullPathName As String, ReturnType As Byte) As String
    Dim strTemp As String, intX As Integer, strPathName As String, strFileName As String
    If Len(strFullPathName) > 0 Then
        strTemp = ""
        intX = Len(strFullPathName)
        Do While strTemp <> "\"
            strTemp = Mid(strFullPathName, intX, 1)
            If strTemp = "\" Then
                strPathName = Left(strFullPathName, intX)
                strFileName = Right(strFullPathName, Len(strFullPathName) - intX)
            End If
            intX = intX - 1
        Loop
        Select Case ReturnType
        Case vbDirectory
            ParsePath = strPathName
        Case vbArchive
            ParsePath = strFileName
        Case Else
            ParsePath = strFullPathName
        End Select
    Else
        ParsePath = ""
    End If
End Function
