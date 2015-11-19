Attribute VB_Name = "modArchivo"
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

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const OFN_EXPLORER = &H80000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const cSingleSelFlags As Long = OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Public Function ExplorarArchivo(ByVal hWnd As Long, ByVal Filtro As String, ByVal Titulo As String, Optional ByVal bSave As Boolean = False) As String
On Error Resume Next
    Dim OFName As OPENFILENAME
    Dim ST As String
    
    ExplorarArchivo = vbNullString

    With OFName
        .lStructSize = Len(OFName)
        .hWndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrFilter = Filtro & Chr$(0) & "Todos los archivos" & Chr$(0) & "*.*"
        .lpstrTitle = Titulo
        .Flags = cSingleSelFlags
        .lpstrFile = Space$(1023)
        .nMaxFile = 1024
    End With

    If bSave = True Then
        If Not GetSaveFileName(OFName) = 0 Then
            ST = Split(Trim$(OFName.lpstrFile), Chr(0))(0)
            ExplorarArchivo = ST
            Exit Function
        End If
    Else
        If Not GetOpenFileName(OFName) = 0 Then
            ST = Split(Trim$(OFName.lpstrFile), Chr(0))(0)
            ExplorarArchivo = ST
            Exit Function
        End If
    End If
    
    ExplorarArchivo = vbNullString
    Exit Function

End Function
