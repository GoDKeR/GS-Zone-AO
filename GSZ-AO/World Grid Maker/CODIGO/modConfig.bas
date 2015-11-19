Attribute VB_Name = "modConfig"
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

Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

Public cVerNumMap As Boolean
Public cVerGrid As Boolean
Public cDirMaps As String
Public cDefaultSize As Integer
Public cUsarPNG As Boolean

Public Sub ReadConfig()
On Error Resume Next
    
    Dim cFile As String
    cFile = App.Path & "\" & App.EXEName & ".ini"

    cDirMaps = GetVar(cFile, "INIT", "DirMaps")
    If LenB(cDirMaps) = 0 Then cDirMaps = App.Path
    cDefaultSize = Val(GetVar(cFile, "INIT", "DefaultSize"))
    If cDefaultSize = 0 Then cDefaultSize = 10
    cVerGrid = IIf(Val(GetVar(cFile, "INIT", "VerGrid")) = 1, True, False)
    cVerNumMap = IIf(Val(GetVar(cFile, "INIT", "VerNumMap")) = 1, True, False)
    cUsarPNG = IIf(Val(GetVar(cFile, "INIT", "UsarPNG")) = 1, True, False)
    
    If FileExist(cDirMaps, vbDirectory) = False Then
        MsgBox "La Carpeta de Mapas es invalida." & vbCrLf & "Por favor, identifique una carpeta valida antes de abrir cualquier WorldGrid, de lo contrario no verá las imagenes." & vbCrLf & vbCrLf & "Puede configurarlo en la ventana de selección de mapa o manualmente editando " & App.EXEName & ".ini, en el campo DirMaps=.", vbInformation + vbOKOnly
    End If
    
End Sub

Public Sub SaveConfig()
On Error Resume Next
    
    Dim cFile As String
    cFile = App.Path & "\" & App.EXEName & ".ini"
    
    WriteVar cFile, "INIT", "DirMaps", cDirMaps
    WriteVar cFile, "INIT", "DefaultSize", Val(cDefaultSize)
    WriteVar cFile, "INIT", "VerGrid", IIf(cVerGrid = True, "1", "0")
    WriteVar cFile, "INIT", "VerNumMap", IIf(cVerNumMap = True, "1", "0")

End Sub



Public Function GetVar(ByRef file As String, ByRef Main As String, ByRef Var As String) As String
On Error Resume Next

    Dim sSpaces As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
    
    szReturn = vbNullString
    sSpaces = Space$(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
    GetPrivateProfileString Main, Var, szReturn, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
End Function

Public Sub WriteVar(ByRef file As String, ByRef Main As String, ByRef Var As String, ByRef value As String)
On Error Resume Next

    writeprivateprofilestring Main, Var, value, file
    
End Sub
