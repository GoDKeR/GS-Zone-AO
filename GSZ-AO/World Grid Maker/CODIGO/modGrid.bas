Attribute VB_Name = "modGrid"
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

Public MaxGrid As Integer
Public WorldGrid() As Integer

Public Function UpdateGrid(ByVal iMap As Integer) As Boolean
On Error Resume Next

    UpdateGrid = False
    
    If MaxGrid <= 0 Or iMap <= 0 Then Exit Function

    Dim iGridPos As Integer
    Dim iGrid As Integer
    
    iGridPos = 0
    For iGrid = 1 To MaxGrid
        If WorldGrid(iGrid) = iMap Then
            iGridPos = iGrid
            Exit For
        End If
    Next
        
    If iGridPos <> 0 Then
        Dim GridSize As Integer
        GridSize = Sqr(MaxGrid)
        
        Dim sInfo As String
    
        sInfo = "--- Mapa " & iMap & " ---" & vbCrLf & "Arriba: "
        If iGridPos > GridSize Then
            sInfo = sInfo & WorldGrid(iGridPos - GridSize)
        Else
            sInfo = sInfo & "0"
        End If
        sInfo = sInfo & vbCrLf & "Abajo: "
        If iGridPos + GridSize <= MaxGrid Then
            sInfo = sInfo & WorldGrid(iGridPos + GridSize)
        Else
            sInfo = sInfo & "0"
        End If
        sInfo = sInfo & vbCrLf & "Derecha: "
        If iGridPos + 1 <= MaxGrid Then
            sInfo = sInfo & WorldGrid(iGridPos + 1)
        Else
            sInfo = sInfo & "0"
        End If
        sInfo = sInfo & vbCrLf & "Izquierda: "
        If iGridPos - 1 >= 1 Then
            sInfo = sInfo & WorldGrid(iGridPos - 1)
        Else
            sInfo = sInfo & "0"
        End If
        
        MsgBox sInfo
    End If
    
End Function

Public Function LoadWorldGrid(ByVal sFile As String) As Boolean
On Error Resume Next

    LoadWorldGrid = False
    
    If FileExist(sFile, vbArchive) = False Then Exit Function
    
    Dim tExp As String
    Dim iGrid As Integer
    Dim Handle As Integer
    Handle = FreeFile
    
    If cUsarPNG = True Then
        tExp = ".png"
    Else
        tExp = ".bmp"
    End If
    
    Open sFile For Binary Access Read As Handle
        Seek Handle, 1
        Get Handle, , MaxGrid
        
        Call SetWorldSize(Sqr(MaxGrid))

        For iGrid = 1 To MaxGrid
            Get Handle, , WorldGrid(iGrid)
        
            If WorldGrid(iGrid) <> 0 Then ' lo dibujamos...
                If FileExist(cDirMaps & "\Mapa" & WorldGrid(iGrid) & tExp, vbArchive) Then
                    If cUsarPNG = True Then
                        Call PngPictureLoad(cDirMaps & "\Mapa" & WorldGrid(iGrid) & tExp, frmMain.pMap(iGrid), False)
                    Else
                        frmMain.pMap(iGrid).Picture = LoadPicture(cDirMaps & "\Mapa" & WorldGrid(iGrid) & tExp)
                    End If
                    frmMain.pMap(iGrid).Tag = vbNullString
                Else
                    frmMain.pMap(iGrid).BackColor = vbRed
                    frmMain.pMap(iGrid).Tag = cDirMaps & "\Mapa" & WorldGrid(iGrid) & tExp
                End If
            End If
        Next
    Close Handle
    
    LoadWorldGrid = True
    Changed = False

End Function

Public Function SaveWorldGrid(ByVal sFile As String) As Boolean
On Error Resume Next

    SaveWorldGrid = False

    If Right$(sFile, 5) <> ".grid" Then sFile = sFile & ".grid"
    If FileExist(sFile, vbArchive) Then Call Kill(sFile)

    Dim iGrid As Integer
    Dim Handle As Integer
    Handle = FreeFile
    
    Open sFile For Binary Access Write As Handle
        Seek Handle, 1
        Put Handle, , MaxGrid
        For iGrid = 1 To MaxGrid
            Put Handle, , WorldGrid(iGrid)
        Next
    Close Handle
    
    SaveWorldGrid = True
    Changed = False
    
End Function
