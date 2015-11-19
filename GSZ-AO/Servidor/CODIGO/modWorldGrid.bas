Attribute VB_Name = "modWorldGrid"
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
'On Error Resume Next

    UpdateGrid = False
    
    If MaxGrid <= 0 Or iMap <= 0 Or iMap > NumMaps Then Exit Function

    Dim iGrid As Integer

    For iGrid = 1 To MaxGrid
        If WorldGrid(iGrid) = iMap Then
            Exit For
        End If
    Next
        
    If (iGrid < (MaxGrid + 1)) Then
    
        Dim GridSize As Integer
        GridSize = Sqr(MaxGrid)
    
        If (iGrid > GridSize) Then
            MapInfo(iMap).Grid(eHeading.NORTH) = WorldGrid(iGrid - GridSize)
        Else
            MapInfo(iMap).Grid(eHeading.NORTH) = 0
        End If

        If (iGrid + GridSize) <= MaxGrid Then
            MapInfo(iMap).Grid(eHeading.SOUTH) = WorldGrid(iGrid + GridSize)
        Else
            MapInfo(iMap).Grid(eHeading.SOUTH) = 0
        End If

        If (iGrid + 1) <= MaxGrid Then
            MapInfo(iMap).Grid(eHeading.EAST) = WorldGrid(iGrid + 1)
        Else
            MapInfo(iMap).Grid(eHeading.EAST) = 0
        End If
        
        If (iGrid - 1) >= 1 Then
            MapInfo(iMap).Grid(eHeading.WEST) = WorldGrid(iGrid - 1)
        Else
            MapInfo(iMap).Grid(eHeading.WEST) = 0
        End If

    End If
    
    UpdateGrid = True

End Function

Public Function LoadWorldGrid(ByVal sFile As String) As Boolean
On Error Resume Next

    LoadWorldGrid = False
    
    If FileExist(sFile, vbArchive) = False Then Exit Function
    
    Dim iGrid As Integer
    Dim handle As Integer
    handle = FreeFile
    
    Open sFile For Binary Access Read As handle
        Seek handle, 1
        Get handle, , MaxGrid
        
        ReDim WorldGrid(1 To MaxGrid) As Integer

        For iGrid = 1 To MaxGrid
            Get handle, , WorldGrid(iGrid)
        Next
    Close handle
    
    LoadWorldGrid = True

End Function
