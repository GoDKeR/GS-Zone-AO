Attribute VB_Name = "modGeneral"
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

Public aPosH As Integer
Public aPosV As Integer
Public SobrePos As Integer
Public aSize As Integer
Public mMap As Integer
Public ActualFile As String
Public SelMapPos As Integer
Public Changed As Boolean

Public Sub SetWorldSize(ByVal iSize As Integer, Optional Resize As Boolean = False)
On Error Resume Next

    If iSize <= 1 Then Exit Sub
    
    Dim iH As Integer
    Dim iV As Integer
    Dim pM As Integer
    
    ' Acomodar Scrolls
    frmMain.HScroll.value = 0
    frmMain.HScroll.Max = iSize - 2
    frmMain.VScroll.value = 0
    frmMain.VScroll.Max = iSize - 2
    aPosH = 0
    aPosV = 0
    
    ' Limpiamos previos...
    If mMap > 1 Then
        For pM = 1 To MaxGrid
            Unload frmMain.pMap(pM)
        Next
    End If
    
    ' Organizar pMap's
    mMap = 1
    pM = 1
    For iV = 1 To iSize
        For iH = 1 To iSize
            Load frmMain.pMap(pM)
            If cVerGrid = False Then frmMain.pMap(pM).BorderStyle = 0
            frmMain.pMap(pM).BackColor = vbBlack
            frmMain.pMap(pM).Top = (((iV - 1) * frmMain.pMap(0).Height))
            frmMain.pMap(pM).Left = (((iH - 1) * frmMain.pMap(0).Width))
            frmMain.pMap(pM).Visible = True
            pM = pM + 1
        Next
        DoEvents
    Next
    
    mMap = pM - 1
    aSize = iSize
    
    MaxGrid = mMap
    If Resize = False Then
        ReDim WorldGrid(1 To MaxGrid)
    Else
        ' Datos originales
        Dim OriginalMaxGrid As Integer
        Dim OriginalGrid() As Integer
        Dim OriginalSize As Integer
        OriginalMaxGrid = UBound(WorldGrid)
        OriginalGrid = WorldGrid
        OriginalSize = Sqr(OriginalMaxGrid)
        
        ' Nuevo tamaño
        ReDim WorldGrid(1 To MaxGrid)
        
        ' Reubicación de datos
        MsgBox "OriginalMaxGrid: " & OriginalMaxGrid & vbCrLf & _
                   " - OriginalSize: " & OriginalSize
        
        
        Dim tN As Integer
        Dim tV As Integer
        Dim tC As Integer
        tC = 0
        tV = 1
        If OriginalSize > iSize Then ' habrá perdida de datos
            ' falta programar :P
        ElseIf OriginalSize <= iSize Then ' no habrá perdida de datos
            For tV = 1 To OriginalMaxGrid
                WorldGrid((tC * iSize) + (tV - (tC * OriginalSize))) = OriginalGrid(tV)
                If tV > ((tC + 1) * OriginalSize) Then
                    tC = tC + 1
                End If
            Next
        End If
        
        Dim iGrid As Integer
        Dim tExp As String
        If cUsarPNG = True Then
            tExp = ".png"
        Else
            tExp = ".bmp"
        End If
        
        For iGrid = 1 To MaxGrid
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
    End If
    
    frmMain.pBase.Refresh
    DoEvents
    
End Sub
