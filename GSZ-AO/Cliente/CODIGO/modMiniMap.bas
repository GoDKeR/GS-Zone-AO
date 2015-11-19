Attribute VB_Name = "modMinimap"
'***************************************GoDKeR*****************************************
Option Explicit

Public Const fMiniMapInit = "MiniMap.init"

Public MiniMapEnabled       As Boolean
Public MiniMapVisible       As Boolean
Public MiniMapDrag          As Boolean
Public DefaultAlphaMiniMap  As Byte

Private AlphaMiniMap        As Byte
Private MMC_Char            As Long ' yo :)

Private DirectD3D As D3DX8

'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type tMinMap
    color(0 To 3) As Long
    X As Integer
    Y As Integer
    Texture As Direct3DTexture8 'Holds the texture of the text
    TextureSize As POINTAPI     'Size of the texture
End Type

Public Minimap As tMinMap ' _Default2 As CustomFont

Public Function LeerMinimapInit()
'**************************************************************
'Author: ^[GS]^
'Last Modify Date: 10/08/2014 - ^[GS]^
'**************************************************************
On Local Error Resume Next

    If FileExist(sPathINIT & fMiniMapInit, vbArchive) Then
        Dim N As Integer
        N = FreeFile
        Open sPathINIT & fMiniMapInit For Binary As #N
            Get #N, , MiCabecera
            Get #N, , Minimap.X
            Get #N, , Minimap.Y
            Get #N, , DefaultAlphaMiniMap
            Get #N, , MiniMapEnabled
        Close #N
    Else ' Default
        Minimap.X = 442
        Minimap.Y = 2
        DefaultAlphaMiniMap = 205
        MiniMapEnabled = True
    End If
    
    ' MiniMap Enabled?
    frmMain.chkMiniMap.Checked = MiniMapEnabled
    
End Function

Public Sub EscribirMinimapInit()
'**************************************************************
'Author: ^[GS]^
'Last Modify Date: 07/08/2013 - ^[GS]^
'**************************************************************
On Local Error Resume Next

    Dim N As Integer
    N = FreeFile
    Open sPathINIT & fMiniMapInit For Binary As #N
        Put #N, , MiCabecera
        Put #N, , Minimap.X
        Put #N, , Minimap.Y
        Put #N, , DefaultAlphaMiniMap
        Put #N, , MiniMapEnabled
    Close #N
    
End Sub

Public Sub MiniMap_Init()
'**************************************************************
'Author: GoDKeR
'Last Modify Date: 06/08/2013 - GoDKeR
'**************************************************************
    MMC_Char = D3DColorARGB(150, 255, 0, 0)
    AlphaMiniMap = DefaultAlphaMiniMap
    Call LeerMinimapInit
End Sub

Public Sub MiniMap_Render(ByVal X As Long, ByVal Y As Long)
'**************************************************************
'Author: GoDKeR
'Last Modify Date: 06/08/2013 - GoDKeR
'**************************************************************
    If MiniMapEnabled = False Then Exit Sub
    
    Dim VertexArray(0 To 3) As TLVERTEX
    Dim SrcWidth            As Integer
    Dim Width               As Integer
    Dim SrcHeight           As Integer
    Dim Height              As Integer
    Dim SrcBitmapWidth      As Long
    Dim SrcBitmapHeight     As Long
    Dim SRDesc              As D3DSURFACE_DESC
    
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
    
    'Set the temp color (or else the first character has no color)
    'TempColor = color
    
    With Minimap
    
        If Not .Texture Is Nothing Then
            .Texture.GetLevelDesc 0, SRDesc
                    
            SrcWidth = 100 'd3dtextures.texwidth
            Width = 100 'd3dtextures.texwidth
            Height = 100 'd3dtextures.texheight
            SrcHeight = 100 'd3dtextures.texheight
                    
            SrcBitmapWidth = SRDesc.Width
            SrcBitmapHeight = SRDesc.Height
                    
            'Set the RHWs (must always be 1)
            VertexArray(0).rhw = 1
            VertexArray(1).rhw = 1
            VertexArray(2).rhw = 1
            VertexArray(3).rhw = 1
                    
            'Find the left side of the rectangle
            VertexArray(0).X = X
            VertexArray(0).tu = (.TextureSize.X / SrcBitmapWidth)
                    
            'Find the top side of the rectangle
            VertexArray(0).Y = Y
            VertexArray(0).tv = (.TextureSize.Y / SrcBitmapHeight)
                    
            'Find the right side of the rectangle
            VertexArray(1).X = X + Width
            VertexArray(1).tu = (.TextureSize.X + SrcWidth) / SrcBitmapWidth
                    
            'These values will only equal each other when not a shadow
            VertexArray(2).X = VertexArray(0).X
            VertexArray(3).X = VertexArray(1).X
                    
            'Find the bottom of the rectangle
            VertexArray(2).Y = Y + Height
            VertexArray(2).tv = (.TextureSize.Y + SrcHeight) / SrcBitmapHeight
                    
            'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
            VertexArray(1).Y = VertexArray(0).Y
            VertexArray(1).tv = VertexArray(0).tv
            VertexArray(2).tu = VertexArray(0).tu
            VertexArray(3).Y = VertexArray(2).Y
            VertexArray(3).tu = VertexArray(1).tu
            VertexArray(3).tv = VertexArray(2).tv
            VertexArray(0).color = .color(0)
            VertexArray(1).color = .color(1)
            VertexArray(2).color = .color(2)
            VertexArray(3).color = .color(3)
                    
            'Set the texture
            DirectDevice.SetTexture 0, .Texture
            DirectDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(AlphaMiniMap, 0, 0, 0)
            
            
            'faster
            DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), LenB(VertexArray(0))
        
            'Correccion del minimap, para que borre el puntito del usuario, se cambio el color del punto ahora es rojo :$
            If AlphaMiniMap > 50 Then
                Call Draw_FillBox(Minimap.X + UserPos.X, Minimap.Y + UserPos.Y, 3, 3, MMC_Char, MMC_Char, False)
            End If
        
            Call MiniMap_ColorSet
        End If
    
    End With

End Sub

Private Sub MiniMap_ColorSet()
'**************************************************************
'Author: GoDKeR
'Last Modify Date: 06/08/2013 - ^[GS]^
'**************************************************************

    On Error GoTo Err
    
    Dim Colorsitus As D3DCOLORVALUE
    
    If frmMain.MouseX > Minimap.X And frmMain.MouseY > Minimap.Y And frmMain.MouseX < Minimap.X + 100 And frmMain.MouseY < Minimap.Y + 100 Then
        If AlphaMiniMap > 0 Then
                AlphaMiniMap = AlphaMiniMap - timerTicksPerFrame * 25
                If AlphaMiniMap < 10 Then AlphaMiniMap = 0
        End If
    Else
        If AlphaMiniMap < DefaultAlphaMiniMap Then
                AlphaMiniMap = AlphaMiniMap + timerTicksPerFrame * 25
                If AlphaMiniMap > DefaultAlphaMiniMap Then AlphaMiniMap = DefaultAlphaMiniMap
        End If
    End If
    
    Colorsitus.r = 255
    Colorsitus.g = 255
    Colorsitus.b = 255
    Colorsitus.a = AlphaMiniMap
    D3DColorToRgbList Minimap.color(), Colorsitus
    
    Exit Sub
    
Err:
    Colorsitus.a = DefaultAlphaMiniMap
    Colorsitus.r = 255
    Colorsitus.g = 255
    Colorsitus.b = 255
    D3DColorToRgbList Minimap.color(), Colorsitus
        
End Sub

Public Sub MiniMap_ChangeTex(iMap As Integer)
'**************************************************************
'Author: GoDKeR
'Last Modify Date: 07/08/2013 - ^[GS]^
'**************************************************************
    On Error GoTo eDebug:

    Dim mapInfo As D3DXIMAGE_INFO_A

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
    
    'Set the texture
    Dim data() As Byte
    If Get_File_Data(DirMapas, "MAPA" & CStr(iMap) & ".BMP", data, 1) = True Then
        Set Minimap.Texture = DirectD3D8.CreateTextureFromFileInMemoryEx(DirectDevice, data(0), UBound(data) + 1, _
                D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
        MiniMapVisible = True
    Else
        Set Minimap.Texture = Nothing
        MiniMapVisible = False
        Call LogError("MiniMap_ChangeTex::El mapa " & iMap & " no tiene vista de MiniMap.")
        Exit Sub
    End If
                
    'Store the size of the texture
    Minimap.TextureSize.X = mapInfo.Width
    Minimap.TextureSize.Y = mapInfo.Height

    Exit Sub

eDebug:

    Call LogError("MiniMap_ChangeTex::Error " & Err.Number & " - " & Err.Description & " - [MAPA: " & iMap & "]")
    If Err.Number = "-2005529767" Then
        Call MsgBox("Error en la textura utilizada del Minimap", vbCritical)
    End If
End Sub


