Attribute VB_Name = "modTileEngine"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Option Explicit

Public vAngle As Single

Public rgbNormal(3) As Long

' // Partículas vbGore

Public WeatherEffectIndex As Integer

Public Const ScreenWidth As Long = 544 'Keep this identical to the value on the server!
Public Const ScreenHeight As Long = 416 'Keep this identical to the value on the server!

Public OffsetCounterX As Single
Public OffsetCounterY As Single

'Screen positioning
Public minY As Integer          'Start Y pos on current screen + tilebuffer
Public maxY As Integer          'End Y pos on current screen
Public minX As Integer          'Start X pos on current screen
Public maxX As Integer          'End X pos on current screen

Public screenminY  As Integer  'Start Y pos on current screen
Public screenmaxY  As Integer  'End Y pos on current screen
Public screenminX  As Integer  'Start X pos on current screen
Public screenmaxX  As Integer  'End X pos on current screen

Public ParticleTexture(1 To 20) As Direct3DTexture8

Public Const PI As Single = 3.14159265358979
 
Public ParticleOffsetX As Long
Public ParticleOffsetY As Long
Public LastOffsetX As Integer
Public LastOffsetY As Integer

' // Partículas vbGore

Public LightRGB_Default(3) As Long
Public LightRGB_Default_Alpha(3) As Long

Public movSpeed As Single

'Private ParticleTimer As Single

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)

'Fonts (Extraido de VBGore)
'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type
Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type
Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type
Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
End Type

Private Const Font_Default_TextureNum As Long = -1   'The texture number used to represent this font - only used for AlternateRendering - keep negative to prevent interfering with game textures
Private cfonts(1 To 2) As CustomFont ' _Default2 As CustomFont

'Standelf
'Static Textures
Private TEXTURERAIN As Direct3DTexture8
Private TEXTUREFOCUS As Direct3DTexture8
Private TEXTUREFOCUSBACK As Direct3DTexture8 'Godker
Private TEXTUREDEAD As Direct3DTexture8
Private FULLRECT As RECT
Private DEADCOLOR(0 To 3) As Long
Public ABD As Integer

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Posicion en un mapa
Public Type Position
    X As Integer
    Y As Integer
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer ' ¿Integer o Long?
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type


'Apariencia del personaje
Public Type Char
    active As Byte
    Heading As E_Heading
    Pos As Position
    
    ' maTih   /   Indice de la particula de la meditación.
    ParticleIndex   As Integer
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FXIndex As Long ' GSZAO
        
    Criminal As Byte
    Atacable As Boolean
    Newbie As Byte ' GSZAO
    bType As Byte ' GSZAO
    
    nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    Invisible As Boolean
    priv As Byte
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    particle_group As Integer
    light_value(3) As Long
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    RenderValue As RVList ' GSZAO
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type mapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    Pk As Boolean ' GSZAO
End Type

'DX8 Objects
Public DirectX As DirectX8
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8
Public DirectDevice As Direct3DDevice8

' Directx8 Fonts
Private Type FontInfo
    MainFont As DxVBLibA.D3DXFont
    MainFontDesc As IFont
    MainFontFormat As New StdFont
    color As Long
End Type: Private Font() As FontInfo

'Public Type TLVERTEX
'    x As Single
'    y As Single
'    Z As Single
'    rhw As Single
'    Color As Long
'    Specular As Long
'    tu As Single
'    tv As Single
'End Type

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer
Private MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime As Single
Public timerTicksPerFrame As Single
Dim engineBaseSpeed As Single


Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer


Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte

Private Enum PARTICLE_STATUS
    Alive = 0
    Dead = 1
End Enum

Private Type Particle
    X As Single     'World Space Coordinates
    Y As Single
    Z As Single
    vX As Single    'Speed and Direction
    vY As Single
    vZ As Single
    StartColor As D3DCOLORVALUE
    EndColor As D3DCOLORVALUE
    CurrentColor As D3DCOLORVALUE
    lifeTime As Long    'How long Mr. Particle Exists
    created As Long 'When this particle was created...
    status As PARTICLE_STATUS 'Does he even exist?
    Radio As Byte
End Type

Private Type pa_gro
    PrtData() As Particle
    PrtVertList() As TLVERTEX
    Position As D3DVECTOR
    light As D3DLIGHT8
    type As Integer
    nParticles As Long
    ParticleSize As Single
    Gravity As Single
    XWind As Single
    ZWind As Single
    YWind As Single
    XVariation As Single
    YVariation As Single
    ZVariation As Single
    X As Single
    Y As Single
    Z As Single
    Activated As Boolean
    Texture As Integer
    Size As Single
    Life As Integer
End Type

Dim particle_group_list() As pa_gro
Dim particle_group_count As Integer
Dim particle_group_last As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public mapInfo As mapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRangoReducido As Boolean 'rango reducido?
Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public CharList(1 To 10000) As Char

' Transparencias de Techos...
Public bTechosTransp As Boolean
Public TechosTransp As Byte
Public TechosColor(3) As Long

' Used by GetTextExtentPoint32
Private Type Size
    cx As Long
    cy As Long
End Type

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

#If ConAlfaB Then
Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long
#End If

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open sPathINIT & "Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open sPathINIT & "Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open sPathINIT & "Cuerpos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
End Sub

Sub CargarEfectos()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    N = FreeFile()
    Open sPathINIT & "Efectos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub

Sub CargarTips()
    Dim N As Integer
    Dim i As Long
    Dim NumTips As Integer
    
    N = FreeFile
    Open sPathINIT & "Tips.ayu" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumTips
    
    'Resize array
    ReDim Tips(1 To NumTips) As String * 255
    
    For i = 1 To NumTips
        Get #N, , Tips(i)
    Next i
    
    Close #N
End Sub

Sub CargarArrayLluvia()
    Dim N As Integer
    Dim i As Long
    Dim Nu As Integer
    
    N = FreeFile()
    Open sPathINIT & "Lluvia.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Nu
    
    'Resize array
    ReDim bLluvia(1 To Nu) As Byte
    
    For i = 1 To Nu
        Get #N, , bLluvia(i)
    Next i
    
    Close #N
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
On Error Resume Next
    tX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With CharList(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .active = 1
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With CharList(CharIndex)
        .active = 0
        .Criminal = 0
        .Newbie = 0 ' GSZAO
        .Atacable = False
        .bType = 0 ' GSZAO
        .FXIndex = 0
        .Invisible = False
               
        .Moving = 0
        .muerto = False
        .nombre = vbNullString
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    CharList(CharIndex).active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until CharList(LastChar).active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
On Error Resume Next

    Dim addx As Integer
    Dim addy As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With CharList(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        
        nX = X + addx
        nY = Y + addy
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        If CharIndex <> UserCharIndex Then
            Call EraseChar(CharIndex)
        End If
    End If
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With CharList(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With CharList(CharIndex)
            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS1, .Pos.X, .Pos.Y)
                Else
                    Call Audio.PlayWave(SND_PASOS2, .Pos.X, .Pos.Y)
                End If
            End If
        End With
    Else
' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Audio.PlayWave(SND_NAVEGANDO, CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y)
    End If
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With CharList(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        If X > 0 And Y > 0 Then ' GSZAO
            MapData(X, Y).CharIndex = 0
        End If
        
        addx = nX - X
        addy = nY - Y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FXIndex = FxMeditar.CHICO Or .FXIndex = FxMeditar.GRANDE Or .FXIndex = FxMeditar.MEDIANO Or .FXIndex = FxMeditar.XGRANDE Or .FXIndex = FxMeditar.XXGRANDE Then
            .FXIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2, True, False)
    End If
    
    Call UpdateUserPos ' GSZAO
        
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim J As Long
    Dim K As Long
    
    For J = UserPos.X - 8 To UserPos.X + 8
        For K = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(J, K) Then
                If MapData(J, K).ObjGrh.GrhIndex = GrhFogata Then
                    location.X = J
                    location.Y = K
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next K
    Next J
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopC As Long
    Dim Dale As Boolean
    
    loopC = 1
    Do While CharList(loopC).active And Dale
        loopC = loopC + 1
        Dale = (loopC <= UBound(CharList))
    Loop
    
    NextOpenChar = loopC
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Open sPathINIT & GraphicsFile For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(0 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        
        If Grh <> 0 Then
            With GrhData(Grh)
                'Get number of frames
                Get handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To GrhData(Grh).NumFrames)
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Get handle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                            GoTo ErrorHandler
                        End If
                    Next Frame
                    
                    Get handle, , .Speed
                    
                    If .Speed <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                Else
                    'Read in normal GRH data
                    Get handle, , .FileNum
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , GrhData(Grh).sX
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .sY
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                End If
            End With
        End If
    Wend
    
    Close handle
    
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(X, Y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
            Exit Function
        End If
        
        With CharList(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If CharList(UserCharIndex).priv > 0 And CharList(UserCharIndex).priv < 6 Then
                    If CharList(UserCharIndex).Invisible = True Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function
Private Sub DDrawGrhtoSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByRef ambient_light() As Long, Optional Alpha As Boolean = False) ' GSZAO
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
On Error GoTo error
        
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * movSpeed
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                        Exit Sub ' 0.13.5
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.GrhIndex = 0 Then Exit Sub ' GSZAO (Error en mapa, no hacer nada mas!)
    
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, ambient_light, Alpha)
    End With
Exit Sub

error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado en DDrawGrhtoSurface, por favor comuníquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef ambient_light() As Long)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, ambient_light, False, CfgDiaNoche)
    End With
End Sub

Sub DDrawTransGrhtoSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByRef ambient_light() As Long, Optional Alpha As Boolean = False, Optional ByVal killAtEnd As Byte = 1) ' GSZAO
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
On Error GoTo error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * movSpeed
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                        Exit Sub ' 0.13.5
                    End If
                End If
            End If
        End If
    End If
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        If X < BackBufferRect.Left Then
            SourceRect.Left = SourceRect.Left - X
            X = 0
        End If
       
        If Y < BackBufferRect.Top Then
            SourceRect.Top = SourceRect.Top - Y
            Y = 0
        End If
        
        'Draw
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, ambient_light, Alpha, True)
    End With
Exit Sub

error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado en DDrawTransGrhtoSurface, por favor comuníquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

#If ConAlfaB = 1 Then

Sub DDrawTransGrhtoSurfaceAlpha(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal killAtEnd As Byte = 1)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
On Error GoTo error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * movSpeed
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                        If killAtEnd Then Exit Sub ' 0.13.5
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
         Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, LightRGB_Default_Alpha, True, CfgDiaNoche)
    End With
Exit Sub

error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuníquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

#End If 'ConAlfaB = 1

Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
    Dim BMHeader As BITMAPFILEHEADER
    Dim BINFOHeader As BITMAPINFOHEADER
    
    Open BmpFile For Binary Access Read As #1
    
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    
    Close #1
    
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight
End Function

'Sub DrawGrhtoHdc(ByVal hdc As Long, ByVal GrhIndex As Integer, ByRef SourceRect As RECT, ByRef destRect As RECT)
'*****************************************************************
'Draws a Grh's portion to the given area of any Device Context
'*****************************************************************
    'Call SurfaceDB.Surface(GrhData(GrhIndex).FileNum).BltToDC(hdc, SourceRect, destRect)
'End Sub

Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, ByVal srchdc As Long, ByRef SourceRect As RECT, ByRef destRect As RECT, ByVal TransparentColor As Long)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 27/07/2012 - ^[GS]^
'*************************************************************
    Dim color As Long
    Dim X As Long
    Dim Y As Long
    
    For X = SourceRect.Left To SourceRect.Right
        For Y = SourceRect.Top To SourceRect.Bottom
            color = GetPixel(srchdc, X, Y)
            
            If color <> TransparentColor Then
                Call SetPixel(dsthdc, destRect.Left + (X - SourceRect.Left), destRect.Top + (Y - SourceRect.Top), color)
            End If
        Next Y
    Next X
End Sub

Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, ByRef Picture As StdPicture, ByVal X1 As Single, ByVal Y1 As Single, Optional Width1, Optional Height1, Optional X2, Optional Y2, Optional Width2, Optional Height2)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 12/28/2009
'Draw Picture in the PictureBox
'*************************************************************

Call PictureBox.PaintPicture(Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2)
End Sub


Sub RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 28/09/2014 - ^[GS]^
'Renders everything to the viewport
'**************************************************************
    Dim Y           As Long     'Keeps track of where on map we are
    Dim X           As Long     'Keeps track of where on map we are
    
    Dim screenx     As Integer  'Keeps track of where to place tile on screen
    Dim screeny     As Integer  'Keeps track of where to place tile on screen
    
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    
    'Figure out Ends and Starts of screen
    screenminY = TileY - HalfWindowTileHeight
    screenmaxY = TileY + HalfWindowTileHeight
    screenminX = TileX - HalfWindowTileWidth
    screenmaxX = TileX + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < YMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        screeny = 1
    End If
    
    If screenmaxY < YMaxMapSize Then
        screenmaxY = screenmaxY + 1
    ElseIf screenmaxY > YMaxMapSize Then
        screenmaxY = YMaxMapSize
    End If
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        screenx = 1
    End If
    
    If screenmaxX < XMaxMapSize Then
        screenmaxX = screenmaxX + 1
    ElseIf screenmaxX > XMaxMapSize Then
        screenmaxX = XMaxMapSize
    End If
    
    ParticleOffsetX = (Engine_PixelPosX(screenminX) - PixelOffsetX)
    ParticleOffsetY = (Engine_PixelPosY(screenminY) - PixelOffsetY)
    
    screenminY = screenminY
    
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            
            'Layer 1 **********************************
            Call DDrawGrhtoSurface(MapData(X, Y).Graphic(1), _
                (screenx - 1) * TilePixelWidth + PixelOffsetX, _
                (screeny - 1) * TilePixelHeight + PixelOffsetY, _
                0, 1, MapData(X, Y).light_value)
            '******************************************
            
            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(2), _
                    (screenx - 1) * TilePixelWidth + PixelOffsetX, _
                    (screeny - 1) * TilePixelHeight + PixelOffsetY, _
                    1, 1, MapData(X, Y).light_value)
            End If
            '******************************************
            
            screenx = screenx + 1
        Next X
        
        'Reset ScreenX to original value and increment ScreenY
        screenx = screenx - X + screenminX
        screeny = screeny + 1
    Next Y
    
    'Draw Transparent Layers
    screeny = (minYOffset - TileBufferSize)
    For Y = minY To maxY
        screenx = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = screenx * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = screeny * TilePixelHeight + PixelOffsetY
            
            With MapData(X, Y)
                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(.ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value)
                End If
                '***********************************************
                
                'Char layer ************************************
                If .CharIndex <> 0 Then
                     Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, MapData(X, Y).light_value)
                End If
                '*************************************************
                
                'GSZAO Dibujamos el valor en el render.
                If .RenderValue.Activated Then
                    modRenderValue.Draw X, Y, PixelOffsetXTemp + 20, PixelOffsetYTemp - 30
                End If

                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    'Draw
                    Call DDrawTransGrhtoSurface(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value)
                End If
                '************************************************
            End With
            
            screenx = screenx + 1
        Next X
        screeny = screeny + 1
    Next Y

    Call Effect_UpdateAll ' GSZAO actualizamos los efectos de particulas aquí!
    'DirectDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(255, 0, 0, 0)
    
    ' GSZAO!
    'Draw blocked tiles and grid
    screeny = minYOffset - TileBufferSize
    For Y = minY To maxY
        screenx = minXOffset - TileBufferSize
        For X = minX To maxX
            'Layer 4 **********************************
            If MapData(X, Y).Graphic(4).GrhIndex Then
                'Draw
                If bTecho Then ' está bajo techo
                    If TechoActivo(X, Y) = False Then ' GSZAO: No hace invisible los techos de las "otras" casas en pantalla!
                        Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                            screenx * TilePixelWidth + PixelOffsetX, _
                            screeny * TilePixelHeight + PixelOffsetY, _
                            1, 1, MapData(X, Y).light_value) ' techo normal
                    ElseIf TechosTransp >= 5 Then
                        Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                            screenx * TilePixelWidth + PixelOffsetX, _
                            screeny * TilePixelHeight + PixelOffsetY, _
                            1, 1, TechosColor())
                        bTechosTransp = True ' hace transparente
                    Else
                        bTechosTransp = False ' no actualizar techos :)
                    End If
                Else ' no está bajo techo
                    If TechosTransp >= 195 Then
                        Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                            screenx * TilePixelWidth + PixelOffsetX, _
                            screeny * TilePixelHeight + PixelOffsetY, _
                            1, 1, MapData(X, Y).light_value) ' techo normal
                        bTechosTransp = False
                    Else
                        Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                            screenx * TilePixelWidth + PixelOffsetX, _
                            screeny * TilePixelHeight + PixelOffsetY, _
                            1, 1, TechosColor()) ' techo apareciendo
                        bTechosTransp = True
                    End If
                End If
            End If
            '**********************************
            screenx = screenx + 1
        Next X
        screeny = screeny + 1
    Next Y
            
    If bRain = True Then ' Llueve (particulas)
         If WeatherEffectIndex <= 0 Then
                WeatherEffectIndex = Effect_Rain_Begin(9, 100)
            ElseIf effect(WeatherEffectIndex).EffectNum <> EffectNum_Rain Then
                Effect_Kill WeatherEffectIndex
                WeatherEffectIndex = Effect_Rain_Begin(9, 100)
            ElseIf Not effect(WeatherEffectIndex).Used Then
                WeatherEffectIndex = Effect_Rain_Begin(9, 100)
            End If
    Else
        If WeatherEffectIndex > 0 Then ' Ya no llueve!
            If effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
        End If
    End If
    
    
    If Y < 100 And X < 100 Then ' GSZAO - Fix
        'Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                                ScreenX * TilePixelWidth + PixelOffsetX, _
                                ScreenY * TilePixelHeight + PixelOffsetY, _
                                1, 1, TechosColor()) ' techo apareciendo
    End If

    If bRain = True Then
        Device_Textured_Render -10, -10, TEXTURERAIN, FULLRECT, rgbNormal(), False, True
    End If
    
    'Foco de visión reducida (GSZAO)
    If bRangoReducido = True Then
        vAngle = vAngle + timerTicksPerFrame * 2 * 0.1
        
        Device_Textured_Render 0, 0, TEXTUREFOCUSBACK, FULLRECT, rgbNormal(), False, True
        Device_Textured_Render 0, 0, TEXTUREFOCUS, FULLRECT, rgbNormal(), False, True, , vAngle 'Giro Horario
        Device_Textured_Render 0, 0, TEXTUREFOCUS, FULLRECT, rgbNormal(), False, True, , -vAngle 'Giro Antihorario
        
        'DrawText 5, 10, vAngle, D3DColorXRGB(255, 255, 255)
    End If
    
    'MINIMAP LAYER**********
    If MiniMapEnabled = True Then
        Call MiniMap_Render(Minimap.X, Minimap.Y)
    End If
    '***********************
    
    'Dead
    If CharList(UserCharIndex).muerto = True Then
        Device_Textured_Render 0, 0, TEXTUREDEAD, FULLRECT, DEADCOLOR(), False, True
    End If
    
    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY
End Sub

Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    End If
    
    DoFogataFx
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            CharList(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
                And CharList(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
                And CharList(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                And CharList(UserCharIndex).Pos.Y <= Y
    End If
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero - complete rewrite
'Last Modify Date: 20/07/2012 - ^[GS]^
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(DirectD3D8, DirGraficos, ClientAOSetup.byMemory)
    
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128
    
    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
    'Dim SurfaceDesc As DDSURFACEDESC2
    'Dim ddck As DDCOLORKEY
    
    movSpeed = 1
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Load Rain Texture
    'Standelf: &HFFFAFFAC is better for *.png
    Dim TexInfo As D3DXIMAGE_INFO_A

    Set TEXTURERAIN = DirectD3D8.CreateTextureFromFileEx(DirectDevice, DirGraficos & "experimental_lluvia.png", D3DX_DEFAULT, _
                D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFFFAFFAC, TexInfo, ByVal 0)
    Set TEXTUREFOCUS = DirectD3D8.CreateTextureFromFileEx(DirectDevice, DirGraficos & "experimental2.png", D3DX_DEFAULT, _
                D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFFFAFFAC, TexInfo, ByVal 0)
    Set TEXTUREFOCUSBACK = DirectD3D8.CreateTextureFromFileEx(DirectDevice, DirGraficos & "experimental2bg.png", D3DX_DEFAULT, _
                D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFFFAFFAC, TexInfo, ByVal 0)
    Set TEXTUREDEAD = DirectD3D8.CreateTextureFromFileEx(DirectDevice, DirGraficos & "muerteexperimental.png", D3DX_DEFAULT, _
                D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFFFAFFAC, TexInfo, ByVal 0)
                
    With FULLRECT
        .Bottom = 416
        .Right = 544
    End With
    
    'Init the common color
    rgbNormal(0) = D3DColorXRGB(255, 255, 255)
    rgbNormal(1) = D3DColorXRGB(255, 255, 255)
    rgbNormal(2) = D3DColorXRGB(255, 255, 255)
    rgbNormal(3) = D3DColorXRGB(255, 255, 255)
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    'Set the dest rect
    With MainDestRect
        .Left = TilePixelWidth * TileBufferSize - TilePixelWidth
        .Top = TilePixelHeight * TileBufferSize - TilePixelHeight
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
On Error GoTo 0
    
    
    If (frmCargando.Visible = True) Then
        frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 4
        Call AddtoRichTextBox(frmCargando.status, "> Cargando indices... ", 255, 255, 255, True, False, True) ' GSZAO
    End If
    
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarEfectos
    
    LTLluvia(0) = 224
    LTLluvia(1) = 352
    LTLluvia(2) = 480
    LTLluvia(3) = 608
    LTLluvia(4) = 736
    
    If (frmCargando.Visible = True) Then
        frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 5
        Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False) ' GSZAO
    End If
    DoEvents
        
    If (frmCargando.Visible = True) Then Call AddtoRichTextBox(frmCargando.status, "> Ajustando graficos... ", 255, 255, 255, True, False, True) ' GSZAO
    
    Call LoadGraphics
    
    If (frmCargando.Visible = True) Then
        frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 6
        Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False) ' GSZAO
    End If
    DoEvents
    
    If (frmCargando.Visible = True) Then Call AddtoRichTextBox(frmCargando.status, "> Cargando fuentes... ", 255, 255, 255, True, False, True) ' GSZAO
    Engine_Init_FontTextures
    Engine_Init_FontSettings
    
    InitTileEngine = True
End Function
Public Sub DirectXInit()

    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
    
    DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    With D3DWindow
    
        .Windowed = True
        .SwapEffect = IIf(ClientAOSetup.bVSync = 0, D3DSWAPEFFECT_COPY, D3DSWAPEFFECT_COPY_VSYNC)
        
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = frmMain.MainViewPic.ScaleWidth
        .BackBufferHeight = frmMain.MainViewPic.ScaleHeight
        
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.MainViewPic.hwnd
        
    End With
    
    Dim ModoD3D As CONST_D3DCREATEFLAGS ' GSZAO
    If ClientAOSetup.bVertex = 0 Then ' Software
        ModoD3D = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    ElseIf ClientAOSetup.bVertex = 1 Then ' Hardware
        ModoD3D = D3DCREATE_HARDWARE_VERTEXPROCESSING
    ElseIf ClientAOSetup.bVertex = 2 Then ' Software & Hardware
        ModoD3D = D3DCREATE_MIXED_VERTEXPROCESSING
    End If

    Set DirectDevice = DirectD3D.CreateDevice( _
                        D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                        frmMain.MainViewPic.hwnd, _
                        ModoD3D, _
                        D3DWindow)
    
    
    
    Engine_Init_ParticleEngine
    Engine_Init_RenderStates
    
    If DirectDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectX. Por favor asegúrese de tener la última versión correctamente instalada."
        Exit Sub
    End If
    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegúrese de tener la última versión correctamente instalada."
        Exit Sub
    End If
    
End Sub
Private Sub Engine_Init_RenderStates()

    DirectDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
        
    'Set the render states
    DirectDevice.SetRenderState D3DRS_LIGHTING, False
    DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DirectDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    DirectDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    DirectDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    DirectDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
End Sub
Public Sub DeinitTileEngine()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Destroys all DX objects
'***************************************************
On Error Resume Next

        'Set no texture in the device to avoid memory leaks

        If Not DirectDevice Is Nothing Then
                DirectDevice.SetTexture 0, Nothing
        End If
        
        Dim i As Long
        
        '   Clean Particles

        For i = 1 To UBound(ParticleTexture)

                If Not ParticleTexture(i) Is Nothing Then Set ParticleTexture(i) = Nothing

        Next i

    Set DirectD3D = Nothing
    
    Set DirectX = Nothing
    Set DirectDevice = Nothing
    
    'Destroy static textures
    Set TEXTURERAIN = Nothing
    Set TEXTUREFOCUS = Nothing
    Set TEXTUREFOCUSBACK = Nothing
    Set TEXTUREDEAD = Nothing
End Sub

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
On Error Resume Next
'***************************************************
'Author: Arron Perkins
'Last Modification: 05/09/2012 - ^[GS]^
'Updates the game's model and renders everything.
'***************************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    '****** Set main view rectangle ******
    MainViewRect.Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
    MainViewRect.Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
       
    'If CfgDiaNoche = True Then Ambient_Set 120, 120, 150     ' GSZ default DIA
    
    If EngineRun Then
    
        DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        DirectDevice.BeginScene
        
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
        
        If CfgDiaNoche = True Then ' GSZ
            If GetTickCount() - AmbientLastCheck >= 10000 Then
                Ambient_Check
                AmbientLastCheck = GetTickCount()
            End If
            If modAmbiente.Fade Then Ambient_Fade
        End If
        
        '****** Update screen ******
        If UserCiego Then
            Call CleanViewPort
        Else
            Call Engine_UpdateTransp ' GSZAO
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
        End If

        If Dialogos.NeedRender Then Call Dialogos.Render ' GSZAO
        If Cartel Then Call DibujarCartel ' GSZAO
        If DialogosClanes.Activo Then Call DialogosClanes.Draw ' GSZAO
        
            DirectDevice.EndScene
        DirectDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        
        'Si está activado el FragShooter y está esperando para sacar una foto, lo hacemos:
        If ClientAOSetup.bActive Then ' 0.13.5
            If FragShooterCapturePending Then
                DoEvents
                Call ScreenCapture(True)
                FragShooterCapturePending = False
            End If
        End If
        
        'FPS update - Dunkansdk
        If fpsLastCheck + 1000 < GetTickCount Then
            FramesPerSecCounter = 1
            FPS = FramesPerSecCounter
            fpsLastCheck = GetTickCount
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
                
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
        FPS = 1000 / timerElapsedTime
        'ParticleTimer = timerElapsedTime * 0.05 GDK: Esto no se usa.
        
    End If
End Sub

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByRef light_value() As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/2012 - ^[GS]^
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long

    With CharList(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            If EstaAtacando = False Then ' animación de arma
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
            End If

            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1

            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        If UserCharIndex = CharIndex Or Not .Invisible Then 'GDK: Re hice esto, lo ordené mejor, y lo dejé andando para el invi y el nombre de la barca
            movSpeed = 0.5
                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, light_value, .Invisible)
            
            If .Head.Head(.Heading).GrhIndex Then

                'Draw Head
                Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, light_value, .Invisible)
        
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, 0, light_value, .Invisible)
        
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, light_value, .Invisible)
        
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, light_value, .Invisible)
        
                'Slash of life by azhiralh
                'If Len(.Nombre) > 0 And UserCharIndex = CharIndex Then ' GSZ
                '    If UserEstado = 0 And (Abs(MouseTileX - .Pos.X) < 2 And (Abs(MouseTileY - .Pos.Y)) < 2) Then
                '        Draw_FillBox PixelOffsetX - 14, PixelOffsetY + 20, 60, 8, D3DColorARGB(100, 0, 0, 0), D3DColorARGB(100, 200, 200, 200)
                '        Draw_FillBox PixelOffsetX - 14, PixelOffsetY + 20, (((UserMinHP / 60) / (UserMaxHP / 60)) * 60), 8, D3DColorARGB(120, 240, 0, 0), D3DColorARGB(1, 200, 200, 200)
                '    End If
                'End If
            End If
            
            'Draw name over head
            If LenB(.nombre) > 0 Then
                If Nombres And (CfgSiempreNombres = True Or (esGM(UserCharIndex) Or Abs(MouseTileX - .Pos.X) < 2 And (Abs(MouseTileY - .Pos.Y)) < 2)) Then
                    Pos = getTagPosition(.nombre)
                    'Pos = InStr(.Nombre, "<")
                    'If Pos = 0 Then Pos = Len(.Nombre) + 2
        
                    If .priv = 0 Then
                        If .Atacable Then
                            color = D3DColorXRGB(ColoresPJ(48).r, ColoresPJ(48).g, ColoresPJ(48).b)
                        Else
                            If .muerto Then
                                color = D3DColorXRGB(ColoresPJ(51).r, ColoresPJ(51).g, ColoresPJ(51).b) ' GSZAO
                            ElseIf .Newbie Then
                                color = D3DColorXRGB(ColoresPJ(47).r, ColoresPJ(47).g, ColoresPJ(47).b) ' GSZAO
                            Else
                                If .Criminal Then
                                    color = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                Else
                                    color = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                End If
                            End If
                        End If
                    Else
                        color = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                    End If
        
                    'Nick
                    line = Left$(.nombre, Pos - 2)
                    Call DrawText(PixelOffsetX - (Len(line) * 6 / 2) + 14, PixelOffsetY + 30, line, color)
        
                    'Clan
                    line = mid$(.nombre, Pos)
                    Call DrawText(PixelOffsetX - (Len(line) * 6 / 2) + 28, PixelOffsetY + 45, line, color)
                End If
            End If
        '    End If
        'ElseIf esGM(UserCharIndex) Then ' GSZAO
        '    'Draw Body
        '    If .Body.Walk(.Heading).GrhIndex Then _
        '        Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, light_value)
        
        End If

        'Debug.Print "inv=" & .invisible
        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex)   '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        movSpeed = 1
        'Draw FX
        If .FXIndex <> 0 Then
#If (ConAlfaB = 1) Then
            Call DDrawTransGrhtoSurfaceAlpha(.fX, PixelOffsetX + FxData(.FXIndex).OffsetX, PixelOffsetY + FxData(.FXIndex).OffsetY, 1, 1)
#Else
            Call DDrawTransGrhtoSurface(.fX, PixelOffsetX + FxData(.FXIndex).OffsetX, PixelOffsetY + FxData(.FXIndex).OffsetY, 1, 1, LightRGB_Default)
#End If
            
            'Check if animation is over
            If .fX.Started = 0 Then _
                .FXIndex = 0
        End If
    End With
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Long, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/03/2012 - ^[GS]^
'Sets an FX to the character.
'***************************************************
    With CharList(CharIndex)
        .FXIndex = fX
        
        If .FXIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If
    End With
End Sub

Private Sub CleanViewPort()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Fills the viewport with black.
'***************************************************
    Dim r As RECT
    'Call BackBufferSurface.BltColorFill(r, vbBlack)
End Sub


Public Sub Device_Textured_Render(ByVal X As Integer, ByVal Y As Integer, ByVal Texture As Direct3DTexture8, ByRef Src_Rect As RECT, ByRef light_value() As Long, Optional Alpha As Boolean = False, Optional ByVal Ambient As Boolean = True, Optional IsInventory As Boolean = False, Optional ByVal Angle As Single = 0)
'***************************************************
'Last Modify Date: 10/04/2013 - ^[GS]^
'***************************************************

    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim SRDesc As D3DSURFACE_DESC

    Dim NewX               As Single
    Dim NewY               As Single
    Dim SinRad             As Single
    Dim CosRad             As Single
    Dim Width              As Single
    Dim Height             As Single
    Dim RadAngle           As Single
    Dim CenterX            As Single
    Dim CenterY            As Single
    Dim Index              As Integer
        
    If Ambient = True And IsInventory = False Then
        If (light_value(0) = 0) Then light_value(0) = base_light
        If (light_value(1) = 0) Then light_value(1) = base_light
        If (light_value(2) = 0) Then light_value(2) = base_light
        If (light_value(3) = 0) Then light_value(3) = base_light
    Else
        'If IsInventory = True Then
        base_light = ARGB(200, 200, 200, 255) ' luz al maximo
        light_value(0) = base_light
        light_value(1) = base_light
        light_value(2) = base_light
        light_value(3) = base_light
    End If

    With dest_rect
        .Bottom = Y + (Src_Rect.Bottom - Src_Rect.Top) ' src_height
        .Left = X
        .Right = X + (Src_Rect.Right - Src_Rect.Left)
        .Top = Y
    End With
    
    Dim texwidth As Long, texheight As Long
    Texture.GetLevelDesc 0, SRDesc
    texwidth = SRDesc.Width
    texheight = SRDesc.Height
        
    Geometry_Create_Box temp_verts(), dest_rect, Src_Rect, light_value(), texwidth, texheight, 0
    
    With DirectDevice

        .SetTexture 0, Texture
    
                    ' ***************** Angulo *****************
    
                    If Angle <> 0 And Angle <> 360 Then
                
                            RadAngle = Angle * DegreeToRadian
    
                            CenterX = X + (texwidth * 0.5)
                            CenterY = Y + (texheight * 0.5)
        
                            SinRad = Sin(RadAngle)
                            CosRad = Cos(RadAngle)
        
                            For Index = 0 To 3
        
                                    NewX = CenterX + (temp_verts(Index).X - CenterX) * -CosRad - (temp_verts(Index).Y - CenterY) * -SinRad
                                    NewY = CenterY + (temp_verts(Index).Y - CenterY) * -CosRad + (temp_verts(Index).X - CenterX) * -SinRad
        
                                    temp_verts(Index).X = NewX
                                    temp_verts(Index).Y = NewY
        
                            Next Index
        
                    End If
    
                    ' ***************** /Angulo *****************
    
        If Alpha Then
            '.SetRenderState D3DRS_SRCBLEND, 3
            Call .SetRenderState(D3DRS_DESTBLEND, 2)
        End If
        
        Call .DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0)))
        
        If Alpha Then
            Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
            Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
            Call .SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        End If
    End With
End Sub
Public Sub Draw_FillBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, color As Long, outlinecolor As Long, Optional UsaOutline As Boolean = True)

'@param: UsaOutline: si esta en false, solo se dibuja el recuadro sin contorno, evitamos usarlo cuando no queremos. GDK
    Static box_rect As RECT
    Static rgb_list(3) As Long
    Static Vertex(3) As TLVERTEX
    
    rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color
    
    If UsaOutline = True Then
        Static Vertex2(3) As TLVERTEX
        Static rgb_list2(3) As Long
        Static Outline As RECT
        
        rgb_list2(0) = outlinecolor
        rgb_list2(1) = outlinecolor
        rgb_list2(2) = outlinecolor
        rgb_list2(3) = outlinecolor
    End If
    
    With box_rect
        .Bottom = Y + Height
        .Left = X
        .Right = X + Width
        .Top = Y
    End With
    
    If UsaOutline = True Then
        With Outline
            .Bottom = Y + Height + 2
            .Left = X - 2
            .Right = X + Width + 2
            .Top = Y - 2
        End With
    
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    End If
    
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), 0, 0
    
    DirectDevice.SetTexture 0, Nothing
    
    If UsaOutline = True Then
        DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    End If
    
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
End Sub

Public Sub DrawText(ByVal Left As Long, ByVal Top As Long, ByVal Text As String, ByVal color As Long, Optional ByVal Alpha As Byte = 255, Optional ByVal Center As Boolean = False, Optional correr As Boolean = True, Optional ByVal fontNum As Byte = 1)
    If Alpha <> 255 Then
        Dim aux As D3DCOLORVALUE
        'Obtener_RGB Color, r, g, b
        ARGBtoD3DCOLORVALUE color, aux
        color = D3DColorARGB(Alpha, aux.r, aux.g, aux.b)
    End If
    Engine_Render_Text cfonts(fontNum), Text, Left, Top, color, Center, Alpha
End Sub

Public Sub Text_Render_Special(ByVal intX As Integer, ByVal intY As Integer, ByRef strText As String, ByVal lngColor As Long, Optional bolCentred As Boolean = False)  ' GSZAO
'*****************************************************************
'Text_Render_Special by ^[GS]^
'*****************************************************************
    
    If LenB(strText) <> 0 Then
        lngColor = ColorToDX8(lngColor)
        Call Engine_Render_Text(cfonts(1), strText, intX, intY, lngColor, bolCentred)
    End If
    
End Sub ' GSZAO

Private Function Es_Emoticon(ByVal ascii As Byte) As Boolean ' GSZAO
'*****************************************************************
'Emoticones by ^[GS]^
'*****************************************************************
    Es_Emoticon = False
    If (ascii = 129 Or ascii = 137 Or ascii = 141 Or ascii = 143 Or ascii = 144 Or ascii = 157 Or ascii = 160) Then
        Es_Emoticon = True
    End If
End Function ' GSZAO

Private Sub Engine_Render_Text(ByRef UseFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal color As Long, Optional ByVal Center As Boolean = False, Optional ByVal Alpha As Byte = 255)
'*****************************************************************
'Render text with a custom font
'*****************************************************************
    Dim TempVA(0 To 3) As TLVERTEX
    Dim tempstr() As String
    Dim Count As Integer
    Dim ascii() As Byte
    Dim Row As Integer
    Dim u As Single
    Dim v As Single
    Dim i As Long
    Dim J As Long
    Dim KeyPhrase As Byte
    Dim TempColor As Long
    Dim ResetColor As Byte
    Dim SrcRect As RECT
    Dim v2 As D3DVECTOR2
    Dim v3 As D3DVECTOR2
    Dim YOffset As Single
    
    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    
    'Analizar mensaje, palabra por palabra... GSZAO
    Dim NewText As String
    
    tempstr = Split(Text, Chr(32))
    NewText = Text
    Text = vbNullString
    
    For i = 0 To UBound(tempstr)
        If tempstr(i) = ":)" Or tempstr(i) = "=)" Then
            tempstr(i) = Chr(129)
        ElseIf tempstr(i) = ":@" Or tempstr(i) = "=@" Then
            tempstr(i) = Chr(137)
        ElseIf tempstr(i) = ":(" Or tempstr(i) = "=(" Then
            tempstr(i) = Chr(141)
        ElseIf tempstr(i) = "^^" Or tempstr(i) = "^_^" Then
            tempstr(i) = Chr(143)
        ElseIf tempstr(i) = ":D" Or tempstr(i) = "=D" Then
            tempstr(i) = Chr(144)
        ElseIf tempstr(i) = "xD" Or tempstr(i) = "XD" Then
            tempstr(i) = Chr(157)
        ElseIf tempstr(i) = ":S" Or tempstr(i) = "=S" Then
            tempstr(i) = Chr(160)
        End If
        Text = Text & Chr(32) & tempstr(i)
    Next
    ' Made by ^[GS]^ for GSZAO
    
    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)
    
    'Set the temp color (or else the first character has no color)
    TempColor = color

    'Set the texture
    DirectDevice.SetTexture 0, UseFont.Texture
    
    If Center Then
        X = X - Engine_GetTextWidth(cfonts(1), Text) * 0.5
    End If
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)
        If Len(tempstr(i)) > 0 Then
            YOffset = i * UseFont.CharHeight
            Count = 0
        
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
        
            'Loop through the characters
            For J = 1 To Len(tempstr(i))

                'Copy from the cached vertex array to the temp vertex array
                CopyMemory TempVA(0), UseFont.HeaderInfo.CharVA(ascii(J - 1)).Vertex(0), 32 * 4
                
                'Set up the verticies
                TempVA(0).X = X + Count
                TempVA(0).Y = Y + YOffset
                
                TempVA(1).X = TempVA(1).X + X + Count
                TempVA(1).Y = TempVA(0).Y
                
                TempVA(2).X = TempVA(0).X
                TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
                
                TempVA(3).X = TempVA(1).X
                TempVA(3).Y = TempVA(2).Y
                
                'Set the colors
                If Es_Emoticon(ascii(J - 1)) Then ' GSZAO los colores no afectan a los emoticones!
                    TempVA(0).color = -1
                    TempVA(1).color = -1
                    TempVA(2).color = -1
                    TempVA(3).color = -1
                    If (ascii(J - 1) <> 157) Then Count = Count + 5   ' Los emoticones tienen tamaño propio (despues hay que cargarlos "correctamente" para evitar hacer esto)
                Else
                    TempVA(0).color = TempColor
                    TempVA(1).color = TempColor
                    TempVA(2).color = TempColor
                    TempVA(3).color = TempColor
                End If
                
                'Draw the verticies
                DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0))
                
                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(ascii(J - 1))
    
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = color
                End If
                
            Next J
            
        End If
    Next i

End Sub

Public Function ARGBtoD3DCOLORVALUE(ByVal ARGB As Long, ByRef color As D3DCOLORVALUE)
Dim dest(3) As Byte
CopyMemory dest(0), ARGB, 4
color.a = dest(3)
color.r = dest(2)
color.g = dest(1)
color.b = dest(0)
End Function

Public Function ARGB(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal a As Long) As Long
        
    Dim c As Long
        
    If a > 127 Then
        a = a - 128
        c = a * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    Else
        c = a * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    End If
    
    ARGB = c

End Function

Private Function Engine_GetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth
'***************************************************
Dim i As Integer

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    'Loop through the text
    For i = 1 To Len(Text)
        
        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
        
    Next i

End Function

Sub Engine_Init_FontTextures()
On Error GoTo eDebug:
'*****************************************************************
'Init the custom font textures
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures
'*****************************************************************
    Dim TexInfo As D3DXIMAGE_INFO_A

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '*** Default font ***
    
    'Set the texture
    Set cfonts(1).Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, FileRequest(DirGraficos & "Font.png"), D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, 0, TexInfo, ByVal 0)
    
    'Store the size of the texture
    cfonts(1).TextureSize.X = TexInfo.Width
    cfonts(1).TextureSize.Y = TexInfo.Height
    
    Exit Sub
eDebug:
    If Err.Number = "-2005529767" Then
        MsgBox "Error en la textura de fuente utilizada " & DirGraficos & "Font.png.", vbCritical
        End
    End If
    End

End Sub

Sub Engine_Init_FontSettings()
'*****************************************************************
'Init the custom font settings
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontSettings
'*****************************************************************
    Dim FileNum As Byte
    Dim LoopChar As Long
    Dim Row As Single
    Dim u As Single
    Dim v As Single

    '*** Default font ***

    'Load the header information
    FileNum = FreeFile
    Open FileRequest(sPathINIT & "Font.dat") For Binary As #FileNum
        Get #FileNum, , cfonts(1).HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    cfonts(1).CharHeight = cfonts(1).HeaderInfo.CellHeight - 4
    cfonts(1).RowPitch = cfonts(1).HeaderInfo.BitmapWidth \ cfonts(1).HeaderInfo.CellWidth
    cfonts(1).ColFactor = cfonts(1).HeaderInfo.CellWidth / cfonts(1).HeaderInfo.BitmapWidth
    cfonts(1).RowFactor = cfonts(1).HeaderInfo.CellHeight / cfonts(1).HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) \ cfonts(1).RowPitch
        u = ((LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) - (Row * cfonts(1).RowPitch)) * cfonts(1).ColFactor
        v = Row * cfonts(1).RowFactor

        'Set the verticies
        With cfonts(1).HeaderInfo.CharVA(LoopChar)
            .Vertex(0).color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).rhw = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).Z = 0
            
            .Vertex(1).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).rhw = 1
            .Vertex(1).tu = u + cfonts(1).ColFactor
            .Vertex(1).tv = v
            .Vertex(1).X = cfonts(1).HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).Z = 0
            
            .Vertex(2).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).rhw = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + cfonts(1).RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = cfonts(1).HeaderInfo.CellHeight
            .Vertex(2).Z = 0
            
            .Vertex(3).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).rhw = 1
            .Vertex(3).tu = u + cfonts(1).ColFactor
            .Vertex(3).tv = v + cfonts(1).RowFactor
            .Vertex(3).X = cfonts(1).HeaderInfo.CellWidth
            .Vertex(3).Y = cfonts(1).HeaderInfo.CellHeight
            .Vertex(3).Z = 0
        End With
        
    Next LoopChar

End Sub

Public Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal Angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
   
    If Angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.Bottom - dest.Top) / 2
       
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.Bottom - y_center) ^ 2)
       
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = 3.1459 - right_point
    End If
   
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - Angle) * radius
        y_Cor = y_center - Sin(-left_point - Angle) * radius
    End If
   
   
    '0 - Bottom left vertex
    If Textures_Width Or Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width + 0.001, (src.Bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - Angle) * radius
        y_Cor = y_center - Sin(left_point - Angle) * radius
    End If
   
   
    '1 - Top left vertex
    If Textures_Width Or Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width + 0.001, src.Top / Textures_Height + 0.001)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - Angle) * radius
        y_Cor = y_center - Sin(-right_point - Angle) * radius
    End If
   
   
    '2 - Bottom right vertex
    If Textures_Width Or Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - Angle) * radius
        y_Cor = y_center - Sin(right_point - Angle) * radius
    End If
   
   
    '3 - Top right vertex
    If Textures_Width Or Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height + 0.001)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If

End Sub

Public Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Public Function Engine_TPtoSPX(ByVal X As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'************************************************************
    
    Engine_TPtoSPX = X * 32 - screenminX * 32 + OffsetCounterX - 16
    
    'Engine_TPtoSPX = Engine_PixelPosX(X - minX) + OffsetCounterX - 288 + TileBufferSize
 
End Function
 
Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'************************************************************
    
    Engine_TPtoSPY = Y * 32 - screenminY * 32 + OffsetCounterY - 16
    'Engine_TPtoSPY = Engine_PixelPosY(Y - minY) + OffsetCounterY - 288 + TileBufferSize
 
End Function
Function Engine_PixelPosX(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************
 
    Engine_PixelPosX = (X - 1) * TilePixelWidth
 
End Function
 
Function Engine_PixelPosY(ByVal Y As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************
 
    Engine_PixelPosY = (Y - 1) * TilePixelHeight
 
End Function

Public Function Engine_SPtoTPX(ByVal X As Long) As Long
        '************************************************************
        'Screen Position to Tile Position
        'Takes the screen pixel position and returns the tile position
        '************************************************************
        
        Engine_SPtoTPX = UserPos.X + X \ TilePixelWidth - WindowTileWidth \ 2
End Function

Public Function Engine_SPtoTPY(ByVal Y As Long) As Long
        '************************************************************
        'Screen Position to Tile Position
        'Takes the screen pixel position and returns the tile position
        '************************************************************
        
        Engine_SPtoTPY = UserPos.Y + Y \ TilePixelHeight - WindowTileHeight \ 2
End Function

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************
Dim SideA As Single
Dim SideC As Single
 
    On Error GoTo ErrOut
 
    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then
 
        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90
 
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then
 
        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360
 
            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)
 
    'Side B = CenterY
 
    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)
 
    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583
 
    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
 
    'Exit function
 
Exit Function
 
    'Check for error
ErrOut:
 
    'Return a 0 saying there was an error
    Engine_GetAngle = 0
 
Exit Function
 
End Function


Private Function Engine_FToDW(F As Single) As Long
Dim buf As D3DXBuffer
    Set buf = DirectD3D8.CreateBuffer(4)
    DirectD3D8.BufferSetData buf, 0, 4, 1, F
    DirectD3D8.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function

Private Function VectorToRGBA(Vec As D3DVECTOR, fHeight As Single) As Long
Dim r As Integer, g As Integer, b As Integer, a As Integer
    r = 127 * Vec.X + 128
    g = 127 * Vec.Y + 128
    b = 127 * Vec.Z + 128
    a = 255 * fHeight
    VectorToRGBA = D3DColorARGB(a, r, g, b)
End Function

Private Sub Engine_UpdateTransp()
'***************************************************
'Author: ^[GS]^
'Last Modification: 11/08/2012 - ^[GS]^
'***************************************************
Static lastTime As Long

    If bTechosTransp = True Then
        If bTecho Then
            If Not Val(TechosTransp) <= 5 Then TechosTransp = Val(TechosTransp) - 5
        Else
            If Not TechosTransp >= 195 Then TechosTransp = TechosTransp + 5
        End If
        TechosColor(0) = D3DColorARGB(TechosTransp, TechosTransp, TechosTransp, TechosTransp)
        TechosColor(1) = D3DColorARGB(TechosTransp, TechosTransp, TechosTransp, TechosTransp)
        TechosColor(2) = D3DColorARGB(TechosTransp, TechosTransp, TechosTransp, TechosTransp)
        TechosColor(3) = D3DColorARGB(TechosTransp, TechosTransp, TechosTransp, TechosTransp)
    End If

    If CharList(UserCharIndex).muerto = True Then
        If GetTickCount - lastTime > 12 Then
            If ABD <> 0 Then
                ABD = ABD - 1
                
                DEADCOLOR(0) = D3DColorARGB(ABD, 255, 255, 255)
                DEADCOLOR(1) = DEADCOLOR(0)
                DEADCOLOR(2) = DEADCOLOR(0)
                DEADCOLOR(3) = DEADCOLOR(0)
            End If
            
        lastTime = GetTickCount
        End If
    End If
End Sub

Public Sub D3DColorToRgbList(rgb_list() As Long, color As D3DCOLORVALUE)

        rgb_list(0) = D3DColorARGB(color.a, color.r, color.g, color.b)
        rgb_list(1) = rgb_list(0)
        rgb_list(2) = rgb_list(0)
        rgb_list(3) = rgb_list(0)

End Sub

Public Function Engine_UTOV_Particle(ByVal UserIndex As Integer, _
                                     ByVal VictimIndex As Integer, _
                                     ByVal Particle_ID As Integer) As Integer

        Dim X         As Long
        Dim Y         As Long
        Dim TempIndex As Integer

'Extraido de Dunkan AO
'Editado en KuviK AO
'Traido a GSZAO

        Select Case Particle_ID

                Case 1              'Fx Telep
                        X = Engine_TPtoSPX(CharList(UserIndex).Pos.X) '+ 16
                        Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y) '+ 16
                        TempIndex = Effect_Fire_Begin(X, Y, 1, 100, , 0)

                Case 2
                        X = Engine_TPtoSPX(CharList(UserIndex).Pos.X) + 5
                        Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y) ' - 25
                        TempIndex = Effect_Rayo_Begin(X, Y, 17, 100)
                        effect(TempIndex).BindToChar = VictimIndex
                        effect(TempIndex).BindSpeed = 7
                
                Case 3 'evolucion de clase
                        X = Engine_TPtoSPX(CharList(UserIndex).Pos.X) + 16
                        Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y) + 16
                        
                        TempIndex = Effect_ChangeClass_Begin(X, Y, 17, 100, 30)
                        
                Case 4
                        X = Engine_TPtoSPX(CharList(UserIndex).Pos.X) + 25
                        Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
                        TempIndex = Effect_LissajousMedit_Begin(X, Y, 13, 50, , , 0, 2, 1, 1)

                Case 5
                        X = Engine_TPtoSPX(CharList(UserIndex).Pos.X) + 25
                        Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
                        TempIndex = Effect_LissajousMedit_Begin(X, Y, 13, 70, , 35, 2, 2, 0, 2)

                Case 6
                        X = Engine_TPtoSPX(CharList(UserIndex).Pos.X) + 25
                        Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
                        TempIndex = Effect_LissajousMedit_Begin(X, Y, 1, 100, , 40, 0, 2, 4, 3)

                Case 16
                        X = Engine_TPtoSPX(CharList(UserIndex).Pos.X) + 25
                        Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
                        TempIndex = Effect_LissajousMedit_Begin(X, Y, 1, 150, , 45, 1, 0, 5, 4)

                Case 34
                        X = Engine_TPtoSPX(CharList(UserIndex).Pos.X) + 25
                        Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
                        TempIndex = Effect_LissajousMedit_Begin(X, Y, 1, 150, , 85, 0.5, 0.7, 1, 5)
        End Select

        Engine_UTOV_Particle = TempIndex
End Function


