Attribute VB_Name = "modGameIni"
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

' GSZAO - Archivos de configuración!
Public Const fAOSetup = "AOSetup.init"
Public Const fConfigInit = "Config.init"

' GSZAO - Las variables de path se definen una sola vez (Ver Sub InitFilePaths())
Public DirGraficos As String
Public DirSound As String
Public DirMidi As String
Public DirMapas As String
Public DirExtras As String
Public DirCursores As String
Public DirGUI As String
Public DirButtons As String

Public Const nDirINIT = "\INIT\" ' Directorio INIT
Public sPathINIT As String ' Path de INITs

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal _
lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tConfigInit
    ' Opciones
    MostrarTips As Byte         ' Activa o desactiva la muestra de tips
    NumParticulas As Integer    ' Numero de particulas
    IndiceGraficos As String    ' Archivo de Indices de Graficos
    
    ' Usuario
    Nombre As String            ' Nombre de usuario
    Password As String          ' Contraseña del usuario
    Recordar As Byte            ' Activado el recordar!
    
    ' Directorio
    DirMultimedia As String     ' Directorio de multimedia
    DirMapas As String          ' Directorio de mapas
    DirGraficos As String       ' Directorio de graficos
    DirFotos As String          ' Directorio de fotos
    DirExtras As String         ' Directorio de extras (dentro de inits)
    DirSonidos As String        ' Directorio de sonidos (dentro de multimedia)
    DirMusicas As String        ' Directorio de musicas (dentro de multimedia)
    DirParticulas As String     ' Directorio de particulas (dentro de graficos)
    DirCursores As String       ' Directorio de cursores (dentro de graficos)
    DirGUI As String            ' Directorio del GUI (dentro de graficos)
    DirBotones As String        ' Directorio de botones (dentro de GUI)
    DirFrags As String          ' Directorio de frags (dentro de fotos)
    DirMuertes As String        ' Directorio de muertes (dentro de fotos)
End Type

Public Type tAOSetup
    ' VIDEO
    bVertex     As Byte     ' GSZAO - Cambia el Vortex de dibujado
    bVSync      As Boolean  ' GSZAO - Utiliza Sincronización Vertical (VSync)
    bDinamic    As Boolean  ' Utilizar carga Dinamica de Graficos o Estatica
    byMemory    As Byte     ' Uso maximo de memoria para la carga Dinamica (exclusivamente)

    ' SONIDO
    bNoMusic    As Boolean  ' Jugar sin Musica
    bNoSound    As Boolean  ' Jugar sin Sonidos
    bNoSoundEffects As Boolean  ' Jugar sin Efectos de sonido (basicamente, sonido que viene de la izquierda y de la derecha)
    lMusicVolume As Long ' Volumen de la Musica
    lSoundVolume As Long ' Volumen de los Sonidos

    ' SCREENSHOTS
    bActive     As Boolean  ' Activa el modo de screenshots
    bDie        As Boolean  ' Obtiene una screenshot al morir (si bActive = True)
    bKill       As Boolean  ' Obtiene una screenshot al matar (si bActive = True)
    byMurderedLevel As Byte ' La screenshot al matar depende del nivel de la victima (si bActive = True)
    
    ' CLAN
    bGuildNews  As Boolean      ' Mostrar Noticias del Clan al inicio
    bGldMsgConsole As Boolean   ' Activa los Dialogos de Clan
    bCantMsgs   As Byte         ' Establece el maximo de mensajes de Clan en pantalla
    
    ' GENERALEs
    bCursores   As Boolean      ' Utilizar Cursores Personalizados
End Type

Public MiCabecera As tCabecera
Public ClientConfigInit As tConfigInit
Public ClientAOSetup As tAOSetup

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
'**************************************************************
'Author: Unknown
'Last Modify Date: 29/08/2012 - ^[GS]^
'**************************************************************
    Cabecera.Desc = "GS-Zone Argentum Online MOD - Copyright GS-Zone 2013 - info@gs-zone.org - Original by Pablo Marquez " ' GSZAO
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
    
End Sub

Public Function LeerConfigInit() As tConfigInit
'**************************************************************
'Author: ^[GS]^
'Last Modify Date: 29/08/2012 - ^[GS]^
'**************************************************************
On Local Error Resume Next

    Dim N As Integer
    Dim ConfigInit As tConfigInit
    N = FreeFile
    Open sPathINIT & fConfigInit For Binary As #N
        Get #N, , MiCabecera
        Get #N, , ConfigInit
    Close #N
    
    ConfigInit.Password = RndCrypt(ConfigInit.Password, App.Path & ConfigInit.Nombre)
    
    LeerConfigInit = ConfigInit
    
End Function

Public Sub EscribirConfigInit(ByRef ImaConfigInit As tConfigInit)
'**************************************************************
'Author: ^[GS]^
'Last Modify Date: 29/08/2012 - ^[GS]^
'**************************************************************
On Local Error Resume Next

    ' GSZAO seguridad para la contraseña
    ImaConfigInit.Password = RndCrypt(ImaConfigInit.Password, App.Path & ImaConfigInit.Nombre)
    
    Dim N As Integer
    N = FreeFile
    Open sPathINIT & fConfigInit For Binary As #N
    Put #N, , MiCabecera
    Put #N, , ImaConfigInit
    Close #N
    
    ' La re convierte para el juego!
    ImaConfigInit.Password = RndCrypt(ImaConfigInit.Password, App.Path & ImaConfigInit.Nombre)
    
End Sub

Public Sub InitFilePaths()
'*************************************************
'Author: ^[GS]^
'Last modified: 25/07/2012 - ^[GS]^
'*************************************************
    If InStr(1, ClientConfigInit.IndiceGraficos, "Graficos") Then
        GraphicsFile = ClientConfigInit.IndiceGraficos
    Else
        GraphicsFile = "Graficos1.ind"
    End If
    DirGraficos = App.Path & "\" & ClientConfigInit.DirGraficos & "\"
    DirSound = App.Path & "\" & ClientConfigInit.DirMultimedia & "\" & ClientConfigInit.DirSonidos & "\"
    DirMidi = App.Path & "\" & ClientConfigInit.DirMultimedia & "\" & ClientConfigInit.DirMusicas & "\"
    DirMapas = App.Path & "\" & ClientConfigInit.DirMapas & "\"
    DirExtras = sPathINIT & "\" & ClientConfigInit.DirExtras & "\"
    DirCursores = App.Path & "\" & ClientConfigInit.DirGraficos & "\" & ClientConfigInit.DirCursores & "\"
    DirGUI = App.Path & "\" & ClientConfigInit.DirGraficos & "\" & ClientConfigInit.DirGUI & "\"
    DirButtons = App.Path & "\" & ClientConfigInit.DirGraficos & "\" & ClientConfigInit.DirGUI & "\" & ClientConfigInit.DirBotones & "\"
End Sub

Public Sub LoadClientAOSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 22/08/2013 - ^[GS]^
'**************************************************************
    Dim fHandle As Integer
    
    ' Por default
    ClientAOSetup.bDinamic = True
    ClientAOSetup.bVertex = 0 ' software
    ClientAOSetup.bVSync = False
    
    If FileExist(sPathINIT & fAOSetup, vbArchive) Then
        fHandle = FreeFile
        Open sPathINIT & fAOSetup For Binary Access Read Lock Write As fHandle
            Get fHandle, , ClientAOSetup
        Close fHandle
    End If
    
    ClientAOSetup.bGuildNews = Not ClientAOSetup.bGuildNews
    Set DialogosClanes = New clsGuildDlg ' 0.13.3
    DialogosClanes.Activo = Not ClientAOSetup.bGldMsgConsole
    DialogosClanes.CantidadDialogos = ClientAOSetup.bCantMsgs
End Sub

Public Sub SaveClientAOSetup()
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 22/08/2013 - ^[GS]^
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    
    ClientAOSetup.bNoMusic = Not Audio.MusicActivated
    ClientAOSetup.bNoSound = Not Audio.SoundActivated
    ClientAOSetup.bNoSoundEffects = Not Audio.SoundEffectsActivated
    ClientAOSetup.bGuildNews = Not ClientAOSetup.bGuildNews
    ClientAOSetup.bGldMsgConsole = Not DialogosClanes.Activo
    ClientAOSetup.bCantMsgs = DialogosClanes.CantidadDialogos
    ClientAOSetup.lMusicVolume = Audio.MusicVolume
    ClientAOSetup.lSoundVolume = Audio.SoundVolume
    
    Open sPathINIT & fAOSetup For Binary As fHandle
        Put fHandle, , ClientAOSetup
    Close fHandle
    
End Sub


Private Function SystemDrive() As String
' GSZ-AO - Obtiene la unidad del sistema
    Dim windows_dir As String
    Dim length As Long
    windows_dir = Space$(255)
    length = GetWindowsDirectory(windows_dir, Len(windows_dir))
    SystemDrive = Left$(windows_dir, 3) ' C:\
End Function

Function GetSerialHD() As Long
' GSZ-AO - Obtiene el numero de serie del disco de sistema
    Dim SerialNum As Long
    Dim res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    res = GetVolumeInformation(SystemDrive(), Temp1, _
    Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialHD = SerialNum
End Function

Public Function SEncriptar(ByVal Cadena As String) As String
' GSZ-AO - Encripta una cadena de texto
    Dim i As Long, RandomNum As Integer
    
    RandomNum = 99 * Rnd
    If RandomNum < 10 Then RandomNum = 10
    For i = 1 To Len(Cadena)
        Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) + RandomNum)
    Next i
    SEncriptar = Cadena & Chr$(Asc(Left$(RandomNum, 1)) + 10) & Chr$(Asc(Right$(RandomNum, 1)) + 10)
    DoEvents

End Function

Public Function SDesencriptar(ByVal Cadena As String) As String
' GSZ-AO - Desencripta una cadena de texto
    Dim i As Long, NumDesencriptar As String
    
    NumDesencriptar = Chr$(Asc(Left$((Right(Cadena, 2)), 1)) - 10) & Chr$(Asc(Right$((Right(Cadena, 2)), 1)) - 10)
    Cadena = (Left$(Cadena, Len(Cadena) - 2))
    For i = 1 To Len(Cadena)
        Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) - NumDesencriptar)
    Next i
    SDesencriptar = Cadena
    DoEvents

End Function

Public Function SXor(ByVal Cadena As String, ByVal Clave As String) As String
' GSZ-AO - Aplicamos un XOR por la clave indicada
    Dim i As Long, c As Integer

    If Len(Cadena) > 0 Then
        c = 1
        For i = 1 To Len(Cadena)
            If c > Len(Clave) Then c = 1
            Mid$(Cadena, i, 1) = Chr(Asc(mid$(Cadena, i, 1)) Xor Asc(mid$(Clave, c, 1)))
            c = c + 1
        Next i
    End If
    SXor = Cadena
    DoEvents

End Function
