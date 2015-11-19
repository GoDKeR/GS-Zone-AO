Attribute VB_Name = "mod_General"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

' Declaraciones de Constantes...
Public Const iniInit As String = "\Init\AOSetup.init" ' Configuración del juego
Public Const iniClient As String = "\Argentum.exe" ' Ejecutable del Cliente
Public Const iniTestWAV As String = "\Multimedia\Wav\18.wav" ' Archivo de Sonido para la Prueba
Public Const iniTestBMP As String = "\Graficos\Geek.bmp" ' Archivo Gráfico para la Prueba
Public Const iniTempDX As String = "\DXTest.txt" ' Archivo "temporal" con la información de DirectX

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

Public ClientAOSetup As tAOSetup

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Const SW_NORMAL As Long = 1

Public Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'Se fija si existe el archivo
'*************************************************
    FileExist = Dir(file, FileType) <> ""
End Function

Public Sub LeerSetup()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/03/2013 - ^[GS]^
'*************************************************
On Error Resume Next
    ' Default
    ' > Video
    frmAOSetup.cProcesar.ListIndex = 0  ' software
    frmAOSetup.chkVSync.value = 0       ' sin vsync
    frmAOSetup.chkDinamico.value = 1    ' dinamico
    frmAOSetup.pMemoria.value = 32      ' 32 MiB
    ' > Sonido
    frmAOSetup.chkSonido.value = 1      ' sonido
    frmAOSetup.chkMusica.value = 1      ' musica
    frmAOSetup.chkEfectosSonido.value = 1 ' efectos 3D
    frmAOSetup.Slider1(1).value = 100   ' volumen de musica
    frmAOSetup.Slider1(0).value = 100   ' volumen de sonido
    ' > Screenshots
    frmAOSetup.chkActScreenshots.value = 1 ' screenshots activadas
    frmAOSetup.chkScreenDie.value = 0      ' no sacar screen cuando muere
    frmAOSetup.chkScreenKill.value = 1     ' sacar screen cuando mata
    frmAOSetup.tLevelShot.Text = 40        ' de nivel 40 para arriba
    ' > Clan
    frmAOSetup.chkNewsGuild.value = 1      ' con noticias de clan al comenzar
    frmAOSetup.chkDlgGuild.value = 1       ' chat de clan en pantalla
    frmAOSetup.tGuildDlgCount.Text = 5     ' 5 mensajes por vez en pantalla
    ' > General
    frmAOSetup.chkCursoresPersonalizados.value = 1 ' cusores personalizados activados
    
    ' Hay alguna configuración?
    If FileExist(App.Path & iniInit, vbArchive) Then
        ' Cargamos...
        Dim handle As Integer
        handle = FreeFile
        Open App.Path & iniInit For Binary As handle
            Get handle, , ClientAOSetup
        Close handle
        
        ' VIDEO
        frmAOSetup.cProcesar.ListIndex = ClientAOSetup.bVertex
        frmAOSetup.chkVSync.value = IIf(ClientAOSetup.bVSync = True, 1, 0)
        If ClientAOSetup.bDinamic Then
            frmAOSetup.chkDinamico.value = 1
            frmAOSetup.lCuantoVideo.ForeColor = vbBlack
            frmAOSetup.pMemoria.EnabledSlider = True
            frmAOSetup.pMemoria.picFillColor = &H8080FF
            frmAOSetup.pMemoria.picForeColor = &H80FF80
        Else
            frmAOSetup.chkDinamico.value = 0
            frmAOSetup.lCuantoVideo.ForeColor = &H808080
            frmAOSetup.pMemoria.EnabledSlider = False
            frmAOSetup.pMemoria.picFillColor = &H808080
            frmAOSetup.pMemoria.picForeColor = &HC0C0C0
        End If
        If ClientAOSetup.byMemory >= 4 And ClientAOSetup.byMemory <= 64 Then
            frmAOSetup.pMemoria.value = ClientAOSetup.byMemory
        End If
           
        ' SONIDO
        frmAOSetup.chkSonido.value = Not ClientAOSetup.bNoSound
        frmAOSetup.chkMusica.value = Not ClientAOSetup.bNoMusic
        frmAOSetup.chkEfectosSonido.value = Not ClientAOSetup.bNoSoundEffects
        frmAOSetup.Slider1(1).value = ClientAOSetup.lMusicVolume
        frmAOSetup.Slider1(0).value = ClientAOSetup.lSoundVolume

        ' SCREENSHOTS
        frmAOSetup.chkActScreenshots.value = IIf(ClientAOSetup.bActive, 1, 0)
        frmAOSetup.chkScreenDie.value = IIf(ClientAOSetup.bDie, 1, 0)
        frmAOSetup.chkScreenKill.value = IIf(ClientAOSetup.bKill, 1, 0)
        frmAOSetup.tLevelShot.Text = Val(ClientAOSetup.byMurderedLevel)
        
        ' CLAN
        frmAOSetup.chkNewsGuild.value = IIf(ClientAOSetup.bGuildNews, 1, 0)
        frmAOSetup.chkDlgGuild.value = IIf(ClientAOSetup.bGldMsgConsole, 1, 0)
        frmAOSetup.tGuildDlgCount.Text = Val(ClientAOSetup.bCantMsgs)
        
        ' GENERAL
        frmAOSetup.chkCursoresPersonalizados.value = IIf(ClientAOSetup.bCursores, 1, 0)
    End If
End Sub
