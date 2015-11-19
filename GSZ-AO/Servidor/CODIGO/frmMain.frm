VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00101010&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GS-Z Argentum Online"
   ClientHeight    =   4470
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   9855
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4470
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer TimerSeg 
      Interval        =   1000
      Left            =   1200
      Top             =   2520
   End
   Begin VB.TextBox txStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   4200
      Top             =   3960
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFC0&
      Height          =   2895
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   4935
   End
   Begin VB.Timer packetResend 
      Interval        =   10
      Left            =   480
      Top             =   1080
   End
   Begin VB.Timer tFXMapas 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   840
      Top             =   2040
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2520
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   480
      Top             =   1560
   End
   Begin VB.Timer tLluvia 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   480
      Top             =   2040
   End
   Begin VB.Timer tEfectoLluvia 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   2040
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   120
      Top             =   1560
   End
   Begin VB.Timer tNpcAtaca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   480
      Top             =   2520
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   120
      Top             =   1080
   End
   Begin VB.Timer tNpcAI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   2520
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00202020&
      Caption         =   "Mensaje a todos los Jugadores"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4800
      TabIndex        =   0
      Top             =   3120
      Width           =   4935
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   600
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Enviar por &Consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Enviar por &Pop-Up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1080
         MaxLength       =   2048
         TabIndex        =   2
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Guardar todos los personajes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hacer un Worldsave"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label Record 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Record:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label lblServidor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GS-Z Argentum Online"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label txtIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      ToolTipText     =   "Click para copiar al portapapeles"
      Top             =   645
      Width           =   2895
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuarios jugando:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tú IP en Inet:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   750
      Width           =   1575
   End
   Begin VB.Label Escuch 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Controles"
      Begin VB.Menu mnuGuardarUsers 
         Caption         =   "&Guardar personajes"
      End
      Begin VB.Menu mnuHacerWorldsave 
         Caption         =   "&Hacer un Worldsave"
      End
      Begin VB.Menu mnuCargarWorldsave 
         Caption         =   "&Cargar último Worldsave"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPausar 
         Caption         =   "&Pausar Servidor"
      End
      Begin VB.Menu mnuReiniciar 
         Caption         =   "&Reiniciar Servidor"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApagarYGuardar 
         Caption         =   "Apagar el Servidor y &guardando cambios"
      End
      Begin VB.Menu mnuApagarSinGuardar 
         Caption         =   "Apagar el Servidor &sin guardar"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuAdmins 
         Caption         =   "&Administración de Personajes"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuTrafico 
         Caption         =   "Ver &Trafico"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSlots 
         Caption         =   "Ver &Slots de Conexión"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuDebugUser 
         Caption         =   "Ver Debug de &Usuarios"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuDebugNPC 
         Caption         =   "Ver Debug de &NPCs"
         Shortcut        =   {F6}
      End
      Begin VB.Menu chkServerHabilitado 
         Caption         =   "Servidor solo para GMs"
      End
      Begin VB.Menu line0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnban 
         Caption         =   "Unban"
         Begin VB.Menu mnuUnbanUsers 
            Caption         =   "...todos los usuarios"
         End
         Begin VB.Menu mnuUnbanIPs 
            Caption         =   "...todas las direcciónes IP"
         End
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReiniciarRespawn 
         Caption         =   "Reiniciar el Respawn de los Guardas"
      End
      Begin VB.Menu mnuDump 
         Caption         =   "&Dump - Guardar log de momento critico"
      End
      Begin VB.Menu mnuReiniciarSock 
         Caption         =   "Reiniciar Sockets"
      End
      Begin VB.Menu mnuReiniciarListen 
         Caption         =   "Reiniciar puerto de conexión [7666]"
      End
   End
   Begin VB.Menu mnuActualizar 
      Caption         =   "&Actualizar"
      Begin VB.Menu mnuRecargarAdministradores 
         Caption         =   "Recargar &Administradores [servidor.ini]"
      End
      Begin VB.Menu mnuObjetos 
         Caption         =   "Listado de &Objetos [objetos.dat]"
      End
      Begin VB.Menu mnuNPCs 
         Caption         =   "Listado de &NPC's [npcs.dat]"
      End
      Begin VB.Menu mnuHechizos 
         Caption         =   "Listado de &Hechizos [hechizos.dat]"
      End
      Begin VB.Menu mnuNombresInvalid 
         Caption         =   "Listado de Nombres &Invalidos [nombresinvalidos.txt]"
      End
      Begin VB.Menu mnuMOTD 
         Caption         =   "&Mensaje del día [motd.ini]"
      End
      Begin VB.Menu mnuBalance 
         Caption         =   "&Balance de razas/clases [balance.dat]"
      End
      Begin VB.Menu mnuServidorIni 
         Caption         =   "Configuración del &Servidor [servidor.ini]"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Panel de Control"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalirGuardando 
         Caption         =   "Apagar el Servidor &guardando cambios"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Apagar el Servidor sin guardar"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private iDay As Integer ' 0.13.5
   
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
    Dim iUserIndex As Long
    
    For iUserIndex = 1 To iniMaxUsuarios
        With UserList(iUserIndex)
            'Conexion activa? y es un usuario loggeado?
            If .ConnID <> -1 And .flags.UserLogged Then
                'Actualiza el contador de inactividad
                If .flags.Traveling = 0 Then
                    .Counters.IdleCount = .Counters.IdleCount + 1
                End If
                
                If .Counters.IdleCount >= iniInactivo Then
                    Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado.")
                    'mato los comercios seguros
                    If .ComUsu.DestUsu > 0 Then
                        If UserList(.ComUsu.DestUsu).flags.UserLogged Then
                            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                                Call WriteMensajes(.ComUsu.DestUsu, eMensajes.Mensaje001) '"Comercio cancelado por el otro usuario."
                                Call FinComerciarUsu(.ComUsu.DestUsu)
                                Call FlushBuffer(.ComUsu.DestUsu) 'flush the buffer to send the message right away
                            End If
                        End If
                        Call FinComerciarUsu(iUserIndex)
                    End If
                    Call Cerrar_Usuario(iUserIndex)
                End If
            End If
        End With
    Next iUserIndex
End Sub

Private Sub Auditoria_Timer()
On Error GoTo errhand
Static centinelSecs As Byte

centinelSecs = centinelSecs + 1

'Saco esto y pongo las llamadas donde en verdad van.
'maTih.-  |  02/03/2012

'Escuch.Caption = NumUsers
'Record.Caption = iniRecord

If iniSoloGMs = 0 Then  ' GSZAO
    chkServerHabilitado.Checked = False
Else
    chkServerHabilitado.Checked = True
End If

If centinelSecs = 5 Then
    'Every 5 seconds, we try to call the player's attention so it will report the code.
    Call modCentinela.CallUserAttention
    
    centinelSecs = 0
End If

Call PasarSegundo 'sistema de desconexion de 10 segs


Exit Sub

errhand:

Call LogError("Error en Timer Auditoria. Err: " & Err.description & " - " & Err.Number)
Resume Next

End Sub

Public Sub UpdateNpcsExp(ByVal Multiplicador As Single) ' 0.13.5
    
    Dim NpcIndex As Long
    For NpcIndex = 1 To LastNPC
        With Npclist(NpcIndex)
            .GiveEXP = .GiveEXP * Multiplicador
            .flags.ExpCount = .flags.ExpCount * Multiplicador
        End With
    Next NpcIndex
End Sub

Private Sub AutoSave_Timer()

On Error GoTo ErrHandler

    'fired every minute
    Static Minutos As Long
    Static MinutosLatsClean As Long
    Static MinsPjesSave As Long
    Static MinsSendMotd As Long
    
    Minutos = Minutos + 1
    MinsPjesSave = MinsPjesSave + 1
    MinsSendMotd = MinsSendMotd + 1
    
    If iniHappyHourActivado = True Then
        Dim tmpHappyHour As Double
    
        ' HappyHour
        iDay = Weekday(Date)
        tmpHappyHour = HappyHourDays(iDay).Multi
         
        If tmpHappyHour <> HappyHour Then ' 0.13.5
            If HappyHourActivated Then
                ' Reestablece la exp de los npcs
               If HappyHour <> 0 Then Call UpdateNpcsExp(1 / HappyHour)
             End If
           
            If tmpHappyHour = 1 Then ' Desactiva
               Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Ha concluido la Happy Hour!", FontTypeNames.FONTTYPE_DIOS))
                 HappyHourActivated = False
           
            Else ' Activa?
                If HappyHourDays(iDay).Hour = Hour(Now) And tmpHappyHour > 0 Then ' GSZAO - Es la hora pautada?
                    UpdateNpcsExp tmpHappyHour
                    
                    If HappyHour <> 1 Then
                       Call SendData(SendTarget.ToAll, 0, _
                           PrepareMessageConsoleMsg("Se ha modificado la Happy Hour, a partir de ahora las criaturas aumentan su experiencia en un " & Round((tmpHappyHour - 1) * 100, 2) & "%", FontTypeNames.FONTTYPE_DIOS))
                    Else
                       Call SendData(SendTarget.ToAll, 0, _
                           PrepareMessageConsoleMsg("¡Ha comenzado la Happy Hour! ¡Las criaturas aumentan su experiencia en un " & Round((tmpHappyHour - 1) * 100, 2) & "%!", FontTypeNames.FONTTYPE_DIOS))
                    End If
                    
                    HappyHourActivated = True
                Else
                    HappyHourActivated = False ' GSZAO
                End If
            End If
         
            HappyHour = tmpHappyHour
        End If
    Else
        ' Si estaba activado, lo deshabilitamos
        If HappyHour <> 0 Then
            Call UpdateNpcsExp(1 / HappyHour)
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Ha concluido la Happy Hour!", FontTypeNames.FONTTYPE_DIOS))
            HappyHourActivated = False
            HappyHour = 0
        End If
    End If
        
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    Call ModAreas.AreasOptimizacion
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    'Actualizamos el centinela
    Call modCentinela.PasarMinutoCentinela
    
    'Actualizamos los objetos con respawn
    Call aMundo.StepMinute
    
    If Minutos = Intervalos(eIntervalos.iWorldSave) - 1 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("WorldSave en 1 minuto...", FontTypeNames.FONTTYPE_SERVER))
    ElseIf Minutos >= Intervalos(eIntervalos.iWorldSave) Then
        Call modFileIO.DoBackUp
        Call aClon.VaciarColeccion
        Minutos = 0
    ElseIf Minutos >= Intervalos(eIntervalos.iWorldSave) - 5 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("WorldSave en " & Intervalos(eIntervalos.iWorldSave) - Minutos & " minutos...", FontTypeNames.FONTTYPE_SERVER))
    End If
    
    If MinsPjesSave = Intervalos(eIntervalos.iGuardarUsuarios) - 1 Then ' 0.13.3
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Guardado de personajes en 1 minuto...", FontTypeNames.FONTTYPE_SERVER))
    ElseIf MinsPjesSave >= Intervalos(eIntervalos.iGuardarUsuarios) Then
        Call modUsuariosParty.ActualizaExperiencias
        Call GuardarUsuarios
        MinsPjesSave = 0
    ElseIf MinsPjesSave >= Intervalos(eIntervalos.iGuardarUsuarios) - 5 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Guardado de personajes en " & Intervalos(eIntervalos.iGuardarUsuarios) - MinsPjesSave & " minutos...", FontTypeNames.FONTTYPE_SERVER))
    End If
    
    If MinutosLatsClean >= 15 Then
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
        Call LimpiarMundo
    Else
        MinutosLatsClean = MinutosLatsClean + 1
    End If
    
    If MinsSendMotd >= Intervalos(eIntervalos.iMinutosMotd) And Intervalos(eIntervalos.iMinutosMotd) <> 0 Then ' 0.13.5
        Dim i As Long
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                Call SendMOTD(i)
            End If
        Next i
        MinsSendMotd = 0
    End If
    
    Call PurgarPenas
    Call CheckIdleUser
    
    '<<<<<-------- Log the number of users online ------>>>
    Dim N As Integer
    N = FreeFile()
    Open App.Path & "\Logs\NumUsers.log" For Output Shared As N
    Print #N, NumUsers
    Close #N
    '<<<<<-------- Log the number of users online ------>>>
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en TimerAutoSave Nro." & Err.Number & ": " & Err.description)
    Resume Next
    
End Sub



Private Sub cmdPausa_Click()



End Sub

Private Sub CantUsuarios_Click()

End Sub

Private Sub chkServerHabilitado_Click()

    If chkServerHabilitado.Checked = False Then
        chkServerHabilitado.Checked = True
    Else
        chkServerHabilitado.Checked = False
    End If
    
    iniSoloGMs = val(chkServerHabilitado.Checked)

End Sub

Private Sub Command1_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
''''''''''''''''SOLO PARA EL TESTEO'''''''
''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
txtChat.Text = txtChat.Text & vbNewLine & "Pop-Up> " & BroadMsg.Text
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call PutSystray
Else
    frmMain.WindowState = vbNormal ' GSZ
    frmMain.Show
    frmMain.SetFocus
End If

End Sub

Private Sub Command18_Click()
    Call mnuGuardarUsers_Click
End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))
''''''''''''''''SOLO PARA EL TESTEO'''''''
''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
End Sub

Private Sub Command23_Click()

End Sub

Private Sub Command4_Click()
    Call mnuHacerWorldsave_Click
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Form_Load()
#If UsarQueSocket = 1 Then
    mnuReiniciarSock.Visible = True
    mnuReiniciarListen.Visible = True
#ElseIf UsarQueSocket = 0 Then
    mnuReiniciarSock.Visible = False
    mnuReiniciarListen.Visible = False
#ElseIf UsarQueSocket = 2 Then
    mnuReiniciarSock.Visible = True
    mnuReiniciarListen.Visible = False
#End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    If MsgBox("¡ATENCIÓN!" & vbCrLf & "El servidor se cerrará SIN GUARDAR los cambios." & vbCrLf & "¿Desea hacerlo de todas maneras?", vbExclamation + vbYesNo) = vbNo Then
        Cancel = True ' Detener el cerrado!
    End If
End If
End Sub

Private Sub Form_Resize()
If frmMain.WindowState = vbMinimized Then Call PutSystray ' GSZ

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

'Save stats!!!
Call modStatistics.DumpStatistics

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
Call LimpiaWsApi
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If

Dim LoopC As Integer

For LoopC = 1 To iniMaxUsuarios
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & time & " servidor cerrado."
Close #N

End

Set SonidosMapas = Nothing

End Sub

Private Sub tFXMapas_Timer()
On Error GoTo hayerror

    Call SonidosMapas.ReproducirSonidosDeMapas

Exit Sub
hayerror:

End Sub

Private Sub GameTimer_Timer()
'********************************************************
'Author: Unknownn
'Last Modification: 03/12/2012 - ^[GS]^
'
'********************************************************
    Dim iUserIndex As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS As Boolean
    
On Error GoTo hayerror

    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To LastUser  'iniMaxUsuarios ' Cambio de iniMaxUsuarios por lastuser (fedudok)

        With UserList(iUserIndex)
           'Conexion activa?
           If .ConnID <> -1 Then
                '¿User valido?
                
                If .ConnIDValida And .flags.UserLogged Then
                    
                    '[Alejo-18-5]
                    bEnviarStats = False
                    bEnviarAyS = False
                    
                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
                    If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                    
                    If .flags.Muerto = 0 Then
                        
                        '[Consejeros]
                        If (.flags.Privilegios And PlayerType.User) Then Call EfectoLava(iUserIndex)
                        
                        If .flags.Desnudo <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoFrio(iUserIndex)
                        
                        If .flags.Meditando Then Call DoMeditar(iUserIndex)
                        
                        If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoVeneno(iUserIndex)
                        
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                        End If
                        
                        If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                        
                        If .flags.AtacablePor <> 0 Then Call EfectoEstadoAtacable(iUserIndex)

                        Call DuracionPociones(iUserIndex)
                        
                        Call HambreYSed(iUserIndex, bEnviarAyS)
                        
                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                            If Lloviendo Then
                                If Not Intemperie(iUserIndex) Then
                                    If Not .flags.Descansar Then
                                    'No esta descansando
                                        Call Sanar(iUserIndex, bEnviarStats, Intervalos(eIntervalos.iSanarSinDescansar))
                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                        Call RecStamina(iUserIndex, bEnviarStats, Intervalos(eIntervalos.iStaminaSinDescansar))
                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                    Else
                                    'esta descansando
                                        Call Sanar(iUserIndex, bEnviarStats, Intervalos(eIntervalos.iSanarDescansando))
                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                        Call RecStamina(iUserIndex, bEnviarStats, Intervalos(eIntervalos.iStaminaDescansando))
                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                        'termina de descansar automaticamente
                                        If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                                            Call WriteRestOK(iUserIndex)
                                            Call WriteMensajes(iUserIndex, eMensajes.Mensaje002) '"Has terminado de descansar."
                                            .flags.Descansar = False
                                        End If
                                        
                                    End If
                                End If
                            Else
                                If Not .flags.Descansar Then
                                'No esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, Intervalos(eIntervalos.iSanarSinDescansar))
                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    Call RecStamina(iUserIndex, bEnviarStats, Intervalos(eIntervalos.iStaminaSinDescansar))
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    
                                Else
                                'esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, Intervalos(eIntervalos.iSanarDescansando))
                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    Call RecStamina(iUserIndex, bEnviarStats, Intervalos(eIntervalos.iStaminaDescansando))
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    'termina de descansar automaticamente
                                    If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                                        Call WriteRestOK(iUserIndex)
                                        Call WriteMensajes(iUserIndex, eMensajes.Mensaje002) '"Has terminado de descansar."
                                        .flags.Descansar = False
                                    End If
                                    
                                End If
                            End If
                        End If
                        
                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                        If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                    Else
                        If .flags.Traveling <> 0 Then Call TravelingEffect(iUserIndex)  ' 0.13.3
                    End If 'Muerto
                Else 'no esta logeado?
                    'Inactive players will be removed!
                    .Counters.IdleCount = .Counters.IdleCount + 1
                    If .Counters.IdleCount > Intervalos(eIntervalos.iCerrarConexionInactivo) Then
                        .Counters.IdleCount = 0
                        Call modTCP.CloseSocket(iUserIndex)
                    End If
                End If 'UserLogged
                
                'If there is anything to be sent, we send it
                Call FlushBuffer(iUserIndex)
            End If
        End With
    Next iUserIndex
Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.description & " UserIndex = " & iUserIndex)
End Sub

Private Sub mnuAdmins_Click()
    frmAdmin.Show
End Sub

Private Sub mnuApagarSinGuardar_Click()
If MsgBox("¡ATENCIÓN!" & vbCrLf & "El servidor se cerrará SIN GUARDAR los cambios." & vbCrLf & "¿Desea hacerlo de todas maneras?", vbExclamation + vbYesNo) = vbYes Then
    Dim f
    For Each f In Forms
        Unload f
    Next
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    Call LimpiaWsApi ' GSZAO, cerramos los sockets con seguridad...
End If
End Sub

Private Sub mnuApagarYGuardar_Click()
If MsgBox("¿Está seguro que desea hacer un WorldSave, guardar los personajes y apagar el servidor?", vbYesNo, "Apagar Magicamente") = vbYes Then
    Me.MousePointer = 11
    frmCargando.Show
    'WorldSave
    Call modFileIO.DoBackUp
    'commit experiencia
    Call modUsuariosParty.ActualizaExperiencias
    'Guardar Pjs
    Call GuardarUsuarios
    'Chauuu
    Unload frmMain
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    Call LimpiaWsApi ' GSZAO, cerramos los sockets con seguridad...
End If
End Sub

Private Sub mnuBalance_Click()
    
    If frmMain.Visible Then frmMain.txStatus.Text = "Cargando Balance de Razas y Clases."
    Call LoadBalance
    If frmMain.Visible Then frmMain.txStatus.Text = "Balance de Razas y Clases cargado."
    
End Sub

Private Sub mnuCerrar_Click()


End Sub

Private Sub mnuCargarWorldsave_Click()
If MsgBox("¿Está seguro que desea cargar el último backup del mundo?", vbYesNo, "Cargar último backup") = vbYes Then

    'Se asegura de que los sockets estan cerrados e ignora cualquier err
    On Error Resume Next
    
    If frmMain.Visible Then frmMain.txStatus.Text = "Reiniciando."
    
    frmCargando.Show
    
    If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\Connect.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
    If FileExist(App.Path & "\logs\Resurrecciones.log", vbNormal) Then Kill App.Path & "\logs\Resurrecciones.log"
    If FileExist(App.Path & "\logs\Teleports.Log", vbNormal) Then Kill App.Path & "\logs\Teleports.Log"
    
    
    #If UsarQueSocket = 1 Then
    Call apiclosesocket(SockListen)
    #ElseIf UsarQueSocket = 0 Then
    frmMain.Socket1.Cleanup
    frmMain.Socket2(0).Cleanup
    #ElseIf UsarQueSocket = 2 Then
    frmMain.Serv.Detener
    #End If
    
    Dim LoopC As Integer
    
    For LoopC = 1 To iniMaxUsuarios
        Call CloseSocket(LoopC)
    Next
      
    
    LastUser = 0
    NumUsers = 0
    
    Call FreeNPCs
    Call FreeCharIndexes
    
    Call LoadSini
    Call CargarBackUp
    Call LoadOBJData
    
    #If UsarQueSocket = 1 Then
    SockListen = ListenForConnect(iniPuerto, hWndMsg, "")
    
    #ElseIf UsarQueSocket = 0 Then
    frmMain.Socket1.AddressFamily = AF_INET
    frmMain.Socket1.Protocol = IPPROTO_IP
    frmMain.Socket1.SocketType = SOCK_STREAM
    frmMain.Socket1.Binary = False
    frmMain.Socket1.Blocking = False
    frmMain.Socket1.BufferSize = 1024
    
    frmMain.Socket2(0).AddressFamily = AF_INET
    frmMain.Socket2(0).Protocol = IPPROTO_IP
    frmMain.Socket2(0).SocketType = SOCK_STREAM
    frmMain.Socket2(0).Blocking = False
    frmMain.Socket2(0).BufferSize = 2048
    
    'Escucha
    frmMain.Socket1.LocalPort = iniPuerto
    frmMain.Socket1.listen
    #End If
    
    If frmMain.Visible Then frmMain.txStatus.Text = "Escuchando conexiones entrantes ..."
End If
End Sub

Private Sub mnuDebugNPC_Click()
    frmDebugNpc.Show
End Sub

Private Sub mnuDebugUser_Click()
    frmUserList.Show
End Sub

Private Sub mnuDump_Click()
On Error Resume Next

    Dim i As Integer
    For i = 1 To iniMaxUsuarios
        Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & _
            ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & _
            " UserLogged: " & UserList(i).flags.UserLogged)
    Next i
    
    Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub mnuGuardarUsers_Click()

    Me.MousePointer = 11
    Call modUsuariosParty.ActualizaExperiencias
    Call GuardarUsuarios
    Me.MousePointer = 0
    frmMain.txStatus.Text = "Guardado de Personajes completo!"
    
End Sub

Private Sub mnuHacerWorldsave_Click()
On Error GoTo eh

    Me.MousePointer = 11
    frmCargando.Show
    Call modFileIO.DoBackUp
    Me.MousePointer = 0
    frmMain.txStatus.Text = "WorldSave completo!"
    
Exit Sub
eh:
    Call LogError("Error en WORLDSAVE")
End Sub

Private Sub mnuHechizos_Click()

    If frmMain.Visible Then frmMain.txStatus.Text = "Cargando Listado de Hechizos."
    Call CargarHechizos
    If frmMain.Visible Then frmMain.txStatus.Text = "Listado de Hechizos cargado."
    
End Sub



Private Sub mnuMOTD_Click()
        
    If frmMain.Visible Then frmMain.txStatus.Text = "Cargando Mensaje del Día."
    Call LoadMotd
    If frmMain.Visible Then frmMain.txStatus.Text = "Mensaje del Día cargado."
    
End Sub

Private Sub mnuNombresInvalid_Click()

    If frmMain.Visible Then frmMain.txStatus.Text = "Cargando Listado de Nombres Invalidos."
    Call CargarForbidenWords
    If frmMain.Visible Then frmMain.txStatus.Text = "Listado de Nombres Invalidos cargado."
    
End Sub

Private Sub mnuNPCs_Click()

    If frmMain.Visible Then frmMain.txStatus.Text = "Cargando Listado de NPC's."
    Call CargaNpcsDat
    If frmMain.Visible Then frmMain.txStatus.Text = "Listado de NPC's cargado."
End Sub

Private Sub mnuObjetos_Click()

    If frmMain.Visible Then frmMain.txStatus.Text = "Cargando Listado de Objetos."
    Call ResetForums
    Call LoadOBJData
    If frmMain.Visible Then frmMain.txStatus.Text = "Listado de Objetos cargado."
    
End Sub

Private Sub mnuPausar_Click()
If EnPausa = False Then
    EnPausa = True
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    mnuPausar.Caption = "&Reanudar Servidor"
    frmMain.txStatus.Text = "Servidor Pausado!"
Else
    EnPausa = False
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    mnuPausar.Caption = "&Pausar Servidor"
    frmMain.txStatus.Text = "Servidor Reanudado!"
End If
End Sub

Private Sub mnuRecargarAdministradores_Click()
    
    If frmMain.Visible Then frmMain.txStatus.Text = "Recargando Administradores."
    Call LoadAdministrativeUsers
    If frmMain.Visible Then frmMain.txStatus.Text = "Administradores recargados."
    
End Sub

Private Sub mnuReiniciar_Click()
If MsgBox("¡ATENCIÓN!" & vbCrLf & "Si reinicia el servidor puede provocar la pérdida de datos de los usarios." & vbCrLf & "¿Desea reiniciar el servidor de todas maneras?", vbYesNo + vbCritical) = vbYes Then
    frmMain.txStatus.Text = "Reiniciando Servidor!"
    Call modGeneral.Restart
End If
End Sub

Private Sub mnuReiniciarListen_Click()
#If UsarQueSocket = 1 Then
    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(iniPuerto, hWndMsg, "")
#End If
End Sub

Private Sub mnuReiniciarRespawn_Click()
    Call ReSpawnOrigPosNpcs
End Sub

Private Sub mnuReiniciarSock_Click()
#If UsarQueSocket = 1 Then

If MsgBox("¿Está seguro que desea reiniciar los sockets? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
    Call WSApiReiniciarSockets
End If

#ElseIf UsarQueSocket = 2 Then

Dim LoopC As Integer

If MsgBox("¿Está seguro que desea reiniciar los sockets? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
    For LoopC = 1 To iniMaxUsuarios
        If UserList(LoopC).ConnID <> -1 And UserList(LoopC).ConnIDValida Then
            Call CloseSocket(LoopC)
        End If
    Next LoopC
    
    Call frmMain.Serv.Detener
    Call frmMain.Serv.Iniciar(iniPuerto)
End If

#End If
End Sub

Private Sub mnusalir_Click()
    Call mnuApagarSinGuardar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog_Timer()
On Error Resume Next

    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
    If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
    If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
    If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
        If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
    End If

End Sub



Private Sub mnuSalirGuardando_Click()
    Call mnuApagarYGuardar_Click
End Sub

Private Sub mnuServidorIni_Click()

    If frmMain.Visible Then frmMain.txStatus.Text = "Cargando Configuración del Servidor."
    Call LoadSini
    If frmMain.Visible Then frmMain.txStatus.Text = "Configuración del Servidor cargada."
    
End Sub

Private Sub mnuSlots_Click()

    frmConID.Show
    
End Sub

Public Sub PutSystray()

    Dim i As Integer
    Dim S As String
    Dim nid As NOTIFYICONDATA
    
    S = "GS-Z Argentum Online"
    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
    i = Shell_NotifyIconA(NIM_ADD, nid)
        
    If WindowState <> vbMinimized Then WindowState = vbMinimized
    Visible = False

End Sub



Private Sub mnuTrafico_Click()

    frmTrafic.Show
    
End Sub

Private Sub mnuUnbanIPs_Click()

    Dim i As Long, N As Long
    Dim sENtrada As String
    
    sENtrada = InputBox("Escribe ""estoy DE acuerdo"" sin comillas y con distinción de mayúsculas minúsculas para desbanear a todos los personajes", "UnBan", "hola")
    If sENtrada = "estoy DE acuerdo" Then
        
        N = BanIPs.Count
        For i = 1 To BanIPs.Count
            BanIPs.Remove 1
        Next i
        
        MsgBox "Se han habilitado " & N & " ipes"
    End If

End Sub

Private Sub mnuUnbanUsers_Click()
On Error Resume Next

    Dim Fn As String
    Dim cad$
    Dim N As Integer, k As Integer
    
    Dim sENtrada As String
    
    sENtrada = InputBox("Escribe ""estoy DE acuerdo"" entre comillas y con distinción de mayúsculas minúsculas para desbanear a todos los personajes.", "UnBan", "hola")
    If sENtrada = "estoy DE acuerdo" Then
    
        Fn = App.Path & "\logs\GenteBanned.log"
        
        If FileExist(Fn, vbNormal) Then
            N = FreeFile
            Open Fn For Input Shared As #N
            Do While Not EOF(N)
                k = k + 1
                Input #N, cad$
                Call UnBan(cad$)
                
            Loop
            Close #N
            MsgBox "Se han habilitado " & k & " personajes."
            Kill Fn
        End If
    End If
    
End Sub

Private Sub tNpcAtaca_Timer()
On Error Resume Next

    Dim npc As Long
    
    For npc = 1 To LastNPC
        Npclist(npc).CanAttack = 1
    Next npc

End Sub

Private Sub packetResend_Timer()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/01/07
'Attempts to resend to the user all data that may be enqueued.
'***************************************************
On Error GoTo ErrHandler:
    Dim i As Long
    
    For i = 1 To iniMaxUsuarios
        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.length > 0 Then
                Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.length))
            End If
        End If
    Next i

Exit Sub

ErrHandler:
    LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.description)
    Resume Next
End Sub

Private Sub tNpcAI_Timer()
'***************************************************
'Autor: Unknown
'Last Modification: 29/07/2012 - ^[GS]^
'***************************************************
On Error GoTo ErrorHandler

    Dim NpcIndex As Long
    Dim mapa As Integer
    Dim e_p As Integer
    
    If Not haciendoBK And Not EnPausa Then
        'Update NPCs
        For NpcIndex = 1 To LastNPC
            
            With Npclist(NpcIndex)
                If .flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
                    ' Chequea si contiua teniendo dueño
                    If .Owner > 0 Then Call ValidarPermanenciaNpc(NpcIndex)
                
                    If .flags.Paralizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)
                    Else
                        ' Preto? Tienen ai especial
                        If .NPCtype = eNPCType.Pretoriano Then
                            Call ClanPretoriano(.ClanIndex).PerformPretorianAI(NpcIndex)
                        Else
                            'Usamos AI si hay algun user en el mapa
                            If .flags.Inmovilizado = 1 Then
                               Call EfectoParalisisNpc(NpcIndex)
                            End If
                            
                            mapa = .Pos.Map
                            
                            If mapa > 0 Then
                                If MapInfo(mapa).NumUsers > 0 Then
                                    If .Movement <> TipoAI.ESTATICO Then
                                        Call NPCAI(NpcIndex)
                                    ElseIf iniAutoSacerdote Then ' GSZAO
                                        If .NPCtype = eNPCType.ResucitadorNewbie Or .NPCtype = eNPCType.Revividor Then
                                            Call NpcAutoSacerdote(NpcIndex)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next NpcIndex
    End If
    
    Exit Sub

Exit Sub

ErrorHandler:
    Call LogError("Error en tAI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.Map)
    Call MuereNpc(NpcIndex, 0)
End Sub

Private Sub TimerSeg_Timer()
'**************************************************************
'Author: ^[GS]^
'Last Modification: 31/03/2013 - ^[GS]^
'**************************************************************
On Error GoTo ErrorHandler

    Dim NpcIndex As Long
    If Not haciendoBK And Not EnPausa Then
        For NpcIndex = 1 To LastNPC
            With Npclist(NpcIndex)
                If .flags.NPCActive Then
                    If (.NPCtype = GuardiaReal Or .NPCtype = GuardiasCaos Or .NPCtype = GuardiasEspeciales) Then
                        If .Contadores.TiempoPersiguiendo > 0 Then  ' Esta persiguiendo a alguien!
                            .Contadores.TiempoPersiguiendo = .Contadores.TiempoPersiguiendo - 1
                            If .Contadores.TiempoPersiguiendo = 0 Then
                                .flags.AttackedBy = vbNullString ' ya deja de buscar al profugo y regresa a su posición!
                            End If
                        End If
                    End If
                End If
            End With
        Next NpcIndex
    End If
    
    If aLluviaDeOro = True Then
        If Intervalos(eIntervalos.iLluviaDeORO) <> 0 Then  ' GSZAO
            Static SegundosSinOro As Long
            If SegundosSinOro >= Intervalos(eIntervalos.iLluviaDeORO) Then
                Dim Oro As Long
                Dim mapa As Integer
                Dim MiObj As Obj
                Dim WhereDrop As WorldPos
                MiObj.ObjIndex = iORO
                Oro = RandomNumber(700, 10000)
                MiObj.Amount = Oro
                WhereDrop.Map = RandomNumber(1, NumMaps)
                Do While MiObj.Amount >= 700
                    WhereDrop.X = RandomNumber(20, 80)
                    WhereDrop.Y = RandomNumber(20, 80)
                    Call TirarItemAlPiso(WhereDrop, MiObj)
                    MiObj.Amount = MiObj.Amount - 500
                Loop
                SegundosSinOro = 0
            Else
                SegundosSinOro = SegundosSinOro + 1
            End If
        End If
    End If
       
Exit Sub

ErrorHandler:
    Call LogError("Error en TimerSeg_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.Map)
    Call MuereNpc(NpcIndex, 0)
    
End Sub

Private Sub tEfectoLluvia_Timer()
On Error GoTo ErrHandler

    If Intervalos(eIntervalos.iEfectoLluvia) = 0 Then Exit Sub
    
    If Lloviendo Then
        Dim iCount As Long
        For iCount = 1 To LastUser
            Call EfectoLluvia(iCount)
       Next iCount
    End If

Exit Sub
ErrHandler:
    Call LogError("tEfectoLluvia " & Err.Number & ": " & Err.description)
End Sub

Private Sub tLluvia_Timer()

On Error GoTo ErrorHandler

    If Intervalos(eIntervalos.iLluvia) = 0 Then Exit Sub
    
    Static MinutosLloviendo As Long
    Static MinutosSinLluvia As Long
    
    If Not Lloviendo Then
        MinutosSinLluvia = MinutosSinLluvia + 1
        If MinutosSinLluvia >= Intervalos(eIntervalos.iLluvia) And MinutosSinLluvia < 1440 Then
            If RandomNumber(1, 100) <= 2 Then
                Lloviendo = True
                MinutosSinLluvia = 0
                Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
            End If
        ElseIf MinutosSinLluvia >= 1440 Then ' minimo 1 vez por dia
            Lloviendo = True
            MinutosSinLluvia = 0
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        End If
    Else
        MinutosLloviendo = MinutosLloviendo + 1
        If MinutosLloviendo >= 5 Then
            Lloviendo = False
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
            MinutosLloviendo = 0
        Else
            If RandomNumber(1, 100) <= 2 Then
                Lloviendo = False
                MinutosLloviendo = 0
                Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
            End If
        End If
    End If
    
    If tEfectoLluvia.Enabled <> Lloviendo Then ' GSZAO - Era inutil tener este timer todo el tiempo enabled si no esta la lluvia activada...
        tEfectoLluvia.Enabled = Lloviendo
    End If

Exit Sub
ErrorHandler:
    Call LogError("Error tLluvia")

End Sub

Private Sub tPiqueteC_Timer()
On Error GoTo ErrHandler

    Dim NuevaA As Boolean
   ' Dim NuevoL As Boolean
    Dim GI As Integer
    Dim i As Long
    
    For i = 1 To LastUser
        With UserList(i)
            If .flags.UserLogged Then
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                    If .flags.Muerto = 0 Then
                        .Counters.PiqueteC = .Counters.PiqueteC + 1
                        Call WriteConsoleMsg(i, "¡¡¡Estás obstruyendo la vía pública, muévete o serás encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                        If .Counters.PiqueteC > 23 Then
                            .Counters.PiqueteC = 0
                            Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                        End If
                    Else
                        .Counters.PiqueteC = 0
                    End If
                Else
                    .Counters.PiqueteC = 0
                End If
                
                'ustedes se preguntaran que hace esto aca?
                'bueno la respuesta es simple: el codigo de AO es una mierda y encontrar
                'todos los puntos en los cuales la alineacion puede cambiar es un dolor de
                'huevos, asi que lo controlo aca, cada 6 segundos, lo cual es razonable
        
                GI = .GuildIndex
                If GI > 0 Then
                    NuevaA = False
                   ' NuevoL = False
                    '[Silver - Sacar alineaciones de Clanes]
                    'If Not modGuilds.m_ValidarPermanencia(i, True, NuevaA) Then
                    '    Call WriteMensajes(i, eMensajes.Mensaje004) '"Has sido expulsado del clan. ¡El clan ha sumado un punto de antifacción!"
                    'End If
                    If NuevaA Then
                        Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("¡El clan ha pasado a tener alineación " & GuildAlignment(GI) & "!", FontTypeNames.FONTTYPE_GUILD))
                        Call LogClanes("¡El clan cambio de alineación!")
                    End If
'                    If NuevoL Then
'                        Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("¡El clan tiene un nuevo líder!", FontTypeNames.FONTTYPE_GUILD))
'                        Call LogClanes("¡El clan tiene nuevo lider!")
'                    End If
                End If
                
                Call FlushBuffer(i)
            End If
        End With
    Next i
Exit Sub

ErrHandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.description)
End Sub





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then





#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub txtChat_Change()

End Sub

Private Sub txtIP_Click()
Call Clipboard.SetText(txtIP.Caption)
frmMain.txStatus.Text = "Dirección IP copiada."
End Sub
