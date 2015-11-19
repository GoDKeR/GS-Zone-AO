Attribute VB_Name = "modAdmin"
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

'INTERVALOS
Public Enum eIntervalos
    ' Generales
    iWorldSave = 1
    iGuardarUsuarios = 2
    iMinutosMotd = 3
    iLluvia = 4
    iCerrarConexion = 5
    iCerrarConexionInactivo = 6
    iEfectoLluvia = 7
    iReproducirFXMapas = 8
    iLluviaDeORO = 9
    ' Estado
    iSanarSinDescansar = 20
    iSanarDescansando = 21
    iStaminaSinDescansar = 22
    iStaminaDescansando = 23
    iSed = 24
    iHambre = 25
    iVeneno = 26
    iParalizado = 27
    iInvisible = 28
    iOculto = 29
    iFrio = 30
    iInvocacion = 31
    ' NPC's
    iNPCPuedeAtacar = 40
    iNPCPuedeUsarAI = 41
     ' Cliente
    iPuedeAtacar = 50
    iPuedeAtacarConFlechas = 51
    iPuedeAtacarConHechizos = 52
    iPuedeUsarItem = 53
    iPuedeUsarPocion = 54
    iPuedeTrabajar = 55
    iComboMagiaGolpe = 56
    iComboGolpeMagia = 57
End Enum
Public Intervalos(60) As Integer

' INTERVALOS FIXED's
Public IntervaloPuedeSerAtacado As Long
Public IntervaloAtacable As Long
Public IntervaloOwnedNpc As Long

Public Const IntervaloParalizadoReducido As Integer = 37 '0.13.3

Public Type tMotd
    texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type
Public Apuestas As tAPuestas

Public tInicioServer As Long

Public iniPuerto As Integer
Public iniWorldBackup As Byte

Public PorcentajeRecuperoMana As Integer

Public Lloviendo As Boolean
Public DeNoche As Boolean

Function VersionOK(ByVal Ver As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    VersionOK = (Ver = iniVersion)
End Function

Sub ReSpawnOrigPosNpcs()
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim i As Integer
    Dim MiNPC As npc
       
    For i = 1 To LastNPC
       'OJO
       If Npclist(i).flags.NPCActive Then
            
            If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                    MiNPC = Npclist(i)
                    Call QuitarNPC(i)
                    Call ReSpawnNpc(MiNPC)
            End If
            
            'tildada por sugerencia de yind
            'If Npclist(i).Contadores.TiempoExistencia > 0 Then
            '        Call MuereNpc(i, 0)
            'End If
       End If
       
    Next i
    
End Sub

Sub WorldSave()
'***************************************************
'Author: Unknownn
'Last Modification: 09/09/2012 - ^[GS]^
'***************************************************

On Error Resume Next

    Dim loopX As Integer
    Dim hFile As Integer
    
    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave", FontTypeNames.FONTTYPE_SERVER))
    
    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    
    Dim j As Integer, k As Integer
    
    For j = 1 To NumMaps
        If MapInfo(j).BackUp = 1 And MapInfo(loopX).MapVersion <> -1 Then k = k + 1
    Next j
    
    frmCargando.pCargar.min = 0
    frmCargando.pCargar.max = k
    frmCargando.pCargar.Value = 0
    
    For loopX = 1 To NumMaps
        'DoEvents
        If MapInfo(loopX).BackUp = 1 And MapInfo(loopX).MapVersion <> -1 Then
            ' GSZAO - Si el mapa tiene MapInfo(loopX).MapVersion = -1, es que esta fallado, por tanto "no se guarda"
            Call GrabarMapa(loopX, App.Path & "\WorldBackUp\Mapa" & loopX)
            frmCargando.pCargar.Value = frmCargando.pCargar.Value + 1
        End If
    
    Next loopX
    
    frmCargando.Visible = False
    
    If FileExist(DatPath & "\NPCs-Backup.dat") Then Kill (DatPath & "NPCs-Backup.dat")

    hFile = FreeFile()
    
    Open DatPath & "\NPCs-Backup.dat" For Output As hFile
        For loopX = 1 To LastNPC
            If Npclist(loopX).flags.BackUp = 1 Then
                Call BackUPnPc(loopX, hFile)
            End If
        Next loopX
    Close hFile
    
    Call SaveForums
    
    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluído.", FontTypeNames.FONTTYPE_SERVER))

End Sub

Public Sub PurgarPenas()
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call WriteMensajes(i, eMensajes.Mensaje016) '"¡Has sido liberado!"
                    
                    Call FlushBuffer(i)
                End If
            End If
        End If
    Next i
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = vbNullString)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    UserList(UserIndex).Counters.Pena = Minutos
    
    
    Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
    
    If LenB(GmName) = 0 Then
        Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    End If
    If UserList(UserIndex).flags.Traveling = 1 Then
        UserList(UserIndex).flags.Traveling = 0
        UserList(UserIndex).Counters.goHome = 0
        Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
    End If
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

On Error Resume Next
    If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
        Kill CharPath & UCase$(UserName) & ".chr"
    End If
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    'Unban the character
    Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")
    
    'Remove it from the banned people database
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")
End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    
    If MD5ClientesActivado = 1 Then
        For i = 0 To UBound(MD5s)
            If (md5formateado = MD5s(i)) Then
                MD5ok = True
                Exit Function
            End If
        Next i
        MD5ok = False
    Else
        MD5ok = True
    End If

End Function


Public Sub BanIpAgrega(ByVal ip As String)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    BanIPs.Add ip
    
    Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim Dale As Boolean
    Dim LoopC As Long
    
    Dale = True
    LoopC = 1
    Do While LoopC <= BanIPs.Count And Dale
        Dale = (BanIPs.Item(LoopC) <> ip)
        LoopC = LoopC + 1
    Loop
    
    If Dale Then
        BanIpBuscar = 0
    Else
        BanIpBuscar = LoopC - 1
    End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim N As Long
    
    N = BanIpBuscar(ip)
    If N > 0 Then
        BanIPs.Remove N
        BanIpGuardar
        BanIpQuita = True
    Else
        BanIpQuita = False
    End If

End Function

Public Sub BanIpGuardar()
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim ArchivoBanIp As String
    Dim ArchN As Long
    Dim LoopC As Long
    
    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN
    
    For LoopC = 1 To BanIPs.Count
        Print #ArchN, BanIPs.Item(LoopC)
    Next LoopC
    
    Close #ArchN

End Sub

Public Sub BanIpCargar()
'***************************************************
'Author: Unknownn
'Last Modification: 10/07/2012 - ^[GS]^
'
'***************************************************

    Dim ArchN As Long
    Dim Tmp As String
    Dim ArchivoBanIp As String
    
    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    Set BanIPs = New Collection
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Input As #ArchN
    
    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanIPs.Add Tmp
    Loop
    
    Close #ArchN

End Sub


Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType
'***************************************************
'Author: Unknownn
'Last Modification: 03/02/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'***************************************************

    If EsAdmin(Name) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSemiDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(Name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = PlayerType.User
    End If
End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/07/2012 - ^[GS]^
'
'***************************************************

    Dim tUser As Integer
    Dim userPriv As Byte
    Dim cantPenas As Byte
    Dim rank As Integer
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)
        If tUser <= 0 Then
            Call WriteMensajes(bannerUserIndex, eMensajes.Mensaje017) '"El usuario no está online."
            
            If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                userPriv = UserDarPrivilegioLevel(UserName)
                
                If (userPriv And rank) > (.flags.Privilegios And rank) Then
                    Call WriteMensajes(bannerUserIndex, eMensajes.Mensaje018) '"No puedes banear a al alguien de mayor jerarquía."
                Else
                    If GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban") <> "0" Then
                        Call WriteMensajes(bannerUserIndex, eMensajes.Mensaje019) '"El personaje ya se encuentra baneado."
                    Else
                        Call LogBanFromName(UserName, bannerUserIndex, Reason)
                        Call SendData(SendTarget.ToAdminsButCounselorsAndRms, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))
                        
                        'ponemos el flag de ban a 1
                        Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                        'ponemos la pena
                        cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)
                        
                        If (userPriv And rank) = (.flags.Privilegios And rank) Then
                            .flags.Ban = 1
                            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                            Call CloseSocket(bannerUserIndex)
                        End If
                        
                        Call LogGM(.Name, "BAN a " & UserName & ". Razón: " & Reason)
                    End If
                End If
            Else
                Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                Call WriteMensajes(bannerUserIndex, eMensajes.Mensaje018) '"No puedes banear a al alguien de mayor jerarquía."
            Else
            
                Call LogBan(tUser, bannerUserIndex, Reason)
                Call SendData(SendTarget.ToAdminsButCounselorsAndRms, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_SERVER))
                
                'Ponemos el flag de ban a 1
                UserList(tUser).flags.Ban = 1
                
                If (UserList(tUser).flags.Privilegios And rank) = (.flags.Privilegios And rank) Then
                    .flags.Ban = 1
                    Call SendData(SendTarget.ToAdminsButCounselorsAndRms, 0, PrepareMessageConsoleMsg(.Name & " ha sido baneado del servidor por un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                    Call CloseSocket(bannerUserIndex)
                End If
                
                Call LogGM(.Name, "BAN a " & UserName & ". Razón: " & Reason)
                
                'ponemos el flag de ban a 1
                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                'ponemos la pena
                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)
                
                Call CloseSocket(tUser)
            End If
        End If
    End With
End Sub


Public Function BanHD_rem(ByVal HD As String) As Boolean
' GSZ-AO - Remueve un SerialHD como baneado

   On Error Resume Next
 
    Dim N As Long
   
    N = BanHD_find(HD) ' buscar
    If N > 0 Then
        BanHDs.Remove N ' quitar
        BanHD_save ' guardar los cambios
        BanHD_rem = True
    Else
        BanHD_rem = False
    End If
   
End Function
Public Sub BanHD_add(ByVal HD As String)
' GSZ-AO - Agrega un nuevo SerialHD como baneado

    Dim N As Long
   
    N = BanHD_find(HD) ' buscar
    If N > 0 Then
        ' ya estaba
    Else
        BanHDs.Add HD ' agregar
        Call BanHD_save ' guardar los cambios
    End If
    
End Sub
Public Function BanHD_find(ByVal HD As String) As Long
' GSZ-AO - Busca si un SerialHD está baneado

    Dim Dale As Boolean
    Dim LoopC As Long
   
    Dale = True
    LoopC = 1
    Do While LoopC <= BanHDs.Count And Dale
        Dale = (BanHDs.Item(LoopC) <> HD)
        LoopC = LoopC + 1
    Loop
   
    If Dale Then
        BanHD_find = 0
    Else
        BanHD_find = LoopC - 1
    End If
    
End Function
Public Sub BanHD_save()
' GSZ-AO - Guarda el listado de SerialHD's baneados
On Error Resume Next
    Dim ArchivoBanHD As String
    Dim ArchN As Long
    Dim LoopC As Long
   
    ArchivoBanHD = App.Path & "\Dat\BanHDs.dat"
       
    ArchN = FreeFile()
    Open ArchivoBanHD For Output As #ArchN
   
    For LoopC = 1 To BanHDs.Count
        Print #ArchN, BanHDs.Item(LoopC)
    Next LoopC
   
    Close #ArchN
   
End Sub
Public Sub BanHD_load()
' GSZ-AO - Carga el listado de SerialHD's baneados
On Error Resume Next

    Dim ArchN As Long
    Dim Tmp As String
    Dim ArchivoBanHD As String
   
    ArchivoBanHD = App.Path & "\Dat\BanHDs.dat"
   
    Set BanHDs = New Collection
   
    ArchN = FreeFile()
    Open ArchivoBanHD For Input As #ArchN
   
    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanHDs.Add Tmp
    Loop
   
    Close #ArchN
End Sub
