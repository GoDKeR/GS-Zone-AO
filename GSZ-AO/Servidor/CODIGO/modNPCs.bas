Attribute VB_Name = "modNPCs"
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


'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Option Explicit

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    
    For i = 1 To MAXMASCOTAS
      If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
         UserList(UserIndex).MascotasIndex(i) = 0
         UserList(UserIndex).MascotasType(i) = 0
         
         UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
         Exit For
      End If
    Next i
End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
End Sub

Public Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'********************************************************
'Author: Unknownn
'Llamado cuando la vida de un NPC llega a cero.
'Last Modification: 13/08/2014 - ^[GS]^
'********************************************************
On Error GoTo ErrHandler
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Dim EraCriminal As Boolean
    Dim PretorianoIndex As Integer
   
   ' Es pretoriano?
    If MiNPC.NPCtype = eNPCType.Pretoriano Then
        Call ClanPretoriano(MiNPC.ClanIndex).MuerePretoriano(NpcIndex) ' 0.13.3
    End If
   
    'Quitamos el npc
    Call QuitarNPC(NpcIndex)
    
    If UserIndex > 0 Then ' Lo mato un usuario?
        With UserList(UserIndex)
        
            If MiNPC.flags.Snd3 > 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.Pos.X, MiNPC.Pos.Y))
            End If
            .flags.TargetNPC = 0
            .flags.TargetNpcTipo = eNPCType.Comun
            
            'El user que lo mato tiene mascotas?
            If .NroMascotas > 0 Then
                Dim T As Integer
                For T = 1 To MAXMASCOTAS
                      If .MascotasIndex(T) > 0 Then
                          If Npclist(.MascotasIndex(T)).TargetNPC = NpcIndex Then
                                  Call FollowAmo(.MascotasIndex(T))
                          End If
                      End If
                Next T
            End If
            
            '[KEVIN]
            If MiNPC.flags.ExpCount > 0 Then
                If .PartyIndex > 0 Then
                    Call modUsuariosParty.ObtenerExito(UserIndex, MiNPC.flags.ExpCount, MiNPC.Pos.Map, MiNPC.Pos.X, MiNPC.Pos.Y)
                Else
                    .Stats.Exp = .Stats.Exp + MiNPC.flags.ExpCount
                    If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
                    Call WriteConsoleMsg(UserIndex, "Has ganado " & MiNPC.flags.ExpCount & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
                End If
                MiNPC.flags.ExpCount = 0
            End If
            
            '[/KEVIN]
            Call WriteMensajes(UserIndex, eMensajes.Mensaje151) '"¡Has matado a la criatura!"
            If .Stats.NPCsMuertos < 32000 Then .Stats.NPCsMuertos = .Stats.NPCsMuertos + 1
            
            EraCriminal = Criminal(UserIndex)
            
            If MiNPC.Stats.Alineacion = 0 Then
            
                If MiNPC.Numero = Guardias Then
                    .Reputacion.NobleRep = 0
                    .Reputacion.PlebeRep = 0
                    .Reputacion.AsesinoRep = .Reputacion.AsesinoRep + 500
                    If .Reputacion.AsesinoRep > MAXREP Then .Reputacion.AsesinoRep = MAXREP
                End If
                
                If MiNPC.MaestroUser = 0 Then
                    .Reputacion.AsesinoRep = .Reputacion.AsesinoRep + vlASESINO
                    If .Reputacion.AsesinoRep > MAXREP Then .Reputacion.AsesinoRep = MAXREP
                End If
                
            ElseIf Not esCaos(UserIndex) Then
                If MiNPC.Stats.Alineacion = 1 Then
                    .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlCAZADOR
                    If .Reputacion.PlebeRep > MAXREP Then _
                        .Reputacion.PlebeRep = MAXREP
                        
                ElseIf MiNPC.Stats.Alineacion = 2 Then
                    .Reputacion.NobleRep = .Reputacion.NobleRep + vlASESINO / 2
                    If .Reputacion.NobleRep > MAXREP Then _
                        .Reputacion.NobleRep = MAXREP
                        
                ElseIf MiNPC.Stats.Alineacion = 4 Then
                    .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlCAZADOR
                    If .Reputacion.PlebeRep > MAXREP Then _
                        .Reputacion.PlebeRep = MAXREP
                        
                End If
                    
            End If
            
            Dim EsCriminal As Boolean
            EsCriminal = Criminal(UserIndex)
            
            ' Cambio de alienacion?
            If EraCriminal <> EsCriminal Then
                
                ' Se volvio pk?
                If EsCriminal Then
                    If esArmada(UserIndex) Then Call ExpulsarFaccionReal(UserIndex)
                
                ' Se volvio ciuda
                Else
                    If esCaos(UserIndex) Then Call ExpulsarFaccionCaos(UserIndex)
                End If
                
                Call RefreshCharStatus(UserIndex)
            End If
            
            Call CheckUserLevel(UserIndex)
            
            If NpcIndex = .flags.ParalizedByNpcIndex Then
                Call RemoveParalisis(UserIndex) ' 0.13.3
            End If
            
        End With
    End If ' Userindex > 0
   
    If MiNPC.MaestroUser = 0 Then
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(MiNPC, MiNPC.NPCtype = eNPCType.Pretoriano)
        'ReSpawn o no
        Call ReSpawnNpc(MiNPC)
    End If
   
    ' GSZAO - Sistema de Quests
    Dim i As Integer
    Dim j As Integer
    For i = 1 To MAXUSERQUESTS
        With UserList(UserIndex).QuestStats.Quests(i)
            If .QuestIndex Then
                If QuestList(.QuestIndex).RequiredNPCs Then
                    For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                        If QuestList(.QuestIndex).RequiredNPC(j).NpcIndex = MiNPC.Numero Then
                            If QuestList(.QuestIndex).RequiredNPC(j).Amount > .NPCsKilled(j) Then
                                .NPCsKilled(j) = .NPCsKilled(j) + 1
                            End If
                        End If
                    Next j
                End If
            End If
        End With
    Next i
    
Exit Sub

ErrHandler:
    Call LogError("Error en MuereNpc - Error: " & Err.Number & " - Desc: " & Err.description)
End Sub

Private Sub ResetNpcFlags(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: 20/03/2013 - ^[GS]^
'***************************************************

    'Clear the npc's flags
    
    With Npclist(NpcIndex).flags
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = vbNullString
        .AttackedFirstBy = vbNullString
        .Backup = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .fAccion = 0
        .Follow = False
        .AtacaDoble = 0
        .LanzaSpells = 0
        .Invisible = 0
        .Maldicion = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
        .UsuariosMatados = 0 ' GSZAO
    End With
End Sub

Private Sub ResetNpcCounters(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: 07/04/2012 - ^[GS]^
'***************************************************

    With Npclist(NpcIndex).Contadores
        .Paralisis = 0
        .TiempoExistencia = 0
        .TiempoPersiguiendo = 0 ' GSZAO
    End With
End Sub

Private Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: -
'***************************************************

    With Npclist(NpcIndex).Char
        .Body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Private Sub ResetNpcCriatures(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: -
'***************************************************

    Dim j As Long
    
    With Npclist(NpcIndex)
        For j = 1 To .NroCriaturas
            .Criaturas(j).NpcIndex = 0
            .Criaturas(j).NpcName = vbNullString
        Next j
        
        .NroCriaturas = 0
    End With
End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim j As Long
    
    With Npclist(NpcIndex)
        For j = 1 To .NroExpresiones
            .Expresiones(j) = vbNullString
        Next j
        
        .NroExpresiones = 0
    End With
End Sub

Private Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: 12/08/2014 - ^[GS]^
'***************************************************

    With Npclist(NpcIndex)
        .Attackable = 0
        .CanAttack = 0
        .Comercia = 0
        .GiveEXP = 0
        .GiveGLD = 0
        .Hostile = 0
        .InvReSpawn = 0
        .QuestNumber = 0 ' GSZAO
        
        If .MaestroUser > 0 Then Call QuitarMascota(.MaestroUser, NpcIndex)
        If .MaestroNpc > 0 Then Call QuitarMascotaNpc(.MaestroNpc)
        If .Owner > 0 Then Call PerdioNpc(.Owner) ' 0.13.3
        
        .MaestroUser = 0
        .MaestroNpc = 0
        .Owner = 0
        
        .Mascotas = 0
        .Movement = 0
        .Name = vbNullString
        .NPCtype = 0
        .Numero = 0
        .Orig.Map = 0
        .Orig.X = 0
        .Orig.Y = 0
        .PoderAtaque = 0
        .PoderEvasion = 0
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .SkillDomar = 0
        .Target = 0
        .TargetNPC = 0
        .TipoItems = 0
        .Veneno = 0
        .desc = vbNullString
        
        .ClanIndex = 0
        
        Dim j As Long
        For j = 1 To .NroSpells
            .Spells(j) = 0
        Next j
    End With
    
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)
End Sub

Public Sub QuitarNPC(ByVal NpcIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 10/08/2011 - ^[GS]^
'16/11/2009: ZaMa - Now NPCs lose their owner
'***************************************************
On Error GoTo ErrHandler

    With Npclist(NpcIndex)
        .flags.NPCActive = False
        
        If InMapBounds(.Pos.Map, .Pos.X, .Pos.Y) Then
            Call EraseNPCChar(NpcIndex)
        End If
    End With
        
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then
        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1
            If LastNPC < 1 Then Exit Do
        Loop
    End If
        
      
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1
    End If
Exit Sub

ErrHandler:
    Call LogError("Error en QuitarNPC")
End Sub

Public Sub QuitarPet(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 18/11/2009
'Kills a pet
'***************************************************
On Error GoTo ErrHandler

    Dim i As Integer
    Dim PetIndex As Integer

    With UserList(UserIndex)
        
        ' Busco el indice de la mascota
        For i = 1 To MAXMASCOTAS
            If .MascotasIndex(i) = NpcIndex Then PetIndex = i
        Next i
        
        ' Poco probable que pase, pero por las dudas..
        If PetIndex = 0 Then Exit Sub
        
        ' Limpio el slot de la mascota
        .NroMascotas = .NroMascotas - 1
        .MascotasIndex(PetIndex) = 0
        .MascotasType(PetIndex) = 0
        
        ' Elimino la mascota
        Call QuitarNPC(NpcIndex)
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarPet. Error: " & Err.Number & " Desc: " & Err.description & " NpcIndex: " & NpcIndex & " UserIndex: " & UserIndex & " PetIndex: " & PetIndex)
End Sub

Private Function TestSpawnTrigger(Pos As WorldPos, Optional PuedeAgua As Boolean = False) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************
    
    If LegalPos(Pos.Map, Pos.X, Pos.Y, PuedeAgua) Then
        TestSpawnTrigger = MapData(Pos.Map, Pos.X, Pos.Y).trigger <> eTrigger.BAJOTECHOSINNPCS And MapData(Pos.Map, Pos.X, Pos.Y).trigger <> eTrigger.BAJOTECHO And MapData(Pos.Map, Pos.X, Pos.Y).trigger <> eTrigger.SINNPCS
    End If
    
End Function

Public Function CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos, _
                         Optional ByVal CustomHead As Integer) As Integer
'***************************************************
'Author: Unknownn
'Last Modification: 07/04/2012 - ^[GS]^
'***************************************************

'Crea un NPC del tipo NRONPC

Dim Pos As WorldPos
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long
Dim PuedeAgua As Boolean
Dim PuedeTierra As Boolean


Dim Map As Integer
Dim X As Integer
Dim Y As Integer

    nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    
    If nIndex > MAXNPCS Then Exit Function
    
    ' Cabeza customizada
    If CustomHead <> 0 Then Npclist(nIndex).Char.Head = CustomHead
     
    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
    
    'Necesita ser respawned en un lugar especifico
    If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
        Map = OrigPos.Map
        X = OrigPos.X
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).Pos = OrigPos
    Else
        Pos.Map = mapa 'mapa
        altpos.Map = mapa
        Do While Not PosicionValida
            Pos.X = RandomNumber(MinXBorder, MaxXBorder)    'Obtenemos posicion al azar en x
            Pos.Y = RandomNumber(MinYBorder, MaxYBorder)    'Obtenemos posicion al azar en y
            
            Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
            If newpos.X <> 0 And newpos.Y <> 0 Then
                altpos.X = newpos.X
                altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn, pero intentando qeu si tenía que ser en el agua, sea en el agua.)
            Else
                Call ClosestLegalPos(Pos, newpos, PuedeAgua)
                If newpos.X <> 0 And newpos.Y <> 0 Then
                    altpos.X = newpos.X
                    altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn)
                End If
            End If
            'Si X e Y son iguales a 0 significa que no se encontro posicion valida
            If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, PuedeAgua) And Not HayPCarea(newpos) And TestSpawnTrigger(newpos, PuedeAgua) Then
                'Asignamos las nuevas coordenas solo si son validas
                Npclist(nIndex).Pos.Map = newpos.Map
                Npclist(nIndex).Pos.X = newpos.X
                Npclist(nIndex).Pos.Y = newpos.Y
                PosicionValida = True
            Else
                newpos.X = 0
                newpos.Y = 0
            End If
            'for debug
            Iteraciones = Iteraciones + 1
            If Iteraciones > MAXSPAWNATTEMPS Then
                If altpos.X <> 0 And altpos.Y <> 0 Then
                    Map = altpos.Map
                    X = altpos.X
                    Y = altpos.Y
                    Npclist(nIndex).Pos.Map = Map
                    Npclist(nIndex).Pos.X = X
                    Npclist(nIndex).Pos.Y = Y
                    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
                    Exit Function
                Else
                    altpos.X = 50
                    altpos.Y = 50
                    Call ClosestLegalPos(altpos, newpos)
                    If newpos.X <> 0 And newpos.Y <> 0 Then
                        Npclist(nIndex).Pos.Map = newpos.Map
                        Npclist(nIndex).Pos.X = newpos.X
                        Npclist(nIndex).Pos.Y = newpos.Y
                        Call MakeNPCChar(True, newpos.Map, nIndex, newpos.Map, newpos.X, newpos.Y)
                        Exit Function
                    Else
                        Call QuitarNPC(nIndex)
                        Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNPC)
                        Exit Function
                    End If
                End If
            End If
        Loop
            
        'Asignamos las nuevas coordenas
        Map = newpos.Map
        X = Npclist(nIndex).Pos.X
        Y = Npclist(nIndex).Pos.Y
    End If
    
    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
    
    CrearNPC = nIndex
            
End Function

Public Sub MakeNPCChar(ByVal ToMap As Boolean, sndIndex As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: 24/07/2012 - ^[GS]^
'
'***************************************************
    
    Dim CharIndex As Integer

    If Npclist(NpcIndex).Char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(NpcIndex).Char.CharIndex = CharIndex
        CharList(CharIndex) = NpcIndex
    End If
    
    MapData(Map, X, Y).NpcIndex = NpcIndex
    
    If Not ToMap Then
        Dim bType As Byte ' GSZAO
        bType = Npclist(NpcIndex).Hostile ' 0 o 1
        If Npclist(NpcIndex).ShowName = True Or (iniNPCNoHostilesConNombre = True And bType = 0) Or (iniNPCHostilesConNombre = True And bType = 1) Then ' GSZAO NPC con nombre ^^
            Call WriteCharacterCreate(sndIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, Npclist(NpcIndex).Char.heading, Npclist(NpcIndex).Char.CharIndex, X, Y, Npclist(NpcIndex).Char.WeaponAnim, Npclist(NpcIndex).Char.ShieldAnim, 0, 0, Npclist(NpcIndex).Char.CascoAnim, Npclist(NpcIndex).Name, 0, 0, bType) ' GSZAO
        Else
            Call WriteCharacterCreate(sndIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, Npclist(NpcIndex).Char.heading, Npclist(NpcIndex).Char.CharIndex, X, Y, Npclist(NpcIndex).Char.WeaponAnim, Npclist(NpcIndex).Char.ShieldAnim, 0, 0, Npclist(NpcIndex).Char.CascoAnim, vbNullString, 0, 0, bType)
        End If
        Call FlushBuffer(sndIndex)
    Else
        Call AgregarNpc(NpcIndex)
    End If
    
    'GSZAO - Los guardias deberían regresar a su punto de control pasado X tiempo
    If Npclist(NpcIndex).NPCtype = GuardiaReal Or Npclist(NpcIndex).NPCtype = GuardiasCaos Then
        If Npclist(NpcIndex).Orig.Map = 0 Then ' solo forzarlo cuando no tienen seteada una pos inicial via DAT
            Npclist(NpcIndex).Orig.Map = Map
            Npclist(NpcIndex).Orig.X = X
            Npclist(NpcIndex).Orig.Y = Y
        End If
    End If
    
End Sub

Public Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal heading As eHeading)
'***************************************************
'Author: Unknownn
'Last Modification: 16/03/2012 - ^[GS]^
'
'***************************************************

    If NpcIndex > 0 Then
        With Npclist(NpcIndex).Char
            .Body = Body
            .Head = Head
            .heading = heading
            
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(Body, Head, heading, .CharIndex, Npclist(NpcIndex).Char.WeaponAnim, Npclist(NpcIndex).Char.ShieldAnim, 0, 0, Npclist(NpcIndex).Char.CascoAnim))
        End With
    End If
End Sub

Private Sub EraseNPCChar(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

If Npclist(NpcIndex).Char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

If Npclist(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar <= 1 Then Exit Do
    Loop
End If

'Quitamos del mapa
MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

'Actualizamos los clientes
Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterRemove(Npclist(NpcIndex).Char.CharIndex))

'Update la lista npc
Npclist(NpcIndex).Char.CharIndex = 0


'update NumChars
NumChars = NumChars - 1


End Sub

Public Function MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte) As Boolean
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 10/08/2011 - ^[GS]^
'06/04/2009: ZaMa - Now NPCs can force to change position with dead character
'01/08/2009: ZaMa - Now NPCs can't force to chance position with a dead character if that means to change the terrain the character is in
'26/09/2010: ZaMa - Turn sub into function to know if npc has moved or not.
'***************************************************

On Error GoTo errh
    Dim nPos As WorldPos
    Dim UserIndex As Integer
    
    With Npclist(NpcIndex)
        nPos = .Pos
        Call HeadtoPos(nHeading, nPos)
        
        ' es una posicion legal
        If LegalPosNPC(nPos.Map, nPos.X, nPos.Y, .flags.AguaValida = 1, .MaestroUser <> 0) Then
            
            If .flags.AguaValida = 0 And HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Function
            If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Function
            
            UserIndex = MapData(.Pos.Map, nPos.X, nPos.Y).UserIndex
            ' Si hay un usuario a donde se mueve el npc, entonces esta muerto
            If UserIndex > 0 Then
                
                ' No se traslada caspers de agua a tierra
                If HayAgua(.Pos.Map, nPos.X, nPos.Y) And Not HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Function
                ' No se traslada caspers de tierra a agua
                If Not HayAgua(.Pos.Map, nPos.X, nPos.Y) And HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Function
                
                With UserList(UserIndex)
                    ' Actualizamos posicion y mapa
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
                    .Pos.X = Npclist(NpcIndex).Pos.X
                    .Pos.Y = Npclist(NpcIndex).Pos.Y
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                        
                    ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, .Pos.X, .Pos.Y))
                    Call WriteForceCharMove(UserIndex, InvertHeading(nHeading))
                End With
            End If
            
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))

            'Update map and user pos
            MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex = 0
            .Pos = nPos
            .Char.heading = nHeading
            MapData(.Pos.Map, nPos.X, nPos.Y).NpcIndex = NpcIndex
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
            ' Npc has moved
            MoveNPCChar = True
        
        ElseIf .MaestroUser = 0 Then
            If .Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                .PFINFO.PathLenght = 0
            End If
        End If
    End With
    
    Exit Function

errh:
    LogError ("Error en move npc " & NpcIndex & ". Error: " & Err.Number & " - " & Err.description)

End Function

Function NextOpenNPC() As Integer
'***************************************************
'Author: Unknownn
'Last Modification: 14/05/2013 - ^[GS]^
'
'***************************************************

On Error GoTo ErrHandler
    Dim LoopC As Long
      
    For LoopC = 1 To MAXNPCS + 1
        If LoopC > MAXNPCS Then Exit For
        If Not Npclist(LoopC).flags.NPCActive Then Exit For
    Next LoopC
      
    NextOpenNPC = LoopC
Exit Function

ErrHandler:
    Call LogError("Error en NextOpenNPC - Nro." & Err.Number & ": " & Err.description)
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: 10/08/2011 - ^[GS]^
'
'***************************************************

    Dim N As Integer
    
    With UserList(UserIndex)
        If .flags.Muerto = 1 Then Exit Sub
        
        N = RandomNumber(1, 100)
        If N < 30 Then
            .flags.Envenenado = 1
            Call WriteMensajes(UserIndex, eMensajes.Mensaje152) '"¡¡La criatura te ha envenenado!!"
        End If
    End With

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 14/05/2013 - ^[GS]^
'
'***************************************************
On Error GoTo ErrHandler

    Dim newpos As WorldPos
    Dim altpos As WorldPos
    Dim nIndex As Integer
    Dim PosicionValida As Boolean
    Dim PuedeAgua As Boolean
    Dim PuedeTierra As Boolean
    
    
    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer
    
    nIndex = OpenNPC(NpcIndex, Respawn)     'Conseguimos un indice
    
    If nIndex > MAXNPCS Then
        SpawnNpc = 0
        Exit Function
    End If
    
    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = Not Npclist(nIndex).flags.TierraInvalida = 1
            
    Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
    Call ClosestLegalPos(Pos, altpos, PuedeAgua)
    'Si X e Y son iguales a 0 significa que no se encontro posicion valida
    
    If newpos.X <> 0 And newpos.Y <> 0 Then
        'Asignamos las nuevas coordenas solo si son validas
        Npclist(nIndex).Pos.Map = newpos.Map
        Npclist(nIndex).Pos.X = newpos.X
        Npclist(nIndex).Pos.Y = newpos.Y
        PosicionValida = True
    Else
        If altpos.X <> 0 And altpos.Y <> 0 Then
            Npclist(nIndex).Pos.Map = altpos.Map
            Npclist(nIndex).Pos.X = altpos.X
            Npclist(nIndex).Pos.Y = altpos.Y
            PosicionValida = True
        Else
            PosicionValida = False
        End If
    End If
    
    If Not PosicionValida Then
        Call QuitarNPC(nIndex)
        SpawnNpc = 0
        Exit Function
    End If
    
    'asignamos las nuevas coordenas
    Map = newpos.Map
    X = Npclist(nIndex).Pos.X
    Y = Npclist(nIndex).Pos.Y
    
    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
    
    If FX Then
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.CharIndex, FXIDs.FXWARP, 0))
    End If
    
    SpawnNpc = nIndex
    
Exit Function

ErrHandler:
    Call LogError("Error en SpawnNPC Nro." & Err.Number & ": " & Err.description & " (NPC " & NpcIndex & ")")

End Function

Sub ReSpawnNpc(MiNPC As npc)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)

End Sub

Private Sub NPCTirarOro(ByRef MiNPC As npc)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

'SI EL NPC TIENE ORO LO TIRAMOS
    If MiNPC.GiveGLD > 0 Then
        Dim MiObj As Obj
        Dim MiAux As Long
        MiAux = MiNPC.GiveGLD
        Do While MiAux > MAX_INVENTORY_OBJS
            MiObj.Amount = MAX_INVENTORY_OBJS
            MiObj.ObjIndex = iORO
            Call TirarItemAlPiso(MiNPC.Pos, MiObj)
            MiAux = MiAux - MAX_INVENTORY_OBJS
        Loop
        If MiAux > 0 Then
            MiObj.Amount = MiAux
            MiObj.ObjIndex = iORO
            Call TirarItemAlPiso(MiNPC.Pos, MiObj)
        End If
    End If
End Sub

Public Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer
'***************************************************
'Author: Unknownn
'Last Modification: 12/08/2014 - ^[GS]^
'
'***************************************************

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'conmigo. Para leer los NPCS se deberá usar la
'nueva clase clsIniManager.
'
'Alejo
'
'###################################################
    Dim NpcIndex As Integer
    Dim Leer As clsIniManager
    Dim LoopC As Long
    Dim ln As String
    
    Set Leer = LeerNPCs
    
    'If requested index is invalid, abort
    If Not Leer.KeyExists("NPC" & NpcNumber) Then
        OpenNPC = MAXNPCS + 1
        Exit Function
    End If
    
    NpcIndex = NextOpenNPC
    
    If NpcIndex > MAXNPCS Then 'Limite de npcs
        OpenNPC = NpcIndex
        Exit Function
    End If
    
    With Npclist(NpcIndex)
        .Numero = NpcNumber
        .Name = Leer.GetValue("NPC" & NpcNumber, "Name")
        .ShowName = val(Leer.GetValue("NPC" & NpcNumber, "ShowName")) ' GSZAO ¿Mostrar nombre?
        .desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
        
        .Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
        .flags.OldMovement = .Movement
        
        .flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
        .flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
        .flags.fAccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))
        .flags.AtacaDoble = val(Leer.GetValue("NPC" & NpcNumber, "AtacaDoble"))
        
        .NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
        
        ' GSZAO Es guardia especial!
        If .NPCtype = eNPCType.GuardiasEspeciales Then
            .AttackLvlLess = val(Leer.GetValue("NPC" & NpcNumber, "AttackLvlLess"))
            .AttackLvlMore = val(Leer.GetValue("NPC" & NpcNumber, "AttackLvlMore"))
        End If
        ' GSZAO
        
        .Char.Body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
        .Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
        .Char.heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
        
        ' GSZAO Tiene escudo, casco o arma?
        .Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "ShieldAnim"))
        .Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "CascoAnim"))
        .Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "WeaponAnim"))
        ' GSZAO
        
        .Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
        .Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
        .Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
        .flags.OldHostil = .Hostile
        
        .GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP")) * iniExp ' GSZAO
        If HappyHourActivated And (HappyHour <> 0) Then .GiveEXP = .GiveEXP * HappyHour ' 0.13.5
        
        .flags.ExpCount = .GiveEXP
        
        .Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
        
        .flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
        
        .GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD")) * iniOro ' GSZAO
        .QuestNumber = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber")) ' GSZAO
        
        .PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
        .PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
        
        .InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
        
        With .Stats
            .MaxHp = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
            .MinHp = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
            .MaxHIT = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
            .MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
            .def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
            .defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
            .Alineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))
        End With
        
        .Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
        For LoopC = 1 To .Invent.NroItems
            ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
            If (LenB(ln) <> 0) Then
                .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
                .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
                ' GSZAO
                .Invent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
                If .Invent.Object(LoopC).Equipped = 1 Then
                    If ObjData(.Invent.Object(LoopC).ObjIndex).OBJType = otESCUDO Then
                        .Char.ShieldAnim = ObjData(.Invent.Object(LoopC).ObjIndex).ShieldAnim
                    ElseIf ObjData(.Invent.Object(LoopC).ObjIndex).OBJType = otCASCO Then
                        .Char.CascoAnim = ObjData(.Invent.Object(LoopC).ObjIndex).CascoAnim
                    ElseIf ObjData(.Invent.Object(LoopC).ObjIndex).OBJType = otWeapon Then
                        .Char.WeaponAnim = ObjData(.Invent.Object(LoopC).ObjIndex).WeaponAnim
                    End If
                End If
                ' GSZAO
            End If
        Next LoopC
        
        For LoopC = 1 To MAX_NPC_DROPS
            ln = Leer.GetValue("NPC" & NpcNumber, "Drop" & LoopC)
            If (LenB(ln) <> 0) Then
                .Drop(LoopC).ObjIndex = val(ReadField(1, ln, 45))
                .Drop(LoopC).Amount = val(ReadField(2, ln, 45))
                ' GSZAO
                .Drop(LoopC).Equipped = val(ReadField(3, ln, 45))
                If .Drop(LoopC).Equipped = 1 Then
                    If ObjData(.Drop(LoopC).ObjIndex).OBJType = otESCUDO Then
                        .Char.ShieldAnim = ObjData(.Drop(LoopC).ObjIndex).ShieldAnim
                    ElseIf ObjData(.Drop(LoopC).ObjIndex).OBJType = otCASCO Then
                        .Char.CascoAnim = ObjData(.Drop(LoopC).ObjIndex).CascoAnim
                    ElseIf ObjData(.Drop(LoopC).ObjIndex).OBJType = otWeapon Then
                        .Char.WeaponAnim = ObjData(.Drop(LoopC).ObjIndex).WeaponAnim
                    End If
                End If
            End If
            ' GSZAO
        Next LoopC

        
        .flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
        If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)
        For LoopC = 1 To .flags.LanzaSpells
            .Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
        Next LoopC
        
        If .NPCtype = eNPCType.Entrenador Then
            .NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
            ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador
            For LoopC = 1 To .NroCriaturas
                .Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
                .Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
            Next LoopC
        End If
        
        With .flags
            .NPCActive = True
            
            If Respawn Then
                .Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
            Else
                .Respawn = 1
            End If
            
            .Backup = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
            .RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
            .AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
            
            .Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
            .Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
            .Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))
        End With
        
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        .NroExpresiones = val(Leer.GetValue("NPC" & NpcNumber, "NROEXP"))
        If .NroExpresiones > 0 Then ReDim .Expresiones(1 To .NroExpresiones) As String
        For LoopC = 1 To .NroExpresiones
            .Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
        Next LoopC
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        
        'Tipo de items con los que comercia
        .TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
        
        .Ciudad = val(Leer.GetValue("NPC" & NpcNumber, "Ciudad"))
    End With
    
    'Update contadores de NPCs
    If NpcIndex > LastNPC Then LastNPC = NpcIndex
    NumNPCs = NumNPCs + 1
    
    'Devuelve el nuevo Indice
    OpenNPC = NpcIndex
End Function

Public Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    With Npclist(NpcIndex)
        If .flags.Follow Then
            .flags.AttackedBy = vbNullString
            .flags.Follow = False
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
        Else
            .flags.AttackedBy = UserName
            .flags.Follow = True
            .Movement = TipoAI.NPCDEFENSA
            .Hostile = 0
        End If
    End With
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    With Npclist(NpcIndex)
        .flags.Follow = True
        .Movement = TipoAI.SigueAmo
        .Hostile = 0
        .Target = 0
        .TargetNPC = 0
    End With
End Sub

Public Sub ValidarPermanenciaNpc(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: -
'Chequea si el npc continua perteneciendo a algún usuario
'***************************************************

    With Npclist(NpcIndex)
        If IntervaloPerdioNpc(.Owner) Then Call PerdioNpc(.Owner)
    End With
End Sub

Sub NPCmataUser(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) ' GSZAO
'***************************************************
'Author: ^[GS]^
'Last Modification: 29/03/2013 - ^[GS]^
'***************************************************

    With Npclist(NpcIndex)
        .flags.UsuariosMatados = .flags.UsuariosMatados + 1
        
        If .flags.UsuariosMatados < 5 Then
            .Stats.def = .Stats.def + CInt(RandomNumber(1, UserList(UserIndex).Stats.ELV / 2))
            .Stats.MaxHIT = .Stats.MaxHIT + CInt(RandomNumber(0, UserList(UserIndex).Stats.ELV / 3))
        Else
            .Stats.def = .Stats.def + CInt(RandomNumber(1, UserList(UserIndex).Stats.ELV / 4))
            .Stats.MaxHIT = .Stats.MaxHIT + CInt(RandomNumber(1, UserList(UserIndex).Stats.ELV / 5))
        End If
    End With
    
End Sub
