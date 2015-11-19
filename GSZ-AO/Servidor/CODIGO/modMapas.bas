Attribute VB_Name = "modMapas"
Option Explicit

''
'Ruta base para los archivos de mapas
Public MapPath As String
Public MapBackupPath As String ' GSZAO
Public MapFlagName As String ' GSZAO

Public Sub LoadMapData()
'***************************************************
'Author: ^[GS]^
'Last Modification: 10/06/2013
'Inicializa las variables para la gestion de los mapas.
'***************************************************
On Error GoTo ErrorReport

    ' Numero total de mapas
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    
    ' Carga AreasStats.dat
    Call InitAreas
        
    ' Nombre de los mapas Mapas (Ejemplo Mapa1.map = Mapa)
    MapFlagName = GetVar(DatPath & "Map.dat", "INIT", "MapFlagName")
    
    ' Path de los Mapas
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    ' Path del Backup de los Mapas
    MapBackupPath = GetVar(DatPath & "Map.dat", "INIT", "MapBackupPath")
        
    ' ¿Existen los directorios?
    If (FileExist(MapPath, vbDirectory) = False) Then
        MsgBox "No existe el directorio de los mapas." & vbCrLf & MapPath, vbCritical + vbOKOnly
        End
    ElseIf (FileExist(MapBackupPath, vbDirectory) = False) Then
        MsgBox "No existe el directorio del backup de los mapas." & vbCrLf & MapBackupPath, vbCritical + vbOKOnly
        End
    End If
    
    ' Reformulamos el tamaño de las variables contenedoras
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
    
ErrorReport:
    Call LogError("LoadMap::Error cargando el Mapa " & Map & " en la Pos: " & X & "," & Y & "." & Err.description)

End Sub

Public Sub LoadMap(ByVal iMap As Long, Optional bBackup As Boolean = False)
'***************************************************
'Author: ^[GS]^
'Last Modification: 10/06/2013
'Carga un mapa en particular.
'***************************************************

'On Error GoTo ErrorReport

    Dim fileActive As String ' solo se utiliza para el reporte de error
    
    Dim fileMapMap As String
    Dim fileMapInf As String
    Dim fileMapDat As String

    If bBackup = True Then
        fileMapMap = App.Path & MapBackupPath & MapFlagName & iMap & ".map"
        fileMapInf = App.Path & MapBackupPath & MapFlagName & iMap & ".inf"
        fileMapDat = App.Path & MapBackupPath & MapFlagName & iMap & ".dat"
        If FileExist(fileMapMap, vbArchive) = False Then fileMapMap = App.Path & MapPath & MapFlagName & iMap & ".map"
        If FileExist(fileMapInf, vbArchive) = False Then fileMapInf = App.Path & MapPath & MapFlagName & iMap & ".inf"
        If FileExist(fileMapDat, vbArchive) = False Then fileMapDat = App.Path & MapPath & MapFlagName & iMap & ".dat"
    Else
        fileMapMap = App.Path & MapPath & MapFlagName & iMap & ".map"
        fileMapInf = App.Path & MapPath & MapFlagName & iMap & ".inf"
        fileMapDat = App.Path & MapPath & MapFlagName & iMap & ".dat"
    End If

    If FileExist(fileMapMap, vbArchive) = False Then ' GSZAO - El Mapa "no existe"!
        MapInfo(iMap).MapVersion = -1 ' marcar como invalido
        Exit Sub
    End If

    Dim hFile As Integer
    Dim X As Long
    Dim Y As Long
    Dim ByFlags As Byte
    Dim NPCfile As String
    Dim Leer As clsIniManager
    Dim MapReader As clsByteBuffer
    Dim InfReader As clsByteBuffer
    Dim Buff() As Byte
    
    Set MapReader = New clsByteBuffer
    Set InfReader = New clsByteBuffer
    Set Leer = New clsIniManager
    
    NPCfile = DatPath & "NPCs.dat"
    hFile = FreeFile

    'map
    fileActive = fileMapMap

    Open fileMapMap For Binary As #hFile
        Seek hFile, 1
        
        ReDim Buff(LOF(hFile) - 1) As Byte
    
        Get #hFile, , Buff
    Close hFile
    
    Call MapReader.initializeReader(Buff)

    'inf
    fileActive = fileMapInf
    
    Open fileMapInf For Binary As #hFile
        Seek hFile, 1

        ReDim Buff(LOF(hFile) - 1) As Byte
    
        Get #hFile, , Buff
    Close hFile
    
    Call InfReader.initializeReader(Buff)
    
    'map Header
    MapInfo(iMap).MapVersion = MapReader.getInteger
    
    MiCabecera.desc = MapReader.getString(Len(MiCabecera.desc))
    MiCabecera.crc = MapReader.getLong
    MiCabecera.MagicWord = MapReader.getLong
    
    Call MapReader.getDouble

    'inf Header
    Call InfReader.getDouble
    Call InfReader.getInteger

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(iMap, X, Y)
                '.map file
                ByFlags = MapReader.getByte

                If ByFlags And 1 Then .Blocked = 1

                .Graphic(1) = MapReader.getInteger

                'Layer 2 used?
                If ByFlags And 2 Then .Graphic(2) = MapReader.getInteger

                'Layer 3 used?
                If ByFlags And 4 Then .Graphic(3) = MapReader.getInteger

                'Layer 4 used?
                If ByFlags And 8 Then .Graphic(4) = MapReader.getInteger

                'Trigger used?
                If ByFlags And 16 Then .trigger = MapReader.getInteger

                '.inf file
                ByFlags = InfReader.getByte

                If ByFlags And 1 Then
                    .TileExit.Map = InfReader.getInteger
                    .TileExit.X = InfReader.getInteger
                    .TileExit.Y = InfReader.getInteger
                End If

                If ByFlags And 2 Then
                    'Get and make NPC
                     .NpcIndex = InfReader.getInteger

                    If .NpcIndex > 0 Then
                        'Si el npc debe hacer respawn en la pos
                        'original la guardamos
                        If val(GetVar(NPCfile, "NPC" & .NpcIndex, "PosOrig")) = 1 Then
                            .NpcIndex = OpenNPC(.NpcIndex)
                            Npclist(.NpcIndex).Orig.Map = iMap
                            Npclist(.NpcIndex).Orig.X = X
                            Npclist(.NpcIndex).Orig.Y = Y
                        Else
                            .NpcIndex = OpenNPC(.NpcIndex)
                        End If

                        Npclist(.NpcIndex).Pos.Map = iMap
                        Npclist(.NpcIndex).Pos.X = X
                        Npclist(.NpcIndex).Pos.Y = Y

                        Call MakeNPCChar(True, 0, .NpcIndex, iMap, X, Y)
                    End If
                End If

                If ByFlags And 4 Then
                    'Get and make Object
                    .ObjInfo.ObjIndex = InfReader.getInteger
                    .ObjInfo.Amount = InfReader.getInteger
                    If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otDestruible Then ' GSZAO
                        .ObjInfo.ExtraLong = ObjData(.ObjInfo.ObjIndex).MaxHp
                    Else
                        .ObjInfo.ExtraLong = 0
                    End If
                End If
            End With
        Next X
    Next Y
    
    fileActive = fileMapDat
    Call Leer.Initialize(fileMapDat)
    
    With MapInfo(Map)
        .Name = Leer.GetValue("Mapa" & iMap, "Name")
        .Music = Leer.GetValue("Mapa" & iMap, "MusicNum")
        .StartPos.Map = val(ReadField(1, Leer.GetValue("Mapa" & iMap, "StartPos"), 45))
        .StartPos.X = val(ReadField(2, Leer.GetValue("Mapa" & iMap, "StartPos"), 45))
        .StartPos.Y = val(ReadField(3, Leer.GetValue("Mapa" & iMap, "StartPos"), 45))
        
        ' 0.13.3
        .OnDeathGoTo.Map = val(ReadField(1, Leer.GetValue("Mapa" & iMap, "OnDeathGoTo"), Asc("-")))
        .OnDeathGoTo.X = val(ReadField(2, Leer.GetValue("Mapa" & iMap, "OnDeathGoTo"), Asc("-")))
        .OnDeathGoTo.Y = val(ReadField(3, Leer.GetValue("Mapa" & iMap, "OnDeathGoTo"), Asc("-")))
        
        .MagiaSinEfecto = val(Leer.GetValue("Mapa" & iMap, "MagiaSinEfecto"))
        .InviSinEfecto = val(Leer.GetValue("Mapa" & iMap, "InviSinEfecto"))
        .ResuSinEfecto = val(Leer.GetValue("Mapa" & iMap, "ResuSinEfecto"))
        
        ' 0.13.3
        .OcultarSinEfecto = val(Leer.GetValue("Mapa" & iMap, "OcultarSinEfecto"))
        .InvocarSinEfecto = val(Leer.GetValue("Mapa" & iMap, "InvocarSinEfecto"))
        
        ' .NoEncriptarMP = val(Leer.GetValue("Mapa" & Map, "NoEncriptarMP")) ' GSZAO - no se utiliza

        .RoboNpcsPermitido = val(Leer.GetValue("Mapa" & iMap, "RoboNpcsPermitido"))
        
        If val(Leer.GetValue("Mapa" & iMap, "Pk")) = 0 Then
            .Pk = True
        Else
            .Pk = False
        End If
        
        .Terreno = TerrainStringToByte(Leer.GetValue("Mapa" & iMap, "Terreno"))
        .Zona = Leer.GetValue("Mapa" & iMap, "Zona")
        .Restringir = RestrictStringToByte(Leer.GetValue("Mapa" & iMap, "Restringir"))
        .Backup = val(Leer.GetValue("Mapa" & iMap, "BACKUP"))
        
        ' WorldGrid
        Call UpdateGrid(iMap)
    End With
    
#If Testeo = 1 Then
    If MaxGrid > 0 Then ' Utiliza Grid
        Dim iX As Integer
        Dim iY As Integer
        For iX = 1 To 100
            For iY = 1 To 100
                If iX = 10 Then
                    MapData(iMap, iX, iY).ObjInfo.ObjIndex = 1
                    MapData(iMap, iX, iY).ObjInfo.Amount = 1
                End If
                If iX = 90 Then
                    MapData(iMap, iX, iY).ObjInfo.ObjIndex = 1
                    MapData(iMap, iX, iY).ObjInfo.Amount = 1
                End If
                If iY = 10 Then
                    MapData(iMap, iX, iY).ObjInfo.ObjIndex = 1
                    MapData(iMap, iX, iY).ObjInfo.Amount = 1
                End If
                If iY = 90 Then
                    MapData(iMap, iX, iY).ObjInfo.ObjIndex = 1
                    MapData(iMap, iX, iY).ObjInfo.Amount = 1
                End If
            Next
        Next
    End If
#End If
    
    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing
    
    Erase Buff
Exit Sub

ErrorReport:
    Call LogError("LoadMap::ERROR " & Err.Number & " (" & Err.description & ") en Mapa " & iMap & " posición: " & X & "," & Y & " durante la carga de " & fileActive & ".")

    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing

End Sub

Public Sub LoadAllMaps(Optional bBackup As Boolean = False)
'***************************************************
'Author: ^[GS]^
'Last Modification: 10/06/2013
'Cargar todos los mapas.
'***************************************************

    Call LoadMapData
End Sub

Public Sub SaveMapBackup(ByVal iMap As Long)
'***************************************************
'Author: ^[GS]^
'Last Modification: 10/06/2013
'Guarda un mapa en particular.
'***************************************************

'On Error GoTo ErrorReport

    Dim fileActive As String ' solo se utiliza para el reporte de error
    
    Dim fileMapMap As String
    Dim fileMapInf As String
    Dim fileMapDat As String

    fileMapMap = App.Path & MapBackupPath & MapFlagName & iMap & ".map"
    fileMapInf = App.Path & MapBackupPath & MapFlagName & iMap & ".inf"
    fileMapDat = App.Path & MapBackupPath & MapFlagName & iMap & ".dat"
    
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim LoopC As Long
    
    ' 0.13.3
    Dim NpcInvalido As Boolean
    Dim MapWriter As clsByteBuffer
    Dim InfWriter As clsByteBuffer
    Dim IniManager As clsIniManager
    Set MapWriter = New clsByteBuffer
    Set InfWriter = New clsByteBuffer
    Set IniManager = New clsIniManager
    
    If FileExist(fileMapMap, vbNormal) Then
        Kill fileMapMap
    End If
    
    If FileExist(fileMapInf, vbNormal) Then
        Kill fileMapInf
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open fileMapMap For Binary As FreeFileMap
    
    Call MapWriter.initializeWriter(FreeFileMap)
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open fileMapInf For Binary As FreeFileInf
    
    Call InfWriter.initializeWriter(FreeFileInf)
    
    'map Header
    Call MapWriter.putInteger(MapInfo(iMap).MapVersion)
        
    Call MapWriter.putString(MiCabecera.desc, False)
    Call MapWriter.putLong(MiCabecera.crc)
    Call MapWriter.putLong(MiCabecera.MagicWord)
    
    Call MapWriter.putDouble(0)
    
    'inf Header
    Call InfWriter.putDouble(0)
    Call InfWriter.putInteger(0)
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(iMap, X, Y)
                ByFlags = 0
                
                If .Blocked Then ByFlags = ByFlags Or 1
                If .Graphic(2) Then ByFlags = ByFlags Or 2
                If .Graphic(3) Then ByFlags = ByFlags Or 4
                If .Graphic(4) Then ByFlags = ByFlags Or 8
                If .trigger Then ByFlags = ByFlags Or 16
                
                Call MapWriter.putByte(ByFlags)
                
                Call MapWriter.putInteger(.Graphic(1))
                
                For LoopC = 2 To 4
                    If .Graphic(LoopC) Then Call MapWriter.putInteger(.Graphic(LoopC))
                Next LoopC
                
                If .trigger Then Call MapWriter.putInteger(CInt(.trigger))
                
                '.inf file
                ByFlags = 0
                
                If .ObjInfo.ObjIndex > 0 Then
                   If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        .ObjInfo.ObjIndex = 0
                        .ObjInfo.Amount = 0
                    End If
                End If
    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                ' No hacer backup de los NPCs inválidos (Pretorianos, Mascotas, Invocados y Centinela)
                If .NpcIndex Then
                    NpcInvalido = (Npclist(.NpcIndex).NPCtype = eNPCType.Pretoriano) Or (Npclist(.NpcIndex).MaestroUser > 0) Or EsCentinela(.NpcIndex)
                    
                    If Not NpcInvalido Then ByFlags = ByFlags Or 2
                End If
                
                If .ObjInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Call InfWriter.putByte(ByFlags)
                
                If .TileExit.Map Then
                    Call InfWriter.putInteger(.TileExit.Map)
                    Call InfWriter.putInteger(.TileExit.X)
                    Call InfWriter.putInteger(.TileExit.Y)
                End If
                
                ' 0.13.3
                If .NpcIndex And Not NpcInvalido Then _
                    Call InfWriter.putInteger(Npclist(.NpcIndex).Numero)
                
                If .ObjInfo.ObjIndex Then
                    Call InfWriter.putInteger(.ObjInfo.ObjIndex)
                    Call InfWriter.putInteger(.ObjInfo.Amount)
                End If
                
                NpcInvalido = False
            End With
        Next X
    Next Y
    
    Call MapWriter.saveBuffer
    Call InfWriter.saveBuffer
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf
    
    Set MapWriter = Nothing
    Set InfWriter = Nothing

    With MapInfo(Map)
        'write .dat file
        Call IniManager.ChangeValue("Mapa" & iMap, "Name", .Name)
        Call IniManager.ChangeValue("Mapa" & iMap, "MusicNum", .Music)
        Call IniManager.ChangeValue("Mapa" & iMap, "MagiaSinefecto", .MagiaSinEfecto)
        Call IniManager.ChangeValue("Mapa" & iMap, "InviSinEfecto", .InviSinEfecto)
        Call IniManager.ChangeValue("Mapa" & iMap, "ResuSinEfecto", .ResuSinEfecto)
        Call IniManager.ChangeValue("Mapa" & iMap, "StartPos", .StartPos.Map & "-" & .StartPos.X & "-" & .StartPos.Y)
        Call IniManager.ChangeValue("Mapa" & iMap, "OnDeathGoTo", .OnDeathGoTo.Map & "-" & .OnDeathGoTo.X & "-" & .OnDeathGoTo.Y) ' 0.13.3
    
        Call IniManager.ChangeValue("Mapa" & iMap, "Terreno", TerrainByteToString(.Terreno))
        Call IniManager.ChangeValue("Mapa" & iMap, "Zona", .Zona)
        Call IniManager.ChangeValue("Mapa" & iMap, "Restringir", RestrictByteToString(.Restringir))
        Call IniManager.ChangeValue("Mapa" & iMap, "BackUp", str$(.Backup))
    
        If .Pk Then
            Call IniManager.ChangeValue("Mapa" & iMap, "Pk", "0")
        Else
            Call IniManager.ChangeValue("Mapa" & iMap, "Pk", "1")
        End If
        
        Call IniManager.ChangeValue("Mapa" & iMap, "OcultarSinEfecto", .OcultarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & iMap, "InvocarSinEfecto", .InvocarSinEfecto)
        ' 0.13.3
        Call IniManager.ChangeValue("Mapa" & iMap, "NoEncriptarMP", .NoEncriptarMP)
        Call IniManager.ChangeValue("Mapa" & iMap, "RoboNpcsPermitido", .RoboNpcsPermitido)
    
        Call IniManager.DumpFile(fileMapDat)

    End With
    
    Set IniManager = Nothing

Exit Sub

ErrorReport:
    Call LogError("SaveMap::ERROR " & Err.Number & " (" & Err.description & ") en Mapa " & Map & " posición: " & X & "," & Y & " durante el guardado de " & fileActive & ".")

    Set MapWriter = Nothing
    Set InfWriter = Nothing
    Set IniManager = Nothing

End Sub

Public Sub SaveAllMapsBackup()
'***************************************************
'Author: ^[GS]^
'Last Modification: 10/06/2013
'Guarda el backup de todos los mapas especificados.
'***************************************************

    Call LoadMapData
End Sub
