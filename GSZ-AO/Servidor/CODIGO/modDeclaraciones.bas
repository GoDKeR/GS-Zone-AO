Attribute VB_Name = "modDeclaraciones"
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

''
' Modulo de modDeclaraciones. Aca hay de todo.
'
' GSZAO
Private Type tDado
    Minimo As Byte
    Base As Byte
    Random As Byte
End Type
Public Dados(4) As tDado

Public aClon As clsAntiMassClon
Public aLimpiarMundo As clsLimpiarMundo ' GSZAO
Public aMundo As clsMundo ' GSZAO

Public aLluviaDeOro As Boolean ' GSZAO

Public Const MAXSPAWNATTEMPS = 60
Public Const INFINITE_LOOPS As Integer = -1
Public Const FXSANGRE = 14

''
' The color of chats over head of dead characters.
Public Const CHAT_COLOR_DEAD_CHAR As Long = &HC0C0C0

''
' The color of yells made by any kind of game administrator.
Public Const CHAT_COLOR_GM_YELL As Long = &HF82FF

''
' Coordinates for normal sounds (not 3D, like rain)
Public Const NO_3D_SOUND As Byte = 0

Public Const iFragataFantasmal = 87
Public Const iFragataReal = 190
Public Const iFragataCaos = 189
Public Const iBarca = 84
Public Const iGalera = 85
Public Const iGaleon = 86

' Embarcaciones ciudas '0.13.3
Public Const iBarcaCiuda = 395
Public Const iBarcaCiudaAtacable = 552
Public Const iGaleraCiuda = 397
Public Const iGaleraCiudaAtacable = 560
Public Const iGaleonCiuda = 399
Public Const iGaleonCiudaAtacable = 556

' Embarcaciones reales '0.13.3
Public Const iBarcaReal = 550
Public Const iBarcaRealAtacable = 553
Public Const iGaleraReal = 558
Public Const iGaleraRealAtacable = 561
Public Const iGaleonReal = 554
Public Const iGaleonRealAtacable = 557

' Embarcaciones pk '0.13.3
Public Const iBarcaPk = 396
Public Const iGaleraPk = 398
Public Const iGaleonPk = 400

' Embarcaciones caos '0.13.3
Public Const iBarcaCaos = 551
Public Const iGaleraCaos = 559
Public Const iGaleonCaos = 555

Public Enum iMinerales
    HierroCrudo = 192
    PlataCruda = 193
    OroCrudo = 194
    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

Public Enum ePrivileges '0.13.3
    Admin = 1
    Dios
    Especial
    SemiDios
    Consejero
    RoleMaster
End Enum

Public Enum eClass
    Mage = 1       'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Thief       'Ladrón
    Bard        'Bardo
    Druid       'Druida
    Bandit      'Bandido
    Paladin     'Paladín
    Hunter      'Cazador
    Worker      'Trabajador
    Pirat       'Pirata
End Enum

Public Enum eRaza
    Humano = 1
    Elfo
    Drow
    Gnomo
    Enano
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum

Public Const LimiteNewbie As Byte = 12

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    crc As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

' 3/10/03
'Cambiado a 2 segundos el 30/11/07
Public Const TIEMPO_INICIO_MEDITAR As Integer = 2000
Public Const TIEMPO_SEND_PING As Integer = 200 ' GSZAO

Public Const NingunEscudo As Integer = 2
Public Const NingunCasco As Integer = 2
Public Const NingunArma As Integer = 2

Public Const EspadaMataDragonesIndex As Integer = 402
Public Const LAUDMAGICO As Integer = 696
Public Const FLAUTAMAGICA As Integer = 208

Public Const LAUDELFICO As Integer = 1049
Public Const FLAUTAELFICA As Integer = 1050

Public Const APOCALIPSIS_SPELL_INDEX As Integer = 25
Public Const DESCARGA_SPELL_INDEX As Integer = 23

Public Const SLOTS_POR_FILA As Byte = 5

Public Const PROB_ACUCHILLAR As Byte = 20
Public Const DAÑO_ACUCHILLAR As Single = 0.2

Public Const MAXMASCOTASENTRENADOR As Byte = 7

Public Enum FXIDs
    FXWARP = 1
    FXMEDITARCHICO = 4
    FXMEDITARMEDIANO = 5
    FXMEDITARGRANDE = 6
    FXMEDITARXGRANDE = 16
    FXMEDITARXXGRANDE = 34
End Enum

Public Const TIEMPO_CARCEL_PIQUETE As Long = 10

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param BAJOTECHOSINNPCS bajo techo y los NPCs no pueden pisar o respawn en tiles con este trigger
' @param SINNPCS los NPCs no pueden pisar o respawn en tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE los usuarios no pueden quedarse parados en este trigger demasiado tiempo o serán encarcelados
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger
    NADA = 0
    BAJOTECHO = 1
    BAJOTECHOSINNPCS = 2
    SINNPCS = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

''
' constantes para el trigger 6
'
' @see eTrigger
' @param TRIGGER6_PERMITE TRIGGER6_PERMITE
' @param TRIGGER6_PROHIBE TRIGGER6_PROHIBE
' @param TRIGGER6_AUSENTE El trigger no aparece
'
Public Enum eTrigger6
    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3
End Enum

'TODO : Reemplazar por un enum
Public Const Bosque As String = "BOSQUE"
Public Const Nieve As String = "NIEVE"
Public Const Desierto As String = "DESIERTO"
Public Const Ciudad As String = "CIUDAD"
Public Const Campo As String = "CAMPO"
Public Const Dungeon As String = "DUNGEON"

Public Enum eTerrain '0.13.3
    terrain_bosque = 0
    terrain_nieve = 1
    terrain_desierto = 2
    terrain_ciudad = 3
    terrain_campo = 4
    terrain_dungeon = 5
End Enum

Public Enum eRestrict '0.13.3
    restrict_no = 0
    restrict_newbie = 1
    restrict_armada = 2
    restrict_caos = 3
    restrict_faccion = 4
End Enum

' <<<<<< Targets >>>>>>
Public Enum TargetType
    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4
End Enum

' <<<<<< Acciona sobre >>>>>>
Public Enum TipoHechizo
    uPropiedades = 1
    uEstado = 2
    uMaterializa = 3    'Nose usa
    uInvocacion = 4
End Enum

Public Const MAXUSERHECHIZOS As Byte = 30


' TODO: Y ESTO ? LO CONOCE GD ?
Public Const EsfuerzoTalarGeneral As Byte = 4
Public Const EsfuerzoTalarLeñador As Byte = 2

Public Const EsfuerzoPescarPescador As Byte = 1
Public Const EsfuerzoPescarGeneral As Byte = 3

Public Const EsfuerzoExcavarMinero As Byte = 2
Public Const EsfuerzoExcavarGeneral As Byte = 5

Public Const FX_TELEPORT_INDEX As Integer = 1

Public Const PORCENTAJE_MATERIALES_UPGRADE As Single = 0.85

' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
Public Enum PartesCuerpo
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public Const Guardias As Integer = 6

Public Const MAX_ORO_EDIT As Long = 5000000
Public Const MAX_VIDA_EDIT As Long = 30000 ' 0.13.3

Public Const STANDARD_BOUNTY_HUNTER_MESSAGE As String = "Se te ha otorgado un premio por ayudar al proyecto reportando bugs, el mismo está disponible en tu bóveda."
Public Const TAG_USER_INVISIBLE As String = "[INVISIBLE]"
Public Const TAG_CONSULT_MODE As String = "[CONSULTA]"

Public Const MAXREP As Long = 6000000
Public Const MAXORO As Long = 90000000
Public Const MAXEXP As Long = 99999999

Public Const MAXUSERMATADOS As Long = 65000

Public Const MAXATRIBUTOS As Byte = 40
Public Const MINATRIBUTOS As Byte = 6

Public Const MAXNPCS As Integer = 10000
Public Const MAXCHARS As Integer = 10000

' Objetios configurados para la Herreria
Public Const MARTILLO_HERRERO As Integer = 389
Public Const MARTILLO_HERRERO_NEWBIE As Integer = 565
Public Const LingoteHierro As Integer = 386
Public Const LingotePlata As Integer = 387
Public Const LingoteOro As Integer = 388
Public Const PIQUETE_MINERO As Integer = 187
Public Const PIQUETE_MINERO_NEWBIE As Integer = 562
Public Const ORO_MINA As Integer = 194
Public Const PLATA_MINA As Integer = 193
Public Const HIERRO_MINA As Integer = 192

' Objetos configurados para la Carpinteria
Public Const SERRUCHO_CARPINTERO As Integer = 198
Public Const SERRUCHO_CARPINTERO_NEWBIE As Integer = 564
Public Const Leña As Integer = 58 ' Madera
Public Const LeñaElfica As Integer = 1006 ' Madera Elfica
Public Const HACHA_LEÑADOR As Integer = 127
Public Const HACHA_LEÑADOR_NEWBIE As Integer = 561
Public Const HACHA_LEÑA_ELFICA As Integer = 1005

' PESCA
Public Const RED_PESCA As Integer = 543
Public Const CAÑA_PESCA As Integer = 138
Public Const CAÑA_PESCA_NEWBIE As Integer = 563 ' 0.13.3

' Sistema de Fogatas
Public Const DAGA As Integer = 15 ' Se requiere UNICAMENTE para crear fogatas
Public Const FOGATA_APAG As Integer = 136 ' Ramitas (Fogata apagada)
Public Const FOGATA As Integer = 63 ' Fogata (encendida)

' Tipos de NPC's
Public Enum eNPCType
    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Noble = 5
    DRAGON = 6
    Timbero = 7
    GuardiasCaos = 8
    ResucitadorNewbie = 9
    Pretoriano = 10
    Gobernador = 11
    GuardiasEspeciales = 12
End Enum

Public Const MIN_APUÑALAR As Byte = 10

'********** CONSTANTANTES ***********

''
' Cantidad de skills
Public Const NUMSKILLS As Byte = 20

''
' Cantidad de Atributos
Public Const NUMATRIBUTOS As Byte = 5

''
' Cantidad de Clases
Public Const NUMCLASES As Byte = 12

''
' Cantidad de Razas
Public Const NUMRAZAS As Byte = 5


''
' Valor maximo de cada skill
Public Const MAXSKILLPOINTS As Byte = 100

''
' Cantidad de Ciudades
Public NUMCIUDADES As Byte ' GSZAO


''
'Direccion
'
' @param NORTH Norte
' @param EAST Este
' @param SOUTH Sur
' @param WEST Oeste
'
Public Enum eHeading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

''
' Cantidad maxima de mascotas
Public Const MAXMASCOTAS As Byte = 3

' 0.13.3
Public Const NUM_PECES As Integer = 4
Public ListaPeces(1 To NUM_PECES) As Integer

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const vlASALTO As Integer = 100
Public Const vlASESINO As Integer = 1000
Public Const vlCAZADOR As Integer = 5
Public Const vlNoble As Integer = 5
Public Const vlLadron As Integer = 25
Public Const vlProleta As Integer = 2

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto As Integer = 8
Public Const iCabezaMuerto As Integer = 500


Public Const iORO As Byte = 12
Public Const Pescado As Byte = 139

Public Enum PECES_POSIBLES
    PESCADO1 = 139
    PESCADO2 = 544
    PESCADO3 = 545
    PESCADO4 = 546
End Enum

Public ObjMatrimonio1 As Integer ' GSZAO - obj que se entrega al casarse (para divorciarse)
Public ObjMatrimonio2 As Integer ' GSZAO - obj necesario "antes" de consumar el matrimonio

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Enum eSkill
    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apuñalar = 6
    Ocultarse = 7
    Supervivencia = 8
    Talar = 9
    Comerciar = 10
    Defensa = 11
    Pesca = 12
    Mineria = 13
    Carpinteria = 14
    Herreria = 15
    Liderazgo = 16
    Domar = 17
    Proyectiles = 18
    Wrestling = 19
    Navegacion = 20
End Enum

' GSZAO
Public Enum eAccionClick
    Matrimonio = 21
    Divorcio = 22
End Enum
' GSZAO

Public Enum eMochilas
    Mediana = 1
    Grande = 2
End Enum

Public Const FundirMetal = 88

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Public Const AdicionalHPGuerrero As Byte = 2 'HP adicionales cuando sube de nivel
Public Const AdicionalHPCazador As Byte = 1 'HP adicionales cuando sube de nivel

Public Const AumentoSTDef As Byte = 15
Public Const AumentoStBandido As Byte = AumentoSTDef + 3 ' 0.13.3
Public Const AumentoSTLadron As Byte = AumentoSTDef + 3
Public Const AumentoSTMago As Byte = AumentoSTDef - 1
Public Const AumentoSTTrabajador As Byte = AumentoSTDef + 25

'Tamaño del mapa
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'Tamaño del tileset
Public Const TileSizeX As Byte = 32
Public Const TileSizeY As Byte = 32

'Tamaño en Tiles de la pantalla de visualizacion
Public Const XWindow As Byte = 17
Public Const YWindow As Byte = 13

'Sonidos
Public Const SND_SWING As Byte = 2
Public Const SND_TALAR As Byte = 13
Public Const SND_PESCAR As Byte = 14
Public Const SND_MINERO As Byte = 15
Public Const SND_WARP As Byte = 3
Public Const SND_PUERTA As Byte = 5
Public Const SND_NIVEL As Byte = 6

Public Const SND_USERMUERTE As Byte = 11
Public Const SND_IMPACTO As Byte = 10
Public Const SND_IMPACTO2 As Byte = 12
Public Const SND_LEÑADOR As Byte = 13
Public Const SND_FOGATA As Byte = 14
Public Const SND_AVE As Byte = 21
Public Const SND_AVE2 As Byte = 22
Public Const SND_AVE3 As Byte = 34
Public Const SND_GRILLO As Byte = 28
Public Const SND_GRILLO2 As Byte = 29
Public Const SND_SACARARMA As Byte = 25
Public Const SND_ESCUDO As Byte = 37
Public Const MARTILLOHERRERO As Byte = 41
Public Const LABUROCARPINTERO As Byte = 42
Public Const SND_BEBER As Byte = 46

''
' Cantidad maxima de objetos por slot de inventario
Public Const MAX_INVENTORY_OBJS As Integer = 10000

''
' Cantidad de "slots" en el inventario con mochila
Public Const MAX_INVENTORY_SLOTS As Byte = 30

''
' Cantidad de "slots" en el inventario sin mochila
Public Const MAX_NORMAL_INVENTORY_SLOTS As Byte = 20

''
' Constante para indicar que se esta usando ORO
Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1


' CATEGORIAS PRINCIPALES
Public Enum eOBJType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otESCUDO = 16
    otCASCO = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otArbolElfico = 36
    otMochilas = 37
    otYacimientoPez = 38    ' 0.13.3
    otPasaje = 39           ' GSZAO
    otDestruible = 40       ' GSZAO
    otMatrimonio = 41       ' GSZAO
    otCualquiera = 1000
End Enum

'Texto
Public Const FONTTYPE_TALK As String = "~255~255~255~0~0"
Public Const FONTTYPE_FIGHT As String = "~255~0~0~1~0"
Public Const FONTTYPE_WARNING As String = "~32~51~223~1~1"
Public Const FONTTYPE_INFO As String = "~65~190~156~0~0"
Public Const FONTTYPE_INFOBOLD As String = "~65~190~156~1~0"
Public Const FONTTYPE_EJECUCION As String = "~130~130~130~1~0"
Public Const FONTTYPE_PARTY As String = "~255~180~255~0~0"
Public Const FONTTYPE_VENENO As String = "~0~255~0~0~0"
Public Const FONTTYPE_GUILD As String = "~255~255~255~1~0"
Public Const FONTTYPE_SERVER As String = "~0~185~0~0~0"
Public Const FONTTYPE_GUILDMSG As String = "~228~199~27~0~0"
Public Const FONTTYPE_CONSEJO As String = "~130~130~255~1~0"
Public Const FONTTYPE_CONSEJOCAOS As String = "~255~60~00~1~0"
Public Const FONTTYPE_CONSEJOVesA As String = "~0~200~255~1~0"
Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~255~50~0~1~0"
Public Const FONTTYPE_CENTINELA As String = "~0~255~0~1~0"

' Roles
Public Const FONTTYPE_DIOS As String = "250~250~150~1~0" ' GSZAO
Public Const FONTTYPE_SEMIDIOS As String = "30~255~30~1~0" ' GSZAO
Public Const FOTNTYPE_CONSEJERO As String = "30~150~30~1~0" ' GSZAO

Public Const FONTTYPE_GOLD As String = "~255~215~0~1~0" ' GSZAO
Public Const FONTTYPE_OBJ As String = "~175~238~238~1~0" ' GSZAO
Public Const FONTTYPE_NPC_WARNING As String = "~235~51~51~1~0" ' GSZAO
Public Const FONTTYPE_NPC_PEACE As String = "~51~235~51~1~0" ' GSZAO
' "Public Enum FontTypeNames" se encuentra en modProtocol del SERVIDOR
' Y tambien debe actualizarse "Public Enum FontTypeNames" en el modProtocol del CLIENTE
' El maximo de la variable Public FontTypes(24) As tFont...
' ...y la función InitFont del CLIENTE...


'Estadisticas
'Public Const STAT_MAXELV As Byte = 255 ' GSZAO
Public Const STAT_MAXHP As Integer = 999
Public Const STAT_MAXSTA As Integer = 999
Public Const STAT_MAXMAN As Integer = 9999
Public Const STAT_MAXHIT_UNDER36 As Byte = 99
Public Const STAT_MAXHIT_OVER36 As Integer = 999
Public Const STAT_MAXDEF As Byte = 99

Public Const ELU_SKILL_INICIAL As Byte = 200
Public Const EXP_ACIERTO_SKILL As Byte = 50
Public Const EXP_FALLO_SKILL As Byte = 20

' **************************************************************
' **************************************************************
' ************************ TIPOS *******************************
' **************************************************************
' **************************************************************


Public Type tObservacion ' 0.13.3
    Creador As String
    Fecha As Date
    
    Detalles As String
End Type

Public Type tRecord ' 0.13.3
    Usuario As String
    Motivo As String
    Creador As String
    Fecha As Date
    
    NumObs As Byte
    Obs() As tObservacion
End Type

Public Type UserOBJ
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
End Type

Public Type tHechizo
    Nombre As String
    GrhIndex As Integer ' GSZAO
    desc As String
    PalabrasMagicas As String
    
    HechizeroMsg As String
    targetMSG As String
    PropioMsg As String
    
'    Resis As Byte
    ExclusivoClase As Byte ' GSZAO
    ExclusivoRaza As Byte ' GSZAO
    
    tipo As TipoHechizo

    PartIndex As Integer 'GSZAO
        
    WAV As Integer
    FXgrh As Integer
    loops As Byte
    
    SubeHP As Byte
    MinHp As Integer
    MaxHp As Integer
    
    SubeMana As Byte
    MiMana As Integer
    MaMana As Integer
    
    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer
    
    SubeHam As Byte
    MinHam As Integer
    MaxHam As Integer
    
    SubeSed As Byte
    MinSed As Integer
    MaxSed As Integer
    
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer
    
    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer
    
    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    RemoverEstupidez As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Morph As Byte
    Mimetiza As Byte
    RemueveInvisibilidadParcial As Byte
    
    Warp As Byte
    Invoca As Byte
    NumNpc As Integer
    cant As Integer

'    Materializa As Byte
'    ItemIndex As Byte
    
    MinSkill As Integer
    ManaRequerido As Integer

    ' 29/9/03
    StaRequerido As Integer

    Target As TargetType
    
    NeedStaff As Integer
    StaffAffected As Boolean
    
    ReqObjNum As Integer ' GSZAO
    ReqObj() As UserOBJ ' GSZAO
End Type

Public Type LevelSkill
    LevelValue As Integer
End Type

Public Type Inventario
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    AnilloEqpObjIndex As Integer
    AnilloEqpSlot As Byte
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    MochilaEqpObjIndex As Integer
    MochilaEqpSlot As Byte
    NroItems As Integer
End Type

Public Type tPartyData
    PIndex As Integer
    RemXP As Double 'La exp. en el server se cuenta con Doubles
    targetUser As Integer 'Para las invitaciones
End Type

Public Type Position
    X As Integer
    Y As Integer
End Type

Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public Type FXdata
    Nombre As String
    GrhIndex As Integer
    Delay As Integer
End Type

'Datos de user o npc
Public Type Char
    CharIndex As Integer
    Head As Integer
    Body As Integer
    
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    FX As Integer
    loops As Integer
    
    heading As eHeading
End Type

'Tipos de objetos
Public Type ObjData
    Name As String 'Nombre del obj
    
    OBJType As eOBJType 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    'Solo contenedores
    MaxItems As Integer
    Conte As Inventario
    Apuñala As Byte
    Acuchilla As Byte
    
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHp As Integer ' Minimo puntos de vida
    MaxHp As Integer ' Maximo puntos de vida
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    Proyectil As Integer
    Municion As Integer
    
    NoLimpiar As Byte ' GSZAO
    Newbie As Integer
    
    'Puntos de Stamina que da
    MinSta As Integer ' Minimo puntos de stamina
    
    'Pociones
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer
    
    MinHIT As Integer 'Minimo golpe
    MaxHIT As Integer 'Maximo golpe
    
    MinHam As Integer
    MinSed As Integer
    
    def As Long
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    WeaponRazaEnanaAnim As Integer
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    
    Valor As Long     ' Precio
    
    Cerrada As Integer
    Llave As Byte
    clave As Long 'si clave=llave la puerta se abre o cierra
    
    Radio As Integer ' Para teleps: El radio para calcular el random de la pos destino
    
    MochilaType As Byte 'Tipo de Mochila (1 la chica, 2 la grande)
    
    Guante As Byte ' Indica si es un guante o no.
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    RazaEnana As Byte
    RazaDrow As Byte
    RazaElfa As Byte
    RazaGnoma As Byte
    RazaHumana As Byte
    
    Mujer As Byte
    Hombre As Byte
    
    Envenena As Byte
    Paraliza As Byte
    
    NoAgarrable As Byte ' GSZAO
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    MaderaElfica As Integer
    
    SkHerreria As Integer
    SkCarpinteria As Integer
    
    texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As eClass
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    Real As Integer
    Caos As Integer
    
    NoSeCae As Integer
    NoSeTira As Integer ' 0.13.5
    NoRobable As Integer ' 0.13.5
    NoComerciable As Integer ' 0.13.5
    Intransferible As Integer ' 0.13.5
    
    Pasaje As WorldPos
    
    StaffPower As Integer
    StaffDamageBonus As Integer
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Refuerzo As Byte
    
    ImpideParalizar As Byte ' 0.13.5
    ImpideInmobilizar As Byte ' 0.13.5
    ImpideAturdir As Byte ' 0.13.5
    ImpideCegar As Byte ' 0.13.5

    Log As Byte 'es un objeto que queremos loguear?
    NoLog As Byte 'es un objeto que esta prohibido loguear?
    
    Bloqueado As Byte ' GSZAO
    Respawn As Integer ' GSZAO
    Upgrade As Integer
End Type

Public Type Obj
    ObjIndex As Integer
    Amount As Integer
    ExtraLong As Long ' GSZAO
End Type

' Extracto de http://www.gs-zone.org/threads/sistema-de-quests-blisseao.82560/
Public Type tQuestNpc ' NPC's Requeridos en las Quests
    NpcIndex As Integer
    Amount As Integer
End Type

Public Type tUserQuest '
    NPCsKilled() As Integer
    QuestIndex As Integer
End Type

Public Type tQuestStats ' Estadisticas del usuario en Quests
    Quests(1 To MAXUSERQUESTS) As tUserQuest
    NumQuestsDone As Integer
    QuestsDone() As Integer
End Type

Public Type tQuest ' Configuración de la Quest
    Nombre As String
    desc As String
    RequiredLevel As Byte
    RequiredOBJs As Byte
    RequiredOBJ() As Obj
    RequiredNPCs As Byte
    RequiredNPC() As tQuestNpc
    RewardGLD As Long
    RewardEXP As Long
    RewardOBJs As Byte
    RewardOBJ() As Obj
End Type

'[Pablo ToxicWaste]
Public Type ModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    AtaqueWrestling As Double
    DañoArmas As Double
    DañoProyectiles As Double
    DañoWrestling As Double
    Escudo As Double
End Type

Public Type ModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type
'[/Pablo ToxicWaste]

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
'[/KEVIN]

'[KEVIN]
Public Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    NroItems As Integer
End Type
'[/KEVIN]

' Determina el color del nick
Public Enum eNickColor
    ieCriminal = &H1
    ieCiudadano = &H2
    ieNewbie = &H3 ' GSZAO
    ieAtacable = &H4
    ieMuerto = &H5 ' GSZAO
End Enum

'*******
'FOROS *
'*******

' Tipos de mensajes
Public Enum eForumMsgType
    ieGeneral
    ieGENERAL_STICKY
    ieREAL
    ieREAL_STICKY
    ieCAOS
    ieCAOS_STICKY
End Enum

' Indica los privilegios para visualizar los diferentes foros
Public Enum eForumVisibility
    ieGENERAL_MEMBER = &H1
    ieREAL_MEMBER = &H2
    ieCAOS_MEMBER = &H4
End Enum

' Indica el tipo de foro
Public Enum eForumType
    ieGeneral
    ieREAL
    ieCAOS
End Enum

' Limite de posts
Public Const MAX_STICKY_POST As Byte = 10
Public Const MAX_GENERAL_POST As Byte = 35

' Estructura contenedora de mensajes
Public Type tForo
    StickyTitle(1 To MAX_STICKY_POST) As String
    StickyPost(1 To MAX_STICKY_POST) As String
    GeneralTitle(1 To MAX_GENERAL_POST) As String
    GeneralPost(1 To MAX_GENERAL_POST) As String
End Type

'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    Promedio As Long
End Type

'Estadisticas de los usuarios
Public Type UserStats
    GLD As Long 'Dinero
    Banco As Long
    
    MaxHp As Integer
    MinHp As Integer
    
    MaxSta As Integer
    MinSta As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHIT As Integer
    MinHIT As Integer
    
    MaxHam As Integer
    MinHam As Integer
    
    MaxAGU As Integer
    MinAGU As Integer
        
    def As Integer
    Exp As Double
    ELV As Byte
    ELU As Long
    UserSkills(1 To NUMSKILLS) As Byte
    UserAtributos(1 To NUMATRIBUTOS) As Byte
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Byte
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Long
    CriminalesMatados As Long
    NPCsMuertos As Integer
    
    SkillPts As Integer
    
    ExpSkills(1 To NUMSKILLS) As Long
    EluSkills(1 To NUMSKILLS) As Long
    
End Type


'Flags
Public Type UserFlags
    'Cuentas
    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Comerciando As Boolean '¿Esta comerciando?
    UserLogged As Boolean '¿Esta online?
    Meditando As Boolean
    Descuento As String
    Hambre As Byte
    Sed As Byte
    PuedeMoverse As Byte
    TimerLanzarSpell As Long
    PuedeTrabajar As Byte
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    
    NoPuedeSerAtacado As Boolean
    AtacablePor As Integer
    ShareNpcWith As Integer
    
    Vuela As Byte
    Navegando As Byte
    Seguro As Boolean
    SeguroResu As Boolean
    
    DuracionEfecto As Long
    TargetNPC As Integer ' Npc señalado por el usuario
    TargetNpcTipo As eNPCType ' Tipo del npc señalado
    OwnedNpc As Integer ' Npc que le pertenece (no puede ser atacado)
    NpcInv As Integer
    
    Ban As Byte
    AdministrativeBan As Byte
    
    targetUser As Integer ' Usuario señalado
    
    targetObj As Integer ' Obj señalado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
    
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
    
    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer
    
    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer
    NPCAtacado As Integer
    Ignorado As Boolean
    
    EnConsulta As Boolean
    SendDenounces As Boolean    ' 0.13.3
    
    StatsChanged As Byte
    Privilegios As PlayerType
    PrivEspecial As Boolean     ' 0.13.3
    
    CaptchaCode(3) As Byte      ' GSZAO
    CaptchaKey As Byte          ' GSZAO
    
    LastCrimMatado As String
    LastCiudMatado As String
    
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    AdminPerseguible As Boolean
    
    ChatColor As Long
    
    '[el oso]
    MD5Reportado As String
    '[/el oso]
    
    '[Barrin 30-11-03]
    TimesWalk As Long
    StartWalk As Long
    CountSH As Long
    '[/Barrin 30-11-03]
    
    '[CDT 17-02-04]
    UltimoMensaje As Byte
    '[/CDT]
    
    Silenciado As Byte
    
    Mimetizado As Byte
    
    CentinelaIndex As Byte ' Indice del centinela que lo revisa ' 0.13.3
    CentinelaOK As Boolean
   
    lastMap As Integer
    Traveling As Byte 'Travelin Band ¿?
    
    ' 0.13.3
    ParalizedBy As String
    ParalizedByIndex As Integer
    ParalizedByNpcIndex As Integer
    
    ' GSZAO
    Matrimonio As String
    FormYesNoType As Byte   ' tipo de form enviado
    FormYesNoA As Integer   ' envio form
    FormYesNoDE As Integer  ' responde form
    SerialHD As Long
End Type

Public Type UserCounters
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    Lava As Integer
    COMCounter As Integer
    AGUACounter As Integer
    Veneno As Integer
    Paralisis As Integer
    Ceguera As Integer
    Estupidez As Integer
    
    Invisibilidad As Integer
    TiempoOculto As Integer
    
    Mimetismo As Integer
    PiqueteC As Long
    Pena As Long
    SendMapCounter As WorldPos
    '[Gonzalo]
    Saliendo As Boolean
    Salir As Integer
    '[/Gonzalo]
    
    tInicioMeditar As Long
    bPuedeMeditar As Boolean
    
    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeUsarArco As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    TimerMagiaGolpe As Long
    TimerGolpeMagia As Long
    TimerGolpeUsar As Long
    TimerPuedeSerAtacado As Long
    TimerPerteneceNpc As Long
    TimerEstadoAtacable As Long
    TimerPuedeSendPing As Long ' GSZAO
    
    Trabajando As Long  ' Para el centinela
    Ocultando As Long   ' Unico trabajo no revisado por el centinela
    
    failedUsageAttempts As Long
    
    goHome As Long
    AsignedSkills As Byte
End Type

'Cosas faccionarias.
Public Type tFacciones
    ArmadaReal As Byte
    FuerzasCaos As Byte
    CriminalesMatados As Long
    CiudadanosMatados As Long
    RecompensasReal As Long
    RecompensasCaos As Long
    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte
    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte
    Reenlistadas As Byte
    NivelIngreso As Integer
    FechaIngreso As String
    MatadosIngreso As Integer 'Para Armadas nada mas
    NextRecompensa As Integer
End Type

Public Type tCrafting
    Cantidad As Long
    PorCiclo As Integer
End Type

Public Enum eCEstado
        Libre = 0
        Ocupado = 1
End Enum

'Tipo de los Usuarios
Public Type User
    IndexPJ As Integer ' ID de la Base de datos /// CASTELLI (FEDUDOK)
    Name    As String
    ID      As Long
    
    Disp    As eCEstado
    
    email   As String
    Password As String

    ShowName As Boolean 'Permite que los GMs oculten su nick con el comando /SHOWNAME
    
    Char As Char 'Define la apariencia
    CharMimetizado As Char
    OrigChar As Char
    
    desc As String ' Descripcion
    DescRM As String
    
    clase As eClass
    raza As eRaza
    Genero As eGenero
    
    Hogar As Byte ' GSZAO
        
    Invent As Inventario
    
    Pos As WorldPos
    
    ConnIDValida As Boolean
    ConnID As Long 'ID
    
    '[KEVIN]
    BancoInvent As BancoInventario
    '[/KEVIN]
    
    Counters As UserCounters
    
    Construir As tCrafting
    
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMascotas As Integer
    
    Stats As UserStats
    flags As UserFlags
    
    Reputacion As tReputacion
    
    fAccion As tFacciones

    ip As String
    
    ComUsu As tCOmercioUsuario
    
    GuildIndex As Integer   'puntero al array global de guilds
    FundandoGuildAlineacion As ALINEACION_GUILD     'esto esta aca hasta que se parchee el cliente y se pongan cadenas de datos distintas para cada alineacion
    EscucheClan As Integer
    
    PartyIndex As Integer   'index a la party q es miembro
    PartySolicitud As Integer   'index a la party q solicito
    
    KeyCrypt As Integer
    
    AreasInfo As AreaInfo
    
    'Outgoing and incoming messages
    outgoingData As clsByteQueue
    incomingData As clsByteQueue
    
    CurrentInventorySlots As Byte
    QuestStats As tQuestStats ' GSZAO
End Type


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type NPCStats
    Alineacion As Integer
    MaxHp As Long
    MinHp As Long
    MaxHIT As Integer
    MinHIT As Integer
    def As Integer
    defM As Integer
End Type

Public Type NpcCounters
    Paralisis As Integer
    TiempoExistencia As Long
    TiempoPersiguiendo As Byte ' GSZAO - 100 seg max
End Type

Public Type NPCFlags
    AfectaParalisis As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean '¿Esta vivo?
    Follow As Boolean
    fAccion As Byte
    AtacaDoble As Byte
    LanzaSpells As Byte
    
    ExpCount As Long
    
    OldMovement As TipoAI
    OldHostil As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    
    Sound As Integer
    AttackedBy As String
    AttackedFirstBy As String
    Backup As Byte
    RespawnOrigPos As Byte
    
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    UsuariosMatados As Integer ' GSZAO
End Type

Public Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

' New type for holding the pathfinding info
Public Type NpcPathFindingInfo
    Path() As tVertice      ' This array holds the path
    Target As Position      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos As Integer       ' Current location of the npc
    targetUser As Integer   ' UserIndex chased
    NoPath As Boolean       ' If it is true there is no path to the target location
    
    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.
    
End Type
' New type for holding the pathfinding info

Public Type tDrops
    ObjIndex As Integer
    Amount As Long
    Equipped As Byte ' GSZAO, para que los NPC's puedan equipar objetos que tienen en el DROP
End Type

Public Const MAX_NPC_DROPS As Byte = 10 ' Objetos maximos que puede tirar el NPC!!

Public Type npc
    Name As String
    ShowName As Boolean ' GSZAO ¿Mostrar Nombre?
    
    Char As Char 'Define como se vera
    desc As String

    NPCtype As eNPCType
    
    QuestNumber As Integer ' GSZAO
    
    AttackLvlMore As Byte ' GSZAO
    AttackLvlLess As Byte ' GSZAO

    Numero As Integer

    InvReSpawn As Byte

    Comercia As Integer
    Target As Long
    TargetNPC As Long
    TipoItems As Integer

    Veneno As Byte

    Pos As WorldPos 'Posicion
    Orig As WorldPos
    SkillDomar As Integer

    Movement As TipoAI
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long

    Owner As Integer

    GiveEXP As Long
    GiveGLD As Long
    Drop(1 To MAX_NPC_DROPS) As tDrops
    
    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    
    Invent As Inventario
    CanAttack As Byte
    
    NroExpresiones As Byte
    Expresiones() As String ' le da vida ;)
    
    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)
    
    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas As Integer
    
    ' New!! Needed for pathfindig
    PFINFO As NpcPathFindingInfo
    AreasInfo As AreaInfo
    
    'Hogar
    Ciudad As Byte
    
    'Para diferenciar entre clanes ' 0.13.3
    ClanIndex As Integer
End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Public Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Integer
    UserIndex As Integer
    NpcIndex As Integer
    ObjInfo As Obj
    TileExit As WorldPos
    trigger As eTrigger
End Type

'Info del mapa
Type MapInfo
    NumUsers As Integer
    Music As String
    Name As String
    StartPos As WorldPos
    OnDeathGoTo As WorldPos     ' 0.13.3

    MapVersion As Integer
    Pk As Boolean
    MagiaSinEfecto As Byte
    NoEncriptarMP As Byte
    
    ' Anti Magias/Habilidades
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    ' 0.13.3
    OcultarSinEfecto As Byte
    InvocarSinEfecto As Byte
    
    RoboNpcsPermitido As Byte
    
    Terreno As String
    Zona As String
    Restringir As Byte
    Backup As Byte
    
    ' GSZAO - Grids's
    Grid(1 To 4) As Integer
    '   Utilizamos eHeading:
    '    NORTH = 1 / Arriba
    '    EAST = 2 / Izquierda
    '    SOUTH = 3 / Abajo
    '    WEST = 4 / Derecha
End Type


'********** V A R I A B L E S     P U B L I C A S ***********

Public SERVERONLINE As Boolean
Public iniNombre As String
Public iniWWW As String
Public iniVersion As String
Public iniDiaNoche As Boolean
Public iniSistemaLuces As Boolean
Public iniSiempreNombres As Boolean
Public iniWorldGrid As String ' GSZAO
' Opciones
Public iniDragDrop As Byte
Public iniTirarOBJZonaSegura As Byte
Public iniMeditarRapido As Boolean
Public iniPrivadoPorConsola As Boolean
Public iniAutoSacerdote As Boolean
Public iniSacerdoteCuraVeneno As Boolean
Public iniNPCHostilesConNombre As Boolean
Public iniNPCNoHostilesConNombre As Boolean
' Balance
Public iniMaxNivel As Byte ' 255 max
Public iniOro As Single ' balance de Oro
Public iniExp As Single ' balance de Exp
Public iniTPesca As Single ' balance de trabajo de Pesca
Public iniTMineria As Single ' balance de trabajo de Mineria
Public iniTTala As Single ' balance de trabajo de Tala
Public iniBilletera As Long ' =0 deshabilitada (se cae siempre el oro)
Public iniBilleteraSegura As Boolean ' =1 no se cae >billetera, con =0 solo se cae el resto y deja la billetera
' Meditar
Public iniFxMedChico As Byte
Public iniFxMedMediano As Byte
Public iniFxMedGrande As Byte
Public iniFxMedExtraGrande As Byte
' Clanes
Public iniCNivel As Byte
Public iniCLiderazgo As Byte
Public iniCRequiereObj As Integer
Public iniCRequiereObjCnt As Integer

Public Backup As Boolean ' TODO: Se usa esta variable ?

Public ListaRazas(1 To NUMRAZAS) As String
Public SkillsNames(1 To NUMSKILLS) As String
Public ListaClases(1 To NUMCLASES) As String
Public ListaAtributos(1 To NUMATRIBUTOS) As String

Public iniRecord As Long

'
'Directorios
'

''
'Ruta base del server, en donde esta el "Servidor.ini"
Public IniPath As String

''
'Ruta base para guardar los chars
Public CharPath As String

''
'Ruta base para los DATs
Public DatPath As String

''
'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

''
'Numero de usuarios actual
Public NumUsers As Integer
Public LastUser As Integer
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumFX As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer
Public NumeroHechizos As Integer
Public iniMultiLogin As Byte
Public iniInactivo As Integer
Public iniMaxUsuarios As Integer
Public iniOculto As Byte
Public LastBackup As String
Public Minutos As String
Public haciendoBK As Boolean
Public iniPuedeCrearPersonajes As Integer
Public iniSoloGMs As Byte
Public NumRecords As Integer    ' 0.13.3

' Sistema de Happy Hour (adaptado de 0.13.5)
Public iniHappyHourActivado As Boolean ' GSZAO
Public HappyHour As Single      ' 0.13.5
Public HappyHourActivated As Boolean      ' 0.13.5
Public Type tHappyHour ' GSZAO
    Multi As Single ' Multi
    Hour As Integer ' Hora
End Type
Public HappyHourDays(1 To 7) As tHappyHour    ' 0.13.5

''
'Esta activada la verificacion MD5 ?
Public MD5ClientesActivado As Byte


Public EnPausa As Boolean
Public iniTesting As Boolean
Public iniLogDesarrollo As Boolean ' GSZAO


'*****************ARRAYS PUBLICOS*************************
Public UserList() As User 'USUARIOS
Public Npclist(1 To MAXNPCS) As npc 'NPCS
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public CharList(1 To MAXCHARS) As Integer
Public ObjData() As ObjData
Public FX() As FXdata
Public SpawnList() As tCriaturasEntrenador
Public LevelSkill(1 To 50) As LevelSkill
Public ForbidenNames() As String
Public MD5s() As String
Public BanIPs As Collection
Public BanHDs As Collection ' GSZAO
Public Parties(1 To MAX_PARTIES) As clsParty
Public ModClase(1 To NUMCLASES) As ModClase
Public ModRaza(1 To NUMRAZAS) As ModRaza
Public ModVida(1 To NUMCLASES) As Double
Public DistribucionEnteraVida(1 To 5) As Integer
Public DistribucionSemienteraVida(1 To 4) As Integer
Public Ciudades() As WorldPos ' GSZAO
Public QuestList() As tQuest ' GSZAO
Public distanceToCities() As HomeDistance
Public Records() As tRecord     ' 0.13.3
' Listados de construcción
Public lHerreroArmas() As Integer
Public lHerreroArmaduras() As Integer
Public lCarpintero() As Integer
'*********************************************************

Type HomeDistance
    distanceToCity(1 To 25) As Integer ' GSZAO = max 25
End Type

Public Prision As WorldPos
Public Libertad As WorldPos

Public Ayuda As clsCola
Public Denuncias As clsCola   ' 0.13.3
Public ConsultaPopular As clsConsultasPopulares
Public SonidosMapas As clsSoundMapInfo

Public Enum e_ObjetosCriticos
    Manzana = 1
    Manzana2 = 2
    ManzanaNewbie = 467
End Enum

' GSZAO
'********** Constantes de daño.
Public Const DAMAGE_PUÑAL    As Byte = 1
Public Const DAMAGE_NORMAL   As Byte = 2
Public Const DAMAGE_MAGIC    As Byte = 3
'********** Constantes de daño.

Public Const MATRIX_INITIAL_MAP As Integer = 1

Public Const GOHOME_PENALTY As Integer = 5
Public Const GM_MAP As Integer = 49

Public Const TELEP_OBJ_INDEX As Integer = 1012

Public Const HUMANO_H_PRIMER_CABEZA As Integer = 1
Public Const HUMANO_H_ULTIMA_CABEZA As Integer = 40 'En verdad es hasta la 51, pero como son muchas estas las dejamos no seleccionables

Public Const ELFO_H_PRIMER_CABEZA As Integer = 101
Public Const ELFO_H_ULTIMA_CABEZA As Integer = 122

Public Const DROW_H_PRIMER_CABEZA As Integer = 201
Public Const DROW_H_ULTIMA_CABEZA As Integer = 221

Public Const ENANO_H_PRIMER_CABEZA As Integer = 301
Public Const ENANO_H_ULTIMA_CABEZA As Integer = 319

Public Const GNOMO_H_PRIMER_CABEZA As Integer = 401
Public Const GNOMO_H_ULTIMA_CABEZA As Integer = 416
'**************************************************
Public Const HUMANO_M_PRIMER_CABEZA As Integer = 70
Public Const HUMANO_M_ULTIMA_CABEZA As Integer = 89

Public Const ELFO_M_PRIMER_CABEZA As Integer = 170
Public Const ELFO_M_ULTIMA_CABEZA As Integer = 188

Public Const DROW_M_PRIMER_CABEZA As Integer = 270
Public Const DROW_M_ULTIMA_CABEZA As Integer = 288

Public Const ENANO_M_PRIMER_CABEZA As Integer = 370
Public Const ENANO_M_ULTIMA_CABEZA As Integer = 384

Public Const GNOMO_M_PRIMER_CABEZA As Integer = 470
Public Const GNOMO_M_ULTIMA_CABEZA As Integer = 484

' Por ahora la dejo constante.. SI se quisiera extender la propiedad de paralziar, se podria hacer
' una nueva variable en el dat.
Public Const GUANTE_HURTO As Integer = 873

' 0.13.3
Public Const ESPADA_VIKINGA As Integer = 123


Public ClanPretoriano() As clsClanPretoriano
Public Administradores As clsIniManager

Public Const MIN_AMOUNT_LOG As Integer = 1000 ' 0.13.5
Public Const MIN_VALUE_LOG As Long = 50000 ' 0.13.5
Public Const MIN_GOLD_AMOUNT_LOG As Long = 10000 ' 0.13.5

Public Const MAX_DENOUNCES As Integer = 20

'Mensajes de los NPCs enlistadores (Nobles):
Public Const MENSAJE_REY_CAOS As String = "¿Esperabas pasar desapercibido, intruso? Los servidores del Demonio no son bienvenidos, ¡Guardias, a él!"
Public Const MENSAJE_REY_CRIMINAL_NOENLISTABLE As String = "Tus pecados son grandes, pero aún así puedes redimirte. El pasado deja huellas, pero aún puedes limpiar tu alma."
Public Const MENSAJE_REY_CRIMINAL_ENLISTABLE As String = "Limpia tu reputación y paga por los delitos cometidos. Un miembro de la Armada Real debe tener un comportamiento ejemplar."

Public Const MENSAJE_DEMONIO_REAL As String = "Lacayo de Tancredo, ve y dile a tu gente que nadie pisará estas tierras si no se arrodilla ante mi."
Public Const MENSAJE_DEMONIO_CIUDADANO_NOENLISTABLE As String = "Tu indecisión te ha condenado a una vida sin sentido, aún tienes elección... Pero ten mucho cuidado, mis hordas nunca descansan."
Public Const MENSAJE_DEMONIO_CIUDADANO_ENLISTABLE As String = "Siento el miedo por tus venas. Deja de ser escoria y únete a mis filas, sabrás que es el mejor camino."


' Funciones del API:
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef destination As Any, ByVal length As Long)
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nsize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nsize As Long) As Long

' GSZAO - Listados para busquedas
Public NpcListNames(1 To MAXNPCS) As String
Public ObjListNames() As String
