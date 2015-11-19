VERSION 5.00
Begin VB.Form frmGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Font.Dat"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdIndex 
      Caption         =   "Indexar Font.Dat"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer Font.Dat"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tCabecera 'Cabecera
    Desc As String * 255
    CRC As Long
    MagicWord As Long
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
    Color As Long
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
Private Type tCustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
End Type

Private cfonts(1 To 2) As tCustomFont ' _Default2 As CustomFont

Private Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    'Cabecera.Desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
    Cabecera.Desc = "GS-Zone Argentum Online MOD - Copyright GS-Zone 2012 - info@gs-zone.org - Original by Pablo Marquez "
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub


Private Function Index()
On Local Error Resume Next

    If LenB(Dir(App.Path & "\Config.Ini", vbArchive)) = 0 Then
        MsgBox "Se requiere Config.Ini en el directorio del programa.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    Dim MiCabecera As tCabecera
    Dim ConfigInit As tConfigInit
    Dim N As Integer
    
    ' CONFIG
    ConfigInit.MostrarTips = Val(GetVar(App.Path & "\Config.Ini", "CONFIG", "MostrarTips"))
    ConfigInit.NumParticulas = Val(GetVar(App.Path & "\Config.Ini", "CONFIG", "NumParticulas"))

    ' DIR
    ConfigInit.DirMultimedia = GetVar(App.Path & "\Config.Ini", "DIR", "Multimedia")
    ConfigInit.DirMapas = GetVar(App.Path & "\Config.Ini", "DIR", "Mapas")
    ConfigInit.DirGraficos = GetVar(App.Path & "\Config.Ini", "DIR", "Graficos")
    ConfigInit.DirFotos = GetVar(App.Path & "\Config.Ini", "DIR", "Fotos")
    ConfigInit.DirExtras = GetVar(App.Path & "\Config.Ini", "DIR", "Extras")
    ConfigInit.DirSonidos = GetVar(App.Path & "\Config.Ini", "DIR", "Sonidos")
    ConfigInit.DirMusicas = GetVar(App.Path & "\Config.Ini", "DIR", "Musicas")
    ConfigInit.DirParticulas = GetVar(App.Path & "\Config.Ini", "DIR", "Particulas")
    ConfigInit.DirGUI = GetVar(App.Path & "\Config.Ini", "DIR", "GUI")
    ConfigInit.DirBotones = GetVar(App.Path & "\Config.Ini", "DIR", "Botones")
    ConfigInit.DirFrags = GetVar(App.Path & "\Config.Ini", "DIR", "Frags")
    ConfigInit.DirMuertes = GetVar(App.Path & "\Config.Ini", "DIR", "Muertes")
    
    ' USUARIO
    ConfigInit.Nombre = GetVar(App.Path & "\Config.Ini", "USUARIO", "Nombre")
    ConfigInit.Password = GetVar(App.Path & "\Config.Ini", "USUARIO", "Password")
    
    If LenB(Dir(App.Path & "\Config.Init", vbArchive)) <> 0 Then
        Kill App.Path & "\Config.Init"
    End If
    
    Call IniciarCabecera(MiCabecera)
    N = FreeFile
    
    Open App.Path & "\Config.Init" For Binary As #N
    Put #N, , MiCabecera
    Put #N, , ConfigInit
    Close #N
    
    MsgBox "Indexación completada!", vbOKOnly

End Function

Private Function Extract()
On Local Error Resume Next

    If LenB(Dir(App.Path & "\Font.dat", vbArchive)) = 0 Then
        MsgBox "Se requiere Font.dat en el directorio del programa.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    Dim MiCabecera As tCabecera
    Dim ConfigInit As tConfigInit
    Dim N As Integer
    
    Call IniciarCabecera(MiCabecera)
    N = FreeFile
    
    Open App.Path & "\Font.dat" For Binary As #N
        Get #N, , cfonts(1).HeaderInfo
    Close #N
    
    If LenB(Dir(App.Path & "\Config.ini", vbArchive)) <> 0 Then
        Kill App.Path & "\Config.ini"
    End If
    
    ' CONFIG
    Call WriteVar(App.Path & "\Config.ini", "CONFIG", "MostrarTips", ConfigInit.MostrarTips)
    Call WriteVar(App.Path & "\Config.ini", "CONFIG", "NumParticulas", ConfigInit.NumParticulas)
    
    ' DIR
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Multimedia", ConfigInit.DirMultimedia)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Mapas", ConfigInit.DirMapas)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Graficos", ConfigInit.DirGraficos)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Fotos", ConfigInit.DirFotos)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Extras", ConfigInit.DirExtras)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Sonidos", ConfigInit.DirSonidos)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Musicas", ConfigInit.DirMusicas)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Particulas", ConfigInit.DirParticulas)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "GUI", ConfigInit.DirGUI)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Botones", ConfigInit.DirBotones)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Frags", ConfigInit.DirFrags)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Muertes", ConfigInit.DirMuertes)
    
    ' USUARIO
    Call WriteVar(App.Path & "\Config.ini", "USUARIO", "Nombre", ConfigInit.Nombre)
    Call WriteVar(App.Path & "\Config.ini", "USUARIO", "Password", ConfigInit.Password)
    
    MsgBox "Extracción completada!", vbOKOnly

End Function

Private Sub cmdExtract_Click()
    Call Extract
End Sub

Private Sub cmdIndex_Click()
    Call Index
End Sub

Private Sub Form_Load()
    Me.Caption = "Configurar Font.Dat v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub
