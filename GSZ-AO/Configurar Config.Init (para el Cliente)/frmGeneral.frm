VERSION 5.00
Begin VB.Form frmGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Config.Init "
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
      Caption         =   "Indexar Config.Init"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer Config.Init"
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

Private Type tConfigInit
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

Private Sub IniciarCabecera(ByRef Cabecera As tCabecera)
'*************************************************
'Author: ^[GS]^
'Last modified: 05/06/2012
'*************************************************
    Cabecera.Desc = "GS-Zone Argentum Online MOD - Copyright GS-Zone 2012 - info@gs-zone.org - Original by Pablo Marquez "
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub


Private Function Index()
'*************************************************
'Author: ^[GS]^
'Last modified: 25/07/2012 - ^[GS]^
'*************************************************
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
    ConfigInit.IndiceGraficos = GetVar(App.Path & "\Config.Ini", "CONFIG", "IndiceGraficos")

    ' DIR
    ConfigInit.DirMultimedia = GetVar(App.Path & "\Config.Ini", "DIR", "Multimedia")
    ConfigInit.DirMapas = GetVar(App.Path & "\Config.Ini", "DIR", "Mapas")
    ConfigInit.DirGraficos = GetVar(App.Path & "\Config.Ini", "DIR", "Graficos")
    ConfigInit.DirFotos = GetVar(App.Path & "\Config.Ini", "DIR", "Fotos")
    ConfigInit.DirExtras = GetVar(App.Path & "\Config.Ini", "DIR", "Extras")
    ConfigInit.DirSonidos = GetVar(App.Path & "\Config.Ini", "DIR", "Sonidos")
    ConfigInit.DirMusicas = GetVar(App.Path & "\Config.Ini", "DIR", "Musicas")
    ConfigInit.DirParticulas = GetVar(App.Path & "\Config.Ini", "DIR", "Particulas")
    ConfigInit.DirCursores = GetVar(App.Path & "\Config.Ini", "DIR", "Cursores")
    ConfigInit.DirGUI = GetVar(App.Path & "\Config.Ini", "DIR", "GUI")
    ConfigInit.DirBotones = GetVar(App.Path & "\Config.Ini", "DIR", "Botones")
    ConfigInit.DirFrags = GetVar(App.Path & "\Config.Ini", "DIR", "Frags")
    ConfigInit.DirMuertes = GetVar(App.Path & "\Config.Ini", "DIR", "Muertes")
    
    ' USUARIO
    ConfigInit.Nombre = GetVar(App.Path & "\Config.Ini", "USUARIO", "Nombre")
    ConfigInit.Password = GetVar(App.Path & "\Config.Ini", "USUARIO", "Password")
    ConfigInit.Recordar = Val(GetVar(App.Path & "\Config.Ini", "USUARIO", "Recordar"))
    
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
'*************************************************
'Author: ^[GS]^
'Last modified: 25/07/2012 - ^[GS]^
'*************************************************
On Local Error Resume Next

    If LenB(Dir(App.Path & "\Config.Init", vbArchive)) = 0 Then
        MsgBox "Se requiere Config.Init en el directorio del programa.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    Dim MiCabecera As tCabecera
    Dim ConfigInit As tConfigInit
    Dim N As Integer
    
    Call IniciarCabecera(MiCabecera)
    N = FreeFile
    
    Open App.Path & "\Config.Init" For Binary As #N
    Get #N, , MiCabecera
    Get #N, , ConfigInit
    Close #N
    
    If LenB(Dir(App.Path & "\Config.ini", vbArchive)) <> 0 Then
        Kill App.Path & "\Config.ini"
    End If
    
    ' CONFIG
    Call WriteVar(App.Path & "\Config.ini", "CONFIG", "MostrarTips", ConfigInit.MostrarTips)
    Call WriteVar(App.Path & "\Config.ini", "CONFIG", "NumParticulas", ConfigInit.NumParticulas)
    Call WriteVar(App.Path & "\Config.ini", "CONFIG", "IndiceGraficos", ConfigInit.IndiceGraficos)
    
    ' DIR
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Multimedia", ConfigInit.DirMultimedia)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Mapas", ConfigInit.DirMapas)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Graficos", ConfigInit.DirGraficos)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Fotos", ConfigInit.DirFotos)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Extras", ConfigInit.DirExtras)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Sonidos", ConfigInit.DirSonidos)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Musicas", ConfigInit.DirMusicas)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Particulas", ConfigInit.DirParticulas)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Cursores", ConfigInit.DirCursores)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "GUI", ConfigInit.DirGUI)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Botones", ConfigInit.DirBotones)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Frags", ConfigInit.DirFrags)
    Call WriteVar(App.Path & "\Config.ini", "DIR", "Muertes", ConfigInit.DirMuertes)
    
    ' USUARIO
    Call WriteVar(App.Path & "\Config.ini", "USUARIO", "Nombre", ConfigInit.Nombre)
    Call WriteVar(App.Path & "\Config.ini", "USUARIO", "Password", ConfigInit.Password)
    Call WriteVar(App.Path & "\Config.ini", "USUARIO", "Recordar", ConfigInit.Recordar)
    
    MsgBox "Extracción completada!", vbOKOnly

End Function

Private Sub cmdExtract_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 05/06/2012
'*************************************************
    Call Extract
End Sub

Private Sub cmdIndex_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 05/06/2012
'*************************************************
    Call Index
End Sub

Private Sub Form_Load()
'*************************************************
'Author: ^[GS]^
'Last modified: 05/06/2012
'*************************************************
    Me.Caption = "Configurar Config.Init v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub
