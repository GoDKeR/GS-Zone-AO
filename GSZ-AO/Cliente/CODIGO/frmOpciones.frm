VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmOpciones 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOpciones.frx":0152
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin GSZAOCliente.uAOCheckbox ChkMusica 
      Height          =   225
      Left            =   435
      TabIndex        =   11
      Top             =   990
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   397
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmOpciones.frx":20977
   End
   Begin VB.TextBox txtCantMensajes 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2310
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "5"
      Top             =   2415
      Width           =   255
   End
   Begin VB.TextBox txtLevel 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "40"
      Top             =   4395
      Width           =   255
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Index           =   0
      Left            =   1380
      TabIndex        =   0
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Max             =   100
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Index           =   1
      Left            =   1380
      TabIndex        =   1
      Top             =   1260
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   10
      Max             =   100
      TickStyle       =   3
   End
   Begin GSZAOCliente.uAOButton cAceptar 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   6555
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      TX              =   "Aceptar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmOpciones.frx":209D5
      PICF            =   "frmOpciones.frx":209F1
      PICH            =   "frmOpciones.frx":20A0D
      PICV            =   "frmOpciones.frx":20A29
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cTutorial 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   6120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      TX              =   "Tutorial"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmOpciones.frx":20A45
      PICF            =   "frmOpciones.frx":20A61
      PICH            =   "frmOpciones.frx":20A7D
      PICV            =   "frmOpciones.frx":20A99
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cCambiarContrasena 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   5640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      TX              =   "Cambiar Contraseña"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmOpciones.frx":20AB5
      PICF            =   "frmOpciones.frx":20AD1
      PICH            =   "frmOpciones.frx":20AED
      PICV            =   "frmOpciones.frx":20B09
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cMensajesPersonalizados 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   5160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      TX              =   "Mensajes Personalizados"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmOpciones.frx":20B25
      PICF            =   "frmOpciones.frx":20B41
      PICH            =   "frmOpciones.frx":20B5D
      PICV            =   "frmOpciones.frx":20B79
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cConfigurarTeclas 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   5160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      TX              =   "Configurar Teclas"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmOpciones.frx":20B95
      PICF            =   "frmOpciones.frx":20BB1
      PICH            =   "frmOpciones.frx":20BCD
      PICV            =   "frmOpciones.frx":20BE9
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cVerMapa 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   5640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      TX              =   "Ver Mapa"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmOpciones.frx":20C05
      PICF            =   "frmOpciones.frx":20C21
      PICH            =   "frmOpciones.frx":20C3D
      PICV            =   "frmOpciones.frx":20C59
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cSoporte 
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   6120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      TX              =   "Soporte"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmOpciones.frx":20C75
      PICF            =   "frmOpciones.frx":20C91
      PICH            =   "frmOpciones.frx":20CAD
      PICV            =   "frmOpciones.frx":20CC9
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOCheckbox ChkSonidos 
      Height          =   225
      Left            =   435
      TabIndex        =   12
      Top             =   1260
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   397
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmOpciones.frx":20CE5
   End
   Begin GSZAOCliente.uAOCheckbox ChkEfectosSonidos 
      Height          =   225
      Left            =   435
      TabIndex        =   13
      Top             =   1545
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   397
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmOpciones.frx":20D43
   End
   Begin GSZAOCliente.uAOCheckbox ChkClanConsola 
      Height          =   225
      Left            =   435
      TabIndex        =   14
      Top             =   2430
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   397
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmOpciones.frx":20DA1
   End
   Begin GSZAOCliente.uAOCheckbox ChkClanPantalla 
      Height          =   225
      Left            =   1950
      TabIndex        =   15
      Top             =   2430
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   397
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmOpciones.frx":20DFF
   End
   Begin GSZAOCliente.uAOCheckbox ChkClanMostrarNoticias 
      Height          =   225
      Left            =   435
      TabIndex        =   16
      Top             =   3315
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   397
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmOpciones.frx":20E5D
   End
   Begin GSZAOCliente.uAOCheckbox ChkActivarFragshooter 
      Height          =   225
      Left            =   435
      TabIndex        =   17
      Top             =   4110
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   397
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmOpciones.frx":20EBB
   End
   Begin GSZAOCliente.uAOCheckbox ChkFragAlMorir 
      Height          =   225
      Left            =   435
      TabIndex        =   18
      Top             =   4740
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   397
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmOpciones.frx":20F19
   End
   Begin GSZAOCliente.uAOCheckbox ChkFragRequiereNivel 
      Height          =   225
      Left            =   435
      TabIndex        =   19
      Top             =   4425
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   397
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmOpciones.frx":20F77
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private clsFormulario As clsFormMovementManager

Private picCheckBox As Picture

Private bMusicActivated As Boolean
Private bSoundActivated As Boolean
Private bSoundEffectsActivated As Boolean

Private loading As Boolean

Private Sub cAceptar_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Unload Me
    If Connected = True Then
        If frmMain.Visible = True Then
            frmMain.SetFocus
        End If
    Else
        If frmConnect.Visible = True Then
            frmConnect.SetFocus
        End If
    End If
End Sub

Private Sub cCambiarContrasena_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call frmNewPassword.Show(vbModal, Me)
End Sub

Private Sub cConfigurarTeclas_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub ChkActivarFragshooter_Click()
    Call Audio.PlayWave(SND_CLICK)
    ClientAOSetup.bActive = ChkActivarFragshooter.Checked
    ChkFragAlMorir.Enabled = ChkActivarFragshooter.Checked
    ChkFragRequiereNivel.Enabled = ChkActivarFragshooter.Checked
    If ChkFragRequiereNivel.Checked = True Then
        txtLevel.Enabled = ChkActivarFragshooter.Checked
    End If
End Sub

Private Sub ChkClanConsola_Click()
    Call Audio.PlayWave(SND_CLICK)
    If ChkClanConsola.Checked = True Then
        DialogosClanes.Activo = False
        ChkClanPantalla.Checked = False
    Else
        DialogosClanes.Activo = True
        ChkClanPantalla.Checked = True
    End If
End Sub

Private Sub ChkClanPantalla_Click()
    Call Audio.PlayWave(SND_CLICK)
    If ChkClanPantalla.Checked = True Then
        DialogosClanes.Activo = True
        ChkClanConsola.Checked = False
        txtCantMensajes.Enabled = True
    Else
        DialogosClanes.Activo = False
        ChkClanConsola.Checked = True
        txtCantMensajes.Enabled = False
    End If
End Sub

Private Sub ChkEfectosSonidos_Click()
    If loading Then Exit Sub
    Call Audio.PlayWave(SND_CLICK)
    bSoundEffectsActivated = Not bSoundEffectsActivated
    Audio.SoundEffectsActivated = bSoundEffectsActivated
    ChkEfectosSonidos.Checked = bSoundEffectsActivated
End Sub

Private Sub ChkFragAlMorir_Click()
    Call Audio.PlayWave(SND_CLICK)
    ClientAOSetup.bDie = ChkFragAlMorir.Checked
End Sub

Private Sub ChkFragRequiereNivel_Click()
    Call Audio.PlayWave(SND_CLICK)
    ClientAOSetup.bKill = ChkFragRequiereNivel.Checked
    txtLevel.Enabled = ChkFragRequiereNivel.Checked
End Sub

Private Sub ChkClanMostrarNoticias_Click()
    Call Audio.PlayWave(SND_CLICK)
    ClientAOSetup.bGuildNews = ChkClanMostrarNoticias.Checked
End Sub

Private Sub ChkMusica_Click()
    If loading Then Exit Sub
    
    Call Audio.PlayWave(SND_CLICK)
    bMusicActivated = ChkMusica.Checked
    If Not bMusicActivated Then
        Audio.MusicActivated = False
        Slider1(0).Enabled = False
    Else
        If Not Audio.MusicActivated Then  'Prevent the music from reloading
            Audio.MusicActivated = True
            Slider1(0).Enabled = True
            Slider1(0).Value = Audio.MusicVolume
        End If
    End If

End Sub

Private Sub ChkSonidos_Click()
    If loading Then Exit Sub
    Call Audio.PlayWave(SND_CLICK)
    bSoundActivated = ChkSonidos.Checked
    If Not bSoundActivated Then
        Audio.SoundActivated = False
        RainBufferIndex = 0
        frmMain.IsPlaying = PlayLoop.plNone
        Slider1(1).Enabled = False
    Else
        Audio.SoundActivated = True
        Slider1(1).Enabled = True
        Slider1(1).Value = Audio.SoundVolume
    End If
End Sub

Private Sub cSoporte_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://www.gs-zone.org/gs_zone_ao_f2j.html", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub cMensajesPersonalizados_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call frmMessageTxt.Show(vbModeless, Me)
End Sub

Private Sub cTutorial_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    frmTutorial.Show vbModeless
End Sub

Private Sub cVerMapa_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call frmMapa.Show(vbModal, Me)
End Sub

Private Sub txtCantMensajes_Change()
    txtCantMensajes.Text = Val(txtCantMensajes.Text)
    If txtCantMensajes.Text > 0 Then
        DialogosClanes.CantidadDialogos = txtCantMensajes.Text
    Else
        txtCantMensajes.Text = 5
    End If
End Sub

Private Sub txtLevel_Change()
    If Not IsNumeric(txtLevel) Then txtLevel = 0
    txtLevel = Trim$(txtLevel)
    ClientAOSetup.byMurderedLevel = CByte(txtLevel)
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGUI & "frmOpciones.jpg")
    
    Dim cControl As Control
    For Each cControl In Me.Controls
        If TypeOf cControl Is uAOButton Then
            cControl.PictureEsquina = LoadPicture(ImgRequest(DirButtons & sty_bEsquina))
            cControl.PictureFondo = LoadPicture(ImgRequest(DirButtons & sty_bFondo))
            cControl.PictureHorizontal = LoadPicture(ImgRequest(DirButtons & sty_bHorizontal))
            cControl.PictureVertical = LoadPicture(ImgRequest(DirButtons & sty_bVertical))
        ElseIf TypeOf cControl Is uAOCheckbox Then
            cControl.Picture = LoadPicture(ImgRequest(DirButtons & sty_cCheckbox2))
        End If
    Next
    
    loading = True      'Prevent sounds when setting check's values
    Call LoadUserConfig
    loading = False     'Enable sounds when setting check's values
    
    If Connected = True Then
        cMensajesPersonalizados.Enabled = True
        cCambiarContrasena.Enabled = True
    Else
        cMensajesPersonalizados.Enabled = False
        cCambiarContrasena.Enabled = False
    End If
End Sub

Private Sub LoadUserConfig()
    ' Load music config
    bMusicActivated = Audio.MusicActivated
    Slider1(0).Enabled = bMusicActivated
    ChkMusica.Checked = bMusicActivated
    
    If bMusicActivated Then
        Slider1(0).Value = Audio.MusicVolume
    End If
    
    ' Load Sound config
    bSoundActivated = Audio.SoundActivated
    Slider1(1).Enabled = bSoundActivated
    ChkSonidos.Checked = bSoundActivated
    
    If bSoundActivated Then
        Slider1(1).Value = Audio.SoundVolume
    End If
    
    ' Load Sound Effects config
    bSoundEffectsActivated = Audio.SoundEffectsActivated
    ChkEfectosSonidos.Checked = bSoundEffectsActivated
    
    ' Clanes
    txtCantMensajes.Text = CStr(DialogosClanes.CantidadDialogos)
    ChkClanPantalla.Checked = DialogosClanes.Activo
    ChkClanMostrarNoticias.Checked = ClientAOSetup.bGuildNews
    
    ' Frags
    ChkActivarFragshooter.Checked = ClientAOSetup.bActive
    ChkFragRequiereNivel.Checked = ClientAOSetup.bKill
    ChkFragAlMorir.Checked = ClientAOSetup.bDie
    txtLevel = ClientAOSetup.byMurderedLevel
End Sub

Private Sub Slider1_Change(Index As Integer)
    Select Case Index
        Case 0
            Audio.MusicVolume = Slider1(0).Value
        Case 1
            Audio.SoundVolume = Slider1(1).Value
    End Select
End Sub

Private Sub Slider1_Scroll(Index As Integer)
    Select Case Index
        Case 0
            Audio.MusicVolume = Slider1(0).Value
        Case 1
            Audio.SoundVolume = Slider1(1).Value
    End Select
End Sub
