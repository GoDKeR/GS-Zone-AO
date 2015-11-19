VERSION 5.00
Begin VB.Form frmNotas 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Notas"
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   Picture         =   "frmNotas.frx":0000
   ScaleHeight     =   7635
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNotas 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6240
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   420
      Width           =   7575
   End
   Begin GSZAOCliente.uAOButton cCerrarSinGuardar 
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   6840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      TX              =   "Cerrar sin guardar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmNotas.frx":13E6E
      PICF            =   "frmNotas.frx":13E8A
      PICH            =   "frmNotas.frx":13EA6
      PICV            =   "frmNotas.frx":13EC2
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   300
      Left            =   8320
      TabIndex        =   1
      Top             =   100
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   529
      TX              =   "X"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmNotas.frx":13EDE
      PICF            =   "frmNotas.frx":13EFA
      PICH            =   "frmNotas.frx":13F16
      PICV            =   "frmNotas.frx":13F32
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cAbrir 
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      TX              =   "Re-abrir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmNotas.frx":13F4E
      PICF            =   "frmNotas.frx":13F6A
      PICH            =   "frmNotas.frx":13F86
      PICV            =   "frmNotas.frx":13FA2
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cGuardar 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      TX              =   "Guardar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmNotas.frx":13FBE
      PICF            =   "frmNotas.frx":13FDA
      PICH            =   "frmNotas.frx":13FF6
      PICV            =   "frmNotas.frx":14012
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cMostrar 
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      TX              =   "Ocultar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmNotas.frx":1402E
      PICF            =   "frmNotas.frx":1404A
      PICH            =   "frmNotas.frx":14066
      PICV            =   "frmNotas.frx":14082
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmNotas"
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
Private sFilename As String

Private Sub cAbrir_Click()
    cAbrir.Enabled = False
    Call OpenNotes
    Call Audio.PlayWave(SND_CLICK)
    cAbrir.Enabled = True
End Sub

Private Sub cCerrarSinGuardar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
End Sub

Private Sub cGuardar_Click()
    cGuardar.Enabled = False
    Call SaveNotes
    Call Audio.PlayWave(SND_CLICK)
    cGuardar.Enabled = True
End Sub

Private Sub cCerrar_Click()
    Call SaveNotes
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
End Sub

Public Sub MostrarNotas()
    If cMostrar.Caption = "Ocultar" Then
        Me.Height = 540
        Me.Width = 4855
        Me.Picture = LoadPicture(DirGUI & "frmNotasMini.jpg")
        cCerrar.Left = 4420
        txtNotas.Visible = False
        cMostrar.Caption = "Mostrar"
    Else
        Me.Picture = LoadPicture(DirGUI & "frmNotas.jpg")
        txtNotas.Visible = True
        cCerrar.Left = 8320
        Me.Height = 7635
        Me.Width = 8745
        cMostrar.Caption = "Ocultar"
    End If
End Sub

Private Sub cMostrar_Click()
    Call MostrarNotas
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub cOcultar_Click()

End Sub

Private Sub Form_Load()

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGUI & "frmNotas.jpg")
    
    Dim cControl As Control
    For Each cControl In Me.Controls
        If TypeOf cControl Is uAOButton Then
            cControl.PictureEsquina = LoadPicture(ImgRequest(DirButtons & sty_bEsquina))
            cControl.PictureFondo = LoadPicture(ImgRequest(DirButtons & sty_bFondo))
            cControl.PictureHorizontal = LoadPicture(ImgRequest(DirButtons & sty_bHorizontal))
            cControl.PictureVertical = LoadPicture(ImgRequest(DirButtons & sty_bVertical))
        ElseIf TypeOf cControl Is uAOCheckbox Then
            cControl.Picture = LoadPicture(ImgRequest(DirButtons & sty_cCheckbox))
        End If
    Next
    
    sFilename = App.Path & "\notas-" & UserName & ".txt"
    Call OpenNotes
    
End Sub

Private Sub OpenNotes()
    If FileExist(sFilename, vbArchive) Then
        Dim rPos As Long
        Dim sFileText As String
        Dim iFileNo As Integer
        rPos = txtNotas.SelStart
        iFileNo = FreeFile
        txtNotas = vbNullString
        Open sFilename For Binary As #iFileNo
            sFileText = Space$(LOF(iFileNo))
            Get #iFileNo, , sFileText
            sFileText = SXor(sFileText, LCase(UserName)) ' Realmente no hay necesidad de complicarsela tanto...
            txtNotas.Text = sFileText
        Close #iFileNo
        If Len(txtNotas) > rPos Then
            txtNotas.SelStart = rPos
        End If
    End If
End Sub

Private Sub SaveNotes()
    If FileExist(sFilename, vbArchive) Then
        Call Kill(sFilename)
    End If
    Dim sFileText As String
    Dim iFileNo As Integer
    iFileNo = FreeFile
    sFileText = vbNullString
    sFileText = txtNotas.Text
    Open sFilename For Binary As #iFileNo
        sFileText = SXor(sFileText, LCase(UserName)) ' Realmente no hay necesidad de complicarsela tanto...
        Put #iFileNo, , sFileText
    Close #iFileNo
End Sub
