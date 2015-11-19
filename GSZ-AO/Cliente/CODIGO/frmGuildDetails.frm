VERSION 5.00
Begin VB.Form frmGuildDetails 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   13935
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuildDetails.frx":0000
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   929
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pbLogoClan 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   7200
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   13
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox txtRuta 
      Height          =   285
      Left            =   9000
      TabIndex        =   12
      Top             =   285
      Width           =   4695
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   12360
      TabIndex        =   11
      Top             =   765
      Width           =   1335
   End
   Begin VB.TextBox txtDesc 
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
      Height          =   1500
      Left            =   405
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   420
      Width           =   6015
   End
   Begin VB.TextBox txtCodex1 
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
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
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
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   3615
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
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
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   3990
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
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
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   4365
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
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
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   5
      Top             =   4740
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
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
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   6
      Top             =   5115
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
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
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   7
      Top             =   5490
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
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
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   8
      Top             =   5865
      Width           =   5835
   End
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildDetails.frx":1F228
      PICF            =   "frmGuildDetails.frx":1F244
      PICH            =   "frmGuildDetails.frx":1F260
      PICV            =   "frmGuildDetails.frx":1F27C
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
   Begin GSZAOCliente.uAOButton cConfirmar 
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   6240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Confirmar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildDetails.frx":1F298
      PICF            =   "frmGuildDetails.frx":1F2B4
      PICH            =   "frmGuildDetails.frx":1F2D0
      PICV            =   "frmGuildDetails.frx":1F2EC
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
End
Attribute VB_Name = "frmGuildDetails"
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

Private Const MAX_DESC_LENGTH As Integer = 520
Private Const MAX_CODEX_LENGTH As Integer = 100

Private Sub cCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
End Sub

Private Sub cConfirmar_Click()
On Error GoTo Fallo
    Call Audio.PlayWave(SND_CLICK)
    Dim fdesc As String
    Dim Codex() As String
    Dim K As Byte
    Dim Cont As Byte

    fdesc = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)

    Cont = 0
    For K = 0 To txtCodex1.UBound
        If LenB(txtCodex1(K).Text) <> 0 Then Cont = Cont + 1
    Next K
    
    If Cont < 4 Then
        MsgBox "Debes definir al menos cuatro mandamientos."
        Exit Sub
    End If
                
    ReDim Codex(txtCodex1.UBound) As String
    For K = 0 To txtCodex1.UBound
        Codex(K) = txtCodex1(K)
    Next K

    If CreandoClan Then
        Call WriteCreateNewGuild(fdesc, ClanName, Site, Codex, txtRuta.Text)
    Else
        Call WriteClanCodexUpdate(fdesc, Codex, txtRuta.Text)
    End If

    CreandoClan = False
    Unload Me
    Exit Sub
Fallo:
    Call MsgBox("Error al intentar definir el codex.", vbCritical + vbOKOnly)
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGUI & "frmGuildDetails.jpg")

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

End Sub

Private Sub txtCodex1_Change(Index As Integer)
    If Len(txtCodex1.Item(Index).Text) > MAX_CODEX_LENGTH Then _
        txtCodex1.Item(Index).Text = Left$(txtCodex1.Item(Index).Text, MAX_CODEX_LENGTH)
End Sub

Private Sub txtDesc_Change()
    If Len(txtDesc.Text) > MAX_DESC_LENGTH Then _
        txtDesc.Text = Left$(txtDesc.Text, MAX_DESC_LENGTH)
End Sub

Private Sub cmdCargar_Click()
Dim pClan As Picture

pbLogoClan.AutoRedraw = True

If Not txtRuta.Text = vbNullString Then
    Set pClan = LoadPicture(txtRuta.Text)
Else
    Set pClan = VB.LoadPicture(DirGraficos & "Sin avatar - Casco.jpg")
End If

If Not ResizePicture(pbLogoClan, pClan) Then
    Set pClan = VB.LoadPicture(DirGraficos & "Sin avatar - Casco.jpg")
    pbLogoClan.Picture = pClan
End If
End Sub


Private Function LoadPicture(ByVal strFileName As String) As Picture
Dim IID As TGUID

        With IID
                .data1 = &H7BF80980
                .data2 = &HBF32
                .Data3 = &H101A
                .Data4(0) = &H8B
                .Data4(1) = &HBB
                .Data4(2) = &H0
                .Data4(3) = &HAA
                .Data4(4) = &H0
                .Data4(5) = &H30
                .Data4(6) = &HC
                .Data4(7) = &HAB
        End With
        
On Error GoTo ERR_LINE

        OleLoadPicturePath StrPtr(strFileName), 0&, 0&, 0&, IID, LoadPicture
        
Exit Function

ERR_LINE:
    Set LoadPicture = VB.LoadPicture(strFileName)
    
End Function

