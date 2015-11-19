VERSION 5.00
Begin VB.Form frmCambiaMotd 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   """ZMOTD"""
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5175
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCambiaMotd.frx":0000
   ScaleHeight     =   361
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMotd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2250
      Left            =   435
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   795
      Width           =   4290
   End
   Begin GSZAOCliente.uAOButton cAzul 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Azul"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":13380
      PICF            =   "frmCambiaMotd.frx":1339C
      PICH            =   "frmCambiaMotd.frx":133B8
      PICV            =   "frmCambiaMotd.frx":133D4
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
   Begin GSZAOCliente.uAOButton cRojo 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Rojo"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":133F0
      PICF            =   "frmCambiaMotd.frx":1340C
      PICH            =   "frmCambiaMotd.frx":13428
      PICV            =   "frmCambiaMotd.frx":13444
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
   Begin GSZAOCliente.uAOButton cBlanco 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Blanco"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":13460
      PICF            =   "frmCambiaMotd.frx":1347C
      PICH            =   "frmCambiaMotd.frx":13498
      PICV            =   "frmCambiaMotd.frx":134B4
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
   Begin GSZAOCliente.uAOButton cGris 
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Gris"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":134D0
      PICF            =   "frmCambiaMotd.frx":134EC
      PICH            =   "frmCambiaMotd.frx":13508
      PICV            =   "frmCambiaMotd.frx":13524
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
   Begin GSZAOCliente.uAOButton cAmarillo 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Amarillo"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":13540
      PICF            =   "frmCambiaMotd.frx":1355C
      PICH            =   "frmCambiaMotd.frx":13578
      PICV            =   "frmCambiaMotd.frx":13594
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cMorado 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Morado"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":135B0
      PICF            =   "frmCambiaMotd.frx":135CC
      PICH            =   "frmCambiaMotd.frx":135E8
      PICV            =   "frmCambiaMotd.frx":13604
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
   Begin GSZAOCliente.uAOButton cVerde 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Verde"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":13620
      PICF            =   "frmCambiaMotd.frx":1363C
      PICH            =   "frmCambiaMotd.frx":13658
      PICV            =   "frmCambiaMotd.frx":13674
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
   Begin GSZAOCliente.uAOButton cMarron 
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Marron"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":13690
      PICF            =   "frmCambiaMotd.frx":136AC
      PICH            =   "frmCambiaMotd.frx":136C8
      PICV            =   "frmCambiaMotd.frx":136E4
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
   Begin GSZAOCliente.uAOButton cAceptar 
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   4680
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      TX              =   "Aceptar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":13700
      PICF            =   "frmCambiaMotd.frx":1371C
      PICH            =   "frmCambiaMotd.frx":13738
      PICV            =   "frmCambiaMotd.frx":13754
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
   Begin GSZAOCliente.uAOCheckbox chkNegrita 
      Height          =   345
      Left            =   1080
      TabIndex        =   10
      Top             =   4320
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmCambiaMotd.frx":13770
   End
   Begin GSZAOCliente.uAOCheckbox chkCursiva 
      Height          =   345
      Left            =   3000
      TabIndex        =   12
      Top             =   4320
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmCambiaMotd.frx":137CE
   End
   Begin VB.Label lCursiva 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cursiva"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lNegrita 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Negrita"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   4320
      Width           =   1095
   End
End
Attribute VB_Name = "frmCambiaMotd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmCambiarMotd.frm
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Private clsFormulario As clsFormMovementManager

Private yNegrita As Byte
Private yCursiva As Byte

Private Sub cAceptar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Dim T() As String
    Dim i As Long, N As Long, Pos As Long
    
    If Len(txtMotd.Text) >= 2 Then
        If Right$(txtMotd.Text, 2) = vbCrLf Then txtMotd.Text = Left$(txtMotd.Text, Len(txtMotd.Text) - 2)
    End If
    
    T = Split(txtMotd.Text, vbCrLf)
    
    For i = LBound(T) To UBound(T)
        N = 0
        Pos = InStr(1, T(i), "~")
        Do While Pos > 0 And Pos < Len(T(i))
            N = N + 1
            Pos = InStr(Pos + 1, T(i), "~")
        Loop
        If N <> 5 Then
            MsgBox "Error en el formato de la linea " & i + 1 & "."
            Exit Sub
        End If
    Next i
    
    Call WriteSetMOTD(txtMotd.Text)
    Unload Me
End Sub

Private Sub cAmarillo_Click()
    Call Audio.PlayWave(SND_CLICK)
    txtMotd.Text = txtMotd & "~244~244~0~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub cAzul_Click()
    Call Audio.PlayWave(SND_CLICK)
    txtMotd.Text = txtMotd & "~50~70~250~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub cBlanco_Click()
    Call Audio.PlayWave(SND_CLICK)
    txtMotd.Text = txtMotd & "~255~255~255~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub cGris_Click()
    Call Audio.PlayWave(SND_CLICK)
    txtMotd.Text = txtMotd & "~157~157~157~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub cMarron_Click()
    Call Audio.PlayWave(SND_CLICK)
    txtMotd.Text = txtMotd & "~97~58~31~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub cMorado_Click()
    Call Audio.PlayWave(SND_CLICK)
    txtMotd.Text = txtMotd & "~128~0~128~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub cRojo_Click()
    Call Audio.PlayWave(SND_CLICK)
    txtMotd.Text = txtMotd & "~255~0~0~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub cVerde_Click()
    Call Audio.PlayWave(SND_CLICK)
    txtMotd.Text = txtMotd & "~23~104~26~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub chkCursiva_Click()
    Call Audio.PlayWave(SND_CLICK)
    yCursiva = IIf(chkCursiva.Checked = True, 1, 0)
End Sub

Private Sub chkNegrita_Click()
    Call Audio.PlayWave(SND_CLICK)
    yNegrita = IIf(chkNegrita.Checked = True, 1, 0)
End Sub

Private Sub lCursiva_Click()
    chkCursiva.Checked = Not chkCursiva.Checked
    Call chkCursiva_Click
End Sub

Private Sub lNegrita_Click()
    chkNegrita.Checked = Not chkNegrita.Checked
    Call chkNegrita_Click
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGUI & "frmCambiaMotd.jpg")
    
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
    
End Sub

