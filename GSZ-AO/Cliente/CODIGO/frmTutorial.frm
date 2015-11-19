VERSION 5.00
Begin VB.Form frmTutorial 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   Picture         =   "frmTutorial.frx":0000
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin GSZAOCliente.uAOCheckbox chkNoMostrar 
      Height          =   345
      Left            =   3120
      TabIndex        =   4
      Top             =   6960
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmTutorial.frx":2E93B
   End
   Begin GSZAOCliente.uAOButton cAnterior 
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   6840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      TX              =   "Anterior"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmTutorial.frx":2E999
      PICF            =   "frmTutorial.frx":2E9B5
      PICH            =   "frmTutorial.frx":2E9D1
      PICV            =   "frmTutorial.frx":2E9ED
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
   Begin GSZAOCliente.uAOButton cSiguiente 
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   6840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      TX              =   "Siguiente"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmTutorial.frx":2EA09
      PICF            =   "frmTutorial.frx":2EA25
      PICH            =   "frmTutorial.frx":2EA41
      PICV            =   "frmTutorial.frx":2EA5D
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
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   300
      Left            =   8370
      TabIndex        =   7
      Top             =   45
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   529
      TX              =   "X"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmTutorial.frx":2EA79
      PICF            =   "frmTutorial.frx":2EA95
      PICH            =   "frmTutorial.frx":2EAB1
      PICV            =   "frmTutorial.frx":2EACD
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   390
      Left            =   525
      TabIndex        =   3
      Top             =   435
      Width           =   7725
   End
   Begin VB.Image imgMostrar 
      Height          =   570
      Left            =   3000
      Picture         =   "frmTutorial.frx":2EAE9
      Top             =   6855
      Width           =   2535
   End
   Begin VB.Label lblMensaje 
      BackStyle       =   0  'Transparent
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
      Height          =   5790
      Left            =   525
      TabIndex        =   2
      Top             =   840
      Width           =   7725
   End
   Begin VB.Label lblPagTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   7365
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblPagActual 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   6870
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmTutorial"
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

Private Type tTutorial
    sTitle As String
    sPage As String
End Type

Private Tutorial() As tTutorial
Private NumPages As Long
Private CurrentPage As Long

Private Sub cAnterior_Click()
    Call Audio.PlayWave(SND_CLICK)
    
    If Not cAnterior.Enabled Then Exit Sub
    
    CurrentPage = CurrentPage - 1
    
    If CurrentPage = 1 Then cAnterior.Enabled = False
    
    If Not cSiguiente.Enabled Then cSiguiente.Enabled = True
    
    Call SelectPage(CurrentPage)
    
End Sub



Private Sub cCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)

    bShowTutorial = False 'Mientras no se pueda tildar/destildar para verlo más tarde, esto queda así :P
    Unload Me
End Sub

Private Sub chkNoMostrar_Click()
    Call Audio.PlayWave(SND_CLICK)

    bShowTutorial = Not bShowTutorial
    chkNoMostrar.Checked = Not bShowTutorial

End Sub

Private Sub cSiguiente_Click()
    Call Audio.PlayWave(SND_CLICK)
    
    If Not cSiguiente.Enabled Then Exit Sub
    
    CurrentPage = CurrentPage + 1
    
    ' DEshabilita el boton siguiente si esta en la ultima pagina
    If CurrentPage = NumPages Then cSiguiente.Enabled = False
    
    ' Habilita el boton anterior
    If Not cAnterior.Enabled Then cAnterior.Enabled = True
    
    Call SelectPage(CurrentPage)
    
End Sub

Private Sub Form_Load()

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGUI & "frmTutorial.jpg")
    
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
    
    chkNoMostrar.Checked = Not bShowTutorial
    
    Call LoadTutorial
    
    CurrentPage = 1
    Call SelectPage(CurrentPage)
End Sub


Private Sub LoadTutorial()
On Error Resume Next
    Dim TutorialPath As String
    Dim lPage As Long
    Dim NumLines As Long
    Dim lLine As Long
    Dim sLine As String
    
    TutorialPath = DirExtras & "Tutorial.dat"
    NumPages = Val(GetVar(TutorialPath, "INIT", "NumPags"))
    
    If NumPages > 0 Then
        ReDim Tutorial(1 To NumPages)
        
        ' Cargo paginas
        For lPage = 1 To NumPages
            NumLines = Val(GetVar(TutorialPath, "PAG" & lPage, "NumLines"))
            
            With Tutorial(lPage)
                
                .sTitle = GetVar(TutorialPath, "PAG" & lPage, "Title")
                
                ' Cargo cada linea de la pagina
                For lLine = 1 To NumLines
                    sLine = GetVar(TutorialPath, "PAG" & lPage, "Line" & lLine)
                    .sPage = .sPage & sLine & vbCrLf
                Next lLine
            End With
            
        Next lPage
    End If
    
    lblPagTotal.Caption = NumPages
End Sub

Private Sub SelectPage(ByVal lPage As Long)
'    lblTitulo.Caption = Tutorial(lPage).sTitle
'    lblMensaje.Caption = Tutorial(lPage).sPage
    lblPagActual.Caption = lPage
End Sub
