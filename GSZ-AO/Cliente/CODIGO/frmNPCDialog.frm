VERSION 5.00
Begin VB.Form frmNPCDialog 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmNPCDialog.frx":0000
   ScaleHeight     =   163
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   309
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmNPCDialog.frx":D127
      PICF            =   "frmNPCDialog.frx":D143
      PICH            =   "frmNPCDialog.frx":D15F
      PICV            =   "frmNPCDialog.frx":D17B
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
   Begin GSZAOCliente.uAOButton cOp 
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   435
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   556
      TX              =   "Opción"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmNPCDialog.frx":D197
      PICF            =   "frmNPCDialog.frx":D1B3
      PICH            =   "frmNPCDialog.frx":D1CF
      PICV            =   "frmNPCDialog.frx":D1EB
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
   Begin GSZAOCliente.uAOButton cOp 
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   765
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   556
      TX              =   "Opción"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmNPCDialog.frx":D207
      PICF            =   "frmNPCDialog.frx":D223
      PICH            =   "frmNPCDialog.frx":D23F
      PICV            =   "frmNPCDialog.frx":D25B
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
   Begin GSZAOCliente.uAOButton cOp 
      Height          =   315
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   1095
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   556
      TX              =   "Opción"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmNPCDialog.frx":D277
      PICF            =   "frmNPCDialog.frx":D293
      PICH            =   "frmNPCDialog.frx":D2AF
      PICV            =   "frmNPCDialog.frx":D2CB
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
   Begin GSZAOCliente.uAOButton cOp 
      Height          =   315
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   1425
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   556
      TX              =   "Opción"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmNPCDialog.frx":D2E7
      PICF            =   "frmNPCDialog.frx":D303
      PICH            =   "frmNPCDialog.frx":D31F
      PICV            =   "frmNPCDialog.frx":D33B
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
End
Attribute VB_Name = "frmNPCDialog"
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

Private Sub cCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
End Sub

Private Sub Form_Load()

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGUI & "frmUserRequest.jpg")

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

