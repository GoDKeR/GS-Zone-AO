VERSION 5.00
Begin VB.Form frmGuildMember 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   Picture         =   "frmGuildMember.frx":0000
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMiembros 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2565
      Left            =   3075
      TabIndex        =   3
      Top             =   675
      Width           =   2610
   End
   Begin VB.ListBox lstClanes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2565
      Left            =   195
      TabIndex        =   2
      Top             =   690
      Width           =   2610
   End
   Begin VB.TextBox txtSearch 
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
      Height          =   225
      Left            =   225
      TabIndex        =   1
      Top             =   3630
      Width           =   2550
   End
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   4920
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildMember.frx":E796
      PICF            =   "frmGuildMember.frx":E7B2
      PICH            =   "frmGuildMember.frx":E7CE
      PICV            =   "frmGuildMember.frx":E7EA
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
   Begin GSZAOCliente.uAOButton cNoticias 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      TX              =   "Noticias"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildMember.frx":E806
      PICF            =   "frmGuildMember.frx":E822
      PICH            =   "frmGuildMember.frx":E83E
      PICV            =   "frmGuildMember.frx":E85A
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
   Begin GSZAOCliente.uAOButton cDetalles 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      TX              =   "Detalles"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildMember.frx":E876
      PICF            =   "frmGuildMember.frx":E892
      PICH            =   "frmGuildMember.frx":E8AE
      PICV            =   "frmGuildMember.frx":E8CA
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
   Begin VB.Label lblCantMiembros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Height          =   195
      Left            =   4635
      TabIndex        =   0
      Top             =   3510
      Width           =   360
   End
End
Attribute VB_Name = "frmGuildMember"
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

Private Sub cDetalles_Click()
    Call Audio.PlayWave(SND_CLICK)
    If lstClanes.ListIndex = -1 Then Exit Sub
    frmGuildBrief.EsLeader = False
    Call WriteGuildRequestDetails(lstClanes.List(lstClanes.ListIndex))
End Sub

Private Sub cNoticias_Click()
    Call Audio.PlayWave(SND_CLICK)
    bShowGuildNews = True ' 0.13.3
    Call WriteShowGuildNews
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Me.Picture = LoadPicture(DirGUI & "frmGuildMember.jpg")
    
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

Private Sub txtSearch_Change()
    Call FiltrarListaClanes(txtSearch.Text)
End Sub

Private Sub txtSearch_GotFocus()
    With txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Public Sub FiltrarListaClanes(ByRef sCompare As String)
    Dim lIndex As Long
    
    If UBound(GuildNames) <> 0 Then
        With lstClanes
            'Limpio la lista
            .Clear
            
            .Visible = False
            
            ' Recorro los arrays
            For lIndex = 0 To UBound(GuildNames)
                ' Si coincide con los patrones
                If InStr(1, UCase$(GuildNames(lIndex)), UCase$(sCompare)) Then
                    ' Lo agrego a la lista
                    .AddItem GuildNames(lIndex)
                End If
            Next lIndex
            
            .Visible = True
        End With
    End If

End Sub

