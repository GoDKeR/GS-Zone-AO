VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMapa.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   375
      Left            =   8080
      TabIndex        =   1
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      TX              =   "X"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMapa.frx":6A66A
      PICF            =   "frmMapa.frx":6A686
      PICH            =   "frmMapa.frx":6A6A2
      PICV            =   "frmMapa.frx":6A6BE
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
   Begin VB.Image imgToogleMap 
      Height          =   255
      Index           =   1
      Left            =   3840
      MousePointer    =   99  'Custom
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgToogleMap 
      Height          =   255
      Index           =   0
      Left            =   3960
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMapa.frx":6A6DA
      BeginProperty Font 
         Name            =   "Morpheus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   8040
      Width           =   8295
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Enum eMaps
    ieGeneral
    ieDungeon
End Enum

Private picMaps(1) As Picture

Private CurrentMap As eMaps

''
' This form is used to show the world map.
' It has two levels. The world map and the dungeons map.
' You can toggle between them pressing the arrows
'
' @file     frmMapa.frm
' @author Marco Vanotti (MarKoxX) marcovanotti15@gmail.com
' @version 1.0.0
' @date 20080724

''
' Checks what Key is down. If the key is const vbKeyDown or const vbKeyUp, it toggles the maps, else the form unloads.
'
' @param KeyCode Specifies the key pressed
' @param Shift Specifies if Shift Button is pressed
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 24/07/08
'
'*************************************************

    Select Case KeyCode
        Case vbKeyDown, vbKeyUp 'Cambiamos el "nivel" del mapa, al estilo Zelda ;D
            ToggleImgMaps
        Case Else
            Unload Me
    End Select
    
End Sub

''
' Toggle which image is visible.
'
Private Sub ToggleImgMaps()
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 24/07/08
'
'*************************************************

    imgToogleMap(CurrentMap).Visible = False
    
    If CurrentMap = eMaps.ieGeneral Then
        cCerrar.Visible = False
        CurrentMap = eMaps.ieDungeon
    Else
        cCerrar.Visible = True
        CurrentMap = eMaps.ieGeneral
    End If
    
    imgToogleMap(CurrentMap).Visible = True
    Me.Picture = picMaps(CurrentMap)
End Sub

''
' Load the images. Resizes the form, adjusts image's left and top and set lblTexto's Top and Left.
'
Private Sub Form_Load()
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 03/08/2012 - ^[GS]^
'*************************************************

On Error GoTo error
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
        
    'Cargamos las imagenes de los mapas
    Set picMaps(eMaps.ieGeneral) = LoadPicture(DirGraficos & "Mapa1.jpg")
    Set picMaps(eMaps.ieDungeon) = LoadPicture(DirGraficos & "Mapa2.jpg")
    
    ' Imagen de fondo
    CurrentMap = eMaps.ieGeneral
    Me.Picture = picMaps(CurrentMap)
    
    imgToogleMap(0).MouseIcon = picMouseIcon
    imgToogleMap(1).MouseIcon = picMouseIcon
    
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
    
    Exit Sub
error:
    MsgBox Err.Description, vbInformation, "Error: " & Err.Number
    Unload Me
End Sub

Private Sub cCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
End Sub

Private Sub imgToogleMap_Click(Index As Integer)
    ToggleImgMaps
End Sub
