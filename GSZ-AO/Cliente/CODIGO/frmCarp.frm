VERSION 5.00
Begin VB.Form frmConstruirCarp 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCarp.frx":0000
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picMaderas4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1830
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   20
      Top             =   3945
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picMaderas3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1830
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   19
      Top             =   3150
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picMaderas2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1830
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   18
      Top             =   2355
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picMaderas1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1830
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   4
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   10
      Top             =   3945
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   3
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   3150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   2
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   2355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   1
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtCantItems 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   5175
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "Ingrese la cantidad total de items a construir."
      Top             =   2925
      Width           =   1050
   End
   Begin VB.ComboBox cboItemsCiclo 
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
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   1
      Left            =   870
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.VScrollBar Scroll 
      Height          =   3135
      Left            =   450
      TabIndex        =   0
      Top             =   1410
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   2
      Left            =   870
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   2355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   3
      Left            =   870
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   3150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   4
      Left            =   870
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   3945
      Visible         =   0   'False
      Width           =   480
   End
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   450
      Left            =   3120
      TabIndex        =   11
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   794
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCarp.frx":1F76C
      PICF            =   "frmCarp.frx":1F788
      PICH            =   "frmCarp.frx":1F7A4
      PICV            =   "frmCarp.frx":1F7C0
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
   Begin GSZAOCliente.uAOButton cConstruir 
      Height          =   615
      Index           =   0
      Left            =   3120
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      TX              =   "Construir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCarp.frx":1F7DC
      PICF            =   "frmCarp.frx":1F7F8
      PICH            =   "frmCarp.frx":1F814
      PICV            =   "frmCarp.frx":1F830
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
   Begin GSZAOCliente.uAOButton cConstruir 
      Height          =   615
      Index           =   1
      Left            =   3120
      TabIndex        =   13
      Top             =   2250
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      TX              =   "Construir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCarp.frx":1F84C
      PICF            =   "frmCarp.frx":1F868
      PICH            =   "frmCarp.frx":1F884
      PICV            =   "frmCarp.frx":1F8A0
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
   Begin GSZAOCliente.uAOButton cConstruir 
      Height          =   615
      Index           =   2
      Left            =   3120
      TabIndex        =   14
      Top             =   3045
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      TX              =   "Construir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCarp.frx":1F8BC
      PICF            =   "frmCarp.frx":1F8D8
      PICH            =   "frmCarp.frx":1F8F4
      PICV            =   "frmCarp.frx":1F910
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
   Begin GSZAOCliente.uAOButton cConstruir 
      Height          =   615
      Index           =   3
      Left            =   3120
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      TX              =   "Construir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCarp.frx":1F92C
      PICF            =   "frmCarp.frx":1F948
      PICH            =   "frmCarp.frx":1F964
      PICV            =   "frmCarp.frx":1F980
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
   Begin GSZAOCliente.uAOCheckbox ChkMacro 
      Height          =   225
      Left            =   5520
      TabIndex        =   16
      Top             =   1935
      Visible         =   0   'False
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   397
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmCarp.frx":1F99C
   End
   Begin VB.Image imgPestania 
      Height          =   375
      Index           =   1
      Left            =   1680
      MousePointer    =   99  'Custom
      Top             =   420
      Width           =   1215
   End
   Begin VB.Image imgPestania 
      Height          =   375
      Index           =   0
      Left            =   720
      MousePointer    =   99  'Custom
      Top             =   420
      Width           =   975
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   4
      Left            =   5280
      Top             =   3780
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   3
      Left            =   5280
      Top             =   2985
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   2
      Left            =   5280
      Top             =   2190
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   1
      Left            =   5280
      Top             =   1395
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   4
      Left            =   720
      Top             =   3780
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   3
      Left            =   720
      Top             =   2985
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   2
      Left            =   720
      Top             =   2190
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   1
      Left            =   720
      Top             =   1395
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoMaderas 
      Height          =   780
      Index           =   4
      Left            =   1680
      Top             =   3780
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgMarcoMaderas 
      Height          =   780
      Index           =   3
      Left            =   1680
      Top             =   2985
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgMarcoMaderas 
      Height          =   780
      Index           =   2
      Left            =   1680
      Top             =   2190
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgMarcoMaderas 
      Height          =   780
      Index           =   1
      Left            =   1680
      Top             =   1395
      Visible         =   0   'False
      Width           =   1260
   End
End
Attribute VB_Name = "frmConstruirCarp"
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

Dim Cargando As Boolean
Private clsFormulario As clsFormMovementManager

Private Enum ePestania
    ieItems
    ieMejorar
End Enum

Private picRecuadroItem As Picture
Private picRecuadroMaderas As Picture

Private Pestanias(1) As Picture
Private UltimaPestania As Byte

Private UsarMacro As Boolean

Private Sub cCerrar_Click()
    ' Cerramos la ventana
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
End Sub

Private Sub cConstruir_Click(Index As Integer)
    ' Es la misma función para Construir que para Mejorar
    Call Audio.PlayWave(SND_CLICK)
    Call Construir(Index + 1)
End Sub

Private Sub ChkMacro_Click()
    ' Vamos a usar macro de construcción?
    UsarMacro = Not UsarMacro
    
    cboItemsCiclo.Visible = UsarMacro

End Sub

Private Sub Form_Load()
'***************************************************
'Author: Unknown
'Last Modification: 10/08/2014 - ^[GS]^
'***************************************************

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    ' Loading...
    Call LoadDefaultValues
    
    ' Default
    Me.Picture = LoadPicture(DirGUI & "frmConstruirCarpItems.jpg")
    
    ' Recuadros
    Dim Index As Integer
    Set picRecuadroItem = LoadPicture(DirGUI & "frmConstruirCarpRecItems.jpg")
    Set picRecuadroMaderas = LoadPicture(DirGUI & "frmConstruirCarpRecMaderas.jpg")
    For Index = 1 To MAX_LIST_ITEMS
        imgMarcoItem(Index).Picture = picRecuadroItem
        imgMarcoUpgrade(Index).Picture = picRecuadroItem
        imgMarcoMaderas(Index).Picture = picRecuadroMaderas
    Next Index
    
    ' Pestañas
    Set Pestanias(ePestania.ieItems) = LoadPicture(DirGUI & "frmConstruirCarpItems.jpg")
    Set Pestanias(ePestania.ieMejorar) = LoadPicture(DirGUI & "frmConstruirCarpMejorar.jpg")
    
    ' Controles
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

    ' Cursores
    imgPestania(ePestania.ieItems).MouseIcon = picMouseIcon
    imgPestania(ePestania.ieMejorar).MouseIcon = picMouseIcon
    
End Sub

Private Sub LoadDefaultValues()
    
    Dim MaxConstItem As Integer
    Dim i As Integer

    Cargando = True
    
    MaxConstItem = CInt((UserLvl - 2) * 0.2) ' 0.13.3
    MaxConstItem = IIf(MaxConstItem < 1, 1, MaxConstItem)
    MaxConstItem = IIf(UserClase = eClass.Worker, MaxConstItem, 1)
     
    For i = 1 To MaxConstItem
        cboItemsCiclo.AddItem i
    Next i
    
    cboItemsCiclo.ListIndex = 0
    
    Scroll.Value = 0
    
    ChkMacro.Checked = True
    UsarMacro = True
    
    UltimaPestania = ePestania.ieItems
    
    Cargando = False
End Sub


Private Sub Construir(ByVal Index As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 06/05/2014 - ^[GS]^
'***************************************************

    Dim ItemIndex As Integer
    Dim CantItemsCiclo As Integer
    
    If Scroll.Visible = True Then ItemIndex = Scroll.Value
    ItemIndex = ItemIndex + Index
    
    Select Case UltimaPestania
        Case ePestania.ieItems
        
            If UsarMacro Then
                CantItemsCiclo = Val(cboItemsCiclo.Text)
                MacroBltIndex = ObjCarpintero(ItemIndex).OBJIndex
                frmMain.ActivarMacroTrabajo
            Else
                ' Que construya el maximo, total si sobra no importa, valida el server
                CantItemsCiclo = Val(cboItemsCiclo.List(cboItemsCiclo.ListCount - 1))
            End If
            
            Call WriteInitCrafting(Val(txtCantItems.Text), CantItemsCiclo)
            Call WriteCraftCarpenter(ObjCarpintero(ItemIndex).OBJIndex)
            
        Case ePestania.ieMejorar
            Call WriteItemUpgrade(CarpinteroMejorar(ItemIndex).OBJIndex)
    End Select
        
    Unload Me

End Sub

Public Sub HideExtraControls(ByVal NumItems As Integer, Optional ByVal Upgrading As Boolean = False)
'***************************************************
'Author: Unknown
'Last Modification: 06/05/2014 - ^[GS]^
'***************************************************
    
    Dim i As Integer
    
    For i = 0 To 3
        If NumItems >= (i + 1) Then
            If Upgrading = False Then
                cConstruir(i).Caption = "Construir"
            Else
                cConstruir(i).Caption = "Mejorar"
            End If
            cConstruir(i).Visible = True
        Else
            cConstruir(i).Visible = False
        End If
    Next
    
    For i = 1 To MAX_LIST_ITEMS
        picItem(i).Visible = (NumItems >= i)
        imgMarcoItem(i).Visible = (NumItems >= i)
        imgMarcoMaderas(i).Visible = (NumItems >= i)

        ' Upgrade
        imgMarcoUpgrade(i).Visible = (NumItems >= i And Upgrading)
        picUpgrade(i).Visible = (NumItems >= i And Upgrading)
    Next i
    
    If NumItems > MAX_LIST_ITEMS Then
        Scroll.Visible = True
        Cargando = True
        Scroll.Max = NumItems - MAX_LIST_ITEMS
        Cargando = False
    Else
        Scroll.Visible = False
    End If
    
    picMaderas1.Visible = (NumItems >= 1)
    picMaderas2.Visible = (NumItems >= 2)
    picMaderas3.Visible = (NumItems >= 3)
    picMaderas4.Visible = (NumItems >= 4)
    
    txtCantItems.Visible = Not Upgrading
    cboItemsCiclo.Visible = Not Upgrading And UsarMacro
    ChkMacro.Visible = Not Upgrading

End Sub

Private Sub RenderItem(ByRef Pic As PictureBox, ByVal GrhIndex As Long)
On Error Resume Next

    Dim SR As RECT
    Dim DR As RECT
    
    With GrhData(GrhIndex)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .pixelWidth
        SR.Bottom = SR.Top + .pixelHeight
    End With
    
    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.Bottom = 32
    
    Call DrawGrhtoHdc(Pic.hdc, GrhIndex, SR, DR)
    Pic.Refresh
End Sub

Public Sub RenderList(ByVal Inicio As Integer)
On Error Resume Next

    Dim i As Long
    Dim NumItems As Integer
    
    NumItems = UBound(ObjCarpintero)
    Inicio = Inicio - 1
    
    For i = 1 To MAX_LIST_ITEMS
        If i + Inicio <= NumItems Then
            With ObjCarpintero(i + Inicio)
                ' Agrego el item
                Call RenderItem(picItem(i), .GrhIndex)
                picItem(i).ToolTipText = .Name
            
                ' Inventario de leños
                Call InvMaderasCarpinteria(i).SetItem(1, 0, .Madera, 0, MADERA_GRH, 0, 0, 0, 0, 0, 0, "Leña")
                Call InvMaderasCarpinteria(i).SetItem(2, 0, .MaderaElfica, 0, MADERA_ELFICA_GRH, 0, 0, 0, 0, 0, 0, "Leña élfica")
            End With
        End If
    Next i
    
End Sub

Public Sub RenderUpgradeList(ByVal Inicio As Integer)
On Error Resume Next

    Dim i As Long
    Dim NumItems As Integer
    
    NumItems = UBound(CarpinteroMejorar)
    Inicio = Inicio - 1
    
    For i = 1 To MAX_LIST_ITEMS
        If i + Inicio <= NumItems Then
            With CarpinteroMejorar(i + Inicio)
                ' Agrego el item
                Call RenderItem(picItem(i), .GrhIndex)
                picItem(i).ToolTipText = .Name
                
                Call RenderItem(picUpgrade(i), .UpgradeGrhIndex)
                picUpgrade(i).ToolTipText = .UpgradeName
            
                ' Inventario de leños
                Call InvMaderasCarpinteria(i).SetItem(1, 0, .Madera, 0, MADERA_GRH, 0, 0, 0, 0, 0, 0, "Leña")
                Call InvMaderasCarpinteria(i).SetItem(2, 0, .MaderaElfica, 0, MADERA_ELFICA_GRH, 0, 0, 0, 0, 0, 0, "Leña élfica")
            End With
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    For i = 1 To MAX_LIST_ITEMS
        Set InvMaderasCarpinteria(i) = Nothing
    Next i

    MirandoCarpinteria = False
End Sub


Private Sub imgPestania_Click(Index As Integer)
On Error Resume Next

    Dim i As Integer
    Dim NumItems As Integer
    
    If Cargando Then Exit Sub
    If UltimaPestania = Index Then Exit Sub
    
    Scroll.Value = 0
    
    Select Case Index
        Case ePestania.ieItems
            ' Background
            Me.Picture = Pestanias(ePestania.ieItems)
            
            NumItems = UBound(ObjCarpintero)
        
            Call HideExtraControls(NumItems)
            
            ' Cargo inventarios e imagenes
            Call RenderList(1)
            

        Case ePestania.ieMejorar
            ' Background
            Me.Picture = Pestanias(ePestania.ieMejorar)
            
            NumItems = UBound(CarpinteroMejorar)
            
            Call HideExtraControls(NumItems, True)
            
            Call RenderUpgradeList(1)
    End Select

    UltimaPestania = Index

End Sub

Private Sub Scroll_Change()
    Dim i As Long
    
    If Cargando Then Exit Sub
    
    i = Scroll.Value
    ' Cargo inventarios e imagenes
    
    Select Case UltimaPestania
        Case ePestania.ieItems
            Call RenderList(i + 1)
        Case ePestania.ieMejorar
            Call RenderUpgradeList(i + 1)
    End Select
End Sub
