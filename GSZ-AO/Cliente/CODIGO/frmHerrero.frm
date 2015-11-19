VERSION 5.00
Begin VB.Form frmConstruirHerrero 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Herrero"
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHerrero.frx":0000
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantItems 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   5175
      MaxLength       =   5
      TabIndex        =   14
      Text            =   "1"
      Top             =   2940
      Width           =   1050
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   1
      Left            =   5430
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   2
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   2355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   3
      Left            =   5430
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   3150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   4
      Left            =   5430
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   3945
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.ComboBox cboItemsCiclo 
      BackColor       =   &H80000006&
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
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   4
      Left            =   870
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   3945
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picLingotes3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   465
      ScaleWidth      =   1440
      TabIndex        =   8
      Top             =   3945
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   3
      Left            =   870
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   3150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picLingotes2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   465
      ScaleWidth      =   1440
      TabIndex        =   6
      Top             =   3150
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   2
      Left            =   870
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   2355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picLingotes1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   465
      ScaleWidth      =   1440
      TabIndex        =   4
      Top             =   2355
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.VScrollBar Scroll 
      Height          =   3135
      Left            =   450
      TabIndex        =   0
      Top             =   1410
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picLingotes0 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   465
      ScaleWidth      =   1440
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   1
      Left            =   870
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   450
      Left            =   3360
      TabIndex        =   15
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   794
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmHerrero.frx":1E447
      PICF            =   "frmHerrero.frx":1E463
      PICH            =   "frmHerrero.frx":1E47F
      PICV            =   "frmHerrero.frx":1E49B
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
      Left            =   3360
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      TX              =   "Construir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmHerrero.frx":1E4B7
      PICF            =   "frmHerrero.frx":1E4D3
      PICH            =   "frmHerrero.frx":1E4EF
      PICV            =   "frmHerrero.frx":1E50B
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
      Left            =   3360
      TabIndex        =   17
      Top             =   2250
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      TX              =   "Construir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmHerrero.frx":1E527
      PICF            =   "frmHerrero.frx":1E543
      PICH            =   "frmHerrero.frx":1E55F
      PICV            =   "frmHerrero.frx":1E57B
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
      Left            =   3360
      TabIndex        =   18
      Top             =   3045
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      TX              =   "Construir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmHerrero.frx":1E597
      PICF            =   "frmHerrero.frx":1E5B3
      PICH            =   "frmHerrero.frx":1E5CF
      PICV            =   "frmHerrero.frx":1E5EB
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
      Left            =   3360
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      TX              =   "Construir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmHerrero.frx":1E607
      PICF            =   "frmHerrero.frx":1E623
      PICH            =   "frmHerrero.frx":1E63F
      PICV            =   "frmHerrero.frx":1E65B
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
      TabIndex        =   20
      Top             =   1935
      Visible         =   0   'False
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   397
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmHerrero.frx":1E677
   End
   Begin VB.Image imgMarcoLingotes 
      Height          =   780
      Index           =   4
      Left            =   1560
      Top             =   3780
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgMarcoLingotes 
      Height          =   780
      Index           =   3
      Left            =   1560
      Top             =   2985
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgMarcoLingotes 
      Height          =   780
      Index           =   2
      Left            =   1560
      Top             =   2190
      Visible         =   0   'False
      Width           =   1740
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
      Index           =   3
      Left            =   5280
      Top             =   2985
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   4
      Left            =   5280
      Top             =   3780
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
   Begin VB.Image picPestania 
      Height          =   375
      Index           =   2
      Left            =   3240
      MousePointer    =   99  'Custom
      Top             =   420
      Width           =   1095
   End
   Begin VB.Image picPestania 
      Height          =   375
      Index           =   1
      Left            =   1680
      MousePointer    =   99  'Custom
      Top             =   420
      Width           =   1455
   End
   Begin VB.Image picPestania 
      Height          =   375
      Index           =   0
      Left            =   720
      MousePointer    =   99  'Custom
      Top             =   420
      Width           =   975
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   1
      Left            =   720
      Top             =   1395
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoLingotes 
      Height          =   780
      Index           =   1
      Left            =   1560
      Top             =   1395
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   1
      Left            =   5280
      Top             =   1395
      Visible         =   0   'False
      Width           =   780
   End
End
Attribute VB_Name = "frmConstruirHerrero"
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

Private Enum ePestania
    ieArmas
    ieArmaduras
    ieMejorar
End Enum

Private picCheck As Picture
Private picRecuadroItem As Picture
Private picRecuadroLingotes As Picture

Private Pestanias(0 To 2) As Picture
Private UltimaPestania As Byte

Private Cargando As Boolean

Private UsarMacro As Boolean
Private Armas As Boolean

Private clsFormulario As clsFormMovementManager

Private Sub ConstruirItem(ByVal Index As Integer)
    Dim ItemIndex As Integer
    Dim CantItemsCiclo As Integer
    
    If Scroll.Visible = True Then ItemIndex = Scroll.Value
    ItemIndex = ItemIndex + Index
    
    Select Case UltimaPestania
        Case ePestania.ieArmas
        
            If UsarMacro Then
                CantItemsCiclo = Val(cboItemsCiclo.Text)
                MacroBltIndex = ArmasHerrero(ItemIndex).OBJIndex
                frmMain.ActivarMacroTrabajo
            Else
                ' Que cosntruya el maximo, total si sobra no importa, valida el server
                CantItemsCiclo = Val(cboItemsCiclo.List(cboItemsCiclo.ListCount - 1))
            End If
            
            Call WriteInitCrafting(Val(txtCantItems.Text), CantItemsCiclo)
            Call WriteCraftBlacksmith(ArmasHerrero(ItemIndex).OBJIndex)
            
        Case ePestania.ieArmaduras
        
            If UsarMacro Then
                CantItemsCiclo = Val(cboItemsCiclo.Text)
                MacroBltIndex = ArmadurasHerrero(ItemIndex).OBJIndex
                frmMain.ActivarMacroTrabajo
             Else
                ' Que cosntruya el maximo, total si sobra no importa, valida el server
                CantItemsCiclo = Val(cboItemsCiclo.List(cboItemsCiclo.ListCount - 1))
            End If
            
            Call WriteInitCrafting(Val(txtCantItems.Text), CantItemsCiclo)
            Call WriteCraftBlacksmith(ArmadurasHerrero(ItemIndex).OBJIndex)
        
        Case ePestania.ieMejorar
            Call WriteItemUpgrade(HerreroMejorar(ItemIndex).OBJIndex)
    End Select
    
    Unload Me

End Sub

Private Sub cCerrar_Click()
    ' Cerramos la ventana
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
End Sub

Private Sub cConstruir_Click(Index As Integer)
    ' Es la misma función para Construir que para Mejorar
    Call Audio.PlayWave(SND_CLICK)
    Call ConstruirItem(Index + 1)
End Sub

Private Sub ChkMacro_Click()
    
    ' Vamos a usar macro de construcción?
    UsarMacro = Not UsarMacro
    
    cboItemsCiclo.Visible = UsarMacro
    
End Sub

Private Sub Form_Load()
'Last Modification: 10/08/2014 - ^[GS]^
'**************************************

    Dim MaxConstItem As Integer
    Dim i As Integer
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    ' Recuadros
    Set picRecuadroItem = LoadPicture(DirGUI & "frmConstruirHerreroRecItemsHerreria.jpg")
    Set picRecuadroLingotes = LoadPicture(DirGUI & "frmConstruirHerreroRecLingotes.jpg")
    For i = 1 To MAX_LIST_ITEMS
        imgMarcoItem(i).Picture = picRecuadroItem
        imgMarcoUpgrade(i).Picture = picRecuadroItem
        imgMarcoLingotes(i).Picture = picRecuadroLingotes
    Next
    
    ' Pestañas
    Set Pestanias(ePestania.ieArmas) = LoadPicture(DirGUI & "frmConstruirHerreroArmas.jpg")
    Set Pestanias(ePestania.ieArmaduras) = LoadPicture(DirGUI & "frmConstruirHerreroArmaduras.jpg")
    Set Pestanias(ePestania.ieMejorar) = LoadPicture(DirGUI & "frmConstruirHerreroMejorar.jpg")
    
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
    
    ' Cargar imagenes
    Set Me.Picture = Pestanias(ePestania.ieArmas)
    
    Cargando = True
    
    MaxConstItem = CInt((UserLvl - 2) * 0.2) ' 0.13.3
    MaxConstItem = IIf(MaxConstItem < 1, 1, MaxConstItem)
    MaxConstItem = IIf(UserClase = eClass.Worker, MaxConstItem, 1)

    For i = 1 To MaxConstItem
        cboItemsCiclo.AddItem i
    Next i
    cboItemsCiclo.ListIndex = 0
    
    ' Cursores
    picPestania(ePestania.ieArmas).MouseIcon = picMouseIcon
    picPestania(ePestania.ieArmaduras).MouseIcon = picMouseIcon
    picPestania(ePestania.ieMejorar).MouseIcon = picMouseIcon

    Cargando = False
    
    UsarMacro = True
    ChkMacro.Checked = UsarMacro
    Armas = True
    UltimaPestania = 0

End Sub

Public Sub HideExtraControls(ByVal NumItems As Integer, Optional ByVal Upgrading As Boolean = False)

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
    
    picLingotes0.Visible = (NumItems >= 1)
    picLingotes1.Visible = (NumItems >= 2)
    picLingotes2.Visible = (NumItems >= 3)
    picLingotes3.Visible = (NumItems >= 4)
    
    For i = 1 To MAX_LIST_ITEMS
        picItem(i).Visible = (NumItems >= i)
        imgMarcoItem(i).Visible = (NumItems >= i)
        imgMarcoLingotes(i).Visible = (NumItems >= i)
        picUpgradeItem(i).Visible = (NumItems >= i And Upgrading)
        imgMarcoUpgrade(i).Visible = (NumItems >= i And Upgrading)
    Next i
    
    ChkMacro.Visible = Not Upgrading
    cboItemsCiclo.Visible = Not Upgrading And UsarMacro
    txtCantItems.Visible = Not Upgrading

    If NumItems > MAX_LIST_ITEMS Then
        Scroll.Visible = True
        Cargando = True
        Scroll.Max = NumItems - MAX_LIST_ITEMS
        Cargando = False
    Else
        Scroll.Visible = False
    End If
    
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

Public Sub RenderList(ByVal Inicio As Integer, ByVal Armas As Boolean)
On Error Resume Next

    Dim i As Long
    Dim NumItems As Integer
    Dim ObjHerrero() As tItemsConstruibles
    
    If Armas Then
        ObjHerrero = ArmasHerrero
    Else
        ObjHerrero = ArmadurasHerrero
    End If
    
    NumItems = UBound(ObjHerrero)
    Inicio = Inicio - 1
    
    For i = 1 To MAX_LIST_ITEMS
        If i + Inicio <= NumItems Then
            With ObjHerrero(i + Inicio)
                ' Agrego el item
                Call RenderItem(picItem(i), .GrhIndex)
                picItem(i).ToolTipText = .Name
     
                Call RenderItem(picUpgradeItem(i), .UpgradeGrhIndex)
                picUpgradeItem(i).ToolTipText = .UpgradeName
                
                 ' Inventariode lingotes
                Call InvLingosHerreria(i).SetItem(1, 0, .LinH, 0, LH_GRH, 0, 0, 0, 0, 0, 0, "Lingotes de Hierro")
                Call InvLingosHerreria(i).SetItem(2, 0, .LinP, 0, LP_GRH, 0, 0, 0, 0, 0, 0, "Lingotes de Plata")
                Call InvLingosHerreria(i).SetItem(3, 0, .LinO, 0, LO_GRH, 0, 0, 0, 0, 0, 0, "Lingotes de Oro")
            End With
        End If
    Next i
End Sub

Public Sub RenderUpgradeList(ByVal Inicio As Integer)
On Error Resume Next
    
    Dim i As Long
    Dim NumItems As Integer
    
    NumItems = UBound(HerreroMejorar)
    Inicio = Inicio - 1
    
    For i = 1 To MAX_LIST_ITEMS
        If i + Inicio <= NumItems Then
            With HerreroMejorar(i + Inicio)
                ' Agrego el item
                Call RenderItem(picItem(i), .GrhIndex)
                picItem(i).ToolTipText = .Name
                
                Call RenderItem(picUpgradeItem(i), .UpgradeGrhIndex)
                picUpgradeItem(i).ToolTipText = .UpgradeName
                
                 ' Inventariode lingotes
                Call InvLingosHerreria(i).SetItem(1, 0, .LinH, 0, LH_GRH, 0, 0, 0, 0, 0, 0, "Lingotes de Hierro")
                Call InvLingosHerreria(i).SetItem(2, 0, .LinP, 0, LP_GRH, 0, 0, 0, 0, 0, 0, "Lingotes de Plata")
                Call InvLingosHerreria(i).SetItem(3, 0, .LinO, 0, LO_GRH, 0, 0, 0, 0, 0, 0, "Lingotes de Oro")
            End With
        End If
    Next i
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Dim i As Integer
    
    For i = 1 To MAX_LIST_ITEMS
        Set InvLingosHerreria(i) = Nothing
    Next i
    
    MirandoHerreria = False
    
End Sub


Private Sub picPestania_Click(Index As Integer)
On Error Resume Next

    Dim i As Integer
    Dim NumItems As Integer
    
    If Cargando Then Exit Sub
    If UltimaPestania = Index Then Exit Sub
    
    Scroll.Value = 0
    
    Select Case Index
        Case ePestania.ieArmas
            ' Background
            Me.Picture = Pestanias(ePestania.ieArmas)
            
            NumItems = UBound(ArmasHerrero)
        
            Call HideExtraControls(NumItems)
            
            ' Cargo inventarios e imagenes
            Call RenderList(1, True)
            
            Armas = True
            
        Case ePestania.ieArmaduras
            ' Background
            Me.Picture = Pestanias(ePestania.ieArmaduras)
            
            NumItems = UBound(ArmadurasHerrero)
        
            Call HideExtraControls(NumItems)
            
            ' Cargo inventarios e imagenes
            Call RenderList(1, False)
            
            Armas = False
            
        Case ePestania.ieMejorar
            ' Background
            Me.Picture = Pestanias(ePestania.ieMejorar)
            
            NumItems = UBound(HerreroMejorar)
            
            Call HideExtraControls(NumItems, True)
            
            Call RenderUpgradeList(1)
    End Select

    UltimaPestania = Index
End Sub

Private Sub Scroll_Change()
On Error Resume Next

    Dim i As Long
    
    If Cargando Then Exit Sub
    
    i = Scroll.Value
    ' Cargo inventarios e imagenes
    
    Select Case UltimaPestania
        Case ePestania.ieArmas
            Call RenderList(i + 1, True)
        Case ePestania.ieArmaduras
            Call RenderList(i + 1, False)
        Case ePestania.ieMejorar
            Call RenderUpgradeList(i + 1)
    End Select
End Sub

Private Sub txtCantItems_Change()
On Error GoTo ErrHandler

    If Val(txtCantItems.Text) < 0 Then
        txtCantItems.Text = 1
    End If
    
    If Val(txtCantItems.Text) > MAX_INVENTORY_OBJS Then
        txtCantItems.Text = MAX_INVENTORY_OBJS
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    txtCantItems.Text = MAX_INVENTORY_OBJS
End Sub

Private Sub txtCantItems_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
    
End Sub

