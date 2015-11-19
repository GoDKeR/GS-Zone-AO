VERSION 5.00
Begin VB.Form frmCerrar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cerrar"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   Picture         =   "frmCerrar.frx":0000
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2700
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCerrar.frx":9B51
      PICF            =   "frmCerrar.frx":9B6D
      PICH            =   "frmCerrar.frx":9B89
      PICV            =   "frmCerrar.frx":9BA5
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cRegresar 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   873
      TX              =   "Regresar a la pantalla de inicio"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCerrar.frx":9BC1
      PICF            =   "frmCerrar.frx":9BDD
      PICH            =   "frmCerrar.frx":9BF9
      PICV            =   "frmCerrar.frx":9C15
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
   Begin GSZAOCliente.uAOButton cSalir 
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   873
      TX              =   "Salir del Juego"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCerrar.frx":9C31
      PICF            =   "frmCerrar.frx":9C4D
      PICH            =   "frmCerrar.frx":9C69
      PICV            =   "frmCerrar.frx":9C85
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
Attribute VB_Name = "frmCerrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub cCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
End Sub

Private Sub cRegresar_Click()
    Call Audio.PlayWave(SND_CLICK)
    If UserParalizado Then 'Inmo
        With FontTypes(FontTypeNames.FONTTYPE_WARNING)
            Call ShowConsoleMsg("No puedes salir estando paralizado.", .red, .green, .blue, .bold, .italic)
        End With
        Exit Sub
    End If
    If frmMain.MacroTrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
    Call WriteQuit
    Unload Me
End Sub

Private Sub cSalir_Click()
    Call Audio.PlayWave(SND_CLICK)
    prgRun = False
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGUI & "frmCerrar.jpg")
    
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

