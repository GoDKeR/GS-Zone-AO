VERSION 5.00
Begin VB.Form frmFormYesNo 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Petición"
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   Picture         =   "frmFormYesNo.frx":0000
   ScaleHeight     =   2910
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tEspera 
      Interval        =   1000
      Left            =   2400
      Top             =   2280
   End
   Begin GSZAOCliente.uAOButton cAceptar 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Tag             =   "Aceptar"
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      TX              =   "Aceptar"
      ENAB            =   0   'False
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmFormYesNo.frx":8801
      PICF            =   "frmFormYesNo.frx":881D
      PICH            =   "frmFormYesNo.frx":8839
      PICV            =   "frmFormYesNo.frx":8855
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
   Begin GSZAOCliente.uAOButton cRechazar 
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      TX              =   "Rechazar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmFormYesNo.frx":8871
      PICF            =   "frmFormYesNo.frx":888D
      PICH            =   "frmFormYesNo.frx":88A9
      PICV            =   "frmFormYesNo.frx":88C5
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
   Begin VB.Label lMensaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Morpheus"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   105
      TabIndex        =   2
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4785
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmFormYesNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private iTiempo As Integer

Private Sub cAceptar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call modProtocol.WriteRequestFormYesNo(frmFormYesNo.Tag, 1)
    bFormYesNo = False
    Unload Me
End Sub

Private Sub cRechazar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call modProtocol.WriteRequestFormYesNo(frmFormYesNo.Tag, 0)
    bFormYesNo = False
    Unload Me
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGUI & "frmFormYesNo.jpg")
    
    bFormYesNo = True
    iTiempo = 0
    cAceptar.Enabled = False
    cAceptar.Caption = "Espere... (5)"
    
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

Private Sub tEspera_Timer()
    iTiempo = iTiempo + 1
    If iTiempo > 4 Then
        cAceptar.Enabled = True
        cAceptar.Caption = cAceptar.Tag
        tEspera.Enabled = False
    Else
        cAceptar.Caption = "Espere... (" & 5 - iTiempo & ")"
    End If
End Sub

