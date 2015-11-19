VERSION 5.00
Begin VB.Form frmQuestInfo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Información de la misión"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   5055
   Icon            =   "frmQuestInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQuestInfo.frx":000C
   ScaleHeight     =   2895
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin GSZAOCliente.uAOButton cRechazar 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      TX              =   "Rechazar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmQuestInfo.frx":880D
      PICF            =   "frmQuestInfo.frx":8829
      PICH            =   "frmQuestInfo.frx":8845
      PICV            =   "frmQuestInfo.frx":8861
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
   Begin GSZAOCliente.uAOButton cAceptar 
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      TX              =   "Aceptar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmQuestInfo.frx":887D
      PICF            =   "frmQuestInfo.frx":8899
      PICH            =   "frmQuestInfo.frx":88B5
      PICV            =   "frmQuestInfo.frx":88D1
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
Attribute VB_Name = "frmQuestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub cAceptar_Click()
    
    Call Audio.PlayWave(SND_CLICK)
    Call WriteQuestAccept
    Unload Me
    
End Sub

Private Sub cRechazar_Click()

    Call Audio.PlayWave(SND_CLICK)
    Unload Me
    
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGUI & "frmFormYesNo.jpg") ' TODO: Falta una ventana para esto
    
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
