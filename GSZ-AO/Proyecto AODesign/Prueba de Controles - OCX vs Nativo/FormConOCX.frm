VERSION 5.00
Object = "{16AB4CA3-BFAE-4164-8806-F8F8C0D97BA7}#1.0#0"; "AOControls.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin AOControls.uAOProgress uAOProgress1 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      Min             =   1
      UseBackground   =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOControls.uAOButton uAOButton1 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      TX              =   "uAOButton1"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "FormConOCX.frx":0000
      PICF            =   "FormConOCX.frx":001C
      PICH            =   "FormConOCX.frx":0038
      PICV            =   "FormConOCX.frx":0054
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOControls.uAOCheckbox uAOCheckbox1 
      Height          =   345
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "FormConOCX.frx":0070
   End
   Begin AOControls.uAOAnimLongLabel uAOAnimLongLabel1 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      Value           =   1
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim cControl As Control
    For Each cControl In Me.Controls
        If TypeOf cControl Is uAOButton Then
            cControl.PictureEsquina = LoadPicture(App.Path & ".\img\uAOButton-e.bmp")
            cControl.PictureFondo = LoadPicture(App.Path & ".\img\uAOButton-f.bmp")
            cControl.PictureHorizontal = LoadPicture(App.Path & ".\img\uAOButton-h.bmp")
            cControl.PictureVertical = LoadPicture(App.Path & ".\img\uAOButton-v.bmp")
        ElseIf TypeOf cControl Is uAOCheckbox Then
            cControl.Picture = LoadPicture(App.Path & ".\img\uAOCheckbox.bmp")
        End If
    Next
End Sub
