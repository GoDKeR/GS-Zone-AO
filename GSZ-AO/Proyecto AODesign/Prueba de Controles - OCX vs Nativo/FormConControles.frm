VERSION 5.00
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
   Begin Proyecto1.uAOProgress uAOProgress1 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      Min             =   1
      UseBackground   =   0   'False
      BackColor       =   8421504
      BorderColor     =   8421504
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
   Begin Proyecto1.uAOButton uAOButton1 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      TX              =   "uAOButton1"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "FormConControles.frx":0000
      PICF            =   "FormConControles.frx":001C
      PICH            =   "FormConControles.frx":0038
      PICV            =   "FormConControles.frx":0054
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
   Begin Proyecto1.uAOCheckbox uAOCheckbox1 
      Height          =   345
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "FormConControles.frx":0070
   End
   Begin Proyecto1.uAOAnimLongLabel uAOAnimLongLabel1 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
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
