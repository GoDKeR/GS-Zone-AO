VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3240
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Text            =   "Text"
      Top             =   960
      Width           =   2295
   End
   Begin Proyecto1.uAOButton uAOButton 
      Height          =   615
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      TX              =   "asd"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "Form1.frx":FD94
      PICF            =   "Form1.frx":FDB0
      PICH            =   "Form1.frx":FDCC
      PICV            =   "Form1.frx":FDE8
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Proyecto1.uAOButton uAOButton 
      Height          =   615
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      TX              =   ""
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "Form1.frx":FE04
      PICF            =   "Form1.frx":FE20
      PICH            =   "Form1.frx":FE3C
      PICV            =   "Form1.frx":FE58
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Proyecto1.uAOButton uAOButton 
      Height          =   615
      Index           =   2
      Left            =   -120
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      TX              =   "asd"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "Form1.frx":FE74
      PICF            =   "Form1.frx":FE90
      PICH            =   "Form1.frx":FEAC
      PICV            =   "Form1.frx":FEC8
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
Private Sub cClose_Click()
    Unload Me
End Sub

Private Sub cMensaje_Click()
    MsgBox "Mensaje :P"

End Sub


