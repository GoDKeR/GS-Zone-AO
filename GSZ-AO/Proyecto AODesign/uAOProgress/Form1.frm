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
   Begin Proyecto1.uAOProgress uAOProgress 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   3615
      _ExtentX        =   2143
      _ExtentY        =   661
      Max             =   200
      Value           =   100
      BackColor       =   192
      BorderColor     =   16711935
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command 
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProgressBar_GotFocus()

End Sub

Private Sub Command_Click()
    uAOProgress.Value = Val(Rnd(1) * 100)
End Sub

