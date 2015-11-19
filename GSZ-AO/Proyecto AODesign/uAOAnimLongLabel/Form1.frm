VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.uAOAnimLongLabel uAOAnimLongLabel1 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   1935
      _extentx        =   3413
      _extenty        =   873
      value           =   1
      font            =   "Form1.frx":0000
   End
   Begin VB.CommandButton Command 
      Height          =   495
      Left            =   960
      TabIndex        =   0
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
    uAOAnimLongLabel1.Value = Val(Rnd(1) * 1102100)
End Sub

Private Sub uAOAnimLongLabel1_Click()
    Me.Caption = "asd"
End Sub

Private Sub uAOAnimLongLabel1_DblClick()
    Me.Caption = "doble click"
End Sub

Private Sub uAOAnimLongLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Caption = "mouse fuera"
End Sub

Private Sub uAOAnimLongLabel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Me.Caption = "mouse move"
End Sub

Private Sub uAOAnimLongLabel1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Caption = "mouse up"
End Sub
