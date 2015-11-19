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
   Begin VB.TextBox Text 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Text            =   "Text"
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command 
      Caption         =   "Captcha..."
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox pCaptcha 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command_Click()
    Dim Capt(3) As Byte
    Capt(0) = Asc(Mid(Text.Text, 1))
    Capt(1) = Asc(Mid(Text.Text, 2))
    Capt(2) = Asc(Mid(Text.Text, 3))
    Capt(3) = Asc(Mid(Text.Text, 4))
    pCaptcha.Cls
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 200))
    pCaptcha.Line (RandomNumber(pCaptcha.ScaleWidth, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleHeight - 10, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 200))
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleWidth))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 200))
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.Line (RandomNumber(pCaptcha.ScaleWidth, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleWidth - 20, pCaptcha.ScaleWidth), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleWidth))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.CurrentX = (pCaptcha.ScaleWidth / 2) - RandomNumber(300, 400)
    pCaptcha.CurrentY = (pCaptcha.ScaleHeight / 2) - RandomNumber(140, 170)
    pCaptcha.ForeColor = RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(60, 255))
    pCaptcha.Print Chr(Capt(0))
    pCaptcha.CurrentX = (pCaptcha.ScaleWidth / 2) - RandomNumber(-60, 100)
    pCaptcha.CurrentY = (pCaptcha.ScaleHeight / 2) - RandomNumber(140, 170)
    pCaptcha.ForeColor = RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(60, 255))
    pCaptcha.Print Chr(Capt(2))
    pCaptcha.CurrentX = (pCaptcha.ScaleWidth / 2) - RandomNumber(-100, -200)
    pCaptcha.CurrentY = (pCaptcha.ScaleHeight / 2) - RandomNumber(140, 170)
    pCaptcha.ForeColor = RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(60, 255))
    pCaptcha.Print Chr(Capt(3))
    pCaptcha.CurrentX = (pCaptcha.ScaleWidth / 2) - RandomNumber(150, 200)
    pCaptcha.CurrentY = (pCaptcha.ScaleHeight / 2) - RandomNumber(150, 170)
    pCaptcha.ForeColor = RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(60, 255))
    pCaptcha.Print Chr(Capt(1))
    pCaptcha.Line (RandomNumber(pCaptcha.ScaleWidth, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleHeight - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.Line (RandomNumber(pCaptcha.ScaleWidth, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleHeight, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
End Sub

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound
End Function
