VERSION 5.00
Begin VB.Form frmConID 
   BackColor       =   &H00101010&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slots de Conexión"
   ClientHeight    =   4710
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   4590
   Icon            =   "frmConID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2980
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1440
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Liberar todos los slots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   4290
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ver &estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   4290
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   4215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00E0E0E0&
      Height          =   510
      Left            =   180
      TabIndex        =   3
      Top             =   2400
      Width           =   4230
   End
End
Attribute VB_Name = "frmConID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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

Private Sub Command1_Click()

    Unload Me
    
End Sub

Private Sub Command2_Click()

    List1.Clear
    
    Dim c As Integer
    Dim i As Integer
    
    For i = 1 To iniMaxUsuarios
        List1.AddItem "UserIndex " & i & " -- " & UserList(i).ConnID
        If UserList(i).ConnID <> -1 Then c = c + 1
    Next i
    
    If c = iniMaxUsuarios Then
        Label1.Caption = "¡No hay slots vacios!"
    Else
        Label1.Caption = "¡Hay " & iniMaxUsuarios - c & " slots vacios!"
    End If

End Sub

Private Sub Command3_Click()

    Dim i As Integer
    
    For i = 1 To iniMaxUsuarios
        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida And Not UserList(i).flags.UserLogged Then Call CloseSocket(i)
    Next i

End Sub

