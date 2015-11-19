VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H00101010&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de Personajes"
   ClientHeight    =   3270
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   3525
   Icon            =   "frmAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Expulsar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.ComboBox cboPjs 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Echar todos los PJS no privilegiados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmAdmin.frm
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Private Sub cboPjs_Change()

    Call ActualizaPjInfo
    
End Sub

Private Sub cboPjs_Click()

    Call ActualizaPjInfo
    
End Sub

Private Sub Command1_Click()

    Dim tIndex As Long
    
    tIndex = NameIndex(cboPjs.Text)
    If tIndex > 0 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & UserList(tIndex).Name & " ha sido echado.", FontTypeNames.FONTTYPE_SERVER))
        Call CloseSocket(tIndex)
    End If

End Sub

Public Sub ActualizaListaPjs()

    Dim LoopC As Long
    
    With cboPjs
        .Clear
        
        For LoopC = 1 To LastUser
            If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
                If UserList(LoopC).flags.Privilegios And PlayerType.User Then
                    .AddItem UserList(LoopC).Name
                    .ItemData(.NewIndex) = LoopC
                End If
            End If
        Next LoopC
    End With

End Sub

Private Sub Command2_Click()

    Unload Me
    
End Sub

Private Sub Command3_Click()

    Call EcharPjsNoPrivilegiados

End Sub

Private Sub ActualizaPjInfo()

    Dim tIndex As Long
    
    tIndex = NameIndex(cboPjs.Text)
    If tIndex > 0 Then
        With UserList(tIndex)
            Text1.Text = .outgoingData.length & " elementos en cola." & vbCrLf
        End With
    End If

End Sub

