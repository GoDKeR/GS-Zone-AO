VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3480
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6840
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2401.958
   ScaleMode       =   0  'User
   ScaleWidth      =   6423.114
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5400
      TabIndex        =   0
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Agradecimientos a Alejandro Santos (AlejoLP) y Kiko (Otto Wallace)"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   4815
   End
   Begin VB.Label lblCredits5 
      Caption         =   "Colaboraciones de Juan Martín Sotuyo Dodero (Maraxus)"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   4815
   End
   Begin VB.Label lblCredits4 
      Caption         =   "ProgressBarSlider Pro ActiveX - Copyright © 2000 por Nik Tupkalov"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   4815
   End
   Begin VB.Label lblCredits3 
      Caption         =   "Chameleon Button - Copyright © 2003 por gonchuki"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label lblCredits2 
      Caption         =   "versión 2.1 - Copyright © 2012 por ^[GS]^"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label lblCredits1 
      Caption         =   "código original - Copyright © por Ivan Leoni y Fernando Costa"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   6310.427
      Y1              =   1242.392
      Y2              =   1242.392
   End
   Begin VB.Label lblDescription 
      Caption         =   "Aplicación de configuración y resolución de problemas de GS-Zone Argentum Online."
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1050
      TabIndex        =   2
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Aplication Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   6310.427
      Y1              =   1242.392
      Y2              =   1242.392
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Acerca de " & App.Title
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub
