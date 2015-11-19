VERSION 5.00
Begin VB.Form frmTorneo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   Picture         =   "frmTorneo.frx":0000
   ScaleHeight     =   3210
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstParticipantes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1785
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin GSZAOCliente.uAOButton cmdAceptar 
      Height          =   480
      Left            =   240
      TabIndex        =   0
      Top             =   2650
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   847
      TX              =   "Aceptar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmTorneo.frx":9B51
      PICF            =   "frmTorneo.frx":9BAF
      PICH            =   "frmTorneo.frx":9C0D
      PICV            =   "frmTorneo.frx":9C6B
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
   Begin GSZAOCliente.uAOButton cmdCerrar 
      Height          =   480
      Left            =   3720
      TabIndex        =   1
      Top             =   2650
      Width           =   1085
      _ExtentX        =   1720
      _ExtentY        =   450
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmTorneo.frx":9CC9
      PICF            =   "frmTorneo.frx":9CE5
      PICH            =   "frmTorneo.frx":9D01
      PICV            =   "frmTorneo.frx":9D1D
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
End
Attribute VB_Name = "frmTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo    : frmTorneo
' Autor     : Facundo Ortega (GoDKeR)
' Fecha     : 07/03/2014
' Propósito : Seguimos la política del patrón a llevar todos los comandos a algo mas gráfico para comodidad del usuario _
(que ironía, el formulario se abre con comando ._.)
'---------------------------------------------------------------------------------------
Option Explicit

Private Sub cmdAceptar_Click()

    If OpcionTorneo = 1 Then 'gm
            WriteTorneoEvento 1, InputBox("Ingrese la cantidad de participantes"), InputBox("Se caen items? 0=No, 1=Si")
    Else
            WriteTorneoEvento 2
    End If
    
End Sub

Private Sub cmdCerrar_Click()
Unload Me

End Sub

Private Sub Form_Load()
'(•_•)_†

If OpcionTorneo = 1 Then
    cmdAceptar.Caption = "Abrir"
Else
    cmdAceptar.Caption = "Participar"
End If

End Sub
