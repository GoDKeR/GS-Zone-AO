VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSInet.ocx"
Begin VB.Form frmPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GS-Zone Auto-Updater"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5010
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPrincipal.frx":030A
   ScaleHeight     =   3195
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet 
      Left            =   3840
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin GSZAU.uAOProgress sTF 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
      _extentx        =   7223
      _extenty        =   450
      backcolor       =   255
      font            =   "frmPrincipal.frx":9E5B
   End
   Begin GSZAU.uAOProgress sSF 
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   4095
      _extentx        =   7223
      _extenty        =   450
      backcolor       =   16711680
      font            =   "frmPrincipal.frx":9E7F
   End
   Begin VB.Label tTF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label tSF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Me.Show
    Me.SetFocus
End Sub

Private Sub Form_LostFocus()
    Me.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Inet.Cancel
    DoEvents
    End
End Sub

Private Sub Inet_StateChanged(ByVal State As Integer)
    Me.Inet.Tag = State
End Sub
