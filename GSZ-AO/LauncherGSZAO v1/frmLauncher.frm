VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmLauncher 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   5520
   ClientLeft      =   2595
   ClientTop       =   570
   ClientWidth     =   8940
   Icon            =   "frmLauncher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLauncher.frx":030A
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   596
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   0
      Top             =   480
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   3360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin LauncherUpdater.uAOProgress sTF 
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   2280
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      BackColor       =   255
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin LauncherUpdater.uAOProgress sSF 
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      BackColor       =   16711680
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   2520
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   4095
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
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblSt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OFFLINE"
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   4095
   End
   Begin VB.Label lblOn 
      BackStyle       =   0  'Transparent
      Caption         =   "Número de usuarios jugando: 0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Image imgGotoGsz 
      Height          =   1095
      Left            =   480
      Top             =   0
      Width           =   7935
   End
   Begin VB.Image imgJugar 
      Height          =   1095
      Left            =   4920
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revisando actualizaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()

Winsock1.Connect "localhost", 7666

'wbb1.Navigate "www.gs-zone.org"

End Sub

Private Sub imgJugar_Click()
Label1.Visible = True

Call Main



End Sub

Private Sub Label2_Click()
Unload Me
End
End Sub


Private Sub Timer1_Timer()
checkOnline
End Sub

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

