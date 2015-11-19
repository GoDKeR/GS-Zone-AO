VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmPrincipal 
   BorderStyle     =   0  'None
   Caption         =   "GS-Zone Argentum Online"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10035
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPrincipal.frx":030A
   ScaleHeight     =   7650
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin GSZAU.uAOCheckbox chkAutoJugar 
      Height          =   345
      Left            =   360
      TabIndex        =   8
      Top             =   6960
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmPrincipal.frx":21E06
   End
   Begin SHDocVwCtl.WebBrowser wNoticias 
      Height          =   4815
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   2415
      ExtentX         =   4260
      ExtentY         =   8493
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin GSZAU.uAOButton cmdSalir 
      Height          =   375
      Left            =   9000
      TabIndex        =   4
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      TX              =   "X"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmPrincipal.frx":22EEC
      PICF            =   "frmPrincipal.frx":23916
      PICH            =   "frmPrincipal.frx":245D8
      PICV            =   "frmPrincipal.frx":2556A
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tTimer 
      Interval        =   1000
      Left            =   8400
      Top             =   4680
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   8880
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin GSZAU.uAOProgress sTF 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   5880
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
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
   Begin GSZAU.uAOProgress sSF 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
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
   Begin GSZAU.uAOButton cmdMinimizar 
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      TX              =   "_"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmPrincipal.frx":2646C
      PICF            =   "frmPrincipal.frx":26E96
      PICH            =   "frmPrincipal.frx":27B58
      PICV            =   "frmPrincipal.frx":28AEA
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAU.uAOButton cmdJugar 
      Height          =   1815
      Left            =   7560
      TabIndex        =   7
      Top             =   5520
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3201
      TX              =   "JUGAR!"
      ENAB            =   0   'False
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmPrincipal.frx":299EC
      PICF            =   "frmPrincipal.frx":2A416
      PICH            =   "frmPrincipal.frx":2B0D8
      PICV            =   "frmPrincipal.frx":2C06A
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label tTS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
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
      Left            =   4680
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      DrawMode        =   2  'Blackness
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1940
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciar el Juego automaticamente"
      BeginProperty Font 
         Name            =   "Morpheus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   6960
      Width           =   3495
   End
   Begin VB.Label tTT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
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
      Left            =   4680
      TabIndex        =   3
      Top             =   6960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label tTF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
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
      Left            =   3000
      TabIndex        =   2
      Top             =   5520
      Width           =   4455
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PrvVal As Long
Private LastVal As Long

Private Sub cmdJugar_Click()
    Call IniciarJuego
End Sub

Private Sub cmdMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.Show
    Me.SetFocus
End Sub

Private Sub Form_Load()
    wNoticias.Navigate "http://gsz-ao.sourceforge.net/noticias.php"
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

Private Sub Label1_Click()
    chkAutoJugar.Checked = Not chkAutoJugar.Checked
End Sub

Private Sub tTimer_Timer()
    If frmPrincipal.sSF.Visible = True And SizeTotal > 0 Then
        Dim Ind As Double
        Dim Res As Double
        Dim iTm(1) As Long
        Dim sRes As String
        If (PrvVal <> 0) Then
            Ind = ((((frmPrincipal.sSF.Value - PrvVal) + LastVal) / 2) / 1024)
            iTm(0) = ((SizeDone) / 1024) ' hecho (en KiB)
            iTm(1) = ((SizeTotal) / 1024) ' por hacer (en KiB)
            frmPrincipal.tTS.Caption = Format(iTm(0) / 1024, "0.0") & " MiB / " & Format(iTm(1) / 1024, "0.0") & " MiB (" & Format((iTm(0) * 100 / iTm(1)), "0.0") & "%)"
            If (Ind > 0) Then
                Res = (((SizeTotal - SizeDone) / 1024) / Ind)
                If Res >= 3600 Then ' horas
                    iTm(0) = Int(Res / 3600) ' hs
                    iTm(1) = Int((Res - iTm(0) * 3600) / 60) ' min
                    sRes = iTm(0) & ":" & Format(iTm(1), "00") & " hs."
                ElseIf Res >= 60 Then ' minutos
                    iTm(1) = Int(Res / 60)
                    sRes = iTm(1) & ":" & Format(Res - (iTm(1) * 60), "00") & " min."
                Else
                    sRes = Res & " segundos."
                End If
            Else
                sRes = "-"
            End If
            frmPrincipal.tTT.Caption = Format(Ind, "0.0") & " KB/s - Restante: " & sRes
            ' TODO: Aquí facilmente se podría añadir el calculo para que nos diga el tiempo restante por archivo
            LastVal = (frmPrincipal.sSF.Value - PrvVal)
        End If
        PrvVal = frmPrincipal.sSF.Value
        If (frmPrincipal.tTT.Visible = False) Then
            frmPrincipal.tTT.Visible = True
        End If
    Else
        frmPrincipal.tTT.Visible = False
        LastVal = 0
        PrvVal = 0
    End If
End Sub

