VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   8970
   ClientLeft      =   330
   ClientTop       =   15
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   5880
      Top             =   1920
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Frame fMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   6840
      TabIndex        =   38
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
      Begin GSZAOCliente.uAOCheckbox chkMiniMap 
         Height          =   225
         Left            =   120
         TabIndex        =   46
         Top             =   120
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   397
         CHCK            =   -1  'True
         ENAB            =   -1  'True
         PICC            =   "frmMain.frx":24EC0
      End
      Begin GSZAOCliente.uAOButton cOpciones 
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         TX              =   "Opciones"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmMain.frx":24F1E
         PICF            =   "frmMain.frx":24F3A
         PICH            =   "frmMain.frx":24F56
         PICV            =   "frmMain.frx":24F72
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GSZAOCliente.uAOButton cMapa 
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         TX              =   "Mapa"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmMain.frx":24F8E
         PICF            =   "frmMain.frx":24FAA
         PICH            =   "frmMain.frx":24FC6
         PICV            =   "frmMain.frx":24FE2
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GSZAOCliente.uAOButton cGrupo 
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         TX              =   "Grupo"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmMain.frx":24FFE
         PICF            =   "frmMain.frx":2501A
         PICH            =   "frmMain.frx":25036
         PICV            =   "frmMain.frx":25052
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GSZAOCliente.uAOButton cEstadisticas 
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         TX              =   "Estadísticas"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmMain.frx":2506E
         PICF            =   "frmMain.frx":2508A
         PICH            =   "frmMain.frx":250A6
         PICV            =   "frmMain.frx":250C2
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GSZAOCliente.uAOButton cClanes 
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         TX              =   "Clanes"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmMain.frx":250DE
         PICF            =   "frmMain.frx":250FA
         PICH            =   "frmMain.frx":25116
         PICV            =   "frmMain.frx":25132
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GSZAOCliente.uAOButton cNotas 
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         TX              =   "Notas"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmMain.frx":2514E
         PICF            =   "frmMain.frx":2516A
         PICH            =   "frmMain.frx":25186
         PICV            =   "frmMain.frx":251A2
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GSZAOCliente.uAOButton cSalir 
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         TX              =   "Salir"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmMain.frx":251BE
         PICF            =   "frmMain.frx":251DA
         PICH            =   "frmMain.frx":251F6
         PICV            =   "frmMain.frx":25212
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblMiniMap 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MiniMap"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   9000
      ScaleHeight     =   190
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   158
      TabIndex        =   14
      Top             =   2400
      Width           =   2400
   End
   Begin GSZAOCliente.uAOAnimLongLabel lblOro 
      Height          =   255
      Left            =   10680
      TabIndex        =   36
      Top             =   6345
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Value           =   1
      ForeColor       =   33023
      BorderColor     =   42461
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOProgress cStatEnergia 
      Height          =   180
      Left            =   9840
      TabIndex        =   29
      Top             =   7515
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   318
      Value           =   1
      BackColor       =   20560
      BorderColor     =   16448
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox SendCMSTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   360
      MaxLength       =   160
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Chat de Clan"
      Top             =   8040
      Visible         =   0   'False
      Width           =   7815
   End
   Begin InetCtlsObjects.Inet iWeb 
      Left            =   5280
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   360
      MaxLength       =   160
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2400
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   4800
      Top             =   1920
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11280
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin VB.PictureBox picSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   9000
      ScaleHeight     =   2865
      ScaleWidth      =   2370
      TabIndex        =   23
      Top             =   2400
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6240
      Left            =   195
      MousePointer    =   99  'Custom
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   22
      Top             =   2235
      Width           =   8160
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   11325
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   21
      Top             =   8445
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   10950
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   20
      Top             =   8445
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   10575
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   19
      Top             =   8445
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   10200
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   18
      Top             =   8445
      Width           =   420
   End
   Begin VB.Timer MacroTrabajo 
      Enabled         =   0   'False
      Left            =   7080
      Top             =   2520
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   5760
      Top             =   2520
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   4920
      Top             =   2520
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1755
      Left            =   180
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Consola de Mensajes"
      Top             =   360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3096
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":2522E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   255
      Left            =   11640
      TabIndex        =   5
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      TX              =   "X"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":252AC
      PICF            =   "frmMain.frx":252C8
      PICH            =   "frmMain.frx":252E4
      PICV            =   "frmMain.frx":25300
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cMinimizar 
      Height          =   255
      Left            =   11400
      TabIndex        =   4
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      TX              =   "_"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":2531C
      PICF            =   "frmMain.frx":25338
      PICH            =   "frmMain.frx":25354
      PICV            =   "frmMain.frx":25370
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cAsignarSkills 
      Height          =   255
      Left            =   11040
      TabIndex        =   24
      Top             =   1605
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      TX              =   "+"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":2538C
      PICF            =   "frmMain.frx":253A8
      PICH            =   "frmMain.frx":253C4
      PICV            =   "frmMain.frx":253E0
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cInventario 
      Height          =   375
      Left            =   8985
      TabIndex        =   0
      Top             =   1965
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      TX              =   "Inventario"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":253FC
      PICF            =   "frmMain.frx":25418
      PICH            =   "frmMain.frx":25434
      PICV            =   "frmMain.frx":25450
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cHechizos 
      Height          =   375
      Left            =   10215
      TabIndex        =   1
      Top             =   1965
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      TX              =   "Hechizos"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":2546C
      PICF            =   "frmMain.frx":25488
      PICH            =   "frmMain.frx":254A4
      PICV            =   "frmMain.frx":254C0
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOProgress cStatMana 
      Height          =   225
      Left            =   8760
      TabIndex        =   30
      Top             =   6720
      Visible         =   0   'False
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   397
      Max             =   999
      Value           =   1
      BackColor       =   8388608
      BorderColor     =   4194304
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOProgress cStatVida 
      Height          =   225
      Left            =   8760
      TabIndex        =   32
      Top             =   7080
      Visible         =   0   'False
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   397
      Max             =   999
      Value           =   1
      BackColor       =   128
      BorderColor     =   64
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOProgress cStatHambre 
      Height          =   180
      Left            =   9840
      TabIndex        =   33
      Top             =   7815
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   318
      Value           =   1
      BackColor       =   24576
      BorderColor     =   16384
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOProgress cStatSed 
      Height          =   180
      Left            =   9840
      TabIndex        =   34
      Top             =   8130
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   318
      Value           =   1
      BackColor       =   8421376
      BorderColor     =   4210688
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOProgress cStatExp 
      Height          =   225
      Left            =   9015
      TabIndex        =   35
      Top             =   1140
      Visible         =   0   'False
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   397
      Value           =   1
      BackColor       =   32768
      BorderColor     =   0
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cMenu 
      Height          =   855
      Left            =   8760
      TabIndex        =   37
      Top             =   7440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      TX              =   "Menu"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":254DC
      PICF            =   "frmMain.frx":254F8
      PICH            =   "frmMain.frx":25514
      PICV            =   "frmMain.frx":25530
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jugando:"
      ForeColor       =   &H006F9BB2&
      Height          =   240
      Left            =   8520
      TabIndex        =   28
      Top             =   15
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FPS:"
      ForeColor       =   &H006F9BB2&
      Height          =   240
      Left            =   9960
      TabIndex        =   27
      Top             =   15
      Width           =   555
   End
   Begin VB.Label lblOnline 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9480
      TabIndex        =   26
      ToolTipText     =   "Jugadores conectados"
      Top             =   15
      Width           =   555
   End
   Begin VB.Label lblSkills 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   10620
      TabIndex        =   25
      Top             =   1620
      Width           =   360
   End
   Begin VB.Label lblFPS 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "101"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10680
      TabIndex        =   17
      ToolTipText     =   "Frames por Segundo"
      Top             =   15
      Width           =   555
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "GS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   9240
      TabIndex        =   16
      Top             =   570
      Width           =   1935
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9480
      TabIndex        =   15
      Top             =   1620
      Width           =   360
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   210
      Left            =   9600
      TabIndex        =   13
      Top             =   6360
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   210
      Left            =   9120
      TabIndex        =   12
      Top             =   6360
      Width           =   210
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "000 X:00 Y:00"
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   8685
      TabIndex        =   11
      ToolTipText     =   "Coordenadas"
      Top             =   8595
      Width           =   1350
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00/00"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   7080
      TabIndex        =   10
      Top             =   8595
      Width           =   735
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00/00"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   5250
      TabIndex        =   9
      Top             =   8595
      Width           =   735
   End
   Begin VB.Label lblHelm 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00/00"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3060
      TabIndex        =   8
      Top             =   8595
      Width           =   735
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00/00"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   8595
      Width           =   735
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      Height          =   6240
      Left            =   180
      Top             =   2235
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Label lblItem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   9000
      TabIndex        =   31
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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

Private last_i As Long
Public UsandoDrag As Boolean
Public UsabaDrag As Boolean

Public WithEvents dragInventory As clsGraphicalInventory
Attribute dragInventory.VB_VarHelpID = -1
Public WithEvents dragSpells As clsGraphicalSpells ' GSZAO
Attribute dragSpells.VB_VarHelpID = -1

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

Private clsFormulario As clsFormMovementManager

Public picSkillStar As Picture

Dim PuedeMacrear As Boolean
Private bKeyBack As Boolean ' GSZAO
Private NeedRedraw As Boolean ' GSZAO
Private AntiSendKey As Boolean ' GSZAO

'Usado para controlar que no se dispare el binding de la tecla CTRL cuando se usa CTRL+Tecla.
Dim CtrlMaskOn As Boolean ' 0.13.3

Public Sub RequestAsignarSkills()
    Dim i As Integer
    LlegaronSkills = False
    Call WriteRequestSkills
    Call FlushBuffer
    Do While Not LlegaronSkills
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    LlegaronSkills = False
    For i = 1 To NUMSKILLS
        frmSkills.text1(i).Caption = UserSkills(i)
    Next i
    Alocados = SkillPoints
    frmSkills.puntos.Caption = SkillPoints
    frmSkills.Show , frmMain
End Sub

Private Sub cAsignarSkills_Click()
    Call RequestAsignarSkills
End Sub

Private Sub cCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmCerrar.Show vbModal
End Sub

Private Sub cClanes_Click()
    Call Audio.PlayWave(SND_CLICK)
    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    
    fMenu.Visible = False
End Sub

Private Sub cEstadisticas_Click()
    Call Audio.PlayWave(SND_CLICK)
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
    Call WriteRequestFame
    Call FlushBuffer
    Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    frmEstadisticas.Iniciar_Labels
    frmEstadisticas.Show , frmMain
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    
    fMenu.Visible = False
End Sub

Private Sub cGrupo_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteRequestPartyForm
    
    fMenu.Visible = False
End Sub

Private Sub cHechizos_Click()
    Call Audio.PlayWave(SND_CLICK)
    ' Activo controles de hechizos
    picSpell.Visible = True
    ' Desactivo controles de inventario
    picInv.Visible = False
    Spells.RenderSpells
    cHechizos.Enabled = False
    cInventario.Enabled = True
    cInventario.SetFocus
    
    UsandoDrag = False
End Sub

Private Sub chkMiniMap_Click()
    modMinimap.MiniMapEnabled = Not modMinimap.MiniMapEnabled
    chkMiniMap.Checked = modMinimap.MiniMapEnabled
    Call modMinimap.EscribirMinimapInit
End Sub

Private Sub cInventario_Click()
    Call Audio.PlayWave(SND_CLICK)
    ' Activo controles de inventario
    picInv.Visible = True
    ' Desactivo controles de hechizo
    picSpell.Visible = False
    Inventario.DrawInv
    cInventario.Enabled = False
    cHechizos.Enabled = True
    cHechizos.SetFocus
    
    UsandoDrag = False
End Sub

Private Sub cMapa_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call frmMapa.Show(vbModeless, frmMain)

    fMenu.Visible = False
End Sub

Private Sub cMenu_Click()
    fMenu.Visible = Not fMenu.Visible
End Sub

Private Sub cMinimizar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Me.WindowState = 1
End Sub
Private Sub cNotas_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call frmNotas.Show(vbModeless, frmMain)
    If frmNotas.cMostrar.Caption = "Mostrar" Then
        Call frmNotas.MostrarNotas
    End If
    
    fMenu.Visible = False
End Sub

Private Sub cOpciones_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call frmOpciones.Show(vbModeless, frmMain)
    
    fMenu.Visible = False
End Sub

Private Sub cPruebaRango_Click()
'    bRangoReducido = Not bRangoReducido
'
'    If bRangoReducido Then
'        vAngle = 0
'    End If
End Sub

Private Sub cSalir_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmCerrar.Show vbModal
    
    fMenu.Visible = False
End Sub

Private Sub dragSpells_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call modProtocol.WriteMoveItem(originalSlot, newSlot, eMoveType.SpellsI)
End Sub

Private Sub Form_Load()

    Dim i As Byte
    Set dragInventory = Inventario
    Set dragSpells = Spells
    
    If NoRes Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me, 120
    End If

    Me.Picture = LoadPicture(DirGUI & "frmMain.jpg")
    
    If SkillPoints > 0 Then
        cAsignarSkills.Visible = True
    Else
        cAsignarSkills.Visible = False
    End If
    
    lblSkills.Caption = SkillPoints ' GSZAO
    
    'lblOro.MouseIcon = picMouseIcon
    
    For i = 0 To 3
        picSM(i).MouseIcon = picMouseIcon
    Next i
        
    Me.Left = 0
    Me.Top = 0
    
    Dim cControl As Control
    For Each cControl In Me.Controls
        If TypeOf cControl Is uAOButton Then
            cControl.PictureEsquina = LoadPicture(ImgRequest(DirButtons & sty_bEsquina))
            cControl.PictureFondo = LoadPicture(ImgRequest(DirButtons & sty_bFondo))
            cControl.PictureHorizontal = LoadPicture(ImgRequest(DirButtons & sty_bHorizontal))
            cControl.PictureVertical = LoadPicture(ImgRequest(DirButtons & sty_bVertical))
        ElseIf TypeOf cControl Is uAOCheckbox Then
            cControl.Picture = LoadPicture(ImgRequest(DirButtons & sty_cCheckbox2))
        End If
    Next
        
    ' Detect links in console
    EnableURLDetect RecTxt.hwnd, Me.hwnd ' 0.13.3
    
    CtrlMaskOn = False ' 0.13.3
    
End Sub

Public Sub LightSkillStar(ByVal bTurnOn As Boolean)
    If bTurnOn Then
        cAsignarSkills.Visible = True
    Else
        cAsignarSkills.Visible = False
    End If
    lblSkills.Caption = SkillPoints
End Sub


Public Sub ActivarMacroHechizos()
    'If Not hlst.Visible Then
    If Not picSpell.Visible Or Spells.SpellSelectedItem = 0 Then  ' GSZAO
        Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, True)
        Exit Sub
    End If
    
    TrainingMacro.Interval = INT_MACRO_HECHIS
    TrainingMacro.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos activado", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mSpells, True)
End Sub

Public Sub DesactivarMacroHechizos()
    TrainingMacro.Enabled = False
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, True)
    Call ControlSM(eSMType.mSpells, False)
End Sub

Public Sub ControlSM(ByVal Index As Byte, ByVal Mostrar As Boolean)
Dim GrhIndex As Long
Dim SR As RECT
Dim DR As RECT

GrhIndex = GRH_INI_SM + Index + SM_CANT * (CInt(Mostrar) + 1)

With GrhData(GrhIndex)
    SR.Left = .sX
    SR.Right = SR.Left + .pixelWidth
    SR.Top = .sY
    SR.Bottom = SR.Top + .pixelHeight
    
    DR.Left = 0
    DR.Right = .pixelWidth
    DR.Top = 0
    DR.Bottom = .pixelHeight
End With

Call DrawGrhtoHdc(picSM(Index).hdc, GrhIndex, SR, DR)
picSM(Index).Refresh

Select Case Index
    Case eSMType.sResucitation
        If Mostrar Then
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, True)
            picSM(Index).ToolTipText = "Seguro de resucitación activado."
        Else
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_OFF, 255, 0, 0, True, False, True)
            picSM(Index).ToolTipText = "Seguro de resucitación desactivado."
        End If
        
    Case eSMType.sSafemode
        If Mostrar Then
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, True)
            picSM(Index).ToolTipText = "Seguro activado."
        Else
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, True)
            picSM(Index).ToolTipText = "Seguro desactivado."
        End If
        
    Case eSMType.mSpells
        If Mostrar Then
            picSM(Index).ToolTipText = "Macro de hechizos activado."
        Else
            picSM(Index).ToolTipText = "Macro de hechizos desactivado."
        End If
        
    Case eSMType.mWork
        If Mostrar Then
            picSM(Index).ToolTipText = "Macro de trabajo activado."
        Else
            picSM(Index).ToolTipText = "Macro de trabajo desactivado."
        End If
End Select

SMStatus(Index) = Mostrar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 11/08/2014 - ^[GS]^
'***************************************************
On Error Resume Next
    
    If Not AntiSendKey Then Exit Sub ' GSZAO
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
        
        'Verificamos si se está presionando la tecla CTRL.
        If Shift = 2 Then ' 0.13.3
            If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
                If KeyCode = vbKey0 Then
                    'Si es CTRL+0 muestro la ventana de configuración de teclas.
                    Call frmCustomKeys.Show(vbModal, Me)
                    
                ElseIf KeyCode >= vbKey1 And KeyCode <= vbKey9 Then
                    'Si es CTRL+1..9 cambio la configuración.
                    If KeyCode - vbKey0 = CustomKeys.CurrentConfig Then Exit Sub
                    
                    CustomKeys.CurrentConfig = KeyCode - vbKey0
                    
                    Dim sMsg As String
                    
                    sMsg = "¡Se ha cargado la configuración "
                    If CustomKeys.CurrentConfig = 0 Then
                        sMsg = sMsg & "default"
                    Else
                        sMsg = sMsg & "perzonalizada número " & CStr(CustomKeys.CurrentConfig)
                    End If
                    sMsg = sMsg & "!"

                    Call ShowConsoleMsg(sMsg, 255, 255, 255, True)
                End If
                
                CtrlMaskOn = True
                Exit Sub
            End If
        End If
        
        If KeyCode = vbKeyControl Then ' 0.13.3
            'Chequeo que no se haya usado un CTRL + tecla antes de disparar las bindings.
            If CtrlMaskOn Then
                CtrlMaskOn = False
                Exit Sub
            End If
        End If
        
        If KeyCode = vbKeyEscape Then ' GSZAO
            frmCerrar.Show vbModal
            Exit Sub
        End If
        
        If KeyCode = 220 Then ' GSZAO
            If picSpell.Visible = True Then
                Call cInventario_Click
            Else
                Call cHechizos_Click
            End If
            Exit Sub
        End If
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    Audio.MusicActivated = Not Audio.MusicActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                    Audio.SoundActivated = Not Audio.SoundActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFxs)
                    Audio.SoundEffectsActivated = Not Audio.SoundEffectsActivated
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    Call WriteSafeToggle

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    Call WriteResuscitationToggle
            End Select
        Else
        
            'Evito que se muestren los mensajes personalizados cuando se cambie una configuración de teclas.
            If Shift = 2 Then Exit Sub ' 0.13.3

            Select Case KeyCode
                'Custom messages!
                Case vbKey0 To vbKey9
                    Dim CustomMessage As String
                    
                    CustomMessage = CustomMessages.Message((KeyCode - 39) Mod 10)
                    If LenB(CustomMessage) <> 0 Then
                        ' No se pueden mandar mensajes personalizados de clan o privado!
                        If UCase$(Left$(CustomMessage, 5)) <> "/CMSG" And _
                            Left$(CustomMessage, 1) <> "\" Then
                            
                            Call ParseUserCommand(CustomMessage)
                        End If
                    End If
            End Select
        End If
    End If
    
    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
            If SendTxt.Visible Then Exit Sub
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And _
              (Not frmMSG.Visible) And (Not MirandoForo) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendCMSTXT.Visible = True
                SendCMSTXT.SetFocus
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call ScreenCapture
                
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            If UserMinMAN = UserMaxMAN Then Exit Sub
            
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                End With
                Exit Sub
            End If
               
            If Not ClMeditarRapido Then ' maTih.-  07/04/2012  -  meditación rápida.
                If Not PuedeMacrear Then
                    AddtoRichTextBox frmMain.RecTxt, "No tan rápido..!", 255, 255, 255, False, False, True
                Else
                    Call WriteMeditate
                    PuedeMacrear = False
                End If
            Else
                Call WriteMeditate
            End If
            
        Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If TrainingMacro.Enabled Then
                DesactivarMacroHechizos
            Else
                ActivarMacroHechizos
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If MacroTrabajo.Enabled Then
                Call DesactivarMacroTrabajo
            Else
                Call ActivarMacroTrabajo
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
            If frmMain.MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteQuit
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
            If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
            Else
                If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
            End If
            
            If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
            If MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
            
            ' 0.13.3
            If frmCustomKeys.Visible Then Exit Sub 'Chequeo si está visible la ventana de configuración de teclas.

            If (CharList(UserCharIndex).Arma.WeaponWalk(CharList(UserCharIndex).Heading).GrhIndex <> 0) Then
                EstaAtacando = True ' animacion de ataque
                CharList(UserCharIndex).Arma.WeaponWalk(CharList(UserCharIndex).Heading).Started = 1
            End If
            
            Call WriteAttack
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            If SendCMSTXT.Visible Then Exit Sub
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And _
              (Not frmMSG.Visible) And (Not MirandoForo) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
            
    End Select
End Sub

Private Sub Form_Paint()
On Error Resume Next
' GSZ-AO - Dibujamos los inventarios ;)
        If picInv.Visible Then
            'picInv.SetFocus
            Call Inventario.DrawInv
        ElseIf picSpell.Visible Then
            'picSpell.SetFocus
            Call Spells.RenderSpells
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisableURLDetect ' 0.13.3
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ChangeCursorMain(cur_Normal)
    If prgRun = True Then
        If UnloadMode <> 1 Then ' Cerrado por código
            frmCerrar.Show vbModal
            Cancel = 1
        End If
    End If
End Sub



Private Sub Image1_Click()

End Sub

Private Sub lblMiniMap_Click()
    Call chkMiniMap_Click
End Sub

Private Sub lblOro_Click()
    Inventario.SelectGold
    If UserGLD > 0 Then
        If Not Comerciando Then frmCantidad.Show , frmMain
    End If
End Sub

'Private Sub imgInvScrollDown_Click()
'    Call Inventario.ScrollInventory(True)
'End Sub

'Private Sub imgInvScrollUp_Click()
'    Call Inventario.ScrollInventory(False)
'End Sub

'Private Sub lblScroll_Click(Index As Integer)
'    Inventario.ScrollInventory (Index = 0)
'End Sub

Private Sub Macro_Timer()
    PuedeMacrear = True
End Sub

Private Sub MacroTrabajo_Timer()
    If Inventario.SelectedItem = 0 Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    'If Not modApplication.IsAppActive() Then
    '    Call DesactivarMacroTrabajo
    '    Exit Sub
    'End If
    
    If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or _
                UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not MirandoHerreria) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
     If Not MirandoCarpinteria Then Call UsarItem
End Sub

Public Sub ActivarMacroTrabajo()
    MacroTrabajo.Interval = INT_MACRO_TRABAJO
    MacroTrabajo.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mWork, True)
End Sub

Public Sub DesactivarMacroTrabajo()
    MacroTrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mWork, False)
End Sub


Private Sub MainViewPic_Click()
    Form_Click

End Sub

Private Sub MainViewPic_DblClick()
    Form_DblClick
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub StopDragInv()
' GSZAO
    UsabaDrag = False
    UsandoDrag = False
    If CurrentCursor <> cur_Action Then
        Call ChangeCursorMain(cur_Normal)
        frmMain.picInv.MousePointer = vbNormal
    End If
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    MouseX = X
    MouseY = Y
    
    Dim selInvSlot      As Byte
    
    'Get new target positions
    ConvertCPtoTP X, Y, tX, tY

    If modMinimap.MiniMapEnabled = True And modMinimap.MiniMapVisible = True Then ' GSZAO
        If modMinimap.MiniMapDrag Then
            If MouseX > Minimap.X - 50 And MouseY > Minimap.Y - 50 And MouseX < Minimap.X + 100 And MouseY < Minimap.Y + 100 Then
                Minimap.X = MouseX
                Minimap.Y = MouseY
                Call EscribirMinimapInit ' Guardamos la posición del minimap
            End If
            Exit Sub 'lo hago para evitar posibles problemas, asi no hace otra accion mientras draggea el minimap
        End If
    End If

    With MapData(tX, tY)
        If UsabaDrag = False Then
            If CurrentCursor <> cur_Action Then
                If .CharIndex <> 0 Then
                    If CharList(.CharIndex).Invisible = False Then
                        If CharList(.CharIndex).bType = 0 Then ' NPC friendly
                            Call ChangeCursorMain(cur_Npc)
                        ElseIf CharList(.CharIndex).bType = 1 Then ' NPC hostile
                            Call ChangeCursorMain(cur_Npc_Hostile)
                        ElseIf CharList(.CharIndex).bType = 2 Then ' User
                            If mapInfo.Pk = False Then
                                Call ChangeCursorMain(cur_User)
                            Else
                                Call ChangeCursorMain(cur_User_Danger)
                            End If
                        End If
                    Else
                        Call ChangeCursorMain(cur_Normal)
                    End If
                ElseIf .ObjGrh.GrhIndex <> 0 Then
                    Call ChangeCursorMain(cur_Obj)
                Else
                    Call ChangeCursorMain(cur_Normal)
                End If
            End If
        Else ' Utiliza Drag
            'Drag de items a posiciones. [maTih.-]
            
            'Get the selected slot of the inventory.
            selInvSlot = Inventario.SelectedItem
            
            'Not selected item?
            If Not selInvSlot <> 0 Then Exit Sub
            
            'There is invalid position?.
            If .Blocked <> 0 Then
               Call ShowConsoleMsg("Posición inválida")
               Call StopDragInv
               Exit Sub
            End If
            
            ' Not Drop on ilegal position; Standelf
            Dim IS_VALID_POS As Boolean
            
            IS_VALID_POS = LegalPos(tX + 1, tY) = False And _
                            LegalPos(tX - 1, tY) = False And _
                            LegalPos(tX, tY - 1) = False And _
                            LegalPos(tX, tY + 1) = False
                
            If IS_VALID_POS Then
                Call ShowConsoleMsg("La posición donde desea tirar el ítem es ilegal.")
                Call StopDragInv
                Exit Sub
            End If
            
            'There is already an object in that position?.
            If Not .CharIndex <> 0 Then
                If .ObjGrh.GrhIndex <> 0 Then
                    Call ShowConsoleMsg("Hay un objeto en esa posición!")
                    Call StopDragInv
                    Exit Sub
                End If
            End If
            
            'Send the package.
            Call modProtocol.WriteDropObj(selInvSlot, tX, tY, 1)
            
            'Reset the flag.
            Call StopDragInv
        End If
    End With
End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub PicMH_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, True)
End Sub

Private Sub Coord_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mundo. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, True)
End Sub

Private Sub PicInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    If Not UsandoDrag Then
        If Button = vbRightButton Then
          
            If Inventario.SelectedItem = 0 Then Exit Sub
            
            If Inventario.GrhIndex(Inventario.SelectedItem) > 0 Then
            
                last_i = Inventario.SelectedItem
                If last_i > 0 And last_i <= MAX_INVENTORY_SLOTS Then
                
                    Dim i As Integer
                    Dim data() As Byte
                    Dim handle As Integer
                    Dim bmpData As StdPicture
                    Dim poss As Integer
                    
                    poss = BuscarI(Inventario.GrhIndex(Inventario.SelectedItem))
                    
                    If poss = 0 Then
                        i = GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum
                        If Get_Image(DirGraficos, CStr(GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum), data, True) Then
                            Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO
                            frmMain.ImageList1.ListImages.Add , CStr("g" & Inventario.GrhIndex(Inventario.SelectedItem)), Picture:=bmpData
                            poss = frmMain.ImageList1.ListImages.Count
                            Set bmpData = Nothing
                        End If
                    End If
                    
                    UsandoDrag = True
                    If frmMain.ImageList1.ListImages.Count <> 0 Then
                        Set picInv.MouseIcon = frmMain.ImageList1.ListImages(poss).ExtractIcon
                    End If
                    frmMain.picInv.MousePointer = vbCustom
                    Exit Sub
                    
                End If
            End If
        Else
            If CurrentCursor <> cur_Action Then
                Call ChangeCursorMain(cur_Normal)
            End If
        End If
    End If
    
End Sub

Private Sub picSM_DblClick(Index As Integer)
Select Case Index
    Case eSMType.sResucitation
        Call WriteResuscitationToggle
        
    Case eSMType.sSafemode
        Call WriteSafeToggle
        
    Case eSMType.mSpells
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
            End With
            Exit Sub
        End If
        
        If TrainingMacro.Enabled Then
            Call DesactivarMacroHechizos
        Else
            Call ActivarMacroHechizos
        End If
        
    Case eSMType.mWork
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
            End With
            Exit Sub
        End If
        
        If MacroTrabajo.Enabled Then
            Call DesactivarMacroTrabajo
        Else
            Call ActivarMacroTrabajo
        End If
End Select
End Sub

Private Sub picSpell_DblClick()
    UsandoDrag = False
End Sub

Private Sub picSpell_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UsaMacro = False
    CnTd = 0

    If Not UsandoDrag Then
        If Button = vbRightButton Then
          
            If Spells.SelectedItem = 0 Then Exit Sub
            
            If Spells.GrhIndex(Spells.SelectedItem) > 0 Then
            
                last_i = Spells.SelectedItem
                If last_i > 0 And last_i <= 30 Then
                
                    Dim i As Integer
                    Dim data() As Byte
                    Dim handle As Integer
                    Dim bmpData As StdPicture
                    Dim poss As Integer
                    
                    poss = BuscarI(Spells.GrhIndex(Spells.SelectedItem))
                    
                    If poss = 0 Then
                        i = GrhData(Spells.GrhIndex(Spells.SelectedItem)).FileNum
                        If Get_Image(DirGraficos, CStr(GrhData(Spells.GrhIndex(Spells.SelectedItem)).FileNum), data, True) Then
                            Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO
                            frmMain.ImageList1.ListImages.Add , CStr("g" & Spells.GrhIndex(Spells.SelectedItem)), Picture:=bmpData
                            poss = frmMain.ImageList1.ListImages.Count
                            Set bmpData = Nothing
                        End If
                    End If
                    
                    UsandoDrag = True
                    If frmMain.ImageList1.ListImages.Count <> 0 Then
                        Set picSpell.MouseIcon = frmMain.ImageList1.ListImages(poss).ExtractIcon
                    End If
                    frmMain.picSpell.MousePointer = vbCustom
                    Exit Sub
                    
                End If
            End If
        Else
            If CurrentCursor <> cur_Action Then
                Call ChangeCursorMain(cur_Normal)
            End If
    End If
    End If
End Sub

Private Sub picSpell_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSpell.MousePointer = vbDefault
    
End Sub

Private Sub RecTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StartCheckingLinks ' 0.13.3
End Sub


Private Sub SendTxt_Click()
    SendTxt.Tag = 0 ' GSZAO
End Sub

'0.13.3
Private Sub SendTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' GSZAO
    Select Case KeyCode
        Case vbKeyBack, vbKeyDelete
            Select Case Len(SendTxt.Text)
                Case Is <> 0
                    bKeyBack = True
            End Select
        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown
            If (SendTxt.Text <> stxtbuffer) Then SendTxt.Text = stxtbuffer
    End Select
    
    ' Control + Shift
    If Shift = 3 Then
        On Error GoTo ErrHandler
        ' Only allow numeric keys
        If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
            ' Get Msg Number
            Dim NroMsg As Integer
            NroMsg = KeyCode - vbKey0 - 1
            ' Pressed "0", so Msg Number is 9
            If NroMsg = -1 Then NroMsg = 9
            'Como es KeyDown, si mantenes _
             apretado el mensaje llena la consola
            If CustomMessages.Message(NroMsg) = SendTxt.Text Then
                Exit Sub
            End If
            CustomMessages.Message(NroMsg) = SendTxt.Text
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡""" & SendTxt.Text & """ fue guardado como mensaje personalizado " & NroMsg + 1 & "!!", .Red, .Green, .Blue, .bold, .italic)
            End With
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    'Did detected an invalid message??
    If Err.Number = CustomMessages.InvalidMessageErrCode Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("El Mensaje es inválido. Modifiquelo por favor.", .Red, .Green, .Blue, .bold, .italic)
        End With
    End If
    
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If (KeyCode = vbKeyReturn) Or (KeyCode = CustomKeys.BindedKey(eKeyType.mKeyTalk)) Then
        stxtbuffer = ComprobarCaracteresLegales(SendTxt.Text)
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = vbNullString
        SendTxt.Text = vbNullString
        KeyCode = 0
        SendTxt.Visible = False
        
        If picInv.Visible Then
            picInv.SetFocus
        Else
            picSpell.SetFocus
        End If
    End If
End Sub

Private Sub Second_Timer()
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
        End With
    Else
        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
            Else
                If Inventario.amount(Inventario.SelectedItem) > 1 Then
                    If Not Comerciando Then frmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem()
    If pausa Then Exit Sub
    
    If Comerciando Then Exit Sub
    
    If TrainingMacro.Enabled Then DesactivarMacroHechizos
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(Inventario.SelectedItem)
End Sub

Private Sub EquiparItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
        End With
    Else
        If Comerciando Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()
On Error Resume Next

    If Not picSpell.Visible Or Spells.SpellSelectedItem = 0 Then ' GSZAO
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    'Macros are disabled if focus is not on Argentum!
    If Not modApplication.IsAppActive() Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    If Comerciando Then Exit Sub
    
    If MainTimer.Check(TimersIndex.CastSpell, False) Then
        Call WriteCastSpell(Spells.SpellSelectedItem)
        Call WriteWork(eSkill.Magia)
    End If
    
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    If UsingSkill = Magia And Not MainTimer.Check(TimersIndex.CastSpell) Then Exit Sub
    
    If UsingSkill = Proyectiles And Not MainTimer.Check(TimersIndex.Attack) Then Exit Sub
    
    Call WriteWorkLeftClick(tX, tY, UsingSkill)
    UsingSkill = 0
End Sub

'Private Sub DespInv_Click(Index As Integer)
'    Inventario.ScrollInventory (Index = 0)
'End Sub

Private Sub Form_Click()
    If Cartel Then Cartel = False

    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
         
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                    If CnTd = 5 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                
                    If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
                    If MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        Call ChangeCursorMain(cur_Normal)
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .Red, .Green, .Blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            Call ChangeCursorMain(cur_Normal)
                            UsingSkill = 0
                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .Red, .Green, .Blue, .bold, .italic)
                            End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                Call ChangeCursorMain(cur_Normal)
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rápido.", .Red, .Green, .Blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                Call ChangeCursorMain(cur_Normal)
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rápido.", .Red, .Green, .Blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            Call ChangeCursorMain(cur_Normal)
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    If CurrentCursor <> cur_Action Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    Call ChangeCursorMain(cur_Normal)
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort ' 0.13.5
                MiniMapDrag = False
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
            End If
        End If
    End If
    
    If modMinimap.MiniMapEnabled = True And modMinimap.MiniMapVisible = True Then
        If MouseBoton = vbRightButton Then
            If MouseX > modMinimap.Minimap.X And MouseY > modMinimap.Minimap.Y And MouseX < modMinimap.Minimap.X + 100 And MouseY < modMinimap.Minimap.Y + 100 Then
                modMinimap.MiniMapDrag = Not modMinimap.MiniMapDrag
            End If
        End If
    End If
End Sub

Private Sub Form_DblClick()
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/27/2007
'12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
'**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteDoubleClick(tX, tY)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - MainViewShp.Left
    MouseY = Y - MainViewShp.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewShp.Width Then
        MouseX = MainViewShp.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewShp.Height Then
        MouseY = MainViewShp.Height
    End If
    
    ' Disable links checking (not over consola)
    StopCheckingLinks ' 0.13.3
    
End Sub



Private Sub picInv_DblClick()

    If MirandoCarpinteria Or MirandoHerreria Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
    
    Call UsarItem
    
    UsandoDrag = False
    
End Sub

Public Sub picSpell_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call modProtocol.WriteMoveItem(originalSlot, newSlot, eMoveType.SpellsI)
    frmMain.picSpell.MousePointer = vbNormal
End Sub

Public Sub picInv_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call modProtocol.WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
    frmMain.picInv.MousePointer = vbNormal
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
    Call ChangeCursorMain(cur_Normal)
    
    If Shift Then ' GSZAO
        Call EquiparItem
    End If
End Sub

Private Sub RecTxt_Change()
On Error Resume Next   'el .SetFocus causaba errores al salir y volver a entrar
    If Not modApplication.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    ElseIf (Not Comerciando) And (Not MirandoAsignarSkills) And _
        (Not frmMSG.Visible) And (Not MirandoForo) And _
        (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) And (Not MirandoParty) And _
        (Not frmCustomKeys.Visible) Then
         
        If picInv.Visible Then
            picInv.SetFocus
        ElseIf picSpell.Visible Then
            picSpell.SetFocus
        End If
    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        picSpell.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 13/03/2012 - ^[GS]^
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = vbNullString ' GSZAO no gastamos datos enviando un mensaje inutil...
    Else
        Dim tempstr As String
        tempstr = ComprobarCaracteresLegales(SendTxt.Text)  ' GSZAO
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
        
        ' GSZAO
        'Select Case (bKeyBack Or Len(SendTxt.Text) < 2)
        '    Case True
        '        bKeyBack = False
        '        Exit Sub
        'End Select
            
        'If (left$(stxtbuffer, 1) = "/") Then
        '    Call AutoComplete_(SendTxt)
        'End If
        
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
    ' GSZAO
    SendTxt.Text = stxtbuffer
End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If (KeyCode = vbKeyReturn) Or (KeyCode = CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)) Then
        'Say
        If LenB(stxtbuffercmsg) <> 0 Then
            Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = vbNullString
        SendCMSTXT.Text = vbNullString
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
        
        If picInv.Visible Then
            picInv.SetFocus
        Else
            picSpell.SetFocus
        End If
    End If
End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        'stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
        stxtbuffercmsg = vbNullString ' GSZAO
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendCMSTXT.Text)
            CharAscii = Asc(mid$(SendCMSTXT.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendCMSTXT.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendCMSTXT.Text = tempstr
        End If
        
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
Private Sub Socket1_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
    
    #If Testeo = 1 Then
        Debug.Print "Socket1_Connect..."
        Call LogTesteo("Socket1_Connect...")
    #End If
    
    Second.Enabled = True
    Call frmConnect.EstadoSocket ' GSZAO

    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
            Call Login
        
        Case E_MODO.Normal
            Call Login
        
        Case E_MODO.Dados
            Call TirarDados ' GSZAO
            
    End Select
End Sub

Private Sub Socket1_Disconnect()

    #If Testeo = 1 Then
        Debug.Print "Socket1_Disconnect..."
        Call LogTesteo("Socket1_Disconnect...")
    #End If
    
    ResetAllInfo
    Socket1.Cleanup
    Call frmConnect.EstadoSocket ' GSZAO
    
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    

    #If Testeo = 1 Then
        Debug.Print "Socket1_LastError..." & ErrorCode & " " & ErrorString
        Call LogTesteo("Socket1_LastError..." & ErrorCode & " " & ErrorString)
    #End If
    
    Select Case ErrorCode
        Case TOO_FAST
            Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
            Exit Sub
        Case REFUSED
            Call MsgBox("El servidor se encuentra cerrado o no te has podido conectar correctamente.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Case ABORTED_OR_OTHER
            Call MsgBox("La conexión ha sido aborada o otro problema ha sucedido.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Case TIME_OUT
            Call MsgBox("El tiempo de espera se ha agotado, intenta nuevamente.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Case NO_AVAILIBLE ' GSZAO
            Call MsgBox("La dirección del servidor no se encuentra disponible.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
            If LCase$(frmMain.Socket1.HostAddress) = "localhost" Then ' El tipico común de que no conecta con localhost ¬¬ Avisamos! by ^[GS]^
                Call MsgBox("Debes utilizar '127.0.0.1' como IP local, no 'localhost'.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
            End If
        Case Else
            Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    End Select

    frmConnect.MousePointer = 1
    Response = 0

    frmMain.Socket1.Disconnect
End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
    Dim RD As String
    Dim data() As Byte
    
    #If Testeo = 1 Then
        Debug.Print "Socket1_Read... " & dataLength
        Call LogTesteo("Socket1_Read... " & dataLength)
    #End If
    
    Call Socket1.Read(RD, dataLength)
    data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub
    
    'Put data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).CharIndex > 0 Then
        If CharList(MapData(tX, tY).CharIndex).Invisible = False Then
        
            Dim i As Long
            Dim M As frmMenuseFashion
            Set M = New frmMenuseFashion
            
            Load M
            M.SetCallback Me
            M.SetMenuId 1
            M.ListaInit 2, False
            
            If LenB(CharList(MapData(tX, tY).CharIndex).nombre) <> 0 Then
                M.ListaSetItem 0, CharList(MapData(tX, tY).CharIndex).nombre, True
            Else
                M.ListaSetItem 0, "<NPC>", True
            End If
            M.ListaSetItem 1, "Comerciar"
            
            M.ListaFin
            M.Show , Me

        End If
    End If
End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        Call WriteLeftClick(tX, tY)
        
    Case 1 'Comerciar
        Call WriteLeftClick(tX, tY)
        Call WriteCommerceStart
    End Select
End Select
End Sub

Public Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call modProtocol.WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
End Sub

Private Function BuscarI(gh As Integer) As Integer
Dim i As Integer
For i = 1 To frmMain.ImageList1.ListImages.Count
    If frmMain.ImageList1.ListImages(i).Key = "g" & CStr(gh) Then
        BuscarI = i
        Exit For
    End If
Next i
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) ' GSZAO
    If Not GetAsyncKeyState(KeyCode) < 0 Then Exit Sub
    AntiSendKey = True
End Sub

