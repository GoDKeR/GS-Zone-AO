VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmAOSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Argentum Online Setup"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frmAOSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Opciones 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   2040
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame Frame7 
         Caption         =   "Musicalización"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   47
         Top             =   1440
         Width           =   6255
         Begin VB.CheckBox chkMusica 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "&Música"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   48
            ToolTipText     =   "Musica del Juego"
            Top             =   360
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   49
            Top             =   360
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            LargeChange     =   10
            Max             =   100
            TickStyle       =   3
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Efectos de Sonido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   0
         TabIndex        =   30
         Top             =   120
         Width           =   6255
         Begin VB.CheckBox chkSonido 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "&Sonido"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   32
            ToolTipText     =   "Sonidos ambiente del Juego"
            Top             =   360
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkEfectosSonido 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "&Efectos de Sonido (3D)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   31
            ToolTipText     =   "Efectos de Sonido de la Interface"
            Top             =   720
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   46
            Top             =   360
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Max             =   100
            TickStyle       =   3
         End
      End
   End
   Begin VB.Frame Opciones 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   3
      Left            =   240
      TabIndex        =   29
      Top             =   2040
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame Frame4 
         Caption         =   "Opciones de Screenshots"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   0
         TabIndex        =   33
         Top             =   120
         Width           =   6375
         Begin VB.TextBox tLevelShot 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5760
            MaxLength       =   4
            TabIndex        =   37
            Text            =   "40"
            Top             =   1080
            Width           =   495
         End
         Begin VB.CheckBox chkScreenKill 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "&Hacer screenshot al matar un personaje de nivel mayor a"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   36
            ToolTipText     =   "Sonidos ambiente del Juego"
            Top             =   1080
            Width           =   5655
         End
         Begin VB.CheckBox chkScreenDie 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "&Hacer screenshot al morir"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   35
            ToolTipText     =   "Sonidos ambiente del Juego"
            Top             =   720
            Width           =   2895
         End
         Begin VB.CheckBox chkActScreenshots 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "&Activado"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   34
            ToolTipText     =   "Sonidos ambiente del Juego"
            Top             =   360
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Opciones 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   4
      Left            =   240
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame Frame6 
         Caption         =   "Opciones de Generales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         TabIndex        =   44
         Top             =   1800
         Width           =   6375
         Begin VB.CheckBox chkCursoresPersonalizados 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Utilizar &Cursores Personalizados"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   45
            ToolTipText     =   "Sonidos ambiente del Juego"
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Opciones de Clan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   0
         TabIndex        =   39
         Top             =   120
         Width           =   6375
         Begin VB.CheckBox chkNewsGuild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Mostrar &Noticias del Clan al iniciar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   42
            ToolTipText     =   "Sonidos ambiente del Juego"
            Top             =   360
            Width           =   3375
         End
         Begin VB.CheckBox chkDlgGuild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Activar &Mensajes Clan en Pantalla"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   41
            ToolTipText     =   "Sonidos ambiente del Juego"
            Top             =   720
            Width           =   3855
         End
         Begin VB.TextBox tGuildDlgCount 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            MaxLength       =   4
            TabIndex        =   40
            Text            =   "5"
            Top             =   1020
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Maximo de Mensajes en Pantalla:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   43
            Top             =   1080
            Width           =   2865
         End
      End
   End
   Begin VB.Frame Opciones 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   6375
      Begin VB.PictureBox fondoVersion 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   6315
         TabIndex        =   18
         Top             =   1800
         Width           =   6375
         Begin VB.Label lDirectX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   2040
            TabIndex        =   20
            Top             =   45
            Width           =   180
         End
         Begin VB.Label lVersionFondo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Versión detectada:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   45
            Width           =   1830
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pruebas de DirectX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   6375
         Begin VB.PictureBox DirectDrawTest 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   1155
            Left            =   4680
            ScaleHeight     =   1095
            ScaleWidth      =   1515
            TabIndex        =   11
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
            Begin VB.Timer Timer1 
               Left            =   120
               Top             =   600
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "DirectX 8"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "DirectDraw"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "DirectSound"
            Top             =   1080
            Width           =   1335
         End
         Begin AOSetup.chameleonButton bProbarSonido 
            Height          =   375
            Left            =   2400
            TabIndex        =   12
            Top             =   960
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Prueba de S&onido"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmAOSetup.frx":0442
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin AOSetup.chameleonButton bProbarVideo 
            Height          =   375
            Left            =   2400
            TabIndex        =   13
            Top             =   600
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Prueba de &Video"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmAOSetup.frx":045E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin VB.Label lblDX 
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   16
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblDD 
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   15
            Top             =   720
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblDS 
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   14
            Top             =   1080
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Line Line5 
            X1              =   120
            X2              =   2280
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   2280
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   2280
            Y1              =   1320
            Y2              =   1320
         End
      End
      Begin AOSetup.chameleonButton cLibrerias 
         Height          =   615
         Left            =   0
         TabIndex        =   17
         Top             =   2280
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Verificar &Librerias"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAOSetup.frx":047A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Opciones 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame Frame1 
         Caption         =   "Opciones de Video"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   0
         TabIndex        =   22
         Top             =   120
         Width           =   6375
         Begin VB.CheckBox chkDinamico 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Usar carga &Dinámica"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3720
            TabIndex        =   25
            ToolTipText     =   "Utilizar carga Dinámica de Graficos"
            Top             =   360
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkVSync 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Utilizar &VSync"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   24
            ToolTipText     =   "Utilizar Sincronización Vertical (VSync)"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.ComboBox cProcesar 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmAOSetup.frx":0496
            Left            =   240
            List            =   "frmAOSetup.frx":04A3
            Style           =   2  'Dropdown List
            TabIndex        =   23
            ToolTipText     =   "Determina si los calculos Vertex de Geometria se realizarán mediante Software, Hardware (GPU) o Mixto."
            Top             =   600
            Width           =   2295
         End
         Begin AOSetup.PBarY pMemoria 
            CausesValidation=   0   'False
            Height          =   255
            Left            =   3720
            TabIndex        =   26
            Top             =   1080
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            Value           =   16
            Min             =   4
            Max             =   64
            BackColor       =   0
            FillColor       =   8421631
            BorderColor     =   16777215
            BorderStyle     =   3
            EnabledSlider   =   0   'False
            MousePointer    =   0
            picForeColor    =   12632256
            picFillColor    =   8421504
            Style           =   1
         End
         Begin VB.Label lCuantoVideo 
            Alignment       =   2  'Center
            Caption         =   "Usar 16 Mb de Memoria"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   3720
            TabIndex        =   28
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label lPro 
            AutoSize        =   -1  'True
            Caption         =   "Procesado mediante..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   1920
         End
      End
   End
   Begin VB.CheckBox cEjecutar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Ejecutar el juego al Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   5280
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin AOSetup.chameleonButton bCancelar 
      Default         =   -1  'True
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   5280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAOSetup.frx":04D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton bAceptar 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Aceptar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632319
      BCOLO           =   12632319
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421631
      MPTR            =   1
      MICON           =   "frmAOSetup.frx":04EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton cCreditos 
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      ToolTipText     =   "Creditos"
      Top             =   1300
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BTYPE           =   5
      TX              =   "?"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAOSetup.frx":0508
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.TabStrip cTab 
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5953
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Pruebas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Video"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Audio"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Screenshots"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Otros"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   6720
      Y1              =   5175
      Y2              =   5175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   120
      X2              =   6720
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1500
      Left            =   120
      Picture         =   "frmAOSetup.frx":0524
      Top             =   120
      Width           =   6675
   End
End
Attribute VB_Name = "frmAOSetup"
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

' sonido
Private Type SoundBuffer
    FileName As String
    looping As Boolean
    x As Byte
    Y As Byte
    normalFq As Long
    Buffer As DirectSoundSecondaryBuffer8
End Type
Dim m_dsBuffer As SoundBuffer
Dim m_bLoaded As Boolean

' video
Private Const SW_SHOWNORMAL = 1

Dim CharWidth As Integer
Dim CharHight As Integer
Dim PostionX As Integer
Dim postionY As Integer
Dim running As Boolean

Private Sub bAceptar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/08/2013 - ^[GS]^
'*************************************************
    ' VIDEO
    ClientAOSetup.bVertex = Val(Me.cProcesar.ListIndex)
    ClientAOSetup.bVSync = CBool(Me.chkVSync.value)
    ClientAOSetup.bDinamic = CBool(Me.chkDinamico.value)
    ClientAOSetup.byMemory = CByte(Me.pMemoria.value)
    
    ' SONIDO
    ClientAOSetup.bNoSound = Not CBool(Me.chkSonido.value)
    ClientAOSetup.bNoMusic = Not CBool(Me.chkMusica.value)
    ClientAOSetup.bNoSoundEffects = Not CBool(Me.chkEfectosSonido.value)
    ClientAOSetup.lMusicVolume = Me.Slider1(1).value
    ClientAOSetup.lSoundVolume = Me.Slider1(0).value

    ' SCREENSHOTS
    ClientAOSetup.bActive = CBool(Me.chkActScreenshots.value)
    ClientAOSetup.bDie = CBool(Me.chkScreenDie.value)
    ClientAOSetup.bKill = CBool(Me.chkScreenKill.value)
    ClientAOSetup.byMurderedLevel = Val(Me.tLevelShot.Text)
    
    ' CLAN
    ClientAOSetup.bGuildNews = CBool(Me.chkNewsGuild.value)
    ClientAOSetup.bGldMsgConsole = CBool(Me.chkDlgGuild.value)
    ClientAOSetup.bCantMsgs = Val(Me.tGuildDlgCount.Text)
    
    ' Generales
    ClientAOSetup.bCursores = CBool(Me.chkCursoresPersonalizados.value)
    
    ' Guardamos...
    Dim handle As Integer
    handle = FreeFile
    Open App.Path & iniInit For Binary As handle
        Put handle, , ClientAOSetup
    Close handle
    DoEvents
    
    ' Ejecutamos
    If cEjecutar.value = 1 Then
        If FileExist(App.Path & iniClient, vbArchive) = True Then _
            Call Shell(App.Path & iniClient)
        DoEvents
    End If
    
    Unload Me
End Sub

Private Sub bCancelar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
    Unload Me
End Sub

Private Sub bProbarSonido_Click()
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 09/07/2012 - ^[GS]^
'*************************************************
On Error Resume Next
    
    If bProbarSonido.value = True Then
        Dim sonido As String
        sonido = App.Path & iniTestWAV
        
        If FileExist(sonido, vbArchive) = False Then
            MsgBox "No se puede probar el sonido porque falta el archivo de pruebas.", vbCritical
            Exit Sub
        End If
        
        DirectSound.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
        If m_bLoaded = False Then
            m_bLoaded = True
            LoadWave 0, sonido
        End If
        Dim flag As Long
        flag = 0
        m_dsBuffer.Buffer.Play flag
        
        If Err.Number <> 0 Then
            MsgBox "Problemas de DirectSound, Reinstale DIRECTX.", vbOKOnly, "Argentum Online Setup"
        End If
    Else
        If m_dsBuffer.Buffer Is Nothing Then Exit Sub
        m_dsBuffer.Buffer.Stop
        m_dsBuffer.Buffer.SetCurrentPosition 0
    End If
End Sub

Sub LoadWave(I As Integer, sfile As String)
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 09/07/2012 - ^[GS]^
'*************************************************

    Dim bufferDesc As DSBUFFERDESC
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    Set m_dsBuffer.Buffer = DirectSound.CreateSoundBufferFromFile(sfile, bufferDesc)
    
    If Err.Number <> 0 Then
        MsgBox "Error en " + sfile
        End
    End If
End Sub

Private Sub bProbarVideo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************

If bProbarVideo.value = True Then
    DirectDrawTest.Visible = True
    Call DirectDrawTestStart
    DoEvents
Else
    DirectDrawTest.Visible = False
    Timer1.Enabled = False
    running = False
End If
End Sub

Private Sub cCreditos_Click()
'*************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last modified: 15/03/06
'*************************************************
    frmAbout.Show vbModal, Me
End Sub

Private Sub chkActScreenshots_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 18/07/2012
'*************************************************
    If chkActScreenshots.value Then
        chkScreenDie.Enabled = True
        chkScreenKill.Enabled = True
    Else
        chkScreenDie.Enabled = False
        chkScreenKill.Enabled = False
    End If
End Sub

Private Sub chkDinamico_Click()
'*************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last modified: 10/03/06
'*************************************************
    If chkDinamico.value Then
        lCuantoVideo.ForeColor = vbBlack
        pMemoria.EnabledSlider = True
        pMemoria.picFillColor = &H8080FF
        pMemoria.picForeColor = &H80FF80
    Else
        lCuantoVideo.ForeColor = &H808080
        pMemoria.EnabledSlider = False
        pMemoria.picFillColor = &H808080
        pMemoria.picForeColor = &HC0C0C0
    End If
End Sub

Private Sub chkDlgGuild_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 18/07/2012
'*************************************************
    If chkDlgGuild.value Then
        tGuildDlgCount.Enabled = True
    Else
        tGuildDlgCount.Enabled = False
    End If
End Sub

Private Sub chkScreenKill_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 18/07/2012
'*************************************************
    If chkScreenKill.value Then
        tLevelShot.Enabled = True
    Else
        tLevelShot.Enabled = False
    End If
End Sub

Private Sub cLibrerias_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
frmLibrerias.Show
End Sub

Private Sub cTab_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 18/07/2012
'*************************************************
On Error Resume Next
    Call OcultarOpciones
    Opciones(cTab.SelectedItem.index - 1).Visible = True
End Sub

Private Sub OcultarOpciones()
'*************************************************
'Author: ^[GS]^
'Last modified: 18/07/2012
'*************************************************
On Error Resume Next
    Dim I As Byte
    For I = 0 To Opciones.UBound
        Opciones(I).Visible = False
    Next I
End Sub

Private Sub Form_Load()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/07/2012 - ^[GS]^
'*************************************************
On Error Resume Next
    Me.Show
    
    DoEvents
    
    Call LeerSetup
    
    If FileExist(App.Path & iniClient, vbArchive) = False Then
        cEjecutar.value = 0
        cEjecutar.Visible = False
    End If
    
    
    Opciones(0).Visible = True
    
    Call mod_DirectX.ProbarDirectX
    Call mod_DirectX.VersionDirectX
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last modified: 27/07/2012 - ^[GS]^
'*************************************************
    If FileExist(App.Path & iniTempDX, vbArchive) Then
        Call Kill(App.Path & iniTempDX)
    End If
    DoEvents
    
    End ' FIN
    
End Sub

Private Sub pMemoria_ChangeValue(NewValue As Long, OldValue As Long)
'*************************************************
'Author: ^[GS]^
'Last modified: 13/07/2012 - ^[GS]^
'*************************************************
    lCuantoVideo.Caption = "Usar " & CStr(NewValue) & " MiB de Memoria"
End Sub




Private Sub tGuildDlgCount_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 18/07/2012
'*************************************************
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub tGuildDlgCount_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 18/07/2012
'*************************************************
    If IsNumeric(tGuildDlgCount.Text) = False Or Val(tGuildDlgCount.Text) < 1 Then
        tGuildDlgCount.Text = 5 ' default
    End If
End Sub

Private Sub Timer1_Timer()
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'*************************************************
    DoEvents
    Call DrawDX8
End Sub

Public Sub DirectDrawTestStart()
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'*************************************************
    If lblDD.ForeColor <> &H8000& Then
        DirectDrawTest.Visible = False
        Exit Sub
    End If

    Call InitDX8
    running = True
  
    PostionX = 0
    postionY = 3
    
    Timer1.Interval = 150
    Timer1.Enabled = True
End Sub

Private Sub tLevelShot_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 18/07/2012
'*************************************************
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub tLevelShot_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 18/07/2012
'*************************************************
    If IsNumeric(tLevelShot.Text) = False Or Val(tGuildDlgCount.Text) < 1 Then
        tLevelShot.Text = 40 ' default
    End If
End Sub
