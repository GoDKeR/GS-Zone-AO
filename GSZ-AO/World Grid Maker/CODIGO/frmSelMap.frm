VERSION 5.00
Begin VB.Form frmSelMap 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00101010&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione el Mapa a utilizar"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7005
   Icon            =   "frmSelMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdNoMapa 
      BackColor       =   &H008080FF&
      Caption         =   "&No usar Mapa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelMapa 
      BackColor       =   &H0080FF80&
      Caption         =   "&Seleccionar Mapa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.PictureBox pBK 
      BackColor       =   &H00000000&
      Height          =   3135
      Left            =   2640
      ScaleHeight     =   3075
      ScaleWidth      =   4155
      TabIndex        =   4
      Top             =   600
      Width           =   4215
      Begin VB.PictureBox pMAP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   1320
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   96
         TabIndex        =   5
         Top             =   720
         Width           =   1470
      End
   End
   Begin VB.FileListBox fMaps 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4170
      Left            =   120
      Pattern         =   "*.bmp"
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdDirMap 
      BackColor       =   &H0080FFFF&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   80
      Width           =   855
   End
   Begin VB.TextBox tDirMap 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   80
      Width           =   3735
   End
   Begin VB.Label lDirMaps 
      BackStyle       =   0  'Transparent
      Caption         =   "Directorio de Mapas:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmSelMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'World Grid Maker
'Copyright (C) 2012 GS-Zone
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
'You can contact me at:
'info@gs-zone.org

Option Explicit

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdDirMap_Click()
On Error Resume Next

    Dim tDir As String
    tDir = ExplorarDirectorio(frmSelMap.hWnd, "Seleccione el Directorio donde se encuentran los Mapas", App.Path)
    If LenB(tDir) <> 0 And FileExist(tDir, vbDirectory) Then
        cDirMaps = tDir
        tDirMap.Text = cDirMaps
        fMaps.Path = cDirMaps
    End If
    
End Sub

Private Sub cmdNoMapa_Click()
On Error Resume Next

    frmMain.pMAP(SelMapPos).Picture = pBK.Picture
    If WorldGrid(SelMapPos) <> 0 Then
        WorldGrid(SelMapPos) = 0
        Changed = True
    End If
    Unload Me
End Sub

Private Sub cmdSelMapa_Click()
On Error Resume Next
    
    If pMAP.Tag <> 0 Then
        If WorldGrid(SelMapPos) <> pMAP.Tag Then
            frmMain.pMAP(SelMapPos).Picture = pMAP.Picture
            WorldGrid(SelMapPos) = pMAP.Tag
            Changed = True
        End If
        Unload Me
    End If
End Sub

Private Sub fMaps_Click()
On Error Resume Next

    pMAP.Tag = 0
    If LenB(fMaps.FileName) <> 0 Then
        pMAP.Cls
        If cUsarPNG = True Then
            Call PngPictureLoad(fMaps.Path & "\" & fMaps.FileName, pMAP, False)
        Else
            pMAP.Picture = LoadPicture(fMaps.Path & "\" & fMaps.FileName)
        End If
        pMAP.Tag = Right$(fMaps.FileName, Len(fMaps.FileName) - 4) ' mapa
        pMAP.Tag = Val(Left(pMAP.Tag, Len(pMAP.Tag) - 4)) ' .bmp/.png
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next

    Dim tExp As String

    If cUsarPNG = True Then
        tExp = ".png"
    Else
        tExp = ".bmp"
    End If
    
    fMaps.Pattern = "*" & tExp

    tDirMap.Text = cDirMaps
    fMaps.Path = cDirMaps
    
    If SelMapPos <> 0 And SelMapPos <= MaxGrid Then
        Dim tP As Integer
        Dim qMap As Integer
        qMap = WorldGrid(SelMapPos)
        For tP = 0 To fMaps.ListCount - 1
            If LCase$(fMaps.List(tP)) = "mapa" & qMap & tExp Then
                fMaps.Selected(tP) = True
                Exit For
            End If
        Next
    End If
    
End Sub
