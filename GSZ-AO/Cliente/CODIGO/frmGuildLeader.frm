VERSION 5.00
Begin VB.Form frmGuildLeader 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Administración del Clan"
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   7875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuildLeader.frx":0000
   ScaleHeight     =   494
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   6240
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   17
      Top             =   360
      Width           =   1500
   End
   Begin VB.TextBox txtFiltrarMiembros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3075
      TabIndex        =   6
      Top             =   2340
      Width           =   2580
   End
   Begin VB.TextBox txtFiltrarClanes 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   195
      TabIndex        =   5
      Top             =   2340
      Width           =   2580
   End
   Begin VB.TextBox txtguildnews 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   195
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3435
      Width           =   5475
   End
   Begin VB.ListBox solicitudes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      ItemData        =   "frmGuildLeader.frx":183FF
      Left            =   195
      List            =   "frmGuildLeader.frx":18401
      TabIndex        =   2
      Top             =   5100
      Width           =   2595
   End
   Begin VB.ListBox members 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":18403
      Left            =   3060
      List            =   "frmGuildLeader.frx":18405
      TabIndex        =   1
      Top             =   540
      Width           =   2595
   End
   Begin VB.ListBox guildslist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":18407
      Left            =   180
      List            =   "frmGuildLeader.frx":18409
      TabIndex        =   0
      Top             =   540
      Width           =   2595
   End
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   495
      Left            =   3030
      TabIndex        =   7
      Top             =   6720
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   873
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildLeader.frx":1840B
      PICF            =   "frmGuildLeader.frx":18427
      PICH            =   "frmGuildLeader.frx":18443
      PICV            =   "frmGuildLeader.frx":1845F
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
   Begin GSZAOCliente.uAOButton cDetallesClan 
      Height          =   375
      Left            =   150
      TabIndex        =   8
      Top             =   2700
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   661
      TX              =   "Detalles"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildLeader.frx":1847B
      PICF            =   "frmGuildLeader.frx":18497
      PICH            =   "frmGuildLeader.frx":184B3
      PICV            =   "frmGuildLeader.frx":184CF
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cDetallesMiembro 
      Height          =   375
      Left            =   3030
      TabIndex        =   9
      Top             =   2700
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   661
      TX              =   "Detalles"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildLeader.frx":184EB
      PICF            =   "frmGuildLeader.frx":18507
      PICH            =   "frmGuildLeader.frx":18523
      PICV            =   "frmGuildLeader.frx":1853F
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cDetallesSolicitud 
      Height          =   375
      Left            =   150
      TabIndex        =   10
      Top             =   6045
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   661
      TX              =   "Detalles"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildLeader.frx":1855B
      PICF            =   "frmGuildLeader.frx":18577
      PICH            =   "frmGuildLeader.frx":18593
      PICV            =   "frmGuildLeader.frx":185AF
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cAbrirElecciones 
      Height          =   375
      Left            =   150
      TabIndex        =   11
      Top             =   6840
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   661
      TX              =   "Abrir Elecciones"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildLeader.frx":185CB
      PICF            =   "frmGuildLeader.frx":185E7
      PICH            =   "frmGuildLeader.frx":18603
      PICV            =   "frmGuildLeader.frx":1861F
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cActualizar 
      Height          =   375
      Left            =   150
      TabIndex        =   12
      Top             =   4230
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   661
      TX              =   "Actualizar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildLeader.frx":1863B
      PICF            =   "frmGuildLeader.frx":18657
      PICH            =   "frmGuildLeader.frx":18673
      PICV            =   "frmGuildLeader.frx":1868F
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
   Begin GSZAOCliente.uAOButton cEditarCodex 
      Height          =   375
      Left            =   3030
      TabIndex        =   13
      Top             =   4680
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   661
      TX              =   "Editar Codex o Descripción"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildLeader.frx":186AB
      PICF            =   "frmGuildLeader.frx":186C7
      PICH            =   "frmGuildLeader.frx":186E3
      PICV            =   "frmGuildLeader.frx":186FF
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
   Begin GSZAOCliente.uAOButton cEditarURL 
      Height          =   375
      Left            =   3030
      TabIndex        =   14
      Top             =   5130
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   661
      TX              =   "Editar URL de la web del clan"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildLeader.frx":1871B
      PICF            =   "frmGuildLeader.frx":18737
      PICH            =   "frmGuildLeader.frx":18753
      PICV            =   "frmGuildLeader.frx":1876F
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
   Begin GSZAOCliente.uAOButton cPropuestasPaz 
      Height          =   495
      Left            =   3030
      TabIndex        =   15
      Top             =   5565
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   873
      TX              =   "Propuestas de Paz"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildLeader.frx":1878B
      PICF            =   "frmGuildLeader.frx":187A7
      PICH            =   "frmGuildLeader.frx":187C3
      PICV            =   "frmGuildLeader.frx":187DF
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cPropuestasAlianzas 
      Height          =   495
      Left            =   3030
      TabIndex        =   16
      Top             =   6120
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   873
      TX              =   "Propuestas de Alianza"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmGuildLeader.frx":187FB
      PICF            =   "frmGuildLeader.frx":18817
      PICH            =   "frmGuildLeader.frx":18833
      PICV            =   "frmGuildLeader.frx":1884F
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1815
      TabIndex        =   3
      Top             =   6510
      Width           =   255
   End
End
Attribute VB_Name = "frmGuildLeader"
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

Private Const MAX_NEWS_LENGTH As Integer = 512
Private clsFormulario As clsFormMovementManager

Private Sub cAbrirElecciones_Click()
On Error Resume Next
    Call Audio.PlayWave(SND_CLICK)
    Call WriteGuildOpenElections
    Unload Me
End Sub

Private Sub cActualizar_Click()
    Dim K As String
    K = Replace(txtguildnews, vbCrLf, "º")
    Call WriteGuildUpdateNews(K)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub cCerrar_Click()
On Error Resume Next
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
    frmMain.SetFocus
End Sub

Private Sub cDetallesClan_Click()
    Call Audio.PlayWave(SND_CLICK)
    If guildslist.ListIndex = -1 Then Exit Sub
    
    frmGuildBrief.EsLeader = True
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))
End Sub

Private Sub cDetallesMiembro_Click()
    Call Audio.PlayWave(SND_CLICK)
    If members.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembers
    Call WriteGuildMemberInfo(members.List(members.ListIndex))
End Sub

Private Sub cDetallesSolicitud_Click()
    Call Audio.PlayWave(SND_CLICK)
    If solicitudes.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    Call WriteGuildMemberInfo(solicitudes.List(solicitudes.ListIndex))
End Sub

Private Sub cEditarCodex_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call frmGuildDetails.Show(vbModal, frmGuildLeader)
End Sub

Private Sub cEditarURL_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call frmGuildURL.Show(vbModeless, frmGuildLeader)
End Sub

Private Sub cPropuestasAlianzas_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteGuildAlliancePropList
End Sub

Private Sub cPropuestasPaz_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteGuildPeacePropList
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGUI & "frmGuildLeader.jpg")
  
    Dim cControl As Control
    For Each cControl In Me.Controls
        If TypeOf cControl Is uAOButton Then
            cControl.PictureEsquina = LoadPicture(ImgRequest(DirButtons & sty_bEsquina))
            cControl.PictureFondo = LoadPicture(ImgRequest(DirButtons & sty_bFondo))
            cControl.PictureHorizontal = LoadPicture(ImgRequest(DirButtons & sty_bHorizontal))
            cControl.PictureVertical = LoadPicture(ImgRequest(DirButtons & sty_bVertical))
        ElseIf TypeOf cControl Is uAOCheckbox Then
            cControl.Picture = LoadPicture(ImgRequest(DirButtons & sty_cCheckbox))
        End If
    Next
  
End Sub

Private Sub txtguildnews_Change()
    If Len(txtguildnews.Text) > MAX_NEWS_LENGTH Then _
        txtguildnews.Text = Left$(txtguildnews.Text, MAX_NEWS_LENGTH)
End Sub

Private Sub txtFiltrarClanes_Change()
    Call FiltrarListaClanes(txtFiltrarClanes.Text)
End Sub

Private Sub txtFiltrarClanes_GotFocus()
    With txtFiltrarClanes
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub FiltrarListaClanes(ByRef sCompare As String)
On Error Resume Next
    Dim lIndex As Long
    
    With guildslist
        'Limpio la lista
        .Clear
        
        .Visible = False
        
        ' Recorro los arrays
        For lIndex = 0 To UBound(GuildNames)
            ' Si coincide con los patrones
            If InStr(1, UCase$(GuildNames(lIndex)), UCase$(sCompare)) Then
                ' Lo agrego a la lista
                .AddItem GuildNames(lIndex)
            End If
        Next lIndex
        
        .Visible = True
    End With

End Sub

Private Sub txtFiltrarMiembros_Change()
    Call FiltrarListaMiembros(txtFiltrarMiembros.Text)
End Sub

Private Sub txtFiltrarMiembros_GotFocus()
    With txtFiltrarMiembros
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub FiltrarListaMiembros(ByRef sCompare As String)
On Error Resume Next
    Dim lIndex As Long
    
    With members
        'Limpio la lista
        .Clear
        
        .Visible = False
        
        ' Recorro los arrays
        For lIndex = 0 To UBound(GuildMembers)
            ' Si coincide con los patrones
            If InStr(1, UCase$(GuildMembers(lIndex)), UCase$(sCompare)) Then
                ' Lo agrego a la lista
                .AddItem GuildMembers(lIndex)
            End If
        Next lIndex
        
        .Visible = True
    End With
End Sub

Public Sub FillLogo(ByVal pStr As String)
Dim pClan As Picture

pbLogo.AutoRedraw = True

If Not pStr = vbNullString Then
    Set pClan = LoadPicture(pStr)
Else
    Set pClan = VB.LoadPicture(DirGraficos & "Sin avatar - Casco.jpg")
End If

If Not ResizePicture(pbLogo, pClan) Then
    Set pClan = VB.LoadPicture(DirGraficos & "Sin avatar - Casco.jpg")
    pbLogo.Picture = pClan
End If

End Sub

Private Function LoadPicture(ByVal strFileName As String) As Picture
Dim IID As TGUID

        With IID
                .Data1 = &H7BF80980
                .Data2 = &HBF32
                .Data3 = &H101A
                .Data4(0) = &H8B
                .Data4(1) = &HBB
                .Data4(2) = &H0
                .Data4(3) = &HAA
                .Data4(4) = &H0
                .Data4(5) = &H30
                .Data4(6) = &HC
                .Data4(7) = &HAB
        End With
        
On Error GoTo ERR_LINE

        OleLoadPicturePath StrPtr(strFileName), 0&, 0&, 0&, IID, LoadPicture
        
Exit Function

ERR_LINE:
    Set LoadPicture = VB.LoadPicture(strFileName)
    
End Function
