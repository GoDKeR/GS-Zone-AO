VERSION 5.00
Begin VB.Form frmParty 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   Picture         =   "frmParty.frx":0000
   ScaleHeight     =   428
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin GSZAOCliente.uAOButton cSalirParty 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   5400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Salir Party"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmParty.frx":1EABD
      PICF            =   "frmParty.frx":1EAD9
      PICH            =   "frmParty.frx":1EAF5
      PICV            =   "frmParty.frx":1EB11
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
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   255
      Left            =   555
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   720
      Width           =   4530
   End
   Begin VB.TextBox txtToAdd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   240
      Left            =   1530
      MaxLength       =   20
      TabIndex        =   1
      Top             =   4365
      Width           =   2580
   End
   Begin VB.ListBox lstMembers 
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
      Height          =   1395
      Left            =   1530
      TabIndex        =   0
      Top             =   1590
      Width           =   2595
   End
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   5400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmParty.frx":1EB2D
      PICF            =   "frmParty.frx":1EB49
      PICH            =   "frmParty.frx":1EB65
      PICV            =   "frmParty.frx":1EB81
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
   Begin GSZAOCliente.uAOButton cAgregar 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Agregar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmParty.frx":1EB9D
      PICF            =   "frmParty.frx":1EBB9
      PICH            =   "frmParty.frx":1EBD5
      PICV            =   "frmParty.frx":1EBF1
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
   Begin GSZAOCliente.uAOButton cLiderGrupo 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Lider Grupo"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmParty.frx":1EC0D
      PICF            =   "frmParty.frx":1EC29
      PICH            =   "frmParty.frx":1EC45
      PICV            =   "frmParty.frx":1EC61
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
   Begin GSZAOCliente.uAOButton cExpulsar 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Expulsar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmParty.frx":1EC7D
      PICF            =   "frmParty.frx":1EC99
      PICH            =   "frmParty.frx":1ECB5
      PICV            =   "frmParty.frx":1ECD1
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
   Begin GSZAOCliente.uAOButton cDisolver 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   5400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Disolver"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmParty.frx":1ECED
      PICF            =   "frmParty.frx":1ED09
      PICH            =   "frmParty.frx":1ED25
      PICV            =   "frmParty.frx":1ED41
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
   Begin VB.Label lblTotalExp 
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Height          =   255
      Left            =   3075
      TabIndex        =   3
      Top             =   3150
      Width           =   1335
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmParty.frm
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

Private clsFormulario As clsFormMovementManager

Private sPartyChat As String
Private Const LEADER_FORM_HEIGHT As Integer = 6015
Private Const NORMAL_FORM_HEIGHT As Integer = 4455
Private Const OFFSET_BUTTONS As Integer = 43 ' pixels


Private Sub cAgregar_Click()
    Call Audio.PlayWave(SND_CLICK)
    If Len(txtToAdd) > 0 Then
        If Not IsNumeric(txtToAdd) Then
            Call WritePartyAcceptMember(Trim$(txtToAdd.Text))
            Unload Me
            Call WriteRequestPartyForm
        End If
    End If
End Sub

Private Sub cCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
End Sub

Private Sub cDisolver_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WritePartyLeave
    Unload Me
End Sub

Private Sub cExpulsar_Click()
    If lstMembers.ListIndex < 0 Then Exit Sub
    Call Audio.PlayWave(SND_CLICK)

    Dim fName As String
    fName = GetName
    
    If LenB(fName) <> 0 Then
        Call WritePartyKick(fName)
        Unload Me
        ' Para que no llame al form si disolvió la party
        If UCase$(fName) <> UCase$(UserName) Then Call WriteRequestPartyForm
    End If
End Sub

Private Sub cLiderGrupo_Click()
    If lstMembers.ListIndex < 0 Then Exit Sub
    Call Audio.PlayWave(SND_CLICK)
    
    Dim sName As String
    sName = GetName
    
    If LenB(sName) <> 0 Then
        Call WritePartySetLeader(sName)
        Unload Me
        Call WriteRequestPartyForm
    End If
End Sub

Private Sub cSalirParty_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WritePartyLeave
    Unload Me
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    lstMembers.Clear
        
    If EsPartyLeader Then
        Me.Picture = LoadPicture(DirGUI & "frmPartyLider.jpg")
        Me.Height = LEADER_FORM_HEIGHT
    Else
        Me.Picture = LoadPicture(DirGUI & "frmPartyMiembro.jpg")
        Me.Height = NORMAL_FORM_HEIGHT
    End If
    
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
                           
    ' Botones visibles solo para el lider
    cExpulsar.Visible = EsPartyLeader
    cLiderGrupo.Visible = EsPartyLeader
    txtToAdd.Visible = EsPartyLeader
    cAgregar.Visible = EsPartyLeader
    
    cDisolver.Visible = EsPartyLeader
    cSalirParty.Visible = Not EsPartyLeader
    
    cSalirParty.Top = Me.ScaleHeight - OFFSET_BUTTONS
    cDisolver.Top = Me.ScaleHeight - OFFSET_BUTTONS
    cCerrar.Top = Me.ScaleHeight - OFFSET_BUTTONS
    MirandoParty = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoParty = False
End Sub

Private Function GetName() As String
'**************************************************************
'Author: ZaMa
'Last Modify Date: 27/07/2012 - ^[GS]^
'**************************************************************
    Dim sName As String
    
    sName = Trim$(mid$(lstMembers.List(lstMembers.ListIndex), 1, InStr(lstMembers.List(lstMembers.ListIndex), " (")))
    If Len(sName) > 0 Then GetName = sName
        
End Function

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 03/08/2012 - ^[GS]^
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        sPartyChat = vbNullString ' GSZAO no gastamos datos enviando un mensaje inutil...
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        sPartyChat = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(sPartyChat) <> 0 Then Call WritePartyMessage(sPartyChat)
        
        sPartyChat = vbNullString
        SendTxt.Text = vbNullString
        KeyCode = 0
        SendTxt.SetFocus
    End If
End Sub

Private Sub txtToAdd_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub txtToAdd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cAgregar_Click
End Sub


