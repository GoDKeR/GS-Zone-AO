VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmComerciarUsu 
   BorderStyle     =   0  'None
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComerciarUsu.frx":0000
   ScaleHeight     =   590
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   667
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picInvOroProp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3450
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   930
      Width           =   960
   End
   Begin VB.TextBox txtAgregar 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   2220
      Width           =   1155
   End
   Begin VB.PictureBox picInvOroOfertaOtro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5610
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   5040
      Width           =   960
   End
   Begin VB.PictureBox picInvOfertaOtro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Left            =   6975
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   4
      Top             =   5040
      Width           =   2400
   End
   Begin VB.PictureBox picInvOfertaProp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Left            =   6960
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   3
      Top             =   930
      Width           =   2400
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
      Left            =   495
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   7965
      Width           =   6060
   End
   Begin VB.PictureBox picInvComercio 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Left            =   630
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   1
      Top             =   945
      Width           =   2400
   End
   Begin VB.PictureBox picInvOroOfertaProp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5610
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   930
      Width           =   960
   End
   Begin RichTextLib.RichTextBox CommerceConsole 
      Height          =   1620
      Left            =   495
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   6030
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   2858
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmComerciarUsu.frx":3001A
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
   Begin GSZAOCliente.uAOButton cCancelar 
      Height          =   735
      Left            =   600
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      TX              =   "Cancelar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciarUsu.frx":30098
      PICF            =   "frmComerciarUsu.frx":300B4
      PICH            =   "frmComerciarUsu.frx":300D0
      PICV            =   "frmComerciarUsu.frx":300EC
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
   Begin GSZAOCliente.uAOButton cConfirmar 
      Height          =   495
      Left            =   7560
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      TX              =   "Confirmar"
      ENAB            =   0   'False
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciarUsu.frx":30108
      PICF            =   "frmComerciarUsu.frx":30124
      PICH            =   "frmComerciarUsu.frx":30140
      PICV            =   "frmComerciarUsu.frx":3015C
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
   Begin GSZAOCliente.uAOButton cAceptar 
      Height          =   495
      Left            =   6840
      TabIndex        =   11
      Top             =   8160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      TX              =   "Aceptar"
      ENAB            =   0   'False
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciarUsu.frx":30178
      PICF            =   "frmComerciarUsu.frx":30194
      PICH            =   "frmComerciarUsu.frx":301B0
      PICV            =   "frmComerciarUsu.frx":301CC
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
   Begin GSZAOCliente.uAOButton cRechazar 
      Height          =   495
      Left            =   8280
      TabIndex        =   12
      Top             =   8160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      TX              =   "Rechazar"
      ENAB            =   0   'False
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciarUsu.frx":301E8
      PICF            =   "frmComerciarUsu.frx":30204
      PICH            =   "frmComerciarUsu.frx":30220
      PICV            =   "frmComerciarUsu.frx":3023C
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
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   255
      Left            =   9690
      TabIndex        =   13
      Top             =   15
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      TX              =   "X"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciarUsu.frx":30258
      PICF            =   "frmComerciarUsu.frx":30274
      PICH            =   "frmComerciarUsu.frx":30290
      PICV            =   "frmComerciarUsu.frx":302AC
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
   Begin GSZAOCliente.uAOButton cAgregar 
      Height          =   375
      Left            =   4500
      TabIndex        =   14
      Top             =   1800
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      TX              =   "Agregar >>"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciarUsu.frx":302C8
      PICF            =   "frmComerciarUsu.frx":302E4
      PICH            =   "frmComerciarUsu.frx":30300
      PICV            =   "frmComerciarUsu.frx":3031C
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
   Begin GSZAOCliente.uAOButton cQuitar 
      Height          =   375
      Left            =   4500
      TabIndex        =   15
      Top             =   2640
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      TX              =   "<< Quitar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciarUsu.frx":30338
      PICF            =   "frmComerciarUsu.frx":30354
      PICH            =   "frmComerciarUsu.frx":30370
      PICV            =   "frmComerciarUsu.frx":3038C
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
End
Attribute VB_Name = "frmComerciarUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmComerciarUsu.frm
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

Private Const GOLD_OFFER_SLOT As Byte = INV_OFFER_SLOTS + 1

Private sCommerceChat As String

Private Sub cAceptar_Click()

    Call Audio.PlayWave(SND_CLICK)
    Call WriteUserCommerceOk
    HabilitarAceptarRechazar False
    
End Sub

Private Sub cAgregar_Click()

    Call Audio.PlayWave(SND_CLICK)
   
    ' No tiene seleccionado ningun item
    If InvComUsu.SelectedItem = 0 Then
        Call PrintCommerceMsg("¡No tienes ningún item seleccionado!", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    ' Numero invalido
    If Not IsNumeric(txtAgregar.Text) Then Exit Sub
    
    HabilitarConfirmar True
    
    Dim OfferSlot As Byte
    Dim amount As Long
    Dim InvSlot As Byte
        
    With InvComUsu
        If .SelectedItem = FLAGORO Then
            If Val(txtAgregar.Text) > InvOroComUsu(0).amount(1) Then
                Call PrintCommerceMsg("¡No tienes esa cantidad!", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            amount = InvOroComUsu(1).amount(1) + Val(txtAgregar.Text)
    
            ' Le aviso al otro de mi cambio de oferta
            Call WriteUserCommerceOffer(FLAGORO, Val(txtAgregar.Text), GOLD_OFFER_SLOT)
            
            ' Actualizo los inventarios
            Call InvOroComUsu(0).ChangeSlotItemAmount(1, InvOroComUsu(0).amount(1) - Val(txtAgregar.Text))
            Call InvOroComUsu(1).ChangeSlotItemAmount(1, amount)
            
            Call PrintCommerceMsg("¡Agregaste " & Val(txtAgregar.Text) & " moneda" & IIf(Val(txtAgregar.Text) = 1, "", "s") & " de oro a tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
            
        ElseIf .SelectedItem > 0 Then
             If Val(txtAgregar.Text) > .amount(.SelectedItem) Then
                Call PrintCommerceMsg("¡No tienes esa cantidad!", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
             
            OfferSlot = CheckAvailableSlot(.SelectedItem, Val(txtAgregar.Text))
            
            ' Hay espacio o lugar donde sumarlo?
            If OfferSlot > 0 Then
            
                Call PrintCommerceMsg("¡Agregaste " & Val(txtAgregar.Text) & " " & .ItemName(.SelectedItem) & " a tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
                
                ' Le aviso al otro de mi cambio de oferta
                Call WriteUserCommerceOffer(.SelectedItem, Val(txtAgregar.Text), OfferSlot)
                
                ' Actualizo el inventario general de comercio
                Call .ChangeSlotItemAmount(.SelectedItem, .amount(.SelectedItem) - Val(txtAgregar.Text))
                
                amount = InvOfferComUsu(0).amount(OfferSlot) + Val(txtAgregar.Text)
                
                ' Actualizo los inventarios
                If InvOfferComUsu(0).OBJIndex(OfferSlot) > 0 Then
                    ' Si ya esta el item, solo actualizo su cantidad en el invenatario
                    Call InvOfferComUsu(0).ChangeSlotItemAmount(OfferSlot, amount)
                Else
                    InvSlot = .SelectedItem
                    ' Si no agrego todo
                    Call InvOfferComUsu(0).SetItem(OfferSlot, .OBJIndex(InvSlot), _
                                                    amount, 0, .GrhIndex(InvSlot), .OBJType(InvSlot), _
                                                    .MaxHit(InvSlot), .MinHit(InvSlot), .MaxDef(InvSlot), .MinDef(InvSlot), _
                                                    .Valor(InvSlot), .ItemName(InvSlot))
                End If
            End If
        End If
    End With
    
End Sub

Private Sub cCancelar_Click()

    Call Audio.PlayWave(SND_CLICK)
    Call WriteUserCommerceEnd
    Unload Me
    
End Sub

Private Sub cCerrar_Click()

    Call Audio.PlayWave(SND_CLICK)
    Call cCancelar_Click
    
End Sub

Private Sub cConfirmar_Click()

    Call Audio.PlayWave(SND_CLICK)
    HabilitarConfirmar False
    cAgregar.Enabled = False
    cQuitar.Enabled = False
    txtAgregar.Enabled = False
    
    Call PrintCommerceMsg("¡Has confirmado tu oferta! Ya no puedes cambiarla.", FontTypeNames.FONTTYPE_CONSEJERO)
    Call WriteUserCommerceConfirm

End Sub

Private Sub cQuitar_Click()

    Call Audio.PlayWave(SND_CLICK)
    
    Dim amount As Long
    Dim InvComSlot As Byte

    ' No tiene seleccionado ningun item
    If InvOfferComUsu(0).SelectedItem = 0 Then
        Call PrintCommerceMsg("¡No tienes ningún ítem seleccionado!", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    ' Numero invalido
    If Not IsNumeric(txtAgregar.Text) Then Exit Sub

    ' Comparar con el inventario para distribuir los items
    If InvOfferComUsu(0).SelectedItem = FLAGORO Then
        amount = IIf(Val(txtAgregar.Text) > InvOroComUsu(1).amount(1), InvOroComUsu(1).amount(1), Val(txtAgregar.Text))
        ' Estoy quitando, paso un valor negativo
        amount = amount * (-1)
        
        ' No tiene sentido que se quiten 0 unidades
        If amount <> 0 Then
            ' Le aviso al otro de mi cambio de oferta
            Call WriteUserCommerceOffer(FLAGORO, amount, GOLD_OFFER_SLOT)
            
            ' Actualizo los inventarios
            Call InvOroComUsu(0).ChangeSlotItemAmount(1, InvOroComUsu(0).amount(1) - amount)
            Call InvOroComUsu(1).ChangeSlotItemAmount(1, InvOroComUsu(1).amount(1) + amount)
        
            Call PrintCommerceMsg("¡¡Quitaste " & amount * (-1) & " moneda" & IIf(Val(txtAgregar.Text) = 1, "", "s") & " de oro de tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
        End If
    Else
        amount = IIf(Val(txtAgregar.Text) > InvOfferComUsu(0).amount(InvOfferComUsu(0).SelectedItem), _
                    InvOfferComUsu(0).amount(InvOfferComUsu(0).SelectedItem), Val(txtAgregar.Text))
        ' Estoy quitando, paso un valor negativo
        amount = amount * (-1)
        
        ' No tiene sentido que se quiten 0 unidades
        If amount <> 0 Then
            With InvOfferComUsu(0)
                
                Call PrintCommerceMsg("¡¡Quitaste " & amount * (-1) & " " & .ItemName(.SelectedItem) & " de tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
    
                ' Le aviso al otro de mi cambio de oferta
                Call WriteUserCommerceOffer(0, amount, .SelectedItem)
            
                ' Actualizo el inventario general
                Call UpdateInvCom(.OBJIndex(.SelectedItem), Abs(amount))
                 
                 ' Actualizo el inventario de oferta
                 If .amount(.SelectedItem) + amount = 0 Then
                     ' Borro el item
                     Call .SetItem(.SelectedItem, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
                 Else
                     ' Le resto la cantidad deseada
                     Call .ChangeSlotItemAmount(.SelectedItem, .amount(.SelectedItem) + amount)
                 End If
            End With
        End If
    End If
    
    ' Si quito todos los items de la oferta, no puede confirmarla
    If Not HasAnyItem(InvOfferComUsu(0)) And _
       Not HasAnyItem(InvOroComUsu(1)) Then HabilitarConfirmar (False)

End Sub

Private Sub cRechazar_Click()
    
    Call Audio.PlayWave(SND_CLICK)
    Call WriteUserCommerceReject

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Comerciando = False

End Sub

Private Sub Form_Load()
'Last Modification: 11/08/2014 - ^[GS]^
'**********************************
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Me.Picture = LoadPicture(DirGUI & "frmComerciarUsu.jpg")
    
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
    
    Call PrintCommerceMsg("> Una vez termines de formar tu oferta, debes presionar en ""Confirmar"", tras lo cual ya no podrás modificarla.", FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> Luego que el otro usuario confirme su oferta, podrás aceptarla o rechazarla. Si la rechazas, se terminará el comercio.", FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> Cuando ambos acepten la oferta del otro, se realizará el intercambio.", FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> Si se intercambian más ítems de los que pueden entrar en tu inventario, es probable que caigan al suelo, así que presta mucha atención a esto.", FontTypeNames.FONTTYPE_GUILDMSG)
    
End Sub

Private Sub Form_LostFocus()
    
    Me.SetFocus

End Sub

Private Sub SubtxtAgregar_Change()
    
    If Val(txtAgregar.Text) < 1 Then txtAgregar.Text = "1"

    If Val(txtAgregar.Text) > 2147483647 Then txtAgregar.Text = "2147483647"

End Sub

Private Sub picInvComercio_Click()
    
    Call InvOroComUsu(0).DeselectItem

End Sub

Private Sub picInvOfertaProp_Click()
    
    InvOroComUsu(1).DeselectItem

End Sub

Private Sub picInvOroOfertaOtro_Click()
    
    ' No se puede seleccionar el oro que oferta el otro :P
    InvOroComUsu(2).DeselectItem

End Sub

Private Sub picInvOroOfertaProp_Click()
    
    InvOfferComUsu(0).SelectGold

End Sub

Private Sub picInvOroProp_Click()
    
    InvComUsu.SelectGold

End Sub

Private Sub txtAgregar_Change()
'**************************************************************
'Author: Unknown
'Last Modification: 05/11/2011 - ^[GS]^
'**************************************************************
    
    'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
    Dim i As Long
    Dim tempstr As String
    Dim CharAscii As Integer
    
    For i = 1 To Len(txtAgregar.Text)
        CharAscii = Asc(mid$(txtAgregar.Text, i, 1))
        
        If CharAscii >= 48 And CharAscii <= 57 Then
            tempstr = tempstr & Chr$(CharAscii)
        End If
    Next i
    
    If tempstr <> txtAgregar.Text Then
        'We only set it if it's different, otherwise the event will be raised
        'constantly and the client will crush
        txtAgregar.Text = tempstr
    End If

End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 13/03/2012 - ^[GS]^
'**************************************************************
    
    If Len(SendTxt.Text) > 160 Then
        sCommerceChat = vbNullString ' GSZAO
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
        
        sCommerceChat = SendTxt.Text
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
        If LenB(sCommerceChat) <> 0 Then Call WriteCommerceChat(sCommerceChat)
        
        sCommerceChat = vbNullString
        SendTxt.Text = vbNullString
        KeyCode = 0
    End If
    
End Sub


Private Sub txtAgregar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or _
            KeyCode = vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
        KeyCode = 0
    End If

End Sub

Private Sub txtAgregar_KeyPress(KeyAscii As Integer)

    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or _
            KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
        'txtCant = KeyCode
        KeyAscii = 0
    End If
    
End Sub

Private Function CheckAvailableSlot(ByVal InvSlot As Byte, ByVal amount As Long) As Byte
'***************************************************
'Author: ZaMa
'Last Modify Date: 30/11/2009
'Search for an available slot to put an item. If found returns the slot, else returns 0.
'***************************************************
    Dim slot As Long
On Error GoTo Err
    ' Primero chequeo si puedo sumar esa cantidad en algun slot que ya tenga ese item
    For slot = 1 To INV_OFFER_SLOTS
        If InvComUsu.OBJIndex(InvSlot) = InvOfferComUsu(0).OBJIndex(slot) Then
            If InvOfferComUsu(0).amount(slot) + amount <= MAX_INVENTORY_OBJS Then
                ' Puedo sumarlo aca
                CheckAvailableSlot = slot
                Exit Function
            End If
        End If
    Next slot
    
    ' No lo puedo sumar, me fijo si hay alguno vacio
    For slot = 1 To INV_OFFER_SLOTS
        If InvOfferComUsu(0).OBJIndex(slot) = 0 Then
            ' Esta vacio, lo dejo aca
            CheckAvailableSlot = slot
            Exit Function
        End If
    Next slot
    Exit Function
Err:
    Debug.Print "Slot: " & slot
    
End Function

Public Sub UpdateInvCom(ByVal OBJIndex As Integer, ByVal amount As Long)

    Dim slot As Byte
    Dim RemainingAmount As Long
    Dim DifAmount As Long
    
    RemainingAmount = amount
    
    For slot = 1 To MAX_INVENTORY_SLOTS
        
        If InvComUsu.OBJIndex(slot) = OBJIndex Then
            DifAmount = Inventario.amount(slot) - InvComUsu.amount(slot)
            If DifAmount > 0 Then
                If RemainingAmount > DifAmount Then
                    RemainingAmount = RemainingAmount - DifAmount
                    Call InvComUsu.ChangeSlotItemAmount(slot, Inventario.amount(slot))
                Else
                    Call InvComUsu.ChangeSlotItemAmount(slot, InvComUsu.amount(slot) + RemainingAmount)
                    Exit Sub
                End If
            End If
        End If
    Next slot
    
End Sub

Public Sub PrintCommerceMsg(ByRef msg As String, ByVal FontIndex As Integer)
    
    With FontTypes(FontIndex)
        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, msg, .Red, .Green, .Blue, .bold, .italic)
    End With
    
End Sub

Public Function HasAnyItem(ByRef Inventory As clsGraphicalInventory) As Boolean

    Dim slot As Long
    
    For slot = 1 To Inventory.MaxObjs
        If Inventory.amount(slot) > 0 Then HasAnyItem = True: Exit Function
    Next slot
    
End Function

Public Sub HabilitarConfirmar(ByVal Habilitar As Boolean)

    cConfirmar.Enabled = Habilitar
    
End Sub

Public Sub HabilitarAceptarRechazar(ByVal Habilitar As Boolean)

    cAceptar.Enabled = Habilitar
    cRechazar.Enabled = Habilitar
    
End Sub
