VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "frmComerciar.frx":0000
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrReDraw 
      Interval        =   50
      Left            =   6000
      Top             =   1440
   End
   Begin VB.TextBox cantidad 
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
      Height          =   285
      Left            =   3150
      TabIndex        =   6
      Text            =   "1"
      Top             =   6570
      Width           =   630
   End
   Begin VB.PictureBox picInvUser 
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
      Height          =   3840
      Left            =   3945
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   5
      Top             =   1965
      Width           =   2400
   End
   Begin VB.PictureBox picInvNpc 
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
      Height          =   3840
      Left            =   600
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   4
      Top             =   1965
      Width           =   2400
   End
   Begin GSZAOCliente.uAOButton cVender 
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   6000
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   873
      TX              =   "Vender"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciar.frx":166CE
      PICF            =   "frmComerciar.frx":166EA
      PICH            =   "frmComerciar.frx":16706
      PICV            =   "frmComerciar.frx":16722
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cComprar 
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   6000
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   873
      TX              =   "Comprar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciar.frx":1673E
      PICF            =   "frmComerciar.frx":1675A
      PICH            =   "frmComerciar.frx":16776
      PICV            =   "frmComerciar.frx":16792
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GSZAOCliente.uAOButton cCerrar 
      Height          =   300
      Left            =   6150
      TabIndex        =   9
      Top             =   435
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      TX              =   "X"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciar.frx":167AE
      PICF            =   "frmComerciar.frx":167CA
      PICH            =   "frmComerciar.frx":167E6
      PICV            =   "frmComerciar.frx":16802
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   3
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000D7FF&
      Height          =   240
      Index           =   1
      Left            =   4695
      TabIndex        =   1
      Top             =   1320
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   240
      Index           =   0
      Left            =   2490
      TabIndex        =   0
      Top             =   720
      Width           =   105
   End
End
Attribute VB_Name = "frmComerciar"
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

Private clsFormulario As clsFormMovementManager

Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public LasActionBuy As Boolean
Private ClickNpcInv As Boolean
Private lIndex As Byte

Private Sub cantidad_Change()
On Error Resume Next
    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
    End If
    
    If ClickNpcInv Then
        If InvComNpc.SelectedItem <> 0 Then
            'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
            Label1(1).Caption = "Precio: " & CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, Val(cantidad.Text))  'No mostramos numeros reales
        End If
    Else
        If InvComUsu.SelectedItem <> 0 Then
            Label1(1).Caption = "Precio: " & CalculateBuyPrice(Inventario.Valor(InvComUsu.SelectedItem), Val(cantidad.Text))  'No mostramos numeros reales
        End If
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteCommerceEnd
    Unload Me
End Sub

Private Sub cComprar_Click()
    ' Debe tener seleccionado un item para comprarlo.
    If InvComNpc.SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
    
    Call Audio.PlayWave(SND_CLICK)
    
    LasActionBuy = True
    If UserGLD >= CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, Val(cantidad.Text)) Then
        Call WriteCommerceBuy(InvComNpc.SelectedItem, Val(cantidad.Text))
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "No tienes suficiente oro.", 2, 51, 223, 1, 1)
        Exit Sub
    End If
End Sub

Private Sub cVender_Click()
    ' Debe tener seleccionado un item para comprarlo.
    If InvComUsu.SelectedItem = 0 Then Exit Sub

    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
    
    Call Audio.PlayWave(SND_CLICK)
    
    LasActionBuy = False

    Call WriteCommerceSell(InvComUsu.SelectedItem, Val(cantidad.Text))
End Sub

Private Sub Form_Activate()
On Error Resume Next
    InvComUsu.DrawInv
    InvComNpc.DrawInv
End Sub

Private Sub Form_GotFocus()
On Error Resume Next
    InvComUsu.DrawInv
    InvComNpc.DrawInv
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    'Cargamos la interfase
    Me.Picture = LoadPicture(DirGUI & "frmComerciar.jpg")
    
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
    
    Comerciando = True

End Sub

''
' Calculates the selling price of an item (The price that a merchant will sell you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.

Private Function CalculateSellPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateSellPrice = CCur(objValue * 1000000) / 1000000 * objAmount + 0.5
    
    Exit Function
error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.Number
End Function
''
' Calculates the buying price of an item (The price that a merchant will buy you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.
Private Function CalculateBuyPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateBuyPrice = Fix(CCur(objValue * 1000000) / 1000000 * objAmount)
    
    Exit Function
error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.Number
End Function

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    InvComUsu.DrawInv
    InvComNpc.DrawInv
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Comerciando = False
End Sub

Private Sub picInvNpc_Click()
On Error Resume Next
    Dim ItemSlot As Byte
    
    ItemSlot = InvComNpc.SelectedItem
    If ItemSlot = 0 Then Exit Sub
    
    ClickNpcInv = True
    InvComUsu.DeselectItem
    Label1(0).Caption = NPCInventory(ItemSlot).Name
    Label1(1).Caption = "Precio: " & CalculateSellPrice(NPCInventory(ItemSlot).Valor, Val(cantidad.Text)) 'No mostramos numeros reales
    
    If NPCInventory(ItemSlot).amount <> 0 Then
    
        Select Case NPCInventory(ItemSlot).OBJType
            Case eObjType.otWeapon
                Label1(2).Caption = "Máx Golpe:" & NPCInventory(ItemSlot).MaxHit
                Label1(3).Caption = "Mín Golpe:" & NPCInventory(ItemSlot).MinHit
                Label1(2).Visible = True
                Label1(3).Visible = True
            Case eObjType.otArmadura, eObjType.otcasco, eObjType.otescudo
                Label1(2).Caption = "Máx Defensa:" & NPCInventory(ItemSlot).MaxDef
                Label1(3).Caption = "Mín Defensa:" & NPCInventory(ItemSlot).MinDef
                Label1(2).Visible = True
                Label1(3).Visible = True
            Case Else
                Label1(2).Visible = False
                Label1(3).Visible = False
        End Select
    Else
        Label1(2).Visible = False
        Label1(3).Visible = False
    End If
End Sub

Private Sub picInvUser_Click()
On Error Resume Next
    Dim ItemSlot As Byte
    
    ItemSlot = InvComUsu.SelectedItem
    
    If ItemSlot = 0 Then Exit Sub
    
    ClickNpcInv = False
    InvComNpc.DeselectItem
    
    Label1(0).Caption = Inventario.ItemName(ItemSlot)
    Label1(1).Caption = "Precio: " & CalculateBuyPrice(Inventario.Valor(ItemSlot), Val(cantidad.Text)) 'No mostramos numeros reales
    
    If Inventario.amount(ItemSlot) <> 0 Then
    
        Select Case Inventario.OBJType(ItemSlot)
            Case eObjType.otWeapon
                Label1(2).Caption = "Máx Golpe:" & Inventario.MaxHit(ItemSlot)
                Label1(3).Caption = "Mín Golpe:" & Inventario.MinHit(ItemSlot)
                Label1(2).Visible = True
                Label1(3).Visible = True
            Case eObjType.otArmadura, eObjType.otcasco, eObjType.otescudo
                Label1(2).Caption = "Máx Defensa:" & Inventario.MaxDef(ItemSlot)
                Label1(3).Caption = "Mín Defensa:" & Inventario.MinDef(ItemSlot)
                Label1(2).Visible = True
                Label1(3).Visible = True
            Case Else
                Label1(2).Visible = False
                Label1(3).Visible = False
        End Select
    Else
        Label1(2).Visible = False
        Label1(3).Visible = False
    End If
End Sub

Private Sub tmrReDraw_Timer()
On Error Resume Next
    InvComUsu.DrawInv
    InvComNpc.DrawInv
End Sub
