VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Private InventoryOffset As Long             'Number of lines we scrolled down from topmost
Private InvSelectedItem As Long             'Currently selected item
Private RGB(3) As Long
Private RGB_A(3) As Long

Private LastRender As Long  ' GSZAO
Private FixY As Boolean ' GSZAO (cuando esta True, es para dibujar el inventario del jugador, sino es otro inventario, de comercio, etc)
Private bItemEmpty As Boolean ' GSZAO
Private ItemEmpty As Direct3DTexture8 ' GSZAO
Private rgbRed(3) As Long ' GSZAO

Private slotDragged As Integer
 
Public Event dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)

Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer

Private ShowValue As Boolean
Private ImgContainer As Boolean
Private AlphaItems As Byte

Private MouseX As Single
Private MouseY As Single

Private Inventory() As Inventory    'User's inventory

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Public Sub ScrollInventory(ByVal Up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 14/01/2010
'Scrolls the graphical inventory up / down
'14/01/2010: ZaMa - Now redraws inventory only if needed
'***************************************************

    'Only move if there is something to show up /down
    If CanScroll(Up) Then
        InventoryOffset = InventoryOffset + IIf(Up, 1, -1)
        
        'Redraw the inventory if needed
        DrawInventory
    End If

End Sub

Public Function CanScroll(ByVal Up As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 14/01/2010
'Determines wether inventory can be scrolled up / down
'***************************************************
    If Up Then
        CanScroll = (InventoryOffset + (InventoryWindow.ScaleHeight \ ownTilePixelHeight)) * _
                    (InventoryWindow.ScaleWidth \ ownTilePixelWidth) < UBound(Inventory)
    Else
        CanScroll = InventoryOffset > 0
    End If
    
End Function

Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
    'Redraw old item to deselect it
    If prevSelectedItem <> FLAGORO Then
        Call DrawInventory
    End If
End Sub

Public Sub DeselectItem()
'***************************************************
'Author: ZaMa
'Last Modify Date: 30/11/2009
'Deselect the currently selected item
'***************************************************
    If InvSelectedItem = 0 Then Exit Sub
    
    Dim ClearSlot As Byte
    
    ClearSlot = InvSelectedItem
    
    'Select nothing
    InvSelectedItem = 0
    
    'Redraw old item to deselect it
    Call DrawInventory
End Sub

Public Sub ChangeSlotItemAmount(ByVal slot As Byte, ByVal amount As Long)
'***************************************************
'Author: ZaMa
'Created Date: 30/11/2009
'Last Modify Date: 09/12/2009 (Budi)
'Change the amount of the required item, and update the slot.
'***************************************************
    
    Inventory(slot).amount = amount
    
    'Update Amount
    Call DrawInventory
End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = Inventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = Inventory(slot).MinHit
End Property

Public Property Get MaxDef(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    MaxDef = Inventory(slot).MaxDef
End Property

Public Property Get MinDef(ByVal slot As Byte) As Integer
'***************************************************
'Author: Budi
'Last Modify Date: 02/1209
'Retrieves the defense of the selected item
'***************************************************
    MinDef = Inventory(slot).MinDef
End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    GrhIndex = Inventory(slot).GrhIndex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = Inventory(slot).Valor
End Property

Public Property Get amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        amount = UserGLD
    ElseIf slot >= LBound(Inventory) And slot <= UBound(Inventory) Then
        amount = Inventory(slot).amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = Inventory(slot).OBJIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = Inventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = Inventory(slot).Name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = Inventory(slot).Equipped
End Property

Public Property Get MaxObjs() As Byte
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Retrieves the capacity of the Inventory
'***************************************************
    MaxObjs = UBound(Inventory)
End Property

Public Sub SetMaxObjs(ByVal MaxObjs As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Set the capacity of the Inventary
'***************************************************
    If UBound(Inventory) = MaxObjs Then Exit Sub
    
    ReDim Preserve Inventory(1 To MaxObjs) As Inventory
End Sub

Public Sub ToggleShowValue(ByVal bShow As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 11/12
'Indicates if the amount of the items in the inventory is shown
'***************************************************
    ShowValue = bShow
    
    Dim slot As Integer
    ' Update Inventory
    For slot = 1 To Me.MaxObjs
        DrawInventory
    Next slot
    
End Sub

Public Sub ClearAllSlots() ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modify Date: 23/11/2011 - ^[GS]^
'Clear All Slots.
'***************************************************
On Error Resume Next

    Dim slot As Long
    For slot = 1 To Me.MaxObjs
        Call SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString)
    Next slot

End Sub

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Long, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal MaxDef As Integer, ByVal MinDef As Integer, ByVal eValor As Single, ByVal eName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 10/08/2014 - ^[GS]^
'Sets all data for a given inventory slot
'***************************************************
On Error GoTo Fallo:

    If slot < 1 Or slot > UBound(Inventory) Then Exit Sub
    
    With Inventory(slot)
        .amount = eAmount
        .MaxDef = MaxDef
        .MinDef = MinDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
    End With
    
    'If InventoryWindow.Visible = False Then InventoryWindow.Visible = True
    
    'Render inventory slot (we don't need to render the whole inventory)
    Call DrawInventory
    
    Exit Sub
Fallo:
    
    Call LogError("SetItem::Error " & Err.Number & " - " & Err.Description & " - Slot: " & slot & " - OBJIndex: " & eOBJIndex)

End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 08/08/2012 - ^[GS]^
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    If (temp_y > 0) And FixY = True Then
        TempItem = TempItem + (temp_y) ' Corrección de filas
    End If
    
    'Make sure it's within limits
    If TempItem <= UBound(Inventory) Then
        'Make sure slot isn't empty
        If Inventory(TempItem).GrhIndex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If
    
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        ClickItem = TempItem
    End If
    
End Function

Private Sub DrawInventory()
On Error Resume Next
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 02/05/2013 - ^[GS]^
'Renders the inventory to the given PictureBox
'***************************************************
    Dim loopC As Long
    Dim SrcRect As RECT
    Dim SrcRecti As RECT
    Dim TempRect As RECT
    Dim destRect As RECT
    
    With TempRect
        .Bottom = InventoryWindow.Height
        .Right = InventoryWindow.Width
    End With
    
    destRect = TempRect
       
    'Clear the inventory window
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
    DirectDevice.BeginScene
    
    For loopC = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 To UBound(Inventory)
            
            startdX = ((loopC - 1) Mod (frmMain.PicInv.Width / 32)) * 32
            startdY = ((loopC - 1) \ (frmMain.PicInv.Width / 32)) * 32
            
            If Inventory(loopC).GrhIndex Then
            
                'Get source rect
                With SrcRect
                    '.left = GrhData(Inventory(loopC).GrhIndex).sX
                    '.top = GrhData(Inventory(loopC).GrhIndex).sY
                    .Right = .Left + ownTilePixelWidth
                    .Bottom = .Top + ownTilePixelHeight
                End With
                
                With SrcRecti
                    .Left = 64
                    .Top = 0
                    .Right = .Left + ownTilePixelWidth
                    .Bottom = .Top + ownTilePixelHeight
                End With
                
                If bItemEmpty = True Then ' GSZAO Fondo!
                    Device_Textured_Render startdX, startdY, ItemEmpty, SrcRect, RGB, , False, True
                End If
                
                If frmMain.UsandoDrag = False Then
                    
                    'Render a box around the selected item
                    If InvSelectedItem = loopC Then
                        If ImgContainer Then
                            Device_Textured_Render startdX, startdY, SurfaceDB.Surface(CLng(2)), SrcRecti, RGB, False, True, True
                        End If
                    End If
                                                                    
                    Device_Textured_Render startdX, startdY, SurfaceDB.Surface(CLng(GrhData(Inventory(loopC).GrhIndex).FileNum)), SrcRect, RGB, , bItemEmpty, True
                    
                    'If equipped we render "E"
                    If Inventory(loopC).Equipped Then
                        DrawText startdX + 23, startdY + 18, "E", -256, 255
                    End If
                                    
                    'Render the item grh and the amount
                    DrawText startdX, startdY, Inventory(loopC).amount, -1, 255
                    
                Else
                    
                    Device_Textured_Render startdX, startdY, SurfaceDB.Surface(CLng(GrhData(Inventory(loopC).GrhIndex).FileNum)), SrcRect, RGB_A, , bItemEmpty, True
                    
                    'If equipped we render "E"
                    If Inventory(loopC).Equipped Then
                        DrawText startdX + 23, startdY + 18, "E", -256, 150
                    End If
                                    
                    'Render the item grh and the amount
                    DrawText startdX, startdY, Inventory(loopC).amount, -1, 150
                    
                End If
                
            Else
            
                If bItemEmpty = True Then ' GSZAO Fondo!
                    If startdY >= 128 Then ' las 2 últimas hileras en rojo, porque es obligatoria la mochila
                        Device_Textured_Render startdX, startdY, ItemEmpty, SrcRect, rgbRed, , True, True
                    Else
                        Device_Textured_Render startdX, startdY, ItemEmpty, SrcRect, RGB, , False, True
                    End If
                End If
            End If
        Next loopC
    DirectDevice.EndScene
    DirectDevice.Present destRect, ByVal 0, InventoryWindow.hwnd, ByVal 0
    
End Sub

Public Sub DrawInv()

    If GetTickCount - LastRender > 56 Then ' GSZAO
        LastRender = GetTickCount
        DoEvents
        Call DrawInventory
    End If
    
End Sub

Public Sub Initialize(ByRef DirectD3D As D3DX8, ByRef InvPic As PictureBox, ByVal MaxObjs As Byte, _
                            Optional ByVal FontSize As Integer = 7, Optional ByVal TileWidth As Integer = 32, _
                            Optional ByVal TileHeight As Integer = 32, Optional ByVal startX As Integer = 0, _
                            Optional ByVal startY As Integer = 0, Optional ByVal bImgContainer As Boolean = True, _
                            Optional ByVal bShowText As Boolean = True, Optional ByVal bFixY As Boolean = False)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 03/12/09
'Sets the reference to the picture box on which to render the inventory
'03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
'***************************************************
    Set InventoryWindow = InvPic
    
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
    ReDim Inventory(1 To MaxObjs) As Inventory
    
    'initialize DX stuff
    'Dim SurfaceDesc As DDSURFACEDESC2
    
    'Make sure DirectD3D was correctly initialized
    If DirectD3D Is Nothing Then Exit Sub

    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    
    FixY = bFixY ' GSZAO
    ImgContainer = bImgContainer
    ShowValue = bShowText
    
    RGB(2) = D3DColorARGB(255, 255, 255, 255)
    RGB(3) = D3DColorARGB(255, 255, 255, 255)
    RGB(1) = D3DColorARGB(255, 255, 255, 255)
    RGB(0) = D3DColorARGB(255, 255, 255, 255)
    
    RGB_A(2) = D3DColorARGB(255, 150, 150, 150)
    RGB_A(3) = D3DColorARGB(255, 150, 150, 150)
    RGB_A(1) = D3DColorARGB(255, 150, 150, 150)
    RGB_A(0) = D3DColorARGB(255, 150, 150, 150)
    
    'TODO : Fonts should be in a separate class / collection
    Dim Font As New StdFont
    Dim Ifnt As IFont
    
    Font.Name = "Verdana"
    Font.bold = True
    Font.italic = False
    Font.Size = FontSize
    Font.Underline = False
    Font.Strikethrough = False
    
    Set Ifnt = Font
    '\TODO
    
    If ImgContainer Then _
        InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
        
    ' Cargamos el grafico por defecto...
    If (FileExist(DirGUI & "mochila.jpg", vbArchive) = True) And (FixY = True) Then ' Solo en FixY=True
        Set ItemEmpty = DirectD3D.CreateTextureFromFileEx(DirectDevice, DirGUI & "mochila.jpg", _
                D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &H0, ByVal 0, ByVal 0)        ' GSZAO
                'ARGB(255, 20, 20, 60)
        rgbRed(0) = D3DColorARGB(255, 255, 180, 180)
        rgbRed(1) = D3DColorARGB(255, 255, 180, 180)
        rgbRed(2) = D3DColorARGB(255, 255, 180, 180)
        rgbRed(3) = D3DColorARGB(255, 255, 180, 180)
        bItemEmpty = True
    Else
        bItemEmpty = False
    End If

End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 30/06/2012 - ^[GS]^
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
        
    MouseX = X
    MouseY = Y
        
    If Button = vbRightButton Then
        
        Call DrawInventory
        
    Else

        temp_x = X \ ownTilePixelWidth
        temp_y = Y \ ownTilePixelHeight
        
        TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
        
        If (temp_y > 0) And FixY = True Then
            TempItem = TempItem + (temp_y) ' Corrección de filas
        End If
        
        If TempItem <= UBound(Inventory) Then
            If FixY = True Then ' inventario del usuario
                frmMain.lblItem.Caption = Inventory(TempItem).Name ' Que nos diga el nombre del objeto cuando pasamos el mouse sobre él
            Else ' "otros" inventarios
                InventoryWindow.ToolTipText = Inventory(TempItem).Name ' Que nos diga el nombre del objeto cuando pasamos el mouse sobre él
            End If
        End If
    
    End If
    
    Call DrawInventory
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
   
    'Exit if it got outside the control's area
    If X < 0 Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        X = 0
    ElseIf Y < 0 Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        Y = 0
    ElseIf X > InventoryWindow.Width Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        X = InventoryWindow.Width
    ElseIf Y > InventoryWindow.Height Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        Y = InventoryWindow.Height
    End If
   
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
   
    prevSelItem = InvSelectedItem
   
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
       
    If Not frmMain.UsabaDrag Then InvSelectedItem = ClickItem(CInt(X), CInt(Y))
           
    Call DrawInventory
        
    If Button = vbRightButton Then
        If InvSelectedItem <> slotDragged Then
            'RaiseEvent dragDone(slotDragged, InvSelectedItem)
            Call frmMain.picInv_dragDone(slotDragged, InvSelectedItem)
            frmMain.UsandoDrag = False
        Else
            frmMain.UsandoDrag = False
        End If
    End If
        

End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer

    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
       
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
       
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
       
    If Button = vbRightButton Then
        TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
        
        If (temp_y > 0) And FixY = True Then
            TempItem = TempItem + (temp_y) ' Corrección de filas
        End If
        
        slotDragged = TempItem
    End If

End Sub
