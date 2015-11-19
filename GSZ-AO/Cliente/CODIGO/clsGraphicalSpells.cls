VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalSpells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private SpellOffset As Long             'Number of lines we scrolled down from topmost
Public SpellSelectedItem As Long              'Currently selected item
Private RGB(3) As Long
Private rgbRed(3) As Long
Private TempItemSelect As Long

Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer

Private slotDragged As Integer
 
Public Event dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)

Private ShowValue As Boolean
Private ImgContainer As Boolean
Private ShowInfo As Boolean

Private LastRender As Long  ' GSZAO
Private bSpellEmpty As Boolean  ' GSZAO
Private SpellEmpty As Direct3DTexture8  ' GSZAO

Private WithEvents SpellWindow As PictureBox    'Placeholder where to render the inventory
Attribute SpellWindow.VB_VarHelpID = -1

Private Sub DrawSpells()

    Dim loopC As Long
    Dim SrcRect As RECT
    Dim SrcRecti As RECT
    Dim TempRect As RECT
    Dim destRect As RECT
    
    With TempRect
        .Bottom = SpellWindow.Height
        .Right = SpellWindow.Width
    End With
    
    destRect = TempRect
    
    'Clear the inventory window
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
    DirectDevice.BeginScene
    For loopC = SpellOffset * (SpellWindow.ScaleWidth \ ownTilePixelWidth) + 1 To UBound(UserHechizos)
        
        startdX = ((loopC - 1) Mod (SpellWindow.Width / 32)) * 32
        startdY = ((loopC - 1) \ (SpellWindow.Width / 32)) * 32
        
        'Get source rect
        With SrcRect
            .Right = .Left + ownTilePixelWidth
            .Bottom = .Top + ownTilePixelHeight
        End With
            
        With SrcRecti
            .Left = 64
            .Top = 0
            .Right = .Left + ownTilePixelWidth
            .Bottom = .Top + ownTilePixelHeight
        End With
        
        If bSpellEmpty = True Then ' GSZAO Fondo!
            Device_Textured_Render startdX, startdY, SpellEmpty, SrcRect, RGB, , False, True
        End If
        
        If UserHechizos(loopC).Name <> "(None)" Then
            If UserHechizos(loopC).GrhIndex = 0 Then
                Device_Textured_Render startdX, startdY, SurfaceDB.Surface(CLng(22031)), SrcRect, RGB, , bSpellEmpty, True
            Else
                With GrhData(UserHechizos(loopC).GrhIndex)
                    SrcRect.Left = .sX
                    SrcRect.Top = .sY
                    SrcRect.Right = SrcRect.Left + .pixelWidth
                    SrcRect.Bottom = SrcRect.Top + .pixelHeight
                    Device_Textured_Render startdX, startdY, SurfaceDB.Surface(CLng(.FileNum)), SrcRect, RGB, , bSpellEmpty, True
                End With
            End If
        End If
        
        If SpellSelectedItem = loopC Then
            Device_Textured_Render startdX, startdY, SurfaceDB.Surface(CLng(2)), SrcRecti, RGB, False, True, True
        End If
        
    Next loopC
        
    DirectDevice.EndScene
    DirectDevice.Present destRect, ByVal 0, SpellWindow.hwnd, ByVal 0
    
End Sub

Public Sub Initialize(ByRef DirectD3D As D3DX8, ByRef Picture As PictureBox, ByVal MaxSpells As Byte, _
                            Optional ByVal TileWidth As Integer = 32, _
                            Optional ByVal TileHeight As Integer = 32, Optional ByVal startX As Integer = 0, _
                            Optional ByVal startY As Integer = 0)

    Set SpellWindow = Picture
    
    SpellWindow.AutoRedraw = True
    
    SpellWindow.ScaleMode = 3
    
    If DirectD3D Is Nothing Then Exit Sub

    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    
    RGB(0) = D3DColorARGB(255, 255, 255, 255)
    RGB(1) = D3DColorARGB(255, 255, 255, 255)
    RGB(2) = D3DColorARGB(255, 255, 255, 255)
    RGB(3) = D3DColorARGB(255, 255, 255, 255)
    
    rgbRed(0) = D3DColorARGB(255, 255, 10, 10)
    rgbRed(1) = D3DColorARGB(255, 255, 10, 10)
    rgbRed(2) = D3DColorARGB(255, 255, 10, 10)
    rgbRed(3) = D3DColorARGB(255, 255, 10, 10)
    
    ' Cargamos el grafico por defecto...
    If (FileExist(DirGUI & "hechizo.jpg", vbArchive) = True) Then
        Set SpellEmpty = DirectD3D.CreateTextureFromFileEx(DirectDevice, DirGUI & "hechizo.jpg", _
                D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &H55000000, ByVal 0, ByVal 0) ' GSZAO
        bSpellEmpty = True
    Else
        bSpellEmpty = False
    End If
End Sub

Public Sub RenderSpells()
    If GetTickCount - LastRender > 56 Then ' GSZAO
        LastRender = GetTickCount
        DoEvents
        Call DrawSpells
    End If
End Sub

Private Sub SpellWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer

    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > SpellWindow.Width Or Y > SpellWindow.Height Then Exit Sub
       
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
       
    TempItemSelect = ClickItem(CInt(X), CInt(Y))
    
    If Button = vbRightButton Then
        TempItem = temp_x + (temp_y + SpellOffset) * (SpellWindow.ScaleWidth \ ownTilePixelWidth) + 1
        
        If (temp_y > 0) Then
            TempItem = TempItem + (temp_y) ' Corrección de filas
        End If
        
        If TempItem <= UBound(UserHechizos) Then
            slotDragged = TempItem
        End If
        
    End If
End Sub

Private Sub SpellWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
    
    UsaMacro = False
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > SpellWindow.Width Or Y > SpellWindow.Height Then _
        Exit Sub
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + SpellOffset) * (SpellWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    If (temp_y > 0) Then
        TempItem = TempItem + (temp_y) ' Corrección de filas
    End If
    
    If TempItem <= UBound(UserHechizos) Then
        frmMain.lblItem.Caption = UserHechizos(TempItem).Name ' Que nos diga el nombre del objeto cuando pasamos el mouse sobre él
        'SpellWindow.ToolTipText = UserHechizos(TempItem).Name ' Que nos diga el nombre del hechizo cuando pasamos el mouse sobre él
    End If
    
    Call DrawSpells
End Sub

Private Sub SpellWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim prevSelItem As Long

    'Exit if it got outside the control's area
    If X < 0 Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        X = 0
    ElseIf Y < 0 Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        Y = 0
    ElseIf X > SpellWindow.Width Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        X = SpellWindow.Width
    ElseIf Y > SpellWindow.Height Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        Y = SpellWindow.Height
    End If
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > SpellWindow.Width Or Y > SpellWindow.Height Then _
        Exit Sub
        
    prevSelItem = SpellSelectedItem
    SpellSelectedItem = ClickItem(CInt(X), CInt(Y))
    
    If Button = 2 Then ' GSZAO - info del hechizo con segundo click
        Call WriteSpellInfo(SpellSelectedItem)
        ShowInfo = True
    End If
    
    

    If Button = vbRightButton Then
        If SpellSelectedItem <> slotDragged Then
            Call frmMain.picSpell_dragDone(slotDragged, TempItemSelect)
            frmMain.UsandoDrag = False
        Else
            frmMain.UsandoDrag = False
        End If
    End If

    Call DrawSpells
    
End Sub

Private Sub SpellWindow_Click()

    If UserHechizos(TempItemSelect).Name <> "(None)" And MainTimer.Check(TimersIndex.Work, False) And UsaMacro = False Then
        If ShowInfo = False Then
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                End With
            Else
                Call WriteCastSpell(SpellSelectedItem)
                Call WriteWork(eSkill.Magia)
                UsaMacro = True
            End If
        Else
            ShowInfo = False
        End If
        frmMain.lblItem.Caption = UserHechizos(SpellSelectedItem).Name
    End If
    
    Call DrawSpells
            
End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long

    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItemSelect = temp_x + (temp_y + SpellOffset) * (SpellWindow.ScaleWidth \ ownTilePixelWidth) + 1

    If (temp_y > 0) Then
        TempItemSelect = TempItemSelect + (temp_y) ' Corrección de filas
    End If

    If TempItemSelect <= UBound(UserHechizos) Then
    
        If UserHechizos(TempItemSelect).Name <> "(None)" Then
            ClickItem = TempItemSelect
        Else
            ClickItem = 0
            frmMain.lblItem.Caption = vbNullString
        End If
        
    End If
    
    Call DrawSpells
        
End Function

Private Function LowMana() As Boolean
   ' Function para hacer el if cuando tenga poca mana se ponga rojo jaja.
End Function

Public Property Get SelectedItem() As Long
'***************************************************
'Author: ^[GS]^
'Last Modify Date: 09/07/2012 - ^[GS]^
'***************************************************
    SelectedItem = SpellSelectedItem
End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Long
'***************************************************
'Author: ^[GS]^
'Last Modify Date: 09/07/2012 - ^[GS]^
'***************************************************
    GrhIndex = UserHechizos(slot).GrhIndex
End Property
