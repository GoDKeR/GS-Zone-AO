Attribute VB_Name = "modGeneral"
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
Public usaParticulas As Boolean ' movelo donde lo creas correcto.

Public iplst As String
Public bFogata As Boolean
Public bLluvia() As Byte ' Array para determinar si
Public lFrameTimer As Long 'debemos mostrar la animacion de la lluvia

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = Fix((UpperBound - LowerBound) * Rnd) + LowerBound ' 0.13.5
End Function

Public Function GetRawName(ByRef sName As String) As String
'***************************************************
'Author: ZaMa
'Last Modify Date: 13/01/2010
'Last Modified By: -
'Returns the char name without the clan name (if it has it).
'***************************************************

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        GetRawName = Trim$(Left$(sName, Pos - 1))
    Else
        GetRawName = sName
    End If

End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim loopC As Long
    Dim ArcH As String
    
    ArcH = sPathINIT & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(ArcH, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopC = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopC).WeaponWalk(1), Val(GetVar(ArcH, "ARMA" & loopC, "Dir1")), 0
        InitGrh WeaponAnimData(loopC).WeaponWalk(2), Val(GetVar(ArcH, "ARMA" & loopC, "Dir2")), 0
        InitGrh WeaponAnimData(loopC).WeaponWalk(3), Val(GetVar(ArcH, "ARMA" & loopC, "Dir3")), 0
        InitGrh WeaponAnimData(loopC).WeaponWalk(4), Val(GetVar(ArcH, "ARMA" & loopC, "Dir4")), 0
    Next loopC
End Sub

Sub CargarColores()
On Error Resume Next
    Dim archivoC As String
    
    archivoC = sPathINIT & "colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego.", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 46 ' 47, 48, 49, 50 y 51 reservados
        ColoresPJ(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).g = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).b = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ' GSZAO Newbie
    ColoresPJ(47).r = CByte(GetVar(archivoC, "NW", "R"))
    ColoresPJ(47).g = CByte(GetVar(archivoC, "NW", "G"))
    ColoresPJ(47).b = CByte(GetVar(archivoC, "NW", "B"))
    
    ' Atacable
    ColoresPJ(48).r = CByte(GetVar(archivoC, "AT", "R"))
    ColoresPJ(48).g = CByte(GetVar(archivoC, "AT", "G"))
    ColoresPJ(48).b = CByte(GetVar(archivoC, "AT", "B"))
 
    ' Ciuda
    ColoresPJ(49).r = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).g = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).b = CByte(GetVar(archivoC, "CI", "B"))
    
    ' Crimi
    ColoresPJ(50).r = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).g = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).b = CByte(GetVar(archivoC, "CR", "B"))

    ' GSZAO Newbie
    ColoresPJ(51).r = CByte(GetVar(archivoC, "MT", "R"))
    ColoresPJ(51).g = CByte(GetVar(archivoC, "MT", "G"))
    ColoresPJ(51).b = CByte(GetVar(archivoC, "MT", "B"))


End Sub

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopC As Long
    Dim ArcH As String
    
    ArcH = sPathINIT & "Escudos.dat"
    
    NumEscudosAnims = Val(GetVar(ArcH, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopC = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopC).ShieldWalk(1), Val(GetVar(ArcH, "ESC" & loopC, "Dir1")), 0
        InitGrh ShieldAnimData(loopC).ShieldWalk(2), Val(GetVar(ArcH, "ESC" & loopC, "Dir2")), 0
        InitGrh ShieldAnimData(loopC).ShieldWalk(3), Val(GetVar(ArcH, "ESC" & loopC, "Dir3")), 0
        InitGrh ShieldAnimData(loopC).ShieldWalk(4), Val(GetVar(ArcH, "ESC" & loopC, "Dir4")), 0
    Next loopC
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************r
    With RichTextBox
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        
        RichTextBox.Refresh
    End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopC As Long
    
    For loopC = 1 To LastChar
        If CharList(loopC).active = 1 Then
            MapData(CharList(loopC).Pos.X, CharList(loopC).Pos.Y).CharIndex = loopC
        End If
    Next loopC
End Sub

Sub SaveConfigInit(Optional Modo As Byte = 0)

    If Modo = 1 Then
        'Grabamos los datos del usuario
        If (frmConnect.chkRecordar.Checked = True) Then
            If ClientConfigInit.Nombre <> frmConnect.txtNombre.Text Then ClientConfigInit.Nombre = frmConnect.txtNombre.Text
            If ClientConfigInit.Password <> frmConnect.txtPasswd.Text Then ClientConfigInit.Password = frmConnect.txtPasswd.Text
            ClientConfigInit.Recordar = 1
        ElseIf (frmConnect.chkRecordar.Checked = False) Then
            ClientConfigInit.Password = vbNullString
            ClientConfigInit.Recordar = 0
        End If
    End If

    Call EscribirConfigInit(ClientConfigInit)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        ' Asc("º") = 186
        If ((car < 97 Or car > 122) Or car = 186) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopC As Long
    Dim CharAscii As Integer
    
    If checkemail And LenB(UserEmail) = 0 Then
        MsgBox ("Dirección de email inválida.")
        Exit Function
    End If
    
    If LenB(UserPassword) = 0 Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopC = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caracter " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopC
    
    If LenB(UserName) = 0 Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopC = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caracter " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopC
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    Call SaveConfigInit(1)
    
    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmConnect
    
    frmMain.lblName.Caption = UserName
    'Load main form
    
    Call SetMusicInfo("Jugando " & NombreCliente & " [" & UserName & "] - [" & SitioOficial & "]", "Games", "{1}{0}") ' GSZAO
    
    frmMain.Visible = True
    
    Call frmMain.ControlSM(eSMType.mSpells, False)
    Call frmMain.ControlSM(eSMType.mWork, False)
    
    FPSFLAG = True

End Sub

'Sub CargarTip()
'    Dim N As Integer
'    N = RandomNumber(1, UBound(Tips))
'
'    frmtip.tip.Caption = Tips(N)
'End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk And Not UserParalizado Then
        'Meditación rápida - maTih 07/04/2012.
        If ClMeditarRapido Then
            If UserMeditar Then UserMeditar = False
        End If
        
        Call WriteWalk(Direccion)
        If Not UserDescansar And Not UserMeditar Then
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If CharList(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    If frmMain.MacroTrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call MoveTo(RandomNumber(NORTH, WEST))
End Sub

Private Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
    Static LastMovement As Long
    
    'No input allowed while Argentum is not the active window
    If Not modApplication.IsAppActive() Then Exit Sub
    
    'No walking when in Request FormYesNo
    If bFormYesNo Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'TODO: Debería informarle por consola?
    If Traveling Then Exit Sub

    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    If GetTickCount - LastMovement > 56 Then
        LastMovement = GetTickCount
    Else
        Exit Sub
    End If
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(NORTH)
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(EAST)
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(SOUTH)
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(WEST)
                Exit Sub
            End If
            
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
            End If
            
            If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
        End If
    End If
End Sub

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMap(ByVal Map As Integer)
'**************************************************************
' Formato de mapas optimizado para reducir el espacio que ocupan.
' Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
' Last Modify Date: 09/09/2012 - ^[GS]^
'**************************************************************
On Error GoTo errorH

    Dim Y As Long
    Dim X As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    
    If CfgDiaNoche = True Then Ambient_Start ' GSZAO
    
    ' Cargamos el mapa de Mapas.AO
    Dim MapReader As clsByteBuffer
    Set MapReader = New clsByteBuffer
    Dim data() As Byte
    
    If Get_File_Data(DirMapas, "MAPA" & CStr(Map) & ".MAP", data, 1) = False Then
        Call MsgBox("El Mapa " & CStr(Map) & " no existe.", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Call MapReader.initializeReader(data)
    
    mapInfo.MapVersion = MapReader.getInteger
    
    MiCabecera.Desc = MapReader.getString(Len(MiCabecera.Desc))
    MiCabecera.CRC = MapReader.getLong
    MiCabecera.MagicWord = MapReader.getLong
    
    Call MapReader.getDouble
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(X, Y)
                '.map file
                ByFlags = MapReader.getByte
                
                MapData(X, Y).particle_group = 0
                
                If ByFlags And 1 Then
                    .Blocked = 1
                Else
                    .Blocked = 0
                End If

                .Graphic(1).GrhIndex = MapReader.getInteger
                InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
                
                'Layer 2 used?
                If ByFlags And 2 Then
                    .Graphic(2).GrhIndex = MapReader.getInteger
                    InitGrh .Graphic(2), .Graphic(2).GrhIndex
                Else
                    .Graphic(2).GrhIndex = 0
                End If
                
                'Layer 3 used?
                If ByFlags And 4 Then
                    .Graphic(3).GrhIndex = MapReader.getInteger
                    InitGrh .Graphic(3), .Graphic(3).GrhIndex
                Else
                    .Graphic(3).GrhIndex = 0
                End If

                'Layer 4 used?
                If ByFlags And 8 Then
                    .Graphic(4).GrhIndex = MapReader.getInteger
                    InitGrh .Graphic(4), .Graphic(4).GrhIndex
                Else
                    .Graphic(4).GrhIndex = 0
                End If

                'Trigger used?
                If ByFlags And 16 Then
                    .Trigger = MapReader.getInteger
                Else
                    .Trigger = 0
                End If
                
                'Erase NPCs
                If .CharIndex > 0 Then
                    Call EraseChar(.CharIndex)
                End If
            
                'Erase OBJs
                .ObjGrh.GrhIndex = 0

                
            End With
        Next X
    Next Y
    
    mapInfo.Name = vbNullString
    mapInfo.Music = vbNullString
    
    CurMap = Map
    
    ' GSZAO - Cargamos el MiniMap
    Call MiniMap_ChangeTex(Map)
     
    ' loadLight's - Hay luces o no. Dunkan
    If CfgSistemaLuces = True Then
        Call Render_All_Lights
    End If
    
    Set MapReader = Nothing
        
    Exit Sub
    
errorH: ' GSZAO
    
    Set MapReader = Nothing
    Call LogError("SwitchMap::Error " & Err.Number & " - " & Err.Description & " - [MAPA: " & Map & "]")
    
    Call MsgBox("Error en el formato del Mapa " & Map, vbCritical + vbOKOnly, "Argentum Online")
    CloseClient ' Cerramos el cliente
    
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Public Function IsIp(ByVal Ip As String) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ServersLst)
        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next i
End Function

Public Sub CargarServidores()
'********************************
'Author: Unknown
'Last Modification: 10/06/2012 - ^[GS]^
'********************************
On Error GoTo errorH

    Dim F As String

    If ObtenerIP = 0 Then
        ' Cargamos los servidores de SInfo.dat
        
        Dim c As Integer
        Dim i As Long
        
        F = sPathINIT & "SInfo.dat"
        c = Val(GetVar(F, "INIT", "Cant"))
        
        ReDim ServersLst(1 To c) As tServerInfo
        For i = 1 To c
            ServersLst(i).Desc = GetVar(F, "S" & i, "Desc")
            ServersLst(i).Ip = Trim$(GetVar(F, "S" & i, "Ip"))
            ServersLst(i).Puerto = CInt(GetVar(F, "S" & i, "PJ"))
            ServersLst(i).PassRecPort = CInt(GetVar(F, "S" & i, "P2"))
        Next i
        
    ElseIf ObtenerIP = 1 Then
        ' Aquí configuramos donde se conecta nuestro cliente de manera "fija".
        
        ReDim ServersLst(1 To 1) As tServerInfo
        ServersLst(1).Desc = "Localhost" ' innecesario, no se usa
        ServersLst(1).Ip = "127.0.0.1" ' tú IP
        ServersLst(1).Puerto = 7666 ' tú puerto
        ServersLst(1).PassRecPort = 7667 ' puerto de recuperación de PJ o auxiliar
        
    ElseIf ObtenerIP = 2 Then
        ' Aquí tenemos que configurar la web de donde obtendra los datos de conexión
        ' Ejemplo:
        ' Localhost:127.0.0.1:7666:7667
        F = frmMain.iWeb.OpenURL("http://www.gs-zone.org/svr.txt") ' Aquí configura la URL de tus datos

        ServersLst(1).Desc = ReadField(1, F, ":")
        ServersLst(1).Ip = ReadField(2, F, ":")
        ServersLst(1).PassRecPort = CInt(ReadField(3, F, ":"))
        ServersLst(1).Puerto = CInt(ReadField(4, F, ":"))
        
    End If
    CurServer = 1
Exit Sub

errorH:
    Call MsgBox("Error Cargando los Servidores, actualicelos de la web.", vbCritical + vbOKOnly, "Argentum Online")
    Call CloseClient
End Sub

Public Sub InitServersList()
On Error Resume Next
    Dim NumServers As Integer
    Dim i As Integer
    Dim Cont As Integer
    
    i = 1
    
    ' Asc(";") = 59
    Do While (ReadField(i, RawServersList, 59) <> "")
        i = i + 1
        Cont = Cont + 1
    Loop
    
    ReDim ServersLst(1 To Cont) As tServerInfo
    
    For i = 1 To Cont
        Dim cur$
        cur$ = ReadField(i, RawServersList, 59)
        ServersLst(i).Ip = ReadField(1, cur$, 59)
        ServersLst(i).Puerto = ReadField(2, cur$, 59)
        ServersLst(i).Desc = ReadField(4, cur$, 59)
        ServersLst(i).PassRecPort = ReadField(3, cur$, 59)
    Next i
    
    CurServer = 1
End Sub

Public Function CurServerPasRecPort() As Integer
    If CurServer <> 0 Then
        CurServerPasRecPort = 7667
    Else
        CurServerPasRecPort = CInt(frmConnect.PortTxt)
    End If
End Function

Public Function CurServerIp() As String
    If CurServer <> 0 Then
        CurServerIp = "127.0.0.1"
    End If
End Function

Public Function CurServerPort() As Integer
    If CurServer <> 0 Then
        CurServerPort = 7666
    End If
End Function

Sub Main()
'********************************
'Author: Unknown
'Last Modification: 24/09/2012 - ^[GS]^
'**************************************

    'If Not StrComp(command$, "1") = 0 Then
    '    MsgBox "Debe ejecutar el launcher para abrir el cliente."
    '    End
    'End If
    
    'Constantes de Inits
    sPathINIT = App.Path & nDirINIT

    'Load config file
    If FileExist(sPathINIT & fConfigInit, vbNormal) Then
        ClientConfigInit = modGameIni.LeerConfigInit()
    Else
        Call MsgBox("Se requiere del archivo de configuración 'Config.Init' en la carpeta INIT del cliente.", vbCritical + vbOKOnly)
        Call CloseClient ' Cerramos el cliente
    End If
    
    NoRes = 1
    
    Call modGameIni.LoadClientAOSetup 'Load AOSetup.dat config file
    Call modGameIni.InitFilePaths 'Init Paths
    
    ' GSZAO Cambiar si los recursos .AO utilizan contraseña...
    ' NOTA: Con "" se deshabilita la utilización de contraseña!
    Call modCompression.GenerateContra("", 0) ' 0 = Graficos.AO
    Call modCompression.GenerateContra("", 1) ' 1 = Mapas.AO
    
    
    If ClientAOSetup.bDinamic Then
        Set SurfaceDB = New clsSurfaceManDyn
    Else
        Set SurfaceDB = New clsSurfaceManStatic
    End If
    
#If Testeo = 0 Then
    If FindPreviousInstance Then
        Call MsgBox("¡Argentum Online ya está siendo ejecutado!" & vbCrLf & "No es posible ejecutar otra instancia del juego." & vbCrLf & "Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If
#Else
    If FileExist(App.Path & "\Testeo.log", vbArchive) Then
        Call Kill(App.Path & "\Testeo.log")
    End If
#End If
    
    'Read command line. Do it AFTER config file is loaded to prevent this from
    'canceling the effects of "/nores" option.
    Call LeerLineaComandos
    
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.Path & "\")
    
    ChDrive App.Path
    ChDir App.Path

    MD5HushYo = "0123456789abcdef"  'We aren't using a real MD5
    
    tipf = ClientConfigInit.MostrarTips
    
    'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
    Call modResolution.SetResolution
    
    ' Load constants, classes, flags, graphics..
    Call LoadInitialConfig ' 0.13.3
    
    Call ChangeCursorMain(cur_Normal) ' GSZAO
    
    
    Call Fade_Initializate 'GSZAO
#If Testeo <> 1 Then
    Dim PresPath As String
    PresPath = DirGraficos & "Presentacion" & RandomNumber(1, 4) & ".jpg"
    
    frmPres.Picture = LoadPicture(PresPath)
    frmPres.Show vbModal    'Es modal, así que se detiene la ejecución de Main hasta que se desaparece
#End If

    frmMain.Socket1.Startup

    frmConnect.MousePointer = vbDefault
    frmConnect.Visible = True
    
    'Inicialización de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    
    ' Intervals
    LoadTimerIntervals ' 0.13.3
        
    'Set the dialog's font
    Dialogos.Font = frmMain.Font
    DialogosClanes.Font = frmMain.Font
    
    lFrameTimer = GetTickCount
    TechosTransp = 200 ' GSZAO
    
    ' Load the form for screenshots
    Call Load(frmScreenshots)

    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            'Play ambient sounds
            Call RenderSounds
            
            Call CheckKeys
        End If
        
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            If FPSFLAG Then frmMain.lblFPS.Caption = modTileEngine.FPS
            DrawText 0, 0, "FPS: " & modTileEngine.FPS, RGB(255, 255, 255)
            lFrameTimer = GetTickCount
        End If
        
        ' If there is anything to be sent, we send it
        Call FlushBuffer
        DoEvents
    Loop
   
    Call CloseClient

End Sub

Private Sub LoadInitialConfig() ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: ^[GS]^ - 22/08/2013
'***************************************************

    Dim i As Long

    frmCargando.Show
    frmCargando.Refresh
    frmCargando.cCargando.Min = 0
    frmCargando.cCargando.Max = 30

    frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    DoEvents
    
    '###########
    ' CONSTANTES
    
    Call AddtoRichTextBox(frmCargando.status, "Iniciando constantes... ", 255, 255, 255, True, False, True)
    Call InicializarNombres
    ' Initialize FONTTYPES
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1
    Call modProtocol.InitFonts
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 2
    
    With frmConnect
        .txtNombre = ClientConfigInit.Nombre
        .txtNombre.SelStart = 0
        .txtNombre.SelLength = Len(.txtNombre)
    End With
    UserMap = 1
    
    ' Mouse Pointer (Loaded before opening any form with buttons in it)
    If FileExist(DirCursores & "d.ico", vbArchive) Then _
        Set picMouseIcon = LoadPicture(DirCursores & "d.ico")
        
    ' Ayuda de Comandos
    Call LoadHelpCommands ' GSZAO
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 3
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    DoEvents

    Call AddtoRichTextBox(frmCargando.status, "Iniciando Clases... ", 255, 255, 255, True, False, True) ' 0.13.5
    
    Set Dialogos = New clsDialogs
    Set Audio = New clsAudio
    Set Inventario = New clsGraphicalInventory
    Set Spells = New clsGraphicalSpells
    Set CustomKeys = New clsCustomKeys
    Set CustomMessages = New clsCustomMessages
    Set incomingData = New clsByteQueue
    Set outgoingData = New clsByteQueue
    Set MainTimer = New clsTimer
    Set clsForos = New clsForum
    
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 4
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    DoEvents
    
    '##############
    ' MOTOR GRAFICO
    
    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectX... ", 255, 255, 255, True, False, True)
    
    DirectXInit ' DX8 engine
    
    If Not InitTileEngine(frmMain.hwnd, 149, 13, 32, 32, 13, 17, 7, 8, 8, 0.018) Then
        Call CloseClient
    End If
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 7
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    DoEvents

    '###################
    ' ANIMACIONES EXTRAS
    
    Call AddtoRichTextBox(frmCargando.status, "Creando animaciones extra... ", 255, 255, 255, True, False, True)
    
    'PARTICULAS TEST
    usaParticulas = True
    
    Call CargarArrayLluvia
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 8
    Call CargarAnimArmas
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 9
    Call CargarAnimEscudos
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 10
    Call CargarColores
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 11
    Call CargarTips
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 12
    
    LightRGB_Default(0) = D3DColorXRGB(255, 255, 255)
    LightRGB_Default(1) = LightRGB_Default(0)
    LightRGB_Default(2) = LightRGB_Default(0)
    LightRGB_Default(3) = LightRGB_Default(0)
    
    LightRGB_Default_Alpha(0) = D3DColorARGB(150, 255, 255, 255)
    LightRGB_Default_Alpha(1) = LightRGB_Default_Alpha(0)
    LightRGB_Default_Alpha(2) = LightRGB_Default_Alpha(0)
    LightRGB_Default_Alpha(3) = LightRGB_Default_Alpha(0)

    ' GSZAO Minimapa en Render
    modMinimap.DefaultAlphaMiniMap = 205   ' primero configuramos la Transparencia base del MiniMap
    modMinimap.MiniMap_Init ' inicializamos el MiniMap
        
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 13
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    DoEvents

    '#############
    ' DIRECT SOUND
    
    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectSound... ", 255, 255, 255, True, False, True)
    
    ' Inicializamos el sonido
    Call Audio.Initialize(DirectX, frmMain.hwnd, DirSound, DirMidi)
    ' Enable / Disable audio
    Audio.MusicActivated = Not ClientAOSetup.bNoMusic
    Audio.SoundActivated = Not ClientAOSetup.bNoSound
    Audio.SoundEffectsActivated = Not ClientAOSetup.bNoSoundEffects
    ' Volumen
    Audio.MusicVolume = ClientAOSetup.lMusicVolume
    Audio.SoundVolume = ClientAOSetup.lSoundVolume
    
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 14
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    DoEvents
    
    Call AddtoRichTextBox(frmCargando.status, "Iniciando Motor de Inventario... ", 255, 255, 255, True, False, True)
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(DirectD3D8, frmMain.PicInv, MAX_INVENTORY_SLOTS, , , , , , , , True)
    
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 15
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    DoEvents
    
    Call AddtoRichTextBox(frmCargando.status, "Iniciando Motor de Hechizos... ", 255, 255, 255, True, False, True)
    
    'Iniciamos los hechizos
    Call Spells.Initialize(DirectD3D8, frmMain.picSpell, MAX_SPELL_SLOTS)
    
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 16
        
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    DoEvents
    
    '######
    ' OTROS
    
    Call AddtoRichTextBox(frmCargando.status, "Iniciando Fuentes... ", 255, 255, 255, True, False, True)
    
    'Set the dialog's font
    Dialogos.Font = frmMain.Font
    DialogosClanes.Font = frmMain.Font
    
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 17
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    DoEvents
    
    'Musica de arranque
    'Call Audio.MusicMP3Play(App.Path & "\MP3\" & MP3_Inicio & ".mp3") ' GSZAO innecesario 99.9999999%
    
    frmCargando.cCargando.Value = frmCargando.cCargando.Value + 1 ' 18
    
    '######
    ' LISTO

    Call AddtoRichTextBox(frmCargando.status, "¡Bienvenido a " & NombreCliente & "!", 0, 255, 0, True, False, True)
    DoEvents
    
    'Give the user enough time to read the welcome text
    Call Sleep(1000)
    
    Unload frmCargando
  
End Sub

Private Sub LoadTimerIntervals() ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 24/09/2012 - ^[GS]^
'Set the intervals of timers
'***************************************************

#If Testeo = True Then
    Exit Sub ' GSZAO TEST (sacamos los intervalos para testear...)
#End If

    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK) ' <<< WTF?
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.SendPing, INT_SEND_PING) ' GSZAO
    
    frmMain.MacroTrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.MacroTrabajo.Enabled = False
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    Call MainTimer.Start(TimersIndex.SendPing) ' GSZAO

End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, Value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
                
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub

''
' Checks the command line parameters, if you are running Ao with /nores command and checks the AoUpdate parameters
'
'

Public Sub LeerLineaComandos()
'*************************************************
'Author: Unknown
'Last modified: 25/11/2008 (BrianPr)
'
'*************************************************
    Dim T() As String
    Dim i As Long
    
    Dim UpToDate As Boolean
    Dim Patch As String
    
    'Parseo los comandos
    T = Split(command, " ")
    For i = LBound(T) To UBound(T)
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
            Case "/UPTODATE"
                UpToDate = True
        End Select
    Next i
    
    'Call AoUpdate(UpToDate, NoRes) ' www.gs-zone.org
End Sub

''
' Runs AoUpdate if we haven't updated yet, patches aoupdate and runs Client normally if we are updated.
'
' @param UpToDate Specifies if we have checked for updates or not
' @param NoREs Specifies if we have to set nores arg when running the client once again (if the AoUpdate is executed).

Private Sub AoUpdate(ByVal UpToDate As Boolean, ByVal NoRes As Boolean)
'*************************************************
'Author: BrianPr
'Created: 25/11/2008
'Last modified: 11/06/2012 - ^[GS]^
'
'*************************************************
On Error GoTo error
    Dim extraArgs As String
    If Not UpToDate Then
        'No recibe update, ejecutar AU
        'Ejecuto el AoUpdate, sino me voy
        If Dir(App.Path & "\AoUpdate.exe", vbArchive) = vbNullString Then
            MsgBox "No se encuentra el archivo de actualización AoUpdate.exe por favor descarguelo y vuelva a intentar", vbCritical
            Call CloseClient ' Cerramos el cliente
        Else
            FileCopy App.Path & "\AoUpdate.exe", App.Path & "\AoUpdateTMP.exe"
            
            If NoRes Then
                extraArgs = " /nores"
            End If
            
            Call ShellExecute(0, "Open", App.Path & "\AoUpdateTMP.exe", App.EXEName & ".exe" & extraArgs, App.Path, SW_SHOWNORMAL)
            Call CloseClient ' Cerramos el cliente
        End If
    Else
        If FileExist(App.Path & "\AoUpdateTMP.exe", vbArchive) Then Kill App.Path & "\AoUpdateTMP.exe"
    End If
Exit Sub

error:
    If Err.Number = 75 Then 'Si el archivo AoUpdateTMP.exe está en uso, entonces esperamos 5 ms y volvemos a intentarlo hasta que nos deje.
        Sleep 5
        Resume
    Else
        MsgBox Err.Description & vbCrLf, vbInformation, "[ " & Err.Number & " ]" & " Error "
        Call CloseClient ' Cerramos el cliente
    End If
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/09/2012 - ^[GS]^
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************

    ' Configuración de las ciudades que se podrán elegir al crear un nuevo Personaje
    Ciudades(1) = "Keeg-Mai"
    Ciudades(2) = "Nork-Dor"
    Ciudades(3) = "Sait-Zeb"
    
    ' Razas
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    
    ' Clases
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clérigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladrón"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladín"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    ' Skills
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasión en combate"
    SkillsNames(eSkill.Armas) = "Combate cuerpo a cuerpo"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Minería"
    SkillsNames(eSkill.Carpinteria) = "Carpintería"
    SkillsNames(eSkill.Herreria) = "Herrería"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegación"

    ' Atributos
    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitución"
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call DialogosClanes.RemoveDialogs
    
    Call Dialogos.RemoveAllDialogs
End Sub

Public Function esGM(CharIndex As Integer) As Boolean
esGM = False
If CharList(CharIndex).priv >= 1 And CharList(CharIndex).priv <= 5 Or CharList(CharIndex).priv = 25 Then _
    esGM = True

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
Dim buf As Integer
buf = InStr(Nick, "<")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
buf = InStr(Nick, "[")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
getTagPosition = Len(Nick) + 2
End Function

Public Sub checkText(ByVal Text As String)
Dim Nivel As Integer
If Right$(Text, Len(MENSAJE_FRAGSHOOTER_TE_HA_MATADO)) = MENSAJE_FRAGSHOOTER_TE_HA_MATADO Then
    Call ScreenCapture(True)
    Exit Sub
End If
If Left$(Text, Len(MENSAJE_FRAGSHOOTER_HAS_MATADO)) = MENSAJE_FRAGSHOOTER_HAS_MATADO Then
    EsperandoLevel = True
    Exit Sub
End If
If EsperandoLevel Then
    If Right$(Text, Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA)) = MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA Then
        If CInt(mid$(Text, Len(MENSAJE_FRAGSHOOTER_HAS_GANADO), (Len(Text) - (Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA) + Len(MENSAJE_FRAGSHOOTER_HAS_GANADO))))) / 2 > ClientAOSetup.byMurderedLevel Then
            Call ScreenCapture(True)
        End If
    End If
End If
EsperandoLevel = False
End Sub

Public Function getStrenghtColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getStrenghtColor = RGB(255 - (M * UserFuerza), (M * UserFuerza), 0)
End Function
Public Function getDexterityColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getDexterityColor = RGB(255, M * UserAgilidad, 0)
End Function

Public Function getCharIndexByName(ByVal Name As String) As Integer
Dim i As Long
For i = 1 To LastChar
    If CharList(i).Nombre = Name Then
        getCharIndexByName = i
        Exit Function
    End If
Next i
End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Returns true if the post is sticky.
'***************************************************
    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
            
    End Select
    
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
'***************************************************
'Author: ZaMa
'Last Modification: 01/03/2010
'Returns the forum alignment.
'***************************************************
    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
            
    End Select
    
End Function

Public Function ColorToDX8(ByVal long_color As Long) As Long
    ' DX8 engine
    Dim temp_color As String
    Dim Red As Integer, Blue As Integer, Green As Integer
    
    temp_color = Hex$(long_color)
    If Len(temp_color) < 6 Then
        'Give is 6 digits for easy RGB conversion.
        temp_color = String(6 - Len(temp_color), "0") + temp_color
    End If
    
    Red = CLng("&H" + mid$(temp_color, 1, 2))
    Green = CLng("&H" + mid$(temp_color, 3, 2))
    Blue = CLng("&H" + mid$(temp_color, 5, 2))
    
    ColorToDX8 = D3DColorXRGB(Red, Green, Blue)

End Function


Public Sub CloseClient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 02/08/2012 - ^[GS]^
'Frees all used resources, cleans up and leaves
'**************************************************************
    ' Allow new instances of the client to be opened
    Call modPrevInstance.ReleaseInstance
    
    EngineRun = False
    
    frmCargando.Show
    
    Call AddtoRichTextBox(frmCargando.status, "Liberando recursos...", 255, 255, 255, 0, True, 0)
    
    DoEvents
    
    Call modResolution.ResetResolution
    
    'Stop tile engine
    Call DeinitTileEngine
    
    Call modGameIni.SaveClientAOSetup
    
    'Destruimos los objetos públicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    'Clear arrays
    Erase GrhData
    Erase BodyData
    Erase HeadData
    Erase FxData
    Erase WeaponAnimData
    Erase ShieldAnimData
    Erase CascoAnimData
    Erase MapData
    Erase CharList
        
    Call UnloadAllForms
    
    'Actualizar tip
    ClientConfigInit.MostrarTips = tipf
    Call EscribirConfigInit(ClientConfigInit)
    
    End ' THE END
    
End Sub

Public Sub ResetAllInfo() ' 0.13.3
'Last Modification: 10/08/2013 - ^[GS]^
'*******************************************************

    ' Save config.ini
    Call SaveConfigInit(0)
    
    ' Disable timers
    frmMain.Second.Enabled = False
    frmMain.MacroTrabajo.Enabled = False
    Connected = False
    
    'Unload all forms except frmMain, frmConnect and frmCrearPersonaje
    Dim frm As Form
    For Each frm In Forms
        If frm.Name <> frmMain.Name And frm.Name <> frmConnect.Name And _
            frm.Name <> frmCrearPersonaje.Name Then
            
            Unload frm
        End If
    Next
    
    On Local Error GoTo 0
    
    ' Return to connection screen
    frmConnect.MousePointer = vbNormal
    If Not frmCrearPersonaje.Visible Then frmConnect.Visible = True
    frmMain.Visible = False
    
    'Stop audio
    Call Audio.StopWave
    frmMain.IsPlaying = PlayLoop.plNone

    ' Reset flags
    pausa = False
    UserMeditar = False
    UserEstupido = False
    UserCiego = False
    UserDescansar = False
    UserParalizado = False
    Traveling = False
    UserNavegando = False
    bRain = False
    bRangoReducido = False
    bFogata = False
    Comerciando = False
    bFormYesNo = False ' GSZAO
    bShowTutorial = False
    
    MirandoAsignarSkills = False
    MirandoCarpinteria = False
    MirandoEstadisticas = False
    MirandoForo = False
    MirandoHerreria = False
    MirandoParty = False
    
    'Delete all kind of dialogs
    Call CleanDialogs

    'Reset some char variables...
    Dim i As Long
    For i = 1 To LastChar
        CharList(i).Invisible = False
    Next i

    ' Reset stats
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = vbNullString
    SkillPoints = 0
    Alocados = 0
    
    ' Reset skills
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    ' Reset attributes
    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    ' Clear inventory slots
    Inventario.ClearAllSlots

    ' Connection screen midi
    Call Audio.PlayMIDI("2.mid")

End Sub

Public Function stringSinTildes(ByRef str As String) As String ' 0.13.5
    Dim temp As String
    
    temp = str
    
    If InStr(1, str, "á") > 0 Then temp = Replace(temp, "á", "a")
    If InStr(1, str, "Á") > 0 Then temp = Replace(temp, "Á", "A")
    
    If InStr(1, str, "é") > 0 Then temp = Replace(temp, "é", "e")
    If InStr(1, str, "É") > 0 Then temp = Replace(temp, "É", "E")
    
    If InStr(1, str, "í") > 0 Then temp = Replace(temp, "í", "i")
    If InStr(1, str, "Í") > 0 Then temp = Replace(temp, "Í", "I")
    
    If InStr(1, str, "ó") > 0 Then temp = Replace(temp, "ó", "o")
    If InStr(1, str, "Ó") > 0 Then temp = Replace(temp, "Ó", "O")
    
    If InStr(1, str, "ú") > 0 Then temp = Replace(temp, "ú", "u")
    If InStr(1, str, "Ú") > 0 Then temp = Replace(temp, "Ú", "U")
    
    stringSinTildes = temp
End Function

Public Function ResizePicture(pBox As PictureBox, pPic As Picture) As Boolean
Dim lWidth      As Single, lHeight    As Single
Dim lNewWidth   As Single, lNewHeight As Single

On Error GoTo Err
    
    ResizePicture = False
    
    'Clear the Picture in the PictureBox
    pBox.Picture = Nothing
    
    'Clear the Image  in the Picturebox
    pBox.Cls
    
    'Get the size of the Image, but in the same Scale than the scale used by the PictureBox
    lWidth = pBox.ScaleX(pPic.Width, vbHimetric, pBox.ScaleMode)
    lHeight = pBox.ScaleY(pPic.Height, vbHimetric, pBox.ScaleMode)
    
    'If image Width > pictureBox Width, resize Width
    If lWidth > pBox.ScaleWidth Then
        lNewWidth = pBox.ScaleWidth              'new Width = PB width
        lHeight = lHeight * (lNewWidth / lWidth) 'Risize Height keeping proportions
    Else
        lNewWidth = lWidth                       'If not, keep the original Width value
    End If
    
    'If the image Height > The pictureBox Height, resize Height
    If lHeight > pBox.ScaleHeight Then
        lNewHeight = pBox.ScaleHeight                   'new Height = PB Height
        lNewWidth = lNewWidth * (lNewHeight / lHeight)  'Risize Width keeping proportions
    Else
        lNewHeight = lHeight                            'If not, use the same value
    End If
    
    'add resized and centered to Picturebox
    pBox.PaintPicture pPic, (pBox.ScaleWidth - lNewWidth) / 2, _
                            (pBox.ScaleHeight - lNewHeight) / 2, _
                            lNewWidth, lNewHeight
                            
    'Update the Picture with the new image if you need it
    Set pBox.Picture = pBox.Image
    
    ResizePicture = True
    
    Exit Function
    
Err:
ResizePicture = False

End Function


