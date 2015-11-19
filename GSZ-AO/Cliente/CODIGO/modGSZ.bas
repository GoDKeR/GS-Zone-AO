Attribute VB_Name = "modGSZ"
Option Explicit

Public Enum eCursorState
    cur_Normal = 0
    cur_Action
    cur_Npc
    cur_Npc_Hostile
    cur_User
    cur_User_Danger
    cur_Obj
End Enum

Public Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type


Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&

Public Const WM_COPYDATA = &H4A
Public CurrentCursor As eCursorState

Type cComando
    command As String
    Desc As String
End Type

Dim arrComandos() As cComando

' función Api para aplicar la transparencia a la ventana
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal crKey As Long, _
     ByVal bAlpha As Byte, _
     ByVal dwFlags As Long) As Long
' Funciones api para los estilos de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long

Function ImgRequest(ByVal sFile As String) As String
    Dim r As Byte
    If LenB(Dir(sFile, vbArchive)) = 0 Then
        r = MsgBox("ERROR: Imagen no encontrada..." & vbCrLf & sFile, vbCritical + vbRetryCancel)
        If r = vbRetry Then
            sFile = ImgRequest(sFile)
        Else
            Call MsgBox("ADVERTENCIA: El juego seguirá funcionando sin alguna imagen!", vbInformation + vbOKOnly)
            sFile = DirGraficos & "blank.bmp"
        End If
    End If
    ImgRequest = sFile
End Function

Function FileRequest(ByVal sFile As String) As String
    If LenB(Dir(sFile, vbArchive)) = 0 Then
        Call MsgBox("ERROR: Archivo no encontrado..." & vbCrLf & sFile & vbCrLf & "¡Imposible continuar el juego!", vbCritical + vbOKOnly)
        DoEvents
        End
    End If
    FileRequest = sFile
End Function

Public Sub TirarDados()
    ' Tira los dados
    Call WriteThrowDices
    Call FlushBuffer
    DoEvents
End Sub

Public Sub Render_All_Lights()

    Light_Cuadrada.Light_Remove_All
    Light_Redonda.Light_Remove_All
    If CurMap = 1 Then
        'Light_Cuadrada.Light_Create 58, 46, &HDCDCDC00, 2, 1
        Light_Redonda.Light_Create 58, 46, 5, 1, 250, 250, 250
        Light_Redonda.Light_Create 45, 46, 7, 1, 250, 250, 250
    End If
    Light_Cuadrada.Light_Render_All
    Light_Redonda.Light_Render_All

End Sub

Public Sub LoadHelpCommands()
'***************************************************
'Author: ^[GS]^
'Last Modification: 13/03/2012
'***************************************************

'TODO: La .desc aun no esta implementada! y falta completar!

    ReDim arrComandos(47) As cComando
    
    ' Comandos de usuario
    arrComandos(0).command = "/comerciar"
    arrComandos(0).Desc = "Permite Comerciar con un NPC."
    arrComandos(1).command = "/online"
    arrComandos(1).Desc = "Nos muestra la cantidad de jugadores conectados en el momento."
    arrComandos(2).command = "/balance"
    arrComandos(2).Desc = "Nos muestra nuestro balance de deposito cuando es un utilizado con un NPC banquero."
    arrComandos(3).command = "/salir"
    arrComandos(3).Desc = "Sale del juego al menú principal."
    arrComandos(4).command = "/salirclan"
    arrComandos(5).command = "/quieto"
    arrComandos(6).command = "/acompañar"
    arrComandos(7).command = "/liberar"
    arrComandos(8).command = "/entrenar"
    arrComandos(9).command = "/descansar"
    arrComandos(10).command = "/meditar"
    arrComandos(10).Desc = "Comienzas a meditar para recuperar Mána, si tu raza lo permite."
    arrComandos(11).command = "/resucitar"
    arrComandos(11).Desc = "Te resucita cuando es utilizado con un NPC sacerdote."
    arrComandos(12).command = "/curar"
    arrComandos(12).Desc = "Te cura la vida cuando es utilizado con un NPC sacerdote."
    arrComandos(13).command = "/ayuda"
    arrComandos(14).command = "/est"
    arrComandos(14).Desc = "Muestra las estadísticas de tu personaje."
    arrComandos(15).command = "/boveda"
    arrComandos(15).Desc = "Abre tu boveda privada cuando es un utilizado con un NPC banquero."
    arrComandos(16).command = "/enlistar"
    arrComandos(17).command = "/informacion"
    arrComandos(18).command = "/recompensa"
    arrComandos(19).command = "/motd"
    arrComandos(20).command = "/uptime"
    arrComandos(21).command = "/salirparty"
    arrComandos(22).command = "/crearparty"
    arrComandos(23).command = "/party"
    arrComandos(24).command = "/encuesta"
    arrComandos(25).command = "/cmsg"
    arrComandos(26).command = "/pmsg"
    arrComandos(27).command = "/centinela"
    arrComandos(27).Desc = "Te permite responder al acertijo del Centinela."
    arrComandos(28).command = "/onlineclan"
    arrComandos(29).command = "/onlineparty"
    arrComandos(30).command = "/bmsg"
    arrComandos(31).command = "/rol"
    arrComandos(32).command = "/gm"
    arrComandos(32).Desc = "Hace un llamado a los GM's para que vengan a responder tu consulta."
    arrComandos(33).command = "/desc"
    arrComandos(34).command = "/voto"
    arrComandos(35).command = "/penas"
    arrComandos(36).command = "/contraseña"
    arrComandos(37).command = "/apostar"
    arrComandos(38).command = "/retirar"
    arrComandos(39).command = "/depositar"
    arrComandos(40).command = "/denunciar"
    arrComandos(40).Desc = "Te permite denunciar de urgencia un comportamiento indebido dentro del juego."
    arrComandos(41).command = "/fundarclan"
    arrComandos(42).command = "/echarparty"
    arrComandos(43).command = "/acceptparty"
    arrComandos(44).command = "/ping"
    arrComandos(44).Desc = "Muestra el tiempo de respuesta entre el servidor y tú."
    arrComandos(45).command = "/compartirnpc"
    arrComandos(46).command = "/nocompartirnpc"
    arrComandos(47).command = "/_bug"
    arrComandos(47).Desc = "Te permite reportar un Bug a los desarrolladores del servidor."

End Sub

Public Sub AutoComplete_(ByRef Textbox As Textbox)
'***************************************************
'Author: ^[GS]^
'Last Modification: 13/03/2012
'***************************************************
    Dim i As Integer
    Dim SelStart As Integer

    If (Len(Textbox.Text) >= 2 And Len(Textbox.Text) <= 20) Then ' solo de 2 a 20 caracteres
        For i = 0 To UBound(arrComandos)
            If InStr(1, arrComandos(i).command, Textbox.Text, vbTextCompare) = 1 Then
                SelStart = Textbox.SelStart
                Textbox.Text = arrComandos(i).command
                Textbox.SelStart = SelStart
                Textbox.SelLength = Len(Textbox.Text) - SelStart
                Exit For
            End If
        Next i
    End If
    
End Sub

Public Function ComprobarCaracteresLegales(ByVal Texto As String) As String
'***************************************************
'Author: ^[GS]^
'Last Modification: 13/03/2012
'***************************************************
    Dim tempstr As String

    If Len(Texto) > 160 Then
        tempstr = vbNullString
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim CharAscii As Integer
        For i = 1 To Len(Texto)
            CharAscii = Asc(mid$(Texto, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
    End If
    
    ComprobarCaracteresLegales = tempstr
    
End Function

' GSZAO - Encriptación basica y rapida para Strings
Public Function RndCrypt(ByVal str As String, ByVal Password As String) As String
    '  Made by Michael Ciurescu
    ' (CVMichael from vbforums.com)
    '  Original thread: http://www.vbforums.com/showthread.php?t=231798
    Dim SK As Long, K As Long

    Rnd -1
    Randomize Len(Password)

    For K = 1 To Len(Password)
        SK = SK + (((K Mod 256) _
        Xor Asc(mid$(Password, K, 1))) _
        Xor Fix(256 * Rnd))
    Next K

    Rnd -1
    Randomize SK
    
    For K = 1 To Len(str)
        Mid$(str, K, 1) = Chr(Fix(256 * Rnd) _
        Xor Asc(mid$(str, K, 1)))
    Next K
    
    RndCrypt = str
End Function

Public Sub ChangeCursorMain(ByVal eCursor As eCursorState)
    If CurrentCursor <> eCursor Then
        If eCursor = cur_Normal Then
            frmMain.MousePointer = vbDefault
        End If
        If ClientAOSetup.bCursores Then
            Call LoadCursor(eCursor, frmMain.MainViewPic.hwnd)
            Call LoadCursor(eCursor, frmMain.hwnd)
        Else
            If eCursor = cur_Action Then
                frmMain.MousePointer = 2
            End If
        End If
        CurrentCursor = eCursor
    End If
End Sub

Private Sub LoadCursor(ByVal sCursor As Byte, lHandle As Long)
    Dim GetCursor As Long
    If FileExist(DirCursores & sCursor & ".ani", vbArchive) = True Then
        GetCursor = LoadCursorFromFile(DirCursores & sCursor & ".ani")
        SetClassLong lHandle, -12, GetCursor
    End If
End Sub

Public Sub SetMusicInfo(ByRef r_sArtist As String, ByRef r_sAlbum As String, ByRef r_sTitle As String, Optional ByRef r_sWMContentID As String = vbNullString, Optional ByRef r_sFormat As String = "{0} - {1}", Optional ByRef r_bShow As Boolean = True)
On Error Resume Next
    Dim udtData As COPYDATASTRUCT
    Dim sBuffer As String
    Dim hMSGRUI As Long
     
    'Total length can Not be longer Then 256 characters!
    'Any longer will simply be ignored by Messenger.
    sBuffer = "\0Games\0" & Abs(r_bShow) & "\0" & r_sFormat & "\0" & r_sArtist & "\0" & r_sTitle & "\0" & r_sAlbum & "\0" & r_sWMContentID & "\0" & vbNullChar
     
    udtData.dwData = &H547
    udtData.lpData = StrPtr(sBuffer)
    udtData.cbData = LenB(sBuffer)
     
    Do
        hMSGRUI = FindWindowEx(0&, hMSGRUI, "MsnMsgrUIManager", vbNullString)
        If (hMSGRUI > 0) Then
            Call SendMessages(hMSGRUI, WM_COPYDATA, 0, VarPtr(udtData))
        End If
    Loop Until (hMSGRUI = 0)
 
End Sub

Public Function TransparenciaControl(ByVal hwnd As Long) As Boolean
On Local Error GoTo ErrSub
    
    Call SetWindowLong(hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    TransparenciaControl = True
    
    Exit Function
ErrSub:
    MsgBox Err.Description, vbCritical, "Error en TrasnaprenciaControl"
End Function

Public Function TransparenciaVentana(ByVal hwnd As Long, Alpha As Byte) As Boolean
On Local Error GoTo ErrSub
    Dim lS As Long
    
    lS = GetWindowLong(hwnd, GWL_EXSTYLE)
    lS = lS Or WS_EX_LAYERED
    Call SetWindowLong(hwnd, GWL_EXSTYLE, lS)
    Call SetLayeredWindowAttributes(hwnd, 0, Alpha, LWA_ALPHA)
    
    If Err Then
        TransparenciaVentana = False
    Else
        TransparenciaVentana = True
    End If
    
    Exit Function
ErrSub:
    MsgBox Err.Description, vbCritical, "Error en TransparenciaVentana"
End Function

Public Function UpdateUserPos()

    'Update pos label
    frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
    'frmMain.pUserMap.Left = UserPos.X - 2
    'frmMain.pUserMap.Top = UserPos.Y - 2
    DoEvents
    
End Function

Public Function TechoActivo(X As Long, Y As Long) As Boolean

    TechoActivo = False
    ' El usuario está bajo techo?
    If (MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2) Then
        Dim lmY As Long
        Dim liY As Long
        Dim lmX As Long
        Dim liX As Long
        Dim tY As Long
        Dim tX As Long
        
        tX = UserPos.X
        For tY = UserPos.Y To maxY
            If MapData(tX, tY).Graphic(4).GrhIndex = 0 Then
                lmY = tY - 1 ' limite superior
                Exit For
            End If
        Next
        For tY = UserPos.Y To minY Step -1
            If MapData(tX, tY).Graphic(4).GrhIndex = 0 Then
                liY = tY + 1 ' limite inferior
                Exit For
            End If
        Next
        
        tY = UserPos.Y
        For tX = UserPos.X To maxX
            If MapData(tX, tY).Graphic(4).GrhIndex = 0 Then
                lmX = tX - 1 ' limite derecho
                Exit For
            End If
        Next
        For tX = UserPos.X To minX Step -1
            If MapData(tX, tY).Graphic(4).GrhIndex = 0 Then
                liX = tX + 1 ' limite izquierdo
                Exit For
            End If
        Next
        
        If (X >= liX And X <= lmX And Y >= liY And Y <= lmY) Then
            TechoActivo = True
        End If
    
    End If
    
End Function

Public Sub LogError(ByVal sDesc As String)
'***************************************************
'Author: ^[GS]^
'Last Modification: 09/10/2012 - ^[GS]^
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\Errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & sDesc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

#If Testeo = 1 Then
    Public Sub LogTesteo(ByVal sDesc As String)
    '***************************************************
    'Author: ^[GS]^
    'Last Modification: 03/12/2012 - ^[GS]^
    '***************************************************
    
    On Error GoTo ErrHandler
    
        Dim nfile As Integer
        nfile = FreeFile ' obtenemos un canal
        Open App.Path & "\Testeo.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & sDesc
        Close #nfile
        
        Exit Sub
    
ErrHandler:
    
    End Sub
#End If
