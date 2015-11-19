Attribute VB_Name = "modGeneral"
Option Explicit
' Otros
Public Declare Function GetTickCount Lib "kernel32" () As Long

' Tipos generales
Private Type tFileInfo
    Filename As String
    FilePath As String
    FileLen As Double
    FileMD5 As String
End Type

' Constantes generales
Private Const S_EXE = ".exe"

' Variables generales
Public Const OFICIAL_URL As String = "http://www.gs-zone.org/"
'Public Const UPDATE_URL As String = "http://localhost/GSZAO/list_update.php?download="
Public Const UPDATE_URL As String = "http://localhost/GSZAO/update/"
Public Const UPDATER_URL As String = "http://localhost/GSZAO/updater/"
'Public Const UPDATER_URL As String = "http://gsz-ao.sourceforge.net/autoupdate/updater/"
Public Const LIST_CHECK_URL As String = "http://localhost/GSZAO/list_update.php"
Public Const UPDATER_CHECK_URL As String = "http://localhost/GSZAO/list_update.php?updater"

Private LocalPath As String
Private AutoInfo As tFileInfo
Private FileInfo As tFileInfo

Private Sub ExtractInfo(ByVal sData As String)
    FileInfo.Filename = ReadField(1, sData, 124)
    If (InStrRev(FileInfo.Filename, "\") <> 0) Then
        FileInfo.FilePath = Left$(FileInfo.Filename, InStrRev(FileInfo.Filename, "\"))
        FileInfo.Filename = Right$(FileInfo.Filename, Len(FileInfo.Filename) - Len(FileInfo.FilePath))
    Else
        FileInfo.FilePath = "\"
    End If
    FileInfo.FileLen = CDbl(ReadField(2, sData, 124))
    FileInfo.FileMD5 = Left$(ReadField(3, sData, 124), Len(ReadField(3, sData, 124)))
End Sub

Private Sub InitTotal(ByVal iTotal As Integer)
    With frmPrincipal
        .sTF.Min = 0
        .sTF.Value = 0
        .sTF.Max = iTotal
    End With
    DoEvents
End Sub

Private Sub NextTotal(ByVal sTitulo As String)
    With frmPrincipal
        .tTF.Caption = sTitulo
        If .sTF.Value < .sTF.Max Then
            .sTF.Value = .sTF.Value + 1
        End If
    End With
    DoEvents
End Sub

Public Sub Main()
'On Error Resume Next
    Dim sTemp As String
    Dim sTemp2 As String
    Dim NeedRename As Boolean
    Dim NeedDownload As Boolean
    
    LocalPath = App.Path & "\"
    
    GoTo CheckFiles
    
    sTemp = App.EXEName ' Nombre sin extensión del ejecutable
    If Right$(sTemp, 1) = "_" Then ' ¿Es una actualización?
        sTemp2 = Left$(sTemp, Len(sTemp) - 1) & S_EXE ' Nombre del ejecutable original
        Call KillProcess(sTemp2)  ' Matamos el proceso, si llega a estar abierto...
        If FileExists(LocalPath & sTemp2) Then
            Call Kill(LocalPath & sTemp2) ' Si existe, lo eliminamos...
        End If
        If Not FileExists(LocalPath & sTemp2) Then ' ¿Ya no existe?
            Call FileCopy(LocalPath & App.EXEName & S_EXE, sTemp2)
            Call Shell(LocalPath & sTemp2)
            End
        Else
            Call MsgBox("Ha ocurrido un error en la aplicación de la actualización." & vbCrLf & "Elimine manualmente " & Left$(App.EXEName, Len(App.EXEName) - 1) & S_EXE & " y ejecute nuevamente " & App.EXEName & S_EXE, vbOKOnly + vbCritical)
            End
        End If
    ElseIf FileExists(LocalPath & sTemp & "_" & S_EXE) Then ' ¿Hay un temporal de actualización remanente?
        Call KillProcess(sTemp & "_" & S_EXE)  ' Matamos el proceso, si llega a estar abierto...
        Call Kill(LocalPath & sTemp & "_" & S_EXE) ' Si existe, lo eliminamos...
    End If
    
    Call InitTotal(3)
    frmPrincipal.tSF.Caption = vbNullString
    
    frmPrincipal.Show
    frmPrincipal.WindowState = 0
    frmPrincipal.Left = (Screen.Width - frmPrincipal.Width) / 2
    frmPrincipal.Top = (Screen.Height - frmPrincipal.Height) / 2
    
    ' Obtenemos datos del auto-updater local
    AutoInfo.Filename = App.EXEName
    AutoInfo.FilePath = LocalPath
    AutoInfo.FileLen = FileLen(AutoInfo.FilePath & AutoInfo.Filename & S_EXE)
    AutoInfo.FileMD5 = MD5File(AutoInfo.FilePath & AutoInfo.Filename & S_EXE)
    
    Call NextTotal("Comprobando programa de actualizaciones desde internet...")
    
    ' Obtenemos datos del auto-updater en internet...
    sTemp = frmPrincipal.Inet.OpenURL(UPDATER_CHECK_URL & "&" & strAntiProxy)
    DoEvents
    
    If (InStr(1, sTemp, "|") = 0) Then
        Call MsgBox("Ha ocurrido un error en la comprobación de actualización." & vbCrLf & "Descargue nuevamente el archivo de actualización de la web " & OFICIAL_URL, vbCritical + vbOKOnly)
        End
    Else
        NeedRename = False
        NeedDownload = False
        
        Call ExtractInfo(sTemp) ' Extraemos la información recibida...

        ' Info valida?
        If LCase(FileInfo.Filename & S_EXE) <> LCase(AutoInfo.Filename & S_EXE) Then
            ' Tiene otro nombre?
            NeedRename = True
        End If
        
        If FileInfo.FileLen <> AutoInfo.FileLen Then
            ' Tiene otro tamaño?
            NeedDownload = True
        ElseIf FileInfo.FileMD5 <> AutoInfo.FileMD5 Then
            ' Tiene otro MD5?
            NeedDownload = True
        End If
        
        If NeedRename = True Then
            Call NextTotal("Verificando programa de actualización...")
            If FileExists(LocalPath & FileInfo.Filename & S_EXE) Then ' Ya existe?
                If FileLen(LocalPath & FileInfo.Filename & S_EXE) = FileInfo.FileLen Then ' Mismo tamaño!
                    If MD5File(LocalPath & FileInfo.Filename & S_EXE) = FileInfo.FileMD5 Then ' Mismo MD5!
                        Call Shell(LocalPath & FileInfo.Filename & S_EXE) ' Ejecitamos el autoupdate correcto...
                        End
                    End If
                End If
                Call Kill(LocalPath & FileInfo.Filename & S_EXE) ' Borrar el archivo que ocupa su lugar...
            End If
            Call FileCopy(LocalPath & AutoInfo.Filename & S_EXE, LocalPath & FileInfo.Filename & S_EXE)
            Call Shell(LocalPath & FileInfo.Filename & S_EXE) ' Ejecitamos el autoupdate correcto...
            End
        End If
        
        If NeedDownload = True Then
            Call NextTotal("Descargando nuevo programa de actualización...")
            Call DownloadFile(UPDATER_URL & FileInfo.Filename & S_EXE & "?" & strAntiProxy, LocalPath & FileInfo.Filename & "_" & S_EXE)
            Call NextTotal("Verificando actualización...")
            If FileExists(LocalPath & FileInfo.Filename & "_" & S_EXE) Then  ' Ya existe?
                If FileLen(LocalPath & FileInfo.Filename & "_" & S_EXE) = FileInfo.FileLen Then  ' Mismo tamaño!
                    If MD5File(LocalPath & FileInfo.Filename & "_" & S_EXE) = FileInfo.FileMD5 Then  ' Mismo MD5!
                        Call Shell(LocalPath & FileInfo.Filename & "_" & S_EXE)  ' Ejecitamos el autoupdate correcto...
                        End
                    End If
                End If
                Call Kill(LocalPath & FileInfo.Filename & "_" & S_EXE)  ' Borrar el archivo fallado...
            End If
            Call MsgBox("Ha ocurrido un error en la descarga de la actualización." & vbCrLf & "Descargue nuevamente el archivo de actualización de la web " & OFICIAL_URL, vbCritical + vbOKOnly)
            End
        End If
        ' Auto-updater ACTUALIZADO!
    End If
    Call NextTotal("Programa de actualización verificado...")
    DoEvents
    
CheckFiles:
    
    frmPrincipal.Show
    frmPrincipal.WindowState = 0
    frmPrincipal.Left = (Screen.Width - frmPrincipal.Width) / 2
    frmPrincipal.Top = (Screen.Height - frmPrincipal.Height) / 2
    
    Call NextTotal("Descargando listado de archivos desde internet...")
    DoEvents
    sTemp = frmPrincipal.Inet.OpenURL(LIST_CHECK_URL & "?" & strAntiProxy)
    DoEvents
    If (InStr(1, sTemp, "|") = 0) Then
        Call MsgBox("Ha ocurrido un error en la comprobación de actualización." & vbCrLf & "Descargue nuevamente el archivo de actualización de la web " & OFICIAL_URL, vbCritical + vbOKOnly)
        End
    Else
        Dim lI As Long
        Dim sFiles() As String
        Dim ActualAOVersion As Long
        Dim NeedPatch As Boolean
        
        sFiles = Split(sTemp, Chr(10))
        Call InitTotal(UBound(sFiles) + 1)
        
        For lI = 0 To UBound(sFiles) - 1
        
            ActualAOVersion = 0
            NeedPatch = False
            NeedDownload = False
            
            Call ExtractInfo(sFiles(lI)) ' Extraemos la información recibida...
            Call NextTotal("Verificando " & FileInfo.FilePath & FileInfo.Filename & "...")
            
            ' Ya existe el archivo?
            If FileExists(LocalPath & FileInfo.FilePath & FileInfo.Filename) Then
                MsgBox "Existe " & FileInfo.Filename
                NeedDownload = True ' Todo archivo es "culpable" hasta que se demuestre lo contrario (Así funciona el primer mundo)
                If FileLen(LocalPath & FileInfo.FilePath & FileInfo.Filename) = FileInfo.FileLen Then
                    If MD5File(LocalPath & FileInfo.FilePath & FileInfo.Filename) = FileInfo.FileMD5 Then
                        NeedDownload = False
                    End If
                End If
                If NeedDownload = True Then
                    MsgBox "Igual lo vamos a descargar " & FileInfo.Filename
                    If LCase$(Right(FileInfo.Filename, 3)) = ".ao" Then ' Es un .AO?
                        MsgBox "entro!"
                        ' ¿Que versión DICE que tiene actualmente?
                        ActualAOVersion = modCompresion.GetVersion(LocalPath & FileInfo.FilePath & FileInfo.Filename)
                        MsgBox "ActualAOVersion=" & ActualAOVersion
                    ElseIf LCase$(Left$(FileInfo.Filename, 5)) = ".init" Then ' Es un .init?
                        NeedDownload = False ' No descargamos los archivos de configuración si ya existen (el cliente debe eliminarlo para forzar la descarga)
                    End If
                    MsgBox "ble ble " & LCase$(Right(FileInfo.Filename, 3))
                End If
            Else ' Si no existe el archivo, requiere que lo descarguemos...
                NeedDownload = True
            End If
            
            If NeedDownload = True Then
                Call NextTotal("Actualizando " & FileInfo.FilePath & FileInfo.Filename & "...")
                ' Necesitamos actualizar este archivo...
                If Not DirExists(LocalPath & FileInfo.FilePath) Then
                    ' Si no existe el path, lo creamos...
                    Call CreatePath(LocalPath & FileInfo.FilePath)
                End If
                Call DownloadFile(UPDATE_URL & FileInfo.FilePath & FileInfo.Filename & "?" & strAntiProxy, LocalPath & FileInfo.FilePath & FileInfo.Filename)
                Call NextTotal("Verificando " & FileInfo.FilePath & FileInfo.Filename & "...")
                If FileExists(LocalPath & FileInfo.FilePath & FileInfo.Filename) Then
                    If FileLen(LocalPath & FileInfo.FilePath & FileInfo.Filename) <> FileInfo.FileLen Then
                        Call MsgBox("Ha ocurrido un error en la verificación de " & FileInfo.Filename & vbCrLf & "El tamaño del archivo es incorrecto.", vbCritical + vbOKOnly)
                        End
                    End If
                    If MD5File(LocalPath & FileInfo.FilePath & FileInfo.Filename) <> FileInfo.FileMD5 Then
                        MsgBox MD5File(LocalPath & FileInfo.FilePath & FileInfo.Filename)
                        Call MsgBox("Ha ocurrido un error en la verificación de " & FileInfo.Filename & vbCrLf & "El archivo se encuentra corrupto, verifique que su sistema no tenga virus.", vbCritical + vbOKOnly)
                        End
                    End If
                Else
                    Call MsgBox("Ha ocurrido un error en la verificación de " & FileInfo.Filename & vbCrLf & "Verifique su conexión a internet y compruebe que se encuentre operativa.", vbCritical + vbOKOnly)
                    End
                End If
                DoEvents
            Else
                Call NextTotal(FileInfo.FilePath & FileInfo.Filename & " se encuentra actualizado...")
            End If
            
            MsgBox "Información:" & vbCrLf & vbCrLf & _
                    "Archivo -> " & FileInfo.Filename & vbCrLf & _
                    "Carpeta -> " & FileInfo.FilePath & vbCrLf & _
                    "Len -> " & FileInfo.FileLen & vbCrLf & _
                    "MD5 -> " & FileInfo.FileMD5
        Next
        'MsgBox sTemp
    End If
    Call NextTotal("Ya tienes la versión actualizada...")
    frmPrincipal.Inet.Cancel
    frmPrincipal.Show
    
End Sub

Public Sub DownloadFile(strURL As String, strDestination As String)
    Dim intFile As Integer, lngBytesReceived As Double, lngFileLength As Double
    Dim b() As Byte, i As Integer
    Const CHUNK_SIZE As Long = 1024
    lngBytesReceived = 0
    DoEvents
    
    ' Fix URL
    strURL = Replace(strURL, "\", "/")
    'strURL = Replace(strURL, "//", "/")
    
    With frmPrincipal.Inet
        .URL = strURL ' Indicamos la URL del archivo
        .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
        
        While .Tag < icError Or .StillExecuting ' Esperamos a que termine el proceso de envio o que devuelva error
            DoEvents
        Wend
        
        If .Tag = icError Then ' Error
            MsgBox .RemoteHost & ":" & .RemotePort & vbCrLf & "Error " & .ResponseCode & ": " & .ResponseInfo
        ElseIf .Tag = icResponseCompleted Then ' Operación completada
            lngFileLength = Val(.GetHeader("Content-Length")) ' Obtenemos el tamaño del archivo
            If lngFileLength > 0 Then
                lngBytesReceived = 0
                intFile = FreeFile()
                ' Anunciamos la descarga del archivo
                frmPrincipal.tSF.Caption = "Descargando " & ReturnFileOrFolder(strURL, True, True) & "..."
                frmPrincipal.sSF.Value = 0
                frmPrincipal.sSF.Max = lngFileLength
                ' Lo hacemos visible...
                frmPrincipal.tSF.Visible = True
                frmPrincipal.sSF.Visible = True
                If FileExists(strDestination) Then  ' Ya existe?
                    Call Kill(strDestination)  ' Borrar el archivo...
                    If FileExists(strDestination) Then  ' Sigue existiendo?!
                        Call MsgBox("Ha ocurrido un error en la descarga de la actualización." & vbCrLf & "Cierre todas las aplicaciones que se puedan" & vbCrLf & "relacionar con el archivo (" & strDestination & ") y vuelve a intentarlo.", vbCritical + vbOKOnly)
                        Exit Sub
                    End If
                End If
                ' Comenzamos a volcar el contenido del archivo en el destino indicado
                Open strDestination For Binary Access Write As #intFile
                Do
                    b = .GetChunk(CHUNK_SIZE, icByteArray)
                    Put #intFile, , b
                    lngBytesReceived = lngBytesReceived + UBound(b, 1) + 1
                    frmPrincipal.sSF.Value = lngBytesReceived ' Actualizamos el progreso de descarga...
                    DoEvents
                Loop While UBound(b, 1) > 0
                ' Cerramos el archivo
                Close #intFile
                ' Ocultamos los controles
                frmPrincipal.tSF.Visible = False
                frmPrincipal.sSF.Visible = False
                ' Nos aseguramos de cerrar la conexión...
                .Execute , "CLOSE"
                DoEvents
            Else
                ' Error en el tamaño del archivo ha descargar
                Call MsgBox("Ha ocurrido un error en la descarga de la actualización." & vbCrLf & "Descargue nuevamente el archivo de actualización de la web " & OFICIAL_URL, vbCritical + vbOKOnly)
                Exit Sub
            End If
        End If
    End With
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 16/03/2012 - ^[GS]^
'Gets a field from a delimited string
'*****************************************************************
' Listado de SepASCII más comunes...
'   = 32
' , = 44
' - = 45
' . = 46
' | = 124

    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If lastPos = 0 And Pos <> 1 Then ' GSZAO, fix
        ReadField = vbNullString
        Exit Function
    End If
    
    If CurrentPos = 0 Then
        ReadField = Mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = Mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
    
End Function

Public Function FileExists(ByRef sFile As String) As Boolean
    FileExists = (LenB(Dir$(sFile)) > 0)
End Function

Private Sub KillProcess(ByVal Filename As String)
On Error GoTo Fallo
    Dim Process As Object
    For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '" & Filename & "'")
        Process.Terminate
    Next
    DoEvents
Fallo:
End Sub

Private Function strAntiProxy() As String
On Error Resume Next
    Call Randomize
    strAntiProxy = "random=" & (Int(Rnd * 30000)) & "&no-cache=" & CLng(GetTickCount)
End Function

Public Function GetHostURL(ByVal FullURL As String) As String
    Dim iT As Integer
    
    iT = InStr(1, FullURL, "://")
    If iT > 0 Then
        GetHostURL = Right$(FullURL, (Len(FullURL) - iT) - 2)
        iT = InStr(1, GetHostURL, "/")
        If iT > 0 Then
            GetHostURL = Left$(GetHostURL, iT - 1)
        End If
    Else
        GetHostURL = vbNullString
    End If
End Function


Public Function ReturnFileOrFolder(ByVal FullPath As String, _
                                   ByVal ReturnFile As Boolean, _
                                   Optional ByVal IsURL As Boolean = False) _
                                   As String
'*************************************************
'Author: Jeff Cockayne
'Last modified: ?/?/?
'*************************************************

' ReturnFileOrFolder:   Returns the filename or path of an
'                       MS-DOS file or URL.
'
' Author:   Jeff Cockayne 4.30.99
'
' Inputs:   FullPath:   String; the full path
'           ReturnFile: Boolean; return filename or path?
'                       (True=filename, False=path)
'           IsURL:      Boolean; Pass True if path is a URL.
'
' Returns:  String:     the filename or path
'
    Dim intDelimiterIndex As Integer
    
    If (IsURL = True) Then
        ' Eliminamos los parametros "extras"
        If InStr(1, FullPath, "?") > 0 Then
            FullPath = Left$(FullPath, InStr(1, FullPath, "?") - 1)
        ElseIf InStr(1, FullPath, "&") > 0 Then
            FullPath = Left$(FullPath, InStr(1, FullPath, "&") - 1)
        End If
    End If
    
    intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
    ReturnFileOrFolder = IIf(ReturnFile, _
                             Right$(FullPath, Len(FullPath) - intDelimiterIndex), _
                             Left$(FullPath, intDelimiterIndex))
End Function

Public Function DirExists(ByRef sDir As String) As Boolean
    DirExists = (LenB(Dir$(sDir, vbDirectory)) > 0)
End Function

Private Function CreatePath(ByVal NewPath As String) As Boolean
On Error GoTo Fallo
    CreatePath = False
    Dim sTemp As String
    If DirExists(NewPath) Then
        CreatePath = True
        Exit Function
    End If
    Do
        sTemp = Left$(NewPath, InStr(Len(sTemp) + 1, NewPath, "\"))
        If Not DirExists(sTemp) Then
            Call MkDir(sTemp)
        ElseIf sTemp = NewPath Then
            Exit Do
        End If
    Loop
    If DirExists(NewPath) Then
        CreatePath = True
    End If
Fallo:
End Function
