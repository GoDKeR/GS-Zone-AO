VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DownloadForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Descargando archivo"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "DownloadForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4320
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Download File"
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.Label StatusLabel 
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label StatusLabel2 
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   4695
      End
   End
End
Attribute VB_Name = "DownloadForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************


' Note: there are a number of useful functions here
' that should go into a public module, but have all
' been lumped into the form for the purpose of this
' example!
'
' See:
' FormatFileSize(), FormatTime(), ReturnFileOrFolder()

Option Explicit

Private CancelSearch As Boolean
Public DownloadSuccess As Boolean
Public BotonCancel As Boolean

Public Function FormatFileSize(ByVal dblFileSize As Double) As String
'*************************************************
'Author: Unkwown
'Last modified: ?/?/?
'*************************************************

' FormatFileSize:   Formats dblFileSize in bytes into
'                   X GB or X MB or X KB or X bytes depending
'                   on size (a la Win9x Properties tab)

Select Case dblFileSize
    Case 0 To 999   ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"
    Case 1000 To 1023999    ' KB
        FormatFileSize = Format(dblFileSize / 1024, "##0.0") & " KB"
    Case 1024000 To (1024 * 10 ^ 6) - 1 ' MB
        FormatFileSize = Format(dblFileSize / (1024 ^ 2), "##0.0#") & " MB"
    Case Is > (1024 * 10 ^ 6)
        FormatFileSize = Format(dblFileSize / (1024 ^ 3), "##0.0#") & " GB"
End Select

End Function

Public Function FormatTime(ByVal sglTime As Single) As String
'*************************************************
'Author: Unkwown
'Last modified: ?/?/?
'*************************************************
                           
' FormatTime:   Formats time in seconds to time in
'               Hours and/or Minutes and/or Seconds

' Determine how to display the time
Select Case sglTime
    Case 0 To 59    ' Seconds
        FormatTime = Format(sglTime, "0") & " sec"
    Case 60 To 3599 ' Minutes Seconds
        FormatTime = Format(Int(sglTime / 60), "#0") & _
                     " min " & _
                     Format(sglTime Mod 60, "0") & " sec"
    Case Else       ' Hours Minutes
        FormatTime = Format(Int(sglTime / 3600), "#0") & _
                     " hr " & _
                     Format(sglTime / 60 Mod 60, "0") & " min"
End Select

End Function

Public Function ReturnFileOrFolder(FullPath As String, _
                                   ReturnFile As Boolean, _
                                   Optional IsURL As Boolean = False) _
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

intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
ReturnFileOrFolder = IIf(ReturnFile, _
                         Right(FullPath, Len(FullPath) - intDelimiterIndex), _
                         Left(FullPath, intDelimiterIndex))

End Function

Public Function DownloadFile(strURL As String, strDestination As String, Optional UserName As String = "", Optional Password As String = "", Optional TACCESO As Integer = 0, Optional PROXY As String = "") As Boolean
'*************************************************
'Author: Jeff Cockayne
'Last modified: ?/?/?
'*************************************************

' Funtion DownloadFile
'
' Author:   Jeff Cockayne
'
' Inputs:   strURL String; the source URL of the file
'           strDestination; valid Win95/NT path to where you want it
'           (i.e. "C:\Program Files\My Stuff\Purina.pdf")
'
' Returns:  Boolean; Was the download successful?

Dim bData() As Byte         ' Data var
Dim intFile As Integer      ' FreeFile var
Dim a As Variant            ' Temp var
Dim intKReceived As Integer ' KB received so far
Dim intKFileLength As Long  ' KB total length of file
Dim lastTime As Single      ' time last chunk received
Dim sglRate As Single       ' var to hold transfer rate
Dim sglTime As Single       ' var to hold time remaining
Dim strFile As String       ' temp filename var
Dim strHeader As String     ' HTTP header store
Dim strHost As String       ' HTTP Host

On Local Error GoTo InternetErrorHandler

' Start with Cancel flag = False
CancelSearch = False
BotonCancel = False
' Get just filename (without dirs) for display
strFile = ReturnFileOrFolder(strDestination, True)
strHost = ReturnFileOrFolder(strURL, False, True)

' Show the status form
Me.Show

StatusLabel2 = "Reciviendo la información del archivo..."
DoEvents

' Download file
With Inet1
    .AccessType = TACCESO
    .PROXY = PROXY
    .URL = strURL
    .UserName = UserName
    .Password = Password
    .Execute , "GET"
End With

StatusLabel = "Guardando:" & vbCr & vbCr & strFile & " desde " _
              & IIf(Len(strHost) < 33, strHost, "..." & Left(strHost, 30))

lastTime = Timer

' While initiating connection, yield CPU to Windows
While Inet1.StillExecuting
    DoEvents
    ' If user pressed Cancel button on StatusForm
    ' then fail, cancel, and exit this download
    If CancelSearch Then
        GoTo ExitDownload
    End If
Wend

' Get first header ("HTTP/X.X XXX ...")
strHeader = Inet1.GetHeader

' Trap common HTTP Errors
Select Case Mid(strHeader, 10, 3)
    Case "200"  ' OK!

    Case "401"  ' Not authorized
        Me.Hide
        MsgBox "No me autorizaron!!", _
               vbCritical, _
               "Desautorizado"
        GoTo ExitDownload
    
    Case "404"  ' File Not Found
        Me.Hide
        MsgBox "El archivo, " & _
               """ & Inet1.URL & """ & _
               " no pudo ser encontrado!", _
               vbCritical, _
               "Archivo no encontrado"
        GoTo ExitDownload
        
    Case vbCrLf
        Me.Hide
        MsgBox "No pude establecer conexión." & vbCr & vbCr & _
               "Verifica la configuración del proxy (o no).", _
               vbExclamation, _
               "Sin conexión"
        GoTo ExitDownload
        
    Case Else
        ' Miscellaneous unexpected errors
        Me.Hide
        strHeader = Left(strHeader, InStr(strHeader, vbCr))
        MsgBox "Respuesta del server:" & vbCr & vbCr & _
               strHeader, _
               vbCritical, _
               "Error"
        GoTo ExitDownload
End Select

' Get file length with "Content-Length" header request
strHeader = Inet1.GetHeader("Content-Length")
intKFileLength = CInt(Val(strHeader) / 1024)
If intKFileLength = 0 Then
    ' Failed; File length would never be 0!
    GoTo ExitDownload
End If

' Prepare display
ProgressBar.value = 0
ProgressBar.Max = intKFileLength
'Animation1.Play
DoEvents

intKReceived = 0

On Local Error GoTo FileErrorHandler

' If no errors occurred, then spank the file to disk
If Inet1.ResponseCode = 0 Then
    intFile = FreeFile()        ' Set intFile to an unused file.
    ' Open a file to write to.
    Open strDestination For Binary Access Write As #intFile
    ' Get the first chunk.
    bData = Inet1.GetChunk(1024, icByteArray)
    a = bData                   ' Must assign array to ANOTHER var cus
                                ' VB has a cow with LenB(bData)!
    
    Do While LenB(a) > 0        ' while there's still data...
        Put #intFile, , bData   ' Put it into our destination file
        ' Get next chunk.
        bData = Inet1.GetChunk(1024, icByteArray)
        a = bData
        If CancelSearch Then
            Close #intFile
            Kill strDestination
            GoTo ExitDownload
        End If
        intKReceived = intKReceived + 1
        If intKReceived < intKFileLength Then   ' to avoid -1's
            sglRate = intKReceived / (Timer - lastTime)
            sglTime = (intKFileLength - intKReceived) / sglRate
            StatusLabel2 = "Tiempo restante estimado: " & _
                           FormatTime(sglTime) & _
                           " (" & _
                           FormatFileSize(intKReceived * 1024#) & _
                           " de " & _
                           FormatFileSize(intKFileLength * 1024#) & _
                           " copiado)" & vbCr & vbCr & _
                           "Velocidad: " & _
                           Format(sglRate, "###,##0.0") & " KB/Sec"
            ProgressBar.value = intKReceived
            Caption = Format((intKReceived / intKFileLength), "##0%") & _
                      " de " & strFile & " completado"
        End If
    Loop
    Put #intFile, , bData
    Close #intFile
End If

StatusLabel2 = Empty
DoEvents

ExitDownload:
If intKReceived >= intKFileLength Then
    StatusLabel = "Download completado!"
    DownloadSuccess = True
Else
    ' Delete partially downloaded file, if it exists
    DownloadSuccess = False
    If Not Dir(strDestination) = Empty Then Kill strDestination
    If Not CancelSearch Then
        StatusLabel = "Descarga fallada!"
        MsgBox "Descarga fallada!", _
        vbCritical, _
        "Error descargando el archivo"
    End If
End If

' Make sure that the Internet connection is closed
Inet1.Cancel
DoEvents
' and exit this function
Unload Me
DoEvents
Exit Function

InternetErrorHandler:
    CancelSearch = True
    Inet1.Cancel
    MsgBox "Error: " & Err.Description & " ocurrido.", _
           vbCritical, _
           "Error descargando el archivo"
    DoEvents
    DownloadSuccess = False
    BotonCancel = True
    Resume Next
    
FileErrorHandler:
    MsgBox "No pude escribir el archivo!", _
           vbCritical, _
           "Error de escritura"
    DownloadSuccess = False
    BotonCancel = True
    Resume Next
    
End Function

Private Sub CancelButton_Click()
'*************************************************
'Author: Unkwown
'Last modified: ?/?/?
'*************************************************
    CancelSearch = True
    BotonCancel = True
    DownloadSuccess = False
    frmLibrerias.descargando = False
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'*************************************************
'Author: Unkwown
'Last modified: ?/?/?
'*************************************************
    frmLibrerias.descargando = False
    CancelSearch = True
    Unload Me
End Sub
