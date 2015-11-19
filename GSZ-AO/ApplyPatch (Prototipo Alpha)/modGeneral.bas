Attribute VB_Name = "modGeneral"
Option Explicit

Const PROCESS_QUERY_INFORMATION = &H400&
Const SYNCHRONIZE = &H100000
Const INFINITE = -1&

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Sub ShellWait(PathName As String, WindowStyle As VbAppWinStyle)
    Dim ID As Long, Proc As Long
    
    ID = Shell(PathName, WindowStyle)
    Proc = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION, 0, ID)
    WaitForSingleObject Proc, INFINITE
    CloseHandle Proc
End Sub

Sub Main()

Dim InCmd() As String
Dim fHelp As Boolean
Dim fSilent As Boolean

fSilent = False
fHelp = False

InCmd = Split(Command, " ", 5)
If UBound(InCmd) >= 1 Then
    Call WriteStdOut("DEBUG Command = " & Command & vbCrLf)
Else
    fHelp = True
End If

If fHelp = True Or fSilent = False Then
    Call WriteStdOut("*********************************************" & vbCrLf & _
                     "******* APPLYPATCH v" & App.Major & "." & App.Minor & "." & App.Revision & " by ^[GS]^ ********" & vbCrLf & _
                     "*********************************************" & vbCrLf)
    If fHelp = True Then
        Call WriteStdOut(vbCrLf & "Modo de Uso:" & vbCrLf & _
                         App.EXEName & ".exe [-v/-c/-a] [-s] archivo.AO parche.AO" & vbCrLf & vbCrLf & _
                         "Operaciones:" & vbCrLf & _
                         "-v archivo.AO               - Nos devuelve la version del archivo.AO" & vbCrLf & _
                         "-c archivo.AO parche.AO     - Nos devuelve si el archivo parche.AO es compatible con archivo.AO" & vbCrLf & _
                         "-a archivo.AO parche.AO     - Aplica el parche.AO a archivo.AO." & vbCrLf & _
                         "-s                          - Modo silencioso. (Se lo puede utilizar en combinacion con las demas operaciones)" & vbCrLf & vbCrLf & _
                         "Para informacion visita WWW.GS-ZONE.ORG" & vbCrLf)
    Else
    
    End If
End If

Call WriteStdOut("DEBUG UBound(InCmd) = " & UBound(InCmd) & vbCrLf)

ShellWait "cmd.exe /C pause", vbNormalFocus

End Sub
