VERSION 5.00
Begin VB.UserControl ctlDownload 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   Picture         =   "ctlAsyncDownload.ctx":0000
   ScaleHeight     =   960
   ScaleWidth      =   960
End
Attribute VB_Name = "ctlDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This component doesn't require any external calls/apis/references
' Obtains an url to a byte array using native VB6 calls
' Asyncronous - no need to wait for data to arrive
' Multiple downloads are accepted at the same time (different URL's, etc)
' If you like this code, please VOTE for it
' You may use this code freely in your projects, but whenever possible,
' include my name 'Filipe Lage' on the 'Help->About' or something ;)
' Cheers :)
'
' Filipe Lage
' fclage@ezlinkng.com
'

Public Event Progress(x As AsyncProperty, percent As Single)
Public Event Finished(x As AsyncProperty)
Public CurrentDownloads As New Collection

Public Function Download(xurl As String) As Boolean
On Error Resume Next
UserControl.AsyncRead xurl, vbAsyncTypeByteArray, xurl, vbAsyncReadForceUpdate
CurrentDownloads.Add xurl, xurl
RefreshStatus
End Function

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
RaiseEvent Finished(AsyncProp)
On Error Resume Next
CurrentDownloads.Remove AsyncProp.PropertyName
RefreshStatus
On Error GoTo 0
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
Dim p As Single
If AsyncProp.BytesMax > 0 Then p = 100 * (AsyncProp.BytesRead / AsyncProp.BytesMax) Else p = 0
RaiseEvent Progress(AsyncProp, p)
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 960
UserControl.Height = 960
End Sub

Public Sub CancelDownload(xurl As String)
On Error Resume Next
UserControl.CancelAsyncRead CurrentDownloads(xurl)
CurrentDownloads.Remove xurl
On Error GoTo 0
End Sub

Private Sub UserControl_Show()
If UIMode = True Then
    Else
    UserControl.Extender.Visible = False
    End If
End Sub

Private Sub UserControl_Terminate()
Do Until CurrentDownloads.Count = 0
CancelDownload CurrentDownloads(1)
Loop
End Sub

Private Sub RefreshStatus()
UserControl.Cls
UserControl.CurrentX = 0
UserControl.CurrentY = 0
UserControl.Print CurrentDownloads.Count
End Sub

Private Function UIMode() As Boolean
On Error Resume Next
Err.Clear
Debug.Print 1 / 0
UIMode = (Err.Number <> 0)
On Error GoTo 0
End Function
