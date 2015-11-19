VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00101010&
   Caption         =   "World Grid Maker"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   12240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pBuffer 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   10200
      ScaleHeight     =   915
      ScaleWidth      =   1260
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdDesplazar 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Desplazar Mundo"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdArr 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Arr."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAba 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Aba."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdDer 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Der."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdIzq 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Izq."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tUpdate 
      Interval        =   1000
      Left            =   6240
      Top             =   120
   End
   Begin VB.PictureBox pBase 
      BackColor       =   &H00000000&
      Height          =   6495
      Left            =   120
      ScaleHeight     =   6435
      ScaleWidth      =   11340
      TabIndex        =   2
      Top             =   600
      Width           =   11400
      Begin VB.PictureBox pMap 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1470
         Index           =   0
         Left            =   4920
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   96
         TabIndex        =   3
         Top             =   2400
         Visible         =   0   'False
         Width           =   1470
      End
   End
   Begin VB.VScrollBar VScroll 
      Height          =   6495
      Left            =   11520
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.HScrollBar HScroll 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7200
      Width           =   11415
   End
   Begin VB.Label pMapI 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Mapa: 0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label pMapI 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Pos: 0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu mnuGuardarComo 
         Caption         =   "Guardar &como..."
      End
      Begin VB.Menu mnuLine0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExportarPNG 
         Caption         =   "Exportar como &PNG"
      End
      Begin VB.Menu mnuExportarJPG 
         Caption         =   "Exportar como &JPG"
      End
      Begin VB.Menu mnuExportarBMP 
         Caption         =   "Exportar como &BMP"
      End
      Begin VB.Menu mnuExportarExcel 
         Caption         =   "Exportar como &XLS de Office Excel"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuGrid 
         Caption         =   "Mallado (Grid)"
      End
      Begin VB.Menu mnuNumMap 
         Caption         =   "Número de Mapa"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuModificarTamaño 
         Caption         =   "Modificar &Tamaño del Mundo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAcerca 
         Caption         =   "&Acerca de..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'World Grid Maker
'Copyright (C) 2012 GS-Zone
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
'You can contact me at:
'info@gs-zone.org

Option Explicit


Private Sub DrawGrid()
On Error Resume Next

    Dim tP As Integer
    
    For tP = 0 To MaxGrid
        If cVerGrid = True Then
            If pMAP(tP).BorderStyle <> 1 Then pMAP(tP).BorderStyle = 1
            If LenB(pMAP(tP).Tag) = 0 Then
                If pMAP(tP).BackColor <> vbBlack Then pMAP(tP).BackColor = vbBlack
            Else
                If pMAP(tP).BackColor <> vbRed Then pMAP(tP).BackColor = vbRed
            End If
        Else
            If pMAP(tP).BorderStyle <> 0 Then pMAP(tP).BorderStyle = 0
            If LenB(pMAP(tP).Tag) = 0 Then
                If pMAP(tP).BackColor <> vbBlack Then pMAP(tP).BackColor = vbBlack
            Else
                If pMAP(tP).BackColor <> vbRed Then pMAP(tP).BackColor = vbRed
            End If
        End If
    Next
    
End Sub

Private Sub cmdAba_Click()
On Error Resume Next

    If aSize <= 1 Then Exit Sub

    Dim tP As Integer
    
    For tP = MaxGrid To (aSize + 1) Step -1
        pMAP(tP).Picture = pMAP(tP - aSize).Picture
        pMAP(tP).Tag = pMAP(tP - aSize).Tag
        pMAP(tP).BackColor = pMAP(tP - aSize).BackColor
        WorldGrid(tP) = WorldGrid(tP - aSize)
    Next
    Changed = True
    
End Sub

Private Sub cmdArr_Click()
On Error Resume Next
    
    If aSize <= 1 Then Exit Sub

    Dim tP As Integer
    
    For tP = (aSize + 1) To MaxGrid
        pMAP(tP - aSize).Picture = pMAP(tP).Picture
        pMAP(tP - aSize).Tag = pMAP(tP).Tag
        pMAP(tP - aSize).BackColor = pMAP(tP).BackColor
        WorldGrid(tP - aSize) = WorldGrid(tP)
    Next
    Changed = True
    
End Sub

Private Sub cmdDer_Click()
    If aSize <= 1 Then Exit Sub

    Dim tP As Integer
    
    For tP = MaxGrid - 1 To 2 Step -1
        pMAP(tP).Picture = pMAP(tP - 1).Picture
        pMAP(tP).Tag = pMAP(tP - 1).Tag
        pMAP(tP).BackColor = pMAP(tP - 1).BackColor
        WorldGrid(tP) = WorldGrid(tP - 1)
    Next
    Changed = True
    
End Sub

Private Sub cmdDesplazar_Click()
If cmdDesplazar.Caption = "&Desplazar Mundo" Then
    cmdArr.Visible = True
    cmdAba.Visible = True
    cmdDer.Visible = True
    cmdIzq.Visible = True
    cmdDesplazar.Caption = "&Ocultar"
Else
    cmdArr.Visible = False
    cmdAba.Visible = False
    cmdDer.Visible = False
    cmdIzq.Visible = False
    cmdDesplazar.Caption = "&Desplazar Mundo"
End If

End Sub

Private Sub cmdIzq_Click()
On Error Resume Next
    If aSize <= 1 Then Exit Sub

    Dim tP As Integer
    
    For tP = 2 To MaxGrid
        pMAP(tP - 1).Picture = pMAP(tP).Picture
        pMAP(tP - 1).Tag = pMAP(tP).Tag
        pMAP(tP - 1).BackColor = pMAP(tP).BackColor
        WorldGrid(tP - 1) = WorldGrid(tP)
    Next
    Changed = True
    
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    Call ReadConfig
    mnuGrid.Checked = cVerGrid
    mnuNumMap.Checked = cVerNumMap
    ActualFile = vbNullString
    Call SetWorldSize(cDefaultSize)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

    If Changed = True Then
        Dim vRes As VbMsgBoxResult
        vRes = MsgBox("Los últimos cambios realizados en el WorldGrid no han sido guardados." & vbCrLf & "¿Está seguro que desea salir?", vbYesNo + vbQuestion, "¿Desea salir SIN Guardar los cambios?")
        If vRes = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Call SaveConfig
    DoEvents
    End
End Sub

Private Sub Form_Resize()
On Error Resume Next

    HScroll.Top = frmMain.Height - 1350
    VScroll.Left = frmMain.Width - 750
    pBase.Height = HScroll.Top - pBase.Top
    pBase.Width = VScroll.Left - pBase.Left
    VScroll.Height = pBase.Height
    HScroll.Width = pBase.Width
End Sub



Private Sub mnuAbrir_Click()
On Error GoTo Fallo

    Dim sF As String
    sF = ExplorarArchivo(frmMain.hWnd, "WorldGrid (*.grid)" & Chr$(0) & "*.grid", "Abrir WorldGrid...")
    If LenB(sF) <> 0 And FileExist(sF, vbArchive) Then
        Call LoadWorldGrid(sF)
        ActualFile = sF
        mnuModificarTamaño.Enabled = True
        'MsgBox "WorldGrid cargado con exito.", vbInformation + vbOKOnly
    End If
    Exit Sub
Fallo:
    MsgBox "Ha ocurrido un error durante la carga del WorldGrid." & vbCrLf & "Error " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
End Sub

Private Sub mnuAcerca_Click()
    MsgBox "Programado por ^[GS]^ para GS-Zone AO" & vbCrLf & "Website: http://www.gs-zone.org" & vbCrLf & "Programa bajo licencia Affero General Public License.", vbInformation + vbOKOnly
End Sub

Private Sub mnuExportarBMP_Click()
On Error GoTo Fallo

    If aSize <= 1 Then Exit Sub
    Dim sFile As String
    If LenB(ActualFile) = 0 Then
        sFile = "Nuevo"
    Else
        sFile = Left$(ActualFile, Len(ActualFile) - 5)
    End If
    If FileExist(sFile & ".bmp", vbArchive) Then Call Kill(sFile & ".bmp")
    
    Call DibujarBuffer
        
    SavePicture pBuffer.Image, sFile & ".bmp"
    
    Call DibujarBuffer(True)
    
    MsgBox "WorldGrid exportado a BMP con exito.", vbInformation + vbOKOnly
    Exit Sub
Fallo:
    MsgBox "Ha ocurrido un error durante la exportación BMP del WorldGrid." & vbCrLf & "Error " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
    
End Sub

Private Sub mnuExportarExcel_Click()
On Error GoTo Fallo

    If aSize <= 1 Then Exit Sub
    Dim sFile As String
    If LenB(ActualFile) = 0 Then
        sFile = "Nuevo"
    Else
        sFile = Left$(ActualFile, Len(ActualFile) - 5)
    End If
    If FileExist(sFile & ".xls", vbArchive) Then Call Kill(sFile & ".xls")
    
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    
    Dim GridSize As Integer
    Dim col As Byte
    Dim fil As Byte
    Dim tP As Integer
    
    GridSize = Sqr(MaxGrid)
    
    col = 1
    fil = 1
    For tP = 1 To MaxGrid
        If (WorldGrid(tP) <> 0) Then
            oSheet.cells(col, fil).value = WorldGrid(tP)
            oSheet.cells(col, fil).interior.Color = vbYellow
        Else
            oSheet.cells(col, fil).value = vbNullString
            oSheet.cells(col, fil).interior.Color = vbBlack
        End If
        fil = fil + 1
        If (fil > GridSize) Then
            fil = 1
            col = col + 1
        End If
    Next
    
    oBook.SaveAs sFile & ".xls"
    oExcel.Quit
    
    MsgBox "WorldGrid exportado a XLS con exito.", vbInformation + vbOKOnly
    
    Exit Sub
Fallo:
    MsgBox "Ha ocurrido un error durante la exportación XLS del WorldGrid." & vbCrLf & "Error " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

Private Sub mnuExportarPNG_Click()
On Error GoTo Fallo

    If aSize <= 1 Then Exit Sub
    Dim sFile As String
    If LenB(ActualFile) = 0 Then
        sFile = "Nuevo"
    Else
        sFile = Left$(ActualFile, Len(ActualFile) - 5)
    End If
    If FileExist(sFile & ".temp.bmp", vbArchive) Then Call Kill(sFile & ".temp.bmp")
    If FileExist(sFile & ".png", vbArchive) Then Call Kill(sFile & ".png")
    
    Call DibujarBuffer
           
    SavePicture pBuffer.Image, sFile & ".temp.bmp"
    
    Call WIAImageProcess(sFile & ".temp.bmp", sFile & ".png", wiaFormatPNG)
    
    Call Kill(sFile & ".temp.bmp")
    
    Call DibujarBuffer(True)
    
    MsgBox "WorldGrid exportado a PNG con exito.", vbInformation + vbOKOnly
    
    Exit Sub
Fallo:
    MsgBox "Ha ocurrido un error durante la exportación PNG del WorldGrid." & vbCrLf & "Error " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
   
End Sub

Private Sub mnuExportarJPG_Click()
On Error GoTo Fallo

    If aSize <= 1 Then Exit Sub
    Dim sFile As String
    If LenB(ActualFile) = 0 Then
        sFile = "Nuevo"
    Else
        sFile = Left$(ActualFile, Len(ActualFile) - 5)
    End If
    If FileExist(sFile & ".temp.bmp", vbArchive) Then Call Kill(sFile & ".temp.bmp")
    If FileExist(sFile & ".jpg", vbArchive) Then Call Kill(sFile & ".jpg")
    
    Call DibujarBuffer
       
    SavePicture pBuffer.Image, sFile & ".temp.bmp"
    
    Call WIAImageProcess(sFile & ".temp.bmp", sFile & ".jpg", wiaFormatJPEG, 100)
    
    Call Kill(sFile & ".temp.bmp")
    
    Call DibujarBuffer(True)
    
    MsgBox "WorldGrid exportado a JPG con exito.", vbInformation + vbOKOnly
    
    Exit Sub
Fallo:
    MsgBox "Ha ocurrido un error durante la exportación JPG del WorldGrid." & vbCrLf & "Error " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
   
End Sub


Private Sub mnuGrid_Click()
On Error Resume Next
    mnuGrid.Checked = Not mnuGrid.Checked
    cVerGrid = mnuGrid.Checked
    Call DrawGrid
End Sub

Private Sub mnuGuardar_Click()
On Error GoTo Fallo

    If LenB(ActualFile) = 0 Then
        mnuGuardarComo_Click
        Exit Sub
    End If
    Call SaveWorldGrid(ActualFile)
    MsgBox "WorldGrid guardado con exito.", vbInformation + vbOKOnly
    Exit Sub
Fallo:
    MsgBox "Ha ocurrido un error durante el guardado del WorldGrid." & vbCrLf & "Error " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
End Sub

Private Sub mnuGuardarComo_Click()
On Error GoTo Fallo

    Dim sF As String
    sF = ExplorarArchivo(frmMain.hWnd, "WorldGrid (*.grid)" & Chr$(0) & "*.grid", "Guardar WorldGrid como...", True)
    If LenB(sF) <> 0 And FileExist(ParsePath(sF, vbDirectory), vbDirectory) Then
        Call SaveWorldGrid(sF)
        MsgBox "WorldGrid guardado con exito.", vbInformation + vbOKOnly
    End If
    Exit Sub
Fallo:
    MsgBox "Ha ocurrido un error durante el guardado del WorldGrid." & vbCrLf & "Error " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

Private Sub mnuModificarTamaño_Click()
On Error Resume Next

    If LenB(ActualFile) = 0 Then
        Exit Sub
    Else
        Dim iT As String
        Dim iC As VbMsgBoxResult
        iT = InputBox("Introduzca el nuevo tamaño del mundo:" & vbCrLf & "Tamaño^2 = Capacidad Total de Mapas" & vbCrLf & "Si el tamaño es 10 la capacidad será de 100 mapas.", "Tamaño", Sqr(MaxGrid))
        If IsNumeric(iT) Then
            iT = Val(iT)
            If (iT ^ 2) > MaxGrid Then
                iC = MsgBox("¿Está seguro que desea expandir el tamaño del mundo de " & MaxGrid & " a " & (iT ^ 2) & "?", vbQuestion + vbYesNo)
                If (iC = vbYes) Then
                    Call SetWorldSize(iT, True)
                    Changed = True
                End If
            ElseIf (iT ^ 2) > 1 And (iT ^ 2) < MaxGrid Then
                iC = MsgBox("¿Está seguro que desea reducir el tamaño del mundo de " & MaxGrid & " a " & (iT ^ 2) & "?" & vbCrLf & vbCrLf & "ADVERTENCIA: Algunos mapas podrían ser eliminados del Grid.", vbQuestion + vbYesNo)
                If (iC = vbYes) Then
                    MsgBox "La opción de reducción aun no se encuentra diposible.", vbInformation + vbOKOnly
                    'Call SetWorldSize(iT, True)
                    'Changed = True
                End If
            End If
        End If
    End If
End Sub

Private Sub mnuNuevo_Click()
On Error Resume Next

    Dim iT As String
    iT = InputBox("Introduzca el tamaño del mundo:" & vbCrLf & "Tamaño^2 = Capacidad Total de Mapas" & vbCrLf & "Si el tamaño es 10 la capacidad será de 100 mapas.", "Tamaño", 10)
    If IsNumeric(iT) Then
        If Val(iT) > 1 Then
            ActualFile = vbNullString
            mnuModificarTamaño.Enabled = False
            Call SetWorldSize(Val(iT))
            Changed = False
        End If
    End If
End Sub

Private Sub mnuNumMap_Click()
    mnuNumMap.Checked = Not mnuNumMap.Checked
    cVerNumMap = mnuNumMap.Checked
    Call DibujarBuffer(True)
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

Private Sub pMap_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    
    If Button = 1 Then
        SelMapPos = Index
        frmSelMap.Show vbModal
    ElseIf Button = 2 Then
        SelMapPos = Index
        Call UpdateGrid(WorldGrid(Index))
        
        'MsgBox "haga click en el grid donde desea mover este mapa"
    End If
End Sub

Private Sub pMap_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Fallo
    If SobrePos <> Index Then
        Dim tP As Integer
        pMapI(0).Caption = "Pos: " & Index
        pMapI(1).Caption = "Mapa: " & WorldGrid(Index)
        For tP = 1 To MaxGrid
            If tP = Index Then
                If pMAP(tP).BorderStyle <> 1 Then pMAP(tP).BorderStyle = 1
                If pMAP(tP).Appearance <> 1 Then pMAP(tP).Appearance = 1
                If LenB(pMAP(tP).Tag) = 0 Then
                    If pMAP(tP).BackColor <> vbBlack Then pMAP(tP).BackColor = vbBlack
                Else
                    If pMAP(tP).BackColor <> vbRed Then pMAP(tP).BackColor = vbRed
                End If
            Else
                If pMAP(tP).BorderStyle <> 0 And cVerGrid = False Then pMAP(tP).BorderStyle = 0
                If pMAP(tP).Appearance <> 0 Then pMAP(tP).Appearance = 0
                If LenB(pMAP(tP).Tag) = 0 Then
                    If pMAP(tP).BackColor <> vbBlack Then pMAP(tP).BackColor = vbBlack
                Else
                    If pMAP(tP).BackColor <> vbRed Then pMAP(tP).BackColor = vbRed
                End If
            End If
        Next
        pBase.Refresh
        SobrePos = Index
    End If
Fallo:

End Sub

Private Sub pMap_Paint(Index As Integer)
On Error Resume Next

    Call UpdateGridCaptions
End Sub

Private Sub tUpdate_Timer()
On Error Resume Next

    Call UpdateCaption
End Sub

Private Sub VScroll_Scroll()
On Error Resume Next
    
    Call VScroll_Change
End Sub


Private Sub VScroll_Change()
On Error Resume Next
    
    ' Desplazamos las vistas internas... >_<
    If MaxGrid < 1 Then Exit Sub
    Dim iP As Integer
    pBase.Visible = False
    For iP = 1 To MaxGrid
        If aPosV < VScroll.value Then
            pMAP(iP).Top = pMAP(iP).Top - ((VScroll.value - aPosV) * pMAP(0).Height)
        ElseIf aPosV > VScroll.value Then
            pMAP(iP).Top = pMAP(iP).Top + ((aPosV - VScroll.value) * pMAP(0).Height)
        End If
    Next
    pBase.Visible = True
    aPosV = VScroll.value

End Sub

Private Sub HScroll_Scroll()
On Error Resume Next
    
    Call HScroll_Change
End Sub

Private Sub HScroll_Change()
On Error Resume Next
    
    ' Desplazamos las vistas internas... >_<
    If MaxGrid < 1 Then Exit Sub
    Dim iP As Integer
    pBase.Visible = False
    For iP = 1 To MaxGrid
        If aPosH < HScroll.value Then
            pMAP(iP).Left = pMAP(iP).Left - ((HScroll.value - aPosH) * pMAP(0).Width)
        ElseIf aPosH > HScroll.value Then
            pMAP(iP).Left = pMAP(iP).Left + ((aPosH - HScroll.value) * pMAP(0).Width)
        End If
    Next
    pBase.Visible = True
    aPosH = HScroll.value

End Sub


Public Sub UpdateCaption()
On Error Resume Next

    frmMain.Caption = "World Grid Maker v" & App.Major & "." & App.Minor & "." & App.Revision & " ["
    If LenB(ActualFile) = 0 Then
        frmMain.Caption = frmMain.Caption & "NUEVO]"
    Else
        If Changed = False Then
            frmMain.Caption = frmMain.Caption & ActualFile & "] [GRID: " & Sqr(MaxGrid) & "]"
        Else
            frmMain.Caption = frmMain.Caption & ActualFile & "*] [GRID: " & Sqr(MaxGrid) & "]"
        End If
    End If

End Sub

Public Sub UpdateGridCaptions()
    If cVerNumMap = False Then Exit Sub

    Dim tP As Integer
    For tP = 1 To MaxGrid
        If cVerNumMap = True Then
            If WorldGrid(tP) <> 0 Then
                ' Sombreado...
                pMAP(tP).ForeColor = vbBlue
                pMAP(tP).CurrentX = 1
                pMAP(tP).CurrentY = 2
                pMAP(tP).Print "Mapa #" & WorldGrid(tP)
                pMAP(tP).CurrentX = 2
                pMAP(tP).CurrentY = 1
                pMAP(tP).Print "Mapa #" & WorldGrid(tP)
                pMAP(tP).CurrentX = 0
                pMAP(tP).CurrentY = 1
                pMAP(tP).Print "Mapa #" & WorldGrid(tP)
                pMAP(tP).CurrentX = 1
                pMAP(tP).CurrentY = 0
                pMAP(tP).Print "Mapa #" & WorldGrid(tP)
                ' Mapa...
                pMAP(tP).ForeColor = vbWhite
                pMAP(tP).CurrentX = 1
                pMAP(tP).CurrentY = 1
                pMAP(tP).Print "Mapa #" & WorldGrid(tP)
            End If
        End If
    Next
    
End Sub

Private Sub DibujarBuffer(Optional Border As Boolean = False)
On Error Resume Next
    
    Dim pH As Integer
    Dim pV As Integer
    Dim tP As Integer
    
    pBuffer.Cls
    
    pBuffer.AutoRedraw = True
    
    pBuffer.Height = pMAP(tP).Height * aSize
    pBuffer.Width = pMAP(tP).Width * aSize
    pBuffer.BackColor = vbGreen

    pH = 1
    pV = 1
    For tP = 1 To MaxGrid
        pMAP(tP).BorderStyle = 0
        pBuffer.PaintPicture pMAP(tP).Image, (pH - 1) * pMAP(0).Width, (pV - 1) * pMAP(0).Height
        If Border = True And cVerGrid = True Then
            pMAP(tP).BorderStyle = 1
        End If
        pH = pH + 1
        If pH > aSize Then
            pH = 1
            pV = pV + 1
        End If
    Next
    
    DoEvents
    
End Sub
