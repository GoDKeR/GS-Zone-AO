Attribute VB_Name = "mod_DirectX"
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

Option Explicit

Private error As Boolean
Public x As String, Y As String, dxVer As String
Public DirectX As New DirectX8
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8
Public DirectDevice As Direct3DDevice8
Public DirectSound As DirectSound8
Public D3DWindow As D3DPRESENT_PARAMETERS
Public TestTexture As Direct3DTexture8
Public aColor As Long
Public aRed As Byte
Public aGreen As Byte
Public aBlue As Byte
Public aEst As Boolean
Public iFPS As Integer

Public Type TLVERTEX
    x As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

Public Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub InitDX8()
'*************************************************
'Author: ^[GS]^
'Last modified: 13/07/2012 - ^[GS]^
'*************************************************

    If FileExist(App.Path & iniTestBMP, vbArchive) = False Then Exit Sub ' [GS] Reviso que exista la imagen

    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate

    DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = IIf(frmAOSetup.chkVSync.value = 0, D3DSWAPEFFECT_COPY, D3DSWAPEFFECT_COPY_VSYNC)
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = 60
        .BackBufferHeight = 64
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmAOSetup.DirectDrawTest.hWnd
    End With

    Dim ModoD3D As CONST_D3DCREATEFLAGS

    If frmAOSetup.cProcesar.ListIndex = 0 Then ' Software
        ModoD3D = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    ElseIf frmAOSetup.cProcesar.ListIndex = 1 Then ' Hardware
        ModoD3D = D3DCREATE_HARDWARE_VERTEXPROCESSING
    ElseIf frmAOSetup.cProcesar.ListIndex = 2 Then ' Software & Hardware
        ModoD3D = D3DCREATE_MIXED_VERTEXPROCESSING
    End If
    
    Set DirectDevice = DirectD3D.CreateDevice( _
                    D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                    frmAOSetup.DirectDrawTest.hWnd, _
                    ModoD3D, _
                    D3DWindow)

                        
    With DirectDevice
      .SetVertexShader FVF
      .SetRenderState D3DRS_LIGHTING, 0
      .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
      .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
      .SetRenderState D3DRS_ALPHABLENDENABLE, False
    End With

    If DirectDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectDevice. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
    
    Set DirectD3D8 = New D3DX8
    
    Set TestTexture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, App.Path & iniTestBMP, _
                D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, RGB(255, 0, 255), ByVal 0, ByVal 0)
    
    aRed = 0
    aGreen = 0
    aBlue = 0
    iFPS = 0
    aEst = True

End Sub



Public Sub DrawDX8()
'*************************************************
'Author: ^[GS]^
'Last modified: 13/07/2012 - ^[GS]^
'*************************************************

    Dim spriteVerts(3)  As D3DTLVERTEX
    Dim I               As Byte
    
    Call ModifColor

    spriteVerts(0) = CrearVertice(0, 0, 0, 1, aColor, 0, 0, 0)
    spriteVerts(1) = CrearVertice(64, 0, 0, 1, aColor, 0, 1, 0)
    spriteVerts(2) = CrearVertice(0, 64, 0, 1, aColor, 0, 0, 1)
    spriteVerts(3) = CrearVertice(64, 64, 0, 1, aColor, 0, 1, 1)
    
    DirectDevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 0#, 0
    DirectDevice.BeginScene
    DirectDevice.SetTexture 0, TestTexture
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, spriteVerts(0), Len(spriteVerts(0))
    DirectDevice.EndScene
    DirectDevice.Present ByVal 0&, ByVal 0&, frmAOSetup.DirectDrawTest.hWnd, ByVal 0&
    iFPS = iFPS + 1
    
End Sub

Private Sub ModifColor()
    If aEst = True Then
        aRed = aRed + 25
        aGreen = aGreen + 25
        aBlue = aBlue + 25
        If aRed >= 235 Then aEst = False
    Else
        aRed = aRed - 25
        aGreen = aGreen - 25
        aBlue = aBlue - 25
        If aRed <= 25 Then aEst = True
    End If
    aColor = RGB(aRed, aGreen, aBlue)
End Sub

Private Function CrearVertice(ByVal x As Single, ByVal Y As Single, ByVal Z As Single, ByVal rhw As Single, ByVal Color As Long, ByVal specular As Long, ByVal tu As Single, ByVal tv As Single) As D3DTLVERTEX
    CrearVertice.sx = x
    CrearVertice.sy = Y
    CrearVertice.sz = Z
    CrearVertice.rhw = rhw
    CrearVertice.Color = Color
    CrearVertice.specular = specular
    CrearVertice.tu = tu
    CrearVertice.tv = tv
End Function


Public Sub VersionDirectX()
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 29/07/2012 - ^[GS]^
'de dxdiag termine para evitar errores (se hace asincrónicamente y puede no terminar antes del open)
'*************************************************
On Error GoTo ErrHandler

    Dim handle As Integer
    
    Call Shell("dxdiag " & App.Path & iniTempDX)
    Do Until FileExist(App.Path & iniTempDX, vbArchive)
        DoEvents
    Loop
    Sleep 500
    
    handle = FreeFile
    Open App.Path & iniTempDX For Input As handle
    Do While Not EOF(handle)
        Line Input #handle, x
        If Left$(LTrim$(x), 15) = "DirectX Version" Then
            Y = LTrim$(x)
            dxVer = Mid$(Y, 17, Len(Y) - 16)
            frmAOSetup.lDirectX.Caption = dxVer
            Exit Do
        End If
    Loop
    Close handle
    If FileExist(App.Path & iniTempDX, vbArchive) Then
        Call Kill(App.Path & iniTempDX)
    End If
    
Exit Sub

ErrHandler:
    If Err.Number = 70 Then
        'Permission denied. Es posible el archivo exista pero como todavía
        'está abierto por dxdiag no podemos abrirlo. Dormimos 5 ms e intentamos de nuevo
        Sleep 5
        Resume
    End If
    frmAOSetup.lDirectX.Caption = "Error"
    If FileExist(App.Path & iniTempDX, vbArchive) Then
        Call Kill(App.Path & iniTempDX)
    End If
End Sub

Public Sub ProbarDirectX()
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 10/03/06
'10/03/06: ^[GS]^ - Adapte el codigo al nuevo formulario principal
'*************************************************
    error = False
    
    'Create DX object
    If IniciarDXobject(DirectX) Then
        frmAOSetup.lblDX.Visible = True
        frmAOSetup.lblDX.Caption = "OK"
        frmAOSetup.lblDX.ForeColor = &H8000&
        frmAOSetup.lblDX.Font.Bold = True
    Else
        frmAOSetup.lblDX.Visible = True
        frmAOSetup.lblDX.Caption = "OK"
        frmAOSetup.lblDX.ForeColor = RGB(255, 0, 0)
        frmAOSetup.lblDX.Font.Bold = True
        error = True
    End If
    
    frmAOSetup.Text1.BackColor = frmAOSetup.lblDX.ForeColor
    DoEvents
    
    'Create DirectSound
    If IniciarDirectSound() Then
        frmAOSetup.lblDS.Visible = True
        frmAOSetup.lblDS.Caption = "OK"
        frmAOSetup.lblDS.ForeColor = &H8000&
        frmAOSetup.lblDS.Font.Bold = True
    Else
        frmAOSetup.lblDS.Visible = True
        frmAOSetup.lblDS.Caption = "OK"
        frmAOSetup.lblDS.ForeColor = RGB(255, 0, 0)
        frmAOSetup.lblDS.Font.Bold = True
        error = True
    End If
    frmAOSetup.Text3.BackColor = frmAOSetup.lblDS.ForeColor
    DoEvents
    
    'Create DirectDraw
    If IniciarDDobject(DirectX) Then
        frmAOSetup.lblDD.Visible = True
        frmAOSetup.lblDD.Caption = "OK"
        frmAOSetup.lblDD.ForeColor = &H8000&
        frmAOSetup.lblDD.Font.Bold = True
    Else
        frmAOSetup.lblDD.Visible = True
        frmAOSetup.lblDD.Caption = "ERROR"
        frmAOSetup.lblDD.ForeColor = RGB(255, 0, 0)
        frmAOSetup.lblDD.Font.Bold = True
        error = True
    End If
    frmAOSetup.Text2.BackColor = frmAOSetup.lblDD.ForeColor
    DoEvents
    
    If error Then
        MsgBox "Necesita reinstalar DirectX", vbCritical, "Argentum Online Setup"
    End If
End Sub

Private Function IniciarDirectSound() As Boolean
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'*************************************************
On Error Resume Next
    Set DirectSound = DirectX.DirectSoundCreate("")
    
    If Err Then
        IniciarDirectSound = False
        Exit Function
    End If
    
    IniciarDirectSound = True
End Function

Private Function IniciarDXobject(ByRef dx As DirectX8) As Boolean
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 09/07/2012 - ^[GS]^
'*************************************************
On Error Resume Next

    Set dx = New DirectX8
    
    If Err Then
        Err.Clear
        IniciarDXobject = False
        Exit Function
    End If
    
    IniciarDXobject = True
End Function

Private Function IniciarDDobject(ByRef Dd As DirectX8) As Boolean
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 09/07/2012 - ^[GS]^
'*************************************************
On Error Resume Next

    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
    
    If Err Then
        Err.Clear
        IniciarDDobject = False
        Exit Function
    End If
    
    IniciarDDobject = True
End Function

