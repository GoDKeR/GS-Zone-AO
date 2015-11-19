Attribute VB_Name = "modAmbiente"
Option Explicit

'AMBIENTE TWIST
Public ColorActual As tColor, ColorFinal As tColor, Fade As Boolean, AmbientLastCheck As Long
Public base_light As Long

Public Sub Ambient_Set(LightR As Byte, LightG As Byte, LightB As Byte, Optional ByVal Fade As Boolean = True)
    Dim light As tColor
    light.r = LightR
    light.g = LightG
    light.b = LightB
    ColorFinal = light
    If Not Fade Then
        ColorActual = light
    End If
End Sub

Public Sub Ambient_Check()
    Dim Hora As Byte, Minutos As Byte
    Hora = Hour(Time)
    Minutos = Minute(Time)
    
    If Hora >= 6 And Hora < 8 Then
        'Amanecer
        Ambient_SetFinal 195, 175, 120
    ElseIf Hora >= 8 And Hora < 16 Then
        'Dia
        Ambient_SetFinal 255, 255, 255
    ElseIf Hora >= 16 And (Hora < 20 Or (Hora = 20 And Minutos < 30)) Then
        'Tarde
        Ambient_SetFinal 150, 150, 160
    ElseIf (Hora = 20 And Minutos >= 30 And Minutos < 35) Then
        'Anochecer
        Ambient_SetFinal 165, 130, 40
    ElseIf (((Hora > 20) Or (Hora = 20 And Minutos >= 35)) And Hora < 23) Or _
              Hora >= 3 And Hora < 6 Then
        'Noche
        Ambient_SetFinal 80, 80, 100
    ElseIf (Hora = 23 Or Hora = 0) Or (Hora > 0 And Hora < 3) Then
        'Noche mas oscura
        Ambient_SetFinal 60, 60, 120
    End If
       ' Ambient_SetFinal 80, 80, 100
    With ColorActual
        'Red
        If .r < ColorFinal.r Then
            .r = .r + 1
        ElseIf .r > ColorFinal.r Then
            .r = .r - 1
        End If
        'Green
        If .g < ColorFinal.g Then
            .g = .g + 1
        ElseIf .g > ColorFinal.g Then
            .g = .g - 1
        End If
        'Blue
        If .b < ColorFinal.b Then
            .b = .b + 1
        ElseIf .b > ColorFinal.b Then
            .b = .b - 1
        End If
        base_light = ARGB(.r, .g, .b, 255)
    End With
End Sub

Public Sub Ambient_Fade()
    With ColorActual
        'Red
        If .r < ColorFinal.r Then
            .r = .r + 1
        ElseIf .r > ColorFinal.r Then
            .r = .r - 1
        End If
        'Green
        If .g < ColorFinal.g Then
            .g = .g + 1
        ElseIf .g > ColorFinal.g Then
            .g = .g - 1
        End If
        'Blue
        If .b < ColorFinal.b Then
            .b = .b + 1
        ElseIf .b > ColorFinal.b Then
            .b = .b - 1
        End If
        base_light = ARGB(.r, .g, .b, 255)
    End With
    'Map_LightRenderAll
    Fade = Not (ColorFinal.r = ColorActual.r And ColorFinal.g = ColorActual.g And ColorFinal.b = ColorActual.b)
End Sub

Public Sub Ambient_Start()
    Dim Hora As Byte, Minutos As Byte
    Hora = Hour(Time)
    Minutos = Minute(Time)
    
    If Hora >= 6 And Hora < 8 Then
        'Amanecer
        Ambient_SetFinal 195, 175, 120
        Ambient_SetActual 195, 175, 120
    ElseIf Hora >= 8 And Hora < 16 Then
        'Dia
        Ambient_SetFinal 255, 255, 255
        Ambient_SetActual 255, 255, 255
    ElseIf Hora >= 16 And (Hora < 20 Or (Hora = 20 And Minutos < 30)) Then
        'Tarde
        Ambient_SetFinal 150, 150, 160
        Ambient_SetActual 150, 150, 160
    ElseIf (Hora = 20 And Minutos >= 30 And Minutos < 35) Then
        'Anochecer
        Ambient_SetFinal 165, 130, 40
        Ambient_SetActual 165, 130, 40
    ElseIf (((Hora > 20) Or (Hora = 20 And Minutos >= 35)) And Hora < 23) Or _
              Hora >= 3 And Hora < 6 Then
        'Noche
        Ambient_SetFinal 80, 80, 100
        Ambient_SetActual 80, 80, 100
    ElseIf (Hora = 23 Or Hora = 0) Or (Hora > 0 And Hora < 3) Then
        'Noche mas oscura
        Ambient_SetFinal 60, 60, 120
        Ambient_SetActual 60, 60, 120
    End If
    With ColorActual
        base_light = ARGB(.r, .g, .b, 255)
    End With
End Sub

Private Sub Ambient_SetFinal(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
    ColorFinal.r = r: ColorFinal.g = g: ColorFinal.b = b
    'Map_LightRenderAll
    Fade = Not (ColorFinal.r = ColorActual.r And ColorFinal.g = ColorActual.g And ColorFinal.b = ColorActual.b)
End Sub

Private Sub Ambient_SetActual(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
    ColorActual.r = r: ColorActual.g = g: ColorActual.b = b
    Fade = Not (ColorFinal.r = ColorActual.r And ColorFinal.g = ColorActual.g And ColorFinal.b = ColorActual.b)
End Sub

'Public Sub ColorAmbiente(ByRef AmbientColor As D3DCOLORVALUE)
'    AmbientColor.r = ColorActual.r
'    AmbientColor.g = ColorActual.g
'    AmbientColor.b = ColorActual.b
'End Sub

'AMBIENTE TWIST
