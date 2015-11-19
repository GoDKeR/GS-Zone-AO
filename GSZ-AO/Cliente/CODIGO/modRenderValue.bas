Attribute VB_Name = "modRenderValue"
Option Explicit

' GS-Zone AO
' Basad en el Sistema de Daño aportado por maTih.- adaptado by ^[GS]^
' Fuente: http://www.gs-zone.org/dano_tds_style_en_mapa_tl6O.html
 
Const RENDER_TIME As Integer = 300

Enum RVType
     ePuñal = 1                'Apuñalo.
     eNormal = 2               'Golpe común.
     eMagic = 3                'Hechizo. ' GSZAO
     eGold = 4                 'Oro ' GSZAO
End Enum
 
Private RVNormalFont    As New StdFont
 
Type RVList
     RenderVal      As Integer  'Cantidad.
     ColorRGB       As Long     'Color.
     RenderType     As RVType   'Tipo, se usa para saber si es apu o no.
     'RenderFont     As New StdFont  'Efecto del apu.
     TimeRendered   As Integer  'Tiempo transcurrido.
     Downloading    As Byte     'Contador para la posicion Y.
     Activated      As Boolean  'Si está activado..
End Type
 
Sub Create(ByVal X As Byte, ByVal Y As Byte, ByVal ColorRGB As Long, ByVal rValue As Integer, ByVal eMode As Byte)
     
    ' @ Agrega un nuevo valor.
     
    With MapData(X, Y).RenderValue
         
         .Activated = True
         .ColorRGB = ColorRGB
         .RenderType = eMode
         .RenderVal = rValue
         .TimeRendered = 0
         .Downloading = 0
         
    End With
 
End Sub
 
Sub Draw(ByVal X As Byte, ByVal Y As Byte, ByVal PixelX As Integer, ByVal PixelY As Integer)
 
    ' @ Dibuja un valor
     
    With MapData(X, Y).RenderValue
         
         If (Not .Activated) Or (Not .RenderVal <> 0) Then Exit Sub
            If .TimeRendered < RENDER_TIME Then
            
                'Sumo el contador del tiempo.
                .TimeRendered = .TimeRendered + 1
                
                If (.TimeRendered / 2) > 0 Then
                    .Downloading = (.TimeRendered / 8)
                End If
                
                .ColorRGB = ModifyColour(.TimeRendered, .RenderType)
                    
                'Dibujo ; D
                If .RenderType <> eGold Then
                    Text_Render_Special (PixelX - 5), (PixelY + 30) - .Downloading, "-" & .RenderVal, .ColorRGB, True ' .RenderFont,
                Else ' el oro es  "+"
                    Text_Render_Special (PixelX - 5), (PixelY + 30) - .Downloading, "+" & .RenderVal, .ColorRGB, True ' .RenderFont,
                End If
               
                'Si llego al tiempo lo limpio
                If .TimeRendered >= RENDER_TIME Then
                   Call Clear(X, Y)
                End If
                
         End If
           
    End With
 
End Sub
 
Private Sub Clear(ByVal X As Byte, ByVal Y As Byte)
 
    ' @ Limpia todo.
     
    With MapData(X, Y).RenderValue
         .Activated = False
         .ColorRGB = 0
         .RenderVal = 0
         .TimeRendered = 0
    End With
 
End Sub
 
Private Function ModifyColour(ByVal TimeNowRendered As Integer, ByVal RenderType As RVType) As Long
 
    ' @ Se usa para los "efectos" en el tiempo.
    
    ' 512 ---- 255
    ' 120 ---- x = 255 * 120 / 512
    
    Dim TimeX2 As Integer
    TimeX2 = TimeNowRendered ' * 2
    If TimeX2 > 255 Then TimeX2 = 255
    
    Select Case RenderType
        Case RVType.ePuñal
            ModifyColour = RGB(255 - TimeX2, 255, TimeX2)
        Case RVType.eNormal
            ModifyColour = RGB(TimeX2, 11, 255 - TimeX2)
        Case RVType.eMagic
            ModifyColour = RGB(255 - TimeX2, 255 - TimeX2, TimeX2)
        Case RVType.eGold
            ModifyColour = RGB(1, 240, 255)
    End Select
 
End Function
 
