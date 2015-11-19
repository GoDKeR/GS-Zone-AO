Attribute VB_Name = "modCarteles"
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

Private Const XPosCartel = 80
Private Const YPosCartel = 80
Private Const MAXLONG = 40

'Carteles
Public Cartel As Boolean
Private Leyenda As String
Private LeyendaFormateada() As String
Private textura As Integer


Sub InitCartel(ByRef Ley As String, ByVal Grh As Integer)

    If Cartel Then Exit Sub
    
    Dim i As Integer, K As Integer, anti As Integer
    
    Leyenda = Ley
    textura = Grh
    Cartel = True
    
    ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2))) As String
                
    anti = 1
    K = 0
    i = 0
    Call DarFormato(Leyenda, i, K, anti)
    i = 0
    Do While ((LenB(LeyendaFormateada(i)) <> 0) And (i < UBound(LeyendaFormateada)))
       i = i + 1
    Loop
    
    ReDim Preserve LeyendaFormateada(0 To i) As String
    
End Sub


Private Function DarFormato(ByRef s As String, ByRef i As Integer, ByRef K As Integer, ByRef anti As Integer)

    If anti + i <= Len(s) + 1 Then
    
        If ((i >= MAXLONG) And mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
            LeyendaFormateada(K) = mid$(s, anti, i + 1)
            K = K + 1
            anti = anti + i + 1
            i = 0
        Else
            i = i + 1
        End If
        
        Call DarFormato(s, i, K, anti)
    End If
    
End Function


Sub DibujarCartel()

    If Not Cartel Then Exit Sub
    
    Dim X As Integer, Y As Integer
    Dim J As Integer, desp As Integer
    
    X = XPosCartel + 20
    Y = YPosCartel + 60
    
    Call DDrawTransGrhIndextoSurface(textura, XPosCartel, YPosCartel, 0, LightRGB_Default)
    
    For J = 0 To UBound(LeyendaFormateada)
        DrawText X, Y + desp, LeyendaFormateada(J), -1, 120
        desp = desp + (frmMain.Font.Size) + 5
    Next
    
End Sub

