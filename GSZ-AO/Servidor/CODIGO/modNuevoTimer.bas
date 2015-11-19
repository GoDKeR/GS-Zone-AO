Attribute VB_Name = "modNuevoTimer"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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

'
' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'

' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 07/10/2012 - ^[GS]^
'***************************************************

Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If getInterval(TActual, UserList(UserIndex).Counters.TimerLanzarSpell) >= Intervalos(eIntervalos.iPuedeAtacarConHechizos) Then  ' 0.13.5
    If Actualizar Then
        UserList(UserIndex).Counters.TimerLanzarSpell = TActual
    End If
    IntervaloPermiteLanzarSpell = True
Else
    IntervaloPermiteLanzarSpell = False
End If

End Function

Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 07/10/2012 - ^[GS]^
'***************************************************

Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If getInterval(TActual, UserList(UserIndex).Counters.TimerPuedeAtacar) >= Intervalos(eIntervalos.iPuedeAtacar) Then  ' 0.13.5
    If Actualizar Then
        UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
        UserList(UserIndex).Counters.TimerGolpeUsar = TActual
    End If
    IntervaloPermiteAtacar = True
Else
    IntervaloPermiteAtacar = False
End If
End Function

Public Function IntervaloPermiteGolpeUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'***************************************************
'Author: ZaMa
'Checks if the time that passed from the last hit is enough for the user to use a potion.
'Last Modification: 07/10/2012 - ^[GS]^
'***************************************************

Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If getInterval(TActual, UserList(UserIndex).Counters.TimerGolpeUsar) >= Intervalos(eIntervalos.iPuedeUsarPocion) Then ' 0.13.5
    If Actualizar Then
        UserList(UserIndex).Counters.TimerGolpeUsar = TActual
    End If
    IntervaloPermiteGolpeUsar = True
Else
    IntervaloPermiteGolpeUsar = False
End If
End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 07/10/2012 - ^[GS]^
'***************************************************
    Dim TActual As Long
    
    With UserList(UserIndex)
        If .Counters.TimerMagiaGolpe > .Counters.TimerLanzarSpell Then
            Exit Function
        End If
        
        TActual = GetTickCount() And &H7FFFFFFF
        
        If getInterval(TActual, .Counters.TimerLanzarSpell) >= Intervalos(eIntervalos.iComboMagiaGolpe) Then ' 0.13.5
            If Actualizar Then
                .Counters.TimerMagiaGolpe = TActual
                .Counters.TimerPuedeAtacar = TActual
                .Counters.TimerGolpeUsar = TActual
            End If
            IntervaloPermiteMagiaGolpe = True
        Else
            IntervaloPermiteMagiaGolpe = False
        End If
    End With
End Function

Public Function IntervaloPermiteGolpeMagia(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 07/10/2012 - ^[GS]^
'***************************************************

    Dim TActual As Long
    
    If UserList(UserIndex).Counters.TimerGolpeMagia > UserList(UserIndex).Counters.TimerPuedeAtacar Then
        Exit Function
    End If
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If getInterval(TActual, UserList(UserIndex).Counters.TimerPuedeAtacar) >= Intervalos(eIntervalos.iComboGolpeMagia) Then ' 0.13.5
        If Actualizar Then
            UserList(UserIndex).Counters.TimerGolpeMagia = TActual
            UserList(UserIndex).Counters.TimerLanzarSpell = TActual
        End If
        IntervaloPermiteGolpeMagia = True
    Else
        IntervaloPermiteGolpeMagia = False
    End If
End Function

' ATAQUE CUERPO A CUERPO
'Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'Dim TActual As Long
'
'TActual = GetTickCount() And &H7FFFFFFF''
'
'If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= Intervalos(eIntervalos.iPuedeAtacar) Then
'    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
'    IntervaloPermiteAtacar = True
'Else
'    IntervaloPermiteAtacar = False
'End If
'End Function

' TRABAJO
Public Function IntervaloPermiteTrabajar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 07/10/2012 - ^[GS]^
'***************************************************

    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If getInterval(TActual, UserList(UserIndex).Counters.TimerPuedeTrabajar) >= Intervalos(eIntervalos.iPuedeTrabajar) Then ' 0.13.5
        If Actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
        IntervaloPermiteTrabajar = True
    Else
        IntervaloPermiteTrabajar = False
    End If
End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 07/10/2012 - ^[GS]^
'***************************************************

    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If getInterval(TActual, UserList(UserIndex).Counters.TimerUsar) >= Intervalos(eIntervalos.iPuedeUsarItem) Then  ' 0.13.5
        If Actualizar Then
            UserList(UserIndex).Counters.TimerUsar = TActual
            'UserList(UserIndex).Counters.failedUsageAttempts = 0
        End If
        IntervaloPermiteUsar = True
    Else
        IntervaloPermiteUsar = False
        
        'UserList(UserIndex).Counters.failedUsageAttempts = UserList(UserIndex).Counters.failedUsageAttempts + 1
        
        'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
        'If UserList(UserIndex).Counters.failedUsageAttempts = 20 Then
            'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " kicked by the server por posible modificación de intervalos.", FontTypeNames.FONTTYPE_FIGHT))
            'Call CloseSocket(UserIndex)
        'End If
    End If

End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 07/10/2012 - ^[GS]^
'***************************************************

    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If getInterval(TActual, UserList(UserIndex).Counters.TimerPuedeUsarArco) >= Intervalos(eIntervalos.iPuedeAtacarConFlechas) Then  ' 0.13.5
        If Actualizar Then UserList(UserIndex).Counters.TimerPuedeUsarArco = TActual
        IntervaloPermiteUsarArcos = True
    Else
        IntervaloPermiteUsarArcos = False
    End If

End Function

Public Function IntervaloPermiteSerAtacado(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = False) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify by: 27/07/2012 - ^[GS]^
'**************************************************************
    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    With UserList(UserIndex)
        ' Inicializa el timer
        If Actualizar Then
            .Counters.TimerPuedeSerAtacado = TActual
            .flags.NoPuedeSerAtacado = True
            IntervaloPermiteSerAtacado = False
        Else
            If getInterval(TActual, .Counters.TimerPuedeSerAtacado) >= IntervaloPuedeSerAtacado Then ' 0.13.5
                .flags.NoPuedeSerAtacado = False
                IntervaloPermiteSerAtacado = True
            Else
                IntervaloPermiteSerAtacado = False
            End If
        End If
    End With

End Function

Public Function IntervaloPerdioNpc(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = False) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify by: 27/07/2012 - ^[GS]^
'**************************************************************
    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    With UserList(UserIndex)
        ' Inicializa el timer
        If Actualizar Then
            .Counters.TimerPerteneceNpc = TActual
            IntervaloPerdioNpc = False
        Else
            If getInterval(TActual, .Counters.TimerPerteneceNpc) >= IntervaloOwnedNpc Then ' 0.13.5
                IntervaloPerdioNpc = True
            Else
                IntervaloPerdioNpc = False
            End If
        End If
    End With

End Function

Public Function IntervaloEstadoAtacable(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = False) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify by: 27/07/2012 - ^[GS]^
'**************************************************************
    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    With UserList(UserIndex)
        ' Inicializa el timer
        If Actualizar Then
            .Counters.TimerEstadoAtacable = TActual
            IntervaloEstadoAtacable = True
        Else
            If getInterval(TActual, .Counters.TimerEstadoAtacable) >= IntervaloAtacable Then ' 0.13.5
                IntervaloEstadoAtacable = False
            Else
                IntervaloEstadoAtacable = True
            End If
        End If
    End With

End Function

Public Function IntervaloGoHome(ByVal UserIndex As Integer, Optional ByVal TimeInterval As Long, Optional ByVal Actualizar As Boolean = False) As Boolean ' 0.13.3
'**************************************************************
'Author: ZaMa
'Last Modification: 10/08/2011 - ^[GS]^
'01/06/2010: ZaMa - Add the Timer which determines wether the user can be teleported to its home or not
'**************************************************************
    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    With UserList(UserIndex)
        ' Inicializa el timer
        If Actualizar Then
            .flags.Traveling = 1
            .Counters.goHome = TActual + TimeInterval
        Else
            If TActual >= .Counters.goHome Then
                IntervaloGoHome = True
            End If
        End If
    End With

End Function

Public Function checkInterval(ByRef startTime As Long, ByVal timeNow As Long, ByVal interval As Long) As Boolean ' 0.13.3
    Dim lInterval As Long
    
    If timeNow < startTime Then
        lInterval = &H7FFFFFFF - startTime + timeNow + 1
    Else
        lInterval = timeNow - startTime
    End If
    
    If lInterval >= interval Then
        startTime = timeNow
        checkInterval = True
    Else
        checkInterval = False
    End If
End Function

Public Function getInterval(ByVal timeNow As Long, ByVal startTime As Long) As Long ' 0.13.5
    If timeNow < startTime Then
        getInterval = &H7FFFFFFF - startTime + timeNow + 1
    Else
        getInterval = timeNow - startTime
    End If
End Function
