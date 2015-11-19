Attribute VB_Name = "modEventoTorneo"
'---------------------------------------------------------------------------------------
' Módulo    : modEventoTorneo
' Autor     : Facundo Ortega (GoDKeR)
' Fecha     : 16/02/2014
' Propósito : Torneo del tipo DeathMatch para Usuarios :)
'---------------------------------------------------------------------------------------

'Mi idea, es tener un enum publico en declaraciones, que indique que cada char si esta en evento o no, de esta forma, evitamos _
bugs, por ejemplo: esta en torneo y que pueda entrar a duelo.

Option Explicit

Private Enum eEstado
        Libre = 0
        Ocupada = 1
        Esperando = 2
End Enum

Private Type tFlags
        
        estado      As eEstado  'Estado de la arena
        CaenItems   As Byte
        'ClaseProhibida as byte 'Sirve?
End Type

Private Type tEvento

        UserIndex() As Integer  'Array de usuarios, evitamos sumonearlos a todos recorriendo 1 to LastUser ;)
        userPos()   As WorldPos 'Guardamos la posicion en la que estaba cuando entro al evento.
        
        cantUsers   As Byte     'Que se te caiga el sv si haces mas de 255 participantes >:(
        
        flags       As tFlags
End Type

Public Torneo   As tEvento

Private Const MAP_EVENT As Byte = 50 '??
Private Const X_EVENT   As Byte = 50
Private Const Y_EVENT   As Byte = 50

Private UsuariosCount   As Byte

Private Const SEPARATOR As String * 1 = vbNullChar


Public Function getNombresParticipantes() As String
    'Godki
    
    Dim tmpStr  As String
    Dim i       As Long
    
    If Torneo.flags.estado = Ocupada Or Torneo.flags.estado = Esperando Then
    
        For i = 1 To UBound(Torneo.UserIndex)
            
            If Torneo.UserIndex(i) <> 0 Then
            
                tmpStr = tmpStr & UserList(Torneo.UserIndex(i)).Name & SEPARATOR
            Else
                tmpStr = tmpStr & "Vacío." & SEPARATOR
            End If
            
        Next
        
        getNombresParticipantes = UCase$(tmpStr)
    Else
        
        getNombresParticipantes = "No hay torneos en este momento."
        
    End If
    
End Function

Public Sub PedirInfoTorneo(ByVal UserIndex As Integer)

        If EsGm(UserIndex) Then
                
                Call WritePedirInfo(UserIndex, 1)
        Else
                Call WritePedirInfo(UserIndex, 2)
        End If
        
End Sub
Public Sub AbrirTorneo(ByVal UserIndex As Integer, ByVal Participantes As Byte, ByVal Items As Byte)
        
        Dim strName As String
        
        If Not EsGm(UserIndex) Then Exit Sub 'Checkear esto pls
        
        If Not Torneo.flags.estado = Libre Then
                Call WriteConsoleMsg(UserIndex, "La sala de torneos se encuentra ocupada en este momento.", FontTypeNames.FONTTYPE_INFO)
                
                Exit Sub
        End If

        ReDim Torneo.UserIndex(1 To Participantes) As Integer
        ReDim Torneo.userPos(1 To Participantes) As WorldPos
        
        Torneo.cantUsers = Participantes
        
        Torneo.flags.estado = Esperando
        Torneo.flags.CaenItems = Items 'TERMINAR
        UsuariosCount = 1
        
        strName = UserList(UserIndex).Name
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg( _
        "El gm " & strName & " está organizando un torneo para: " & Participantes & ", para participar escribe /Participar.", _
        FontTypeNames.FONTTYPE_DIOS))
        
        'Comienza el timer, 5 minutos para inscribirse
End Sub

Public Sub ParticiparTorneo(ByVal UserIndex As Integer)
        
        
        If EsGm(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Vos no podes participar, cara rota >:c.", FontTypeNames.FONTTYPE_WARNING)
            
            Exit Sub
        End If
        
        If Not Torneo.flags.estado = Esperando Then
            Call WriteConsoleMsg(UserIndex, "No te puedes inscribir ahora!", FontTypeNames.FONTTYPE_WARNING)
            
            Exit Sub
        End If
        
        If UserList(UserIndex).Disp = Ocupado Then
            Call WriteConsoleMsg(UserIndex, "No puedes entrar al torneo si estas en algun otro evento.", FontTypeNames.FONTTYPE_WARNING)
            
            Exit Sub
        End If
        
        Dim i As Long
        
        For i = 1 To UBound(Torneo.UserIndex) 'No se deberia llegar jamás a esta comprobación, lo que se hace arriba deberia ser suficiente...
                If UserIndex = Torneo.UserIndex(i) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes entrar al torneo, ya estas participando.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub
                End If
        Next
        
        If Torneo.cantUsers > UsuariosCount Then
        
            Torneo.UserIndex(UsuariosCount) = UserIndex
            
            Torneo.userPos(UsuariosCount).Map = UserList(UserIndex).Pos.Map
            Torneo.userPos(UsuariosCount).X = UserList(UserIndex).Pos.X
            Torneo.userPos(UsuariosCount).Y = UserList(UserIndex).Pos.Y
            
            UsuariosCount = UsuariosCount + 1
            
            Call WriteConsoleMsg(UserIndex, "Has entrado al torneo!", FontTypeNames.FONTTYPE_INFO)
            
            UserList(UserIndex).Disp = Ocupado
            
            If UsuariosCount = Torneo.cantUsers Then
                Torneo.flags.estado = Ocupada
                TraerUsuarios
            End If
            
        Else
            Call WriteConsoleMsg(UserIndex, "Lo siento, el cupo se ha llenado.", FontTypeNames.FONTTYPE_WARNING)
        End If
        
End Sub

Private Sub TraerUsuarios()
        
        Dim i As Long
        
        'If Not EsGm(UserIndex) Then Exit Sub
        
        If Torneo.flags.estado = Ocupada Then
        
            If UsuariosCount = Torneo.cantUsers Then
            
                For i = 1 To UBound(Torneo.UserIndex)
                    Call WarpUserChar(Torneo.UserIndex(i), MAP_EVENT, X_EVENT, Y_EVENT, True)
                Next
            'Else
             '   Call WriteConsoleMsg(UserIndex, "El cupo aún no esta lleno.", FontTypeNames.FONTTYPE_WARNING)
            End If
        'Else
        '    Call WriteConsoleMsg(UserIndex, "No hay ningun torneo activo.", FontTypeNames.FONTTYPE_WARNING)
        End If
End Sub

Public Sub UserDieAtTorneo(ByVal UserIndex As Integer)
        
        If UserList(UserIndex).Disp = Ocupado Then
            If UserList(UserIndex).Pos.Map = MAP_EVENT Then
                
                If getIndexEnTorneo(UserIndex) <> 0 Then 'esto nunca jamas deberia pasar
                        
                        If UsuariosCount > 1 Then
                            UsuariosCount = UsuariosCount - 1
                            
                            Torneo.UserIndex(getIndexEnTorneo(UserIndex)) = 0
                        Else
                            DarGanador
                        End If
                        
                        Call WarpUserChar(UserIndex, Torneo.userPos(getIndexEnTorneo(UserIndex)).Map, _
                                        Torneo.userPos(getIndexEnTorneo(UserIndex)).X, _
                                        Torneo.userPos(getIndexEnTorneo(UserIndex)).Y, True)
                End If
            End If
        End If
        
End Sub

Private Sub DarGanador()
Dim i As Long 'recorremos el array para ver el unico indice que no quedo en 0
    
        For i = 1 To Torneo.cantUsers
                If Torneo.UserIndex(i) <> 0 Then
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El usuario: " & UserList(Torneo.UserIndex(i)).Name & ",ganó el torneo!", FontTypeNames.FONTTYPE_GMMSG))
                        
                        Call WarpUserChar(getIndexEnTorneo(i), Torneo.userPos(getIndexEnTorneo(i)).Map, _
                                Torneo.userPos(getIndexEnTorneo(i)).X, _
                                Torneo.userPos(getIndexEnTorneo(i)).Y, True)
                        'DAMOS PREMIOS :D
                End If
        Next
        
End Sub
Public Sub CloseSocketAtTorneo(ByVal UserIndex As Integer)
        
        If UserList(UserIndex).Disp = Ocupado Then
            If UserList(UserIndex).Pos.Map = MAP_EVENT Then
                
                If getIndexEnTorneo(UserIndex) <> 0 Then 'esto nunca jamas deberia pasar
                        
                        If UsuariosCount > 1 Then
                                UsuariosCount = UsuariosCount - 1
                                Torneo.UserIndex(getIndexEnTorneo(UserIndex)) = 0
                        Else
                            DarGanador
                        End If
                        
                        Call WarpUserChar(UserIndex, Torneo.userPos(getIndexEnTorneo(UserIndex)).Map, _
                                        Torneo.userPos(getIndexEnTorneo(UserIndex)).X, _
                                        Torneo.userPos(getIndexEnTorneo(UserIndex)).Y, True)
                End If
            End If
        End If
        
End Sub


Private Function getIndexEnTorneo(ByVal UserIndex As Integer) As Integer
Dim i As Long
        
        For i = 1 To Torneo.cantUsers
                
                If UserIndex = Torneo.UserIndex(i) Then
                    
                    getIndexEnTorneo = i
                    
                    Exit Function
                End If
        Next
        
getIndexEnTorneo = 0
End Function
