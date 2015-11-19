Attribute VB_Name = "modFadeEffects"
' ########################################
' Mod Fade versión BETA1.0a+
' Este Módulo fue programado por Standelf para diversos usos relacionados o no con el Argentum Online
' ##############################
' Explicación:   Este módulo es multiuso, ya que con el podremos _
                    crear tanto un FadeIn, como un FadeOut, o un Fade Continuo, _
                    a su vez, este se puede utilizar como cuenta regresiva o como contador.
' ##############################
' ¡¡¡¡¡ ATENCIÓN !!!!!
'   Este módulo fue pensado para aplicaciones con un búcle principal, _
    en caso de no tener tal bucle, podrá tambien ser utilizado en un For, un Do, o un Timer, _
    o a su libre gusto e imaginación.
' ##############################
' Funciones incluidas:
' Ver detalle sobre cada Función
' ########################################

' #### Enum eFadeMode, Utilizado para facilitar la compresión al crear un Fade
Public Enum eFadeMode
    ' #### Fade de: aparición, contador.
    FadeIn = 1
    ' #### Fade de: cuenta regresiva, desaparición.
    FadeOut = 2
    ' #### Fade infinito, de aparición y desaparición contínua.
    LoopFade = 3
End Enum

' #### Type tFade Estructura, Estructura principal de cada Fade
Public Type tFade
    active As Boolean
    
    MaxVal As Integer
    MinVal As Integer

    ActVal As Byte
    
    Interval As Integer
    Mode As eFadeMode
    DeleteOnFinish As Boolean

    LastCheck As Long
    Substract As Boolean
    
    Static As Boolean 'GSZAO
End Type

' #### Función GetTickCount
'Private Declare Function GetTickCount Lib "kernel32" () As Long ' Ya se encuentra globalmente

' #### Lista y Numero de Fade
Public FadeList() As tFade
Private TotalFade As Byte
Private i As Long

' ########################################
' Fade_Initializate, función pública que será utilizada a la hora de inicializar el módulo Fade.
' Return Byte: N/A
' Argumentos: N/A
' Ultima modificación: 10/02/2013
' ########################################

Public Function Fade_Initializate()
    ReDim FadeList(1 To 20) As tFade

    With FadeList(1)
        .active = True
        .Static = True 'Seteamos static para que no pueda ser reutilizado
        
        .MinVal = 1
        .MaxVal = 255
        .Interval = 10
        .Mode = eFadeMode.FadeIn
        .ActVal = 1
        .DeleteOnFinish = False
    End With
    
End Function

' ########################################
' Fade_Create, función pública que será utilizada a la hora de crear un nuevo Fade.
' Return Byte: Devuelve el ID del Fade Creado.
' Argumentos:   MinVal, Valor integer que determina el valor mínimo del fade.
'                       MaxVal, Valor integer que determina el valor máximo del fade.
'                       Mode, Valor eFadeMode, determina el cíclo que realiza el fade.
'                       Interval, Valor integer OPCIONAL, determina el intervalo en milisegundos en el que el fade se actualizará, por defecto serán 25ms.
'                       DeteleOnFinish, Valor booleano OPCIONAL, que determina si el fade será desactivado al finalizar su ciclo, por defecto será FALSO, El valor de esta variable no afectará si el Mode es LoopFade
' Ultima modificación: 19/01/2013
' ########################################

Public Function Fade_Create(ByVal MinVal As Integer, MaxVal As Integer, Mode As eFadeMode, Optional ByVal Interval As Integer = 25, Optional ByVal DeteleOnFinish As Boolean = False) As Byte

    Dim tmpID As Byte
            tmpID = Fade_FindFree()
            
    With FadeList(tmpID)
        .active = True
        .MinVal = MinVal
        .MaxVal = MaxVal
        .Interval = Interval
        .Mode = Mode
        .Static = False
        Select Case .Mode
            Case eFadeMode.FadeIn, eFadeMode.LoopFade
                .ActVal = .MinVal
            Case eFadeMode.FadeOut
                .ActVal = .MaxVal
        End Select
        
        .DeleteOnFinish = DeteleOnFinish
        
    End With
    
    Fade_Create = tmpID
    
End Function

' ########################################
' Fade_FindFree, función privada utilizada para obtener un ID
' vacio, o crear un nuevo ID.
' Return Byte: Devuelve el ID vacio que será utilizado al crear un Fade.
' Argumentos: N/A
' Ultima modificación: 19/01/2013
' ########################################

Private Function Fade_FindFree() As Byte
    
    'Esto no haria falta, por que nunca estaria en 0
    'If TotalFade = 0 Then
    '    ReDim FadeList(1 To 1) As tFade
    '    Fade_FindFree = 1
    '    TotalFade = 1
    '    Exit Function
    'End If

    For i = 1 To TotalFade
        If FadeList(i).active = False And FadeList(i).Static = False Then
            Fade_FindFree = i
            Exit Function
        End If
    Next i
    
    TotalFade = TotalFade + 1
    ReDim Preserve FadeList(1 To TotalFade) As tFade
    Fade_FindFree = TotalFade
    
    Exit Function
    
End Function


' ########################################
' Fade_UpdateAll, función pública que actualiza el valor de todos los fade juntos.
' Return Byte: N/A
' Argumentos: N/A
' Ultima modificación: 19/01/2013
' ########################################

Public Function Fade_UpdateAll()
    
    If TotalFade = 0 Then Exit Function
    
    Dim Time As Long
    Time = GetTickCount
    
    For i = 1 To TotalFade
        If FadeList(i).active = True Then
            With FadeList(i)
                If Time - .LastCheck > .Interval Then
                    .LastCheck = Time
                    Select Case .Mode
                        
                        Case eFadeMode.FadeIn
                            If .ActVal <> .MaxVal Then
                                .ActVal = .ActVal + 1
                            End If
                            If .DeleteOnFinish And .ActVal = .MaxVal Then .active = False
                            
                        Case eFadeMode.FadeOut
                            If .ActVal <> .MinVal Then
                                .ActVal = .ActVal - 1
                            End If
                            If .DeleteOnFinish And .ActVal = .MinVal Then .active = False
                            
                        Case eFadeMode.LoopFade
                            If .Substract Then
                                .ActVal = .ActVal - 1
                                If .ActVal = .MinVal Then .Substract = False
                            Else
                                .ActVal = .ActVal + 1
                                If .ActVal = .MaxVal Then .Substract = True
                            End If
                        
                    End Select
                End If
            End With
        End If
    Next i
        
End Function

' ########################################
' Fade_UpdateByID, función pública que actualiza el valor de fade correspondiente al ID seleccionado
' Return Byte: N/A
' Argumentos: ID, Valor byte que determina el número identificador del fade a actualizar.
' Ultima modificación: 19/01/2013
' ########################################

Public Function Fade_UpdateByID(ByVal id As Byte)
    If TotalFade = 0 Or id < 0 Or id > TotalFade Then Exit Function

    Dim Time As Long
    Time = GetTickCount
    
    If FadeList(id).active = True Then
            
        With FadeList(id)
            
            If Time - .LastCheck > .Interval Then
                        
                .LastCheck = Time
                        
                Select Case .Mode
                        
                    Case eFadeMode.FadeIn
                        If .ActVal <> .MaxVal Then
                            .ActVal = .ActVal + 1
                        End If
                                
                        If .DeleteOnFinish And .ActVal = .MaxVal Then .active = False
                            
                    Case eFadeMode.FadeOut
                        If .ActVal <> .MinVal Then
                            .ActVal = .ActVal - 1
                        End If
                            
                        If .DeleteOnFinish And .ActVal = .MinVal Then .active = False
                            
                    Case eFadeMode.LoopFade
                        If .Substract Then
                            .ActVal = .ActVal - 1
                            If .ActVal = .MinVal Then .Substract = False
                        Else
                            .ActVal = .ActVal + 1
                            If .ActVal = .MaxVal Then .Substract = True
                        End If
                        
                End Select
                        
            End If
                
        End With
            
    End If
End Function

