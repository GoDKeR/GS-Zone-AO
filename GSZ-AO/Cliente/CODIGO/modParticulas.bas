Attribute VB_Name = "modParticulas"
'/////////////////////////////Motor Grafico en DirectX 8///////////////////////////////
'////////////////////////Extraccion de varios motores por ShaFTeR//////////////////////
'///////////////////ORE - VBGORE - GSZAO - KKAO y algunos ejemplos de webs/////////////
'**************************************************************************************
Option Explicit

'Texture for particle effects - this is handled differently then the rest of the graphics


Private Type effect

        X As Single                 'Location of effect
        Y As Single
        GoToX As Single             'Location to move to
        GoToY As Single
        KillWhenAtTarget As Boolean     'If the effect is at its target (GoToX/Y), then Progression is set to 0
        KillWhenTargetLost As Boolean   'Kill the effect if the target is lost (sets progression = 0)
        Gfx As Byte                 'Particle texture used
        Used As Boolean             'If the effect is in use
        EffectNum As Byte           'What number of effect that is used
        Modifier As Integer         'Misc variable (depends on the effect)
        FloatSize As Long           'The size of the particles
        Direction As Integer        'Misc variable (depends on the effect)
        Particles() As clsParticle     'Information on each particle
        Progression As Single       'Progression state, best to design where 0 = effect ends
        PartVertex() As TLVERTEX    'Used to point render particles
        PreviousFrame As Long       'Tick time of the last frame
        ParticleCount As Integer    'Number of particles total
        ParticlesLeft As Integer    'Number of particles left - only for non-repetitive effects
        BindToChar As Integer       'Setting this value will bind the effect to move towards the character
        BindSpeed As Single         'How fast the effect moves towards the character
        BoundToMap As Byte          'If the effect is bound to the map or not (used only by the map editor)
        r As Single
        g As Single
        b As Single
        EcuationCount As Byte
        Sng As Single 'Misc variable

End Type

Public NumEffects                       As Byte   'Maximum number of effects at once
Public effect()                         As effect   'List of all the active effects

'Constants With The Order Number For Each Effect

Public Const EffectNum_Fire             As Byte = 1         'Burn baby, burn! Flame from a central point that blows in a specified direction
Public Const EffectNum_Snow             As Byte = 2         'Snow that covers the screen - weather effect
Public Const EffectNum_Heal             As Byte = 3         'Healing effect that can bind to a character, ankhs float up and fade
Public Const EffectNum_Bless            As Byte = 4         'Following three effects are same: create a circle around the central point
Public Const EffectNum_Protection       As Byte = 5         ' (often the character) and makes the given particle on the perimeter
Public Const EffectNum_Strengthen       As Byte = 6         ' which float up and fade out
Public Const EffectNum_Rain             As Byte = 7         'Exact same as snow, but moves much faster and more alpha value - weather effect
Public Const EffectNum_EquationTemplate As Byte = 8         'Template for creating particle effects through equations - a page with some equations can be found here: http://www.vbgore.com/modules.php?name=Forums&file=viewtopic&t=221
Public Const EffectNum_Waterfall        As Byte = 9         'Waterfall effect
Public Const EffectNum_Summon           As Byte = 10        'Summon effect
Public Const EffectNum_Meditate         As Byte = 11        'Medit effect
Public Const EffectNum_Portal           As Byte = 12        'Portal effect
Public Const EffectNum_Atomic           As Byte = 13        'Atomic Effect
Public Const EffectNum_Circle           As Byte = 14        'Outlined Circle Effect
Public Const EffectNum_Raro             As Byte = 15
Public Const EffectNum_Lissajous        As Byte = 16
Public Const EffectNum_Apocalipsis      As Byte = 17
Public Const EffectNum_Humo             As Byte = 18
Public Const EffectNum_CherryBlossom    As Byte = 19
Public Const EffectNum_BloodSpray       As Byte = 20
Public Const EffectNum_BloodSplatter    As Byte = 21
Public Const EffectNum_LevelUP          As Byte = 22        'Level Up Effect
Public Const EffectNum_AnimatedSign     As Byte = 23
Public Const EffectNum_Galaxy           As Byte = 24
Public Const EffectNum_FancyThickCircle As Byte = 25
Public Const EffectNum_Flower           As Byte = 26
Public Const EffectNum_Wormhole         As Byte = 27
Public Const EffectNum_HouseTeleport    As Byte = 28        'Teleport To House Effect
Public Const EffectNum_GuildTeleport    As Byte = 29        'Teleport To Guild Meeting
Public Const EffectNum_Rayo             As Byte = 30        'Tormenta de Fuego
Public Const EffectNum_LissajousMedit   As Byte = 31
Public Const EffectNum_Inmovilizar      As Byte = 32
Public Const EffectNum_ChangeClass      As Byte = 33
Public Const EffectNum_Armada           As Byte = 34        'Particula armada ultimo rango
Public Const EffectNum_ButterflyCurve   As Byte = 35

Sub Engine_Init_ParticleEngine(Optional ByVal SkipToTextures As Boolean = False)
'*****************************************************************
'Loads all particles into memory - unlike normal textures, these stay in memory. This isn't
'done for any reason in particular, they just use so little memory since they are so small
'*****************************************************************
Dim i As Byte
 
    If Not SkipToTextures Then
   
        'Set the particles texture
        NumEffects = 35 'ClientConfigInit.NumParticulas ' GSZAO
        ReDim effect(1 To NumEffects)
   
    End If
   
    For i = 1 To UBound(ParticleTexture())
        If Not ParticleTexture(i) Is Nothing Then Set ParticleTexture(i) = Nothing
        Set ParticleTexture(i) = DirectD3D8.CreateTextureFromFileEx(DirectDevice, DirGraficos & ClientConfigInit.DirParticulas & "\p" & i & ".png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
    Next i
 
End Sub

Function Effect_EquationTemplate_Begin(ByVal X As Single, _
                                       ByVal Y As Single, _
                                       ByVal Gfx As Integer, _
                                       ByVal Particles As Integer, _
                                       Optional ByVal Progression As Single = 1) As Integer

        '*****************************************************************
        'Particle effect template for effects as described on the
        'wiki page: http://www.vbgore.com/Particle_effect_equations
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_EquationTemplate_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_EquationTemplate  'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True                     'Enable the effect
        effect(EffectIndex).X = X                           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx                       'Set the graphic
        effect(EffectIndex).Progression = Progression       'If we loop the effect
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_EquationTemplate_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_EquationTemplate_Reset(ByVal EffectIndex As Integer, _
                                          ByVal Index As Long)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
        '*****************************************************************

        Dim X As Single
        Dim Y As Single
        Dim r As Single

        effect(EffectIndex).Progression = effect(EffectIndex).Progression + 0.1
        r = (Index / 20) * Exp(Index / effect(EffectIndex).Progression Mod 3)
        X = r * Cos(Index)
        Y = r * Sin(Index)
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + X, effect(EffectIndex).Y + Y, 0, 0, 0, 0
        effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 1, 0.2 + (Rnd * 0.2)
End Sub

Private Sub Effect_EquationTemplate_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_EquationTemplate_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_Bless_Begin(ByVal X As Single, _
                            ByVal Y As Single, _
                            ByVal Gfx As Integer, _
                            ByVal Particles As Integer, _
                            Optional ByVal Size As Byte = 30, _
                            Optional ByVal Time As Single = 10) As Integer

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Bless_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Bless     'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True             'Enabled the effect
        effect(EffectIndex).X = X                   'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx               'Set the graphic
        effect(EffectIndex).Modifier = Size         'How large the circle is
        effect(EffectIndex).Progression = Time      'How long the effect will last
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Bless_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Bless_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Reset
        '*****************************************************************

        Dim a As Single
        Dim X As Single
        Dim Y As Single

        'Get the positions
        a = Rnd * 360 * DegreeToRadian
        X = effect(EffectIndex).X - (Sin(a) * effect(EffectIndex).Modifier)
        Y = effect(EffectIndex).Y + (Cos(a) * effect(EffectIndex).Modifier)
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
        effect(EffectIndex).Particles(Index).ResetColor 1, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
End Sub

Private Sub Effect_Bless_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Bless_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_Fire_Begin(ByVal X As Single, _
                           ByVal Y As Single, _
                           ByVal Gfx As Integer, _
                           ByVal Particles As Integer, _
                           Optional ByVal Direction As Integer = 180, _
                           Optional ByVal Progression As Single = 1) As Integer

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Fire_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Fire      'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True     'Enabled the effect
        effect(EffectIndex).X = X           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx       'Set the graphic
        effect(EffectIndex).Direction = Direction       'The direction the effect is animat
        effect(EffectIndex).Progression = Progression   'Loop the effect
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Fire_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Fire_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
        '*****************************************************************
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X - 10 + Rnd * 20, effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Cos((effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
        effect(EffectIndex).Particles(Index).ResetColor 1, 0.2, 0.2, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)
End Sub

Private Sub Effect_Fire_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Fire_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Private Function Effect_FToDW(ByVal F As Single) As Long

        '*****************************************************************
        'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_FToDW
        '*****************************************************************

        Dim buf As D3DXBuffer

        'Converts a single into a long (Float to DWORD)
        Set buf = DirectD3D8.CreateBuffer(4)
        DirectD3D8.BufferSetData buf, 0, 4, 1, F
        DirectD3D8.BufferGetData buf, 0, 4, 1, Effect_FToDW
End Function

Function Effect_Heal_Begin(ByVal X As Single, _
                           ByVal Y As Single, _
                           ByVal Gfx As Integer, _
                           ByVal Particles As Integer, _
                           Optional ByVal Progression As Single = 1) As Integer

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Heal_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Heal      'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True     'Enabled the effect
        effect(EffectIndex).X = X           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx       'Set the graphic
        effect(EffectIndex).Progression = Progression   'Loop the effect
        effect(EffectIndex).KillWhenAtTarget = True     'End the effect when it reaches the target (progression = 0)
        effect(EffectIndex).KillWhenTargetLost = True   'End the effect if the target is lost (progression = 0)
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Heal_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Heal_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Reset
        '*****************************************************************
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X - 10 + Rnd * 20, effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Cos((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), 0, 0
        effect(EffectIndex).Particles(Index).ResetColor 0.8, 0.2, 0.2, 0.6 + (Rnd * 0.2), 0.01 + (Rnd * 0.5)
End Sub

Private Sub Effect_Heal_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate the time difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression <> 0 Then
                                        'Reset the particle
                                        Effect_Heal_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Sub Effect_Kill(Optional ByVal EffectIndex As Integer = 1, _
                Optional ByVal KillAll As Boolean = False)

        '*****************************************************************
        'Kills (stops) a single effect or all effects
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Kill
        '*****************************************************************

        Dim loopC As Long

        'Check If To Kill All Effects

        If KillAll = True Then

                'Loop Through Every Effect

                For loopC = 1 To NumEffects
                        'Stop The Effect
                        effect(loopC).Used = False

                Next

        Else
                'Stop The Selected Effect
                effect(EffectIndex).Used = False
        End If

End Sub

Private Function Effect_NextOpenSlot() As Integer

        '*****************************************************************
        'Finds the next open effects index
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_NextOpenSlot
        '*****************************************************************

        Dim EffectIndex As Integer

        'Find The Next Open Effect Slot

        Do
                EffectIndex = EffectIndex + 1   'Check The Next Slot

                If EffectIndex > NumEffects Then    'Dont Go Over Maximum Amount
                        Effect_NextOpenSlot = -1

                        Exit Function

                End If

        Loop While effect(EffectIndex).Used = True    'Check Next If Effect Is In Use

        'Return the next open slot
        Effect_NextOpenSlot = EffectIndex
        'Clear the old information from the effect
          
        Erase effect(EffectIndex).Particles()
        Erase effect(EffectIndex).PartVertex()
          
        ZeroMemory effect(EffectIndex), LenB(effect(EffectIndex))
        effect(EffectIndex).GoToX = -30000
        effect(EffectIndex).GoToY = -30000
          
End Function

Function Effect_Protection_Begin(ByVal X As Single, _
                                 ByVal Y As Single, _
                                 ByVal Gfx As Integer, _
                                 ByVal Particles As Integer, _
                                 Optional ByVal Size As Byte = 30, _
                                 Optional ByVal Time As Single = 10) As Integer

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Protection_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Protection    'Set the effect number
        effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
        effect(EffectIndex).Used = True             'Enabled the effect
        effect(EffectIndex).X = X                   'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx               'Set the graphic
        effect(EffectIndex).Modifier = Size         'How large the circle is
        effect(EffectIndex).Progression = Time      'How long the effect will last
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Protection_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Protection_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Reset
        '*****************************************************************

        Dim a As Single
        Dim X As Single
        Dim Y As Single

        'Get the positions
        a = Rnd * 360 * DegreeToRadian
        X = effect(EffectIndex).X - (Sin(a) * effect(EffectIndex).Modifier)
        Y = effect(EffectIndex).Y + (Cos(a) * effect(EffectIndex).Modifier)
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
        effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
End Sub

Private Sub Effect_UpdateOffset(ByVal EffectIndex As Integer)
        '***************************************************
        'Update an effect's position if the screen has moved
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateOffset
        '***************************************************

        If UserCharIndex <> 0 Then
                If EffectIndex <> CharList(UserCharIndex).ParticleIndex Then
                        effect(EffectIndex).X = effect(EffectIndex).X + (LastOffsetX - ParticleOffsetX)
                        effect(EffectIndex).Y = effect(EffectIndex).Y + (LastOffsetY - ParticleOffsetY)
                End If
        End If

End Sub

Private Sub Effect_UpdateBinding(ByVal EffectIndex As Integer)

        '***************************************************
        'Updates the binding of a particle effect to a target, if
        'the effect is bound to a character
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateBinding
        '***************************************************

        Dim TargetI As Integer
        Dim TargetA As Single
        Dim RetNum  As Integer 'fao
        
        'Update position through character binding

        If effect(EffectIndex).BindToChar > 0 Then
                'Store the character index
                TargetI = effect(EffectIndex).BindToChar

                'Check for a valid binding index

                If TargetI > LastChar Then
                        effect(EffectIndex).BindToChar = 0

                        If effect(EffectIndex).KillWhenTargetLost Then
                                effect(EffectIndex).Progression = 0

                                Exit Sub

                        End If

                ElseIf CharList(TargetI).active = 0 Then
                        effect(EffectIndex).BindToChar = 0

                        If effect(EffectIndex).KillWhenTargetLost Then
                                effect(EffectIndex).Progression = 0

                                Exit Sub

                        End If

                Else
                        'Calculate the X and Y positions
                        effect(EffectIndex).GoToX = Engine_TPtoSPX(CharList(effect(EffectIndex).BindToChar).Pos.X) + 10
                        effect(EffectIndex).GoToY = Engine_TPtoSPY(CharList(effect(EffectIndex).BindToChar).Pos.Y)
                End If
        End If

        'Move to the new position if needed

        If effect(EffectIndex).GoToX > -30000 Or effect(EffectIndex).GoToY > -30000 Then
                If effect(EffectIndex).GoToX <> effect(EffectIndex).X Or effect(EffectIndex).GoToY <> effect(EffectIndex).Y Then
                        'Calculate the angle
                        TargetA = Engine_GetAngle(effect(EffectIndex).X, effect(EffectIndex).Y, effect(EffectIndex).GoToX, effect(EffectIndex).GoToY) + 180
                        'Update the position of the effect
                        effect(EffectIndex).X = effect(EffectIndex).X - Sin(TargetA * DegreeToRadian) * effect(EffectIndex).BindSpeed '* timerElapsedTime
                        effect(EffectIndex).Y = effect(EffectIndex).Y + Cos(TargetA * DegreeToRadian) * effect(EffectIndex).BindSpeed '* timerElapsedTime

                        'Check if the effect is close enough to the target to just stick it at the target

                        If effect(EffectIndex).GoToX > -30000 Then
                                If Abs(effect(EffectIndex).X - effect(EffectIndex).GoToX) < 6 Then effect(EffectIndex).X = effect(EffectIndex).GoToX
                        End If

                        If effect(EffectIndex).GoToY > -30000 Then
                                If Abs(effect(EffectIndex).Y - effect(EffectIndex).GoToY) < 6 Then effect(EffectIndex).Y = effect(EffectIndex).GoToY
                        End If

                        'Check if the position of the effect is equal to that of the target

                        If effect(EffectIndex).X = effect(EffectIndex).GoToX Then
                                If effect(EffectIndex).Y = effect(EffectIndex).GoToY Then

                                        'For some effects, if the position is reached, we want to end the effect

                                        If effect(EffectIndex).KillWhenAtTarget Then
                                            
                                                'Explosion

                                                If effect(EffectIndex).Progression <> 0 Then
                                                
                                                        If effect(EffectIndex).EffectNum = EffectNum_Rayo Then
                                                                RetNum = Effect_Inmovilizar_Begin(effect(EffectIndex).X - 16, effect(EffectIndex).Y - 16, 17, 100, , 5)
                                                                'elseif otros efectos
                                                        End If
                                                
                                                End If
                                
                                                If RetNum > 0 Then
                                                        effect(RetNum).BindToChar = effect(EffectIndex).BindToChar
                                                        effect(RetNum).BindSpeed = 10
                                                End If
                                                
                                                effect(EffectIndex).BindToChar = 0
                                                effect(EffectIndex).Progression = 0
                                                effect(EffectIndex).GoToX = effect(EffectIndex).X
                                                effect(EffectIndex).GoToY = effect(EffectIndex).Y
                                        End If

                                        Exit Sub    'The effect is at the right position, don't update

                                End If
                        End If
                End If
        End If

End Sub

Private Sub Effect_Protection_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Protection_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Public Sub Effect_Render(ByVal EffectIndex As Integer, _
                         Optional ByVal SetRenderStates As Boolean = True)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Render
        '*****************************************************************

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Set the render state for the size of the particle
    DirectDevice.SetRenderState D3DRS_POINTSIZE, effect(EffectIndex).FloatSize
        
    'Set the render state to point blitting
    If SetRenderStates Then DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    
    'Set the last texture to a random number to force the engine to reload the texture
    'LastTexture = -65489

    'Set the texture
    DirectDevice.SetTexture 0, ParticleTexture(effect(EffectIndex).Gfx)

    'Draw all the particles at once
    DirectDevice.DrawPrimitiveUP D3DPT_POINTLIST, effect(EffectIndex).ParticleCount, effect(EffectIndex).PartVertex(0), Len(effect(EffectIndex).PartVertex(0))
                    
    'Reset the render state back to normal
    If SetRenderStates Then DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub

Function Effect_Snow_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Snow_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Snow      'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True     'Enabled the effect
        effect(EffectIndex).Gfx = Gfx       'Set the graphic
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Snow_Reset EffectIndex, loopC, 1

        Next loopC

        'Set the initial time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Snow_Reset(ByVal EffectIndex As Integer, _
                              ByVal Index As Long, _
                              Optional ByVal FirstReset As Byte = 0)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Reset
        '*****************************************************************

        If FirstReset = 1 Then
                'The very first reset
                effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (frmMain.ScaleWidth + 400)), Rnd * (frmMain.ScaleHeight + 50), Rnd * 5, 5 + Rnd * 3, 0, 0
        Else
                'Any reset after first
                effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (frmMain.ScaleWidth + 400)), -15 - Rnd * 185, Rnd * 5, 5 + Rnd * 3, 0, 0

                If effect(EffectIndex).Particles(Index).sngX < -20 Then effect(EffectIndex).Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
                If effect(EffectIndex).Particles(Index).sngX > frmMain.ScaleWidth Then effect(EffectIndex).Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
                If effect(EffectIndex).Particles(Index).sngY > frmMain.ScaleHeight Then effect(EffectIndex).Particles(Index).sngX = Rnd * (frmMain.ScaleWidth + 50)
        End If

        'Set the color
        effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.8, 0
End Sub

Private Sub Effect_Snow_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate the time difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check if particle is in use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if to reset the particle

                        If effect(EffectIndex).Particles(loopC).sngX < -200 Then effect(EffectIndex).Particles(loopC).sngA = 0
                        If effect(EffectIndex).Particles(loopC).sngX > (frmMain.ScaleWidth + 200) Then effect(EffectIndex).Particles(loopC).sngA = 0
                        If effect(EffectIndex).Particles(loopC).sngY > (frmMain.ScaleHeight + 200) Then effect(EffectIndex).Particles(loopC).sngA = 0

                        'Time for a reset, baby!

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then
                                'Reset the particle
                                Effect_Snow_Reset EffectIndex, loopC
                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_Strengthen_Begin(ByVal X As Single, _
                                 ByVal Y As Single, _
                                 ByVal Gfx As Integer, _
                                 ByVal Particles As Integer, _
                                 Optional ByVal Size As Byte = 30, _
                                 Optional ByVal Time As Single = 10) As Integer

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Strengthen_Begin = EffectIndex
        'Set the effect's variables
        effect(EffectIndex).EffectNum = EffectNum_Strengthen    'Set the effect number
        effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
        effect(EffectIndex).Used = True             'Enabled the effect
        effect(EffectIndex).X = X                   'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx               'Set the graphic
        effect(EffectIndex).Modifier = Size         'How large the circle is
        effect(EffectIndex).Progression = Time      'How long the effect will last
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Strengthen_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Strengthen_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Reset
        '*****************************************************************

        Dim a As Single
        Dim X As Single
        Dim Y As Single

        'Get the positions
        a = Rnd * 360 * DegreeToRadian
        X = effect(EffectIndex).X - (Sin(a) * effect(EffectIndex).Modifier)
        Y = effect(EffectIndex).Y + (Cos(a) * effect(EffectIndex).Modifier)
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
        effect(EffectIndex).Particles(Index).ResetColor 0.2, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
End Sub

Private Sub Effect_Strengthen_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate the time difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check if particle is in use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update the particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Strengthen_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Sub Effect_UpdateAll()

        '*****************************************************************
        'Updates all of the effects and renders them
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateAll
        '*****************************************************************

        Dim loopC As Long

        'Make sure we have effects

        If NumEffects = 0 Then Exit Sub
        'Set the render state for the particle effects
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

        'Update every effect in use

        For loopC = 1 To NumEffects

                'Make sure the effect is in use

                If effect(loopC).Used Then
                        'Update the effect position if the screen has moved
                        Effect_UpdateOffset loopC
                        'Update the effect position if it is binded
                        Effect_UpdateBinding loopC

                        'Find out which effect is selected, then update it

                        If effect(loopC).EffectNum = EffectNum_Fire Then Effect_Fire_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Snow Then Effect_Snow_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Heal Then Effect_Heal_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Bless Then Effect_Bless_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Protection Then Effect_Protection_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Strengthen Then Effect_Strengthen_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Rain Then Effect_Rain_Update loopC
                        If effect(loopC).EffectNum = EffectNum_EquationTemplate Then Effect_EquationTemplate_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Waterfall Then Effect_Waterfall_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Summon Then Effect_Summon_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Meditate Then Effect_Meditate_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Portal Then Effect_Portal_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Atomic Then Effect_Atomic_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Circle Then Effect_Circle_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Raro Then Effect_Raro_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Lissajous Then Effect_Lissajous_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Apocalipsis Then Effect_Apocalipsis_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Humo Then Effect_Humo_Update loopC
                        If effect(loopC).EffectNum = EffectNum_CherryBlossom Then Effect_CherryBlossom_Update loopC
                        If effect(loopC).EffectNum = EffectNum_BloodSpray Then Effect_BloodSpray_Update loopC
                        If effect(loopC).EffectNum = EffectNum_BloodSplatter Then Effect_BloodSplatter_Update loopC
                        If effect(loopC).EffectNum = EffectNum_LevelUP Then Effect_Spawn_Update EffectNum_LevelUP, loopC
                        If effect(loopC).EffectNum = EffectNum_AnimatedSign Then Effect_Spawn_Update EffectNum_AnimatedSign, loopC
                        If effect(loopC).EffectNum = EffectNum_Galaxy Then Effect_Spawn_Update EffectNum_Galaxy, loopC
                        If effect(loopC).EffectNum = EffectNum_FancyThickCircle Then Effect_Spawn_Update EffectNum_FancyThickCircle, loopC
                        If effect(loopC).EffectNum = EffectNum_Flower Then Effect_Spawn_Update EffectNum_Flower, loopC
                        If effect(loopC).EffectNum = EffectNum_Wormhole Then Effect_Spawn_Update EffectNum_Wormhole, loopC
                        If effect(loopC).EffectNum = EffectNum_HouseTeleport Then Effect_Spawn_Update EffectNum_HouseTeleport, loopC
                        If effect(loopC).EffectNum = EffectNum_GuildTeleport Then Effect_Spawn_Update EffectNum_GuildTeleport, loopC
                        If effect(loopC).EffectNum = EffectNum_Rayo Then Effect_Rayo_Update loopC
                        If effect(loopC).EffectNum = EffectNum_LissajousMedit Then Effect_LissajousMedit_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Inmovilizar Then Effect_Inmovilizar_Update loopC
                        If effect(loopC).EffectNum = EffectNum_ChangeClass Then Effect_ChangeClass_Update loopC
                        If effect(loopC).EffectNum = EffectNum_Armada Then Effect_Armada_Update loopC
                        If effect(loopC).EffectNum = EffectNum_ButterflyCurve Then Effect_Butterfly_Update loopC
                        
                        'Render the effect
                        Effect_Render loopC, False
                End If

        Next

        'Set the render state back for normal rendering
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub

Function Effect_Rain_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Rain_Begin = EffectIndex
        'Set the effect's variables
        effect(EffectIndex).EffectNum = EffectNum_Rain      'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True     'Enabled the effect
        effect(EffectIndex).Gfx = Gfx       'Set the graphic
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(10)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Rain_Reset EffectIndex, loopC, 1

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Rain_Reset(ByVal EffectIndex As Integer, _
                              ByVal Index As Long, _
                              Optional ByVal FirstReset As Byte = 0)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Reset
        '*****************************************************************

        If FirstReset = 1 Then
                'The very first reset
                effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (frmMain.ScaleWidth + 400)), Rnd * (frmMain.ScaleHeight + 50), Rnd * 5, 25 + Rnd * 12, 0, 0
        Else
                'Any reset after first
                effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * 1200), -15 - Rnd * 185, Rnd * 5, 25 + Rnd * 12, 0, 0

                If effect(EffectIndex).Particles(Index).sngX < -20 Then effect(EffectIndex).Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
                If effect(EffectIndex).Particles(Index).sngX > frmMain.ScaleWidth Then effect(EffectIndex).Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
                If effect(EffectIndex).Particles(Index).sngY > frmMain.ScaleHeight Then effect(EffectIndex).Particles(Index).sngX = Rnd * (frmMain.ScaleWidth + 50)
        End If

        'Set the color
        effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.4, 0
End Sub

Private Sub Effect_Rain_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate the time difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check if the particle is in use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update the particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if to reset the particle

                        If effect(EffectIndex).Particles(loopC).sngX < -200 Then effect(EffectIndex).Particles(loopC).sngA = 0
                        If effect(EffectIndex).Particles(loopC).sngX > (frmMain.ScaleWidth + 200) Then effect(EffectIndex).Particles(loopC).sngA = 0
                        If effect(EffectIndex).Particles(loopC).sngY > (frmMain.ScaleHeight + 200) Then effect(EffectIndex).Particles(loopC).sngA = 0

                        'Time for a reset, baby!

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then
                                'Reset the particle
                                Effect_Rain_Reset EffectIndex, loopC
                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Public Sub Effect_Begin(ByVal EffectIndex As Integer, _
                        ByVal X As Single, _
                        ByVal Y As Single, _
                        ByVal GfxIndex As Byte, _
                        ByVal Particles As Byte, _
                        Optional ByVal Direction As Single = 180, _
                        Optional ByVal BindToMap As Boolean = False)

        '*****************************************************************
        'A very simplistic form of initialization for particle effects
        'Should only be used for starting map-based effects
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Begin
        '*****************************************************************

        Dim RetNum As Byte

        Select Case EffectIndex

                Case 1
                        RetNum = Effect_Fire_Begin(X, Y, GfxIndex, Particles, Direction, 1)

                Case 4
                        RetNum = Effect_Waterfall_Begin(X, Y, GfxIndex, Particles)
                        
                Case 7
                        RetNum = Effect_Portal_Begin(X, Y, GfxIndex, Particles, 100)
        End Select

        'Bind the effect to the map if needed

        If BindToMap Then effect(RetNum).BoundToMap = 1
End Sub

Function Effect_Waterfall_Begin(ByVal X As Single, _
                                ByVal Y As Single, _
                                ByVal Gfx As Integer, _
                                ByVal Particles As Integer) As Integer

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Waterfall_Begin = EffectIndex
        'Set the effect's variables
        effect(EffectIndex).EffectNum = EffectNum_Waterfall     'Set the effect number
        effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
        effect(EffectIndex).Used = True             'Enabled the effect
        effect(EffectIndex).X = X                   'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx               'Set the graphic
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Waterfall_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Waterfall_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Reset
        '*****************************************************************

        If Int(Rnd * 10) = 1 Then
                effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + (Rnd * 60), effect(EffectIndex).Y + (Rnd * 130), 0, 8 + (Rnd * 6), 0, 0
        Else
                effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + (Rnd * 60), effect(EffectIndex).Y + (Rnd * 10), 0, 8 + (Rnd * 6), 0, 0
        End If

        effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0
End Sub

Private Sub Effect_Waterfall_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                With effect(EffectIndex).Particles(loopC)

                        'Check if the particle is in use

                        If .Used Then
                                'Update The Particle
                                .UpdateParticle ElapsedTime

                                'Check if the particle is ready to die

                                If (.sngY > effect(EffectIndex).Y + 140) Or (.sngA = 0) Then
                                        'Reset the particle
                                        Effect_Waterfall_Reset EffectIndex, loopC
                                Else
                                        'Set the particle information on the particle vertex
                                        effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                                        effect(EffectIndex).PartVertex(loopC).X = .sngX
                                        effect(EffectIndex).PartVertex(loopC).Y = .sngY
                                End If
                        End If

                End With

        Next loopC

End Sub

Function Effect_Summon_Begin(ByVal X As Single, _
                             ByVal Y As Single, _
                             ByVal Gfx As Integer, _
                             ByVal Particles As Integer, _
                             Optional ByVal Progression As Single = 0) As Integer

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Summon_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Summon    'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True                     'Enable the effect
        effect(EffectIndex).X = X                           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx                       'Set the graphic
        effect(EffectIndex).Progression = Progression       'If we loop the effect
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Summon_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Summon_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Reset
        '*****************************************************************

        Dim X As Single
        Dim Y As Single
        Dim r As Single

        If effect(EffectIndex).Progression > 1000 Then
                effect(EffectIndex).Progression = effect(EffectIndex).Progression + 1.4
        Else
                effect(EffectIndex).Progression = effect(EffectIndex).Progression + 0.5
        End If

        r = (Index / 30) * Exp(Index / effect(EffectIndex).Progression)
        X = r * Cos(Index)
        Y = r * Sin(Index)
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + X, effect(EffectIndex).Y + Y, 0, 0, 0, 0
        effect(EffectIndex).Particles(Index).ResetColor 0, Rnd, 0, 0.9, 0.2 + (Rnd * 0.2)
End Sub

Private Sub Effect_Summon_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression < 1800 Then
                                        'Reset the particle
                                        Effect_Summon_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_Meditate_Begin(ByVal X As Single, _
                               ByVal Y As Single, _
                               ByVal Gfx As Integer, _
                               ByVal Particles As Integer, _
                               Optional ByVal Size As Byte = 30, _
                               Optional ByVal Time As Single = 10) As Integer

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Partic ... tate_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Meditate_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Meditate     'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True             'Enabled the effect
        effect(EffectIndex).X = X                   'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx               'Set the graphic
        effect(EffectIndex).Modifier = Size         'How large the circle is
        effect(EffectIndex).Progression = Time      'How long the effect will last
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Meditate_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function
 
Private Sub Effect_Meditate_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Partic ... tate_Reset
        '*****************************************************************

        Dim a  As Single
        Dim X  As Single
        Dim Y  As Single
        Dim rR As Single
        Dim rG As Single
        Dim rB As Single

        'Get the positions
        a = Rnd * 360 * DegreeToRadian
        X = effect(EffectIndex).X - (Sin(a) * effect(EffectIndex).Modifier)
        Y = effect(EffectIndex).Y + (Cos(a) * effect(EffectIndex).Modifier / 2.5)
        'Load Colours
        rR = (0.1 - 0.05) * Rnd + 0.03
        rG = 0.8
        rB = 0.5
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
        effect(EffectIndex).Particles(Index).ResetColor rR, rG, rB, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
End Sub
 
Private Sub Effect_Meditate_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Partic ... ate_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Meditate_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_Portal_Begin(ByVal X As Single, _
                             ByVal Y As Single, _
                             ByVal Gfx As Integer, _
                             ByVal Particles As Integer, _
                             Optional ByVal Size As Byte = 30, _
                             Optional ByVal Time As Single = 10) As Long

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Partic ... rtal_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        
        'Return the index of the used slot
        Effect_Portal_Begin = EffectIndex
        
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Portal     'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True             'Enabled the effect
        effect(EffectIndex).X = X                   'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx               'Set the graphic
        effect(EffectIndex).Modifier = Size         'How large the circle is
        effect(EffectIndex).Progression = Time      'How long the effect will last
        
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
        
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Portal_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function
 
Private Sub Effect_Portal_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Partic ... rtal_Reset
        '*****************************************************************

        Dim a  As Single
        Dim X  As Single
        Dim Y  As Single
        Dim rR As Single
        Dim rG As Single
        Dim rB As Single

        'Get the positions
        a = Rnd * 360 * DegreeToRadian
        
        effect(EffectIndex).Progression = effect(EffectIndex).Progression + 0.14

        If Rnd > Rnd Then
                X = effect(EffectIndex).X - (Sin(a) * effect(EffectIndex).Modifier / 1.8) * Rnd * 1.1
                Y = effect(EffectIndex).Y + (Cos(a) * effect(EffectIndex).Modifier * 1.1) * Rnd * 1.1
                rR = (0.1 - 0.05) * Rnd + 0.03
                rG = 0.2
                rB = 0.8
        Else
                X = effect(EffectIndex).X - (Sin(a) * effect(EffectIndex).Modifier / 3)
                Y = effect(EffectIndex).Y + (Cos(a) * effect(EffectIndex).Modifier / 1.5)
                rR = (0.2 - 0.06) * Rnd + 0.04
                rG = 0.3
                rB = 0.2
        End If

        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, 0 '-2
        effect(EffectIndex).Particles(Index).ResetColor rR, rG, rB, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
End Sub
 
Private Sub Effect_Portal_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Partic ... tal_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Portal_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_Atomic_Begin(ByVal X As Single, _
                             ByVal Y As Single, _
                             ByVal Gfx As Integer, _
                             ByVal Particles As Integer, _
                             Optional ByVal Size As Byte = 30, _
                             Optional ByVal Time As Single = 10) As Integer

        '*****************************************************************
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Atomic_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Atomic        'Set the effect number
        effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
        effect(EffectIndex).Used = True             'Enabled the effect
        effect(EffectIndex).X = X                   'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx               'Set the graphic
        effect(EffectIndex).Modifier = Size         'How large the circle is
        effect(EffectIndex).Progression = Time      'How long the effect will last
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Atomic_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Atomic_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        '*****************************************************************

        Dim r As Single
        Dim X As Single
        Dim Y As Single

        'Get the positions
        r = 10 + Sin(2 * (Index / 10)) * 50
        X = r * Cos(Index / 30)
        Y = r * Sin(Index / 30)
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + X, effect(EffectIndex).Y + Y, 0, 0, 0, 0
        effect(EffectIndex).Particles(Index).ResetColor 200, 50, 1, 1, 0.9 + (Rnd * 0.2)
End Sub

Private Sub Effect_Atomic_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Atomic_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_Circle_Begin(ByVal X As Single, _
                             ByVal Y As Single, _
                             ByVal Gfx As Integer, _
                             ByVal Particles As Integer, _
                             Optional ByVal Size As Byte = 30, _
                             Optional ByVal Time As Single = 10) As Integer

        '*****************************************************************
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Circle_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Circle        'Set the effect number
        effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
        effect(EffectIndex).Used = True             'Enabled the effect
        effect(EffectIndex).X = X                   'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx               'Set the graphic
        effect(EffectIndex).Modifier = Size         'How large the circle is
        effect(EffectIndex).Progression = Time      'How long the effect will last
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Circle_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Circle_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        '*****************************************************************

        Dim a As Single
        Dim X As Single
        Dim Y As Single

        'Get the positions
        a = Rnd * 360 * DegreeToRadian 'The point on the circumference to be used
        X = effect(EffectIndex).X - (Sin(a) * 40) 'The 40s state the radius of circle
        Y = effect(EffectIndex).Y + (Cos(a) * 40)
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, -2
        effect(EffectIndex).Particles(Index).ResetColor 1 * Rnd + 0.4, 0, 1, 1, 0.2 + (Rnd * 0.2)
End Sub

Private Sub Effect_Circle_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Circle_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_Raro_Begin(ByVal X As Single, _
                           ByVal Y As Single, _
                           ByVal Gfx As Integer, _
                           ByVal Particles As Integer, _
                           Optional ByVal Size As Byte = 30, _
                           Optional ByVal Time As Single = 10) As Integer

        '*****************************************************************
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Raro_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Raro        'Set the effect number
        effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
        effect(EffectIndex).Used = True             'Enabled the effect
        effect(EffectIndex).X = X                   'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx               'Set the graphic
        effect(EffectIndex).Modifier = Size         'How large the circle is
        effect(EffectIndex).Progression = Time      'How long the effect will last
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Raro_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Raro_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        '*****************************************************************

        Dim X As Single
        Dim Y As Single
        Dim i As Single

        'Get the positions
        'a = Rnd * 360 * DegreeToRadian 'The point on the circumference to be used

        For i = 0 To 360 Step 30
                X = effect(EffectIndex).X - Cos(i)
                Y = effect(EffectIndex).Y + Sin(i) + Rnd
                'Reset the particle
                effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 1, 0.2 + (Rnd * 0.2)

        Next i

End Sub

Private Sub Effect_Raro_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Raro_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_Lissajous_Begin(ByVal X As Single, _
                                ByVal Y As Single, _
                                ByVal Gfx As Integer, _
                                ByVal Particles As Integer, _
                                Optional ByVal Progression As Single = 0, _
                                Optional Size As Byte = 30, _
                                Optional r As Single = 100, _
                                Optional g As Single = 100, _
                                Optional b As Single = 100) As Long

        '*****************************************************************
        'Particle effect Lissajous equation
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Lissajous_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Lissajous 'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True                     'Enable the effect
        effect(EffectIndex).X = X                           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx                       'Set the graphic
        effect(EffectIndex).Modifier = Size                 'How large the circle is
        effect(EffectIndex).Progression = Progression
        effect(EffectIndex).r = r
        effect(EffectIndex).g = g
        effect(EffectIndex).b = b
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Lissajous_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Lissajous_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
        '*****************************************************************

        Dim X As Single
        Dim Y As Single
        Dim a As Single

        '2
        '1
        '1
        '2
        effect(EffectIndex).Progression = effect(EffectIndex).Progression + 0.01
        a = effect(EffectIndex).Progression

        If RandomNumber(1, 2) = 1 Then
                X = effect(EffectIndex).X - (Sin(2 * a) * effect(EffectIndex).Modifier) - 20
                Y = effect(EffectIndex).Y + (Sin(1 * a) * effect(EffectIndex).Modifier)
                'Reset the particle
                effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                effect(EffectIndex).Particles(Index).ResetColor effect(EffectIndex).r * effect(EffectIndex).Progression, effect(EffectIndex).g * effect(EffectIndex).Progression, effect(EffectIndex).b, 0.2, 0.2 + (Rnd * 0.2)
        Else
                X = effect(EffectIndex).X - (Sin(1 + 5 * a) * effect(EffectIndex).Modifier) - 20
                Y = effect(EffectIndex).Y + (Sin(2 + 7 * a) * effect(EffectIndex).Modifier)
                'Reset the particle
                effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                effect(EffectIndex).Particles(Index).ResetColor effect(EffectIndex).r * effect(EffectIndex).Progression, effect(EffectIndex).g * effect(EffectIndex).Progression, effect(EffectIndex).b, 0.2, 0.2 + (Rnd * 0.2)
        End If

End Sub

Private Sub Effect_Lissajous_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Lissajous_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_Apocalipsis_Begin(ByVal X As Single, _
                                  ByVal Y As Single, _
                                  ByVal Gfx As Integer, _
                                  ByVal Particles As Integer, _
                                  Optional ByVal Progression As Single = 0) As Integer

        '*****************************************************************
        'Particle effect template for effects as described on the
        'wiki page: http://www.vbgore.com/Particle_effect_equations
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Apocalipsis_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Apocalipsis  'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True                     'Enable the effect
        effect(EffectIndex).X = X                           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx                       'Set the graphic
        effect(EffectIndex).Progression = Progression
        effect(EffectIndex).KillWhenAtTarget = True     'End the effect when it reaches the target (progression = 0)
        effect(EffectIndex).KillWhenTargetLost = True   'End the effect if the target is lost (progression = 0)
        
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Apocalipsis_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Apocalipsis_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
        '*****************************************************************

        Dim X As Single
        Dim Y As Single
        Dim a As Single

        effect(EffectIndex).Progression = effect(EffectIndex).Progression + 0.01
        a = effect(EffectIndex).Progression

        If RandomNumber(1, 2) = 1 Then
                X = effect(EffectIndex).X - (Sin(a)) * 120
                Y = effect(EffectIndex).Y + Cos(5 * a) * 20 'The 40s state the radius of circle
                effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                effect(EffectIndex).Particles(Index).ResetColor 5, 0, 3, 1, 0.2 + (Rnd * 0.2)
        Else
                X = effect(EffectIndex).X - (Sin(a)) * 120
                Y = effect(EffectIndex).Y - Cos(5 * a) * 20 'The 40s state the radius of circle
                '
                effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                effect(EffectIndex).Particles(Index).ResetColor 0, 5, 2, 1, 0.2 + (Rnd * 0.2)
        End If

End Sub

Private Sub Effect_Apocalipsis_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Apocalipsis_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_Humo_Begin(ByVal X As Single, _
                           ByVal Y As Single, _
                           ByVal Gfx As Integer, _
                           ByVal Particles As Integer, _
                           Optional ByVal Direction As Integer = 180, _
                           Optional ByVal Progression As Single = 1) As Integer

        '*****************************************************************
        'More info: http://svn2.assembla.com/svn/vblore/trunk/Code/Common%20Code/Particles.bas
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Humo_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Humo      'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True     'Enabled the effect
        effect(EffectIndex).X = X           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx       'Set the graphic
        effect(EffectIndex).Direction = Direction       'The direction the effect is animat
        effect(EffectIndex).Progression = Progression   'Loop the effect
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(30)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Humo_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Humo_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
        '*****************************************************************
        'More info: http://svn2.assembla.com/svn/vblore/trunk/Code/Common%20Code/Particles.bas
        '*****************************************************************
        'Reset the particle
        'Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
        'Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2, 0.2, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)
        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X - 10 + Rnd * 50, effect(EffectIndex).Y - 10 + Rnd * 50, -Sin((effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 5, Cos((effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0.5, 0
        effect(EffectIndex).Particles(Index).ResetColor 0.2, 0.2, 0.2, 0.2 + (Rnd * 0.2), 0.03 + (Rnd * 0.01)
        'Reset the particle
        'Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 50, Effect(EffectIndex).Y - 10 + Rnd * 80, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
        'Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.1, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)
        'Effect(EffectIndex).Particles(index).ResetColor 0.1, 0.1, 0.1, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)
End Sub

Private Sub Effect_Humo_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://svn2.assembla.com/svn/vblore/trunk/Code/Common%20Code/Particles.bas
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression <> 0 Then
                                        'Reset the particle
                                        Effect_Humo_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_CherryBlossom_Begin(ByVal X As Single, _
                                    ByVal Y As Single, _
                                    ByVal Gfx As Integer, _
                                    ByVal Particles As Integer) As Integer

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_CherryBlossom_Begin = EffectIndex
        'Set the effect's variables
        effect(EffectIndex).EffectNum = EffectNum_CherryBlossom     'Set the effect number
        effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
        effect(EffectIndex).Used = True             'Enabled the effect
        effect(EffectIndex).X = X                   'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx               'Set the graphic
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_CherryBlossom_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_CherryBlossom_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Reset
        '*****************************************************************

        If Int(Rnd * 10) = 1 Then
                effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + (Rnd * 60), effect(EffectIndex).Y + (Rnd * 130), 2 + (Rnd * 2), 2 + (Rnd * 2), 0, 0
        Else
                effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + (Rnd * 60), effect(EffectIndex).Y + (Rnd * 10), 2 + (Rnd * 2), 2 + (Rnd * 2), 0, 0
        End If

        effect(EffectIndex).Particles(Index).ResetColor 1#, 0.7, 0.75, 0.6 + (Rnd * 0.4), 0
End Sub

Private Sub Effect_CherryBlossom_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                With effect(EffectIndex).Particles(loopC)

                        'Check if the particle is in use

                        If .Used Then
                                'Update The Particle
                                .UpdateParticle ElapsedTime

                                'Check if the particle is ready to die

                                If (.sngY > effect(EffectIndex).Y + 140) Or (.sngA = 0) Then
                                        'Reset the particle
                                        Effect_CherryBlossom_Reset EffectIndex, loopC
                                Else
                                        'Set the particle information on the particle vertex
                                        effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                                        effect(EffectIndex).PartVertex(loopC).X = .sngX
                                        effect(EffectIndex).PartVertex(loopC).Y = .sngY
                                End If
                        End If

                End With

        Next loopC

End Sub

Function Effect_BloodSpray_Begin(ByVal X As Single, _
                                 ByVal Y As Single, _
                                 ByVal Particles As Integer, _
                                 ByVal Direction As Single, _
                                 Optional ByVal Intensity As Single = 1) As Integer

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_BloodSpray_Begin = EffectIndex
        'Set the effect's variables
        effect(EffectIndex).EffectNum = EffectNum_BloodSpray  'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True                     'Enable the effect
        effect(EffectIndex).X = X                           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
        effect(EffectIndex).Direction = Direction           'Direction
        effect(EffectIndex).Modifier = Intensity
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(7)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_BloodSpray_Reset EffectIndex, loopC

        Next loopC

        'Set the initial time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_BloodSpray_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        'Reset the particle

        With effect(EffectIndex)
                .Particles(Index).ResetIt .X + (Rnd * 16) - 8, .Y + (Rnd * 32) - 16, Sin((.Direction - 10 + (Rnd * 20)) * DegreeToRadian) * (30 * .Modifier * Rnd), -Cos((.Direction - 10 + (Rnd * 20)) * DegreeToRadian) * (30 * .Modifier * Rnd), 0, 0, -10, -2 - (Rnd * 30), 8 + Rnd * 4
                .Particles(Index).ResetColor 1, 1, 1, 0.8, 0
        End With

End Sub

Private Sub Effect_BloodSpray_Update(ByVal EffectIndex As Integer)

        Dim ElapsedTime As Single
        Dim loopC       As Long
        Dim TileX       As Long
        Dim TileY       As Long

        'Calculate the time difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                With effect(EffectIndex).Particles(loopC)

                        'Check if particle is in Use

                        If .Used Then
                                'Update the particle
                                .UpdateParticle ElapsedTime
                                'Don't pass any walls/etc
                                TileX = Engine_SPtoTPX(.sngX)
                                TileY = Engine_SPtoTPY(.sngY)

                                If TileX < 1 Then
                                        .sngZ = 1.1
                                ElseIf TileY < 1 Then
                                        .sngZ = 1.1
                                ElseIf TileX > 92 Then
                                        .sngZ = 1.1
                                ElseIf TileY > 92 Then
                                        .sngZ = 1.1
                                End If

                                If .sngZ <> 1.1 Then
                                        'If MapData(TileX, TileY).BlockedAttack Then
                                        '.sngZ = 1.1
                                        'End If
                                End If

                                'Blood trails

                                If loopC = 0 Or loopC Mod 15 = 0 Then
                                        If Int(Rnd * 3) = 0 Then
                                                If Int(Rnd * 2) = 0 Then
                                                        'Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY, 2
                                                Else
                                                        'Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY, 1
                                                End If
                                        End If
                                End If

                                'Check if to kill off the particle

                                If .sngZ > 1 Then
                                        'Disable the particle
                                        .Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).Particles(loopC).sngA = 0

                                        'Check if we lost all the particles

                                        If effect(EffectIndex).ParticlesLeft <= 0 Then effect(EffectIndex).Used = False
                                        'Create the blood splatter
                                        'Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY, 0
                                Else
                                        'Set the particle information on the particle vertex
                                        effect(EffectIndex).PartVertex(loopC).color = 1258291200
                                        effect(EffectIndex).PartVertex(loopC).X = .sngX
                                        effect(EffectIndex).PartVertex(loopC).Y = .sngY
                                End If
                        End If

                End With

        Next loopC

End Sub

Function Effect_BloodSplatter_Begin(ByVal X As Single, _
                                    ByVal Y As Single, _
                                    ByVal Particles As Integer) As Integer

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_BloodSplatter_Begin = EffectIndex
        'Set the effect's variables
        effect(EffectIndex).EffectNum = EffectNum_BloodSplatter  'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True                     'Enable the effect
        effect(EffectIndex).X = X                           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(7)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_BloodSplatter_Reset EffectIndex, loopC

        Next loopC

        'Set the initial time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_BloodSplatter_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        Dim Direction As Single

        'Find the direction
        Direction = Rnd * 360
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + (Rnd * 16) - 8, effect(EffectIndex).Y + (Rnd * 32) - 16, Sin(Direction * DegreeToRadian) * (24 * Rnd), -Cos(Direction * DegreeToRadian) * (24 * Rnd), 0, 0, -25, -3 - (Rnd * 40), 10 + Rnd * 4
        effect(EffectIndex).Particles(Index).ResetColor 1, 0, 0, 0.8, 0
End Sub

Private Sub Effect_BloodSplatter_Update(ByVal EffectIndex As Integer)

        Dim ElapsedTime As Single
        Dim loopC       As Long
        Dim TileX       As Long
        Dim TileY       As Long

        'Calculate the time difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                With effect(EffectIndex).Particles(loopC)

                        'Check if particle is in Use

                        If .Used Then
                                'Update the particle
                                .UpdateParticle ElapsedTime
                                'Don't pass any walls/etc
                                TileX = Engine_SPtoTPX(.sngX)
                                TileY = Engine_SPtoTPY(.sngY)

                                If TileX < 1 Then
                                        .sngZ = 1.1
                                ElseIf TileY < 1 Then
                                        .sngZ = 1.1
                                ElseIf TileY > 92 Then
                                        .sngZ = 1.1
                                ElseIf TileY > 92 Then
                                        .sngZ = 1.1
                                End If

                                If .sngZ <> 1.1 Then
                                        'If MapData(TileX, TileY).BlockedAttack Then
                                        '.sngZ = 1.1
                                        'End If
                                End If

                                'Blood trails

                                If loopC = 0 Or loopC Mod 10 = 0 Then
                                        If Int(Rnd * 3) = 0 Then
                                                If Int(Rnd * 2) = 0 Then
                                                       ' Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY, 2
                                                Else
                                                       ' Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY, 1
                                                End If
                                        End If
                                End If

                                'Check if to kill off the particle

                                If .sngZ > 1 Then
                                        'Disable the particle
                                        .Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).Particles(loopC).sngA = 0

                                        'Check if we lost all the particles

                                        If effect(EffectIndex).ParticlesLeft <= 0 Then effect(EffectIndex).Used = False
                                        'Create the blood splatter
                                        'Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY, 0
                                Else
                                        'Set the particle information on the particle vertex
                                        effect(EffectIndex).PartVertex(loopC).color = 1258291200
                                        effect(EffectIndex).PartVertex(loopC).X = .sngX
                                        effect(EffectIndex).PartVertex(loopC).Y = .sngY
                                End If
                        End If

                End With

        Next loopC

End Sub

Function Effect_Spawn_Begin(ByVal EffectNum As Byte, _
                            ByVal X As Single, _
                            ByVal Y As Single, _
                            ByVal Gfx As Integer, _
                            ByVal Particles As Integer, _
                            Optional ByVal Size As Byte = 30, _
                            Optional ByVal Time As Single = 10, _
                            Optional ByVal Red As Single = -1, _
                            Optional ByVal Green As Single = -1, _
                            Optional ByVal Blue As Single = -1, _
                            Optional ByVal Alpha As Single = -1) As Integer

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Spawn_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum     'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True             'Enabled the effect
        effect(EffectIndex).X = X                   'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx               'Set the graphic
        effect(EffectIndex).Modifier = Size         'How large the circle is
        effect(EffectIndex).Progression = Time      'How long the effect will last
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Spawn_Reset EffectNum, EffectIndex, loopC, Red, Green, Blue, Alpha

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Spawn_Reset(ByVal EffectNum As Byte, _
                               ByVal EffectIndex As Integer, _
                               ByVal Index As Long, _
                               Optional ByVal Red As Single = -1, _
                               Optional ByVal Green As Single = -1, _
                               Optional ByVal Blue As Single = -1, _
                               Optional ByVal Alpha As Single = -1)

        Dim X As Single
        Dim Y As Single
        Dim r As Single

        'Determine if deafults are used

        If Red = -2 Then Red = Rnd
        If Green = -2 Then Green = Rnd
        If Blue = -2 Then Blue = Rnd
        If Alpha = -2 Then Alpha = Rnd
        'store
        effect(EffectIndex).Particles(Index).Red = Red
        effect(EffectIndex).Particles(Index).Green = Green
        effect(EffectIndex).Particles(Index).Blue = Blue
        effect(EffectIndex).Particles(Index).Alpha = Alpha

        Select Case EffectNum

                Case EffectNum_HouseTeleport
                        r = Sin(20 / (Index + 1)) * 100
                        X = r * Cos((Index))
                        Y = r * Sin((Index))
                        'Reset the particle
                        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + X, effect(EffectIndex).Y + Y, 0, 0, 0, 0

                        'Determine if deafults are used

                        If Red = -1 Then Red = Rnd
                        If Green = -1 Then Green = Rnd
                        If Blue = -1 Then Blue = 1
                        If Alpha = -1 Then Alpha = Rnd
                        effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.5)

                Case EffectNum_GuildTeleport
                        r = 150 + Cos(Index * Rnd) * Sin(Index * Rnd)
                        X = r * Cos(Index) * Rnd
                        Y = r * Sin(Index) * Rnd

                        'Determine if deafults are used

                        If Red = -1 Then Red = Rnd
                        If Green = -1 Then Green = Rnd
                        If Blue = -1 Then Blue = 0.5
                        If Alpha = -1 Then Alpha = Rnd
                        'Reset the particle
                        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + X, effect(EffectIndex).Y + Y, 0, 0, 0, 0
                        effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.2)

                Case EffectNum_LevelUP
                        r = 10 + Sin(2 * (Index / 10)) * 50 + (30 * Rnd)
                        X = r * Cos(Index / 30)
                        Y = r * Sin(Index / 30)

                        'Determine if deafults are used

                        If Red = -1 Then Red = 1
                        If Green = -1 Then Green = 0.3 + Rnd / 2
                        If Blue = -1 Then Blue = Rnd / 3
                        If Alpha = -1 Then Alpha = Rnd / 2
                        'Reset the particle
                        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + X, effect(EffectIndex).Y + Y, 0, 0, 0, 0
                        effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.005 + (Rnd * 0.2)

                Case EffectNum_AnimatedSign

                        If Index = 0 Then effect(EffectIndex).Modifier = effect(EffectIndex).Modifier + 1
                        effect(EffectIndex).Progression = effect(EffectIndex).Progression + effect(EffectIndex).Direction

                        If effect(EffectIndex).Progression > 100 Then effect(EffectIndex).Direction = -0.02
                        If effect(EffectIndex).Progression < -100 Then effect(EffectIndex).Direction = 0.02
                        r = effect(EffectIndex).Progression + 2 * Cos(2 * Index) * 40
                        X = r * Cos(Index / (effect(EffectIndex).Modifier + 1) * 5)
                        Y = r * Sin(Index / (effect(EffectIndex).Modifier + 1) * 5)

                        'Determine if deafults are used

                        If Red = -1 Then Red = 1
                        If Green = -1 Then Green = 1
                        If Blue = -1 Then Blue = 1
                        If Alpha = -1 Then Alpha = 1
                        'Reset the particle
                        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + X, effect(EffectIndex).Y + Y, 0, 0, 0, 0
                        effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.2)

                Case EffectNum_Galaxy
                        r = Sin(20 / (Index + 1)) * 100
                        X = r * Cos((Index))
                        Y = r * Sin((Index))

                        'Determine if deafults are used

                        If Red = -1 Then Red = 1
                        If Green = -1 Then Green = 1
                        If Blue = -1 Then Blue = 1
                        If Alpha = -1 Then Alpha = 1
                        'Reset the particle
                        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + X, effect(EffectIndex).Y + Y, 0, 0, 0, 0
                        effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.2)

                Case EffectNum_FancyThickCircle
                        r = 50 + Rnd * 15 * Cos(2 * Index)
                        X = r * Cos(Index / 30)
                        Y = r * Sin(Index / 30)

                        'Determine if deafults are used

                        If Red = -1 Then Red = 1
                        If Green = -1 Then Green = 1
                        If Blue = -1 Then Blue = 1
                        If Alpha = -1 Then Alpha = 1
                        'Reset the particle
                        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + X, effect(EffectIndex).Y + Y, 0, 0, 0, 0
                        effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.2)

                Case EffectNum_Flower
                        r = Cos(2 * (Index / 10)) * 50
                        X = r * Cos(Index / 10)
                        Y = r * Sin(Index / 10)

                        'Determine if deafults are used

                        If Red = -1 Then Red = 1
                        If Green = -1 Then Green = 1
                        If Blue = -1 Then Blue = 1
                        If Alpha = -1 Then Alpha = 1
                        'Reset the particle
                        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + X, effect(EffectIndex).Y + Y, 0, 0, 0, 0
                        effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.2)

                Case EffectNum_Wormhole
                        effect(EffectIndex).Progression = effect(EffectIndex).Progression + 0.1
                        r = (Index / 20) * Exp(Index / effect(EffectIndex).Progression Mod 3)
                        X = r * Cos(Index)
                        Y = r * Sin(Index)

                        'Determine if deafults are used

                        If Red = -1 Then Red = 1
                        If Green = -1 Then Green = 1
                        If Blue = -1 Then Blue = 1
                        If Alpha = -1 Then Alpha = 1
                        'Reset the particle
                        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + X, effect(EffectIndex).Y + Y, 0, 0, 0, 0
                        effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.2)
        End Select

End Sub

Private Sub Effect_Spawn_Update(ByVal EffectNum As Byte, ByVal EffectIndex As Integer)

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Spawn_Reset EffectNum, EffectIndex, loopC, effect(EffectIndex).Particles(loopC).Red, effect(EffectIndex).Particles(loopC).Green, effect(EffectIndex).Particles(loopC).Blue, effect(EffectIndex).Particles(loopC).Alpha
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Public Sub Effect_Create(ByVal QuienLanza As Byte, _
                         ByVal CharIndex As Integer, _
                         ByVal Effecto As Byte)

        With CharList(CharIndex)

                Select Case Effecto

                        Case 1
                               ' .ParticleIndex = Effect_BloodSplatter_Begin(Engine_TPtoSPX(CharList(CharIndex).Pos.X), Engine_TPtoSPY(CharList(CharIndex).Pos.Y), 20 + Rnd * 40)

                        Case 2
                                .ParticleIndex = Effect_Rayo_Begin(Engine_TPtoSPX(CharList(QuienLanza).Pos.X), Engine_TPtoSPY(CharList(QuienLanza).Pos.Y), 13, 100)
                                effect(Effecto).BindToChar = CharIndex
                                effect(Effecto).BindSpeed = 3
                End Select

        End With

End Sub

Function Effect_Rayo_Begin(ByVal X As Single, _
                           ByVal Y As Single, _
                           ByVal Gfx As Integer, _
                           ByVal Particles As Integer, _
                           Optional ByVal Progression As Single = 1) As Integer

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rayo_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Rayo_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Rayo      'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True     'Enabled the effect
        effect(EffectIndex).X = X           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx       'Set the graphic
        effect(EffectIndex).Progression = Progression   'Loop the effect
        effect(EffectIndex).KillWhenAtTarget = True     'End the effect when it reaches the target (progression = 0)
        effect(EffectIndex).KillWhenTargetLost = True   'End the effect if the target is lost (progression = 0)
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Rayo_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Rayo_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rayo_Reset
        '*****************************************************************
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X - 10 + Rnd * 20, effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Cos((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), 0, 0
        'Effect(EffectIndex).Particles(Index).ResetColor 0, 0.8, 0.8, 0.6 + (Rnd * 0.2), 0.001 + (Rnd * 0.5)
        effect(EffectIndex).Particles(Index).ResetColor (Rnd * 0.8), (Rnd * 0.8), (Rnd * 0.8), 0.6 + (Rnd * 0.2), 0.001 + (Rnd * 0.5)
End Sub

Private Sub Effect_Rayo_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rayo_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate the time difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression <> 0 Then
                                        'Reset the particle
                                        Effect_Rayo_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_LissajousMedit_Begin(ByVal X As Single, _
                                     ByVal Y As Single, _
                                     ByVal Gfx As Integer, _
                                     ByVal Particles As Integer, _
                                     Optional ByVal Progression As Single = 0, _
                                     Optional Size As Byte = 30, _
                                     Optional r As Single = 100, _
                                     Optional g As Single = 100, _
                                     Optional b As Single = 100, _
                                     Optional ByVal EcuationCount As Byte = 1) As Long

        '*****************************************************************
        'Particle effect Lissajous equation
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_LissajousMedit_Begin = EffectIndex
        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_LissajousMedit 'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True                     'Enable the effect
        effect(EffectIndex).X = X                           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx                       'Set the graphic
        effect(EffectIndex).Modifier = Size                 'How large the circle is
        effect(EffectIndex).Progression = Progression
        effect(EffectIndex).r = r
        effect(EffectIndex).g = g
        effect(EffectIndex).b = b
        effect(EffectIndex).EcuationCount = EcuationCount
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_LissajousMedit_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_LissajousMedit_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
        '*****************************************************************

        Dim X As Single
        Dim Y As Single
        Dim a As Single

        '2
        '1
        '1
        '2
        effect(EffectIndex).Progression = effect(EffectIndex).Progression + 0.01
        a = effect(EffectIndex).Progression

        With effect(EffectIndex)

                If .EcuationCount = 1 Then
                        X = effect(EffectIndex).X - (Sin(1 * a + 1) * effect(EffectIndex).Modifier) - 20
                        Y = effect(EffectIndex).Y + (Sin(1 * a) * effect(EffectIndex).Modifier)
                        'Reset the particle
                        effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                        effect(EffectIndex).Particles(Index).ResetColor effect(EffectIndex).r * effect(EffectIndex).Progression, effect(EffectIndex).g * effect(EffectIndex).Progression, effect(EffectIndex).b, 0.2, 0.2 + (Rnd * 0.2)
                ElseIf .EcuationCount = 2 Then

                        If RandomNumber(1, 2) = 1 Then
                                X = effect(EffectIndex).X - (Sin(1 * a + 1) * effect(EffectIndex).Modifier) - 20
                                Y = effect(EffectIndex).Y + (Sin(1 * a) * effect(EffectIndex).Modifier)
                                'Reset the particle
                                effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                                effect(EffectIndex).Particles(Index).ResetColor effect(EffectIndex).r * effect(EffectIndex).Progression, effect(EffectIndex).g * effect(EffectIndex).Progression, effect(EffectIndex).b, 0.2, 0.2 + (Rnd * 0.2)
                        Else
                                X = .X - (Sin(1 * a) * .Modifier) - 20
                                Y = .Y + (Sin(1 * a) * .Modifier)
                                'Reset the particle
                                .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                                .Particles(Index).ResetColor .r * .Progression, .g * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)
                        End If

                ElseIf .EcuationCount = 3 Then

                        If RandomNumber(1, 2) = 1 Then
                                X = .X - (Sin(2 * a) * .Modifier) - 20
                                Y = .Y + (Sin(1 * a) * .Modifier)
                                'Reset the particle
                                .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                                .Particles(Index).ResetColor .r * .Progression, .g * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)
                        Else
                                X = .X - (Sin(1 * a) * .Modifier) - 20
                                Y = .Y + (Sin(2 * a) * .Modifier)
                                'Reset the particle
                                .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                                .Particles(Index).ResetColor .r * .Progression, .g * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)
                        End If

                ElseIf .EcuationCount = 4 Then

                        If RandomNumber(1, 2) = 1 Then
                                X = .X - (Sin(4 * a) * .Modifier) - 20
                                Y = .Y + (Sin(2 * a) * .Modifier)
                                'Reset the particle
                                .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                                .Particles(Index).ResetColor .r * .Progression, .g * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)
                        Else
                                X = .X - (Sin(2 * a) * .Modifier) - 20
                                Y = .Y + (Sin(4 * a) * .Modifier)
                                'Reset the particle
                                .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                                .Particles(Index).ResetColor .r * .Progression, .g * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)
                        End If

                ElseIf .EcuationCount = 5 Then

                        If RandomNumber(1, 2) = 1 Then
                                X = .X - (Sin(2 * a) * .Modifier) - 20
                                Y = .Y + (Sin(1 * a) * .Modifier)
                                'Reset the particle
                                .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                                .Particles(Index).ResetColor .r * .Progression, .g * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)
                        Else
                                X = .X - (Sin(1 + 5 * a) * .Modifier) - 20
                                Y = .Y + (Sin(2 + 7 * a) * .Modifier)
                                'Reset the particle
                                .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
                                .Particles(Index).ResetColor .r * .Progression, .g * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)
                        End If
                End If

        End With

End Sub

Private Sub Effect_LissajousMedit_Update(ByVal EffectIndex As Integer)

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
        '*****************************************************************

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_LissajousMedit_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_Inmovilizar_Begin(ByVal X As Single, _
                                  ByVal Y As Single, _
                                  ByVal Gfx As Integer, _
                                  ByVal Particles As Integer, _
                                  Optional ByVal Size As Byte = 30, _
                                  Optional ByVal Time As Single = 10) As Long

        '*****************************************************************
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Begin
        '*****************************************************************

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function
        'Return the index of the used slot
        Effect_Inmovilizar_Begin = EffectIndex
        'Set the effect's variables
        effect(EffectIndex).EffectNum = EffectNum_Inmovilizar    'Set the effect number
        effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
        effect(EffectIndex).Used = True             'Enabled the effect
        effect(EffectIndex).X = X                   'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx               'Set the graphic
        effect(EffectIndex).Modifier = Size         'How large the circle is
        effect(EffectIndex).Progression = Time      'How long the effect will last
        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount
        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles
        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Inmovilizar_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime
End Function

Private Sub Effect_Inmovilizar_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        '// Posiciones

        Dim X As Single
        Dim Y As Single

        'Get the positions
        X = effect(EffectIndex).X + (Rnd * 60)
        Y = effect(EffectIndex).Y + (Rnd * 60)

        '// Colores
        '// Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt X, Y, Rnd * 1 - 0.5, Rnd * 1 - 0.5, Rnd - 0.5, Rnd * -0.9 + 0.45
        effect(EffectIndex).Particles(Index).ResetColor (Rnd * 0.8), (Rnd * 0.8), (Rnd * 0.8), 0.6 + (Rnd * 0.2), 0.07 + Rnd * 0.01
End Sub

Private Sub Effect_Inmovilizar_Update(ByVal EffectIndex As Integer)

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate the time difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime

        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go through the particle loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check if particle is in use

                If effect(EffectIndex).Particles(loopC).Used Then
                        'Update the particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then
                                        'Reset the particle
                                        Effect_Inmovilizar_Reset EffectIndex, loopC
                                Else
                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False
                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False
                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0
                                End If

                        Else
                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY
                        End If
                End If

        Next loopC

End Sub

Function Effect_ChangeClass_Begin(ByVal X As Single, _
                                  ByVal Y As Single, _
                                  ByVal Gfx As Integer, _
                                  ByVal Particles As Integer, _
                                  Optional ByVal Progression As Single = 1) As Integer

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function

        'Return the index of the used slot
        Effect_ChangeClass_Begin = EffectIndex

        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_ChangeClass  'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True                     'Enable the effect
        effect(EffectIndex).X = X                           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx                       'Set the graphic
        effect(EffectIndex).Progression = Progression       'If we loop the effect

        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount

        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_ChangeClass_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_ChangeClass_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        Dim X As Single
        Dim Y As Single

        effect(EffectIndex).Sng = effect(EffectIndex).Sng + 0.03

        If effect(EffectIndex).Sng > 360 * DegreeToRadian Then effect(EffectIndex).Sng = effect(EffectIndex).Sng - 360 * DegreeToRadian
        effect(EffectIndex).Modifier = effect(EffectIndex).Modifier + 1
   
        'Get the positions
        X = effect(EffectIndex).X - (Sin(effect(EffectIndex).Sng) * 40) + Rnd * 10
        Y = effect(EffectIndex).Y + (Cos(effect(EffectIndex).Sng) * 40) - (effect(EffectIndex).Modifier / 10) + Rnd * 10
 
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        effect(EffectIndex).Particles(Index).ResetColor 1, 1, 0, 1, 0.2 + (Rnd * 0.2)

End Sub

Private Sub Effect_ChangeClass_Update(ByVal EffectIndex As Integer)

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime
   
        'Update the life span

        If effect(EffectIndex).Progression > 0 Then effect(EffectIndex).Progression = effect(EffectIndex).Progression - ElapsedTime

        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then

                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 0 Then

                                        'Reset the particle
                                        Effect_ChangeClass_Reset EffectIndex, loopC

                                Else

                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False

                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False

                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0

                                End If

                        Else

                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY

                        End If

                End If

        Next loopC

End Sub

Function Effect_Armada_Begin(ByVal X As Single, _
                             ByVal Y As Single, _
                             ByVal Gfx As Integer, _
                             ByVal Particles As Integer, _
                             Optional ByVal Progression As Single = 1) As Integer

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function

        'Return the index of the used slot
        Effect_Armada_Begin = EffectIndex

        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_Armada       'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True                     'Enable the effect
        effect(EffectIndex).X = X                           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx                       'Set the graphic
        effect(EffectIndex).Progression = Progression       'If we loop the effect

        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount

        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Armada_Reset EffectIndex, loopC

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Armada_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

        Dim X  As Single
        Dim Y  As Single

        Dim a  As Single
        Dim b  As Integer
        Dim aL As Single
        Dim rG As Single

        aL = 3.1415 / 2
        a = 3
        b = 4
    
        rG = (Rnd * 0.6)
    
        effect(EffectIndex).Progression = effect(EffectIndex).Progression + 0.14
    
        X = a * Sin(effect(EffectIndex).Progression * aL / 12) * 6 + (Rnd * 0.5) - 1
        Y = b * Sin(effect(EffectIndex).Progression / 12) * 6 - 10 + (Rnd * 0.5) - 10
    
        'Reset the particle
    
        If RandomNumber(1, 2) = 1 Then
                effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X + X, effect(EffectIndex).Y + Y, 0, 0, 0, 0
                effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.2, rG, 0.1 + (Rnd * 0.2), 0.1
        Else
                effect(EffectIndex).Particles(Index).ResetIt effect(EffectIndex).X - X, effect(EffectIndex).Y - Y - 50, 0, 0, 0, 0
                effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.2, rG, 0.1 + (Rnd * 0.1), 0.036
        End If

End Sub

Private Sub Effect_Armada_Update(ByVal EffectIndex As Integer)

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime
    
        'Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.001
        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then

                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0.5 And RandomNumber(1, 4) = 1 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 1 Then

                                        'Reset the particle
                                        Effect_Armada_Reset EffectIndex, loopC

                                Else

                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False

                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False

                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0

                                End If

                        Else

                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY

                        End If

                End If

        Next loopC

End Sub

Function Effect_Butterfly_Begin(ByVal X As Single, _
                                ByVal Y As Single, _
                                ByVal Gfx As Integer, _
                                ByVal Particles As Integer, _
                                Optional ByVal Progression As Single = 1, _
                                Optional ByVal Aura As Byte = 0) As Integer

        Dim EffectIndex As Integer
        Dim loopC       As Long

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function

        'Return the index of the used slot
        Effect_Butterfly_Begin = EffectIndex

        'Set The Effect's Variables
        effect(EffectIndex).EffectNum = EffectNum_ButterflyCurve       'Set the effect number
        effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
        effect(EffectIndex).Used = True                     'Enable the effect
        effect(EffectIndex).X = X                           'Set the effect's X coordinate
        effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
        effect(EffectIndex).Gfx = Gfx                       'Set the graphic
        effect(EffectIndex).Progression = Progression       'If we loop the effect

        'Set the number of particles left to the total avaliable
        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticleCount

        'Set the float variables
        effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

        'Redim the number of particles
        ReDim effect(EffectIndex).Particles(0 To effect(EffectIndex).ParticleCount)
        ReDim effect(EffectIndex).PartVertex(0 To effect(EffectIndex).ParticleCount)

        'Create the particles

        For loopC = 0 To effect(EffectIndex).ParticleCount
                Set effect(EffectIndex).Particles(loopC) = New clsParticle
                effect(EffectIndex).Particles(loopC).Used = True
                effect(EffectIndex).PartVertex(loopC).rhw = 1
                Effect_Butterfly_Reset EffectIndex, loopC, Aura

        Next loopC

        'Set The Initial Time
        effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Butterfly_Reset(ByVal EffectIndex As Integer, _
                                   ByVal Index As Long, _
                                   ByVal AuraType As Byte)

        Dim X As Single
        Dim Y As Single
    
        If AuraType = 1 Then 'la posta del auratype es mandarle una cantidad de 300 particulas
                effect(EffectIndex).Sng = effect(EffectIndex).Sng + 0.003
        Else
                effect(EffectIndex).Sng = effect(EffectIndex).Sng + 0.03
        End If
    
        If effect(EffectIndex).Sng > 360 * DegreeToRadian Then effect(EffectIndex).Sng = effect(EffectIndex).Sng - 360 * DegreeToRadian
        effect(EffectIndex).Modifier = effect(EffectIndex).Modifier + 1
   
        'Get the positions
        'x = Effect(EffectIndex).x - (Sin(Effect(EffectIndex).Sng) * 40) + Rnd * 10
        'y = Effect(EffectIndex).y + (Cos(Effect(EffectIndex).Sng) * 40) - (Effect(EffectIndex).Modifier / 10) + Rnd * 10
    
        With effect(EffectIndex)
        
                If AuraType = 1 Then
                        If RandomNumber(1, 2) = 1 Then
                                X = .X + (.X * (Sin(.Sng) * (Exp(Cos(.Sng)) - 2 * Cos(4 * .Sng) - Sin((.Sng / 12) ^ 5)))) / 5
                                Y = .Y - (.Y * (Cos(.Sng) * (Exp(Cos(.Sng)) - 2 * Cos(4 * .Sng) - Sin((.Sng / 12) ^ 5)))) / 5
                        Else
                                X = .X + (.X * -(Sin(.Sng) * (Exp(Cos(.Sng)) - 2 * Cos(4 * .Sng) - Sin((.Sng / 12) ^ 5)))) / 5
                                Y = .Y - (.Y * -(Cos(.Sng) * (Exp(Cos(.Sng)) - 2 * Cos(4 * .Sng) - Sin((.Sng / 12) ^ 5)))) / 5

                        End If

                Else
                        X = .X + (.X * (Sin(.Sng) * (Exp(Cos(.Sng)) - 2 * Cos(4 * .Sng) - Sin((.Sng / 12) ^ 5)))) / 5
                        Y = .Y - (.Y * (Cos(.Sng) * (Exp(Cos(.Sng)) - 2 * Cos(4 * .Sng) - Sin((.Sng / 12) ^ 5)))) / 5

                End If
    
        End With
    
        'Reset the particle
        effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        effect(EffectIndex).Particles(Index).ResetColor 1, 1, 0, 1, 0.2 + (Rnd * 0.2)
End Sub

Private Sub Effect_Butterfly_Update(ByVal EffectIndex As Integer)

        Dim ElapsedTime As Single
        Dim loopC       As Long

        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - effect(EffectIndex).PreviousFrame) * 0.01
        effect(EffectIndex).PreviousFrame = timeGetTime
    
        'Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.001
        'Go Through The Particle Loop

        For loopC = 0 To effect(EffectIndex).ParticleCount

                'Check If Particle Is In Use

                If effect(EffectIndex).Particles(loopC).Used Then

                        'Update The Particle
                        effect(EffectIndex).Particles(loopC).UpdateParticle ElapsedTime

                        'Check if the particle is ready to die

                        If effect(EffectIndex).Particles(loopC).sngA <= 0.5 And RandomNumber(1, 4) = 1 Then

                                'Check if the effect is ending

                                If effect(EffectIndex).Progression > 1 Then

                                        'Reset the particle
                                        Effect_Armada_Reset EffectIndex, loopC

                                Else

                                        'Disable the particle
                                        effect(EffectIndex).Particles(loopC).Used = False

                                        'Subtract from the total particle count
                                        effect(EffectIndex).ParticlesLeft = effect(EffectIndex).ParticlesLeft - 1

                                        'Check if the effect is out of particles

                                        If effect(EffectIndex).ParticlesLeft = 0 Then effect(EffectIndex).Used = False

                                        'Clear the color (dont leave behind any artifacts)
                                        effect(EffectIndex).PartVertex(loopC).color = 0

                                End If

                        Else

                                'Set the particle information on the particle vertex
                                effect(EffectIndex).PartVertex(loopC).color = D3DColorMake(effect(EffectIndex).Particles(loopC).sngR, effect(EffectIndex).Particles(loopC).sngG, effect(EffectIndex).Particles(loopC).sngB, effect(EffectIndex).Particles(loopC).sngA)
                                effect(EffectIndex).PartVertex(loopC).X = effect(EffectIndex).Particles(loopC).sngX
                                effect(EffectIndex).PartVertex(loopC).Y = effect(EffectIndex).Particles(loopC).sngY

                        End If

                End If

        Next loopC

End Sub


