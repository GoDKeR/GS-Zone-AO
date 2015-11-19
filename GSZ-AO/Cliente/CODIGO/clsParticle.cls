VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
'Values Stored For Each Particle
Private mvarused          As Boolean
Private mvarsngX          As Single
Private mvarsngY          As Single
Private mvarsngXSpeed     As Single
Private mvarsngYSpeed     As Single
Private mvarsngXAccel     As Single
Private mvarsngYAccel     As Single
Private mvarsngR          As Single
Private mvarsngG          As Single
Private mvarsngB          As Single
Private mvarsngA          As Single
Private mvarsngAlphaDecay As Single

Private mvarsngZ          As Single
Private mvarsngZSpeed     As Single
Private mvarsngZAccel     As Single

Private mvarAlpha         As Single
Private mvarRed           As Single
Private mvarGreen         As Single
Private mvarBlue          As Single

Public Sub ResetColor(sngRed As Single, sngGreen As Single, sngBlue As Single, sngAlpha As Single, sngDecay As Single)
'************************************************************
'Change the particle's color values
'More info: http://www.vbgore.com/CommonCode.Particle.ResetColor" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************
 
    sngR = sngRed
    sngG = sngGreen
    sngB = sngBlue
    sngA = sngAlpha
    sngAlphaDecay = sngDecay
 
End Sub
 
Public Sub ResetIt(X As Single, Y As Single, XSpeed As Single, YSpeed As Single, XAcc As Single, YAcc As Single, Optional Z As Single = 0, _
                   Optional ZSpeed As Single = 0, _
                   Optional ZAcc As Single = 0)
'************************************************************
'Reset the particle's variables
'More info: http://www.vbgore.com/CommonCode.Particle.ResetIt" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************
 
    sngX = X
    sngY = Y
    sngXSpeed = XSpeed
    sngYSpeed = YSpeed
    sngXAccel = XAcc
    sngYAccel = YAcc
 
End Sub

Public Property Get Alpha() As Single
        Alpha = mvarAlpha
End Property

Public Property Get Red() As Single
        Red = mvarRed
End Property

Public Property Get Green() As Single
        Green = mvarGreen
End Property

Public Property Get Blue() As Single
        Blue = mvarBlue
End Property

Public Property Let Alpha(ByVal vData As Single)
        mvarAlpha = vData
End Property

Public Property Let Red(ByVal vData As Single)
        mvarRed = vData
End Property

Public Property Let Green(ByVal vData As Single)
        mvarGreen = vData
End Property

Public Property Let Blue(ByVal vData As Single)
        mvarBlue = vData
End Property
Public Property Get sngA() As Single
'Lorwik> Particulas de vbgore http://www.****.net" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
    sngA = mvarsngA
 
End Property
 
Public Property Let sngA(ByVal vData As Single)
 
    mvarsngA = vData
 
End Property
 
Public Property Get sngAlphaDecay() As Single
 
    sngAlphaDecay = mvarsngAlphaDecay
 
End Property
 
Public Property Let sngAlphaDecay(ByVal vData As Single)
 
    mvarsngAlphaDecay = vData
 
End Property
 
Public Property Get sngB() As Single
 
    sngB = mvarsngB
 
End Property
 
Public Property Let sngB(ByVal vData As Single)
 
    mvarsngB = vData
 
End Property
 
Public Property Get sngG() As Single
 
    sngG = mvarsngG
 
End Property
 
Public Property Let sngG(ByVal vData As Single)
 
    mvarsngG = vData
 
End Property
 
Public Property Get sngR() As Single
 
    sngR = mvarsngR
 
End Property
 
Public Property Let sngR(ByVal vData As Single)
 
    mvarsngR = vData
 
End Property
 
Public Property Get sngX() As Single
 
    sngX = mvarsngX
 
End Property
 
Public Property Let sngX(ByVal vData As Single)
 
    mvarsngX = vData
 
End Property
 
Public Property Get sngXAccel() As Single
 
    sngXAccel = mvarsngXAccel
 
End Property
 
Public Property Let sngXAccel(ByVal vData As Single)
 
    mvarsngXAccel = vData
 
End Property
 
Public Property Get sngXSpeed() As Single
 
    sngXSpeed = mvarsngXSpeed
 
End Property
 
Public Property Let sngXSpeed(ByVal vData As Single)
 
    mvarsngXSpeed = vData
 
End Property
 
Public Property Get sngY() As Single
 
    sngY = mvarsngY
 
End Property
 
Public Property Let sngY(ByVal vData As Single)
 
    mvarsngY = vData
 
End Property
 
Public Property Get sngYAccel() As Single
 
    sngYAccel = mvarsngYAccel
 
End Property
 
Public Property Let sngYAccel(ByVal vData As Single)
 
    mvarsngYAccel = vData
 
End Property
 
Public Property Get sngYSpeed() As Single
 
    sngYSpeed = mvarsngYSpeed
 
End Property
 
Public Property Let sngYSpeed(ByVal vData As Single)
 
    mvarsngYSpeed = vData
 
End Property

Public Property Get sngZSpeed() As Single

        sngZSpeed = mvarsngZSpeed
        
End Property

Public Property Let sngZSpeed(ByVal vData As Single)

        mvarsngZSpeed = vData
        
End Property

Public Property Get sngZ() As Single

        sngZ = mvarsngZ
        
End Property

Public Property Let sngZ(ByVal vData As Single)

        mvarsngZ = vData
        
End Property

Public Property Get sngZAccel() As Single

        sngZAccel = mvarsngZAccel
        
End Property

Public Property Let sngZAccel(ByVal vData As Single)

        mvarsngZAccel = vData
        
End Property

Public Sub UpdateParticle(sngTime As Single)
'************************************************************
'Update the particle's variables
'More info: http://www.vbgore.com/CommonCode.Partic ... teParticle" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************
'Lorwik> Particulas de vbgore http://www.****.net" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
    
    sngX = sngX + (LastOffsetX - ParticleOffsetX) + sngXSpeed * sngTime
    sngY = sngY + (LastOffsetY - ParticleOffsetY) + sngYSpeed * sngTime
    
    sngZ = sngZ + sngZSpeed * sngTime
     
    sngXSpeed = sngXSpeed + sngXAccel * sngTime
    sngYSpeed = sngYSpeed + sngYAccel * sngTime
    sngZSpeed = sngZSpeed + sngZAccel * sngTime
    
    sngA = sngA - sngAlphaDecay * sngTime
 
End Sub
 
Public Property Get Used() As Boolean
 
    Used = mvarused
 
End Property
 
Public Property Let Used(ByVal vData As Boolean)
 
    mvarused = vData
 
End Property

