VERSION 5.00
Begin VB.UserControl PBarY 
   BackColor       =   &H00000000&
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   FillStyle       =   0  'Solid
   ScaleHeight     =   615
   ScaleWidth      =   615
   ToolboxBitmap   =   "PBarY.ctx":0000
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "PBarY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ProgressBarSlider Pro ActiveX © 2000 by Nik Tupkalov
' This ActiveX Control was writen by Nik Tupkalov
'Default Property Values:
Const m_def_Style = 0
Const m_def_BackStyle = 0
Const m_def_picForeColor = &H404040
Const m_def_picFillColor = &HFFFF00
Const m_def_picStep = 50
Const m_def_MousePointer = 9
Const m_def_EnabledSlider = True
Const m_def_BorderStyle = 0
Const m_def_Value = 25
Const m_def_Min = 0
Const m_def_Max = 100
'Property Variables:
Dim m_Style As bView
Dim m_BackStyle As bStyle
Dim m_picForeColor As OLE_COLOR
Dim m_picFillColor As OLE_COLOR
Dim m_picStep As Integer
Dim m_MousePointer As bMouse
Dim m_EnabledSlider As Boolean
Dim m_BorderStyle As rStyle
Dim m_Value As Long
Dim m_Min As Integer
Dim m_Max As Integer
Dim Ref As Boolean
'Event Declarations:
Event Click()
Event ChangeValue(NewValue As Long, OldValue As Long)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

Public Enum bView
    Normal
    Digital
   End Enum
   
Public Enum bStyle
    Flat
    b3D
   End Enum

Public Enum rStyle
    Transparent
    Solid
    Dash
    Dot
    DashDot
    DashDotDot
    InsideSolid
    End Enum
    
Public Enum bMouse
    Default
    Arrow
    Cross
    Beam
    Icon
    Size
    SizeNES
    SizeNS
    SizeNWS
    SizeWE
    UpArrow
    Hourglass
    NoDrop
    ArrowG
    ArrowH
    SizeAll
    Custom = 99
   End Enum
   
Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
RaiseEvent MouseDown(Button, Shift, X, y)
    If Not m_EnabledSlider Then
        UserControl.MousePointer = Default: Exit Sub
    Else
UserControl.MousePointer = m_MousePointer
    End If
            GetValue X
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, X, y)
        If Not m_EnabledSlider Then
    UserControl.MousePointer = Default: Exit Sub
        Else
            UserControl.MousePointer = m_MousePointer
        End If
    If Button <> 1 Then Exit Sub
            GetValue X
End Sub

Private Sub GetValue(ByVal X As Single)
    If X < 0 Then X = 0
        If X > ScaleWidth Then X = ScaleWidth

Static o_Value As Long

        o_Value = m_Value
    m_Value = X / ScaleWidth * (m_Max - m_Min) + m_Min
        If m_Style = Normal Then
    If Ref Then Ref = False: Cls
            Shape1.Visible = True
        Shape1.Width = ScaleWidth * (m_Value - m_Min) / (m_Max - m_Min)
    Else
        Shape1.Visible = False
            If Ref Then Ref = False: Cls
Static X1 As Single

            For X1 = 0 To ScaleWidth Step m_picStep
        If X1 <= X Then
            Line (X1, 0)-(X1, ScaleHeight), m_picFillColor, BF
        Else
            Line (X1, 0)-(X1, ScaleHeight), m_picForeColor, BF
        End If
            Next X1
    End If
    PropertyChanged "Value"
RaiseEvent ChangeValue(m_Value, o_Value)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
RaiseEvent MouseUp(Button, Shift, X, y)
End Sub

Private Sub UserControl_Paint()
Ref = True: RefreshBar
End Sub

Private Sub UserControl_Resize()
Shape1.Height = ScaleHeight
Ref = True: RefreshBar
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get value() As Long
    value = m_Value
End Property

Public Property Let value(ByVal New_Value As Long)
        m_Value = New_Value
    If m_Value < m_Min Then m_Value = m_Min
If m_Value > m_Max Then m_Value = m_Max
        PropertyChanged "Value"
    RefreshBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Min() As Integer
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    m_Min = New_Min
  PropertyChanged "Min"
Shape1.Width = ScaleWidth * (m_Value - m_Min) / (m_Max - m_Min)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,100
Public Property Get Max() As Integer
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    m_Max = New_Max
  PropertyChanged "Max"
Shape1.Width = ScaleWidth * (m_Value - m_Min) / (m_Max - m_Min)
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_BorderStyle = m_def_BorderStyle
    m_EnabledSlider = m_def_EnabledSlider
    m_MousePointer = m_def_MousePointer
    m_BackStyle = m_def_BackStyle
    m_picForeColor = m_def_picForeColor
    m_picFillColor = m_def_picFillColor
    m_picStep = m_def_picStep
    m_Style = m_def_Style
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000007)
    Shape1.FillColor = PropBag.ReadProperty("FillColor", &HFFFF&)
    Shape1.BorderColor = PropBag.ReadProperty("BorderColor", &HFF&)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    UserControl.BorderStyle = m_BackStyle
    Shape1.BorderStyle = m_BorderStyle
    m_EnabledSlider = PropBag.ReadProperty("EnabledSlider", m_def_EnabledSlider)
    m_picForeColor = PropBag.ReadProperty("picForeColor", m_def_picForeColor)
    m_picFillColor = PropBag.ReadProperty("picFillColor", m_def_picFillColor)
    m_picStep = PropBag.ReadProperty("picStep", m_def_picStep)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    UserControl.BorderStyle = m_BackStyle
    RefreshBar
    If Not m_EnabledSlider Then Exit Sub
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    m_MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
    UserControl.MousePointer = m_MousePointer
End Sub

Private Sub UserControl_Show()
Ref = True: RefreshBar
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000007)
    Call PropBag.WriteProperty("FillColor", Shape1.FillColor, &HFFFF&)
    Call PropBag.WriteProperty("BorderColor", Shape1.BorderColor, &HFF&)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("EnabledSlider", m_EnabledSlider, m_def_EnabledSlider)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("picForeColor", m_picForeColor, m_def_picForeColor)
    Call PropBag.WriteProperty("picFillColor", m_picFillColor, m_def_picFillColor)
    Call PropBag.WriteProperty("picStep", m_picStep, m_def_picStep)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = Shape1.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    Shape1.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,BorderColor
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = Shape1.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    Shape1.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=22,0,0,0
Public Property Get BorderStyle() As rStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As rStyle)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
Shape1.BorderStyle = m_BorderStyle
RefreshBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get EnabledSlider() As Boolean
Attribute EnabledSlider.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    EnabledSlider = m_EnabledSlider
End Property

Public Property Let EnabledSlider(ByVal New_EnabledSlider As Boolean)
    m_EnabledSlider = New_EnabledSlider
    PropertyChanged "EnabledSlider"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,0,0,0
Public Property Get MousePointer() As bMouse
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As bMouse)
    m_MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
    UserControl.MousePointer = m_MousePointer
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,0
Public Property Get BackStyle() As bStyle
Attribute BackStyle.VB_Description = "Returns/sets the border style for an object."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As bStyle)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
    UserControl.BorderStyle = m_BackStyle
Ref = True
RefreshBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbMagenta
Public Property Get picForeColor() As OLE_COLOR
    picForeColor = m_picForeColor
End Property

Public Property Let picForeColor(ByVal New_picForeColor As OLE_COLOR)
    m_picForeColor = New_picForeColor
    PropertyChanged "picForeColor"
If m_Style = Digital Then RefreshBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbBlue
Public Property Get picFillColor() As OLE_COLOR
    picFillColor = m_picFillColor
End Property

Public Property Let picFillColor(ByVal New_picFillColor As OLE_COLOR)
    m_picFillColor = New_picFillColor
    PropertyChanged "picFillColor"
If m_Style = Digital Then RefreshBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,50
Public Property Get picStep() As Integer
    picStep = m_picStep
End Property

Public Property Let picStep(ByVal New_picStep As Integer)
    If New_picStep < 10 Then New_picStep = 10
If New_picStep > ScaleWidth / 10 Then New_picStep = ScaleWidth / 10
    m_picStep = New_picStep
        PropertyChanged "picStep"
If m_Style = Digital Then Ref = True: RefreshBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=24,0,0,0
Public Property Get style() As bView
    style = m_Style
End Property

Public Property Let style(ByVal New_Style As bView)
    Ref = True
    m_Style = New_Style
    PropertyChanged "Style"
RefreshBar
End Property


Private Sub RefreshBar(Optional ByVal value As Long)
If value = Empty Then value = m_Value
GetValue ScaleWidth * (value - m_Min) / (m_Max - m_Min)
End Sub


