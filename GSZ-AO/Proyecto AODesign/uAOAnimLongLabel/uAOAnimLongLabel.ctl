VERSION 5.00
Begin VB.UserControl uAOAnimLongLabel 
   BackStyle       =   0  'Transparent
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   ScaleHeight     =   870
   ScaleWidth      =   2295
   ToolboxBitmap   =   "uAOAnimLongLabel.ctx":0000
   Begin VB.Timer tTimer 
      Interval        =   10
      Left            =   1680
      Top             =   360
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   405
      TabIndex        =   0
      Top             =   0
      Width           =   285
   End
   Begin VB.Shape shpBack 
      BorderColor     =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   2085
   End
End
Attribute VB_Name = "uAOAnimLongLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%                                                     %
'%                   AO ANIMLABEL v1.0                 %
'%               Copyright © 2013 by ^[GS]^            %
'%                    www.GS-ZONE.org                  %
'%                                                     %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%  Este control animado de label contador.            %
'%                                                     %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%  Changelog:                                         %
'%   30/04/2013 - Se termino la primera versión        %
'%                funcional, con eventos de click,     %
'%                doble click y de cursor. (^[GS]^)    %
'%   25/04/2013 - Se inicio el proyecto. (^[GS]^)      %
'%                                                     %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Option Explicit

Private iValue As Long
Private iNewValue As Long
Private bEnabled As Boolean
Private bUseBackground As Boolean
Private lForeColor As Long
Private lBackColor As Long
Private lBorderColor As Long
Private fTextFont As Font

Private bAnimate As Boolean
Private bAnimating As Boolean

Private Const MAX_LONG As Long = 2147483647

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub lblLabel_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If bEnabled Then
        RaiseEvent Click
    End If
    
End Sub

Private Sub lblLabel_DblClick()
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If bEnabled Then
        RaiseEvent DblClick
    End If
    
End Sub

Private Sub lblLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If bEnabled Then
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
    
End Sub

Private Sub lblLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If bEnabled Then
        RaiseEvent MouseMove(Button, Shift, X, Y)
    End If
    
End Sub

Private Sub lblLabel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If bEnabled Then
        RaiseEvent MouseUp(Button, Shift, X, Y)
    End If

End Sub

Private Sub UserControl_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If bEnabled Then
        RaiseEvent Click
    End If
    
End Sub

Private Sub UserControl_DblClick()
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If bEnabled Then
        RaiseEvent DblClick
    End If
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If bEnabled Then
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If bEnabled Then
        RaiseEvent MouseMove(Button, Shift, X, Y)
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If bEnabled Then
        RaiseEvent MouseUp(Button, Shift, X, Y)
    End If

End Sub

Private Sub DrawLabel()
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If bEnabled = False Then Exit Sub

    If bAnimate = False Then
        iNewValue = iValue
        lblLabel.Caption = iNewValue
    Else
        If iNewValue = iValue Then
            tTimer.Enabled = False
        Else
            tTimer.Enabled = True
        End If
        Dim lDif As Long
        lDif = Abs(iValue - iNewValue)
        If iNewValue < iValue Then
            iNewValue = iNewValue + 1
            Select Case lDif
                Case Is > 5000
                    iNewValue = iNewValue + (lDif / 8)
                Case Is > 1000
                    iNewValue = iNewValue + (lDif / 14)
                Case Is > 10
                    iNewValue = iNewValue + (lDif / 18)
            End Select
            If iNewValue > iValue Then iNewValue = iValue
            bAnimating = True
        ElseIf iNewValue > iValue Then
            iNewValue = iNewValue - 1
            Select Case lDif
                Case Is > 5000
                    iNewValue = iNewValue - (lDif / 8)
                Case Is > 1000
                    iNewValue = iNewValue - (lDif / 14)
                Case Is > 10
                    iNewValue = iNewValue - (lDif / 18)
            End Select
            If iNewValue < iValue Then iNewValue = iValue
            bAnimating = True
        Else
            iNewValue = iValue
            bAnimating = False
        End If
        lblLabel.Caption = iNewValue
    End If
    
End Sub

Private Sub tTimer_Timer()
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If bEnabled = False Then
        tTimer.Enabled = False
        Exit Sub
    End If
    If bAnimate = True Then
        Call DrawLabel
    End If
    
End Sub

Private Sub ResizeLabel()
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    lblLabel.Left = 0
    lblLabel.Width = UserControl.Width
    lblLabel.Top = (UserControl.Height / 2) - ((lblLabel.Height / 2))
    Call DrawLabel
    
End Sub

Private Sub UserControl_InitProperties()
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next

    iValue = 1
    bEnabled = True
    bAnimate = True
    bUseBackground = True
    lForeColor = RGB(255, 255, 255)
    lBackColor = RGB(0, 0, 0)
    lBorderColor = RGB(200, 200, 200)
    
End Sub

Private Sub UserControl_Resize()
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    shpBack.Height = UserControl.Height
    shpBack.Width = UserControl.Width
    
    Call ResizeLabel
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    With PropBag
        iValue = .ReadProperty("Value", 50)
        bEnabled = .ReadProperty("Enabled", True)
        bAnimate = .ReadProperty("Animate", True)
        bUseBackground = .ReadProperty("UseBackground", True)
        lForeColor = .ReadProperty("ForeColor", RGB(255, 255, 255))
        lBackColor = .ReadProperty("BackColor", RGB(0, 0, 0))
        lBorderColor = .ReadProperty("BorderColor", RGB(200, 200, 200))
        Set lblLabel.Font = .ReadProperty("FONT", lblLabel.Font)
    End With
    
    lblLabel.ForeColor = lForeColor
    shpBack.FillColor = lBackColor
    shpBack.BorderColor = lBorderColor
    shpBack.Visible = bUseBackground
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    With PropBag
        .WriteProperty "Value", iValue, 50
        .WriteProperty "Enabled", bEnabled, True
        .WriteProperty "Animate", bAnimate, True
        .WriteProperty "UseBackground", bUseBackground, True
        .WriteProperty "ForeColor", lForeColor, RGB(255, 255, 255)
        .WriteProperty "BackColor", lBackColor, RGB(0, 0, 0)
        .WriteProperty "BorderColor", lBorderColor, RGB(200, 200, 200)
        Call .WriteProperty("FONT", lblLabel.Font)
    End With
    
End Sub

Public Property Get Enabled() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    Enabled = bEnabled
    
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    bEnabled = NewValue
    PropertyChanged "Enabled"
    
    UserControl.Enabled = False
    
End Property

Public Property Get Animado() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    Animado = bAnimate
    
End Property

Public Property Let Animado(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    bAnimate = NewValue
    PropertyChanged "Animate"
    
    Call DrawLabel
    
End Property

Public Property Get UseBackground() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    UseBackground = bUseBackground
    
End Property

Public Property Let UseBackground(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    bUseBackground = NewValue
    PropertyChanged "UseBackground"
    
    shpBack.Visible = bUseBackground
    
End Property

Public Property Get Font() As Font
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    Set Font = lblLabel.Font
    
End Property

Public Property Set Font(ByRef newFont As Font)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    Set lblLabel.Font = newFont

    Call ResizeLabel

    PropertyChanged "FONT"
    
End Property

Public Property Get FontBold() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    FontBold = lblLabel.FontBold
    
End Property

Public Property Let FontBold(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    lblLabel.FontBold = NewValue
    
    Call ResizeLabel
    
End Property

Public Property Get FontItalic() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    FontItalic = lblLabel.FontItalic
    
End Property

Public Property Let FontItalic(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    lblLabel.FontItalic = NewValue

    Call ResizeLabel
    
End Property

Public Property Get FontUnderline() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    FontUnderline = lblLabel.FontUnderline
    
End Property

Public Property Let FontUnderline(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    lblLabel.FontUnderline = NewValue

    Call ResizeLabel
    
End Property

Public Property Get FontSize() As Integer
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    FontSize = lblLabel.FontSize
    
End Property

Public Property Let FontSize(ByVal NewValue As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    lblLabel.FontSize = NewValue

    Call ResizeLabel
    
End Property

Public Property Get FontName() As String
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    FontName = lblLabel.FontName
    
End Property

Public Property Let FontName(ByVal NewValue As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    lblLabel.FontName = NewValue
    
    Call ResizeLabel
    
End Property

Public Property Get ForeColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    ForeColor = lForeColor
    
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    lForeColor = NewValue
    PropertyChanged "ForeColor"
    
    lblLabel.ForeColor = lForeColor
    
End Property

Public Property Get BackColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    BackColor = lBackColor
    
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    lBackColor = NewValue
    PropertyChanged "BackColor"
    
    shpBack.FillColor = lBackColor
    
End Property

Public Property Get BorderColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    BorderColor = lBorderColor
    
End Property

Public Property Let BorderColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    lBorderColor = NewValue
    PropertyChanged "BorderColor"

    shpBack.BorderColor = lBorderColor
    
End Property

Public Property Let Value(ByVal NewValue As Long)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    If NewValue > MAX_LONG Then NewValue = MAX_LONG
    If NewValue < 0 Then NewValue = 0
    iValue = NewValue
    
    PropertyChanged "Value"
    
    Call DrawLabel
    
End Property

Public Property Get Value() As Long
'*************************************************
'Author: ^[GS]^
'Last modified: 30/04/2013
'*************************************************

On Error Resume Next
    
    Value = iValue
    
End Property
