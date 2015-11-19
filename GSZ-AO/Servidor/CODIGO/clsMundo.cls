VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMundo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Diseñado y programado por ^[GS]^
'Original para GSZAO

Private Type tItemInfo
    ObjIndex As Integer
    TimeLeft As Integer
End Type

Private Const MAX_ITEMS = 1000
Private p_iteminfo(1 To MAX_ITEMS) As tItemInfo
Private p_items(1 To MAX_ITEMS) As WorldPos
Private p_cant As Integer

Public Sub Class_Initialize()
    p_cant = 0
End Sub

Public Sub Class_Terminate()
    ' limpiar variables?
End Sub

Public Property Get CantItems() As Integer
    CantItems = p_cant
End Property

Public Function AddItem(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal ObjIndex As Integer, ByVal TimeLeft As Integer) As Boolean
    Dim i As Integer
    Dim i_free As Integer
    
    AddItem = False

    If p_cant = MAX_ITEMS Then
        AddItem = False ' Esta al maximo de capacidad!
    Else
        For i = 1 To p_cant
            If Map = p_items(i).Map Then
                If X = p_items(i).X Then
                    If Y = p_items(i).Y Then
                        AddItem = True ' ya está!
                        Exit For
                    End If
                End If
            ElseIf p_items(i).Map = 0 Then
                If i_free <> 0 Then  ' posición libre!
                    i_free = i
                End If
            End If
        Next i
        If Not AddItem Then
            If i_free <> 0 Then
                i = i_free ' usar posición libre!
            Else
                p_cant = p_cant + 1 ' nuevo item!
                i = p_cant
            End If
            p_items(i).Map = Map
            p_items(i).X = X
            p_items(i).Y = Y
            p_iteminfo(i).ObjIndex = ObjIndex
            p_iteminfo(i).TimeLeft = TimeLeft
            AddItem = True
        End If
    
    End If
End Function

Public Function RemoveItem(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional Index As Integer = 0) As Boolean
    Dim i As Integer
    
    RemoveItem = False
    
    If Index > 0 Then
        If Index <= p_cant Then
            If p_items(Index).Map <> 0 Then
                If p_cant = Index Then p_cant = p_cant - 1
                p_items(Index).Map = 0
                p_items(Index).X = 0
                p_items(Index).Y = 0
                p_iteminfo(Index).ObjIndex = 0
                p_iteminfo(Index).TimeLeft = 0
                RemoveItem = True ' liberado
            End If
        End If
    Else
        For i = 1 To p_cant
            If Map = p_items(i).Map Then
                If X = p_items(i).X Then
                    If Y = p_items(i).Y Then
                        p_items(i).Map = 0
                        p_items(i).X = 0
                        p_items(i).Y = 0
                        p_iteminfo(i).ObjIndex = 0
                        p_iteminfo(i).TimeLeft = 0
                        RemoveItem = True ' liberado
                        Exit For
                    End If
                End If
            End If
        Next i
    End If

End Function

Public Function StepMinute() As Boolean
    Dim i As Integer
    Dim t_iteminfo As tItemInfo
    Dim t_obj As Obj
    Dim t_WorldPos As WorldPos
        
    StepMinute = False
    
    For i = 1 To p_cant
        If p_items(i).Map <> 0 Then
            With p_iteminfo(i)
                If .TimeLeft <= 1 Then
                    t_iteminfo = p_iteminfo(i)
                    t_obj.ObjIndex = t_iteminfo.ObjIndex
                    t_obj.Amount = 1
                    t_obj.ExtraLong = ObjData(t_iteminfo.ObjIndex).MaxHp
                    t_WorldPos = p_items(i)
                    If RemoveItem(0, 0, 0, i) Then
                        Call EraseObj(10000, t_WorldPos.Map, t_WorldPos.X, t_WorldPos.Y)
                        Call MakeObj(t_obj, t_WorldPos.Map, t_WorldPos.X, t_WorldPos.Y)
                    End If
                    p_items(i).Map = 0
                    p_items(i).X = 0
                    p_items(i).Y = 0
                ElseIf .TimeLeft > 1 Then
                    .TimeLeft = .TimeLeft - 1
                End If
            End With
        End If
    Next
    
    StepMinute = True

End Function
