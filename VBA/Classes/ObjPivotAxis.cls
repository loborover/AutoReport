Private vXaxis As Single
Private vYaxis As Single
Private vZaxis As Single
'초기화 이벤트 메서드
'Private Sub Class_Initialize()
'End Sub
'소멸 이벤트 메서드
'Private Sub Class_Terminate()
'End Sub
Public Property Let X(value As Single)
    vXaxis = value
End Property
Public Property Get X() As Single
    X = vXaxis
End Property

Public Property Let y(value As Single)
    vYaxis = value
End Property
Public Property Get y() As Single
    y = vYaxis
End Property

Public Property Let Z(value As Single)
    vZaxis = value
End Property
Public Property Get Z() As Single
    Z = vZaxis
End Property

Public Function Copy() As ObjPivotAxis
    Dim CopiedObj As New ObjPivotAxis
    
    With CopiedObj
        .X = Me.X
        .y = Me.y
        .Z = Me.Z
    End With
    Set Copy = CopiedObj
End Function