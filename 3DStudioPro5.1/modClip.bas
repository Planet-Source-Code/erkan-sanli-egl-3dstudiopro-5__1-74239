Attribute VB_Name = "modClip"
Option Explicit

Public CLines() As CLIPLINE

Dim Pts As Byte

Public Function ClipTriangle() As CanStat
    
    Dim P1 As POINTAPI
    Dim P2 As POINTAPI
    Dim P3 As POINTAPI
    
    Pts = 0
    With g_Meshs(g_idxM)
        P1 = .Screen(.Faces(g_idxF).A)
        P2 = .Screen(.Faces(g_idxF).B)
        P3 = .Screen(.Faces(g_idxF).C)
    End With
    
'Face = Small, Canvas = Big
    IsInRectangle P1
    IsInRectangle P2
    IsInRectangle P3
    If Pts = 3 Then
        ClipTriangle = CanIn
        Exit Function
    End If

'Face = Big, Canvas = Small
    IsInTriangle P1, P2, P3, g_SafeFrame.Left, g_SafeFrame.Top
    IsInTriangle P1, P2, P3, g_SafeFrame.Right, g_SafeFrame.Top
    IsInTriangle P1, P2, P3, g_SafeFrame.Left, g_SafeFrame.Bottom
    IsInTriangle P1, P2, P3, g_SafeFrame.Right, g_SafeFrame.Bottom

'Face and Canvas intersection
    RectLineIntersection P1, P2
    RectLineIntersection P2, P3
    RectLineIntersection P3, P1

    If Pts > 0 Then
        ClipTriangle = CanClip
    Else
'Face out of Canvas, No intersection
        ClipTriangle = CanOut
    End If
    
End Function

Private Function IsInTriangle(P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI, PX As Long, PY As Long) As Boolean

    Dim Val1 As Single, Val2 As Single, Val3 As Single
    Dim K1 As Double, K2 As Double
    Dim F1 As Double, F2 As Double, F3 As Double, F4 As Double

    On Error GoTo Jump
    
    F1 = P1.X - PX
    F2 = P2.Y - PY
    F3 = P2.X - PX
    F4 = P1.Y - PY
    K1 = F1 * F2
    K2 = F3 * F4
    Val1 = CSng(K1 - K2)
    
    F1 = P2.X - PX
    F2 = P3.Y - PY
    F3 = P3.X - PX
    F4 = P2.Y - PY
    K1 = F1 * F2
    K2 = F3 * F4
    Val2 = CSng(K1 - K2)

    F1 = P3.X - PX
    F2 = P1.Y - PY
    F3 = P1.X - PX
    F4 = P3.Y - PY
    K1 = F1 * F2
    K2 = F3 * F4
    Val3 = CSng(K1 - K2)

    If (Val1 > 0 And Val2 > 0 And Val3 > 0) Or _
       (Val1 < 0 And Val2 < 0 And Val3 < 0) Then
        AddPoint
        IsInTriangle = True
    End If
    Exit Function

Jump:
    IsInTriangle = False
End Function

Private Function IsInRectangle(P1 As POINTAPI) As Boolean

    If P1.X > g_SafeFrame.Left And P1.X < g_SafeFrame.Right And _
       P1.Y > g_SafeFrame.Top And P1.Y < g_SafeFrame.Bottom Then
       AddPoint
       IsInRectangle = True
    End If

End Function

Private Sub RectLineIntersection(P1 As POINTAPI, P2 As POINTAPI)

    With g_SafeFrame
        LineLineIntersection .Left, .Top, .Right, .Top, P1, P2
        LineLineIntersection .Left, .Bottom, .Left, .Top, P1, P2
        LineLineIntersection .Right, .Top, .Right, .Bottom, P1, P2
        LineLineIntersection .Right, .Bottom, .Left, .Bottom, P1, P2
    End With

End Sub

Private Sub LineLineIntersection(P1X As Long, P1Y As Long, P2X As Long, P2Y As Long, P3 As POINTAPI, P4 As POINTAPI)

    Dim F1  As Long
    Dim F2  As Long
    Dim F3  As Long
    Dim F4  As Long
    Dim D1  As Single
    Dim D2  As Single
    Dim N1 As Double
    Dim N2 As Double

    On Error Resume Next

    F1 = P2X - P1X
    F2 = P2Y - P1Y
    F3 = P4.X - P3.X
    F4 = P4.Y - P3.Y

    N2 = F3 * F2 - F4 * F1
    If N2 = 0 Then Exit Sub
    N1 = F1 * (P3.Y - P1Y) + F2 * (P1X - P3.X)
    D1 = Div(CSng(N1), CSng(N2))

    N1 = F3 * (P1Y - P3.Y) + F4 * (P3.X - P1X)
    N2 = F4 * F1 - F3 * F2
    D2 = Div(CSng(N1), CSng(N2))
    If (D1 >= 0# And D1 <= 1# And D2 >= 0# And D2 <= 1#) Then AddPoint

End Sub

Private Sub AddPoint()
        
    Pts = Pts + 1
    
End Sub

Public Function ClipTriangleWireframe(Optional Box As Boolean = False) As CanStat
    
    Dim P1 As POINTAPI
    Dim P2 As POINTAPI
    Dim P3 As POINTAPI
    
    Pts = 0
    ReDim CLines(2)
    
    With g_Meshs(g_idxM)
        If Box = False Then
            P1 = .Screen(.Faces(g_idxF).A)
            P2 = .Screen(.Faces(g_idxF).B)
            P3 = .Screen(.Faces(g_idxF).C)
        Else
            P1 = .BoxScreen(.BoxFace(g_idxF).A)
            P2 = .BoxScreen(.BoxFace(g_idxF).B)
            P3 = .BoxScreen(.BoxFace(g_idxF).C)
        End If
    End With
        
'Face = Small, Canvas = Big
    IsInRectangle P1
    IsInRectangle P2
    IsInRectangle P3
    If Pts = 3 Then
        ClipTriangleWireframe = CanIn
        Exit Function
    End If

    ClipTriangleLine 0, P1, P2 'AB
    ClipTriangleLine 1, P2, P3 'BC
    ClipTriangleLine 2, P3, P1 'CA

    If Pts > 0 Then
'Face and Canvas intersection
        ClipTriangleWireframe = CanClip
    Else
'Face out of Canvas, No intersection
        ClipTriangleWireframe = CanOut
    End If
    
End Function

Private Sub ClipTriangleLine(idx As Byte, P1 As POINTAPI, P2 As POINTAPI)
    
    If IsInRectangle(P1) Then
        CLines(idx).P1 = P1
        If IsInRectangle(P2) Then
            CLines(idx).P2 = P2
            Exit Sub
        Else
            With g_SafeFrame
                LineLineIntersectionLine idx, .Left, .Top, .Right, .Top, P1, P2
                LineLineIntersectionLine idx, .Left, .Bottom, .Left, .Top, P1, P2
                LineLineIntersectionLine idx, .Right, .Top, .Right, .Bottom, P1, P2
                LineLineIntersectionLine idx, .Right, .Bottom, .Left, .Bottom, P1, P2
            End With
        End If
    ElseIf IsInRectangle(P2) Then
        CLines(idx).P1 = P2
        If IsInRectangle(P1) Then
            CLines(idx).P2 = P1
            Exit Sub
        Else
            With g_SafeFrame
                LineLineIntersectionLine idx, .Left, .Top, .Right, .Top, P1, P2
                LineLineIntersectionLine idx, .Left, .Bottom, .Left, .Top, P1, P2
                LineLineIntersectionLine idx, .Right, .Top, .Right, .Bottom, P1, P2
                LineLineIntersectionLine idx, .Right, .Bottom, .Left, .Bottom, P1, P2
            End With
        End If
    End If

End Sub

Private Sub LineLineIntersectionLine(idx As Byte, P1X As Long, P1Y As Long, P2X As Long, P2Y As Long, P3 As POINTAPI, P4 As POINTAPI)

    Dim F1  As Long
    Dim F2  As Long
    Dim F3  As Long
    Dim F4  As Long
    Dim D1  As Single
    Dim D2  As Single
    Dim N1 As Double
    Dim N2 As Double

    On Error Resume Next

    F1 = P2X - P1X
    F2 = P2Y - P1Y
    F3 = P4.X - P3.X
    F4 = P4.Y - P3.Y

    N2 = F3 * F2 - F4 * F1
    If N2 = 0 Then Exit Sub
    N1 = F1 * (P3.Y - P1Y) + F2 * (P1X - P3.X)
    D1 = Div(CSng(N1), CSng(N2))

    N1 = F3 * (P1Y - P3.Y) + F4 * (P3.X - P1X)
    N2 = F4 * F1 - F3 * F2
    D2 = Div(CSng(N1), CSng(N2))
    If (D1 >= 0# And D1 <= 1# And D2 >= 0# And D2 <= 1#) Then
        CLines(idx).P2.X = P1X + D2 * F1
        CLines(idx).P2.Y = P1Y + D2 * F2
    End If

End Sub

