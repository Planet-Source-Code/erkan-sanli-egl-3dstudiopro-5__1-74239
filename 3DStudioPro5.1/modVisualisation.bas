Attribute VB_Name = "modVisualisation"
Option Explicit

Public Const ApproachVal As Single = 0.000001
Private Const GRADIENT_FILL_TRIANGLE    As Long = &H2

Public Enum GradMode
    GRADIENT_FILL_RECT_H
    GRADIENT_FILL_RECT_V
End Enum

Type GRADIENT_RECT
    UpperLeft   As Long
    LowerRight  As Long
End Type

Private Type TRIVERTEX
    X           As Long
    Y           As Long
    Red         As Integer
    Green       As Integer
    Blue        As Integer
    Alpha       As Integer
End Type

Private Type GRADIENT_TRIANGLE
    Vertex1     As Long
    Vertex2     As Long
    Vertex3     As Long
End Type

Private Type GRATEXEL
    Y1          As Long
    Y2          As Long
    M1          As MAPCOORD
    M2          As MAPCOORD
    C1          As COLORRGB_SNG
    C2          As COLORRGB_SNG
    Used        As Boolean
End Type

Private Declare Function CreatePen Lib "gdi32" ( _
                        ByVal nPenStyle As Long, _
                        ByVal nWidth As Long, _
                        ByVal crColor As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" ( _
                        ByVal crColor As Long) As Long

Private Declare Function SelectObject Lib "gdi32" ( _
                        ByVal hDC As Long, _
                        ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" ( _
                        ByVal hObject As Long) As Long

Private Declare Function Polygon Lib "gdi32" ( _
                        ByVal hDC As Long, _
                        lpPoint As POINTAPI, _
                        ByVal nCount As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" ( _
                        ByVal hDC As Long, _
                        ByVal X As Long, _
                        ByVal Y As Long, _
                        lpPoint As POINTAPI) As Long

Private Declare Function LineTo Lib "gdi32" ( _
                        ByVal hDC As Long, _
                        ByVal X As Long, _
                        ByVal Y As Long) As Long

Public Declare Function BitBlt Lib "gdi32" ( _
                        ByVal hDestDC As Long, _
                        ByVal X As Long, _
                        ByVal Y As Long, _
                        ByVal nWidth As Long, _
                        ByVal nHeight As Long, _
                        ByVal hSrcDC As Long, _
                        ByVal xSrc As Long, _
                        ByVal ySrc As Long, _
                        ByVal dwRop As Long) As Long

Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" ( _
                        ByVal hDC As Long, pVertex As TRIVERTEX, _
                        ByVal dwNumVertex As Long, _
                        pMesh As GRADIENT_TRIANGLE, _
                        ByVal dwNumMesh As Long, _
                        ByVal dwMode As Long) As Long

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" ( _
                        ByVal hDC As Long, pVertex As TRIVERTEX, _
                        ByVal dwNumVertex As Long, _
                        pMesh As GRADIENT_RECT, _
                        ByVal dwNumMesh As Long, _
                        ByVal dwMode As Long) As Long

Private triVert(2) As TRIVERTEX

Public g_idxM   As Integer
Public g_idxF   As Long
Public g_idxMt  As Integer
Public g_FColor As RGBQUAD

Private minX        As Long
Private maxX        As Long
Private minY        As Long
Private maxY        As Long
Private StepU       As Single
Private StepV       As Single
Private StepR       As Single
Private StepG       As Single
Private StepB       As Single
Dim Gratexels()    As GRATEXEL

Public Sub Render()
        
    Select Case g_VStyle
        
        Case Dot
            ViewDot
        
        Case Box
            ViewBox
            
        Case Wireframe
            If g_blnDoubleSide Then
                ViewWireframe
            Else
                If SortFaces(True) >= 0 Then ViewWireframeHidden
            End If

        Case WireframeMap
            ViewWireframeMap

        Case Flat
            If g_blnDoubleSide Then
                If SortFaces(False) >= 0 Then ViewFlat
            End If
            If SortFaces(True) >= 0 Then ViewFlat

        Case FlatMap
            If g_blnDoubleSide Then
                If SortFaces(False) >= 0 Then ViewFlatMap
            End If
            If SortFaces(True) >= 0 Then ViewFlatMap
        
        Case Gouraud
            If g_blnDoubleSide Then
                If SortFaces(False) >= 0 Then ViewGouraud
            End If
            If SortFaces(True) >= 0 Then ViewGouraud

        Case GouraudMap
            If g_blnDoubleSide Then
                If SortFaces(False) >= 0 Then ViewGouraudMap
            End If
            If SortFaces(True) >= 0 Then ViewGouraudMap
            
    End Select

End Sub

Public Sub RenderShadow()
    
    g_FColor.rgbRed = g_Lights(0).ShadowColor.R
    g_FColor.rgbGreen = g_Lights(0).ShadowColor.G
    g_FColor.rgbBlue = g_Lights(0).ShadowColor.B
    For g_idxM = 0 To g_idxMesh
        For g_idxF = 0 To g_Meshs(g_idxM).idxFace
            If IsInCamera Then
                Select Case ClipTriangle
                    Case CanIn:     DrawTriangleFlatAPI
                    Case CanClip:   DrawTriangleFlat
                End Select
            End If
        Next g_idxF
    Next g_idxM

End Sub

'======================================================================================
'Start Views

Private Sub ViewDot()

    Dim idxPoint        As Long
    Dim byteRGBColor    As RGBQUAD

    GetFaceColor
    For g_idxM = 0 To g_idxMesh
        With g_Meshs(g_idxM)
            For idxPoint = 0 To .idxVert
                If PointInCanvas(.Screen(idxPoint)) And ZValInCamera(.Vertices(idxPoint).VectorsT.Z) Then
                    byteRGBColor = ColIntToColByte(.Vertices(idxPoint).intRGBColor)
                    g_dibCanvas.uBI.Bits(.Screen(idxPoint).X, .Screen(idxPoint).Y) = byteRGBColor
                    g_dibCanvas.uBI.Bits(.Screen(idxPoint).X + 1, .Screen(idxPoint).Y) = byteRGBColor
                    g_dibCanvas.uBI.Bits(.Screen(idxPoint).X + 1, .Screen(idxPoint).Y + 1) = byteRGBColor
                    g_dibCanvas.uBI.Bits(.Screen(idxPoint).X, .Screen(idxPoint).Y + 1) = byteRGBColor
                End If
            Next idxPoint
        End With
    Next g_idxM

End Sub

Private Sub ViewBox()

    GetFaceColor
    For g_idxM = 0 To g_idxMesh
        For g_idxF = 0 To UBound(g_Meshs(g_idxM).BoxFace)
            If IsInCamera Then
                g_FColor = ColorAverage2
                Select Case ClipTriangleWireframe(True)
                    Case CanIn:     DrawTriangleBoxAPI
                    Case CanClip:   DrawTriangleBoxClip
                End Select
            End If
        Next g_idxF
    Next g_idxM

End Sub

Private Sub ViewWireframe()
    
    GetFaceColor
    For g_idxM = 0 To g_idxMesh
        For g_idxF = 0 To g_Meshs(g_idxM).idxFace
            If IsInCamera Then DrawWireframe
        Next g_idxF
    Next g_idxM
    
End Sub

Private Sub ViewWireframeHidden()

    Dim idxOrder As Long
    Dim OldVal As Single
    
    GetFaceColorO
    For idxOrder = 0 To UBound(g_MeshsOrder)
        g_idxM = g_MeshsOrder(idxOrder).idxMeshO
        g_idxF = g_MeshsOrder(idxOrder).idxFaceO
        g_idxMt = g_Meshs(g_idxM).Faces(g_idxF).idxMat
        If IsInCamera Then
            OldVal = g_Materials(g_idxMt).OpcValue1
            g_Materials(g_idxMt).OpcValue1 = 0
            DrawTriangleFlat
            DrawWireframe
            g_Materials(g_idxMt).OpcValue1 = OldVal
        End If
    Next idxOrder

End Sub

Private Sub ViewWireframeMap()
    
    GetFaceColor
    For g_idxM = 0 To g_idxMesh
        With g_Meshs(g_idxM)
            For g_idxF = 0 To .idxFace
                g_idxMt = .Faces(g_idxF).idxMat
                If IsInCamera Then
                    If .MapCoorsOK And g_Materials(g_idxMt).MapUse Then
                        If ClipTriangleWireframe <> CanOut Then DrawTriangleWireframeMap
                    Else
                        DrawWireframe
                    End If
                End If
            Next g_idxF
        End With
    Next g_idxM
    
End Sub

Private Sub ViewFlat()
    
    Dim idxOrder As Long
        
    GetFaceColorO
    For idxOrder = 0 To UBound(g_MeshsOrder)
        g_idxM = g_MeshsOrder(idxOrder).idxMeshO
        g_idxF = g_MeshsOrder(idxOrder).idxFaceO
        g_idxMt = g_Meshs(g_idxM).Faces(g_idxF).idxMat
        If IsInCamera Then DrawFlat
    Next idxOrder
    
End Sub

Private Sub ViewFlatMap()

    Dim idxOrder    As Long
    
    GetFaceColorO
    For idxOrder = 0 To UBound(g_MeshsOrder)
        g_idxM = g_MeshsOrder(idxOrder).idxMeshO
        g_idxF = g_MeshsOrder(idxOrder).idxFaceO
        g_idxMt = g_Meshs(g_idxM).Faces(g_idxF).idxMat
        If IsInCamera Then
            If g_Meshs(g_idxM).MapCoorsOK And g_Materials(g_idxMt).MapUse Then
                If ClipTriangle <> CanOut Then
                    g_FColor = ColorAverage2
                    DrawTriangleFlatMap
                End If
            Else
                DrawFlat
            End If
        End If
    Next idxOrder
    
End Sub

Private Sub ViewGouraud()

    Dim idxOrder As Long
    
    GetFaceColorO
    For idxOrder = 0 To UBound(g_MeshsOrder)
        g_idxM = g_MeshsOrder(idxOrder).idxMeshO
        g_idxF = g_MeshsOrder(idxOrder).idxFaceO
        g_idxMt = g_Meshs(g_idxM).Faces(g_idxF).idxMat
        If IsInCamera Then DrawGouraud
    Next idxOrder

End Sub

Private Sub ViewGouraudMap()

    Dim idxOrder    As Long
    
    GetFaceColorO
    For idxOrder = 0 To UBound(g_MeshsOrder)
        g_idxM = g_MeshsOrder(idxOrder).idxMeshO
        g_idxF = g_MeshsOrder(idxOrder).idxFaceO
        g_idxMt = g_Meshs(g_idxM).Faces(g_idxF).idxMat
        If IsInCamera Then
            If g_Meshs(g_idxM).MapCoorsOK And g_Materials(g_idxMt).MapUse Then
                If ClipTriangle <> CanOut Then DrawTriangleGradMap
            Else
                DrawGouraud
            End If
        End If
    Next
    
End Sub
'End Views
'=====================================================================================
'
'=====================================================================================
'Start Draws

Private Sub DrawLine(P1 As POINTAPI, P2 As POINTAPI, lColor As Long)
    
    Dim Pen     As Long
    Dim pAPI    As POINTAPI

    Pen = SelectObject(g_dibCanvas.hDC, CreatePen(0, 1, lColor))
    MoveToEx g_dibCanvas.hDC, P1.X, P1.Y, pAPI
    LineTo g_dibCanvas.hDC, P2.X, P2.Y
    Call DeleteObject(Pen)

End Sub

Private Sub DrawLineMap(P1 As POINTAPI, P2 As POINTAPI, M1 As MAPCOORD, M2 As MAPCOORD)

    Dim DeltaX As Single
    
    Dim X1 As Long
    Dim Y1 As Single
    Dim X2 As Long
    Dim Y2 As Single
    Dim U1 As Single
    Dim V1 As Single
    Dim U2 As Single
    Dim V2 As Single
    
    Dim DeltaY As Single
    Dim StartX  As Single, StartY   As Single
    Dim EndX    As Single, EndY     As Single
    Dim StepX   As Single, StepY    As Single
    Dim absDeltaY As Single
    Dim idx     As Long
    Dim DeltaU  As Single, DeltaV   As Single
    Dim StartU  As Single, StartV   As Single
    Dim EndU    As Single, EndV     As Single
    X1 = P1.X
    X2 = P2.X
    Y1 = P1.Y
    Y2 = P2.Y
    
    U1 = M1.U
    U2 = M2.U
    V1 = M1.V
    V2 = M2.V
    
    If X1 < X2 Then
        StartX = X1:    EndX = X2
        StartY = Y1:    EndY = Y2
        StartU = U1:    EndU = U2
        StartV = V1:    EndV = V2
    Else
        StartX = X2:    EndX = X1
        StartY = Y2:    EndY = Y1
        StartU = U2:    EndU = U1
        StartV = V2:    EndV = V1
    End If

    DeltaX = EndX - StartX
    DeltaY = EndY - StartY
    DeltaU = EndU - StartU
    DeltaV = EndV - StartV
    If DeltaX > Abs(DeltaY) Then
        StepY = Div(DeltaY, DeltaX)
        StepU = Div(DeltaU, DeltaX)
        StepV = Div(DeltaV, DeltaX)
        For StartX = StartX To EndX
            If StartX > g_SafeFrame.Left And StartX < g_SafeFrame.Right And StartY > g_SafeFrame.Top And StartY < g_SafeFrame.Bottom Then
                Call FindTexel(StartU, StartV)
                g_dibCanvas.uBI.Bits(StartX, StartY) = g_Materials(g_idxMt).Texture.dibTexT.uBI.Bits(Fix(StartU), Fix(StartV))
            End If
            StartY = StartY + StepY
            StartU = StartU + StepU
            StartV = StartV + StepV
        Next
    Else
        absDeltaY = Abs(DeltaY)
        StepX = Div(DeltaX, absDeltaY)
        StepY = Div(DeltaY, absDeltaY)
        StepU = Div(DeltaU, absDeltaY)
        StepV = Div(DeltaV, absDeltaY)
        For idx = 0 To absDeltaY
            If StartX > g_SafeFrame.Left And StartX < g_SafeFrame.Right And StartY > g_SafeFrame.Top And StartY < g_SafeFrame.Bottom Then
                Call FindTexel(StartU, StartV)
                g_dibCanvas.uBI.Bits(StartX, StartY) = g_Materials(g_idxMt).Texture.dibTexT.uBI.Bits(Fix(StartU), Fix(StartV))
            End If
            StartX = StartX + StepX
            StartY = StartY + StepY
            StartU = StartU + StepU
            StartV = StartV + StepV
        Next
    End If

End Sub

'Private Sub DrawSafeFrame()
'
'    Dim P1  As POINTAPI
'    Dim P2  As POINTAPI
'    Dim Col As Long
'
'    Col = RGB(200, 200, 250)
'
'    P1.X = g_SafeFrame.Left:   P1.Y = g_SafeFrame.Top
'    P2.X = g_SafeFrame.Right:   P2.Y = g_SafeFrame.Top
'    Call DrawLine(P1, P2, Col)
'
'    P1.X = g_SafeFrame.Right:   P1.Y = g_SafeFrame.Top
'    P2.X = g_SafeFrame.Right:   P2.Y = g_SafeFrame.Bottom
'    Call DrawLine(P1, P2, Col)
'
'    P1.X = g_SafeFrame.Right:   P1.Y = g_SafeFrame.Bottom
'    P2.X = g_SafeFrame.Left:   P2.Y = g_SafeFrame.Bottom
'    Call DrawLine(P1, P2, Col)
'
'    P1.X = g_SafeFrame.Left:   P1.Y = g_SafeFrame.Bottom
'    P2.X = g_SafeFrame.Left:   P2.Y = g_SafeFrame.Top
'    Call DrawLine(P1, P2, Col)
'
'End Sub

Private Sub DrawTriangleBoxAPI()
    
    Dim lColor      As Long
    
    lColor = RGB(g_FColor.rgbRed, g_FColor.rgbGreen, g_FColor.rgbBlue)
    With g_Meshs(g_idxM)
        If .BoxFace(g_idxF).AB Then
            Call DrawLine(.BoxScreen(.BoxFace(g_idxF).A), .BoxScreen(.BoxFace(g_idxF).B), lColor)
        End If
        If .BoxFace(g_idxF).BC Then
            Call DrawLine(.BoxScreen(.BoxFace(g_idxF).B), .BoxScreen(.BoxFace(g_idxF).C), lColor)
        End If
        If .BoxFace(g_idxF).CA Then
            Call DrawLine(.BoxScreen(.BoxFace(g_idxF).C), .BoxScreen(.BoxFace(g_idxF).A), lColor)
        End If
    End With
    
End Sub

Private Sub DrawTriangleBoxClip()
    
    Dim lColor      As Long
    
    lColor = RGB(g_FColor.rgbRed, g_FColor.rgbGreen, g_FColor.rgbBlue)
    With g_Meshs(g_idxM)
        If .BoxFace(g_idxF).AB Then
            Call DrawLine(CLines(0).P1, CLines(0).P2, lColor)
        End If
        If .BoxFace(g_idxF).BC Then
            Call DrawLine(CLines(1).P1, CLines(1).P2, lColor)
        End If
        If .BoxFace(g_idxF).CA Then
            Call DrawLine(CLines(2).P1, CLines(2).P2, lColor)
        End If
    End With
    
End Sub

Private Sub DrawWireframe()
    
    g_FColor = ColorAverage2
    Select Case ClipTriangleWireframe
        Case CanIn:     DrawTriangleWireframeAPI
        Case CanClip:   DrawTriangleWireframeClip
    End Select
  
End Sub

Private Sub DrawTriangleWireframeAPI()
    
    Dim lColor      As Long
    
    lColor = RGB(g_FColor.rgbRed, g_FColor.rgbGreen, g_FColor.rgbBlue)
    With g_Meshs(g_idxM)
        If .Faces(g_idxF).AB Then
            Call DrawLine(.Screen(.Faces(g_idxF).A), .Screen(.Faces(g_idxF).B), lColor)
        End If
        If .Faces(g_idxF).BC Then
            Call DrawLine(.Screen(.Faces(g_idxF).B), .Screen(.Faces(g_idxF).C), lColor)
        End If
        If .Faces(g_idxF).CA Then
            Call DrawLine(.Screen(.Faces(g_idxF).C), .Screen(.Faces(g_idxF).A), lColor)
        End If
    End With
    
End Sub

Private Sub DrawTriangleWireframeClip()
    
    Dim lColor      As Long
    
    lColor = RGB(g_FColor.rgbRed, g_FColor.rgbGreen, g_FColor.rgbBlue)
    With g_Meshs(g_idxM)
        If .Faces(g_idxF).AB Then
            Call DrawLine(CLines(0).P1, CLines(0).P2, lColor)
        End If
        If .Faces(g_idxF).BC Then
            Call DrawLine(CLines(1).P1, CLines(1).P2, lColor)
        End If
        If .Faces(g_idxF).CA Then
            Call DrawLine(CLines(2).P1, CLines(2).P2, lColor)
        End If
    End With
    
End Sub

Private Sub DrawTriangleWireframeMap()
    
    Dim P1 As POINTAPI
    Dim P2 As POINTAPI
    Dim P3 As POINTAPI
    Dim M1 As MAPCOORD
    Dim M2 As MAPCOORD
    Dim M3 As MAPCOORD
    
    With g_Meshs(g_idxM)
        CalculateUV
        M1 = .TScreen(.Faces(g_idxF).A)
        M2 = .TScreen(.Faces(g_idxF).B)
        M3 = .TScreen(.Faces(g_idxF).C)
    End With
    
    RedimGratexels P1, P2, P3
    With g_Meshs(g_idxM)
        If .Faces(g_idxF).AB Then
            Call DrawLineMap(P1, P2, M1, M2)
        End If
        If .Faces(g_idxF).BC Then
            Call DrawLineMap(P2, P3, M2, M3)
        End If
        If .Faces(g_idxF).CA Then
            Call DrawLineMap(P3, P1, M3, M1)
        End If
    End With
    
End Sub

Private Sub DrawFlat()
    
    g_FColor = ColorAverage2
    If g_Materials(g_idxMt).OpcValue1 = 1 Then
        Select Case ClipTriangle
            Case CanIn:     DrawTriangleFlatAPI
            Case CanClip:   DrawTriangleFlat
        End Select
    Else
        If ClipTriangle <> CanOut Then DrawTriangleFlat
    End If

End Sub

Private Sub DrawTriangleFlatAPI()
    
    Dim pAPI(2)     As POINTAPI
    Dim lColor      As Long
    Dim Pen         As Long
    Dim Brush       As Long
    
    lColor = RGBQuadToLong(g_FColor)
    Pen = SelectObject(g_dibCanvas.hDC, CreatePen(0, 1, lColor))
    Brush = SelectObject(g_dibCanvas.hDC, CreateSolidBrush(lColor))
    With g_Meshs(g_idxM)
        pAPI(0) = .Screen(.Faces(g_idxF).A)
        pAPI(1) = .Screen(.Faces(g_idxF).B)
        pAPI(2) = .Screen(.Faces(g_idxF).C)
    End With
    Polygon g_dibCanvas.hDC, pAPI(0), 3
    DeleteObject Pen
    DeleteObject Brush
                
End Sub

Private Sub DrawTriangleFlat()
    
    Dim P1      As POINTAPI
    Dim P2      As POINTAPI
    Dim P3      As POINTAPI
    
    RedimGratexels P1, P2, P3
    InterpolateFlat P1, P2
    InterpolateFlat P2, P3
    InterpolateFlat P3, P1
    For minX = minX To maxX
        FillFlat
    Next

End Sub

Private Sub DrawTriangleFlatMap()
    
    Dim P1      As POINTAPI
    Dim P2      As POINTAPI
    Dim P3      As POINTAPI
    Dim M1      As MAPCOORD
    Dim M2      As MAPCOORD
    Dim M3      As MAPCOORD

    With g_Meshs(g_idxM)
        CalculateUV
        M1 = .TScreen(.Faces(g_idxF).A)
        M2 = .TScreen(.Faces(g_idxF).B)
        M3 = .TScreen(.Faces(g_idxF).C)
    End With
    
    RedimGratexels P1, P2, P3
    InterpolateFlaTex P1, P2, M1, M2
    InterpolateFlaTex P2, P3, M2, M3
    InterpolateFlaTex P3, P1, M3, M1
    For minX = minX To maxX
        FillFlaTex
    Next

End Sub

Private Sub DrawGouraud()

    If g_Materials(g_idxMt).OpcValue1 = 1 Then
        Select Case ClipTriangle
            Case CanIn:     DrawTriangleGradientAPI
            Case CanClip:   DrawTriangleGradient
        End Select
    Else
        If ClipTriangle <> CanOut Then DrawTriangleGradient
    End If

End Sub

Private Sub DrawTriangleGradientAPI()

    Dim vert(2) As TRIVERTEX
    Dim gTri    As GRADIENT_TRIANGLE
    
    With g_Meshs(g_idxM)
    
        vert(0).X = .Screen(.Faces(g_idxF).A).X
        vert(0).Y = .Screen(.Faces(g_idxF).A).Y
        vert(0).Red = ConvertUShort(.Vertices(.Faces(g_idxF).A).intRGBColor.R)
        vert(0).Green = ConvertUShort(.Vertices(.Faces(g_idxF).A).intRGBColor.G)
        vert(0).Blue = ConvertUShort(.Vertices(.Faces(g_idxF).A).intRGBColor.B)
        
        vert(1).X = .Screen(.Faces(g_idxF).B).X
        vert(1).Y = .Screen(.Faces(g_idxF).B).Y
        vert(1).Red = ConvertUShort(.Vertices(.Faces(g_idxF).B).intRGBColor.R)
        vert(1).Green = ConvertUShort(.Vertices(.Faces(g_idxF).B).intRGBColor.G)
        vert(1).Blue = ConvertUShort(.Vertices(.Faces(g_idxF).B).intRGBColor.B)
    
        vert(2).X = .Screen(.Faces(g_idxF).C).X
        vert(2).Y = .Screen(.Faces(g_idxF).C).Y
        vert(2).Red = ConvertUShort(.Vertices(.Faces(g_idxF).C).intRGBColor.R)
        vert(2).Green = ConvertUShort(.Vertices(.Faces(g_idxF).C).intRGBColor.G)
        vert(2).Blue = ConvertUShort(.Vertices(.Faces(g_idxF).C).intRGBColor.B)
    
    End With
    gTri.Vertex1 = 0
    gTri.Vertex2 = 1
    gTri.Vertex3 = 2
    Call GradientFillTriangle(g_dibCanvas.hDC, vert(0), 3, gTri, 1, GRADIENT_FILL_TRIANGLE)

End Sub

Public Sub DrawTriangleGradient()
    
    Dim P1      As POINTAPI
    Dim P2      As POINTAPI
    Dim P3      As POINTAPI
    Dim C1      As COLORRGB_SNG
    Dim C2      As COLORRGB_SNG
    Dim C3      As COLORRGB_SNG

    With g_Meshs(g_idxM)
        C1 = ColIntToColSng(.Vertices(.Faces(g_idxF).A).intRGBColor)
        C2 = ColIntToColSng(.Vertices(.Faces(g_idxF).B).intRGBColor)
        C3 = ColIntToColSng(.Vertices(.Faces(g_idxF).C).intRGBColor)
    End With
    
    RedimGratexels P1, P2, P3
    InterpolateGra P1, P2, C1, C2
    InterpolateGra P2, P3, C2, C3
    InterpolateGra P3, P1, C3, C1
    For minX = minX To maxX
        FillGra
    Next

End Sub

Private Sub DrawTriangleGradMap()
    
    Dim P1      As POINTAPI
    Dim P2      As POINTAPI
    Dim P3      As POINTAPI
    Dim M1      As MAPCOORD
    Dim M2      As MAPCOORD
    Dim M3      As MAPCOORD
    Dim C1      As COLORRGB_SNG
    Dim C2      As COLORRGB_SNG
    Dim C3      As COLORRGB_SNG

    With g_Meshs(g_idxM)
        CalculateUV
        M1 = .TScreen(.Faces(g_idxF).A)
        M2 = .TScreen(.Faces(g_idxF).B)
        M3 = .TScreen(.Faces(g_idxF).C)
        C1 = ColIntToColSng(.Vertices(.Faces(g_idxF).A).intRGBColor)
        C2 = ColIntToColSng(.Vertices(.Faces(g_idxF).B).intRGBColor)
        C3 = ColIntToColSng(.Vertices(.Faces(g_idxF).C).intRGBColor)
    End With
    
    RedimGratexels P1, P2, P3
    InterpolateGraTex P1, P2, M1, M2, C1, C2
    InterpolateGraTex P2, P3, M2, M3, C2, C3
    InterpolateGraTex P3, P1, M3, M1, C3, C1
    For minX = minX To maxX
        FillGraTex
    Next

End Sub
' End Draws
'==============================================================================================

Private Sub RedimGratexels(P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI)

'Get points values
    With g_Meshs(g_idxM)
        P1 = .Screen(.Faces(g_idxF).A)
        P2 = .Screen(.Faces(g_idxF).B)
        P3 = .Screen(.Faces(g_idxF).C)
    End With
'Redim Gratexels
    minX = IIf(P1.X < P2.X, P1.X, P2.X)
    If P3.X < minX Then minX = P3.X
    maxX = IIf(P1.X > P2.X, P1.X, P2.X)
    If P3.X > maxX Then maxX = P3.X
    ReDim Gratexels(minX To maxX)
    If minX < g_SafeFrame.Left Then minX = g_SafeFrame.Left
    If maxX > g_SafeFrame.Right Then maxX = g_SafeFrame.Right

End Sub

Private Function ConvertUShort(Color As Integer) As Integer
    
    Dim Unsigned As Long
    
    Unsigned = Color * 256&
    If Unsigned < &H8000& Then
        ConvertUShort = CInt(Unsigned)
    Else
        ConvertUShort = CInt(Unsigned - &H10000)
    End If
        
End Function

Private Sub InterpolateFlat(P1 As POINTAPI, P2 As POINTAPI)

    Dim DeltaX  As Long
    Dim X1      As Long
    Dim X2      As Long
    Dim Y1      As Single
    Dim Y2      As Single
    Dim StepY   As Single
    
    X1 = P1.X
    X2 = P2.X
    Y1 = P1.Y
    Y2 = P2.Y

    If X1 < X2 Then
        DeltaX = X2 - X1
        StepY = Div(Y2 - Y1, DeltaX)
        For X1 = X1 To X2
            With Gratexels(X1)
                If .Used Then
                    If .Y1 < Y1 Then .Y1 = Y1
                    If .Y2 > Y1 Then .Y2 = Y1
                Else
                    .Y1 = Y1
                    .Y2 = Y1
                    .Used = True
                End If
            End With
            Y1 = Y1 + StepY
        Next
    Else
        DeltaX = X1 - X2
        StepY = Div(Y1 - Y2, DeltaX)
        For X2 = X2 To X1
            With Gratexels(X2)
                If .Used Then
                    If .Y1 < Y2 Then .Y1 = Y2
                    If .Y2 > Y2 Then .Y2 = Y2
                Else
                    .Y1 = Y2
                    .Y2 = Y2
                    .Used = True
                End If
            End With
            Y2 = Y2 + StepY
        Next
   End If
End Sub

Private Sub InterpolateGra(P1 As POINTAPI, P2 As POINTAPI, _
                           C1 As COLORRGB_SNG, C2 As COLORRGB_SNG)
                                    
    Dim DeltaX  As Long
    Dim X1      As Long
    Dim X2      As Long
    Dim Y1      As Single
    Dim Y2      As Single
    Dim CC1     As COLORRGB_SNG
    Dim CC2     As COLORRGB_SNG
    Dim StepY   As Single
    Dim StepR   As Single
    Dim StepG   As Single
    Dim StepB   As Single

    X1 = P1.X
    X2 = P2.X
    Y1 = P1.Y
    Y2 = P2.Y
    CC1 = C1
    CC2 = C2
    
    If X1 < X2 Then
        DeltaX = X2 - X1
        StepY = Div(Y2 - Y1, DeltaX)
        StepR = Div(CC2.R - CC1.R, DeltaX)
        StepG = Div(CC2.G - CC1.G, DeltaX)
        StepB = Div(CC2.B - CC1.B, DeltaX)
        For X1 = X1 To X2
            With Gratexels(X1)
                If .Used Then
                    If .Y1 < Y1 Then .Y1 = Y1: .C1 = CC1
                    If .Y2 > Y1 Then .Y2 = Y1: .C2 = CC1
                Else
                    .Y1 = Y1:  .C1 = CC1
                    .Y2 = Y1:  .C2 = CC1
                    .Used = True
                End If
            End With
            Y1 = Y1 + StepY
            CC1.R = CC1.R + StepR
            CC1.G = CC1.G + StepG
            CC1.B = CC1.B + StepB
        Next
    Else
        DeltaX = X1 - X2
        StepY = Div(Y1 - Y2, DeltaX)
        StepR = Div(CC1.R - CC2.R, DeltaX)
        StepG = Div(CC1.G - CC2.G, DeltaX)
        StepB = Div(CC1.B - CC2.B, DeltaX)
        For X2 = X2 To X1
            With Gratexels(X2)
                If .Used Then
                    If .Y1 < Y2 Then .Y1 = Y2:  .C1 = CC2
                    If .Y2 > Y2 Then .Y2 = Y2:  .C2 = CC2
                Else
                    .Y1 = Y2: .C1 = CC2
                    .Y2 = Y2: .C2 = CC2
                    .Used = True
                End If
            End With
            Y2 = Y2 + StepY
            CC2.R = CC2.R + StepR
            CC2.G = CC2.G + StepG
            CC2.B = CC2.B + StepB
        Next
   End If
End Sub

Private Sub InterpolateFlaTex(P1 As POINTAPI, P2 As POINTAPI, _
                              M1 As MAPCOORD, M2 As MAPCOORD)
                              
    Dim DeltaX  As Long
    Dim X1      As Long
    Dim X2      As Long
    Dim Y1      As Single
    Dim Y2      As Single
    Dim MM1     As MAPCOORD
    Dim MM2     As MAPCOORD
    Dim StepY   As Single
    Dim StepU   As Single
    Dim StepV   As Single
    
    X1 = P1.X
    X2 = P2.X
    Y1 = P1.Y
    Y2 = P2.Y
    MM1 = M1
    MM2 = M2

    If X1 < X2 Then
        DeltaX = X2 - X1
        StepY = Div(Y2 - Y1, DeltaX)
        StepU = Div(MM2.U - MM1.U, DeltaX)
        StepV = Div(MM2.V - MM1.V, DeltaX)
        For X1 = X1 To X2
            With Gratexels(X1)
                If .Used Then
                    If .Y1 < Y1 Then .Y1 = Y1: .M1 = MM1
                    If .Y2 > Y1 Then .Y2 = Y1: .M2 = MM1
                Else
                    .Y1 = Y1: .M1 = MM1
                    .Y2 = Y1: .M2 = MM1
                    .Used = True
                End If
            End With
            Y1 = Y1 + StepY
            MM1.U = MM1.U + StepU
            MM1.V = MM1.V + StepV
        Next
    Else
        DeltaX = X1 - X2
        StepY = Div(Y1 - Y2, DeltaX)
        StepU = Div(MM1.U - MM2.U, DeltaX)
        StepV = Div(MM1.V - MM2.V, DeltaX)
        For X2 = X2 To X1
            With Gratexels(X2)
                If .Used Then
                    If .Y1 < Y2 Then .Y1 = Y2: .M1 = MM2
                    If .Y2 > Y2 Then .Y2 = Y2: .M2 = MM2
                Else
                    .Y1 = Y2: .M1 = MM2
                    .Y2 = Y2: .M2 = MM2
                    .Used = True
                End If
            End With
            Y2 = Y2 + StepY
            MM2.U = MM2.U + StepU
            MM2.V = MM2.V + StepV
        Next
   End If
End Sub

Private Sub InterpolateGraTex(P1 As POINTAPI, P2 As POINTAPI, _
                              M1 As MAPCOORD, M2 As MAPCOORD, _
                              C1 As COLORRGB_SNG, C2 As COLORRGB_SNG)
                                    
    Dim DeltaX  As Long
    Dim X1      As Long
    Dim X2      As Long
    Dim Y1      As Single
    Dim Y2      As Single
    Dim MM1     As MAPCOORD
    Dim MM2     As MAPCOORD
    Dim CC1     As COLORRGB_SNG
    Dim CC2     As COLORRGB_SNG
    Dim StepY   As Single
    Dim StepU   As Single
    Dim StepV   As Single
    Dim StepR   As Single
    Dim StepG   As Single
    Dim StepB   As Single

    X1 = P1.X
    X2 = P2.X
    Y1 = P1.Y
    Y2 = P2.Y
    MM1 = M1
    MM2 = M2
    CC1 = C1
    CC2 = C2

    If X1 < X2 Then
        DeltaX = X2 - X1
        StepY = Div(Y2 - Y1, DeltaX)
        StepU = Div(MM2.U - MM1.U, DeltaX)
        StepV = Div(MM2.V - MM1.V, DeltaX)
        StepR = Div(CC2.R - CC1.R, DeltaX)
        StepG = Div(CC2.G - CC1.G, DeltaX)
        StepB = Div(CC2.B - CC1.B, DeltaX)
        
        For X1 = X1 To X2
            With Gratexels(X1)
                If .Used Then
                    If .Y1 < Y1 Then .Y1 = Y1: .M1 = MM1: .C1 = CC1
                    If .Y2 > Y1 Then .Y2 = Y1: .M2 = MM1: .C2 = CC1
                Else
                    .Y1 = Y1: .M1 = MM1: .C1 = CC1
                    .Y2 = Y1: .M2 = MM1: .C2 = CC1
                    .Used = True
                End If
            End With
            Y1 = Y1 + StepY
            MM1.U = MM1.U + StepU
            MM1.V = MM1.V + StepV
            CC1.R = CC1.R + StepR
            CC1.G = CC1.G + StepG
            CC1.B = CC1.B + StepB
        Next
    Else
        DeltaX = X1 - X2
        StepY = Div(Y1 - Y2, DeltaX)
        StepU = Div(MM1.U - MM2.U, DeltaX)
        StepV = Div(MM1.V - MM2.V, DeltaX)
        StepR = Div(CC1.R - CC2.R, DeltaX)
        StepG = Div(CC1.G - CC2.G, DeltaX)
        StepB = Div(CC1.B - CC2.B, DeltaX)

        For X2 = X2 To X1
            With Gratexels(X2)
                If .Used Then
                    If .Y1 < Y2 Then .Y1 = Y2: .M1 = MM2: .C1 = CC2
                    If .Y2 > Y2 Then .Y2 = Y2: .M2 = MM2: .C2 = CC2
                Else
                    .Y1 = Y2: .M1 = MM2: .C1 = CC2
                    .Y2 = Y2: .M2 = MM2: .C2 = CC2
                    .Used = True
                End If
            End With
            Y2 = Y2 + StepY
            MM2.U = MM2.U + StepU
            MM2.V = MM2.V + StepV
            CC2.R = CC2.R + StepR
            CC2.G = CC2.G + StepG
            CC2.B = CC2.B + StepB
        Next
   End If
End Sub

Private Sub FillFlat()

    Dim DeltaY  As Long
    
    On Error Resume Next
    
    With Gratexels(minX)
        DeltaY = .Y1 - .Y2
        If .Y2 < g_SafeFrame.Top Then
            minY = g_SafeFrame.Top
        Else
            minY = .Y2
        End If
        maxY = IIf(.Y1 > g_SafeFrame.Bottom, g_SafeFrame.Bottom, .Y1)
        Select Case g_Materials(g_idxMt).OpcValue1
            Case 0:     FillFlat0
            Case 1:     FillFlat1
            Case Else:  FillFlatBlend
        End Select
    End With

End Sub

Private Sub FillGra()

    Dim DeltaY  As Long
    Dim Val     As Single
    
    On Error Resume Next
    
    With Gratexels(minX)
        DeltaY = .Y1 - .Y2
        StepR = Div(.C1.R - .C2.R, DeltaY)
        StepG = Div(.C1.G - .C2.G, DeltaY)
        StepB = Div(.C1.B - .C2.B, DeltaY)
        If .Y2 < g_SafeFrame.Top Then
            minY = g_SafeFrame.Top
            Val = Abs(g_SafeFrame.Top - .Y2)
            .C2.R = .C2.R + (StepR * Val)
            .C2.G = .C2.G + (StepG * Val)
            .C2.B = .C2.B + (StepB * Val)
        Else
            minY = .Y2
        End If
        maxY = IIf(.Y1 > g_SafeFrame.Bottom, g_SafeFrame.Bottom, .Y1)
        Select Case g_Materials(g_idxMt).OpcValue1
            Case 1:     FillGra1
            Case Else:  FillGraBlend
        End Select
    End With

End Sub

Private Sub FillFlaTex()

    Dim DeltaY  As Long
    Dim Val     As Single
    
    On Error Resume Next
    
    With Gratexels(minX)
        DeltaY = .Y1 - .Y2
        StepU = Div(.M1.U - .M2.U, DeltaY)
        StepV = Div(.M1.V - .M2.V, DeltaY)
        If .Y2 < g_SafeFrame.Top Then
            minY = g_SafeFrame.Top
            Val = Abs(g_SafeFrame.Top - .Y2)
            .M2.U = .M2.U + (StepU * Val)
            .M2.V = .M2.V + (StepV * Val)
        Else
            minY = .Y2
        End If
        maxY = IIf(.Y1 > g_SafeFrame.Bottom, g_SafeFrame.Bottom, .Y1)
        
        If g_Materials(g_idxMt).Texture.Filtre = None Then
            If g_Materials(g_idxMt).OpcValue1 = 1 Then
                If g_Lights(0).Enabled Then
                    FillFlatMapLightOn
                Else
                    FillFlatMap
                End If
           Else
                If g_Lights(0).Enabled Then
                    FillFlatMapLightOnBlend
                Else
                    FillFlatMapBlend
                End If
            End If
        Else
            If g_Materials(g_idxMt).OpcValue1 = 1 Then
                If g_Lights(0).Enabled Then
                    FillFlatMapBLinLightOn
                Else
                    FillFlatMapBLin
                End If
            Else
                If g_Lights(0).Enabled Then
                    FillFlatMapBLinLightOnBlend
                Else
                    FillFlatMapBLinBlend
                End If
            End If
        End If
    End With

End Sub

Private Sub FillGraTex()
    
    Dim DeltaY  As Long
    Dim Val     As Single
    
    On Error Resume Next
    
    With Gratexels(minX)
        DeltaY = .Y1 - .Y2
        StepU = Div(.M1.U - .M2.U, DeltaY)
        StepV = Div(.M1.V - .M2.V, DeltaY)
        StepR = Div(.C1.R - .C2.R, DeltaY)
        StepG = Div(.C1.G - .C2.G, DeltaY)
        StepB = Div(.C1.B - .C2.B, DeltaY)
        
        If .Y2 < g_SafeFrame.Top Then
            minY = g_SafeFrame.Top
            Val = Abs(g_SafeFrame.Top - .Y2)
            .M2.U = .M2.U + (StepU * Val)
            .M2.V = .M2.V + (StepV * Val)
            .C2.R = .C2.R + (StepR * Val)
            .C2.G = .C2.G + (StepG * Val)
            .C2.B = .C2.B + (StepB * Val)
        Else
            minY = .Y2
        End If
        maxY = IIf(.Y1 > g_SafeFrame.Bottom, g_SafeFrame.Bottom, .Y1)
        
        If g_Materials(g_idxMt).Texture.Filtre = None Then
            If g_Materials(g_idxMt).OpcValue1 = 1 Then
                If g_Lights(0).Enabled Then
                    FillGradMapLightOn
                Else
                    FillGradMap
                End If
           Else
                If g_Lights(0).Enabled Then
                    FillGradMapLightOnBlend
                Else
                    FillGradMapBlend
                End If
            End If
        Else
            If g_Materials(g_idxMt).OpcValue1 = 1 Then
                If g_Lights(0).Enabled Then
                    FillGradMapBLinLightOn
                Else
                    FillGradMapBLin
                End If
            Else
                If g_Lights(0).Enabled Then
                    FillGradMapBLinLightOnBlend
                Else
                    FillGradMapBLinBlend
                End If
            End If
        End If
        
    End With

End Sub

Private Sub FillFlat0()

    For minY = minY To maxY
        g_dibCanvas.uBI.Bits(minX, minY) = g_dibCanvas.uBI.Bits(minX, minY)
    Next

End Sub

Private Sub FillFlat1()
    
    For minY = minY To maxY
        g_dibCanvas.uBI.Bits(minX, minY) = g_FColor
    Next

End Sub

Private Sub FillGra1()

    For minY = minY To maxY
        g_dibCanvas.uBI.Bits(minX, minY) = ColorSToColorB(Gratexels(minX).C2)
        NextStepMC
    Next

End Sub

Private Sub FillFlatBlend()
        
    For minY = minY To maxY - 1
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendB(g_FColor, _
                                                       g_dibCanvas.uBI.Bits(minX, minY), _
                                                       g_Materials(g_idxMt).OpcValue1, _
                                                       g_Materials(g_idxMt).OpcValue2)
    Next
    
End Sub

Private Sub FillGraBlend()
    
    For minY = minY To maxY - 1
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendB(ColorSToColorB(Gratexels(minX).C2), _
                                                       g_dibCanvas.uBI.Bits(minX, minY), _
                                                       g_Materials(g_idxMt).OpcValue1, _
                                                       g_Materials(g_idxMt).OpcValue2)
        NextStepMC
    Next
    
End Sub

Private Sub FillFlatMap()

    For minY = minY To maxY
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        g_dibCanvas.uBI.Bits(minX, minY) = g_Materials(g_idxMt).Texture.dibTexT.uBI.Bits(Gratexels(minX).M2.U, Gratexels(minX).M2.V)
        NextStepMC
    Next

End Sub

Private Sub FillGradMap()

    For minY = minY To maxY
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        g_dibCanvas.uBI.Bits(minX, minY) = g_Materials(g_idxMt).Texture.dibTexT.uBI.Bits(Gratexels(minX).M2.U, Gratexels(minX).M2.V)
        NextStepMC
    Next

End Sub

Private Sub FillFlatMapLightOn()
        
    For minY = minY To maxY
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendB(g_Materials(g_idxMt).Texture.dibTexT.uBI.Bits(Gratexels(minX).M2.U, Gratexels(minX).M2.V), _
                                        g_FColor, _
                                        g_Lights(0).IntValue1, _
                                        g_Lights(0).IntValue2)
        NextStepMC
    Next

End Sub

Private Sub FillGradMapLightOn()
        
    For minY = minY To maxY
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendS(g_Materials(g_idxMt).Texture.dibTexT.uBI.Bits(Gratexels(minX).M2.U, Gratexels(minX).M2.V), _
                                                       Gratexels(minX).C2, _
                                                       g_Lights(0).IntValue1, _
                                                       g_Lights(0).IntValue2)
        NextStepMC
    Next

End Sub

Private Sub FillFlatMapBLin()
        
    For minY = minY To maxY
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        g_dibCanvas.uBI.Bits(minX, minY) = BLin
        NextStepMC
    Next

End Sub

Private Sub FillGradMapBLin()
        
    For minY = minY To maxY
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        g_dibCanvas.uBI.Bits(minX, minY) = BLin
        NextStepMC
    Next

End Sub

Private Sub FillFlatMapBLinLightOn()
    
    For minY = minY To maxY
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendB(BLin, _
                                                       g_FColor, _
                                                       g_Lights(0).IntValue1, _
                                                       g_Lights(0).IntValue2)
        NextStepMC
    Next

End Sub

Private Sub FillGradMapBLinLightOn()
    
    For minY = minY To maxY
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendS(BLin, _
                                                       Gratexels(minX).C2, _
                                                       g_Lights(0).IntValue1, _
                                                       g_Lights(0).IntValue2)
        NextStepMC
    Next

End Sub

Private Sub FillFlatMapBlend()
        
    For minY = minY To maxY - 1
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendB(g_Materials(g_idxMt).Texture.dibTexT.uBI.Bits(Gratexels(minX).M2.U, Gratexels(minX).M2.V), _
                                                       g_dibCanvas.uBI.Bits(minX, minY), _
                                                       g_Materials(g_idxMt).OpcValue1, _
                                                       g_Materials(g_idxMt).OpcValue2)
        NextStepMC
    Next
    
End Sub

Private Sub FillGradMapBlend()
        
    For minY = minY To maxY - 1
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendB(g_Materials(g_idxMt).Texture.dibTexT.uBI.Bits(Gratexels(minX).M2.U, Gratexels(minX).M2.V), _
                                                       g_dibCanvas.uBI.Bits(minX, minY), _
                                                       g_Materials(g_idxMt).OpcValue1, _
                                                       g_Materials(g_idxMt).OpcValue2)
        NextStepMC
    Next
    
End Sub

Private Sub FillFlatMapLightOnBlend()
    
    Dim Temp As RGBQUAD
    
    For minY = minY To maxY - 1
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        Temp = ColorBlendB(g_Materials(g_idxMt).Texture.dibTexT.uBI.Bits(Gratexels(minX).M2.U, Gratexels(minX).M2.V), _
                           g_dibCanvas.uBI.Bits(minX, minY), _
                           g_Materials(g_idxMt).OpcValue1, _
                           g_Materials(g_idxMt).OpcValue2)
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendB(Temp, _
                                                       g_FColor, _
                                                       g_Lights(0).IntValue1, _
                                                       g_Lights(0).IntValue2)
        NextStepMC
    Next

End Sub

Private Sub FillGradMapLightOnBlend()
    
    Dim Temp As RGBQUAD
    
    For minY = minY To maxY - 1
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        Temp = ColorBlendB(g_Materials(g_idxMt).Texture.dibTexT.uBI.Bits(Gratexels(minX).M2.U, Gratexels(minX).M2.V), _
                           g_dibCanvas.uBI.Bits(minX, minY), _
                           g_Materials(g_idxMt).OpcValue1, _
                           g_Materials(g_idxMt).OpcValue2)
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendS(Temp, _
                                                       Gratexels(minX).C2, _
                                                       g_Lights(0).IntValue1, _
                                                       g_Lights(0).IntValue2)
        NextStepMC
    Next

End Sub

Private Sub FillFlatMapBLinBlend()
        
    For minY = minY To maxY - 1
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendB(BLin, _
                                                       g_dibCanvas.uBI.Bits(minX, minY), _
                                                       g_Materials(g_idxMt).OpcValue1, _
                                                       g_Materials(g_idxMt).OpcValue2)
        NextStepMC
    Next

End Sub

Private Sub FillGradMapBLinBlend()
        
    For minY = minY To maxY - 1
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendB(BLin, _
                                                       g_dibCanvas.uBI.Bits(minX, minY), _
                                                       g_Materials(g_idxMt).OpcValue1, _
                                                       g_Materials(g_idxMt).OpcValue2)
        NextStepMC
    Next

End Sub

Private Sub FillFlatMapBLinLightOnBlend()
    
    Dim Temp As RGBQUAD
        
    For minY = minY To maxY - 1
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        Temp = ColorBlendB(BLin, _
                           g_dibCanvas.uBI.Bits(minX, minY), _
                           g_Materials(g_idxMt).OpcValue1, _
                           g_Materials(g_idxMt).OpcValue2)
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendB(Temp, _
                                                       g_FColor, _
                                                       g_Lights(0).IntValue1, _
                                                       g_Lights(0).IntValue2)
        NextStepMC
    Next
    
End Sub

Private Sub FillGradMapBLinLightOnBlend()
    
    Dim Temp As RGBQUAD
        
    For minY = minY To maxY - 1
        FindTexel Gratexels(minX).M2.U, Gratexels(minX).M2.V
        Temp = ColorBlendB(BLin, _
                           g_dibCanvas.uBI.Bits(minX, minY), _
                           g_Materials(g_idxMt).OpcValue1, _
                           g_Materials(g_idxMt).OpcValue2)
        g_dibCanvas.uBI.Bits(minX, minY) = ColorBlendS(Temp, _
                                                       Gratexels(minX).C2, _
                                                       g_Lights(0).IntValue1, _
                                                       g_Lights(0).IntValue2)
        NextStepMC
    Next
    
End Sub

Private Function ColorSToColorB(C1 As COLORRGB_SNG) As RGBQUAD

    Dim C2 As COLORRGB_SNG
    
    C2 = ColorLimit_SNG(C1)
    ColorSToColorB.rgbRed = CByte(C2.R)
    ColorSToColorB.rgbGreen = CByte(C2.G)
    ColorSToColorB.rgbBlue = CByte(C2.B)

End Function

Private Function ColorBlendB(C1 As RGBQUAD, C2 As RGBQUAD, A1 As Single, A2 As Single) As RGBQUAD

    ColorBlendB.rgbRed = CByte(C1.rgbRed * A1 + C2.rgbRed * A2)
    ColorBlendB.rgbGreen = CByte(C1.rgbGreen * A1 + C2.rgbGreen * A2)
    ColorBlendB.rgbBlue = CByte(C1.rgbBlue * A1 + C2.rgbBlue * A2)

End Function

Private Function ColorBlendS(C1 As RGBQUAD, C2 As COLORRGB_SNG, A1 As Single, A2 As Single) As RGBQUAD

    Dim C3 As COLORRGB_SNG

    C3 = ColorLimit_SNG(C2)
    ColorBlendS.rgbRed = CByte(C1.rgbRed * A1 + C3.R * A2)
    ColorBlendS.rgbGreen = CByte(C1.rgbGreen * A1 + C3.G * A2)
    ColorBlendS.rgbBlue = CByte(C1.rgbBlue * A1 + C3.B * A2)

End Function

Private Sub NextStepMC()
    
    With Gratexels(minX)
        .M2.U = .M2.U + StepU
        .M2.V = .M2.V + StepV
        .C2.R = .C2.R + StepR
        .C2.G = .C2.G + StepG
        .C2.B = .C2.B + StepB
    End With
                
End Sub

Private Function BLin() As RGBQUAD
    
    Dim U1      As Long
    Dim V1      As Long
    Dim U2      As Long
    Dim V2      As Long
    Dim fU      As Single 'Fraction U
    Dim fV      As Single 'Fraction V
    Dim osfU    As Single 'OneSubtractFraction U
    Dim osfV    As Single 'OneSubtractFraction V
    Dim F1      As Single
    Dim F2      As Single
    Dim F3      As Single
    Dim F4      As Single

    U1 = Fix(Gratexels(minX).M2.U)
    V1 = Fix(Gratexels(minX).M2.V)
    U2 = U1 + 1
    V2 = V1 + 1
    If U2 > g_Materials(g_idxMt).Texture.dibTexT.Width - 1 Then U2 = g_Materials(g_idxMt).Texture.dibTexT.Width - 1
    If V2 > g_Materials(g_idxMt).Texture.dibTexT.Height - 1 Then V2 = g_Materials(g_idxMt).Texture.dibTexT.Height - 1
    fU = Gratexels(minX).M2.U - U1
    fV = Gratexels(minX).M2.V - V1
    osfU = 1 - fU
    osfV = 1 - fV
    F1 = fU * fV
    F2 = fU * osfV
    F3 = osfU * fV
    F4 = osfU * osfV
            
    With g_Materials(g_idxMt).Texture.dibTexT.uBI
        BLin.rgbRed = F1 * .Bits(U2, V2).rgbRed + _
                      F2 * .Bits(U2, V1).rgbRed + _
                      F3 * .Bits(U1, V2).rgbRed + _
                      F4 * .Bits(U1, V1).rgbRed
        
        BLin.rgbGreen = F1 * .Bits(U2, V2).rgbGreen + _
                        F2 * .Bits(U2, V1).rgbGreen + _
                        F3 * .Bits(U1, V2).rgbGreen + _
                        F4 * .Bits(U1, V1).rgbGreen
        
        BLin.rgbBlue = F1 * .Bits(U2, V2).rgbBlue + _
                       F2 * .Bits(U2, V1).rgbBlue + _
                       F3 * .Bits(U1, V2).rgbBlue + _
                       F4 * .Bits(U1, V1).rgbBlue
    End With

End Function

Public Function Div(ByVal R1 As Single, ByVal R2 As Single) As Single
    
    If R2 = 0 Then R2 = ApproachVal
    Div = CSng(R1 / R2)

End Function

Public Sub FindTexel(U As Single, V As Single)
    
    Dim texWidth As Long
    Dim texHeight As Long
    
    texWidth = g_Materials(g_idxMt).Texture.dibTexT.Width - 1
    texHeight = g_Materials(g_idxMt).Texture.dibTexT.Height - 1
    
    If U > texWidth Then
        Do
            U = U - texWidth
        Loop Until U < texWidth
    End If
    
    If U < 0 Then
        Do
            U = U + texWidth
        Loop Until U > 0
    End If
    
    If V > texHeight Then
        Do
            V = V - (texHeight)
        Loop Until V < texHeight
    End If
    
    If V < 0 Then
        Do
            V = V + (texHeight)
        Loop Until V > 0
    End If
    
End Sub

Public Sub CalculateUV()

    Dim texWidth    As Long
    Dim texHeight   As Long
    Dim sUScale    As Single
    Dim sVScale    As Single
    Dim sUOffset    As Single
    Dim sVOffset    As Single
    Dim Angle       As Single
    Dim sngCosine   As Single
    Dim sngSine     As Single

    With g_Meshs(g_idxM)
        If .MapCoorsOK And g_Materials(g_idxMt).MapUse Then
            texWidth = g_Materials(g_idxMt).Texture.dibTexT.Width - 1
            texHeight = g_Materials(g_idxMt).Texture.dibTexT.Height - 1
            sUScale = g_Materials(g_idxMt).Texture.UScale * texWidth
            sVScale = g_Materials(g_idxMt).Texture.VScale * texHeight
            sUOffset = -g_Materials(g_idxMt).Texture.UOffset * texWidth
            sVOffset = -g_Materials(g_idxMt).Texture.VOffset * texHeight
            'A
            .TScreen(.Faces(g_idxF).A).U = sUOffset + (.MCoors(.Faces(g_idxF).A).U * sUScale)
            .TScreen(.Faces(g_idxF).A).V = sVOffset + (.MCoors(.Faces(g_idxF).A).V * sVScale)
            If g_Materials(g_idxMt).Texture.Angle <> 0 Then
                Angle = ConvertDeg2Rad(g_Materials(g_idxMt).Texture.Angle)
                sngCosine = Round(Cos(Angle), 6)
                sngSine = Round(Sin(Angle), 6)
                .TScreen(.Faces(g_idxF).A).U = (sngCosine * .TScreen(.Faces(g_idxF).A).U) - _
                                                (sngSine * .TScreen(.Faces(g_idxF).A).V)
                .TScreen(.Faces(g_idxF).A).V = (sngSine * .TScreen(.Faces(g_idxF).A).U) + _
                                                (sngCosine * .TScreen(.Faces(g_idxF).A).V)
            End If
            'B
            .TScreen(.Faces(g_idxF).B).U = sUOffset + (.MCoors(.Faces(g_idxF).B).U * sUScale)
            .TScreen(.Faces(g_idxF).B).V = sVOffset + (.MCoors(.Faces(g_idxF).B).V * sVScale)
            If g_Materials(g_idxMt).Texture.Angle <> 0 Then
                Angle = ConvertDeg2Rad(g_Materials(g_idxMt).Texture.Angle)
                sngCosine = Round(Cos(Angle), 6)
                sngSine = Round(Sin(Angle), 6)
                .TScreen(.Faces(g_idxF).B).U = (sngCosine * .TScreen(.Faces(g_idxF).B).U) - _
                                                (sngSine * .TScreen(.Faces(g_idxF).B).V)
                .TScreen(.Faces(g_idxF).B).V = (sngSine * .TScreen(.Faces(g_idxF).B).U) + _
                                                (sngCosine * .TScreen(.Faces(g_idxF).B).V)
            End If
            'C
            .TScreen(.Faces(g_idxF).C).U = sUOffset + (.MCoors(.Faces(g_idxF).C).U * sUScale)
            .TScreen(.Faces(g_idxF).C).V = sVOffset + (.MCoors(.Faces(g_idxF).C).V * sVScale)
            If g_Materials(g_idxMt).Texture.Angle <> 0 Then
                Angle = ConvertDeg2Rad(g_Materials(g_idxMt).Texture.Angle)
                sngCosine = Round(Cos(Angle), 6)
                sngSine = Round(Sin(Angle), 6)
                .TScreen(.Faces(g_idxF).C).U = (sngCosine * .TScreen(.Faces(g_idxF).C).U) - _
                                                (sngSine * .TScreen(.Faces(g_idxF).C).V)
                .TScreen(.Faces(g_idxF).C).V = (sngSine * .TScreen(.Faces(g_idxF).C).U) + _
                                                (sngCosine * .TScreen(.Faces(g_idxF).C).V)
            End If
        End If
    End With

End Sub

Private Sub GetFaceColor()
        
    For g_idxM = 0 To g_idxMesh
        With g_Meshs(g_idxM)
            For g_idxF = 0 To .idxFace
                .Vertices(.Faces(g_idxF).A).Used = False
                .Vertices(.Faces(g_idxF).B).Used = False
                .Vertices(.Faces(g_idxF).C).Used = False
            Next g_idxF
            If g_Lights(0).Enabled Then
                For g_idxF = 0 To .idxFace
                    GetVertexColorS
                Next g_idxF
            Else
                For g_idxF = 0 To .idxFace
                    GetVertexColor
                Next g_idxF
            End If
        End With
    Next g_idxM
    
End Sub

Private Sub GetFaceColorO()
    
    Dim idxOrder    As Long
    
    For idxOrder = 0 To UBound(g_MeshsOrder)
        g_idxM = g_MeshsOrder(idxOrder).idxMeshO
        g_idxF = g_MeshsOrder(idxOrder).idxFaceO
        With g_Meshs(g_idxM)
            .Vertices(.Faces(g_idxF).A).Used = False
            .Vertices(.Faces(g_idxF).B).Used = False
            .Vertices(.Faces(g_idxF).C).Used = False
        End With
    Next idxOrder
    
    If g_Lights(0).Enabled Then
        For idxOrder = 0 To UBound(g_MeshsOrder)
            g_idxM = g_MeshsOrder(idxOrder).idxMeshO
            g_idxF = g_MeshsOrder(idxOrder).idxFaceO
            GetVertexColorS
        Next idxOrder
    Else
        For idxOrder = 0 To UBound(g_MeshsOrder)
            g_idxM = g_MeshsOrder(idxOrder).idxMeshO
            g_idxF = g_MeshsOrder(idxOrder).idxFaceO
            GetVertexColor
        Next idxOrder
    End If
        
End Sub

Private Sub GetVertexColorS()
    
    Dim vNormal As VECTOR
    
    With g_Meshs(g_idxM)
        vNormal = FaceNormal(.Vertices(.Faces(g_idxF).A).VectorsS, _
                             .Vertices(.Faces(g_idxF).B).VectorsS, _
                             .Vertices(.Faces(g_idxF).C).VectorsS)
        If .Vertices(.Faces(g_idxF).A).Used = False Then
            .Vertices(.Faces(g_idxF).A).intRGBColor = Shade(vNormal, .Faces(g_idxF).A)
            .Vertices(.Faces(g_idxF).A).Used = True
        End If
        If .Vertices(.Faces(g_idxF).B).Used = False Then
            .Vertices(.Faces(g_idxF).B).intRGBColor = Shade(vNormal, .Faces(g_idxF).B)
            .Vertices(.Faces(g_idxF).B).Used = True
        End If
        If .Vertices(.Faces(g_idxF).C).Used = False Then
            .Vertices(.Faces(g_idxF).C).intRGBColor = Shade(vNormal, .Faces(g_idxF).C)
            .Vertices(.Faces(g_idxF).C).Used = True
        End If
    End With

End Sub

Private Sub GetVertexColor()
        
    With g_Meshs(g_idxM)
        If .Vertices(.Faces(g_idxF).A).Used = False Then
            .Vertices(.Faces(g_idxF).A).intRGBColor = g_Materials(.Faces(g_idxF).idxMat).Diffuse
            .Vertices(.Faces(g_idxF).A).Used = True
        End If
        If .Vertices(.Faces(g_idxF).B).Used = False Then
            .Vertices(.Faces(g_idxF).B).intRGBColor = g_Materials(.Faces(g_idxF).idxMat).Diffuse
            .Vertices(.Faces(g_idxF).B).Used = True
        End If
        If .Vertices(.Faces(g_idxF).C).Used = False Then
            .Vertices(.Faces(g_idxF).C).intRGBColor = g_Materials(.Faces(g_idxF).idxMat).Diffuse
            .Vertices(.Faces(g_idxF).C).Used = True
        End If
    End With

End Sub

Private Function PointInCanvas(Point As POINTAPI) As Boolean
    
    If Point.X > g_SafeFrame.Left And Point.X < g_SafeFrame.Right And _
       Point.Y > g_SafeFrame.Top And Point.Y < g_SafeFrame.Bottom Then PointInCanvas = True

End Function

Private Function ZValInCamera(ZVal As Single) As Boolean
    
    If (ZVal > g_Cameras(0).ClipNear) And _
       (ZVal < g_Cameras(0).ClipFar) Then ZValInCamera = True
                   
End Function

Private Function IsInCamera() As Boolean
    
    With g_Meshs(g_idxM)
        If ZValInCamera(.Vertices(.Faces(g_idxF).A).VectorsT.Z) And _
           ZValInCamera(.Vertices(.Faces(g_idxF).B).VectorsT.Z) And _
           ZValInCamera(.Vertices(.Faces(g_idxF).C).VectorsT.Z) Then IsInCamera = True
    End With

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Background~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub GradientBack()

    Dim gTri As GRADIENT_TRIANGLE
    gTri.Vertex1 = 0
    gTri.Vertex2 = 1
    gTri.Vertex3 = 2
    Call GradientFillTriangle(g_dibBack.hDC, triVert(0), 3, gTri, 1, GRADIENT_FILL_TRIANGLE)

End Sub

Public Sub Gradient0()
    
    triVert(0).X = 0
    triVert(0).Y = 0
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    triVert(1).X = g_CanvasWidth
    triVert(1).Y = 0
    triVert(1).Red = ConvertUShort(BColor1.R)
    triVert(1).Green = ConvertUShort(BColor1.G)
    triVert(1).Blue = ConvertUShort(BColor1.B)

    triVert(2).X = 0
    triVert(2).Y = g_CanvasHeight
    triVert(2).Red = ConvertUShort(BColor1.R)
    triVert(2).Green = ConvertUShort(BColor1.G)
    triVert(2).Blue = ConvertUShort(BColor1.B)
    
    Call GradientBack
    
    triVert(0).X = g_CanvasWidth
    triVert(0).Y = g_CanvasHeight
    
    Call GradientBack

End Sub

Public Sub Gradient1()
    
    triVert(0).X = 0
    triVert(0).Y = 0
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    triVert(1).X = g_CanvasWidth
    triVert(1).Y = 0
    triVert(1).Red = ConvertUShort(BColor1.R)
    triVert(1).Green = ConvertUShort(BColor1.G)
    triVert(1).Blue = ConvertUShort(BColor1.B)

    triVert(2).X = 0
    triVert(2).Y = g_CanvasHeight
    triVert(2).Red = ConvertUShort(BColor2.R)
    triVert(2).Green = ConvertUShort(BColor2.G)
    triVert(2).Blue = ConvertUShort(BColor2.B)
    
    Call GradientBack
    
    triVert(0).X = g_CanvasWidth
    triVert(0).Y = g_CanvasHeight
    triVert(0).Red = ConvertUShort(BColor2.R)
    triVert(0).Green = ConvertUShort(BColor2.G)
    triVert(0).Blue = ConvertUShort(BColor2.B)
    
    Call GradientBack

End Sub

Public Sub Gradient2()
    
    triVert(0).X = 0
    triVert(0).Y = 0
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    triVert(1).X = g_CanvasWidth
    triVert(1).Y = 0
    triVert(1).Red = ConvertUShort(BColor2.R)
    triVert(1).Green = ConvertUShort(BColor2.G)
    triVert(1).Blue = ConvertUShort(BColor2.B)

    triVert(2).X = 0
    triVert(2).Y = g_CanvasHeight
    triVert(2).Red = ConvertUShort(BColor2.R)
    triVert(2).Green = ConvertUShort(BColor2.G)
    triVert(2).Blue = ConvertUShort(BColor2.B)
    
    Call GradientBack
    
    triVert(0).X = g_CanvasWidth
    triVert(0).Y = g_CanvasHeight
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    Call GradientBack

End Sub

Public Sub Gradient3()
    
    triVert(0).X = 0
    triVert(0).Y = 0
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    triVert(1).X = g_CanvasWidth
    triVert(1).Y = 0
    triVert(1).Red = ConvertUShort(BColor1.R)
    triVert(1).Green = ConvertUShort(BColor1.G)
    triVert(1).Blue = ConvertUShort(BColor1.B)

    triVert(2).X = g_OriginX
    triVert(2).Y = g_OriginY
    triVert(2).Red = ConvertUShort(BColor2.R)
    triVert(2).Green = ConvertUShort(BColor2.G)
    triVert(2).Blue = ConvertUShort(BColor2.B)
    
    Call GradientBack
    
    triVert(0).X = g_CanvasWidth
    triVert(0).Y = g_CanvasHeight
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    Call GradientBack

    triVert(1).X = 0
    triVert(1).Y = g_CanvasHeight
    triVert(1).Red = ConvertUShort(BColor1.R)
    triVert(1).Green = ConvertUShort(BColor1.G)
    triVert(1).Blue = ConvertUShort(BColor1.B)
    
    Call GradientBack
    
    triVert(0).X = 0
    triVert(0).Y = 0
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    Call GradientBack

End Sub

Public Sub GradientRectangle(Pic As PictureBox, C1 As COLORRGB_INT, C2 As COLORRGB_INT, Mode As GradMode)

    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT
    
    With vert(0)
        .X = 0
        .Y = 0
        .Red = ConvertUShort(C1.R)
        .Green = ConvertUShort(C1.G)
        .Blue = ConvertUShort(C1.B)
        .Alpha = 0&
    End With

    With vert(1)
        .X = Pic.ScaleWidth
        .Y = Pic.ScaleHeight
        .Red = ConvertUShort(C2.R)
        .Green = ConvertUShort(C2.G)
        .Blue = ConvertUShort(C2.B)
        .Alpha = 0&
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    GradientFillRect Pic.hDC, vert(0), 2, gRect, 1, Mode
    Pic.Refresh
 
End Sub
