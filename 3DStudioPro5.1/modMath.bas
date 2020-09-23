Attribute VB_Name = "modMath"
Option Explicit

Private Const dblPIDiv180   As Double = 1.74532925199433E-02
Private Const sng180DivPI   As Single = 57.2957795130823
Private Const sngPIDiv360   As Single = 0.008726!
Private Const sng360DivPI   As Single = 114.5916!

Public Function ConvertDeg2Rad(Degress As Single) As Single
    
    ConvertDeg2Rad = Degress * dblPIDiv180
    
End Function

Public Function ConvertRad2Deg(Radians As Single) As Single
    
    ConvertRad2Deg = Radians * sng180DivPI
    
End Function

Public Function ConvertZoomtoFOV(Zoom As Single) As Single
    
    ConvertZoomtoFOV = sng360DivPI * Atn(1 / Zoom)
    
End Function

Public Function ConvertFOVtoZoom(FOV As Single) As Single
    
    ConvertFOVtoZoom = 1 / Tan(FOV * sngPIDiv360)
    
End Function

Public Function VectorSet(X As Single, Y As Single, Z As Single) As VECTOR

    VectorSet.X = X
    VectorSet.Y = Y
    VectorSet.Z = Z
    VectorSet.W = 1
    
End Function

Public Function VectorPlus(V1 As VECTOR, Val As Single) As VECTOR

    VectorPlus.X = V1.X + Val
    VectorPlus.Y = V1.Y + Val
    VectorPlus.Z = V1.Z + Val
    VectorPlus.W = 1
    
End Function

Public Function VectorMinus(V1 As VECTOR, Val As Single) As VECTOR

    VectorMinus.X = V1.X - Val
    VectorMinus.Y = V1.Y - Val
    VectorMinus.Z = V1.Z - Val
    VectorMinus.W = 1
    
End Function

Public Function VectorAddition(V1 As VECTOR, V2 As VECTOR) As VECTOR
        
    VectorAddition.X = V1.X + V2.X
    VectorAddition.Y = V1.Y + V2.Y
    VectorAddition.Z = V1.Z + V2.Z
    VectorAddition.W = 1

End Function

Public Function VectorSubtract(V1 As VECTOR, V2 As VECTOR) As VECTOR
        
    VectorSubtract.X = V1.X - V2.X
    VectorSubtract.Y = V1.Y - V2.Y
    VectorSubtract.Z = V1.Z - V2.Z
    VectorSubtract.W = 1

End Function

Public Function VectorScale(V1 As VECTOR, Scalar As Single) As VECTOR

    VectorScale.X = V1.X * Scalar
    VectorScale.Y = V1.Y * Scalar
    VectorScale.Z = V1.Z * Scalar

End Function

Public Function VectorNormalize(V1 As VECTOR) As VECTOR

    Dim sngLength As Single
    
    sngLength = VectorLength(V1)
    If sngLength = 0 Then sngLength = 1
    VectorNormalize.X = V1.X / sngLength
    VectorNormalize.Y = V1.Y / sngLength
    VectorNormalize.Z = V1.Z / sngLength
    VectorNormalize.W = V1.W
    
End Function

Public Function VectorLength(V1 As VECTOR) As Single
    
    VectorLength = Sqr(V1.X * V1.X + V1.Y * V1.Y + V1.Z * V1.Z)
    
End Function

Public Function VectorAngle(V1 As VECTOR, V2 As VECTOR) As Single

    VectorAngle = DotProduct(VectorNormalize(V1), VectorNormalize(V2))

End Function

Public Function VectorReflect(V1 As VECTOR, V2 As VECTOR) As VECTOR

    If VectorAngle(V1, V2) < 0 Then
        VectorReflect = VectorAddition(V1, VectorScale(VectorScale(VectorNormalize(V2), DotProduct(V1, VectorNormalize(V2))), -2))
    End If

End Function

Public Function VectorDistance(V1 As VECTOR, V2 As VECTOR) As Single

    VectorDistance = VectorLength(VectorSubtract(V1, V2))

End Function

Public Function FaceNormal(V1 As VECTOR, V2 As VECTOR, V3 As VECTOR) As VECTOR

    FaceNormal = VectorNormalize(CrossProduct(VectorSubtract(V1, V2), VectorSubtract(V3, V2)))

End Function
'
'Public Function FaceCenter(V1 As VECTOR, V2 As VECTOR, V3 As VECTOR) As VECTOR
'
'    FaceCenter.X = (V1.X + V2.X + V3.X) * sng1Div3
'    FaceCenter.Y = (V1.Y + V2.Y + V3.Y) * sng1Div3
'    FaceCenter.Z = (V1.Z + V2.Z + V3.Z) * sng1Div3
'    FaceCenter.W = 1
'
'End Function

Public Function CrossProduct(V1 As VECTOR, V2 As VECTOR) As VECTOR
    
    CrossProduct.X = V1.Y * V2.Z - V1.Z * V2.Y
    CrossProduct.Y = V1.Z * V2.X - V1.X * V2.Z
    CrossProduct.Z = V1.X * V2.Y - V1.Y * V2.X
    CrossProduct.W = 1
    
End Function

Public Function DotProduct(V1 As VECTOR, V2 As VECTOR) As Single

    DotProduct = V1.X * V2.X + V1.Y * V2.Y + V1.Z * V2.Z
    
End Function

Public Function MatrixIdentity() As MATRIX
    
    With MatrixIdentity
        .rc11 = 1: .rc12 = 0: .rc13 = 0: .rc14 = 0
        .rc21 = 0: .rc22 = 1: .rc23 = 0: .rc24 = 0
        .rc31 = 0: .rc32 = 0: .rc33 = 1: .rc34 = 0
        .rc41 = 0: .rc42 = 0: .rc43 = 0: .rc44 = 1
    End With
    
End Function

Public Function MatrixRotationX(Radians As Single) As MATRIX
    
    Dim sngCosine As Double
    Dim sngSine As Double
    
    sngCosine = Cos(Radians)
    sngSine = Sin(Radians)
    MatrixRotationX = MatrixIdentity()
    MatrixRotationX.rc22 = sngCosine
    MatrixRotationX.rc23 = -sngSine
    MatrixRotationX.rc32 = sngSine
    MatrixRotationX.rc33 = sngCosine

End Function

Public Function MatrixRotationY(Radians As Single) As MATRIX
    
    Dim sngCosine As Double
    Dim sngSine As Double
    
    sngCosine = Cos(Radians)
    sngSine = Sin(Radians)
    MatrixRotationY = MatrixIdentity()
    MatrixRotationY.rc11 = sngCosine
    MatrixRotationY.rc31 = -sngSine
    MatrixRotationY.rc13 = sngSine
    MatrixRotationY.rc33 = sngCosine
    
End Function

Public Function MatrixRotationZ(Radians As Single) As MATRIX
    
    Dim sngCosine As Double
    Dim sngSine As Double
    
    sngCosine = Cos(Radians)
    sngSine = Sin(Radians)
    MatrixRotationZ = MatrixIdentity()
    MatrixRotationZ.rc11 = sngCosine
    MatrixRotationZ.rc21 = sngSine
    MatrixRotationZ.rc12 = -sngSine
    MatrixRotationZ.rc22 = sngCosine

End Function

Public Function MatrixScale(X As Single, Y As Single, Z As Single, W As Single) As MATRIX
    
    MatrixScale = MatrixIdentity()
    MatrixScale.rc11 = X
    MatrixScale.rc22 = Y
    MatrixScale.rc33 = Z
    MatrixScale.rc44 = W
    
End Function

Public Function MatrixTranslation(X As Single, Y As Single, Z As Single) As MATRIX
    
    MatrixTranslation = MatrixIdentity()
    MatrixTranslation.rc14 = X
    MatrixTranslation.rc24 = Y
    MatrixTranslation.rc34 = Z
    
End Function

Public Function MatrixMultiplyVector(M1 As MATRIX, V1 As VECTOR) As VECTOR
    
    With MatrixMultiplyVector
        .X = M1.rc11 * V1.X + M1.rc12 * V1.Y + M1.rc13 * V1.Z + M1.rc14 * V1.W
        .Y = M1.rc21 * V1.X + M1.rc22 * V1.Y + M1.rc23 * V1.Z + M1.rc24 * V1.W
        .Z = M1.rc31 * V1.X + M1.rc32 * V1.Y + M1.rc33 * V1.Z + M1.rc34 * V1.W
        .W = M1.rc41 * V1.X + M1.rc42 * V1.Y + M1.rc43 * V1.Z + M1.rc44 * V1.W
    End With
    
End Function

Public Function MatrixMultiply(M1 As MATRIX, M2 As MATRIX) As MATRIX

    With MatrixMultiply
        .rc11 = M1.rc11 * M2.rc11 + M1.rc21 * M2.rc12 + M1.rc31 * M2.rc13
        .rc12 = M1.rc12 * M2.rc11 + M1.rc22 * M2.rc12 + M1.rc32 * M2.rc13
        .rc13 = M1.rc13 * M2.rc11 + M1.rc23 * M2.rc12 + M1.rc33 * M2.rc13
        .rc14 = M1.rc14 * M2.rc11 + M1.rc24 * M2.rc12 + M1.rc34 * M2.rc13 + M2.rc14
        .rc21 = M1.rc11 * M2.rc21 + M1.rc21 * M2.rc22 + M1.rc31 * M2.rc23
        .rc22 = M1.rc12 * M2.rc21 + M1.rc22 * M2.rc22 + M1.rc32 * M2.rc23
        .rc23 = M1.rc13 * M2.rc21 + M1.rc23 * M2.rc22 + M1.rc33 * M2.rc23
        .rc24 = M1.rc14 * M2.rc21 + M1.rc24 * M2.rc22 + M1.rc34 * M2.rc23 + M2.rc24
        .rc31 = M1.rc11 * M2.rc31 + M1.rc21 * M2.rc32 + M1.rc31 * M2.rc33
        .rc32 = M1.rc12 * M2.rc31 + M1.rc22 * M2.rc32 + M1.rc32 * M2.rc33
        .rc33 = M1.rc13 * M2.rc31 + M1.rc23 * M2.rc32 + M1.rc33 * M2.rc33
        .rc34 = M1.rc14 * M2.rc31 + M1.rc24 * M2.rc32 + M1.rc34 * M2.rc33 + M2.rc34
    End With

End Function

Public Function WorldMatrix() As MATRIX

    Dim sX As Single, sY As Single, sZ As Single
    Dim tX As Single, tY As Single, tZ As Single
    Dim SinZ As Double, CosZ As Double
    Dim SinX As Double, CosX As Double
    Dim SinY As Double, CosY As Double

    CosX = Cos(ConvertDeg2Rad(g_MeshsRotation.X))
    SinX = Sin(ConvertDeg2Rad(g_MeshsRotation.X))
    CosY = Cos(ConvertDeg2Rad(g_MeshsRotation.Y))
    SinY = Sin(ConvertDeg2Rad(g_MeshsRotation.Y))
    CosZ = Cos(ConvertDeg2Rad(g_MeshsRotation.Z))
    SinZ = Sin(ConvertDeg2Rad(g_MeshsRotation.Z))
    sX = g_MeshsScales.X
    sY = g_MeshsScales.Y
    sZ = g_MeshsScales.Z
    tX = g_MeshsTranslation.X
    tY = g_MeshsTranslation.Y
    tZ = g_MeshsTranslation.Z

    With WorldMatrix
        .rc11 = sX * CosY * CosZ
        .rc12 = sY * (SinX * SinY * CosZ + CosX * -SinZ)
        .rc13 = sZ * (CosX * SinY * CosZ + -SinX * -SinZ)
        .rc14 = tX
        .rc21 = sX * CosY * SinZ
        .rc22 = sY * (SinX * SinY * SinZ + CosX * CosZ)
        .rc23 = sZ * (CosX * SinY * SinZ + -SinX * CosZ)
        .rc24 = tY
        .rc31 = sX * -SinY
        .rc32 = sY * SinX * CosY
        .rc33 = sZ * CosX * CosY
        .rc34 = tZ
    End With

End Function

Public Function ViewMatrix() As MATRIX

    Dim vVPN    As VECTOR
    Dim vVUP    As VECTOR
    Dim vVRP    As VECTOR
    Dim vN      As VECTOR
    Dim vU      As VECTOR
    Dim vV      As VECTOR

    With g_Cameras(0)
        vVRP = .WorldPosition
        vVPN = VectorSubtract(.WorldPosition, .LookAtPoint)
        If (vVPN.X = 0) And (vVPN.Y = 0) And (vVPN.Z = 0) Then vVPN = VectorSet(0, 0, 1)
        vVUP = .VUP
    End With

    vN = VectorNormalize(vVPN)
    vU = VectorNormalize(CrossProduct(vN, vVUP))
    vV = CrossProduct(vU, vN)

    With ViewMatrix
        .rc11 = vU.X
        .rc12 = vU.Y
        .rc13 = vU.Z
        .rc14 = -vVRP.X * vU.X + -vVRP.Y * vU.Y + -vVRP.Z * vU.Z
        .rc21 = vV.X
        .rc22 = vV.Y
        .rc23 = vV.Z
        .rc24 = -vVRP.X * vV.X + -vVRP.Y * vV.Y + -vVRP.Z * vV.Z
        .rc31 = vN.X
        .rc32 = vN.Y
        .rc33 = vN.Z
        .rc34 = -vVRP.X * vN.X + -vVRP.Y * vN.Y + -vVRP.Z * vN.Z
    End With

End Function

''Public Function WorldMatrix() As MATRIX
''
''    Dim vTranslate As VECTOR, vRotate As VECTOR, VScale As VECTOR
''    Dim matTranslation As MATRIX
''    Dim matRotateX As MATRIX
''    Dim matRotateY As MATRIX
''    Dim matRotateZ As MATRIX
''    Dim matUniformScale As MATRIX
''
''    vTranslate = MeshsTranslation
''    vRotate = MeshsRotation
''    VScale = MeshsScales
''
''    matUniformScale = MatrixScale(VScale.X, VScale.Y, VScale.Z, 1)
''    matRotateX = MatrixRotationX(ConvertDeg2Rad(vRotate.X))
''    matRotateY = MatrixRotationY(ConvertDeg2Rad(vRotate.Y))
''    matRotateZ = MatrixRotationZ(ConvertDeg2Rad(vRotate.Z))
''    matTranslation = MatrixTranslation(vTranslate.X, vTranslate.Y, vTranslate.Z)
''
''    WorldMatrix = MatrixIdentity()
''    WorldMatrix = MatrixMultiply(WorldMatrix, matUniformScale)
''    WorldMatrix = MatrixMultiply(WorldMatrix, matRotateX)
''    WorldMatrix = MatrixMultiply(WorldMatrix, matRotateY)
''    WorldMatrix = MatrixMultiply(WorldMatrix, matRotateZ)
''    WorldMatrix = MatrixMultiply(WorldMatrix, matTranslation)
''
''End Function

''Public Function ViewMatrix() As MATRIX
''
''    Dim vVPN    As VECTOR
''    Dim vVUP    As VECTOR
''    Dim vVRP    As VECTOR
''    Dim vN   As VECTOR
''    Dim vU   As VECTOR
''    Dim vV   As VECTOR
''    Dim matRotateVRC    As MATRIX
''    Dim matTranslateVRP As MATRIX
''
''    With Cameras(0)
''        vVRP = .WorldPosition
''        vVPN = VectorSubtract(.WorldPosition, .LookAtPoint)
''        If (vVPN.X = 0) And (vVPN.Y = 0) And (vVPN.Z = 0) Then vVPN = VectorSet(0, 0, 1)
''        vVUP = .VUP
''    End With
''
''    vN = VectorNormalize(vVPN)
''    vU = VectorNormalize(CrossProduct(vN, vVUP))
''    vV = CrossProduct(vU, vN)
''    matRotateVRC = MatrixIdentity()
''    With matRotateVRC
''        .rc11 = vU.X: .rc12 = vU.Y: .rc13 = vU.Z
''        .rc21 = vV.X: .rc22 = vV.Y: .rc23 = vV.Z
''        .rc31 = vN.X: .rc32 = vN.Y: .rc33 = vN.Z
''    End With
''    matTranslateVRP = MatrixTranslation(-vVRP.X, -vVRP.Y, -vVRP.Z)
''    ViewMatrix = MatrixIdentity()
''    ViewMatrix = MatrixMultiply(ViewMatrix, matTranslateVRP)
''    ViewMatrix = MatrixMultiply(ViewMatrix, matRotateVRC)
''
''End Function

Public Function MatrixShadow(LightPosition As VECTOR, PlaneEquation As VECTOR) As MATRIX
        
    Dim sngDP As Single
    
    sngDP = DotProduct2(LightPosition, PlaneEquation)
    
    With MatrixShadow
        .rc11 = sngDP - (LightPosition.X * PlaneEquation.X)
        .rc12 = -(LightPosition.X * PlaneEquation.Y)
        .rc13 = -(LightPosition.X * PlaneEquation.Z)
        .rc14 = -(LightPosition.X * PlaneEquation.W)
        
        .rc21 = -(LightPosition.Y * PlaneEquation.X)
        .rc22 = sngDP - (LightPosition.Y * PlaneEquation.Y)
        .rc23 = -(LightPosition.Y * PlaneEquation.Z)
        .rc24 = -(LightPosition.Y * PlaneEquation.W)
        
        .rc31 = -(LightPosition.Z * PlaneEquation.X)
        .rc32 = -(LightPosition.Z * PlaneEquation.Y)
        .rc33 = sngDP - (LightPosition.Z * PlaneEquation.Z)
        .rc34 = -(LightPosition.Z * PlaneEquation.W)
        
        .rc41 = -(LightPosition.W * PlaneEquation.X)
        .rc42 = -(LightPosition.W * PlaneEquation.Y)
        .rc43 = -(LightPosition.W * PlaneEquation.Z)
        .rc44 = sngDP - (LightPosition.W * PlaneEquation.W)
    End With

End Function

Public Function DotProduct2(V1 As VECTOR, V2 As VECTOR) As Single
    
    DotProduct2 = V1.X * V2.X + V1.Y * V2.Y + V1.Z * V2.Z + V1.W * V2.W

End Function

