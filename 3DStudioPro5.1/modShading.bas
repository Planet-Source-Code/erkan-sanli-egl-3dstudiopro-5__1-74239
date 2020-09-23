Attribute VB_Name = "modShading"
Option Explicit

Public Function Shade(Normal As VECTOR, idxVert As Long) As COLORRGB_INT

    Dim Alpha           As Single 'Not transparency value , this is angle
    Dim Beta            As Single
    Dim Epsilon         As Single
    Dim Gamma           As Single
    Dim VectorT         As VECTOR
    Dim Color           As COLORRGB_INT

    VectorT = g_Meshs(g_idxM).Vertices(idxVert).VectorsT
    Color = g_Materials(g_Meshs(g_idxM).Faces(g_idxF).idxMat).Diffuse
    With g_Lights(0)
        If .Enabled = True Then
            .OriginT = MatrixMultiplyVector(g_Cameras(0).ViewMatrix, .Origin)
            .DirectionT = MatrixMultiplyVector(g_Cameras(0).ViewMatrix, .Direction)

            If .Falloff > 0 Then
                Epsilon = VectorAngle(.DirectionT, VectorSubtract(VectorT, .OriginT))
                If .Falloff <> .Hotspot Then
                    Epsilon = (.Falloff - Epsilon) / (.Falloff - .Hotspot)
                    If Epsilon < 0 Then Epsilon = 0
                    If Epsilon > 1 Then Epsilon = 1
                Else
                    Exit Function
                End If
            Else
                Epsilon = 1
            End If

            Alpha = VectorAngle(VectorSubtract(.OriginT, VectorT), Normal) * .Diffusion
            If Alpha < 0 Then Alpha = 0
            Beta = VectorAngle(VectorReflect(VectorSubtract(VectorT, .OriginT), Normal), VectorSubtract(g_Cameras(0).WorldPosition, VectorT)) * .Specular
            If Beta < 0 Then Beta = 0

            If .AttenEnable Then
                If .DarkRange <> .BrightRange Then
                    Gamma = (.DarkRange - VectorDistance(VectorT, .OriginT)) / (.DarkRange - .BrightRange)
                    If Gamma < 0 Then Gamma = 0
                    If Gamma > 1 Then Gamma = 1
                End If
            Else
                Gamma = 1
            End If
            Shade = ColorScale(.Color, (Alpha + Beta) * Gamma * Epsilon)
            Shade = ColorInterpolateInt(Shade, Color, 0.5)
            Shade = ColorAdd(Shade, .Ambient)
        End If
    End With
    Call ColorLimit(Shade)

End Function

'Public Function Shade(Normal As VECTOR, idxVert As Long) As COLORRGB_INT
'
'    Dim eyeX As Single
'    Dim eyeY As Single
'    Dim eyeZ As Single
'    Dim PX As Single
'    Dim PY As Single
'    Dim Pz As Single
'    Dim Nx As Single
'    Dim Ny As Single
'    Dim Nz As Single
'
'    Dim DiffKr As Single
'    Dim DiffKg As Single
'    Dim DiffKb As Single
'
'    Dim AmbKr As Single
'    Dim AmbKg As Single
'    Dim AmbKb As Single
'
'    Dim SpecK As Single
'    Dim SpecN As Single
'
'    ' Vectors:
'    Dim Vx As Single        'V: p to viewpoint
'    Dim Vy As Single
'    Dim Vz As Single
'    Dim Vlen As Single
'    Dim Lx As Single        'L: p to lightsource
'    Dim Ly As Single
'    Dim Lz As Single
'    Dim Llen As Single
'    Dim LMx As Single       'LM: Light source mirror vector
'    Dim LMy As Single
'    Dim LMz As Single
'
'    ' Dot products:
'    Dim LdotN As Single
'    Dim VdotN As Single
'    Dim LMdotV As Single
'
'    ' Colors:
'    Dim TotalR As Double 'Single
'    Dim TotalG As Double 'Single
'    Dim TotalB As Double 'Single
'
'    Dim spec As Double
'
'    eyeX = g_Cameras(0).WorldPosition.X
'    eyeY = g_Cameras(0).WorldPosition.Y
'    eyeZ = g_Cameras(0).WorldPosition.Z
'
'    PX = g_Meshs(g_idxM).Vertices(idxVert).VectorsT.X
'    PY = g_Meshs(g_idxM).Vertices(idxVert).VectorsT.Y
'    Pz = g_Meshs(g_idxM).Vertices(idxVert).VectorsT.Z
'
'    Nx = Normal.X
'    Ny = Normal.Y
'    Nz = Normal.Z
'
'
'    'Get vector V
'    Vx = eyeX - PX
'    Vy = eyeY - PY
'    Vz = eyeZ - Pz
'    Vlen = Sqr(Vx * Vx + Vy * Vy + Vz * Vz)
'    Vx = Vx / Vlen
'    Vy = Vy / Vlen
'    Vz = Vz / Vlen
''
''    ' Consider each lightsource
''    For Each Light_Source In LightSources
''        ' Find vector L not normalized
'        g_Lights(0).OriginT = MatrixMultiplyVector(g_Cameras(0).ViewMatrix, g_Lights(0).Origin)
'        Lx = g_Lights(0).OriginT.X - PX
'        Ly = g_Lights(0).OriginT.Y - PY
'        Lz = g_Lights(0).OriginT.Z - Pz
''
'        ' Normalize vector L
'        Llen = Sqr(Lx * Lx + Ly * Ly + Lz * Lz)
'        Lx = Lx / Llen
'        Ly = Ly / Llen
'        Lz = Lz / Llen
'
'        ' See if the viewpoint is on the same side
'        ' of the surface as the Surface Normal
'        VdotN = Vx * Nx + Vy * Ny + Vz * Nz
'
'        ' See if the LightSrc is on the same side
'        ' of the surface as the Surface Normal
'        LdotN = Lx * Nx + Ly * Ny + Lz * Nz
'
'        ' We only have specular and diffuse lighting
'        ' components if the viewpoint and light are
'        ' on the same side of the surface, and if we
'        ' are not shadowed
'
'            DiffKr = g_Lights(0).Color.R * sng1Div255
'            DiffKg = g_Lights(0).Color.G * sng1Div255
'            DiffKb = g_Lights(0).Color.B * sng1Div255
'
'            AmbKr = g_Materials(g_Meshs(g_idxM).Faces(g_idxF).idxMat).Diffuse.R * sng1Div255 '\ 255 '0.3
'            AmbKg = g_Materials(g_Meshs(g_idxM).Faces(g_idxF).idxMat).Diffuse.G * sng1Div255 '\ 255 '0.3
'            AmbKb = g_Materials(g_Meshs(g_idxM).Faces(g_idxF).idxMat).Diffuse.B * sng1Div255 '\ 255 '0.3
'
'            SpecK = g_Materials(g_idxMt).SpecK ' 0.35
'            SpecN = g_Materials(g_idxMt).SpecN ' 20
'
'        If (VdotN >= 0) And (LdotN >= 0) Then 'And (Not Shadowed) Then
'            ' The light is shining on the surface
'
'            ' ####################
'            ' # Diffuse lighting #
'            ' ####################
'            ' There is a diffuse component
'            TotalR = TotalR + g_Lights(0).Color.R * DiffKr * LdotN
'            TotalG = TotalG + g_Lights(0).Color.G * DiffKg * LdotN
'            TotalB = TotalB + g_Lights(0).Color.B * DiffKb * LdotN
'
'            ' #####################
'            ' # Specular lighting #
'            ' #####################
'            ' Find the light mirror vector LM
'            LMx = 2 * Nx * LdotN - Lx
'            LMy = 2 * Ny * LdotN - Ly
'            LMz = 2 * Nz * LdotN - Lz
'
'            ' Get LM dot V
'            LMdotV = LMx * Vx + LMy * Vy + LMz * Vz
'            If LMdotV > 0 Then
'                spec = SpecK * (LMdotV ^ SpecN)
'                TotalR = TotalR + g_Lights(0).Color.R * spec
'                TotalG = TotalG + g_Lights(0).Color.G * spec
'                TotalB = TotalB + g_Lights(0).Color.B * spec
'            End If
'        End If
''    Next Light_Source
''
'    ' ####################
'    ' # Ambient lighting #
'    ' ####################
'    TotalR = TotalR + g_Lights(0).Ambient.R * AmbKr
'    TotalG = TotalG + g_Lights(0).Ambient.G * AmbKg
'    TotalB = TotalB + g_Lights(0).Ambient.B * AmbKb
'    If TotalR > 255 Then TotalR = 255
'    If TotalG > 255 Then TotalG = 255
'    If TotalB > 255 Then TotalB = 255
'    If TotalR < 0 Then TotalR = 0
'    If TotalG < 0 Then TotalG = 0
'    If TotalB < 0 Then TotalB = 0
'
'    ' Set the ByRef-passed color components
'    Shade.R = TotalR
'    Shade.G = TotalG
'    Shade.B = TotalB
'
'End Function


