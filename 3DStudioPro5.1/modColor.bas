Attribute VB_Name = "modColor"
Option Explicit

Public Const sng1Div3     As Single = 0.3333333
Public Const sng1Div255   As Single = 0.0039215

Public Function ColorSet(Red As Integer, Green As Integer, Blue As Integer) As COLORRGB_INT

    ColorSet.R = Red
    ColorSet.G = Green
    ColorSet.B = Blue

End Function

Public Function ColorScale(C1 As COLORRGB_INT, Scalar As Single) As COLORRGB_INT

    ColorScale.R = C1.R * Scalar
    ColorScale.G = C1.G * Scalar
    ColorScale.B = C1.B * Scalar

End Function

Public Function ColorAdd(C1 As COLORRGB_INT, C2 As COLORRGB_INT) As COLORRGB_INT

    ColorAdd.R = C1.R + C2.R
    ColorAdd.G = C1.G + C2.G
    ColorAdd.B = C1.B + C2.B

End Function

Public Function ColorSub(C1 As COLORRGB_INT, C2 As COLORRGB_INT) As COLORRGB_INT

    ColorSub.R = C1.R - C2.R
    ColorSub.G = C1.G - C2.G
    ColorSub.B = C1.B - C2.B

End Function

Public Function ColorPlus(ByVal R As Integer, ByVal G As Integer, ByVal B As Integer, Val As Integer) As RGBQUAD
    
    ColorPlus.rgbRed = CByte(ColorLimitVal(R + Val))
    ColorPlus.rgbGreen = CByte(ColorLimitVal(G + Val))
    ColorPlus.rgbBlue = CByte(ColorLimitVal(B + Val))

End Function

Function ColorInterpolate(CC1 As RGBQUAD, CC2 As RGBQUAD, Alpha As Single) As RGBQUAD

    Dim C1 As COLORRGB_INT
    Dim C2 As COLORRGB_INT
    C1.R = CC1.rgbRed
    C1.G = CC1.rgbGreen
    C1.B = CC1.rgbBlue
    C2.R = CC2.rgbRed
    C2.G = CC2.rgbGreen
    C2.B = CC2.rgbBlue

    ColorInterpolate.rgbRed = CByte(ColorLimitVal(((C2.R - C1.R) * Alpha) + C1.R))
    ColorInterpolate.rgbGreen = CByte(ColorLimitVal(((C2.G - C1.G) * Alpha) + C1.G))
    ColorInterpolate.rgbBlue = CByte(ColorLimitVal(((C2.B - C1.B) * Alpha) + C1.B))

End Function

Function ColorInterpolateInt(C1 As COLORRGB_INT, C2 As COLORRGB_INT, Alpha As Single) As COLORRGB_INT

    ColorInterpolateInt.R = ((C2.R - C1.R) * Alpha) + C1.R
    ColorInterpolateInt.G = ((C2.G - C1.G) * Alpha) + C1.G
    ColorInterpolateInt.B = ((C2.B - C1.B) * Alpha) + C1.B

End Function

Public Sub ColorLimit(C1 As COLORRGB_INT)

    If (C1.R > 255) Then C1.R = 255 Else If (C1.R < 0) Then C1.R = 0
    If (C1.G > 255) Then C1.G = 255 Else If (C1.G < 0) Then C1.G = 0
    If (C1.B > 255) Then C1.B = 255 Else If (C1.B < 0) Then C1.B = 0

End Sub

Public Function ColorLimitVal(ByVal iColor As Integer) As Integer

    ColorLimitVal = iColor
    If ColorLimitVal < 0 Then ColorLimitVal = 0
    If ColorLimitVal > 255 Then ColorLimitVal = 255
    
End Function

Public Function ColIntToColByte(C1 As COLORRGB_INT) As RGBQUAD

    ColIntToColByte.rgbRed = CByte(C1.R)
    ColIntToColByte.rgbGreen = CByte(C1.G)
    ColIntToColByte.rgbBlue = CByte(C1.B)

End Function

Public Function ColIntToColSng(C1 As COLORRGB_INT) As COLORRGB_SNG

    ColIntToColSng.R = CSng(C1.R)
    ColIntToColSng.G = CSng(C1.G)
    ColIntToColSng.B = CSng(C1.B)

End Function

Public Function ColorGray(C1 As COLORRGB_INT) As Integer

    ColorGray = CInt((C1.R + C1.G + C1.B) * sng1Div3)
    If (ColorGray > 255) Then ColorGray = 255
    
End Function

'Public Function ColorAverage(C1 As COLORRGB_INT, C2 As COLORRGB_INT, C3 As COLORRGB_INT) As RGBQUAD
'
'    ColorAverage.rgbRed = CByte(ColorLimitVal(((C1.R + C2.R + C3.R) * sng1Div3)))
'    ColorAverage.rgbGreen = CByte(ColorLimitVal(((C1.G + C2.G + C3.G) * sng1Div3)))
'    ColorAverage.rgbBlue = CByte(ColorLimitVal(((C1.B + C2.B + C3.B) * sng1Div3)))
'
'End Function

Public Function ColorAverage2() As RGBQUAD

    With g_Meshs(g_idxM)
        ColorAverage2.rgbRed = CByte(ColorLimitVal(((.Vertices(.Faces(g_idxF).A).intRGBColor.R + _
                                                    .Vertices(.Faces(g_idxF).B).intRGBColor.R + _
                                                    .Vertices(.Faces(g_idxF).C).intRGBColor.R) * _
                                                    sng1Div3)))
        ColorAverage2.rgbGreen = CByte(ColorLimitVal(((.Vertices(.Faces(g_idxF).A).intRGBColor.G + _
                                                      .Vertices(.Faces(g_idxF).B).intRGBColor.G + _
                                                      .Vertices(.Faces(g_idxF).C).intRGBColor.G) * _
                                                      sng1Div3)))
        ColorAverage2.rgbBlue = CByte(ColorLimitVal(((.Vertices(.Faces(g_idxF).A).intRGBColor.B + _
                                                     .Vertices(.Faces(g_idxF).B).intRGBColor.B + _
                                                     .Vertices(.Faces(g_idxF).C).intRGBColor.B) * _
                                                     sng1Div3)))
    End With

End Function

Public Function ColorDiffuse(C1 As COLORRGB_INT, C2 As COLORRGB_INT) As COLORRGB_INT

    ColorDiffuse.R = C1.R * C2.R * sng1Div255
    ColorDiffuse.G = C1.G * C2.G * sng1Div255
    ColorDiffuse.B = C1.B * C2.B * sng1Div255

End Function

Public Function ColorLongToRGB(lColor As Long) As COLORRGB_INT

    ColorLongToRGB.R = (lColor And &HFF&)
    ColorLongToRGB.G = (lColor And &HFF00&) / &H100&
    ColorLongToRGB.B = (lColor And &HFF0000) / &H10000

End Function

Public Function ColorRGBToLong(C1 As COLORRGB_INT) As Long

    ColorRGBToLong = RGB(C1.R, C1.G, C1.B)

End Function

Public Function RGBQuadToLong(C1 As RGBQUAD) As Long

    RGBQuadToLong = RGB(C1.rgbRed, C1.rgbGreen, C1.rgbBlue)

End Function


Public Function ColorLimit_SNG(C1 As COLORRGB_SNG) As COLORRGB_SNG

    ColorLimit_SNG = C1
    If (ColorLimit_SNG.R > 255) Then ColorLimit_SNG.R = 255
    If (ColorLimit_SNG.G > 255) Then ColorLimit_SNG.G = 255
    If (ColorLimit_SNG.B > 255) Then ColorLimit_SNG.B = 255
    If (ColorLimit_SNG.R < 0) Then ColorLimit_SNG.R = 0
    If (ColorLimit_SNG.G < 0) Then ColorLimit_SNG.G = 0
    If (ColorLimit_SNG.B < 0) Then ColorLimit_SNG.B = 0

End Function


