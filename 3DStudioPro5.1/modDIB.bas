Attribute VB_Name = "modDIB"
Option Explicit

Private Const DIB_RGB_COLORS As Long = 0

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Public Type RGBQUAD
    rgbBlue         As Byte
    rgbGreen        As Byte
    rgbRed          As Byte
    rgbReserved     As Byte
End Type

Private Type BITMAPINFO
    Header          As BITMAPINFOHEADER
    Bits()          As RGBQUAD
End Type

Private Type SAFEARRAYBOUND
    cElements       As Long
    lLbound         As Long
End Type

Private Type SAFEARRAY2D
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As Long
    pvData          As Long
    Bounds(1)       As SAFEARRAYBOUND
End Type

Public Type DIB
    hDC         As Long
    hDIB        As Long
    hOldDIB     As Long
    lpBits      As Long
    Width       As Long
    Height      As Long
    uBI         As BITMAPINFO
    uSA         As SAFEARRAY2D
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long
Private Declare Function VarPtrArray Lib "MSVBVM60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal numBytes As Long)


Public Sub CreateArray(di As DIB, Optional Orientation As Boolean = False)
    
    Call Destroy(di)
    Call CreateDC(di, Orientation)
    
    If (di.hDC <> 0) Then
        di.hDIB = CreateDIBSection(di.hDC, di.uBI, DIB_RGB_COLORS, di.lpBits, ByVal 0&, ByVal 0&)
        If (di.hDIB <> 0) Then
            di.hOldDIB = SelectObject(di.hDC, di.hDIB)
          Else
            Call Destroy(di)
        End If
    End If
        
End Sub

Public Sub CreateArrayFromPicBox2(Picture As PictureBox, di As DIB, Optional Orientation As Boolean = False)

    Call CreateArray(di, Orientation)
    SetStretchBltMode di.hDC, vbPaletteModeNone
    StretchBlt di.hDC, 0, 0, di.Width, di.Height, Picture.hDC, 0, 0, Picture.ScaleWidth, Picture.ScaleHeight, vbSrcCopy
    Call Pic2Array(di)

End Sub

Private Sub CreateDC(di As DIB, Optional Orientation As Boolean = False)

    With di.uBI.Header
        .biBitCount = 32
        .biPlanes = 1
        .biSize = 40
        .biWidth = di.Width
        .biHeight = IIf(Orientation, di.Height, -di.Height)
        .biSizeImage = di.Width * di.Height * 4 '4 * ((.biWidth * .biBitCount + 31) \ 32) * .biHeight '
    End With
    di.hDC = CreateCompatibleDC(0)
    
End Sub

Public Sub Pic2Array(di As DIB)
    
    With di.uSA
        .cbElements = 4
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = di.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = di.Width
        .pvData = di.lpBits
    End With
    Call CopyMemory(ByVal VarPtrArray(di.uBI.Bits()), VarPtr(di.uSA), 4)

End Sub

Public Sub Destroy(di As DIB)

    If (di.hDC <> 0) Then
        If (di.hDIB <> 0) Then
            Call SelectObject(di.hDC, di.hOldDIB)
            Call DeleteObject(di.hDIB)
        End If
        Call DeleteDC(di.hDC)
    End If
    Call ZeroMemory(di.uBI.Header, Len(di.uBI.Header))
    Call CopyMemory(ByVal VarPtrArray(di.uBI.Bits()), 0&, 4)
    di.hDC = 0
    di.hDIB = 0
    di.hOldDIB = 0
    di.lpBits = 0
    
End Sub

Public Sub Clear(di As DIB)

    ZeroMemory di.uBI.Bits(0, 0), di.uBI.Header.biSizeImage ' di.Width * di.Height * 4

End Sub


