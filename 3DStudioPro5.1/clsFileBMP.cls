VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type BITMAPFILEHEADER
    bfType          As Integer
    bfSize          As Long
    bfReserved1     As Integer
    bfReserved2     As Integer
    bfOffBits       As Long
End Type

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

Private Const DIB_RGB_COLORS    As Long = 0
Private Const bpp24             As Byte = 24

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
  
Public Sub WriteBMP24(FileName As String)
    
    Dim hFile               As Integer
    Dim tBFH                As BITMAPFILEHEADER ' 14 bytes
    Dim tBIH                As BITMAPINFOHEADER ' 40 bytes
    Dim aBI()               As Byte
    Dim aData()             As Byte
    Dim BytesPerScanLine    As Long
    
    BytesPerScanLine = ((g_SafeFrame.Right * bpp24 + 31) \ 32) * 4
    ReDim aData(BytesPerScanLine - 1, g_SafeFrame.Bottom)
    
    With tBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = bpp24
        .biWidth = g_SafeFrame.Right
        .biHeight = g_SafeFrame.Bottom
    End With
    
    With tBFH
        .bfType = &H4D42
        .bfSize = Len(tBIH) + BytesPerScanLine * g_SafeFrame.Bottom
        .bfOffBits = Len(tBFH) + Len(tBIH)
    End With
    
    ReDim aBI(1 To 40)
    
    Call CopyMemory(aBI(1), tBIH, 40)
    Call GetDIBits(g_dibCanvas.hDC, g_dibCanvas.hDIB, 0, g_SafeFrame.Bottom, aData(0, 0), tBIH, DIB_RGB_COLORS)
    
'Write file
    hFile = FreeFile
    Open FileName For Binary Access Write As #hFile
      Put #hFile, , tBFH
      Put #hFile, , aBI()
      Put #hFile, , aData()
    Close #hFile
     
End Sub

