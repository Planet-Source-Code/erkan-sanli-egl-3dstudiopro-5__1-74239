VERSION 5.00
Begin VB.Form frmPick 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmPick.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmPick.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   45
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picNew 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Dim rDC     As Long

Private Sub Form_Load()
              
    rDC = GetDC(0&)
    BitBlt Me.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, rDC, 0, 0, vbSrcCopy
    picNew.Move rc.Left, rc.Top
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ReleaseDC 0&, rDC

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim rPixel As Long
    
    rPixel = GetPixel(rDC, X - 16, Y + 16)
    PickRGB = ColorLongToRGB(rPixel)
    picNew.BackColor = RGB(PickRGB.R, PickRGB.G, PickRGB.B)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Unload Me

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Unload Me

End Sub

