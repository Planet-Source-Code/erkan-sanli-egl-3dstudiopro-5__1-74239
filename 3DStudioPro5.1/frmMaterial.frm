VERSION 5.00
Begin VB.Form frmMaterial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Material"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Filter"
      Height          =   855
      Left            =   6240
      TabIndex        =   34
      Top             =   5280
      Width           =   1575
      Begin VB.OptionButton optFilter 
         Caption         =   "Bilinear"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   20
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Colors"
      Height          =   2175
      Left            =   120
      TabIndex        =   15
      Top             =   4320
      Width           =   1935
      Begin EGL_3DStudioPro5.UpDown udSpecN 
         Height          =   300
         Left            =   1080
         TabIndex        =   42
         Top             =   1560
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
      End
      Begin EGL_3DStudioPro5.UpDown udSpecK 
         Height          =   300
         Left            =   1080
         TabIndex        =   41
         Top             =   1200
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
      End
      Begin VB.PictureBox picAmbient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1080
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   39
         Top             =   360
         Width           =   600
      End
      Begin VB.PictureBox picDiffuse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1080
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   16
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label11 
         Caption         =   "Ambient"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Diffuse"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Opacity"
      Height          =   855
      Left            =   4800
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
      Begin EGL_3DStudioPro5.UpDown udOpacity 
         Height          =   300
         Left            =   360
         TabIndex        =   37
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Map"
      Height          =   5055
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.Frame Frame2 
         Caption         =   "Coordinates"
         Height          =   1575
         Left            =   120
         TabIndex        =   21
         Top             =   3360
         Width           =   4215
         Begin VB.TextBox txtUOffset 
            Height          =   300
            Left            =   480
            TabIndex        =   27
            Text            =   "UOffset"
            Top             =   600
            Width           =   600
         End
         Begin VB.TextBox txtVOffset 
            Height          =   300
            Left            =   480
            TabIndex        =   26
            Text            =   "VOffset"
            Top             =   960
            Width           =   600
         End
         Begin VB.TextBox txtUScale 
            Height          =   300
            Left            =   1440
            TabIndex        =   25
            Text            =   "UScale"
            Top             =   600
            Width           =   600
         End
         Begin VB.TextBox txtVScale 
            Height          =   300
            Left            =   1440
            TabIndex        =   24
            Text            =   "VScale"
            Top             =   960
            Width           =   600
         End
         Begin VB.TextBox txtRotate 
            Height          =   300
            Left            =   2760
            TabIndex        =   23
            Text            =   "WRotate"
            Top             =   600
            Width           =   600
         End
         Begin VB.CommandButton cmdCoordApply 
            Caption         =   "Apply"
            Height          =   375
            Left            =   3120
            TabIndex        =   22
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Offset"
            Height          =   300
            Left            =   480
            TabIndex        =   33
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Scale"
            Height          =   300
            Left            =   1440
            TabIndex        =   32
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label3 
            Caption         =   "U"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   600
            Width           =   135
         End
         Begin VB.Label Label4 
            Caption         =   "V"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Rotate"
            Height          =   300
            Left            =   2760
            TabIndex        =   29
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label8 
            Caption         =   "W"
            Height          =   255
            Left            =   2520
            TabIndex        =   28
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdMeshTex 
         Caption         =   "Browse"
         Height          =   375
         Left            =   3240
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optFlip 
         Caption         =   "Vertical"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.OptionButton optFlip 
         Caption         =   "Horizontal"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton optFlip 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Left            =   240
         ScaleHeight     =   120
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1800
      End
      Begin VB.CommandButton cmdMirrorApply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   2760
         Width           =   615
      End
      Begin VB.CheckBox chkMirror 
         Height          =   200
         Index           =   0
         Left            =   2520
         TabIndex        =   5
         Top             =   1560
         Width           =   200
      End
      Begin VB.CheckBox chkMirror 
         Height          =   200
         Index           =   1
         Left            =   2520
         TabIndex        =   4
         Top             =   1800
         Width           =   200
      End
      Begin VB.Label lblLoadTexture 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "-"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "-"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Flip"
         Height          =   180
         Left            =   3120
         TabIndex        =   11
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label9 
         Caption         =   "U"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "V"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Mirror"
         Height          =   300
         Left            =   2280
         TabIndex        =   6
         Top             =   1320
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Material List"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstMaterials 
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private WithEvents frmColorDiffuse     As frmColorDialog
Attribute frmColorDiffuse.VB_VarHelpID = -1

Dim dibBuffer   As DIB
Dim idxSelMat   As Integer

Private Sub cmdClose_Click()
    
    Destroy dibBuffer
    Unload Me

End Sub

Private Sub Form_Load()

    Dim idx As Integer
        
    For idx = 0 To g_idxMat
        lstMaterials.AddItem CStr(idx) & ": " & g_Materials(idx).Name
    Next
    idxSelMat = 0
    lstMaterials.Selected(idxSelMat) = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not frmColorDiffuse Is Nothing Then Set frmColorDiffuse = Nothing

End Sub

Private Sub lstMaterials_Click()
    
    idxSelMat = lstMaterials.ListIndex
    
    If g_Materials(idxSelMat).MapUse Then
        dibBuffer.Width = g_Materials(idxSelMat).Texture.dibTex.Width * 2
        dibBuffer.Height = g_Materials(idxSelMat).Texture.dibTex.Height * 2
        Call CreateArray(dibBuffer)
        Call Pic2Array(dibBuffer)
        lblLoadTexture.Caption = "Loaded"
    Else
        lblLoadTexture.Caption = "File Not Found"
        picPreview.Cls
    End If
    txtUOffset.Text = g_Materials(idxSelMat).Texture.UOffset
    txtVOffset.Text = g_Materials(idxSelMat).Texture.VOffset
    txtUScale.Text = g_Materials(idxSelMat).Texture.UScale
    txtVScale.Text = g_Materials(idxSelMat).Texture.VScale
    txtRotate.Text = g_Materials(idxSelMat).Texture.Angle
    udSpecK.Value = g_Materials(idxSelMat).SpecK * 100
    udSpecN.Value = g_Materials(idxSelMat).SpecN
    lblName.Caption = g_Materials(idxSelMat).Texture.FileName
    picDiffuse.BackColor = RGB(g_Materials(idxSelMat).Diffuse.R, g_Materials(idxSelMat).Diffuse.G, g_Materials(idxSelMat).Diffuse.B)
    udOpacity.Value = g_Materials(idxSelMat).Opacity
    optFilter(g_Materials(idxSelMat).Texture.Filtre).Value = True
    Call RefreshPreview

End Sub

Private Sub chkMirror_Click(Index As Integer)
    
    Dim Ustat As Byte
    Dim Vstat As Byte
    
    If g_Materials(idxSelMat).MapUse = False Then Exit Sub

    Ustat = IIf(chkMirror(0).Value, 1, 0)
    Vstat = IIf(chkMirror(1).Value, 10, 0)
    g_Materials(idxSelMat).Texture.Mirroring = Ustat + Vstat
    Call RefreshPreview

End Sub

Private Sub optFilter_Click(Index As Integer)
    
    g_Materials(idxSelMat).Texture.Filtre = Index

End Sub

Private Sub optFlip_Click(Index As Integer)
    
    If g_Materials(idxSelMat).MapUse = False Then Exit Sub

    g_Materials(idxSelMat).Texture.Flipping = CByte(Index)
    Call RefreshPreview

End Sub

Private Sub RefreshPreview()

'    g_Materials(idxSelMat).Opacity = udOpacity.Value
    
    If g_Materials(idxSelMat).MapUse = False Then Exit Sub
    
    Call SetStretchBltMode(picPreview.hDC, vbPaletteModeNone)
    
    Select Case g_Materials(idxSelMat).Texture.Mirroring
        
        Case 0
            Call StretchBlt(dibBuffer.hDC, 0, 0, dibBuffer.Width, dibBuffer.Height, _
                            g_Materials(idxSelMat).Texture.dibTex.hDC, 0, 0, g_Materials(idxSelMat).Texture.dibTex.Width, g_Materials(idxSelMat).Texture.dibTex.Height, vbSrcCopy)
        
        Case 1
            Call StretchBlt(dibBuffer.hDC, 0, 0, g_Materials(idxSelMat).Texture.dibTex.Width, dibBuffer.Height, _
                            g_Materials(idxSelMat).Texture.dibTex.hDC, 0, 0, g_Materials(idxSelMat).Texture.dibTex.Width, g_Materials(idxSelMat).Texture.dibTex.Height, vbSrcCopy)
            Call StretchBlt(dibBuffer.hDC, g_Materials(idxSelMat).Texture.dibTex.Width, 0, g_Materials(idxSelMat).Texture.dibTex.Width, dibBuffer.Height, _
                            g_Materials(idxSelMat).Texture.dibTex.hDC, g_Materials(idxSelMat).Texture.dibTex.Width, 0, -g_Materials(idxSelMat).Texture.dibTex.Width, g_Materials(idxSelMat).Texture.dibTex.Height, vbSrcCopy)
        
        Case 10
            Call StretchBlt(dibBuffer.hDC, 0, 0, dibBuffer.Width, g_Materials(idxSelMat).Texture.dibTex.Height, _
                            g_Materials(idxSelMat).Texture.dibTex.hDC, 0, 0, g_Materials(idxSelMat).Texture.dibTex.Width, g_Materials(idxSelMat).Texture.dibTex.Height, vbSrcCopy)
            Call StretchBlt(dibBuffer.hDC, 0, g_Materials(idxSelMat).Texture.dibTex.Height, dibBuffer.Width, g_Materials(idxSelMat).Texture.dibTex.Height, _
                            g_Materials(idxSelMat).Texture.dibTex.hDC, 0, g_Materials(idxSelMat).Texture.dibTex.Height, g_Materials(idxSelMat).Texture.dibTex.Width, -g_Materials(idxSelMat).Texture.dibTex.Height, vbSrcCopy)
        Case 11
            Call StretchBlt(dibBuffer.hDC, 0, 0, g_Materials(idxSelMat).Texture.dibTex.Width, g_Materials(idxSelMat).Texture.dibTex.Height, _
                            g_Materials(idxSelMat).Texture.dibTex.hDC, 0, 0, g_Materials(idxSelMat).Texture.dibTex.Width, g_Materials(idxSelMat).Texture.dibTex.Height, vbSrcCopy)
            Call StretchBlt(dibBuffer.hDC, g_Materials(idxSelMat).Texture.dibTex.Width, 0, g_Materials(idxSelMat).Texture.dibTex.Width, g_Materials(idxSelMat).Texture.dibTex.Height, _
                            g_Materials(idxSelMat).Texture.dibTex.hDC, g_Materials(idxSelMat).Texture.dibTex.Width, 0, -g_Materials(idxSelMat).Texture.dibTex.Width, g_Materials(idxSelMat).Texture.dibTex.Height, vbSrcCopy)
            Call StretchBlt(dibBuffer.hDC, 0, g_Materials(idxSelMat).Texture.dibTex.Height, g_Materials(idxSelMat).Texture.dibTex.Width, g_Materials(idxSelMat).Texture.dibTex.Height, _
                            g_Materials(idxSelMat).Texture.dibTex.hDC, 0, g_Materials(idxSelMat).Texture.dibTex.Height, g_Materials(idxSelMat).Texture.dibTex.Width, -g_Materials(idxSelMat).Texture.dibTex.Height, vbSrcCopy)
            Call StretchBlt(dibBuffer.hDC, g_Materials(idxSelMat).Texture.dibTex.Width, g_Materials(idxSelMat).Texture.dibTex.Height, g_Materials(idxSelMat).Texture.dibTex.Width, g_Materials(idxSelMat).Texture.dibTex.Height, _
                            g_Materials(idxSelMat).Texture.dibTex.hDC, g_Materials(idxSelMat).Texture.dibTex.Width, g_Materials(idxSelMat).Texture.dibTex.Height, -g_Materials(idxSelMat).Texture.dibTex.Width, -g_Materials(idxSelMat).Texture.dibTex.Height, vbSrcCopy)
    End Select
    
    Select Case g_Materials(idxSelMat).Texture.Flipping
        Case 0
            Call StretchBlt(picPreview.hDC, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight, _
                            dibBuffer.hDC, 0, 0, dibBuffer.Width, dibBuffer.Height, vbSrcCopy)
        Case 1
            Call StretchBlt(picPreview.hDC, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight, _
                            dibBuffer.hDC, 0, dibBuffer.Height, dibBuffer.Width, -dibBuffer.Height, vbSrcCopy)
        Case 2
            Call StretchBlt(picPreview.hDC, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight, _
                            dibBuffer.hDC, dibBuffer.Width, 0, -dibBuffer.Width, dibBuffer.Height, vbSrcCopy)
    End Select
    
End Sub

Private Sub cmdCoordApply_Click()
    
    If g_Materials(idxSelMat).MapUse = False Then Exit Sub

    g_Materials(idxSelMat).Texture.UScale = VerifyText(txtUScale)
    g_Materials(idxSelMat).Texture.VScale = VerifyText(txtVScale)
    g_Materials(idxSelMat).Texture.UOffset = VerifyText(txtUOffset)
    g_Materials(idxSelMat).Texture.VOffset = VerifyText(txtVOffset)
    g_Materials(idxSelMat).Texture.Angle = VerifyText(txtRotate)

End Sub

Private Sub cmdMirrorApply_Click()

    If g_Materials(idxSelMat).MapUse = False Then Exit Sub
    
    Select Case g_Materials(idxSelMat).Texture.Flipping
        Case 0
            Call StretchBlt(g_Materials(idxSelMat).Texture.dibTexT.hDC, 0, 0, g_Materials(idxSelMat).Texture.dibTexT.Width, g_Materials(idxSelMat).Texture.dibTexT.Height, _
                            dibBuffer.hDC, 0, 0, dibBuffer.Width, dibBuffer.Height, vbSrcCopy)
        Case 1
            Call StretchBlt(g_Materials(idxSelMat).Texture.dibTexT.hDC, 0, 0, g_Materials(idxSelMat).Texture.dibTexT.Width, g_Materials(idxSelMat).Texture.dibTexT.Height, _
                            dibBuffer.hDC, 0, dibBuffer.Height, dibBuffer.Width, -dibBuffer.Height, vbSrcCopy)
        Case 2
            Call StretchBlt(g_Materials(idxSelMat).Texture.dibTexT.hDC, 0, 0, g_Materials(idxSelMat).Texture.dibTexT.Width, g_Materials(idxSelMat).Texture.dibTexT.Height, _
                            dibBuffer.hDC, dibBuffer.Width, 0, -dibBuffer.Width, dibBuffer.Height, vbSrcCopy)
    End Select

End Sub

Private Sub picDiffuse_Click()
    
    If g_bLoadOK = False Then Exit Sub
    
    If frmColorDiffuse Is Nothing Then
        Set frmColorDiffuse = New frmColorDialog
        frmColorDiffuse.Caption = "Select Mesh Color"
    End If
    frmColorDiffuse.Show vbModal, frmMaterial

End Sub

Private Sub txtRotate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then cmdCoordApply_Click

End Sub

Private Sub txtUOffset_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdCoordApply_Click

End Sub

Private Sub txtUScale_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdCoordApply_Click

End Sub

Private Sub txtVOffset_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdCoordApply_Click

End Sub

Private Sub txtVScale_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdCoordApply_Click

End Sub

Private Sub frmColorDiffuse_Active(R As Integer, G As Integer, B As Integer)

    R = g_Materials(idxSelMat).Diffuse.R
    G = g_Materials(idxSelMat).Diffuse.G
    B = g_Materials(idxSelMat).Diffuse.B

End Sub

Private Sub frmColorDiffuse_Change(R As Integer, G As Integer, B As Integer)

    g_Materials(idxSelMat).Diffuse = ColorSet(R, G, B)
    picDiffuse.BackColor = RGB(R, G, B)

End Sub

Private Sub cmdMeshTex_Click()

    Dim FileName As String

    FileName = GetMyPicturesFolder(Me.hwnd)
    FileName = PicturePath(FileName, "Texture File")
    Call LoadTexture(idxSelMat, FileName)

End Sub

Private Sub udOpacity_Change()

    g_Materials(idxSelMat).Opacity = udOpacity.Value
    g_Materials(idxSelMat).OpcValue1 = udOpacity.Value * 0.01
    g_Materials(idxSelMat).OpcValue2 = 1 - g_Materials(idxSelMat).OpcValue1

End Sub

Private Sub udSpecK_Change()
    
    g_Materials(idxSelMat).SpecK = udSpecK.Value * 0.01

End Sub

Private Sub udSpecN_Change()

    g_Materials(idxSelMat).SpecN = udSpecN.Value

End Sub
