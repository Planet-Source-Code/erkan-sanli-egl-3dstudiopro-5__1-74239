VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "3D Studio Pro 5"
   ClientHeight    =   5010
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   334
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrProcess 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   0
   End
   Begin VB.PictureBox picLoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   2520
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
         Begin VB.Menu mnuImport3DS 
            Caption         =   "3DS"
         End
         Begin VB.Menu mnuImportOBJ 
            Caption         =   "OBJ"
         End
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Begin VB.Menu mnuExportBMP 
            Caption         =   "BMP"
         End
      End
      Begin VB.Menu tire 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuRenderTime 
         Caption         =   "Show Render Time"
      End
   End
   Begin VB.Menu mnuVisualStyleC 
      Caption         =   "Visual &Style"
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Dot"
         Index           =   0
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Box"
         Index           =   1
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Wireframe"
         Index           =   2
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Wireframe Map"
         Index           =   3
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Flat"
         Index           =   4
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Flat Map"
         Index           =   5
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Gouraud"
         Index           =   6
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Gouraud Map"
         Index           =   7
      End
      Begin VB.Menu tire2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackface 
         Caption         =   "Backface"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuMesh 
         Caption         =   "Mesh"
      End
      Begin VB.Menu mnuCamera 
         Caption         =   "Camera"
      End
      Begin VB.Menu mnuMaterial 
         Caption         =   "Material"
      End
      Begin VB.Menu mnuLight 
         Caption         =   "Light"
      End
      Begin VB.Menu mnuBackground 
         Caption         =   "Background"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuControl 
         Caption         =   "Control"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cTimer      As New clsTiming
Dim MinWidth    As Long
Dim MinHeight   As Long
Dim FormLoadOK  As Boolean

Dim matOutput   As MATRIX
Dim matOutputShadow   As MATRIX
Dim matShadow   As MATRIX

Private Sub Form_Load()

    Me.ScaleMode = vbPixels
    MinWidth = Screen.TwipsPerPixelX * 648
    MinHeight = Screen.TwipsPerPixelY * 526
    Me.Width = MinWidth
    Me.Height = MinHeight
    
    g_idxMat = -1
    g_idxMesh = -1
    
    BFileName = App.Path & "\Default.jpg"
    BColor1 = ColorSet(200, 200, 250)
    BColor2 = ColorSet(200, 250, 250)
    g_BkType = BkPic
    ValZoom = 1000

    mnuRenderTime_Click
    FormLoadOK = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim idx As Integer
    
    tmrProcess.Enabled = False
    Destroy g_dibCanvas
    Destroy g_dibBack
    If g_idxMat > -1 Then
        For idx = 0 To g_idxMat
            Destroy g_Materials(idx).Texture.dibTex
            Destroy g_Materials(idx).Texture.dibTexT
        Next
    End If
    End

End Sub

Private Sub Form_Resize()
    
    If FormLoadOK Then ResizeCanvas

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then PopupMenu mnuVisualStyleC

End Sub

Private Sub ResizeCanvas()
        
    
    If frmMain.WindowState = vbMinimized Then Exit Sub
    If Me.Width < MinWidth Then Me.Width = MinWidth
    If Me.Height < MinHeight Then Me.Height = MinHeight
    
    g_CanvasWidth = Me.ScaleWidth
    g_CanvasHeight = Me.ScaleHeight
    g_OriginX = CLng(g_CanvasWidth * 0.5)
    g_OriginY = CLng(g_CanvasHeight * 0.5)
    picCanvas.Move 0, 0, g_CanvasWidth, g_CanvasHeight

'Canvas
    g_dibCanvas.Width = g_CanvasWidth
    g_dibCanvas.Height = g_CanvasHeight
    Call CreateArray(g_dibCanvas)
    Call Pic2Array(g_dibCanvas)

'Back
    g_dibBack.Width = g_CanvasWidth
    g_dibBack.Height = g_CanvasHeight
    Call CreateArray(g_dibBack)
    Call Pic2Array(g_dibCanvas)
    RefreshBack

'Canvas Rect
    g_SafeFrame.Left = 0
    g_SafeFrame.Top = 0
    g_SafeFrame.Right = g_CanvasWidth - 1
    g_SafeFrame.Bottom = g_CanvasHeight - 1

End Sub

Private Sub Screening(Vertices() As VERTEX, Screen() As POINTAPI, matOut As MATRIX)
    
    Dim idx         As Long
        
    For idx = 0 To UBound(Vertices)
        Vertices(idx).VectorsT = MatrixMultiplyVector(matOut, Vertices(idx).Vectors)
        Vertices(idx).VectorsS = Vertices(idx).VectorsT
        If Vertices(idx).VectorsT.Z <> 0 Then
            Vertices(idx).VectorsT.X = (Vertices(idx).VectorsT.X / Vertices(idx).VectorsT.Z)
            Vertices(idx).VectorsT.Y = (Vertices(idx).VectorsT.Y / Vertices(idx).VectorsT.Z)
        End If
        Screen(idx).X = Vertices(idx).VectorsT.X * ValZoom + g_OriginX
        Screen(idx).Y = Vertices(idx).VectorsT.Y * ValZoom + g_OriginY
    Next idx
    
End Sub

Private Sub tmrProcess_Timer()
    
    Dim idx As Integer
    
    If g_bLoadOK Then
        DoEvents
        If g_blnRenderTime Then Call cTimer.Reset
        Call RefreshEvents(picCanvas.hwnd)
        
        g_Cameras(0).ViewMatrix = ViewMatrix
        g_MeshsWorldMatrix = WorldMatrix
        ValZoom = g_Cameras(0).Zoom * 1000
        
        BitBlt g_dibCanvas.hDC, 0, 0, g_CanvasWidth, g_CanvasHeight, g_dibBack.hDC, 0, 0, vbSrcCopy
        
        
        If g_Lights(0).ShadowEnable Then
            matShadow = MatrixShadow(g_Lights(0).Origin, g_Lights(0).ShadowPlane)
            matOutputShadow = MatrixMultiply(g_MeshsWorldMatrix, matShadow)
            matOutputShadow = MatrixMultiply(matOutputShadow, g_Cameras(0).ViewMatrix)
            For idx = 0 To g_idxMesh
                If g_Meshs(idx).idxVert > 0 Then
                    If g_VStyle <> Box Then
                        Call Screening(g_Meshs(idx).Vertices, g_Meshs(idx).Screen, matOutputShadow)
                    End If
                End If
            Next
            Call RenderShadow
        End If
        
        matOutput = MatrixMultiply(g_MeshsWorldMatrix, g_Cameras(0).ViewMatrix)
        For idx = 0 To g_idxMesh
            If g_Meshs(idx).idxVert > 0 Then
                If g_VStyle = Box Then
                    Call Screening(g_Meshs(idx).Box, g_Meshs(idx).BoxScreen, matOutput)
                Else
                    Call Screening(g_Meshs(idx).Vertices, g_Meshs(idx).Screen, matOutput)
                End If
            End If
        Next
        
        Call Render
        
        BitBlt picCanvas.hDC, 0, 0, g_CanvasWidth, g_CanvasHeight, g_dibCanvas.hDC, 0, 0, vbSrcCopy

        If g_blnRenderTime Then
            picCanvas.CurrentY = 2
            picCanvas.Print cTimer.Elapsed
        End If
    Else
        tmrProcess.Enabled = False
    End If
 
End Sub

'=====================================================================================
'Start Menus

'Menu : File

Private Sub mnuImport3DS_Click()
    
    Dim idx As Integer
    Dim FileName As String
    
    Set g_CDialog = New clsCommonDialog
    With g_CDialog
        .Filter = "3D Studio File|*.3ds"
        .InitDir = App.Path & "\Sample"
        .FileName = ""
        .ShowOpen
        If .FileName = vbNullString Then Exit Sub
        If Not FileExist(.FileName) Then Exit Sub
        g_Path3DS = GetFilePath(.FileName)
        Set g_File3DS = New clsFile3DS
        g_bLoadOK = g_File3DS.Read3DS(.FileName)
        Set g_File3DS = Nothing
    End With
    Set g_CDialog = Nothing
    If g_bLoadOK Then
        Call BestFit
        For idx = 0 To g_idxMat
            FileName = g_Path3DS & g_Materials(idx).Texture.FileName
            Call LoadTexture(idx, FileName)
        Next
        mnuEdit.Enabled = True
        tmrProcess.Enabled = True
        frmMain.picCanvas.SetFocus
    Else
        mnuEdit.Enabled = False
        MsgBox "Loading Error"
    End If
    
    mnuVisualStyle_Click Gouraud
    
End Sub

Private Sub mnuImportOBJ_Click()
    
    Dim idx As Integer
    Dim FileName As String
    
    Set g_CDialog = New clsCommonDialog
    With g_CDialog
        .Filter = "Wavefront Object|*.obj"
        .InitDir = App.Path & "\Sample"
        .FileName = ""
        .ShowOpen
        If .FileName = vbNullString Then Exit Sub
        If Not FileExist(.FileName) Then Exit Sub
        g_PathOBJ = GetFilePath(.FileName)
        Set g_FileOBJ = New clsFileOBJ
        g_bLoadOK = g_FileOBJ.ReadOBJ(.FileName)
        Set g_FileOBJ = Nothing
    End With
    Set g_CDialog = Nothing
    If g_bLoadOK Then
        Call BestFit
'        For idx = 0 To g_idxMat
'            FileName = g_Path3DS & g_Materials(idx).Texture.FileName
'            Call LoadTexture(idx, FileName)
'        Next
        mnuEdit.Enabled = True
        tmrProcess.Enabled = True
        frmMain.picCanvas.SetFocus
    Else
        mnuEdit.Enabled = False
        MsgBox "Loading Error"
    End If
    
    mnuVisualStyle_Click FlatMap

End Sub

Private Sub mnuExportBMP_Click()
    
    Dim cfBMP   As clsFileBMP
    Dim Result  As Long
    Dim strMsg  As String
    Dim FileName As String
    
    If g_bLoadOK = False Then Exit Sub
    
    tmrProcess.Enabled = False
    Set g_CDialog = New clsCommonDialog
ReOpen:
    With g_CDialog
        .Filter = "24-bit Bitmap |*.bmp"
        .DefaultExt = ".bmp"
        .DialogTitle = "Save Bitmap"
        .InitDir = GetMyPicturesFolder(Me.hwnd)
        .ShowSave
        FileName = Left$(.FileName, InStr(1, .FileName, Chr$(0)) - 1)
        If Len(FileName) <> 0 Then
            If FileExist(FileName) Then
                strMsg = FileName & " already exists." & vbCrLf & "Do you want to replace it."
                Result = MsgBox(strMsg, vbYesNo)
                If Result = vbYes Then
                    Kill FileName
                Else
                    GoTo ReOpen
                End If
            End If
            Set cfBMP = New clsFileBMP
            Call cfBMP.WriteBMP24(FileName)
            Set cfBMP = Nothing
            strMsg = FileName & vbCrLf & " saved."
            MsgBox strMsg
        End If
    End With
    tmrProcess.Enabled = True

End Sub

Private Sub mnuExit_Click()
    
    Unload Me

End Sub

'Menu : View

Private Sub mnuRenderTime_Click()
    
    mnuRenderTime.Checked = Not mnuRenderTime.Checked
    g_blnRenderTime = mnuRenderTime.Checked

End Sub

'Menu : Visual Style

Public Sub mnuVisualStyle_Click(Index As Integer)
    
    Dim idx As Integer
    
    For idx = 0 To mnuVisualStyle.UBound
        mnuVisualStyle(idx).Checked = False
    Next
    mnuVisualStyle(Index).Checked = True
    g_VStyle = Index

End Sub

Private Sub mnuBackface_Click()

    mnuBackface.Checked = Not mnuBackface.Checked
    g_blnDoubleSide = mnuBackface.Checked

End Sub

'Menu : Edit

Private Sub mnuMesh_Click()

    frmMesh.Show vbModal, frmMain

End Sub

Private Sub mnuCamera_Click()
    
    frmCamera.Show vbModal, frmMain

End Sub

Private Sub mnuMaterial_Click()
    
    frmMaterial.Show vbModal, frmMain

End Sub

Private Sub mnuLight_Click()
    
    frmLight.Show vbModal, frmMain

End Sub

Private Sub mnuBackground_Click()
    
    frmBackground.Show vbModeless, frmMain

End Sub

'Menu : Help

Private Sub mnuControl_Click()
    
    frmControl.Show vbModal, frmMain

End Sub

'End Menus
'========================================================================================
