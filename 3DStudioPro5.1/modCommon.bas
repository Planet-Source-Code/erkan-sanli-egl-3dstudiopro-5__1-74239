Attribute VB_Name = "modCommon"
Option Explicit

Public Enum VisualStyle
    [Dot]
    [Box]
    [Wireframe]
    [WireframeMap]
    [Flat]
    [FlatMap]
    [Gouraud]
    [GouraudMap]
End Enum

Public Enum CanStat
    [CanOut]
    [CanClip]
    [CanIn]
End Enum

Public Enum BackType
    [Blank]     '= 0
    [Solid]     '= 1
    [Grad1]     '= 2
    [Grad2]     '= 3
    [Grad3]     '= 4
    [BkPic]     '= 5
End Enum

Public Enum Filter
    [None]      '= 0
    [Bilinear]  '= 1
End Enum

Public Type POINTAPI
    X               As Long
    Y               As Long
End Type

Public Type MAPCOORD
    U               As Single
    V               As Single
    'W               As Single
End Type

Public Type COLORRGB_SNG
    R               As Single
    G               As Single
    B               As Single
End Type

Public Type COLORRGB_INT
    R               As Integer
    G               As Integer
    B               As Integer
End Type

Public Type COLORHSV
    H               As Integer
    S               As Integer
    V               As Integer
End Type

Public Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

Public Type VECTOR
    X               As Single
    Y               As Single
    Z               As Single
    W               As Single
End Type

Public Type VERTEX
    Vectors         As VECTOR
    VectorsT        As VECTOR
    VectorsS        As VECTOR
    intRGBColor     As COLORRGB_INT
    Used            As Boolean
End Type

Public Type FACE
    A               As Long
    B               As Long
    C               As Long
    AB              As Boolean
    BC              As Boolean
    CA              As Boolean
    Mapping         As Boolean
    Normal          As VECTOR
    idxMat          As Integer
End Type

Public Type MATRIX
    rc11 As Single: rc12 As Single: rc13 As Single: rc14 As Single
    rc21 As Single: rc22 As Single: rc23 As Single: rc24 As Single
    rc31 As Single: rc32 As Single: rc33 As Single: rc34 As Single
    rc41 As Single: rc42 As Single: rc43 As Single: rc44 As Single
End Type

Public Type ORDER
    idxMeshO        As Integer
    idxFaceO        As Long
    ZValue          As Single
End Type

Public Type MAP
    FileName        As String
    UScale          As Single
    VScale          As Single
    UOffset         As Single
    VOffset         As Single
    Angle           As Single
    Mirroring       As Byte
    Flipping        As Byte
    Filtre          As Filter
    dibTex          As DIB
    dibTexT         As DIB
End Type

Public Type MATERIAL
    Name            As String
    Ambient         As COLORRGB_INT
    Diffuse         As COLORRGB_INT
    'Specular        As COLORRGB_INT
    SpecK           As Single
    SpecN           As Single
    Opacity         As Integer
    OpcValue1       As Single           'Opacity percent
    OpcValue2       As Single           'Opacity percent
    
    TwoSided        As Boolean          'Two side flag
    MapUse          As Boolean
    Texture         As MAP
'    Opacity         As MAP
'    Bump            As MAP
End Type

Public Type MESH
    Name            As String
    
    idxVert         As Long
    idxFace         As Integer
    idxTVert        As Long

    Vertices()      As VERTEX
    Faces()         As FACE
    Screen()        As POINTAPI
    FaceV()         As ORDER
    FaceInfo        As Boolean
    MapCoorsOK      As Boolean
    
    MCoors()        As MAPCOORD
    TScreen()       As MAPCOORD

    Box(7)          As VERTEX
    BoxScreen(7)    As POINTAPI
    BoxFace(11)     As FACE
    
'    Mtrl            As MATERIAL
    
'    Rotation        As VECTOR
'    Translation     As VECTOR
'    Scales          As VECTOR
'    WorldMatrix     As MATRIX
End Type

Public Type CAMERA
    WorldPosition   As VECTOR
    LookAtPoint     As VECTOR
    VUP             As VECTOR
    PRP             As VECTOR
    Zoom            As Single
    FOV             As Single
    Yaw             As Single
    ClipFar         As Single
    ClipNear        As Single
    ViewMatrix      As MATRIX
End Type

Public Type LIGHT
    Origin          As VECTOR
    OriginT         As VECTOR
    Direction       As VECTOR
    DirectionT      As VECTOR
    
    Color           As COLORRGB_INT
    Falloff         As Single
    Hotspot         As Single
    BrightRange     As Single
    DarkRange       As Single

    Ambient         As COLORRGB_INT
    Diffusion       As Single
    Specular        As Single
    AttenEnable     As Boolean
    Enabled         As Boolean
    
    Intensity       As Integer
    IntValue1       As Single
    IntValue2       As Single
    
    ShadowEnable    As Boolean
    ShadowColor     As COLORRGB_INT
    ShadowPlane     As VECTOR
End Type

Public Type CLIPLINE
    P1              As POINTAPI
    P2              As POINTAPI
End Type

Public g_CanvasWidth      As Long
Public g_CanvasHeight     As Long
Public g_OriginX          As Long
Public g_OriginY          As Long
Public g_dibCanvas        As DIB
Public g_dibBack          As DIB
Public g_SafeFrame        As RECT
Public g_BkType           As BackType
Public g_VStyle           As VisualStyle
Public g_CDialog          As clsCommonDialog
Public g_File3DS          As clsFile3DS
Public g_Path3DS          As String
Public g_FileOBJ          As clsFileOBJ
Public g_PathOBJ          As String

Public g_idxMat           As Integer
Public g_idxMesh          As Integer

Public g_Materials()      As MATERIAL
Public g_Cameras(0)       As CAMERA
Public g_Lights(0)        As LIGHT
Public g_Meshs()          As MESH
Public g_MeshsOrder()     As ORDER
Public g_MeshsRotation    As VECTOR
Public g_MeshsTranslation As VECTOR
Public g_MeshsScales      As VECTOR

Public g_MeshsWorldMatrix As MATRIX
Public FileVer            As Long
Public MeshVer            As Long
Public BColor1          As COLORRGB_INT
Public BColor2          As COLORRGB_INT
Public BFileName        As String
Public rc               As RECT
Public PickRGB          As COLORRGB_INT
Public ValZoom          As Single

'Flags
Public g_bLoadOK          As Boolean
Public g_blnDoubleSide    As Boolean
Public g_blnRenderTime    As Boolean

Private ReLocalTra      As VECTOR
Private ReLocalSca      As VECTOR

Public Sub RefreshBack()
    
    Select Case g_BkType
        Case Blank: Clear g_dibBack
        Case Solid: Gradient0
        Case Grad1: Gradient1
        Case Grad2: Gradient2
        Case Grad3: Gradient3
        Case BkPic: LoadBackPicture
    End Select

End Sub

Public Sub BestFit()

    Dim idxMesh         As Integer
    Dim idxVect         As Long
    Dim idxBox          As Integer
    Dim MinVector       As VECTOR
    Dim MaxVector       As VECTOR
    Dim MinVectorAll    As VECTOR
    Dim MaxVectorAll    As VECTOR
    Dim CenterVectorAll As VECTOR
    Dim Dx              As Single
    Dim Dy              As Single
    Dim MaxH            As Single
    
    MaxVectorAll = g_Meshs(0).Box(0).Vectors
    MinVectorAll = g_Meshs(0).Box(0).Vectors
    For idxMesh = 0 To g_idxMesh
        With g_Meshs(idxMesh)
            MaxVector = .Box(0).Vectors
            MinVector = .Box(0).Vectors
            For idxBox = 1 To UBound(.Box)
                If MaxVector.X < .Box(idxBox).Vectors.X Then MaxVector.X = .Box(idxBox).Vectors.X
                If MaxVector.Y < .Box(idxBox).Vectors.Y Then MaxVector.Y = .Box(idxBox).Vectors.Y
                If MaxVector.Z < .Box(idxBox).Vectors.Z Then MaxVector.Z = .Box(idxBox).Vectors.Z
                If MinVector.X > .Box(idxBox).Vectors.X Then MinVector.X = .Box(idxBox).Vectors.X
                If MinVector.Y > .Box(idxBox).Vectors.Y Then MinVector.Y = .Box(idxBox).Vectors.Y
                If MinVector.Z > .Box(idxBox).Vectors.Z Then MinVector.Z = .Box(idxBox).Vectors.Z
            Next
            If MaxVectorAll.X < MaxVector.X Then MaxVectorAll.X = MaxVector.X
            If MaxVectorAll.Y < MaxVector.Y Then MaxVectorAll.Y = MaxVector.Y
            If MaxVectorAll.Z < MaxVector.Z Then MaxVectorAll.Z = MaxVector.Z
            If MinVectorAll.X > MinVector.X Then MinVectorAll.X = MinVector.X
            If MinVectorAll.Y > MinVector.Y Then MinVectorAll.Y = MinVector.Y
            If MinVectorAll.Z > MinVector.Z Then MinVectorAll.Z = MinVector.Z
        End With
    Next
    CenterVectorAll.X = (MaxVectorAll.X + MinVectorAll.X) * 0.5
    CenterVectorAll.Y = (MaxVectorAll.Y + MinVectorAll.Y) * 0.5
    CenterVectorAll.Z = (MaxVectorAll.Z + MinVectorAll.Z) * 0.5
    CenterVectorAll.W = 1
    
    Dx = MaxVectorAll.X - MinVectorAll.X
    Dy = MaxVectorAll.Y - MinVectorAll.Y
    MaxH = IIf(Dx > Dy, Dx, Dy)
    MaxH = Div(g_SafeFrame.Bottom, MaxH) * 0.05
    ReLocalTra = VectorSet(-CenterVectorAll.X, -CenterVectorAll.Y, -CenterVectorAll.Z)
    ReLocalSca = VectorSet(MaxH, MaxH, MaxH)
    ResetMeshParameters
    
End Sub

Public Sub ResetMeshParameters()

    g_MeshsRotation = VectorSet(0, 0, 0)
    g_MeshsTranslation = ReLocalTra
    g_MeshsScales = ReLocalSca

End Sub

Public Sub ResetCameraParameters()

    With g_Cameras(0)
        .WorldPosition = VectorSet(0, 0, 120)
        .LookAtPoint = VectorSet(0, 0, 0)
        .VUP = VectorSet(0, 1, 0)
        .PRP = VectorSet(0, 0, 1)
        .Zoom = 1
        .FOV = ConvertZoomtoFOV(.Zoom)
        .Yaw = 0
        .ClipFar = 0
        .ClipNear = -1000
    End With
    
End Sub

Public Sub ResetLightParameters()

    With g_Lights(0)
        .Origin = VectorSet(0, 0, 1)
        .Direction = VectorSet(0, 0, 0)
        
        .Color = ColorSet(130, 130, 130)
        .Ambient = ColorSet(0, 0, 0)
        
        .Diffusion = 2
        .Specular = 2
        .Falloff = 5
        .Hotspot = 0.5
        .DarkRange = 100
        .BrightRange = 0
        .AttenEnable = False
        .Enabled = True
        .Intensity = 50
        .IntValue1 = .Intensity * 0.01
        .IntValue2 = 1 - .IntValue1
        
        .ShadowEnable = False
        .ShadowColor = ColorSet(100, 100, 100)
        .ShadowPlane = VectorSet(0, 1, 0)
    End With

End Sub

Public Sub ResetMaterialParameters(idxMat As Integer)

    With g_Materials(idxMat)
        .Name = "Default"
        .TwoSided = False
        '.Ambient
        '.Diffusion
        '.Specular
        .Texture.FileName = ""
        .Texture.UOffset = 0
        .Texture.VOffset = 0
        .Texture.UScale = 1
        .Texture.VScale = 1
        .Texture.Angle = 0
        .Texture.Mirroring = 0
        .Texture.Flipping = 0
        Erase .Texture.dibTex.uBI.Bits
        Erase .Texture.dibTexT.uBI.Bits
    End With
        
End Sub

Public Sub CreateBox(idxMesh As Integer)

    Dim idx         As Long
    Dim MinVector   As VECTOR
    Dim MaxVector   As VECTOR

    With g_Meshs(idxMesh)
        MaxVector = .Vertices(0).Vectors
        MinVector = .Vertices(0).Vectors
        For idx = 1 To .idxVert
            If MaxVector.X < .Vertices(idx).Vectors.X Then MaxVector.X = .Vertices(idx).Vectors.X
            If MaxVector.Y < .Vertices(idx).Vectors.Y Then MaxVector.Y = .Vertices(idx).Vectors.Y
            If MaxVector.Z < .Vertices(idx).Vectors.Z Then MaxVector.Z = .Vertices(idx).Vectors.Z
            If MinVector.X > .Vertices(idx).Vectors.X Then MinVector.X = .Vertices(idx).Vectors.X
            If MinVector.Y > .Vertices(idx).Vectors.Y Then MinVector.Y = .Vertices(idx).Vectors.Y
            If MinVector.Z > .Vertices(idx).Vectors.Z Then MinVector.Z = .Vertices(idx).Vectors.Z
        Next
        
'           +Y
'           ^
'           3--------2
'          /|       /|
'         / |      / |
'        7--+-----6  |
'        |  |     |  |
'        |  0-----+--1 >+X
'        | /      | /
'        |/       |/
'        4--------5
'       /
'      +Z
        .Box(0).Vectors = MinVector
        .Box(1).Vectors = MinVector:  .Box(1).Vectors.X = MaxVector.X
        .Box(2).Vectors = MaxVector:  .Box(2).Vectors.Z = MinVector.Z
        .Box(3).Vectors = MinVector:  .Box(3).Vectors.Y = MaxVector.Y
        .Box(4).Vectors = MinVector:  .Box(4).Vectors.Z = MaxVector.Z
        .Box(5).Vectors = MaxVector:  .Box(5).Vectors.Y = MinVector.Y
        .Box(6).Vectors = MaxVector
        .Box(7).Vectors = MaxVector:  .Box(7).Vectors.X = MinVector.X
        .BoxFace(0).A = 0: .BoxFace(0).B = 1: .BoxFace(0).C = 2: .BoxFace(0).AB = True: .BoxFace(0).BC = True: .BoxFace(0).CA = False
        .BoxFace(1).A = 2: .BoxFace(1).B = 3: .BoxFace(1).C = 0: .BoxFace(1).AB = True: .BoxFace(1).BC = True: .BoxFace(1).CA = False
        .BoxFace(2).A = 4: .BoxFace(2).B = 5: .BoxFace(2).C = 6: .BoxFace(2).AB = True: .BoxFace(2).BC = True: .BoxFace(2).CA = False
        .BoxFace(3).A = 6: .BoxFace(3).B = 7: .BoxFace(3).C = 4: .BoxFace(3).AB = True: .BoxFace(3).BC = True: .BoxFace(3).CA = False
        
        .BoxFace(4).A = 1: .BoxFace(4).B = 2: .BoxFace(4).C = 6: .BoxFace(4).AB = True: .BoxFace(4).BC = True: .BoxFace(4).CA = False
        .BoxFace(5).A = 6: .BoxFace(5).B = 5: .BoxFace(5).C = 1: .BoxFace(5).AB = True: .BoxFace(5).BC = True: .BoxFace(5).CA = False
        
        .BoxFace(6).A = 0: .BoxFace(6).B = 3: .BoxFace(6).C = 7: .BoxFace(6).AB = True: .BoxFace(6).BC = True: .BoxFace(6).CA = False
        .BoxFace(7).A = 7: .BoxFace(7).B = 4: .BoxFace(7).C = 0: .BoxFace(7).AB = True: .BoxFace(7).BC = True: .BoxFace(7).CA = False
        
        .BoxFace(8).A = 2: .BoxFace(8).B = 3: .BoxFace(8).C = 7: .BoxFace(8).AB = True: .BoxFace(8).BC = True: .BoxFace(8).CA = False
        .BoxFace(9).A = 7: .BoxFace(9).B = 6: .BoxFace(9).C = 2: .BoxFace(9).AB = True: .BoxFace(9).BC = True: .BoxFace(9).CA = False
        
        .BoxFace(10).A = 4: .BoxFace(10).B = 5: .BoxFace(10).C = 1: .BoxFace(10).AB = True: .BoxFace(10).BC = True: .BoxFace(10).CA = False
        .BoxFace(11).A = 1: .BoxFace(11).B = 0: .BoxFace(11).C = 4: .BoxFace(11).AB = True: .BoxFace(11).BC = True: .BoxFace(11).CA = False
    End With

End Sub

Public Function VerifyText(txt As TextBox) As Single

    VerifyText = IIf(IsNumeric(txt.Text), CSng(txt.Text), 0)

End Function

