VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile3DS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAIN3DS                 As Long = &H4D4D
Private Const EDIT3DS                 As Long = &H3D3D
Private Const EDIT_OBJECT             As Long = &H4000
Private Const EDIT_MATERIAL           As Long = &HAFFF

Private Const OBJ_TRIMESH             As Long = &H4100
Private Const FILE_VERSION            As Long = &H2
Private Const MESH_VERSION            As Long = &H3D3E

Private Const MAT_DIFFUSE             As Long = &HA020
Private Const MAT_NAME                As Long = &HA000
Private Const MAT_TEXMAP              As Long = &HA200
Private Const MAT_MAPNAME             As Long = &HA300
Private Const MAT_MAP_TILING          As Long = &HA351
Private Const MAT_MAP_USCALE          As Long = &HA354
Private Const MAT_MAP_VSCALE          As Long = &HA356
Private Const MAT_MAP_UOFFSET         As Long = &HA358
Private Const MAT_MAP_VOFFSET         As Long = &HA35A
Private Const MAT_MAP_ANG             As Long = &HA35C
Private Const MAT_OPACITY             As Long = &HA050

Private Const TRI_VERTEXL             As Long = &H4110
Private Const TRI_FACEL1              As Long = &H4120
Private Const TRI_MATERIAL            As Long = &H4130
Private Const TRI_MAPPINGCOORS        As Long = &H4140

Private Const NULL_CHUNK              As Long = &H0

Private Type CHUNK
    Header      As Integer '2
    Length      As Long    '4
End Type

Private hFile   As Long
Private Pointer As Long
Private ObjName As String
Private Chnk    As CHUNK
Private ErrorLoad As Boolean

Dim tex As String

Public Function Read3DS(FileName As String) As Boolean
        
    On Error Resume Next
    Dim idx As Integer
    
    If g_idxMat > -1 Then
        For idx = 0 To g_idxMat
            Destroy g_Materials(idx).Texture.dibTex
            Destroy g_Materials(idx).Texture.dibTexT
        Next
    End If
    g_idxMat = -1
    g_idxMesh = -1
        
    ResetCameraParameters
    ResetLightParameters
    
    Pointer = 1
    hFile = FreeFile
    Open FileName For Binary As #hFile
        While Not EOF(hFile)
            If Not ErrorLoad Then
                Call ReadChunk
            Else
                GoTo Jump
            End If
        Wend
    Close #hFile
    Read3DS = True
    Exit Function
    
Jump:
    Close #hFile
    Read3DS = False

End Function

Private Sub ReadChunk()

    On Error Resume Next
    
    Get #hFile, Pointer, Chnk
    
    Select Case Chnk.Header

'GO TO SUBCHUNK
        Case MAIN3DS, EDIT3DS, EDIT_MATERIAL, OBJ_TRIMESH, MAT_TEXMAP, NULL_CHUNK:
                                Chnk.Length = 6
        Case FILE_VERSION:      ReadFileVersion
        Case MESH_VERSION:      ReadMeshVersion

'MATERIALS
        Case MAT_NAME:          ReadMaterialName
        Case MAT_DIFFUSE:       ReadDiffuse
        Case MAT_MAPNAME:       ReadMapName
        Case MAT_MAP_USCALE:    ReadUScale
        Case MAT_MAP_VSCALE:    ReadVScale
        Case MAT_MAP_UOFFSET:   ReadUOffset
        Case MAT_MAP_VOFFSET:   ReadVOffset
        Case MAT_MAP_ANG:       ReadMapAng
        Case MAT_OPACITY:       ReadOpacity
        
'OBJECTS
        Case EDIT_OBJECT:       ReadObjectName
            
'READ MESH STUFF
        Case TRI_VERTEXL:       ReadVertex
        Case TRI_FACEL1:        ReadFace
        Case TRI_MATERIAL:      ReadFaceMaterial
        Case TRI_MAPPINGCOORS:  ReadMappingCoors
        
    End Select
    Pointer = Pointer + Chnk.Length
    
End Sub
        
Private Sub ReadFileVersion()

    Get #hFile, , FileVer
    
End Sub

Private Sub ReadMeshVersion()

    Get #hFile, , MeshVer

End Sub

Private Sub ReadMaterialName()
    
    Dim strRead As String
    
    strRead = ReadString
    g_idxMat = g_idxMat + 1
    ReDim Preserve g_Materials(g_idxMat)
    Call ResetMaterialParameters(g_idxMat)
    g_Materials(g_idxMat).Name = strRead
    Chnk.Length = Len(strRead) + 7
    
End Sub

Private Sub ReadDiffuse()
    
    Dim intRead As Integer
    Dim lngRead As Long
    Dim bytRead As Byte

    Get #hFile, , intRead  'Data type (skip)
    Get #hFile, , lngRead  'Color Length (skip)
    Get #hFile, , bytRead: g_Materials(g_idxMat).Diffuse.R = CInt(bytRead)
    Get #hFile, , bytRead: g_Materials(g_idxMat).Diffuse.G = CInt(bytRead)
    Get #hFile, , bytRead: g_Materials(g_idxMat).Diffuse.B = CInt(bytRead)
    'Get #hFile, , intRead 'SubChunk ID (skip)
    g_Materials(g_idxMat).SpecK = 0.35
    g_Materials(g_idxMat).SpecN = 20

End Sub

Private Sub ReadMapName()
    
    g_Materials(g_idxMat).Texture.FileName = ReadString
    
End Sub

Private Sub ReadUScale()
    
    Get #hFile, , g_Materials(g_idxMat).Texture.UScale

End Sub

Private Sub ReadVScale()
    
    Get #hFile, , g_Materials(g_idxMat).Texture.VScale

End Sub

Private Sub ReadUOffset()
    
    Get #hFile, , g_Materials(g_idxMat).Texture.UOffset

End Sub

Private Sub ReadVOffset()
    
    Dim sngRead As Single
    
    Get #hFile, , sngRead
    g_Materials(g_idxMat).Texture.VOffset = -sngRead

End Sub

Private Sub ReadMapAng()
    
    Dim sngRead As Single
    
    Get #hFile, , sngRead
    g_Materials(g_idxMat).Texture.Angle = -sngRead

End Sub

Private Sub ReadOpacity()
    
    Dim intRead As Integer
    Dim lngRead As Long
    
    Get #hFile, , intRead 'skip
    Get #hFile, , lngRead 'skip
    Get #hFile, , intRead
    g_Materials(g_idxMat).Opacity = 100 - intRead
    g_Materials(g_idxMat).OpcValue1 = g_Materials(g_idxMat).Opacity * 0.01
    g_Materials(g_idxMat).OpcValue2 = 1 - g_Materials(g_idxMat).OpcValue1
    

End Sub

Private Sub ReadObjectName()

    ObjName = ReadString
    Chnk.Length = Len(ObjName) + 7

End Sub

Private Function ReadString() As String

    Dim strChar   As String
    Dim strAll    As String
    
    strChar = StrConv(InputB(1, #hFile), vbUnicode)
    While strChar <> Chr(0)
        strAll = strAll & strChar
        strChar = StrConv(InputB(1, #hFile), vbUnicode)
    Wend
    ReadString = strAll
  
End Function

Private Sub ReadVertex()
    
    Dim idx         As Long
    Dim intRead     As Integer
    Dim sngRead     As Single
    
    g_idxMesh = g_idxMesh + 1
    ReDim Preserve g_Meshs(g_idxMesh)

    With g_Meshs(g_idxMesh)
        .Name = ObjName
        Get #hFile, , intRead
        If intRead > 0 Then
            .idxVert = intRead - 1
            ReDim .Vertices(.idxVert)
            ReDim .Screen(.idxVert)
            For idx = 0 To .idxVert
                Get #hFile, , .Vertices(idx).Vectors.X
                Get #hFile, , .Vertices(idx).Vectors.Y
                Get #hFile, , .Vertices(idx).Vectors.Z
                              .Vertices(idx).Vectors.W = 1
            Next
            Call CreateBox(g_idxMesh)
        Else
            .idxVert = -1 'error
        End If
    End With
        
End Sub

Private Sub ReadFace()
    
    Dim idx         As Long
    Dim intRead     As Integer
    Dim intTemp     As Integer
    Dim intABC()    As Integer

    With g_Meshs(g_idxMesh)
        Get #hFile, , intRead
        If intRead > 0 Then
            .idxFace = intRead - 1
            ReDim .Faces(.idxFace)
            ReDim intABC(3, .idxFace) '0=A, 1=B, 2=C, 3=AB BC CA
            Get #hFile, , intABC
            For idx = 0 To .idxFace
                .Faces(idx).A = CLng(intABC(0, idx))
                .Faces(idx).B = CLng(intABC(1, idx))
                .Faces(idx).C = CLng(intABC(2, idx))
                
                intTemp = intABC(3, idx) And 1 '0000000X - CA
                .Faces(idx).CA = CBool(intTemp)
                
                intTemp = intABC(3, idx) And 2 '000000X0 - BC
                intTemp = intTemp / 2
                .Faces(idx).BC = CBool(intTemp)
                
                intTemp = intABC(3, idx) And 4 '00000X00 - AB
                intTemp = intTemp / 4
                .Faces(idx).AB = CBool(intTemp)
                
                intTemp = intABC(3, idx) And 8 '0000X000 - Mapping
                intTemp = intTemp / 8
                .Faces(idx).Mapping = CBool(intTemp)
            Next
            .FaceInfo = True
        Else
            .idxFace = -1 'error
        End If
        Chnk.Length = CLng(intRead) * 8 + 8
    End With
    
End Sub

Private Sub ReadFaceMaterial()
    
    Dim idxMat      As Integer
    Dim idxFace     As Integer
    Dim intRead     As Integer
    Dim intReadA()  As Integer
    Dim strRead     As String
    
    strRead = ReadString
    Get #hFile, , intRead
    If intRead > 0 Then
        ReDim intReadA(intRead - 1)
        Get #hFile, , intReadA
        For idxMat = 0 To g_idxMat
            If g_Materials(idxMat).Name = strRead Then
                For idxFace = 0 To UBound(intReadA)
                    g_Meshs(g_idxMesh).Faces(intReadA(idxFace)).idxMat = idxMat
                Next
                Exit For
            End If
        Next
    End If
    
End Sub

Private Sub ReadMappingCoors()

    Dim idx         As Long
    Dim intRead     As Integer

    On Error Resume Next

    With g_Meshs(g_idxMesh)
        Get #hFile, , intRead
        .idxTVert = intRead - 1
        ReDim .MCoors(.idxTVert)
        ReDim .TScreen(.idxTVert)
        For idx = 0 To .idxTVert
            Get #hFile, , .MCoors(idx).U
            Get #hFile, , .MCoors(idx).V
        Next
        .MapCoorsOK = True
    End With

End Sub
