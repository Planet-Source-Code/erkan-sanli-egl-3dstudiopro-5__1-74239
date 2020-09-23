VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileOBJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'######  UNCOMPLETED  ######

'
' Wavefront OBJ File parser
'
Dim hFile       As Long
Private Lines() As String

Public Function ReadOBJ(FileName As String) As Boolean
    
    Dim strData     As String
    Dim char        As String
    Dim idx         As Long
    Dim idxVert     As Long
    Dim idxFace     As Long
    Dim idxMap      As Long
    Dim Segments()  As String
    
    On Error Resume Next
    
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

    hFile = FreeFile
    Open FileName For Input As #hFile
        strData = Input(LOF(1) - 1, #hFile)
    Close #hFile
    
    Erase Lines
    Lines = Split(strData, vbLf)
    g_idxMesh = 0
    ReDim g_Meshs(g_idxMesh)
    g_idxMat = 0
    ReDim g_Materials(g_idxMat)
    ResetMaterialParameters g_idxMat
    g_Materials(g_idxMat).Diffuse = ColorSet(0, 200, 0)
    g_Materials(g_idxMat).SpecK = 0.35
    g_Materials(g_idxMat).SpecN = 20
    g_Materials(g_idxMat).Opacity = 100 ' - 50
    g_Materials(g_idxMat).OpcValue1 = g_Materials(g_idxMat).Opacity * 0.01
    g_Materials(g_idxMat).OpcValue2 = 1 - g_Materials(g_idxMat).OpcValue1

    With g_Meshs(g_idxMesh)
        .idxVert = -1
        .idxFace = -1
'        .idxTVert = -1
        For idx = 0 To UBound(Lines)
            If (Len(Lines(idx))) Then
                Segments = Split(Lines(idx), " ")
                char = Trim(Segments(0))
                Select Case char
                    Case "v"
                        .idxVert = .idxVert + 1
                        ReDim Preserve .Vertices(.idxVert)
                        .Vertices(.idxVert).Vectors = GetVectorValue(Lines(idx))
'                    Case "vt"
'                        ReDim Preserve .MCoors(idxMap)
'                        .MCoors(idxMap) = GetMapValue(Lines(idx))
'                        idxMap = idxMap + 1
                    Case "f":
                        .idxFace = .idxFace + 1
                        ReDim Preserve .Faces(.idxFace)
                        .Faces(.idxFace) = GetFaceValue(Lines(idx))
                        .Faces(.idxFace).AB = True
                        .Faces(.idxFace).BC = True
                        .Faces(.idxFace).CA = True
                End Select
            End If
        Next
        
        ReDim .Screen(.idxVert)
    End With
    Call CreateBox(0)
    ReadOBJ = True
    
End Function

Private Function GetVectorValue(Line As String) As VECTOR

    Dim Value As String
    Dim Segments() As String
    
    Segments = Split(Line, " ")
    
'X Value
    Value = Segments(UBound(Segments) - 2)
    GetVectorValue.X = CSng(Replace(Value, ".", ","))
'Y Value
    Value = Segments(UBound(Segments) - 1)
    GetVectorValue.Y = CSng(Replace(Value, ".", ","))
'Z Value
    Value = Segments(UBound(Segments))
    GetVectorValue.Z = CSng(Replace(Value, ".", ","))
'W Value
    GetVectorValue.W = 1

End Function

Private Function GetFaceValue(Line As String) As FACE

    Dim Value As String
    Dim Values() As String
    Dim Segments() As String
    Dim Segments2() As String
    
    
    Segments = Split(Line, " ")
    
'A Value
    Value = Segments(UBound(Segments) - 2)
    Values = Split(Value, "/")
    GetFaceValue.A = CLng(Values(0)) - 1
'B Value
    Value = Segments(UBound(Segments) - 1)
    Values = Split(Value, "/")
    GetFaceValue.B = CLng(Values(0)) - 1
'C Value
    Value = Segments(UBound(Segments))
    Values = Split(Value, "/")
    GetFaceValue.C = CLng(Values(0)) - 1

End Function

Private Function GetMapValue(Line As String) As MAPCOORD

    Dim Value As String
    Dim Segments() As String

    Segments = Split(Line, " ")

'U Value
    Value = Segments(UBound(Segments) - 1)
    GetMapValue.U = CSng(Replace(Value, ".", ","))
'V Value
    Value = Segments(UBound(Segments))
    GetMapValue.V = CSng(Replace(Value, ".", ","))

End Function

'Public Sub WriteOBJ(FileName As String)
'
'    Dim strTemp     As String
'    Dim idx         As Long
'    Dim idxMesh     As Long
'    Dim idxFace     As Long
'    Dim TotalNumFace As Long
'    Dim strX        As String
'    Dim strY        As String
'    Dim strZ        As String
'
'    On Error GoTo Jump
'
'    If FileExist(FileName) Then Kill FileName
'
'    strTemp = "#" & vbCrLf & "# Create : EGL Ege Graphic Library V1.0" & vbCrLf & "#" & vbCrLf
'
'    hFile = FreeFile
'    Open FileName For Binary As hFile
'
''Description
'        Put #hFile, , strTemp
'
''Vertices
'        For idx = 0 To UBound(Meshs(0).Vertices)
'            strX = GetVal(Meshs(0).Vertices(idx).Vectors.X)
'            strY = GetVal(Meshs(0).Vertices(idx).Vectors.Y)
'            strZ = GetVal(Meshs(0).Vertices(idx).Vectors.Z)
'
'            strTemp = "v  " & strX & " " & strY & " " & strZ & vbCrLf
'            Put #hFile, , strTemp
'        Next
'        strTemp = "# " & CStr(UBound(Meshs(0).Vertices)) & " vertices" & vbCrLf & vbCrLf & "g Object" & vbCrLf
'        Put #hFile, , strTemp
'
''Faces
'        'For idxMesh = 1 To UBound(Meshs(0).Faces)
'            For idxFace = 0 To UBound(Meshs(0).Faces) 'Meshs(0).Meshs(idxMesh).NumFaces
'                With Meshs(0).Faces(idxFace)
'                    strTemp = "f " & CStr(.A) & " " & CStr(.B) & " " & CStr(.C) & vbCrLf
'                    Put #hFile, , strTemp
'                    TotalNumFace = TotalNumFace + 1
'                End With
'            Next
'        'Next
'        strTemp = "# " & CStr(TotalNumFace) & " faces" & vbCrLf & vbCrLf & "g " & vbCrLf
'        Put #hFile, , strTemp
'
'    Close #hFile
'    Exit Sub
'Jump:
'    Close #hFile
'End Sub

'Private Function GetVal(Val As Single) As String
'
'    Val = Round(Val, 6)
'    GetVal = CStr(Val)
'    GetVal = Replace(GetVal, ",", ".")
'
'End Function

