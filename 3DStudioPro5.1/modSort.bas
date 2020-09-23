Attribute VB_Name = "modSort"
Option Explicit

Public Function SortFaces(Optional Visible As Boolean = True) As Integer
    
    Dim idxMesh     As Integer
    Dim idxFace     As Long
    Dim idxOrder    As Long
    
    idxOrder = -1
    Erase g_MeshsOrder
    
    If Visible Then
        For idxMesh = 0 To g_idxMesh
            For idxFace = 0 To g_Meshs(idxMesh).idxFace
                If VisibleFace(idxMesh, idxFace) Then AddFace idxOrder, idxMesh, idxFace
            Next
        Next
    Else
        For idxMesh = 0 To g_idxMesh
            For idxFace = 0 To g_Meshs(idxMesh).idxFace
                If Not VisibleFace(idxMesh, idxFace) Then AddFace idxOrder, idxMesh, idxFace
            Next
        Next
    End If
    If idxOrder > -1 Then QSort 0, idxOrder
    SortFaces = idxOrder

End Function

Public Function VisibleFace(idxM As Integer, idxF As Long) As Boolean
    
    
    With g_Meshs(idxM)
        VisibleFace = ((.Vertices(.Faces(idxF).A).VectorsT.X - .Vertices(.Faces(idxF).B).VectorsT.X) * _
                       (.Vertices(.Faces(idxF).C).VectorsT.Y - .Vertices(.Faces(idxF).B).VectorsT.Y) - _
                       (.Vertices(.Faces(idxF).A).VectorsT.Y - .Vertices(.Faces(idxF).B).VectorsT.Y) * _
                       (.Vertices(.Faces(idxF).C).VectorsT.X - .Vertices(.Faces(idxF).B).VectorsT.X) > 0)
    End With

End Function

Private Sub AddFace(idxOrder As Long, idxM As Integer, idxF As Long)

    idxOrder = idxOrder + 1
    ReDim Preserve g_MeshsOrder(idxOrder)
    
    With g_Meshs(idxM)
        g_MeshsOrder(idxOrder).ZValue = (.Vertices(.Faces(idxF).A).VectorsT.Z + _
                                         .Vertices(.Faces(idxF).B).VectorsT.Z + _
                                         .Vertices(.Faces(idxF).C).VectorsT.Z)
        g_MeshsOrder(idxOrder).idxFaceO = idxF
        g_MeshsOrder(idxOrder).idxMeshO = idxM
    End With

End Sub

Private Sub QSort(ByVal First As Long, ByVal Last As Long)

    Dim FirstIdx    As Long
    Dim MidIdx      As Long
    Dim LastIdx     As Long
    Dim MidVal      As Single
    Dim TempOrder   As ORDER
    
    If (First < Last) Then
            MidIdx = (First + Last) * 0.5
            MidVal = g_MeshsOrder(MidIdx).ZValue
            FirstIdx = First
            LastIdx = Last
            Do
                Do While g_MeshsOrder(FirstIdx).ZValue < MidVal
                    FirstIdx = FirstIdx + 1
                Loop
                Do While g_MeshsOrder(LastIdx).ZValue > MidVal
                    LastIdx = LastIdx - 1
                Loop
                If (FirstIdx <= LastIdx) Then
                    TempOrder = g_MeshsOrder(LastIdx)
                    g_MeshsOrder(LastIdx) = g_MeshsOrder(FirstIdx)
                    g_MeshsOrder(FirstIdx) = TempOrder
                    FirstIdx = FirstIdx + 1
                    LastIdx = LastIdx - 1
                End If
            Loop Until FirstIdx > LastIdx

            If (LastIdx <= MidIdx) Then
                QSort First, LastIdx
                QSort FirstIdx, Last
            Else
                QSort FirstIdx, Last
                QSort First, LastIdx
            End If
    End If

End Sub
