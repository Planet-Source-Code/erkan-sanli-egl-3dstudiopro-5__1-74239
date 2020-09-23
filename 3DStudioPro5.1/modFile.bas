Attribute VB_Name = "modFile"
Option Explicit

Private Type ITEMID
    cb      As Long
    abID    As Integer
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMID) As Long

Public Function GetMyPicturesFolder(hwnd As Long) As String
    
    Dim retval As Long
    Dim tiID As ITEMID
    Dim Folder As String
    
    Folder = Space$(260)
    retval = SHGetSpecialFolderLocation(hwnd, &H27, tiID) 'MyPictures= &H27
    retval = SHGetPathFromIDList(ByVal tiID.cb, ByVal Folder)
    If retval Then GetMyPicturesFolder = Left$(Folder, InStr(1, Folder, Chr$(0)) - 1) & "\"

End Function

Public Function FileExist(FileName As String) As Boolean
    
    On Error Resume Next
    FileExist = CBool(FileLen(FileName))
    
End Function

Public Sub LoadBackPicture()
    
    On Error GoTo Jump
    
    With frmMain
        If FileExist(BFileName) Then
            .picLoad.Picture = LoadPicture(BFileName)
            g_dibBack.Width = g_CanvasWidth
            g_dibBack.Height = g_CanvasHeight
            CreateArrayFromPicBox2 .picLoad, g_dibBack
            g_BkType = BkPic
            Exit Sub
        End If
    End With

Jump:
    g_BkType = Blank
End Sub

Public Sub LoadTexture(idx As Integer, FileName As String)
   
    With frmMain
        On Error GoTo Jump
        If FileExist(FileName) Then
            g_Materials(idx).MapUse = True
            .picLoad.Picture = LoadPicture(FileName)
            g_Materials(idx).Texture.dibTex.Width = .picLoad.ScaleWidth
            g_Materials(idx).Texture.dibTex.Height = .picLoad.ScaleHeight
            CreateArrayFromPicBox2 .picLoad, g_Materials(idx).Texture.dibTex, True
            g_Materials(idx).Texture.dibTexT.Width = .picLoad.ScaleWidth
            g_Materials(idx).Texture.dibTexT.Height = .picLoad.ScaleHeight
            CreateArrayFromPicBox2 .picLoad, g_Materials(idx).Texture.dibTexT, True
            Exit Sub
        End If
    End With
    
Jump:
     g_Materials(idx).MapUse = False

End Sub

Public Function PicturePath(InitDir As String, Title As String) As String
    
    Set g_CDialog = New clsCommonDialog
    With g_CDialog
        .Filter = "All Picture Files|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur |" & _
                  "Bitmaps (*.bmp;*.dib)|*.bmp;*.dib|" & _
                  "GIF Images (*.gif)|*.gif|" & _
                  "JPEG Images (*.jpg)|*.jpg,*.jpeg|" & _
                  "All Files (*.*)|*.*"
        .DialogTitle = Title
        .InitDir = InitDir
        .FileName = ""
        .ShowOpen
        PicturePath = .FileName
    End With

End Function

Public Function GetFilePath(strFilePath As String) As String

    Dim FilenameEx  As String
    Dim Length      As Long
    
    FilenameEx = GetFileNameEx(strFilePath)
    Length = Len(strFilePath) - Len(FilenameEx)
    GetFilePath = Left(strFilePath, Length)
    
End Function

Public Function GetFileNameEx(strFilePath As String) As String

    Dim Segments() As String
    
    If Len(strFilePath) <> 0 Then
        Segments = Split(strFilePath, "\")
        GetFileNameEx = Segments(UBound(Segments))
    Else
        GetFileNameEx = "-"
    End If
    
End Function



