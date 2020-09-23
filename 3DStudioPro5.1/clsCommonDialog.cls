VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type OPENFILENAME
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type

Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_FILEMUSTEXIST = &H1000

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private ofn         As OPENFILENAME

Public Sub ShowOpen()
    
    Dim retval  As Long
    
    ofn.lpstrFile = Space$(254)
    retval = GetOpenFileName(ofn)
    If (retval) Then
        FileName = Trim$(ofn.lpstrFile)
    Else
        FileName = vbNullString
    End If
    
End Sub

Public Function ShowSave() As String

    Dim retval  As Long
    
    ofn.lpstrFile = Space$(254)
    retval = GetSaveFileName(ofn)
    If (retval) Then
        FileName = Trim$(ofn.lpstrFile)
    Else
        FileName = vbNullString
    End If

End Function

Public Property Get hwnd() As Long
    hwnd = ofn.hwndOwner
End Property

Public Property Let hwnd(NewhWnd As Long)
    ofn.hwndOwner = NewhWnd
End Property

Public Property Get InitDir() As String
    InitDir = ofn.lpstrInitialDir
End Property

Public Property Let InitDir(NewInitDir As String)
    ofn.lpstrInitialDir = NewInitDir
End Property

Public Property Get DialogTitle() As String
    DialogTitle = ofn.lpstrTitle
End Property

Public Property Let DialogTitle(NewTitle As String)
    ofn.lpstrTitle = NewTitle
End Property

Public Property Get DefaultExt() As String
    DefaultExt = ofn.lpstrDefExt
End Property

Public Property Let DefaultExt(NewExt As String)
    ofn.lpstrDefExt = NewExt
End Property


Public Property Get FileName() As String
    FileName = ofn.lpstrFile
End Property

Public Property Let FileName(NewFileName As String)
       
    Dim LenName As Integer
    
    LenName = 254 - Len(NewFileName)
    ofn.lpstrFile = NewFileName & Space$(LenName)
    
End Property

Public Property Get Filter() As String
    Filter = ofn.lpstrFilter
End Property

Public Property Let Filter(strFilter As String)
    If Right$(strFilter, 1) <> "|" Then strFilter = strFilter + "|"
    ofn.lpstrFilter = Replace(strFilter, "|", Chr$(0))
End Property

Private Sub Class_Initialize()

    With ofn
        .lStructSize = Len(ofn)
        .hInstance = App.hInstance
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    End With

End Sub
