VERSION 5.00
Begin VB.Form frmControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control"
   ClientHeight    =   8775
   ClientLeft      =   2760
   ClientTop       =   3675
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtKeyboard 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   4575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   7920
      Width           =   1215
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim strData     As String
    Dim strFileName As String
    Dim hFile       As Long
    
    txtKeyboard.Move 0, 0, Me.Width
    strFileName = App.Path & "\Control.txt"
    
    hFile = FreeFile
    Open strFileName For Input As hFile
        strData = Input(LOF(1) - 1, hFile)
    Close hFile
    
    txtKeyboard.Text = strData

End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub txtKeyboard_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

