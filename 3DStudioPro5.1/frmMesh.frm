VERSION 5.00
Begin VB.Form frmMesh 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mesh"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   22
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Info"
      Height          =   2055
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   6495
      Begin VB.ListBox lstInfo 
         Height          =   1425
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   6015
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Local"
      Height          =   1815
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   2280
      Width           =   3375
      Begin VB.CommandButton cmdParams 
         Caption         =   "Go"
         Height          =   300
         Index           =   1
         Left            =   2200
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtTra 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Text            =   "traX"
         Top             =   480
         Width           =   910
      End
      Begin VB.TextBox txtTra 
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Text            =   "traY"
         Top             =   720
         Width           =   910
      End
      Begin VB.TextBox txtTra 
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   10
         Text            =   "traZ"
         Top             =   960
         Width           =   910
      End
      Begin VB.TextBox txtRot 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   9
         Text            =   "rotX"
         Top             =   480
         Width           =   910
      End
      Begin VB.TextBox txtRot 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Text            =   "rotY"
         Top             =   720
         Width           =   910
      End
      Begin VB.TextBox txtRot 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   7
         Text            =   "rotZ"
         Top             =   960
         Width           =   910
      End
      Begin VB.TextBox txtSca 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   6
         Text            =   "scaX"
         Top             =   480
         Width           =   910
      End
      Begin VB.TextBox txtSca 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   5
         Text            =   "scaY"
         Top             =   720
         Width           =   910
      End
      Begin VB.TextBox txtSca 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   4
         Text            =   "scaZ"
         Top             =   960
         Width           =   910
      End
      Begin VB.CommandButton cmdParams 
         Caption         =   "Reset"
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Scale"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   19
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Rotation"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Translation"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   17
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Z"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   1000
         Width           =   300
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Y"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   750
         Width           =   300
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "X"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mesh List"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   3015
      Begin VB.ListBox lstMeshs 
         Height          =   1230
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Dim idx As Long
    
    lstInfo.Clear
    lstInfo.AddItem "File Version: " & FileVer
    lstInfo.AddItem "Mesh Version: " & MeshVer
    
    lstMeshs.Clear
    For idx = 0 To g_idxMesh
        lstMeshs.AddItem idx & " - " & g_Meshs(idx).Name
    Next
    
    txtTra(0).Text = g_MeshsTranslation.X
    txtTra(1).Text = g_MeshsTranslation.Y
    txtTra(2).Text = g_MeshsTranslation.Z
    txtRot(0).Text = g_MeshsRotation.X
    txtRot(1).Text = g_MeshsRotation.Y
    txtRot(2).Text = g_MeshsRotation.Z
    txtSca(0).Text = g_MeshsScales.X
    txtSca(1).Text = g_MeshsScales.X '.Y
    txtSca(2).Text = g_MeshsScales.X '.Z

    
End Sub

Private Sub cmdParams_Click(Index As Integer)
    
    Select Case Index
        Case 0
            Call ResetMeshParameters
        Case 1
            g_MeshsTranslation = VectorSet(VerifyText(txtTra(0)), VerifyText(txtTra(1)), VerifyText(txtTra(2)))
            g_MeshsRotation = VectorSet(VerifyText(txtRot(0)), VerifyText(txtRot(1)), VerifyText(txtRot(2)))
            g_MeshsScales = VectorSet(VerifyText(txtSca(0)), VerifyText(txtSca(0)), VerifyText(txtSca(0)))
    End Select
    frmMain.tmrProcess.Enabled = True

End Sub

Private Sub txtRot_GotFocus(Index As Integer)
    
    frmMain.tmrProcess.Enabled = False

End Sub

Private Sub txtRot_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdParams_Click (1) ' 13=Enter

End Sub

Private Sub txtSca_GotFocus(Index As Integer)
    
    frmMain.tmrProcess.Enabled = False

End Sub

Private Sub txtSca_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdParams_Click (1) ' 13=Enter

End Sub

Private Sub txtTra_GotFocus(Index As Integer)
    
    frmMain.tmrProcess.Enabled = False

End Sub

Private Sub txtTra_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdParams_Click (1)  ' 13=Enter

End Sub

Private Sub cmdClose_Click()
    
    Unload Me

End Sub

