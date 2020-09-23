VERSION 5.00
Begin VB.Form frmCamera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Camera"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      Caption         =   "Camera"
      Height          =   3015
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdParams 
         Caption         =   "Reset"
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   2505
         Width           =   855
      End
      Begin VB.TextBox txtClpNear 
         Height          =   285
         Left            =   900
         TabIndex        =   14
         Text            =   "clpNear"
         Top             =   2085
         Width           =   2400
      End
      Begin VB.TextBox txtCamZoom 
         Height          =   285
         Left            =   900
         TabIndex        =   13
         Text            =   "zoom"
         Top             =   1485
         Width           =   2400
      End
      Begin VB.TextBox txtCamFOV 
         Height          =   285
         Left            =   900
         TabIndex        =   12
         Text            =   "fov"
         Top             =   1185
         Width           =   2400
      End
      Begin VB.CommandButton cmdParams 
         Caption         =   "Go"
         Height          =   300
         Index           =   1
         Left            =   2200
         TabIndex        =   11
         Top             =   2505
         Width           =   855
      End
      Begin VB.TextBox txtCamVUP 
         Height          =   285
         Index           =   2
         Left            =   2500
         TabIndex        =   10
         Text            =   "vupZ"
         Top             =   885
         Width           =   800
      End
      Begin VB.TextBox txtCamLAP 
         Height          =   285
         Index           =   2
         Left            =   2500
         TabIndex        =   9
         Text            =   "lapZ"
         Top             =   585
         Width           =   800
      End
      Begin VB.TextBox txtCamWP 
         Height          =   285
         Index           =   2
         Left            =   2500
         TabIndex        =   8
         Text            =   "wpZ"
         Top             =   285
         Width           =   800
      End
      Begin VB.TextBox txtCamVUP 
         Height          =   285
         Index           =   1
         Left            =   1712
         TabIndex        =   7
         Text            =   "vupY"
         Top             =   885
         Width           =   800
      End
      Begin VB.TextBox txtCamLAP 
         Height          =   285
         Index           =   1
         Left            =   1712
         TabIndex        =   6
         Text            =   "lapY"
         Top             =   585
         Width           =   800
      End
      Begin VB.TextBox txtCamWP 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Text            =   "wpY"
         Top             =   285
         Width           =   800
      End
      Begin VB.TextBox txtCamVUP 
         Height          =   285
         Index           =   0
         Left            =   900
         TabIndex        =   4
         Text            =   "vupX"
         Top             =   885
         Width           =   800
      End
      Begin VB.TextBox txtCamLAP 
         Height          =   285
         Index           =   0
         Left            =   900
         TabIndex        =   3
         Text            =   "lapX"
         Top             =   585
         Width           =   800
      End
      Begin VB.TextBox txtCamWP 
         Height          =   285
         Index           =   0
         Left            =   900
         TabIndex        =   2
         Text            =   "wpX"
         Top             =   285
         Width           =   800
      End
      Begin VB.TextBox txtClpFar 
         Height          =   285
         Left            =   900
         TabIndex        =   1
         Text            =   "clpFar"
         Top             =   1785
         Width           =   2400
      End
      Begin VB.Label lbl 
         Caption         =   "World Pos."
         Height          =   255
         Index           =   6
         Left            =   75
         TabIndex        =   22
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lbl 
         Caption         =   "Look At P."
         Height          =   255
         Index           =   7
         Left            =   75
         TabIndex        =   21
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lbl 
         Caption         =   "Zoom"
         Height          =   255
         Index           =   10
         Left            =   75
         TabIndex        =   20
         Top             =   1500
         Width           =   780
      End
      Begin VB.Label lbl 
         Caption         =   "FOV"
         Height          =   255
         Index           =   9
         Left            =   75
         TabIndex        =   19
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lbl 
         Caption         =   "VUP"
         Height          =   255
         Index           =   8
         Left            =   75
         TabIndex        =   18
         Top             =   900
         Width           =   780
      End
      Begin VB.Label lbl 
         Caption         =   "Clip Far"
         Height          =   255
         Index           =   11
         Left            =   75
         TabIndex        =   17
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label lbl 
         Caption         =   "Clip Near"
         Height          =   255
         Index           =   12
         Left            =   75
         TabIndex        =   16
         Top             =   2100
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
        
    With g_Cameras(0)
        frmCamera.txtCamWP(0).Text = .WorldPosition.X
        frmCamera.txtCamWP(1).Text = .WorldPosition.Y
        frmCamera.txtCamWP(2).Text = .WorldPosition.Z
        frmCamera.txtCamLAP(0).Text = .LookAtPoint.X
        frmCamera.txtCamLAP(1).Text = .LookAtPoint.Y
        frmCamera.txtCamLAP(2).Text = .LookAtPoint.Z
        frmCamera.txtCamVUP(0).Text = .VUP.X
        frmCamera.txtCamVUP(1).Text = .VUP.Y
        frmCamera.txtCamVUP(2).Text = .VUP.Z
        frmCamera.txtCamFOV.Text = .FOV
        frmCamera.txtCamZoom.Text = .Zoom
        frmCamera.txtClpFar.Text = .ClipFar
        frmCamera.txtClpNear.Text = .ClipNear
    End With

End Sub

Private Sub cmdParams_Click(Index As Integer)
    
    Select Case Index
        Case 0
             Call ResetCameraParameters
        Case 1
            With g_Cameras(0)
                .WorldPosition = VectorSet(VerifyText(txtCamWP(0)), _
                                           VerifyText(txtCamWP(1)), _
                                           VerifyText(txtCamWP(2)))
                .LookAtPoint = VectorSet(VerifyText(txtCamLAP(0)), _
                                         VerifyText(txtCamLAP(1)), _
                                         VerifyText(txtCamLAP(2)))
                .VUP = VectorSet(VerifyText(txtCamVUP(0)), _
                                 VerifyText(txtCamVUP(1)), _
                                 VerifyText(txtCamVUP(2)))
                .FOV = VerifyText(txtCamFOV)
                .Zoom = VerifyText(txtCamZoom)
                .ClipFar = VerifyText(txtClpFar)
                .ClipNear = VerifyText(txtClpNear)
            End With
    End Select
    frmMain.tmrProcess.Enabled = True
    
End Sub

Private Sub txtCamFOV_GotFocus()
    
    frmMain.tmrProcess.Enabled = False

End Sub

Private Sub txtCamFOV_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdParams_Click (1) ' 13=Enter

End Sub

Private Sub txtCamFOV_Validate(Cancel As Boolean)
    
    txtCamZoom.Text = ConvertFOVtoZoom(CSng(txtCamFOV.Text))

End Sub

Private Sub txtCamLAP_GotFocus(Index As Integer)
    
    frmMain.tmrProcess.Enabled = False

End Sub

Private Sub txtCamLAP_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdParams_Click (1) ' 13=Enter

End Sub

Private Sub txtCamVUP_GotFocus(Index As Integer)
    
    frmMain.tmrProcess.Enabled = False

End Sub

Private Sub txtCamVUP_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdParams_Click (1) ' 13=Enter

End Sub

Private Sub txtCamWP_GotFocus(Index As Integer)
    
    frmMain.tmrProcess.Enabled = False

End Sub

Private Sub txtCamWP_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdParams_Click (1) ' 13=Enter

End Sub

Private Sub txtCamZoom_GotFocus()
    
    frmMain.tmrProcess.Enabled = False

End Sub

Private Sub txtCamZoom_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdParams_Click (1) ' 13=Enter

End Sub

Private Sub txtCamZoom_Validate(Cancel As Boolean)
    
    txtCamFOV.Text = ConvertZoomtoFOV(CSng(txtCamZoom.Text))

End Sub

Private Sub txtClpFar_GotFocus()
    
    frmMain.tmrProcess.Enabled = False

End Sub

Private Sub txtClpFar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdParams_Click (1) ' 13=Enter

End Sub

Private Sub txtClpNear_GotFocus()
    
    frmMain.tmrProcess.Enabled = False

End Sub

Private Sub txtClpNear_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdParams_Click (1) ' 13=Enter

End Sub
