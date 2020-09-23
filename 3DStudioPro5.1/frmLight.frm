VERSION 5.00
Begin VB.Form frmLight 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Light"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Shadow"
      Height          =   1095
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   2535
      Begin VB.CheckBox chkShadow 
         Caption         =   "Enable"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   24
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Color"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Color"
      Height          =   1575
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   2535
      Begin VB.PictureBox picDiffuse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox picAmbient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin EGL_3DStudioPro5.UpDown udIntensity 
         Height          =   300
         Left            =   1320
         TabIndex        =   17
         Top             =   1080
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
      End
      Begin VB.Label Label1 
         Caption         =   "Ambient"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Diffuse"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Intensity"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   20
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.CheckBox chkLight 
      Caption         =   "Enable"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame fra 
      Caption         =   "Local"
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   2535
      Begin VB.CommandButton cmdParams 
         Caption         =   "Go"
         Height          =   300
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtCen 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Text            =   "cenX"
         Top             =   480
         Width           =   910
      End
      Begin VB.TextBox txtCen 
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Text            =   "cenY"
         Top             =   720
         Width           =   910
      End
      Begin VB.TextBox txtCen 
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   5
         Text            =   "cenZ"
         Top             =   960
         Width           =   910
      End
      Begin VB.TextBox txtTar 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   4
         Text            =   "tarX"
         Top             =   480
         Width           =   910
      End
      Begin VB.TextBox txtTar 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Text            =   "tarY"
         Top             =   720
         Width           =   910
      End
      Begin VB.TextBox txtTar 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   2
         Text            =   "tarZ"
         Top             =   960
         Width           =   910
      End
      Begin VB.CommandButton cmdParams 
         Caption         =   "Reset"
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Target"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Center"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Z"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1000
         Width           =   300
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Y"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   750
         Width           =   300
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "X"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmColorLight        As frmColorDialog
Attribute frmColorLight.VB_VarHelpID = -1
Private WithEvents frmColorAmbient      As frmColorDialog
Attribute frmColorAmbient.VB_VarHelpID = -1
Private WithEvents frmColorShadow       As frmColorDialog
Attribute frmColorShadow.VB_VarHelpID = -1

Private Sub chkLight_Click()

    g_Lights(0).Enabled = chkLight.Value

End Sub

Private Sub chkShadow_Click()
    
    g_Lights(0).ShadowEnable = chkShadow.Value

End Sub

Private Sub cmdParams_Click(Index As Integer)
    
    Select Case Index
        Case 0
            Call ResetLightParameters
        Case 1
            g_Lights(0).Origin = VectorSet(VerifyText(txtCen(0)), VerifyText(txtCen(1)), VerifyText(txtCen(2)))
            g_Lights(0).Direction = VectorSet(VerifyText(txtTar(0)), VerifyText(txtTar(1)), VerifyText(txtTar(2)))
    End Select

End Sub

Private Sub Form_Load()
    
    picDiffuse.BackColor = ColorRGBToLong(g_Lights(0).Color)
    picAmbient.BackColor = ColorRGBToLong(g_Lights(0).Ambient)
    picShadow.BackColor = ColorRGBToLong(g_Lights(0).ShadowColor)
    chkLight.Value = IIf(g_Lights(0).Enabled, vbChecked, vbUnchecked)
    chkShadow.Value = IIf(g_Lights(0).ShadowEnable, vbChecked, vbUnchecked)
    udIntensity.Value = g_Lights(0).Intensity
    txtCen(0).Text = g_Lights(0).Origin.X
    txtCen(1).Text = g_Lights(0).Origin.Y
    txtCen(2).Text = g_Lights(0).Origin.Z
    txtTar(0).Text = g_Lights(0).Direction.X
    txtTar(1).Text = g_Lights(0).Direction.Y
    txtTar(2).Text = g_Lights(0).Direction.Z

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not frmColorLight Is Nothing Then Set frmColorLight = Nothing
    If Not frmColorAmbient Is Nothing Then Set frmColorAmbient = Nothing
    If Not frmColorShadow Is Nothing Then Set frmColorShadow = Nothing

End Sub

Private Sub frmColorLight_Active(R As Integer, G As Integer, B As Integer)

    R = g_Lights(0).Color.R
    G = g_Lights(0).Color.G
    B = g_Lights(0).Color.B

End Sub

Private Sub frmColorLight_Change(R As Integer, G As Integer, B As Integer)

    g_Lights(0).Color = ColorSet(R, G, B)
    picDiffuse.BackColor = RGB(R, G, B)

End Sub

Private Sub frmColorAmbient_Active(R As Integer, G As Integer, B As Integer)

    R = g_Lights(0).Ambient.R
    G = g_Lights(0).Ambient.G
    B = g_Lights(0).Ambient.B

End Sub

Private Sub frmColorAmbient_Change(R As Integer, G As Integer, B As Integer)

    g_Lights(0).Ambient = ColorSet(R, G, B)
    picAmbient.BackColor = RGB(R, G, B)

End Sub

Private Sub frmColorShadow_Change(R As Integer, G As Integer, B As Integer)

    g_Lights(0).ShadowColor = ColorSet(R, G, B)
    picShadow.BackColor = RGB(R, G, B)

End Sub

Private Sub frmColorShadow_Active(R As Integer, G As Integer, B As Integer)

    R = g_Lights(0).ShadowColor.R
    G = g_Lights(0).ShadowColor.G
    B = g_Lights(0).ShadowColor.B

End Sub

Private Sub picAmbient_Click()
    
    If g_bLoadOK = False Then Exit Sub
    
    If frmColorAmbient Is Nothing Then
        Set frmColorAmbient = New frmColorDialog
        frmColorAmbient.Caption = "Select Ambient Color"
    End If
    frmColorAmbient.Show vbModal, frmLight

End Sub

Private Sub picDiffuse_Click()
    
    If g_bLoadOK = False Then Exit Sub
    
    If frmColorLight Is Nothing Then
        Set frmColorLight = New frmColorDialog
        frmColorLight.Caption = "Select Light Color"
    End If
    frmColorLight.Show vbModal, frmLight

End Sub

Private Sub picShadow_Click()
    
    If g_bLoadOK = False Then Exit Sub
    
    If frmColorShadow Is Nothing Then
        Set frmColorShadow = New frmColorDialog
        frmColorShadow.Caption = "Select Shadow Color"
    End If
    frmColorShadow.Show vbModal, frmLight

End Sub

Private Sub cmdClose_Click()
    
    Unload Me

End Sub

Private Sub txtSca_Change(Index As Integer)

End Sub

Private Sub udIntensity_Change()
    
    With g_Lights(0)
        .Intensity = udIntensity.Value
        .IntValue1 = .Intensity * 0.01
        .IntValue2 = 1 - .IntValue1
    End With

End Sub
