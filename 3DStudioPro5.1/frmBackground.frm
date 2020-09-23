VERSION 5.00
Begin VB.Form frmBackground 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Background"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   159
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwap 
      Caption         =   "Swap"
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox chkRefresh 
      Caption         =   "Auto Refresh"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   1335
   End
   Begin VB.PictureBox picColor 
      Height          =   375
      Index           =   2
      Left            =   720
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   375
      Index           =   1
      Left            =   720
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdPicture 
      Caption         =   "Picture"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdBType 
      Height          =   375
      Index           =   5
      Left            =   2520
      Picture         =   "frmBackground.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   500
   End
   Begin VB.CommandButton cmdBType 
      Height          =   375
      Index           =   4
      Left            =   2040
      Picture         =   "frmBackground.frx":15F9
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   500
   End
   Begin VB.CommandButton cmdBType 
      Height          =   375
      Index           =   3
      Left            =   1560
      Picture         =   "frmBackground.frx":1F0B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   500
   End
   Begin VB.CommandButton cmdBType 
      Height          =   375
      Index           =   2
      Left            =   1080
      Picture         =   "frmBackground.frx":2959
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   500
   End
   Begin VB.CommandButton cmdBType 
      Height          =   375
      Index           =   1
      Left            =   600
      Picture         =   "frmBackground.frx":30C9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   500
   End
   Begin VB.CommandButton cmdBType 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "frmBackground.frx":339C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Color 2"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Color 1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblBType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmBColor1    As frmColorDialog
Attribute frmBColor1.VB_VarHelpID = -1
Private WithEvents frmBColor2    As frmColorDialog
Attribute frmBColor2.VB_VarHelpID = -1

Private Sub cmdSwap_Click()

    Dim Temp As COLORRGB_INT
    
    Temp = BColor1
    BColor1 = BColor2
    BColor2 = Temp
    picColor(1).BackColor = RGB(BColor1.R, BColor1.G, BColor1.B)
    picColor(2).BackColor = RGB(BColor2.R, BColor2.G, BColor2.B)
    cmdBType_Click (g_BkType)

End Sub

Private Sub Form_Load()
        
    picColor(1).BackColor = RGB(BColor1.R, BColor1.G, BColor1.B)
    picColor(2).BackColor = RGB(BColor2.R, BColor2.G, BColor2.B)
    chkRefresh.Value = RegRead_AutoRefresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not frmBColor1 Is Nothing Then
        Unload frmBColor1
        Set frmBColor1 = Nothing
    End If
    
    If Not frmBColor2 Is Nothing Then
        Unload frmBColor2
        Set frmBColor2 = Nothing
    End If
    
    RegWrite_AutoRefresh chkRefresh.Value

End Sub

Private Sub frmBColor1_Active(R As Integer, G As Integer, B As Integer)
    
    R = BColor1.R
    G = BColor1.G
    B = BColor1.B
    
End Sub

Private Sub frmBColor2_Active(R As Integer, G As Integer, B As Integer)
    
    R = BColor2.R
    G = BColor2.G
    B = BColor2.B
    
End Sub

Private Sub frmBColor1_Change(R As Integer, G As Integer, B As Integer)
    
    BColor1 = ColorSet(R, G, B)
    picColor(1).BackColor = RGB(R, G, B)
    If (chkRefresh.Value) Then cmdBType_Click (g_BkType)
    
End Sub

Private Sub frmBColor2_Change(R As Integer, G As Integer, B As Integer)
    
    BColor2 = ColorSet(R, G, B)
    picColor(2).BackColor = RGB(R, G, B)
    If (chkRefresh.Value) Then cmdBType_Click (g_BkType)

End Sub

Private Sub cmdOK_Click()
    
    cmdBType_Click (g_BkType)
    Unload Me

End Sub

Private Sub cmdPicture_Click()

    Dim FileName As String
    
    FileName = PicturePath(App.Path & "\Background\", "Background Picture")
    If FileExist(FileName) Then
        BFileName = FileName
    Else
        BFileName = App.Path & "\Background\Default.jpg"
    End If
    Call cmdBType_Click(BkPic)

End Sub

Private Sub cmdBType_Click(Index As Integer)

    g_BkType = Index
    lblBType.Caption = Choose(Index + 1, "Blank", "Solid", "Gradient 1", "Gradient 2", "Gradient 3", "Picture")
    RefreshBack
    
End Sub

Private Sub picColor_Click(Index As Integer)
    
    Select Case Index
        Case 1
            If frmBColor1 Is Nothing Then
                Set frmBColor1 = New frmColorDialog
                frmBColor1.Caption = "Select Color 1"
            End If
            frmBColor1.Show vbModeless, frmMain
        Case 2
            If frmBColor2 Is Nothing Then
                Set frmBColor2 = New frmColorDialog
                frmBColor2.Caption = "Select Color 2"
            End If
            frmBColor2.Show vbModeless, frmMain
    End Select

End Sub

Private Sub chkRefresh_Click()
    
    cmdRefresh.Enabled = IIf(chkRefresh.Value = vbChecked, False, True)
    cmdBType_Click (g_BkType)

End Sub

Private Sub cmdRefresh_Click()
    
    cmdBType_Click (g_BkType)

End Sub

