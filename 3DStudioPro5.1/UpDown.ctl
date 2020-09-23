VERSION 5.00
Begin VB.UserControl UpDown 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   41
   ToolboxBitmap   =   "UpDown.ctx":0000
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   0
      Top             =   1800
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   0
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   360
   End
   Begin VB.TextBox txtVal 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "255"
      Top             =   15
      Width           =   390
   End
   Begin VB.Image imgDown 
      Appearance      =   0  'Flat
      Height          =   150
      Left            =   405
      Top             =   150
      Width           =   195
   End
   Begin VB.Image imgUp 
      Appearance      =   0  'Flat
      Height          =   150
      Left            =   405
      Top             =   0
      Width           =   195
   End
End
Attribute VB_Name = "UpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim lVal    As Long
Dim lMin    As Long
Dim lMax    As Long
Dim bEnable As Boolean

Event Change()
Event MouseUp()

' Down=================================================
Private Sub imgDown_DblClick()
    
    imgDown_MouseDown 1, 0, 0, 0

End Sub

Private Sub imgDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    imgDown.Picture = LoadResPicture(104, vbResBitmap)
    If Button = 1 Then
        bEnable = True
        Timer1.Enabled = True
        If lVal <> lMin Then Value = lVal - 1
    End If
End Sub

Private Sub imgDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    imgDown.Picture = LoadResPicture(103, vbResBitmap)
    bEnable = False
    Timer1.Enabled = False
    Timer2.Enabled = False
    RaiseEvent MouseUp
    
End Sub

Private Sub Timer1_Timer()
    
    Timer2.Enabled = True

End Sub

Private Sub Timer2_Timer()
    
    bEnable = True
    If lVal <> lMin Then Value = lVal - 1

End Sub


'Up======================================================================================

Private Sub imgUp_DblClick()
    
    imgUp_MouseDown 1, 0, 0, 0

End Sub

Private Sub imgUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    imgUp.Picture = LoadResPicture(102, vbResBitmap)
    If Button = 1 Then
        bEnable = True
        Timer3.Enabled = True
        If lVal <> lMax Then Value = lVal + 1
    End If
    
End Sub

Private Sub imgUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    imgUp.Picture = LoadResPicture(101, vbResBitmap)
    bEnable = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    RaiseEvent MouseUp

End Sub

Private Sub Timer3_Timer()
    
    Timer4.Enabled = True

End Sub

Private Sub Timer4_Timer()
    
    bEnable = True
    If lVal <> lMax Then Value = lVal + 1
        
End Sub

Private Sub txtVal_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If IsNumeric(txtVal.Text) Then
            If txtVal.Text > lMin And txtVal.Text < lMax Then
                bEnable = True
                Value = txtVal.Text
            Else
                GoTo err
            End If
        Else
            GoTo err
        End If
    End If
    
Exit Sub

err:
    txtVal.Text = Value
    txtVal.SelLength = Len(txtVal.Text)

End Sub

'=======================================================================================

Private Sub UserControl_Initialize()

    imgUp.Picture = LoadResPicture(101, vbResBitmap)
    imgDown.Picture = LoadResPicture(103, vbResBitmap)
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  
  lVal = PropBag.ReadProperty("Value", 0)
  lMin = PropBag.ReadProperty("Min", 0)
  lMax = PropBag.ReadProperty("Max", 255)
  txtVal.Text = lVal

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  
  PropBag.WriteProperty "Value", lVal, 0
  PropBag.WriteProperty "Min", lMin, 0
  PropBag.WriteProperty "Max", lMax, 255

End Sub

Private Sub UserControl_Resize()
    
    UserControl.Width = 615
    UserControl.Height = 300

End Sub

Private Sub UserControl_EnterFocus()
 
    txtVal.SelLength = Len(txtVal.Text)

End Sub

Public Property Get Value() As Long
    
    bEnable = False
    Value = lVal
    
End Property

Public Property Let Value(ByVal NewValue As Long)

    lVal = NewValue
    txtVal.Text = lVal
    PropertyChanged "Value"
    If bEnable Then RaiseEvent Change
    bEnable = False
    
End Property

Public Property Get Min() As Long
    
    Min = lMin

End Property

Public Property Let Min(ByVal NewValue As Long)

    lMin = NewValue
    PropertyChanged "Min"
    
End Property

Public Property Get Max() As Long
    
    Max = lMax

End Property

Public Property Let Max(ByVal NewValue As Long)

    lMax = NewValue
    PropertyChanged "Max"
    
End Property
