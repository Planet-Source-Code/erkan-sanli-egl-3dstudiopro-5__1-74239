Attribute VB_Name = "modEvents"
Option Explicit

'Mouse___________________________________________________________________________________
'Private Const SM_MOUSEWHEELPRESENT  As Long = 75
'Private Const WM_MOUSEWHEEL         As Integer = &H20A
'
'Private Type MSG
'    hwnd        As Long
'    message     As Long
'    wParam      As Long
'    lParam      As Long
'    time        As Long
'    pt          As POINTAPI
'End Type
'
'Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
'Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
'Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
'
'Public m_blnWheelPresent    As Boolean

'Keyboard________________________________________________________________________________
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Dim sngStep As Single
'________________________________________________________________________________________

Private Function State(Key As KeyCodeConstants) As Boolean

    State = (GetKeyState(Key) And &H8000)

End Function

Public Sub RefreshEvents(hwnd As Long)
    
'    Dim lResult         As Long
'    Dim tMouseCords     As POINTAPI
'    Dim lCurrentHwnd    As Long
'    Dim tMSG            As MSG
'    Dim iDir            As Single

'MouseWheel
'    If m_blnWheelPresent Then
'        lResult = GetCursorPos(tMouseCords)
'        lCurrentHwnd = WindowFromPoint(tMouseCords.X, tMouseCords.Y)
'        If lCurrentHwnd = hwnd Then
'            lResult = GetMessage(tMSG, frmMain.hwnd, 0, 0)
'            lResult = TranslateMessage(tMSG)
'            lResult = DispatchMessage(tMSG)
'            If tMSG.message = WM_MOUSEWHEEL Then
'                iDir = Sgn(tMSG.wParam \ &H7FFF)
'                Call ActSca(iDir)
'            End If
'        End If
'    End If
    
'End If
'
'Public Sub UpdateParameters()

    sngStep = 5
    'With Meshs(0)
'R: Mesh Rotate
        If State(vbKeyR) Then
            If State(vbKeyUp) Then g_MeshsRotation.X = g_MeshsRotation.X + sngStep
            If State(vbKeyDown) Then g_MeshsRotation.X = g_MeshsRotation.X - sngStep
            If State(vbKeyLeft) Then g_MeshsRotation.Y = g_MeshsRotation.Y - sngStep
            If State(vbKeyRight) Then g_MeshsRotation.Y = g_MeshsRotation.Y + sngStep
            If State(vbKeyPageUp) Then g_MeshsRotation.Z = g_MeshsRotation.Z - sngStep
            If State(vbKeyPageDown) Then g_MeshsRotation.Z = g_MeshsRotation.Z + sngStep
            g_MeshsRotation.X = g_MeshsRotation.X Mod 360
            g_MeshsRotation.Y = g_MeshsRotation.Y Mod 360
            g_MeshsRotation.Z = g_MeshsRotation.Z Mod 360
'S: Mesh Scale
        ElseIf State(vbKeyS) Then
            If State(vbKeyUp) Then g_MeshsScales = VectorPlus(g_MeshsScales, 0.1)
            If State(vbKeyDown) Then g_MeshsScales = VectorMinus(g_MeshsScales, 0.1)
            If State(vbKeyLeft) Then g_MeshsScales = VectorPlus(g_MeshsScales, 1)
            If State(vbKeyRight) Then g_MeshsScales = VectorMinus(g_MeshsScales, 1)
'            If State(vbKeyPageUp) Then g_MeshsScales = VectorPlus(g_MeshsScales, 2.5)
'            If State(vbKeyPageDown) Then g_MeshsScales = VectorMinus(g_MeshsScales, 2.5)
            If g_MeshsScales.X < 0.1 Then g_MeshsScales = VectorSet(0.1, 0.1, 0.1)
'T: Mesh Translate
        ElseIf State(vbKeyT) Then
            If State(vbKeyUp) Then g_MeshsTranslation.Y = g_MeshsTranslation.Y + g_MeshsScales.X
            If State(vbKeyDown) Then g_MeshsTranslation.Y = g_MeshsTranslation.Y - g_MeshsScales.X
            If State(vbKeyLeft) Then g_MeshsTranslation.X = g_MeshsTranslation.X - g_MeshsScales.X
            If State(vbKeyRight) Then g_MeshsTranslation.X = g_MeshsTranslation.X + g_MeshsScales.X
            If State(vbKeyPageUp) Then g_MeshsTranslation.Z = g_MeshsTranslation.Z + g_MeshsScales.X
            If State(vbKeyPageDown) Then g_MeshsTranslation.Z = g_MeshsTranslation.Z - g_MeshsScales.X
        End If
    'End With
    
    With g_Cameras(0)
'Q: Camera Look At Point
        If State(vbKeyQ) Then
            If State(vbKeyUp) Then .LookAtPoint.Y = .LookAtPoint.Y + sngStep
            If State(vbKeyDown) Then .LookAtPoint.Y = .LookAtPoint.Y - sngStep
            If State(vbKeyLeft) Then .LookAtPoint.X = .LookAtPoint.X - sngStep
            If State(vbKeyRight) Then .LookAtPoint.X = .LookAtPoint.X + sngStep
            If State(vbKeyPageUp) Then .LookAtPoint.Z = .LookAtPoint.Z + sngStep
            If State(vbKeyPageDown) Then .LookAtPoint.Z = .LookAtPoint.Z - sngStep
            Call UpdateVUP
'W: Camera World Position
        ElseIf State(vbKeyW) Then
            If State(vbKeyUp) Then .WorldPosition.Y = .WorldPosition.Y + sngStep
            If State(vbKeyDown) Then .WorldPosition.Y = .WorldPosition.Y - sngStep
            If State(vbKeyLeft) Then .WorldPosition.X = .WorldPosition.X - sngStep
            If State(vbKeyRight) Then .WorldPosition.X = .WorldPosition.X + sngStep
            If State(vbKeyPageUp) Then .WorldPosition.Z = .WorldPosition.Z + sngStep
            If State(vbKeyPageDown) Then .WorldPosition.Z = .WorldPosition.Z - sngStep
            Call UpdateVUP
'Y: Camera Yaw Angle
        ElseIf State(vbKeyY) Then
            If State(vbKeyUp) Then .Yaw = .Yaw + 1
            If State(vbKeyDown) Then .Yaw = .Yaw - 1
            .Yaw = .Yaw Mod 360
            Call UpdateVUP
'Z: Camera Zoom
        ElseIf State(vbKeyZ) Then
            If State(vbKeyPageUp) And .Zoom > 0.05 Then .Zoom = .Zoom - 0.05
            If State(vbKeyPageDown) Then .Zoom = .Zoom + 0.05
            .FOV = ConvertZoomtoFOV(.Zoom)
            Call UpdateCanvas
        End If
    End With

    With g_Lights(0)
'L: Light World Position
        If State(vbKeyL) Then
            If State(vbKeyUp) Then .Origin.Y = .Origin.Y + sngStep '* 10
            If State(vbKeyDown) Then .Origin.Y = .Origin.Y - sngStep '* 10
            If State(vbKeyLeft) Then .Origin.X = .Origin.X - sngStep '* 10
            If State(vbKeyRight) Then .Origin.X = .Origin.X + sngStep '* 10
            If State(vbKeyPageUp) Then .Origin.Z = .Origin.Z + sngStep '* 10
            If State(vbKeyPageDown) Then .Origin.Z = .Origin.Z - sngStep ' * 10
'K: Light Direction
        ElseIf State(vbKeyK) Then
            If State(vbKeyUp) Then .Direction.Y = .Direction.Y + sngStep
            If State(vbKeyDown) Then .Direction.Y = .Direction.Y - sngStep
            If State(vbKeyLeft) Then .Direction.X = .Direction.X - sngStep
            If State(vbKeyRight) Then .Direction.X = .Direction.X + sngStep
            If State(vbKeyPageUp) Then .Direction.Z = .Direction.Z + sngStep
            If State(vbKeyPageDown) Then .Direction.Z = .Direction.Z - sngStep
        End If
    End With
'C,X,Esc
    If State(vbKeyC) Then Call ResetCameraParameters
    If State(vbKeyX) Then Call ResetMeshParameters
    If State(vbKeyV) Then Call ResetLightParameters
    If State(vbKeyEscape) Then Unload frmMain
'Visual Styles
    If State(vbKeyShift) Then
        If State(vbKey1) Then frmMain.mnuVisualStyle_Click (Dot)
        If State(vbKey2) Then frmMain.mnuVisualStyle_Click (Box)
        If State(vbKey3) Then frmMain.mnuVisualStyle_Click (Wireframe)
        If State(vbKey4) Then frmMain.mnuVisualStyle_Click (WireframeMap)
        If State(vbKey5) Then frmMain.mnuVisualStyle_Click (Flat)
        If State(vbKey6) Then frmMain.mnuVisualStyle_Click (Gouraud)
        If State(vbKey7) Then frmMain.mnuVisualStyle_Click (FlatMap)
    End If

End Sub

Public Sub UpdateCanvas()

    On Error Resume Next
    
    With frmMain
        .ScaleWidth = 2 / g_Cameras(0).Zoom
        .ScaleHeight = .ScaleWidth * .Height / .Width
        .ScaleLeft = -.ScaleWidth * 0.5
        .ScaleTop = -.ScaleHeight * 0.5
    End With

End Sub

Private Sub UpdateVUP()
    
    With g_Cameras(0)
        If .LookAtPoint.Z = .WorldPosition.Z Then .WorldPosition.Z = .WorldPosition.Z + 1
        If .LookAtPoint.X - .WorldPosition.X = 0 And .LookAtPoint.Y - .WorldPosition.Y = 0 Then
            .VUP.X = 0
            .VUP.Y = 1
        ElseIf .LookAtPoint.Z - .WorldPosition.Z < 0 Then
            .VUP.X = .LookAtPoint.X - .WorldPosition.X
            .VUP.Y = .LookAtPoint.Y - .WorldPosition.Y
        Else
            .VUP.X = .WorldPosition.X - .LookAtPoint.X '.VUP.X = -(.LookAtPoint.X - .WorldPosition.X)
            .VUP.Y = .WorldPosition.Y - .LookAtPoint.Y '.VUP.Y = -(.LookAtPoint.Y - .WorldPosition.Y)
        End If
        .VUP = MatrixMultiplyVector(MatrixRotationZ(ConvertDeg2Rad(.Yaw)), .VUP)
    End With

End Sub

'Public Sub MouseInit()
'
'    m_blnWheelPresent = GetSystemMetrics(SM_MOUSEWHEELPRESENT)
'
'End Sub

Public Sub ActSca(Step As Single)
            
    With g_Meshs(0)
        If Step < 0 Then g_MeshsScales = VectorPlus(g_MeshsScales, 0.25)
        If Step > 0 Then g_MeshsScales = VectorMinus(g_MeshsScales, 0.25)
        If g_MeshsScales.X < 0.1 Then g_MeshsScales = VectorSet(0.1, 0.1, 0.1)
    End With

End Sub

