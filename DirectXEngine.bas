Attribute VB_Name = "DirectXEngine"
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Const PlayerHeightStand = 7
Public Const PlayerHeightCrouch = 3
Public Const PlayerSpeedStand = 0.4
Public Const PlayerSpeedCrouch = 0.2
Public Const ZoomIn = 10
Public Const ZoomOut = 3

Public Const SWidth = 320  '640  '800  '...
Public Const SHeight = 240 '480  '600  '...

Public PlayerHeight As Single
Public PlayerSpeed As Single

Public DX As DirectX8
Public D3DX As New D3DX8
Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8
Public DS As DirectSound8

Public DI As DirectInput8
Public DIDevice As DirectInputDevice8
Public DIDevice2 As DirectInputDevice8
Public DIState As DIKEYBOARDSTATE

Public mMesh As New cMesh
Public mWeapon As New cWeapon
Public mSky As New ESkyBox
Public mEmitter As New cEmitter
Public Zoomed As Integer

Public EyePosition As D3DVECTOR
Public EyeLookAt As D3DVECTOR
Public EyeLookDir As D3DVECTOR
Public Yaw As Single
Public Pitch As Single
Public matView As D3DMATRIX

Public Light As D3DLIGHT8

Public MainFont As D3DXFont
Public MainFontDesc As IFont
Public TextRect As RECT
Public Fnt As New StdFont

Public Type CHVERTEX
    Pos As D3DVECTOR
    Col As Long
End Type

Public CHList(7) As CHVERTEX
Public Const FVFCH = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

Public Black(0 To 5) As FVERTEX

Public Type FVERTEX
    Pos As D3DVECTOR
    Col As Long
    tu As Single
    TV As Single
End Type

Public Const FVFF = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Public Breather As Single

Private DevData(1 To 10) As DIDEVICEOBJECTDATA
Private NumEvents As Long

Private LastUpdate As Long

Public Function Init(HWND As Long) As Boolean

    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    If D3D Is Nothing Then Exit Function
    
    Dim DispMode As D3DDISPLAYMODE
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

    Set DS = DX.DirectSoundCreate(vbNullString)
    DS.SetCooperativeLevel frmMain.HWND, DSSCL_PRIORITY

    Dim D3DPP As D3DPRESENT_PARAMETERS
    
    With D3DPP
        .Windowed = 0
        .BackBufferHeight = SHeight
        .BackBufferWidth = SWidth
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        .BackBufferFormat = DispMode.Format
        .hDeviceWindow = HWND
        .BackBufferCount = 1
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
    End With
    
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, HWND, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DPP)
    If D3DDevice Is Nothing Then Exit Function
    
    D3DDevice.SetRenderState D3DRS_CULLMODE, 0
    D3DDevice.SetRenderState D3DRS_ZENABLE, 1

    D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
    D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE
    
    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    D3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_LINEAR
    
    D3DDevice.SetRenderState D3DRS_CULLMODE, 1
    
    D3DDevice.SetRenderState D3DRS_FILLMODE, 3
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, 13
    D3DDevice.SetRenderState D3DRS_DESTBLEND, 13
    
    Fnt.Name = "Verdana"
    Fnt.Size = 10
    Fnt.Bold = True
    
    Set MainFontDesc = Fnt
    Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
    
    CHList(0).Pos = MakeVector(0, 0.05, 1)
    CHList(1).Pos = MakeVector(0, 0.01, 1)
    CHList(2).Pos = MakeVector(-0.04, 0, 1)
    CHList(3).Pos = MakeVector(-0.01, 0, 1)
    CHList(4).Pos = MakeVector(0, -0.05, 1)
    CHList(5).Pos = MakeVector(0, -0.01, 1)
    CHList(6).Pos = MakeVector(0.04, 0, 1)
    CHList(7).Pos = MakeVector(0.01, 0, 1)
    
    For k = 0 To 7
        CHList(k).Col = &H99FFFFFF
    Next k
    
    Black(0).Pos = MakeVector(-2, -2, 1)
    Black(1).Pos = MakeVector(2, -2, 1)
    Black(2).Pos = MakeVector(-2, 2, 1)
    Black(3).Pos = MakeVector(2, -2, 1)
    Black(4).Pos = MakeVector(2, 2, 1)
    Black(5).Pos = MakeVector(-2, 2, 1)
    
    For k = 0 To 5
        Black(k).Col = &HFFFFFFFF
    Next k
    
    InitDI HWND
    SetupMatrices
    InitLight
    
    mMesh.Init App.Path + "\Test.x"
    mWeapon.Init App.Path + "\Wep.x", App.Path + "\Rel.x"
    mSky.Init App.Path + "\sky.jpg"
    
    GameInit
    
    Init = True
End Function

Public Function InitLight()
    D3DDevice.SetRenderState D3DRS_LIGHTING, 1
    D3DDevice.SetRenderState D3DRS_AMBIENT, RGB(50, 50, 50)
    
    Light.Type = D3DLIGHT_DIRECTIONAL
    Light.diffuse.R = 0.6
    Light.diffuse.G = 0.6
    Light.diffuse.B = 0.6
    Light.Direction = MakeVector(1, -1, 0.7)
End Function

Public Function InitDI(HWND As Long)
    Set DI = DX.DirectInputCreate()
    Set DIDevice = DI.CreateDevice("GUID_SysKeyboard")
    
    DIDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DIDevice.SetCooperativeLevel HWND, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    
    DIDevice.Acquire
    
    Set DIDevice2 = DI.CreateDevice("guid_SysMouse")
    
    DIDevice2.SetCommonDataFormat DIFORMAT_MOUSE
    DIDevice2.SetCooperativeLevel HWND, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
    
    Dim DIProp As DIPROPLONG
    DIProp.lHow = DIPH_DEVICE
    DIProp.lObj = 0
    DIProp.lData = 10
    
    DIDevice2.SetProperty "DIPROP_BUFFERSIZE", DIProp
    
    DIDevice2.Acquire
End Function

Public Function SetupMatrices()
    Dim matView As D3DMATRIX
    
    EyePosition = MakeVector(0, 5, 0)
    EyeLookAt = MakeVector(2, 5, 0)
    Yaw = PI / 2
    PlayerHeight = PlayerHeightStand
    D3DXVec3Subtract EyeLookDir, EyeLookAt, EyePosition
    D3DXVec3Normalize EyeLookDir, EyeLookDir
    
    D3DXMatrixLookAtLH matView, EyePosition, EyeLookAt, MakeVector(0, 1, 0)
    
    D3DDevice.SetTransform D3DTS_VIEW, matView

    Dim matProj As D3DMATRIX
    
    D3DXMatrixPerspectiveFovLH matProj, 3.14152 / 3, 1, 1, 10000
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj

End Function

Public Function RenderAll()
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, RGB(125, 0, 0), 1#, 0

    If GODI = 0 Then
        Breather = Sin(GetTickCount / 400) / 15 + 0.2
    End If
    
    D3DXVec3Add EyeLookAt, MakeVector(EyePosition.X, EyePosition.Y + PlayerHeight + Breather, EyePosition.Z), EyeLookDir
    D3DXMatrixLookAtLH matView, MakeVector(EyePosition.X, EyePosition.Y + PlayerHeight + Breather, EyePosition.Z), EyeLookAt, MakeVector(0, 1, 0)

    D3DDevice.SetTransform D3DTS_VIEW, matView

    If GODI = 0 Then
        mWeapon.Adjust Yaw, Pitch, MakeVector(EyePosition.X, EyePosition.Y + PlayerHeight, EyePosition.Z)
        mMesh.Update
        mSky.Update
    End If
    
    D3DDevice.BeginScene
    
        D3DDevice.SetLight 0, Light
        D3DDevice.LightEnable 0, 1
        
        mSky.Render
        RenderGame
        mMesh.Render
        mWeapon.Render
        
        If GODI = 1 Then
            DrawBlack
            D3DX.DrawText MainFont, &H99CCCCFF, "Game Over!", TextRect, DT_CENTER Or DT_VCENTER
            D3DX.DrawText MainFont, &H99CCCCFF, "SturmNacht", TextRect, DT_BOTTOM Or DT_RIGHT
        ElseIf GODI = 2 Then
            DrawBlack
            D3DX.DrawText MainFont, &H99CCCCFF, "You Have Won!", TextRect, DT_CENTER Or DT_VCENTER
            D3DX.DrawText MainFont, &H99CCCCFF, "SturmNacht", TextRect, DT_BOTTOM Or DT_RIGHT
        Else
            DrawCrosshair
            
            TextRect.Top = 0
            TextRect.Left = 0
            TextRect.bottom = SHeight
            TextRect.Right = SWidth
            
            D3DX.DrawText MainFont, &H99CCCCFF, "Life: " + CStr(PlayerHealth), TextRect, DT_TOP Or DT_LEFT
            D3DX.DrawText MainFont, &H99CCCCFF, "Enemylife: " + CStr(mEnemy.tLife), TextRect, DT_TOP Or DT_RIGHT
            D3DX.DrawText MainFont, &H99CCCCFF, "Bullets: " + CStr(mWeapon.Bullets), TextRect, DT_BOTTOM Or DT_LEFT
            D3DX.DrawText MainFont, &H99CCCCFF, "Clips: " + CStr(mWeapon.Clips), TextRect, DT_BOTTOM Or DT_RIGHT
            LastUpdate = GetTickCount
        End If
        
    D3DDevice.EndScene
    
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Function

Public Function DrawBlack()
    Dim Leng As Single
    Leng = 255 - (GetTickCount - LastUpdate) / 10
    If Leng <= 0 Then Leng = 0
    
    For k = 0 To 5
        Black(k).Col = ARGB2LONG(CLng(Leng), 0, 0, 0)
    Next k
    
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    D3DXMatrixIdentity matView
    D3DDevice.SetTransform D3DTS_WORLD, matView
    D3DDevice.SetTransform D3DTS_VIEW, matView
    
    D3DDevice.SetTexture 0, Nothing
    
    D3DDevice.SetVertexShader FVFCH
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, Black(0), Len(Black(0))
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    D3DDevice.SetRenderState D3DRS_LIGHTING, 1
End Function

Public Function DrawCrosshair()
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    D3DXMatrixIdentity matView
    D3DDevice.SetTransform D3DTS_WORLD, matView
    D3DDevice.SetTransform D3DTS_VIEW, matView
    
    D3DDevice.SetTexture 0, Nothing
    
    D3DDevice.SetVertexShader FVFCH
    D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 4, CHList(0), Len(CHList(0))
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    D3DDevice.SetRenderState D3DRS_LIGHTING, 1
End Function

Public Function MainLoop()
    On Error Resume Next
    If GODI = 0 Then
        PlayerHeight = PlayerHeightStand
        PlayerSpeed = PlayerSpeedStand
        DIDevice.GetDeviceStateKeyboard DIState
        With DIState
            If .Key(DIK_LSHIFT) <> 0 Then
                PlayerHeight = PlayerHeightCrouch
                PlayerSpeed = PlayerSpeedCrouch
            End If
            If .Key(DIK_A) <> 0 Then
                mMesh.Move 0, -PlayerSpeed
            End If
            If .Key(DIK_D) <> 0 Then
                mMesh.Move 0, PlayerSpeed
            End If
            If .Key(DIK_W) <> 0 Then
                mMesh.Move PlayerSpeed, 0
            End If
            If .Key(DIK_S) <> 0 Then
                mMesh.Move -PlayerSpeed, 0
            End If
            If .Key(DIK_R) <> 0 Then
                mWeapon.Reload
            End If
            If .Key(DIK_SPACE) <> 0 Then
                mMesh.Jump
            End If
        End With
        NumEvents = DIDevice2.GetDeviceData(DevData, DIGDD_DEFAULT)
        
        For k% = 1 To NumEvents
            Select Case DevData(k%).lOfs
                Case DIMOFS_X
                    RotateCamera CSng(DevData(k%).lData / 100), 0
                Case DIMOFS_Y
                    RotateCamera 0, CSng(DevData(k%).lData / 100)
                Case DIMOFS_BUTTON1 ' Rechte taste
                    Zoomed = 1 Xor Zoomed
                    Zoom Zoomed
                Case DIMOFS_BUTTON0 ' Linke Taste
                    mWeapon.Shoot
            End Select
        Next k%
    End If
    
    RenderAll
    DoEvents
End Function

Public Function Zoom(InOutZ As Integer)
    Dim matProj As D3DMATRIX
    If InOutZ = 1 Then
        D3DXMatrixPerspectiveFovLH matProj, 3.14152 / ZoomIn, 1, 1, 10000
        D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    Else
        D3DXMatrixPerspectiveFovLH matProj, 3.14152 / ZoomOut, 1, 1, 10000
        D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    End If
End Function

Public Function RotateCamera(Sideways As Single, Updown As Single)
    Dim Matr As D3DMATRIX
    Yaw = Yaw + Sideways
    Pitch = Pitch + Updown
    If Pitch >= PI / 2 Then Pitch = PI / 2
    If Pitch <= -PI / 2 Then Pitch = -PI / 2
    D3DXMatrixRotationYawPitchRoll Matr, Yaw, Pitch, 0
    D3DXVec3TransformCoord EyeLookDir, MakeVector(0, 0, 1), Matr
End Function

Public Function KillApp()
    Set DVB = Nothing
    Set D3DDevice = Nothing
    Set D3D = Nothing
    End
End Function
