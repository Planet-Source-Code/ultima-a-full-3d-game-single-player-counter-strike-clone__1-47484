VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWeapon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const ReloadTime = 40
Private Const FSize = 0.5
Private Const NPLife = 1000
Private Const CDown = 300

Private Mesh As D3DXMesh
Private NumMaterials As Long
Private Materials() As D3DMATERIAL8
Private Textures() As Direct3DTexture8
Private mMatrix As D3DMATRIX
Private mIdentity As D3DMATRIX
Private Interval As Single
Private Recoil As Single
Private mRecoil As Single
Private LastUpdate As Long
Private mYaw As Single
Private mPitch As Single
Private mPos As D3DVECTOR
Private Change As Boolean
Private Mouth As D3DVECTOR
Private TempM As D3DMATRIX
Private BulletStart As D3DVECTOR

Private NoAction As Single
Private Interpol() As D3DVERTEX
Private IStart() As D3DVERTEX
Private IReload() As D3DVERTEX

Private SoundShoot As New cSound
Private SoundReload As New cSound

Private TempV As D3DVECTOR

Private LastShot As Long

Public Clips As Long
Public Bullets As Long

Private SVisible As Integer
Private SHList(1) As CHVERTEX

Private pLife As Single

Private FList(11) As FVERTEX
Private FlashTexture As Direct3DTexture8
Private FlashTrans As D3DMATRIX

Public Function Init(Path As String, RelPath As String)
    Dim mtrlBuffer As D3DXBuffer
    Dim I As Long
    Dim TextureFile As String
    Dim ModelFile As String
    Dim MReload As D3DXMesh
    
    Set MReload = D3DX.LoadMeshFromX(RelPath, D3DXMESH_MANAGED, D3DDevice, Nothing, mtrlBuffer, NumMaterials)
    Set Mesh = D3DX.LoadMeshFromX(Path, D3DXMESH_MANAGED, D3DDevice, Nothing, mtrlBuffer, NumMaterials)

    ReDim Materials(NumMaterials) As D3DMATERIAL8
    ReDim Textures(NumMaterials) As Direct3DTexture8

    For I = 0 To NumMaterials - 1
    
        D3DX.BufferGetMaterial mtrlBuffer, I, Materials(I)
        
        Materials(I).Ambient = Materials(I).diffuse

        TextureFile = D3DX.BufferGetTextureName(mtrlBuffer, I)

        If TextureFile <> "" Then
            Set Textures(I) = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\WEP-1" + TextureFile)
        End If
        
    Next I
    
    Set FlashTexture = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\flash.bmp")
    
    Mouth.X = 0.378
    Mouth.Y = 0.658
    Mouth.Z = 4.9
    
    mRecoil = 0.5
    Bullets = 30
    Clips = 5
    
    InitFlash
    
    ReDim Interpol(Mesh.GetNumVertices) As D3DVERTEX
    ReDim IStart(Mesh.GetNumVertices) As D3DVERTEX
    ReDim IReload(Mesh.GetNumVertices) As D3DVERTEX
    
    D3DXMeshVertexBuffer8GetData Mesh, 0, Len(Interpol(0)) * Mesh.GetNumVertices, 0, IStart(0)
    D3DXMeshVertexBuffer8GetData Mesh, 0, Len(Interpol(0)) * Mesh.GetNumVertices, 0, Interpol(0)
    D3DXMeshVertexBuffer8GetData MReload, 0, Len(Interpol(0)) * Mesh.GetNumVertices, 0, IReload(0)
    
    SHList(0).Col = &HCCFFFF00
    SHList(1).Col = &HCCFFFF00
    
    SoundShoot.Create App.Path + "\shoot.wav", 0
    SoundReload.Create App.Path + "\reload.wav", 0
    
    D3DXMatrixIdentity mIdentity
    
    LastUpdate = GetTickCount
End Function

Private Function InitFlash()
    FList(0) = MakeFVertex(-FSize, 0, 0, 0, 1, &HFFFFFFFF)
    FList(1) = MakeFVertex(0, 0, 0, 1, 1, &HFFFFFF)
    FList(2) = MakeFVertex(0, FSize * 2, 0, 1, 0, &HFFFFFFFF)
    FList(3) = MakeFVertex(FSize, 0, 0, 0, 1, &HFFFFFFFF)
    FList(4) = MakeFVertex(0, 0, 0, 1, 1, &HFFFFFF)
    FList(5) = MakeFVertex(0, FSize * 2, 0, 1, 0, &HFFFFFFFF)
    
    FList(6) = MakeFVertex(-FSize, 0, 0, 0, 1, &HFFFFFFFF)
    FList(7) = MakeFVertex(0, 0, 0, 1, 1, &HFFFFFF)
    FList(8) = MakeFVertex(0, -FSize * 2, 0, 1, 0, &HFFFFFFFF)
    FList(9) = MakeFVertex(FSize, 0, 0, 0, 1, &HFFFFFFFF)
    FList(10) = MakeFVertex(0, 0, 0, 1, 1, &HFFFFFF)
    FList(11) = MakeFVertex(0, -FSize * 2, 0, 1, 0, &HFFFFFFFF)
    
    D3DXMatrixTranslation FlashTrans, Mouth.X, -Mouth.Y, Mouth.Z
End Function

Public Function Adjust(Yaw As Single, Pitch As Single, Pos As D3DVECTOR)
    mYaw = Yaw
    mPitch = Pitch
    mPos = Pos
    Change = True
End Function

Public Function Render()
    Interval = (GetTickCount - LastUpdate) / 100
    LastUpdate = GetTickCount
    
    If pLife > 0 Then
        D3DDevice.SetTransform D3DTS_WORLD, mIdentity
        mEmitter.Update
        mEmitter.Render
        pLife = pLife - Interval
    End If
    
    If Recoil > 0 Then
        D3DXMatrixTranslation TempM, 0, 0, -Recoil
        D3DXMatrixRotationYawPitchRoll mMatrix, mYaw, mPitch, 0
        D3DXMatrixMultiply mMatrix, TempM, mMatrix
        D3DXMatrixTranslation TempM, mPos.X, mPos.Y + Breather, mPos.Z
        D3DXMatrixMultiply mMatrix, mMatrix, TempM
        D3DDevice.SetTransform D3DTS_WORLD, mMatrix
        Recoil = Recoil - Interval
    ElseIf Change = True Then
        D3DXMatrixRotationYawPitchRoll mMatrix, mYaw, mPitch, 0
        D3DXMatrixTranslation TempM, mPos.X, mPos.Y + Breather, mPos.Z
        D3DXMatrixMultiply mMatrix, mMatrix, TempM
        D3DDevice.SetTransform D3DTS_WORLD, mMatrix
        Change = False
    End If
    
    If NoAction > 0 Then
    
        If NoAction >= ReloadTime / 2 Then
            For k = 0 To Mesh.GetNumVertices
                D3DXVec3Lerp TempV, MakeVector(IStart(k).X, IStart(k).Y, IStart(k).Z), _
                                    MakeVector(IReload(k).X, IReload(k).Y, IReload(k).Z), 2 - (2 * (NoAction / ReloadTime))
                Interpol(k).X = TempV.X
                Interpol(k).Y = TempV.Y
                Interpol(k).Z = TempV.Z
            Next k
        
            
            NoAction = NoAction - 5 * Interval
        Else
            For k = 0 To Mesh.GetNumVertices
                D3DXVec3Lerp TempV, MakeVector(IReload(k).X, IReload(k).Y, IReload(k).Z), _
                                    MakeVector(IStart(k).X, IStart(k).Y, IStart(k).Z), 1 - (2 * (NoAction / ReloadTime))
                Interpol(k).X = TempV.X
                Interpol(k).Y = TempV.Y
                Interpol(k).Z = TempV.Z
            Next k
            
            NoAction = NoAction - 5 * Interval

                
        End If
        
        If NoAction < 0 Then
            NoAction = 0
            For k = 1 To Mesh.GetNumVertices
                Interpol(k) = IStart(k)
            Next k
        End If
        
        D3DXMeshVertexBuffer8SetData Mesh, 0, Len(Interpol(0)) * Mesh.GetNumVertices, 0, Interpol(0)
        
    End If
    
    For k = 0 To NumMaterials - 1
        D3DDevice.SetMaterial Materials(k)
        D3DDevice.SetTexture 0, Textures(k)
        Mesh.DrawSubset k
    Next k
    
    If SVisible > 0 Then
        DrawShot
        DrawFlash
        SVisible = SVisible - 1
    End If
    
    D3DDevice.SetTransform D3DTS_WORLD, mIdentity
End Function

Public Function DrawShot()
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    D3DDevice.SetTransform D3DTS_WORLD, mIdentity

    D3DDevice.SetTexture 0, Nothing

    D3DDevice.SetVertexShader FVFCH
    D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, SHList(0), Len(SHList(0))
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    D3DDevice.SetRenderState D3DRS_LIGHTING, 1
End Function

Public Function DrawFlash()
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    D3DDevice.SetTransform D3DTS_WORLD, FlashTrans
    D3DDevice.SetTransform D3DTS_VIEW, mIdentity

    D3DDevice.SetTexture 0, FlashTexture

    D3DDevice.SetVertexShader FVFF
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 4, FList(0), Len(FList(0))
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    D3DDevice.SetRenderState D3DRS_LIGHTING, 1
End Function

Public Function Shoot()
    If Bullets > 0 And NoAction = 0 And CDown + LastShot <= GetTickCount Then
        Dim HitD As Single
        SVisible = 3
        BulletStart = MakeVector(EyePosition.X + EyeLookDir.Z * Mouth.X + EyeLookDir.X * Mouth.Z, _
                                 EyePosition.Y + PlayerHeight - Mouth.Y + Breather + EyeLookDir.Y * Mouth.Z, _
                                 EyePosition.Z - EyeLookDir.X * Mouth.X + EyeLookDir.Z * Mouth.Z)
        If mMesh.CollisionDetection(MakeVector(EyePosition.X, EyePosition.Y + PlayerHeight, EyePosition.Z), EyeLookDir, 1000, SHList(1).Pos, HitD) = False Then
            D3DXVec3Scale SHList(1).Pos, EyeLookDir, 1000
            D3DXVec3Add SHList(1).Pos, SHList(1).Pos, BulletStart
            HitD = 1000
        Else
            mEmitter.Init SHList(1).Pos, MakeVector(0, 1, 0), 2, 2, 0.1, 125, 125, 125, 2, 500
            pLife = NPLife
        End If
        If mEnemy.Collision(BulletStart, EyeLookDir, HitD, TempV) = True Then
            mEnemy.IsHit
            mEmitter.Init TempV, MakeVector(0, 1, 0), 5, 3, 0.1, 200, 0, 0, 2, 500
            pLife = NPLife
            mEnemy.SpotEnemy
        End If
        SHList(0).Pos = BulletStart
        Recoil = mRecoil
        Bullets = Bullets - 1
        SoundShoot.PlaySound False
        LastShot = GetTickCount
    End If
End Function

Public Function Reload()
    If Clips > 0 And NoAction = 0 Then
        Clips = Clips - 1
        Bullets = 30
        NoAction = ReloadTime
        SoundReload.PlaySound False
    End If
End Function
