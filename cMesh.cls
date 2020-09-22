VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const FlightTime As Long = 300
Private Const JSpeed As Single = 0.25

Private GravityAcc As Single
Private Gravity As Single
Private ColHeight As Single
Private LastUpdate As Long
Private Interval As Single
Private Mesh As D3DXMesh
Private NumMaterials As Long
Private Materials() As D3DMATERIAL8
Private Textures() As Direct3DTexture8
Private mIdentity As D3DMATRIX

Private InAir As Long

Private TEMP1 As Single

Public Function Init(Path As String)
    Dim mtrlBuffer As D3DXBuffer
    Dim I As Long
    Dim TextureFile As String
    Dim ModelFile As String
    
    Set Mesh = D3DX.LoadMeshFromX(Path, D3DXMESH_MANAGED, D3DDevice, Nothing, mtrlBuffer, NumMaterials)

    ReDim Materials(NumMaterials) As D3DMATERIAL8
    ReDim Textures(NumMaterials) As Direct3DTexture8

    For I = 0 To NumMaterials - 1
    
        D3DX.BufferGetMaterial mtrlBuffer, I, Materials(I)
        
        Materials(I).Ambient = Materials(I).diffuse

        TextureFile = D3DX.BufferGetTextureName(mtrlBuffer, I)

        If TextureFile <> "" Then
            Set Textures(I) = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\MAP-1" + TextureFile)
        End If
        
    Next I
    GravityAcc = 20
    Gravity = 0.5
    ColHeight = 4.7
    InAir = -99
    LastUpdate = GetTickCount
    D3DXMatrixIdentity mIdentity
End Function

Public Function Render()
    D3DDevice.SetTransform D3DTS_WORLD, mIdentity
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    For k = 0 To NumMaterials - 1
         D3DDevice.SetMaterial Materials(k)
         D3DDevice.SetTexture 0, Textures(k)
         Mesh.DrawSubset k
    Next k
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Function

Public Function CollisionDetection(StartP As D3DVECTOR, Dir As D3DVECTOR, mDist As Single, Hit As D3DVECTOR, aDist As Single) As Boolean
    Dim Hits As Long, PU As Single, PV As Single, Dist As Single, nHits As Long, Ind As Long
    D3DX.Intersect Mesh, StartP, Dir, Hits, Ind, PU, PV, Dist, nHits
    CollisionDetection = False
    If nHits >= 1 Then
        If mDist >= Dist Then
            CollisionDetection = True
            aDist = Dist
            Hit = MakeVector(StartP.X + Dir.X * Dist, StartP.Y + Dir.Y * Dist, StartP.Z + Dir.Z * Dist)
            Exit Function
        End If
    End If
End Function

Private Function Collision(StartP As D3DVECTOR, Dir As D3DVECTOR, mDist As Single, Hit As D3DVECTOR) As Boolean
    Dim Hits As Long, PU As Single, PV As Single, Dist As Single, nHits As Long, Ind As Long
    D3DX.Intersect Mesh, StartP, Dir, Hits, Ind, PU, PV, Dist, nHits
    If nHits >= 1 Then
        If mDist >= Dist Then
            Collision = True
            Hit = MakeVector(StartP.X + Dir.X * (Dist), StartP.Y + Dir.Y * (Dist), StartP.Z + Dir.Z * (Dist))
            TEMP1 = Dist
            Exit Function
        End If
    End If
    Collision = False
End Function

Private Function CollisionDown(StartP As D3DVECTOR, mDist As Single, Hit As D3DVECTOR) As Boolean
    Dim Hits As Long, PU As Single, PV As Single, Dist As Single, nHits As Long, Ind As Long
    D3DX.Intersect Mesh, StartP, MakeVector(0, -1, 0), Hits, Ind, PU, PV, Dist, nHits
    If nHits >= 1 Then
        If mDist >= Dist Then
            CollisionDown = True
            Hit = MakeVector(StartP.X, StartP.Y - Dist, StartP.Z)
            Exit Function
        End If
    End If
    CollisionDown = False
End Function

Public Function Update()
    Dim Res As D3DVECTOR
    Interval = (GetTickCount - LastUpdate) / 10

    If InAir <= 0 Then
        If CollisionDown(EyePosition, Gravity, EyePosition) = False Then
            EyePosition.Y = EyePosition.Y - Gravity
        Else
            InAir = -99
        End If
    ElseIf InAir > 0 Then
        InAir = InAir - (Interval * 10)
        EyePosition.Y = EyePosition.Y + (JSpeed * Interval) * (InAir / FlightTime)
    End If
    
    LastUpdate = GetTickCount
End Function

Public Function Move(ForW As Single, SideW As Single)
    Dim MoveDir As D3DVECTOR, Temp As Single
    Dim Res As D3DVECTOR
    
    EyePosition.Y = EyePosition.Y + ColHeight
    
    D3DXVec3Normalize MoveDir, MakeVector(EyeLookDir.X, 0, EyeLookDir.Z)
    
    If ForW <> 0 Then
        If ForW < 0 Then
            Temp = -ForW * Interval
            If Collision(EyePosition, MakeVector(-MoveDir.X, 0, -MoveDir.Z), Temp * 4, Res) = False Then
                If mEnemy.Collision(EyePosition, MakeVector(-MoveDir.X, 0, -MoveDir.Z), Temp * 4, Res) = False Then
                    EyePosition.X = EyePosition.X - MoveDir.X * Temp
                    EyePosition.Z = EyePosition.Z - MoveDir.Z * Temp
                End If
            End If
        Else
            Temp = ForW * Interval
            If Collision(EyePosition, MoveDir, Temp * 4, Res) = False Then
                If mEnemy.Collision(EyePosition, MoveDir, Temp * 4, Res) = False Then
                    EyePosition.X = EyePosition.X + MoveDir.X * Temp
                    EyePosition.Z = EyePosition.Z + MoveDir.Z * Temp
                End If
            End If
        End If
        If CollisionDown(EyePosition, ColHeight, EyePosition) = False Then
            EyePosition.Y = EyePosition.Y - ColHeight
        Else
            InAir = -99
        End If
    End If
    If SideW <> 0 Then
        MoveDir = MakeVector(MoveDir.Z, 0, -MoveDir.X)
        If SideW < 0 Then
            Temp = -SideW * Interval
            If Collision(EyePosition, MakeVector(-MoveDir.X, 0, -MoveDir.Z), Temp * 4, Res) = False Then
                EyePosition.X = EyePosition.X - MoveDir.X * Temp
                EyePosition.Z = EyePosition.Z - MoveDir.Z * Temp
            End If
        Else
            Temp = SideW * Interval
            If Collision(EyePosition, MoveDir, Temp * 4, Res) = False Then
                EyePosition.X = EyePosition.X + MoveDir.X * Temp
                EyePosition.Z = EyePosition.Z + MoveDir.Z * Temp
            End If
        End If
        If CollisionDown(EyePosition, ColHeight, EyePosition) = False Then
            EyePosition.Y = EyePosition.Y - ColHeight
        Else
            InAir = -99
        End If
    End If
End Function

Public Function Jump()
    If InAir = -99 Then
        InAir = FlightTime
    End If
End Function









