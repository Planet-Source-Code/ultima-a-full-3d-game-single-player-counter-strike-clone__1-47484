VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const Life As Long = 10
Private Const FSize = 0.5
Private Const CDown As Long = 1000

Private StartAnim As Long
Private CurrAnim As Integer

Private AnimLength(1 To 3) As Long

Private KeyFrame1() As D3DVERTEX
Private KeyFrame2() As D3DVERTEX
Private KeyFrame3() As D3DVERTEX
Private KeyFrame4() As D3DVERTEX
Private KeyFrame5() As D3DVERTEX
Private Interpol() As D3DVERTEX

Private Infin As Integer

Private TempV As D3DVECTOR
Private Mesh As D3DXMesh
Private CTime As Long
Private Interval As Single

Private NumMaterials As Long
Private Materials() As D3DMATERIAL8
Private Textures() As Direct3DTexture8

Private Position As D3DVECTOR
Public tLife As Long

Private EMove As D3DMATRIX
Private ERotate As D3DMATRIX
Private EMatrix As D3DMATRIX
Public CAngle As Single
Private LookDir As D3DVECTOR

Private mIdentity As D3DMATRIX
Private TV As D3DVECTOR

Private WayPoints(1 To 9) As D3DVECTOR
Private TurnPoints(1 To 9) As Single
Private CurrWaypoint As Integer
Private TTN As Single
Private LastUpdate As Long

Private Mouth As D3DVECTOR
Private SVisible As Integer
Private SHList(1) As CHVERTEX
Private ShootCool As Single

Private Death As Boolean
Private FList(11) As FVERTEX
Private SoundShoot As New cSound
Private FlashTexture As Direct3DTexture8
Private FlashTrans As D3DMATRIX

Private Spotted As Boolean

Public Function Init(Path As String)
    Dim mtrlBuffer As D3DXBuffer
    Dim I As Long
    Dim TextureFile As String
    
    Death = False
    
    Set Mesh = D3DX.LoadMeshFromX(Path, D3DXMESH_MANAGED, D3DDevice, Nothing, mtrlBuffer, NumMaterials)

    ReDim Materials(NumMaterials) As D3DMATERIAL8
    ReDim Textures(NumMaterials) As Direct3DTexture8

    For I = 0 To NumMaterials - 1
    
        D3DX.BufferGetMaterial mtrlBuffer, I, Materials(I)
        
        Materials(I).Ambient = Materials(I).diffuse

        TextureFile = D3DX.BufferGetTextureName(mtrlBuffer, I)

        If TextureFile <> "" Then
            Set Textures(I) = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\ENE-1" + TextureFile)
        End If
        
    Next I
    
    Set FlashTexture = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\flash.bmp")
    
    ReDim KeyFrame1(0 To Mesh.GetNumVertices) As D3DVERTEX
    ReDim KeyFrame2(0 To Mesh.GetNumVertices) As D3DVERTEX
    ReDim KeyFrame3(0 To Mesh.GetNumVertices) As D3DVERTEX
    ReDim KeyFrame4(0 To Mesh.GetNumVertices) As D3DVERTEX
    ReDim KeyFrame5(0 To Mesh.GetNumVertices) As D3DVERTEX
    ReDim Interpol(0 To Mesh.GetNumVertices) As D3DVERTEX
    
    D3DXMeshVertexBuffer8GetData Mesh, 0, Len(Interpol(0)) * Mesh.GetNumVertices, 0, KeyFrame1(0)
    D3DXMeshVertexBuffer8GetData Mesh, 0, Len(Interpol(0)) * Mesh.GetNumVertices, 0, Interpol(0)
    Set Mesh = D3DX.LoadMeshFromX(App.Path + "\Ene2.x", D3DXMESH_MANAGED, D3DDevice, Nothing, mtrlBuffer, NumMaterials)
    D3DXMeshVertexBuffer8GetData Mesh, 0, Len(Interpol(0)) * Mesh.GetNumVertices, 0, KeyFrame2(0)
    Set Mesh = D3DX.LoadMeshFromX(App.Path + "\Ene3.x", D3DXMESH_MANAGED, D3DDevice, Nothing, mtrlBuffer, NumMaterials)
    D3DXMeshVertexBuffer8GetData Mesh, 0, Len(Interpol(0)) * Mesh.GetNumVertices, 0, KeyFrame3(0)
    Set Mesh = D3DX.LoadMeshFromX(App.Path + "\Ene4.x", D3DXMESH_MANAGED, D3DDevice, Nothing, mtrlBuffer, NumMaterials)
    D3DXMeshVertexBuffer8GetData Mesh, 0, Len(Interpol(0)) * Mesh.GetNumVertices, 0, KeyFrame4(0)
    Set Mesh = D3DX.LoadMeshFromX(App.Path + "\Ene5.x", D3DXMESH_MANAGED, D3DDevice, Nothing, mtrlBuffer, NumMaterials)
    D3DXMeshVertexBuffer8GetData Mesh, 0, Len(Interpol(0)) * Mesh.GetNumVertices, 0, KeyFrame5(0)
    
    AnimLength(1) = 300
    AnimLength(2) = 300
    AnimLength(3) = 100
    
    TurnPoints(5) = PI / 2
    TurnPoints(6) = PI
    TurnPoints(7) = PI * 5 / 4
    TurnPoints(8) = PI
    TurnPoints(9) = PI / 2
    TurnPoints(1) = 0
    TurnPoints(2) = PI * 7 / 4
    TurnPoints(3) = 0
    TurnPoints(4) = PI * 3 / 2
    
    WayPoints(5) = MakeVector(-1.845205, 0, 2.014498)
    WayPoints(6) = MakeVector(86.57822, 0, 2.272206)
    WayPoints(7) = MakeVector(82.34792, 0, -53.47541)
    WayPoints(8) = MakeVector(-0.3356083, 0, -136.6905)
    WayPoints(9) = MakeVector(-2.144823, 0, -204.9477)
    WayPoints(1) = MakeVector(173.6679, 0, -205.8015)
    WayPoints(2) = MakeVector(174.1602, 0, -143.1451)
    WayPoints(3) = MakeVector(94.5705, 0, -57.34747)
    WayPoints(4) = MakeVector(84.67476, 0, 1.902351)
    
    CurrWaypoint = 1
    
    SHList(0).Col = &HCCFFFF00
    SHList(1).Col = &HCCFFFF00
    
    LookDir.Z = 1
    
    Position = WayPoints(1)

    D3DXMatrixRotationY ERotate, CAngle
    D3DXVec3TransformCoord LookDir, MakeVector(0, 0, 1), ERotate
    D3DXMatrixTranslation EMove, Position.X, Position.Y, Position.Z
    D3DXMatrixMultiply EMatrix, ERotate, EMove
    
    SoundShoot.Create App.Path + "\shoot.wav", -1000
    
    tLife = Life
    
    D3DXMatrixIdentity mIdentity
    
    Mouth.X = 4.5
    Mouth.Y = 6.864
    Mouth.Z = 0.3
    
    InitFlash
    
    Spotted = False
    
    PlayAnimation 1
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
    
    D3DXMatrixTranslation FlashTrans, Mouth.X, Mouth.Y, Mouth.Z
End Function

Public Function Render()
    Dim Leng As Single, vOut As D3DVECTOR
    Interval = (GetTickCount - LastUpdate) / 100
    CTime = GetTickCount
    
    If Death = False And GODI = 0 Then
        D3DXVec3Subtract TV, MakeVector(EyePosition.X, EyePosition.Y + PlayerHeight, EyePosition.Z), Position
        Leng = Sqr(TV.X * TV.X + TV.Y * TV.Y + TV.Z * TV.Z)
        If Leng > 0 Then
            TV.X = TV.X / Leng
            TV.Y = TV.Y / Leng
            TV.Z = TV.Z / Leng
            If (mMesh.CollisionDetection(MakeVector(Position.X, Position.Y + 4, Position.Z), TV, Leng, vOut, Leng) = True) _
               Or (DotProduct(TV, LookDir) < 0 And Spotted = False) Then
                CurrAnim = 1
                TTN = TTN + Interval / 100
                If CurrWaypoint < 9 Then
                    D3DXVec3Lerp Position, WayPoints(CurrWaypoint), WayPoints(CurrWaypoint + 1), TTN
                    CAngle = TurnPoints(CurrWaypoint)
                    If TTN >= 1 Then
                        CurrWaypoint = CurrWaypoint + 1
                        TTN = 0
                    End If
                Else
                    D3DXVec3Lerp Position, WayPoints(CurrWaypoint), WayPoints(1), TTN
                    CAngle = TurnPoints(CurrWaypoint)
                    If TTN >= 1 Then
                        CurrWaypoint = 1
                        TTN = 0
                    End If
                End If
            
                D3DXMatrixRotationY ERotate, CAngle
                D3DXVec3TransformCoord LookDir, MakeVector(0, 0, 1), ERotate
                D3DXMatrixTranslation EMove, Position.X, Position.Y, Position.Z
                D3DXMatrixMultiply EMatrix, ERotate, EMove
                Spotted = False
            Else
                SoundShoot.SetVolume -Leng * 10
                D3DXVec3Normalize TV, MakeVector(TV.X, 0, TV.Z)
                Leng = DotProduct(TV, MakeVector(0, 0, 1))
                Leng = Asin(Sqr(1 - Leng * Leng))
                If TV.X <= 0 Then Leng = 2 * PI - Leng
                If TV.Z <= 0 Then Leng = PI - Leng
                CAngle = 0
                D3DXMatrixRotationY ERotate, Leng
                D3DXVec3TransformCoord LookDir, MakeVector(0, 0, 1), ERotate
                D3DXMatrixTranslation EMove, Position.X, Position.Y, Position.Z
                D3DXMatrixMultiply EMatrix, ERotate, EMove
                If CurrAnim = 1 Then
                    PlayAnimation 3
                    Spotted = True
                End If
                
                If ShootCool <= 0 Then
                    SoundShoot.PlaySound False
                    HitPlayer
                    ShootCool = CDown
                    SHList(0).Pos = MakeVector(Position.X + LookDir.X * Mouth.X + LookDir.Z * Mouth.Z, _
                                               Position.Y + Mouth.Y, _
                                               Position.Z + LookDir.Z * Mouth.X - LookDir.X * Mouth.Z)
                    SHList(1).Pos = MakeVector(EyePosition.X, EyePosition.Y + PlayerHeight - 1, EyePosition.Z)
                    SVisible = 3
                    D3DXMatrixTranslation FlashTrans, SHList(0).Pos.X, SHList(0).Pos.Y, SHList(0).Pos.Z
                    D3DXMatrixMultiply FlashTrans, ERotate, FlashTrans
                Else
                    ShootCool = ShootCool - (CTime - LastUpdate)
                End If
            End If
        End If
    End If
    
    If CurrAnim <> 0 Then
        Select Case CurrAnim
            Case 1
                Leng = (CTime - StartAnim) / AnimLength(1)
                If Leng > 1 Then Leng = 1
                If Infin = 0 Then
                    For k = 0 To Mesh.GetNumVertices
                        D3DXVec3Lerp TempV, _
                                     MakeVector(KeyFrame2(k).X, KeyFrame2(k).Y, KeyFrame2(k).Z), _
                                     MakeVector(KeyFrame3(k).X, KeyFrame3(k).Y, KeyFrame3(k).Z), _
                                     (CTime - StartAnim) / AnimLength(1)
                        Interpol(k).X = TempV.X
                        Interpol(k).Y = TempV.Y
                        Interpol(k).Z = TempV.Z
                    Next k
                Else
                    For k = 0 To Mesh.GetNumVertices
                        D3DXVec3Lerp TempV, _
                                     MakeVector(KeyFrame3(k).X, KeyFrame3(k).Y, KeyFrame3(k).Z), _
                                     MakeVector(KeyFrame2(k).X, KeyFrame2(k).Y, KeyFrame2(k).Z), _
                                     (CTime - StartAnim) / AnimLength(1)
                        Interpol(k).X = TempV.X
                        Interpol(k).Y = TempV.Y
                        Interpol(k).Z = TempV.Z
                    Next k
                End If
                If Leng = 1 Then
                    StartAnim = GetTickCount
                    Infin = Infin Xor 1
                End If
            Case 2
                Leng = (CTime - StartAnim) / AnimLength(2)
                If Leng > 1 Then Leng = 1
                For k = 0 To Mesh.GetNumVertices
                    D3DXVec3Lerp TempV, _
                                 MakeVector(KeyFrame1(k).X, KeyFrame1(k).Y, KeyFrame1(k).Z), _
                                 MakeVector(KeyFrame4(k).X, KeyFrame4(k).Y, KeyFrame4(k).Z), _
                                 Leng
                    Interpol(k).X = TempV.X
                    Interpol(k).Y = TempV.Y
                    Interpol(k).Z = TempV.Z
                Next k
                If Leng = 1 Then
                    CurrAnim = 0
                End If
            Case 3
                Leng = (CTime - StartAnim) / AnimLength(3)
                If Leng > 1 Then Leng = 1
                For k = 0 To Mesh.GetNumVertices
                    D3DXVec3Lerp TempV, _
                                 MakeVector(KeyFrame1(k).X, KeyFrame1(k).Y, KeyFrame1(k).Z), _
                                 MakeVector(KeyFrame5(k).X, KeyFrame5(k).Y, KeyFrame5(k).Z), _
                                 Leng
                    Interpol(k).X = TempV.X
                    Interpol(k).Y = TempV.Y
                    Interpol(k).Z = TempV.Z
                Next k
                If Leng = 1 Then
                    CurrAnim = 0
                End If
        End Select
        
        D3DXMeshVertexBuffer8SetData Mesh, 0, Len(Interpol(0)) * Mesh.GetNumVertices, 0, Interpol(0)
    End If
    
    D3DDevice.SetTransform D3DTS_WORLD, EMatrix
    
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
    
    LastUpdate = GetTickCount
    
End Function

Public Function SpotEnemy()
    Spotted = True
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

    D3DDevice.SetTexture 0, FlashTexture

    D3DDevice.SetVertexShader FVFF
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 4, FList(0), Len(FList(0))
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    D3DDevice.SetRenderState D3DRS_LIGHTING, 1
End Function

Public Function MoveRot(ForW As Single, SideW As Single, aY As Single)
    CAngle = CAngle + aY
    If CAngle < 0 Then CAngle = CAngle + 2 * PI
    If CAngle > 2 * PI Then CAngle = CAngle - 2 * PI
    D3DXMatrixRotationY ERotate, CAngle
    D3DXVec3TransformCoord LookDir, MakeVector(0, 0, 1), ERotate
    D3DXMatrixTranslation EMove, LookDir.X * ForW + Position.X, Position.Y, LookDir.Z * ForW + Position.Z
    D3DXVec3TransformCoord Position, MakeVector(0, 0, 0), EMove
    D3DXMatrixMultiply EMatrix, ERotate, EMove
End Function

Public Function Collision(StartPo As D3DVECTOR, Dir As D3DVECTOR, mDist As Single, Hit As D3DVECTOR) As Boolean
    Dim Hits As Long, PU As Single, PV As Single, Dist As Single, nHits As Long, Ind As Long
    Dim StartP As D3DVECTOR, TempM As D3DMATRIX
    D3DXMatrixTranslation TempM, -Position.X, -Position.Y, -Position.Z
    D3DXVec3TransformCoord StartP, StartPo, TempM
    D3DX.Intersect Mesh, StartP, Dir, Hits, Ind, PU, PV, Dist, nHits
    If Hits = 1 Then
        If mDist >= Dist Then
            Collision = True
            Hit = MakeVector(StartPo.X + Dir.X * Dist, StartPo.Y + Dir.Y * Dist, StartPo.Z + Dir.Z * Dist)
            Exit Function
        End If
    End If
    Collision = False
End Function

Public Function PlayAnimation(NumAn As Integer)
    If Death = False Then
        If NumAn = 2 Then Death = True
        StartAnim = GetTickCount
        CurrAnim = NumAn
    End If
End Function

Public Function IsHit()
    If tLife <= 1 Then
        PlayAnimation 2
        tLife = 0
        GODI = 2
    Else
        tLife = tLife - 1
    End If
End Function

Public Function Revive()
    Death = False
    CurrAnim = 1
    tLife = Life
End Function
