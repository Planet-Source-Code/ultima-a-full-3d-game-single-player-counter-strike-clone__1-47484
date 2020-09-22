VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'RENDER MODE @ 13
'        D3DDevice.SetRenderState D3DRS_SRCBLEND, 13
'        D3DDevice.SetRenderState D3DRS_DESTBLEND, 13
        
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const NPARTS = 100
Private Const FVF_PARTICLEVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

Private Speed As Single
Private mDir As D3DVECTOR
Private Spread As Single
Private ParticleSize As Single
Private Growth As Single
Private Created As Long
Private tLife As Long

Private k As Long
Private t As Long

Private LDir As D3DVECTOR
Private Quad(1 To 4) As D3DVECTOR
Private TEMP1 As D3DVECTOR, Temp2 As D3DVECTOR

Private Type Particle
    Pos As D3DVECTOR
    Dir As D3DVECTOR
    Col As Long
    Life As Single
    Size As Single
End Type

Private Type PARTICLEV
    Pos As D3DVECTOR
    Color As Long
End Type

Private VertSize As Long
Private Particles(1 To 6, 1 To NPARTS) As PARTICLEV
Private ParticlesDesc(1 To NPARTS) As Particle

Private Cols(1 To 3) As Byte

Private LastUpdate As Long
Private Interval As Single

Public Function Init(cPos As D3DVECTOR, cDir As D3DVECTOR, cSpread As Single, cSpeed As Single, cPSize As Single, cR As Byte, cG As Byte, cB As Byte, cGrowth As Single, cLife As Single)
    Spread = cSpread
    Speed = cSpeed
    mDir = cDir
    ParticleSize = cPSize
    Growth = cGrowth
    tLife = cLife
    
    Cols(1) = cR
    Cols(2) = cG
    Cols(3) = cB
    
    Randomize Timer

    For k = 1 To NPARTS
        ParticlesDesc(k).Dir = MakeVector(mDir.X + Spread / 2 - Spread * Rnd, _
                                          mDir.Y + Spread / 2 - Spread * Rnd, _
                                          mDir.Z + Spread / 2 - Spread * Rnd)
        ParticlesDesc(k).Size = cPSize + Rnd
        ParticlesDesc(k).Pos = cPos
    Next k
    
    VertSize = Len(Particles(1, 1))
    LastUpdate = GetTickCount
    Created = GetTickCount
End Function

Public Function Render()
    D3DDevice.SetRenderState D3DRS_FILLMODE, 3
    D3DDevice.SetVertexShader FVF_PARTICLEVERTEX
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 0
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0

    D3DDevice.SetTexture 0, Nothing

    For k = 1 To NPARTS
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 4, Particles(1, k), VertSize
    Next k
 
    D3DDevice.SetRenderState D3DRS_LIGHTING, 1
    D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 1
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Function

Public Function Update() As Boolean
    Dim Move As D3DVECTOR
    Dim Grow As Single
    Dim Col As Single

    If Created + tLife >= GetTickCount Then
        Col = (1 - (((Created + tLife) - GetTickCount) / tLife)) * 0.5 + 0.5
        Interval = (GetTickCount - LastUpdate) / 1000
        LastUpdate = GetTickCount
    
        Grow = (1 + Interval * Growth)
        
        TEMP1 = MakeVector(EyeLookDir.Z, 0, -EyeLookDir.X)
        D3DXVec3Normalize TEMP1, TEMP1
        D3DXVec3Cross Temp2, TEMP1, EyeLookDir
        D3DXVec3Scale TEMP1, TEMP1, ParticleSize
        D3DXVec3Scale Temp2, Temp2, ParticleSize
        D3DXVec3Add Quad(1), TEMP1, Temp2
        D3DXVec3Add Quad(2), TEMP1, MakeVector(-Temp2.X, -Temp2.Y, -Temp2.Z)
        D3DXVec3Add Quad(3), MakeVector(-TEMP1.X, -TEMP1.Y, -TEMP1.Z), MakeVector(-Temp2.X, -Temp2.Y, -Temp2.Z)
        D3DXVec3Add Quad(4), MakeVector(-TEMP1.X, -TEMP1.Y, -TEMP1.Z), Temp2
        
        
        For k = 1 To NPARTS
            D3DXVec3Scale Move, ParticlesDesc(k).Dir, Interval * Speed
            D3DXVec3Add ParticlesDesc(k).Pos, ParticlesDesc(k).Pos, Move
    
            Particles(1, k).Pos = ParticlesDesc(k).Pos
                
            If Growth = 0 Then
                D3DXVec3Add Particles(2, k).Pos, Quad(1), ParticlesDesc(k).Pos
                D3DXVec3Add Particles(3, k).Pos, Quad(2), ParticlesDesc(k).Pos
                D3DXVec3Add Particles(4, k).Pos, Quad(3), ParticlesDesc(k).Pos
                D3DXVec3Add Particles(5, k).Pos, Quad(4), ParticlesDesc(k).Pos
                Particles(6, k).Pos = Particles(2, k).Pos
            Else
                With ParticlesDesc(k)
                    D3DXVec3Add Particles(2, k).Pos, MakeVector(Quad(1).X * Grow, Quad(1).Y * Growth, Quad(1).Z * Growth), .Pos
                    D3DXVec3Add Particles(3, k).Pos, MakeVector(Quad(2).X * Grow, Quad(2).Y * Growth, Quad(2).Z * Growth), .Pos
                    D3DXVec3Add Particles(4, k).Pos, MakeVector(Quad(3).X * Grow, Quad(3).Y * Growth, Quad(3).Z * Growth), .Pos
                    D3DXVec3Add Particles(5, k).Pos, MakeVector(Quad(4).X * Grow, Quad(4).Y * Growth, Quad(4).Z * Growth), .Pos
                End With
                Particles(6, k).Pos = Particles(2, k).Pos
            End If
            Particles(1, k).Color = ARGB2LONG(255 * Col, CLng(Cols(1)), CLng(Cols(2)), CLng(Cols(3)))
            For t = 2 To 6
                Particles(t, k).Color = ARGB2LONG(255, CLng(Cols(1)), CLng(Cols(2)), CLng(Cols(3)))
            Next t
        Next k
        Growth = Growth * 1.001
    End If
End Function

