VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ESkyBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private LastX As Single, LastZ As Single, LastY As Single
Private Texture As Direct3DTexture8
Private STexture As Direct3DTexture8
Private Verts(1 To 30) As FVERTEX
Private mVB As Direct3DVertexBuffer8
Private VertexSizeInBytes

Private Const MX As Single = 1000
Private Const MY As Single = 1000
Private Const TY As Single = MY
Private Const TX As Single = MX
Private Const GM As Single = -1000

Public Function Init(FPath As String)
    Set Texture = D3DX.CreateTextureFromFile(D3DDevice, FPath)
    
    VertexSizeInBytes = Len(Verts(1))
    
    Verts(1) = MakeFVertex(-TX, TY, -TX, 0, 0, &HFFFFFF)
    Verts(2) = MakeFVertex(TX, TY, -TX, 0.5, 0, &HFFFFFF)
    Verts(3) = MakeFVertex(TX, TY, TX, 0.5, 1, &HFFFFFF)
    Verts(4) = Verts(1)
    Verts(5) = Verts(3)
    Verts(6) = MakeFVertex(-TX, TY, TX, 0, 1, &HFFFFFF)
    
    Verts(7) = MakeFVertex(-MX, MY, -MX, 0.51, 0, &HFFFFFF)
    Verts(8) = MakeFVertex(-MX, GM, -MX, 0.51, 1, &HFFFFFF)
    Verts(9) = MakeFVertex(MX, GM, -MX, 0.99, 1, &HFFFFFF)
    Verts(10) = Verts(9)
    Verts(11) = Verts(7)
    Verts(12) = MakeFVertex(MX, MY, -MX, 0.99, 0, &HFFFFFF)
    
    Verts(13) = Verts(12)
    Verts(14) = Verts(9)
    Verts(15) = MakeFVertex(MX, GM, MX, 0.51, 1, &HFFFFFF)
    Verts(16) = Verts(12)
    Verts(17) = Verts(15)
    Verts(18) = MakeFVertex(MX, MY, MX, 0.51, 0, &HFFFFFF)
    
    Verts(19) = Verts(18)
    Verts(20) = Verts(15)
    Verts(21) = MakeFVertex(-MX, MY, MX, 0.99, 0, &HFFFFFF)
    Verts(22) = Verts(15)
    Verts(23) = Verts(21)
    Verts(24) = MakeFVertex(-MX, GM, MX, 0.99, 1, &HFFFFFF)
    
    Verts(25) = Verts(24)
    Verts(26) = Verts(21)
    Verts(27) = Verts(7)
    Verts(28) = Verts(7)
    Verts(29) = Verts(24)
    Verts(30) = Verts(8)
    
    Set mVB = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 30, 0, FVFF, D3DPOOL_DEFAULT)
    
    ReloadVB
End Function

Public Function Render()
    D3DDevice.SetStreamSource 0, mVB, VertexSizeInBytes
    D3DDevice.SetVertexShader FVFF
    D3DDevice.SetTexture 0, Texture
    D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 10
End Function

Public Function Update()
    Dim DifX As Single, DifY As Single, DifZ As Single
    DifX = EyePosition.X - LastX
    DifY = EyePosition.Y - LastY
    DifZ = EyePosition.Z - LastZ
    For k = 1 To 30
        With Verts(k).Pos
            .X = .X + DifX
            .Y = .Y + DifY
            .Z = .Z + DifZ
        End With
    Next k
    LastX = EyePosition.X
    LastY = EyePosition.Y
    LastZ = EyePosition.Z
    ReloadVB
End Function

Private Function ReloadVB()
    D3DVertexBuffer8SetData mVB, 0, VertexSizeInBytes * 30, 0, Verts(1)
End Function
