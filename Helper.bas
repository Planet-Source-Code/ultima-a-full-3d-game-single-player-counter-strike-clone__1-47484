Attribute VB_Name = "Helper"
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Const PI = 3.141

Public Type XMesh
    Mesh As D3DXMesh
    NumMaterials As Long
    MeshTextures() As Direct3DTexture8
    MeshMaterials() As D3DMATERIAL8
    Transformation As D3DMATRIX
End Type

Public Function MakeFVertex(X As Single, Y As Single, Z As Single, U As Single, V As Single, Col As Long) As FVERTEX
    With MakeFVertex
        .Pos.X = X
        .Pos.Y = Y
        .Pos.Z = Z
        .tu = U
        .TV = V
        .Col = Col
    End With
End Function

Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    With MakeVector
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function

Public Function Min(X As Single, Y As Single) As Single
    If X < Y Then
        Min = X
    Else
        Min = Y
    End If
End Function

Public Function Max(X As Single, Y As Single) As Single
    If X > Y Then
        Max = X
    Else
        Max = Y
    End If
End Function

Public Function CrossProduct(V1 As D3DVECTOR, V2 As D3DVECTOR) As D3DVECTOR
    With CrossProduct
        .X = V1.Y * V2.Z - V1.Z * V2.Y
        .Y = V1.Z * V2.X - V1.X * V2.Z
        .Z = V1.X * V2.Y - V1.Y * V2.X
    End With
End Function

Public Function DotProduct(V1 As D3DVECTOR, V2 As D3DVECTOR) As Single
    DotProduct = V1.X * V2.X + V1.Y * V2.Y + V1.Z * V2.Z
End Function

Public Function Dist(V1 As D3DVECTOR, V2 As D3DVECTOR) As Single
    Dist = Sqr((V2.X - V1.X) * (V2.X - V1.X) + (V2.Y - V1.Y) * (V2.Y - V1.Y) + (V2.Z - V1.Z) * (V2.Z - V1.Z))
End Function

Public Function Length(V1 As D3DVECTOR) As Single
    Length = Sqr(V1.X * V1.X + V1.Y * V1.Y + V1.Z * V1.Z)
End Function

Public Function Acos(XX As Single) As Single
    If Abs(XX) < 1 Then
        Acos = Atn(-XX / Sqr(-XX * XX + 1)) + 1.5707963267949
    ElseIf XX = 1 Then
        Acos = 0
    ElseIf XX = -1 Then
        Acos = PI
    End If
End Function

Public Function Asin(XX As Single) As Single
    If Abs(XX) < 1 Then
        Asin = Atn(XX / Sqr(-XX * XX + 1))
    ElseIf XX = 1 Then
        Asin = 1.5707963267948
    ElseIf XX = -1 Then
        Asin = 4.7123889803846
    End If
End Function

Public Function ARGB2LONG(A As Long, R As Long, G As Long, B As Long) As Long
    If A > 127 Then
        ARGB2LONG = &H80000000 Or (A - 128) * 16777216
        ARGB2LONG = ARGB2LONG Or R * 65536
        ARGB2LONG = ARGB2LONG Or G * 256
        ARGB2LONG = ARGB2LONG Or B
    Else
        ARGB2LONG = A * 16777216
        ARGB2LONG = ARGB2LONG Or R * 65536
        ARGB2LONG = ARGB2LONG Or G * 256
        ARGB2LONG = ARGB2LONG Or B
    End If
End Function
