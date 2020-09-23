Attribute VB_Name = "Object"
Option Explicit

'Just create a simple plane

Public Sub InitializeMesh()
 With Mesh.Vertices(0): .X = 10:  .Y = 10:  .Z = 0.5: .RHW = 1: .Color = XRGB(&H0, &HFF, &HFF): .Color1 = XRGB(&HFF, &HFF, &HFF): .tU = 0: .tV = 0: .tU1 = 0: .tV1 = 0: End With
 With Mesh.Vertices(1): .X = 395: .Y = 10:  .Z = 0.5: .RHW = 1: .Color = XRGB(&HFF, &HFF, &H0): .Color1 = XRGB(&HFF, &HFF, &HFF): .tU = 1: .tV = 0: .tU1 = 1: .tV1 = 0: End With
 With Mesh.Vertices(2): .X = 395: .Y = 320: .Z = 0.5: .RHW = 1: .Color = XRGB(&HFF, &H0, &H0):  .Color1 = XRGB(&H0, &H0, &H0):    .tU = 1: .tV = 1: .tU1 = 1: .tV1 = 1: End With
 With Mesh.Vertices(3): .X = 10:  .Y = 320: .Z = 0.5: .RHW = 1: .Color = XRGB(&H0, &H0, &HFF):  .Color1 = XRGB(&H0, &H0, &H0):    .tU = 0: .tV = 1: .tU1 = 0: .tV1 = 1: End With
End Sub

Public Sub PositionMesh(PositionX As Single, PositionY As Single)
 Dim tmpVertices(3) As ShaderVertex__1
 Dim SizeOfVertex As Long
 Dim i As Long
 
 SizeOfVertex = Len(Mesh.Vertices(0))
 Call D3DVertexBuffer8GetData(DX_VB, 0, SizeOfVertex * 4, 0, tmpVertices(0))
 For i = 0 To 3
  tmpVertices(i) = Mesh.Vertices(i)
  tmpVertices(i).X = Mesh.Vertices(i).X + PositionX
  tmpVertices(i).Y = Mesh.Vertices(i).Y + PositionY
 Next
 Call D3DVertexBuffer8SetData(DX_VB, 0, SizeOfVertex * 4, 0, tmpVertices(0))
    
End Sub

Private Function XRGB(R As Long, G As Long, B As Long) As Long
 XRGB = B
 XRGB = XRGB Or (G * (2 ^ 8))
 XRGB = XRGB Or (R * (2 ^ 16))
End Function
