Attribute VB_Name = "mRender"
Option Explicit

'This is where the actual rendering gets done!

Public Sub Render()
    Dim vertexSize As Single
    Dim loopval As Single
    
    Dim sizeVert As VERTEX
    Dim sizeNormVert As NORMALVERTEX
    Dim sizeTexVert As TEXVERTEX
    
    If currentApp = Normal Then
        vertexSize = Len(sizeVert)
    ElseIf currentApp = Lighting Then
        vertexSize = Len(sizeNormVert)
    ElseIf currentApp = Texturing Then
        vertexSize = Len(sizeTexVert)
    End If
    
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, RGB(225, 100, 100), 1#, 0
    D3DDevice.BeginScene
         
        
    If currentApp = Texturing Then
        D3DDevice.SetStreamSource 0, vertexBuffer, vertexSize
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 12
    
    ElseIf currentApp = Modelling Then
        
        If cubeModel.nMaterials = 0 Then
            cubeModel.modelMesh.DrawSubset 0
        
        Else
            For loopval = 0 To cubeModel.nMaterials
                D3DDevice.setTexture 0, cubeModel.MeshTextures(loopval)
                D3DDevice.SetMaterial cubeModel.MeshMaterials(loopval)
                cubeModel.modelMesh.DrawSubset loopval
         
            Next loopval
        
        End If
        
    Else
        D3DDevice.SetStreamSource 0, vertexBuffer, vertexSize
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 8
    End If
    
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        
End Sub
