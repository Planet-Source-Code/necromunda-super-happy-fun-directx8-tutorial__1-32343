Attribute VB_Name = "mModels"
Option Explicit

Type modelType
    modelMesh As D3DXMesh
    nMaterials As Long
    MeshMaterials() As D3DMATERIAL8
    MeshTextures() As Direct3DTexture8
End Type

Public currentModel As String
Public cubeModel As modelType

Sub loadModel(modelName As String)
    Dim mtrlBuffer As D3DXBuffer
    Dim textureLoc As String
    Dim textureName As String
    Dim loopval As Long
    
    On Local Error GoTo report

    If modelName = "default" Then
        Set cubeModel.modelMesh = D3DX.CreateTeapot(D3DDevice, mtrlBuffer)
        Exit Sub
    End If
    
    textureLoc = Left(modelName, InStrRev(modelName, "\"))

    Set cubeModel.modelMesh = D3DX.LoadMeshFromX(modelName, D3DXMESH_MANAGED, _
                                  D3DDevice, Nothing, mtrlBuffer, cubeModel.nMaterials)
    currentModel = modelName

    ReDim cubeModel.MeshMaterials(cubeModel.nMaterials)
    ReDim cubeModel.MeshTextures(cubeModel.nMaterials)
    
    For loopval = 0 To cubeModel.nMaterials - 1
        
        D3DX.BufferGetMaterial mtrlBuffer, loopval, cubeModel.MeshMaterials(loopval)
        textureName = D3DX.BufferGetTextureName(mtrlBuffer, loopval)
            
        cubeModel.MeshMaterials(loopval).Ambient = cubeModel.MeshMaterials(loopval).diffuse
    
        If textureName <> "" Then
            Set cubeModel.MeshTextures(loopval) = _
            D3DX.CreateTextureFromFileEx(D3DDevice, textureLoc & textureName, _
                                         256, 256, D3DX_DEFAULT, 0, _
                                         D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                         D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, _
                                         0, ByVal 0, ByVal 0)
        End If
    
    Next loopval

Exit Sub
report:

    If Err.Number <> D3D_OK Then
        MsgBox "Error loading model file! Check that all files are present", vbExclamation + vbOKOnly, "Models"
        Exit Sub
    End If
    

End Sub

Sub SetupModelLights()
Dim Material As D3DMATERIAL8
    
    Material.Ambient = loadColours(1, 1, 1, 1)
    Material.diffuse = loadColours(1, 1, 1, 1)
    
    D3DDevice.SetMaterial Material
    
    With Light
        .Type = D3DLIGHT_POINT
        
        .position = D3DVec(0, 3, -2)
        .Direction = D3DVec(0, 0, 0)
                    
        .diffuse = loadColours(1, 2, 0, 0)
                
        .Range = 20
        .Attenuation1 = 0.2   'Set the Linear Attenuation to 0.05
    End With

    D3DDevice.SetLight 0, Light
    
    With Light
        .Type = D3DLIGHT_POINT
        
        .position = D3DVec(0, -2, -3)
        .Direction = D3DVec(0, 0, 0)
        
        .diffuse = loadColours(1, 0, 2, 0)
        
        .Range = 20

        .Attenuation1 = 0.5   'Set the Linear Attenuation to 0.05
    End With
    
    D3DDevice.SetLight 1, Light
    
    D3DDevice.LightEnable 0, 1
    D3DDevice.LightEnable 1, 1
    


End Sub

Private Function loadColours(ByVal a As Single, ByVal r As Single, ByVal g As Single, ByVal b As Single) As D3DCOLORVALUE
    
    With loadColours
        .a = a
        .r = r
        .g = g
        .b = b
    End With

End Function

