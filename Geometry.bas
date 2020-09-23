Attribute VB_Name = "mGeometry"
Option Explicit

'#Requires
'# Direct3DVertexBuffer8
'# That initD3D has been run

'# Constant for untransformed and lit vertices - XYZ and DIFFUSE
Public Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

'# Type for lit vertex
Type VERTEX
  X As Single
  Y As Single
  Z As Single
  Colour As Long
End Type

'-------------------------
'#initGeometry

'# Creates 24 vertices (8 triangles) in an octahedral
'# shape.  The vertices are also coloured.
'# Then loads the vertices into a Vertex Buffer

'#initMatrix

'# Initialises the view and projection matrices.
'# Sets the viewport (eye) to look at (0,0,0) with up set
'# as (0,1,0).  The projection is set to a ratio of pi/4
'# with a range of 1 to 10 metres.

'#createVertex
'# Creates a lit vertex

'#D3DVec
'# Creates a D3DVECTOR
'-------------------------

'Now, here's where we make up our luvly little shapes, using equally nice
'vectors.  Vectors are great wee things that tell us where we are in
'3D space, using an X, Y and Z value. A point in 3D space (like the corner
'of a cube) is known as a vertex.

'Eg.
'Using the normal type co-ordinate system (X across the screen,
'Y up the screen, Z into the screen):

'verts(0) = createVertex(-1, 0, -1, vbBlue) means go -1 in the X-axis,
'0 in the Y-axis, and -1 in the Z-axis, then make the vertex there.
'And make it blue.

'As we are making an octahedron (think of 2 pyramids
'stuck base to base) here we need to make 24 verticies. This works out as
'8 triangles - 3 verticies per triangle.

'For a cube, we would need 36 verticies
'(think about it - cube has 6 faces, but each face is made up of 2 triangles.
'so that makes 12 triangles = 36 verticies). It really helps if you work out
'the points that you need beforehand, using pen and paper...


'Right. That's the basic theory explained...now the VB stuff.
'In the declarations section you'll notice a type (FVF_VERTEX) and
'a constant (VERTEX).  These describe what type of vertex that you're
'gonna be using, using the Flexible Vertex Format system in DirectX.
'This lovely system means that you can mix and match your vertex types.

'In this module we only need a basic vertex type,
'the X,Y and Z of each vertex, and to make it look nice, the colour.
'For the X,Y and Z we need to add D3DFVF_XYZ to the constant, and to
'the type:

'Type VERTEX
'  X As Single \
'  Y As Single  = D3DFVF_XYZ
'  Z As Single /
'End Type

'For the colour we add D3DFVF_DIFFUSE to the constant, making the type:

'Type VERTEX
'  X As Single \
'  Y As Single  = D3DFVF_XYZ
'  Z As Single /
'  Colour As Long } D3DFVF_DIFFUSE
'End Type

'Looking at this way can make it a bit easier to understand. You tell
'DirectX what qualities you wish your vertex to have, and build up the
'UDT from that.
'Below is a list of some of the FVF types you are liely to use, and
'a wee bit about what they do..

'Type EXAMPLEVERTEX
'   X as Single   \
'   Y as Single    | D3DFVF_XYZ - the x,y,z of each vertex. Fairly essential.
'   Z as Single   /
'   rhw as Single/ D3DFVF_XYZRHW - for doing 2D graphics. You cant do any matrix stuff with this one.
'   nX as Single \
'   nY as Single  = D3DFVF_NORMAL - for lighting use, specifies the normal of each vertex
'   nZ as Single /
'   tu as Single \ D3DFVF_TEX1 - for basic texturing (texture coordinates use u,v,w instead of x,y,z)
'   tv as single /
'   Colour as Long = D3DFVF_DIFFUSE - not too tough
'   Specular as Long = D3DFVF_SPECULAR - for making things look shiny...
'End Type

'NOTE: You cannot use the D3DFVF_XYZRHW flag with the XYZ and NORMAL flags

'These should allow you to do most of the basic things.
'When making your vertex type, cut out the bits you dont really need
'(you won't need normals if you are having making unlit textured cube etc.)
'For a full ist of the FVF things lookit up in the DirectX 8 SDK...



Sub initGeometry()
Dim verts(23) As VERTEX
Dim rt2 As Single


    rt2 = Sqr(2)

    'Making an octahedron with coloured corners.
    'I've separated them into each triangle (8 in all) - which makes it easier
    ' to understand and debug.

    verts(0) = createVertex(-1, 0, -1, vbBlue)
    verts(1) = createVertex(0, rt2, 0, vbCyan)
    verts(2) = createVertex(1, 0, -1, vbRed)

    verts(3) = createVertex(-1, 0, 1, vbRed)
    verts(4) = createVertex(0, rt2, 0, vbCyan)
    verts(5) = createVertex(1, 0, 1, vbBlue)

    verts(6) = createVertex(1, 0, 1, vbBlue)
    verts(7) = createVertex(0, rt2, 0, vbCyan)
    verts(8) = createVertex(1, 0, -1, vbRed)

    verts(9) = createVertex(-1, 0, 1, vbRed)
    verts(10) = createVertex(0, rt2, 0, vbCyan)
    verts(11) = createVertex(-1, 0, -1, vbBlue)

    verts(12) = createVertex(-1, 0, -1, vbBlue)
    verts(13) = createVertex(0, -rt2, 0, vbCyan)
    verts(14) = createVertex(1, 0, -1, vbRed)

    verts(15) = createVertex(-1, 0, 1, vbRed)
    verts(16) = createVertex(0, -rt2, 0, vbCyan)
    verts(17) = createVertex(1, 0, 1, vbBlue)

    verts(18) = createVertex(1, 0, 1, vbBlue)
    verts(19) = createVertex(0, -rt2, 0, vbCyan)
    verts(20) = createVertex(1, 0, -1, vbRed)

    verts(21) = createVertex(-1, 0, 1, vbRed)
    verts(22) = createVertex(0, -rt2, 0, vbCyan)
    verts(23) = createVertex(-1, 0, -1, vbBlue)
    
    
    
    'Set DirectX to use out vertex type that we've made
    D3DDevice.SetVertexShader FVF_VERTEX
    
    'Enable lighting, the z-buffer and set the cullmode (see lighting) to none.
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    
    'Now we come to the vertex buffer.
    'This is an area set aside especially for storing verticies, which
    'makes it generally easier to deal with all those verticies.
    'The verticies are first loaded into an array of the corresponding vertex type.
    'You then simply load them into the vertex buffer (declared in objects)
    'by stating the size of them (the size of one multiplied by the number there are)
    'and setting the FVF type. The usage and pool should be set to 0 and D3DPOOL_DEFAULT
    'most of the time, you will rarely need to change them.  Once they have been
    'loaded into the vertexbuffer you inform DirectX that it exists and fill in the
    'parameters. Simple.
    

    
    Set vertexBuffer = D3DDevice.CreateVertexBuffer(Len(verts(0)) * 24, _
                                                    0, _
                                                    FVF_VERTEX, _
                                                    D3DPOOL_DEFAULT)
        
    D3DVertexBuffer8SetData vertexBuffer, 0, Len(verts(0)) * 24, 0, verts(0)
   

End Sub

Sub initMatrix()
    Dim matView As D3DMATRIX
    Dim matProj As D3DMATRIX
    
    'Right, now the matrix stuff.
    'Matricies are used to set up the scene so that it actually looks
    '3D, and to do stuff to our nice shapes, like rotating and
    'scaling them.
    
    'These two matrix functions are the initialisation ones.
    'You wont really need to change them at all during the run of your program
    '(unless u wanna do crazy stuff - zooming etc.)
        
    'The D3DTS_VIEW matrix can best be described as the 'eye' or 'camera'
    'The D3DXMatrixLookAtLH function requires that you specify where the
    'camera is, where to look at, and which way up is. This stuff is loaded
    'into a matrix passed into the function (matView). Once it has been
    'created, we get DirectX to use it with the SetTransform function.
    'The main transform types you will use are the view, projection and
    'world ones.
        
    D3DXMatrixLookAtLH matView, _
        D3DVec(1#, 2#, -5.5), _
        D3DVec(0#, 0#, 0#), _
        D3DVec(0#, 1#, 0#)
    D3DDevice.SetTransform D3DTS_VIEW, matView
    
    'The projection matrix is how the scene is projected. Well, obviously -
    'it specifies things like the viewing angle and perspective.
    
    D3DXMatrixPerspectiveFovLH matProj, _
        PI / 4, _
        1, _
        0.1, 100
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj

End Sub

Sub moveCamera(location As D3DVECTOR)
    Dim matMove As D3DMATRIX

    D3DXMatrixLookAtLH matMove, _
        D3DVec(location.X, location.Y, location.Z), _
        D3DVec(0#, 0#, 0#), _
        D3DVec(0#, 1#, 0#)
    D3DDevice.SetTransform D3DTS_VIEW, matMove
    
End Sub


Private Function createVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal Colour As Long) As VERTEX
    
    'A function that makes it easier to create all the verticies.
    'It loads all the given into the vertex...

    With createVertex
        .X = X
        .Y = Y
        .Z = Z
        .Colour = Colour
    End With

End Function

Function D3DVec(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
    
    'Create a D3DVECTOR
    With D3DVec
        .X = X
        .Y = Y
        .Z = Z
    End With

End Function
