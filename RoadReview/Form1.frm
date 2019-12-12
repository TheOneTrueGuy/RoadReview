VERSION 5.00
Begin VB.Form RoadView 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   5925
   ClientTop       =   4920
   ClientWidth     =   7830
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7830
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   7365
      Top             =   2685
   End
   Begin VB.CommandButton Command1 
      Height          =   435
      Left            =   7350
      TabIndex        =   1
      Top             =   315
      Width           =   450
   End
   Begin VB.PictureBox Picture1 
      Height          =   4710
      Left            =   480
      ScaleHeight     =   4650
      ScaleWidth      =   6645
      TabIndex        =   0
      Top             =   300
      Width           =   6705
   End
End
Attribute VB_Name = "RoadView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim g_DX As New DirectX8
Dim g_D3D As Direct3D8              'Used to create the D3DDevice
Dim g_D3DDevice As Direct3DDevice8  'Our rendering device
Dim g_VB As Direct3DVertexBuffer8


' A structure for our custom vertex type
' representing a point on the screen
Private Type CUSTOMVERTEX
    x As Double         'x in screen space
    y As Double         'y in screen space
    z  As Double        'normalized z
    color As Long       'vertex color
End Type

' Our custom FVF, which describes our custom vertex structure
Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)
Const g_pi = 3.1415
Dim vertcount As Integer
Dim Vertices() As CUSTOMVERTEX
Dim datar() As Double
Private Sub Command1_Click()

datar = Data_Entry2.getDatar
Debug.Print "datar(1, 1):" & datar(1, 1)
render
End Sub

Private Sub Form_Activate()
' reload grid data from entry form

End Sub

'-----------------------------------------------------------------------------
' Name: Form_Load()
'-----------------------------------------------------------------------------
Private Sub Form_Load()
    Dim b As Boolean
    
    ' Allow the form to become visible
    Me.Show
    DoEvents
    
    ' Initialize D3D and D3DDevice
    b = InitD3D(Picture1.hWnd)
    If Not b Then
        MsgBox "Unable to CreateDevice (see InitD3D() source for comments)"
        End
    End If
    
    
    ' Initialize Vertex Buffer with Geometry
    b = InitGeometry()
    If Not b Then
        MsgBox "Unable to Create VertexBuffer"
        End
    End If
    
    
    ' Enable Timer to update
'    Timer1.Enabled = True
    
End Sub

'-----------------------------------------------------------------------------
' Name: Timer1_Timer()
'-----------------------------------------------------------------------------
Private Sub Timer1_Timer()
'Static countit As Integer
'countit = countit + 1
'If countit = 2 Then Timer1.Enabled = False

    render
End Sub

'-----------------------------------------------------------------------------
' Name: Form_Unload()
'-----------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
   Cleanup
End Sub


'-----------------------------------------------------------------------------
' Name: InitD3D()
' Desc: Initializes Direct3D
'-----------------------------------------------------------------------------
Private Function InitD3D(hWnd As Long) As Boolean

  
  
  
  '  On Local Error Resume Next
'''======= these lines from livetime

'Dim VertexSizeInBytes As Long
'dx3ddev.SetVertexShader (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE)
'For tzl = 0 To 999
'With Vertices(tzl)
'.z = 0
'.rhw = 1
'.color = &HFFFFFF00
'End With
'Next tzl


'    ' Create the D3D object
'    Set dx3d = dx.Direct3DCreate()
'    If dx3d Is Nothing Then Exit Function
'
'    ' Get the current display mode
'    Dim mode As D3DDISPLAYMODE
'    dx3d.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode
'
'    ' Fill in the type structure used to create the device
'    Dim d3dpp As D3DPRESENT_PARAMETERS
'    d3dpp.Windowed = True
'    d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
'    d3dpp.BackBufferFormat = mode.Format
'
'    ' Create the D3DDevice
'    ' If you do not have hardware 3d acceleration. Enable the reference rasterizer
'    ' using the DirectX control panel and change D3DDEVTYPE_HAL to D3DDEVTYPE_REF
'
'    Set dx3ddev = dx3d.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, _
'                                      D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
'    If dx3ddev Is Nothing Then Exit Function
'
'    ' Device state would normally be set here
'    ' Turn off culling, so we see the front and back of the triangle
'dx3ddev.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
'dx3ddev.SetRenderState D3DRS_LINEPATTERN, 0
'    ' Turn off D3D lighting, since we are providing our own vertex colors
'    dx3ddev.SetRenderState D3DRS_LIGHTING, 0
'
'    InitD3D = True



'    On Local Error Resume Next
    
    ' Create the D3D object
    Set g_D3D = g_DX.Direct3DCreate()
    If g_D3D Is Nothing Then Exit Function
    
    ' Get the current display mode
    Dim mode As D3DDISPLAYMODE
    g_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode
        
    ' Fill in the type structure used to create the device
    Dim d3dpp As D3DPRESENT_PARAMETERS
    d3dpp.Windowed = 1
    d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    d3dpp.BackBufferFormat = mode.Format
    
    ' Create the D3DDevice
    ' If you do not have hardware 3d acceleration. Enable the reference rasterizer
    ' using the DirectX control panel and change D3DDEVTYPE_HAL to D3DDEVTYPE_REF
    
    Set g_D3DDevice = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, _
                                      D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    If g_D3DDevice Is Nothing Then Exit Function
    
    ' Device state would normally be set here
    ' Turn off culling, so we see the front and back of the triangle
    g_D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    ' Turn off D3D lighting, since we are providing our own vertex colors
    g_D3DDevice.SetRenderState D3DRS_LIGHTING, 0

    InitD3D = True
End Function


'-----------------------------------------------------------------------------
' Name: SetupMatrices()
' Desc: Sets up the world, view, and projection transform matrices.
'-----------------------------------------------------------------------------
Private Sub SetupMatrices()

    
    ' The transform Matrix is used to position and orient the objects
    ' you are drawing
    ' For our world matrix, we will just rotate the object about the y-axis.
    Dim matWorld As D3DMATRIX
    D3DXMatrixRotationY matWorld, Timer * 4
    g_D3DDevice.SetTransform D3DTS_WORLD, matWorld


    ' The view matrix defines the position and orientation of the camera
    ' Set up our view matrix. A view matrix can be defined given an eye point,
    ' a point to lookat, and a direction for which way is up. Here, we set the
    ' eye five units back along the z-axis and up three units, look at the
    ' origin, and define "up" to be in the y-direction.
    
    
    Dim matView As D3DMATRIX
    D3DXMatrixLookAtLH matView, vec3(0#, 3#, -17#), _
                                 vec3(0#, 0#, -2#), _
                                 vec3(0#, 1#, 0#)
                                 
    g_D3DDevice.SetTransform D3DTS_VIEW, matView

    ' The projection matrix describes the camera's lenses
    ' For the projection matrix, we set up a perspective transform (which
    ' transforms geometry from 3D view space to 2D viewport space, with
    ' a perspective divide making objects smaller in the distance). To build
    ' a perpsective transform, we need the field of view (1/4 pi is common),
    ' the aspect ratio, and the near and far clipping planes (which define at
    ' what distances geometry should be no longer be rendered).
    Dim matProj As D3DMATRIX
    D3DXMatrixPerspectiveFovLH matProj, g_pi / 3, 1, 1, 1000
    g_D3DDevice.SetTransform D3DTS_PROJECTION, matProj

End Sub



'-----------------------------------------------------------------------------
' Name: InitGeometry()
' Desc: Creates a vertex buffer and fills it with our vertices.
'-----------------------------------------------------------------------------
Function InitGeometry() As Boolean

    ' Initialize three vertices for rendering a triangle
   ReDim Vertices(6) 'As CUSTOMVERTEX
    Dim VertexSizeInBytes As Long
    
    VertexSizeInBytes = Len(Vertices(0))
    
    With Vertices(0): .x = -1: .y = -1: .z = 0: .color = &HFFFF0000: End With
    With Vertices(1): .x = 1: .y = -1: .z = 0:  .color = &HFF00FF00: End With
    With Vertices(2): .x = 0: .y = 1: .z = 0:  .color = &HFF00FFFF: End With
    With Vertices(3): .x = 0: .y = 1: .z = 0:  .color = &HFFAFFFFF: End With
    With Vertices(4): .x = 5: .y = 2: .z = 0:  .color = &H0: End With
    With Vertices(5): .x = 0: .y = 7: .z = 3:  .color = &HDDFFFF00: End With

'Dim Vertices(5) As CUSTOMVERTEX
'With Vertices(0): .x = -5#: .y = -5#: .z = 0#: .color = &HFFFF0000: End With
'With Vertices(1): .x = 0#: .y = 5#: .z = 0#: .color = &HFFFF00FF: End With
'With Vertices(2): .x = 5#: .y = -5#: .z = 0#: .color = &HFFFFFF00: End With
'With Vertices(3): .x = 10#: .y = 5#: .z = 0#: .color = &HFFFF0000: End With
'With Vertices(4): .x = 15#: .y = -5#: .z = 0#: .color = &HFFFF00FF: End With
'With Vertices(5): .x = 20#: .y = 5#: .z = 0#: .color = &HFFFF0000: End With


    ' Create the vertex buffer.
    Set g_VB = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 6, _
                     0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VB Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VB, 0, VertexSizeInBytes * 6, 0, Vertices(0)

    InitGeometry = True
End Function



'-----------------------------------------------------------------------------
' Name: Cleanup()
' Desc: Releases all previously initialized objects
'-----------------------------------------------------------------------------
Sub Cleanup()
    Set g_VB = Nothing
    Set g_D3DDevice = Nothing
    Set g_D3D = Nothing
End Sub

'-----------------------------------------------------------------------------
' Name: Render()
' Desc: Draws the scene
'-----------------------------------------------------------------------------
'Sub Render()
'
'    Dim v As CUSTOMVERTEX
'    Dim sizeOfVertex As Long
'
'
'    If g_D3DDevice Is Nothing Then Exit Sub
'
'    ' Clear the backbuffer to a blue color (ARGB = 000000ff)
'    '
'    ' To clear the entire back buffer we send down
'    g_D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, &HFF&, 1#, 0
'
'
'    ' Begin the scene
'    g_D3DDevice.BeginScene
'
'
'    ' Setup the world, view, and projection matrices
'    SetupMatrices
'
'    'Draw the triangles in the vertex buffer
'    sizeOfVertex = Len(v)
'    g_D3DDevice.SetStreamSource 0, g_VB, sizeOfVertex
'    g_D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
'    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 4
'
'
'    ' End the scene
'    g_D3DDevice.EndScene
'
'
'    ' Present the backbuffer contents to the front buffer (screen)
'    g_D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
'
'End Sub
Public Sub render()
'' transfer bytebuffer into vertex array and call drawprimitive
  Dim v As CUSTOMVERTEX
    Dim sizeOfVertex As Long
'
'
    If dx3ddev Is Nothing Then Exit Sub
'
'    ' Clear the backbuffer to a blue color (ARGB = 000000ff)
'    '
'    ' To clear the entire back buffer we send down
  dx3ddev.Clear 0, ByVal 0, D3DCLEAR_TARGET, &HFF&, 1#, 0
'
'
'    ' Begin the scene
dx3ddev.BeginScene
sizeOfVertex = Len(v)
'dx3ddev.SetStreamSource 0, g_VB, sizeOfVertex
'dx3ddev.SetVertexShader D3DFVF_CUSTOMVERTEX
'dx3ddev.DrawPrimitive D3DPT_LINESTRIP, 0, samplesize - 1
dx3ddev.DrawPrimitiveUP D3DPT_LINESTRIP, samplesize, Vertices(0), sizeOfVertex

'D3DPT_LINESTRIP
'
'    ' End the scene
dx3ddev.EndScene
'
'
'    ' Present the backbuffer contents to the front buffer (screen)
dx3ddev.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

'-----------------------------------------------------------------------------
' Name: vec3()
' Desc: helper function
'-----------------------------------------------------------------------------
Function vec3(x As Double, y As Double, z As Double) As D3DVECTOR
    vec3.x = x
    vec3.y = y
    vec3.z = z
End Function
Private Sub loadRoadseg(centerpoints() As Poynt, WL As Double, WR As Double)
' load vertices from 2 four sided polygons or 4 triangles
' calculate distance for X and Z, extract elev
' for the entire set of elevations

End Sub



Public Sub loadVertice(lat As Double, longi As Double, elev As Double)
ReDim Preserve Vertices(vertcount)
Vertices(vertcount).color = RGB(255, 255, 255)
'vertices(vertcount).rhw = 1
Vertices(vertcount).x = longi
Vertices(vertcount).y = elev
Vertices(vertcount).z = lat

vertcount = vertcount + 1
End Sub


Public Sub vertLoader()
'' was in a timer event
'If timing Then Exit Sub
'timing = True
'Dim tzl
'
''' convert bytebuffer into picture1 lines
'dscb.GetCurrentPosition capCURS
'Dim countit As Integer
'dscb.ReadBuffer capCURS.lWrite - samplesize, samplesize, ByteBuffer(0), DSCBLOCK_DEFAULT
'
''####################
'' fill vertices() with  samples
'countit = 0
' For tzl = 0 To samplesize - 1
' countit = countit + 1
' With Vertices(countit)
'  .x = (tzl): .y = (((ByteBuffer(tzl) * scaleamp) / 680) + (halfhyt / scaleamp))
'
''   Debug.Print ByteBuffer(tzl); 'vertices(tzl).y
'blockbuffer(tzl + (UBound(blockbuffer) - 2000)) = ByteBuffer(tzl)
'
' End With
' Next tzl
'
'If UBound(blockbuffer) < 1000000 Then
'    ReDim Preserve blockbuffer(UBound(ByteBuffer) + UBound(blockbuffer))
'Else
'    ReDim blockbuffer(2000)
'End If
'
'Render
''If dscb.GetStatus = 0 Then Stop
''Debug.Print ByteBuffer(10), ByteBuffer(22), dscb.GetStatus, pos
'timing = False
End Sub
