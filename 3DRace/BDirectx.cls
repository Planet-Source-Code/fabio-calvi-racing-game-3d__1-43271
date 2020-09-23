VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BDirectx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event DirecXNotInstalled()
Public Event PostRender()
Public Event Error4(Errstr As String)
Public Event Attemt4(Attemtstr As String)

Private Type DeviceCharacteristics
    bDither As Long
    Name As String
    Quality As Long
    RenderMode As Long
    Shades As Long
    TexQ As Long
    WFoptions As Long
End Type

Private Type ViewportCharacteristics
    Back As Single
    Field As Single
    Front As Single
    Name As String
    Projection As Long
    left As Single
    right As Single
    bottom As Single
    top As Single
    scaling As Long
End Type

'- direct x object
Dim m_dx As New DirectX7

'- direct draw objects
Dim m_dd As DirectDraw4
Dim m_ddClip As DirectDrawClipper
Dim m_frontBuffer As DirectDrawSurface4
Dim m_backBuffer As DirectDrawSurface4

'- direct 3drm objects
Dim m_rm As Direct3DRM3
Dim m_rmDevice As Direct3DRMDevice3
Dim m_rmViewport As Direct3DRMViewport2
Dim m_rmFrameScene As Direct3DRMFrame3
Dim m_rmFrameCamera As Direct3DRMFrame3
Dim m_rmFrameDirLight As Direct3DRMFrame3
Dim m_rmFrameAmbientLight As Direct3DRMFrame3
Dim m_rmDirLight As Direct3DRMLight
Dim m_rmAmbientLight As Direct3DRMLight


'- state
Dim m_strDDGuid As String               'DirectDraw device guid
Dim m_strD3DGuid As String              'Direct3DRM device guid
Dim m_hwnd As Long                      'hwnd (either FSWindow or our ocx)
Dim m_binit As Boolean                  'are we initailized?
Dim m_lastFPS As Long                   'time stamp of last FPS update
Dim m_fps As Single                     'frame per second
Dim m_bCreateFromClipper As Boolean     'Use a clipper to start the RM
Dim m_DevInfo As DeviceCharacteristics
Dim m_ViewInfo As ViewportCharacteristics
Dim m_bfullscreen As Boolean
Dim m_bUseSoftwareOnly As Boolean
Dim Attemt As String
Dim m_emptyrect As RECT
Dim wndr As RECT


Dim NoEdgeW As DDCOLORKEY
Dim hud As DirectDrawSurface4
Dim Stretching As DDSURFACEDESC2
Dim hudWidth
Dim hudHeight
Dim hudAttributes As RECT
Dim hudX As Integer
Dim hudY As Integer

'Directsoundinner
Public DSOUND70 As DirectSound
Public DsoundPri70 As DirectSoundBuffer
Public DsoundLis70 As DirectSound3DListener
Private Property Let Attemtstr(Str As String)
    Attemt = Str
    RaiseEvent Attemt4(Attemt)
End Property


'Direct sound Functions
Public Sub InitDsound70()
    On Local Error GoTo errout
    
    Attemtstr = "Creating Directsound"
    
    Set DSOUND70 = m_dx.DirectSoundCreate("")
    Attemtstr = "Setting DirectsoundCooperativeLevel"
    DSOUND70.SetCooperativeLevel m_hwnd, DSSCL_PRIORITY
    
    Attemtstr = "Creating primarySoundBuffer"
    Dim dsbd As DSBUFFERDESC
    Dim wfmt As WAVEFORMATEX
    dsbd.lFlags = DSBCAPS_PRIMARYBUFFER Or DSBCAPS_CTRL3D
    Set DsoundPri70 = DSOUND70.CreateSoundBuffer(dsbd, wfmt)
    
    Attemtstr = "Creating Directsoundlisner"
    Set DsoundLis70 = DsoundPri70.GetDirectSound3DListener()
    
    Exit Sub
errout:
    Err.Clear
    RaiseEvent Error4("Error in " + Attemt)
End Sub
Public Function Create2DsBuffromfile70(Sfile As String) As DirectSoundBuffer
    On Local Error GoTo errout
    
    Dim dsbd As DSBUFFERDESC
    Dim wfmt As WAVEFORMATEX
    
    Attemtstr = "Creating 2DSoundBufferFromfile " + Sfile
    dsbd.lFlags = DSBCAPS_CTRL3D Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    Set Create2DsBuffromfile70 = DSOUND70.CreateSoundBufferFromFile(Sfile, dsbd, wfmt)
    
    Exit Function
errout:
    Err.Clear
    RaiseEvent Error4("Error in " + Attemt)
End Function
Public Function Create3DSBUFfrom2Dbuf70(Buf As DirectSoundBuffer) As DirectSound3DBuffer
    On Local Error GoTo errout
    
    Attemtstr = "Creating 3DSound Buffer"
    Set Create3DSBUFfrom2Dbuf70 = Buf.GetDirectSound3DBuffer
    Exit Function
errout:
    Err.Clear
    RaiseEvent Error4("Error in " + Attemt)
End Function

'-============================================================
' StartWindowed
'-============================================================
Public Function StartWindowed() As Boolean
    Dim b As Boolean
                
    b = InitWindowed("", "IID_IDirect3DHALDevice")
    If b = True Then
        StartWindowed = True
        Exit Function
    End If
    
    b = InitWindowed("", "IID_IDirect3DRGBDevice")
    StartWindowed = b
        
End Function
'-============================================================
' InitWindowed
'-============================================================
Public Function InitWindowed(ddrawguid As String, d3dguid As String) As Boolean
    Dim b As Boolean
    Dim ddsd As DDSURFACEDESC2
    
    On Local Error GoTo errout
        
    m_binit = False
    
    Attemtstr = "make sure we have com out of fullscreen mode"
    If Not (m_dd Is Nothing) Then m_dd.RestoreDisplayMode
    If Not (m_dd Is Nothing) Then m_dd.SetCooperativeLevel 0, DDSCL_NORMAL

    Cleanup
    
    Attemtstr = "get rid of our current rm device.."
    Set m_rmDevice = Nothing
    Set m_rmViewport = Nothing
    
    m_strDDGuid = ddrawguid
    m_strD3DGuid = d3dguid
    
    If d3dguid = "" Then m_strD3DGuid = "IID_IDirect3DRGBDevice"
    If m_bUseSoftwareOnly = True Then m_strD3DGuid = "IID_IDirect3DRGBDevice"

    Attemtstr = "creating requested DirectDraw object from ddrawguid"
    Set m_dd = m_dx.DirectDraw4Create(m_strDDGuid)
    
    Attemtstr = "Setting The CooperativeLevel"
    m_dd.SetCooperativeLevel m_hwnd, DDSCL_NORMAL
    
    Attemtstr = "Creating screen surface from DirectDraw"
    ddsd.lFlags = DDSD_CAPS
    ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set m_frontBuffer = m_dd.CreateSurface(ddsd)
    
    Attemtstr = "setting the clipper"
    Set m_ddClip = m_dd.CreateClipper(0)
    m_ddClip.SetHWnd m_hwnd
    m_frontBuffer.SetClipper m_ddClip
    
    b = ResizeWindowedDevice(m_strD3DGuid)
    If b = False Then GoTo errout
    
    SetDeviceDefaults
    
    m_binit = True
    m_bfullscreen = False
    InitWindowed = True
    Exit Function
    
errout:
    Cleanup
    RaiseEvent Error4("Error in " + Attemt)
End Function

'-============================================================
' ResizeWindowedDevice
'-============================================================
Private Function ResizeWindowedDevice(d3dg As String) As Boolean
    If m_dd Is Nothing Then Exit Function
    
    If m_bfullscreen Then Exit Function
    
    On Local Error GoTo errout
    
    
    Dim memflags As Long
    Dim r As RECT
    Dim ddsd As DDSURFACEDESC2
    
    'Get window extent
    Call m_dx.GetWindowRect(m_hwnd, r)
    ddsd.lWidth = r.right - r.left
    ddsd.lHeight = r.bottom - r.top
    
    Set m_rmViewport = Nothing
    Set m_rmDevice = Nothing
    Set m_backBuffer = Nothing
    
    
    'Take care of createFromWindowed shortcut
    If m_bCreateFromClipper Then
        
        Attemtstr = "Creating RM Device or Viewport for current window size"
        Set m_rmDevice = m_rm.CreateDeviceFromClipper(m_ddClip, d3dg, ddsd.lWidth, ddsd.lHeight)
        Set m_rmViewport = m_rm.CreateViewport(m_rmDevice, m_rmFrameCamera, 0, 0, ddsd.lWidth, ddsd.lHeight)
        
        ResizeWindowedDevice = True
        Exit Function
    
    End If
    
    
    If UCase(d3dg) = "IID_IDIRECT3DHALDEVICE" Then
        memflags = DDSCAPS_VIDEOMEMORY
    Else
        memflags = DDSCAPS_SYSTEMMEMORY
    End If

    
    'CreateBacksurface
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE Or memflags
           
    Attemtstr = "creating backbuffer for current window size - try setting Use3DHardware=FALSE"
    Set m_backBuffer = m_dd.CreateSurface(ddsd)
    
    Attemtstr = "creating RM Device or Viewport for current window size"
    Set m_rmViewport = Nothing
    Set m_rmDevice = Nothing
    Set m_rmDevice = m_rm.CreateDeviceFromSurface(d3dg, m_dd, m_backBuffer, 0)
    Set m_rmViewport = m_rm.CreateViewport(m_rmDevice, m_rmFrameCamera, 0, 0, ddsd.lWidth, ddsd.lHeight)
    
        
    ResizeWindowedDevice = True
    Exit Function
    
errout:

    Err.Clear
    Set m_rmDevice = Nothing
    Set m_rmViewport = Nothing
    Set m_backBuffer = Nothing
    
    ResizeWindowedDevice = False
    m_binit = False
    RaiseEvent Error4("Error in " + Attemt)

End Function


'-============================================================
' InitFullScreen
'-============================================================
Public Function InitFullScreen(ddrawguid As String, d3dguid As String, w As Long, h As Long, bpp As Long, Optional TripleBuffering As Boolean) As Boolean
    
    On Local Error GoTo errout
    
    m_binit = False
    
    'get rid of our current rm device..
    Cleanup
    
    'make sure fs window is up
    
    DoEvents
    
    m_strDDGuid = ddrawguid
    If d3dguid = "" Then m_strD3DGuid = "IID_IDirect3DRGBDevice"
    
    'DirectDrawCreate
    Attemtstr = "Creating DirectDraw Object Create for given ddrawguid"
    Set m_dd = dX.DirectDraw4Create(m_strDDGuid)
    
    'Set cooperative level
    Attemtstr = "Setting CooperativeLevel for fullscreen operation"
    m_dd.SetCooperativeLevel m_hwnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWMODEX

    'set the display mode
    If w <> 0 And h <> 0 And bpp <> 0 Then
        Attemtstr = "setting full screen display mode at requested w h and bpp"
        m_dd.SetDisplayMode w, h, bpp, 0, DDSDM_DEFAULT
    End If

    Dim ddsd As DDSURFACEDESC2
    If TripleBuffering = False Then
        'create Flipping Surfaces - one front and 1 back buffer
        ddsd.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
        ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_VIDEOMEMORY Or DDSCAPS_3DDEVICE
        ddsd.lBackBufferCount = 1
        Attemtstr = "Creating filipable surface for fullscreen operation"
        Set m_frontBuffer = m_dd.CreateSurface(ddsd)
    Else
        'create Flipping Surfaces - one front and 2 back buffer
        ddsd.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
        ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_VIDEOMEMORY Or DDSCAPS_3DDEVICE
        ddsd.lBackBufferCount = 2
        Attemtstr = "Creating filipable surface for fullscreen operation Using TripleBuffering"
        Set m_frontBuffer = m_dd.CreateSurface(ddsd)
    End If
    'Setup a clipper
    Attemtstr = "settingup the clipper"
    Set m_ddClip = m_dd.CreateClipper(0)
    m_ddClip.SetHWnd m_hwnd
    m_frontBuffer.SetClipper m_ddClip
            
    'Get backbuffer
    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Attemtstr = "getting the fullscreen backbuffer"
    Set m_backBuffer = m_frontBuffer.GetAttachedSurface(caps)
    
    'get backbuffer description
    Dim ddsd2 As DDSURFACEDESC2
    Attemtstr = "getting the fullscreen backbuffer description"
    m_backBuffer.GetSurfaceDesc ddsd2
    
    
    'see if they turned hw off
    If m_bUseSoftwareOnly Then
        m_strD3DGuid = "IID_IDirect3DRGBDevice"
    Else
        m_strD3DGuid = d3dguid
    End If
    
    'create the rm device from surface
    Attemtstr = "Creating the Retained Mode device - try a smaller resolution or try setting  Use3DHardware=false"
    Set m_rmDevice = m_rm.CreateDeviceFromSurface(m_strD3DGuid, m_dd, m_backBuffer, D3DRMDEVICE_DEFAULT)
    Set m_rmViewport = m_rm.CreateViewport(m_rmDevice, m_rmFrameCamera, 0, 0, ddsd2.lWidth, ddsd2.lHeight)
    
    SetDeviceDefaults
    
    m_binit = True
    m_bfullscreen = True
    InitFullScreen = True
    Exit Function

errout:
    m_binit = False
    Cleanup
    RaiseEvent Error4("Error in " + Attemt)
End Function


'Run Time R/O access properties

'-============================================================
' Dx
'-============================================================
Public Function dX() As DirectX7
    Set dX = m_dx

End Function

'-============================================================
' DDraw
'-============================================================
Public Function DDraw() As DirectDraw4
    Set DDraw = m_dd
End Function

'-============================================================
' BackBuffer
'-============================================================
Public Function BackBuffer() As DirectDrawSurface4
    Set BackBuffer = m_backBuffer
End Function

'-============================================================
' D3drm
'-============================================================
Public Function D3DRM() As Direct3DRM3
    Set D3DRM = m_rm
End Function

'-============================================================
' Device
'-============================================================
Public Function Device() As Direct3DRMDevice3
    Set Device = m_rmDevice
End Function

'-============================================================
' Viewport
'-============================================================
Public Function Viewport() As Direct3DRMViewport2
    Set Viewport = m_rmViewport
End Function
 
'-============================================================
' DirLightFrame
'-============================================================
Public Function DirLightFrame() As Direct3DRMFrame3
    Set DirLightFrame = m_rmFrameDirLight
End Function
'-============================================================
Public Function SceneFrame() As Direct3DRMFrame3
    Set SceneFrame = m_rmFrameScene
End Function
 
'-============================================================
' CameraFrame
'-============================================================
Public Function CameraFrame() As Direct3DRMFrame3
    Set CameraFrame = m_rmFrameCamera
End Function
 
'-============================================================
' DirLight
'-============================================================
Public Function DirLight() As Direct3DRMLight
    Set DirLight = m_rmDirLight
End Function

'-============================================================
' AmbientLight
'-============================================================
Public Function AmbientLight() As Direct3DRMLight
    Set AmbientLight = m_rmAmbientLight
End Function
 
 
'-============================================================
' Use3DHardware
'-============================================================
Property Let Use3DHardware(b As Boolean)
    m_bUseSoftwareOnly = Not b
End Property

Property Get Use3DHardware() As Boolean
    Use3DHardware = Not m_bUseSoftwareOnly
End Property
'-============================================================
' UseBackBuffer
'-============================================================
Property Let UseBackBuffer(b As Boolean)
    m_bCreateFromClipper = Not b
End Property

Property Get UseBackBuffer() As Boolean
    UseBackBuffer = Not m_bCreateFromClipper
End Property
'-============================================================
' DirectDrawGuid
'-============================================================
Property Get DirectDrawGuid() As String
     DirectDrawGuid = m_strDDGuid
End Property
'-============================================================
' Direct3DGuid
'-============================================================
Property Get Direct3DGuid() As String
     Direct3DGuid = m_strD3DGuid
End Property

'- Runtime only List Functions
'
'-============================================================
' Devices
'-============================================================
Public Function Devices(Optional ddrawguid = "") As Direct3DEnumDevices
    On Local Error GoTo exitOut:
    
    Dim dd As DirectDraw7
    Dim d3d As Direct3D7
    
    Set dd = dX.DirectDrawCreate(CStr(ddrawguid))
        
    Set d3d = dd.GetDirect3D()
    
    Set Devices = d3d.GetDevicesEnum()
    
    Set dd = Nothing
    Set d3d = Nothing
    Exit Function
    
exitOut:

End Function

'-============================================================
' VideoCards
'-============================================================
Public Function VideoCards() As DirectDrawEnum
    Set VideoCards = m_dx.GetDDEnum()
End Function

'-============================================================
' DisplayModes
'-============================================================
Public Function DisplayModes(Optional ddrawguid = "") As DirectDrawEnumModes
    On Local Error GoTo exitOut
    Dim dd As DirectDraw4
    Set dd = dX.DirectDraw4Create(CStr(ddrawguid))
    Dim ddsd As DDSURFACEDESC2
    Set DisplayModes = dd.GetDisplayModesEnum(0, ddsd)
    Set dd = Nothing
exitOut:
End Function
'-============================================================
' IsFullScreen
'-============================================================
Property Get IsFullScreen() As Boolean
    IsFullScreen = m_bfullscreen
End Property
'-============================================================
' FPS
'-============================================================
Public Property Get FPS() As Single
    FPS = m_fps
End Property
'-============================================================
' Render
'-============================================================
Public Sub Render()
    On Local Error GoTo errout
    If m_binit = False Then Exit Sub
    
    Dim t As Long
    Dim delta As Single
    Static fcount As Long
    
    t = dX.TickCount()
    
    m_rmViewport.Clear D3DRMCLEAR_ZBUFFER Or D3DRMCLEAR_TARGET
    
    ' clear Rect
    ClearRec(0).X1 = m_rmViewport.GetX
    ClearRec(0).Y1 = m_rmViewport.GetY
    ClearRec(0).X2 = m_rmViewport.GetWidth
    ClearRec(0).Y2 = m_rmViewport.GetHeight
    ' calculate the Background X to display
    bakm = (6.28 - Angle) / 6.28 * m_rmViewport.GetWidth * 2
    If bakm > m_rmViewport.GetWidth * 2 Then bakm = m_rmViewport.GetWidth * 2
    If bakm < 0 Then bakm = 0
    ' display the part of the background
     m_backBuffer.Blt REC(0, 0, 0, 0), Background, REC(bakm, 0, bakm + m_rmViewport.GetWidth / 2, m_rmViewport.GetHeight), DDBLT_WAIT
    
     m_rmViewport.Render m_rmFrameScene
     RaiseEvent PostRender

     m_rmDevice.Update
    
     If m_bfullscreen Then
        m_frontBuffer.Flip Nothing, DDFLIP_WAIT
    Else
        If m_bCreateFromClipper = False Then
           Call m_dx.GetWindowRect(m_hwnd, wndr)
           m_frontBuffer.Blt wndr, m_backBuffer, m_emptyrect, DDBLT_WAIT
        End If
    End If
    fcount = fcount + 1
    If fcount = 30 Then
        t = dX.TickCount()
        m_fps = 30000 / (t - m_lastFPS)
        fcount = 0
        m_lastFPS = t
    End If
    
errout:
End Sub
'-============================================================
' hwnd
'-============================================================
Public Property Let hwnd(handle As Long)
    m_hwnd = handle
End Property
Public Sub Resize(w As Long, h As Long)
    On Local Error Resume Next
    
    Dim b As Boolean
    
    If m_binit = False Then
        Exit Sub
    End If
    
    'full screen apps shouldnt resize
    If m_bfullscreen Then Exit Sub
    
    SaveDeviceViewportCharacteristics
   
    If Not m_bUseSoftwareOnly Then
        b = InitWindowed(m_strDDGuid, "IID_IDirect3DHALDevice")
    End If
    If Not b Then
       b = InitWindowed(m_strDDGuid, "IID_IDirect3DRGBDevice")
    End If
    
    RestoreDeviceViewportCharacteristics
    
End Sub

'-============================================================
' Class_Initialize
'-============================================================
Private Sub class_Initialize()
    Dim b As Boolean
    
    m_bCreateFromClipper = True
    
    b = InitSceneGraph()
    If Not b Then
        RaiseEvent DirecXNotInstalled
        Exit Sub
    End If
End Sub

'-============================================================
' Class_Terminate
'-============================================================
Private Sub Class_Terminate()
    Cleanup
    CleanupRMObjects
    Set DSOUND70 = Nothing
    Set DsoundPri70 = Nothing
    Set DsoundLis70 = Nothing
End Sub

'-============================================================
' Cleanup objects that can hold onto vmem
'-============================================================
Private Sub Cleanup()
    Err.Clear
    On Local Error Resume Next
    m_dd.RestoreDisplayMode
    m_dd.SetCooperativeLevel m_hwnd, DDSCL_NORMAL
    Set m_backBuffer = Nothing
    Set m_frontBuffer = Nothing
    Set m_dd = Nothing
    Set m_ddClip = Nothing
    Set m_rmViewport = Nothing
    Set m_rmDevice = Nothing
    m_bfullscreen = False
    m_binit = False
End Sub

'-============================================================
' Cleanup rest of RM objects
'-============================================================
Private Sub CleanupRMObjects()
    Set m_rmFrameCamera = Nothing
    Set m_rmFrameScene = Nothing
    Set m_rmFrameDirLight = Nothing
    Set m_rmFrameAmbientLight = Nothing
    Set m_rmDirLight = Nothing
    Set m_rmAmbientLight = Nothing
End Sub
'-====================================================
' RestoreDeviceViewportCharacteristics
'
' when the viewport is destroyed for whatever reason (resize)
' this function allows us to retain the characteristics
' of the viewport we just destroyed
'-====================================================
Private Sub RestoreDeviceViewportCharacteristics()
    With m_DevInfo
        m_rmDevice.SetDither .bDither
        m_rmDevice.SetName .Name
        m_rmDevice.SetQuality .Quality
        m_rmDevice.SetRenderMode .RenderMode
        m_rmDevice.SetShades .Shades
        m_rmDevice.SetTextureQuality .TexQ
    End With
    With m_ViewInfo
        m_rmViewport.SetBack .Back
        m_rmViewport.SetField .Field
        m_rmViewport.SetFront .Front
        m_rmViewport.SetName .Name
        m_rmViewport.SetProjection .Projection
        m_rmViewport.SetPlane .left, .right, .bottom, .top
        m_rmViewport.SetUniformScaling .scaling
    End With
End Sub
'-====================================================
' SaveDeviceViewportCharacteristics
'
' we need to retain certain characteristics about the
' viewport and device so that they look the same
' when recreated after a resize
'-====================================================
Private Sub SaveDeviceViewportCharacteristics()
    
    With m_DevInfo
        .bDither = m_rmDevice.GetDither
        .Name = m_rmDevice.GetName
        .Quality = m_rmDevice.GetQuality
        .RenderMode = m_rmDevice.GetRenderMode
        .Shades = m_rmDevice.GetShades
        .TexQ = m_rmDevice.GetTextureQuality
        .WFoptions = m_rmDevice.GetWireframeOptions
    End With
    With m_ViewInfo
        .Back = m_rmViewport.GetBack
        .Field = m_rmViewport.GetField
        .Front = m_rmViewport.GetFront
        .Name = m_rmViewport.GetName
        .Projection = m_rmViewport.GetProjection
        .scaling = m_rmViewport.GetUniformScaling
        m_rmViewport.GetPlane .left, .right, .bottom, .top
        
    End With
End Sub
'-====================================================
' SetDeviceDefaults
'-====================================================
Private Sub SetDeviceDefaults()
    m_rmDevice.SetQuality D3DRMRENDER_GOURAUD
End Sub
'-====================================================
' InitSceneGraph
'
' create default lighting and cameras
'-====================================================
Private Function InitSceneGraph() As Boolean
    On Local Error GoTo errout
    'create a skeletal scene graph
    Set m_rm = m_dx.Direct3DRMCreate()
    Set m_rmFrameScene = m_rm.CreateFrame(Nothing)
    Set m_rmFrameCamera = m_rm.CreateFrame(m_rmFrameScene)
    m_rmFrameCamera.SetPosition Nothing, 0, 0, -10
    
    'create a bright directional light
    Set m_rmFrameDirLight = m_rm.CreateFrame(m_rmFrameScene)
    Set m_rmDirLight = m_rm.CreateLightRGB(D3DRMLIGHT_DIRECTIONAL, 1, 1, 1)
    
    'create a dull ambient light
    Set m_rmAmbientLight = m_rm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.2, 0.2, 0.2)
    
    'add the lights to the scene graph
    m_rmFrameDirLight.AddLight m_rmDirLight
    m_rmFrameScene.AddLight m_rmAmbientLight
    m_rmFrameDirLight.SetPosition Nothing, 5, 5, -5
    m_rmFrameDirLight.LookAt m_rmFrameScene, Nothing, 0
       
    InitSceneGraph = True
    Exit Function
errout:
    InitSceneGraph = False
End Function
'-============================================================
' GetBoundingBox
'-============================================================
Public Sub GetBoundingBox(frame As Direct3DRMFrame3, ByRef xmin As Single, ByRef ymin As Single, ByRef zmin As Single, ByRef xMax As Single, ByRef yMax As Single, ByRef zmax As Single)
    Dim box1 As D3DRMBOX
    Dim mb As Direct3DRMMeshBuilder3
    
    Set mb = m_rm.CreateMeshBuilder()
    mb.AddFrame frame
    
    mb.GetBox box1
    
    xmin = box1.Min.x
    ymin = box1.Min.y
    zmin = box1.Min.z
    
    xMax = box1.Max.x
    yMax = box1.Max.y
    zmax = box1.Max.z
        
End Sub
Public Function CreateSheetMesh(nSides As Integer, Height As Single, Width As Single, TU As Single, TV As Single) As Direct3DRMMeshBuilder3

    Dim m As Direct3DRMMeshBuilder3
    Dim f As Direct3DRMFace2
    Set m = m_rm.CreateMeshBuilder()
        
    Dim dX  As Single
    Dim dy  As Single
    dy = Height / 2
    dX = Width / 2
    
    'Front Face
    Set f = m_rm.CreateFace()
    f.AddVertex dX, dy, 0
    f.AddVertex dX, -dy, 0
    f.AddVertex -dX, -dy, 0
    f.AddVertex -dX, dy, 0
    m.AddFace f
    
    m.SetTextureCoordinates 3, 0, 0
    m.SetTextureCoordinates 2, 0, TV
    m.SetTextureCoordinates 1, TU, TV
    m.SetTextureCoordinates 0, TU, 0
    
    
    If nSides > 1 Then
        'Back Face
        Set f = m_rm.CreateFace()
        f.AddVertex -dX, dy, 0
        f.AddVertex -dX, -dy, 0
        f.AddVertex dX, -dy, 0
        f.AddVertex dX, dy, 0
        m.AddFace f
    
        m.SetTextureCoordinates 7, 0, 0
        m.SetTextureCoordinates 6, 0, TV
        m.SetTextureCoordinates 5, TU, TV
        m.SetTextureCoordinates 4, TU, 0
    
    End If
    
    Set CreateSheetMesh = m

End Function
Public Function CreateDDSFromBMP4(Sfile As String, Optional SWidth As Long, Optional Sheight As Long, Optional Memoryusage As MemoryAlloc) As DirectDrawSurface4
    On Local Error GoTo errout
    
    Attemtstr = "Creating DirectdrawSurface From Bitmap " + Sfile
    
    Dim ddsd As DDSURFACEDESC2
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If Memoryusage = USESYSTEMMEMORY Then
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Else
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    End If
    ddsd.lWidth = SWidth
    ddsd.lHeight = Sheight
    If Sfile <> "" Then
        Set CreateDDSFromBMP4 = m_dd.CreateSurfaceFromFile(Sfile, ddsd)
    Else
        Set CreateDDSFromBMP4 = m_dd.CreateSurface(ddsd)
    End If
    
errout:
    RaiseEvent Error4("Error in " + Attemt)
End Function
Private Function CreateDDSFromCDIB4(cDib As cDIBSection, Optional SWID As Long, Optional SHei As Long, Optional Memoryusage As MemoryAlloc) As DirectDrawSurface4
    Dim ddsd As DDSURFACEDESC2
    Dim dds As DirectDrawSurface4
    Dim hdcSurface As Long
    Dim TMPDIB As cDIBSection
    
    If SWID = 0 Then
        SWID = cDib.Width
    End If
    If SHei = 0 Then
        SHei = cDib.Height
    End If
    
    With ddsd
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .lWidth = SWID
        .lHeight = SHei
    End With
    If Memoryusage = USESYSTEMMEMORY Then
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Else
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    End If
    
    Set dds = m_dd.CreateSurface(ddsd)
    dds.restore
    hdcSurface = dds.GetDC
    Set TMPDIB = cDib.Resample(SHei, SWID)
    TMPDIB.PaintPicture hdcSurface
    dds.ReleaseDC hdcSurface
    
    Set TMPDIB = Nothing
    Set CreateDDSFromCDIB4 = dds
    Set dds = Nothing
End Function
Public Function CreateDDSFromCDIBFILE4(file As String, Optional SWID As Long, Optional SHei As Long, Optional Memoryusage As MemoryAlloc) As DirectDrawSurface4
    On Local Error GoTo errout
    Attemtstr = "Creating DirectdrawSurface From FILE " + file
    
    Dim cDib As New cDIBSection
    cDib.CreateFromPicture LoadPicture(file)
    Set CreateDDSFromCDIBFILE4 = CreateDDSFromCDIB4(cDib, SWID, SHei, Memoryusage)
    Set cDib = Nothing
    
    Exit Function
errout:
    RaiseEvent Error4("Error in " + Attemt)
End Function

