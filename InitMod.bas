Attribute VB_Name = "InitMod"
Option Explicit

'-----------------------------------------
'   DX Variables
'-----------------------------------------
Public DX As DirectX8
Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8
Public D3DX As New D3DX8
Public MtrlBuffer As D3DXBuffer

'-----------------------------------------
'   Global Matrix
'-----------------------------------------
Public matView As D3DMATRIX
Public matProj As D3DMATRIX

'-----------------------------------------
'   Screen Stuff
'-----------------------------------------
Public Const ScreenWidth As Integer = 1024
Public Const ScreenHeight As Integer = 768
Public DispMode As D3DDISPLAYMODE
Public D3DWindow As D3DPRESENT_PARAMETERS

Public Running As Boolean 'used to check if the program should be stoped

Public Lights() As D3DLIGHT8 'out lighting array

Public Const RadPerDeg As Single = 1.74532925199433E-02 'this means there are 1.7~~E-02 degrees per radian
Public Const DegPerRad As Single = 57.2957795130823     'this meand there are 57.2~~ radians per degree

Public Const PI    As Single = 3.14159265358979
Public Const D_1   As Single = PI / 180                 '1   degrees in radians
Public Const D_90  As Single = PI / 2                   '90  degrees in radians
Public Const D_180 As Single = PI                       '180 degrees in radians
Public Const D_270 As Single = (PI / 2) * 3             '270 degrees in radians
Public Const D_360 As Single = PI * 2                   '360 degrees in radians

Public DIMouse      As DirectInputDevice8 ' Mouse device
Public DIMState     As DIMOUSESTATE       ' to check mouse movements and clicks
Public MouX         As Single
Public MouY         As Single

Public Const TSPEED     As Single = D_1 + 0.03
Public Const MOUSESPEED As Integer = 1
Public MSPEED       As Integer
Public PSPEED       As Single

Public TPOWER       As Single   'used to move the camera when the snowman turns

Public CamObDist    As Single
Public CamGrDist    As Single
Public Angle As Single, ChrAngle As Single 'stores the direction the snowman and camera are facing
Public Pitch As Single
Public ChrX As Single, ChrY As Single
Public CAMX As Single, CAMY As Single, CAMZ As Single
Public nKills As Integer

'-----------------------------------------
'   Mesh information - 3D objects
'-----------------------------------------

Public Type MeshData
    Matrix   As D3DMATRIX
    Mesh     As D3DXMesh
    Mat()    As D3DMATERIAL8
    Tex()    As Direct3DTexture8
    MatCount As Long
    MX       As Single
    MY       As Single
    MZ       As Single
    MAngle   As Single
    MWidth   As Long
    MLength  As Long
    MHeight  As Long
    RenderMe As Boolean
    LifeSpan As Integer
    Turns    As Integer
End Type

Public WallMesh()    As MeshData
Public GateMesh()    As MeshData
Public TreeMesh()    As MeshData
Public WorldMesh()   As MeshData
Public HouseMesh()   As MeshData
Public RoadMesh()    As MeshData
Public SnowMesh()    As MeshData
Public SnowEvlMesh() As MeshData
Public ThrowMesh()   As MeshData
Public DropMesh()    As MeshData
Public FormMesh()    As MeshData

'--------------------------------------------------------------------------
'   template meshes, used so we only have to load one of each object once
'   from then on we can just assign the template to the new object
'--------------------------------------------------------------------------

Public TemplateWallMesh    As MeshData
Public TemplateGateMesh    As MeshData
Public TemplateTreeMesh    As MeshData
Public TemplateWorldMesh   As MeshData
Public TemplateHouseMesh   As MeshData
Public TemplateRoadMesh    As MeshData
Public TemplateSnowMesh    As MeshData
Public TemplateSnowEvlMesh As MeshData
Public TemplateThrowMesh   As MeshData
Public TemplateDropMesh    As MeshData
Public TemplateFormMesh    As MeshData

Public ThrowCount As Integer, DropCount As Integer, ExpCount As Integer
Public LastThrowTime As Long
Public LastDropTime  As Long
Public Const ThrowSpeed As Integer = 250 'throw a ball every 0.25 seconds
Public Const DropSpeed  As Integer = 250 'drop a christmas tree every 0.25 seconds

'-----------------------------------------
'   Animation Data, keyframe animation
'-----------------------------------------

Public Type AnimFrames
    AnimFMesh As D3DXMesh
    AnimMat() As D3DMATERIAL8
    AnimTex() As Direct3DTexture8
    AnimMCount As Long
    AnimTIndex As Long
    AnimTLength As Long
End Type

Public Type AnimMeshData
    AnimMatrix As D3DMATRIX
    AnimDMesh() As AnimFrames
    AnimTCurrent As Long
    AnimX As Single
    AnimY As Single
    AnimZ As Single
    AnimAngle As Single
    RenderMe As Boolean
End Type

Public TemplateExpMesh As AnimMeshData
Public ExpMesh() As AnimMeshData

'-----------------------------------------
'   Evil Snowman Variables
'-----------------------------------------
Public LastEvlAIChoice As Long
Public LastAIMove As Long
Public EvlAI As String
Public dNum As Integer 'direction number
Public EvlHealth As Integer

'-----------------------------------------
'   FPS Variables
'-----------------------------------------
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public FrameCount As Long
Public LastFrameCount As Long
Public LastTickCount As Long

'-----------------------------------------
'   OnScreen Text
'-----------------------------------------
Public D3DFont(1) As D3DXFont
Public FontDesc As IFont
Public TextRect(3) As RECT
Public TempFont As New StdFont

'-----------------------------------------
'   Game Time Limit variables
'-----------------------------------------
Public Const tLIMIT As Long = 300000 '5 minutes
Public LasttLIMITCheck As Long
Public LastcDownCheck As Long
Public CountDown As Long
'=================================================================================='
Public Function DTR(Deg As Single) As Single
    DTR = Deg * DegPerRad
End Function
'=================================================================================='
Public Function RTD(Rad As Single) As Single
    RTD = Rad * RadPerDeg
End Function
'=================================================================================='
Private Function SetDispMode()
    DispMode.Format = D3DFMT_R5G6B5
    DispMode.Width = ScreenWidth
    DispMode.Height = ScreenHeight
    D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
    D3DWindow.Windowed = 0
    D3DWindow.BackBufferCount = 1
    D3DWindow.BackBufferFormat = D3DFMT_R5G6B5
    D3DWindow.BackBufferWidth = ScreenWidth
    D3DWindow.BackBufferHeight = ScreenHeight
    D3DWindow.hDeviceWindow = frmMain.hWnd
    D3DWindow.EnableAutoDepthStencil = 1
    D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
End Function
'=================================================================================='
Public Function InitDX() As Boolean
    On Error GoTo BailOut
    
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    SetDispMode 'set up the Screen display settings
    
    D3DDevice.SetRenderState D3DRS_LIGHTING, 1  'enable the lighting
    D3DDevice.SetRenderState D3DRS_ZENABLE, 1   'enable the ZBuffer
    D3DDevice.SetRenderState D3DRS_AMBIENT, &HDDDDDD    'Setup the ambient scene lighting
    
    'the following to line make our textures look better, try to comment them out and see what happens
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
           
    'setup our view of the world
    D3DXMatrixIdentity matView
    D3DXMatrixLookAtLH matView, MakeVector(0, -D_360, D_1), MakeVector(0, 0, 0), MakeVector(0, 1, 0)
    D3DDevice.SetTransform D3DTS_WORLD, matView 'tell the world how we are looking at it
    
    D3DXMatrixPerspectiveFovLH matProj, PI / 4, PI * 0.3, 5, 750
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    
'-------------------------------------------
'   InGame Text
'-------------------------------------------
    TempFont.Name = "Times New Roman"
    TempFont.Size = 10
    Set FontDesc = TempFont
    Set D3DFont(0) = D3DX.CreateFont(D3DDevice, FontDesc.hFont)    'Create the font
'-------------------------------------------
'   Game Over Text
'-------------------------------------------
    TempFont.Name = "Times New Roman"
    TempFont.Size = 14
    Set FontDesc = TempFont
    Set D3DFont(1) = D3DX.CreateFont(D3DDevice, FontDesc.hFont)    'Create the font
    
    Init3DObj   'Init our 3d objects
    InitLights  'Init our Lights
    InitKeyBoardMouse
    
    Running = True
    
    ThrowCount = 0: DropCount = 0: ExpCount = 0: EvlHealth = 100: TPOWER = 4
    LasttLIMITCheck = GetTickCount()
    LastcDownCheck = GetTickCount()
    
    InitDX = True
    Exit Function
BailOut:
    InitDX = False
End Function
'=================================================================================='
Public Function Init3DObj()
    'Create the Template Mesh's and The Acutal Mesh's
    MyReadMeshFromX "\MultiXFiles\MultiX(Wall).txt", TemplateWallMesh, WallMesh()
    MyReadMeshFromX "\MultiXFiles\MultiX(Gate).txt", TemplateGateMesh, GateMesh()
    MyReadMeshFromX "\MultiXFiles\MultiX(Ground).txt", TemplateWorldMesh, WorldMesh()
    MyReadMeshFromX "\MultiXFiles\MultiX(Snowman).txt", TemplateSnowMesh, SnowMesh()
    MyReadMeshFromX "\MultiXFiles\MultiX(SnowmanEvl).txt", TemplateSnowEvlMesh, SnowEvlMesh()
    MyReadMeshFromX "\MultiXFiles\MultiX(Tree).txt", TemplateTreeMesh, TreeMesh()
    MyReadMeshFromX "\MultiXFiles\MultiX(House).txt", TemplateHouseMesh, HouseMesh()
    MyReadMeshFromX "\MultiXFiles\MultiX(Road).txt", TemplateRoadMesh, RoadMesh()
        
    CreateMesh "\XFiles\SnowBall.x", TemplateThrowMesh, 8, 8
    CreateMesh "\XFiles\Marker(Tree).x", TemplateDropMesh, 16, 16
    
    CreateMesh "\XFiles\FormBackDrop.x", TemplateFormMesh, 128, 128
    ReDim FormMesh(0) As MeshData: FormMesh(0) = TemplateFormMesh
        
    'Create the keyframe Animation Mesh's
    ReadAnimFile "\AnimXFiles\KeyFrame(Snowball).txt", TemplateExpMesh
End Function
'=================================================================================='
Private Function InitLights()
    On Error GoTo BailOut
    Dim Mtrl As D3DMATERIAL8 'Material
    Dim Col As D3DCOLORVALUE 'Color
    'Create color
    Col.a = 1: Col.r = 1: Col.g = 1: Col.B = 1
    'Apply material
    Mtrl.Ambient = Col: Mtrl.diffuse = Col
    D3DDevice.SetMaterial Mtrl
    
    'Create directional light
    Lights(0).Type = D3DLIGHT_DIRECTIONAL
    Lights(0).diffuse.r = 1
    Lights(0).diffuse.g = 1
    Lights(0).diffuse.B = 1
    Lights(0).Direction = MakeVector(0, 0, 50)

    'Apply lights to device
    D3DDevice.SetLight 0, Lights(0)
    
    InitLights = True
    Exit Function
BailOut:
    InitLights = False
End Function
'=================================================================================='
Private Function InitKeyBoardMouse() As Boolean
On Local Error GoTo BailOut:
'-----------------------------------------------------------------------------------
'   Keyboard Init Section
'-----------------------------------------------------------------------------------
    'Create a new DirectX8 Input device
    Set DKI = DX.DirectInputCreate
    'Tell this new device it is going to be a keyboard device
    Set DKIDevice = DKI.CreateDevice("GUID_SysKeyboard")
    'Set the new keyboard devive to the common keyboard format
    DKIDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    'Set its cooperative level this is basically for telling the keyboard if it has soul right to the input from the keyboard or can the windows environment also use the input
    DKIDevice.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DevProp.lHow = DIPH_DEVICE
    DevProp.lData = 10 'set the size of the keybuffer
    DKIDevice.SetProperty DIPROP_BUFFERSIZE, DevProp
    'start using the keyboard device
    DKIDevice.Acquire
'-----------------------------------------------------------------------------------
'   Mouse Init Section
'-----------------------------------------------------------------------------------
    'Tell this device it is going to be a Mouse device
    Set DIMouse = DKI.CreateDevice("GUID_SysMouse")
    'set the mouse's format to the common mouse format
    DIMouse.SetCommonDataFormat DIFORMAT_MOUSE
    DIMouse.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE 'DISCL_FOREGROUND Or DISCL_EXCLUSIVE
    DIMouse.Acquire 'start using the mouse
'-----------------------------------------------------------------------------------
    InitKeyBoardMouse = True
    Exit Function
BailOut:
    InitKeyBoardMouse = False
End Function
'=================================================================================='
Public Function MyReadMeshFromX(FileName As String, TemplateMesh As MeshData, InMesh() As MeshData)
    Dim InText, XFileName As String, FileNum, LoopCount, MeshNum, ISizeX As Integer, ISizeY As Integer, IX, IY, IZ, IWid, ILen, IHei As Long, IAngle As Single
    FileNum = FreeFile
    Open App.Path + FileName For Input As FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, InText
            If (Not (InText = "") Or (InText = ";")) Then
                Select Case InText
                Case "<FileName>"
                    Input #FileNum, XFileName
                Case "<TextureSize>"
                    Input #FileNum, ISizeX, ISizeY
                Case "<Location>"
                    'create a template of this Mesh
                    CreateMesh XFileName, TemplateMesh, ISizeX, ISizeY
                    Input #FileNum, LoopCount
                    For MeshNum = 0 To LoopCount
                        'read from the file all our values for this new mesh
                        Input #FileNum, IX, IY, IZ, IWid, ILen, IHei, IAngle
                        ReDim Preserve InMesh(MeshNum)
                        'Assign the new mesh the template values, this is muxh faster then create a new mesh each time
                        InMesh(MeshNum) = TemplateMesh
                        InMesh(MeshNum).MX = IX 'Assign the x,y,z values
                        InMesh(MeshNum).MY = IY
                        InMesh(MeshNum).MZ = IZ
                        'Setup the Width of an object depending on the angle it is facing
                        If IAngle = 1 Then InMesh(MeshNum).MWidth = ILen Else InMesh(MeshNum).MWidth = IWid
                        'Setup the Length of an object depending on the angle it is facing
                        If IAngle = 1 Then InMesh(MeshNum).MLength = IWid Else InMesh(MeshNum).MLength = ILen
                        InMesh(MeshNum).MHeight = IHei 'setup the height of the object
                        'Setup the Angle of the object
                        If IAngle = 1 Then InMesh(MeshNum).MAngle = D_90 Else InMesh(MeshNum).MAngle = 0
                        InMesh(MeshNum).RenderMe = True
                    Next MeshNum
                End Select
            End If
        Loop
    Close FileNum
End Function
'==================================================================================
Public Function CreateMesh(FileName As String, InMesh As MeshData, SizeX As Integer, SizeY As Integer)
Dim TextureName As String, q As Integer
    Set InMesh.Mesh = D3DX.LoadMeshFromX(App.Path & FileName, D3DXMESH_MANAGED, D3DDevice, Nothing, MtrlBuffer, InMesh.MatCount)
    ReDim InMesh.Mat(InMesh.MatCount - 1) As D3DMATERIAL8   'setup the materials array
    ReDim InMesh.Tex(InMesh.MatCount - 1) As Direct3DTexture8   'setup the texture array
    For q = 0 To InMesh.MatCount - 1
        D3DX.BufferGetMaterial MtrlBuffer, q, InMesh.Mat(q)
        'setup the ambient lighting
        InMesh.Mat(q).Ambient = InMesh.Mat(q).diffuse
        'get the texture name from the 3D object
        TextureName = D3DX.BufferGetTextureName(MtrlBuffer, q)
        'assign the texture to the new 3d object
        If TextureName <> "" Then Set InMesh.Tex(q) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Textures\" & TextureName, SizeX, SizeY, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    Next q
    Set MtrlBuffer = Nothing 'Clear the Material buffer
End Function
'==================================================================================
Public Function ReadAnimFile(FileName As String, TemplateMesh As AnimMeshData)
    Dim InText As String, XFileName As String
    Dim nFrames As Integer, aLength As Integer, FileNum As Integer
    Dim SizeX As Integer, SizeY As Integer, mNum As Long
    FileNum = FreeFile
    Open App.Path + FileName For Input As FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, InText
            If (Not (InText = "") Or (InText = ";")) Then
                Select Case InText
                Case "<NumFrames>":  Input #FileNum, nFrames
                Case "<AnimLength>": Input #FileNum, aLength
                Case "<FileName>"
                    ReDim TemplateMesh.AnimDMesh(nFrames) As AnimFrames
                    For mNum = 0 To nFrames
                        Input #FileNum, XFileName, SizeX, SizeY
                        'create a template of the animation
                        CreateNewAnimMesh XFileName, TemplateMesh, mNum, SizeX, SizeY
                        If mNum = 0 Then TemplateMesh.AnimDMesh(mNum).AnimTIndex = 0 Else TemplateMesh.AnimDMesh(mNum).AnimTIndex = aLength * (mNum / (nFrames + 1))
                        If mNum = nFrames Then TemplateMesh.AnimDMesh(mNum).AnimTLength = aLength Else TemplateMesh.AnimDMesh(mNum).AnimTLength = aLength * ((mNum + 1) / (nFrames + 1))
                        TemplateMesh.AnimTCurrent = 0
                    Next mNum
                End Select
            End If
        Loop
    Close FileNum
End Function
'==================================================================================
Public Function CreateNewAnimMesh(FileName As String, TempMesh As AnimMeshData, fNum As Long, SizeX As Integer, SizeY As Integer)
Dim TextureName As String, q As Integer
    Set TempMesh.AnimDMesh(fNum).AnimFMesh = D3DX.LoadMeshFromX(App.Path & FileName, D3DXMESH_MANAGED, D3DDevice, Nothing, MtrlBuffer, TempMesh.AnimDMesh(fNum).AnimMCount)
    ReDim TempMesh.AnimDMesh(fNum).AnimMat(TempMesh.AnimDMesh(fNum).AnimMCount - 1) As D3DMATERIAL8
    ReDim TempMesh.AnimDMesh(fNum).AnimTex(TempMesh.AnimDMesh(fNum).AnimMCount - 1) As Direct3DTexture8
    For q = 0 To TempMesh.AnimDMesh(fNum).AnimMCount - 1
        D3DX.BufferGetMaterial MtrlBuffer, q, TempMesh.AnimDMesh(fNum).AnimMat(q)
        TempMesh.AnimDMesh(fNum).AnimMat(q).Ambient = TempMesh.AnimDMesh(fNum).AnimMat(q).diffuse
        TextureName = D3DX.BufferGetTextureName(MtrlBuffer, q)
        If TextureName <> "" Then Set TempMesh.AnimDMesh(fNum).AnimTex(q) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Textures\" & TextureName, SizeX, SizeY, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    Next q
    Set MtrlBuffer = Nothing
End Function
'==================================================================================
Public Function CreateSnowMeshObj(TemplateMesh As MeshData, InMesh() As MeshData, InMeshNum As Integer, InZ As Integer, InW As Integer, InH As Integer, InL As Integer, InLifeSpan As Integer, InTurns As Integer, SrcX As Single, SrcY As Single, SrcAngle As Single)
    'this function is used to create an object whilst the program is running eg a snowball
    InMesh(InMeshNum) = TemplateMesh
    InMesh(InMeshNum).MX = SrcX
    InMesh(InMeshNum).MY = SrcY
    InMesh(InMeshNum).MZ = InZ
    InMesh(InMeshNum).MWidth = InW
    InMesh(InMeshNum).MLength = InL
    InMesh(InMeshNum).MHeight = InH
    InMesh(InMeshNum).MAngle = -SrcAngle
    InMesh(InMeshNum).LifeSpan = InLifeSpan
    InMesh(InMeshNum).Turns = InTurns
    InMesh(InMeshNum).RenderMe = True
End Function
'==================================================================================
Public Function CreateAnimMeshObj(TemplateMesh As AnimMeshData, InMesh As AnimMeshData, SrcX As Single, SrcY As Single, SrcZ As Single, SrcAngle As Single)
    InMesh = TemplateMesh
    InMesh.AnimX = SrcX
    InMesh.AnimY = SrcY
    InMesh.AnimZ = SrcZ
    InMesh.AnimAngle = SrcAngle
    InMesh.RenderMe = True
End Function
'==================================================================================
Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    MakeVector.X = X: MakeVector.Y = Y: MakeVector.Z = Z
End Function
'==================================================================================

