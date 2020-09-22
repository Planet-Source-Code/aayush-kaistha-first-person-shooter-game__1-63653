Attribute VB_Name = "modInitialize"
'#########################################################
'#                                                       #
'#      A First Person Shooting game (Incomplete)        #
'#                                                       #
'#      By: Aayush Kaistha                               #
'#      Place: UIET, Panjab University Chandigarh, India #
'#      Contact: aayushkaistha@gmail.com                 #
'#                                                       #
'#########################################################


Option Explicit

Public Function InitD3D() As Boolean

On Error GoTo Hell:

Dim D3DWindow As D3DPRESENT_PARAMETERS
Dim DispMode As D3DDISPLAYMODE

'initialize and allocate memory 4 directX objects
Set DX = New DirectX8
Set D3D = DX.Direct3DCreate
Set D3DX = New D3DX8

DispMode.Format = CheckDisplayMode(640, 480, 32)
If DispMode.Format > D3DFMT_UNKNOWN Then
    Debug.Print "Using 32-Bit format"
Else
    DispMode.Format = CheckDisplayMode(640, 480, 16)
    If DispMode.Format > D3DFMT_UNKNOWN Then
        Debug.Print "32-Bit format not supported. Using 16-Bit format"
    Else
        MsgBox "Neither 16-Bit nor 32-Bit Display Mode Supported", vbInformation, "ERROR"
        Unload frmMain
        End
    End If
End If

With D3DWindow
    .BackBufferCount = 1
    .BackBufferFormat = DispMode.Format
    .BackBufferWidth = 640
    .BackBufferHeight = 480
    .hDeviceWindow = frmMain.hWnd
    .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
End With

If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D32) = D3D_OK Then
    D3DWindow.AutoDepthStencilFormat = D3DFMT_D32
    D3DWindow.EnableAutoDepthStencil = 1
Else
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D24X8) = D3D_OK Then
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D24X8
        D3DWindow.EnableAutoDepthStencil = 1
    Else
        If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
            D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
            D3DWindow.EnableAutoDepthStencil = 1
        Else
            D3DWindow.EnableAutoDepthStencil = 0
            MsgBox "Depth buffer could not be enabled", vbInformation, "Depth buffer not supported"
        End If
    End If
End If

Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)

With D3DDevice
    .SetVertexShader FVF_VERTEX
    .SetRenderState D3DRS_LIGHTING, 1
    .SetRenderState D3DRS_AMBIENT, D3DColorXRGB(150, 150, 150)
    .SetRenderState D3DRS_ZENABLE, 1
End With

D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
    
D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
D3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_LINEAR

D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

'initialize our matrices
D3DXMatrixIdentity matWorld
D3DDevice.SetTransform D3DTS_WORLD, matWorld

D3DXMatrixLookAtLH matView, MakeVector(-2, 2, -2), MakeVector(0, 0, 0), MakeVector(0, 1, 0)
D3DDevice.SetTransform D3DTS_VIEW, matView

D3DXMatrixPerspectiveFovLH matProj, PI / 3, 1, 1, 10000
D3DDevice.SetTransform D3DTS_PROJECTION, matProj

'font settings
SetFont "Verdana", 12, True

InitD3D = True

Exit Function
Hell:
MsgBox "ERROR initializing D3D ", vbCritical, "ERROR"
InitD3D = False

End Function

Public Sub InitDInput()

Set DI = DX.DirectInputCreate
Set DIDevice = DI.CreateDevice("guid_SysMouse")
Call DIDevice.SetCommonDataFormat(DIFORMAT_MOUSE)
Call DIDevice.SetCooperativeLevel(frmMain.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE)

Dim diProp As DIPROPLONG
diProp.lHow = DIPH_DEVICE
diProp.lObj = 0
diProp.lData = 10
    
Call DIDevice.SetProperty("DIPROP_BUFFERSIZE", diProp)
hEvent = DX.CreateEvent(frmMain)
DIDevice.SetEventNotification hEvent

DIDevice.Acquire

End Sub

Public Sub InitDSound()
On Error Resume Next

Dim DSBDesc As DSBUFFERDESC

Set DSEnum = DX.GetDSEnum
Set DS = DX.DirectSoundCreate(DSEnum.GetGuid(1))
DS.SetCooperativeLevel frmMain.hWnd, DSSCL_NORMAL

DSBDesc.lFlags = DSBCAPS_CTRLVOLUME
Set sndShoot = DS.CreateSoundBufferFromFile(SoundPath + "shoot.wav", DSBDesc)

If sndShoot Is Nothing Then Exit Sub
sndShoot.SetVolume -1000

End Sub

Public Function CheckRangeFog(adapter As Byte) As Boolean
On Local Error Resume Next
    Dim DX As New DirectX8
    Dim D3D As Direct3D8
    Dim Caps As D3DCAPS8
    
    Set D3D = DX.Direct3DCreate
    
    D3D.GetDeviceCaps adapter - 1, D3DDEVTYPE_HAL, Caps
    
    If Caps.RasterCaps And D3DPRASTERCAPS_FOGRANGE Then
        CheckRangeFog = True
    Else
        CheckRangeFog = False
    End If
End Function

Public Sub LoadModels()

LoadFromX Floor, "floor.x"
LoadFromX Build(0), "build1.x"
LoadFromX Build(1), "build2.x"
LoadFromX Build(2), "build3.x"
LoadFromX Build(3), "build4.x"
LoadFromX Build(4), "build5.x"
LoadFromX Build(5), "build6.x"
LoadFromX Build(6), "build7.x"
LoadFromX Lawn, "lawn.x"
LoadFromX HardFloor, "hardfloor.x"
LoadFromX Pole, "pole.x"
LoadFromX Wall, "wall.x"
LoadFromX SWall, "swall.x"
LoadFromX MDoor(0), "mdoor1.x"
LoadFromX MDoor(1), "mdoor2.x"
LoadFromX Gun1, "gun1.x"
LoadFromX Gun2, "gun2.x"

End Sub

Public Sub LoadFromX(ByRef tmpObj As Object3D, XFile As String)

On Error GoTo Out:

Dim mtrlBuffer As D3DXBuffer
Dim i As Long, j As Integer

Set tmpObj.Mesh = D3DX.LoadMeshFromX(ModelPath + XFile, D3DXMESH_MANAGED, D3DDevice, Nothing, mtrlBuffer, tmpObj.nMaterials)
    
ReDim tmpObj.Materials(tmpObj.nMaterials) As D3DMATERIAL8
ReDim tmpObj.Textures(tmpObj.nMaterials) As Direct3DTexture8

For i = 0 To tmpObj.nMaterials - 1
    D3DX.BufferGetMaterial mtrlBuffer, i, tmpObj.Materials(i)
    tmpObj.Materials(i).Ambient = tmpObj.Materials(i).diffuse
    tmpObj.TextureFile = D3DX.BufferGetTextureName(mtrlBuffer, i)
    If tmpObj.TextureFile <> "" Then
        Set tmpObj.Textures(i) = D3DX.CreateTextureFromFile(D3DDevice, TexturePath + tmpObj.TextureFile)
    End If
Next

Exit Sub
Out:
    MsgBox "Error loading models", vbCritical, "ERROR"
End Sub

Public Sub SetFont(Name As String, Size As Integer, Bold As Boolean)

fnt.Name = Name
fnt.Size = Size
fnt.Bold = Bold
Set MainFontDesc = fnt
Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)

End Sub

Public Sub Initialize()
Dim i As Integer

bRunning = InitD3D
InitDInput
InitDSound
InitCH
InitFlash
LoadModels
LoadMap
SetLights
'CalModelDimensions
ResetVariables

Set FlashTex = D3DX.CreateTextureFromFileEx(D3DDevice, TexturePath + "flash.tga", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_A1R5G5B5, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, &HFF000000, ByVal 0, ByVal 0)
Set TreeTex = D3DX.CreateTextureFromFileEx(D3DDevice, TexturePath & "tree.tga", 256, 256, D3DX_DEFAULT, 0, D3DFMT_A1R5G5B5, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, &HFF000000, ByVal 0, ByVal 0)
CreateTree

For i = 0 To 5
    Set SkyTex(i) = D3DX.CreateTextureFromFile(D3DDevice, SkyTexFile(i))
Next
CreateSky

ShowCursor 0

End Sub

Public Sub ResetVariables()

With Player
    .Pos = MakeVector(0, 70, -1900)
    .Rotation = 0
    .Health = 3
    .Ammo = 20
    .Dead = False
    .Hit = False
    .Score = 0
End With

CamPitch = 0

End Sub

Public Sub InitCH()

CH(0) = CreateTLVertex(300, 240, 0, 1, &HFF00&, 0, 0, 0)
CH(1) = CreateTLVertex(315, 240, 0, 1, &HFF00&, 0, 0, 0)
CH(2) = CreateTLVertex(325, 240, 0, 1, &HFF00&, 0, 0, 0)
CH(3) = CreateTLVertex(340, 240, 0, 1, &HFF00&, 0, 0, 0)

CH(4) = CreateTLVertex(320, 220, 0, 1, &HFF00&, 0, 0, 0)
CH(5) = CreateTLVertex(320, 235, 0, 1, &HFF00&, 0, 0, 0)
CH(6) = CreateTLVertex(320, 245, 0, 1, &HFF00&, 0, 0, 0)
CH(7) = CreateTLVertex(320, 260, 0, 1, &HFF00&, 0, 0, 0)

End Sub

Public Sub InitFlash()

Flash(0) = CreateLitVertex(-1, 1, 13, &HFFFFFF, 0, 0, 0)
Flash(1) = CreateLitVertex(1, 1, 13, &HFFFFFF, 0, 1, 0)
Flash(2) = CreateLitVertex(-1, -1, 13, &HFFFFFF, 0, 0, 1)
Flash(3) = CreateLitVertex(1, -1, 13, &HFFFFFF, 0, 1, 1)

End Sub

Public Function CheckDisplayMode(Width As Long, Height As Long, Depth As Long) As CONST_D3DFORMAT
Dim i As Long
Dim DispMode As D3DDISPLAYMODE
    
For i = 0 To D3D.GetAdapterModeCount(0) - 1
    D3D.EnumAdapterModes 0, i, DispMode
    If DispMode.Width = Width Then
        If DispMode.Height = Height Then
            If (DispMode.Format = D3DFMT_R5G6B5) Or (DispMode.Format = D3DFMT_X1R5G5B5) Or (DispMode.Format = D3DFMT_X4R4G4B4) Then
                '16 bit mode
                If Depth = 16 Then
                    CheckDisplayMode = DispMode.Format
                    Exit Function
                End If
            ElseIf (DispMode.Format = D3DFMT_R8G8B8) Or (DispMode.Format = D3DFMT_X8R8G8B8) Then
                '32bit mode
                If Depth = 32 Then
                    CheckDisplayMode = DispMode.Format
                    Exit Function
                End If
            End If
        End If
    End If
Next i
CheckDisplayMode = D3DFMT_UNKNOWN
End Function

Public Sub CreateRectangle(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Color As Long, ByRef tmp() As TLVERTEX)

tmp(0) = CreateTLVertex(X1, Y1, 0, 1, Color, 0, 0, 0)
tmp(1) = CreateTLVertex(X2, Y1, 0, 1, Color, 0, 1, 0)
tmp(2) = CreateTLVertex(X1, Y2, 0, 1, Color, 0, 0, 1)
tmp(3) = CreateTLVertex(X2, Y2, 0, 1, Color, 0, 1, 1)

End Sub

Public Sub CreateTree()

Tree(0) = CreateVertex(-100, 300, 0, -0.316, 0.316, 0.316, 0, 0)
Tree(1) = CreateVertex(100, 300, 0, 0.316, 0.316, 0.316, 1, 0)
Tree(2) = CreateVertex(100, 0, 0, 1, 1, 1, 1, 1)
Tree(3) = CreateVertex(100, 0, 0, 1, 1, 1, 1, 1)
Tree(4) = CreateVertex(-100, 0, 0, -1, 1, 1, 0, 1)
Tree(5) = CreateVertex(-100, 300, 0, -0.316, 0.316, 0.316, 0, 0)

End Sub

Public Sub CreateSky()

Sky(0) = CreateVertex(-5000, -5000, 5000, -0.577, -0.577, 0.577, 1, 1)
Sky(1) = CreateVertex(-5000, 5000, 5000, -0.577, 0.577, 0.577, 1, 0)
Sky(2) = CreateVertex(5000, -5000, 5000, 0.577, -0.577, 0.577, 0, 1)
Sky(3) = CreateVertex(5000, 5000, 5000, 0.577, 0.577, 0.577, 0, 0)

Sky(4) = CreateVertex(-5000, -5000, -5000, -0.577, -0.577, -0.577, 0, 1)
Sky(5) = CreateVertex(5000, -5000, -5000, 0.577, -0.577, -0.577, 1, 1)
Sky(6) = CreateVertex(-5000, 5000, -5000, -0.577, 0.577, -0.577, 0, 0)
Sky(7) = CreateVertex(5000, 5000, -5000, 0.577, 0.577, -0.577, 1, 0)

Sky(8) = CreateVertex(5000, -5000, -5000, 0.577, -0.577, -0.577, 0, 1)
Sky(9) = CreateVertex(5000, -5000, 5000, 0.577, -0.577, 0.577, 1, 1)
Sky(10) = CreateVertex(5000, 5000, -5000, 0.577, 0.577, -0.577, 0, 0)
Sky(11) = CreateVertex(5000, 5000, 5000, 0.577, 0.577, 0.577, 1, 0)

Sky(12) = CreateVertex(-5000, -5000, 5000, -0.577, -0.577, 0.577, 0, 1)
Sky(13) = CreateVertex(-5000, -5000, -5000, -0.577, -0.577, -0.577, 1, 1)
Sky(14) = CreateVertex(-5000, 5000, 5000, -0.577, 0.577, 0.577, 0, 0)
Sky(15) = CreateVertex(-5000, 5000, -5000, -0.577, 0.577, -0.577, 1, 0)

Sky(16) = CreateVertex(5000, 5000, -5000, 0.577, 0.577, -0.577, 0, 1)
Sky(17) = CreateVertex(5000, 5000, 5000, 0.577, 0.577, 0.577, 1, 1)
Sky(18) = CreateVertex(-5000, 5000, -5000, -0.577, 0.577, -0.577, 0, 0)
Sky(19) = CreateVertex(-5000, 5000, 5000, -0.577, 0.577, 0.577, 1, 0)

Sky(20) = CreateVertex(5000, -5000, 5000, 0.577, -0.577, 0.577, 1, 0)
Sky(21) = CreateVertex(5000, -5000, -5000, 0.577, -0.577, -0.577, 0, 0)
Sky(22) = CreateVertex(-5000, -5000, 5000, -0.577, -0.577, 0.577, 1, 1)
Sky(23) = CreateVertex(-5000, -5000, -5000, -0.577, -0.577, -0.577, 0, 1)

End Sub

Public Sub CalModelDimensions()


End Sub

Public Sub CalMeshDimen(ByRef tmpObj As Object3D)
Dim min As D3DVECTOR, max As D3DVECTOR

D3DX.ComputeBoundingBoxFromMesh tmpObj.Mesh, min, max
With tmpObj.Radius
    .X = (max.X - min.X) / 2
    .Y = (max.Y - min.Y) / 2
    .z = (max.z - min.z) / 2
End With

End Sub
