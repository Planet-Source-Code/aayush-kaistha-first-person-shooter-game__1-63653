Attribute VB_Name = "modDeclarations"
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

Public DX As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8
Public DI As DirectInput8, DIDevice As DirectInputDevice8
Public hEvent As Long 'a handle for an event...
Public DS As DirectSound8
Public DSEnum As DirectSoundEnum8
Public sndShoot As DirectSoundSecondaryBuffer8
Public sndReload As DirectSoundSecondaryBuffer8

Public ModelPath As String, TexturePath As String
Public SoundPath As String
Public BulHit As Long, HitDist As Single, HitPos As D3DVECTOR
Public bRunning As Boolean, GameOver As Boolean
Public UpKey As Boolean, DownKey As Boolean
Public Restart As Boolean, EndState As Integer
Public LeftKey As Boolean, RightKey As Boolean
Public WKey As Boolean, SKey As Boolean
Public Day As Boolean, Zoom As Boolean
Public EnemySaw As Long, CanHit As Boolean
Public EneDir As D3DVECTOR
Public WarningMsg As String
Public Fire As Boolean, FireTimer As Long
Public UseFog As Boolean, FogCol As Long
Public FogRange As Single

Public Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)
Public Const FVF_LVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)
Public Const FVF_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR)

Public Type LVERTEX
    X As Single
    Y As Single
    z As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type
Public Flash(3) As LVERTEX
Public FlashTex As Direct3DTexture8

Public Type VERTEX
    X As Single
    Y As Single
    z As Single
    nx As Single
    ny As Single
    nz As Single
    tu As Single
    tv As Single
End Type
Public Tree(5) As VERTEX
Public Sky(23) As VERTEX
Public SkyTex(5) As Direct3DTexture8
Public SkyTexFile(5) As String

Public Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type
Public CH(7) As TLVERTEX

'all these variables r required to load 3d objects in directX
Public Type Object3D
    nMaterials As Long
    Materials() As D3DMATERIAL8
    Textures() As Direct3DTexture8
    TextureFile As String
    Mesh As D3DXMesh
    Radius As D3DVECTOR
End Type
Public Floor As Object3D
Public Lawn As Object3D
Public HardFloor As Object3D
Public Pole As Object3D
Public Wall As Object3D
Public SWall As Object3D
Public MDoor(1) As Object3D
Public Gun As Object3D
Public Gun1 As Object3D
Public Gun2 As Object3D

Public Const no_build = 6
Public Build(no_build) As Object3D

Public Type Plyr_Data
    Pos As D3DVECTOR
    Rotation As Single
    MoveSpeed As Single
    Health As Integer
    Hit As Boolean
    Dead As Boolean
    DieTime As Long
    Score As Integer
    Ammo As Integer
End Type

Public EyeLookDir As D3DVECTOR, EyeLookAt As D3DVECTOR
Public matCam As D3DMATRIX

'this only holds data req to calculate frames per second
Public Type FPS_data
    Count As Long
    Value As Long
    Last As Long
End Type
Public Fps As FPS_data

Public MainFont As D3DXFont
Public MainFontDesc As IFont
Public fnt As New StdFont

Public Type Mesh_Dimen
    Center As D3DVECTOR
    Radius As D3DVECTOR
End Type

Public CamPitch As Single
Public Player As Plyr_Data
Public TreeTex As Direct3DTexture8

Public Const floor_seg = 39
Public Const no_poles = 3
Public Const no_trees = 6
Public Const no_walls = 24
Public Const no_swalls = 3

Public matFloor(floor_seg) As D3DMATRIX
Public matBuild(no_build) As D3DMATRIX
Public matLawn(floor_seg) As D3DMATRIX
Public matHFloor(floor_seg) As D3DMATRIX
Public matPole(no_poles) As D3DMATRIX
Public matTree(no_trees) As D3DMATRIX
Public TreePos(no_trees) As D3DVECTOR
Public matWall(no_walls) As D3DMATRIX
Public matSWall(no_swalls) As D3DMATRIX
Public matMDoor1(1) As D3DMATRIX
Public matMDoor2(1) As D3DMATRIX
Public matGun As D3DMATRIX
Public matFlash As D3DMATRIX

Public matProj As D3DMATRIX 'this holds the camera settings
Public matView As D3DMATRIX 'this tells where the camera is n where it is looking at
Public matWorld As D3DMATRIX 'this holds the reference coordinates of entire 3d world

Public Const PI = 3.14159
Public Const Rad = PI / 180
Public Const DEG = 180 / PI

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function CreateLitVertex(X As Single, Y As Single, z As Single, Colour As Long, Specular As Long, tu As Single, tv As Single) As LVERTEX
    CreateLitVertex.X = X
    CreateLitVertex.Y = Y
    CreateLitVertex.z = z
    CreateLitVertex.Color = Colour
    CreateLitVertex.Specular = Specular
    CreateLitVertex.tu = tu
    CreateLitVertex.tv = tv
End Function

Public Function CreateVertex(X As Single, Y As Single, z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single) As VERTEX
    CreateVertex.X = X
    CreateVertex.Y = Y
    CreateVertex.z = z
    CreateVertex.nx = nx
    CreateVertex.ny = ny
    CreateVertex.nz = nz
    CreateVertex.tu = tu
    CreateVertex.tv = tv
End Function

Public Function CreateTLVertex(X As Single, Y As Single, z As Single, rhw As Single, _
                                                Color As Long, Specular As Long, tu As Single, _
                                                tv As Single) As TLVERTEX
    CreateTLVertex.X = X
    CreateTLVertex.Y = Y
    CreateTLVertex.z = z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.Color = Color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
End Function

Public Function MakeVector(X As Single, Y As Single, z As Single) As D3DVECTOR
    MakeVector.X = X: MakeVector.Y = Y: MakeVector.z = z
End Function

Public Function MakeRect(Left As Single, Right As Single, Top As Single, Bottom As Single) As RECT
    MakeRect.Left = Left
    MakeRect.Right = Right
    MakeRect.Top = Top
    MakeRect.Bottom = Bottom
End Function

