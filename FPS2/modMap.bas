Attribute VB_Name = "modMap"
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
Dim matTemp As D3DMATRIX, PolePos(no_poles) As D3DVECTOR

Public Sub LoadMap()

CreateFloor
PlaceBuildings
PlaceWalls
PlaceObjects

End Sub

Public Sub CreateFloor()
Dim i As Integer, j As Single

j = -2000
For i = 0 To floor_seg
    TranslateMatrix matFloor(i), MakeVector(0, 0, j), True
    TranslateMatrix matLawn(i), MakeVector(450, 0, j), True
    TranslateMatrix matHFloor(i), MakeVector(800, 0, j), True
    j = j + 100
Next

End Sub

Public Sub PlaceWalls()
Dim i As Integer, j As Single

j = 300
For i = 0 To 2
    TranslateMatrix matWall(i), MakeVector(j, 0, -2060), True
    j = j + 200
Next

j = 300
For i = 3 To 5
    TranslateMatrix matWall(i), MakeVector(j, 0, 1960), True
    j = j + 200
Next

j = -1850
For i = 6 To no_walls
    RotateMatrixY matWall(i), PI / 2
    TranslateMatrix matWall(i), MakeVector(910, 0, j), False
    j = j + 200
Next

TranslateMatrix matSWall(0), MakeVector(850, 0, -2060), True

TranslateMatrix matSWall(1), MakeVector(850, 0, 1960), True

RotateMatrixY matSWall(2), PI / 2
TranslateMatrix matSWall(2), MakeVector(910, 0, -2000), False

RotateMatrixY matSWall(3), PI / 2
TranslateMatrix matSWall(3), MakeVector(910, 0, 1900), False

TranslateMatrix matMDoor1(0), MakeVector(-100, 0, -2060), True
TranslateMatrix matMDoor1(1), MakeVector(100, 0, 1960), True
TranslateMatrix matMDoor2(0), MakeVector(100, 0, -2060), True
TranslateMatrix matMDoor2(1), MakeVector(-100, 0, 1960), True

End Sub

Public Sub PlaceBuildings()

TranslateMatrix matBuild(0), MakeVector(-250, 0, -150), True
TranslateMatrix matBuild(1), MakeVector(-250, 0, 350), True
TranslateMatrix matBuild(2), MakeVector(-250, 0, -650), True
TranslateMatrix matBuild(3), MakeVector(-250, 0, -1100), True
TranslateMatrix matBuild(4), MakeVector(-250, 0, -1600), True
TranslateMatrix matBuild(5), MakeVector(-250, 0, 800), True
TranslateMatrix matBuild(6), MakeVector(-250, 0, 1500), True

End Sub

Public Sub PlaceObjects()
Dim i As Integer

PolePos(0) = MakeVector(675, 0, -1500)
PolePos(1) = MakeVector(675, 0, -500)
PolePos(2) = MakeVector(675, 0, 500)
PolePos(3) = MakeVector(675, 0, 1500)

For i = 0 To no_poles
    TranslateMatrix matPole(i), PolePos(i), True
Next

TreePos(0) = MakeVector(425, 0, -1700)
TreePos(1) = MakeVector(350, 0, -1300)
TreePos(2) = MakeVector(550, 0, -700)
TreePos(3) = MakeVector(300, 0, -100)
TreePos(4) = MakeVector(450, 0, 300)
TreePos(5) = MakeVector(600, 0, 750)
TreePos(6) = MakeVector(400, 0, 1200)

End Sub

Public Sub SetLights()
Dim i As Integer

If Day Then
    Dim Light As D3DLIGHT8
    
    FogCol = D3DColorXRGB(80, 80, 80)
    D3DDevice.SetRenderState D3DRS_AMBIENT, D3DColorXRGB(255, 255, 255)
    
    Light.Type = D3DLIGHT_DIRECTIONAL
    Light.Position = MakeVector(0, 1000, 0)
    Light.Direction = MakeVector(0, -1, 0)
    With Light.diffuse
        .a = 1: .b = 1: .g = 1: .r = 1
    End With
    Light.Range = 1
    
    D3DDevice.SetLight 0, Light
    If Not UseFog Then D3DDevice.LightEnable 0, 1
Else
    Dim StreetLight(3) As D3DLIGHT8
    
    FogCol = D3DColorXRGB(50, 50, 50)
    D3DDevice.SetRenderState D3DRS_AMBIENT, D3DColorXRGB(150, 150, 150)
    
    StreetLight(0).Position = MakeVector(PolePos(0).X, 70, PolePos(0).z)
    StreetLight(1).Position = MakeVector(PolePos(1).X, 70, PolePos(1).z)
    StreetLight(2).Position = MakeVector(PolePos(2).X, 70, PolePos(2).z)
    StreetLight(3).Position = MakeVector(PolePos(3).X, 70, PolePos(3).z)

    For i = 0 To 3
        StreetLight(i).Type = D3DLIGHT_POINT
        With StreetLight(i).diffuse
            .a = 1: .b = 1: .g = 1: .r = 1
        End With
        StreetLight(i).Range = 200#
        StreetLight(i).Attenuation0 = 0.1
        StreetLight(i).Attenuation1 = 0#
        StreetLight(i).Attenuation2 = 0#

        D3DDevice.SetLight i, StreetLight(i)
        D3DDevice.LightEnable i, 1
    Next
End If

D3DDevice.SetRenderState D3DRS_FOGENABLE, UseFog 'set to 0 to disable
D3DDevice.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_NONE 'dont use table fog
D3DDevice.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_LINEAR 'use standard linear fog
D3DDevice.SetRenderState D3DRS_RANGEFOGENABLE, 1 'enable range based fog, hw dependent
D3DDevice.SetRenderState D3DRS_FOGSTART, FloatToDWord(1)
D3DDevice.SetRenderState D3DRS_FOGEND, FloatToDWord(FogRange)
D3DDevice.SetRenderState D3DRS_FOGCOLOR, FogCol

End Sub

Public Sub TranslateMatrix(ByRef Mat As D3DMATRIX, Pos As D3DVECTOR, Reset As Boolean)
Dim matTemp As D3DMATRIX

If Reset Then D3DXMatrixIdentity Mat
Mat.m41 = Pos.X
Mat.m42 = Pos.Y
Mat.m43 = Pos.z

End Sub

Public Sub RotateMatrixY(ByRef Mat As D3DMATRIX, Ang As Single)
Dim matTemp As D3DMATRIX

D3DXMatrixIdentity matTemp
D3DXMatrixIdentity Mat
D3DXMatrixRotationY matTemp, Ang
D3DXMatrixMultiply Mat, Mat, matTemp

End Sub

Public Function FloatToDWord(f As Single) As Long
    'this function packs a 32bit floating point number
    'into a 32bit integer number; quite slow - dont overuse.
    'DXCopyMemory or CopyMemory() (win32 api) would
    'probably be faster...
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, l
    FloatToDWord = l
End Function
