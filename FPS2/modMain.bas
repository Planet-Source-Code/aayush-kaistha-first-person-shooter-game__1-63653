Attribute VB_Name = "modMain"
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

Public Sub Main()

Dim LastUpdated As Long, matTemp As D3DMATRIX

frmStart.Hide
frmMain.Show

Initialize 'initialize the game

Fps.Last = GetTickCount
LastUpdated = GetTickCount

Gun = Gun1
Do While bRunning
    LastUpdated = GetTickCount
    If Not Player.Dead Then
        CheckKeys

        D3DXMatrixIdentity matCam
        D3DXMatrixRotationYawPitchRoll matCam, Player.Rotation, CamPitch, 0
        D3DXVec3TransformCoord EyeLookDir, MakeVector(0, 0, 1), matCam
    
        D3DXVec3Add EyeLookAt, Player.Pos, EyeLookDir
        D3DXMatrixLookAtLH matView, Player.Pos, EyeLookAt, MakeVector(0, 1, 0)
        D3DDevice.SetTransform D3DTS_VIEW, matView
    End If
    
    D3DXMatrixScaling matWorld, 1, 1, 1
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
    Render
    
    Fps.Count = Fps.Count + 1
    If ((GetTickCount - Fps.Last) >= 1000) Then
        Fps.Value = Fps.Count
        Fps.Count = 0
        Fps.Last = GetTickCount
    End If
    
    DoEvents
    
    Player.MoveSpeed = ((GetTickCount - LastUpdated) / 1000) * 250
Loop

DestroyApp

End Sub

Public Sub Render()
Dim i As Long, j As Integer, temp As D3DMATRIX
Dim tmpAng As Single

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, FogCol, 1#, 0

D3DDevice.BeginScene
        
    If Not UseFog Then DrawSky
    
    For i = 0 To floor_seg
        D3DDevice.SetTransform D3DTS_WORLD, matFloor(i)
        RenderXFile Floor
    
        D3DDevice.SetTransform D3DTS_WORLD, matLawn(i)
        RenderXFile Lawn
    
        D3DDevice.SetTransform D3DTS_WORLD, matHFloor(i)
        RenderXFile HardFloor
    Next
    
    
    For i = 0 To no_build
        D3DDevice.SetTransform D3DTS_WORLD, matBuild(i)
        RenderXFile Build(i)
    Next
    
    
    For i = 0 To no_walls
        D3DDevice.SetTransform D3DTS_WORLD, matWall(i)
        RenderXFile Wall
    Next
    
    For i = 0 To no_swalls
        D3DDevice.SetTransform D3DTS_WORLD, matSWall(i)
        RenderXFile SWall
    Next
    
    For i = 0 To 1
        D3DDevice.SetTransform D3DTS_WORLD, matMDoor1(i)
        RenderXFile MDoor(0)
        D3DDevice.SetTransform D3DTS_WORLD, matMDoor2(i)
        RenderXFile MDoor(1)
    Next

    If Not Day Then D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    For i = 0 To no_poles
        D3DDevice.SetTransform D3DTS_WORLD, matPole(i)
        RenderXFile Pole
    Next
    D3DDevice.SetRenderState D3DRS_LIGHTING, 1
    
    If Not Zoom Then DrawGun
    DrawTrees
    DrawCH
    If Fire Then If Not Zoom Then DrawFlash
        
    D3DDevice.SetRenderState D3DRS_FOGENABLE, 0
    D3DX.DrawText MainFont, &HFFFFFFFF, "FPS = " + Str(Fps.Value), MakeRect(10, 300, 10, 30), DT_TOP Or DT_LEFT
    D3DX.DrawText MainFont, &HFF00FF00, "Aayush Kaistha", MakeRect(10, 300, 40, 60), DT_TOP Or DT_LEFT
    D3DX.DrawText MainFont, &HFFFFFF00, "aayushkaistha@gmail.com", MakeRect(10, 300, 70, 90), DT_TOP Or DT_LEFT
    If UseFog Then D3DDevice.SetRenderState D3DRS_FOGENABLE, 1
    
D3DDevice.EndScene

D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Public Sub RenderXFile(ByRef tmpObj As Object3D)
'draws objects loaded from x files
Dim i As Long
    For i = 0 To tmpObj.nMaterials - 1
        D3DDevice.SetMaterial tmpObj.Materials(i)
        D3DDevice.SetTexture 0, tmpObj.Textures(i)
        tmpObj.Mesh.DrawSubset i
    Next
End Sub

Public Sub SortTrees()
Dim d1 As Single, d2 As Single
Dim t1 As Single, t2 As Single
Dim i As Integer, j As Integer
Dim tmp As D3DVECTOR, changed As Boolean

For j = 0 To no_trees - 1
    changed = False
    For i = 0 To no_trees - j - 1
        t1 = TreePos(i).X - Player.Pos.X
        t2 = TreePos(i).z - Player.Pos.z
        d1 = Sqr((t1 * t1) + (t2 * t2))
        t1 = TreePos(i + 1).X - Player.Pos.X
        t2 = TreePos(i + 1).z - Player.Pos.z
        d2 = Sqr((t1 * t1) + (t2 * t2))
        If (d1 < d2) Then
            tmp = TreePos(i)
            TreePos(i) = TreePos(i + 1)
            TreePos(i + 1) = tmp
            changed = True
        End If
    Next
    If Not changed Then Exit For
Next

End Sub

Public Sub DrawTrees()
Dim matTemp As D3DMATRIX, i As Integer, dis As Single

SortTrees
For i = 0 To no_trees

    RotateMatrixY matTree(i), Player.Rotation
    TranslateMatrix matTree(i), TreePos(i), False
    
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    D3DDevice.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL

    D3DDevice.SetTransform D3DTS_WORLD, matTree(i)
    D3DDevice.SetTexture 0, TreeTex
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, Tree(0), Len(Tree(0))

    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
Next

End Sub

Public Sub DrawGun()
Dim tmpAng As Single, temp As D3DMATRIX

If Fire Then
    If ((GetTickCount - FireTimer) <= 150) Then
        Gun = Gun2
    Else
        Fire = False
        Gun = Gun1
    End If
End If

tmpAng = Player.Rotation - (PI / 2)
D3DXMatrixIdentity temp
D3DXMatrixIdentity matGun
D3DXMatrixTranslation temp, Player.Pos.X + Sin(Player.Rotation) - (Sin(tmpAng) * 2), 67, Player.Pos.z + Cos(Player.Rotation) - (Cos(tmpAng) * 2)
D3DXMatrixMultiply matGun, matCam, temp

D3DDevice.SetTransform D3DTS_WORLD, matGun
RenderXFile Gun

End Sub

Public Sub DrawSky()

D3DXMatrixScaling matWorld, 1, 1, 1
D3DDevice.SetTransform D3DTS_WORLD, matWorld

D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
D3DDevice.SetRenderState D3DRS_ZENABLE, 0
D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 0

D3DDevice.SetTexture 0, SkyTex(0)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Sky(0), Len(Sky(0))

D3DDevice.SetTexture 0, SkyTex(1)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Sky(4), Len(Sky(0))

D3DDevice.SetTexture 0, SkyTex(2)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Sky(8), Len(Sky(0))

D3DDevice.SetTexture 0, SkyTex(3)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Sky(12), Len(Sky(0))

D3DDevice.SetTexture 0, SkyTex(4)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Sky(16), Len(Sky(0))

D3DDevice.SetTexture 0, SkyTex(5)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Sky(20), Len(Sky(0))

D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 1
D3DDevice.SetRenderState D3DRS_ZENABLE, 1
D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW

End Sub

Public Sub DrawCH()

D3DDevice.SetRenderState D3DRS_LIGHTING, 0
D3DDevice.SetRenderState D3DRS_FOGENABLE, 0
D3DDevice.SetVertexShader FVF_TLVERTEX
D3DDevice.SetTexture 0, Nothing

D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 4, CH(0), Len(CH(0))

D3DDevice.SetRenderState D3DRS_LIGHTING, 1
If UseFog Then D3DDevice.SetRenderState D3DRS_FOGENABLE, 1
D3DDevice.SetVertexShader FVF_VERTEX

End Sub

Public Sub DrawFlash()
Dim matTemp As D3DMATRIX, tmpAng As Single

tmpAng = Player.Rotation - (PI / 2)
D3DXMatrixIdentity matTemp
D3DXMatrixIdentity matFlash
D3DXMatrixTranslation matTemp, Player.Pos.X + Sin(Player.Rotation) - (Sin(tmpAng) * 2), 68, Player.Pos.z + Cos(Player.Rotation) - (Cos(tmpAng) * 2)
D3DXMatrixMultiply matFlash, matCam, matTemp

D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

D3DDevice.SetRenderState D3DRS_LIGHTING, 0
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
D3DDevice.SetVertexShader FVF_LVERTEX
D3DDevice.SetTexture 0, FlashTex

D3DDevice.SetTransform D3DTS_WORLD, matFlash
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Flash(0), Len(Flash(0))

D3DDevice.SetRenderState D3DRS_LIGHTING, 1
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
D3DDevice.SetVertexShader FVF_VERTEX

D3DDevice.SetTransform D3DTS_WORLD, matWorld
D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW

End Sub

Public Sub CheckKeys()
Dim tmpAng As Integer

If UpKey Then
    'If Colliding(Player.Pos.X + (Sin(Player.Rotation) * 10), Player.Pos.Z + (Cos(Player.Rotation) * Player.MoveSpeed)) = False Then
        Player.Pos.X = Player.Pos.X + (Sin(Player.Rotation) * Player.MoveSpeed)
    'If Colliding(Player.Pos.X + (Sin(Player.Rotation) * Player.MoveSpeed), Player.Pos.Z + (Cos(Player.Rotation) * 10)) = False Then
        Player.Pos.z = Player.Pos.z + (Cos(Player.Rotation) * Player.MoveSpeed)
End If
If DownKey Then
    'If Colliding(Player.Pos.X - (Sin(Player.Rotation) * 10), Player.Pos.Z - (Cos(Player.Rotation) * Player.MoveSpeed)) = False Then
        Player.Pos.X = Player.Pos.X - (Sin(Player.Rotation) * Player.MoveSpeed)
    'If Colliding(Player.Pos.X - (Sin(Player.Rotation) * Player.MoveSpeed), Player.Pos.Z - (Cos(Player.Rotation) * 10)) = False Then
        Player.Pos.z = Player.Pos.z - (Cos(Player.Rotation) * Player.MoveSpeed)
End If
If LeftKey Then
    tmpAng = Player.Rotation + (PI / 2)
    'If Colliding(Player.Pos.X - (Sin(tmpAng) * 10), Player.Pos.Z - (Cos(tmpAng) * Player.MoveSpeed)) = False Then
        Player.Pos.X = Player.Pos.X - (Sin(tmpAng) * Player.MoveSpeed / 2)
    'If Colliding(Player.Pos.X - (Sin(tmpAng) * Player.MoveSpeed), Player.Pos.Z - (Cos(tmpAng) * 10)) = False Then
        Player.Pos.z = Player.Pos.z - (Cos(tmpAng) * Player.MoveSpeed / 2)
End If
If RightKey Then
    tmpAng = Player.Rotation - (PI / 2)
    'If Colliding(Player.Pos.X - (Sin(tmpAng) * 10), Player.Pos.Z - (Cos(tmpAng) * Player.MoveSpeed)) = False Then
        Player.Pos.X = Player.Pos.X - (Sin(tmpAng) * Player.MoveSpeed / 2)
    'If Colliding(Player.Pos.X - (Sin(tmpAng) * Player.MoveSpeed), Player.Pos.Z - (Cos(tmpAng) * 10)) = False Then
        Player.Pos.z = Player.Pos.z - (Cos(tmpAng) * Player.MoveSpeed / 2)
End If

If Player.Pos.X > 850 Then Player.Pos.X = 850
If Player.Pos.X < -150 Then Player.Pos.X = -150
If Player.Pos.z > 1900 Then Player.Pos.z = 1900
If Player.Pos.z < -2000 Then Player.Pos.z = -2000

If SKey Then D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
If WKey Then D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME

End Sub

Public Sub DestroyApp()
On Error Resume Next

If hEvent <> 0 Then DX.DestroyEvent hEvent
Set DIDevice = Nothing
Set DI = Nothing

Set D3DX = Nothing
Set D3DDevice = Nothing
Set D3D = Nothing
Set DX = Nothing

ShowCursor 1

End

End Sub
