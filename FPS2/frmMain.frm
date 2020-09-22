VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Implements DirectXEvent8 'for an event based system we need a callback function

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
Dim tmpAng As Single

On Error Resume Next
If Player.Dead Then Exit Sub

If Not (eventid = hEvent) Then Exit Sub

Dim DevData(1 To 10) As DIDEVICEOBJECTDATA 'storage for the event data
Dim nEvents As Long 'how many events have just happened (usually 1)
Dim i As Long 'looping variables
        
'1. retrieve the data from the device.
nEvents = DIDevice.GetDeviceData(DevData, DIGDD_DEFAULT)
        
'2. loop through all the events
For i = 1 To nEvents
    Select Case DevData(i).lOfs
        Case DIMOFS_X
            'the mouse has moved along the X Axis
            tmpAng = DevData(i).lData * 0.005
            Player.Rotation = Player.Rotation + tmpAng
            If Player.Rotation < 0 Then Player.Rotation = 2 * PI
            If Player.Rotation > 2 * PI Then Player.Rotation = 0
 
        Case DIMOFS_Y
            'the mouse has moved along the Y axis
            CamPitch = CamPitch + (DevData(i).lData * 0.005)
            If CamPitch < -PI / 4 Then CamPitch = -PI / 4
            If CamPitch > PI / 4 Then CamPitch = PI / 4
             
        Case DIMOFS_BUTTON0
            'the first (left) button has been pressed
            If DevData(i).lData <> 0 Then
                Fire = True
                sndShoot.Stop
                sndShoot.SetCurrentPosition 0
                sndShoot.Play DSBPLAY_DEFAULT
                FireTimer = GetTickCount
            End If
            
        Case DIMOFS_BUTTON1
            'the second (right) button has been pressed
            Dim Ang As Single
            If DevData(i).lData <> 0 Then
                Zoom = True
                Ang = PI / 10
            Else
                Zoom = False
                Ang = PI / 3
            End If
            D3DXMatrixPerspectiveFovLH matProj, Ang, 1, 1, 10000
            D3DDevice.SetTransform D3DTS_PROJECTION, matProj

        Case DIMOFS_BUTTON2
            'the third (middle usually) button has been pressed
    End Select
Next i
    
DoEvents 'let windows catch up on things...

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then UpKey = True
If KeyCode = vbKeyDown Then DownKey = True
If KeyCode = vbKeyLeft Then LeftKey = True
If KeyCode = vbKeyRight Then RightKey = True

If KeyCode = vbKeyS Then SKey = True
If KeyCode = vbKeyW Then WKey = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then UpKey = False
If KeyCode = vbKeyDown Then DownKey = False
If KeyCode = vbKeyLeft Then LeftKey = False
If KeyCode = vbKeyRight Then RightKey = False

If KeyCode = vbKeyEscape Then bRunning = False

If KeyCode = vbKeyS Then SKey = False
If KeyCode = vbKeyW Then WKey = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

DestroyApp

End Sub

