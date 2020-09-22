VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Fog"
      Height          =   2535
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   2175
      Begin VB.Frame Frame3 
         Caption         =   "Fog Density"
         Height          =   1575
         Left            =   480
         TabIndex        =   6
         Top             =   720
         Width           =   1575
         Begin VB.OptionButton optDensity 
            Caption         =   "Medium"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton optDensity 
            Caption         =   "Light"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   8
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton optDensity 
            Caption         =   "Dense"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkFog 
         Caption         =   "Use Fog"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Time"
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      Begin VB.OptionButton optDay 
         Caption         =   "Night"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optDay 
         Caption         =   "Day"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmStart"
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

Private Sub chkFog_Click()

If chkFog.Value = 0 Then
    Frame3.Enabled = False
    optDensity(0).Enabled = False
    optDensity(1).Enabled = False
    optDensity(2).Enabled = False
Else
    Frame3.Enabled = True
    optDensity(0).Enabled = True
    optDensity(1).Enabled = True
    optDensity(2).Enabled = True
End If

End Sub

Private Sub cmdOk_Click()

ModelPath = App.Path + "\models\"
TexturePath = App.Path + "\textures\"
SoundPath = App.Path + "\sound\"

If chkFog.Value = 0 Then
    UseFog = False
Else
    UseFog = True
    If optDensity(0).Value Then
        FogRange = 1500
    ElseIf optDensity(1).Value Then
        FogRange = 2000
    Else
        FogRange = 3000
    End If
End If

If optDay(0).Value Then
    Day = True
    SkyTexFile(0) = TexturePath + "sky_front_day.jpg"
    SkyTexFile(1) = TexturePath + "sky_back_day.jpg"
    SkyTexFile(2) = TexturePath + "sky_right_day.jpg"
    SkyTexFile(3) = TexturePath + "sky_left_day.jpg"
    SkyTexFile(4) = TexturePath + "sky_up_day.jpg"
    SkyTexFile(5) = TexturePath + "sky_down_day.jpg"
Else
    Day = False
    SkyTexFile(0) = TexturePath + "sky_front.jpg"
    SkyTexFile(1) = TexturePath + "sky_back.jpg"
    SkyTexFile(2) = TexturePath + "sky_right.jpg"
    SkyTexFile(3) = TexturePath + "sky_left.jpg"
    SkyTexFile(4) = TexturePath + "sky_up.jpg"
    SkyTexFile(5) = TexturePath + "sky_down.jpg"
End If

Main

End Sub

Private Sub Form_Load()
    optDay(0).Value = True
    optDensity(1).Value = True
End Sub
