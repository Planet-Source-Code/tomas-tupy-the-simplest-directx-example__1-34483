VERSION 5.00
Object = "{08216199-47EA-11D3-9479-00AA006C473C}#2.1#0"; "RMCONTROL.OCX"
Begin VB.Form Form1 
   Caption         =   "Simplest 3D (RMCanvas Example)"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Index           =   3
      Interval        =   1
      Left            =   3960
      Top             =   3000
   End
   Begin VB.Timer Timer2 
      Index           =   2
      Interval        =   1
      Left            =   3480
      Top             =   3000
   End
   Begin VB.Timer Timer2 
      Index           =   1
      Interval        =   1
      Left            =   3960
      Top             =   2520
   End
   Begin VB.Timer Timer2 
      Index           =   0
      Interval        =   1
      Left            =   3480
      Top             =   2520
   End
   Begin RMControl7.RMCanvas RMCanvas1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9551
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CarMesh As Direct3DRMMeshBuilder3  'MESHES
Dim GroundMesh As Direct3DRMMeshBuilder3

Dim Car As Direct3DRMFrame3
Dim Ground As Direct3DRMFrame3      'FRAMES

Dim MovSpeed 'movement speed
Const RotSpeed = 2 'rotational speed
Const PI = 3.1415926535

Private Sub Form_Load()

With RMCanvas1
    .StartWindowed
    .CameraFrame.SetPosition Nothing, 0, 1, 0 'Set Camera position (at the beginning of track)
    .Viewport.SetBack (9000) 'set the depth of vision possible
    .SceneFrame.SetSceneBackgroundRGB 0.2, 0.4, 1 'set the background color
    'Set Meshes
    Set GroundMesh = .D3DRM.CreateMeshBuilder()
    Set CarMesh = .D3DRM.CreateMeshBuilder()
    'Set Objects
    Set Car = .D3DRM.CreateFrame(.SceneFrame)
    Set Ground = .D3DRM.CreateFrame(.SceneFrame)
End With

    
'load car.x
CarMesh.LoadFromFile "car.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing  'Load Car file
Car.AddVisual CarMesh 'Make car
Car.SetPosition Nothing, 0, 3, 40 'Position of car (x(side to side), y(p down), z(Forwad Back))(at the beginning of track)
Car.SetMaterialMode D3DRMMATERIAL_FROMFRAME 'Material of Car
'Car.SetColorRGB 0, 0, 0 'Color of car (Black)


'load ground.x
GroundMesh.LoadFromFile "ground.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
Ground.AddScale D3DRMCOMBINE_BEFORE, 0.1, 0.001, 0.1 'Size of Ground
Ground.SetMaterialMode D3DRMMATERIAL_FROMFRAME 'tell the mesh to become whatever color I want it to, as opposed to having it be the color of the mesh
Ground.SetColorRGB 0, 1, 0 'make the ground green
Ground.AddVisual GroundMesh 'add the ground mesh to the frame

    
    
MovSpeed = 3 'set the movement speed

RMCanvas1.Update
End Sub

Private Sub Form_Resize()
RMCanvas1.Width = Form1.ScaleWidth - RMCanvas1.Left
RMCanvas1.Height = Form1.ScaleHeight - RMCanvas1.Top
End Sub

Private Sub RMCanvas1_KeyPress(KeyAscii As Integer)
Dim CamV As D3DVECTOR 'camera vector (position)
RMCanvas1.CameraFrame.GetPosition Nothing, CamV 'get the camera's position
   
'axal movement
If LCase(Chr(KeyAscii)) = "a" Then Call RMCanvas1.CameraFrame.SetPosition(RMCanvas1.CameraFrame, -MovSpeed, 0, 0)
If LCase(Chr(KeyAscii)) = "d" Then Call RMCanvas1.CameraFrame.SetPosition(RMCanvas1.CameraFrame, MovSpeed, 0, 0)
If LCase(Chr(KeyAscii)) = "w" Then Call RMCanvas1.CameraFrame.SetPosition(RMCanvas1.CameraFrame, 0, 0, MovSpeed)
If LCase(Chr(KeyAscii)) = "s" Then Call RMCanvas1.CameraFrame.SetPosition(RMCanvas1.CameraFrame, 0, 0, -MovSpeed)

'rotatation
If LCase(Chr(KeyAscii)) = "q" Then RMCanvas1.CameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -RotSpeed * PI / 180
If LCase(Chr(KeyAscii)) = "e" Then RMCanvas1.CameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, RotSpeed * PI / 180
End Sub

Private Sub Timer2_Timer(Index As Integer)
'makes sure the RMCanvas stays updated
RMCanvas1.Update
End Sub
