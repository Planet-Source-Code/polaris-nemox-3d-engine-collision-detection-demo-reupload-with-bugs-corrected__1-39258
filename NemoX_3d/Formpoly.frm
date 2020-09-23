VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'main engine
Dim Nemo As New NemoX


'for world datas and geometry
Dim MM As New cNemo_Mesh






'----------------------------------------
'Name: GameLoop
'----------------------------------------
Sub GameLoop()
Nemo.BackBuffer_ClearCOLOR = 0

Nemo.Camera_SetPosition Vector(0#, 8#, -8#), _
                                Vector(0#, 8#, 500#)
                                
Nemo.Camera_SetRotation 0, 0, 0



Dim VC As D3DVECTOR

    Nemo.Set_EngineTextureFilter NEMO_FILTER_BILINEAR
 
     
     Nemo.Set_ViewFrustum 10, 5500, PI / 8


'Nemo.Set_CullMode D3DCULL_NONE

Do
  
  DoEvents
  Call Me.GetKey
  If Nemo.Get_KeyPress(DIK_ESCAPE) Then GoTo End_it
 GLOB.SetUpFrustum
  
    
    'prevent device missing or badly created
    If Nemo.GetD3dDevice Is Nothing Then Exit Sub


   
Nemo.Begin3D
    



      'MM.Render2
      'check for collision
      If MM.CheckCollisionSliding(Nemo.Camera_GetPosition, VC, 15) Then Nemo.Camera_set_EYE VC 'Nemo.Camera_Recall
      MM.Render
     
     Nemo.Draw_Text "FPS=" + Str(Nemo.Framesperseconde), 10, 1
     
   
Nemo.End3D
 
  
Loop

End_it:
Form_Unload 0
End Sub








'----------------------------------------
'Name: GetKey
'
'this sub checks if player has pressed a key
'key Left,Right for camera rotation
'key Up,Down for moving
'key '+' '-' for moving verticaly
'key '8' '2' for rotate camera horizontaly
'----------------------------------------
Sub GetKey()


If Nemo.Get_KeyPress(NEMO_KEY_LEFT) Then _
    Nemo.Camera_Turn_Left 1 / 50
If Nemo.Get_KeyPress(NEMO_KEY_RIGHT) Then _
    Nemo.Camera_Turn_Right 1 / 50
    
    
    If Nemo.Get_KeyPress(NEMO_KEY_UP) Then _
    Nemo.Camera_Move_Foward 1
    If Nemo.Get_KeyPress(NEMO_KEY_RCONTROL) Then _
    Nemo.Camera_Move_Foward 5
   If Nemo.Get_KeyPress(NEMO_KEY_DOWN) Then _
    Nemo.Camera_Move_Backward 1



  If Nemo.Get_KeyPress(NEMO_KEY_ADD) Then _
    Nemo.Camera_Strafe_UP 1
If Nemo.Get_KeyPress(NEMO_KEY_SUBTRACT) Then _
    Nemo.Camera_Strafe_DOWN 1


 If Nemo.Get_KeyPress(NEMO_KEY_NUMPAD8) Then
    Nemo.Camera_Turn_UP 1 / 50
End If
If Nemo.Get_KeyPress(NEMO_KEY_NUMPAD2) Then _
    Nemo.Camera_Turn_DOWN 1 / 50
    
If Nemo.Get_KeyPress(NEMO_KEY_SPACE) Then _
 Nemo.Camera_SetPosition Vector(0#, 8#, -8#), _
                                Vector(0#, 8#, 500#)

If Nemo.Get_KeyPress(NEMO_KEY_S) Then _
 Nemo.Take_SnapShot App.Path + "\Shot.bmp"




End Sub



'----------------------------------------
'Name: Form_Load
'Object: Form
'Event: Load
'----------------------------------------
Private Sub Form_Load()
 Me.Show
  
  DoEvents ' Let the PC do what it has to do
  
  'InitD3D ' Initialize Direct3D
  
  'for choosing windowed mode
  Nemo.INIT_ShowDeviceDLG Me.hwnd
  
  'windowed mode
  'If Not Nemo.Initialize(Me.hwnd) Then
    'Form_Unload 0
  'End If
  
  
  GeoMetry
End Sub



'----------------------------------------
'Name: Geo
'----------------------------------------
Sub GeoMetry()
Dim P As Long

ReDim POLYGON(58) As NemoPOLYGON


MM.LoadMesh App.Path + "\MESH.nmsh"

 
 
 GameLoop
End Sub


'----------------------------------------
'Name: Form_Unload
'Object: Form
'Event: Unload
'----------------------------------------
Private Sub Form_Unload(Cancel As Integer)
 Nemo.Free
 
 End
End Sub




