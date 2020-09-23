VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NemoX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'=====================
'
'18-03-2002 Beginning of the development
'    -Add main properties
'    -Add inititialization
'    -Add Input Keyboard and Input
'    -Add clearBuffer for
'    -Add keyevent
'
'19-03-2002
'    -Add Flip
'    -Frames Perseconde conter
'    -Add world transformation routines
'
'    -Add a Xmesh class
'
'
'20-03-2002
'    -Add Camera
'    -Strafe left/right/down/up
'    -Move forward/backward
'    -Roll Left/right
'21-03-2002
'
'22-03-2002
'
'    -Add A particle engine class
'    -Add Fog built in fog control
'
'24-03-2002
'
'    -Add Textdrawing tools
'
'27-03-2002
'
'    -Add Good skybox texture
'    -Rebuild of landscape engine
'    -Add extern filtering mode support
'    -Add extern fill mode support
'
'    -Add Ray Picking
'    -ADD Landscape class
'
'
'29-03-2002
'
'    -Add Device enumeration
'
'30-03-2002
'
'    -Add SnapShop
'
'3-04-2002
'    -Add Drawtext
'    -OBJ models loading

'6-04-2002
'    -Start  of PlaneAI class for any Flying objects
'8-04-2002
'    -Add gamma control
'    -Add sphere mapping
'11-04-2002
'    -Add New effect for particle system,fountain,smoke
'
''14-04-2002
'    -Add New effect for particle system Rain and snow

'17-04-2002
'    -Add New effect for particle system now good Explosion
'18-04-2002
'    -Add New effect for particle
'       Now Effect are fast and accurate
'       60FPS per 1000 particle
'20-04-2002
'    -Add New Lensflare class
'
'24-05-2002
'    -Add New effect for particle
'       Now milkshape 3d models are supported
'
'12-08-2002
'    -Add Q3level support with lighmap
'    -Add multitexturing quick acces
'    -Add milshpae 3d texture environment
'
'14-08-2002
'    -Add a new 3ds class with material

'
'24-08-2002
'    -Engine hierarchie has changed The 3d interface are global
'


Option Explicit
Private Const version = 0.06

'This project needs 2 Buttons
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursor Lib "user32" () As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long


Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

'========Main 3d Object==============
Private lpD3dDevice As Direct3DDevice8
Private lpDX8 As DirectX8
Private lpD3D As Direct3D8
Private lpD3DX As D3DX8

Private Nemo_DispModeParam As D3DDISPLAYMODE

'========Input========
'keyboard
Private DI As DirectInput8 'this is DirectInput, used to monitor the keys on the keyboard in my case
Private DInputDevice As DirectInputDevice8 'this device will be the keyboard
Private DIState As DIKEYBOARDSTATE 'to check the state of keys

Public Enum NEMO_KEY_CONST

    NEMO_KEY_0 = DIK_0
    NEMO_KEY_1 = DIK_1
    NEMO_KEY_2 = DIK_2
    NEMO_KEY_3 = DIK_3
    NEMO_KEY_4 = DIK_4
    NEMO_KEY_5 = DIK_5
    NEMO_KEY_6 = DIK_6
    NEMO_KEY_7 = DIK_7
    NEMO_KEY_8 = DIK_8
    NEMO_KEY_9 = DIK_9
    NEMO_KEY_A = DIK_A
    'NEMO_KEY_ABNT_C1 = DIK_ABNT_C1
    'NEMO_KEY_ABNT_C2 = DIK_ABNT_C2
    NEMO_KEY_ADD = DIK_ADD
    NEMO_KEY_APOSTROPHE = DIK_APOSTROPHE
    NEMO_KEY_APPS = DIK_APPS
    NEMO_KEY_AT = DIK_AT
    NEMO_KEY_AX = DIK_AX
    NEMO_KEY_B = DIK_B
    NEMO_KEY_BACKSPACE = DIK_BACK
    NEMO_KEY_BACKSLASH = DIK_BACKSLASH
    NEMO_KEY_C = DIK_C
    'NEMO_KEY_CALCULATOR = DIK_CALCULATOR
    NEMO_KEY_CAPITAL = DIK_CAPITAL
    'NEMO_KEY_CAPSLOCK = DIK_CAPSLOCK
    NEMO_KEY_CIRCUMFLEX = DIK_CIRCUMFLEX
    NEMO_KEY_COLON = DIK_COLON
    NEMO_KEY_COMMA = DIK_COMMA
    NEMO_KEY_CONVERT = DIK_CONVERT
    NEMO_KEY_D = DIK_D
    NEMO_KEY_DECIMAL = DIK_DECIMAL
    NEMO_KEY_DELETE = DIK_DELETE
    NEMO_KEY_DIVIDE = DIK_DIVIDE
    NEMO_KEY_DOWN = DIK_DOWN
    'NEMO_KEY_DOWNARROW = DIK_DOWNARROW
    NEMO_KEY_E = DIK_E
    NEMO_KEY_END = DIK_END
    NEMO_KEY_EQUALS = DIK_EQUALS
    NEMO_KEY_ESCAPE = DIK_ESCAPE
    NEMO_KEY_F = DIK_F
    NEMO_KEY_F1 = DIK_F1
    NEMO_KEY_F2 = DIK_F2
    NEMO_KEY_F3 = DIK_F3
    NEMO_KEY_F4 = DIK_F4
    NEMO_KEY_F5 = DIK_F5
    NEMO_KEY_F6 = DIK_F6
    NEMO_KEY_F7 = DIK_F7
    NEMO_KEY_F8 = DIK_F8
    NEMO_KEY_F9 = DIK_F9
    NEMO_KEY_F10 = DIK_F10
    NEMO_KEY_F11 = DIK_F11
    NEMO_KEY_F12 = DIK_F12
    NEMO_KEY_F13 = DIK_F13
    NEMO_KEY_F14 = DIK_F14
    NEMO_KEY_F15 = DIK_F15
    NEMO_KEY_G = DIK_G
    NEMO_KEY_GRAVE = DIK_GRAVE
    NEMO_KEY_H = DIK_H
    NEMO_KEY_HOME = DIK_HOME
    NEMO_KEY_I = DIK_I
    NEMO_KEY_INSERT = DIK_INSERT
    NEMO_KEY_J = DIK_J
    NEMO_KEY_K = DIK_K
    NEMO_KEY_KANA = DIK_KANA
    NEMO_KEY_KANJI = DIK_KANJI
    NEMO_KEY_L = DIK_L
    'NEMO_KEY_ALT_LEFT = DIK_LALT
    NEMO_KEY_LEFTBRACKET = DIK_LBRACKET
    NEMO_KEY_LEFTCONTROL = DIK_LCONTROL
    NEMO_KEY_LEFT = DIK_LEFT
    'NEMO_KEY_LEFTARROW = DIK_LEFTARROW
    NEMO_KEY_LEFTMENU = DIK_LMENU
    NEMO_KEY_LEFTSHIFT = DIK_LSHIFT
    NEMO_KEY_LEFTWINDOWS = DIK_LWIN
    NEMO_KEY_M = DIK_M
    'NEMO_KEY_MAIL = DIK_MAIL
    'NEMO_KEY_MEDIASELECT = DIK_MEDIASELECT
    'NEMO_KEY_MEDIASTOP = DIK_MEDIASTOP
    NEMO_KEY_MINUS = DIK_MINUS
    NEMO_KEY_MULTIPLY = DIK_MULTIPLY
    'NEMO_KEY_MUTE = DIK_MUTE
    'NEMO_KEY_MYCOMPUTER = DIK_MYCOMPUTER
    NEMO_KEY_N = DIK_N
    NEMO_KEY_NEXT = DIK_NEXT
    'NEMO_KEY_NEXTTRACK = DIK_NEXTTRACK
    NEMO_KEY_NOCONVERT = DIK_NOCONVERT
    NEMO_KEY_NUMLOCK = DIK_NUMLOCK
    NEMO_KEY_NUMPAD0 = DIK_NUMPAD0
    NEMO_KEY_NUMPAD1 = DIK_NUMPAD1
    NEMO_KEY_NUMPAD2 = DIK_NUMPAD2
    NEMO_KEY_NUMPAD3 = DIK_NUMPAD3
    NEMO_KEY_NUMPAD4 = DIK_NUMPAD4
    NEMO_KEY_NUMPAD5 = DIK_NUMPAD5
    NEMO_KEY_NUMPAD6 = DIK_NUMPAD6
    NEMO_KEY_NUMPAD7 = DIK_NUMPAD7
    NEMO_KEY_NUMPAD8 = DIK_NUMPAD8
    NEMO_KEY_NUMPAD9 = DIK_NUMPAD9
    NEMO_KEY_NUMPADCOMMA = DIK_NUMPADCOMMA
    NEMO_KEY_NUMPADENTER = DIK_NUMPADENTER
    NEMO_KEY_NUMPADEQUALS = DIK_NUMPADEQUALS

    NEMO_KEY_O = DIK_O

    NEMO_KEY_P = DIK_P
    NEMO_KEY_PAUSE = DIK_PAUSE
    NEMO_KEY_PERIOD = DIK_PERIOD

    NEMO_KEY_Q = DIK_Q
    NEMO_KEY_R = DIK_R
    'NEMO_KEY_ALT_RIGHT = DIK_RALT
    NEMO_KEY_RIGHTBRACKET = DIK_RBRACKET
    NEMO_KEY_RCONTROL = DIK_RCONTROL
    NEMO_KEY_RETURN = DIK_RETURN
    NEMO_KEY_RIGHT = DIK_RIGHT
    'NEMO_KEY_RIGHTARROW = DIK_RIGHTARROW
    NEMO_KEY_RIGHTMENU = DIK_RMENU
    NEMO_KEY_RIGHTSHIFT = DIK_RSHIFT
    NEMO_KEY_RWIN = DIK_RWIN
    NEMO_KEY_S = DIK_S
    NEMO_KEY_SCROLL = DIK_SCROLL
    NEMO_KEY_SEMICOLON = DIK_SEMICOLON
    NEMO_KEY_SLASH = DIK_SLASH
    'NEMO_KEY_SLEEP = DIK_SLEEP
    NEMO_KEY_SPACE = DIK_SPACE
    NEMO_KEY_STOP = DIK_STOP
    NEMO_KEY_SUBTRACT = DIK_SUBTRACT
    NEMO_KEY_SYSRQ = DIK_SYSRQ
    NEMO_KEY_T = DIK_T
    NEMO_KEY_TAB = DIK_TAB
    NEMO_KEY_U = DIK_U
    NEMO_KEY_UNDERLINE = DIK_UNDERLINE
    NEMO_KEY_UNLABELED = DIK_UNLABELED
    NEMO_KEY_UP = DIK_UP
    'NEMO_KEY_UPARROW = DIK_UPARROW
    NEMO_KEY_V = DIK_V

    NEMO_KEY_W = DIK_W

    NEMO_KEY_X = DIK_X
    NEMO_KEY_Y = DIK_Y
    NEMO_KEY_YEN = DIK_YEN
    NEMO_KEY_Z = DIK_Z

End Enum

'for Capabilities

Enum NEMO_FILTERING_MODE
    NEMO_FILTER_POINT
    NEMO_FILTER_BILINEAR
    NEMO_FILTER_TRILINEAR
    NEMO_FILTER_ANISOTROPIC
End Enum

Enum NEMO_SHADING_MODE
    NEMO_SHADE_FLAT = D3DSHADE_FLAT
    NEMO_SHADE_GOURAUD = D3DSHADE_GOURAUD
    NEMO_SHADE_PHONG = D3DSHADE_PHONG

End Enum

Enum NEMO_FILL_MODE
    NEMO_FILL_WIREFRAME
    NEMO_FILL_SOLID

End Enum

'Mouse




Private DIMouse As DirectInputDevice8 ' Mouse device
Private DIMState As DIMOUSESTATE ' to check mouse movements and clicks

Enum NEMO_MOUSE_BUTTON

    NK_LEFT_Button = 0
    NK_RIGHT_Button = 1
    NK_MIDDLE_Button = 2

End Enum

'**************************
'*  for the frames per seconde calcul
'*************************
Private Iclip_count
Private tLast As Single
Private FPS As Single

Private Ang As Single, NemoFPS_2 As Single, OldFPS As Long, NBImage As Integer
Private OldTime
Private NBFrame

Dim NemoFPS_LastCheck As Long
Dim NemoFPS_Count As Long
Dim NemoFPS_Current As Integer
Dim LastTimeUpdatedMatrix As Long
Dim LastUpdatedParticles As Long


Dim Dauer, Takt
Dim FPS_Current As Single            'Unsere FPS
Dim FPS_timefactor As Currency       'Zeit die vergangen ist seid dem
                                        'der letzte Frame dargestellt wurde
                                        'timefactor lässt sich zum Zeit basierten steuern von Abläufen
                                        'einsetzen (wenn eine Taste gedrückt wurde kann man
                                        'ermitteln wie lange sie schon gedrückt wurde und daraus die
                                        'stärke der Auswirkung errechnen)
Dim FPS_tLastFrame As Currency       'Speichert die System-Zeit als der letzte Frame gezeichnet wurde
Dim FPS_tLastFPSdisplay As Currency  'Speichert zeit der letzten FPS-Berchnung
Dim FPS_Counter As Currency          'Zählt die Frames





'===========================================
'   For the camera
'================================

Private Type CAM
    EYE As D3DVECTOR
    RotVec As D3DVECTOR
    DEG As Single
    ANGx As Single
    ANGy As Single
    ANGz As Single
    Dest_at As D3DVECTOR
    Mouse_Prev_X As Long
    Mouse_Prev_Y As Long

End Type

Private Type NEMO_CAPABILITIES
    Filter_Bilinear As Boolean
    Filter_Trilinear As Boolean
    Filter_Anisotropic As Boolean
    CanDoMultiTexture As Boolean

    TandL_Device As Boolean
    BumpMapping As Boolean
End Type

Private Type NEMO_Frustrum
    Near As Single
    Far As Single
    FovAngle As Single
    Aspect As Single

End Type

Private Type NemoCFG
    lpState As Nemo_SaveState
    Buffer_Width As Integer
    Buffer_Height As Integer
    Is_Windowed As Boolean
    BPP As Integer
    Is_engineActive As Boolean
    BackBuff_ClearColor As Long
    MatProjec As D3DMATRIX
    matView As D3DMATRIX
    MatSphereMapping As D3DMATRIX
    EYES As CAM
    SaveCAM As CAM
    MainFont As D3DXFont
    StFont As StdFont
    FontDesc As IFont
    Capa As NEMO_CAPABILITIES
    ViewFrust As NEMO_Frustrum
    
    Hwindow As Long
    Device As NEMO_DEVICE_TYPE
    GamaLevel As Single
End Type

'#########
'## FONTS ##
'#########
Dim MainFont As D3DXFont
Dim MainFontDesc As IFont
Dim TextRect As RECT
Dim fnt As New StdFont

Enum EnnFrustumSide

    frustRight = 0  ''       ''//The RIGHT side of the frustum
    frustLeft = 1   ''       ''//The LEFT  side of the frustum
    frustbottom = 2 ''       ''//The BOTTOM side of the frustum
    frustTop = 3    ''       ''//The TOP side of the frustum
    frustBack = 4   ''       ''//The BACK side of the frustum
    frustFront = 5           '''//The FRONT side of the frustum
End Enum

'''//Like above''instead of saying a number for the ABC and D of the plane''we
'''//want to be more descriptive.
Enum EnnPlaneData

    FrustA = 0 ''             ''//The X value of the plane's normal
    Frustb = 1 ''             ''//The Y value of the plane's normal
    Frustc = 2 ''             ''//The Z value of the plane's normal
    frustd = 3               '''//The distance the plane is from the origin
End Enum

Dim m_Frustum(6, 4) As Single

Private Data As NemoCFG


'----------------------------------------
'Name: Set_Font
'Object: Set
'Event: Font
'----------------------------------------
Sub Set_Font(FontName As String, PoliceSize As Integer)

    Set Data.StFont = New StdFont

    Data.StFont.Name = FontName
    Data.StFont.size = PoliceSize
    Set Data.FontDesc = Data.StFont

    Set Data.MainFont = D3DX.CreateFont(D3dDevice, Data.FontDesc.hFont)

End Sub


'----------------------------------------
'Name: Draw_Text
'Object: Draw
'Event: Text
'----------------------------------------
Sub Draw_Text(ST As String, x, y, Optional color As Long = &HFFFFCC00)

  Dim RC As RECT

    TextRect.Top = y

    TextRect.Left = x

    'disable fog if anable
    If D3dDevice.GetRenderState(D3DRS_FOGENABLE) Then
        D3dDevice.SetRenderState D3DRS_FOGENABLE, 0
        D3DX.DrawText Data.MainFont, color, ST, TextRect, DT_TOP Or DT_LEFT
        D3dDevice.SetRenderState D3DRS_FOGENABLE, 1
      Else 'D3dDevice.GETRENDERSTATE(D3DRS_FOGENABLE) = FALSE
        D3DX.DrawText Data.MainFont, color, ST, TextRect, DT_TOP Or DT_LEFT
    End If

End Sub


'----------------------------------------
'Name: PrintTextGDI
'----------------------------------------
Sub PrintTextGDI(x, y, TXT As String)

    TextOut GetDC(Data.Hwindow), x, y, TXT, Len(TXT)

End Sub


'----------------------------------------
'Name: INITIALIZE_Windowed
'Object: INITIALIZE
'Event: Windowed
'----------------------------------------
Function INITIALIZE_Windowed(Hand As Long, Optional DeviceType As NEMO_DEVICE_TYPE = NEMO_HAL_DEVICE, Optional UseTnL As Boolean = 0) As Boolean

    Initialize Hand, True, , , , UseTnL

End Function


'----------------------------------------
'Name: INITIALIZE_FullScreen
'Object: INITIALIZE
'Event: FullScreen
'----------------------------------------
Function INITIALIZE_FullScreen(Hand As Long, Optional DeviceType As NEMO_DEVICE_TYPE = NEMO_HAL_DEVICE, Optional UseTnL As Boolean = 0, Optional Wi = 640, Optional HI = 480, Optional BPP = 16, Optional VertSincOFF As Boolean = True, Optional GAMA = 1) As Boolean

    Initialize Hand, DeviceType, False, Wi, HI, BPP, UseTnL, VertSincOFF, GAMA

End Function


'----------------------------------------
'Name: Initialize
'----------------------------------------
Function Initialize(Hand As Long, Optional DeviceType As NEMO_DEVICE_TYPE = NEMO_HAL_DEVICE, Optional Windowed As Boolean = 1, Optional Wi = 640, Optional HI = 480, Optional BPP = 16, Optional TnL As Boolean = False, Optional VertSincOFF As Boolean = True, Optional GamaLevel = 1, Optional NumBUFFERS = 1) As Boolean

    On Local Error Resume Next
        Set DX8 = New DirectX8
        Set D3DX = New D3DX8
        Set D3D = DX8.Direct3DCreate() ' Create D3D
        If D3D Is Nothing Then GoTo D3DError

      Dim DispMode As D3DDISPLAYMODE '//Describes our Display Mode':( Move line to top of current Function
      Dim D3DWindow As D3DPRESENT_PARAMETERS '//Describes our Viewport':( Move line to top of current Function

        'Set DX8 = New DirectX8  '//Create our Master Object
        'Set D3D = dx8.Direct3DCreate() '//Make our Master Object create the Direct3D Interface
        'Set D3DX = New D3DX8 '//Create our helper library...

        Data.Hwindow = Hand
        'DispMode.Format = D3DFMT_X8R8G8B8
        DispMode.format = D3DFMT_R5G6B5 'If this mode doesn't work try the commented one above...

        'd3d.GetAdapterDisplayMode 0, DispMode

        If Not Windowed Then
            DispMode.width = Wi
            DispMode.Height = HI
            D3DWindow.SwapEffect = D3DSWAPEFFECT_DISCARD
            'D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE 'or D3DPRESENT_INTERVAL_ONE
            If VertSincOFF Then
                D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
              Else 'VERTSINCOFF = FALSE
                D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_DEFAULT
            End If
          Else 'NOT NOT...
            'Dim RectClient As RECT
            'GetWindowRect Hand, RectClient
            'DispMode.width = RectClient.Right - RectClient.Left
            'DispMode.Height = RectClient.bottom - RectClient.Top
            'D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_DEFAULT
            
            
                D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_DEFAULT
           
            
            D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY

        End If

        D3DWindow.BackBufferCount = NumBUFFERS '//1 backbuffer only
        D3DWindow.BackBufferWidth = DispMode.width
        D3DWindow.BackBufferHeight = DispMode.Height
        D3DWindow.hDeviceWindow = Hand
        D3DWindow.Windowed = Windowed

        Nemo_DispModeParam = DispMode

      Dim Hbpp ':( Move line to top of current Function
      Dim DevType ':( Move line to top of current Function

        If BPP = 16 Then
            D3DWindow.BackBufferFormat = DispMode.format 'What we specified earlier
            Hbpp = D3DFMT_D16
          ElseIf BPP = 32 Then 'NOT BPP...
      Dim Mode As D3DDISPLAYMODE ':( Move line to top of current Function
            D3D.GetAdapterDisplayMode 0, Mode

            D3DWindow.BackBufferFormat = Mode.format
            Hbpp = D3DFMT_D16
        End If

        If DeviceType = NEMO_HAL_DEVICE Or DeviceType = NEMO_TnT_DEVICE Then DevType = D3DDEVTYPE_HAL
        If DeviceType = NEMO_HEL_REF Then DevType = D3DDEVTYPE_REF

        If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, DevType, DispMode.format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, Hbpp) = D3D_OK Then
            'We can use a 16 bit Z-Buffer
            D3DWindow.AutoDepthStencilFormat = Hbpp '//16 bit Z-Buffer
            D3DWindow.EnableAutoDepthStencil = 1 '//Only enable it if we can actually use a z-buffer
          Else 'NOT d3d.CHECKDEVICEFORMAT(D3DADAPTER_DEFAULT,...
            'We could now check for different modes available...
        End If

        '//This line creates a device that uses a hardware device if possible; software vertex processing and uses the form as it's target

        If TnL Then
            Set D3dDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, DevType, Hand, D3DCREATE_HARDWARE_VERTEXPROCESSING, _
                D3DWindow)
          Else 'TNL = FALSE
            Set D3dDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, DevType, Hand, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                D3DWindow)
        End If

        If D3dDevice Is Nothing Then GoTo D3DError

        'D3dDevice.SetRenderState D3DRS_DITHERENABLE, 1

        ' Turn on the zbuffer
        D3dDevice.SetRenderState D3DRS_ZENABLE, 1

        'D3dDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
        'D3dDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP

        Me.Set_EngineTextureFilter NEMO_FILTER_BILINEAR
        Me.Set_EngineFillMode NEMO_FILL_SOLID
        ' Set up a material. The material here just has the diffuse and ambient
        ' colors set to yellow. Note that only one material can be used at a time.
      Dim COL As D3DCOLORVALUE ':( Move line to top of current Function
      Dim mtrl As D3DMATERIAL8 ':( Move line to top of current Function
        With COL

            .R = 1
            .G = 1
            .B = 1
            .A = 1

        End With 'COL
        mtrl.diffuse = COL
        mtrl.Ambient = COL
        D3dDevice.SetMaterial mtrl
        Me.Set_LightAmbient &HFFFFFFFF

        'setup the projection
        'setup the projection matrix

        Me.Set_ViewFrustum 1, 9800, PI / 8, 1
        D3dDevice.SetRenderState D3DRS_LIGHTING, True

        With Data
            If Windowed Then

            End If

            .Buffer_Height = HI
            .Buffer_Width = Wi
            .Is_Windowed = Windowed
            .Is_engineActive = True
            .BPP = BPP
            .Device = DeviceType
            .GamaLevel = GamaLevel

        End With 'DATA

        Initialize = True
        Call Init_Input(Hand)

        'set default font
        Set_Font "Arial", 12
        Call Get_Capabilities

        'create spheremapping environnment
      Dim Mat As D3DMATRIX ':( Move line to top of current Function
        Data.MatSphereMapping.m11 = 0.5
        Data.MatSphereMapping.m12 = 0#
        Data.MatSphereMapping.m13 = 0#
        Data.MatSphereMapping.m14 = 0#
        Data.MatSphereMapping.m21 = 0#
        Data.MatSphereMapping.m22 = -0.5
        Data.MatSphereMapping.m23 = 0#
        Data.MatSphereMapping.m24 = 0#
        Data.MatSphereMapping.m31 = 0#
        Data.MatSphereMapping.m32 = 0#
        Data.MatSphereMapping.m33 = 1#
        Data.MatSphereMapping.m34 = 0#
        Data.MatSphereMapping.m41 = 0.5
        Data.MatSphereMapping.m42 = 0.5
        Data.MatSphereMapping.m43 = 0#
        Data.MatSphereMapping.m44 = 1#

        'set gamma level
        ChangeGamma Data.GamaLevel

        '
        Set GLOB.LpGLOBAL_NEMO = Me

    Exit Function

D3DError:

        MsgBox "Unable to CreateDevice ", vbCritical, "Error"

End Function ':( On Error Resume still active


'----------------------------------------
'Name: INIT_ShowDeviceDLG
'Object: INIT
'Event: ShowDeviceDLG
'----------------------------------------
Sub INIT_ShowDeviceDLG(appHandle As Long)

    frmMain.SHOW_DIALOG appHandle, Me

End Sub



'----------------------------------------
'Name: GetD3DX
'----------------------------------------
Function GetD3DX() As D3DX8
   Set GetD3DX = D3DX
End Function


'----------------------------------------
'Name: GetD3D8
'----------------------------------------
Function GetD3D8() As Direct3D8
   Set GetD3D8 = D3D
End Function


'----------------------------------------
'Name: GetD3dDevice
'----------------------------------------
Function GetD3dDevice() As Direct3DDevice8
   Set GetD3dDevice = D3dDevice
End Function



'----------------------------------------
'Name: ChangeGamma
'----------------------------------------
Private Sub ChangeGamma(gamma As Single)

  Dim A As D3DGAMMARAMP, i, G

    D3dDevice.GetGammaRamp A
  Dim T As Long
    For i = 0 To 255
        T = i * gamma * 255
        G = 0
        If T > 32767 Then G = 1: T = T - 32767
        If T > 32767 Then T = 32767
        'We need this bit thing to handle unsigned integer
        A.red(i) = T Or (&H8000 * G)
        A.green(i) = T Or (&H8000 * G)
        A.blue(i) = T Or (&H8000 * G)

    Next i

    D3dDevice.SetGammaRamp 0, A

End Sub












'----------------------------------------
'Name: Mtransform
'----------------------------------------
Private Function Mtransform(Matrix As D3DMATRIX, V As D3DVECTOR) As D3DVECTOR

    D3DXVec3TransformCoord Mtransform, V, Matrix

End Function




'----------------------------------------
'Name: Get_EngineVersion
'Object: Get
'Event: EngineVersion
'----------------------------------------
Function Get_EngineVersion() As Single

    Get_EngineVersion = version

End Function


'----------------------------------------
'Name: Get_Matrix
'Object: Get
'Event: Matrix
'----------------------------------------
Function Get_Matrix(RetMatrix As D3DMATRIX, MatTYPE As CONST_D3DTRANSFORMSTATETYPE)

    Call D3dDevice.GetTransform(MatTYPE, RetMatrix)

End Function


'----------------------------------------
'Name: Get_MatrixRet
'Object: Get
'Event: MatrixRet
'----------------------------------------
Function Get_MatrixRet(MatTYPE As CONST_D3DTRANSFORMSTATETYPE) As D3DMATRIX

    Call D3dDevice.GetTransform(MatTYPE, Get_MatrixRet)

End Function


'----------------------------------------
'Name: Get_Material
'Object: Get
'Event: Material
'----------------------------------------
Function Get_Material(Amb As D3DCOLORVALUE, Diff As D3DCOLORVALUE) As D3DMATERIAL8

    Get_Material.Ambient = Amb
    Get_Material.diffuse = Diff
    Get_Material.emissive = Amb
    Get_Material.Specular = Diff
    Get_Material.power = 1

End Function


'----------------------------------------
'Name: Set_MousePOS
'Object: Set
'Event: MousePOS
'----------------------------------------
Sub Set_MousePOS(x, y)

    SetCursorPos x, y

End Sub


'----------------------------------------
'Name: Get_RayPick
'Object: Get
'Event: RayPick
'----------------------------------------
Sub Get_RayPick(ByRef Vorigin As D3DVECTOR, ByRef vDir As D3DVECTOR)

  Dim vPickRayDir As D3DVECTOR
  Dim vPickRayOrig As D3DVECTOR
  Dim Mview As D3DMATRIX

    ''//Get the pick ray from the mouse position

  Dim ptCursor As POINTAPI
    Call GetCursorPos(ptCursor)
    'Call ClientToScreen(Data.Hwindow, ptCursor)

    If Data.Is_Windowed Then
  Dim RC As RECT
        Call GetWindowRect(Data.Hwindow, RC)
        ptCursor.x = ptCursor.x - RC.Left
        ptCursor.y = ptCursor.y - RC.Top
    End If

    ''//Compute the vector of the pick ray in screen space
  Dim V As D3DVECTOR

    V.x = (((2# * ptCursor.x) / Data.Buffer_Width) - 1) / Data.MatProjec.m11
    V.y = -(((2# * ptCursor.y) / Data.Buffer_Height) - 1) / Data.MatProjec.m22
    V.z = 1#

    ''//Get the inverse view matrix
  Dim M As D3DMATRIX
    D3dDevice.GetTransform D3DTS_VIEW, Mview
    Call D3DXMatrixInverse(M, 0, Mview)

    ''//Transform the screen space pick ray into 3D space
    vPickRayDir.x = V.x * M.m11 + V.y * M.m21 + V.z * M.m31
    vPickRayDir.y = V.x * M.m12 + V.y * M.m22 + V.z * M.m32
    vPickRayDir.z = V.x * M.m13 + V.y * M.m23 + V.z * M.m33
    Call D3DXVec3Normalize(vPickRayDir, vPickRayDir)
    vPickRayOrig.x = M.m41
    vPickRayOrig.y = M.m42
    vPickRayOrig.z = M.m43

    ''//calc origin as intersection with near frustum

    'vPickRayOrig+=vPickRayDir*NEAR_Z
    vDir = vPickRayDir
    D3DXVec3Scale vPickRayDir, vPickRayDir, Data.ViewFrust.Near
    D3DXVec3Add Vorigin, vPickRayOrig, vPickRayDir

End Sub


'----------------------------------------
'Name: Get_Capabilities
'Object: Get
'Event: Capabilities
'----------------------------------------
Private Sub Get_Capabilities()

  Dim d3dCaps As D3DCAPS8, i, j

    D3dDevice.GetDeviceCaps d3dCaps

    'check bump mapping
    

    Data.Capa.BumpMapping = 1
    ''//Does this device support the two bump mapping blend operations?
    If (d3dCaps.TextureOpCaps And _
       (D3DTEXOPCAPS_BUMPENVMAP Or D3DTEXOPCAPS_BUMPENVMAPLUMINANCE) = 0) Then
        Data.Capa.BumpMapping = 0
    End If

    ''//Does this device support up to three blending stages?
    If d3dCaps.MaxTextureBlendStages < 3 Then
        Data.Capa.BumpMapping = 0
    End If

    ''//Does this device support multitexturing
    If d3dCaps.MaxSimultaneousTextures > 1 Then
        Data.Capa.CanDoMultiTexture = 1
    End If

    'anisotropic filter
    If d3dCaps.RasterCaps And D3DPRASTERCAPS_ANISOTROPY Then
        Data.Capa.Filter_Anisotropic = True
    End If

End Sub


'----------------------------------------
'Name: Set_ViewFrustum
'Object: Set
'Event: ViewFrustum
'----------------------------------------
Sub Set_ViewFrustum(Near, Far, FovAngle, Optional Aspect = 1)

    With Data.ViewFrust
        .Aspect = Aspect
        .Near = Near
        .Far = Far
        .FovAngle = FovAngle

    End With 'DATA.VIEWFRUST

    D3DXMatrixPerspectiveFovLH Data.MatProjec, Data.ViewFrust.FovAngle, Data.ViewFrust.Aspect, Data.ViewFrust.Near, Data.ViewFrust.Far
    D3dDevice.SetTransform D3DTS_PROJECTION, Data.MatProjec

End Sub



'----------------------------------------
'Name: Push_Renderstate
'Object: Push
'Event: Renderstate
'----------------------------------------
Function Push_Renderstate(StateVal As NEMO_RENDERSTATE)
   
  
   Data.lpState.m_State(StateVal) = D3dDevice.GetRenderState(StateVal)

End Function


'----------------------------------------
'Name: Pop_Renderstate
'Object: Pop
'Event: Renderstate
'----------------------------------------
Function Pop_Renderstate(StateVal As NEMO_RENDERSTATE)
   D3dDevice.SetRenderState StateVal, Data.lpState.m_State(StateVal)
End Function




'=======
' Engine Renderstate Parameters
'
'===


'----------------------------------------
'Name: Set_EngineTextureFilter
'Object: Set
'Event: EngineTextureFilter
'----------------------------------------
Sub Set_EngineTextureFilter(Filter As NEMO_FILTERING_MODE)

    If Filter = NEMO_FILTER_POINT Then

        Call D3dDevice.SetTextureStageState(0, D3DTSS_MAGFILTER, D3DTEXF_POINT)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_MINFILTER, D3DTEXF_POINT)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_MIPFILTER, D3DTEXF_POINT)

        If Data.Capa.CanDoMultiTexture Then
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MAGFILTER, D3DTEXF_POINT)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MINFILTER, D3DTEXF_POINT)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MIPFILTER, D3DTEXF_POINT)
        End If

    End If

    If Filter = NEMO_FILTER_BILINEAR Then
        Call D3dDevice.SetTextureStageState(0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_MINFILTER, D3DTEXF_LINEAR)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_MIPFILTER, D3DTEXF_POINT)

        If Data.Capa.CanDoMultiTexture Then
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MINFILTER, D3DTEXF_LINEAR)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MIPFILTER, D3DTEXF_POINT)
        End If

    End If

    If Filter = NEMO_FILTER_TRILINEAR Then
        Call D3dDevice.SetTextureStageState(0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_MINFILTER, D3DTEXF_LINEAR)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_MIPFILTER, D3DTEXF_LINEAR)

        If Data.Capa.CanDoMultiTexture Then
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MINFILTER, D3DTEXF_LINEAR)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MIPFILTER, D3DTEXF_LINEAR)
        End If

    End If

    If Filter = NEMO_FILTER_ANISOTROPIC And Data.Capa.Filter_Anisotropic Then
        Call D3dDevice.SetTextureStageState(0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_MIPFILTER, D3DTEXF_ANISOTROPIC)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_MAXANISOTROPY, 4)

        If Data.Capa.CanDoMultiTexture Then
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MIPFILTER, D3DTEXF_ANISOTROPIC)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_MAXANISOTROPY, 4)

        End If

    End If

End Sub


'----------------------------------------
'Name: Set_EngineShadeMode
'Object: Set
'Event: EngineShadeMode
'----------------------------------------
Sub Set_EngineShadeMode(ShadingModel As NEMO_SHADING_MODE)

    Call D3dDevice.SetRenderState(D3DRS_SHADEMODE, ShadingModel)

End Sub


'----------------------------------------
'Name: Set_EngineFillMode
'Object: Set
'Event: EngineFillMode
'----------------------------------------
Sub Set_EngineFillMode(FillMode As NEMO_FILL_MODE)

    If FillMode = NEMO_FILL_WIREFRAME Then
        Call D3dDevice.SetRenderState(D3DRS_FILLMODE, D3DFILL_WIREFRAME)
        Call D3dDevice.SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        Call D3dDevice.SetTexture(0, Nothing)

      Else 'NOT FILLMODE...
        Call D3dDevice.SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)
        Call D3dDevice.SetRenderState(D3DRS_CULLMODE, D3DCULL_CCW)
    End If

End Sub


'----------------------------------------
'Name: Set_EngineVertexShader
'Object: Set
'Event: EngineVertexShader
'----------------------------------------
Sub Set_EngineVertexShader(FVF_format As NEMO_VERTEX_FVF)

    Call D3dDevice.SetVertexShader(FVF_format)

End Sub


'----------------------------------------
'Name: Set_LightAmbient
'Object: Set
'Event: LightAmbient
'----------------------------------------
Sub Set_LightAmbient(color As Long)

    D3dDevice.SetRenderState D3DRS_AMBIENT, color

End Sub


'----------------------------------------
'Name: Set_light
'Object: Set
'Event: light
'----------------------------------------
Sub Set_light(Enable As Boolean)

    D3dDevice.SetRenderState D3DRS_LIGHTING, Enable

End Sub


'----------------------------------------
'Name: Set_TEX_SphereMapping
'Object: Set
'Event: TEX_SphereMapping
'----------------------------------------
Sub Set_TEX_SphereMapping(Enable As Boolean, Optional Stage = 0)

    If Enable Then
        D3dDevice.SetTransform D3DTS_TEXTURE0, Data.MatSphereMapping
        D3dDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        D3dDevice.SetTextureStageState Stage, D3DTSS_TEXTURETRANSFORMFLAGS, D3DTTFF_COUNT2
        D3dDevice.SetTextureStageState Stage, D3DTSS_TEXCOORDINDEX, D3DTSS_TCI_CAMERASPACEREFLECTIONVECTOR
        D3dDevice.SetRenderState D3DRS_NORMALIZENORMALS, 1
      Else 'ENABLE = FALSE
        D3dDevice.SetTransform D3DTS_TEXTURE0, GLOB.GetIDmatrix
        D3dDevice.SetTextureStageState Stage, D3DTSS_TEXTURETRANSFORMFLAGS, D3DTTFF_DISABLE
        D3dDevice.SetTextureStageState Stage, D3DTSS_TEXCOORDINDEX, D3DTSS_TCI_PASSTHRU

    End If

End Sub


'----------------------------------------
'Name: Set_EngineRenderState
'Object: Set
'Event: EngineRenderState
'----------------------------------------
Sub Set_EngineRenderState(RenderState As CONST_D3DRENDERSTATETYPE, Val As Long)

    D3dDevice.SetRenderState RenderState, Val

End Sub


'----------------------------------------
'Name: Set_EngineBlendingMode
'Object: Set
'Event: EngineBlendingMode
'----------------------------------------
Public Sub Set_EngineBlendingMode(Destination As CONST_D3DBLEND, src As CONST_D3DBLEND)

    D3dDevice.SetRenderState D3DRS_SRCBLEND, src
    D3dDevice.SetRenderState D3DRS_DESTBLEND, Destination
    

End Sub


'----------------------------------------
'Name: Set_TEX_texture
'Object: Set
'Event: TEX_texture
'----------------------------------------
Sub Set_TEX_texture(TextureSurface As Direct3DBaseTexture8, Optional Stage As Long = 0)

    D3dDevice.SetTexture Stage, TextureSurface

End Sub


'----------------------------------------
'Name: Set_CullMode
'Object: Set
'Event: CullMode
'----------------------------------------
Sub Set_CullMode(CullMode As CONST_D3DCULL)

    D3dDevice.SetRenderState D3DRS_CULLMODE, CullMode

End Sub


'----------------------------------------
'Name: Set_Fog
'Object: Set
'Event: Fog
'----------------------------------------
Public Sub Set_Fog(Enable As Boolean, Optional density As Double = 1 / 500, Optional color As Long = &HFFFFFFFF, Optional FogMode As CONST_D3DFOGMODE = D3DFOG_EXP)

    If density > 0 Then
        density = density
        D3dDevice.SetRenderState D3DRS_FOGDENSITY, FtoDW(CSng(density))
    End If

    D3dDevice.SetRenderState D3DRS_FOGCOLOR, color
    D3dDevice.SetRenderState D3DRS_FOGVERTEXMODE, FogMode
    D3dDevice.SetRenderState D3DRS_FOGENABLE, Enable

End Sub


'----------------------------------------
'Name: Set_Fog_EX
'Object: Set
'Event: Fog_EX
'----------------------------------------
Sub Set_Fog_EX(FogStart As Single, FogEnd As Single, Optional density As Double = 1 / 1000, Optional color As Long = &HFFFFFF)

  '' Turn on fog

    With D3dDevice

        .SetRenderState D3DRS_FOGENABLE, 1 ' True
        .SetRenderState D3DRS_FOGCOLOR, color
        .SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_NONE
        .SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_LINEAR
        .SetRenderState D3DRS_RANGEFOGENABLE, 0 'False
        .SetRenderState D3DRS_FOGSTART, FtoDW(FogStart)
        .SetRenderState D3DRS_FOGEND, FtoDW(FogEnd)
    End With 'D3DDEVICE

End Sub


'----------------------------------------
'Name: Take_SnapShot
'Object: Take
'Event: SnapShot
'----------------------------------------
Public Sub Take_SnapShot(ByVal sFilename As String)

  Dim oSurface As Direct3DSurface8
  Dim SrcPalette As PALETTEENTRY
  Dim SrcRect As RECT
  Dim DispMode As D3DDISPLAYMODE

    'get display dimensions
    D3dDevice.GetDisplayMode DispMode
    GetWindowRect Data.Hwindow, SrcRect
    'create a surface to put front buffer on,
    'GetFrontBuffer always returns D3DFMT_A8R8G8B8
    Set oSurface = D3dDevice.CreateImageSurface(DispMode.width, _
        DispMode.Height, _
        D3DFMT_A8R8G8B8)

    'get data from front buffer
    D3dDevice.GetFrontBuffer oSurface

    'we are saving entire area of this surface
    With SrcRect
        .Left = 0
        .Right = DispMode.width
        .Top = 0
        .bottom = DispMode.Height
    End With 'SRCRECT

    'save this surface to a BMP file
    D3DX.SaveSurfaceToFile sFilename, _
                           D3DXIFF_BMP, _
                           oSurface, _
                           SrcPalette, _
                           SrcRect

End Sub


'----------------------------------------
'Name: Set_Multitexture
'Object: Set
'Event: Multitexture
'----------------------------------------
Sub Set_Multitexture(Optional style = 0)

  'If USE_stage0 Then call D3dDevice.SetTexture(0, TEX1)

    Select Case style
        'Dark_Mapping
      Case 0
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_SELECTARG1)

        '' Set darkmap
        'call D3dDevice.SetTexture(1, Tex1)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG2, D3DTA_CURRENT)

      Case 1 'smooth_satur
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_DIFFUSE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG2)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE)

        '' Set darkmap
        'call D3dDevice.SetTexture(1, Tex1)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLOROP, D3DTOP_ADD)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG2, D3DTA_CURRENT)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1)

      Case 2 'DetailMODULATE
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_DIFFUSE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG2)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE)

        '' Set darkmap
        'call D3dDevice.SetTexture(1, TEX1_stage1)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLOROP, D3DTOP_MODULATE2X)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG2, D3DTA_CURRENT)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1)

      Case 3 'Normal

        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_DIFFUSE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_ADDSIGNED)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_DIFFUSE)

        'Call D3dDevice.SetTexture(1, Text_LEV_1)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG2, D3DTA_DIFFUSE)

      Case 4 'Dark_Diffu_color_ligh
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_DIFFUSE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_ADDSIGNED)

        '' Set darkmap
        'call D3dDevice.SetTexture(1, Tex1)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG2, D3DTA_CURRENT)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLOROP, D3DTOP_MODULATE)

      Case 5 'Specular_mapping
        '// Program Stage 0:
        'call D3dDevice.SetTexture(0, Tex0)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_DIFFUSE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)

        '// Program Stage 1:
        'call D3dDevice.SetTexture(1, Tex1)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG2, D3DTA_CURRENT)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLOROP, D3DTOP_ADD)

      Case 6
        'Blending the texture with material color

        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_DIFFUSE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_ADD)

      Case 7
        'Darkmap blended with material color

        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_DIFFUSE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)

        '// Set darkmap
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG2, D3DTA_CURRENT)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLOROP, D3DTOP_MODULATE)

      Case 8

        Call D3dDevice.SetTextureStageState(0, D3DTSS_TEXCOORDINDEX, 0)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_SELECTARG1)

        '// Set glow map
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG2, D3DTA_CURRENT)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLOROP, D3DTOP_ADD)

      Case 9

        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_SELECTARG1)

        '// Set detail map
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG2, D3DTA_CURRENT)
        Call D3dDevice.SetTextureStageState(1, D3DTSS_COLOROP, D3DTOP_ADDSIGNED)

      Case 10

        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_DIFFUSE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)

        Call D3dDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE)
        Call D3dDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)

      Case 11

        Call D3dDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        Call D3dDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR)
        Call D3dDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR)

    End Select

End Sub


'----------------------------------------
'Name: Set_MultiTextureDisable
'Object: Set
'Event: MultiTextureDisable
'----------------------------------------
Sub Set_MultiTextureDisable()

    Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG1, D3DTOP_DISABLE)
    Call D3dDevice.SetTextureStageState(1, D3DTSS_COLOROP, D3DTOP_DISABLE)
    Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG2, D3DTOP_DISABLE)
    Call D3dDevice.SetTextureStageState(1, D3DTSS_ALPHAARG1, D3DTOP_DISABLE)
    Call D3dDevice.SetTextureStageState(1, D3DTSS_ALPHAOP, D3DTOP_DISABLE)

    Call D3dDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, False)

End Sub

'=======
' World transformation
'
'===

'
'===Rotation
'
'----------------------------------------
'Name: Set_WorldRotation
'Object: Set
'Event: WorldRotation
'----------------------------------------
Sub Set_WorldRotation(Vrot As D3DVECTOR)

  Dim MatZ As D3DMATRIX

    Call Set_WorldRotationEX(Vrot.x, Vrot.y, Vrot.z)

End Sub


'----------------------------------------
'Name: Set_WorldRotationEX
'Object: Set
'Event: WorldRotationEX
'----------------------------------------
Sub Set_WorldRotationEX(x, y, z)

  Dim matTemp As D3DMATRIX

    D3DXMatrixIdentity matTemp

    GLOB.MRotate matTemp, x, y, z
    D3dDevice.SetTransform D3DTS_WORLD, matTemp

End Sub

'
'===Translation
'
'----------------------------------------
'Name: Set_WorldTranslation
'Object: Set
'Event: WorldTranslation
'----------------------------------------
Sub Set_WorldTranslation(Vtrans As D3DVECTOR)

  Dim MatZ As D3DMATRIX

    Call Set_WorldRotationEX(Vtrans.x, Vtrans.y, Vtrans.z)

End Sub


'----------------------------------------
'Name: Set_WorldTranslationEX
'Object: Set
'Event: WorldTranslationEX
'----------------------------------------
Sub Set_WorldTranslationEX(x, y, z)

  Dim matTemp As D3DMATRIX

    D3DXMatrixIdentity matTemp
    matTemp = GLOB.MTranslate(matTemp, x, y, z)
    D3dDevice.SetTransform D3DTS_WORLD, matTemp

End Sub


'----------------------------------------
'Name: Set_WorldScale
'Object: Set
'Event: WorldScale
'----------------------------------------
Sub Set_WorldScale(VScale As D3DVECTOR)

  Dim MatZ As D3DMATRIX

    Call Set_WorldRotationEX(VScale.x, VScale.y, VScale.z)

End Sub


'----------------------------------------
'Name: Set_WorldScaleEX
'Object: Set
'Event: WorldScaleEX
'----------------------------------------
Sub Set_WorldScaleEX(x, y, z)

  Dim matTemp As D3DMATRIX

    D3DXMatrixIdentity matTemp
    matTemp = GLOB.MScale(matTemp, x, y, z)
    D3dDevice.SetTransform D3DTS_WORLD, matTemp

End Sub


'----------------------------------------
'Name: Set_WorldRotate_Scale_Translate
'Object: Set
'Event: WorldRotate_Scale_Translate
'----------------------------------------
Sub Set_WorldRotate_Scale_Translate(RotRadVec As D3DVECTOR, ScalVec As D3DVECTOR, TransVec As D3DVECTOR)

  Dim matTemp As D3DMATRIX

    D3DXMatrixIdentity matTemp
    matTemp = GLOB.Matrix_Get(ScalVec.x, ScalVec.y, ScalVec.z, RotRadVec.x, RotRadVec.y, RotRadVec.z, TransVec.x, TransVec.y, TransVec.z)
    D3dDevice.SetTransform D3DTS_WORLD, matTemp

End Sub


'----------------------------------------
'Name: Get_WorldRotate_Scale_Translate
'Object: Get
'Event: WorldRotate_Scale_Translate
'----------------------------------------
Sub Get_WorldRotate_Scale_Translate(RetMat As D3DMATRIX, RotRadVec As D3DVECTOR, ScalVec As D3DVECTOR, TransVec As D3DVECTOR)

  Dim matTemp As D3DMATRIX

    D3DXMatrixIdentity matTemp
    matTemp = GLOB.Matrix_Get(ScalVec.x, ScalVec.y, ScalVec.z, RotRadVec.x, RotRadVec.y, RotRadVec.z, TransVec.x, TransVec.y, TransVec.z)
    RetMat = matTemp

End Sub


'----------------------------------------
'Name: Set_WorldRotate_Scale
'Object: Set
'Event: WorldRotate_Scale
'----------------------------------------
Sub Set_WorldRotate_Scale(RotRadVec As D3DVECTOR, ScalVec As D3DVECTOR)

  Dim matTemp As D3DMATRIX
  Dim ScalMat As D3DMATRIX
  Dim Trans As D3DMATRIX

    D3DXMatrixIdentity matTemp
    D3DXMatrixIdentity ScalMat
    D3DXMatrixIdentity Trans

    D3DXMatrixScaling ScalMat, ScalVec.x, ScalVec.x, ScalVec.z
    MRotate Trans, RotRadVec.x, RotRadVec.y, RotRadVec.z
    D3DXMatrixMultiply matTemp, Trans, ScalMat

    D3dDevice.SetTransform D3DTS_WORLD, matTemp

End Sub


'----------------------------------------
'Name: Set_WorldRotate_Scale_TranslateEX
'Object: Set
'Event: WorldRotate_Scale_TranslateEX
'----------------------------------------
Sub Set_WorldRotate_Scale_TranslateEX(RadX, RadY, RadZ, ScalX, ScalY, ScalZ, TransX, TransY, TransZ)

    Call Set_WorldRotate_Scale_Translate(Vector(RadX, RadY, RadZ), Vector(ScalX, ScalY, ScalZ), Vector(TransX, TransY, TransZ))

End Sub


'----------------------------------------
'Name: Set_World_Scale_TranslateEX
'Object: Set
'Event: World_Scale_TranslateEX
'----------------------------------------
Sub Set_World_Scale_TranslateEX(ScalX, ScalY, ScalZ, TransX, TransY, TransZ)

  Dim matTemp As D3DMATRIX
  Dim ScalMat As D3DMATRIX
  Dim Trans As D3DMATRIX

    D3DXMatrixIdentity matTemp
    D3DXMatrixIdentity ScalMat
    D3DXMatrixIdentity Trans

    D3DXMatrixScaling ScalMat, ScalX, ScalY, ScalZ
    D3DXMatrixTranslation Trans, TransX, TransY, TransZ
    D3DXMatrixMultiply matTemp, Trans, ScalMat

    D3dDevice.SetTransform D3DTS_WORLD, matTemp

End Sub


'----------------------------------------
'Name: Init_Input
'Object: Init
'Event: Input
'----------------------------------------
Private Sub Init_Input(Hand As Long)

  ' Create Direct Input

    Set DI = DX8.DirectInputCreate()

    ' Create keyboard device
    Set DInputDevice = DI.CreateDevice("GUID_SysKeyboard")
    ' Set common data format to keyboard
    DInputDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DInputDevice.SetCooperativeLevel Hand, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DInputDevice.Acquire

    ' Create Mouse device
    Set DIMouse = DI.CreateDevice("GUID_SysMouse")
    'Set common data format to mouse
    DIMouse.SetCommonDataFormat DIFORMAT_MOUSE
    DIMouse.SetCooperativeLevel Hand, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DIMouse.Acquire

End Sub


'----------------------------------------
'Name: Get_KeyPress
'Object: Get
'Event: KeyPress
'----------------------------------------
Function Get_KeyPress(Key As NEMO_KEY_CONST) As Boolean

    DInputDevice.GetDeviceStateKeyboard DIState

    ' If key is pressed then
    If DIState.Key(Key) <> 0 Then Get_KeyPress = True

End Function


'----------------------------------------
'Name: Get_MouseBouton
'Object: Get
'Event: MouseBouton
'----------------------------------------
Function Get_MouseBouton(BUTTON As NEMO_MOUSE_BUTTON) As Boolean

  ' Lets get the mouse state

    DIMouse.GetDeviceStateMouse DIMState
    If DIMState.Buttons(BUTTON) <> 0 Then Get_MouseBouton = True

End Function


'----------------------------------------
'Name: Get_MouseXmove
'Object: Get
'Event: MouseXmove
'----------------------------------------
Function Get_MouseXmove() As Integer

  ' Lets get the mouse state

    DIMouse.GetDeviceStateMouse DIMState
    Get_MouseXmove = DIMState.lX

End Function


'----------------------------------------
'Name: Get_MouseYmove
'Object: Get
'Event: MouseYmove
'----------------------------------------
Function Get_MouseYmove() As Integer

  ' Lets get the mouse state

    DIMouse.GetDeviceStateMouse DIMState
    Get_MouseYmove = DIMState.lY

End Function


'----------------------------------------
'Name: Get_MOUSE_Cursor_API_X
'Object: Get
'Event: MOUSE_Cursor_API_X
'----------------------------------------
Function Get_MOUSE_Cursor_API_X() As Integer

  Dim pos As POINTAPI

    Call GetCursorPos(pos)
    Get_MOUSE_Cursor_API_X = pos.x

End Function


'----------------------------------------
'Name: Get_MOUSE_Cursor_API_Y
'Object: Get
'Event: MOUSE_Cursor_API_Y
'----------------------------------------
Function Get_MOUSE_Cursor_API_Y() As Integer

  Dim pos As POINTAPI

    Call GetCursorPos(pos)
    Get_MOUSE_Cursor_API_Y = pos.y

End Function


'----------------------------------------
'Name: Let BackBuffer_ClearCOLOR
'Object: Let BackBuffer
'Event: ClearCOLOR
'----------------------------------------
Property Let BackBuffer_ClearCOLOR(NewColor As Long)

    Data.BackBuff_ClearColor = NewColor

End Property


'----------------------------------------
'Name: Get BackBuffer_ClearCOLOR
'Object: Get BackBuffer
'Event: ClearCOLOR
'----------------------------------------
Property Get BackBuffer_ClearCOLOR() As Long

    BackBuffer_ClearCOLOR = Data.BackBuff_ClearColor

End Property


'----------------------------------------
'Name: Get Engine_Active
'Object: Get Engine
'Event: Active
'----------------------------------------
Property Get Engine_Active() As Boolean

    Engine_Active = Data.Is_engineActive

End Property


'----------------------------------------
'Name: Free
'----------------------------------------
Sub Free()

    Data.Is_engineActive = False

    Set D3dDevice = Nothing
    Set DX8 = Nothing
    Set D3D = Nothing
    Set D3DX = Nothing

    'free input device
    Set DI = Nothing

    DInputDevice.Unacquire
    DIMouse.Unacquire

    Set DInputDevice = Nothing
    Set DIMouse = Nothing

End Sub


'----------------------------------------
'Name: Clear3D
'----------------------------------------
Sub Clear3D()

    If Data.Is_engineActive Then

        On Error Resume Next
        D3dDevice.Clear ByVal 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, Data.BackBuff_ClearColor, 1, 0

    End If

End Sub


'----------------------------------------
'Name: Get ElapseFrame
'----------------------------------------
Public Property Get ElapseFrame() As Single

    ElapseFrame = NBFrame

End Property

'Public Property Get FramesPerSec() As Single
'  FramesPerSec = FPS
'End Property
'
'Public Property Get ImparSeconde() As Single
'  ImparSeconde = NemoFPS_2
'End Property



'----------------------------------------
'Name: QPTimer
'----------------------------------------
Private Function QPTimer() As Currency
  If Takt = 0 Then
    QueryPerformanceFrequency Takt
  End If
  
  QueryPerformanceCounter Dauer
  
  QPTimer = Dauer / Takt
End Function


'----------------------------------------
'Name: Get Framesperseconde
'----------------------------------------
Public Property Get Framesperseconde() As Single

    Framesperseconde = NemoFPS_Current

End Property


'----------------------------------------
'Name: GetFPS
'----------------------------------------
Function GetFPS()
  GetFPS = FPS_Current
End Function



'----------------------------------------
'Name: Begin3D
'----------------------------------------
Sub Begin3D()

    Clear3D
    D3dDevice.BeginScene

End Sub


'----------------------------------------
'Name: End3D
'----------------------------------------
Sub End3D()

    D3dDevice.EndScene
    
    Flip

End Sub


'----------------------------------------
'Name: Engine_GetDisplaymode
'Object: Engine
'Event: GetDisplaymode
'----------------------------------------
Sub Engine_GetDisplaymode(lpDisp As D3DDISPLAYMODE)

    lpDisp = Nemo_DispModeParam

End Sub


'----------------------------------------
'Name: Flip
'----------------------------------------
Sub Flip()
Dim tt
Dim RectClient As RECT
Dim calc


    If Data.Is_engineActive Then

        '*****************************
        '*  for Frames per seconde
        '***************************

       FPS_timefactor = QPTimer - FPS_tLastFrame
        FPS_tLastFrame = QPTimer
        
        FPS_Counter = FPS_Counter + 1
        If QPTimer - FPS_tLastFPSdisplay >= 1 Then
            FPS_Current = FPS_Counter / (QPTimer - FPS_tLastFPSdisplay)
            FPS_Counter = 0
            FPS_tLastFPSdisplay = QPTimer
        End If
       
       
       

        '//Frame rate calculations
        If GetTickCount() - NemoFPS_LastCheck >= 1000 Then
            NemoFPS_Current = NemoFPS_Count
            NemoFPS_Count = 0 'reset the counter
            NemoFPS_LastCheck = GetTickCount()
        End If
        NemoFPS_Count = NemoFPS_Count + 1

        NBFrame = GetTickCount - OldTime + 1
        OldTime = GetTickCount

        GLOB.NEMO_NbElapseFrame = NBFrame

        'ComputeClipPlanes

        If Data.Is_Windowed Then
            GetWindowRect Data.Hwindow, RectClient
            Data.Buffer_Width = RectClient.Right - RectClient.Left
            Data.Buffer_Height = RectClient.bottom - RectClient.Top

        End If

        If D3dDevice Is Nothing Then Exit Sub
        On Error Resume Next
            D3dDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        End If

End Sub ':( On Error Resume still active


'----------------------------------------
'Name: CreateTexture
'----------------------------------------
Function CreateTexture(TextureFile As String) As Direct3DTexture8

  Dim Tex As Direct3DTexture8

    Set Tex = D3DX.CreateTextureFromFile(D3dDevice, TextureFile)
    Set CreateTexture = Tex

End Function


'----------------------------------------
'Name: CLearMATRIX
'----------------------------------------
Sub CLearMATRIX()

  Dim MatZ As D3DMATRIX

    Call GLOB.IdentityMatrix(MatZ)
    D3dDevice.SetTransform D3DTS_WORLD, MatZ

End Sub

'==========================================
' Camera Stuff
'
'
'===========================================
'----------------------------------------
'Name: Camera_Turn_Left
'Object: Camera
'Event: Turn_Left
'----------------------------------------
Public Sub Camera_Turn_Left(StepAngleDEG)

    SaveCAM
    Data.EYES.ANGy = Data.EYES.ANGy + DegToRad(StepAngleDEG) * NBFrame
    'If DATA.EYES.ANGy < 0 Then DATA.EYES.ANGy = PI_360 - (-DATA.EYES.ANGy)

    Camera_update

End Sub


'----------------------------------------
'Name: Camera_Turn_Right
'Object: Camera
'Event: Turn_Right
'----------------------------------------
Public Sub Camera_Turn_Right(StepAngleDEG)

    SaveCAM
    Data.EYES.ANGy = Data.EYES.ANGy - DegToRad(StepAngleDEG) * NBFrame
    'If DATA.EYES.ANGy < 0 Then DATA.EYES.ANGy = PI_360 + (-DATA.EYES.ANGy)

    Camera_update

End Sub


'----------------------------------------
'Name: Camera_Move_Foward
'Object: Camera
'Event: Move_Foward
'----------------------------------------
Public Sub Camera_Move_Foward(StepZ)

    SaveCAM
  Dim Ang
    Ang = PI_360 - Data.EYES.ANGy
    Data.EYES.EYE.x = Data.EYES.EYE.x + Sin(Ang) * StepZ
    Data.EYES.EYE.z = Data.EYES.EYE.z + Cos(Ang) * StepZ

    Camera_update

End Sub


'----------------------------------------
'Name: Camera_Move_Backward
'Object: Camera
'Event: Move_Backward
'----------------------------------------
Public Sub Camera_Move_Backward(StepZ)

    Camera_Move_Foward -StepZ

End Sub


'----------------------------------------
'Name: Camera_Turn_UP
'Object: Camera
'Event: Turn_UP
'----------------------------------------
Public Sub Camera_Turn_UP(StepAngleDEG)

    SaveCAM
    Data.EYES.ANGx = Data.EYES.ANGx + DegToRad(StepAngleDEG) * NBFrame
    Camera_update

End Sub


'----------------------------------------
'Name: Camera_Turn_DOWN
'Object: Camera
'Event: Turn_DOWN
'----------------------------------------
Public Sub Camera_Turn_DOWN(StepAngleDEG)

    Camera_Turn_UP -StepAngleDEG

End Sub


'----------------------------------------
'Name: Camera_Strafe_UP
'Object: Camera
'Event: Strafe_UP
'----------------------------------------
Public Sub Camera_Strafe_UP(StepY)

    SaveCAM
    Data.EYES.EYE.y = Data.EYES.EYE.y + StepY

    Camera_update

End Sub


'----------------------------------------
'Name: Camera_Strafe_DOWN
'Object: Camera
'Event: Strafe_DOWN
'----------------------------------------
Public Sub Camera_Strafe_DOWN(StepY)

    Me.Camera_Strafe_UP -StepY

End Sub


'----------------------------------------
'Name: Camera_Strafe_Left
'Object: Camera
'Event: Strafe_Left
'----------------------------------------
Public Sub Camera_Strafe_Left(Step)

    SaveCAM
  Dim Ang
    Step = Step

    Ang = PI_360 - Data.EYES.ANGy

    Ang = Ang - PI_90
    Data.EYES.EYE.x = Data.EYES.EYE.x + Sin(Ang) * Step
    Data.EYES.EYE.z = Data.EYES.EYE.z + Cos(Ang) * Step

    Camera_update

End Sub


'----------------------------------------
'Name: Camera_Strafe_Right
'Object: Camera
'Event: Strafe_Right
'----------------------------------------
Public Sub Camera_Strafe_Right(Step)

    Me.Camera_Strafe_Left -Step

End Sub


'----------------------------------------
'Name: Camera_Roll_Left
'Object: Camera
'Event: Roll_Left
'----------------------------------------
Public Sub Camera_Roll_Left(SetpAngleDEG)

    SaveCAM
    Data.EYES.ANGz = Data.EYES.ANGz - SetpAngleDEG
    Camera_update

End Sub


'----------------------------------------
'Name: Camera_Roll_Right
'Object: Camera
'Event: Roll_Right
'----------------------------------------
Public Sub Camera_Roll_Right(SetpAngle)

    Camera_Roll_Left -SetpAngle

End Sub


'----------------------------------------
'Name: Get CAM_angle
'Object: Get CAM
'Event: angle
'----------------------------------------
Property Get CAM_angle()

    CAM_angle = Data.EYES.ANGy

End Property


'----------------------------------------
'Name: Set CAM_angle
'Object: Set CAM
'Event: angle
'----------------------------------------
Property Set CAM_angle(Ang)

    Data.EYES.DEG = Ang

End Property


'----------------------------------------
'Name: Get CAM_angleDEGRE
'Object: Get CAM
'Event: angleDEGRE
'----------------------------------------
Property Get CAM_angleDEGRE()

    CAM_angleDEGRE = DegToRad(Data.EYES.ANGy)

End Property


'----------------------------------------
'Name: Get CAM_angleBilBOARD
'Object: Get CAM
'Event: angleBilBOARD
'----------------------------------------
Property Get CAM_angleBilBOARD()

    CAM_angleBilBOARD = (Data.EYES.ANGy) - PI_90

End Property


'----------------------------------------
'Name: Set CAM_angleDEGRE
'Object: Set CAM
'Event: angleDEGRE
'----------------------------------------
Property Set CAM_angleDEGRE(ANG_degree)

    Data.EYES.DEG = ANG_degree

End Property


'----------------------------------------
'Name: Camera_set_EYE
'Object: Camera
'Event: set_EYE
'----------------------------------------
Sub Camera_set_EYE(veye As D3DVECTOR)

    Data.EYES.EYE = veye
    Me.Camera_update

End Sub


'----------------------------------------
'Name: Camera_GetLookAt
'Object: Camera
'Event: GetLookAt
'----------------------------------------
Function Camera_GetLookAt() As D3DVECTOR

    Camera_GetLookAt = Data.EYES.Dest_at

End Function


'----------------------------------------
'Name: Camera_GetPosition
'Object: Camera
'Event: GetPosition
'----------------------------------------
Function Camera_GetPosition() As D3DVECTOR

    Camera_GetPosition = Data.EYES.EYE

End Function


'----------------------------------------
'Name: Camera_GetLastPosition
'Object: Camera
'Event: GetLastPosition
'----------------------------------------
Function Camera_GetLastPosition() As D3DVECTOR

    Camera_GetLastPosition = Data.SaveCAM.EYE

End Function

'Data.SaveCAM
'----------------------------------------
'Name: Camera_GetVelocity
'Object: Camera
'Event: GetVelocity
'----------------------------------------
Function Camera_GetVelocity() As D3DVECTOR

    D3DXVec3Subtract Camera_GetVelocity, Data.EYES.EYE, Data.SaveCAM.EYE

End Function


'----------------------------------------
'Name: Camera_GetDirection
'Object: Camera
'Event: GetDirection
'----------------------------------------
Function Camera_GetDirection() As D3DVECTOR

  Dim VC As D3DVECTOR

    D3DXVec3Subtract VC, Data.EYES.EYE, Data.SaveCAM.EYE
    D3DXVec3Normalize Camera_GetDirection, VC

End Function


'----------------------------------------
'Name: Camera_SetPositionEX
'Object: Camera
'Event: SetPositionEX
'----------------------------------------
Function Camera_SetPositionEX(vPos As D3DVECTOR) As D3DVECTOR

    Data.EYES.EYE = vPos
    Me.Camera_update

End Function


'----------------------------------------
'Name: Camera_SetXRotation
'Object: Camera
'Event: SetXRotation
'----------------------------------------
Sub Camera_SetXRotation(Xdeg)

    Data.EYES.ANGx = Xdeg

    Me.Camera_update

End Sub


'----------------------------------------
'Name: Camera_SetYRotation
'Object: Camera
'Event: SetYRotation
'----------------------------------------
Sub Camera_SetYRotation(Ydeg)

    Data.EYES.ANGy = DegToRad(Ydeg)

    Me.Camera_update

End Sub


'----------------------------------------
'Name: Camera_SetZRotation
'Object: Camera
'Event: SetZRotation
'----------------------------------------
Sub Camera_SetZRotation(Zdeg)

    Data.EYES.ANGz = DegToRad(Zdeg)

    Me.Camera_update

End Sub


'----------------------------------------
'Name: Camera_SetRotation
'Object: Camera
'Event: SetRotation
'----------------------------------------
Sub Camera_SetRotation(Xrad, Yrad, Zrad)

    Data.EYES.ANGy = Yrad
    Data.EYES.ANGx = Xrad
    Data.EYES.ANGz = Zrad
    Me.Camera_update

End Sub


'----------------------------------------
'Name: Camera_SetRotationEX
'Object: Camera
'Event: SetRotationEX
'----------------------------------------
Sub Camera_SetRotationEX(Xdeg, Ydeg, Zdeg)

    Data.EYES.ANGy = DegToRad(Ydeg)
    Data.EYES.ANGx = DegToRad(Xdeg)
    Data.EYES.ANGz = DegToRad(Zdeg)
    Me.Camera_update

End Sub


'----------------------------------------
'Name: Camera_SetPosition
'Object: Camera
'Event: SetPosition
'----------------------------------------
Sub Camera_SetPosition(Campos As D3DVECTOR, LookAtVec As D3DVECTOR)

    Data.EYES.EYE = Campos
    Data.EYES.Dest_at = LookAtVec
    D3DXMatrixLookAtLH Data.matView, Campos, LookAtVec, Vector(0, 1, 0)
    D3dDevice.SetTransform D3DTS_VIEW, Data.matView

End Sub


'----------------------------------------
'Name: SaveCAM
'----------------------------------------
Sub SaveCAM()

    Data.SaveCAM = Data.EYES

End Sub


'----------------------------------------
'Name: Camera_Recall
'Object: Camera
'Event: Recall
'----------------------------------------
Sub Camera_Recall()

    Data.EYES = Data.SaveCAM

End Sub


'----------------------------------------
'Name: Camera_update
'Object: Camera
'Event: update
'----------------------------------------
Sub Camera_update()

  Dim matView As D3DMATRIX
  Dim matRotation As D3DMATRIX
  Dim matPitch As D3DMATRIX
  Dim MatRoll As D3DMATRIX
  Dim matLook As D3DMATRIX
  Dim matPos As D3DMATRIX
  Dim matWorld As D3DMATRIX
  Dim matProj As D3DMATRIX

    'setup world matrix
    '  D3DXMatrixIdentity matWorld
    '  D3dDevice.SetTransform D3DTS_WORLD, matWorld

    'make them identity matrices
    D3DXMatrixIdentity matView
    D3DXMatrixIdentity matPos
    D3DXMatrixIdentity matRotation
    D3DXMatrixIdentity matLook
    D3DXMatrixIdentity MatRoll
    'rotate around x and y, for angle and pitch

    D3DXMatrixRotationY matRotation, Data.EYES.ANGy
    D3DXMatrixRotationX matPitch, Data.EYES.ANGx
    D3DXMatrixRotationZ MatRoll, Data.EYES.ANGz

    'multiply angle and pitch matrices together to create one 'look' matrix
    D3DXMatrixMultiply matLook, matRotation, matPitch
    D3DXMatrixMultiply matLook, matLook, MatRoll

    'put the position of the camera into the translation matrix, matPos
    D3DXMatrixTranslation matPos, -Data.EYES.EYE.x, -Data.EYES.EYE.y, -Data.EYES.EYE.z
    'multiply that with the look matrix to make the complete view matrix
    D3DXMatrixMultiply matView, matPos, matLook

    'which we can then set as the view matrix:
    D3dDevice.SetTransform D3DTS_VIEW, matView

    'compute viewat
    Data.EYES.Dest_at.x = Data.EYES.EYE.x + Sin(Data.EYES.ANGy)
    Data.EYES.Dest_at.z = Data.EYES.EYE.z + Cos(Data.EYES.ANGz)
    Data.EYES.Dest_at.y = Data.EYES.EYE.y + Sin(Data.EYES.ANGx)

    'Data.EYES.Dest_at = GLOB.Rotate(Data.EYES.EYE, Data.EYES.ANGx, Data.EYES.ANGy, Data.EYES.ANGz)

    'x '=z*sin(yaw)+x*cos(yaw)
    'y '=y
    'z '=z*cos(yaw)-x*sin(yaw)
    'x "=x'"
    'y "=y'*cos(pitch)-z'*sin(pitch)"
    'z "=y'*sin(pitch)+z'*cos(pitch) "
    '
    'x"'=y"*sin(roll)+x"*cos(roll)
    'y"'=y"*cos(roll)-x"*sin(roll)
    'z "'=z"

End Sub


'----------------------------------------
'Name: Class_Initialize
'Object: Class
'Event: Initialize
'----------------------------------------
Private Sub Class_Initialize()

  Data.EYES.ANGy = GLOB.PI_90

End Sub

'=================================================
'geometry stuff
'==============================================


'----------------------------------------
'Name: Create_GroundVertex
'Object: Create
'Event: GroundVertex
'----------------------------------------
Sub Create_GroundVertex(LpVertexBuffer As Direct3DVertexBuffer8, VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional vertextype As NEMO_VERTEX_FVF = NEMO_CUSTOM_VERTEX, Optional tu = 1, Optional tv = 1)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR, size As Long

    v1 = VBotomLeft
    v2 = VTopRight

    If vertextype = NEMO_CUSTOM_VERTEX Then
  Dim V(3) As NEMO_VERTEX
        size = Len(V(0))
    End If

    V(0) = GLOB.Make_NemoVerxtex(v1.x, v2.y, v2.z, 0, 1, 0, 0, 0)
    V(1) = GLOB.Make_NemoVerxtex(v2.x, v2.y, v2.z, 0, 1, 0, tu, 0)
    V(2) = GLOB.Make_NemoVerxtex(v1.x, v2.y, v1.z, 0, 1, 0, 0, tv)
    V(3) = GLOB.Make_NemoVerxtex(v2.x, v2.y, v1.z, 0, 1, 0, tu, tv)

    Set LpVertexBuffer = D3dDevice.CreateVertexBuffer(size * 4, _
        0, vertextype, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData LpVertexBuffer, 0, size * 4, 0, V(0)

End Sub


'----------------------------------------
'Name: Render_GroundVertex
'Object: Render
'Event: GroundVertex
'----------------------------------------
Sub Render_GroundVertex(LpvBuffer As Direct3DVertexBuffer8, Optional vertextype As NEMO_VERTEX_FVF = NEMO_CUSTOM_VERTEX)

  Dim V As NEMO_VERTEX
  Dim sizeOfVertex As Long
  Dim i As Integer

    'Set_light 1
    D3dDevice.SetVertexShader vertextype

    sizeOfVertex = Len(V)

    D3dDevice.SetStreamSource 0, LpvBuffer, sizeOfVertex
    D3dDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2

End Sub

':) Ulli's VB Code Formatter V2.13.6 (17/08/2002 18:31:52) 406 + 1886 = 2292 Lines



