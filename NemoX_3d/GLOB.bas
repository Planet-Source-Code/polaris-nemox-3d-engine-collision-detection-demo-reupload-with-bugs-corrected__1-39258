Attribute VB_Name = "GLOB"

Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Dim P_dCalculatorMatrix As D3DMATRIX
Dim PIFactor As Single

Public Const MAX_LIGHTS = 32
' The number Pi
Public Const PI As Double = 3.14159265358979
Public Const PIdiv180 = PI / 180
Public Const RAD = PI / 180
Public Const PI_90 = PI / 2
Public Const PI_45 = PI / 4
Public Const PI_180 = PI
Public Const PI_360 = PI * 2
Public Const PI_270 = (PI / 2) * 3
Public Const NEMO_EPSILON = 0.0001

'========Main 3d Object==============
Public D3dDevice As Direct3DDevice8
Public DX8 As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8


Global NumTextureInpool As Long
Global NumLighmapsInpool As Long

Public POOL_texture() As Direct3DBaseTexture8
Public POOL_Lightmaps() As Direct3DBaseTexture8


Public LpGLOBAL_NEMO As NemoX
Public Enum NEMO_VERTEX_FVF
    'Global FVF Vertex format
    NEMO_CUSTOM_VERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)
    NEMO_CUSTOM_VERTEX2 = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX2)
    NEMO_CUSTOM_LVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)
    NEMO_CUSTOM_LVERTEX2 = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2)
    NEMO_CUSTOM_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)

End Enum

Public Enum NEMO_DEVICE_TYPE
    NEMO_HEL_REF = 1
    NEMO_HAL_DEVICE = 0
    NEMO_TnT_DEVICE = 2

End Enum




Public Enum NEMO_RENDERSTATE

    D3DRS_ZWRITEENABLE = 14
    D3DRS_ALPHATESTENABLE = 15
    NEMORS_LASTPIXEL = 16                 '(&H10)
    NEMORS_SRCBLEND = 19                  '(&H13)
    NEMORS_DESTBLEND = 20                 '(&H14)

    NEMORS_ZENABLE = 7
    NEMORS_FILLMODE = 8
    NEMORS_SHADEMODE = 9
    NEMORS_ZWRITEENABLE = 14
    NEMORS_ALPHATESTENABLE = 15
    NEMORS_CULLMODE = 22                  '(&H16)
    NEMORS_ZFUNC = 23                     '(&H17)
    NEMORS_ALPHAREF = 24                  '(&H18)
    NEMORS_ALPHAFUNC = 25                 '(&H19)
    NEMORS_DITHERENABLE = 26              '(&H1A)
    NEMORS_ALPHABLENDENABLE = 27          '(&H1B)
    NEMORS_FOGENABLE = 28                 '(&H1C)
    NEMORS_SPECULARENABLE = 29            '(&H1D)
    NEMORS_ZVISIBLE = 30                  '(&H1E)
    NEMORS_EDGEANTIALIAS = 40             '(&H28)
    NEMORS_ZBIAS = 47                     '(&H2F)
    NEMORS_RANGEFOGENABLE = 48            '(&H30)
    NEMORS_STENCILENABLE = 52             '(&H34)
    NEMORS_STENCILFAIL = 53               '(&H35)
    NEMORS_STENCILZFAIL = 54              '(&H36)
    NEMORS_STENCILPASS = 55               '(&H37)
    NEMORS_STENCILFUNC = 56               '(&H38)
    NEMORS_STENCILREF = 57                '(&H39)
    NEMORS_STENCILMASK = 58               '(&H3A)
    NEMORS_STENCILWRITEMASK = 59          '(&H3B)
    NEMORS_TEXTUREFACTOR = 60             '(&H3C)
    NEMORS_WRAP0 = 128                    '(&H80)
    NEMORS_WRAP1 = 129                    '(&H81)
    NEMORS_WRAP2 = 130                    '(&H82)
    NEMORS_WRAP3 = 131                    '(&H83)
    NEMORS_WRAP4 = 132                    '(&H84)
    NEMORS_WRAP5 = 133                    '(&H85)
    NEMORS_WRAP6 = 134                    '(&H86)
    NEMORS_WRAP7 = 135                    '(&H87)
    NEMORS_CLIPPING = 136                 '(&H88)
    NEMORS_LIGHTING = 137                 '(&H89)
    NEMORS_AMBIENT = 139                  '(&H8B)
    NEMORS_FOGVERTEXMODE = 140            '(&H8C)
    NEMORS_COLORVERTEX = 141              '(&H8D)
    NEMORS_LOCALVIEWER = 142              '(&H8E)
    NEMORS_NORMALIZENORMALS = 143         '(&H8F)
    NEMORS_DIFFUSEMATERIALSOURCE = 145    '(&H91)
    NEMORS_SPECULARMATERIALSOURCE = 146   '(&H92)
    NEMORS_AMBIENTMATERIALSOURCE = 147    '(&H93)
    NEMORS_EMISSIVEMATERIALSOURCE = 148   '(&H94)
    NEMORS_VERTEXBLEND = 151              '(&H97)
    NEMORS_CLIPPLANEENABLE = 152          '(&H98)
    NEMORS_SOFTWAREVERTEXPROCESSING = 153 '(&H99)
    NEMORS_POINTSIZE = 154                '(&H9A)
    NEMORS_POINTSIZE_MIN = 155            '(&H9B)
    NEMORS_POINTSPRITEENABLE = 156        '(&H9C)
    NEMORS_POINTSCALEENABLE = 157         '(&H9D)
    NEMORS_POINTSCALE_A = 158             '(&H9E)
    NEMORS_POINTSCALE_B = 159             '(&H9F)
    NEMORS_POINTSCALE_C = 160             '(&HA0)
    NEMORS_MULTISAMPLEANTIALIAS = 161     '(&HA1)
    NEMORS_MULTISAMPLEMASK = 162          '(&HA2)
    NEMORS_PATCHEDGESTYLE = 163           '(&HA3)
    NEMORS_PATCHSEGMENTS = 164            '(&HA4)
    NEMORS_DEBUGMONITORTOKEN = 165        '(&HA5)
    NEMORS_POINTSIZE_MAX = 166            '(&HA6)
    NEMORS_INDEXVERTEXBLENDENABLE = 167   '(&HA7)
    NEMORS_COLORWRITEENABLE = 168         '(&HA8)
    NEMORS_TWEENFACTOR = 170              '(&HAA)
    NEMORS_BLENDOP = 171                  '(&HAB)

End Enum

Type Nemo_SaveState
    m_State(7 To 171) As Long
    MATERIAL_state As D3DMATERIAL8
    VIEWMAT_state As D3DMATRIX
    PROJMAT_state As D3DMATRIX
    WORLDMAT_state As D3DMATRIX

End Type

Public Type NEMO_CFG_INI

    width As Integer
    Height As Integer
    format As Long
    USE_from_Dialog As Boolean
    MaxFramePerSec As Long
    USE_TnL As Boolean
    DeviceTyp As NEMO_DEVICE_TYPE
    ForceVerSINC As Boolean
    appHandle As Long
    IS_FullScreen As Boolean
    GamaLevel As Single
    BPP As Integer
    BufferCount As Integer
End Type

Public NEMO_NbElapseFrame As Single

Global lpFRUST As NEMO_FRUSTUM

'Global lpBSP As cNemo_Q3BSP

'DOC: see D3DUtil_Timer
Public Enum TIMER_COMMAND
    TIMER_RESET = 1         '- to reset the timer
    TIMER_start = 2         '- to start the timer
    TIMER_STOP = 3          '- to stop (or pause) the timer
    TIMER_ADVANCE = 4       '- to advance the timer by 0.1 seconds
    TIMER_GETABSOLUTETIME = 5 '- to get the absolute system time
    TIMER_GETAPPTIME = 6      '- to get the current time
    TIMER_GETELLAPSEDTIME = 7 '- to get the ellapsed time
End Enum

Private Type Q2_miptex_t

    Name As String * 32
    width As Long
    Height  As Long
    offsets(4 - 1) As Long ''''  four mip maps stored
    animname As String * 32   ''''  next frame in animation chain
    Flags As Long
    Contents As Long
    value As Long
End Type

Public Type D3DRMBOX
    min As D3DVECTOR
    max As D3DVECTOR
End Type

Public Type NEMO_BBOX
    min As D3DVECTOR
    max As D3DVECTOR
End Type

'=====================
'
'  some APIZ
'================
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Const INVALID_HANDLE_VALUE = -1 ':( As Integer ?

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

Private Frustum(5) As D3DPLANE
Private Enum FrustumSide
    FS_RIGHT = 0          ' The RIGHT side of the frustum
    FS_LEFT = 1           ' The LEFT  side of the frustum
    FS_BOTTOM = 2         ' The BOTTOM side of the frustum
    FS_TOP = 3            ' The TOP side of the frustum
    FS_BACK = 4           ' The BACK side of the frustum
    FS_FRONT = 5          ' The FRONT side of the frustum
End Enum

Private Type BITMAPHEADER
    'intMagic As Integer
    lngSize As Long
    intReserved1 As Integer
    intReserved2 As Integer
    lngOffset As Long

    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
'Die beiden Header
Private BMH As BITMAPHEADER
Private Const BSPLMAP_SIZE = 49152     'Die byte-daten sind 49125 groß

Dim matClip As D3DMATRIX





'----------------------------------------
'Name: Add_TextureToPool
'Object: Add
'Event: TextureToPool
'----------------------------------------
Function Add_TextureToPool(TexName As String)
 
 If FileiS_valid(TexName) Then
  NumTextureInpool = NumTextureInpool + 1
  ReDim Preserve POOL_texture(NumTextureInpool - 1)
  Set POOL_texture(NumTextureInpool - 1) = D3DX.CreateTextureFromFile(D3dDevice, TexName)
  Add_TextureToPool = NumTextureInpool - 1
 End If
End Function



'----------------------------------------
'Name: Add_LightmapToPool
'Object: Add
'Event: LightmapToPool
'----------------------------------------
Function Add_LightmapToPool(TexName As String)
 
 If FileiS_valid(TexName) Then
  NumLighmapsInpool = NumLighmapsInpool + 1
  ReDim Preserve POOL_Lightmaps(NumLighmapsInpool - 1)
  Set POOL_Lightmaps(NumLighmapsInpool - 1) = D3DX.CreateTextureFromFile(D3dDevice, TexName)
   Add_LightmapToPool = NumLighmapsInpool - 1
 End If
End Function



'----------------------------------------
'Name: Add_LightmapToPoolEX
'Object: Add
'Event: LightmapToPoolEX
'----------------------------------------
Function Add_LightmapToPoolEX(Tex As Direct3DBaseTexture8)
 
 
  NumLighmapsInpool = NumLighmapsInpool + 1
  ReDim Preserve POOL_Lightmaps(NumLighmapsInpool - 1)
  Set POOL_Lightmaps(NumLighmapsInpool - 1) = Tex
 Add_LightmapToPoolEX = NumLighmapsInpool - 1
End Function




'----------------------------------------
'Name: CheckPointVisible
'----------------------------------------
Public Function CheckPointVisible(Vect As D3DVECTOR) As Boolean

    CheckPointVisible = FRUST_PointInFrustum(lpFRUST, Vect.x, Vect.y, Vect.z)

End Function


'----------------------------------------
'Name: CheckBoxVisible
'----------------------------------------
Function CheckBoxVisible(BoxMin As D3DVECTOR, BoxMax As D3DVECTOR) As Boolean

    CheckBoxVisible = FRUST_BoxInFrustum(lpFRUST, BoxMin.x, BoxMin.y, BoxMin.z, BoxMax.x, BoxMax.y, BoxMax.z)

End Function


'----------------------------------------
'Name: CheckBoxVisibleNear
'----------------------------------------
Function CheckBoxVisibleNear(BoxMin As D3DVECTOR, BoxMax As D3DVECTOR) As Boolean

  Dim P As D3DPLANE

    CopyMemory P, lpFRUST.PLANE(0), Len(P)
    CheckBoxVisibleNear = True

    If (P.A * BoxMin.x + P.B * BoxMin.y + P.c * BoxMin.z + P.d > 0) Then CheckBoxVisibleNear = False ':( Expand Structure
    If (P.A * BoxMax.x + P.B * BoxMin.y + P.c * BoxMin.z + P.d > 0) Then CheckBoxVisibleNear = False ':( Expand Structure
    If (P.A * BoxMin.x + P.B * BoxMax.y + P.c * BoxMin.z + P.d > 0) Then CheckBoxVisibleNear = False ':( Expand Structure
    If (P.A * BoxMax.x + P.B * BoxMax.y + P.c * BoxMin.z + P.d > 0) Then CheckBoxVisibleNear = False ':( Expand Structure
    If (P.A * BoxMin.x + P.B * BoxMin.y + P.c * BoxMax.z + P.d > 0) Then CheckBoxVisibleNear = False ':( Expand Structure
    If (P.A * BoxMax.x + P.B * BoxMin.y + P.c * BoxMax.z + P.d > 0) Then CheckBoxVisibleNear = False ':( Expand Structure
    If (P.A * BoxMin.x + P.B * BoxMax.y + P.c * BoxMax.z + P.d > 0) Then CheckBoxVisibleNear = False ':( Expand Structure
    If (P.A * BoxMax.x + P.B * BoxMax.y + P.c * BoxMax.z + P.d > 0) Then CheckBoxVisibleNear = False ':( Expand Structure

End Function


'----------------------------------------
'Name: CheckBoxVisibleNearEX
'----------------------------------------
Function CheckBoxVisibleNearEX(x1, y1, z1, x2, Y2, z2) As Boolean

  Dim P As D3DPLANE

  Dim BoxMin As D3DVECTOR
  Dim BoxMax As D3DVECTOR

    BoxMin = Vector(x1, y1, z1)
    BoxMax = Vector(x2, Y2, z2)

    CopyMemory P, lpFRUST.PLANE(0), Len(P)
    CheckBoxVisibleNearEX = True

    If (P.A * BoxMin.x + P.B * BoxMin.y + P.c * BoxMin.z + P.d > 0) Then CheckBoxVisibleNearEX = False ':( Expand Structure
    If (P.A * BoxMax.x + P.B * BoxMin.y + P.c * BoxMin.z + P.d > 0) Then CheckBoxVisibleNearEX = False ':( Expand Structure
    If (P.A * BoxMin.x + P.B * BoxMax.y + P.c * BoxMin.z + P.d > 0) Then CheckBoxVisibleNearEX = False ':( Expand Structure
    If (P.A * BoxMax.x + P.B * BoxMax.y + P.c * BoxMin.z + P.d > 0) Then CheckBoxVisibleNearEX = False ':( Expand Structure
    If (P.A * BoxMin.x + P.B * BoxMin.y + P.c * BoxMax.z + P.d > 0) Then CheckBoxVisibleNearEX = False ':( Expand Structure
    If (P.A * BoxMax.x + P.B * BoxMin.y + P.c * BoxMax.z + P.d > 0) Then CheckBoxVisibleNearEX = False ':( Expand Structure
    If (P.A * BoxMin.x + P.B * BoxMax.y + P.c * BoxMax.z + P.d > 0) Then CheckBoxVisibleNearEX = False ':( Expand Structure
    If (P.A * BoxMax.x + P.B * BoxMax.y + P.c * BoxMax.z + P.d > 0) Then CheckBoxVisibleNearEX = False ':( Expand Structure

End Function


'----------------------------------------
'Name: CheckBoxVisibleEX
'----------------------------------------
Function CheckBoxVisibleEX(BoxMin As D3DVECTOR, BoxMax As D3DVECTOR) As Boolean

End Function


'----------------------------------------
'Name: CheckSphereVisible
'----------------------------------------
Function CheckSphereVisible(Center As D3DVECTOR, radius As Single) As Integer

End Function


'----------------------------------------
'Name: SetUpFrustumEX
'----------------------------------------
Sub SetUpFrustumEX()

  Dim M As D3DMATRIX

    Call D3dDevice.GetTransform(D3DTS_VIEW, M) '

    '// Near plane
    lpFRUST.PLANE(0).A = M.m14 + M.m13 '
    lpFRUST.PLANE(0).B = M.m24 + M.m23 '
    lpFRUST.PLANE(0).c = M.m34 + M.m33 '
    lpFRUST.PLANE(0).d = M.m44 + M.m43 '

    '// Left plane
    lpFRUST.PLANE(1).A = M.m14 + M.m11 '
    lpFRUST.PLANE(1).B = M.m24 + M.m21 '
    lpFRUST.PLANE(1).c = M.m34 + M.m31 '
    lpFRUST.PLANE(1).d = M.m44 + M.m41 '

    '// Right plane
    lpFRUST.PLANE(2).A = M.m14 - M.m11 '
    lpFRUST.PLANE(2).B = M.m24 - M.m21 '
    lpFRUST.PLANE(2).c = M.m34 - M.m31 '
    lpFRUST.PLANE(2).d = M.m44 - M.m41 '

    '// Bottom plane
    lpFRUST.PLANE(3).A = M.m14 + M.m12 '
    lpFRUST.PLANE(3).B = M.m24 + M.m22 '
    lpFRUST.PLANE(3).c = M.m34 + M.m32 '
    lpFRUST.PLANE(3).d = M.m44 + M.m42 '

    '// Top plane
    lpFRUST.PLANE(4).A = M.m14 - M.m12 '
    lpFRUST.PLANE(4).B = M.m24 - M.m22 '
    lpFRUST.PLANE(4).c = M.m34 - M.m32 '
    lpFRUST.PLANE(4).d = M.m44 - M.m42 '

    '// front
    lpFRUST.PLANE(5).A = M.m14 + M.m13 '
    lpFRUST.PLANE(5).B = M.m24 + M.m23 '
    lpFRUST.PLANE(5).c = M.m34 + M.m33 '
    lpFRUST.PLANE(5).d = M.m44 + M.m43 '

End Sub


'----------------------------------------
'Name: SetUpFrustumNear
'----------------------------------------
Sub SetUpFrustumNear()

  Dim M As D3DMATRIX

    Call D3dDevice.GetTransform(D3DTS_VIEW, M) '

    '// Near plane
    lpFRUST.PLANE(0).A = M.m14 + M.m13 '
    lpFRUST.PLANE(0).B = M.m24 + M.m23 '
    lpFRUST.PLANE(0).c = M.m34 + M.m33 '
    lpFRUST.PLANE(0).d = M.m44 + M.m43 '

End Sub


'----------------------------------------
'Name: SetUpFrustum
'----------------------------------------
Public Sub SetUpFrustum()
  Dim clip As D3DMATRIX
  Dim matView As D3DMATRIX
  Dim matProj As D3DMATRIX, j

    D3dDevice.GetTransform D3DTS_VIEW, matView
    D3dDevice.GetTransform D3DTS_PROJECTION, matProj

   ' FRUST_CalculateFrustum lpFRUST, MatToNemoMATRIX(matView), MatToNemoMATRIX(matProj)
   'J = lpFRUST.Plane(4).b
'Exit Sub

    D3DXMatrixMultiply matClip, matView, matProj

    clip.m11 = matView.m11 * matProj.m11 + matView.m12 * matProj.m21 + matView.m13 * matProj.m31 + matView.m14 * matProj.m41
    clip.m12 = matView.m11 * matProj.m12 + matView.m12 * matProj.m22 + matView.m13 * matProj.m32 + matView.m14 * matProj.m42
    clip.m13 = matView.m11 * matProj.m13 + matView.m12 * matProj.m23 + matView.m13 * matProj.m33 + matView.m14 * matProj.m43
    clip.m14 = matView.m11 * matProj.m14 + matView.m12 * matProj.m24 + matView.m13 * matProj.m34 + matView.m14 * matProj.m44

    clip.m21 = matView.m21 * matProj.m11 + matView.m22 * matProj.m21 + matView.m23 * matProj.m31 + matView.m24 * matProj.m41
    clip.m22 = matView.m21 * matProj.m12 + matView.m22 * matProj.m22 + matView.m23 * matProj.m32 + matView.m24 * matProj.m42
    clip.m23 = matView.m21 * matProj.m13 + matView.m22 * matProj.m23 + matView.m23 * matProj.m33 + matView.m24 * matProj.m43
    clip.m24 = matView.m21 * matProj.m14 + matView.m22 * matProj.m24 + matView.m23 * matProj.m34 + matView.m24 * matProj.m44

    clip.m31 = matView.m31 * matProj.m11 + matView.m32 * matProj.m21 + matView.m33 * matProj.m31 + matView.m34 * matProj.m41
    clip.m32 = matView.m31 * matProj.m12 + matView.m32 * matProj.m22 + matView.m33 * matProj.m32 + matView.m34 * matProj.m42
    clip.m33 = matView.m31 * matProj.m13 + matView.m32 * matProj.m23 + matView.m33 * matProj.m33 + matView.m34 * matProj.m43
    clip.m34 = matView.m31 * matProj.m14 + matView.m32 * matProj.m24 + matView.m33 * matProj.m34 + matView.m34 * matProj.m44

    clip.m41 = matView.m41 * matProj.m11 + matView.m42 * matProj.m21 + matView.m43 * matProj.m31 + matView.m44 * matProj.m41
    clip.m42 = matView.m41 * matProj.m12 + matView.m42 * matProj.m22 + matView.m43 * matProj.m32 + matView.m44 * matProj.m42
    clip.m43 = matView.m41 * matProj.m13 + matView.m42 * matProj.m23 + matView.m43 * matProj.m33 + matView.m44 * matProj.m43
    clip.m44 = matView.m41 * matProj.m14 + matView.m42 * matProj.m24 + matView.m43 * matProj.m34 + matView.m44 * matProj.m44

    'Right
    Frustum(FS_RIGHT).A = clip.m14 - clip.m11
    Frustum(FS_RIGHT).B = clip.m24 - clip.m21
    Frustum(FS_RIGHT).c = clip.m34 - clip.m31
    Frustum(FS_RIGHT).d = clip.m44 - clip.m41
    NormalizePlane Frustum(), FS_RIGHT
    'Left
    Frustum(FS_LEFT).A = clip.m14 + clip.m11
    Frustum(FS_LEFT).B = clip.m24 + clip.m21
    Frustum(FS_LEFT).c = clip.m34 + clip.m31
    Frustum(FS_LEFT).d = clip.m44 + clip.m41
    NormalizePlane Frustum(), FS_LEFT
    'Bottom
    Frustum(FS_BOTTOM).A = clip.m14 + clip.m12
    Frustum(FS_BOTTOM).B = clip.m24 + clip.m22
    Frustum(FS_BOTTOM).c = clip.m34 + clip.m32
    Frustum(FS_BOTTOM).d = clip.m44 + clip.m42
    NormalizePlane Frustum(), FS_BOTTOM
    'Top
    Frustum(FS_TOP).A = clip.m14 - clip.m12
    Frustum(FS_TOP).B = clip.m24 - clip.m22
    Frustum(FS_TOP).c = clip.m34 - clip.m32
    Frustum(FS_TOP).d = clip.m44 - clip.m42
    NormalizePlane Frustum(), FS_TOP
    'Back
    Frustum(FS_BACK).A = clip.m14 - clip.m13
    Frustum(FS_BACK).B = clip.m24 - clip.m23
    Frustum(FS_BACK).c = clip.m34 - clip.m33
    Frustum(FS_BACK).d = clip.m44 - clip.m43
    NormalizePlane Frustum(), FS_BACK
    'Front
    Frustum(FS_FRONT).A = clip.m14 + clip.m13
    Frustum(FS_FRONT).B = clip.m24 + clip.m23
    Frustum(FS_FRONT).c = clip.m34 + clip.m33
    Frustum(FS_FRONT).d = clip.m44 + clip.m43
    NormalizePlane Frustum(), FS_FRONT

    CopyMemory lpFRUST.PLANE(0), Frustum(0), Len(lpFRUST)

End Sub


'----------------------------------------
'Name: NormalizePlane
'----------------------------------------
Private Function NormalizePlane(aFrustum() As D3DPLANE, side As Long)
Dim magnitude
magnitude = Sqr(aFrustum(side).A * aFrustum(side).A + _
                aFrustum(side).B * aFrustum(side).B + _
                aFrustum(side).c * aFrustum(side).c)

'Then we divide the plane's values by it's magnitude.
'This makes it easier to work with.
aFrustum(side).A = aFrustum(side).A / magnitude
aFrustum(side).B = aFrustum(side).B / magnitude
aFrustum(side).c = aFrustum(side).c / magnitude
On Error Resume Next
aFrustum(side).d = aFrustum(side).d / magnitude
End Function


'----------------------------------------
'Name: FileiS_valid
'Object: FileiS
'Event: valid
'----------------------------------------
Function FileiS_valid(ByVal Filename As String) As Boolean

  Dim WFD As WIN32_FIND_DATA ':( Duplicated Name
  Dim hFile As Long ':( Duplicated Name
  Dim fn As String

    If Right$(Filename, 1) <> Chr$(0) Then
        fn = Filename & Chr$(0)
      Else 'NOT RIGHT$(FILENAME,...
        fn = Filename
    End If
    hFile = FindFirstFile(Filename, WFD)
    FileiS_valid = (hFile <> INVALID_HANDLE_VALUE)
    FindClose hFile

End Function


'----------------------------------------
'Name: Sleep
'----------------------------------------
Sub Sleep(Sec)

  Dim KK ':( As Variant ?

    For KK = 0 To Sec * 5000
        DoEvents
    Next KK

End Sub


'----------------------------------------
'Name: PathWin32
'----------------------------------------
Function PathWin32(Path As String) As String

    Path = Mid$(Path, 1, InStr(Path, Chr$(0)) - 1)
    PathWin32 = Replace(Path, "/", "\")

End Function


'----------------------------------------
'Name: Make_NemoVerxtex
'Object: Make
'Event: NemoVerxtex
'----------------------------------------
Function Make_NemoVerxtex(x, y, z, Nx, Ny, Nz, tu, tv) As NEMO_VERTEX

    Make_NemoVerxtex.x = x
    Make_NemoVerxtex.y = y
    Make_NemoVerxtex.z = z

    Make_NemoVerxtex.Nx = Nx
    Make_NemoVerxtex.Ny = Ny
    Make_NemoVerxtex.Nz = Nz

    Make_NemoVerxtex.tu = tu
    Make_NemoVerxtex.tv = tv

End Function


'----------------------------------------
'Name: Mtransform
'----------------------------------------
Function Mtransform(Mat As D3DMATRIX, _
                  vSrc As D3DVECTOR _
                  ) As D3DVECTOR

    D3DXVec3TransformCoord Mtransform, vSrc, Mat

Exit Function

  Dim vDest As D3DVECTOR ':( Move line to top of current Function
  Dim x, y, z, w ':( As Variant ?':( As Variant ?':( Move line to top of current Function

    x = vSrc.x * Mat.m11 + vSrc.y * Mat.m21 + vSrc.z * Mat.m31 + Mat.m41
    y = vSrc.x * Mat.m12 + vSrc.y * Mat.m22 + vSrc.z * Mat.m32 + Mat.m42
    z = vSrc.x * Mat.m13 + vSrc.y * Mat.m23 + vSrc.z * Mat.m33 + Mat.m43
    w = vSrc.x * Mat.m14 + vSrc.y * Mat.m24 + vSrc.z * Mat.m34 + Mat.m44

    If Abs(w) < 0.0000000001 Then

        Exit Function '>---> Bottom
    End If

    vDest.x = x / w
    vDest.y = y / w
    vDest.z = z / w

    Mtransform = vDest

End Function


'----------------------------------------
'Name: VICTOR
'----------------------------------------
Public Function VICTOR(A As Single, B As Single, c As Single) As D3DVECTOR

  Dim vecOut As D3DVECTOR

    With vecOut
        .x = A
        .y = B
        .z = c
    End With 'VECOUT

    VICTOR = vecOut

End Function


'----------------------------------------
'Name: Vector
'----------------------------------------
Function Vector(x, y, z) As D3DVECTOR

    On Error Resume Next
    Vector.x = x
    Vector.y = y
    Vector.z = z

End Function


'----------------------------------------
'Name: VectorN
'----------------------------------------
Function VectorN(x, y, z) As NemoVECTOR

    VectorN.x = x
    VectorN.y = y
    VectorN.z = z

End Function


'----------------------------------------
'Name: VectorN2
'----------------------------------------
Function VectorN2(VEC As D3DVECTOR) As NemoVECTOR

    VectorN2.x = VEC.x
    VectorN2.y = VEC.y
    VectorN2.z = VEC.z

End Function


'----------------------------------------
'Name: VlenghtABS
'----------------------------------------
Function VlenghtABS(V As D3DVECTOR) ':( As Variant ?

    VlenghtABS = Sqr(V.x * V.x) + _
                 (V.y * V.y) + Sqr(V.z * V.z)

End Function


'----------------------------------------
'Name: Vlenght2
'----------------------------------------
Function Vlenght2(V As D3DVECTOR) ':( As Variant ?

    Vlenght2 = (V.x * V.x) + _
               (V.y * V.y) + (V.z * V.z)

End Function


'----------------------------------------
'Name: Vaverage
'----------------------------------------
Function Vaverage(V As D3DVECTOR) ':( As Variant ?

    Vaverage = (V.x + V.y + V.z) / 3

End Function


'----------------------------------------
'Name: VectorRADtoDEG
'----------------------------------------
Function VectorRADtoDEG(VEC As D3DVECTOR) As D3DVECTOR

    VectorRADtoDEG.x = VEC.x / RAD
    VectorRADtoDEG.y = VEC.y / RAD
    VectorRADtoDEG.z = VEC.z / RAD

End Function


'----------------------------------------
'Name: VectorDEGtoRAD
'----------------------------------------
Function VectorDEGtoRAD(VEC As D3DVECTOR) As D3DVECTOR

    VectorDEGtoRAD.x = VEC.x * RAD
    VectorDEGtoRAD.y = VEC.y * RAD
    VectorDEGtoRAD.z = VEC.z * RAD

End Function


'----------------------------------------
'Name: Vector_Dir
'Object: Vector
'Event: Dir
'----------------------------------------
Function Vector_Dir(yaw, pitch) As D3DVECTOR

    Vector_Dir.x = -Sin(yaw * RAD) * Cos(pitch * RAD)
    Vector_Dir.y = Sin(pitch * RAD)
    Vector_Dir.z = Cos(pitch * RAD) * Cos(yaw * RAD)

End Function

'Function Vector_Dir2(V) As D3DVECTOR
'
'          Vector_Dir.x = -Sin(yaw * RAD) * Cos(pitch * RAD)
'          Vector_Dir.y = Sin(pitch * RAD)
'          Vector_Dir.z = Cos(pitch * RAD) * Cos(yaw * RAD)
'
'
'End Function


'----------------------------------------
'Name: Vector_Lerp
'Object: Vector
'Event: Lerp
'----------------------------------------
Function Vector_Lerp(Vstart As D3DVECTOR, Vend As D3DVECTOR, Val) As D3DVECTOR

  Dim T ':( As Variant ?

    T = 1 - Val

    Vector_Lerp.x = (T * Vstart.x) + (Val * Vend.x)
    Vector_Lerp.y = (T * Vstart.y) + (Val * Vend.y)
    Vector_Lerp.z = (T * Vstart.z) + (Val * Vend.z)

End Function


'----------------------------------------
'Name: VectorAddAndScale
'----------------------------------------
Sub VectorAddAndScale(dest As D3DVECTOR, S1 As Single, v1 As D3DVECTOR, s2 As Single, v2 As D3DVECTOR)

    dest.x = S1 * v1.x + s2 * v2.x
    dest.y = S1 * v1.y + s2 * v2.y
    dest.z = S1 * v1.z + s2 * v2.z

End Sub

'=================================
' VectorCopy
'=================================
'----------------------------------------
'Name: VectorCopy
'----------------------------------------
Sub VectorCopy(dest As D3DVECTOR, src As D3DVECTOR)

    dest.x = src.x
    dest.y = src.y
    dest.z = src.z

End Sub


'----------------------------------------
'Name: RVertex
'----------------------------------------
Public Function RVertex(x, y, z, Nx, Ny, Nz, tu, tv) As D3DVERTEX

    With RVertex
        .x = x
        .y = y
        .z = z
        .Nx = Nx
        .Ny = Ny
        .Nz = Nz
        .tu = tu
        .tv = tv
    End With 'RVERTEX

End Function


'----------------------------------------
'Name: RVertex2
'----------------------------------------
Public Function RVertex2(x, y, z, Nx, Ny, Nz, tu, tv, tu2, tv2) As D3DVERTEX2

    With RVertex2
        .x = x
        .y = y
        .z = z
        .Nx = Nx
        .Ny = Ny
        .Nz = Nz
        .tu1 = tu
        .tv1 = tv
        .tu2 = tu2
        .tv2 = tv2

    End With 'RVERTEX2

End Function


'----------------------------------------
'Name: LVertex
'----------------------------------------
Public Function LVertex(x, y, z, color, tu, tv) As D3DLVERTEX

    With LVertex
        .x = x
        .y = y
        .z = z
        .color = color
        .Specular = color
        .tu = tu
        .tv = tv
    End With 'LVERTEX

End Function


'----------------------------------------
'Name: TLVERTEX
'----------------------------------------
Function TLVERTEX(x, y, z, color, tu, tv) As D3DTLVERTEX

    TLVERTEX.Sx = x
    TLVERTEX.Sy = y
    TLVERTEX.SZ = z
    TLVERTEX.color = color
    TLVERTEX.Specular = color
    TLVERTEX.rhw = 0.5
    TLVERTEX.tu = tu
    TLVERTEX.tv = tv

End Function


'----------------------------------------
'Name: Write_VERTEX
'Object: Write
'Event: VERTEX
'----------------------------------------
Function Write_VERTEX(STfile As String, lpVERT() As D3DVERTEX)

  Dim i, j, z, fFIL ':( As Variant ?

    i = LBound(lpVERT)
    j = UBound(lpVERT)

    fFIL = FreeFile
  Dim ST As String ':( Move line to top of current Function

    Open STfile For Output As #fFIL

    For z = i To j
        ST = Str(lpVERT(z).x) + "," + Str(lpVERT(z).y) + "," + Str(lpVERT(z).z) + "," + Str(lpVERT(z).tu) + "," + Str(lpVERT(z).tv) + ", _"

        Print #fFIL, ST

    Next z

    Close #fFIL

End Function


'----------------------------------------
'Name: Write_INdices
'Object: Write
'Event: INdices
'----------------------------------------
Function Write_INdices(STfile As String, lpVERT() As Integer)

  Dim i, j, z, fFIL, NUMin, Step, K, L ':( As Variant ?

    i = LBound(lpVERT)
    j = UBound(lpVERT)

    NUMin = j

    Step = Int(NUMin / 6)

    fFIL = FreeFile
  Dim ST As String ':( Move line to top of current Function

    Open STfile For Output As #fFIL

    For L = 1 To Step
        ST = Str(lpVERT(z)) + "," + Str(lpVERT(z + 1)) + "," + Str(lpVERT(z + 2)) + "," + Str(lpVERT(z + 3)) + "," + Str(lpVERT(z + 4)) + "," + Str(lpVERT(z + 5)) + ", _"

        Print #fFIL, ST

        z = z + 6
    Next L

  Dim RESTE ':( As Variant ?':( Move line to top of current Function
    RESTE = NUMin - (6 * Step)
    If RESTE < 1 Then GoTo finiT ':( Expand Structure
    For K = 1 To RESTE
        ST = Str(lpVERT(z + K))

        Print #fFIL, ST

    Next K

finiT:

    Close #fFIL

End Function


'----------------------------------------
'Name: Floor
'----------------------------------------
Public Function Floor(value As Single) As Single

    If Int(value) < value Then
        Floor = Int(value)
      Else 'NOT INT(VALUE)...
        Floor = value
    End If

End Function


'----------------------------------------
'Name: Ceil
'----------------------------------------
Public Function Ceil(value As Single) As Single

    If Int(value) < value Then
        Ceil = Int(value) + 1
      Else 'NOT INT(VALUE)...
        Ceil = value
    End If

End Function


'----------------------------------------
'Name: FtoDW
'----------------------------------------
Function FtoDW(F As Single) As Long

  Dim buf As D3DXBuffer
  Dim L As Long
  Dim D2dX As New D3DX8

    Set buf = D2dX.CreateBuffer(4)
    D2dX.BufferSetData buf, 0, 4, 1, F
    D2dX.BufferGetData buf, 0, 4, 1, L
    FtoDW = L
    Set D2dX = Nothing

End Function

'-----------------------------------------------------------------------------
' Name: D3DCOLORVALUEtoLONG
'-----------------------------------------------------------------------------


'----------------------------------------
'Name: ColorValue4
'----------------------------------------
Function ColorValue4(A As Single, R As Single, G As Single, B As Single) As D3DCOLORVALUE
    Dim c As D3DCOLORVALUE
    c.A = A
    c.R = R
    c.G = G
    c.B = B
    ColorValue4 = c
End Function


'----------------------------------------
'Name: D3DCOLORVALUEtoLONG
'----------------------------------------
Function D3DCOLORVALUEtoLONG(cv As D3DCOLORVALUE) As Long

  Dim R As Long
  Dim G As Long
  Dim B As Long
  Dim A As Long
  Dim c As Long

    R = cv.R * 255
    G = cv.G * 255
    B = cv.B * 255
    A = cv.A * 255

    If A > 127 Then
        A = A - 128
        c = A * 2 ^ 24 Or &H80000000
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or B
      Else 'NOT A...
        c = A * 2 ^ 24
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or B
    End If

    D3DCOLORVALUEtoLONG = c

End Function

'-----------------------------------------------------------------------------
' Name: LONGtoD3DCOLORVALUE
'-----------------------------------------------------------------------------
'----------------------------------------
'Name: LONGtoD3DCOLORVALUE
'----------------------------------------
Function LONGtoD3DCOLORVALUE(color As Long) As D3DCOLORVALUE

  Dim A As Long, R As Long, G As Long, B As Long

    If color < 0 Then
        A = ((color And (&H7F000000)) / (2 ^ 24)) Or &H80&
      Else 'NOT COLOR...
        A = color / (2 ^ 24)
    End If
    R = (color And &HFF0000) / (2 ^ 16)
    G = (color And &HFF00&) / (2 ^ 8)
    B = (color And &HFF&)

    LONGtoD3DCOLORVALUE.A = A / 255
    LONGtoD3DCOLORVALUE.R = R / 255
    LONGtoD3DCOLORVALUE.G = G / 255
    LONGtoD3DCOLORVALUE.B = B / 255

End Function


'----------------------------------------
'Name: GetSquaredSize
'----------------------------------------
Sub GetSquaredSize(ByRef CWidth As Long, ByRef CHeight As Long)

    If CWidth > 2 And CWidth <= 4 Then CWidth = 4 ':( Expand Structure
    If CWidth > 4 And CWidth <= 8 Then CWidth = 8 ':( Expand Structure
    If CWidth > 8 And CWidth <= 16 Then CWidth = 16 ':( Expand Structure
    If CWidth > 16 And CWidth <= 32 Then CWidth = 32 ':( Expand Structure
    If CWidth > 32 And CWidth <= 64 Then CWidth = 64 ':( Expand Structure
    If CWidth > 64 And CWidth <= 128 Then CWidth = 128 ':( Expand Structure
    If CWidth > 128 Then CWidth = 256 ':( Expand Structure
    If CHeight > 2 And CHeight <= 4 Then CHeight = 4 ':( Expand Structure
    If CHeight > 4 And CHeight <= 8 Then CHeight = 8 ':( Expand Structure
    If CHeight > 8 And CHeight <= 16 Then CHeight = 16 ':( Expand Structure
    If CHeight > 16 And CHeight <= 32 Then CHeight = 32 ':( Expand Structure
    If CHeight > 32 And CHeight <= 64 Then CHeight = 64 ':( Expand Structure
    If CHeight > 64 And CHeight <= 128 Then CHeight = 128 ':( Expand Structure
    If CHeight > 128 Then CHeight = 256 ':( Expand Structure

End Sub


'----------------------------------------
'Name: DXColor
'----------------------------------------
Function DXColor(R, G, B, A) As D3DCOLORVALUE

    DXColor.A = A
    DXColor.B = B
    DXColor.G = G
    DXColor.R = R

End Function


'----------------------------------------
'Name: DXCToL
'----------------------------------------
Function DXCToL(cv As D3DCOLORVALUE) As Long

  Dim R As Long
  Dim G As Long
  Dim B As Long
  Dim A As Long
  Dim c As Long

    R = cv.R * 255
    G = cv.G * 255
    B = cv.B * 255
    A = cv.A * 255
    If R > 255 Or G > 255 Or B > 255 Or A > 255 Then Exit Function ':( Expand Structure or consider reversing Condition
    If A > 127 Then
        A = A - 128

        c = A * 2 ^ 24 Or &H80000000
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or B
      Else 'NOT A...
        c = A * 2 ^ 24
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or B
    End If

    DXCToL = c

End Function


'----------------------------------------
'Name: CreateTextureFromBuffer
'----------------------------------------
Function CreateTextureFromBuffer(BufferByte() As Byte, Wi, HI) As Direct3DBaseTexture8

  Dim SURF As Direct3DBaseTexture8
  Dim MAP() As Byte

    ReDim MAP(CLng(Wi * HI * 3) + 54)

    With BMH
        '.intMagic = 19778
        .lngSize = (Wi * HI * 3) + 54
        .lngOffset = 54

        .biBitCount = 24
        .biWidth = Wi
        .biHeight = HI
        .biSize = 40
        .biPlanes = 1
        .biSizeImage = Wi * HI * 3
        .biCompression = 0
        .biXPelsPerMeter = 50
        .biYPelsPerMeter = 50
    End With 'BMH

  Dim BMPheader(Len(BMH) + 1) As Byte ':( Move line to top of current Function
    CopyMemory BMPheader(2), BMH, Len(BMH)
    BMPheader(0) = 66
    BMPheader(1) = 77

    'Copy  data into an array
    CopyMemory MAP(0), BMPheader(0), 54
    CopyMemory MAP(54), BufferByte(0), Wi * HI * 3

    

    Set SURF = D3DX.CreateTextureFromFileInMemoryEx(D3dDevice, MAP(0), UBound(MAP()), Wi, HI, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &H0, ByVal 0, ByVal 0)

    Set CreateTextureFromBuffer = SURF

End Function


'----------------------------------------
'Name: CreateTextureFromWAL
'----------------------------------------



'----------------------------------------
'Name: CreateTextureFromWALBuffer
'----------------------------------------
Function CreateTextureFromWALBuffer(WalFIleRawByte() As Byte, ByRef RetWI, ByRef RetHI) As Direct3DBaseTexture8

End Function


'----------------------------------------
'Name: SaveFile
'----------------------------------------
Sub SaveFile(lpBUFFER() As Byte, File As String)

    Open File For Binary As 255
    'Write the byte-array
    Put 255, , lpBUFFER()
    Close 255

End Sub


'----------------------------------------
'Name: CreateNemoVertex
'----------------------------------------
Sub CreateNemoVertex(x, y, z, Nx, Ny, Nz, tu, tv, lpVERT As NEMO_VERTEX)
 With lpVERT
        .Nx = Nx
        .Ny = Ny
        .Nz = Nz
        .tu = tu
        .tv = tv
        .x = x
        .y = y
        .z = z
    End With 'RET

End Sub


'----------------------------------------
'Name: CopyToNemoVector
'----------------------------------------
Sub CopyToNemoVector(DestVec As NemoVECTOR, OrgVec As D3DVECTOR)
  CopyMemory DestVec, OrgVec, Len(DestVec)
  
End Sub




'----------------------------------------
'Name: RenderWireFrameBOX
'----------------------------------------
Sub RenderWireFrameBOX(NemoBoxMin As NemoVECTOR, NemoBoxMax As NemoVECTOR)
Dim VERT() As NEMO_VERTEX
CreateCubeEX VERT(), NemoBoxMin, NemoBoxMax

GLOB.LpGLOBAL_NEMO.Set_EngineFillMode NEMO_FILL_WIREFRAME

 RenderCubeEX VERT

GLOB.LpGLOBAL_NEMO.Set_EngineFillMode NEMO_FILL_SOLID


End Sub



'----------------------------------------
'Name: RenderCubeEX
'----------------------------------------
Sub RenderCubeEX(Vertices() As NEMO_VERTEX)
  
  D3dDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertices(0), Len(Vertices(0))
  D3dDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertices(4), Len(Vertices(0))
  
  D3dDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertices(8), Len(Vertices(0))
  D3dDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertices(12), Len(Vertices(0))
  
  D3dDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertices(16), Len(Vertices(0))
  D3dDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertices(20), Len(Vertices(0))
  
  
  
End Sub



'----------------------------------------
'Name: CreateCubeEX
'----------------------------------------
Sub CreateCubeEX(Vertices() As NEMO_VERTEX, Vmin As NemoVECTOR, Vmax As NemoVECTOR, Optional TileX = 1, Optional TileY = 1)
     ReDim Preserve Vertices(23)
      Dim v1 As D3DVECTOR
     Dim v2 As D3DVECTOR
     Dim Sx, Sy, SZ
    
     
     'V1 = vectorn(
     'V2 = Vector(Sx / 2, Sy, Sz / 2)
     
     CopyNemoToD3Dvec Vmin, v1
     CopyNemoToD3Dvec Vmax, v2
     
           
    ' Create vertices describing the front face of the cube.
    CreateNemoVertex v1.x, v2.y, v1.z, 0, 0, -1, 0, 0, Vertices(0)
    CreateNemoVertex v2.x, v2.y, v1.z, 0, 0, -1, TileX, 0, Vertices(1)
    CreateNemoVertex v1.x, v1.y, v1.z, 0, 0, -1, 0, TileY, Vertices(2)
    CreateNemoVertex v2.x, v1.y, v1.z, 0, 0, -1, TileX, TileY, Vertices(3)
        
    ' Create vertices describing the back face of the cube.
    CreateNemoVertex v1.x, v2.y, v2.z, 0, 0, 1, TileX, 0, Vertices(4)
    CreateNemoVertex v1.x, v1.y, v2.z, 0, 0, 1, TileX, TileY, Vertices(5)
    CreateNemoVertex v2.x, v2.y, v2.z, 0, 0, 1, 0, 0, Vertices(6)
    CreateNemoVertex v2.x, v1.y, v2.z, 0, 0, 1, 0, TileY, Vertices(7)
        
    ' Create vertices describing the top face of the cube.
    CreateNemoVertex v1.x, v2.y, v2.z, 0, 1, 0, 0, 0, Vertices(8)
    CreateNemoVertex v2.x, v2.y, v2.z, 0, 1, 0, TileX, 0, Vertices(9)
    CreateNemoVertex v1.x, v2.y, v1.z, 0, 1, 0, 0, TileY, Vertices(10)
    CreateNemoVertex v2.x, v2.y, v1.z, 0, 1, 0, TileX, TileY, Vertices(11)
        
    ' Create vertices describing the bottom face of the cube.
    CreateNemoVertex v1.x, v1.y, v2.z, 0, -1, 0, 0, 0, Vertices(12)
    CreateNemoVertex v1.x, v1.y, v1.z, 0, -1, 0, 0, TileY, Vertices(13)
    CreateNemoVertex v2.x, v1.y, v2.z, 0, -1, 0, TileX, 0, Vertices(14)
    CreateNemoVertex v2.x, v1.y, v1.z, 0, -1, 0, TileX, TileY, Vertices(15)
        
    ' Create vertices describing the right face of the cube.
    CreateNemoVertex v2.x, v2.y, v1.z, 1, 0, 0, 0, 0, Vertices(16)
    CreateNemoVertex v2.x, v2.y, v2.z, 1, 0, 0, TileX, 0, Vertices(17)
    CreateNemoVertex v2.x, v1.y, v1.z, 1, 0, 0, 0, TileY, Vertices(18)
    CreateNemoVertex v2.x, v1.y, v2.z, 1, 0, 0, TileX, TileY, Vertices(19)
        
    ' Create vertices describing the left face of the cube.
    CreateNemoVertex v1.x, v2.y, v1.z, -1, 0, 0, TileX, 0, Vertices(20)
    CreateNemoVertex v1.x, v1.y, v1.z, -1, 0, 0, TileX, TileY, Vertices(21)
    CreateNemoVertex v1.x, v2.y, v2.z, -1, 0, 0, 0, 0, Vertices(22)
    CreateNemoVertex v1.x, v1.y, v2.z, -1, 0, 0, 0, TileY, Vertices(23)
    
    
End Sub




'----------------------------------------
'Name: RGBA
'----------------------------------------
Function RGBA(R, G, B, A) As Long

    RGBA = DXCToL(DXColor(R, G, B, A))

End Function



'----------------------------------------
'Name: CopyNemoToD3Dvec
'----------------------------------------
Sub CopyNemoToD3Dvec(vNem As NemoVECTOR, Vd3d As D3DVECTOR)
   CopyMemory Vd3d, vNem, Len(vNem)
   
End Sub
'----------------------------------------
'Name: CopyD3dToNemvec
'----------------------------------------
Sub CopyD3dToNemvec(vNem As NemoVECTOR, Vd3d As D3DVECTOR)
   CopyMemory vNem, Vd3d, Len(vNem)
   
End Sub


'----------------------------------------
'Name: CreateMatrix
'----------------------------------------
Function CreateMatrix(RotDEGVec As D3DVECTOR, ScalVec As D3DVECTOR, TransVec As D3DVECTOR) As D3DMATRIX ':( Missing Scope

  Dim CosRx As Single, CosRy As Single, CosRz As Single
  Dim SinRx As Single, SinRy As Single, SinRz As Single

  Dim Rx As Single, Ry As Single, Rz As Single, Sx As Single, _
          Sy As Single, SZ As Single, Tx As Single, Ty As Single, Tz As Single

    Rx = RotDEGVec.x
    Ry = RotDEGVec.y
    Rz = RotDEGVec.z

    Sx = ScalVec.x
    Sy = ScalVec.y
    SZ = ScalVec.z

    Tx = TransVec.x
    Ty = TransVec.y
    Tz = TransVec.z

  Dim Zpi ':( As Variant ?':( Move line to top of current Function':( As Variant ?':( Move line to top of current Function
    Zpi = 3.14564875441112
  Dim Zrad ':( As Variant ?':( Move line to top of current Function':( As Variant ?':( Move line to top of current Function
    Zrad = Zpi / 180

    CosRx = Cos(Rx * Zrad) 'Used 6x
    CosRy = Cos(Ry * Zrad) 'Used 4x
    CosRz = Cos(Rz * Zrad) 'Used 4x
    SinRx = Sin(Rx * Zrad) 'Used 5x
    SinRy = Sin(Ry * Zrad) 'Used 5x
    SinRz = Sin(Rz * Zrad) 'Used 5x

    With CreateMatrix
        .m11 = (Sx * CosRy * CosRz)
        .m12 = (Sx * CosRy * SinRz)
        .m13 = -(Sx * SinRy)

        .m21 = -(Sy * CosRx * SinRz) + (Sy * SinRx * SinRy * CosRz)
        .m22 = (Sy * CosRx * CosRz) + (Sy * SinRx * SinRy * SinRz)
        .m23 = (Sy * SinRx * CosRy)

        .m31 = (SZ * SinRx * SinRz) + (SZ * CosRx * SinRy * CosRz)
        .m32 = -(SZ * SinRx * CosRx) + (SZ * CosRx * SinRy * SinRz)
        .m33 = (SZ * CosRx * CosRy)

        .m41 = Tx
        .m42 = Ty
        .m43 = Tz
        .m44 = 1#
    End With 'CREATEMATRIX

    'Set DXc = Nothing

End Function '


'----------------------------------------
'Name: MatToNemoMATRIX
'----------------------------------------
Function MatToNemoMATRIX(Mat As D3DMATRIX) As NemoMATRIX

    With MatToNemoMATRIX
        .m11 = Mat.m11
        .m12 = Mat.m12
        .m13 = Mat.m13
        .m14 = Mat.m14

        .m21 = Mat.m21
        .m22 = Mat.m22
        .m23 = Mat.m23
        .m24 = Mat.m24

        .m31 = Mat.m31
        .m32 = Mat.m32
        .m33 = Mat.m33
        .m34 = Mat.m34

        .m41 = Mat.m41
        .m42 = Mat.m42
        .m43 = Mat.m43
        .m44 = Mat.m44

    End With 'MATTONEMOMATRIX

End Function


'----------------------------------------
'Name: VecToNemoVECTOR
'----------------------------------------
Function VecToNemoVECTOR(VEC As D3DVECTOR) As NemoVECTOR

    With VecToNemoVECTOR
        .x = VEC.x
        .y = VEC.y
        .z = VEC.z
    End With 'VECTONEMOVECTOR

End Function


'----------------------------------------
'Name: NemotoD3DVECTOR
'----------------------------------------
Function NemotoD3DVECTOR(VEC As NemoVECTOR) As D3DVECTOR

    With NemotoD3DVECTOR
        .x = VEC.x
        .y = VEC.y
        .z = VEC.z
    End With 'NEMOTOD3DVECTOR

End Function


'----------------------------------------
'Name: NemoVerttoD3DVECTOR
'----------------------------------------
Function NemoVerttoD3DVECTOR(VEC As NEMO_VERTEX) As D3DVECTOR

    With NemoVerttoD3DVECTOR
        .x = VEC.x
        .y = VEC.y
        .z = VEC.z
    End With 'NEMOTOD3DVECTOR

End Function


'----------------------------------------
'Name: NemoVert2toD3DVECTOR
'----------------------------------------
Function NemoVert2toD3DVECTOR(VEC As NEMO_VERTEX2) As D3DVECTOR

    With NemoVert2toD3DVECTOR
        .x = VEC.x
        .y = VEC.y
        .z = VEC.z
    End With 'NEMOTOD3DVECTOR

End Function

'Function VecToNemoVECTOR(Vec As D3DVECTOR) As NemoVECTOR
'    CopyMemory VecToNemoVECTOR, Vec, Len(Vec)
'
'End Function


'----------------------------------------
'Name: MatrixN
'----------------------------------------
Function MatrixN(RotRadVec As D3DVECTOR, ScalVec As D3DVECTOR, TransVec As D3DVECTOR) As NemoMATRIX

  'Dim M As NemoMATRIX
  
  Dim M0 As D3DMATRIX

    M0 = CreateMatrixEX(RotRadVec, ScalVec, TransVec)

    MatrixN = GLOB.Get_NemoMATRIX(M0)

End Function


'----------------------------------------
'Name: CreateMatrixEX
'----------------------------------------
Function CreateMatrixEX(RotRadVec As D3DVECTOR, ScalVec As D3DVECTOR, TransVec As D3DVECTOR) As D3DMATRIX ':( Missing Scope

  Dim CosRx As Single, CosRy As Single, CosRz As Single
  Dim SinRx As Single, SinRy As Single, SinRz As Single

  Dim Rx As Single, Ry As Single, Rz As Single, Sx As Single, _
          Sy As Single, SZ As Single, Tx As Single, Ty As Single, Tz As Single

    Rx = RotRadVec.x
    Ry = RotRadVec.y
    Rz = RotRadVec.z

    Sx = ScalVec.x
    Sy = ScalVec.y
    SZ = ScalVec.z

    Tx = TransVec.x
    Ty = TransVec.y
    Tz = TransVec.z

  Dim Zpi ':( As Variant ?':( Move line to top of current Function':( As Variant ?':( Move line to top of current Function
    Zpi = 3.14564875441112
  Dim Zrad ':( As Variant ?':( Move line to top of current Function':( As Variant ?':( Move line to top of current Function
    Zrad = Zpi / 180

    CosRx = Cos(Rx)  'Used 6x
    CosRy = Cos(Ry)  'Used 4x
    CosRz = Cos(Rz)  'Used 4x
    SinRx = Sin(Rx)  'Used 5x
    SinRy = Sin(Ry)  'Used 5x
    SinRz = Sin(Rz)  'Used 5x

    With CreateMatrixEX
        .m11 = (Sx * CosRy * CosRz)
        .m12 = (Sx * CosRy * SinRz)
        .m13 = -(Sx * SinRy)

        .m21 = -(Sy * CosRx * SinRz) + (Sy * SinRx * SinRy * CosRz)
        .m22 = (Sy * CosRx * CosRz) + (Sy * SinRx * SinRy * SinRz)
        .m23 = (Sy * SinRx * CosRy)

        .m31 = (SZ * SinRx * SinRz) + (SZ * CosRx * SinRy * CosRz)
        .m32 = -(SZ * SinRx * CosRx) + (SZ * CosRx * SinRy * SinRz)
        .m33 = (SZ * CosRx * CosRy)

        .m41 = Tx
        .m42 = Ty
        .m43 = Tz
        .m44 = 1#
    End With 'CREATEMATRIX'CREATEMATRIXEX

    'Set DXc = Nothing

End Function '


'----------------------------------------
'Name: OBJECT_Scale_Rotate_MOV
'Object: OBJECT
'Event: Scale_Rotate_MOV
'----------------------------------------
Sub OBJECT_Scale_Rotate_MOV(Xscal, Yscal, Zscal, Xrot, Yrot, Zrot, Xmov, Ymov, Zmov)

  Dim MatZ As D3DMATRIX
  Dim ROTz As D3DMATRIX
  Dim MOVz As D3DMATRIX
  Dim TEMPz As D3DMATRIX

    D3DXMatrixIdentity MatZ
    D3DXMatrixIdentity MOVz
    D3DXMatrixIdentity ROTz

    Call D3DXMatrixScaling(MatZ, Xscal, Yscal, Zscal)
    Call MRotate(ROTz, Xrot, Yrot, Zrot)
    Call D3DXMatrixTranslation(MOVz, Xmov, Ymov, Zmov)

    D3DXMatrixMultiply TEMPz, MatZ, ROTz
    D3DXMatrixMultiply MatZ, TEMPz, MOVz

End Sub


'----------------------------------------
'Name: Matrix_Get
'Object: Matrix
'Event: Get
'----------------------------------------
Function Matrix_Get(Xscal, Yscal, Zscal, Xrot, Yrot, Zrot, Xmov, Ymov, Zmov) As D3DMATRIX

  Dim MatZ As D3DMATRIX
  Dim ROTz As D3DMATRIX
  Dim MOVz As D3DMATRIX
  Dim TEMPz As D3DMATRIX

    D3DXMatrixIdentity MatZ
    D3DXMatrixIdentity MOVz
    D3DXMatrixIdentity ROTz

    Call D3DXMatrixScaling(MatZ, Xscal, Yscal, Zscal)
    Call MRotate(ROTz, Xrot, Yrot, Zrot)
    Call D3DXMatrixTranslation(MOVz, Xmov, Ymov, Zmov)

    D3DXMatrixMultiply TEMPz, MatZ, ROTz
    D3DXMatrixMultiply MatZ, TEMPz, MOVz

    Matrix_Get = MatZ

End Function


'----------------------------------------
'Name: Matrix_GetEX
'Object: Matrix
'Event: GetEX
'----------------------------------------
Function Matrix_GetEX(Vscal As D3DVECTOR, Vrot As D3DVECTOR, Vtrans As D3DVECTOR) As D3DMATRIX

  Dim MatZ As D3DMATRIX
  Dim ROTz As D3DMATRIX
  Dim MOVz As D3DMATRIX
  Dim TEMPz As D3DMATRIX

    D3DXMatrixIdentity MatZ
    D3DXMatrixIdentity MOVz
    D3DXMatrixIdentity ROTz

    Call D3DXMatrixScaling(MatZ, Vscal.x, Vscal.y, Vscal.z)
    Call MRotate(ROTz, Vrot.x, Vrot.y, Vrot.z)
    Call D3DXMatrixTranslation(MOVz, Vtrans.x, Vtrans.y, Vtrans.z)

    D3DXMatrixMultiply TEMPz, MatZ, ROTz
    D3DXMatrixMultiply MatZ, TEMPz, MOVz

    Matrix_GetEX = MatZ

End Function


'----------------------------------------
'Name: VectorMatrixMultiply
'----------------------------------------
Public Function VectorMatrixMultiply(vDest As D3DVECTOR, _
                                     vSrc As D3DVECTOR, _
                                     Mat As D3DMATRIX) ':( As Variant ?

  Dim x, y, z, w ':( As Variant ?':( As Variant ?

    x = vSrc.x * Mat.m11 + vSrc.y * Mat.m21 + vSrc.z * Mat.m31 + Mat.m41
    y = vSrc.x * Mat.m12 + vSrc.y * Mat.m22 + vSrc.z * Mat.m32 + Mat.m42
    z = vSrc.x * Mat.m13 + vSrc.y * Mat.m23 + vSrc.z * Mat.m33 + Mat.m43
    w = vSrc.x * Mat.m14 + vSrc.y * Mat.m24 + vSrc.z * Mat.m34 + Mat.m44

    If Abs(w) < 0.0000000001 Then
        VectorMatrixMultiply = 0
        Exit Function '>---> Bottom
    End If

    vDest.x = x / w
    vDest.y = y / w
    vDest.z = z / w

    VectorMatrixMultiply = 1

End Function


'---------------------------------------_________
'Name: VerTEX_MatMULTILPY
'Object: VerTEX
'Event: MatMULTILPY
'----------------------------------------
Function VerTEX_MatMULTILPY(VERT() As D3DVERTEX, MATA As D3DMATRIX, retVert() As D3DVERTEX)

  Dim i, j, z ':( As Variant ?

    i = LBound(VERT)
    j = UBound(VERT)
    'Dim retVert() As D3DVERTEX
    ReDim Preserve retVert(i To j) As D3DVERTEX
    retVert = VERT
  Dim VC As D3DVECTOR ':( Move line to top of current Function
    For z = i To j
        VC = VICTOR(retVert(z).x, retVert(z).y, retVert(z).z)
        VectorMatrixMultiply VC, VC, MATA

        retVert(z).x = VC.x '(retVert(Z).x * MATA.m11) + (retVert(Z).Y * MATA.m21) + (retVert(Z).Z * MATA.m31) + MATA.m41
        retVert(z).y = VC.y '(retVert(Z).x * MATA.m12) + (retVert(Z).Y * MATA.m22) + (retVert(Z).Z * MATA.m32) + MATA.m42
        retVert(z).z = VC.z '(retVert(Z).x * MATA.m13) + (retVert(Z).Y * MATA.m23) + (retVert(Z).Z * MATA.m33) + MATA.m43
    Next z

End Function


'----------------------------------------
'Name: VDist
'----------------------------------------
Public Function VDist(VA As D3DVECTOR, VB As D3DVECTOR) As Single

    VDist = Sqr((VB.x - VA.x) ^ 2 + (VB.y - VA.y) ^ 2 + (VB.z - VA.z) ^ 2)

End Function


'----------------------------------------
'Name: Vsub
'----------------------------------------
Public Function Vsub(A As D3DVECTOR, B As D3DVECTOR) As D3DVECTOR

  Dim dest As D3DVECTOR

    dest.x = A.x - B.x
    dest.y = A.y - B.y
    dest.z = A.z - B.z
    Vsub = dest

End Function


'----------------------------------------
'Name: Vmiddle
'----------------------------------------
Public Function Vmiddle(VA As D3DVECTOR, VB As D3DVECTOR) As D3DVECTOR

    Vmiddle.x = (VA.x + VB.x) / 2
    Vmiddle.y = (VA.y + VB.y) / 2
    Vmiddle.z = (VA.z + VB.z) / 2

End Function


'----------------------------------------
'Name: min
'----------------------------------------
Function min(v1, v2) ':( As Variant ?

    If v1 < v2 Then min = v1 ':( Expand Structure
    If v2 < v1 Then min = v2 ':( Expand Structure
    If v2 = v1 Then min = v1 ':( Expand Structure

End Function


'----------------------------------------
'Name: max
'----------------------------------------
Function max(v1, v2) ':( As Variant ?

    If v1 > v2 Then max = v1 ':( Expand Structure
    If v2 > v1 Then max = v2 ':( Expand Structure
    If v2 = v1 Then max = v1 ':( Expand Structure

End Function


'----------------------------------------
'Name: MAX_3
'Object: MAX
'Event: 3
'----------------------------------------
Function MAX_3(v1, v2, v3) ':( As Variant ?

    If v1 > v2 And v1 > v3 Then MAX_3 = v1 ':( Expand Structure
    If v2 > v1 And v2 > v3 Then MAX_3 = v2 ':( Expand Structure
    If v3 > v2 And v3 > v1 Then MAX_3 = v3 ':( Expand Structure

    If MAX_3 = 0 Then
        If v1 > v2 And v1 = v3 Then MAX_3 = v1 ':( Expand Structure
        If v2 > v3 And v2 = v1 Then MAX_3 = v2 ':( Expand Structure
        If v3 > v2 And v3 = v1 Then MAX_3 = v3 ':( Expand Structure
    End If

    If MAX_3 = 0 Then
        If v1 = v2 And v1 = v3 Then MAX_3 = v1 ':( Expand Structure

    End If

End Function


'----------------------------------------
'Name: Vmult
'----------------------------------------
Public Function Vmult(SrcVec As D3DVECTOR, ByVal Val) As D3DVECTOR

    Vmult.x = SrcVec.x * Val
    Vmult.y = SrcVec.y * Val
    Vmult.z = SrcVec.z * Val

End Function


'----------------------------------------
'Name: VAdd
'----------------------------------------
Public Function VAdd(A As D3DVECTOR, B As D3DVECTOR) As D3DVECTOR

  Dim dest As D3DVECTOR

    dest.x = A.x + B.x
    dest.y = A.y + B.y
    dest.z = A.z + B.z
    VAdd = dest

End Function


'----------------------------------------
'Name: VScale
'----------------------------------------
Public Function VScale(src As D3DVECTOR, S As Single) As D3DVECTOR

  Dim dest As D3DVECTOR

    dest.x = src.x * S
    dest.y = src.y * S
    dest.z = src.z * S
    VScale = dest

End Function


'----------------------------------------
'Name: RetIdentityMatrix
'----------------------------------------
Function RetIdentityMatrix() As D3DMATRIX ':( Missing Scope

  Dim RET As D3DMATRIX

    IdentityMatrix RET

End Function

'=================================
' IdentityMatrix
'=================================


'----------------------------------------
'Name: GetIDmatrix
'----------------------------------------
Function GetIDmatrix() As D3DMATRIX

    D3DXMatrixIdentity GetIDmatrix

End Function


'----------------------------------------
'Name: IdentityMatrix
'----------------------------------------
Sub IdentityMatrix(dest As D3DMATRIX) ':( Missing Scope

   D3DXMatrixIdentity dest

End Sub


'----------------------------------------
'Name: MIdentity
'----------------------------------------
Public Sub MIdentity(dMatrix As D3DMATRIX)

    dMatrix.m11 = 1
    dMatrix.m12 = 0
    dMatrix.m13 = 0
    dMatrix.m14 = 0
    dMatrix.m21 = 0
    dMatrix.m22 = 1
    dMatrix.m23 = 0
    dMatrix.m24 = 0
    dMatrix.m31 = 0
    dMatrix.m32 = 0
    dMatrix.m33 = 1
    dMatrix.m34 = 0
    dMatrix.m41 = 0
    dMatrix.m42 = 0
    dMatrix.m43 = 0
    dMatrix.m44 = 1

End Sub


'----------------------------------------
'Name: MRotate
'----------------------------------------
Sub MRotate(DestMat As D3DMATRIX, ByVal nValueX As Single, ByVal nValueY As Single, ByVal nValueZ As Single)

  Dim MatX As D3DMATRIX
  Dim MatY As D3DMATRIX
  Dim MatZ As D3DMATRIX
  Dim matTemp As D3DMATRIX

    D3DXMatrixIdentity matTemp

  Dim matWorld As D3DMATRIX ':( Move line to top of current Sub
    D3DXMatrixRotationX MatX, nValueX
    D3DXMatrixRotationY MatY, nValueY
    D3DXMatrixRotationZ MatZ, nValueZ

    D3DXMatrixMultiply matTemp, MatX, MatY
    D3DXMatrixMultiply matTemp, matTemp, MatZ

    DestMat = matTemp

End Sub


'----------------------------------------
'Name: MatrixToArray
'----------------------------------------
Sub MatrixToArray(ByRef dest, Mat As D3DMATRIX)

  'CopyMemory dest(0), mat, Len(mat.m11) * 16

    dest(0) = Mat.m11
    dest(1) = Mat.m12
    dest(2) = Mat.m13
    dest(3) = Mat.m14
    dest(4) = Mat.m21
    dest(5) = Mat.m22
    dest(6) = Mat.m23
    dest(7) = Mat.m24
    dest(8) = Mat.m31
    dest(9) = Mat.m32
    dest(10) = Mat.m33
    dest(11) = Mat.m34
    dest(12) = Mat.m41
    dest(13) = Mat.m42
    dest(14) = Mat.m43
    dest(15) = Mat.m44

End Sub


'----------------------------------------
'Name: MatToArray
'----------------------------------------
Sub MatToArray(ByRef dest() As Single, Mat As D3DMATRIX)

    ReDim dest(15)
    CopyMemory dest(0), Mat, Len(Mat)

End Sub


'----------------------------------------
'Name: Get_NemoMATRIX
'Object: Get
'Event: NemoMATRIX
'----------------------------------------
Function Get_NemoMATRIX(SrcMat As D3DMATRIX) As NemoMATRIX

  Dim MAZ As NemoMATRIX

    CopyMemory MAZ, SrcMat, Len(SrcMat)
  Dim i ':( As Variant ?':( Move line to top of current Function
    Get_NemoMATRIX = MAZ

End Function

' MTRANSLATE: Translates a matrix along the axis
'----------------------------------------
'Name: MTranslate
'----------------------------------------
Public Function MTranslate(dSource As D3DMATRIX, ByVal nValueX As Single, ByVal nValueY As Single, ByVal nValueZ As Single) As D3DMATRIX

  ' Reset to identity matrix

    MIdentity P_dCalculatorMatrix

    ' Add calculations
    With P_dCalculatorMatrix
        .m41 = nValueX
        .m42 = nValueY
        .m43 = nValueZ
    End With 'P_DCALCULATORMATRIX

    ' Apply transformations

    ' Return result
    MTranslate = MMultiply(P_dCalculatorMatrix, dSource)

End Function


'----------------------------------------
'Name: Matrix_Invert
'Object: Matrix
'Event: Invert
'----------------------------------------
Sub Matrix_Invert(DestMat As D3DMATRIX, Mat As D3DMATRIX)

  Dim FdetInv ':( As Variant ?

    If Abs((Mat.m44 - 1#) > 0.001) Then Exit Sub ':( Expand Structure or consider reversing Condition
    If (Abs(Mat.m14) > 0.001 And Abs(Mat.m24) > 0.001 And Abs(Mat.m34) > 0.001) Then Exit Sub ':( Expand Structure or consider reversing Condition

    FdetInv = 1# / (Mat.m11 * (Mat.m22 * Mat.m33 - Mat.m23 * Mat.m32) - _
              Mat.m12 * (Mat.m21 * Mat.m33 - Mat.m23 * Mat.m31) + _
              Mat.m13 * (Mat.m21 * Mat.m32 - Mat.m22 * Mat.m31))

    DestMat.m11 = FdetInv * (Mat.m22 * Mat.m33 - Mat.m23 * Mat.m32)
    DestMat.m12 = -FdetInv * (Mat.m12 * Mat.m33 - Mat.m13 * Mat.m32)
    DestMat.m13 = FdetInv * (Mat.m12 * Mat.m23 - Mat.m13 * Mat.m22)
    DestMat.m14 = 0#

    DestMat.m21 = -FdetInv * (Mat.m21 * Mat.m33 - Mat.m23 * Mat.m31)
    DestMat.m22 = FdetInv * (Mat.m11 * Mat.m33 - Mat.m13 * Mat.m31)
    DestMat.m23 = -FdetInv * (Mat.m11 * Mat.m23 - Mat.m13 * Mat.m21)
    DestMat.m24 = 0#

    DestMat.m31 = FdetInv * (Mat.m21 * Mat.m32 - Mat.m22 * Mat.m31)
    DestMat.m32 = -FdetInv * (Mat.m11 * Mat.m32 - Mat.m12 * Mat.m31)
    DestMat.m33 = FdetInv * (Mat.m11 * Mat.m22 - Mat.m12 * Mat.m21)
    DestMat.m34 = 0#

    DestMat.m41 = -(Mat.m41 * DestMat.m11 + Mat.m42 * DestMat.m21 + Mat.m43 * DestMat.m31)
    DestMat.m42 = -(Mat.m41 * DestMat.m12 + Mat.m42 * DestMat.m22 + Mat.m43 * DestMat.m32)
    DestMat.m43 = -(Mat.m41 * DestMat.m13 + Mat.m42 * DestMat.m23 + Mat.m43 * DestMat.m33)
    DestMat.m44 = 1#

End Sub


'----------------------------------------
'Name: MTrans_EX
'Object: MTrans
'Event: EX
'----------------------------------------
Public Function MTrans_EX(dSource As D3DMATRIX, VEC As D3DVECTOR) As D3DMATRIX

  ' Reset to identity matrix

    MIdentity P_dCalculatorMatrix

    ' Add calculations
    With P_dCalculatorMatrix
        .m41 = VEC.x
        .m42 = VEC.y
        .m43 = VEC.z
    End With 'P_DCALCULATORMATRIX

    ' Apply transformations

    ' Return result
    MTrans_EX = MMultiply(P_dCalculatorMatrix, dSource)

End Function


'----------------------------------------
'Name: MATRIX_Scale_Rotate_MOV
'Object: MATRIX
'Event: Scale_Rotate_MOV
'----------------------------------------
Function MATRIX_Scale_Rotate_MOV(Vscal As D3DVECTOR, Vrot As D3DVECTOR, Vmov As D3DVECTOR) As D3DMATRIX ':( Missing Scope

  Dim MatZ As D3DMATRIX
  Dim ROTz As D3DMATRIX
  Dim MOVz As D3DMATRIX
  Dim TEMPz As D3DMATRIX

    D3DXMatrixIdentity MatZ
    D3DXMatrixIdentity MOVz
    D3DXMatrixIdentity ROTz

    Call D3DXMatrixScaling(MatZ, Vscal.x, Vscal.y, Vscal.z)
    Call MRotate(ROTz, Vrot.x, Vrot.y, Vrot.z)
    Call D3DXMatrixTranslation(MOVz, Vmov.x, Vmov.y, Vmov.z)

    D3DXMatrixMultiply TEMPz, MatZ, ROTz
    D3DXMatrixMultiply MatZ, TEMPz, MOVz

    MATRIX_Scale_Rotate_MOV = MatZ

End Function


'----------------------------------------
'Name: MMultiply
'----------------------------------------
Public Function MMultiply(dM1 As D3DMATRIX, dM2 As D3DMATRIX) As D3DMATRIX

  ' Calculate multiply ...

    D3DXMatrixMultiply MMultiply, dM2, dM1

End Function

' MSCALE: Scales a matrix along the axis
'----------------------------------------
'Name: MScale
'----------------------------------------
Public Function MScale(dSource As D3DMATRIX, ByVal nValueX As Single, ByVal nValueY As Single, ByVal nValueZ As Single) As D3DMATRIX

  ' Reset to identity matrix

    MIdentity P_dCalculatorMatrix

    ' Add calculations
    With P_dCalculatorMatrix
        .m11 = nValueX
        .m22 = nValueY
        .m33 = nValueZ
    End With 'P_DCALCULATORMATRIX

    ' Apply transformations

    ' Return result
    ' dSource =

    MScale = MMultiply(P_dCalculatorMatrix, dSource)

End Function


'----------------------------------------
'Name: Class_Initialize
'Object: Class
'Event: Initialize
'----------------------------------------
Sub Class_Initialize()

    PIFactor = 3.14598745148 / 180

End Sub


'----------------------------------------
'Name: RetColor
'----------------------------------------
Function RetColor(R255, G255, B255, AlPHA255) As D3DCOLORVALUE

    With RetColor
        .R = R255 / 255
        .G = G255 / 255
        .B = B255 / 255
        .A = AlPHA255 / 255
    End With 'RETCOLOR

End Function


'----------------------------------------
'Name: Rotate
'----------------------------------------
Public Function Rotate(V As D3DVECTOR, xA, yA, zA) As D3DVECTOR

  Dim XPP1 As Double, YPP1 As Double, ZPP3 As Double, XPP2 As Double, ZPP2 As Double, YPP3 As Double
  Dim x                          As Double
  Dim y                          As Double
  Dim z                          As Double
  Dim cXA                        As Double
  Dim cYA                        As Double
  Dim cZA                        As Double
  Dim sXA                        As Double
  Dim sYA                        As Double
  Dim sZA                        As Double

    cXA = Cos(xA)
    cYA = Cos(yA)
    cZA = Cos(zA)
    sXA = Sin(xA)
    sYA = Sin(yA)
    sZA = Sin(zA)
    x = V.x
    y = V.y
    z = V.z
    XPP1 = x * cXA + y * sXA
    YPP1 = y * cXA - x * sXA
    XPP2 = XPP1 * cYA + z * sYA
    ZPP2 = z * cYA - XPP1 * sYA
    YPP3 = YPP1 * cZA - ZPP2 * sZA
    ZPP3 = ZPP2 * cZA + YPP1 * sZA
    Rotate.x = XPP2
    Rotate.y = YPP3
    Rotate.z = ZPP3

End Function


'----------------------------------------
'Name: RayTri
'----------------------------------------
Public Function RayTri(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, vDir As D3DVECTOR, vOrig As D3DVECTOR, T As Single, U As Single, V As Single, Optional CNormal As Boolean = False) As Boolean

  Dim edge1 As D3DVECTOR
  Dim edge2 As D3DVECTOR
  Dim pvec As D3DVECTOR
  Dim tvec As D3DVECTOR
  Dim qvec As D3DVECTOR
  Dim det As Single
  Dim fInvDet As Single
  Dim CollisionImpact As D3DVECTOR
  Dim CollisionNormal As D3DVECTOR

    'find vectors for the two edges sharing vert0
    D3DXVec3Subtract edge1, v1, v0
    D3DXVec3Subtract edge2, v2, v0

    'begin calculating the determinant - also used to caclulate u parameter
    D3DXVec3Cross pvec, vDir, edge2

    'if determinant is nearly zero, ray lies in plane of triangle
    det = D3DXVec3Dot(edge1, pvec)
    If Abs(det < 0.0001) Then
        Exit Function '>---> Bottom
    End If

    'calculate distance from vert0 to ray origin
    D3DXVec3Subtract tvec, vOrig, v0

    'calculate u parameter and test bounds
    U = D3DXVec3Dot(tvec, pvec)
    If (U < 0 Or U > det) Then
        Exit Function '>---> Bottom
    End If

    'prepare to test v parameter
    D3DXVec3Cross qvec, tvec, edge1

    'calculate v parameter and test bounds
    V = D3DXVec3Dot(vDir, qvec)
    If (V < 0 Or (U + V > det)) Then
        Exit Function '>---> Bottom
    End If

    'calculate t, scale parameters, ray intersects triangle
    T = D3DXVec3Dot(edge2, qvec)
    fInvDet = 1 / det
    T = T * fInvDet
    U = U * fInvDet
    V = V * fInvDet
    If T = 0 Then Exit Function ':( Expand Structure or consider reversing Condition
    If CNormal = True Then ':( Remove Pleonasm
        D3DXVec3Cross CollisionNormal, VNormalize2(edge1), VNormalize2(edge2)
        CollisionImpact = VAdd(vOrig, VScale(vDir, T))
    End If
    'CollisionNormal = D3D
    RayTri = 1

End Function


'----------------------------------------
'Name: VERTEXtoVEC
'----------------------------------------
Function VERTEXtoVEC(VERT As D3DVERTEX) As D3DVECTOR

    VERTEXtoVEC.x = VERT.x
    VERTEXtoVEC.y = VERT.y
    VERTEXtoVEC.z = VERT.z

End Function


'----------------------------------------
'Name: LVERTEXtoVEC
'----------------------------------------
Function LVERTEXtoVEC(VERT As D3DLVERTEX) As D3DVECTOR

    LVERTEXtoVEC.x = VERT.x
    LVERTEXtoVEC.y = VERT.y
    LVERTEXtoVEC.z = VERT.z

End Function


'----------------------------------------
'Name: GET_fileName
'Object: GET
'Event: fileName
'----------------------------------------

Function KillSpace(ST As String) As String
Dim i, K, SS As String
Dim SZ As String
SZ = "/\ABCDEFGHIJKLMNOPQRSTUVWXYZ.123456789_-"
    
    For i = 1 To Len(ST)
      If Mid(ST, i, 1) <> " " Then SS = SS + Mid(ST, i, 1)
      
    Next


  KillSpace = SS
End Function

Function GET_fileName(LongNAME As String) As String

  Dim ST As String
  Dim ZZ, i ':( As Variant ?':( As Variant ?
  Dim EXT ':( As Variant ?':( As Variant ?
  Dim ST1 As String

    For i = Len(LongNAME) To 1 Step -1

        If Mid$(LongNAME, i, 1) = "." Then Exit For ':( Expand Structure or consider reversing Condition':( Expand Structure or consider reversing Condition
    Next i

    EXT = Len(LongNAME) - i

    ZZ = Len(LongNAME)

    For i = ZZ To 1 Step -1
        If Mid$(LongNAME, i, 1) = "\" Or Mid$(LongNAME, i, 1) = "/" Then Exit For ':( Expand Structure or consider reversing Condition':( Expand Structure or consider reversing Condition
    Next ':( Repeat For-Variable: I':( Repeat For-Variable: I

    ST1 = Right$(LongNAME, ZZ - i)

    On Error GoTo ooF
    If InStr(LongNAME, ".") > 0 Then GET_fileName = Left$(ST1, Len(ST1) - EXT - 1) ':( Expand Structure':( Expand Structure
    If InStr(LongNAME, ".") < 1 Then GET_fileName = Left$(ST1, Len(ST1)) ':( Expand Structure':( Expand Structure

Exit Function

ooF:
    GET_fileName = ST1

End Function



Function GET_fileNameEX(LongNAME As String) As String

  Dim ST As String
  Dim ZZ, i ':( As Variant ?':( As Variant ?
  Dim EXT ':( As Variant ?':( As Variant ?
  Dim ST1 As String

    For i = Len(LongNAME) To 1 Step -1

        If Mid$(LongNAME, i, 1) = "." Then Exit For ':( Expand Structure or consider reversing Condition':( Expand Structure or consider reversing Condition
    Next i

    EXT = Len(LongNAME) - i

    ZZ = Len(LongNAME)

    For i = ZZ To 1 Step -1
        If Mid$(LongNAME, i, 1) = "\" Or Mid$(LongNAME, i, 1) = "/" Then Exit For ':( Expand Structure or consider reversing Condition':( Expand Structure or consider reversing Condition
    Next ':( Repeat For-Variable: I':( Repeat For-Variable: I

    ST1 = Right$(LongNAME, ZZ - i)

    On Error GoTo ooF
    If InStr(LongNAME, ".") > 0 Then GET_fileNameEX = Left$(ST1, Len(ST1))   ':( Expand Structure':( Expand Structure
    If InStr(LongNAME, ".") < 1 Then GET_fileNameEX = Left$(ST1, Len(ST1)) ':( Expand Structure':( Expand Structure

Exit Function

ooF:
    GET_fileNameEX = ST1

End Function


'----------------------------------------
'Name: GetLongFolDerName
'----------------------------------------
Function GetLongFolDerName(STlongFilename As String) ':( As Variant ?

  Dim ST As String
  Dim Temp As String
  Dim i As Integer

    ST = GET_fileName(STlongFilename)
    i = InStr(STlongFilename, ST)
    Temp = Left(STlongFilename, i - 2)
    GetLongFolDerName = Temp

End Function


'----------------------------------------
'Name: GET_pathName
'Object: GET
'Event: pathName
'----------------------------------------
Function GET_pathName(LongNAME As String) As String

  Dim ST As String
  Dim ZZ, i ':( As Variant ?':( As Variant ?
  Dim ST1 As String

    ZZ = Len(LongNAME)

    For i = ZZ To 1 Step -1
        If Mid$(LongNAME, i, 1) = "\" Or Mid$(LongNAME, i, 1) = "/" Then Exit For ':( Expand Structure or consider reversing Condition':( Expand Structure or consider reversing Condition
    Next ':( Repeat For-Variable: I':( Repeat For-Variable: I

    ST1 = Left$(LongNAME, i - 1)
    GET_pathName = ST1

End Function


'----------------------------------------
'Name: GET_LastpathName
'Object: GET
'Event: LastpathName
'----------------------------------------
Function GET_LastpathName(LongNAME As String) As String

  Dim ST As String
  Dim ST1 As String
  Dim ZZ, i, j, Pos1, Pos2 ':( As Variant ?

    ZZ = Len(LongNAME)

    For i = ZZ To 1 Step -1
        j = j + 1
        If Mid(LongNAME, i, 1) = "\" Or Mid(LongNAME, i, 1) = "/" Then
            If Pos1 = 0 Then
                Pos1 = i
                GoTo NEXTA
            End If
            If Pos1 > 0 And Pos2 = 0 Then Pos2 = i ':( Expand Structure
        End If
NEXTA:

        If i = 1 Or Pos2 > 0 Then
            If Pos2 = 0 Then Pos2 = 1 ':( Expand Structure
        End If

    Next ':( Repeat For-Variable: I

    ST1 = Mid(LongNAME, Pos2 + 1, Pos1 - Pos2 - 1)
    GET_LastpathName = ST1

End Function


'----------------------------------------
'Name: Get_BoundingBOX
'Object: Get
'Event: BoundingBOX
'----------------------------------------
Sub Get_BoundingBOX(lpVERT() As D3DVERTEX, ByRef RETBOX As NEMO_BBOX)

  Dim i, j, ZZ ':( As Variant ?
  Dim VEC As D3DVECTOR

    ' bounding box
  Dim min As D3DVECTOR ':( Move line to top of current Sub
  Dim max As D3DVECTOR ':( Move line to top of current Sub
    min = Vector(99999999, 999999999, 999999999)
    max = Vector(-999999999, -999999999, -999999999)

    i = LBound(lpVERT)
    j = UBound(lpVERT)

    For ZZ = i To j
        With lpVERT(ZZ)
            If .x > max.x Then max.x = .x ':( Expand Structure
            If .y > max.y Then max.y = .y ':( Expand Structure
            If .z > max.z Then max.z = .z ':( Expand Structure
            If .x < min.x Then min.x = .x ':( Expand Structure
            If .y < min.y Then min.y = .y ':( Expand Structure
            If .z < min.z Then min.z = .z ':( Expand Structure

        End With 'LPVERT(ZZ)
    Next ZZ
    RETBOX.max = max
    RETBOX.min = min

End Sub


'----------------------------------------
'Name: Get_BoundingBoxEX
'Object: Get
'Event: BoundingBoxEX
'----------------------------------------
Sub Get_BoundingBoxEX(lpVERT() As D3DVERTEX, ByRef RETBOX As NemoBoundBOX)

  Dim i, j, ZZ ':( As Variant ?
  Dim VEC As D3DVECTOR

    ' bounding box
  Dim min As NemoVECTOR ':( Move line to top of current Sub
  Dim max As NemoVECTOR ':( Move line to top of current Sub
    min = VectorN(99999999, 999999999, 999999999)
    max = VectorN(-999999999, -999999999, -999999999)

    i = LBound(lpVERT)
    j = UBound(lpVERT)

    For ZZ = i To j
        With lpVERT(ZZ)
            If .x > max.x Then max.x = .x ':( Expand Structure
            If .y > max.y Then max.y = .y ':( Expand Structure
            If .z > max.z Then max.z = .z ':( Expand Structure
            If .x < min.x Then min.x = .x ':( Expand Structure
            If .y < min.y Then min.y = .y ':( Expand Structure
            If .z < min.z Then min.z = .z ':( Expand Structure

        End With 'LPVERT(ZZ)
    Next ZZ
    RETBOX.Vmax = max
    RETBOX.Vmin = min

End Sub


'----------------------------------------
'Name: Get_BoundingSphere
'Object: Get
'Event: BoundingSphere
'----------------------------------------
Sub Get_BoundingSphere(lpVERT() As D3DVERTEX, ByRef RetSphereCenter As D3DVECTOR, RetRadius)

  Dim i, j, ZZ ':( As Variant ?
  Dim VEC As D3DVECTOR

    ' bounding box
  Dim Vmin As D3DVECTOR ':( Move line to top of current Sub
  Dim Vmax As D3DVECTOR ':( Move line to top of current Sub
    Vmin = Vector(99999999, 999999999, 999999999)
    Vmax = Vector(-999999999, -999999999, -999999999)

    i = LBound(lpVERT)
    j = UBound(lpVERT)

    For ZZ = i To j
        With lpVERT(ZZ)
            If .x > Vmax.x Then Vmax.x = .x ':( Expand Structure
            If .y > Vmax.y Then Vmax.y = .y ':( Expand Structure
            If .z > Vmax.z Then Vmax.z = .z ':( Expand Structure
            If .x < Vmin.x Then Vmin.x = .x ':( Expand Structure
            If .y < Vmin.y Then Vmin.y = .y ':( Expand Structure
            If .z < Vmin.z Then Vmin.z = .z ':( Expand Structure

        End With 'LPVERT(ZZ)
    Next ZZ

    RetRadius = MAX_3(Vmax.x - Vmin.x, Vmax.y - Vmin.y, Vmax.z - Vmin.z) / 2
    RetSphereCenter = Vmiddle(Vmax, Vmin)

End Sub

'colision Detection
'Function BOX_GetBoxCollision(Box1 As D3DRMBOX, Box2 As D3DRMBOX) As Boolean
'Dim vec1 As D3DVECTOR
'Dim vec2 As D3DVECTOR
'Dim vec3 As D3DVECTOR
'Dim vec4 As D3DVECTOR
'Dim vec5 As D3DVECTOR
'Dim Vec6 As D3DVECTOR
'
'With Box2
'vec1 = VICTOR(.min.x, .min.y, .max.z)
'vec2 = VICTOR(.max.x, .min.y, .max.z)
'vec3 = VICTOR(.min.x, .max.y, .min.z)
'vec4 = VICTOR(.min.x, .max.y, .max.z)
'
'vec5 = VICTOR(.max.x, .min.y, .min.z)
'Vec6 = VICTOR(.max.x, .max.y, .min.z)
'
'End With
'
'If BOX_GetPointBOXCollision(Box1, Box2.max) Then
'  BOX_GetBoxCollision = True
'  Exit Function
'ElseIf BOX_GetPointBOXCollision(Box1, Box2.min) Then
'  BOX_GetBoxCollision = True
'  Exit Function
'ElseIf BOX_GetPointBOXCollision(Box1, vec1) Then
'  BOX_GetBoxCollision = True
'  Exit Function
'ElseIf BOX_GetPointBOXCollision(Box1, vec2) Then
'  BOX_GetBoxCollision = True
'  Exit Function
'ElseIf BOX_GetPointBOXCollision(Box1, vec3) Then
'  BOX_GetBoxCollision = True
'  Exit Function
'
'ElseIf BOX_GetPointBOXCollision(Box1, vec4) Then
'  BOX_GetBoxCollision = True
'  Exit Function
'
'ElseIf BOX_GetPointBOXCollision(Box1, vec5) Then
'  BOX_GetBoxCollision = True
'  Exit Function
'
'ElseIf BOX_GetPointBOXCollision(Box1, Vec6) Then
'  BOX_GetBoxCollision = True
'  Exit Function
'
'
'End If
'
'
'End Function


'----------------------------------------
'Name: CreateBox
'----------------------------------------
Sub CreateBox(Vmin As D3DVECTOR, Vmax As D3DVECTOR, RETBOX As D3DRMBOX)

    RETBOX.min = Vmin
    RETBOX.max = Vmax

End Sub


'----------------------------------------
'Name: BOX_GetPointBOXCollision
'Object: BOX
'Event: GetPointBOXCollision
'----------------------------------------
Function BOX_GetPointBOXCollision(OBJboxMax As D3DVECTOR, OBJboxMin As D3DVECTOR, Tpoint As D3DVECTOR) As Boolean

  '===============Teste de colision

    If Tpoint.x > OBJboxMin.x And Tpoint.x < OBJboxMax.x And _
       Tpoint.y > OBJboxMin.y And Tpoint.y < OBJboxMax.y And _
       Tpoint.z > OBJboxMin.z And Tpoint.z < OBJboxMax.z Then _
       BOX_GetPointBOXCollision = True ':( Expand Structure

End Function


'----------------------------------------
'Name: VNormalize
'----------------------------------------
Function VNormalize(dest As D3DVECTOR) As D3DVECTOR

  '  On Local Error Resume Next
  '  Dim l As Double
  '  l = dest.x * dest.x + dest.y * dest.y + dest.z * dest.z
  '  l = Sqr(l)
  '  If l = 0 Then
  '    dest.x = 0
  '    dest.y = 0
  '    dest.z = 0
  '    Exit Sub
  '  End If
  '  dest.x = dest.x / l
  '  dest.y = dest.y / l
  '  dest.z = dest.z / l

    D3DXVec3Normalize dest, dest
    VNormalize = dest

End Function


'----------------------------------------
'Name: VCompare
'----------------------------------------
Public Function VCompare(v1 As D3DVECTOR, v2 As D3DVECTOR) As Integer

    If v1.x = v2.x And v1.y = v2.y And v1.z = v2.z Then VCompare = 1 ':( Expand Structure

End Function


'----------------------------------------
'Name: VNormalize2
'----------------------------------------
Function VNormalize2(dest As D3DVECTOR) As D3DVECTOR

    On Local Error Resume Next
      Dim L As Double ':( Move line to top of current Function
        L = dest.x * dest.x + dest.y * dest.y + dest.z * dest.z
        L = Sqr(L)
        If L = 0 Then
            dest.x = 0
            dest.y = 0
            dest.z = 0
            Exit Function '>---> Bottom
        End If
        VNormalize2.x = dest.x / L
        VNormalize2.y = dest.y / L
        VNormalize2.z = dest.z / L

End Function ':( On Error Resume still active


'----------------------------------------
'Name: VectorNormalize
'----------------------------------------
Sub VectorNormalize(dest As D3DVECTOR)

    On Local Error Resume Next
      Dim L As Double ':( Move line to top of current Sub
        L = dest.x * dest.x + dest.y * dest.y + dest.z * dest.z
        L = Sqr(L)
        If L = 0 Then
            dest.x = 0
            dest.y = 0
            dest.z = 0
            Exit Sub '>---> Bottom
        End If
        dest.x = dest.x / L
        dest.y = dest.y / L
        dest.z = dest.z / L

End Sub ':( On Error Resume still active






'----------------------------------------
'Name: Vector_Rotate
'Object: Vector
'Event: Rotate
'----------------------------------------
Function Vector_Rotate(VEC As D3DVECTOR, angleX, angleY, angleZ) As D3DVECTOR

  Dim TempVec As D3DVECTOR

    TempVec = VEC
  Dim qx  As Single, qy    As Single, qz    As Single, qw    As Single, tmpx    As Single, tmpy    As Single, tmpz As Single ':( Move line to top of current Function
    Call Euler2quat(angleX, angleY, angleZ, qw, qx, qy, qz)
    tmpx = TempVec.x * (1 - 2 * qy * qy - 2 * qz * qz) + TempVec.y * (2 * qx * qy - 2 * qw * qz) + TempVec.z * (2 * qx * qz + 2 * qw * qy)
    tmpy = TempVec.x * (2 * qx * qy + 2 * qw * qz) + TempVec.y * (1 - 2 * qx * qx - 2 * qz * qz) + TempVec.z * (2 * qy * qz - 2 * qw * qx)
    tmpz = TempVec.x * (2 * qx * qz - 2 * qw * qy) + TempVec.y * (2 * qy * qz + 2 * qw * qx) + TempVec.z * (1 - 2 * qx * qx - 2 * qy * qy)
    Vector_Rotate.x = tmpx
    Vector_Rotate.y = tmpy
    Vector_Rotate.z = tmpz

End Function


'----------------------------------------
'Name: Vector_RotateEX
'Object: Vector
'Event: RotateEX
'----------------------------------------
Function Vector_RotateEX(VEC As D3DVECTOR, rotQuat As D3DQUATERNION) As D3DVECTOR

  Dim tmpx, tmpy, tmpz ':( As Variant ?

    tmpx = VEC.x * (1 - 2 * rotQuat.y * rotQuat.y - 2 * rotQuat.z * rotQuat.z) + VEC.y * (2 * rotQuat.x * rotQuat.y - 2 * rotQuat.z * rotQuat.z) + VEC.z * (2 * rotQuat.x * rotQuat.z + 2 * rotQuat.z * rotQuat.y)
    tmpy = VEC.x * (2 * rotQuat.x * rotQuat.y + 2 * rotQuat.z * rotQuat.z) + VEC.y * (1 - 2 * rotQuat.x * rotQuat.x - 2 * rotQuat.z * rotQuat.z) + VEC.z * (2 * rotQuat.y * rotQuat.z - 2 * rotQuat.z * rotQuat.x)
    tmpz = VEC.x * (2 * rotQuat.x * rotQuat.z - 2 * rotQuat.z * rotQuat.y) + VEC.y * (2 * rotQuat.y * rotQuat.z + 2 * rotQuat.z * rotQuat.x) + VEC.z * (1 - 2 * rotQuat.x * rotQuat.x - 2 * rotQuat.y * rotQuat.y)
    Vector_RotateEX.x = tmpx
    Vector_RotateEX.y = tmpy
    Vector_RotateEX.z = tmpz

End Function


'----------------------------------------
'Name: Acos
'----------------------------------------
Function Acos(x) ':( As Variant ?

    On Error Resume Next
        Acos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)

End Function ':( On Error Resume still active


'----------------------------------------
'Name: Quaternion_ToMatrix
'Object: Quaternion
'Event: ToMatrix
'----------------------------------------
Sub Quaternion_ToMatrix(Quat As D3DQUATERNION, RetMat As D3DMATRIX)

    RetMat = RetIdentityMatrix
  Dim wx, wy, wz, XX, yY, yz, xy, xz, ZZ, x2, Y2, z2 ':( As Variant ?':( Move line to top of current Sub

    x2 = Quat.x + Quat.x
    Y2 = Quat.y + Quat.y
    z2 = Quat.z + Quat.z

    XX = Quat.x * x2
    xy = Quat.x * Y2
    xz = Quat.x * z2
    yY = Quat.y * Y2
    yz = Quat.y * z2
    ZZ = Quat.z * z2
    wx = Quat.z * x2
    wy = Quat.z * Y2
    wz = Quat.z * z2

    RetMat.m11 = 1# - (yY + ZZ)
    RetMat.m12 = xy - wz
    RetMat.m13 = xz + wy
    RetMat.m14 = 0#
    RetMat.m21 = xy + wz
    RetMat.m22 = 1# - (XX + ZZ)
    RetMat.m23 = yz - wx
    RetMat.m24 = 0#
    RetMat.m31 = xz - wy
    RetMat.m32 = yz + wx
    RetMat.m33 = 1# - (XX + yY)
    RetMat.m34 = 0#
    RetMat.m41 = 0
    RetMat.m42 = 0
    RetMat.m43 = 0
    RetMat.m44 = 1

End Sub


'----------------------------------------
'Name: Quaternion_FromEulerAngle
'Object: Quaternion
'Event: FromEulerAngle
'----------------------------------------
Sub Quaternion_FromEulerAngle(ByVal Ax As Single, ByVal Ay As Single, ByVal Az As Single, ByRef w As Single, ByRef x As Single, ByRef y As Single, ByRef z)

  Dim ex  As Double, ey    As Double, ez As Double
  Dim cr  As Double, cp    As Double, cy    As Double, sr    As Double, sp    As Double, Sy    As Double, cpcy    As Double, spsy As Double
  Dim norm As Double

    ex = DegToRad(Ax) / 2#
    ey = DegToRad(Ay) / 2#
    ez = DegToRad(Az) / 2#

    cr = Cos(ex)
    cp = Cos(ey)
    cy = Cos(ez)

    sr = Sin(ex)
    sp = Sin(ey)
    Sy = Sin(ez)

    cpcy = cp * cy
    spsy = sp * Sy

    w = (cr * cpcy + sr * spsy)
    x = (sr * cpcy - cr * spsy)
    y = (cr * sp * cy + sr * cp * Sy)
    z = (cr * cp * Sy - sr * sp * cy)
    norm = Sqr((x * x) + Sqr(y * y) + Sqr(z * z) + Sqr(w * w))
    If (norm = 0#) Then Exit Sub ':( Expand Structure or consider reversing Condition
    w = w / norm
    x = x / norm
    y = y / norm
    z = z / norm

End Sub


'----------------------------------------
'Name: Quaternion_FromEulerAngleEX
'Object: Quaternion
'Event: FromEulerAngleEX
'----------------------------------------
Sub Quaternion_FromEulerAngleEX(ByVal Ax As Single, ByVal Ay As Single, ByVal Az As Single, RetQuat As D3DQUATERNION)

  Dim ex  As Double, ey    As Double, ez As Double
  Dim cr  As Double, cp    As Double, cy    As Double, sr    As Double, sp    As Double, Sy    As Double, cpcy    As Double, spsy As Double
  Dim norm As Double

    ex = DegToRad(Ax) / 2#
    ey = DegToRad(Ay) / 2#
    ez = DegToRad(Az) / 2#

    cr = Cos(ex)
    cp = Cos(ey)
    cy = Cos(ez)

    sr = Sin(ex)
    sp = Sin(ey)
    Sy = Sin(ez)

    cpcy = cp * cy
    spsy = sp * Sy

    RetQuat.w = (cr * cpcy + sr * spsy)
    RetQuat.x = (sr * cpcy - cr * spsy)
    RetQuat.y = (cr * sp * cy + sr * cp * Sy)
    RetQuat.z = (cr * cp * Sy - sr * sp * cy)
    norm = Sqr((RetQuat.x * RetQuat.x) + Sqr(RetQuat.y * RetQuat.y) + Sqr(RetQuat.z * RetQuat.z) + Sqr(RetQuat.z * RetQuat.z))
    If (norm = 0#) Then Exit Sub ':( Expand Structure or consider reversing Condition
    RetQuat.w = RetQuat.w / norm
    RetQuat.x = RetQuat.x / norm
    RetQuat.y = RetQuat.y / norm
    RetQuat.z = RetQuat.z / norm

End Sub


'----------------------------------------
'Name: Quaternion_GetAxisAngle
'Object: Quaternion
'Event: GetAxisAngle
'----------------------------------------
Function Quaternion_GetAxisAngle(Quat As D3DQUATERNION, ByRef x, ByRef y, ByRef z, ByRef Angle) ':( As Variant ?

  Dim temp_angle As Double, Scal As Double
  Dim Length As Single

    temp_angle = Acos(Quat.z)
    Scal = Sqr(Quat.x * Quat.x + Quat.y * Quat.y + Quat.z * Quat.z)

    If (Scal = 0) Then
        Angle = 0#
        x = 0
        y = 0
        z = 0#
      Else 'NOT (SCAL...
        Angle = (temp_angle * 2#)      '// angle in radians
        x = (Quat.x / Scal)
        y = (Quat.y / Scal)
        z = (Quat.z / Scal)
        Length = Sqr(x * x + y * y + z * z)
        If (Length <> 0) Then
            x = x / Length
            y = y / Length
            z = z / Length
        End If
        Angle = RadToDeg(Angle)
    End If

End Function


'----------------------------------------
'Name: Quaternion_MultWith
'Object: Quaternion
'Event: MultWith
'----------------------------------------
Sub Quaternion_MultWith(DestQuat As D3DQUATERNION, q1 As D3DQUATERNION, Q2 As D3DQUATERNION)

  Dim Rx As Double, Ry As Double, Rz As Double, rw As Double

    rw = q1.z * Q2.z - q1.x * Q2.x - q1.y * Q2.y - q1.z * Q2.z

    Rx = q1.z * Q2.x + q1.x * Q2.z + q1.y * Q2.z - q1.z * Q2.y
    Ry = q1.z * Q2.y + q1.y * Q2.z + q1.z * Q2.x - q1.x * Q2.z
    Rz = q1.z * Q2.z + q1.z * Q2.z + q1.x * Q2.y - q1.y * Q2.x

    DestQuat.x = Rx
    DestQuat.y = Ry
    DestQuat.z = Rz
    DestQuat.z = rw

End Sub


'----------------------------------------
'Name: Quaternion_Normalize
'Object: Quaternion
'Event: Normalize
'----------------------------------------
Sub Quaternion_Normalize(DestQuat As D3DQUATERNION)

  Dim norm As Double

    norm = Sqr(DestQuat.x * DestQuat.x + DestQuat.y * DestQuat.y + DestQuat.z * DestQuat.z + DestQuat.z * DestQuat.z)

    If (norm = 0) Then Exit Sub ':( Expand Structure or consider reversing Condition

    DestQuat.x = (DestQuat.x / norm)
    DestQuat.y = (DestQuat.y / norm)
    DestQuat.z = (DestQuat.z / norm)
    DestQuat.z = (DestQuat.z / norm)

End Sub


'----------------------------------------
'Name: Vector_RotateAroundVector
'Object: Vector
'Event: RotateAroundVector
'----------------------------------------
Function Vector_RotateAroundVector(VecToRot As D3DVECTOR, angleX, angleY, angleZ, VcenterRot As D3DVECTOR) As D3DVECTOR

  Dim TempVec As D3DVECTOR

    TempVec = VecToRot

    TempVec.x = TempVec.x - VcenterRot.x
    TempVec.y = TempVec.y - VcenterRot.y
    TempVec.z = TempVec.z - VcenterRot.z

    TempVec = Vector_Rotate(TempVec, angleX, angleY, angleZ)

    TempVec.x = TempVec.x + VcenterRot.x
    TempVec.y = TempVec.y + VcenterRot.y
    TempVec.z = TempVec.z + VcenterRot.z

    Vector_RotateAroundVector = TempVec

End Function


'----------------------------------------
'Name: DegToRad
'----------------------------------------
Function DegToRad(XX) ':( As Variant ?

    DegToRad = XX * RAD

End Function


'----------------------------------------
'Name: RadToDeg
'----------------------------------------
Function RadToDeg(XX) ':( As Variant ?

    RadToDeg = ((XX) * 180#) / PI

End Function


'----------------------------------------
'Name: Euler2quat
'----------------------------------------
Sub Euler2quat(ByVal Ax As Single, ByVal Ay As Single, ByVal Az As Single, ByRef w As Single, ByRef x As Single, ByRef y As Single, ByRef z)

  Dim ex  As Double, ey    As Double, ez As Double
  Dim cr  As Double, cp    As Double, cy    As Double, sr    As Double, sp    As Double, Sy    As Double, cpcy    As Double, spsy As Double
  Dim norm As Double

    ex = DegToRad(Ax) / 2#
    ey = DegToRad(Ay) / 2#
    ez = DegToRad(Az) / 2#

    cr = Cos(ex)
    cp = Cos(ey)
    cy = Cos(ez)

    sr = Sin(ex)
    sp = Sin(ey)
    Sy = Sin(ez)

    cpcy = cp * cy
    spsy = sp * Sy

    w = (cr * cpcy + sr * spsy)
    x = (sr * cpcy - cr * spsy)
    y = (cr * sp * cy + sr * cp * Sy)
    z = (cr * cp * Sy - sr * sp * cy)
    norm = Sqr((x * x) + Sqr(y * y) + Sqr(z * z) + Sqr(w * w))
    If (norm = 0#) Then Exit Sub ':( Expand Structure or consider reversing Condition
    w = w / norm
    x = x / norm
    y = y / norm
    z = z / norm

End Sub


'----------------------------------------
'Name: Euler2quatEX
'----------------------------------------
Sub Euler2quatEX(ByVal Ax As Single, ByVal Ay As Single, ByVal Az As Single, RetQuat As D3DQUATERNION)

  Dim ex  As Double, ey    As Double, ez As Double
  Dim cr  As Double, cp    As Double, cy    As Double, sr    As Double, sp    As Double, Sy    As Double, cpcy    As Double, spsy As Double
  Dim norm As Double

    ex = DegToRad(Ax) / 2#
    ey = DegToRad(Ay) / 2#
    ez = DegToRad(Az) / 2#

    cr = Cos(ex)
    cp = Cos(ey)
    cy = Cos(ez)

    sr = Sin(ex)
    sp = Sin(ey)
    Sy = Sin(ez)

    cpcy = cp * cy
    spsy = sp * Sy

    RetQuat.z = (cr * cpcy + sr * spsy)
    RetQuat.x = (sr * cpcy - cr * spsy)
    RetQuat.y = (cr * sp * cy + sr * cp * Sy)
    RetQuat.z = (cr * cp * Sy - sr * sp * cy)
    norm = Sqr((RetQuat.x * RetQuat.x) + Sqr(RetQuat.y * RetQuat.y) + Sqr(RetQuat.z * RetQuat.z) + Sqr(RetQuat.z * RetQuat.z))
    If (norm = 0#) Then Exit Sub ':( Expand Structure or consider reversing Condition
    RetQuat.z = RetQuat.z / norm
    RetQuat.x = RetQuat.x / norm
    RetQuat.y = RetQuat.y / norm
    RetQuat.z = RetQuat.z / norm

End Sub

'=================================
' VectorNegate
'=================================
'----------------------------------------
'Name: VectorNegate
'----------------------------------------
Sub VectorNegate(V As D3DVECTOR)

    V.x = -V.x
    V.y = -V.y
    V.z = -V.z

End Sub


'----------------------------------------
'Name: Vabs
'----------------------------------------
Function Vabs(V As D3DVECTOR) As D3DVECTOR

    Vabs.x = Abs(V.x)
    Vabs.y = Abs(V.y)
    Vabs.z = Abs(V.z)

End Function


'----------------------------------------
'Name: VNegate
'----------------------------------------
Function VNegate(V As D3DVECTOR) As D3DVECTOR

  Dim RET As D3DVECTOR

    RET.x = -V.x
    RET.y = -V.y
    RET.z = -V.z
    VNegate = RET

End Function

'=================================
' Vector subtract
'=================================
'----------------------------------------
'Name: VectorSubtract
'----------------------------------------
Sub VectorSubtract(dest As D3DVECTOR, A As D3DVECTOR, B As D3DVECTOR)

    dest.x = A.x - B.x
    dest.y = A.y - B.y
    dest.z = A.z - B.z

End Sub

'=================================
' VectorAdd
'=================================
'----------------------------------------
'Name: VectorAdd
'----------------------------------------
Sub VectorAdd(dest As D3DVECTOR, A As D3DVECTOR, B As D3DVECTOR)

    dest.x = A.x + B.x
    dest.y = A.y + B.y
    dest.z = A.z + B.z

End Sub

'=================================
' VectorCrossProduct
'=================================
' can be used to compute normals.
'
'----------------------------------------
'Name: VectorCrossProduct
'----------------------------------------
Sub VectorCrossProduct(dest As D3DVECTOR, A As D3DVECTOR, B As D3DVECTOR)

    dest.x = A.y * B.z - A.z * B.y
    dest.y = A.z * B.x - A.x * B.z
    dest.z = A.x * B.y - A.y * B.x

End Sub


'----------------------------------------
'Name: VCross
'----------------------------------------
Function VCross(A As D3DVECTOR, B As D3DVECTOR) As D3DVECTOR

  Dim dest As D3DVECTOR

    dest.x = A.y * B.z - A.z * B.y
    dest.y = A.z * B.x - A.x * B.z
    dest.z = A.x * B.y - A.y * B.x
    VCross = dest

End Function

'=================================
' VectorNormalize
'=================================
' creates a vector of length 1 in the same direction
'

'=================================
' VectorDotProduct
'=================================
'----------------------------------------
'Name: VectorDotProduct
'----------------------------------------
Function VectorDotProduct(A As D3DVECTOR, B As D3DVECTOR) As Single

    VectorDotProduct = A.x * B.x + A.y * B.y + A.z * B.z

End Function


'----------------------------------------
'Name: VDot
'----------------------------------------
Function VDot(A As D3DVECTOR, B As D3DVECTOR) As Single

    VDot = A.x * B.x + A.y * B.y + A.z * B.z

End Function


'----------------------------------------
'Name: VCopy
'----------------------------------------
Function VCopy(src As D3DVECTOR) As D3DVECTOR

  Dim dest As D3DVECTOR

    dest.x = src.x
    dest.y = src.y
    dest.z = src.z
    VCopy = dest

End Function

'=================================
' VectorScale
'=================================
' scale a vector by a scalar
'----------------------------------------
'Name: VectorScale
'----------------------------------------
Sub VectorScale(dest As D3DVECTOR, src As D3DVECTOR, S As Single)

    dest.x = src.x * S
    dest.y = src.y * S
    dest.z = src.z * S

End Sub

'=================================
' MakeVector
'=================================
'----------------------------------------
'Name: MakeVector
'----------------------------------------
Sub MakeVector(V As D3DVECTOR, x As Single, y As Single, z As Single)

    V.x = x
    V.y = z
    V.z = y

End Sub


'----------------------------------------
'Name: RVector
'----------------------------------------
Function RVector(x As Single, y As Single, z As Single) As D3DVECTOR

  Dim V As D3DVECTOR

    V.x = x
    V.y = y
    V.z = z
    RVector = V

End Function

'=================================
' MakeVertex
'=================================
'----------------------------------------
'Name: MakeVertex
'----------------------------------------
Sub MakeVertex(RET As D3DVERTEX, Vect As D3DVECTOR, vNorm As D3DVECTOR, tu As Single, tv As Single)

    With RET
        .Nx = vNorm.x
        .Ny = vNorm.y
        .Nz = vNorm.z
        .tu = tu
        .tv = tv
        .x = Vect.x
        .y = Vect.y
        .z = Vect.z
    End With 'RET

End Sub

'=================================
' MakeLVertex
'=================================
'----------------------------------------
'Name: MakeLVertex
'----------------------------------------
Sub MakeLVertex(RET As D3DLVERTEX, x As Single, y As Single, z As Single, color As Long, Specular As Single, tu As Single, tv As Single)

    With RET
        .Specular = Specular
        .tu = tu
        .tv = tv
        .x = x
        .y = y
        .z = z
        .color = color
    End With 'RET

End Sub


'----------------------------------------
'Name: RLVertex
'----------------------------------------
Function RLVertex(x As Single, y As Single, z As Single, color As Long, Specular As Single, tu As Single, tv As Single) As D3DLVERTEX

  Dim RET As D3DLVERTEX

    With RET
        .Specular = Specular
        .tu = tu
        .tv = tv
        .x = x
        .y = y
        .z = z
        .color = color
    End With 'RET
    RLVertex = RET

End Function

'=================================
' MakeTLVertex
'=================================
'----------------------------------------
'Name: MakeTLVertex
'----------------------------------------
Function MakeTLVertex(VERT As D3DTLVERTEX, Sx As Single, Sy As Single, SZ As Single, w As Single, c As Long, S As Single, U As Single, V As Single) As D3DTLVERTEX

    VERT.Sx = Sx
    VERT.Sy = Sy
    VERT.SZ = SZ
    VERT.rhw = w
    VERT.color = c
    VERT.Specular = S
    VERT.tu = U
    VERT.tv = V

End Function


'----------------------------------------
'Name: RTLVertex
'----------------------------------------
Function RTLVertex(Sx As Single, Sy As Single, SZ As Single, w As Single, c As Long, S As Single, U As Single, V As Single) As D3DTLVERTEX

  Dim VERT As D3DTLVERTEX

    VERT.Sx = Sx
    VERT.Sy = Sy
    VERT.SZ = SZ
    VERT.rhw = w
    VERT.color = c
    VERT.Specular = S
    VERT.tu = U
    VERT.tv = V
    RTLVertex = VERT

End Function

'=================================
' MakeRect
'=================================
'----------------------------------------
'Name: MakeRect
'----------------------------------------
Function MakeRect(RET As RECT, x1 As Single, y1 As Single, x2 As Single, Y2 As Single) ':( As Variant ?

    With RET
        .Left = x1
        .Top = y1
        .Right = x2
        .bottom = Y2
    End With 'RET

End Function

'-----------------------------------------------------------------------------
' Name: TimerGet()
' Desc: Performs timer opertations. Use the following commands:
'          TIMER_RESET           - to reset the timer
'          TIMER_START           - to start the timer
'          TIMER_STOP            - to stop (or pause) the timer
'          TIMER_ADVANCE         - to advance the timer by 0.1 seconds
'          TIMER_GETABSOLUTETIME - to get the absolute system time
'          TIMER_GETAPPTIME      - to get the current time
'          TIMER_GETELLAPSEDTIME - to get the ellapsed time between calls
'-----------------------------------------------------------------------------
'----------------------------------------
'Name: TimerGet
'----------------------------------------
Function TimerGet(command As TIMER_COMMAND) As Single

    On Local Error Resume Next

      Static m_bTimerInitialized  As Boolean ':( Move line to top of current Function
      Static m_bUsingQPF         As Boolean ':( Move line to top of current Function
      Static m_fSecsPerTick  As Single ':( Move line to top of current Function
      Static m_fBaseTime    As Single ':( Move line to top of current Function
      Static m_fStopTime     As Single ':( Move line to top of current Function
      Static m_fLastTime As Single ':( Move line to top of current Function

      Dim fTime As Single ':( Move line to top of current Function

        ' Initialize the timer
        If (False = m_bTimerInitialized) Then
            m_bTimerInitialized = True
        End If

        If m_fLastTime > Timer Then command = TIMER_RESET 'For the midnight wrap':( Expand Structure
        fTime = Timer

        ' Reset the timer
        If (command = TIMER_RESET) Then
            m_fBaseTime = fTime
            m_fStopTime = 0
            m_fLastTime = 0
            TimerGet = 0
            Exit Function '>---> Bottom
        End If

        ' Return the current time
        If (command = TIMER_GETAPPTIME) Then
            TimerGet = fTime - m_fBaseTime
            Exit Function '>---> Bottom
        End If

        ' Start the timer
        If (command = TIMER_start) Then
            m_fBaseTime = m_fBaseTime + fTime - m_fStopTime
            m_fLastTime = m_fLastTime + fTime - m_fStopTime
            m_fStopTime = 0
        End If

        ' Stop the timer
        If (command = TIMER_STOP) Then
            m_fStopTime = fTime
        End If

        ' Advance the timer by 1/10th second
        If (command = TIMER_ADVANCE) Then
            m_fBaseTime = m_fBaseTime + fTime - (m_fStopTime + 0.1)
        End If

        ' Return ellapsed time
        If (command = TIMER_GETELLAPSEDTIME) Then
            TimerGet = fTime - m_fLastTime
            m_fLastTime = fTime
            If TimerGet < 0 Then TimerGet = 0 ':( Expand Structure
            Exit Function '>---> Bottom
        End If

        TimerGet = fTime

End Function ':( On Error Resume still active

'=================================
' ZeroMatrix
'=================================


'----------------------------------------
'Name: ZeroMatrix
'----------------------------------------
Sub ZeroMatrix(dest As D3DMATRIX)

    dest.m11 = 0
    dest.m12 = 0
    dest.m13 = 0
    dest.m14 = 0
    dest.m21 = 0
    dest.m22 = 0
    dest.m23 = 0
    dest.m24 = 0
    dest.m31 = 0
    dest.m32 = 0
    dest.m33 = 0
    dest.m34 = 0
    dest.m41 = 0
    dest.m42 = 0
    dest.m43 = 0
    dest.m44 = 0

End Sub


'----------------------------------------
'Name: MatrixMult
'----------------------------------------
Sub MatrixMult(result As D3DMATRIX, A As D3DMATRIX, B As D3DMATRIX)

  Dim RET As D3DMATRIX
  Dim tmp As Double
  Dim i As Integer
  Dim j As Integer
  Dim K As Integer

    Call ZeroMatrix(RET)
    RET.m11 = B.m11 * A.m11 + B.m21 * A.m12 + B.m31 * A.m13 + B.m41 * A.m14
    RET.m12 = B.m12 * A.m11 + B.m22 * A.m12 + B.m32 * A.m13 + B.m42 * A.m14
    RET.m13 = B.m13 * A.m11 + B.m23 * A.m12 + B.m33 * A.m13 + B.m43 * A.m14
    RET.m14 = B.m14 * A.m11 + B.m24 * A.m12 + B.m34 * A.m13 + B.m44 * A.m14
    RET.m21 = B.m11 * A.m21 + B.m21 * A.m22 + B.m31 * A.m23 + B.m41 * A.m24
    RET.m22 = B.m12 * A.m21 + B.m22 * A.m22 + B.m32 * A.m23 + B.m42 * A.m24
    RET.m23 = B.m13 * A.m21 + B.m23 * A.m22 + B.m33 * A.m23 + B.m43 * A.m24
    RET.m24 = B.m14 * A.m21 + B.m24 * A.m22 + B.m34 * A.m23 + B.m44 * A.m24
    RET.m31 = B.m11 * A.m31 + B.m21 * A.m32 + B.m31 * A.m33 + B.m41 * A.m34
    RET.m32 = B.m12 * A.m31 + B.m22 * A.m32 + B.m32 * A.m33 + B.m42 * A.m34
    RET.m33 = B.m13 * A.m31 + B.m23 * A.m32 + B.m33 * A.m33 + B.m43 * A.m34
    RET.m34 = B.m14 * A.m31 + B.m24 * A.m32 + B.m34 * A.m33 + B.m44 * A.m34
    RET.m41 = B.m11 * A.m41 + B.m21 * A.m42 + B.m31 * A.m43 + B.m41 * A.m44
    RET.m42 = B.m12 * A.m41 + B.m22 * A.m42 + B.m32 * A.m43 + B.m42 * A.m44
    RET.m43 = B.m13 * A.m41 + B.m23 * A.m42 + B.m33 * A.m43 + B.m43 * A.m44
    RET.m44 = B.m14 * A.m41 + B.m24 * A.m42 + B.m34 * A.m43 + B.m44 * A.m44
    result = RET

End Sub


'----------------------------------------
'Name: RetMatrixMult
'----------------------------------------
Function RetMatrixMult(A As D3DMATRIX, B As D3DMATRIX) As D3DMATRIX

  Dim RET As D3DMATRIX

    RET.m11 = B.m11 * A.m11 + B.m21 * A.m12 + B.m31 * A.m13 + B.m41 * A.m14
    RET.m12 = B.m12 * A.m11 + B.m22 * A.m12 + B.m32 * A.m13 + B.m42 * A.m14
    RET.m13 = B.m13 * A.m11 + B.m23 * A.m12 + B.m33 * A.m13 + B.m43 * A.m14
    RET.m14 = B.m14 * A.m11 + B.m24 * A.m12 + B.m34 * A.m13 + B.m44 * A.m14
    RET.m21 = B.m11 * A.m21 + B.m21 * A.m22 + B.m31 * A.m23 + B.m41 * A.m24
    RET.m22 = B.m12 * A.m21 + B.m22 * A.m22 + B.m32 * A.m23 + B.m42 * A.m24
    RET.m23 = B.m13 * A.m21 + B.m23 * A.m22 + B.m33 * A.m23 + B.m43 * A.m24
    RET.m24 = B.m14 * A.m21 + B.m24 * A.m22 + B.m34 * A.m23 + B.m44 * A.m24
    RET.m31 = B.m11 * A.m31 + B.m21 * A.m32 + B.m31 * A.m33 + B.m41 * A.m34
    RET.m32 = B.m12 * A.m31 + B.m22 * A.m32 + B.m32 * A.m33 + B.m42 * A.m34
    RET.m33 = B.m13 * A.m31 + B.m23 * A.m32 + B.m33 * A.m33 + B.m43 * A.m34
    RET.m34 = B.m14 * A.m31 + B.m24 * A.m32 + B.m34 * A.m33 + B.m44 * A.m34
    RET.m41 = B.m11 * A.m41 + B.m21 * A.m42 + B.m31 * A.m43 + B.m41 * A.m44
    RET.m42 = B.m12 * A.m41 + B.m22 * A.m42 + B.m32 * A.m43 + B.m42 * A.m44
    RET.m43 = B.m13 * A.m41 + B.m23 * A.m42 + B.m33 * A.m43 + B.m43 * A.m44
    RET.m44 = B.m14 * A.m41 + B.m24 * A.m42 + B.m34 * A.m43 + B.m44 * A.m44
    RetMatrixMult = RET

End Function

'=================================
' TranslateMatrix
'=================================
' used to position an object


'----------------------------------------
'Name: TranslateMatrix
'----------------------------------------
Sub TranslateMatrix(M As D3DMATRIX, V As D3DVECTOR)

    Call IdentityMatrix(M)
    M.m41 = V.x
    M.m42 = V.y
    M.m43 = V.z

End Sub


'----------------------------------------
'Name: RetTranslateMatrix
'----------------------------------------
Function RetTranslateMatrix(V As D3DVECTOR) As D3DMATRIX

  Dim M As D3DMATRIX

    Call IdentityMatrix(M)
    M.m41 = V.x
    M.m42 = V.y
    M.m43 = V.z
    RetTranslateMatrix = M

End Function

'=================================
' RotateXMatrix
'=================================
' rotate an object about x axis rad radians


'----------------------------------------
'Name: RotateXMatrix
'----------------------------------------
Sub RotateXMatrix(RET As D3DMATRIX, rads As Single)

  Dim cosine As Single
  Dim sine As Single

    cosine = Cos(rads)
    sine = Sin(rads)
    Call IdentityMatrix(RET)
    RET.m22 = cosine
    RET.m33 = cosine
    RET.m23 = -sine
    RET.m32 = sine

End Sub


'----------------------------------------
'Name: RetRotateXMatrix
'----------------------------------------
Function RetRotateXMatrix(rads As Single) As D3DMATRIX

  Dim cosine As Single
  Dim sine As Single
  Dim RET As D3DMATRIX

    cosine = Cos(rads)
    sine = Sin(rads)
    Call IdentityMatrix(RET)
    RET.m22 = cosine
    RET.m33 = cosine
    RET.m23 = -sine
    RET.m32 = sine
    RetRotateXMatrix = RET

End Function

'=================================
' RotateYMatrix
'=================================
' rotate an object about y axis rad radians


'----------------------------------------
'Name: RotateYMatrix
'----------------------------------------
Sub RotateYMatrix(RET As D3DMATRIX, rads As Single)

  Dim cosine As Single
  Dim sine As Single

    cosine = Cos(rads)
    sine = Sin(rads)
    Call IdentityMatrix(RET)
    RET.m11 = cosine
    RET.m33 = cosine
    RET.m13 = sine
    RET.m31 = -sine

End Sub


'----------------------------------------
'Name: RetRotateYMatrix
'----------------------------------------
Function RetRotateYMatrix(rads As Single) As D3DMATRIX

  Dim cosine As Single
  Dim sine As Single
  Dim RET As D3DMATRIX

    cosine = Cos(rads)
    sine = Sin(rads)
    Call IdentityMatrix(RET)
    RET.m11 = cosine
    RET.m33 = cosine
    RET.m13 = sine
    RET.m31 = -sine
    RetRotateYMatrix = RET

End Function

'=================================
' RotateZMatrix
'=================================
' rotate an object about z axis rad radians


'----------------------------------------
'Name: RotateZMatrix
'----------------------------------------
Sub RotateZMatrix(RET As D3DMATRIX, rads As Single)

  Dim cosine As Single
  Dim sine As Single

    cosine = Cos(rads)
    sine = Sin(rads)
    Call IdentityMatrix(RET)
    RET.m11 = cosine
    RET.m22 = cosine
    RET.m12 = -sine
    RET.m21 = sine

End Sub


'----------------------------------------
'Name: RetRotateZMatrix
'----------------------------------------
Function RetRotateZMatrix(rads As Single) As D3DMATRIX

  Dim RET As D3DMATRIX
  Dim cosine As Single
  Dim sine As Single

    cosine = Cos(rads)
    sine = Sin(rads)
    Call IdentityMatrix(RET)
    RET.m11 = cosine
    RET.m22 = cosine
    RET.m12 = -sine
    RET.m21 = sine
    RetRotateZMatrix = RET

End Function


'----------------------------------------
'Name: ViewMatrix
'----------------------------------------
Sub ViewMatrix(view As D3DMATRIX, from As D3DVECTOR, At As D3DVECTOR, world_up As D3DVECTOR, roll As Single)

  Dim Up As D3DVECTOR
  Dim Right As D3DVECTOR
  Dim view_Dir As D3DVECTOR

    Call IdentityMatrix(view)
    Call VectorSubtract(view_Dir, At, from)
    Call VectorNormalize(view_Dir)

    'think lefthanded coords
    Call VectorCrossProduct(Right, world_up, view_Dir)
    Call VectorCrossProduct(Up, view_Dir, Right)

    Call VectorNormalize(Right)
    Call VectorNormalize(Up)

    view.m11 = Right.x
    view.m21 = Right.y
    view.m31 = Right.z
    view.m12 = Up.x   'AK? should this be negative?
    view.m22 = Up.y
    view.m32 = Up.z
    view.m13 = view_Dir.x
    view.m23 = view_Dir.y
    view.m33 = view_Dir.z

    view.m41 = -VectorDotProduct(Right, from)
    view.m42 = -VectorDotProduct(Up, from)
    view.m43 = -VectorDotProduct(view_Dir, from)

    ' Set roll
    If (roll <> 0#) Then
  Dim rotZMat As D3DMATRIX ':( Move line to top of current Sub
        Call RotateZMatrix(rotZMat, -roll)
        Call MatrixMult(view, rotZMat, view)
    End If

End Sub


'----------------------------------------
'Name: PowerOf2
'----------------------------------------
Function PowerOf2(x) ':( As Variant ?

    If (x > 256) Then PowerOf2 = 512: Exit Function ':( Expand Structure or consider reversing Condition
    If (x > 128) Then PowerOf2 = 256: Exit Function ':( Expand Structure or consider reversing Condition
    If (x > 64) Then PowerOf2 = 128: Exit Function ':( Expand Structure or consider reversing Condition
    If (x > 32) Then PowerOf2 = 64: Exit Function ':( Expand Structure or consider reversing Condition
    If (x > 16) Then PowerOf2 = 32: Exit Function ':( Expand Structure or consider reversing Condition
    If (x > 8) Then PowerOf2 = 16: Exit Function ':( Expand Structure or consider reversing Condition
    If (x > 4) Then PowerOf2 = 8: Exit Function ':( Expand Structure or consider reversing Condition
    If (x > 2) Then PowerOf2 = 4: Exit Function ':( Expand Structure or consider reversing Condition
    If (x > 1) Then PowerOf2 = 2: Exit Function ':( Expand Structure or consider reversing Condition

    PowerOf2 = 1

Exit Function

End Function


'----------------------------------------
'Name: IsPowerOf2
'----------------------------------------
Function IsPowerOf2(x) As Boolean

  Dim P ':( As Variant ?

    P = 1
  Dim i ':( As Variant ?':( Move line to top of current Function
    For i = 0 To 31
        P = P * 2 ^ 1
        If P = x Then
            IsPowerOf2 = True
        End If
    Next i

End Function


'----------------------------------------
'Name: Arcsin
'----------------------------------------
Function Arcsin(x) ':( As Variant ?

  Dim T ':( As Variant ?

    T = x / Sqr(-x * x + 1)
    Arcsin = Atn(T)

End Function


'----------------------------------------
'Name: Arccos
'----------------------------------------
Function Arccos(x) ':( As Variant ?

    Arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)

End Function


'----------------------------------------
'Name: Arcsec
'----------------------------------------
Function Arcsec(x) ':( As Variant ?

    Arcsec = Atn(x / Sqr(x * x - 1)) + Sgn((x) - 1) * (2 * Atn(1))

End Function


'----------------------------------------
'Name: Arccosec
'----------------------------------------
Function Arccosec(x) ':( As Variant ?

    Arccosec = Atn(x / Sqr(x * x - 1)) + (Sgn(x) - 1) * (2 * Atn(1))

End Function


'----------------------------------------
'Name: Arccotan
'----------------------------------------
Function Arccotan(x) ':( As Variant ?

    Arccotan = Atn(x) + 2 * Atn(1)

End Function


'----------------------------------------
'Name: HSin
'----------------------------------------
Function HSin(x) ':( As Variant ?

    HSin = (Exp(x) - Exp(-x)) / 2

End Function


'----------------------------------------
'Name: HCos
'----------------------------------------
Function HCos(x) ':( As Variant ?

    HCos = (Exp(x) + Exp(-x)) / 2

End Function


'----------------------------------------
'Name: HTan
'----------------------------------------
Function HTan(x) ':( As Variant ?

    HTan = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))

End Function


'----------------------------------------
'Name: HSec
'----------------------------------------
Function HSec(x) ':( As Variant ?

    HSec = 2 / (Exp(x) + Exp(-x))

End Function


'----------------------------------------
'Name: HCosec
'----------------------------------------
Function HCosec(x) ':( As Variant ?

    HCosec = 2 / (Exp(x) - Exp(-x))

End Function


'----------------------------------------
'Name: HCotan
'----------------------------------------
Function HCotan(x) ':( As Variant ?

    HCotan = (Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x))

End Function


'----------------------------------------
'Name: HArcsin
'----------------------------------------
Function HArcsin(x) ':( As Variant ?

    HArcsin = Log(x + Sqr(x * x + 1))

End Function


'----------------------------------------
'Name: HArccos
'----------------------------------------
Function HArccos(x) ':( As Variant ?

    HArccos = Log(x + Sqr(x * x - 1))

End Function


'----------------------------------------
'Name: HArctan
'----------------------------------------
Function HArctan(x) ':( As Variant ?

    HArctan = Log((1 + x) / (1 - x)) / 2

End Function


'----------------------------------------
'Name: HArcsec
'----------------------------------------
Function HArcsec(x) ':( As Variant ?

    HArcsec = Log((Sqr(-x * x + 1) + 1) / x)

End Function


'----------------------------------------
'Name: HArccosec
'----------------------------------------
Function HArccosec(x) ':( As Variant ?

    HArccosec = Log((Sgn(x) * Sqr(x * x + 1) + 1) / x)

End Function


'----------------------------------------
'Name: HArccotan
'----------------------------------------
Function HArccotan(x) ':( As Variant ?

    HArccotan = Log((x + 1) / (x - 1)) / 2

End Function


'----------------------------------------
'Name: LogN
'----------------------------------------
Function LogN(x) ':( As Variant ?

    LogN = Log(x) / Log(10)

End Function


'----------------------------------------
'Name: RenderBSPEX
'----------------------------------------
Function RenderBSPEX(ByVal i As Long) As Boolean

    'lpBSP.RenderFace (i)

    'RenderBSPEX = True

End Function



'----------------------------------------
'Name: RenderFastPoly
'----------------------------------------
Function RenderFastPoly(P As NemoPOLYGON) As Boolean

If P.LightmapID > -1 Then
  RenderFastPoly2 P

Else
   Dim i
   Dim p1(0) As NemoPOLYGON
   Dim VERT(3) As NEMO_VERTEX
   
   p1(0) = P
   NemoPOLYGON_BuildVertex p1(0), VERT(0), 1
   
   
   D3dDevice.SetTexture 0, POOL_texture(P.TextureID)
   D3dDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 1, VERT(0), Len(VERT(0))
   
 End If
   
   
  RenderFastPoly = True
End Function



'----------------------------------------
'Name: RenderFastPoly2
'----------------------------------------
Private Function RenderFastPoly2(P As NemoPOLYGON)


    'set shaders and texture
 Dim i
   Dim p1(0) As NemoPOLYGON
   Dim VERT(3) As NEMO_VERTEX2
   
   p1(0) = P
   i = P.Vertex(1).x
   NemoPOLYGON_BuildVertex2 p1(0), VERT(0), 1
   
   D3dDevice.SetVertexShader NEMO_CUSTOM_VERTEX2
   D3dDevice.SetTexture 0, POOL_texture(P.TextureID)
   D3dDevice.SetTexture 1, POOL_Lightmaps(P.LightmapID)
   
    Call D3dDevice.SetTextureStageState(1, D3DTSS_COLOROP, D3DTOP_MODULATE)
    Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG2, D3DTA_TEXTURE)
    Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG1, D3DTA_CURRENT)

    ''Light Map
    
            Call D3dDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
            Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_TEXTURE)
            Call D3dDevice.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_DIFFUSE)

      

    'Render stuff
   D3dDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, VERT(0), Len(VERT(0))

   
            Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG1, D3DTOP_DISABLE)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_COLOROP, D3DTOP_DISABLE)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_COLORARG2, D3DTOP_DISABLE)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_ALPHAARG1, D3DTOP_DISABLE)
            Call D3dDevice.SetTextureStageState(1, D3DTSS_ALPHAOP, D3DTOP_DISABLE)

    

End Function



'----------------------------------------
'Name: Secante
'----------------------------------------
Function Secante(x) ':( As Variant ?

    Secante = 1 / Cos(x)

End Function


'----------------------------------------
'Name: Cosecante
'----------------------------------------
Function Cosecante(x) ':( As Variant ?

    Cosecante = 1 / Sin(x)

End Function


'----------------------------------------
'Name: Cotangente
'----------------------------------------
Function Cotangente(x) ':( As Variant ?

    Cotangente = 1 / Tan(x)

End Function


