VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNemo_Mesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'=============================================
' Nemo Mesh class by Johna on Sept 8th 2002
'
'this class provides easy and fast methods for
'building a mesh from..poly,vertices ect..
'
'it uses Poly patition wich are defined by Unique TextureID
'
'
'19 sept
'Now there are 2 rendering method
'  render with only textures
'  render2 with lighmaps
' all the poly with common textures are sorted for fast
' rendering
'
' render2 allow per-poly frustumculling
'
'===============================================

Option Explicit
Private Const MAX_VERTICES = 30000

Private Type tSinglePOLY
    'POLY() As NemoPOLYGON
    NumVERT As Long
    VERT() As NEMO_VERTEX2
    ZZindex As Long
End Type

Private Type NemoMeshTextureInfo
    Filename As String * 64
End Type

Private Type Tmesh
    TextureInfo() As NemoMeshTextureInfo
    TexturesID() As Long 'local texture index
    LightmapsID() As Long 'local lighmap index
    PolyTexIndex() As Long
    Vertex() As NEMO_VERTEX2
    'Vertex2() As NEMO_VERTEX2
    'Lvertex() as nemo_v
    NumVERT As Long
    NumText As Long
    NumLightmaps As Long
    NumPOLY As Long
    NumPrimitives As Long

    BBOX As NemoBoundBOX
    IS_Loaded As Boolean

    POLY_sorted() As tSinglePOLY
    POLY_unsorted As Long
    POLYcoldet() As NemoPOLYGON
    Light() As NemoStaticLight
    numLight As Long
End Type

Private DrawStyle As CONST_D3DPRIMITIVETYPE

Private mMESH As Tmesh
Dim PolyCount As Long
Dim PolyIndex
Dim LoadFromFile As Boolean

'----------------------------------------
'Name: SetTriangleDrawing
'----------------------------------------
Private Sub SetTriangleDrawing(Optional TriangleDrawinstyle As CONST_D3DPRIMITIVETYPE = D3DPT_TRIANGLELIST)

    DrawStyle = TriangleDrawinstyle
    Select Case TriangleDrawinstyle

      Case D3DPT_TRIANGLESTRIP
        mMESH.NumPrimitives = (mMESH.NumVERT - 3) + 1
      Case D3DPT_TRIANGLELIST
        mMESH.NumPrimitives = Int((mMESH.NumVERT) / 3)

    End Select

End Sub

'----------------------------------------
'Name: Class_Initialize
'Object: Class
'Event: Initialize
'----------------------------------------
Private Sub Class_Initialize()

    DrawStyle = D3DPT_TRIANGLELIST
    ReDim Preserve mMESH.POLYcoldet(PolyCount)
    PolyIndex = 0

End Sub

'----------------------------------------
'Name: Add_Texture
'Object: Add
'Event: Texture
'----------------------------------------
Sub Add_Texture(TextureFile As String)

  'If GLOB.FileiS_valid(TextureFile) Then

    mMESH.NumText = mMESH.NumText + 1
    ReDim Preserve mMESH.TexturesID(mMESH.NumText - 1)
    ReDim Preserve mMESH.TextureInfo(mMESH.NumText - 1)
  Dim tempStr As String
    tempStr = GLOB.GET_fileNameEX(TextureFile)

    CopyMemory mMESH.TextureInfo(mMESH.NumText - 1).Filename, tempStr, Len(tempStr)

    'Set mMESH.Textures(mMESH.NumText - 1) = GLOB.LpGLOBAL_NEMO.CreateTexture(TextureFile)
    'End If
    'ReDim Preserve mMESH.TexturesID(mMESH.NumText - 1)
    mMESH.TexturesID(mMESH.NumText - 1) = GLOB.Add_TextureToPool(TextureFile)

End Sub

'Sub Add_Texture(TextureFile As String)

'If GLOB.FileiS_valid(TextureFile) Then
' mMESH.NumText = mMESH.NumText + 1
'ReDim Preserve mMESH.TexturesID(mMESH.NumText - 1)
'Set mMESH.Textures(mMESH.NumText - 1) = GLOB.LpGLOBAL_NEMO.CreateTexture(TextureFile)
'End If
'ReDim Preserve mMESH.TexturesID(mMESH.NumText - 1)
'mMESH.TexturesID(mMESH.NumText - 1) = GLOB.Add_TextureToPool(TextureFile)

'End Sub

'----------------------------------------
'Name: Add_Light
'Object: Add
'Event: Light
'----------------------------------------
Sub Add_Light(R, G, B, radius, Position As D3DVECTOR, Optional brightness = 100)

    mMESH.numLight = mMESH.numLight + 1
    ReDim Preserve mMESH.Light(mMESH.numLight - 1)

    With mMESH.Light(mMESH.numLight - 1)
        .radius = radius
        .Position.x = Position.x
        .Position.y = Position.y
        .Position.z = Position.z

        .red = R
        .green = G
        .blue = B
        .brightness = brightness

    End With 'MMESH.LIGHT(MMESH.NUMLIGHT

End Sub

'----------------------------------------
'Name: Add_Vertex
'Object: Add
'Event: Vertex
'----------------------------------------
Sub Add_Vertex(x As Single, y As Single, z As Single, Nx As Single, Ny As Single, Nz As Single, tu As Single, tv As Single, Optional TextureID As Long = 0, Optional LightmapID As Long = -1)

    If PolyIndex > 2 Or LoadFromFile Then
        PolyIndex = 0
        PolyCount = PolyCount + 1
        ReDim Preserve mMESH.POLYcoldet(PolyCount)
        LoadFromFile = False
    End If

    'fill poly for collision detection
    mMESH.POLYcoldet(PolyCount).Vertex(PolyIndex).x = x
    mMESH.POLYcoldet(PolyCount).Vertex(PolyIndex).y = y
    mMESH.POLYcoldet(PolyCount).Vertex(PolyIndex).z = z

    mMESH.POLYcoldet(PolyCount).Vertex(PolyIndex).tu1 = tu
    mMESH.POLYcoldet(PolyCount).Vertex(PolyIndex).tv1 = tv

    mMESH.POLYcoldet(PolyCount).Vertex(PolyIndex).Nx = Nx
    mMESH.POLYcoldet(PolyCount).Vertex(PolyIndex).Ny = Ny
    mMESH.POLYcoldet(PolyCount).Vertex(PolyIndex).Nz = Nz

    mMESH.POLYcoldet(PolyCount).TextureID = mMESH.TexturesID(TextureID)
    mMESH.POLYcoldet(PolyCount).LightmapID = LightmapID

    If TextureID = 3 Then
    
      Debug.Print "uuu"
    End If

    PolyIndex = PolyIndex + 1

    With mMESH
        .NumVERT = .NumVERT + 1
        ReDim Preserve .PolyTexIndex(.NumVERT - 1)
        ReDim Preserve .Vertex(.NumVERT - 1)

        'fill the vertex and its texture ID
        .PolyTexIndex(.NumVERT - 1) = TextureID
        .Vertex(.NumVERT - 1).x = x
        .Vertex(.NumVERT - 1).y = y
        .Vertex(.NumVERT - 1).z = z
        .Vertex(.NumVERT - 1).Nx = Nx
        .Vertex(.NumVERT - 1).Ny = Ny
        .Vertex(.NumVERT - 1).Nz = Nz
        .Vertex(.NumVERT - 1).tu1 = tu
        .Vertex(.NumVERT - 1).tv1 = tv

        .IS_Loaded = True
    End With 'MMESH

End Sub

'----------------------------------------
'Name: Add_VertexEX
'Object: Add
'Event: VertexEX
'----------------------------------------
Sub Add_VertexEX(vPos As D3DVECTOR, Vnormal As D3DVECTOR, TexUV As D3DVECTOR2, Optional TextureID As Long = 0)

    Call Add_Vertex(vPos.x, vPos.y, vPos.z, Vnormal.x, Vnormal.y, Vnormal.z, TexUV.x, TexUV.y, TextureID)

End Sub

'----------------------------------------
'Name: Add_VertexFromPoly
'Object: Add
'Event: VertexFromPoly
'----------------------------------------
Sub Add_VertexFromPoly(POLY As NemoPOLYGON, Optional TextureID As Long = 0)

    Call Add_Vertex(POLY.Vertex(0).x, POLY.Vertex(0).y, POLY.Vertex(0).z, POLY.Vertex(0).Nx, POLY.Vertex(0).Ny, POLY.Vertex(0).Nz, POLY.Vertex(0).tv1, POLY.Vertex(0).tv1, TextureID)
    Call Add_Vertex(POLY.Vertex(1).x, POLY.Vertex(1).y, POLY.Vertex(1).z, POLY.Vertex(1).Nx, POLY.Vertex(1).Ny, POLY.Vertex(1).Nz, POLY.Vertex(1).tv1, POLY.Vertex(1).tv1, TextureID)
    Call Add_Vertex(POLY.Vertex(2).x, POLY.Vertex(2).y, POLY.Vertex(2).z, POLY.Vertex(2).Nx, POLY.Vertex(2).Ny, POLY.Vertex(2).Nz, POLY.Vertex(2).tv1, POLY.Vertex(2).tv1, TextureID)

End Sub

'----------------------------------------
'Name: Add_VertexFromPolys
'Object: Add
'Event: VertexFromPolys
'----------------------------------------
Sub Add_VertexFromPolys(POLY() As NemoPOLYGON)

  Dim i As Long

    For i = 0 To UBound(POLY())
        Call Add_Vertex(POLY(i).Vertex(0).x, POLY(i).Vertex(0).y, POLY(i).Vertex(0).z, POLY(i).Vertex(0).Nx, POLY(i).Vertex(0).Ny, POLY(i).Vertex(0).Nz, POLY(i).Vertex(0).tu1, POLY(i).Vertex(0).tv1, POLY(i).TextureID)
        Call Add_Vertex(POLY(i).Vertex(1).x, POLY(i).Vertex(1).y, POLY(i).Vertex(1).z, POLY(i).Vertex(1).Nx, POLY(i).Vertex(1).Ny, POLY(i).Vertex(1).Nz, POLY(i).Vertex(1).tu1, POLY(i).Vertex(1).tv1, POLY(i).TextureID)
        Call Add_Vertex(POLY(i).Vertex(2).x, POLY(i).Vertex(2).y, POLY(i).Vertex(2).z, POLY(i).Vertex(2).Nx, POLY(i).Vertex(2).Ny, POLY(i).Vertex(2).Nz, POLY(i).Vertex(2).tu1, POLY(i).Vertex(2).tv1, POLY(i).TextureID)
    Next i

End Sub

'============================================
'
'
'REORGANIZE POLYGONS BY TEXTURE
'
'
'
'================================================

'----------------------------------------
'Name: BuilMesh
'----------------------------------------
Sub BuilMesh()

  Dim i, j, z

    SetTriangleDrawing

    With mMESH
        ReDim .POLY_sorted(.NumText - 1)
        For i = 0 To .NumText - 1
            ReDim .POLY_sorted(i).VERT(MAX_VERTICES - 1)
        Next i

        For i = 0 To .NumVERT - 1

            .POLY_sorted(.PolyTexIndex(i)).VERT(.POLY_sorted(.PolyTexIndex(i)).ZZindex).x = .Vertex(i).x
            .POLY_sorted(.PolyTexIndex(i)).VERT(.POLY_sorted(.PolyTexIndex(i)).ZZindex).y = .Vertex(i).y
            .POLY_sorted(.PolyTexIndex(i)).VERT(.POLY_sorted(.PolyTexIndex(i)).ZZindex).z = .Vertex(i).z

            .POLY_sorted(.PolyTexIndex(i)).VERT(.POLY_sorted(.PolyTexIndex(i)).ZZindex).Nx = .Vertex(i).Nx
            .POLY_sorted(.PolyTexIndex(i)).VERT(.POLY_sorted(.PolyTexIndex(i)).ZZindex).Ny = .Vertex(i).Ny
            .POLY_sorted(.PolyTexIndex(i)).VERT(.POLY_sorted(.PolyTexIndex(i)).ZZindex).Nz = .Vertex(i).Nz

            .POLY_sorted(.PolyTexIndex(i)).VERT(.POLY_sorted(.PolyTexIndex(i)).ZZindex).tu1 = .Vertex(i).tu1
            .POLY_sorted(.PolyTexIndex(i)).VERT(.POLY_sorted(.PolyTexIndex(i)).ZZindex).tv1 = .Vertex(i).tv1
            .POLY_sorted(.PolyTexIndex(i)).ZZindex = .POLY_sorted(.PolyTexIndex(i)).ZZindex + 1

        Next i

        For i = 0 To .NumText - 1
            ReDim Preserve .POLY_sorted(i).VERT(.POLY_sorted(i).ZZindex)
            .POLY_sorted(i).NumVERT = .POLY_sorted(i).ZZindex

        Next i

    End With 'MMESH

    Call Buildlightmap

End Sub

Sub GetMesmPolygons(ByRef lpArray() As NemoPOLYGON)

    ReDim lpArray(PolyCount)

    CopyMemory lpArray(0), mMESH.POLYcoldet(0), Len(mMESH.POLYcoldet(0)) * (PolyCount + 1)

End Sub

Sub SaveMesh(File As String)

    mMESH.NumPOLY = PolyCount + 1

  Dim F

    F = FreeFile
    Open File For Binary As F
    Put #F, , mMESH
    Close F

End Sub

Sub LoadMesh(File As String)

  Dim Pmesh As Tmesh
  Dim TextID() As Long
  

  Dim F, i

    F = FreeFile

    Open File For Binary As F
    Get #F, , Pmesh
    Close F

  Dim ST As String
  Dim MeshPath As String
    MeshPath = GLOB.GET_pathName(File)
    ReDim TextID(Pmesh.NumText - 1)
    For i = 0 To Pmesh.NumText - 1
        ST = GLOB.KillSpace(Pmesh.TextureInfo(i).Filename)
        Add_Texture MeshPath + "\" + ST
        TextID(i) = mMESH.TexturesID(i)
    Next i
  
    PolyCount = Pmesh.NumPOLY - 1
    
    'restaure real texturepool info
    
    
    mMESH = Pmesh
    
    For i = 0 To Pmesh.NumText - 1
       
         mMESH.TexturesID(i) = TextID(i)
    Next i
    
    
        LoadFromFile = True
        
End Sub

'----------------------------------------
'Name: Buildlightmap
'----------------------------------------
Private Sub Buildlightmap()

    If mMESH.numLight <= 0 Then Exit Sub ':( Expand Structure or consider reversing Condition

  Dim LIT() As Byte
    ReDim LIT(16 * 16 * 3)
  Dim LI() As Nemolight16
  Dim i

    ReDim LI(PolyCount) As Nemolight16

    NemoPOLYGON_MakeLightmaps PolyCount + 1, mMESH.POLYcoldet(0), mMESH.Light(0), mMESH.numLight, LI(0)

    i = LI(0).RGB(210)

    For i = 0 To PolyCount
        mMESH.POLYcoldet(i).LightmapID = -1
        CopyMemory LIT(0), LI(i).RGB(0), Len(LIT(0)) * 16 * 16 * 3
        mMESH.POLYcoldet(i).LightmapID = GLOB.Add_LightmapToPoolEX(GLOB.CreateTextureFromBuffer(LIT, 16, 16))

    Next i

End Sub

'----------------------------------------
'Name: Add_WallFloor
'Object: Add
'Event: WallFloor
'----------------------------------------
Sub Add_WallFloor(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional tu As Single = 1, Optional tv As Single = 1, Optional TextureID As Long = 0)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight

    Call Add_Vertex(v1.x, v2.y, v2.z, 0, 1, 0, 0, 0, TextureID)
    Call Add_Vertex(v2.x, v2.y, v2.z, 0, 1, 0, tu, 0, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, 0, 1, 0, 0, tv, TextureID)

    Call Add_Vertex(v2.x, v2.y, v1.z, 0, 1, 0, tu, tv, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, 0, 1, 0, 0, tv, TextureID)
    Call Add_Vertex(v2.x, v2.y, v2.z, 0, 1, 0, tu, 0, TextureID)

End Sub

'----------------------------------------
'Name: Add_WallRoof
'Object: Add
'Event: WallRoof
'----------------------------------------
Sub Add_WallRoof(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional tu As Single = 1, Optional tv As Single = 1, Optional TextureID As Long = 0)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight

    Call Add_Vertex(v2.x, v2.y, v2.z, 0, -1, 0, tu, 0, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, 0, -1, 0, 0, tv, TextureID)
    Call Add_Vertex(v2.x, v2.y, v1.z, 0, -1, 0, tu, tv, TextureID)

    Call Add_Vertex(v1.x, v2.y, v1.z, 0, -1, 0, 0, tv, TextureID)
    Call Add_Vertex(v2.x, v2.y, v2.z, 0, -1, 0, tu, 0, TextureID)
    Call Add_Vertex(v1.x, v2.y, v2.z, 0, -1, 0, 0, 0, TextureID)

End Sub

'----------------------------------------
'Name: Add_WallLeft
'Object: Add
'Event: WallLeft
'----------------------------------------
Sub Add_WallLeft(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional tu As Single = 1, Optional tv As Single = 1, Optional TextureID As Long = 0)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight

    Call Add_Vertex(v1.x, v1.y, v1.z, 1, 0, 0, 0, 0, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, 1, 0, 0, tu, 0, TextureID)
    Call Add_Vertex(v1.x, v1.y, v2.z, 1, 0, 0, 0, tv, TextureID)

    Call Add_Vertex(v2.x, v2.y, v2.z, 1, 0, 0, tu, tv, TextureID)
    Call Add_Vertex(v1.x, v1.y, v2.z, 1, 0, 0, 0, tv, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, 1, 0, 0, tu, 0, TextureID)

End Sub

'----------------------------------------
'Name: Add_WallRight
'Object: Add
'Event: WallRight
'----------------------------------------
Sub Add_WallRight(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional tu As Single = 1, Optional tv As Single = 1, Optional TextureID As Long = 0)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight

    Call Add_Vertex(v1.x, v2.y, v1.z, -1, 0, 0, tu, 0, TextureID)
    Call Add_Vertex(v1.x, v1.y, v2.z, -1, 0, 0, 0, tv, TextureID)
    Call Add_Vertex(v2.x, v2.y, v2.z, -1, 0, 0, tu, tv, TextureID)

    Call Add_Vertex(v1.x, v1.y, v2.z, -1, 0, 0, 0, tv, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, -1, 0, 0, tu, 0, TextureID)
    Call Add_Vertex(v1.x, v1.y, v1.z, -1, 0, 0, 0, 0, TextureID)

End Sub

'----------------------------------------
'Name: Add_WallBack
'Object: Add
'Event: WallBack
'----------------------------------------
Sub Add_WallBack(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional tu As Single = 1, Optional tv As Single = 1, Optional TextureID As Long = 0)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight

    Call Add_Vertex(v1.x, v1.y, v1.z, 0, -1, 0, 0, tv, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, 0, -1, 0, 0, 0, TextureID)
    Call Add_Vertex(v2.x, v1.y, v2.z, 0, -1, 0, tu, tv, TextureID)

    Call Add_Vertex(v1.x, v2.y, v1.z, 0, -1, 0, 0, 0, TextureID)
    Call Add_Vertex(v2.x, v2.y, v2.z, 0, -1, 0, tu, 0, TextureID)
    Call Add_Vertex(v2.x, v1.y, v1.z, 0, -1, 0, tu, tv, TextureID)

End Sub

'----------------------------------------
'Name: Add_WallFront
'Object: Add
'Event: WallFront
'----------------------------------------
Sub Add_WallFront(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional tu As Single = 1, Optional tv As Single = 1, Optional TextureID As Long = 0)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight

    Call Add_Vertex(v2.x, v1.y, v2.z, 0, 1, 0, tu, tv, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, 0, 1, 0, 0, 0, TextureID)
    Call Add_Vertex(v1.x, v1.y, v1.z, 0, 1, 0, 0, tv, TextureID)

    '_________
    Call Add_Vertex(v2.x, v1.y, v1.z, 0, 1, 0, tu, tv, TextureID)
    Call Add_Vertex(v2.x, v2.y, v2.z, 0, 1, 0, tu, 0, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, 0, 1, 0, 0, 0, TextureID)

End Sub

'----------------------------------------
'Name: Add_CorridorBlock
'Object: Add
'Event: CorridorBlock
'----------------------------------------
Sub Add_CorridorBlock(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional TextureIDLeft As Long = 0, Optional TextureIDRight As Long = 0, Optional TextureIDRoof As Long = 0, Optional TextureIDfloor As Long = 0, Optional tu As Single = 1, Optional tv As Single = 1)

    Add_WallFloor Vector(VBotomLeft.x, VBotomLeft.y, VBotomLeft.z), Vector(VTopRight.x, VBotomLeft.y, VTopRight.z), tu, tv, TextureIDfloor
    Add_WallLeft Vector(VBotomLeft.x, VBotomLeft.y, VBotomLeft.z), Vector(VBotomLeft.x, VTopRight.y, VTopRight.z), tu, tv, TextureIDLeft
    Add_WallRight Vector(VTopRight.x, VBotomLeft.y, VBotomLeft.z), Vector(VTopRight.x, VTopRight.y, VTopRight.z), tu, tv, TextureIDRight
    Add_WallRoof Vector(VBotomLeft.x, VTopRight.y, VBotomLeft.z), Vector(VTopRight.x, VTopRight.y, VTopRight.z), tu, tv, TextureIDRoof

End Sub

'----------------------------------------
'Name: Add_Box
'Object: Add
'Event: Box
'----------------------------------------
Sub Add_Box(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional TextureIDLeft As Long = 0, Optional TextureIDRight As Long = 0, Optional TextureIDTop As Long = 0, Optional TextureIDBottom As Long = 0, Optional TextureIDFront As Long = 0, Optional TextureIDBack As Long = 0, Optional tu As Single = 1, Optional tv As Single = 1)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight

    ' Create vertices describing the front face of the cube.
    Call Add_Vertex(v1.x, v1.y, v1.z, 0, -1, 0, 0, tv, TextureIDFront)
    Call Add_Vertex(v1.x, v2.y, v1.z, 0, -1, 0, 0, 0, TextureIDFront)
    Call Add_Vertex(v2.x, v1.y, v1.z, 0, -1, 0, tu, tv, TextureIDFront)

    Call Add_Vertex(v1.x, v2.y, v1.z, 0, -1, 0, 0, 0, TextureIDFront)
    Call Add_Vertex(v2.x, v2.y, v1.z, 0, -1, 0, tu, 0, TextureIDFront)
    Call Add_Vertex(v2.x, v1.y, v1.z, 0, -1, 0, tu, tv, TextureIDFront)

    ' Create vertices describing the back face of the cube.
    Call Add_Vertex(v2.x, v1.y, v2.z, 0, -1, 0, 0, tv, TextureIDBack)
    Call Add_Vertex(v2.x, v2.y, v2.z, 0, -1, 0, 0, 0, TextureIDBack)
    Call Add_Vertex(v1.x, v1.y, v2.z, 0, -1, 0, tu, tv, TextureIDBack)

    Call Add_Vertex(v2.x, v2.y, v2.z, 0, 1, 0, 0, 0, TextureIDBack)
    Call Add_Vertex(v1.x, v2.y, v2.z, 0, 1, 0, tu, 0, TextureIDBack)
    Call Add_Vertex(v1.x, v1.y, v2.z, 0, 1, 0, tu, tv, TextureIDBack)

    ' Create vertices describing the left face of the cube.

    Call Add_Vertex(v1.x, v2.y, v1.z, -1, 0, 0, tu, 0, TextureIDLeft)
    Call Add_Vertex(v1.x, v1.y, v2.z, -1, 0, 0, 0, tv, TextureIDLeft)
    Call Add_Vertex(v1.x, v2.y, v2.z, -1, 0, 0, tu, tv, TextureIDLeft)

    Call Add_Vertex(v1.x, v1.y, v2.z, -1, 0, 0, 0, tv, TextureIDLeft)
    Call Add_Vertex(v1.x, v2.y, v1.z, -1, 0, 0, tu, 0, TextureIDLeft)
    Call Add_Vertex(v1.x, v1.y, v1.z, -1, 0, 0, 0, 0, TextureIDLeft)

    ' Create vertices describing the right face of the cube.
    Call Add_Vertex(v2.x, v1.y, v1.z, 1, 0, 0, 0, 0, TextureIDRight)
    Call Add_Vertex(v2.x, v2.y, v1.z, 1, 0, 0, tu, 0, TextureIDRight)
    Call Add_Vertex(v2.x, v1.y, v2.z, 1, 0, 0, 0, tv, TextureIDRight)

    Call Add_Vertex(v2.x, v2.y, v2.z, 1, 0, 0, tu, tv, TextureIDRight)
    Call Add_Vertex(v2.x, v1.y, v2.z, 1, 0, 0, 0, tv, TextureIDRight)
    Call Add_Vertex(v2.x, v2.y, v1.z, 1, 0, 0, tu, 0, TextureIDRight)

    ' Create vertices describing the top face of the cube.
    Call Add_Vertex(v1.x, v2.y, v2.z, 0, 1, 0, 0, 0, TextureIDTop)
    Call Add_Vertex(v2.x, v2.y, v2.z, 0, 1, 0, tu, 0, TextureIDTop)
    Call Add_Vertex(v1.x, v2.y, v1.z, 0, 1, 0, 0, tv, TextureIDTop)

    Call Add_Vertex(v2.x, v2.y, v1.z, 0, 1, 0, tu, tv, TextureIDTop)
    Call Add_Vertex(v1.x, v2.y, v1.z, 0, 1, 0, 0, tv, TextureIDTop)
    Call Add_Vertex(v2.x, v2.y, v2.z, 0, 1, 0, tu, 0, TextureIDTop)

    ' Create vertices describing the Bottom face of the cube.
    Call Add_Vertex(v2.x, v1.y, v2.z, 0, -1, 0, tu, 0, TextureIDBottom)
    Call Add_Vertex(v1.x, v1.y, v1.z, 0, -1, 0, 0, tv, TextureIDBottom)
    Call Add_Vertex(v2.x, v1.y, v1.z, 0, -1, 0, tu, tv, TextureIDBottom)

    Call Add_Vertex(v1.x, v1.y, v1.z, 0, -1, 0, 0, tv, TextureIDBottom)
    Call Add_Vertex(v2.x, v1.y, v2.z, 0, -1, 0, tu, 0, TextureIDBottom)
    Call Add_Vertex(v1.x, v1.y, v2.z, 0, -1, 0, 0, 0, TextureIDBottom)

End Sub

Sub Add_Cilynder(Center As D3DVECTOR, radius, Height, Optional Segment = 10, Optional TextureIDBody As Long)

  Dim DeltaSegAngle
  Dim SegmentLength, CurrentSegMent
  Dim i, X0, Z0, tu0 As Single, tv0 As Single

    DeltaSegAngle = (2# * PI / Segment)
    SegmentLength = 1# / Segment

    For CurrentSegMent = 0 To Segment

        X0 = radius * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.x
        Z0 = radius * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.z

        Me.Add_Vertex CSng(X0), Center.y, CSng(Z0), 0, 0, CSng(Z0), 1# - SegmentLength * (CurrentSegMent + 1), 0, TextureIDBody

        X0 = radius * Sin(CurrentSegMent * DeltaSegAngle) + Center.x
        Z0 = radius * Cos(CurrentSegMent * DeltaSegAngle) + Center.z

        Me.Add_Vertex CSng(X0), Center.y + Height, CSng(Z0), 0, 0, CSng(Z0), 1# - SegmentLength * (CurrentSegMent), 1, TextureIDBody

        Me.Add_Vertex CSng(X0), Center.y, CSng(Z0), 0, 0, CSng(Z0), 1# - SegmentLength * (CurrentSegMent), 0, TextureIDBody




        X0 = radius * Sin(CurrentSegMent * DeltaSegAngle) + Center.x
        Z0 = radius * Cos(CurrentSegMent * DeltaSegAngle) + Center.z

        Me.Add_Vertex CSng(X0), Center.y + Height, CSng(Z0), 0, 0, CSng(Z0), 1# - SegmentLength * (CurrentSegMent), 1, TextureIDBody
        



        X0 = radius * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.x
        Z0 = radius * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.z

        Me.Add_Vertex CSng(X0), Center.y, CSng(Z0), 0, 0, CSng(Z0), 1# - SegmentLength * (CurrentSegMent + 1), 0, TextureIDBody


        X0 = radius * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.x
        Z0 = radius * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.z

        Me.Add_Vertex CSng(X0), Center.y + Height, CSng(Z0), 0, 0, CSng(Z0), 1# - SegmentLength * (CurrentSegMent + 1), 1, TextureIDBody
        


     
    Next CurrentSegMent
    
    
    
    
    'add top
    
     For CurrentSegMent = 0 To Segment
      



        X0 = radius * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.x
        Z0 = radius * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.z

        
         tu0 = (0.5 * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle)) + 0.5
         tv0 = (0.5 * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle)) + 0.5
        
        Me.Add_Vertex CSng(X0), Center.y + Height, CSng(Z0), 0, 1, 0, tu0, tv0, TextureIDBody


      
      

        Me.Add_Vertex Center.x, Center.y + Height, Center.z, 0, 1, tu0, tv0, 0, TextureIDBody



        X0 = radius * Sin(CurrentSegMent * DeltaSegAngle) + Center.x
        Z0 = radius * Cos(CurrentSegMent * DeltaSegAngle) + Center.z

        
         tu0 = (0.5 * Sin(CurrentSegMent * DeltaSegAngle)) + 0.5
         tv0 = (0.5 * Cos(CurrentSegMent * DeltaSegAngle)) + 0.5
        
        Me.Add_Vertex CSng(X0), Center.y + Height, CSng(Z0), 0, 1, 0, tu0, tv0, TextureIDBody


      
    Next CurrentSegMent
    
    
    
    
    
     'add top bottom
    
     For CurrentSegMent = 0 To Segment
      



    





        X0 = radius * Sin(CurrentSegMent * DeltaSegAngle) + Center.x
        Z0 = radius * Cos(CurrentSegMent * DeltaSegAngle) + Center.z

        
         tu0 = (0.5 * Sin(CurrentSegMent * DeltaSegAngle)) + 0.5
         tv0 = (0.5 * Cos(CurrentSegMent * DeltaSegAngle)) + 0.5
        
        Me.Add_Vertex CSng(X0), Center.y, CSng(Z0), 0, 1, 0, tu0, tv0, TextureIDBody



        Me.Add_Vertex Center.x, Center.y, Center.z, 0, 1, tu0, tv0, 0, TextureIDBody


        X0 = radius * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.x
        Z0 = radius * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.z

        
         tu0 = (0.5 * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle)) + 0.5
         tv0 = (0.5 * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle)) + 0.5
        
        Me.Add_Vertex CSng(X0), Center.y, CSng(Z0), 0, 1, 0, tu0, tv0, TextureIDBody

      
    Next CurrentSegMent
    
    

End Sub







Sub Add_Cone(Center As D3DVECTOR, radius, Height, Optional Segment = 10, Optional TextureIDBody As Long)



 Dim DeltaSegAngle
  Dim SegmentLength, CurrentSegMent
  Dim i, X0, Z0, tu0 As Single, tv0 As Single, ny0 As Single

    DeltaSegAngle = (2# * PI / Segment)
    SegmentLength = 1# / Segment


    ny0 = (PI / 2 - farcTan(Height / radius)) / PI / 2

 
     For CurrentSegMent = 0 To Segment
      




        X0 = radius * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.x
        Z0 = radius * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.z

        
         tu0 = 1# - (SegmentLength * CurrentSegMent + CurrentSegMent)
        
         
        Me.Add_Vertex CSng(X0), Center.y, CSng(Z0), 0, ny0, 0, tu0, 1, TextureIDBody
 
 
         Me.Add_Vertex Center.x, Center.y + Height, Center.z, 0, ny0, 0, tu0, 0, TextureIDBody


        X0 = radius * Sin(CurrentSegMent * DeltaSegAngle) + Center.x
        Z0 = radius * Cos(CurrentSegMent * DeltaSegAngle) + Center.z

        
         tu0 = 1# - (SegmentLength * CurrentSegMent)
         
         
        Me.Add_Vertex CSng(X0), Center.y, CSng(Z0), 0, ny0, 0, tu0, 1, TextureIDBody
 
        
    Next CurrentSegMent




End Sub






'----------------------------------------
'Name: Render
'----------------------------------------
Sub Render()

  Dim i

    D3dDevice.SetVertexShader NEMO_CUSTOM_VERTEX

    For i = 0 To mMESH.NumText - 1
        With mMESH
            D3dDevice.SetTexture 0, POOL_texture(.TexturesID(i))
            D3dDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, Int(.POLY_sorted(i).NumVERT / 3), .POLY_sorted(i).VERT(0), Len(.POLY_sorted(i).VERT(0))
        End With 'MMESH
    Next i

End Sub

'----------------------------------------
'Name: Render2
'----------------------------------------
Sub Render2()

  Dim i

    D3dDevice.SetVertexShader NEMO_CUSTOM_VERTEX

    'For I = 0 To mMESH.NumText - 1
    With mMESH

        RenderNemoPoly .POLYcoldet(0), PolyCount + 1, lpFRUST, 1, AddressOf RenderFastPoly
    End With 'MMESH
    'Next I

End Sub

'----------------------------------------
'Name: CheckCollision
'----------------------------------------
Function CheckCollision(PointVec As D3DVECTOR, Optional radius As Single = 30) As Boolean

  Dim B As Boolean
  Dim PAK As NemoColResult

    PAK.eRadius.y = radius
    PAK.eRadius.x = radius
    PAK.eRadius.z = radius

    CopyMemory PAK.sourcePoint, PointVec, Len(PointVec)
  Dim TempV As NemoVECTOR ':( Move line to top of current Function
    Call NemoPOLYGON_CheckForCollision(TempV, PAK, mMESH.POLYcoldet(0), PolyCount + 1)

    CheckCollision = PAK.foundCollision

    'Debug.Print PAK.foundCollision

End Function

Function CheckCollisionSliding(PointVec As D3DVECTOR, RetPos As D3DVECTOR, Optional radius As Single = 30) As Boolean

  Dim B As Boolean
  Dim PAK As NemoColResult

    PAK.eRadius.y = radius
    PAK.eRadius.x = radius
    PAK.eRadius.z = radius

    CopyMemory PAK.sourcePoint, PointVec, Len(PointVec)
  Dim TempV As NemoVECTOR ':( Move line to top of current Function
    Call NemoPOLYGON_CheckForCollision(TempV, PAK, mMESH.POLYcoldet(0), PolyCount + 1)

    CopyMemory RetPos, TempV, Len(TempV)

    CheckCollisionSliding = PAK.foundCollision

    'Debug.Print PAK.foundCollision

End Function

'----------------------------------------
'Name: Class_Terminate
'Object: Class
'Event: Terminate
'----------------------------------------
Private Sub Class_Terminate()

    If mMESH.IS_Loaded Then
        'Erase mMESH.Textures
        Erase mMESH.Vertex
        'Erase mMESH.Vertex2
    End If

End Sub





