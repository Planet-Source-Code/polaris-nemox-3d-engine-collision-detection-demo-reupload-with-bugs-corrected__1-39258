VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nemo Engine Enumeration Dialogue"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Engine Extra Option"
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   2040
      TabIndex        =   15
      Top             =   2520
      Width           =   3615
      Begin VB.OptionButton Option4 
         BackColor       =   &H00000000&
         Caption         =   "Tripple Buffering"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00000000&
         Caption         =   "Double Buffering"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Vertical Sync Disabled"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Rendering Mode"
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   1695
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Windowed"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "FullScreen"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000008&
      Caption         =   "Video parametters"
      ForeColor       =   &H80000005&
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox NemoCmbRes 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   960
         Width           =   3255
      End
      Begin VB.ComboBox NemoCmbDevice 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1560
         Width           =   3255
      End
      Begin VB.ComboBox NemoCmbAdapters 
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000007&
         Caption         =   "Acceleration "
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Screen Resolution"
         ForeColor       =   &H00FFFF80&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Caption         =   "Display Adaptators"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Gamma Level"
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   5535
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4080
         TabIndex        =   2
         Text            =   "1.00"
         Top             =   240
         Width           =   615
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Max             =   100
         TabIndex        =   1
         Top             =   240
         Value           =   50
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'   ============================================================
'    ----------------------------------------------------------
'     Application Name:
'     Developer/Programmer:
'    ----------------------------------------------------------
'     Module Name: frmMain
'     Module File: SourcePSC\frmENUM.frm
'     Module Type: Form
'     Module Description:
'    ----------------------------------------------------------
'     © Copyright 2002
'    ----------------------------------------------------------
'   ============================================================


Dim TempDX8 As DirectX8          'The Root Object
Dim TempD3D8 As Direct3D8      'The Direct3D Interface

Dim nAdapters As Long 'How many adapters we found
Dim AdapterInfo As D3DADAPTER_IDENTIFIER8 'A Structure holding information on the adapter

Dim nModes As Long 'How many display modes we found
Dim CFG As NEMO_CFG_INI
Dim LPnemo As NemoX







'----------------------------------------
'Name: EnumerateAdapters
'----------------------------------------
Private Sub EnumerateAdapters()
    
    Dim i As Integer, sTemp As String, j As Integer
    
    '//This'll either be 1 or 2
    nAdapters = TempD3D8.GetAdapterCount
    
    For i = 0 To nAdapters - 1
        'Get the relevent Details
        TempD3D8.GetAdapterIdentifier i, 0, AdapterInfo
        
        'Get the name of the current adapter - it's stored as a long
        'list of character codes that we need to parse into a string
        ' - Dont ask me why they did it like this; seems silly really :)
        sTemp = "" 'Reset the string ready for our use
        For j = 0 To 511
            sTemp = sTemp & Chr$(AdapterInfo.Description(j))
        Next j
        sTemp = Replace(sTemp, Chr$(0), " ")
        NemoCmbAdapters.AddItem sTemp
    Next i
End Sub


'----------------------------------------
'Name: EnumerateDevices
'----------------------------------------
Private Sub EnumerateDevices()
On Local Error Resume Next '//We want to handle the errors...
Dim CAPS As D3DCAPS8

TempD3D8.GetDeviceCaps NemoCmbAdapters.ListIndex, D3DDEVTYPE_HAL, CAPS
    If Err.Number = D3DERR_NOTAVAILABLE Then
        'There is no hardware acceleration
        NemoCmbDevice.AddItem "Reference Rasterizer (REF)" 'Reference device will always be available
    Else
        NemoCmbDevice.AddItem "Hardware Acceleration (HAL)"
        NemoCmbDevice.AddItem "Reference Rasterizer (REF)" 'Reference device will always be available
    End If
    
End Sub


'----------------------------------------
'Name: EnumerateDispModes
'----------------------------------------
Private Sub EnumerateDispModes(Renderer As Long)
NemoCmbRes.Clear '//Remove any existing entries...

Dim i As Integer, ModeTemp As D3DDISPLAYMODE

nModes = TempD3D8.GetAdapterModeCount(NemoCmbAdapters.ListIndex)

For i = 0 To nModes - 1 '//Cycle through them and collect the data...
    Call TempD3D8.EnumAdapterModes(NemoCmbAdapters.ListIndex, i, ModeTemp)
    
    'First we parse the modes into two catergories - 16bit and 32bit
    If ModeTemp.format = D3DFMT_R8G8B8 Or ModeTemp.format = D3DFMT_X8R8G8B8 Or ModeTemp.format = D3DFMT_A8R8G8B8 Then
        'Check that the device is acceptable and valid...
        If TempD3D8.CheckDeviceType(NemoCmbAdapters.ListIndex, Renderer, ModeTemp.format, ModeTemp.format, False) >= 0 Then
            'then add it to the displayed list
            NemoCmbRes.AddItem ModeTemp.width & "x" & ModeTemp.Height & " 32 bit" '& "    [FMT: " & ModeTemp.Format & "]"
        End If
    Else
        If TempD3D8.CheckDeviceType(NemoCmbAdapters.ListIndex, Renderer, ModeTemp.format, ModeTemp.format, False) >= 0 Then
            NemoCmbRes.AddItem ModeTemp.width & "x" & ModeTemp.Height & " 16 bit" '& "    [FMT: " & ModeTemp.Format & "]"
        End If
    End If
    
Next i

NemoCmbRes.ListIndex = NemoCmbRes.ListCount - 1
End Sub


'----------------------------------------
'Name: EnumerateHardware
'----------------------------------------
Private Sub EnumerateHardware(Renderer As Long)
'Renderer = Renderer + 1 '//We need it on a base 1 scale (not base 0)
'List1.Clear '//Clear our list

Dim CAPS As D3DCAPS8 '//Holds all our information...
'
TempD3D8.GetDeviceCaps NemoCmbAdapters.ListIndex, Renderer, CAPS

If CAPS.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT Then

    If NemoCmbDevice.ListCount < 3 Then _
    NemoCmbDevice.AddItem "Hardware Transform and lighting (TnL)"

End If

'
If CAPS.Caps2 And D3DCAPS2_FULLSCREENGAMMA Then
     Frame1.Enabled = True
End If


End Sub


'----------------------------------------
'Name: Check1_Click
'Object: Check1
'Event: Click
'----------------------------------------
Private Sub Check1_Click()
 
 CFG.ForceVerSINC = Check1.value
 
 
End Sub


'----------------------------------------
'Name: NemoCmbAdapters_Click
'Object: NemoCmbAdapters
'Event: Click
'----------------------------------------
Private Sub NemoCmbAdapters_Click()
If UCase(Left(NemoCmbDevice.Text, 3)) = "REF" Then
    EnumerateDispModes 2
Else
    EnumerateDispModes 1
End If
End Sub


'----------------------------------------
'Name: NemoCmbDevice_Click
'Object: NemoCmbDevice
'Event: Click
'----------------------------------------
Private Sub NemoCmbDevice_Click()


CFG.DeviceTyp = NemoCmbDevice.ListIndex
If NemoCmbDevice.ListIndex = 2 Then CFG.USE_TnL = True

End Sub



'----------------------------------------
'Name: Init
'----------------------------------------
Private Sub Init()
With CFG
  CFG.BPP = 16
  .DeviceTyp = NEMO_HAL_DEVICE
  .ForceVerSINC = True
  .GamaLevel = 1
  .BufferCount = 1

End With

Text1.Text = Str(CFG.GamaLevel)
Option1.value = 1
NemoCmbRes.Enabled = 1
Check1.value = 1

End Sub





'----------------------------------------
'Name: NemoCmbRes_Click
'Object: NemoCmbRes
'Event: Click
'----------------------------------------
Private Sub NemoCmbRes_Click()
Dim v1, v2, v3, SS As String

SS = NemoCmbRes.List(NemoCmbRes.ListIndex)

v1 = InStr(SS, "x")
CFG.width = Val(Left(SS, v1 - 1))

v2 = InStr(v1, SS, " ")
CFG.Height = Val(Mid(SS, v1 + 1, v2 - v1))


CFG.BPP = Val(Left(Right(SS, 6), 2))



End Sub


'----------------------------------------
'Name: Command1_Click
'Object: Command1
'Event: Click
'----------------------------------------
Private Sub Command1_Click()
  End
End Sub



'----------------------------------------
'Name: SHOW_DIALOG
'Object: SHOW
'Event: DIALOG
'----------------------------------------
Sub SHOW_DIALOG(LpHandle As Long, lpNemoEngine As NemoX)
  Load Me
  CFG.appHandle = LpHandle
  Set LPnemo = lpNemoEngine
  Me.Show vbModal
  
  

End Sub


'----------------------------------------
'Name: Command2_Click
'Object: Command2
'Event: Click
'----------------------------------------
Private Sub Command2_Click()
 Me.Hide
 LPnemo.Initialize CFG.appHandle, CFG.DeviceTyp, Not (CFG.IS_FullScreen), CFG.width, CFG.Height, 16, CFG.USE_TnL, CFG.ForceVerSINC, CFG.GamaLevel, CFG.BufferCount
 Unload Me

End Sub


'----------------------------------------
'Name: Form_Load
'Object: Form
'Event: Load
'----------------------------------------
Private Sub Form_Load()

Dim N As New NemoX
Me.Caption = "Nemo Engine Version" + Str(N.Get_EngineVersion) + " Enumeration Dialogue"
Set N = Nothing
Init
Me.Frame1.Enabled = 0
'//1. Create any relevent objects
Set TempDX8 = New DirectX8
Set TempD3D8 = TempDX8.Direct3DCreate

'//2. Collect the data
EnumerateAdapters 'Information on physical hardware cards available
    NemoCmbAdapters.ListIndex = 0 'Set it to the first entry in the list
    
EnumerateDevices 'what rendering devices they support
    NemoCmbDevice.ListIndex = 0 'Set it to the first entry
    
If UCase(Left(NemoCmbDevice.Text, 3)) = "REF" Then
    EnumerateDispModes 2
Else
    EnumerateDispModes 1
End If
    NemoCmbRes.ListIndex = NemoCmbRes.ListCount - 1 'Set it to the first entry
    
If UCase(Left(NemoCmbDevice.Text, 3)) = "REF" Then
    EnumerateHardware 2 'Reference device
Else
    EnumerateHardware 1 'hardware device
End If

If NemoCmbRes.ListCount > 6 Then NemoCmbRes.ListIndex = 6
End Sub


'----------------------------------------
'Name: HScroll1_Change
'Object: HScroll1
'Event: Change
'----------------------------------------
Private Sub HScroll1_Change()
 CFG.GamaLevel = (HScroll1 / 100) * 2
 Text1.Text = Str(CFG.GamaLevel)

End Sub


'----------------------------------------
'Name: Option1_Click
'Object: Option1
'Event: Click
'----------------------------------------
Private Sub Option1_Click()
   Set_Fmode
End Sub


'----------------------------------------
'Name: Option2_Click
'Object: Option2
'Event: Click
'----------------------------------------
Private Sub Option2_Click()
   Set_Fmode
End Sub


'----------------------------------------
'Name: Set_Fmode
'Object: Set
'Event: Fmode
'----------------------------------------
Sub Set_Fmode()
  CFG.IS_FullScreen = Option1.value
  Check1.Enabled = Int(Option1.value)
  Frame1.Enabled = Option1.value
  HScroll1.Enabled = Option1.value
  Text1.Enabled = Option1.value
  NemoCmbRes.Enabled = Option1.value
  
End Sub


'----------------------------------------
'Name: Option3_Click
'Object: Option3
'Event: Click
'----------------------------------------
Private Sub Option3_Click()
 If Option3.value = 1 Then CFG.BufferCount = 2
 
End Sub


'----------------------------------------
'Name: Option4_Click
'Object: Option4
'Event: Click
'----------------------------------------
Private Sub Option4_Click()
 If Option4.value = 1 Then CFG.BufferCount = 3

End Sub



