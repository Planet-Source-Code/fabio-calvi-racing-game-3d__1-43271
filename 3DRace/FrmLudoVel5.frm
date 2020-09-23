VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   72
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   318
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   428
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   0
      ScaleHeight     =   457
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin VB.Timer Tmr_time 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   240
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "LOADING..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   2040
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const pi = 3.14159265358979

Private WithEvents RM_E As BDirectx
Attribute RM_E.VB_VarHelpID = -1

Dim RMC As New BDirectx
Dim running As Boolean

Dim Sfondo As Direct3DRMTexture3

Dim scene As Direct3DRMFrame3
Dim cam As Direct3DRMFrame3
Dim mesh As Direct3DRMMeshBuilder3
Dim wallpos As D3DVECTOR
Dim LightFrame As Direct3DRMFrame3
Dim ViewFrame() As Direct3DRMFrame3
Dim Wall(111) As Direct3DRMMeshBuilder3
Dim m_objectFrame(11) As Direct3DRMFrame3
Dim m_meshBuilder(11) As Direct3DRMMeshBuilder3
Dim ombreFrame As Direct3DRMFrame3
Dim ombreMesh As Direct3DRMMeshBuilder3

Dim BonusFrame() As Direct3DRMFrame3
Dim BonusMesh() As Direct3DRMMeshBuilder3

Dim XFileTex As Direct3DRMTexture3

Dim Dir As D3DVECTOR
Dim up As D3DVECTOR
Dim pos As D3DVECTOR

Dim Playersnd As DirectSoundBuffer
Dim Playersnd3D As DirectSound3DBuffer
Dim Driversnd(10) As DirectSoundBuffer
Dim Driversnd3D(10) As DirectSound3DBuffer
Dim Boingsnd As DirectSoundBuffer
Dim Boingsnd3D As DirectSound3DBuffer
Dim Beepsnd As DirectSoundBuffer
Dim Beepsnd3D As DirectSound3DBuffer
Dim Hit1snd As DirectSoundBuffer
Dim Hit1snd3D As DirectSound3DBuffer
Dim Hit2snd As DirectSoundBuffer
Dim Hit2snd3D As DirectSound3DBuffer
Dim Bonussnd As DirectSoundBuffer
Dim Bonussnd3D As DirectSound3DBuffer

Dim D3Pos As D3DVECTOR  ' Holds position of player
Dim D3Ori As D3DVECTOR  ' Holds orientation of player
Dim D3Nor As D3DVECTOR  ' Holds normal of player
Dim Heading As Single      ' Current heading in radians
Dim I_nBanking As Single
Dim Velocity As Single     ' Current velocity
Dim MaxVel As Single       ' Velocity max

Dim TCase As Integer
Dim collide As Boolean, distance As Single
Dim SystemFrame As Single, old As Single, oldtime As Single
Dim vel As Single, nvel As Integer, bp, lv As Single
Dim avanti As Boolean, indietro As Boolean, destra As Boolean, sinistra As Boolean, fine As Boolean, retro As Boolean

Dim xx As Integer, yy As Integer
Dim iFree As Integer
Dim x As Integer
Dim y As Integer
Dim MapNumber As Integer
Dim MapSizeX As Integer
Dim MapSizeY As Integer
Dim StartX As Integer
Dim StartY As Integer
Dim tile() As String
Dim map() As Integer

Dim DINPUT As DirectInput
Dim DIdevice As DirectInputDevice
Dim mat As Direct3DRMMaterial2
Dim keyb As DIKEYBOARDSTATE

Dim dds As DDSURFACEDESC2

Dim t1 As Long, fogcolor As Single
Dim Starttick As Long, LastTick As Long
Dim D3Ray As D3DRMRAY                ' Ray for picking
Dim L_oD3PDA As Direct3DRMPick2Array ' Result of picking
Dim D3PD As D3DRMPICKDESC2           ' Result of picking
Dim L_oD3Visual As Direct3DRMVisual  ' Visual picked
Dim L_bColliding As Boolean
Dim D3Tmp As D3DVECTOR

Dim oview As Single, hview As Single, dview As Single
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Type RacerType
    Name        As String
    Location_X  As Single
    Location_Z  As Single
    Direction   As Single
    Speed       As Single
    MaxSpeed    As Single
    MinSpeed    As Single
    TurnRatio   As Single   'We rotate this much every tick
    Acceleration As Single
    BrakeSpeed  As Single
    BackAngle   As Single   'Used in speed calculations, see tutorial for proper explanation
    Radius      As Integer  'The size of the car
    NextNode    As Integer  'The number of the node we're heading for
    Lap         As Integer
    Time        As String
    Position    As Integer
End Type

Private Type NodeType
    Location_X  As Single
    Location_Y  As Single
    Dir As String
End Type

Private Type TargetType
    Location_X  As Single
    Location_Z  As Single
    Pass As Boolean
    Size As Integer
End Type

Private Type bonusType
    Location_X  As Single
    Location_Z  As Single
    Active As Boolean
    sec As Integer
End Type

Private Racers(10) As RacerType   'Opponent racers
Private DummyRacers(10) As RacerType
Private Racerspos(10) As Byte

Private Nodes() As NodeType      'An array of nodes marking our route
Private Targets() As TargetType
Private Bonus() As bonusType

Private Const NodeRadius = 15    'Size of the nodes
            
Private CircleX As Single        'They would be local vars except
Private CircleY As Single
Private CircleRadius As Single

Dim ArrNum15() As Integer, ret As Integer
Dim TargetsCount As Integer, Lap As Integer
Dim Numdown As Integer, MyFont As New StdFont, cont As Byte
Dim Timemin As Single, Timesec As Single, Besttime As String
Dim F1 As Boolean, F2 As Boolean, F3 As Boolean, F4 As Boolean, F5 As Boolean, F6 As Boolean
Dim F7 As Boolean, F8 As Boolean, F9 As Boolean, F10 As Boolean, F11 As Boolean, F12 As Boolean

Dim Lastnodes As Boolean

'Main sub
Public Sub form_load()
Dim ind As Integer

MusicVolume = 75 'Max(100%) -> MusicVolume = 100
                 'High      -> MusicVolume = 75
                 'Medium    -> MusicVolume = 50
                 'Low       -> MusicVolume = 25
                 'Null(0%)  -> MusicVolume = 0
'Set music
Music_mod.Initialize_Music
Music_mod.Load_Music (0)
Music_mod.SetMusic (MusicVolume)

Me.Show

Pic.ScaleWidth = frmMain.ScaleWidth
Pic.ScaleHeight = frmMain.ScaleHeight
Label1.left = (frmMain.ScaleWidth - Label1.Width) / 2
Label1.top = (frmMain.ScaleHeight - Label1.Height) / 2

DoEvents

    ShowCursor 0
    TCase = 60
    Set RM_E = RMC
    RMC.hwnd = Pic.hwnd
    RMC.UseBackBuffer = True
    RMC.Use3DHardware = True
    RMC.StartWindowed
    InitDeviceobjects
    InitSounds
    
    Set DINPUT = RMC.dx.DirectInputCreate
    Set DIdevice = DINPUT.CreateDevice("GUID_SysKeyboard")
    DIdevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DIdevice.SetCooperativeLevel Me.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
     
    ' Create rotating background
    dds.lFlags = DDSD_WIDTH Or DDSD_HEIGHT
    dds.lHeight = RMC.Viewport.GetHeight  '/ 2  + 12
    dds.lWidth = RMC.Viewport.GetWidth * 5 / 2
    'I don't have downloaded the original background (too big)
    Set Background = RMC.DDraw.CreateSurfaceFromFile(App.path & "\back.bmp", dds)
  
    Initwalls
    OpenMap
    CreatePlayers
    
    For ind = 1 To UBound(Racers)
        RandNum
    Next
    InitRacers
    
    doTargets
    
    nvel = 0
    Angle = 0
    Heading = 0
    Velocity = 0
    MaxVel = 6
    hview = 5
    dview = -25
    cont = 1
    Lastnodes = False
    
    F1 = False
    F2 = False
    F3 = False
    F4 = False
    F5 = False
    F6 = False
    F7 = False
    F8 = False
    F9 = False
    F10 = False

    running = True
    Numdown = 0
    Lap = 1
    
    Timemin = 0
    Timesec = 0
    
    dview = -50
    
    Music_mod.PlayMusic
    
    ' main loop
    Do While running = True
       SystemFrame = RMC.dx.TickCount - oldtime
       oldtime = RMC.dx.TickCount
       If oldtime > old + 1000 Then
          old = oldtime
          If Tmr_time.Enabled = False Then Numdown = Numdown + 1
          For ind = 1 To UBound(Bonus)
              If Bonus(ind).Active = False Then
                 Bonus(ind).sec = Bonus(ind).sec + 1
                 If Bonus(ind).sec > 5 Then
                    Bonus(ind).sec = 0
                    Bonus(ind).Active = True
                    If BonusFrame(ind).GetVisualCount = 0 Then BonusFrame(ind).AddVisual BonusMesh(ind)
                 End If
              End If
          Next
       End If
       If Lap = 1 And dview < -25 Then dview = dview + 0.5
       If Numdown <= 4 Then '3,2,1,GO!
          m_objectFrame(1).SetPosition Nothing, StartX * TCase, 2, StartY * TCase
          ombreFrame.SetPosition m_objectFrame(1), 0, -1.95, 0
          RMC.CameraFrame.SetPosition m_objectFrame(1), 0, hview, dview
          RMC.CameraFrame.LookAt m_objectFrame(1), Nothing, D3DRMCONSTRAIN_Z
          For ind = 1 To UBound(Racers)
              m_objectFrame(ind + 1).SetPosition Nothing, Racers(ind).Location_X, 1, Racers(ind).Location_Z
          Next
          If Numdown = 4 Then Tmr_time.Enabled = True
       Else 'let's move the racers
          If Lap < 4 Then
             MovePlayer
          Else
             Finish
          End If
          For ind = 1 To UBound(Racers)
              If Racers(ind).Lap < 4 Then
                 RacersAI (ind)
              Else
                 If Racers(ind).Speed > 0 Then Racers(ind).Speed = Racers(ind).Speed - 1
                 If Racers(ind).Speed < 0 Then Racers(ind).Speed = 0
              End If
              MoveRacers (ind)
          Next ind
       End If
       'Render
       RMC.Render
       'Loop Music
       Call Music_mod.LoopMusic
       
       If fine = True Then Cleanup
       
       DoEvents
    
    Loop
End Sub
Private Sub MovePlayer()
Dim i As Integer, distance As Integer
    
'       DIdevice.Acquire
'       DIdevice.GetDeviceStateKeyboard keyb
       
       m_objectFrame(1).GetPosition Nothing, D3Pos
       m_objectFrame(1).GetOrientation Nothing, D3Ori, D3Nor
        
       nvel = 0
       
       If avanti = False And indietro = False Then
          'Velocity = Velocity - 0.05
           If Velocity > 0 Then Velocity = Velocity - 0.1
           If Velocity < 0 Then Velocity = Velocity + 0.1
           MaxVel = 6
       End If
        
        'Move forward
        If avanti = True Then
           Velocity = Velocity + 0.5
           nvel = 1
'           Playersnd.Play DSBPLAY_DEFAULT
        End If
        
        'Move back
        If indietro = True Then
           Velocity = Velocity - 0.5
           MaxVel = 6
           nvel = -1
'           Playersnd.Play DSBPLAY_DEFAULT
        End If
        
        'Rotate left
        If sinistra = True Then
           I_nBanking = I_nBanking + 0.01
        End If

        'Rotate right
        If destra = True Then
           I_nBanking = I_nBanking - 0.01
        End If
        
        Heading = Heading + I_nBanking * 0.5
        ' Banking dekeys
        I_nBanking = I_nBanking * 0.95
     
        ' Reset colission
        L_bColliding = False
                    
        ' Prepare ray
        RMC.dx.VectorCopy D3Ray.pos, D3Pos
        D3Ray.Dir.x = D3Ori.x
        D3Ray.Dir.y = 0
        D3Ray.Dir.z = D3Ori.z
        If dview > -25 Then dview = dview - 1
        
        For xx = 0 To MapSizeX - 1
           For yy = 0 To MapSizeY - 1
                    
                    ' Cast ray
                    Set L_oD3PDA = ViewFrame(xx, yy).RayPick(Nothing, D3Ray, D3DRMRAYPICK_IGNOREFURTHERPRIMITIVES)
                    ' Retrieve results
                    If Not (L_oD3PDA.GetSize = 0) Then
                       If L_oD3PDA.GetPickFrame(0, D3PD).GetSize > 0 Then
                          Set L_oD3Visual = L_oD3PDA.GetPickVisual(0, D3PD)
                          RMC.dx.VectorSubtract D3Tmp, D3Ray.pos, D3PD.vPostion
                          L_bColliding = (RMC.dx.VectorModulus(D3Tmp) <= 15)
                          If L_bColliding Then
                             Boingsnd3D.SetPosition D3Pos.x, D3Pos.y, D3Pos.z, DS3D_IMMEDIATE
                             Boingsnd.Play DSBPLAY_DEFAULT
                             Velocity = -Velocity '/ 2
                             MaxVel = 6
                             If dview < -20 Then dview = dview + 2 Else dview = -20
                          End If
                       End If
                    End If
           
           Next
        Next
       
        If Velocity > MaxVel Then Velocity = MaxVel
        If Velocity < -5 Then Velocity = -5
        
        D3Pos.x = D3Pos.x + Cos(Heading) * Velocity
        D3Pos.y = D3Pos.y
        D3Pos.z = D3Pos.z + Sin(Heading) * Velocity
                                
        ' Calculate Orientation
        D3Ori.x = Cos(Heading)
        D3Ori.y = 0
        D3Ori.z = Sin(Heading)
                    
        ' Calculate Normal
        D3Nor.x = I_nBanking * 2 * -Sin(Heading)
        D3Nor.y = 1
        D3Nor.z = I_nBanking * 2 * Cos(Heading)
                    
        ' Normalize data
        RMC.dx.VectorNormalize D3Ori
        RMC.dx.VectorNormalize D3Nor
        
        ' Check for collision with other drivers
        For i = 1 To UBound(Racers)
            distance = GetDistance(D3Pos.x, D3Pos.z, Racers(i).Location_X, Racers(i).Location_Z)
            If distance < 20 Then
               Driversnd3D(i).SetPosition Racers(i).Location_X, 1, Racers(i).Location_Z, DS3D_IMMEDIATE
               Driversnd(i).Play DSBPLAY_DEFAULT
            End If
            If distance = 10 Then
               Beepsnd3D.SetPosition Racers(i).Location_X, 1, Racers(i).Location_Z, DS3D_IMMEDIATE
               Beepsnd.Play DSBPLAY_DEFAULT
            End If
            If distance < 2.5 + Racers(i).Radius Then
               Hit1snd3D.SetPosition D3Pos.x, D3Pos.y, D3Pos.z, DS3D_IMMEDIATE
               Hit1snd.Play DSBPLAY_DEFAULT
               Hit2snd3D.SetPosition Racers(i).Location_X, 1, Racers(i).Location_Z, DS3D_IMMEDIATE
               Hit2snd.Play DSBPLAY_DEFAULT
               Velocity = -Velocity
               Racers(i).Speed = -Racers(i).Speed / 2
            End If
        Next
        
        For i = 1 To UBound(Bonus)
            distance = GetDistance(D3Pos.x, D3Pos.z, Bonus(i).Location_X, Bonus(i).Location_Z)
            If distance < 5 And Bonus(i).Active = True Then
               'You can drop down this limit below and increase an unlimit(well, integer limit) turbo
               'and try to run at max speed as you can
               If MaxVel < 10 Then MaxVel = MaxVel + 1
               Bonussnd3D.SetPosition D3Pos.x, D3Pos.y, D3Pos.z, DS3D_IMMEDIATE
               Bonussnd.Play DSBPLAY_DEFAULT
               Bonus(i).Active = False
               If BonusFrame(i).GetVisualCount > 0 Then BonusFrame(i).DeleteVisual BonusMesh(i)
            End If
            If Bonus(i).Active = True Then BonusFrame(i).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 5 * (pi / 180)
        Next
                
        For i = 1 To UBound(Targets)
            distance = GetDistance(D3Pos.x, D3Pos.z, Targets(i).Location_X, Targets(i).Location_Z)
            If Targets(i).Pass = False And distance < 60 Then
               Targets(i).Pass = True
               If i = 1 Then
                  Targets(9).Pass = False
                  TargetsCount = TargetsCount + 1
               End If
               TargetsCount = TargetsCount - 1
            End If
        Next
        
        If TargetsCount = 0 Then
           Lap = Lap + 1
           doTargets
        End If
        
        'Set values ...
        m_objectFrame(1).SetPosition Nothing, D3Pos.x, D3Pos.y, D3Pos.z
        m_objectFrame(1).SetOrientation Nothing, D3Ori.x, D3Ori.y, D3Ori.z, D3Nor.x, D3Nor.y, D3Nor.z
        m_objectFrame(1).GetPosition Nothing, D3Pos
        Playersnd3D.SetPosition D3Pos.x, D3Pos.y, D3Pos.z, DS3D_IMMEDIATE
        ombreFrame.SetPosition m_objectFrame(1), 0, -1.95, 0
        If retro = True Then 'back view
           RMC.CameraFrame.SetPosition m_objectFrame(1), 0, 2, 0
           If ret = 1 Then
              RMC.CameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 180 * pi / 180
              ret = 2
           End If
        Else 'front view
           RMC.CameraFrame.SetPosition m_objectFrame(1), 0, hview, dview
           RMC.CameraFrame.LookAt m_objectFrame(1), Nothing, D3DRMCONSTRAIN_Z
        End If
       
       'These fixed camera views are good for this map only
       If F1 = True Then
          RMC.CameraFrame.SetPosition Nothing, 16 * TCase, 15, 5 * TCase
          RMC.CameraFrame.SetOrientation Nothing, 0.5, 0, -0.5, 0, 1, 0
       End If
       If F2 = True Then
          RMC.CameraFrame.SetPosition Nothing, 23.35 * TCase, 8, 0.45 * TCase
          RMC.CameraFrame.SetOrientation Nothing, -0.5, 0, 1.2, 0, 1, 0
       End If
       If F3 = True Then
          RMC.CameraFrame.SetPosition Nothing, 23 * TCase, 15, 15 * TCase
          RMC.CameraFrame.SetOrientation Nothing, 0.5, 0, -0.5, 0, 1, 0
       End If
       If F4 = True Then
          RMC.CameraFrame.SetPosition Nothing, 45 * TCase, 25, 4.5 * TCase
          RMC.CameraFrame.SetOrientation Nothing, -0.65, 0, 0.35, 0, 1, 0
       End If
       If F5 = True Then
          RMC.CameraFrame.SetPosition Nothing, 50 * TCase, 5, 8.7 * TCase
          RMC.CameraFrame.SetOrientation Nothing, -0.65, 0, 0.35, 0, 1, 0
       End If
       If F6 = True Then
          RMC.CameraFrame.SetPosition Nothing, 47 * TCase, 20, 30 * TCase
          RMC.CameraFrame.SetOrientation Nothing, -1, 0, -1, 0, 1, 0
       End If
       If F7 = True Then
          RMC.CameraFrame.SetPosition Nothing, 41 * TCase, 5, 43 * TCase
          RMC.CameraFrame.SetOrientation Nothing, 0, 0, -1, 0, 1, 0
       End If
       If F8 = True Then
          RMC.CameraFrame.SetPosition Nothing, 29 * TCase, 20, 46 * TCase
          RMC.CameraFrame.SetOrientation Nothing, 1, 0, 0, 0, 1, 0
       End If
       If F9 = True Then
          RMC.CameraFrame.SetPosition Nothing, 18 * TCase, 15, 37 * TCase
          RMC.CameraFrame.SetOrientation Nothing, 0.5, 0, 1, 0, 1, 0
       End If
       If F10 = True Then
          RMC.CameraFrame.SetPosition Nothing, 7 * TCase, 7.5, 37 * TCase
          RMC.CameraFrame.SetOrientation Nothing, 0, 0, 0, 0, 1, 0
       End If
       If F11 = True Then
          RMC.CameraFrame.SetPosition Nothing, 7.5 * TCase, 15, 12 * TCase
          RMC.CameraFrame.SetOrientation Nothing, -0.5, 0, 1, 0, 1, 0
       End If
       If F12 = True Then
          RMC.CameraFrame.SetPosition Nothing, 3 * TCase, 7.5, 2 * TCase
          RMC.CameraFrame.SetOrientation Nothing, 0, 0, 1, 0, 1, 0
       End If
      
        RMC.DsoundLis70.SetPosition D3Pos.x, D3Pos.y, D3Pos.z, DS3D_IMMEDIATE
        RMC.DsoundLis70.SetOrientation D3Ori.x, D3Ori.y, D3Ori.z, D3Nor.x, D3Nor.y, D3Nor.z, DS3D_IMMEDIATE
        
        'Angle = Cos(Heading)
        
        'fix the angle
        If Angle > pi * 2 Then Angle = Angle - pi * 2
        If Angle < 0 Then Angle = Angle + pi * 2
     
End Sub
  

Private Sub Initwalls()
Dim i, appo As String, blocco As String
For i = 0 To 111
    If i < 10 Then
       appo = "00" & CStr(i)
    ElseIf i < 100 Then
       appo = "0" & CStr(i)
    Else
       appo = CStr(i)
    End If
    blocco = App.path & "\Mesh\bloc_" & appo & ".x"
    Set Wall(i) = CaricaX(blocco)
Next i
End Sub
Private Function CaricaX(Nom$) As Direct3DRMMeshBuilder3
Dim i As Byte, appo As String

Set CaricaX = RMC.D3DRM.CreateMeshBuilder()
CaricaX.LoadFromFile Nom$, 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
appo = Mid(Nom$, Len(Nom$) - 9, 8)
'I cannot test the mesh directly: I get an Automation error(I do not know why);
'so I have to test it with this "if" command below:
If appo = "bloc_009" Or appo = "bloc_010" Or appo = "bloc_011" Or appo = "bloc_012" Or appo = "bloc_013" Or _
   appo = "bloc_014" Or appo = "bloc_015" Or appo = "bloc_016" Or appo = "bloc_017" Or appo = "bloc_018" Or _
   appo = "bloc_019" Or appo = "bloc_020" Or appo = "bloc_021" Or appo = "bloc_022" Or appo = "bloc_023" Or _
   appo = "bloc_024" Or appo = "bloc_025" Or appo = "bloc_026" Or appo = "bloc_027" Or appo = "bloc_028" Or _
   appo = "bloc_029" Or appo = "bloc_030" Or appo = "bloc_031" Or appo = "bloc_032" Or appo = "bloc_033" Or _
   appo = "bloc_034" Or appo = "bloc_035" Or appo = "bloc_036" Or appo = "bloc_037" Or appo = "bloc_038" Or _
   appo = "bloc_039" Or appo = "bloc_040" Or appo = "bloc_041" Or appo = "bloc_042" Or appo = "bloc_043" Or _
   appo = "bloc_044" Or appo = "bloc_045" Or appo = "bloc_046" Or appo = "bloc_047" Or appo = "bloc_048" Or _
   appo = "bloc_049" Or appo = "bloc_050" Or appo = "bloc_051" Or appo = "bloc_052" Or appo = "bloc_053" Or _
   appo = "bloc_054" Or appo = "bloc_055" Or appo = "bloc_056" Or appo = "bloc_057" Or appo = "bloc_058" Or _
   appo = "bloc_059" Or appo = "bloc_060" Or appo = "bloc_061" Or appo = "bloc_062" Or appo = "bloc_063" Or _
   appo = "bloc_064" Or appo = "bloc_065" Or appo = "bloc_066" Or appo = "bloc_067" Or appo = "bloc_068" Or _
   appo = "bloc_069" Or appo = "bloc_070" Or appo = "bloc_071" Or appo = "bloc_072" Or appo = "bloc_073" Or _
   appo = "bloc_074" Or appo = "bloc_075" Or appo = "bloc_076" Or appo = "bloc_077" Or appo = "bloc_078" Or _
   appo = "bloc_079" Or appo = "bloc_080" Or appo = "bloc_081" Or appo = "bloc_082" Or appo = "bloc_083" Or _
   appo = "bloc_084" Or appo = "bloc_085" Or appo = "bloc_086" Or appo = "bloc_087" Or appo = "bloc_088" Or _
   appo = "bloc_089" Or appo = "bloc_090" Or appo = "bloc_091" Or appo = "bloc_092" Or appo = "bloc_093" Or _
   appo = "bloc_094" Or appo = "bloc_095" Or appo = "bloc_096" Or appo = "bloc_097" Or appo = "bloc_098" Or _
   appo = "bloc_099" Or appo = "bloc_100" Or appo = "bloc_101" Or appo = "bloc_102" Or appo = "bloc_103" Or _
   appo = "bloc_104" Or appo = "bloc_105" Or appo = "bloc_106" Or appo = "bloc_107" Or appo = "bloc_108" Or _
   appo = "bloc_109" Or appo = "bloc_110" Or appo = "bloc_111" Then
   For i = 0 To CaricaX.GetFaceCount - 1
       CaricaX.GetFace(i).GetTexture.SetDecalTransparency D_TRUE
       CaricaX.GetFace(i).GetTexture.SetDecalTransparentColor RGB(255, 0, 255)
   Next
End If
End Function
Private Sub CreatePlayers()
Dim i As Integer
For i = 1 To UBound(Racers) + 1
    Set m_objectFrame(i) = RMC.D3DRM.CreateFrame(RMC.SceneFrame)
    Set m_meshBuilder(i) = RMC.D3DRM.CreateMeshBuilder()
    If i = 1 Then m_meshBuilder(i).LoadFromFile App.path & "\Mesh\player.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    If i = 2 Or i = 3 Then m_meshBuilder(i).LoadFromFile App.path & "\Mesh\kazi.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    If i = 4 Or i = 5 Then m_meshBuilder(i).LoadFromFile App.path & "\Mesh\cop.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    If i = 6 Or i = 7 Then m_meshBuilder(i).LoadFromFile App.path & "\Mesh\iron.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    If i = 8 Or i = 9 Then m_meshBuilder(i).LoadFromFile App.path & "\Mesh\leroy.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    If i = 10 Or i = 11 Then m_meshBuilder(i).LoadFromFile App.path & "\Mesh\cubik.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    m_objectFrame(i).AddScale D3DRMCOMBINE_REPLACE, 1, 1, 1
    m_objectFrame(i).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
    m_objectFrame(i).AddVisual m_meshBuilder(i)
    Set m_meshBuilder(i) = Nothing
Next i
Set ombreFrame = RMC.D3DRM.CreateFrame(RMC.SceneFrame)
Set ombreMesh = RMC.D3DRM.CreateMeshBuilder()
ombreMesh.LoadFromFile App.path & "\Mesh\ombre.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
'ombreMesh.SetQuality D3DRMRENDER_UNLITFLAT
ombreMesh.GetFace(0).GetTexture.SetDecalTransparency D_TRUE
ombreMesh.GetFace(0).GetTexture.SetDecalTransparentColor RGB(255, 255, 255)
Dim fin As Direct3DRMFrame3
Set fin = RMC.D3DRM.CreateFrame(ombreFrame)
Dim mo As D3DRMMATERIALOVERRIDE
With mo
     .lFlags = D3DRMMATERIALOVERRIDE_DIFFUSE_RGBONLY
     .dcDiffuse.a = 0.1
End With
fin.SetMaterialOverride mo
fin.AddVisual ombreMesh
ombreFrame.AddVisual ombreMesh
Set ombreMesh = Nothing
    
m_objectFrame(1).SetPosition RMC.SceneFrame, StartX * TCase, 2, StartY * TCase
ombreFrame.SetPosition RMC.SceneFrame, StartX * TCase, 0.1, StartY * TCase

End Sub
Private Sub RandNum()
' Random routine for 5 numbers
Static count As Integer, i As Integer
Dim intNum As Integer
count = count + 1
If count > UBound(Racers) Then
    count = 1
    Exit Sub
End If
ReDim Preserve ArrNum15(1 To UBound(Racers))
Randomize
intNum = Int((UBound(Racers) * Rnd) + 1)
                             
If count > 1 Then

    For i = 1 To count - 1
        Do Until ArrNum15(i) < intNum
            If ArrNum15(i) = intNum Then
               intNum = Int((UBound(Racers) * Rnd) + 1)
               i = 1
            Else
               Exit Do
            End If
        Loop
    Next i
End If
ArrNum15(count) = intNum
End Sub
Private Sub InitRacers()
Dim i As Integer, rand As Integer

For i = 1 To UBound(Racers)
    Racers(i).Radius = 2.5
    Racers(i).Acceleration = 0.25
    Racers(i).BrakeSpeed = 1
    Racers(i).Direction = 0
    Racers(ArrNum15(i)).Speed = (i + 1) * 0.05
    Racers(i).MaxSpeed = 8
    Racers(i).MinSpeed = 1
    Racers(ArrNum15(i)).TurnRatio = (i + 1) * 0.05
    If Racers(ArrNum15(i)).TurnRatio > 0.3 Then Racers(ArrNum15(i)).TurnRatio = 0.3
    Racers(i).BackAngle = pi / 3 '* 2
    Racers(i).NextNode = 1
    Racers(i).Lap = 1
    Racers(i).Time = 0
Next

Racers(1).Name = "KAZIM 1"
Racers(1).Location_X = 6.9 * TCase
Racers(1).Location_Z = 2 * TCase
Racers(1).Position = 3
Racers(2).Name = "KAZIM 2"
Racers(2).Location_X = 6.8 * TCase
Racers(2).Location_Z = 1.6 * TCase
Racers(2).Position = 5
Racers(3).Name = "COP 1"
Racers(3).Location_X = 6.7 * TCase
Racers(3).Location_Z = 2.2 * TCase
Racers(3).Position = 10
Racers(4).Name = "COP 2"
Racers(4).Location_X = 6.9 * TCase
Racers(4).Location_Z = 1.8 * TCase
Racers(4).Position = 1
Racers(5).Name = "IRON 1"
Racers(5).Location_X = 6.9 * TCase
Racers(5).Location_Z = 2.1 * TCase
Racers(5).Position = 2
Racers(6).Name = "IRON 2"
Racers(6).Location_X = 6.7 * TCase
Racers(6).Location_Z = 1.9 * TCase
Racers(6).Position = 9
Racers(7).Name = "LEROY 1"
Racers(7).Location_X = 6.8 * TCase
Racers(7).Location_Z = 1.8 * TCase
Racers(7).Position = 6
Racers(8).Name = "LEROY 2"
Racers(8).Location_X = 6.9 * TCase
Racers(8).Location_Z = 2.3 * TCase
Racers(8).Position = 4
Racers(9).Name = "CUBIK 1"
Racers(9).Location_X = 6.7 * TCase
Racers(9).Location_Z = 1.7 * TCase
Racers(9).Position = 8
Racers(10).Name = "CUBIK 2"
Racers(10).Location_X = 6.8 * TCase
Racers(10).Location_Z = 2.1 * TCase
Racers(10).Position = 7

End Sub
Private Sub OpenMap()
Dim campo, rot As Byte, num As Integer, nodi As Integer, ind As Integer
    
       iFree = FreeFile
       Open App.path & "\mappa1.map" For Input As #iFree
       Input #iFree, MapNumber, MapSizeX, MapSizeY, StartX, StartY
    
       'resize and clear out the map arrays
       ReDim map(MapSizeX, MapSizeY)
       ReDim ViewFrame(MapSizeX, MapSizeY)
       ReDim tile(MapSizeX, MapSizeY)
    
       For xx = 0 To MapSizeX - 1
           For yy = 0 To MapSizeY - 1
               Input #iFree, tile(xx, yy)
               num = CInt(Mid(tile(xx, yy), 1, 3))
               rot = CInt(Mid(tile(xx, yy), 4, 1))
               Set ViewFrame(xx, yy) = RMC.D3DRM.CreateFrame(RMC.SceneFrame)
               With ViewFrame(xx, yy)
                    .AddScale D3DRMCOMBINE_REPLACE, 0.5, 0.5, 0.5
                    .SetPosition RMC.SceneFrame, xx * 60, 0, yy * 60
                    If rot = 1 Then .AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 0 * (pi / 180)
                    If rot = 2 Then .AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
                    If rot = 3 Then .AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 180 * (pi / 180)
                    If rot = 4 Then .AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 270 * (pi / 180)
                    .AddVisual Wall(num)
               End With
           Next yy
           Input #iFree, campo
       Next xx
       
      Input #iFree, num
    
      ReDim BonusFrame(num)
      ReDim BonusMesh(num)
      ReDim Bonus(num)
      For ind = 1 To num
          Input #iFree, xx, yy
          Set BonusFrame(ind) = RMC.D3DRM.CreateFrame(RMC.SceneFrame)
          Set BonusMesh(ind) = RMC.D3DRM.CreateMeshBuilder()
          BonusMesh(ind).LoadFromFile App.path & "\Mesh\bonus.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
          BonusFrame(ind).AddScale D3DRMCOMBINE_REPLACE, 1, 1, 1
          BonusFrame(ind).SetPosition RMC.SceneFrame, xx * TCase, 0, yy * TCase
          Bonus(ind).Location_X = xx * TCase
          Bonus(ind).Location_Z = yy * TCase
          Bonus(ind).Active = True
          Bonus(ind).sec = 0
          BonusFrame(ind).AddVisual BonusMesh(ind)
       Next
       
       Close #iFree
       
       Open App.path & "\path.map" For Input As #iFree
       Input #iFree, nodi
       ReDim Nodes(nodi)
       For ind = 1 To nodi
           Input #iFree, Nodes(ind).Location_X, Nodes(ind).Location_Y, Nodes(ind).Dir
           Nodes(ind).Location_X = Nodes(ind).Location_X * TCase
           Nodes(ind).Location_Y = Nodes(ind).Location_Y * TCase
       Next
       Close #iFree
    
End Sub
Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case vbKeyUp: avanti = True
    Case vbKeyDown: indietro = True
    Case vbKeyLeft: sinistra = True
    Case vbKeyRight: destra = True
    Case vbKeyEscape: fine = True
    Case vbKeyV: retro = True: If ret = 0 Then ret = 1 Else ret = 2
    Case vbKeyF1: F1 = True
    Case vbKeyF2: F2 = True
    Case vbKeyF3: F3 = True
    Case vbKeyF4: F4 = True
    Case vbKeyF5: F5 = True
    Case vbKeyF6: F6 = True
    Case vbKeyF7: F7 = True
    Case vbKeyF8: F8 = True
    Case vbKeyF9: F9 = True
    Case vbKeyF10: F10 = True
    Case vbKeyF11: F11 = True
    Case vbKeyF12: F12 = True
 
 End Select
End Sub
Private Sub pic_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp: avanti = False
    Case vbKeyDown: indietro = False
    Case vbKeyLeft: sinistra = False
    Case vbKeyRight: destra = False
    Case vbKeyEscape: fine = False
    Case vbKeyV: retro = False: ret = 0
    Case vbKeyF1: F1 = False
    Case vbKeyF2: F2 = False
    Case vbKeyF3: F3 = False
    Case vbKeyF4: F4 = False
    Case vbKeyF5: F5 = False
    Case vbKeyF6: F6 = False
    Case vbKeyF7: F7 = False
    Case vbKeyF8: F8 = False
    Case vbKeyF9: F9 = False
    Case vbKeyF10: F10 = False
    Case vbKeyF11: F11 = False
    Case vbKeyF12: F12 = False
  
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    running = False
    Cleanup
End Sub
Private Sub Form_Resize()
    Pic.Width = Me.ScaleWidth
    Pic.Height = Me.ScaleHeight
    If running = False Then Exit Sub
    If RMC.IsFullScreen = True Then Exit Sub
    RMC.Resize Pic.ScaleWidth, Pic.ScaleHeight
End Sub
Private Sub InitDeviceobjects()
Dim fogcolor
fogcolor = 0 * 65536 + 125 * 256 + 125

    Dim vp As Direct3DRMViewport2
    Set vp = RMC.Viewport
    vp.SetBack 1000
    vp.SetFront 1
    vp.SetProjection D3DRMPROJECT_PERSPECTIVE
    
    Set mat = RMC.D3DRM.CreateMaterial(0)
    With mat
        .SetAmbient 1, 1, 1
    End With
    
    RMC.SceneFrame.SetSceneBackground fogcolor
    RMC.SceneFrame.SetSceneFogEnable D_TRUE
    RMC.SceneFrame.SetSceneFogMethod D3DRMFOGMETHOD_TABLE
    RMC.SceneFrame.SetSceneFogColor fogcolor
    RMC.SceneFrame.SetSceneFogMode D3DRMFOG_LINEAR
    RMC.SceneFrame.SetSceneFogParams 520, 1000, 1
    
    With RMC.Device
        .SetTextureQuality D3DRMTEXTURE_LINEARMIPLINEAR
        .SetQuality D3DRMFILL_SOLID Or D3DRMLIGHT_ON Or D3DRMRENDER_GOURAUD Or D3DRMSHADE_GOURAUD
        .SetDither D_TRUE
        .SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY Or D3DRMRENDERMODE_SORTEDTRANSPARENCY
    End With
    
    RMC.AmbientLight.SetColorRGB 255, 255, 255
    RMC.DirLight.SetType D3DRMLIGHT_POINT
    RMC.DirLightFrame.SetPosition Nothing, 1 * TCase, 10 * TCase, 1 * TCase
    
End Sub
Private Sub RM_E_DirecXNotInstalled()
    MsgBox "DirectX7 is not installed", vbCritical
    End
End Sub
Private Sub RM_E_Error4(Errstr As String)
    MsgBox Errstr, vbCritical, "Error"
    End
End Sub
Private Sub RM_E_PostRender()
Dim i As Integer, Annuncio As String
    
    On Local Error Resume Next
       If Numdown < 4 Then
          MyFont.Name = "Comic Sans MS"
          MyFont.Size = 200
          MyFont.Bold = True
          RMC.BackBuffer.SetFont MyFont
          RMC.BackBuffer.SetForeColor RGB(255, 255, 255)
       End If
       If Numdown = 1 Then RMC.BackBuffer.DrawText (frmMain.ScaleWidth - 200) / 2, (frmMain.ScaleHeight - 400) / 2, "3", False
       If Numdown = 2 Then RMC.BackBuffer.DrawText (frmMain.ScaleWidth - 200) / 2, (frmMain.ScaleHeight - 400) / 2, "2", False
       If Numdown = 3 Then RMC.BackBuffer.DrawText (frmMain.ScaleWidth - 200) / 2, (frmMain.ScaleHeight - 400) / 2, "1", False
       If Numdown = 4 Then RMC.BackBuffer.DrawText (frmMain.ScaleWidth - 350) / 2, (frmMain.ScaleHeight - 400) / 2, "GO!", False
       If Numdown > 4 And Lap < 4 Then
          MyFont.Name = "Impact"
          MyFont.Size = 15
          MyFont.Bold = True
          RMC.BackBuffer.SetFont MyFont
          RMC.BackBuffer.SetForeColor RGB(255, 255, 255)  'vbWhite
          RMC.BackBuffer.DrawText 20, frmMain.ScaleHeight - 30, "Velocity: " + CStr(CInt(Velocity)), False
          RMC.BackBuffer.DrawText (frmMain.ScaleWidth - 20) / 2, frmMain.ScaleHeight - 30, "Time: " + CStr(Format(Timemin, "00")) + "." & CStr(Format(Timesec, "00")), False
          RMC.BackBuffer.DrawText (frmMain.ScaleWidth - 100), frmMain.ScaleHeight - 30, "Lap: " + CStr(Lap) & "/3", False
       End If
       If Lap > 3 Then
          For i = 1 To UBound(Racers)
              If Racers(i).Position = 1 Then
                 If Racers(i).Lap = 4 And Racers(i).Time < CStr(Format(Timemin, "00")) + CStr(Format(Timesec, "00")) Then
                    Annuncio = "WINNER IS " & CStr(Racers(i).Name)
                    Besttime = CStr(Mid(Racers(i).Time, 1, 2)) & "." & CStr(Mid(Racers(i).Time, 3, 2))
                 Else
                    Annuncio = "   WINNER IS YOU"
                    Besttime = CStr(Format(Timemin, "00")) + "." & CStr(Format(Timesec, "00"))
                 End If
              End If
          Next
          
          MyFont.Name = "Comic Sans MS"
          MyFont.Size = 50
          MyFont.Bold = True
          RMC.BackBuffer.SetFont MyFont
          RMC.BackBuffer.SetForeColor RGB(255, 255, 255)
          RMC.BackBuffer.DrawText 120, (frmMain.ScaleHeight - 400) / 2, Annuncio, False
          MyFont.Size = 20
          RMC.BackBuffer.DrawText 200, (frmMain.ScaleHeight - 400) / 2 + 150, "Best Time: " + Besttime, False
          RMC.BackBuffer.DrawText 200, (frmMain.ScaleHeight - 400) / 2 + 250, "Your Time: " + CStr(Format(Timemin, "00")) + "." & CStr(Format(Timesec, "00")), False
          MyFont.Name = "Impact"
          MyFont.Size = 15
          MyFont.Bold = True
          RMC.BackBuffer.SetFont MyFont
          RMC.BackBuffer.SetForeColor RGB(255, 255, 255)
          RMC.BackBuffer.DrawText (frmMain.ScaleWidth - 20) / 2, frmMain.ScaleHeight - 30, "Press ESC to Quit", False
       
       End If
'       RMC.BackBuffer.DrawText 10, 10, "Frame/Sec:" + CStr(RMC.FPS), False
End Sub
Private Sub Cleanup()
Dim i As Integer, j As Integer
For i = 0 To MapSizeX - 1
    For j = 0 To MapSizeY - 1
        Set ViewFrame(i, j) = Nothing
    Next j
Next i
For i = 0 To 111
    Set Wall(i) = Nothing
Next i
Set ombreFrame = Nothing
For i = 1 To UBound(Racers) + 1
    Set m_objectFrame(i) = Nothing
    Set m_meshBuilder(i) = Nothing
Next i

Music_mod.EndMusic

Set Playersnd = Nothing
Set Playersnd3D = Nothing
For i = 1 To UBound(Racers)
    Set Driversnd(i) = Nothing
    Set Driversnd3D(i) = Nothing
Next i
Set Boingsnd = Nothing
Set Boingsnd3D = Nothing
Set Beepsnd = Nothing
Set Beepsnd3D = Nothing
Set Hit1snd = Nothing
Set Hit1snd3D = Nothing
Set Hit2snd = Nothing
Set Hit2snd3D = Nothing
Set Bonussnd = Nothing
Set Bonussnd3D = Nothing
    
Set RM_E = Nothing
Set mat = Nothing
    
Set DINPUT = Nothing
Set DIdevice = Nothing

ShowCursor 1

End

End Sub
Private Function GetDistance(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
    'calculates the distance between two points
    GetDistance = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
End Function

Private Sub RacersAI(ind As Integer)
    Dim AbsoluteAngle   As Single
    Dim RelativeAngle   As Single
    Dim NextNode        As Integer
    Dim distance        As Integer
    NextNode = Racers(ind).NextNode  'we could reference this via the object/udt but its neater like this
    
'## Check for Node Collisions
    distance = GetDistance(Racers(ind).Location_X, Racers(ind).Location_Z, Nodes(NextNode).Location_X, Nodes(NextNode).Location_Y)
    If distance < Racers(ind).Radius + NodeRadius Then
        If Nodes(NextNode).Dir = "+" Then
           m_objectFrame(ind + 1).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
        ElseIf Nodes(NextNode).Dir = "-" Then
           m_objectFrame(ind + 1).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -90 * (pi / 180)
        End If
        NextNode = NextNode + 1
        
        If NextNode > UBound(Nodes) Then
           NextNode = 1
           Lastnodes = True
        End If
        Racers(ind).NextNode = NextNode
    End If

'## Calculate Direction
    AbsoluteAngle = FindAngle(Racers(ind).Location_X, Racers(ind).Location_Z, Nodes(NextNode).Location_X, Nodes(NextNode).Location_Y)
    
    'We now have the angle from us to our next node
    'we're going to re-orient it so its a relative angle
    CheckAngle Racers(ind).Direction
    CheckAngle AbsoluteAngle
    RelativeAngle = AbsoluteAngle - Racers(ind).Direction
    CheckAngle RelativeAngle
    If RelativeAngle > 0 Then
        Racers(ind).Direction = Racers(ind).Direction + Racers(ind).TurnRatio
    Else
        Racers(ind).Direction = Racers(ind).Direction - Racers(ind).TurnRatio
    End If



'## Calculate Speed
    'work out our turning circle
    If RelativeAngle > 0 Then
        ProjectCircle CircleX, CircleY, CircleRadius, True, ind  'the first 3 params are where the data is returned to after
    Else
        ProjectCircle CircleX, CircleY, CircleRadius, False, ind 'the sub executes.
    End If
    'check if the node is within our turning circle
    'ie if we dont slow down we'll miss it cause we cant turn enough
    distance = GetDistance(CircleX, CircleY, Nodes(NextNode).Location_X, Nodes(NextNode).Location_Y)
    If distance < CircleRadius Then
        Racers(ind).Speed = Racers(ind).Speed - Racers(ind).BrakeSpeed
        If Racers(ind).Speed < Racers(ind).MinSpeed Then Racers(ind).Speed = Racers(ind).MinSpeed
    ElseIf Abs(RelativeAngle) > Racers(ind).BackAngle Then
    'Check if node is behind us and slow down if it is
        Racers(ind).Speed = Racers(ind).Speed - Racers(ind).BrakeSpeed
        If Racers(ind).Speed < Racers(ind).MinSpeed Then Racers(ind).Speed = Racers(ind).MinSpeed
    Else 'we are a race car, we should speed up if we can
        Racers(ind).Speed = Racers(ind).Speed + Racers(ind).Acceleration
        If Racers(ind).Speed > Racers(ind).MaxSpeed Then Racers(ind).Speed = Racers(ind).MaxSpeed
    End If
    
    If Lastnodes = True And GetDistance(Racers(ind).Location_X, Racers(ind).Location_Z, Nodes(UBound(Nodes)).Location_X, Nodes(UBound(Nodes)).Location_Y) < 60 Then
       Racers(ind).Lap = Racers(ind).Lap + 1
       Racers(ind).Time = Format(CStr(Timemin), "00") + Format(CStr(Timesec), "00")
       Racers(ind).Position = cont
       Lastnodes = False
       
       Racerspos(ind) = cont
       cont = cont + 1
       If cont > 10 Then cont = 1
    End If

End Sub

Private Sub CheckAngle(ByRef Dir As Single)
    'simply ensures that a direction is within a range of -1*Pi to +1*Pi
    While Dir > pi Or Dir < -pi
        If Dir > pi Then Dir = Dir - 2 * pi
        If Dir < -pi Then Dir = Dir + 2 * pi
    Wend
End Sub

Private Sub MoveRacers(ind As Integer)
Dim i As Integer, distance As Integer
'Simple collision through the racers (could be implemented)
For i = 1 To UBound(Racers)
    If i <> ind Then
       distance = GetDistance(Racers(ind).Location_X, Racers(ind).Location_Z, Racers(i).Location_X, Racers(i).Location_Z)
       If distance < Racers(ind).Radius + Racers(i).Radius Then
             Racers(ind).Speed = -Racers(ind).Speed
             Racers(i).Speed = -Racers(i).Speed
'             Hit1snd3D.SetPosition Racers(ind).Location_X, 1, Racers(ind).Location_Z, DS3D_IMMEDIATE
'             Hit1snd.Play DSBPLAY_DEFAULT
'             Hit2snd3D.SetPosition Racers(i).Location_X, 1, Racers(i).Location_Z, DS3D_IMMEDIATE
'             Hit2snd.Play DSBPLAY_DEFAULT
       End If
   End If
Next

'Position the racers
Racers(ind).Location_X = Racers(ind).Location_X + Racers(ind).Speed * Sin(Racers(ind).Direction)
Racers(ind).Location_Z = Racers(ind).Location_Z - Racers(ind).Speed * Cos(Racers(ind).Direction)
m_objectFrame(ind + 1).SetPosition Nothing, Racers(ind).Location_X, 1, Racers(ind).Location_Z
'Driversnd3D(ind).SetPosition Racers(ind).Location_X, 1, Racers(ind).Location_Z, DS3D_IMMEDIATE
'Driversnd(ind).Play DSBPLAY_DEFAULT

End Sub

Private Sub ProjectCircle(rtnOriginX As Single, rtnOriginY As Single, rtnRadius As Single, TurnLeft As Boolean, ind As Integer)
'We calculate out the minimum circle we can turn at our current speed
'NOTE: I could not have done this bit so efficiently without the help of a great guy called Brykovian
    
    Dim VelocityX1  As Single
    Dim VelocityY1  As Single
    Dim VelocityX2  As Single
    Dim VelocityY2  As Single
    Dim AccelX      As Single
    Dim AccelY      As Single
    Dim AccelTot    As Single
    Dim Radius      As Single
    Dim OriginX     As Single
    Dim OriginY     As Single
    Dim t           As Integer
    Const NumTicks = 10
    DummyRacers(ind) = Racers(ind)
    
    'First we must calculate the seperate X & Y velocities
    'We project the motion of the racers for a few ticks to get more accuracy
    For t = 1 To NumTicks
        If TurnLeft Then
            DummyRacers(ind).Direction = DummyRacers(ind).Direction + DummyRacers(ind).TurnRatio
        Else 'Turn Right
            DummyRacers(ind).Direction = DummyRacers(ind).Direction - DummyRacers(ind).TurnRatio
        End If
        DummyRacers(ind).Location_X = DummyRacers(ind).Location_X + DummyRacers(ind).Speed * Sin(DummyRacers(ind).Direction)
        DummyRacers(ind).Location_Z = DummyRacers(ind).Location_Z - DummyRacers(ind).Speed * Cos(DummyRacers(ind).Direction)
    Next t
    VelocityX1 = Sin(Racers(ind).Direction) * Racers(ind).Speed
    VelocityY1 = Cos(Racers(ind).Direction) * Racers(ind).Speed
    VelocityX2 = Sin(DummyRacers(ind).Direction) * DummyRacers(ind).Speed
    VelocityY2 = Cos(DummyRacers(ind).Direction) * DummyRacers(ind).Speed

    'Now we calculate the acceleration towards the center of the circle
    AccelX = (VelocityX2 - VelocityX1) / NumTicks
    AccelY = (VelocityY2 - VelocityY1) / NumTicks
    AccelTot = Sqr(AccelX * AccelX + AccelY * AccelY)

    'Finally we can work out the radius of our circle using
    'On Error GoTo OverFlowHandler
    If Radius > 0 Then
       Radius = (Racers(ind).Speed * Racers(ind).Speed) / AccelTot
    Else
       Radius = 0
    End If
    'On Error GoTo 0

    'now it just remains of cource to calculate the origin of our circle
    If TurnLeft Then
        OriginX = Racers(ind).Location_X + Radius * Sin(Racers(ind).Direction + pi / 2)
        OriginY = Racers(ind).Location_Z - Radius * Cos(Racers(ind).Direction + pi / 2)
    Else 'Turn Right
        OriginX = Racers(ind).Location_X + Radius * Sin(Racers(ind).Direction - pi / 2)
        OriginY = Racers(ind).Location_Z - Radius * Cos(Racers(ind).Direction - pi / 2)
    End If

    'And now really finally we pass the results back
    rtnOriginX = OriginX
    rtnOriginY = OriginY
    rtnRadius = Radius
    Exit Sub
OverFlowHandler:
'   When/If the car slows to exactly 0 we get an overflow error
    Radius = 0
    Resume Next
End Sub

Private Function FindAngle(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
    Dim sngXComp As Single
    Dim sngYComp As Single

    'Find the angle between the 2 coords
    sngXComp = X2 - X1
    sngYComp = Y1 - Y2
    If Sgn(sngYComp) > 0 Then FindAngle = Atn(sngXComp / sngYComp)
    If Sgn(sngYComp) < 0 Then FindAngle = Atn(sngXComp / sngYComp) + pi
End Function
Private Sub Tmr_time_Timer()
Numdown = 5
Timesec = Timesec + 1
If Timesec > 59 Then
   Timesec = 0
   Timemin = Timemin + 1
End If
End Sub
Private Sub InitSounds()
Dim i As Byte

    RMC.InitDsound70
    
    'Player car sound
    Set Playersnd = RMC.Create2DsBuffromfile70(App.path & "\Audio\engine1.wav")
    Set Playersnd3D = RMC.Create3DSBUFfrom2Dbuf70(Playersnd)
    Playersnd3D.SetMinDistance 1 * TCase, DS3D_IMMEDIATE
    Playersnd3D.SetMaxDistance 2 * TCase, DS3D_IMMEDIATE
   
    'Drivers car sound
    For i = 1 To UBound(Racers)
        If i Mod 2 = 0 Then
           Set Driversnd(i) = RMC.Create2DsBuffromfile70(App.path & "\Audio\engine1.wav")
           Set Driversnd3D(i) = RMC.Create3DSBUFfrom2Dbuf70(Driversnd(i))
           Driversnd3D(i).SetMinDistance 1 * TCase, DS3D_IMMEDIATE
           Driversnd3D(i).SetMaxDistance 3 * TCase, DS3D_IMMEDIATE
        Else
           Set Driversnd(i) = RMC.Create2DsBuffromfile70(App.path & "\Audio\engine1.wav")
           Set Driversnd3D(i) = RMC.Create3DSBUFfrom2Dbuf70(Driversnd(i))
           Driversnd3D(i).SetMinDistance 1 * TCase, DS3D_IMMEDIATE
           Driversnd3D(i).SetMaxDistance 3 * TCase, DS3D_IMMEDIATE
        End If
    Next
    
    'Trumpet
    Set Beepsnd = RMC.Create2DsBuffromfile70(App.path & "\Audio\Beep.wav")
    Set Beepsnd3D = RMC.Create3DSBUFfrom2Dbuf70(Beepsnd)
    Beepsnd3D.SetMinDistance 10, DS3D_IMMEDIATE
    Beepsnd3D.SetMaxDistance 20, DS3D_IMMEDIATE
    
    'Collision with walls sound
    Set Boingsnd = RMC.Create2DsBuffromfile70(App.path & "\Audio\Boing.wav")
    Set Boingsnd3D = RMC.Create3DSBUFfrom2Dbuf70(Boingsnd)
    Boingsnd3D.SetMinDistance 1, DS3D_IMMEDIATE
    Boingsnd3D.SetMaxDistance 1, DS3D_IMMEDIATE
    
    'Collision with drivers
    Set Hit1snd = RMC.Create2DsBuffromfile70(App.path & "\Audio\Hit1.wav")
    Set Hit1snd3D = RMC.Create3DSBUFfrom2Dbuf70(Hit1snd)
    Hit1snd3D.SetMinDistance 10, DS3D_IMMEDIATE
    Hit1snd3D.SetMaxDistance 20, DS3D_IMMEDIATE
    
    'Collision between racers
    Set Hit2snd = RMC.Create2DsBuffromfile70(App.path & "\Audio\Hit2.wav")
    Set Hit2snd3D = RMC.Create3DSBUFfrom2Dbuf70(Hit2snd)
    Hit2snd3D.SetMinDistance 5, DS3D_IMMEDIATE
    Hit2snd3D.SetMaxDistance 10, DS3D_IMMEDIATE
    
    'Pick bonus
    Set Bonussnd = RMC.Create2DsBuffromfile70(App.path & "\Audio\Bonus.wav")
    Set Bonussnd3D = RMC.Create3DSBUFfrom2Dbuf70(Bonussnd)
    Bonussnd3D.SetMinDistance 1, DS3D_IMMEDIATE
    Bonussnd3D.SetMaxDistance 1, DS3D_IMMEDIATE

End Sub
Sub doTargets()
'For other maps these point have to stay in a file
'they are good only for this demo map
ReDim Targets(9)
Targets(1).Location_X = 23 * TCase
Targets(1).Location_Z = 7 * TCase
Targets(1).Pass = False
Targets(1).Size = 60
Targets(2).Location_X = 28 * TCase
Targets(2).Location_Z = 14 * TCase
Targets(2).Pass = False
Targets(2).Size = 60
Targets(3).Location_X = 43 * TCase
Targets(3).Location_Z = 8 * TCase
Targets(3).Pass = False
Targets(3).Size = 60
Targets(4).Location_X = 49 * TCase
Targets(4).Location_Z = 15 * TCase
Targets(4).Pass = False
Targets(4).Size = 60
Targets(5).Location_X = 41 * TCase
Targets(5).Location_Z = 37 * TCase
Targets(5).Pass = False
Targets(5).Size = 60
Targets(6).Location_X = 29 * TCase
Targets(6).Location_Z = 46 * TCase
Targets(6).Pass = False
Targets(6).Size = 60
Targets(7).Location_X = 12 * TCase
Targets(7).Location_Z = 37 * TCase
Targets(7).Pass = False
Targets(7).Size = 60
Targets(8).Location_X = 3 * TCase
Targets(8).Location_Z = 5 * TCase
Targets(8).Pass = False
Targets(8).Size = 60
Targets(9).Location_X = 8 * TCase
Targets(9).Location_Z = 2 * TCase
Targets(9).Pass = False
Targets(9).Size = 60
TargetsCount = 9
End Sub
Sub Finish()
   Tmr_time.Enabled = False
   If Velocity > 0 Then Velocity = Velocity - 0.5
   If Velocity < 0 Then Velocity = 0
   D3Pos.x = D3Pos.x + Cos(Heading) * Velocity
   D3Pos.y = D3Pos.y
   D3Pos.z = D3Pos.z + Sin(Heading) * Velocity
   If dview < 50 Then dview = dview + 0.5
   If oview > -50 Then oview = oview - 0.5
   If hview < 20 Then hview = hview + 0.05
   m_objectFrame(1).SetPosition Nothing, D3Pos.x, D3Pos.y, D3Pos.z
   ombreFrame.SetPosition m_objectFrame(1), 0, -1.95, 0
   RMC.CameraFrame.SetPosition m_objectFrame(1), oview, hview, dview
   RMC.CameraFrame.LookAt m_objectFrame(1), Nothing, D3DRMCONSTRAIN_Z
End Sub
