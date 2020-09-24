Attribute VB_Name = "basGame"
Option Explicit

'Author: Grzegorz Holdys (Wroclaw, Poland)
'E-mail: gregor@kn.pl

'Debugging and Modifications: Ulli (UMGEDV GmbH)
'E-mail: umgedv@yahoo.com

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function BeginIdleDetection Lib "Msidle.dll" Alias "#3" (ByVal pfnCallback As Long, ByVal dwIdleMin As Long, ByVal dwReserved As Long) As Long
Private Declare Function EndIdleDetection Lib "Msidle.dll" Alias "#4" (ByVal dwReserved As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long 'good for more than 24 days continuous Windows up time
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPenPosition Lib "gdi32" Alias "MoveToEx" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function DrawLine Lib "gdi32" Alias "LineTo" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function Beeper Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Const MF_BYPOSITION     As Long = &H400
Private Const MF_GRAYED         As Long = 1
Private Const MenuCloseItem     As Long = 6
Private Const WM_NCPAINT        As Long = &H85
Private Const USER_IDLE_BEGIN   As Long = 1
Private Const USER_IDLE_END     As Long = 2

Private hDCViewport             As Long

Private Enum PLAYER_CONTROL
    PC_HUMAN = 1
    PC_CPU = 2
End Enum
#If False Then 'Spoof to preserve Enum capitalization
Private PC_HUMAN, PC_CPU
#End If

Private Enum GAME_STAGE
    GM_INITIAL_LOOP = 0
    GM_BILLARDS_IN_MOTION = 1
    GM_PLAYERS_AIMING_INIT = 2
    GM_PLAYERS_AIMING = 3
    GM_SHOWING_RESULTS_INIT = 4
    GM_SHOWING_RESULTS = 5
    GM_FREEBALL_INIT = 6
    GM_FREEBALL = 7
    GM_8BALL_INIT = 8
    GM_8BALL = 9
End Enum
#If False Then
Private GM_INITIAL_LOOP, GM_BILLARDS_IN_MOTION, GM_PLAYERS_AIMING_INIT, GM_PLAYERS_AIMING, GM_SHOWING_RESULTS_INIT, GM_SHOWING_RESULTS, _
        GM_FREEBALL_INIT, GM_FREEBALL, GM_8BALL_INIT, GM_8BALL
#End If

Public Enum STICKER_CAPTIONS
    SC_PLAYER1 = 1
    SC_PLAYER2 = 2
    SC_FREEBALL = 3
    SC_EIGHTBALL = 4
    SC_WINS = 5
    SC_BEGINS = 6
    SC_FOUL = 7
End Enum
#If False Then 'Spoof to preserve Enum capitalization
Private SC_PLAYER1, SC_PLAYER2, SC_FREEBALL, SC_EIGHTBALL, SC_WINS, SC_BEGINS, SC_FOUL
#End If

'DirectX8 objects
Private m_DX                    As DirectX8
Private m_D3DX                  As D3DX8
Private m_D3D                   As Direct3D8
Private m_D3DDevice             As Direct3DDevice8

'Viewport
Public Viewport                 As D3DVIEWPORT8
Public Down                     As Boolean

'The table...
Private m_Table                 As clsTable
'...its dimensions...
Private m_TableWidth            As Single
Private m_TableLength           As Single
Private m_GameAreaMinZ          As Single
Private m_GameAreaMaxZ          As Single
Private m_GameAreaMinX          As Single
Private m_GameAreaMaxX          As Single
'...and its pockets
Private m_Pockets(1 To 6)       As D3DVECTOR
Private m_PocketRadius          As Single

'The billards...
Public m_Billards               As clsBillards
Private m_numBillards           As Long
Private m_BPocketNumbers()      As Long
Private m_BillRadius            As Single

'Two cameras - one is moveable and the other is fixed hanging from the ceiling and looking down
Public Enum CamTypes
    CamMoveable = 1
    CamFixed = 2
End Enum
#If False Then
Private CamFixed, CamMoveable
#End If
Private m_Cameras(CamMoveable To CamFixed) As clsCamera
Public m_ActiveCam              As Long

'Collision controller
Private m_PCC                   As clsCollisionController

'Lights
Private m_Lamp                  As D3DLIGHT8
Private LightPower              As Single

'Sounds
Private m_DirectSound8          As DirectSound8
Public m_DSOops                 As DirectSoundSecondaryBuffer8
Public m_DSTransitionLong       As DirectSoundSecondaryBuffer8
Private m_DSClap                As DirectSoundSecondaryBuffer8
Private m_DSBillBillHit         As DirectSoundSecondaryBuffer8
Private m_DSBillTableHit        As DirectSoundSecondaryBuffer8
Private m_DSCueBallLaunched     As DirectSoundSecondaryBuffer8
Private m_DSPocketHit           As DirectSoundSecondaryBuffer8

Private m_VolumeBase            As Single
Public Enum VolumeBase
    m_VolumeMute = -9999
    m_VolumeLow = 6500
    m_VolumeMedium = 8600
    m_VolumeMax = 10000
End Enum
#If False Then
Private m_VolumeMute, m_VolumeLow, m_VolumeMedium, m_VolumeMax
#End If

'Players
Private m_Players(1 To 2)       As PLAYER_CONTROL
Public ShotCounts(1 To 2)       As Long
Public SinkCounts(1 To 2)       As Long
Public PrevShotCounts(1 To 2)   As Long
Public PrevSinkCounts(1 To 2)   As Long

Public m_Designated8BallPocket  As Long
Public m_CurrentPlayer          As Long
Private m_PreviousPlayer        As Long
Private m_AndTheWinnerIs        As Long

'Billboard
Private m_Billboard             As clsBillboard

'FILE NAMES
'Textures and pictures
Public Const BillardBMP         As String = "\Billard#.BMP"
Public Const WoodBMP            As String = "\Wood.BMP"
Public Const ClothBMP           As String = "\Cloth.BMP"

'Sounds
Public Const OopsWAV            As String = "\Oops.WAV"
Public Const BallBallHitWAV     As String = "\BallBallHit.WAV"
Public Const BallTableHitWAV    As String = "\BallTableHit.WAV"
Public Const PocketHitWAV       As String = "\PocketHit.WAV"
Public Const CueBallLaunchWAV   As String = "\CueBallLaunch.WAV"
Public Const TransitionLongWAV  As String = "\TransitionLong.WAV"
Public Const ClapWAV            As String = "\Clap.WAV"

'DirectDraw
Public Const Player1DDS         As String = "\Player1.DDS"
Public Const Player2DDS         As String = "\Player2.DDS"
Public Const FreeBallDDS        As String = "\FreeBall.DDS"
Public Const EightBallDDS       As String = "\EightBall.DDS"
Public Const WinsDDS            As String = "\Wins.DDS"
Public Const BeginsDDS          As String = "\Begins.DDS"
Public Const ArrowDDS           As String = "\Arrow.DDS"
Public Const BlackHoleDDS       As String = "\BlackHole.DDS"
Public Const FoulDDS            As String = "\Foul.DDS"

'Sprite and sprite images
Private Const m_numStickers     As Long = 7
Private m_Sprite                As D3DXSprite
Public m_Stickers(1 To m_numStickers)  As clsScreenSticker

'Current game stage
Private m_GameStage             As GAME_STAGE

'A boolean indicating that a new game has been started
Private m_bNewGame              As Boolean

Public AimingLine               As Boolean
Public JoyPresent               As Boolean
Public m_bButtonIsDown          As Boolean
Public InNormalMode             As Boolean 'not in cheat mode
Public FastDecelerate           As Boolean 'true when 8 ball was sunk, causes fast deceleration of all moving billards

'Time raster g_dt
Public g_dt                     As Single   'master clock
Private PerfFreq                As Currency 'high speed counter
Private NextTick                As Currency
Private FPSLastCheck            As Currency

'some other goodies
Public Const SloMoDelay         As Single = 4 'normal speed / 4
Public SloMoFactor              As Single
Private Const TransTime         As Single = 3 'secs time for a camera transition
Public TransFrames              As Long  'number of frames for camera transition (time im secs for transit)
Public g_Pi                     As Single
Public g_PiHalf                 As Single
Public g_2Pi                    As Single
Private Const AimingElevation   As Single = 0.33
Public Const Shoot              As String = "Shoot"

'Helpers and iterators
Private PreviousVolume          As Long
Public Determinant              As Single
Private i                       As Long
Private j                       As Long
Private FPSCount                As Long
Private vVctr1                  As D3DVECTOR
Private vVctr2                  As D3DVECTOR
Private Declare Sub InitCommonControls Lib "comctl32" ()

Public Sub ActiveWait(Time As Single) 'seconds

  Dim Cnt As Long

    For Cnt = 1 To Time / g_dt
        Render
    Next Cnt

End Sub

Private Sub ClearFrames()

    With frmPool
        .lblPl(m_CurrentPlayer - 1).BorderStyle = vbBSNone
        .lblPl(1 Xor (m_CurrentPlayer - 1)).BorderStyle = vbBSNone
    End With 'FRMPOOL

End Sub

Public Sub CloseDown()

  'Clears all objects and terminates

    Set m_Table = Nothing
    Set m_Billards = Nothing
    Set m_PCC = Nothing
    Set m_Billboard = Nothing
    Set m_DirectSound8 = Nothing
    Set m_DSOops = Nothing
    Set m_DSBillBillHit = Nothing
    Set m_DSBillTableHit = Nothing
    Set m_DSCueBallLaunched = Nothing
    Set m_DSPocketHit = Nothing
    Set m_DSTransitionLong = Nothing
    Set m_DSClap = Nothing
    Set m_Sprite = Nothing
    Set m_D3DDevice = Nothing
    Set m_D3D = Nothing
    Set m_D3DX = Nothing
    Set m_DX = Nothing

    Erase m_Cameras, m_BPocketNumbers, m_Stickers

End Sub

Public Sub DisplayCounts()

    With frmPool.lblPl(m_CurrentPlayer - 1)
        .Caption = "Player " & m_CurrentPlayer & "  -  Shots: " & ShotCounts(m_CurrentPlayer) & "   Sunk: " & SinkCounts(m_CurrentPlayer)
        .Visible = True
    End With 'FRMPOOL.LBLPL(M_CURRENTPLAYER

End Sub

Public Function DotProduct(v1 As D3DVECTOR, v2 As D3DVECTOR) As Single

    With v1
        DotProduct = .x * v2.x + .y * v2.y + .z * v2.z
    End With 'V1

End Function

Private Sub EvalLastStrike()

  'Tests winning/loosing conditions and decides about who shoots next.

  Dim bPlayerScored As Boolean

  'check the 8th ball

    If m_Billards.FellInPocketNumber(8) Then
        'If it fell into other pocket than it was supposed to, then the other player wins
        If m_Billards.FellInPocketNumber(8) <> m_Designated8BallPocket Then
            SinkCounts(m_CurrentPlayer) = SinkCounts(m_CurrentPlayer) - 1
            m_AndTheWinnerIs = m_CurrentPlayer Xor 3
          Else 'NOT M_BILLARDS.FELLINPOCKETNUMBER(8)...
            m_AndTheWinnerIs = m_CurrentPlayer
            PlaySound m_DSClap, 1
        End If
        'Change the game's mode, from GM_BILLARDS_IN_MOTION to the initial results showing mode.
        m_GameStage = GM_SHOWING_RESULTS_INIT
        m_Billards.ClearPocketNumbers
        'Show relevant messages
        m_Stickers(m_AndTheWinnerIs).Visible = True
        m_Stickers(SC_WINS).Visible = True
        ClearFrames
        frmPool.mnuUndo.Enabled = False
      Else 'M_BILLARDS.FELLINPOCKETNUMBER(8) = FALSE/0
        'Now, check the cue-ball. If it fell into any pocket then the players change places
        If m_Billards.FellInPocketNumber(0) Then
            m_CurrentPlayer = m_CurrentPlayer Xor 3
            m_Billards.ClearPocketNumbers
            'Show relevant messages
            m_Stickers(m_CurrentPlayer).Visible = True
            m_Stickers(SC_FREEBALL).Visible = True
            m_Billboard.Visible False
            m_GameStage = GM_FREEBALL_INIT
            SetFrames
          Else 'M_BILLARDS.FELLINPOCKETNUMBER(0) = FALSE/0
            'Now, we can check the other billards. If any of them fell into any pocket,
            'then the current player plays on, if not, the players swap
            For i = 1 To m_numBillards - 1
                If m_Billards.FellInPocketNumber(i) > 0 Then
                    bPlayerScored = True
                    Exit For 'loop varying i
                End If
            Next i
            m_Billards.ClearPocketNumbers
            If Not bPlayerScored Then
                m_CurrentPlayer = 3 - m_CurrentPlayer
                frmPool.lblPlayer = vbNullString
            End If
            'If the only billards left on the table are the cue-ball
            'and the 8'th ball, then initiate a special loop, that
            'enables the player to specify the pocket for the 8'th ball
            For i = 1 To m_numBillards - 1
                If i <> 8 Then
                    If m_Billards.InTheGame(i) Then
                        Exit For 'loop varying i
                    End If
                End If
            Next i

            'Hide the arrow
            m_Billboard.Visible False

            'Change the game's mode, from GM_BILLARDS_IN_MOTION to
            'the aiming mode or the 8'th ball pocket designation mode.
            If i >= m_numBillards Then
                m_GameStage = GM_8BALL_INIT
                'Show relevant messages
                m_Stickers(m_CurrentPlayer).Visible = True
                m_Stickers(SC_EIGHTBALL).Visible = True
              Else 'NOT I...
                m_GameStage = GM_PLAYERS_AIMING_INIT
                'Show relevant messages
                If Not bPlayerScored Then
                    m_Stickers(m_CurrentPlayer).Visible = True
                End If
                SetFrames
            End If
        End If
    End If

End Sub

Public Sub FireCueBall(ByVal Velocity As Single)

  'Sets the cue-ball's initial velocity vector

  Dim vDir As D3DVECTOR

    If m_GameStage = GM_PLAYERS_AIMING Then
        Do Until LightPower = 1
            Render
        Loop
        m_PreviousPlayer = m_CurrentPlayer
        For i = 1 To 2
            PrevShotCounts(i) = ShotCounts(i)
            PrevSinkCounts(i) = SinkCounts(i)
        Next i
        ShotCounts(m_CurrentPlayer) = ShotCounts(m_CurrentPlayer) + 1
        DisplayCounts
        D3DXVec3Subtract vDir, m_Billards.BillardPosition(0), m_Cameras(CamMoveable).Position
        vDir.y = 0
        D3DXVec3Normalize vDir, vDir
        With frmPool
            If (.mnuTgAfterHit.Checked Or Not InNormalMode) And m_ActiveCam = CamMoveable Then
                ToggleCameras
            End If
        End With 'FRMPOOL
        m_Billards.FireCueBall ScaleVector(vDir, Velocity)
        m_GameStage = GM_BILLARDS_IN_MOTION
        'Play cue-ball launch sound
        PlaySound m_DSCueBallLaunched, Velocity ^ 0.2
    End If

End Sub

Private Sub FreeBall()

  Dim vCueball As D3DVECTOR

    If m_Players(m_CurrentPlayer) = PC_HUMAN Then
        m_ActiveCam = CamFixed
    End If
    m_Billards.ReappearCueBall
    vCueball = m_Billards.BillardPosition(0)
    vCueball.y = m_BillRadius
    m_Billards.BillardPosition(0) = vCueball
    Do
        'wait for the user to place the cue-ball on the table.
        Render
    Loop While m_GameStage = GM_FREEBALL Or Down 'or Down for the rare case when the only balls left are
    '                                             the cue ball and the eight ball and the cue ball was sunk
    '                                             resulting in a free ball for the other player
    '                                             to enable the the 8-ball loop below - in particular the
    '                                             resulting game stage GM_8BALL_INIT - the mouse must be up
    With frmPool
        .mnuGameExit.Enabled = True
        EnableMenuItem GetSystemMenu(.hwnd, False), MenuCloseItem, MF_BYPOSITION 'enable main window sysmenu close
        SendMessage .hwnd, WM_NCPAINT, 1&, 0& 'repaint the frame and sysmenu
    End With 'FRMPOOL

    SetMousePosition

    'If the only billards left on the table are the cue-ball and
    'the 8'th ball, then initiate a special loop, that
    'enables the player to specify the pocket for the 8'th ball
    For i = 1 To m_numBillards - 1
        If i <> 8 Then
            If m_Billards.InTheGame(i) Then
                Exit For 'loop varying i
            End If
        End If
    Next i

    If i >= m_numBillards Then
        m_GameStage = GM_8BALL_INIT
      Else 'NOT I...
        m_GameStage = GM_PLAYERS_AIMING_INIT
    End If

End Sub

Private Sub FreeBallInit()

  Dim iFrames     As Long

    With frmPool
        'cannot close down now until the ball has been released
        .mnuGameExit.Enabled = False
        EnableMenuItem GetSystemMenu(.hwnd, False), MenuCloseItem, MF_BYPOSITION Or MF_GRAYED 'disable (gray) main window sysmenu close
        SendMessage .hwnd, WM_NCPAINT, 1&, 0& 'repaint the frame and sysmenu
        With .picViewport
            SetCursorPos frmPool.Left / 15 + .Left + .Width / 2, frmPool.Top / 15 + 44 + .Top + .Height / 2
        End With '.PICVIEWPORT
    End With 'FRMPOOL
    If m_ActiveCam = CamMoveable Then 'move the camera into the position of camera #2.
        With m_Cameras(CamMoveable)
            .StartTransitToFixed TransFrames
            Do
                If m_bNewGame Then
                    Exit Sub '---> Bottom
                End If
                .UpdateCamPos
                Render
            Loop While .CamInTransit
        End With 'M_CAMERAS(CAMMOVEABLE)
        m_ActiveCam = CamFixed
      Else 'wait a little 'NOT M_ACTIVECAM...
        For iFrames = 0 To 1.5 / g_dt
            If m_bNewGame Then
                Exit Sub '---> Bottom
            End If
            Render
        Next iFrames
    End If

    'Set the lblPlayer's caption and font size for FreeBall loop
    Info vbWhite, 12, "Move the  cue-ball  with  your mouse to desired position on the table and  release the ball with a left click"
    HideStickers
    m_GameStage = GM_FREEBALL
    m_ActiveCam = CamFixed

End Sub

Public Function GetLightDir() As D3DVECTOR4

  'Retrieves the light direction vector

    With GetLightDir
        .x = m_Lamp.Direction.x
        .y = m_Lamp.Direction.y
        .z = m_Lamp.Direction.z
        If m_Lamp.Type = D3DLIGHT_DIRECTIONAL Then
            .w = 0
          Else 'NOT M_LAMP.TYPE...
            .w = 1
        End If
    End With 'GETLIGHTDIR

End Function

Public Sub HideCrsr()

    Do 'show cursor
    Loop Until ShowCursor(False) < 0

End Sub

Private Sub HideStickers()

  'Make all stickers disappear

    For i = 1 To m_numStickers
        m_Stickers(i).Visible = False
    Next i

End Sub

Private Sub IdleBeginDetection(ByVal IdleMinutes As Long)

    BeginIdleDetection AddressOf IdleCallBack, IdleMinutes, 0&

End Sub

Private Sub IdleCallBack(ByVal dwState As Long)

    Select Case dwState
      Case USER_IDLE_BEGIN
        If m_Billards.AnyBillardInMotion Then
            IdleStopDetection
            LightsOn
          Else 'M_BILLARDS.ANYBILLARDINMOTION = FALSE/0
            m_Billboard.Visible m_Billboard.Showing, False
            LightsOff
        End If
      Case USER_IDLE_END
        m_Billboard.Visible m_Billboard.Showing
        LightsOn
    End Select

End Sub

Public Sub IdleStopDetection()

    EndIdleDetection 0&

End Sub

Private Sub Info(Color As Long, Size As Long, Text As String)

    With frmPool.lblPlayer
        .ForeColor = Color
        .FontSize = Size
        .Caption = Text
    End With 'FRMPOOL.LBLPLAYER

End Sub

Private Sub InitD3D(ByVal hWndViewport As Long)

  'Attempts to create Direct3DDevice8 object.

  Dim Mode As D3DDISPLAYMODE
  Dim D3DPP As D3DPRESENT_PARAMETERS

  'Create the D3D object

    Set m_D3D = m_DX.Direct3DCreate()

    'Get The current Display Mode format
    m_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode

    'Set up the structure used to create the D3DDevice.
    With D3DPP
        .Windowed = 1
        .SwapEffect = D3DSWAPEFFECT_FLIP 'D3DSWAPEFFECT_COPY_VSYNC
        .BackBufferFormat = Mode.Format
        .BackBufferCount = 1
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
    End With 'D3DPP

    'Create the D3DDevice
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'If you do not have hardware 3d acceleration enable the reference rasterizer
    'using the DirectX control panel and change D3DDEVTYPE_HAL to D3DDEVTYPE_REF
    Set m_D3DDevice = m_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWndViewport, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)
    'ie comment out the line above and un-comment the line below
    'Set m_D3DDevice = m_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, hWndViewport, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    'Device state would normally be set here
    'Turn off culling, so we see the front and back of the triangle
    With m_D3DDevice
        .SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
        .SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_DIFFUSE
        .SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_TEXTURE
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR

        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_DITHERENABLE, 1
        .SetRenderState D3DRS_FOGENABLE, 0
        .SetRenderState D3DRS_FILLMODE, 0
        .SetRenderState D3DRS_LASTPIXEL, 1
        .SetRenderState D3DRS_ZENABLE, 1
    End With 'M_D3DDEVICE

End Sub

Private Sub InitialLoop(ByVal Count As Long)

  ' this loops a little before the player ist taking aim, mainly to adjust the master clock

    With frmPool
        Do
            Count = Count - 1
            Render
        Loop While Count
        SetFrames
        .picViewport.Enabled = True
    End With 'FRMPOOL
    m_GameStage = GM_PLAYERS_AIMING_INIT

End Sub

Private Sub LightsOff()

    Do
        Render
        SetupLights LightPower
        LightPower = LightPower - g_dt / 3
        If LightPower < 0.1 Then
            LightPower = 0.1
            Exit Do 'loop 
        End If
    Loop

End Sub

Private Sub LightsOn()

    Do
        SetupLights LightPower
        Render
        LightPower = LightPower + g_dt / 3
    Loop While LightPower < 1
    LightPower = 1
    SetupLights 1
    IdleBeginDetection 5

End Sub

Public Sub Main()

  'Programm's entry point and main loop.

  Dim Dummy     As Single

    If App.PrevInstance Then
        For i = 1 To 40
            Beeper 2000 - 45 * i, 10
        Next i
        MsgBox "This Program makes extensive use of DirectX Graphics and running two or" & vbCrLf & _
               "more instances concurrently might slow down your computer considerably." & vbCrLf & vbCrLf & _
               "Please click on OK button or press Enter.", vbCritical Or vbSystemModal, App.ProductName & " (Aditional Program Instance)"
      Else 'APP.PREVINSTANCE = FALSE/0

        HideCrsr

        InitCommonControls

        'create the global objects objects
        Set m_DX = New DirectX8
        Set m_D3DX = New D3DX8
        Set m_Table = New clsTable
        Set m_Billards = New clsBillards
        Set m_Cameras(CamMoveable) = New clsCamera
        Set m_Cameras(CamFixed) = New clsCamera
        Set m_PCC = New clsCollisionController
        Set m_Billboard = New clsBillboard
        For i = 1 To m_numStickers
            Set m_Stickers(i) = New clsScreenSticker
        Next i

        'some math
        g_Pi = 4 * Atn(1)
        g_PiHalf = 2 * Atn(1)
        g_2Pi = 2 * g_Pi

        'high speed counter freq
        QueryPerformanceFrequency PerfFreq
        SloMoFactor = 1

        g_dt = 0.03 'initial guess
        TransFrames = TransTime / g_dt
        InNormalMode = True

        Randomize

        'The viewport
        frmPool.picViewport.ScaleMode = vbPixels

        SetViewportSize

        'Initialize D3D and D3DDevice
        InitD3D frmPool.picViewport.hwnd

        'Clear the rendering surface and the Z-buffer then set the viewport
        With m_D3DDevice
            .Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1, 0
            .SetViewport Viewport
        End With 'M_D3DDEVICE
        SetupLights 0

        'Setup sounds
        SetupSounds

        'Create the cameras
        'The first camera will move around the table
        m_Cameras(CamMoveable).Setup

        'The second camera will hang from the ceiling, pointing downwards
        m_Cameras(CamFixed).Setup

        'Initially the active camera is the fixed one
        m_ActiveCam = CamFixed

        'Create the table
        m_TableWidth = 2.25
        m_TableLength = 4
        With m_Table
            .Create m_TableWidth, m_TableLength, m_D3DDevice, m_D3DX
            'After creating the table "ask" it for the position and radius of its pockets.
            'We will need these data for selecting the pocket for the eight-ball
            'at the end of the game.
            .GetPockets m_Pockets, m_PocketRadius
            'We also need to know, how large the game area really is.
            .GetGameArea m_GameAreaMinZ, m_GameAreaMaxZ, m_GameAreaMinX, m_GameAreaMaxX
        End With 'M_TABLE

        'Create the billards
        With m_Billards
            .Create m_D3DDevice, m_D3DX
            'After creating the billards, we can "ask" them for their number...
            m_numBillards = .NumBillards
        End With 'M_BILLARDS
        '...so that we can set the dimensions of the array holding indexes of pockets,
        'into which the billards fell
        ReDim m_BPocketNumbers(0 To m_numBillards - 1)

        'Create the collision controller
        m_PCC.Setup m_Billards, m_Table

        'Create the billboard with an arrow indicating the pocket
        'for the 8'th ball
        On Error Resume Next
            m_Billboard.Setup App.Path & ArrowDDS, 0.2, 0.2, m_D3DDevice, m_D3DX

            'Create the D3DXSprite object to enable rendering screen stickers.
            Set m_Sprite = m_D3DX.CreateSprite(m_D3DDevice)
            With Viewport
                'Create the stickers
                m_Stickers(SC_PLAYER1).Setup App.Path & Player1DDS, D3DFMT_A8R8G8B8, (.Width - 305) / 2, (.Height - 84) / 2, 305, 84, &HFFFFFFFF, m_D3DDevice, m_D3DX
                m_Stickers(SC_PLAYER2).Setup App.Path & Player2DDS, D3DFMT_A8R8G8B8, (.Width - 316) / 2, (.Height - 84) / 2, 316, 84, &HFFFFFFFF, m_D3DDevice, m_D3DX
                m_Stickers(SC_FREEBALL).Setup App.Path & FreeBallDDS, D3DFMT_A8R8G8B8, (.Width - 335) / 2, (.Height + 70) / 2, 335, 65, &HFFFFFFFF, m_D3DDevice, m_D3DX
                m_Stickers(SC_EIGHTBALL).Setup App.Path & EightBallDDS, D3DFMT_A8R8G8B8, (.Width - 355) / 2, (.Height + 90) / 2, 355, 83, &HFFFFFFFF, m_D3DDevice, m_D3DX
                m_Stickers(SC_WINS).Setup App.Path & WinsDDS, D3DFMT_A8R8G8B8, (.Width - 176) / 2, (.Height + 80) / 2, 176, 64, &HFFFFFFFF, m_D3DDevice, m_D3DX
                m_Stickers(SC_BEGINS).Setup App.Path & BeginsDDS, D3DFMT_A8R8G8B8, (.Width - 255) / 2, (.Height + 90) / 2, 255, 83, &HFFFFFFFF, m_D3DDevice, m_D3DX
                m_Stickers(SC_FOUL).Setup App.Path & FoulDDS, D3DFMT_A8R8G8B8, (.Width - 170) / 2, 20, 255, 83, &HFFFFFFFF, m_D3DDevice, m_D3DX
            End With 'VIEWPORT
        On Error GoTo 0

        m_Billards.GetPhysBillConstants Dummy, Dummy, m_BillRadius, Dummy
        Sleep 2000
        'Show the form
        With frmPool
            hDCViewport = .picViewport.hDC
            .lbTitle(0).Visible = False
            .lbTitle(1).Visible = False
            .lblLoading.Visible = False
            DoEvents
            If .mnuGameNew.Enabled Then
                StartNewGame True
                DoEvents
                'Main game loop
                Do
                    Select Case m_GameStage
                      Case GM_INITIAL_LOOP
                        InitialLoop 1 / g_dt
                        .mnuTg.Enabled = True
                      Case GM_BILLARDS_IN_MOTION
                        ShowShootingControls False
                        UpdateBillards
                      Case GM_PLAYERS_AIMING_INIT
                        PlayerTakingAimInit m_Players(m_CurrentPlayer)
                        ShowShootingControls True
                        .mnuSloMo.Enabled = True
                      Case GM_PLAYERS_AIMING
                        .mnuGameNew.Enabled = True
                        AimingLine = .mnuAiming.Checked
                        PlayersTakingAim m_Players(m_CurrentPlayer)
                        AimingLine = False
                      Case GM_FREEBALL_INIT
                        .mnuTg.Enabled = False
                        FreeBallInit
                      Case GM_FREEBALL
                        FreeBall
                        .mnuTg.Enabled = True
                      Case GM_8BALL_INIT
                        .mnuTg.Enabled = False
                        Set8BallTargetInit
                      Case GM_8BALL
                        Set8BallTarget
                        .mnuTg.Enabled = True
                      Case GM_SHOWING_RESULTS_INIT
                        .mnuTg.Enabled = False
                        ShowResultsInit
                      Case GM_SHOWING_RESULTS
                        ShowResults
                    End Select
                Loop
            End If
        End With 'FRMPOOL
    End If

End Sub

Private Function MakeBGR(Color As OLE_COLOR) As OLE_COLOR

    MakeBGR = RGB((Color And vbBlue) / &H10000, (Color And vbGreen) / &H100&, Color And vbRed)

End Function

Public Function MakeColor(ByVal r As Single, ByVal g As Single, ByVal b As Single, Optional ByVal a As Single = 0) As D3DCOLORVALUE

    With MakeColor
        .r = r
        .g = g
        .b = b
        .a = a
    End With 'MakeColor

End Function

Public Function MakeVector(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR

    With MakeVector
        .x = x
        .y = y
        .z = z
    End With 'MakeVector

End Function

Public Sub MouseKeyboardEventHandler(ByVal Button As Integer, ByVal MouseX As Single, ByVal MouseY As Single, ByVal MouseMoveX As Single, ByVal MouseMoveY As Single)

  'Acts as a middle-man between the user and objects (mainly cameras) in the game.

  Dim vCueball      As D3DVECTOR        'Position vector of the cue-ball.
  Dim vCamPos       As D3DVECTOR        'Position vector of the active camera.
  Dim vDist         As D3DVECTOR        'A vector linking the camera with the cue ball
  Dim mProj         As D3DMATRIX        'Projection transformation matrix of the active camera (used for "unprojection")
  Dim mView         As D3DMATRIX        'View transformation matrix of the active camera (same as sbove)
  Dim mWorld        As D3DMATRIX        'World transformation matrix of the active camera (again same as above)
  Dim Lambda        As Single           'A utility variable used for scaling vectors.
  Dim CBDist        As Single           'Distance between cam and cue ball
  Dim mCamGen       As D3DMATRIX        'Camera generator matrix.
  Dim v8BallTarget  As D3DVECTOR        'The result of unprojecting the mouse cursor coordinates

    Select Case m_GameStage
      Case GM_PLAYERS_AIMING
        'If a human player is aiming the camera should rotate around the cue-ball
        'and translate along a line that runs from the cue-ball to the camera
        'to aid aiming. However, if the player is CPU driven, the camera should
        'move in exactly the same way as it does, when the billards are in motion.
        MouseMoveY = MouseMoveY * 4
        MouseMoveX = MouseMoveX * 4
        If m_ActiveCam = CamMoveable Then
            Select Case Button
              Case vbLeftButton
                With m_Cameras(CamMoveable)
                    If m_Players(m_CurrentPlayer) = PC_CPU Then
                        vCamPos = .Position
                        .PivotCam vCamPos, MouseMoveY, MouseMoveX
                      Else 'NOT M_PLAYERS(M_CURRENTPLAYER)...
                        vCueball = m_Billards.BillardPosition(0)
                        If MouseY < 385 Then 'above top of cue ball
                            .PivotCam vCueball, MouseMoveY, -MouseMoveX
                          Else 'NOT MOUSEY...
                            .PivotCam vCueball, MouseMoveY, MouseMoveX
                        End If
                    End If
                End With 'M_CAMERAS(M_ACTIVECAM) 'M_CAMERAS(CAMMOVEABLE)
              Case vbRightButton
                mCamGen = m_Cameras(CamMoveable).Generators
                With mCamGen
                    If m_Players(m_CurrentPlayer) = PC_CPU Then
                        m_Cameras(CamMoveable).MoveCam MakeVector(-.m11 * MouseMoveX - .m31 * MouseMoveY, -.m12 * MouseMoveX, -.m13 * MouseMoveX - .m33 * MouseMoveY)
                      Else 'NOT M_PLAYERS(M_CURRENTPLAYER)...
                        vCueball = m_Billards.BillardPosition(0)
                        vCamPos = m_Cameras(CamMoveable).Position
                        D3DXVec3Subtract vDist, vCamPos, vCueball
                        CBDist = D3DXVec3Length(vDist) 'distance between cam and cue ball
                        If Abs(MouseMoveY) > 0.015 Then 'prevent overshooting
                            MouseMoveY = 0.015 * Sgn(MouseMoveY)
                        End If
                        If (CBDist > 0.35 Or MouseMoveY > 0) And (CBDist < 5 Or MouseMoveY < 0) Then
                            m_Cameras(CamMoveable).MoveCam MakeVector(-.m31 * MouseMoveY, -.m32 * MouseMoveY, -.m33 * MouseMoveY)
                        End If
                    End If
                End With 'MCAMGEN
            End Select
        End If
      Case GM_FREEBALL
        'Free-ball is relativelly simple. There are no problems with the cameras
        'as the only active camera allowed in this game mode is the fixed Camera
        'Mouse cursor position has a different role now - it indicates the spot on the table's
        'surface where the user whants to place the cue-ball.
        Select Case Button
          Case 0
            'Button = 0 means that no button is pressed. This in turn means that the user is moving the
            'cue-ball around. Make sure that it stays within the table and does not "hit" any other billard.
            'Get the camera's position
            vCamPos = m_Cameras(m_ActiveCam).Position
            'Unproject the point specified by MouseX and MouseY
            vVctr1 = MakeVector(MouseX, MouseY, 1)    'Z = 1 as we want the point to be placed at the back of the viewport.
            mView = m_Cameras(m_ActiveCam).ViewMatrix
            mProj = m_Cameras(m_ActiveCam).ProjectionMatrix
            D3DXMatrixIdentity mWorld
            D3DXVec3Unproject vVctr1, vVctr1, Viewport, mProj, mView, mWorld
            'vVctr1 is now the point (MouseX, MouseY, 1) in 3D space.
            'Now all we need to do is run a line through this point and the camera position vector
            'and find the spot where this line intersects with the table plane
            '(i.e. when y = 0, though not exactly - explained later)
            D3DXVec3Subtract vVctr2, vVctr1, vCamPos
            Lambda = (m_BillRadius - vCamPos.y) / vVctr2.y 'the billards'radius.
            D3DXVec3Add vCueball, vCamPos, ScaleVector(vVctr2, Lambda) 'Cue-ball's new position.
            'Just to be sure that the cue-ball is on exactly the same level as other billards
            vCueball.y = m_BillRadius 'Without this the cue-ball tends to end up slightly below, or above
            'This would have devastating consequences for collision response.
            'Now, check whether the new cue-ball position fits within the game area.
            If vCueball.x > m_GameAreaMinX + m_BillRadius And vCueball.x < m_GameAreaMaxX - m_BillRadius And vCueball.z > m_GameAreaMinZ + m_BillRadius And vCueball.z < m_GameAreaMaxZ - m_BillRadius Then
                'Now, check for any overlapping between the cue-ball and other billards
                For i = 1 To m_numBillards - 1
                    D3DXVec3Subtract vVctr1, m_Billards.BillardPosition(i), vCueball
                    'If a billard overlapps with the cue-ball set CueBallCleared to false
                    If vVctr1.x * vVctr1.x + vVctr1.z * vVctr1.z < 0.01 Then
                        Exit For 'loop varying i
                    End If
                Next i
                'The new cue-ball position can be applied to the cue-ball only if the billard is cleared
                If i = m_numBillards Then
                    m_Billards.BillardPosition(0) = vCueball
                End If
            End If
          Case vbLeftButton
            'When the left mouse button is pressed the user wants to release the cue ball
            frmPool.lblPlayer = vbNullString 'so erase the hint...
            m_GameStage = GM_PLAYERS_AIMING_INIT '...and proceed to next game stage
        End Select
      Case GM_8BALL
        'In this mode the user is expected to point, with the mouse cursor, at the pocket,
        'they want to shoot the eight-ball into.
        If Button = vbLeftButton Then 'a pocket has been chosen.
            'Get the camera's position.
            vCamPos = m_Cameras(m_ActiveCam).Position
            'Unproject the coordinates specified by MouseX and MouseY
            vVctr1 = MakeVector(MouseX, MouseY, 1)
            mView = m_Cameras(m_ActiveCam).ViewMatrix
            mProj = m_Cameras(m_ActiveCam).ProjectionMatrix
            D3DXMatrixIdentity mWorld
            D3DXVec3Unproject vVctr1, vVctr1, Viewport, mProj, mView, mWorld
            D3DXVec3Subtract vVctr2, vVctr1, vCamPos
            Lambda = -vCamPos.y / vVctr2.y        'This time we use y = 0.
            D3DXVec3Add v8BallTarget, vCamPos, ScaleVector(vVctr2, Lambda)
            'If a distance between a pocket and the v8BallTarget point is smaller
            'than the radius of a pocket then we've found the pocket indicated by the user.
            For i = 1 To 6
                D3DXVec3Subtract vDist, m_Pockets(i), v8BallTarget
                If vDist.x * vDist.x + vDist.z * vDist.z <= m_PocketRadius * m_PocketRadius Then
                    m_Designated8BallPocket = i
                    m_Billboard.BasePoint = MakeVector(m_Pockets(i).x, 0.1, m_Pockets(i).z)
                    'We found the right pocket, thus set the current game mode to aiming mode.
                    m_GameStage = GM_PLAYERS_AIMING_INIT
                    frmPool.lblPlayer = vbNullString
                    Exit For 'loop varying i
                End If
            Next i
        End If
    End Select

End Sub

Private Sub PlayersTakingAim(ByVal pc As PLAYER_CONTROL)

    If pc = PC_CPU Then

        'Code for computer player aiming goes here

      Else 'NOT PC...
        SetMousePosition
        ShowCrsr
        Do
            'Do nothing. Just wait for user's aiming input.
            m_Cameras(m_ActiveCam).UpdateCamPos
            Render
        Loop While m_GameStage = GM_PLAYERS_AIMING
    End If
    frmPool.SetCursorIcons vbArrow, vbSizePointer, vbSizePointer

End Sub

Private Sub PlayerTakingAimInit(ByVal pc As PLAYER_CONTROL)

  'Sets the scene before calling PlayerTakingAim.

  Dim vCueball    As D3DVECTOR
  Dim vCamPos     As D3DVECTOR
  Dim iFrames     As Long

    m_bNewGame = False
    If pc = PC_CPU Then

        'Code for computer player aiming goes here

      Else 'NOT PC...
        'The first camera should be placed "behind" the cue-ball
        vCueball = m_Billards.BillardPosition(0)
        vCueball.y = 0 'nomalize with y = zero
        D3DXVec3Normalize vCamPos, vCueball
        vCueball.y = m_Billards.Diameter 'the camera's line of sight shouldn't go through the bottom of the cue-ball but rather through its top
        D3DXVec3Add vCamPos, vCueball, ScaleVector(vCamPos, AimingElevation * 1.5)
        vCamPos.y = AimingElevation
        'prep camera's transition loop
        If m_ActiveCam = CamFixed Then 'switch to moveable cam first
            m_Cameras(CamMoveable).ChangeView
            m_ActiveCam = CamMoveable
        End If

        With m_Cameras(CamMoveable)
            .StartTransit vCamPos, vCueball, MakeVector(0, 2, 0), TransFrames
            Do
                If m_bNewGame Then
                    Exit Sub '---> Bottom
                End If
                .UpdateCamPos
                Render
            Loop While .CamInTransit

        End With 'M_CAMERAS(CAMMOVEABLE)

        'Cursors for aiming
        frmPool.SetCursorIcons vbArrow, vbSizePointer, vbSizeNS
    End If

    'Set the correct caption and font size for the lblPlayer label
    Info IIf(m_CurrentPlayer = 1, &HD0D0&, &HFF80FF), 24, "Player " & m_CurrentPlayer
    HideStickers
    m_GameStage = GM_PLAYERS_AIMING

End Sub

Public Sub PlaySound(ByVal dsBuffer As DirectSoundSecondaryBuffer8, ByVal Volume As Single)

  Dim CurrentVolume As Long

    If Not (dsBuffer Is Nothing) And m_VolumeBase > m_VolumeMute Then 'sound track present and not muted
        CurrentVolume = Volume * m_VolumeBase - m_VolumeMax
        Select Case CurrentVolume 'limit volume
          Case Is > 0
            CurrentVolume = 0
          Case Is < -m_VolumeMax
            CurrentVolume = -m_VolumeMax
        End Select
        With dsBuffer
            If .GetStatus <> DSBSTATUS_PLAYING Then 'currently no sound
                PreviousVolume = -m_VolumeMax - 1 'reset so that it accepts this volume
            End If
            If CurrentVolume > PreviousVolume Then 'this sound is louder than sound currently playing
                .Stop 'stop current sound
                .SetCurrentPosition 0 'set sound track to start
                .SetVolume CurrentVolume 'volume -10000 .. 0
                .Play DSBPLAY_DEFAULT  'play
                PreviousVolume = CurrentVolume 'and remember current volune
            End If
        End With 'DSBUFFER
    End If

End Sub

Public Sub Render()

  'Draws the scene

  Dim mView           As D3DMATRIX
  Dim mProj           As D3DMATRIX
  Dim x               As Long
  Dim CurrTick        As Currency

    If Not m_D3DDevice Is Nothing Then
        mView = m_Cameras(m_ActiveCam).ViewMatrix
        mProj = m_Cameras(m_ActiveCam).ProjectionMatrix
        m_D3DDevice.SetTransform D3DTS_VIEW, mView
        m_D3DDevice.SetTransform D3DTS_PROJECTION, mProj

        'Clear the backbuffer to picViewport.ForeColor, clear the Z buffer too
        m_D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, MakeBGR(frmPool.picViewport.ForeColor), 1, 0

        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        m_D3DDevice.BeginScene

        'The table...
        m_Table.RenderTable
        '...and the billards
        m_Billards.RenderBillards
        'Sprites
        m_Sprite.Begin
        For j = 1 To m_numStickers
            With m_Stickers(j)
                If .Visible Then
                    .Draw m_Sprite
                End If
            End With 'M_STICKERS(J)
        Next j
        m_Sprite.End
        'The arrow billboard
        If m_ActiveCam = CamMoveable Then
            vVctr1 = MakeVector(0, 1, 0)
          Else 'NOT M_ACTIVECAM...
            vVctr1 = m_Billboard.BasePoint
        End If
        m_Billboard.RenderBillboard m_Cameras(m_ActiveCam).Position, vVctr1

        m_D3DDevice.EndScene
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        'Present the backbuffer contents to the front buffer (screen)
        m_D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        With frmPool
            If m_ActiveCam = CamMoveable And AimingLine Then 'draw aiming line
                With .picViewport
                    x = .ScaleWidth / 2
                    SetPenPosition hDCViewport, x, 0, ByVal 0
                    DrawLine hDCViewport, x, .ScaleHeight / 2 - 8 'down to a little above the cueball
                End With '.PICVIEWPORT
            End If
            Do
                DoEvents
            Loop While .WindowState = vbMinimized Or .mnuFreeze.Checked

            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            'internal time raster and variable master clock
            x = 0
            Do
                QueryPerformanceCounter CurrTick
                x = x + 1 'count idle time
            Loop While CurrTick < NextTick 'wait until time raster ticks again
            If SloMoFactor = 1 Then 'not in slow motion mode
                Select Case x
                  Case Is < 2 'if idle time is to small
                    g_dt = g_dt * 1.01 'then increase g_dt by 1 %
                  Case Is > 2 'if idle time is to big
                    g_dt = g_dt / 1.01 'then decrease g_dt by 1 %
                End Select
            End If
            NextTick = CurrTick + PerfFreq * g_dt * SloMoFactor 'adjust raster time for next cycle
            'adjust number of transit frames
            TransFrames = TransTime / g_dt
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            'compute frame rate
            If CurrTick - FPSLastCheck >= PerfFreq Then
                .lblFpS = FPSCount & " FpS"
                FPSCount = 0
                FPSLastCheck = CurrTick
            End If
            FPSCount = FPSCount + 1
        End With 'FRMPOOL
    End If

End Sub

Public Function ScaleVector(v As D3DVECTOR, ScaleFactor As Single) As D3DVECTOR

    With ScaleVector
        .x = v.x * ScaleFactor
        .y = v.y * ScaleFactor
        .z = v.z * ScaleFactor
    End With 'ScaleVector

End Function

Private Sub Set8BallTarget()

  'Allows the players to pick the pocket for the eight-ball

    ShowCrsr
    If m_Players(m_CurrentPlayer) = PC_HUMAN Then
        Do
            'This loop will go on and on until the user marks the pocket for the 8'th ball.
            '            m_Cameras(m_ActiveCam).UpdateCamPos
            Render
        Loop While m_GameStage = GM_8BALL
      Else 'NOT M_PLAYERS(M_CURRENTPLAYER)...
        'the computer chooses the pocket
    End If
    If m_GameStage = GM_PLAYERS_AIMING_INIT Then
        m_Billboard.Visible True
        m_PCC.Sinking8Illegal = False 'tell collision controller that sinking 8-ball is now legal
        HideStickers 'and hide the stickers
        HideCrsr
    End If

End Sub

Private Sub Set8BallTargetInit()

  'Prepares the scene for the Set8BallTarget loop

  Dim iFrames As Long

  'Show relevant messages

    m_Stickers(m_CurrentPlayer).Visible = True
    m_Stickers(SC_EIGHTBALL).Visible = True

    'Disable camera toggling for the transition loop
    If m_ActiveCam = CamMoveable Then 'move the camera into the position of camera #2.
        With m_Cameras(CamMoveable)
            .StartTransitToFixed TransFrames
            Do
                If m_bNewGame Then
                    Exit Sub '---> Bottom
                End If
                .UpdateCamPos
                Render
            Loop While .CamInTransit
        End With 'M_CAMERAS(CAMMOVEABLE)

    End If

    'Set the lblPlayer's caption and font size for Set8BallTarget loop
    Info vbRed, 12, "Left-click on the pocket you want to shoot the eight-ball into"

    'The Set8BallTarget loop requires that the player sees the table from above
    m_ActiveCam = CamFixed
    'Set the new game mode
    m_GameStage = GM_8BALL
    'Change the cursors on the form.
    frmPool.SetCursorIcons vbArrow, vbArrow, vbArrow

End Sub

Private Sub SetFrames()

    With frmPool
        .lblPl(m_CurrentPlayer - 1).BorderStyle = vbFixedSingle
        .lblPl(1 Xor (m_CurrentPlayer - 1)).BorderStyle = vbBSNone
    End With 'FRMPOOL

End Sub

Public Sub SetMousePosition()

    With frmPool.picViewport
        SetCursorPos frmPool.Left / 15 + .Left + .Width / 2, frmPool.Top / 15 + 44 + .Top + .Height * 0.9
    End With 'FRMPOOL.PICVIEWPORT

End Sub

Public Sub SetSoundVolumeBase(ByVal VolBase As Single)

    m_VolumeBase = VolBase

End Sub

Public Sub SetToPrevPosn()

    AimingLine = False
    m_Billards.PrevPosns
    m_CurrentPlayer = m_PreviousPlayer
    m_Stickers(m_CurrentPlayer).Visible = True
    For i = 1 To 2
        ShotCounts(i) = PrevShotCounts(i)
        SinkCounts(i) = PrevSinkCounts(i)
    Next i
    DisplayCounts
    SetFrames
    m_Designated8BallPocket = 0
    m_Billboard.Visible False
    SetMousePosition
    PlayerTakingAimInit PC_HUMAN
    AimingLine = frmPool.mnuAiming.Checked

End Sub

Private Sub SetupLights(Brightness As Single)

  'Introduces a source of directional light to the scene and sets it's brightness.

    With m_Lamp  'We have one lamp above the table
        .Ambient = MakeColor(0.05 * Brightness, 0.05 * Brightness, 0.1 * Brightness, Brightness)
        .Type = D3DLIGHT_DIRECTIONAL
        .diffuse = MakeColor(Brightness, Brightness, Brightness, Brightness)
        .specular = .diffuse
        .Direction = MakeVector(0, -3, 0.6)
    End With 'M_LAMP

    With m_D3DDevice
        .SetLight 0, m_Lamp
        .LightEnable 0, 1
        .SetRenderState D3DRS_LIGHTING, 1
        .SetRenderState D3DRS_SPECULARENABLE, 1
        .SetRenderState D3DRS_AMBIENT, D3DColorRGBA(100, 100, 100, 0)
    End With 'M_D3DDEVICE

End Sub

Private Sub SetupSounds()

  'Loads sound files into buffers

  Dim dsBufDesc As DSBUFFERDESC

  'Create a default DirectSound object

    Set m_DirectSound8 = m_DX.DirectSoundCreate(vbNullString)

    'Set the cooperation level
    m_DirectSound8.SetCooperativeLevel frmPool.hwnd, DSSCL_PRIORITY

    'Create and fill in the buffer description structure...
    dsBufDesc.lFlags = DSBCAPS_CTRLVOLUME

    'Create the sound buffers from ".wav" files. If any file is missing, just skip it
    On Error Resume Next
        Set m_DSOops = m_DirectSound8.CreateSoundBufferFromFile(App.Path & OopsWAV, dsBufDesc)
        Set m_DSBillBillHit = m_DirectSound8.CreateSoundBufferFromFile(App.Path & BallBallHitWAV, dsBufDesc)
        Set m_DSBillTableHit = m_DirectSound8.CreateSoundBufferFromFile(App.Path & BallTableHitWAV, dsBufDesc)
        Set m_DSPocketHit = m_DirectSound8.CreateSoundBufferFromFile(App.Path & PocketHitWAV, dsBufDesc)
        Set m_DSCueBallLaunched = m_DirectSound8.CreateSoundBufferFromFile(App.Path & CueBallLaunchWAV, dsBufDesc)
        Set m_DSTransitionLong = m_DirectSound8.CreateSoundBufferFromFile(App.Path & TransitionLongWAV, dsBufDesc)
        Set m_DSClap = m_DirectSound8.CreateSoundBufferFromFile(App.Path & ClapWAV, dsBufDesc)
    On Error GoTo 0

    'set initial volume
    m_VolumeBase = m_VolumeMedium

End Sub

Public Sub SetViewportSize()

    With Viewport
        'size...
        .Height = frmPool.picViewport.ScaleHeight
        .Width = frmPool.picViewport.ScaleWidth
        '...visibility...
        .MaxZ = 1
        .MinZ = 0
        '...and origin
        .x = 0
        .y = 0
    End With 'VIEWPORT

End Sub

Public Sub ShowCrsr()

    Do 'show cursor
    Loop Until ShowCursor(True) >= 0

End Sub

Private Sub ShowResults()

    m_bNewGame = False
    LightsOff
    IdleStopDetection
    Do
        Render
    Loop Until m_bNewGame
    LightsOn

End Sub

Private Sub ShowResultsInit()

  'Prepares the scene for showing game results

    m_Billboard.Visible False
    m_Stickers(SC_FOUL).Visible = False
    m_bNewGame = False
    If m_ActiveCam = CamMoveable Then
        With m_Cameras(CamMoveable)
            .StartTransitToFixed TransFrames
            Do
                If m_bNewGame Then
                    Exit Sub '---> Bottom
                End If
                .UpdateCamPos
                Render
            Loop While .CamInTransit
        End With 'M_CAMERAS(CAMMOVEABLE)
    End If
    m_ActiveCam = CamFixed
    frmPool.SetCursorIcons vbNormal, vbNormal, vbNormal
    ShowCrsr

    'Set the lblPlayer's caption and font size for ShowResults loop
    Info vbCyan, 12, "Press F2 to start a new game or F3 to quit"

    'Set the current game mode to GM_SHOWING_RESULTS
    m_GameStage = GM_SHOWING_RESULTS

End Sub

Public Sub ShowShootingControls(Visible As Boolean)

    With frmPool
        .lblShoot.Visible = Visible
        .shpShoot.Visible = Visible
        .picPower.Visible = Visible
        .mnuJoy.Enabled = Visible And JoyPresent
    End With 'FRMPOOL

End Sub

Public Sub StartNewGame(FirstGame As Boolean)

    HideCrsr
    LightsOn
    m_Players(1) = PC_HUMAN
    m_Players(2) = PC_HUMAN 'once computer player is ready player 1 or player 2 could optionally be the computer
    For i = 1 To 2
        ShotCounts(i) = 0
        SinkCounts(i) = 0
    Next i
    m_CurrentPlayer = 2
    DisplayCounts
    m_CurrentPlayer = 1
    DisplayCounts
    m_GameStage = GM_INITIAL_LOOP
    m_bNewGame = True
    frmPool.lblPlayer = vbNullString
    HideStickers
    m_Billboard.Visible False
    AimingLine = False
    SloMoFactor = 1
    FastDecelerate = False

    With frmPool
        .mnuSloMo.Checked = False
        .mnuSloMo.Enabled = False
        .lblSloMo.Visible = False
        .lblFrozen.Visible = False
        .mnuGameNew.Enabled = False
        .mnuUndo.Enabled = False
        InNormalMode = Not .mnuCheat.Checked
        'aiming line
        ShowShootingControls False
        .lblShoot = Shoot
        .picPower.Cls
        .ShootingPower = 0
        ClearFrames
        'Set the cursors for GM_PLAYERS_AIMING_INIT mode
        .SetCursorIcons vbArrow, vbSizePointer, vbSizeNS
    End With 'FRMPOOL

    If m_ActiveCam = CamMoveable Then
        With m_Cameras(CamMoveable)
            .StartTransitToFixed TransFrames
            Do While .CamInTransit
                .UpdateCamPos
                Render
            Loop
        End With 'M_CAMERAS(CAMMOVEABLE)
    End If

    m_ActiveCam = CamFixed
    'Put the billards on the table
    m_Billards.InitialPositions FirstGame
    'Show the startup stickers
    m_Stickers(SC_PLAYER1).Visible = True
    m_Stickers(SC_BEGINS).Visible = True
    m_Designated8BallPocket = 0 'ULLI: this stmnt was origianlly missing
    m_PCC.Sinking8Illegal = True 'tell collision controller that sinking the 8-ball is illegal

End Sub

Public Sub StopSound(ByVal dsBuffer As DirectSoundSecondaryBuffer8)

    dsBuffer.Stop

End Sub

Public Sub ToggleCameras()

  'Toggles between available cameras.

    m_ActiveCam = m_ActiveCam Xor (CamMoveable Or CamFixed)
    With frmPool
        If m_ActiveCam = CamMoveable Then
            .SetCursorIcons vbCustom, vbSizePointer, vbSizeNS
            ShowShootingControls (m_GameStage = GM_PLAYERS_AIMING)
          Else 'NOT M_ACTIVECAM...
            .SetCursorIcons
            ShowShootingControls False
        End If
    End With 'FRMPOOL

End Sub

Private Sub UpdateBillards()

  'Runs collisions tests for billards and makes appropriate noises.

    m_bNewGame = False
    HideCrsr
    Do While m_Billards.AnyBillardInMotion
        If m_bNewGame Then
            Exit Sub '---> Bottom
        End If
        With m_PCC
            .CollisionDetection
            'Play collision sounds
            If .BillBillCollisionMomentum > 0 Then
                PlaySound m_DSBillBillHit, (.BillBillCollisionMomentum ^ 0.2) * 1.6
            End If
            If .BillTableCollisionMomentum > 0 Then
                PlaySound m_DSBillTableHit, (.BillTableCollisionMomentum ^ 0.2) * 1.3
            End If
            If .PocketHitDetected Then
                PlaySound m_DSPocketHit, 1
            End If
            If .CheatTimeOut Then 'timeout cheat-mode
                InNormalMode = True
            End If
        End With 'M_PCC
        m_Billards.NextFrame
        Render
    Loop

    'After the billards have stopped, enable the undo menu item, test the winning/loosing conditions
    'and set the initial value for the frames counter
    frmPool.mnuUndo.Enabled = True
    EvalLastStrike

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-29 14:15)  Decl: 223  Code: 1340  Total: 1563 Lines
':) CommentOnly: 188 (12%)  Commented: 130 (8,3%)  Empty: 269 (17,2%)  Max Logic Depth: 7
