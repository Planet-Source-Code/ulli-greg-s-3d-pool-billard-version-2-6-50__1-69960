VERSION 5.00
Begin VB.Form frmPool 
   Appearance      =   0  '2D
   BackColor       =   &H00004000&
   Caption         =   "Greg's and Ulli's 3D Pool Game"
   ClientHeight    =   8025
   ClientLeft      =   960
   ClientTop       =   2640
   ClientWidth     =   10995
   Icon            =   "Pool.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   535
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   733
   Begin VB.PictureBox picViewport 
      Align           =   1  'Oben ausrichten
      Appearance      =   0  '2D
      BackColor       =   &H00004000&
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      DrawStyle       =   2  'Punkt
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   0
      MouseIcon       =   "Pool.frx":08CA
      MousePointer    =   99  'Benutzerdefiniert
      ScaleHeight     =   425
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   733
      TabIndex        =   0
      Top             =   0
      Width           =   10995
      Begin VB.Label lblGoodBye 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Good bye and come back soon..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Old English Text MT"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   510
         Left            =   3045
         TabIndex        =   10
         Top             =   5685
         Visible         =   0   'False
         Width           =   5805
      End
      Begin VB.Image imgCW 
         Enabled         =   0   'False
         Height          =   1290
         Left            =   4500
         Picture         =   "Pool.frx":1194
         Top             =   2415
         Width           =   1620
      End
      Begin VB.Label lbTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Old English Text MT"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   765
         Index           =   1
         Left            =   2040
         TabIndex        =   6
         Top             =   1035
         Width           =   165
      End
      Begin VB.Label lbTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Old English Text MT"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   0
         Left            =   1995
         TabIndex        =   5
         Top             =   975
         Width           =   165
      End
   End
   Begin VB.PictureBox picPower 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FF80&
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   1800
      ScaleHeight     =   15
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   100
      TabIndex        =   7
      ToolTipText     =   "Click to launch cue-ball"
      Top             =   6915
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Timer tmrPercent 
      Interval        =   30
      Left            =   285
      Top             =   6810
   End
   Begin VB.Label lblFrozen 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Frozen"
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   4605
      TabIndex        =   13
      Top             =   6585
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblFpS 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   10860
      TabIndex        =   12
      Top             =   7695
      Width           =   45
   End
   Begin VB.Label lblLoading 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   510
      Left            =   870
      TabIndex        =   11
      Top             =   6435
      Width           =   1635
   End
   Begin VB.Shape shpShoot 
      BackColor       =   &H00000000&
      BorderColor     =   &H00008000&
      BorderStyle     =   6  'Innen ausgefüllt
      Height          =   600
      Left            =   5010
      Shape           =   4  'Gerundetes Rechteck
      Top             =   6765
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lblPl 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   240
      Index           =   1
      Left            =   405
      TabIndex        =   9
      Top             =   7245
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblPl 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   0
      Left            =   405
      TabIndex        =   8
      Top             =   6825
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblSloMo 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Slow Motion"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   4605
      TabIndex        =   4
      Top             =   6435
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "© Grzegorz Holdys && Ulli"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4477
      TabIndex        =   3
      Top             =   7425
      Width           =   2040
   End
   Begin VB.Label lblShoot 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   " Shoot "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   480
      Left            =   4995
      TabIndex        =   2
      ToolTipText     =   "Click & hold down; then release to launch the cue-ball"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label lblPlayer 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   480
      Left            =   7935
      TabIndex        =   1
      ToolTipText     =   "Current player"
      Top             =   6840
      Width           =   3360
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGamebar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "&Over        "
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuStgs 
      Caption         =   "&Options"
      Begin VB.Menu mnuTg 
         Caption         =   "&Toggle Cameras"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSV 
         Caption         =   "&Sound Volume"
         Begin VB.Menu mnuSVMute 
            Caption         =   "&Off"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuSVLow 
            Caption         =   "&Low"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuSVMedium 
            Caption         =   "&Medium"
            Checked         =   -1  'True
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuSVHigh 
            Caption         =   "&High"
            Shortcut        =   {F8}
         End
      End
      Begin VB.Menu mnuSloMo 
         Caption         =   "Slow &Motion"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuTgAfterHit 
         Caption         =   "&Full view after shot"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuAiming 
         Caption         =   "&Aiming Line"
         Checked         =   -1  'True
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFreeze 
         Caption         =   "Free&ze"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuJoy 
         Caption         =   "Use &Joysick"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuSW 
         Caption         =   "Size &Window"
         Begin VB.Menu mnuSize 
            Caption         =   "&Maximized"
            Checked         =   -1  'True
            Index           =   0
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuSize 
            Caption         =   "&Large"
            Index           =   1
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuSize 
            Caption         =   "M&edium"
            Index           =   2
            Shortcut        =   ^E
         End
         Begin VB.Menu mnuSize 
            Caption         =   "&Small"
            Index           =   3
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuSize 
            Caption         =   "Mi&nimum"
            Index           =   4
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuSize 
            Caption         =   "&Iconized"
            Index           =   5
            Shortcut        =   ^I
         End
      End
      Begin VB.Menu mnusep34 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheat 
         Caption         =   "&Cheat Mode"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSep33 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   ""
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep41 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMail 
         Caption         =   "Sen&d Mail to Author"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL As Long = 1
Private Const SE_NO_ERROR   As Long = 33  'Values below 33 are error returns

Private Declare Function GetJoyDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Private Const MAXPNAMELEN As Long = 32  '  max product name length (including NULL)
Private Type JOYCAPS
    wMid        As Integer
    wPid        As Integer
    szPname     As String * MAXPNAMELEN
    wXmin       As Long
    wXmax       As Long
    wYmin       As Long
    wYmax       As Long
    wZmin       As Long
    wZmax       As Long
    wNumButtons As Long
    wPeriodMin  As Long
    wPeriodMax  As Long
End Type
Private JoyCapabs As JOYCAPS

Private Declare Sub GetJoyPos Lib "winmm.dll" Alias "joyGetPos" (ByVal uJoyID As Long, pji As JOYINFO)
Private Type JOYINFO
    wXpos       As Long
    wYpos       As Long
    wZpos       As Long 'used for range
    wButtons    As Long
End Type
Private ThisJoyInf          As JOYINFO
Private PrevJoyInf          As JOYINFO
Private MinJoyInf           As JOYINFO
Private MaxJoyInf           As JOYINFO 'wZpos here is ymax - ymin = deltaY
'                                       wXPos = 1 3rd wYpos = 2 3rd of x-rabge
Private Calibrated          As Boolean

Private ControlAreaTop      As Long
Private ControlAreaHeight   As Long
Private PrevWidth           As Long
Private PrevHeight          As Long

Private m_CurrentX          As Single
Private m_CurrentY          As Single
Private Const m_VptSideRel  As Single = 0.78
Private Const m_MinHeight   As Long = 8100
Private Const m_MinWidth    As Long = m_MinHeight / m_VptSideRel
Private LastSize            As Integer
Private UserName            As String

'Cursors
Private m_LButtonCur        As Long
Private m_RButtonCur        As Long
Private m_NoButtonCur       As Long

Private m_Power             As Single
Private Counting            As Boolean
Private Repet               As Single
Private Dispcount           As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If mnuFreeze.Checked = False Then
        m_bButtonIsDown = True
        Select Case KeyCode

            'shooting
          Case vbKeySpace, vbKeyReturn
            lblShoot_Click
          Case vbKeyEscape
            lblShoot = Shoot
            m_Power = 0
          Case vbKeyNumpad1 To vbKeyNumpad9
            picPower.Cls
            m_Power = (KeyCode - vbKeyNumpad0) * 10
            lblShoot = vbNullString
            lblShoot = m_Power & " %"
          Case vbKey1 To vbKey9
            picPower.Cls
            m_Power = (KeyCode - vbKey0) * 10
            lblShoot = vbNullString
            lblShoot = m_Power & " %"
          Case vbKey0, vbKeyNumpad0
            picPower.Cls
            m_Power = 100
            lblShoot = vbNullString
            lblShoot = m_Power & " %"
          Case vbKeyS, vbKeyU
            Counting = True
            tmrPercent_Timer
            Counting = False
          Case vbKeyD
            m_Power = m_Power - Sgn(m_Power)
            lblShoot = m_Power & " %"

            'cameras, movement, aiming
          Case vbKeyHome
            If mnuTg.Enabled Then
                ToggleCameras
            End If
          Case vbKeyUp
            m_ActiveCam = CamMoveable
            MouseKeyboardEventHandler 1, 0, 0, 0, -0.002
          Case vbKeyDown
            m_ActiveCam = CamMoveable
            MouseKeyboardEventHandler 1, 0, 0, 0, 0.002
          Case vbKeyLeft
            m_ActiveCam = CamMoveable
            Repet = Repet - (Repet < 30) / 4
            MouseKeyboardEventHandler 1, 0, 0, 0.0001 * Repet, 0
          Case vbKeyRight
            m_ActiveCam = CamMoveable
            Repet = Repet - (Repet < 30) / 4
            MouseKeyboardEventHandler 1, 0, 0, -0.0001 * Repet, 0
          Case vbKeyPageUp
            m_ActiveCam = CamMoveable
            MouseKeyboardEventHandler 2, 0, 0, 0, -0.003
          Case vbKeyPageDown
            m_ActiveCam = CamMoveable
            MouseKeyboardEventHandler 2, 0, 0, 0, 0.003

            'help
          Case vbKeyF1
            mnuAbout_Click

            'undo
          Case vbKeyBack
            If mnuUndo.Enabled Then
                mnuUndo_Click
            End If

        End Select
        m_bButtonIsDown = False
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
      Case vbKeyLeft, vbKeyRight
        Repet = 0
    End Select

End Sub

Private Sub Form_Load()

  Dim k As Long

    k = 128
    UserName = String$(k, 0)
    GetUserName UserName, k
    UserName = Left$(UserName, k + (Asc(Mid$(UserName, k, 1)) = 0))
    lblLoading = lblLoading & " " & UserName
    mnuSize_Click 1
    mnuUndo.Caption = "&Undo last shot" & vbTab & "Backspace"
    JoyPresent = (GetJoyDevCaps(0, JoyCapabs, Len(JoyCapabs)) = 0)
    mnuJoy.Enabled = JoyPresent
    Show
    DoEvents
    With picViewport
        .ForeColor = BackColor
        .BackColor = BackColor
    End With 'PICVIEWPORT

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbAppWindows Then
        On Error Resume Next 'in case a modal window is open
            Cancel = True
            Enabled = False
            frmWinTerm.Show vbModeless, Me
            If Err Then
                Cancel = False
            End If
        On Error GoTo 0
    End If

End Sub

Private Sub Form_Resize()

    ReleaseCapture
    If WindowState <> vbMinimized Then
        ControlAreaTop = ScaleHeight * m_VptSideRel
        picViewport.Height = ControlAreaTop
        ControlAreaHeight = ScaleHeight - ControlAreaTop

        With imgCW
            .Move (picViewport.Width - .Width) / 2, (ControlAreaTop - .Height) / 1.4
        End With 'IMGCW

        With lbTitle(0)
            .Caption = Caption
            .Move (picViewport.Width - .Width + 5) / 2, (ControlAreaTop - .Height + 5) / 2.5
        End With 'LBTITLE(0)

        With lbTitle(1)
            .Caption = Caption
            .Move (picViewport.Width - .Width - 5) / 2, (ControlAreaTop - .Height - 5) / 2.5
        End With 'LBTITLE(1)

        With lblGoodBye
            .Move (ScaleWidth - .Width) / 2, (ScaleHeight - .Height) * 0.75
        End With 'IMGCW 'LBLGOODBYE

        With lblLoading
            .Move (ScaleWidth - .Width) / 2, ControlAreaTop + ControlAreaHeight / 2 - .Height
        End With 'LBLLOADING

        'The "shoot button" label
        With lblShoot
            .FontSize = 20
            .Move (ScaleWidth - .Width) / 2, ControlAreaTop + (ControlAreaHeight - .Height) * 0.72
            shpShoot.Move .Left - 4, .Top - 4, .Width + 8, .Height + 8
        End With 'LBLSHOOT

        With picPower
            .Move (ScaleWidth - .Width) / 2, ControlAreaTop + (ControlAreaHeight - .Height) * 0.28
        End With 'PICPOWER

        With lblAuthor
            .Move (ScaleWidth - .Width) / 2, ControlAreaTop + ControlAreaHeight - .Height - 2
        End With 'LBLAUTHOR

        'The label displaying current player number
        With lblPlayer
            .Move ScaleWidth - .Width - 20, ControlAreaTop + (ControlAreaHeight - .Height) / 2
        End With 'LBLPLAYER

        With lblSloMo
            .Move (ScaleWidth - .Width) / 2, ControlAreaTop
            lblFrozen.Move .Left, .Top + .Height
        End With 'LBLSLOMO

        With lblPl(0)
            .Move 10, ControlAreaTop + (ControlAreaHeight - .Height) * 0.28
        End With 'LBLPL(0)

        With lblPl(1)
            .Move 10, ControlAreaTop + (ControlAreaHeight - .Height) * 2 / 3
        End With 'LBLPL(1)
        With lblFpS
            .Move ScaleWidth - .Width - 5, lblAuthor.Top
        End With 'LBLFPS
    End If
    Select Case WindowState
      Case vbMaximized
        mnuSize_Click 0
      Case vbNormal
        mnuSize_Click LastSize
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim Color     As Long
  Dim CsrVis    As Long
  Dim WasMinim  As Boolean

    CsrVis = ShowCursor(True) 'is cursor visible?
    HideCrsr
    ShowCrsr
    If mnuFreeze.Checked Then
        mnuFreeze_Click
    End If
    picViewport.Cls
    SetCntlsVisible False
    If WindowState = vbMinimized Then
        WasMinim = True
        WindowState = IIf(LastSize = 0, vbMaximized, vbNormal)
    End If
    If Enabled Then
        With frmStop
            .Move Left + (Width - .Width) / 2, Top + ControlAreaTop * 15 - .Height / 3
            .Show vbModal, Me
        End With 'FRMSTOP
      Else 'ENABLED = FALSE/0
        Unload frmStop
    End If
    If frmStop.optYesNo(0) Or Enabled = False Then
        HideCrsr
        mnuGame.Enabled = False
        mnuStgs.Enabled = False
        mnuAbout.Enabled = False
        lblFpS = vbNullString
        CloseDown
        DoEvents
        If frmPool.WindowState <> vbMinimized Then
            With lblGoodBye
                .Visible = True
                For Color = &HFF& To &H40& Step -2
                    Sleep 30
                    .ForeColor = RGB(0, Color, 0)
                    .Refresh
                Next Color
            End With 'LBLGOODBYE
            Sleep 666
            frmPool.WindowState = vbMinimized
            DoEvents
        End If
        ShowCrsr
        IdleStopDetection
        Rem Mark Off Silent
        End 'since this can be called from almost anywhere we have to force End and not return
        Rem Mark On
      Else 'NOT FRMSTOP.OPTYESNO(0)...
        Cancel = True
        If CsrVis <= 0 Then 'cursor was not visible
            HideCrsr
        End If
        SetCntlsVisible True
        Unload frmStop
        If WasMinim Then
            WindowState = vbMinimized
        End If
    End If

End Sub

Private Sub lblPlayer_Change()

    lblPlayer.Top = ControlAreaTop + (ControlAreaHeight - lblPlayer.Height) / 2
    lblPlayer.Visible = True
    Dispcount = 0

End Sub

Private Sub lblShoot_Change()

    If m_bButtonIsDown Then
        picPower.Cls
        picPower.Line (0, 0)-(Val(lblShoot), picPower.Height), , BF
    End If

End Sub

Private Sub lblShoot_Click()

    If lblShoot.Visible And Val(lblShoot) And Not mnuFreeze.Checked Then
        SloMoFactor = IIf(mnuSloMo.Checked, SloMoDelay, 1)
        picPower.Line (0, 0)-(Val(lblShoot), picPower.Height), , BF
        FireCueBall m_Power / (g_dt * 500) + (Rnd - Rnd) / 20 'faster computers can shoot harder
    End If

End Sub

Private Sub lblShoot_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    m_bButtonIsDown = True
    lblShoot.BorderStyle = vbFixedSingle
    m_Power = 0
    picPower.Cls
    Counting = True

End Sub

Private Sub lblShoot_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Counting = False
    lblShoot.BorderStyle = vbBSNone
    m_bButtonIsDown = False

End Sub

Private Sub mnuAbout_Click()

    lblFpS = vbNullString
    Render
    lblFpS = "0 FpS"
    Load frmAbout
    With frmAbout
        .Theme = Timer Mod 27 + 1
        .AppIcon(&HFFE0C0) = Icon
        .Title(&H80FF&) = App.ProductName
        .Version(&HFFC0C0) = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        .Copyright(vbYellow) = App.LegalCopyright
        .Otherstuff1(&HF0FFF0) = "Original Author: Grzegorz Holdys (gregor@kn.pl)"
        .Otherstuff2(&HE0C0C0) = "Enhancements and Debugging: Ulli (umgedv@yahoo.com)" & vbCrLf & "Enjoy The Game!"
        .Show vbModal, Me
    End With 'FRMABOUT

End Sub

Private Sub mnuAiming_Click()

    mnuAiming.Checked = Not mnuAiming.Checked
    AimingLine = mnuAiming.Checked

End Sub

Private Sub mnuCheat_Click()

    mnuCheat.Checked = Not mnuCheat.Checked
    InNormalMode = Not mnuCheat.Checked

End Sub

Private Sub mnuFreeze_Click()

    mnuFreeze.Checked = Not mnuFreeze.Checked
    lblFrozen.Visible = mnuFreeze.Checked
    picViewport.Enabled = Not mnuFreeze.Checked
    ShowCursor mnuFreeze.Checked = False
    If mnuFreeze.Checked Then
        lblFpS = "0 FpS"
    End If

End Sub

Private Sub mnuGameExit_Click()

    Unload Me

End Sub

Private Sub mnuGameNew_Click()

    Counting = False
    StartNewGame False

End Sub

Private Sub mnuJoy_Click()

    mnuJoy.Checked = Not mnuJoy.Checked
    If mnuJoy.Checked Then
        If Not Calibrated Then
            Enabled = False
            frmCalibrate.Show vbModeless, Me
            MinJoyInf.wXpos = 100000
            MinJoyInf.wYpos = 100000
            MaxJoyInf.wXpos = 0
            MaxJoyInf.wYpos = 0
            tmrPercent.Enabled = False
            With PrevJoyInf
                Do
                    Do
                        Render
                        DoEvents
                        GetJoyPos 0, PrevJoyInf
                        If .wXpos < MinJoyInf.wXpos Then
                            MinJoyInf.wXpos = .wXpos
                        End If
                        If .wYpos < MinJoyInf.wYpos Then
                            MinJoyInf.wYpos = .wYpos
                        End If
                        If .wXpos > MaxJoyInf.wXpos Then
                            MaxJoyInf.wXpos = .wXpos
                        End If
                        If .wYpos > MaxJoyInf.wYpos Then
                            MaxJoyInf.wYpos = .wYpos
                        End If
                    Loop Until MinJoyInf.wXpos + 30000 < MaxJoyInf.wXpos And MinJoyInf.wYpos + 30000 < MaxJoyInf.wYpos
                Loop Until .wButtons = 1
                With MaxJoyInf
                    .wZpos = .wYpos - MinJoyInf.wYpos 'delta
                    .wXpos = (.wXpos - MinJoyInf.wXpos) / 3 - MinJoyInf.wXpos
                    .wYpos = (.wYpos - MinJoyInf.wYpos) * 2 / 3 - MinJoyInf.wYpos
                End With 'MAXJOYINF

                Do
                    DoEvents
                    GetJoyPos 0, PrevJoyInf
                Loop Until .wButtons = 0
            End With 'PREVJOYINF
            tmrPercent.Enabled = True
            Calibrated = True
            Unload frmCalibrate
            Enabled = True
        End If
    End If

End Sub

Private Sub mnuMail_Click()

    With App
        If ShellExecute(hwnd, vbNullString, "mailto:UMGEDV@Yahoo.com?subject=" & .ProductName & " V" & .Major & "." & .Minor & "." & .Revision & " &body=Hi Ulli,<br><br>[your message]<br><br>Best regards from " & UserName, vbNullString, .Path, SW_SHOWNORMAL) < SE_NO_ERROR Then
            MessageBeep vbCritical
            MsgBox "Cannot send Mail from this System.", vbCritical, "Mail disabled/not installed"
        End If
    End With 'APP

End Sub

Private Sub mnuSize_Click(Index As Integer)

  Dim Diff As Long

    If Index <> 5 Then
        For Diff = 0 To 4
            mnuSize(Diff).Checked = False
        Next Diff
    End If
    Diff = (Screen.Width - m_MinWidth) / 4

    Select Case Index
      Case 0 'maximized"
        WindowState = vbMaximized
        mnuSize(Index).Checked = True

      Case 1 To 4
        LastSize = Index
        WindowState = vbNormal
        Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2, Screen.Width - Diff * Index, Width * m_VptSideRel
        mnuSize(Index).Checked = True

      Case 5 'iconized
        WindowState = vbMinimized

    End Select

End Sub

Private Sub mnuSloMo_Click()

    mnuSloMo.Checked = Not mnuSloMo.Checked
    lblSloMo.Visible = mnuSloMo.Checked
    If mnuSloMo.Checked = False Then
        SloMoFactor = 1
      Else 'NOT MNUSLOMO.CHECKED...
        If m_Billards.AnyBillardInMotion Then
            SloMoFactor = SloMoDelay
        End If
    End If

End Sub

Private Sub mnuSVHigh_Click()

    SetChecked 8
    SetSoundVolumeBase m_VolumeMax

End Sub

Private Sub mnuSVLow_Click()

    SetChecked 2
    SetSoundVolumeBase m_VolumeLow

End Sub

Private Sub mnuSVMedium_Click()

    SetChecked 4
    SetSoundVolumeBase m_VolumeMedium

End Sub

Private Sub mnuSVMute_Click()

    SetChecked 1
    SetSoundVolumeBase m_VolumeMute

End Sub

Private Sub mnuTG_Click()

    ToggleCameras

End Sub

Private Sub mnuTgAfterHit_Click()

    mnuTgAfterHit.Checked = Not mnuTgAfterHit.Checked

End Sub

Private Sub mnuUndo_Click()

    mnuUndo.Enabled = False
    SetToPrevPosn

End Sub

Private Sub picPower_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    m_bButtonIsDown = True
    m_Power = Fix(x)
    lblShoot = m_Power & " %"
    'lblShoot_Change

End Sub

Private Sub picPower_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If x >= 0 And x <= 100 Then
        If Button = vbLeftButton Then
            picPower_MouseDown Button, Shift, x, y
          Else 'NOT BUTTON...
            picPower.Cls
            picPower.Line (x, 0)-(x, picPower.ScaleHeight), vbBlack
            m_Power = Fix(x)
            lblShoot = m_Power & " %"
        End If
    End If

End Sub

Private Sub picPower_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Not (x < 0 Or x > 100 Or y < 0 Or y > picPower.Height) Then
        lblShoot_Click
        lblShoot_MouseUp 0, 0, 0, 0
    End If
    m_bButtonIsDown = False

End Sub

Private Sub picViewport_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Down = True
    With Viewport
        m_CurrentX = x * .Width / ScaleWidth
        m_CurrentY = y * .Height / ScaleHeight / m_VptSideRel
    End With 'VIEWPORT
    MouseKeyboardEventHandler Button, m_CurrentX, m_CurrentY, 0, 0
    'Set the cursor
    If Button = vbLeftButton Then
        picViewport.MousePointer = m_LButtonCur
      ElseIf Button = vbRightButton Then 'NOT BUTTON...
        picViewport.MousePointer = m_RButtonCur
    End If

End Sub

Private Sub picViewport_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim ShiftX As Single
  Dim ShiftY As Single

  'Set the cursor

    If Button = vbLeftButton Then
        picViewport.MousePointer = m_LButtonCur
      ElseIf Button = vbRightButton Then 'NOT BUTTON...
        picViewport.MousePointer = m_RButtonCur
    End If
    With Viewport
        x = x * .Width / picViewport.Width
        y = y * .Height / ControlAreaTop
        ShiftX = x - m_CurrentX
        ShiftY = y - m_CurrentY
        m_CurrentX = x
        m_CurrentY = y
        If m_CurrentX >= 0 And m_CurrentX <= .Width And m_CurrentY >= 0 And m_CurrentY <= .Height Then
            ShiftX = ShiftX / picViewport.Width * 0.2
            ShiftY = ShiftY / ControlAreaTop * 0.2
            MouseKeyboardEventHandler Button, x, y, ShiftX, ShiftY
          Else 'NOT M_CURRENTX...
            picViewport.MousePointer = m_NoButtonCur
        End If
    End With 'VIEWPORT

End Sub

Private Sub picViewport_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    picViewport.MousePointer = IIf(m_ActiveCam = CamMoveable, vbCustom, vbArrow)
    Down = False

End Sub

Private Sub SetChecked(Cond As Long)

    mnuSVMute.Checked = (Cond And 1)
    mnuSVLow.Checked = (Cond And 2)
    mnuSVMedium.Checked = (Cond And 4)
    mnuSVHigh.Checked = (Cond And 8)

End Sub

Private Sub SetCntlsVisible(State As Boolean)

    tmrPercent.Enabled = State
    lblPlayer.Visible = State
    lblSloMo.Visible = State And mnuSloMo.Checked = True
    lblFrozen.Visible = State And mnuFreeze.Checked = True
    lblPl(0).Visible = State
    lblPl(1).Visible = State
    lblFpS.Visible = State
    ShowShootingControls State
    lbTitle(0).Visible = Not State
    lbTitle(1).Visible = Not State
    lblGoodBye.Enabled = Not State

    DoEvents

End Sub

Friend Sub SetCursorIcons(Optional NoButtonCur As Integer, Optional LButtonCur As Integer, Optional RButtonCur As Integer)

    m_NoButtonCur = NoButtonCur
    m_LButtonCur = LButtonCur
    m_RButtonCur = RButtonCur
    picViewport.MousePointer = IIf(m_ActiveCam = CamMoveable, vbCustom, vbArrow)

End Sub

Public Property Let ShootingPower(nuPower As Long)

    m_Power = nuPower

End Property

Private Sub tmrPercent_Timer()

    If mnuJoy.Checked Then
        DoEvents
        GetJoyPos 0, ThisJoyInf
        With ThisJoyInf
            'smoothing
            .wXpos = (.wXpos + 3 * PrevJoyInf.wXpos) / 4
            .wYpos = (.wYpos + 3 * PrevJoyInf.wYpos) / 4

            PrevJoyInf = ThisJoyInf
            m_Power = 100 - Int(100 * (.wYpos - MinJoyInf.wYpos) / MaxJoyInf.wZpos)
            Select Case m_Power
              Case Is < 0
                m_Power = 0
              Case Is > 100
                m_Power = 100
            End Select
            m_bButtonIsDown = True
            lblShoot = m_Power & " %"
            m_bButtonIsDown = False
            If .wButtons = 1 Then
                lblShoot_Click
            End If
            If Dispcount And 1 Then
                Select Case .wXpos
                  Case Is < MaxJoyInf.wXpos
                    Form_KeyDown vbKeyLeft, 0
                  Case Is > MaxJoyInf.wYpos
                    Form_KeyDown vbKeyRight, 0
                  Case Else
                    Form_KeyUp vbKeyLeft, 0
                End Select
            End If
        End With 'THISJOYINF
      Else 'MNUJOY.CHECKED = FALSE/0
        If m_Power < 100 And Counting Then
            m_Power = m_Power + 1
            lblShoot = m_Power & " %"
        End If
    End If
    Dispcount = (Dispcount + 1) Mod 70
    lblPlayer.Visible = (Dispcount < 50)

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-29 14:15)  Decl: 63  Code: 699  Total: 762 Lines
':) CommentOnly: 16 (2,1%)  Commented: 43 (5,6%)  Empty: 154 (20,2%)  Max Logic Depth: 7
