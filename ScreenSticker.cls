VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreenSticker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Represents the Sticker Controler
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'References to DirectX objects
Private m_D3DDevice8        As Direct3DDevice8
Private m_D3DX8             As D3DX8
Private m_Texture           As Direct3DTexture8

'Sticker characteristics
Private m_Rect              As Rect
Private m_vRotCentre        As D3DVECTOR2
Private m_Rotation          As Single
Private m_vScale            As D3DVECTOR2
Private m_vTranslate        As D3DVECTOR2
Private m_Color             As Long
Private m_bVisible          As Boolean

Private Sub Class_Terminate()

  'Clears all objects

    Set m_D3DDevice8 = Nothing
    Set m_D3DX8 = Nothing
    Set m_Texture = Nothing

End Sub

Friend Sub Draw(ByVal Sprite As D3DXSprite)

  'Draws the screen sticker image using a D3DxSprite object provided by the caller.

    Sprite.Draw m_Texture, m_Rect, m_vScale, m_vRotCentre, m_Rotation, m_vTranslate, m_Color

End Sub

Friend Sub Setup(ByRef TexSrc As String, ByVal PixelFormat As CONST_D3DFORMAT, ByVal SpriteLeft As Single, ByVal SpriteTop As Single, ByVal SpriteWidth As Single, ByVal SpriteHeight As Single, ByVal SpriteColor As Long, ByVal D3DDevice8 As Direct3DDevice8, ByVal D3DX8 As D3DX8)

  'Copy the references to DirectX objects

    Set m_D3DDevice8 = D3DDevice8
    Set m_D3DX8 = D3DX8

    'Create the sprite bounding rectangle
    With m_Rect
        .Right = SpriteWidth
        .Bottom = SpriteHeight
    End With 'M_RECT

    'And the translation vector
    With m_vTranslate
        .y = SpriteTop
        .x = SpriteLeft
    End With 'M_VTRANSLATE

    'Copy the sprite's color
    m_Color = SpriteColor

    'The scaling vector is by default (1,1)
    m_vScale.x = 1
    m_vScale.y = 1

    'The centre of rotation is by default in the middle of m_Rect
    With m_Rect
        m_vRotCentre.x = .Left + (.Right - .Left) / 2
        m_vRotCentre.y = .Top + (.Bottom - .Top) / 2
    End With 'M_RECT
    'Rotation is initially set to zero.
    m_Rotation = 0

    'Create the texture
    Set m_Texture = m_D3DX8.CreateTextureFromFileEx(m_D3DDevice8, TexSrc, D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, PixelFormat, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, 0, ByVal 0, ByVal 0)

End Sub

Friend Property Get Visible() As Boolean

  'Returns true if the sticker is visible and false if it's not

    Visible = m_bVisible

End Property

Friend Property Let Visible(bVal As Boolean)

    m_bVisible = (bVal <> False)

End Property

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-29 14:15)  Decl: 19  Code: 74  Total: 93 Lines
':) CommentOnly: 18 (19,4%)  Commented: 3 (3,2%)  Empty: 29 (31,2%)  Max Logic Depth: 2
