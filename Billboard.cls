VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBillboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Represents the Arrow Billboard
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'A billboard is simply a transparent rectangular shape
'covered with a texture. Its most significant feature
'is the fact that it is always facing the camera.

Private Type BILLBOARDVERTEX
    x       As Single
    y       As Single
    z       As Single
    Color   As Long
    tu      As Single
    tv      As Single
End Type
Private Const D3DFVF_BILLBOARDVERTEX As Long = D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1

'Geometry data
Private m_vBasePt           As D3DVECTOR    'A point at the bottom of the billboard around which the billboard rotates.
Private m_mGenerators       As D3DMATRIX    'A generator matrix.

'DirectX objects
Private m_VertexBuffer      As Direct3DVertexBuffer8
Private m_D3DDevice8        As Direct3DDevice8
Private m_D3DX8             As D3DX8
Private m_Texture           As Direct3DTexture8

'The size, in bytes, of the BILLBOARDVERTEX data type
Private m_VertexSize        As Long

'A flag indicating whether the billboard is visible or not
Private m_bVisible          As Boolean
Private m_bShowing          As Boolean

Friend Property Get BasePoint() As D3DVECTOR

    BasePoint = m_vBasePt

End Property

Friend Property Let BasePoint(vPt As D3DVECTOR)

    m_vBasePt = vPt

End Property

Private Sub Class_Terminate()

    Set m_D3DDevice8 = Nothing
    Set m_D3DX8 = Nothing
    Set m_VertexBuffer = Nothing
    Set m_Texture = Nothing

End Sub

Friend Sub RenderBillboard(vEye As D3DVECTOR, vUp As D3DVECTOR)

  Dim vAxisZ  As D3DVECTOR    'The Z-axis of the billboard
  Dim vAxisX  As D3DVECTOR    'The x-axis of the billboard
  Dim vAxisY  As D3DVECTOR    'The y-axis of the billboard

    If Not m_D3DDevice8 Is Nothing And m_bVisible Then
        'We need vEye because we need to know where the observer is
        'if we want to "face him" (in fact we face the active cam)
        'The vUp vector tells us, where "up" is (and where "down" is)
        D3DXVec3Subtract vAxisZ, vEye, m_vBasePt
        D3DXVec3Cross vAxisX, vAxisZ, vUp
        D3DXVec3Normalize vAxisZ, vAxisZ
        D3DXVec3Normalize vAxisX, vAxisX
        D3DXVec3Cross vAxisY, vAxisX, vAxisZ
        'Since vectors vAxisZ and vAxisX were normalized before the
        'last cross product, the result of this operation,
        'the vAxisY vector, is also normalized.

        'With these three vectors (vAxisZ, vAxisX, vAxisY) and the centre point vector
        'we can create the generator matrix for this billboard

        With m_mGenerators
            .m11 = vAxisX.x
            .m12 = vAxisX.y
            .m13 = vAxisX.z
            .m14 = 0
            .m21 = vAxisY.x
            .m22 = vAxisY.y
            .m23 = vAxisY.z
            .m24 = 0
            .m31 = vAxisZ.x
            .m32 = vAxisZ.y
            .m33 = vAxisZ.z
            .m34 = 0
            .m41 = m_vBasePt.x
            .m42 = m_vBasePt.y
            .m43 = m_vBasePt.z
            .m44 = 1
        End With 'M_MGENERATORS

        With m_D3DDevice8
            .SetRenderState D3DRS_ALPHABLENDENABLE, 1
            .SetRenderState D3DRS_LIGHTING, 0
            .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
            .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
            .SetVertexShader D3DFVF_BILLBOARDVERTEX
            .SetStreamSource 0, m_VertexBuffer, m_VertexSize
            .SetTexture 0, m_Texture
            .SetTransform D3DTS_WORLD, m_mGenerators
            .DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
            .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
            .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE
            .SetRenderState D3DRS_LIGHTING, 1
            .SetRenderState D3DRS_ALPHABLENDENABLE, 0
        End With 'M_D3DDEVICE8
    End If

End Sub

Friend Sub Setup(ByRef TexSrc As String, ByVal BillHeight As Single, ByVal BillWidth As Single, ByVal D3DDevice8 As Direct3DDevice8, ByVal D3DX8 As D3DX8)

  Dim Vertices(3) As BILLBOARDVERTEX

  'Copy the references to DirectX objects

    Set m_D3DDevice8 = D3DDevice8
    Set m_D3DX8 = D3DX8

    'Vertices are numbered according to this
    'diagram
    '0 --- 1     y
    '|     |     |
    '|     |     |
    '|     |     |
    '2 -+- 3     +----- x

    With Vertices(0)
        .Color = &HFFFFFFFF
        .x = -BillWidth / 2
        .y = BillHeight
        .z = 0
        .tu = 0
        .tv = 0
    End With 'VERTICES(0)

    With Vertices(1)
        .Color = &HFFFFFFFF
        .x = BillWidth / 2
        .y = BillHeight
        .z = 0
        .tu = 1
        .tv = 0
    End With 'VERTICES(1)

    With Vertices(2)
        .Color = &HFFFFFFFF
        .x = -BillWidth / 2
        .y = 0
        .z = 0
        .tu = 0
        .tv = 1
    End With 'VERTICES(2)

    With Vertices(3)
        .Color = &HFFFFFFFF
        .x = BillWidth / 2
        .y = 0
        .z = 0
        .tu = 1
        .tv = 1
    End With 'VERTICES(3)

    'Create and fill the vertex buffer
    m_VertexSize = Len(Vertices(0))
    Set m_VertexBuffer = m_D3DDevice8.CreateVertexBuffer(4 * m_VertexSize, 0, D3DFVF_BILLBOARDVERTEX, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData m_VertexBuffer, 0, 4 * m_VertexSize, 0, Vertices(0)

    'Create the texture
    Set m_Texture = m_D3DX8.CreateTextureFromFile(m_D3DDevice8, TexSrc)

End Sub

Friend Property Get Showing() As Boolean

    Showing = m_bShowing

End Property

Friend Sub Visible(bVal As Boolean, Optional LightIsOn As Boolean = True)

    m_bShowing = (bVal <> False)
    m_bVisible = m_bShowing And LightIsOn

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-29 14:15)  Decl: 36  Code: 160  Total: 196 Lines
':) CommentOnly: 30 (15,3%)  Commented: 11 (5,6%)  Empty: 42 (21,4%)  Max Logic Depth: 3
