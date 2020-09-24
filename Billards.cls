VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBillards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Represents all balls
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Type BILLARDVERTEX
    XYZ         As D3DVECTOR
    Normal      As D3DVECTOR
    Tex1        As D3DVECTOR2
End Type
Private Const D3DFVF_BILLARDVERTEX              As Long = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1

Private Type SHADOWVERTEX
    XYZ         As D3DVECTOR
    Color       As Long
End Type
Private Const D3DFVF_SHADOWVERTEX               As Long = D3DFVF_XYZ Or D3DFVF_DIFFUSE

Private Const m_numBillards                     As Long = 16
Private Const m_MinSpeed                        As Single = 1 / 40000
Private m_D3DX8                                 As D3DX8
Private m_D3DDevice8                            As Direct3DDevice8
Private m_BillardMesh                           As D3DXMesh
Private m_WhiteMaterial                         As D3DMATERIAL8
Private m_Textures(0 To m_numBillards - 1)      As Direct3DTexture8

'Shadows
Private m_ShadowVB                              As Direct3DVertexBuffer8
Private m_numShadowVB                           As Long
Private m_sizeShadowVB                          As Long
Private m_plnShadow                             As D3DPLANE

'Light direction
Private m_vLightDir                             As D3DVECTOR4

'Array of generators
Private m_mGenerators(0 To m_numBillards - 1)   As D3DMATRIX

Private m_PocketNumber(m_numBillards - 1)       As Long
Private m_NumFramesSpentInPocket(0 To m_numBillards - 1) As Long
Private m_MaxFramesSpentInPocket                As Long
Private m_AnyBillardInMotion                    As Boolean

'Physical data
Private m_vVelocities(0 To m_numBillards - 1)   As D3DVECTOR
Private m_vSpins(0 To m_numBillards - 1)        As D3DVECTOR
Private m_Mass                                  As Single
Private m_Inertia                               As Single
Private m_Radius                                As Single
Private m_Elasticity                            As Single

'for PrevPosns
Private m_PrevGenerators(0 To m_numBillards - 1)    As D3DMATRIX
Private m_PrevNumFramesSpentInPocket(0 To m_numBillards - 1) As Long
'Helpers
Private mMtrx1                                  As D3DMATRIX
Private mMtrx2                                  As D3DMATRIX
Private vVctr1                                  As D3DVECTOR
Private vVctr2                                  As D3DVECTOR
Private i                                       As Long
Private j                                       As Long

Friend Property Get AnyBillardInMotion() As Boolean

    AnyBillardInMotion = m_AnyBillardInMotion

End Property

Public Property Get BillardPosition(ByVal Indx As Long) As D3DVECTOR

    If Indx < m_numBillards And Indx > -1 Then
        With BillardPosition
            .x = m_mGenerators(Indx).m41
            .y = m_mGenerators(Indx).m42
            .z = m_mGenerators(Indx).m43
        End With 'BILLARDPOSITION
    End If

End Property

Friend Property Let BillardPosition(ByVal Indx As Long, PosVctr As D3DVECTOR)

    If Indx < m_numBillards And Indx > -1 Then
        With m_mGenerators(Indx)
            .m41 = PosVctr.x
            .m42 = PosVctr.y
            .m43 = PosVctr.z
        End With 'M_MGENERATORS(Indx)
    End If

End Property

Friend Property Get BillardSpin(ByVal Indx As Long) As D3DVECTOR

    If Indx < m_numBillards And Indx > -1 Then
        BillardSpin = m_vSpins(Indx)
    End If

End Property

Friend Property Let BillardSpin(ByVal Indx As Long, SpinVctr As D3DVECTOR)

    If Indx < m_numBillards And Indx > -1 Then
        m_vSpins(Indx) = SpinVctr
    End If

End Property

Friend Property Get BillardVelocity(ByVal Indx As Long) As D3DVECTOR

    If Indx < m_numBillards And Indx > -1 Then
        BillardVelocity = m_vVelocities(Indx)
    End If

End Property

Friend Property Let BillardVelocity(ByVal Indx As Long, VelocityVctr As D3DVECTOR)

    If Indx < m_numBillards And Indx > -1 Then
        m_vVelocities(Indx) = VelocityVctr
    End If

End Property

Private Sub Class_Terminate()

  'Called when an instance of this class is destroyed

    Erase m_Textures

    Set m_ShadowVB = Nothing
    Set m_BillardMesh = Nothing
    Set m_D3DDevice8 = Nothing
    Set m_D3DX8 = Nothing

End Sub

Friend Sub ClearPocketNumbers()

  'Sets all values in the pocket number array to 0 and adds another frame to the m_numFramesSpentInPocket array.

    For i = 0 To m_numBillards - 1
        If m_PocketNumber(i) > 0 Then
            m_NumFramesSpentInPocket(i) = m_MaxFramesSpentInPocket + 1
            m_PocketNumber(i) = 0
        End If
    Next i

End Sub

Friend Sub Create(ByVal D3DDevice8 As Direct3DDevice8, ByVal D3DX8 As D3DX8)

  'Creates sphere mesh, sets materials etc.

  Dim Angle As Single
  Dim AdjBuffer   As D3DXBuffer
  Dim Mesh As D3DXMesh
  Dim VBuffer As Direct3DVertexBuffer8
  Dim VArray() As BILLARDVERTEX
  Dim numVertices As Long
  Dim ShadowVertices() As SHADOWVERTEX

    Set m_D3DDevice8 = D3DDevice8
    Set m_D3DX8 = D3DX8

    'Billard's radius, mass and inertia momentum
    m_Radius = 0.05
    m_Mass = 0.1
    m_Inertia = m_Mass * m_Radius '* m_Radius
    m_Elasticity = 0.92

    'Create billard mesh
    Set Mesh = m_D3DX8.CreateSphere(m_D3DDevice8, m_Radius, 48, 36, AdjBuffer)
    Set m_BillardMesh = Mesh.CloneMeshFVF(ByVal 0, D3DFVF_BILLARDVERTEX, m_D3DDevice8)
    numVertices = m_BillardMesh.GetNumVertices
    ReDim VArray(1 To numVertices)
    Set VBuffer = m_BillardMesh.GetVertexBuffer
    D3DVertexBuffer8GetData VBuffer, 0, numVertices * Len(VArray(1)), 0, VArray(1)
    For i = 1 To numVertices
        VArray(i).Tex1.x = (-Sgn(VArray(i).XYZ.z) * VArray(i).XYZ.x + m_Radius) / (2 * m_Radius)
        VArray(i).Tex1.y = -(VArray(i).XYZ.y + m_Radius) / (2 * m_Radius)
    Next i
    D3DVertexBuffer8SetData VBuffer, 0, numVertices * Len(VArray(1)), 0, VArray(1)

    'Cleanup
    Set AdjBuffer = Nothing
    Set Mesh = Nothing
    Set VBuffer = Nothing
    Erase VArray

    'Textures
    On Error Resume Next
        For i = 0 To m_numBillards - 1
            Set m_Textures(i) = m_D3DX8.CreateTextureFromFile(m_D3DDevice8, App.Path & Replace$(BillardBMP, "#", i))
        Next i
    On Error GoTo 0

    'All billards are made of the same shiny, white material
    With m_WhiteMaterial
        .Ambient = MakeColor(1, 1, 1)
        .diffuse = .Ambient
        .specular = .Ambient
        .power = 10
    End With 'M_WHITEMATERIAL

    'Shadows'vertex buffer
    'The shadow is simply a (flat) circle with the same radius as the billard mesh.
    'When rendering the contents of this buffer we will use the D3DPT_TRIANGLEFAN primitive type,
    'thus, we don't need an index buffer
    m_numShadowVB = 25 '24 is round enough for the shadow. The additional vertex is the middle one
    ReDim ShadowVertices(0 To m_numShadowVB - 1)

    Angle = g_2Pi / (m_numShadowVB - 2)
    ShadowVertices(0).XYZ = MakeVector(0, 0, 0)
    ShadowVertices(0).Color = D3DColorMake(0, 0, 0, 1)
    For i = 0 To m_numShadowVB - 2
        ShadowVertices(i + 1).XYZ = MakeVector(m_Radius * Sin(i * Angle), 0, m_Radius * Cos(i * Angle))
        ShadowVertices(i + 1).Color = D3DColorMake(0, 0, 0, 0.1)
    Next i
    m_sizeShadowVB = Len(ShadowVertices(0))
    'Create the vertex buffer
    Set m_ShadowVB = m_D3DDevice8.CreateVertexBuffer(m_numShadowVB * m_sizeShadowVB, 0, D3DFVF_SHADOWVERTEX, D3DPOOL_DEFAULT)
    'Fill it
    D3DVertexBuffer8SetData m_ShadowVB, 0, m_numShadowVB * m_sizeShadowVB, 0, ShadowVertices(0)

    'Shadow plane
    D3DXPlaneFromPoints m_plnShadow, MakeVector(0, 0.005, 0), MakeVector(1, 0.005, 0), MakeVector(0, 0.005, 1)

    'Light direction
    m_vLightDir = GetLightDir

    Erase ShadowVertices

End Sub

Friend Property Get Diameter() As Single

    Diameter = m_Radius + m_Radius

End Property

Friend Property Get FellInPocketNumber(ByVal Indx As Long) As Long

    If Indx < m_numBillards And Indx > -1 Then
        FellInPocketNumber = m_PocketNumber(Indx)
    End If

End Property

Friend Property Let FellInPocketNumber(ByVal Indx As Long, ByVal iPocket As Long)

    If Indx < m_numBillards And Indx > -1 Then
        m_PocketNumber(Indx) = iPocket
    End If

End Property

Friend Sub FireCueBall(VelVctr As D3DVECTOR)

  'Fires the cue-ball with a given initial velocity vector

    For i = 0 To m_numBillards - 1
        With m_mGenerators(i)
            m_PrevGenerators(i) = m_mGenerators(i)
            m_PrevNumFramesSpentInPocket(i) = m_NumFramesSpentInPocket(i)
        End With 'M_MGENERATORS(I)
    Next i
    m_vVelocities(0) = VelVctr
    m_AnyBillardInMotion = True

End Sub

Friend Sub GetPhysBillConstants(ByRef retMass As Single, ByRef retInertia As Single, ByRef retRadius As Single, ByRef retElasticity As Single)

    retMass = m_Mass
    retInertia = m_Inertia
    retRadius = m_Radius
    retElasticity = m_Elasticity

End Sub

Friend Sub InitialPositions(FirstGame As Boolean)

  'Places the billards on the table in a triangle.

  Dim Inaccuracy As Single
  Dim Used       As Long
  Dim BallNum    As Long

    Do 'Slightly randomize the initial positions
        Inaccuracy = (Rnd - Rnd)
    Loop While Abs(Inaccuracy) > 0.05

    Randomize Timer

    'None of the billards is initially moving or in any pocket and all are visible, thus
    For BallNum = 0 To m_numBillards - 1
        m_PocketNumber(BallNum) = 0
        m_NumFramesSpentInPocket(BallNum) = 0
        m_vVelocities(BallNum) = MakeVector(0, 0, 0)
        If FirstGame Then
            'place balls in a stright line
            D3DXMatrixTranslation m_mGenerators(BallNum), 0, m_Radius, -BallNum * m_Radius * 2
        End If
    Next BallNum

    m_MaxFramesSpentInPocket = 1 / g_dt 'one second
    Render

    If FirstGame Then
        ActiveWait 1
    End If

    'balls are grouped in a triangular shape
    For BallNum = 1 To 5
        GoSub XlatBallNum
        D3DXMatrixTranslation m_mGenerators(j), 0.306 - BallNum * 0.102 + (Rnd - Rnd) / 333, m_Radius, 1.08 + Inaccuracy
    Next BallNum
    For BallNum = 6 To 9
        GoSub XlatBallNum
        D3DXMatrixTranslation m_mGenerators(j), 0.254 - (BallNum - 5) * 0.102 + (Rnd - Rnd) / 333, m_Radius, 0.99 + Inaccuracy
    Next BallNum
    For BallNum = 10 To 12
        GoSub XlatBallNum
        D3DXMatrixTranslation m_mGenerators(j), 0.202 - (BallNum - 9) * 0.102 + (Rnd - Rnd) / 333, m_Radius, 0.9 + Inaccuracy
    Next BallNum
    For BallNum = 13 To 14
        GoSub XlatBallNum
        D3DXMatrixTranslation m_mGenerators(j), 0.15 - (BallNum - 12) * 0.102 + (Rnd - Rnd) / 333, m_Radius, 0.81 + Inaccuracy
    Next BallNum
    GoSub XlatBallNum '15th ball
    D3DXMatrixTranslation m_mGenerators(j), (Rnd - Rnd) / 333, m_Radius, 0.72 + Inaccuracy

    'and finally the cue-ball
    D3DXMatrixTranslation m_mGenerators(0), (Rnd - Rnd) / 333, m_Radius, -1 + Inaccuracy / 2

Exit Sub

XlatBallNum:

    Render

    '--------------------------------------------------------

    'place balls in numerical order
    'j = BallNum

    '--------------------------------------------------------

    'alternatively place balls in random order
    Do
        j = (Int(Rnd * 15) + 1)
    Loop While 2 ^ j And Used
    Used = Used Or 2 ^ j

    '--------------------------------------------------------
    If FirstGame Then
        ActiveWait 0.1
      Else 'FIRSTGAME = FALSE/0
        ActiveWait 0.05
    End If
    Return

End Sub

Friend Property Get InTheGame(ByVal Indx As Long) As Boolean

    InTheGame = (m_NumFramesSpentInPocket(Indx) = 0)

End Property

Friend Sub NextFrame()

  'Computes new generator matrices, velocity and spin vectors

  Dim vTranslation    As D3DVECTOR
  Dim vRotation       As D3DVECTOR
  Dim mRotation       As D3DMATRIX
  Dim Angle           As Single

    m_AnyBillardInMotion = False
    'Update the generators
    For i = 0 To m_numBillards - 1
        'First, check if the billard is still "in the game"
        If m_PocketNumber(i) > 0 Or m_NumFramesSpentInPocket(i) > 0 Then
            If m_NumFramesSpentInPocket(i) <= m_MaxFramesSpentInPocket Then
                m_NumFramesSpentInPocket(i) = m_NumFramesSpentInPocket(i) + 1
            End If
          Else 'NOT M_POCKETNUMBER(I)...
            'Generate spin caused by the billard's linear velocity
            D3DXVec3Scale vVctr1, m_vVelocities(i), -1
            D3DXVec3Cross vVctr2, MakeVector(0, -1, 0), vVctr1
            D3DXVec3Scale m_vSpins(i), vVctr2, 1 / m_Radius
            'Integrate the velocity vector into translation vector
            D3DXVec3Scale vTranslation, m_vVelocities(i), g_dt
            'Translate the billard
            With m_mGenerators(i)
                .m41 = .m41 + vTranslation.x
                .m42 = .m42 + vTranslation.y
                .m43 = .m43 + vTranslation.z
            End With 'M_MGENERATORS(I)

            If FastDecelerate Then '8 ball was sunk - decelerate fast
                D3DXVec3Scale m_vVelocities(i), m_vVelocities(i), 1 - 5 * g_dt
              Else 'FastDecelerate = FALSE/0
                If InNormalMode Or i = 0 Or i = 8 Then
                    'Reduce velocity due to friction
                    D3DXVec3Scale m_vVelocities(i), m_vVelocities(i), 1 - g_dt * 0.95
                  Else 'NOT I... 'NOT INNORMALMODE...
                    D3DXVec3Scale m_vVelocities(i), m_vVelocities(i), 1 + g_dt * 0.15
                End If
            End If

            'If the velocity is small enough then trucate it to 0; the spin will be zeroed further down.
            If D3DXVec3LengthSq(m_vVelocities(i)) < m_MinSpeed Then 'thats pretty slow
                m_vVelocities(i) = MakeVector(0, 0, 0)
              Else 'NOT D3DXVEC3LENGTHSQ(M_VVELOCITIES(I))...
                m_AnyBillardInMotion = True
            End If
            'Transform the spin vector into the coordinate system of the i'th billard
            'First create the transformation matrix
            D3DXMatrixInverse mMtrx1, Determinant, m_mGenerators(i)
            'Now, integrate the spin vector into rotation vector
            D3DXVec3Scale vRotation, m_vSpins(i), g_dt
            'Conduct the transformation
            D3DXVec3TransformNormal vRotation, vRotation, mMtrx1
            'Normalize the rotation vector
            NormalizeD3DVECTOR vRotation, Angle, vRotation
            'And, if the angle is greater than zero...
            If Angle > 0 Then
                '... create a transformation matrix for this rotation
                D3DXMatrixRotationAxis mRotation, vRotation, Angle
                'Conduct the rotation transformation (rotation comes first as it is local)
                D3DXMatrixMultiply m_mGenerators(i), mRotation, m_mGenerators(i)
            End If
        End If
    Next i

End Sub

Private Sub NormalizeD3DVECTOR(retV As D3DVECTOR, retModule As Single, vVctr As D3DVECTOR)

    retModule = Sqr(vVctr.x * vVctr.x + vVctr.y * vVctr.y + vVctr.z * vVctr.z)
    If retModule > 0 Then
        With vVctr
            retV.x = .x / retModule
            retV.y = .y / retModule
            retV.z = .z / retModule
        End With 'VVCTR
    End If

End Sub

Friend Property Get NumBillards() As Long

    NumBillards = m_numBillards

End Property

Friend Sub PrevPosns()

    For i = 0 To m_numBillards - 1
        m_mGenerators(i) = m_PrevGenerators(i)
        m_NumFramesSpentInPocket(i) = m_PrevNumFramesSpentInPocket(i)
    Next i

End Sub

Friend Sub ReappearCueBall()

  'Sets the number of frames spent in a pocket by the cue-ball to 0.

    m_NumFramesSpentInPocket(0) = 0

End Sub

Friend Sub RenderBillards()

  'Draws the scene

  Dim matBlack As D3DMATERIAL8
  Dim mTranslation As D3DMATRIX

    If Not m_D3DDevice8 Is Nothing Then
        'The material is the same for all billards
        With m_D3DDevice8
            .SetRenderState D3DRS_WRAP0, D3DWRAP_U
            .SetMaterial m_WhiteMaterial
            .SetVertexShader D3DFVF_BILLARDVERTEX
            For i = 0 To m_numBillards - 1
                If m_NumFramesSpentInPocket(i) < m_MaxFramesSpentInPocket Then
                    'Set the world transformation matrix and the texture for this subset
                    .SetTransform D3DTS_WORLD, m_mGenerators(i)
                    .SetTexture 0, m_Textures(i)
                    'Draw the mesh subset
                    m_BillardMesh.DrawSubset 0
                End If
            Next i

            'Shadows
            .SetVertexShader D3DFVF_SHADOWVERTEX
            .SetRenderState D3DRS_SPECULARENABLE, 0
            .SetRenderState D3DRS_LIGHTING, 0
            .SetRenderState D3DRS_ALPHABLENDENABLE, 1
            .SetStreamSource 0, m_ShadowVB, m_sizeShadowVB
            .SetMaterial matBlack
            For i = 0 To m_numBillards - 1
                If m_NumFramesSpentInPocket(i) < m_MaxFramesSpentInPocket Then
                    'Set the world transformation matrix for the billard's shadow
                    With m_mGenerators(i)
                        D3DXMatrixTranslation mTranslation, .m41, .m42, .m43
                    End With 'M_MGENERATORS(I)
                    D3DXMatrixShadow mMtrx1, m_vLightDir, m_plnShadow
                    D3DXMatrixMultiply mMtrx2, mTranslation, mMtrx1
                    .SetTransform D3DTS_WORLD, mMtrx2
                    'Draw the shadow
                    .DrawPrimitive D3DPT_TRIANGLEFAN, 0, m_numShadowVB - 2
                End If
            Next i
            .SetRenderState D3DRS_LIGHTING, 1
            .SetRenderState D3DRS_SPECULARENABLE, 1
            .SetRenderState D3DRS_ALPHABLENDENABLE, 0
            .SetRenderState D3DRS_WRAP0, 0
        End With 'M_D3DDEVICE8
    End If

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-29 14:15)  Decl: 62  Code: 469  Total: 531 Lines
':) CommentOnly: 63 (11,9%)  Commented: 20 (3,8%)  Empty: 120 (22,6%)  Max Logic Depth: 6
