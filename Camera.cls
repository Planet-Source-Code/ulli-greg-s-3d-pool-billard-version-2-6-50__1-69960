VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Represents a Camera
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'Main camera parametres
Private m_mGenerators       As D3DMATRIX            'The generator matrix for this camera
Private m_mProjection       As D3DMATRIX            'Projection transformation matrix for this camera
Private m_mView             As D3DMATRIX            'View transformation matrix for this cmera

'Pivot point and lever
Private m_vPivotPt          As D3DVECTOR            'Position vector indicating the point (in scene coordinates) around which the camera will rotate
Private m_mPivToOrigin      As D3DMATRIX            'A matrix describing the translation from the pivot point to the origin of the coordinate system
Private m_mOriginToPiv      As D3DMATRIX            'A matrix describing the translation from the origin of the coordinate system to the pivot point

'Angular and linear velocities
Private m_SpinLclX          As Single               'The value of spin around a local (camera's) X-axis
Private m_SpinGlblY         As Single               'The value of spin around a global (scene's) Y_axis
Private m_vVel              As D3DVECTOR            'Vector of linear velocity.

'Transition variables
Private m_vDestEye          As D3DVECTOR            'The position of the camera at the end of a transition loop
Private m_vDestUp           As D3DVECTOR            'The vertical axis of the camera at the end of a transition loop
Private m_vDestAt           As D3DVECTOR            'The point observed by the camera at the end of a transition loop
Private m_vSrcEye           As D3DVECTOR            'The position of the camera at the start of a transition loop
Private m_vSrcUp            As D3DVECTOR            'The vertical axis of the camera at the start of a transition loop
Private m_vSrcAt            As D3DVECTOR            'The point observed by the camera at the start of a transition loop
Private m_numFramesLeft     As Long                 'The number of frames left in a transition loop
Private m_numFramesTotal    As Long                 'The total number of frames in a transition loop
Private m_blnTransition     As Boolean              '"True" if a transition is in progress

'Helpers
Private vVctr1              As D3DVECTOR
Private vVctr2              As D3DVECTOR
Private vVctr3              As D3DVECTOR

Friend Property Get CamInTransit() As Boolean

    CamInTransit = m_blnTransition

End Property

Friend Sub ChangeView()

  'unconditional and rapid change of the camera's position

    D3DXMatrixLookAtLH m_mView, MakeVector(0, 3, 0), MakeVector(0, 0, 0), MakeVector(1, 0, 0)
    D3DXMatrixInverse m_mGenerators, Determinant, m_mView

End Sub

Friend Property Get Generators() As D3DMATRIX

    Generators = m_mGenerators

End Property

Friend Sub MoveCam(glbl_vShift As D3DVECTOR)

  'Moves the camera along a vector.

    If Not m_blnTransition Then  'If the camera is not performing a transition, update its linear velocity vector
        D3DXVec3Add m_vVel, m_vVel, ScaleVector(glbl_vShift, 1 / g_dt)
    End If

End Sub

Friend Sub PivotCam(glbl_vPivotPt As D3DVECTOR, ByVal HorizRot As Single, ByVal VertRot As Single)

  'Enables pivotting the camera around a point in space.

    If Not m_blnTransition Then 'Ignore the Pivot call if the camera is performing a transition
        m_vPivotPt = glbl_vPivotPt
        'Update the spins
        m_SpinLclX = m_SpinLclX + HorizRot / g_dt
        m_SpinGlblY = m_SpinGlblY + VertRot / g_dt
        With m_vPivotPt
            'Matrix of translation from pivot point to the origin
            D3DXMatrixTranslation m_mPivToOrigin, -.x, -.y, -.z
            'Matrix of translation from the origin to the pivot point
            D3DXMatrixTranslation m_mOriginToPiv, .x, .y, .z
        End With 'M_VPIVOTPT
    End If

End Sub

Friend Property Get Position() As D3DVECTOR

    With Position
        .x = m_mGenerators.m41
        .y = m_mGenerators.m42
        .z = m_mGenerators.m43
    End With 'POSITION

End Property

Friend Property Get ProjectionMatrix() As D3DMATRIX

    ProjectionMatrix = m_mProjection

End Property

Friend Sub Setup()

    ChangeView
    D3DXMatrixPerspectiveFovLH m_mProjection, g_Pi / 4, Viewport.Height / Viewport.Width, 0.01, 10
    D3DXMatrixIdentity m_mOriginToPiv
    D3DXMatrixIdentity m_mPivToOrigin

End Sub

Friend Sub StartTransit(vDestEye As D3DVECTOR, vDestAt As D3DVECTOR, vDestUp As D3DVECTOR, ByVal numFrames As Long)

  'Starts a camera transition loop

  Dim vDist     As D3DVECTOR

  '1) The initial camera vectors

    With m_mGenerators
        m_vSrcEye = MakeVector(.m41, .m42, .m43) 'current position of the camera
        m_vSrcAt = MakeVector(.m41 + .m31, .m42 + .m32, .m43 + .m33) 'point currently looked at by the camera
        m_vSrcUp = MakeVector(.m21, .m22, .m23) 'current vertical axis of the camera
    End With 'M_MGENERATORS

    '2) The final camera vectors
    m_vDestEye = vDestEye
    m_vDestAt = vDestAt
    m_vDestUp = vDestUp

    '3) Compute the transit distance from m_vSrcEye to m_vDestEye
    D3DXVec3Subtract vDist, m_vSrcEye, m_vDestEye

    '4) The total number of frames for this transition (depend on distance)
    m_numFramesTotal = numFrames * D3DXVec3Length(vDist) ^ 0.25 / 1.6 'these values are empirical

    '5) The remaining number of frames (initially equal to the total number).
    m_numFramesLeft = m_numFramesTotal

    '6) The boolean indicating, that the camera is in transition mode
    m_blnTransition = True

    '7) Make some sound
    PlaySound m_DSTransitionLong, 1

    '8) Disable camera toggle
    frmPool.mnuTg.Enabled = False

    SloMoFactor = 1

End Sub

Friend Sub StartTransitToFixed(ByVal Frames As Long)

    StartTransit MakeVector(0, 3, 0), MakeVector(0, 0, 0), MakeVector(1, 0, 0), Frames / 2

End Sub

Friend Sub UpdateCamPos()

  Dim RotLclX         As Single           'Rotation around a local X-Axis
  Dim PitchAngle      As Single           'Pitch angle...
  Dim CosPitchAngle   As Single           'Cosine of the pitch angle
  Dim mRotationGlblY  As D3DMATRIX        'Matrix of rotation around a global Y-Axis
  Dim mRotationLclX   As D3DMATRIX        'Matrix of rotation around a local X-Axis
  Dim mTranslation    As D3DMATRIX        'Matrix of translation
  Dim mFinalTrans     As D3DMATRIX        'The final (summed) transformation matrix
  Dim BlendWeight     As Single           'Used for transition lerping

    If m_blnTransition Then
        'The process of transition is very similiar to a technique called tweening, used for
        'animating flexible bodies. It is built around Linear Interpolation (Lerp) of vectors.

        'Sinus transition function; BlendWeight curve looks like this

        ' BlendWeight -> 1 |          ..
        '                  |        .
        '                  |       ·
        '                  |      ·
        '                  |    .
        '                0 | ..·
        '                  o-----------> NumFrames

        BlendWeight = -Sin((m_numFramesLeft / m_numFramesTotal - 0.5) * g_Pi) / 2 + 0.5

        'lerp the vectors
        D3DXVec3Lerp vVctr1, m_vSrcEye, m_vDestEye, BlendWeight
        D3DXVec3Lerp vVctr2, m_vSrcAt, m_vDestAt, BlendWeight
        D3DXVec3Lerp vVctr3, m_vSrcUp, m_vDestUp, BlendWeight

        'create the new matrices
        D3DXMatrixLookAtLH m_mView, vVctr1, vVctr2, vVctr3
        D3DXMatrixInverse m_mGenerators, Determinant, m_mView

        'count frames
        m_numFramesLeft = m_numFramesLeft - 1
        m_blnTransition = (m_numFramesLeft >= 0)
        frmPool.mnuTg.Enabled = Not m_blnTransition
        If m_blnTransition = False Then 'stop the transit sound
            StopSound m_DSTransitionLong
        End If

      Else 'm_blnTransition = FALSE/0

        'Moving Cam by hand

        'The angle between camera's line of sight and the vertical axis
        CosPitchAngle = m_mGenerators.m32
        If CosPitchAngle > -1 Then
            PitchAngle = Atn(-CosPitchAngle / Sqr(-CosPitchAngle * CosPitchAngle + 1)) + g_PiHalf
          Else 'NOT COSPITCHANGLE...
            PitchAngle = 0
        End If

        'The PitchAngle of rotation around local X axis
        RotLclX = m_SpinLclX * g_dt
        'The rotation is limited to angles between a litle more than pi/2 and a little less than pi
        '(measured form the vertical axis)
        Select Case PitchAngle + RotLclX
          Case Is <= g_PiHalf + 0.07, Is >= g_Pi - 0.07
            RotLclX = 0
            m_SpinLclX = 0
        End Select

        'Matrix of rotation around local X axis
        D3DXMatrixRotationAxis mRotationLclX, MakeVector(m_mGenerators.m11, m_mGenerators.m12, m_mGenerators.m13), RotLclX
        'Matrix of rotation around the global Y axis
        D3DXMatrixRotationY mRotationGlblY, m_SpinGlblY * g_dt
        'Matrix of translation along the linear velocity vector
        D3DXMatrixTranslation mTranslation, m_vVel.x * g_dt, m_vVel.y * g_dt, m_vVel.z * g_dt

        'Create the transformation matrix
        '1) Translate from pivot point to the origin
        D3DXMatrixMultiply mFinalTrans, m_mGenerators, m_mPivToOrigin
        '2) Rotate around the local X axis
        D3DXMatrixMultiply mFinalTrans, mFinalTrans, mRotationLclX
        '3) Rotate around the global Y axis
        D3DXMatrixMultiply mFinalTrans, mFinalTrans, mRotationGlblY
        '4) Return to the pivot point
        D3DXMatrixMultiply mFinalTrans, mFinalTrans, m_mOriginToPiv
        '5) Move the camera along integrated m_vVel vector
        D3DXMatrixMultiply mFinalTrans, mFinalTrans, mTranslation

        'The new generator matrix
        m_mGenerators = mFinalTrans
        'And the new view transformation matrix
        D3DXMatrixInverse m_mView, Determinant, m_mGenerators

        'Reduce the spins and the linear velocity by a constant coefficient
        m_SpinLclX = m_SpinLclX * 0.66
        m_SpinGlblY = m_SpinGlblY * 0.66
        D3DXVec3Scale m_vVel, m_vVel, 0.66
    End If

End Sub

Friend Property Get ViewMatrix() As D3DMATRIX

    ViewMatrix = m_mView

End Property

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-29 14:15)  Decl: 36  Code: 229  Total: 265 Lines
':) CommentOnly: 55 (20,8%)  Commented: 38 (14,3%)  Empty: 72 (27,2%)  Max Logic Depth: 3
