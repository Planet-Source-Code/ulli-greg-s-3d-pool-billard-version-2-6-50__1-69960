VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollisionController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Represents the Collision Controler
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'+-----------------------------------------------------------------------+
'|   GENERAL NOTES:                                                      |
'|                                                                       |
'|   - A capital "C" after "m_" or at the beginning of a variable name   |
'|     (eg. "m_CBillBillMomentum") stands for "Collision".               |
'|   - A capital "B" after "m_" or at the beginning of a variable name   |
'|     (eg. "m_BMass") stands for "Billard".                             |
'|   - A capital "T" after "m_" or at the beginning of a variable name   |
'|                                                                       |
'+-----------------------------------------------------------------------+

'Reference to objects taking part in collisions.
Private m_cBillards                         As clsBillards
Private m_Table                             As clsTable

'Physical data
'The billards
Private m_vBPositions()                     As D3DVECTOR    'Array of billards'position vectors
Private m_vBVelocities()                    As D3DVECTOR    'Array of billards'linear velocity vectors
Private m_vBSpins()                         As D3DVECTOR    'Array of billards'spin (angular velocity) vectors
Private m_BMass                             As Single       'Billard's mass
Private m_BInertia                          As Single       'Billard's inertia (around any axis).
Private m_BRadius                           As Single       'Billard radius.
Private m_2BRadSqrd                         As Single       'Billard radius, squared and multiplied by two.
Private m_BElasticity                       As Single
'The table
Private m_TBumperElasticity                 As Single       'Bumper elasticity coefficient (used for collision reactions).
Private m_TBumpers(1 To 6, 1 To 5)          As D3DVECTOR    'Array of points (vectors) representing bumpers.
Private m_TBumperNorms(1 To 6, 1 To 5)      As D3DVECTOR    'Array of normal vectors for the bumpers.

'Billard "presence" list
Private m_bBInTheGame()                     As Boolean      'True when billard() is still present
Private m_bBSinking8Illegal                 As Boolean      'True when sinking the 8-ball is illegal

'Pockets
Private m_vTPockets(1 To 6)                 As D3DVECTOR    'Array of pocket positions.
Private m_TPocketRadSqrd                    As Single       'Squared radius of table pockets.

'Collisions detected
Private m_CBillBillMomentum                 As Single       'The scalar value of momentum for the collision between two billards.
'                                                            (currently used only for setting collision sound volume)
Private m_CBillTableMomentum                As Single       'The scalar value of momentum for the collision between a billard and the table.
'                                                            (currently used only for setting collision sound volume)
Private m_bPocketHitDetected                As Boolean      'True when a billard falls into a pocket.

'Counters
Private m_numBillards                       As Long         'Yeah, You guessed it: the total number of billards in the game.
Private Const m_numTableBumpers             As Long = 6     'The number of table bumpers

'To time out an unpromising cheat mode situation
Private BallsMoving                         As Long         'Number balls currently moving
Private TimeoutCount                        As Long         'Number of straight-edge bumper hits in a row

'Helpers
Private vVctr1                              As D3DVECTOR
Private a                                   As Single
Private b                                   As Single
Private c                                   As Single
Private i                                   As Long
Private j                                   As Long
Private k                                   As Long

Friend Property Get BillBillCollisionMomentum() As Single

    BillBillCollisionMomentum = m_CBillBillMomentum

End Property

Friend Property Get BillTableCollisionMomentum() As Single

    BillTableCollisionMomentum = m_CBillTableMomentum

End Property

Friend Property Get CheatTimeOut() As Boolean

  'This becomes True when all moving billards have accumulated
  'an average of 10 or more straight edge bumper contacts in a row,
  'this being an indication that probably one or more billards are
  'cought in a loop.

  'Bill-Bill collisions reset the TimeoutCount as do pocket hits
  'and sloped bumper edge contacts that is - billard hits one of
  'the edges that lead into the pockets.

    CheatTimeOut = ((TimeoutCount \ BallsMoving) > 9)

End Property

Friend Sub CollisionDetection()

    With m_cBillards
        'First, get the newest generators, velocity and spin vectors into private arrays
        For i = 0 To m_numBillards - 1
            m_vBPositions(i) = .BillardPosition(i)
            m_vBVelocities(i) = .BillardVelocity(i)
            m_vBSpins(i) = .BillardSpin(i)
            m_bBInTheGame(i) = .InTheGame(i)
        Next i
    End With 'M_CBILLARDS

    'We haven't detected any collisions yet, thus
    m_CBillBillMomentum = 0
    m_CBillTableMomentum = 0
    m_bPocketHitDetected = False

    'Now, run the collision tests and eventual reactions
    'First, for collisions between billards...
    CollisionDetection_BillardBillard
    '... then for collisions between the billards and the table bumpers
    CollisionDetection_BillardBumper   'This sub also looks for pocket hits

    'Finally, update the generators, velocity and spin vectors from the private arrays
    With m_cBillards
        For i = 0 To m_numBillards - 1
            .BillardPosition(i) = m_vBPositions(i)
            .BillardVelocity(i) = m_vBVelocities(i)
            .BillardSpin(i) = m_vBSpins(i)
        Next i
    End With 'M_CBILLARDS

End Sub

Private Sub CollisionDetection_BillardBillard()

  'Detects collisions between billards and generates collision response

  Dim vBillard_i      As D3DVECTOR    'Position vector of the i'th billard
  Dim vBillard_j      As D3DVECTOR    'Position vector of the j'th billard
  Dim vVel_i          As D3DVECTOR    'Velocity vector of the i'th billard
  Dim vVel_j          As D3DVECTOR    'Velocity vector of the j'th billard
  Dim vDist           As D3DVECTOR    'A vector linking the centres of both billards.
  Dim vCNorm          As D3DVECTOR    'Collision Normal vector.
  Dim vVel_ij         As D3DVECTOR    'Relative velocity between the i'th and j'th billard.
  Dim vCPVel_i        As D3DVECTOR    'Velocity vector of the i'th billard AT POINT OF COLLISION.
  Dim vCPVel_j        As D3DVECTOR    'Velocity vector of the j'th billard AT POINT OF COLLISION.
  Dim vCPVel_ij       As D3DVECTOR    'Relative velocity between the i'th and j'th billard at collision point.
  Dim Delta           As Single       'A utility variable.
  Dim Ct1             As Single       'First possible point in time, when a collision occurs.
  Dim Ct2             As Single       'Second possible point in time, when a collision occurs.
  Dim CT              As Single       'The real collision point in time.
  Dim CMmntm          As Single       'The momentum of collision.

    For i = m_numBillards - 1 To 0 Step -1  'Run through all billards from the back of the array to its front
        'Don't bother looking for collisions, if the billard is not on the table any more
        If m_bBInTheGame(i) Then
            'Copy the i'th billard's position and velocity to local variables
            vBillard_i = m_vBPositions(i)
            vVel_i = m_vBVelocities(i)
            'For every i'th billard iterate through all other billards in search for collisions
            For j = m_numBillards - 1 To 0 Step -1
                'Don't bother looking for collisions, if the j'th billard is not on the table any more, or if j = i
                If m_bBInTheGame(j) Then
                    If j <> i Then
                        'Copy the j'th billard's position and velocity
                        'to local variables
                        vBillard_j = m_vBPositions(j)
                        vVel_j = m_vBVelocities(j)
                        D3DXVec3Subtract vDist, vBillard_i, vBillard_j  'A vector linking the centres of both billards
                        D3DXVec3Subtract vVel_ij, vVel_i, vVel_j        'Relative velocity between billard "i" and "j"
                        If (vDist.x * vDist.x + vDist.z * vDist.z) <= m_2BRadSqrd Then
                            'If the condition above is true, the billards are touching or overlapping. Thus the collision has already occured.
                            'We will pretend that it has happend at this particular point in time, hence
                            CT = 0
                          Else 'NOT (VDIST.X...
                            'If the condition is false, then we need to check whether a collision will occur within this time frame.
                            'To do this we use a standard solution for a binominal equation
                            With vVel_ij
                                a = .x * .x + .z * .z
                                If a = 0 Then
                                    Delta = -1
                                  Else 'NOT A...
                                    b = 2 * (.x * vDist.x + .z * vDist.z)
                                    c = vDist.x * vDist.x + vDist.z * vDist.z - m_2BRadSqrd
                                    Delta = b * b - 4 * a * c
                                End If
                            End With 'VVEL_IJ
                            If Delta < 0 Then
                                CT = -1
                              Else 'NOT DELTA...
                                CT = (-b - Sqr(Delta)) / (a + a)
                            End If
                        End If
                        If CT >= 0 And CT < g_dt Then 'we have or will have a collision if current velocity of bodies is enough for them to touch
                            'Collision response
                            'First move both billards along their trajectories to the point of collision
                            D3DXVec3Add vBillard_i, vBillard_i, ScaleVector(vVel_i, CT)
                            D3DXVec3Add vBillard_j, vBillard_j, ScaleVector(vVel_j, CT)
                            D3DXVec3Subtract vDist, vBillard_i, vBillard_j

                            'Collision's Normal vector
                            D3DXVec3Normalize vCNorm, vDist

                            'i-th body velocity at Collision Point
                            vCPVel_i = vVel_i
                            vCPVel_j = vVel_j

                            'relative velocity at Collision Point
                            D3DXVec3Subtract vCPVel_ij, vCPVel_i, vCPVel_j

                            If DotProduct(vCPVel_ij, vCNorm) <= 0 Then 'we have (or will have) a REAL collision
                                'The Collision Impulse value
                                CMmntm = (1 + m_BElasticity) * DotProduct(vCPVel_ij, vCNorm) / (2 / m_BMass)

                                'Update velocities
                                D3DXVec3Add m_vBVelocities(i), m_vBVelocities(i), ScaleVector(vCNorm, -CMmntm / m_BMass)
                                D3DXVec3Add m_vBVelocities(j), m_vBVelocities(j), ScaleVector(vCNorm, CMmntm / m_BMass)

                                'ULLI:
                                'Add a fraction of the original velocities back for the effect of rotational inertia
                                D3DXVec3Add m_vBVelocities(i), m_vBVelocities(i), ScaleVector(vCPVel_i, m_BInertia)
                                D3DXVec3Add m_vBVelocities(j), m_vBVelocities(j), ScaleVector(vCPVel_j, m_BInertia)

                                'Update positions
                                m_vBPositions(i) = vBillard_i
                                m_vBPositions(j) = vBillard_j
                                'Update the value of momentum for a billard - billard collision,
                                m_CBillBillMomentum = m_CBillBillMomentum - CMmntm
                                TimeoutCount = 0 'restart timeout count on bill-bill collision
                            End If
                        End If
                    End If
                End If
            Next j
        End If
    Next i

End Sub

Private Sub CollisionDetection_BillardBumper()

  'Detects collisions between billards and the bumpers and generates collision response

  Dim vBillardPos     As D3DVECTOR    'Position vector of the i'th billard.
  Dim vBillardVel     As D3DVECTOR    'Velocity vector of a billard.
  Dim vCNorm          As D3DVECTOR    'Collision normal vector.
  Dim vCP_i           As D3DVECTOR    'Collision Point relative to the centre of the i'th billard
  Dim vCPVel_i        As D3DVECTOR    'Velocity vector of the i'th billard AT POINT OF COLLISION.
  Dim vBumperToBill   As D3DVECTOR    'A vetor linking one vertex of a bumper with a billard's centre point
  Dim vDist           As D3DVECTOR    'A vector linking the billard's and the pocket's centre points.
  Dim D1              As Single       'A utility variable.
  Dim CT              As Single       'A point in time when a collision was detected.
  Dim LatestCt        As Single       'A point in time, when the most recent collision with an edge took place.
  Dim iLatestEdge     As Single       'The edge, which took part in the latest collision.
  Dim CMmntm          As Single       'Collision Momentum.
  Dim MinDistToEdge   As Single       'Minimal recorded distance to an edge.
  Dim BVelDotTBNorm   As Single       'Result of a dot product between Billard's velocity vector and a bumper normal.
  Dim MinDistTime     As Single       'A point in time, when the distance between a pocket and a billard is in the shortest.
  Dim BillardVelSqrd  As Single       'Billard's velocity squared.
  Dim iNorm           As Long         'Index of a normal vector in the normal vector array
  Dim iClosestEdge    As Long         'Index of the edge, which is at the closest distance to the billard.
  Dim bCDetected      As Boolean      'Set to "True" if a collision was detected.

    BallsMoving = 0
    For i = 0 To m_numBillards - 1
        If m_bBInTheGame(i) Then 'is still in the gane
            vBillardPos = m_vBPositions(i)
            vBillardVel = m_vBVelocities(i)
            BillardVelSqrd = DotProduct(vBillardVel, vBillardVel)
            'We are not interested in the billards that are not moving
            If BillardVelSqrd Then
                For j = 1 To m_numTableBumpers
                    iLatestEdge = 0
                    iClosestEdge = 0
                    LatestCt = -1
                    bCDetected = True
                    MinDistToEdge = 10

                    For k = 1 To 5
                        'Computing the distance between the k'th edge of table bumper and the billard
                        D3DXVec3Subtract vBumperToBill, vBillardPos, m_TBumpers(j, k)
                        D1 = DotProduct(vBumperToBill, m_TBumperNorms(j, k))
                        If D1 > m_BRadius Then
                            bCDetected = False
                          Else 'NOT D1...
                            If -D1 < MinDistToEdge Then
                                MinDistToEdge = -D1
                                iClosestEdge = k
                            End If
                            BVelDotTBNorm = DotProduct(vBillardVel, m_TBumperNorms(j, k))
                            If BVelDotTBNorm Then
                                CT = (m_BRadius - D1) / BVelDotTBNorm
                                If CT > LatestCt And CT <= 0 And CT >= -g_dt Then
                                    LatestCt = CT
                                    iLatestEdge = k
                                End If
                            End If
                        End If
                    Next k
                    If iLatestEdge = 0 Then
                        iNorm = iClosestEdge
                      Else 'NOT ILATESTEDGE...
                        iNorm = iLatestEdge
                    End If
                    If bCDetected Then  'we have a possible collision
                        vCNorm = m_TBumperNorms(j, iNorm)
                        D3DXVec3Scale vCP_i, vCNorm, -m_BRadius
                        D3DXVec3Cross vCPVel_i, m_vBSpins(i), vCP_i
                        'i'th billard's velocity at Collision Point
                        D3DXVec3Add vCPVel_i, vCPVel_i, vBillardVel
                        'Collision response
                        If DotProduct(vCPVel_i, vCNorm) <= 0 Then 'We are sure that we have a collision
                            'The Collision Momentum value
                            CMmntm = -(1 + m_BElasticity * m_TBumperElasticity) * DotProduct(vCPVel_i, vCNorm) / (1 / m_BMass)
                            If LatestCt > -g_dt Then
                                'Update billard's position by moving it along its trajectory to the point of collision
                                D3DXVec3Add vBillardPos, vBillardPos, ScaleVector(vBillardVel, LatestCt)
                                m_vBPositions(i) = vBillardPos
                            End If
                            'Update velocities
                            D3DXVec3Add vBillardVel, vBillardVel, ScaleVector(vCNorm, CMmntm / m_BMass)
                            'ULLI:
                            'Add a fraction of the original velocity for rotational inertia
                            D3DXVec3Add vBillardVel, vBillardVel, ScaleVector(m_vBVelocities(i), m_BInertia * 2)

                            m_vBVelocities(i) = vBillardVel
                            'Update the value of momentum for a billard - table collision
                            m_CBillTableMomentum = m_CBillTableMomentum + CMmntm
                            If iNorm = 3 Then 'ball hits straight edge of bumper
                                TimeoutCount = TimeoutCount + 1
                              Else 'NOT INORM...
                                TimeoutCount = 0 'restart timeout count on sloped edge hit
                            End If
                        End If
                    End If
                Next j
                'Now check for a pocket hit
                'Compute the point in time, at which the billard will
                'be at the shortest distance from the pocket and then
                'check exactly how close will it be.
                With vDist
                    For k = 1 To 6
                        D3DXVec3Subtract vDist, vBillardPos, m_vTPockets(k)
                        'First check if the billard is in the pocket already
                        If (.x * .x + .z * .z) <= m_TPocketRadSqrd Then
                            GoSub PocketHit
                            'Jump to another billard
                            Exit For 'loop varying k
                          Else 'NOT (.X...
                            'The second method for finding pocket hits is to compute the point in time,
                            'when the billard is at the shortest distance from the pocket's centre
                            MinDistTime = -DotProduct(vDist, vBillardVel) / BillardVelSqrd
                            'If the point in time is between the beginning of the time frame (0) and its end (g_dt)...
                            If MinDistTime >= 0 And MinDistTime <= g_dt Then
                                D3DXVec3Add vDist, vDist, ScaleVector(vBillardVel, MinDistTime)
                                '... and if the squared module of the minimum distance vector is less or equal to
                                'the squared radius of the pocket, then we have a pocket hit.
                                If (.x * .x + .z * .z) <= m_TPocketRadSqrd Then
                                    'The i'th billard is about to fall into the k'th pocket
                                    GoSub PocketHit
                                    'Jump to another billard
                                    Exit For 'loop varying k
                                End If
                            End If
                        End If
                    Next k
                End With 'VDIST
                BallsMoving = BallsMoving + 1
            End If
        End If
    Next i

Exit Sub

PocketHit:
    m_bPocketHitDetected = True
    'Notify the i'th billard that it has fallen into the k'th pocket
    m_cBillards.FellInPocketNumber(i) = k
    'Set the billard's position to the centre of the k'th pocket...
    m_vBPositions(i) = m_vTPockets(k)
    '... and stop its movement.
    m_vBVelocities(i) = MakeVector(0, 0, 0)
    FastDecelerate = FastDecelerate Or (i = 8)
    If i = 0 Or (FastDecelerate And m_bBSinking8Illegal) Then 'cue-ball or 8-ball illegally in pocket
        m_Stickers(SC_FOUL).Visible = True
        PlaySound m_DSOops, 1
      Else 'NOT I...
        SinkCounts(m_CurrentPlayer) = SinkCounts(m_CurrentPlayer) + 1
        DisplayCounts
    End If
    TimeoutCount = 0 'restart timeout count on pocket hit
    Return

End Sub

Friend Property Get PocketHitDetected() As Boolean

    PocketHitDetected = m_bPocketHitDetected

End Property

Friend Sub Setup(ByVal Billards As clsBillards, ByVal Table As clsTable)

    Set m_cBillards = Billards
    Set m_Table = Table

    'Get the relevant data from the Billards object
    With m_cBillards
        'We will need the total number of billards...
        m_numBillards = .NumBillards
        '... and their physical properties
        .GetPhysBillConstants m_BMass, m_BInertia, m_BRadius, m_BElasticity
    End With 'M_CBILLARDS
    'Resize the relevant arrays
    ReDim m_vBPositions(0 To m_numBillards - 1)
    ReDim m_vBVelocities(0 To m_numBillards - 1)
    ReDim m_vBSpins(0 To m_numBillards - 1)
    ReDim m_bBInTheGame(0 To m_numBillards - 1)
    'Compute the m_2BRadSqrd variable
    m_2BRadSqrd = (2 * m_BRadius) ^ 2

    'Get the relevant data from the PoolTable object
    With m_Table
        'We will need the bumpers,
        .GetBumpers m_TBumpers
        'their physical properties,
        m_TBumperElasticity = .GetPhysTableConstants
        'and the pockets
        .GetPockets m_vTPockets, a 'a has the pocket radius
    End With 'M_TABLE
    'compute m_TPocketRadSqrd
    m_TPocketRadSqrd = a * a * 1.66 '1.66 is a "pocket attractor" to make sinking a little easier

    'Bumper architecture
    'A typical bumper looks like this

    '                1|2                            5
    '                  \                           /
    '                   \                         /
    '                    3 --------------------- 4

    'The numbers refer to vertices (vectors) in the m_TBumpers array.
    'The first and second vertices are actually the same vertex.
    'This way we will have five normals and five vertices. This will proove
    'very convenient, when the bumper vertices and bumper normal vectors will be
    'used together for collision detection.
    'Compute the bumpers'normal vectors.
    'This is done by calculating the cross product of a vector running along the bumper's edge and the Y-axis
    For i = 1 To m_numTableBumpers
        For j = 2 To 4
            D3DXVec3Subtract vVctr1, m_TBumpers(i, j + 1), m_TBumpers(i, j) 'vVctr1 is the vector running along the bumper's edge
            vVctr1.y = 0 '"Flatten" the vector, just to be sure...
            D3DXVec3Cross vVctr1, vVctr1, MakeVector(0, 1, 0) 'This is the cross product with the y-axis.
            D3DXVec3Normalize m_TBumperNorms(i, j), vVctr1  'The normal vector should be of unit length.
        Next j
        'The first and the last normal vectors are obtained with a cross product
        'of the third normal and the y-axis (pointing upwards or downwards).
        D3DXVec3Cross m_TBumperNorms(i, 1), m_TBumperNorms(i, 3), MakeVector(0, 1, 0)
        D3DXVec3Cross m_TBumperNorms(i, 5), m_TBumperNorms(i, 3), MakeVector(0, -1, 0)
    Next i

End Sub

Public Property Let Sinking8Illegal(bVal As Boolean)

    m_bBSinking8Illegal = (bVal <> False)

End Property

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-29 14:15)  Decl: 67  Code: 400  Total: 467 Lines
':) CommentOnly: 114 (24,4%)  Commented: 89 (19,1%)  Empty: 63 (13,5%)  Max Logic Depth: 9
