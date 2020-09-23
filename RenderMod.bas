Attribute VB_Name = "RenderMod"
Option Explicit
Private ColFlag As Boolean  'collision flag
Public Function Render() 'normal render state
    Dim W As Integer: ColFlag = False
    
    'Clear the screen so we have a blank canvas to paint on
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    
    'start the rendering scene *note : anything that is to be rendered must be within the begin and end scene tags
    D3DDevice.BeginScene
        'if they are allowed to be rendered then draw our objects
        For W = 0 To UBound(WallMesh)
            If WallMesh(W).RenderMe Then RenderMesh WallMesh(W), WallMesh(W).MX, WallMesh(W).MY, WallMesh(W).MZ, WallMesh(W).MAngle
        Next W
        For W = 0 To UBound(GateMesh)
            If GateMesh(W).RenderMe Then RenderMesh GateMesh(W), GateMesh(W).MX, GateMesh(W).MY, GateMesh(W).MZ
        Next W
        For W = 0 To UBound(WorldMesh)
            If WorldMesh(W).RenderMe Then RenderMesh WorldMesh(W), WorldMesh(W).MX, WorldMesh(W).MY, WorldMesh(W).MZ
        Next W
        For W = 0 To UBound(TreeMesh)
            If TreeMesh(W).RenderMe Then RenderMesh TreeMesh(W), TreeMesh(W).MX, TreeMesh(W).MY, TreeMesh(W).MZ
        Next W
        For W = 0 To UBound(HouseMesh)
            If HouseMesh(W).RenderMe Then RenderMesh HouseMesh(W), HouseMesh(W).MX, HouseMesh(W).MY, HouseMesh(W).MZ
        Next W
        For W = 0 To UBound(RoadMesh)
            If RoadMesh(W).RenderMe Then RenderMesh RoadMesh(W), RoadMesh(W).MX, RoadMesh(W).MY, RoadMesh(W).MZ
        Next W
        
        'Setup the New Matrix (*move the objects*)
        MatrixSetUp
                
        'Do the FPS count
        If (GetTickCount() - LastTickCount) >= 1000 Then
            LastFrameCount = FrameCount
            FrameCount = 0
            LastTickCount = GetTickCount()
        Else: FrameCount = FrameCount + 1
        End If
        
        'Setup our Text boxes on the screen
        TextRect(0).Top = 1: TextRect(0).bottom = 300: TextRect(0).Left = 1: TextRect(0).Right = 250
        TextRect(1).Top = 1: TextRect(1).Left = ScreenWidth - 200: TextRect(1).bottom = 200: TextRect(1).Right = ScreenWidth
        TextRect(2).Top = ScreenHeight - 200: TextRect(2).bottom = ScreenHeight: TextRect(2).Left = ScreenWidth - 400: TextRect(2).Right = ScreenWidth
        TextRect(3).Top = ScreenHeight - 200: TextRect(3).bottom = ScreenHeight: TextRect(3).Left = 0: TextRect(3).Right = 200
        
        'Write in the text boxes
        D3DX.DrawText D3DFont(0), &HFFFFCC00, "SnowMan's Stats" _
            & vbCrLf & "--------------" _
            & vbCrLf & "Life " & "100" & "/" & "100" _
            & vbCrLf & "Number of Kills : " & nKills, _
            TextRect(0), DT_TOP Or DT_LEFT
        
        D3DX.DrawText D3DFont(0), &HFFFFCC00, " Evil SnowMan's Stats" _
            & vbCrLf & "--------------------" _
            & vbCrLf & "Life " & EvlHealth & "/" & "100", _
            TextRect(1), DT_TOP Or DT_RIGHT
        
        D3DX.DrawText D3DFont(0), &HFFFFCC00, "SavageVB's 3D World" _
            & vbCrLf & "Current FPS : " & LastFrameCount, TextRect(2), DT_BOTTOM Or DT_RIGHT

        '------------------------------------------------------
        '   Count Down Section
        '------------------------------------------------------
        If (GetTickCount() - LastcDownCheck) >= 1000 Then CountDown = CountDown + 1: LastcDownCheck = GetTickCount()
        D3DX.DrawText D3DFont(0), &HFFFFCC00, "Time Left : " _
            & (tLIMIT / 1000) - CountDown & " Seconds", TextRect(3), DT_BOTTOM Or DT_LEFT
        
        'enable the lights
        D3DDevice.LightEnable 0, True
    
    'End the rendering Scene
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Function
'==================================================================================
Public Function RenderForm(): Dim W As Integer
    Dim matTemp As D3DMATRIX, matCamera As D3DMATRIX, matRotation As D3DMATRIX
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0  '//Clear the screen black
    D3DDevice.BeginScene
        For W = 0 To UBound(WallMesh)
            RenderMesh WallMesh(W), WallMesh(W).MX, WallMesh(W).MY, WallMesh(W).MZ, WallMesh(W).MAngle
        Next W
        For W = 0 To UBound(WorldMesh)
            RenderMesh WorldMesh(W), WorldMesh(W).MX, WorldMesh(W).MY, WorldMesh(W).MZ
        Next W
        For W = 0 To UBound(TreeMesh)
            RenderMesh TreeMesh(W), TreeMesh(W).MX, TreeMesh(W).MY, TreeMesh(W).MZ
        Next W
        For W = 0 To UBound(HouseMesh)
            RenderMesh HouseMesh(W), HouseMesh(W).MX, HouseMesh(W).MY, HouseMesh(W).MZ
        Next W
        For W = 0 To UBound(RoadMesh)
            RenderMesh RoadMesh(W), RoadMesh(W).MX, RoadMesh(W).MY, RoadMesh(W).MZ
        Next W
        
        SnowMesh(0).MAngle = SnowMesh(0).MAngle - (TSPEED / 2)
        If SnowMesh(0).MAngle < 0 Then SnowMesh(0).MAngle = D_360 + SnowMesh(0).MAngle
        RenderMesh SnowMesh(0), -12, -40, 1, SnowMesh(0).MAngle
        
        SnowEvlMesh(0).MAngle = SnowEvlMesh(0).MAngle + (TSPEED / 2)
        If SnowEvlMesh(0).MAngle > D_360 Then SnowEvlMesh(0).MAngle = 0 + (SnowEvlMesh(0).MAngle - D_360)
        RenderMesh SnowEvlMesh(0), -38, -40, 1, SnowEvlMesh(0).MAngle
        
        RenderMesh FormMesh(0), -25, -40, 10
        
        TextRect(0).bottom = ScreenHeight - 300
        TextRect(0).Left = 200
        TextRect(0).Right = ScreenWidth - 200
        TextRect(0).Top = 400
        
        D3DX.DrawText D3DFont(1), &HFFFFCC00, "Congratulations you killed " & nKills _
        & " Evil Snowmen" & vbCrLf & vbCrLf & "Please Press Esc to finish", TextRect(0), DT_TOP Or DT_CENTER
        
        D3DXMatrixRotationY matRotation, 0
        D3DXMatrixMultiply matCamera, matCamera, matRotation
        
        D3DXMatrixTranslation matCamera, CAMX, CAMY + 20, -15
        D3DXMatrixMultiply matCamera, matCamera, matView
        
        D3DDevice.SetTransform D3DTS_VIEW, matCamera
        
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Function
'==================================================================================
Private Sub MatrixSetUp()
    
    Dim matTemp As D3DMATRIX, matCamera As D3DMATRIX, matRotation As D3DMATRIX
    Dim TempChrX As Single, TempChrY As Single, TempCAMX As Single, TempCAMY As Single, TempEvlX As Single, TempEvlY As Single
    
    D3DXMatrixIdentity matCamera
    D3DXMatrixIdentity matRotation
    
    TempEvlX = SnowEvlMesh(0).MX
    TempEvlY = SnowEvlMesh(0).MY
    TempChrX = ChrX: TempChrY = ChrY
    TempCAMX = CAMX: TempCAMY = CAMY
    
    DIMouse.GetDeviceStateMouse DIMState
    DKIDevice.GetDeviceStateKeyboard DKIState
    
'======================================================================================
'   Movement by Keyboard Section
'======================================================================================
    
    If DKIState.Key(30) <> 0 Then
        If DKIState.Key(42) <> 0 Then MSPEED = 3 Else MSPEED = 2
        ChrX = ChrX + (Sin(ChrAngle + D_90) * MSPEED)
        ChrY = ChrY + (Cos(ChrAngle + D_90) * MSPEED)
        CAMX = CAMX + (Sin(ChrAngle + D_90) * MSPEED)
        CAMY = CAMY + (Cos(ChrAngle + D_90) * MSPEED)
    ElseIf DKIState.Key(32) <> 0 Then
        If DKIState.Key(42) <> 0 Then MSPEED = 3 Else MSPEED = 2
        ChrX = ChrX - (Sin(ChrAngle + D_90) * MSPEED)
        ChrY = ChrY - (Cos(ChrAngle + D_90) * MSPEED)
        CAMX = CAMX - (Sin(ChrAngle + D_90) * MSPEED)
        CAMY = CAMY - (Cos(ChrAngle + D_90) * MSPEED)
    End If
    
    If DKIState.Key(200) <> 0 Then
        'Move the camera and character (Snowman) forwards
        If DKIState.Key(42) <> 0 Then MSPEED = 3 Else MSPEED = 2
        ChrX = ChrX - (Sin(D_360 - ChrAngle) * MSPEED)
        ChrY = ChrY + (Cos(D_360 - ChrAngle) * MSPEED)
        CAMX = CAMX - (Sin(D_360 - ChrAngle) * MSPEED)
        CAMY = CAMY + (Cos(D_360 - ChrAngle) * MSPEED)
    ElseIf DKIState.Key(208) <> 0 Then
        'Move the camera and character (Snowman) backwards
        If DKIState.Key(42) <> 0 Then MSPEED = 3 Else MSPEED = 2
        ChrX = ChrX + (Sin(D_360 - ChrAngle) * MSPEED)
        ChrY = ChrY - (Cos(D_360 - ChrAngle) * MSPEED)
        CAMX = CAMX + (Sin(D_360 - ChrAngle) * MSPEED)
        CAMY = CAMY - (Cos(D_360 - ChrAngle) * MSPEED)
    End If
    
'======================================================================================
'    Collision detection for Snowman Section
'======================================================================================
    Dim I As Integer
    For I = 0 To UBound(HouseMesh())
        If MeshColDetect(HouseMesh(I), SnowMesh(0), ChrX, ChrY) Then ChrX = TempChrX: ChrY = TempChrY: CAMX = TempCAMX: CAMY = TempCAMY: Exit For
    Next I
    For I = 0 To UBound(TreeMesh())
        If MeshColDetect(TreeMesh(I), SnowMesh(0), ChrX, ChrY) Then ChrX = TempChrX: ChrY = TempChrY: CAMX = TempCAMX: CAMY = TempCAMY: Exit For
    Next I
    For I = 0 To UBound(WallMesh())
        If MeshColDetect(WallMesh(I), SnowMesh(0), ChrX, ChrY) Then ChrX = TempChrX: ChrY = TempChrY: CAMX = TempCAMX: CAMY = TempCAMY: Exit For
    Next I
    For I = 0 To UBound(GateMesh())
        If MeshColDetect(GateMesh(I), SnowMesh(0), ChrX, ChrY) Then ChrX = TempChrX: ChrY = TempChrY: CAMX = TempCAMX: CAMY = TempCAMY: Exit For
    Next I
    If SnowEvlMesh(0).RenderMe Then If MeshColDetect(SnowEvlMesh(0), SnowMesh(0), ChrX, ChrY) Then ChrX = TempChrX: ChrY = TempChrY: CAMX = TempCAMX: CAMY = TempCAMY
    SnowMesh(0).MX = -ChrX: SnowMesh(0).MY = -ChrY
    
'======================================================================================
'   Movement by Mouse Section
'======================================================================================
    If MouX < 100 Then
        CAMX = CAMX - (Cos(-Angle) * TPOWER)
        CAMY = CAMY - (Sin(-Angle) * TPOWER)
        Angle = Angle + TSPEED
        If Angle > D_360 Then Angle = 0 + (Angle - D_360)
    ElseIf MouX > (frmMain.Width - 100) Then
        Angle = Angle - TSPEED
        If Angle < 0 Then Angle = D_360 + Angle
        CAMX = CAMX + (Cos(-Angle) * TPOWER)
        CAMY = CAMY + (Sin(-Angle) * TPOWER)
    End If
    
'======================================================================================
'    Turning Section
'======================================================================================
    If DKIState.Key(205) <> 0 Then
        'Rotate the snowman to the right while moving the camera to the left
        Angle = Angle - TSPEED: ChrAngle = ChrAngle - TSPEED
        If Angle < 0 Then Angle = D_360 + Angle
        If ChrAngle < 0 Then ChrAngle = D_360 + ChrAngle
        CAMX = CAMX + (Cos(-Angle) * TPOWER)
        CAMY = CAMY + (Sin(-Angle) * TPOWER)
    End If
    
    If DKIState.Key(203) <> 0 Then
        'Rotate the Snowman to the left while moving the camera to the right
        CAMX = CAMX - (Cos(-Angle) * TPOWER)
        CAMY = CAMY - (Sin(-Angle) * TPOWER)
        Angle = Angle + TSPEED: ChrAngle = ChrAngle + TSPEED
        If Angle > D_360 Then Angle = 0 + (Angle - D_360)
        If ChrAngle > D_360 Then ChrAngle = 0 + (ChrAngle - D_360)
    End If
    
    SnowMesh(0).MAngle = (D_360 - ChrAngle) + D_180 'Rotate the snowman

'======================================================================================
'   Speedup Section
'======================================================================================
    'This might not do much on a project this size, but on a larger project it could make quite a bit of difference
    DoEvents
    For I = 0 To UBound(TreeMesh): CheckWithinSite TreeMesh(I), Angle: Next I
    For I = 0 To UBound(WallMesh): CheckWithinSite WallMesh(I), Angle: Next I
    For I = 0 To UBound(GateMesh): CheckWithinSite GateMesh(I), Angle: Next I
        
'======================================================================================
'   Snowball Section
'======================================================================================
    '----Create a snow ball it the left control has been pressed---
    If DKIState.Key(29) <> 0 Then
        If (GetTickCount() - LastThrowTime) >= ThrowSpeed Then
            ReDim Preserve ThrowMesh(ThrowCount)
            CreateSnowMeshObj TemplateThrowMesh, ThrowMesh, ThrowCount, 15, 2.5, 2.5, 2.5, 100, 0, _
            -(SnowMesh(0).MX + (SnowMesh(0).MWidth / 2)), -SnowMesh(0).MY, SnowMesh(0).MAngle
            ThrowCount = ThrowCount + 1
            LastThrowTime = GetTickCount()
        End If
    End If

    If ThrowCount > 0 Then
        Dim W As Integer, OldX As Single, OldY As Single
        For W = 0 To UBound(ThrowMesh)
            If ThrowMesh(W).RenderMe = True Then
                OldX = ThrowMesh(W).MX: OldY = ThrowMesh(W).MY
                                
                '----Move the Snowball and do some collision detection----
                ThrowMesh(W).MX = ThrowMesh(W).MX + (Sin(D_360 - ThrowMesh(W).MAngle) * 3.5)
                ThrowMesh(W).MY = ThrowMesh(W).MY - (Cos(D_360 - ThrowMesh(W).MAngle) * 3.5)
                For I = 0 To UBound(TreeMesh)
                    If MeshColDetect(TreeMesh(I), ThrowMesh(W), ThrowMesh(W).MX, ThrowMesh(W).MY) Then GoTo SkipColDetect
                Next I
                For I = 0 To UBound(HouseMesh)
                    If MeshColDetect(HouseMesh(I), ThrowMesh(W), ThrowMesh(W).MX, ThrowMesh(W).MY) Then GoTo SkipColDetect
                Next I
                For I = 0 To UBound(WallMesh)
                    If MeshColDetect(WallMesh(I), ThrowMesh(W), ThrowMesh(W).MX, ThrowMesh(W).MY) Then GoTo SkipColDetect
                Next I
                For I = 0 To UBound(GateMesh())
                    If MeshColDetect(GateMesh(I), ThrowMesh(W), ThrowMesh(W).MX, ThrowMesh(W).MY) Then GoTo SkipColDetect
                Next I
                
                '----Check for any collisions against the evil snowman------
                If MeshColDetect(SnowEvlMesh(0), ThrowMesh(W), ThrowMesh(W).MX, ThrowMesh(W).MY) Then
                    EvlHealth = EvlHealth - Int((15 * Rnd) + 5): Randomize
                    If EvlHealth <= 0 Then EvlHealth = 100: nKills = nKills + 1
                    'When the evil snow man gets hit teleport it to a new location
                    Select Case Int((4 * Rnd) + 1)
                    Case 1: SnowEvlMesh(0).MX = -200: SnowEvlMesh(0).MY = -200
                    Case 2: SnowEvlMesh(0).MX = 200: SnowEvlMesh(0).MY = -200
                    Case 3: SnowEvlMesh(0).MX = -200: SnowEvlMesh(0).MY = 200
                    Case 4: SnowEvlMesh(0).MX = 200: SnowEvlMesh(0).MY = 200
                    End Select
                End If
                
SkipColDetect:
                '----Create an explosion if there was a collision-----
                If ColFlag Then
                    ThrowMesh(W).MX = OldX: ThrowMesh(W).MY = OldY
                    ThrowMesh(W).RenderMe = False
                    ReDim Preserve ExpMesh(ExpCount)
                    CreateAnimMeshObj TemplateExpMesh, ExpMesh(ExpCount), _
                        ThrowMesh(W).MX, ThrowMesh(W).MY, ThrowMesh(W).MZ, ThrowMesh(W).MAngle
                    ExpCount = ExpCount + 1
                End If
                '----Render the Snowball if it hasn't finished its life----
                If ThrowMesh(W).Turns < ThrowMesh(W).LifeSpan Then
                    ThrowMesh(W).Turns = ThrowMesh(W).Turns + 1
                    RenderMesh ThrowMesh(W), ThrowMesh(W).MX, ThrowMesh(W).MY, ThrowMesh(W).MZ, ThrowMesh(W).MAngle
                Else: ThrowMesh(W).RenderMe = False
                End If
            End If
        Next W
    End If
        
    '-----Render any explosions-------
    If ExpCount > 0 Then
        For I = 0 To UBound(ExpMesh)
            If ExpMesh(I).RenderMe Then
                If Not ExpMesh(I).AnimDMesh(ExpMesh(I).AnimTCurrent).AnimTIndex >= ExpMesh(I).AnimDMesh(ExpMesh(I).AnimTCurrent).AnimTLength Then
                    ExpMesh(I).AnimDMesh(ExpMesh(I).AnimTCurrent).AnimTIndex = ExpMesh(I).AnimDMesh(ExpMesh(I).AnimTCurrent).AnimTIndex + 1
                    RenderAnim ExpMesh(I), ExpMesh(I).AnimTCurrent
                Else
                    If ExpMesh(I).AnimTCurrent = 5 Then
                        ExpMesh(I).RenderMe = True
                    Else: ExpMesh(I).AnimTCurrent = ExpMesh(I).AnimTCurrent + 1
                    End If
                End If
            End If
        Next I
    End If

'======================================================================================
'   Drop Marker Section
'======================================================================================
    If DKIState.Key(57) <> 0 Then
        If (GetTickCount() - LastDropTime) >= DropSpeed Then
            ReDim Preserve DropMesh(DropCount)
            CreateSnowMeshObj TemplateDropMesh, DropMesh, DropCount, 0, 5, 5, 5, 0, 0, _
            -SnowMesh(0).MX, -SnowMesh(0).MY, SnowMesh(0).MAngle
            DropCount = DropCount + 1
            LastDropTime = GetTickCount()
        End If
    End If
    
    If DropCount > 0 Then
        For W = 0 To UBound(DropMesh)
            'Rotate and render the dropped object
            DropMesh(W).MAngle = DropMesh(W).MAngle + TSPEED
            If DropMesh(W).MAngle > D_360 Then DropMesh(W).MAngle = 0 + (DropMesh(W).MAngle - D_360)
            RenderMesh DropMesh(W), DropMesh(W).MX, DropMesh(W).MY, DropMesh(W).MZ, DropMesh(W).MAngle
        Next W
    End If

'======================================================================================
'   Evil Snowman Section
'======================================================================================
    GenerateMovementForAI
    RenderMesh SnowEvlMesh(0), SnowEvlMesh(0).MX, SnowEvlMesh(0).MY, 0, -SnowEvlMesh(0).MAngle + D_180

'======================================================================================
'   Matrix Section
'======================================================================================
    
    RenderMesh SnowMesh(0), -SnowMesh(0).MX, -SnowMesh(0).MY, 0, SnowMesh(0).MAngle
    
    D3DXMatrixTranslation matCamera, -CAMX, -CAMY + CamObDist, -CAMZ - CamGrDist
    D3DXMatrixMultiply matCamera, matCamera, matView
    
    D3DXMatrixRotationY matRotation, Angle
    D3DXMatrixMultiply matCamera, matCamera, matRotation
    
    D3DXMatrixRotationX matRotation, Pitch
    D3DXMatrixMultiply matCamera, matCamera, matRotation
    
    D3DDevice.SetTransform D3DTS_VIEW, matCamera
End Sub
'==================================================================================
Public Function RenderMesh(InMatrix As MeshData, InX As Single, InY As Single, InZ As Single, Optional RAngle As Single = 0, Optional RPitch As Single = 0): Dim RenderTempMat As D3DMATRIX, I As Integer
    D3DXMatrixIdentity RenderTempMat
    D3DXMatrixIdentity InMatrix.Matrix
    D3DXMatrixRotationZ InMatrix.Matrix, RAngle
    D3DXMatrixMultiply RenderTempMat, RenderTempMat, InMatrix.Matrix
    D3DXMatrixRotationX InMatrix.Matrix, RPitch
    D3DXMatrixMultiply RenderTempMat, RenderTempMat, InMatrix.Matrix
    D3DXMatrixTranslation InMatrix.Matrix, InX, InY, InZ
    D3DXMatrixMultiply RenderTempMat, RenderTempMat, InMatrix.Matrix
    D3DDevice.SetTransform D3DTS_WORLD, RenderTempMat

    For I = 0 To InMatrix.MatCount - 1
        D3DDevice.SetMaterial InMatrix.Mat(I)
        D3DDevice.SetTexture 0, InMatrix.Tex(I)
        InMatrix.Mesh.DrawSubset I
    Next
End Function
Public Function RenderAnim(InMatrix As AnimMeshData, rNum As Long): Dim RenderTempMat As D3DMATRIX, I As Integer
    D3DXMatrixIdentity RenderTempMat
    D3DXMatrixIdentity InMatrix.AnimMatrix
    D3DXMatrixRotationZ InMatrix.AnimMatrix, InMatrix.AnimAngle
    D3DXMatrixMultiply RenderTempMat, RenderTempMat, InMatrix.AnimMatrix
    D3DXMatrixTranslation InMatrix.AnimMatrix, InMatrix.AnimX, InMatrix.AnimY, InMatrix.AnimZ
    D3DXMatrixMultiply RenderTempMat, RenderTempMat, InMatrix.AnimMatrix
    D3DDevice.SetTransform D3DTS_WORLD, RenderTempMat
    
    For I = 0 To InMatrix.AnimDMesh(rNum).AnimMCount - 1
        D3DDevice.SetMaterial InMatrix.AnimDMesh(rNum).AnimMat(I)
        D3DDevice.SetTexture 0, InMatrix.AnimDMesh(rNum).AnimTex(I)
        InMatrix.AnimDMesh(rNum).AnimFMesh.DrawSubset I
    Next
End Function
'==================================================================================
Public Function MeshColDetect(InMesh As MeshData, InChar As MeshData, IX As Single, IY As Single) As Boolean: MeshColDetect = False: ColFlag = False
    'do some basic collision detection, i know this is by far the best way of doing it but it works and its pretty fast
    If ((IX + (InChar.MWidth / 2)) > (InMesh.MX - (InMesh.MWidth / 2))) And ((IX - (InChar.MWidth / 2)) < (InMesh.MX + (InMesh.MWidth / 2))) Then
        If ((IY + (InChar.MLength / 2)) > (InMesh.MY - (InMesh.MLength / 2))) And ((IY - (InChar.MLength / 2)) < (InMesh.MY + (InMesh.MLength / 2))) Then
            MeshColDetect = True: ColFlag = True
        End If
    End If
End Function
'==================================================================================
Public Function CheckWithinSite(InMesh As MeshData, CAngle As Single)
    'This function basically cuts the world into 90 degree parts and checks if the
    'inmesh object is not in sight of the camera, this doesn't work properly on larger object such as the ground
    If CAngle < D_90 Then
        If InMesh.MY < (ChrY - CamObDist) Then InMesh.RenderMe = False Else InMesh.RenderMe = True
    ElseIf CAngle < D_180 And CAngle > D_90 Then
        If InMesh.MX < (ChrX - CamObDist) Then InMesh.RenderMe = False Else InMesh.RenderMe = True
    ElseIf CAngle < D_270 And CAngle > D_180 Then
        If InMesh.MY > (ChrY + CamObDist) Then InMesh.RenderMe = False Else InMesh.RenderMe = True
    ElseIf CAngle > D_270 Then
        If InMesh.MX > (ChrX + CamObDist) Then InMesh.RenderMe = False Else InMesh.RenderMe = True
    End If
End Function
Public Function GenerateMovementForAI(): Dim tAngle As Single: Randomize
    'Move the evil snowman depending on a random direction, then do some collision detection on the evil snowman
    SnowEvlMesh(0).MX = SnowEvlMesh(0).MX - (Sin(D_360 - SnowEvlMesh(0).MAngle))
    SnowEvlMesh(0).MY = SnowEvlMesh(0).MY + (Cos(D_360 - SnowEvlMesh(0).MAngle))
    If (GetTickCount() - LastAIMove) >= 200 Then dNum = Int((100 * Rnd) + 1): LastAIMove = GetTickCount()
    If dNum < 20 Then SnowEvlMesh(0).MAngle = SnowEvlMesh(0).MAngle + TSPEED
    If dNum > 20 And dNum < 80 Then SnowEvlMesh(0).MAngle = SnowEvlMesh(0).MAngle
    If dNum > 80 Then SnowEvlMesh(0).MAngle = SnowEvlMesh(0).MAngle - TSPEED
    Dim I As Integer
    For I = 0 To UBound(HouseMesh())
        If MeshColDetect(HouseMesh(I), SnowEvlMesh(0), SnowEvlMesh(0).MX, SnowEvlMesh(0).MY) Then SnowEvlMesh(0).MAngle = SnowEvlMesh(0).MAngle + D_180: Exit For
    Next I
    For I = 0 To UBound(TreeMesh())
        If MeshColDetect(TreeMesh(I), SnowEvlMesh(0), SnowEvlMesh(0).MX, SnowEvlMesh(0).MY) Then SnowEvlMesh(0).MAngle = SnowEvlMesh(0).MAngle + D_180: Exit For
    Next I
    For I = 0 To UBound(WallMesh())
        If MeshColDetect(WallMesh(I), SnowEvlMesh(0), SnowEvlMesh(0).MX, SnowEvlMesh(0).MY) Then SnowEvlMesh(0).MAngle = SnowEvlMesh(0).MAngle + D_180: Exit For
    Next I
    For I = 0 To UBound(GateMesh())
        If MeshColDetect(GateMesh(I), SnowEvlMesh(0), SnowEvlMesh(0).MX, SnowEvlMesh(0).MY) Then SnowEvlMesh(0).MAngle = SnowEvlMesh(0).MAngle + D_180: Exit For
    Next I
End Function
