VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   14295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load(): Dim I As Integer
    'Init all the variables to their starting values
    CAMX = 0: CAMY = -160: CAMZ = 0: ChrX = 0: ChrY = -160: MouX = 0: MouY = 0
    CamObDist = 75: CamGrDist = 75: Pitch = -0.55: Angle = D_360: MSPEED = 2
    ChrAngle = D_360
    
    InitDX 'call the initDX function

    Do While Running
        KeyCheck 'check for any key presses
        'check to see how long the program has been running and if its > the tLimit display the end scene
        If (GetTickCount() - LasttLIMITCheck) < tLIMIT Then
            Render  'normal render state
        Else
            'This inner while loop is to allow these variables below to initialized only once
            SnowEvlMesh(0).MAngle = 0:  SnowMesh(0).MAngle = 0
            CAMX = 25: CAMY = 60
            Do While Running
                KeyCheck 'check for any key presses
                RenderForm 'end scene
            Loop
        End If
    Loop
    
    'Do A Little house keeping
    Set DKIDevice = Nothing
    Set DIMouse = Nothing
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set DX = Nothing
    'then kill the program
    End
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouX = X: MouY = Y 'just get and set the mouse co-ordinates
End Sub
