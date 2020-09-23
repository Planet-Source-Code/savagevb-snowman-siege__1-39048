Attribute VB_Name = "KeyboardMod"
Option Explicit
Public DKI As DirectInput8
Public DKIDevice As DirectInputDevice8
Public DKIState As DIKEYBOARDSTATE
Public KeyState(0 To 255) As Boolean    'so we can detect if the key has gone up or down!
Public Const KeyBufferSize As Long = 10 'how many events the buffer holds.
Public DevProp As DIPROPLONG
'====================================================================================
Public Function KeyCheck(): Dim B As Integer
    DKIDevice.GetDeviceStateKeyboard DKIState 'get the keyboard state
    For B = LBound(DKIState.Key()) To UBound(DKIState.Key())
        DoEvents
        If DKIState.Key(B) = 128 And (Not KeyState(B) = True) Then
            KeyState(B) = True: KeyEvent B, "D"
        ElseIf DKIState.Key(B) = 0 And (Not KeyState(B) = False) Then
            KeyState(B) = False: KeyEvent B, "U"
        End If
    Next B
End Function
'====================================================================================
Private Function KeyEvent(ByVal B As Integer, ByVal KTag As String): If KTag = Empty Then Exit Function
    If KTag = "D" Then
        'do down functions
    ElseIf KTag = "U" Then  'do the functions i want on keyup
        Select Case B
            Case 1: Running = False 'the Esc key was pressed and released
        End Select
    End If
End Function
'====================================================================================
