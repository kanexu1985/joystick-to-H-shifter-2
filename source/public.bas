Attribute VB_Name = "public"
Option Explicit

'v0.9 defs:
Public justPressed As Boolean
Public justCentered As Boolean
Public strDemo As String

'v0.2 defs:

Public center As Boolean
Public centerl As Boolean
Public dz As Long
Public dzl As Long

Public vers As String

Public temp1 As Long, temp2 As Long


Public joyID As Integer            'which joystick

Public Type JOYCAPS
  wMid As Integer
  wPid As Integer
  szPname As String * 32
  wXmin As Long
  wXmax As Long
  wYmin As Long
  wYmax As Long
  wZmin As Long
  wZmax As Long
  wNumButtons As Long
  wPeriodMin As Long
  wPeriodMax As Long
  wRmin As Long
  wRmax As Long
  wUmin As Long
  wUmax As Long
  wVmin As Long
  wVmax As Long
  wMaxAxes As Long
  wNumAxes As Long
  wMaxButtons As Long
  szRegKey As String * 32
  szOEMVxD As String * 240
End Type

Public JoyDInfo As JOYCAPS  ' receives joystick information
Public joyDriver As String  ' will be set to the joystick's driver name


Public Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long


Public Const JOY_BUTTON1 = &H1
Public Const JOY_BUTTON2 = &H2
Public Const JOY_BUTTON3 = &H4
Public Const JOY_BUTTON4 = &H8
Public Const JOY_BUTTON5 = &HF
Public Const JOY_BUTTON6 = &H20
Public Const JOY_BUTTON7 = &H40
Public Const JOY_BUTTON8 = &H80


Public step1 As Long
Public step2 As Long
Public step3 As Long
Public step4 As Long

Public Type JOYINFO
  wXpos As Long
  wYpos As Long
  wZpos As Long
  wButtons As Long
End Type

Public joyPos As JOYINFO   'joy actions
Public joyPos2 As JOYINFO
Public joyCursor As JOYINFO 'cursor displayed as green x mark
Public pointDelta As JOYINFO 'delta data between current joypos and effect zone --while draging zone

Public Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long

Public Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer

Public JoyButton As Long
Public KeyButton As Integer

Public userJoy As JOYINFO
Public userKey As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


'switch state of joystick buttons 1 & 2:
Public joyBut1 As Boolean
Public joyBut2 As Boolean







Public Function Press(step As Long, Delayt As Integer, if_state As Boolean) As Boolean

    If if_state = True Or step > 0 Then step = step + 1

    If step > Delayt Then step = 0

    If step = 1 Then
      Press = True
      Else
      Press = False
    End If

End Function












