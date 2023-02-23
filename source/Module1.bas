Attribute VB_Name = "Module_sendInput"
Option Explicit


Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CANCEL = &H3
Public Const VK_MBUTTON = &H4
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_CLEAR = &HC
Public Const VK_RETURN = &HD
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_PAUSE = &H13
Public Const VK_CAPITAL = &H14
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_SELECT = &H29
Public Const VK_PRINT = &H2A
Public Const VK_EXECUTE = &H2B
Public Const VK_SNAPSHOT = &H2C
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_HELP = &H2F
Public Const VK_0 = &H30
Public Const VK_1 = &H31
Public Const VK_2 = &H32
Public Const VK_3 = &H33
Public Const VK_4 = &H34
Public Const VK_5 = &H35
Public Const VK_6 = &H36
Public Const VK_7 = &H37
Public Const VK_8 = &H38
Public Const VK_9 = &H39
Public Const VK_A = &H41
Public Const VK_B = &H42
Public Const VK_C = &H43
Public Const VK_D = &H44
Public Const VK_E = &H45
Public Const VK_F = &H46
Public Const VK_G = &H47
Public Const VK_H = &H48
Public Const VK_I = &H49
Public Const VK_J = &H4A
Public Const VK_K = &H4B
Public Const VK_L = &H4C
Public Const VK_M = &H4D
Public Const VK_N = &H4E
Public Const VK_O = &H4F
Public Const VK_P = &H50
Public Const VK_Q = &H51
Public Const VK_R = &H52
Public Const VK_S = &H53
Public Const VK_T = &H54
Public Const VK_U = &H55
Public Const VK_V = &H56
Public Const VK_W = &H57
Public Const VK_X = &H58
Public Const VK_Y = &H59
Public Const VK_Z = &H5A
Public Const VK_STARTKEY = &H5B
Public Const VK_CONTEXTKEY = &H5D
Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87
Public Const VK_NUMLOCK = &H90
Public Const VK_OEM_SCROLL = &H91
Public Const VK_OEM_1 = &HBA
Public Const VK_OEM_PLUS = &HBB
Public Const VK_OEM_COMMA = &HBC
Public Const VK_OEM_MINUS = &HBD
Public Const VK_OEM_PERIOD = &HBE
Public Const VK_OEM_2 = &HBF
Public Const VK_OEM_3 = &HC0
Public Const VK_OEM_4 = &HDB
Public Const VK_OEM_5 = &HDC
Public Const VK_OEM_6 = &HDD
Public Const VK_OEM_7 = &HDE
Public Const VK_OEM_8 = &HDF
Public Const VK_ICO_F17 = &HE0
Public Const VK_ICO_F18 = &HE1
Public Const VK_OEM102 = &HE2
Public Const VK_ICO_HELP = &HE3
Public Const VK_ICO_00 = &HE4
Public Const VK_ICO_CLEAR = &HE6
Public Const VK_OEM_RESET = &HE9
Public Const VK_OEM_JUMP = &HEA
Public Const VK_OEM_PA1 = &HEB
Public Const VK_OEM_PA2 = &HEC
Public Const VK_OEM_PA3 = &HED
Public Const VK_OEM_WSCTRL = &HEE
Public Const VK_OEM_CUSEL = &HEF
Public Const VK_OEM_ATTN = &HF0
Public Const VK_OEM_FINNISH = &HF1
Public Const VK_OEM_COPY = &HF2
Public Const VK_OEM_AUTO = &HF3
Public Const VK_OEM_ENLW = &HF4
Public Const VK_OEM_BACKTAB = &HF5
Public Const VK_ATTN = &HF6
Public Const VK_CRSEL = &HF7
Public Const VK_EXSEL = &HF8
Public Const VK_EREOF = &HF9
Public Const VK_PLAY = &HFA
Public Const VK_ZOOM = &HFB
Public Const VK_NONAME = &HFC
Public Const VK_PA1 = &HFD
Public Const VK_OEM_CLEAR = &HFE


Public Const SK_0 = &HB
Public Const SK_9 = &HA
Public Const SK_8 = &H9
Public Const SK_7 = &H8
Public Const SK_6 = &H7
Public Const SK_5 = &H6
Public Const SK_4 = &H5
Public Const SK_3 = &H4

'v0.9 def:
Public SK(8) As Long
Public VK(8) As Integer


Public Const KEYEVENTF_KEYUP = &H2
Public Const KEYEVENTF_EXTENDEDKEY = &H1

Public Const INPUT_KEYBOARD = 1



Public Type KEYBDINPUT
      wVk As Integer
      wScan As Integer
      dwFlags As Long
      time As Long
      dwExtraInfo As Long
End Type



Public Type INPUT_TYPE
      dwType As Long
      xi(0 To 23) As Byte
End Type



Public Timer2Count As Integer


Public Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbSize As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)





'**********************************************************







Public Function Pk(sKey As Long)
    Dim GInput As INPUT_TYPE
    Dim KInput As KEYBDINPUT
    Dim rtv As Long
    
    KInput.wVk = 0
    KInput.wScan = sKey
    KInput.dwFlags = 8
    KInput.time = 0
    KInput.dwExtraInfo = 0

    
    
   
    GInput.dwType = INPUT_KEYBOARD
    CopyMemory GInput.xi(0), KInput, Len(KInput)
    'MsgBox Len(KInput)
    
    
    rtv = SendInput(1, GInput, Len(GInput))
'    MsgBox rtv
'Form2.Label6.Caption = Str(sKey) + " (scan code) pressing"

If Form2.Option2.Value = True Then
    
    Sleep 50
    PKrelease sKey

        
End If


End Function


Public Function PKrelease(sKey As Long)
    Dim GInput As INPUT_TYPE
    Dim KInput As KEYBDINPUT
    Dim rtv As Long
    
        KInput.wVk = 0
        KInput.wScan = sKey
        KInput.dwFlags = 2
        KInput.time = 0
        KInput.dwExtraInfo = 0
    
        
       
        GInput.dwType = INPUT_KEYBOARD
        CopyMemory GInput.xi(0), KInput, Len(KInput)
        'MsgBox Len(KInput)
        
        
        rtv = SendInput(1, GInput, Len(GInput))
        'Form2.Label6.Caption = Str(sKey) + " (scan code) released"
           
End Function

'**********************************************************
















Public Sub SingleKybdEvent(bKey As Integer, EventType As Long)
    Dim GInput As INPUT_TYPE
    Dim KInput As KEYBDINPUT
    KInput.wVk = bKey  'the key we're going to press
    KInput.dwFlags = EventType 'just send whatever keybd event you wish...
    'copy the structure into the input array's buffer.
    GInput.dwType = INPUT_KEYBOARD   ' keyboard input
    CopyMemory GInput.xi(0), KInput, Len(KInput)
    
    'send the input now
    Call SendInput(1, GInput, Len(GInput))
End Sub






'Public Sub SingleKybdEvent(bKey As Integer, EventType As Long)
'    Dim GInput As INPUT_TYPE
'    Dim KInput As KEYBDINPUT
'
'    GInput.dwType = INPUT_KEYBOARD   ' keyboard input
'
'
'
'    KInput.wVk = 0  '"trick"
'    KInput.wScan = bKey  'the key we're going to press
'    KInput.dwFlags = EventType 'just send whatever keybd event you wish...
'
'
'    'copy the structure into the input array's buffer.
'    CopyMemory GInput.xi(0), KInput, Len(KInput)
'
'
'    'send the input now
'    Call SendInput(1, GInput, Len(GInput))
'
'End Sub


'now use like this:

'SingleKybdEvent bKey, 0   'key down
''do some stuff, wait, etc.
'
'SingleKybdEvent bKey, KEYEVENTEF_KEYUP  'key up


Public Sub SingleKybdEventPro(bKey As Integer)

SingleKybdEvent bKey, KEYEVENTF_EXTENDEDKEY

If Form2.Option2.Value = True Then    'press and release mode
    Sleep 50
    SingleKybdEvent bKey + &H80, KEYEVENTF_KEYUP 'key up
End If

'note:
'for press and holding actions,
'disable "sleep 50" and move "SingleKybdEvent bKey, &H2  'key up" to where Y-axis is centered.
'for press and release actions,
'enable the above two sentance and disable the related sentances where Y-axis is centered.



End Sub


''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''
Public Sub SingleKybdEventOld(bKey As Integer, EventType As Long)
    Dim GInput As INPUT_TYPE
    Dim KInput As KEYBDINPUT
    
    GInput.dwType = INPUT_KEYBOARD   ' keyboard input
    
    
    
    KInput.wVk = bKey  'the key we're going to press
    KInput.dwFlags = EventType 'just send whatever keybd event you wish...
    
    
    'copy the structure into the input array's buffer.
    CopyMemory GInput.xi(0), KInput, Len(KInput)
    
    
    'send the input now
    Call SendInput(1, GInput, Len(GInput))
    
End Sub


'now use like this:

'SingleKybdEvent bKey, 0   'key down
''do some stuff, wait, etc.
'
'SingleKybdEvent bKey, KEYEVENTEF_KEYUP  'key up


Public Sub SingleKybdEventProOld(bKey As Integer)

SingleKybdEventOld bKey, 0
If Form2.Option2.Value = True Then    'press and release mode
    Sleep 50
    SingleKybdEvent bKey, KEYEVENTF_KEYUP 'key up
End If

End Sub
