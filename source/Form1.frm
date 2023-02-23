VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6420
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "?"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mode/模式"
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "Press then release"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Press and hold"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "按住不放"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "按后释放"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   2760
      Max             =   6554
      TabIndex        =   5
      Top             =   1680
      Width           =   3135
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2895
      Left            =   4200
      Max             =   6554
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1440
      TabIndex        =   3
      Text            =   "2"
      Top             =   360
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   360
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Yctl"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Ycentered"
      Enabled         =   0   'False
      Height          =   2415
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Text            =   "Test Area"
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   855
      Left            =   1320
      TabIndex        =   15
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "设备 ID："
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Device ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim retval As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub Form_Load()
Form2.Show

'set value for hit point range:
dz = 5000

'set value for netural position range:
dzl = 30000



'''''''''''''''''''''''''''''''''''''''''''''''
'Y: MAX                     MIN
' |dz/dzl|....|center|....|dz/dzl|

'''''''''''''''''''''''''''''''''''''''''''''''

vers = "Joystick to H-shifter 0.2"
Form1.Caption = vers

joyID = Val(Text1.Text)


End Sub

Private Sub Command1_Click()
Dim str As String

Dim str10 As String
Dim str11 As String
Dim str20 As String
Dim str21 As String
Dim str30 As String
Dim str31 As String
Dim str40 As String
Dim str41 As String

Dim str90 As String
Dim str91 As String
Dim str92 As String
Dim str100 As String
Dim str101 As String

str10 = "Welcome to " + vers
str11 = "欢迎使用" + vers

str20 = "Choose your joystick by changing the 'Device ID' "
str21 = "改变“设备ID”选择所用的摇杆"

str30 = "Choose the Press&Holding or Press/Release mode. ( for LFS, use Press&Holding mode )"
str31 = "选择 按住不放 或 按后释放 模式，LFS只识别 按住不放"

str40 = "This program is still under construction, any suggestions are welcomed :-) "
str41 = "此程序尚未完成，欢迎提出宝贵意见"

str90 = "Kane Xu"
str91 = "徐时开"
str100 = "shikai.xu@hotmail.com" + Chr(13) + "http://ssnakemsn.spaces.live.com"
str101 = "Special thanks to Jonathan,the creator, designer and programmer of Xpadder," + Chr(13) + "to help me with my program!!!"

str92 = "2008.4.28"


str = str10 + Chr(13) + str11 + Chr(13) + Chr(13) _
     + str20 + Chr(13) + str21 + Chr(13) + Chr(13) _
     + str30 + Chr(13) + str31 + Chr(13) + Chr(13) _
     + str40 + Chr(13) + str41 + Chr(13) + Chr(13) _
                + Chr(13) + Chr(13) _
     + str90 + Chr(13) + str91 + Chr(13) + str92 + Chr(13) + Chr(13) _
     + str100 + Chr(13) + Chr(13) + str101
     
MsgBox str



End Sub







'########################################################


'########################################################





Private Sub Text1_Change()

joyID = Val(Text1.Text)
Text1.Text = joyID

End Sub




Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Timer1_Timer()

'get joystick position data:
joyGetPos joyID, joyPos

'print position data:
Label1.Caption = "X: " + str(joyPos.wXpos) + Chr(13) + "Y: " + str(joyPos.wYpos)

VScroll1.Value = joyPos.wYpos / 10
HScroll1.Value = joyPos.wXpos / 10


'set center
If joyPos.wYpos > 0 + dz And joyPos.wYpos < 65535 - dz Then
 center = True
End If

'while Y-axis return to center:
If joyPos.wYpos > 0 + dzl And joyPos.wYpos < 65535 - dzl Then
 centerl = True
 
 If Option1.Value = True Then   'press and holding mode
 
 PKrelease SK_9
 PKrelease SK_8
 PKrelease SK_7
 PKrelease SK_6
 PKrelease SK_5
 PKrelease SK_4

 
'    SingleKybdEvent VK_9 + &H8, KEYEVENTF_KEYUP 'key up
'    SingleKybdEvent VK_8 + &H8, KEYEVENTF_KEYUP 'key up
'    SingleKybdEvent VK_7, KEYEVENTF_KEYUP  'key up
'    SingleKybdEvent VK_6, KEYEVENTF_KEYUP  'key up
'    SingleKybdEvent VK_5, KEYEVENTF_KEYUP  'key up
'    SingleKybdEvent VK_4, KEYEVENTF_KEYUP  'key up
 End If
 
 
End If



'ft
If center = True Then Check1.Value = 1
If center = False Then Check1.Value = 0

If centerl = True Then Check2.Value = 1
If centerl = False Then Check2.Value = 0

'ftftftftftftftftftftftftftftftftftftftftftftftftftftftftftftftftftftftft

'ftftftftftftftftftftftftftftftftftftftftftftftftftftftftftftftftftftftft


'H shifter

If joyPos.wXpos < dz Then
    '1
    If centerl = True And center = True And joyPos.wYpos < dz Then
        Pk SK_9
        center = False
        centerl = False
    End If
    '2
    If centerl = True And center = True And joyPos.wYpos > 65535 - dz Then
        Pk SK_8
        center = False
        centerl = False
    End If
    
ElseIf joyPos.wXpos > 2 * dz And joyPos.wXpos < 65535 - 2 * dz Then
    '3
    If centerl = True And center = True And joyPos.wYpos < dz Then
        Pk SK_7
        center = False
        centerl = False
    End If
    
    '4
    If centerl = True And center = True And joyPos.wYpos > 65535 - dz Then
        Pk SK_6
        center = False
        centerl = False
    End If
    
ElseIf joyPos.wXpos > 65535 - dz Then
    '5
    If centerl = True And center = True And joyPos.wYpos < dz Then
        Pk SK_5
        center = False
        centerl = False
    End If
    
    '6
    If centerl = True And center = True And joyPos.wYpos > 65535 - dz Then
        Pk SK_4
        center = False
        centerl = False
    End If



End If







End Sub



'Private Sub Timer2_Timer()
'Timer2Count = Timer2Count + 1
'If Timer2Count = 50 Then
'    PKrelease SK_9
'    PKrelease SK_8
'    PKrelease SK_7
'    PKrelease SK_6
'    PKrelease SK_5
'    PKrelease SK_4
'    Timer2.Enabled = False
'End If
'End Sub


Private Sub Timer2_Timer()

End Sub
