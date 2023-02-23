VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   Caption         =   "Form2"
   ClientHeight    =   6540
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8610
   LinkTopic       =   "Form2"
   ScaleHeight     =   6540
   ScaleWidth      =   8610
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   6480
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   495
      Left            =   7080
      TabIndex        =   44
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   7440
      TabIndex        =   1
      Text            =   "2"
      Top             =   240
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text13 
      Height          =   615
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   41
      Text            =   "Form2.frx":0000
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   495
      Left            =   7080
      TabIndex        =   40
      Top             =   4680
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resize Effect Zone/重设效果区"
      Height          =   2055
      Left            =   120
      TabIndex        =   25
      Top             =   4200
      Width           =   3015
      Begin VB.OptionButton OptionG 
         Caption         =   "R"
         Enabled         =   0   'False
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   45
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Redraw/重画"
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   2055
      End
      Begin VB.OptionButton OptionG 
         Caption         =   "0"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   32
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton OptionG 
         Caption         =   "1"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptionG 
         Caption         =   "2"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton OptionG 
         Caption         =   "3"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   29
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptionG 
         Caption         =   "4"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   28
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton OptionG 
         Caption         =   "5"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptionG 
         Caption         =   "6"
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   26
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load File/载入文件"
      Height          =   615
      Left            =   3240
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Modify Zones/调整区域"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Width           =   4815
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6600
      Top             =   4200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Default Data/默认数据"
      Height          =   615
      Left            =   4800
      TabIndex        =   20
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3276
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3276
      Begin VB.Label Labelx 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         ForeColor       =   &H0000FF00&
         Height          =   200
         Left            =   1440
         TabIndex        =   22
         Top             =   840
         Width           =   200
      End
      Begin VB.Label Label_G 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "R"
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   7
         Left            =   2520
         TabIndex        =   39
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label Label_G 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         Caption         =   "6"
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   6
         Left            =   2280
         TabIndex        =   19
         Top             =   2160
         Width           =   315
      End
      Begin VB.Label Label_G 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         Caption         =   "5"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   2760
         TabIndex        =   18
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label_G 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "4"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   1560
         TabIndex        =   17
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label_G 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   195
      End
      Begin VB.Label Label_G 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   15
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label_G 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   14
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label_G 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "C/NZ"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   1440
         Width           =   2835
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Mode/模式"
      Height          =   2775
      Left            =   3600
      TabIndex        =   35
      Top             =   840
      Width           =   4815
      Begin VB.CheckBox Check5 
         Caption         =   "R Gear/倒档"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "<-- ?"
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "<-- ?"
         Height          =   255
         Left            =   3960
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check4 
         Caption         =   "N Gear/空档"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "<-- ?"
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   2280
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Normal Joystick? ( Auto Center )"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   4095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<-- ?"
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "LFS"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Press and hold"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Press then release"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "普通摇杆？（自动回中）"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "按后释放"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "按住不放"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Label Label9 
      Caption         =   "DEMO beta1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3360
      TabIndex        =   46
      Top             =   4800
      Width           =   4935
   End
   Begin VB.Label Label8 
      Caption         =   "Device ID:"
      Height          =   255
      Left            =   6240
      TabIndex        =   43
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "设备 ID："
      Height          =   255
      Left            =   6240
      TabIndex        =   42
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Label3"
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   3240
      TabIndex        =   23
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   3600
      TabIndex        =   21
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu MenuOptFile 
      Caption         =   "File&&Option/文件与设置"
      Begin VB.Menu MenuFile 
         Caption         =   "Profile/配置文件"
         Begin VB.Menu MenuOpen 
            Caption         =   "Open/打开..."
         End
         Begin VB.Menu MenuSave 
            Caption         =   "Save/保存..."
         End
      End
      Begin VB.Menu MenuKey 
         Caption         =   "Key/按键"
         Begin VB.Menu MenuOpenKey 
            Caption         =   "Open/打开..."
         End
         Begin VB.Menu MenuDash1 
            Caption         =   "-"
         End
         Begin VB.Menu MenuReload 
            Caption         =   "Reload/重载 key.cfg"
         End
      End
      Begin VB.Menu MenuDash 
         Caption         =   "-"
      End
      Begin VB.Menu MenuReset 
         Caption         =   "Reset all/重置所有选项"
      End
   End
   Begin VB.Menu MenuCreateFile 
      Caption         =   "Generate File/生成文件"
      Begin VB.Menu MenuCrtDefault 
         Caption         =   "default.cfg"
      End
      Begin VB.Menu MenuCrtKey 
         Caption         =   "key.cfg"
      End
      Begin VB.Menu MenuCrtHelp 
         Caption         =   "help.html"
      End
   End
   Begin VB.Menu MenuUt 
      Caption         =   "Utlities/工具"
      Begin VB.Menu MenuNotepad 
         Caption         =   "Notepad/记事本"
      End
      Begin VB.Menu MenuAppFolder 
         Caption         =   "AppFolder/程序目录"
      End
      Begin VB.Menu MenuVClist 
         Caption         =   "Virtual Code List"
      End
      Begin VB.Menu MenuSClist 
         Caption         =   "Scan Code List"
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "Help/帮助"
      Begin VB.Menu Menu111 
         Caption         =   "111"
      End
      Begin VB.Menu Menu222 
         Caption         =   "222"
      End
      Begin VB.Menu MenuMore 
         Caption         =   "....(more)"
      End
      Begin VB.Menu MenuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu MenuAboutCN 
         Caption         =   "关于"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit








Private Sub Check2_Click()
 If Check2.Value = 1 Then
    Option1.Value = 1
    Check4.Value = 0
 End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Check4.Value = 0
        Check4.Enabled = False
        

    Else
        Check4.Enabled = True

    End If
End Sub



Private Sub Check5_Click()
    drawZone
End Sub

Private Sub Form_Load()
Dim i As Integer

'====COPYED FREOM OLD form2=====================================

vers = "Joystick To H-shifter 2 DEMO beta1"
Form2.Caption = vers


'-----------------------------------------------------------

strDemo = "Sorry, not available in DEMO version." + Chr(13) + "抱歉，演示版暂不支持。"


Label2.Caption = ""
Label3.Caption = ""

CommonDialog1.InitDir = App.Path


justPressed = False
justCentered = True

joyBut1 = False
joyBut2 = False






resetAll

'load default.cfg
If loadFile("profile", True) = False Then
   
   If MsgBox("default.cfg not found, do you want to generate one? (program can be run without default.cfg)" + Chr(13) + Chr(13) + "未找到default.cfg，是否现在生成？（无default.cfg程序也可运行）", vbYesNo + vbDefaultButton1) = vbYes Then
    'code here
     MenuCrtDefault_Click
   End If
Else
End If


'load key.cfg
If loadFile("keyfile", True) = False Then
   
    MsgBox "key.cfg not found, do you want to generate one?"
    'code here
    MenuCrtKey_Click
Else
End If



End Sub













Private Sub Check1_Click()
    Dim i As Integer
    If Check1.Value = 1 Then
        Form2.Height = 7200
        
        For i = 0 To 7
            OptionG(i).Enabled = True
            OptionG(i).Value = False
        Next
        Label3.Caption = "Please select a gear..." + Chr(13) + "请选择一个档位……"
        
        'ft
        MsgBox "Press and hold button 1 to drag GEZ"
    Else
        Form2.Height = 5000
        
        For i = 0 To 7
            OptionG(i).Enabled = False
        Next
        Label3.Caption = "Modifying disabled." + Chr(13) + "调整已关闭"
    End If
End Sub

Private Sub Command1_Click()
    If MsgBox("All data will be reset, are you sure?" + Chr(13) + "重设所有数据，确定？", vbOKCancel + vbDefaultButton2) = vbOK Then
        resetZone
        drawZone
    End If

End Sub



Private Sub Command2_Click()
    If MsgBox("Current data will lost, are you sure?" + Chr(13) + "当前数据将要丢失，确定？", vbOKCancel + vbDefaultButton2) = vbOK Then
'        resetZone
'        drawZone
    End If
End Sub

Private Sub Command3_Click()
    Dim id As Integer
    Dim xc As Long
    Dim yc As Long
    
    id = checkOptionG
    
    xc = 0.5 * (gearZone(id).xmin + gearZone(id).xmax)
    yc = 0.5 * (gearZone(id).ymin + gearZone(id).ymax)
    
    gearZone(id).xmin = xc - 25
    gearZone(id).xmax = xc + 25
    gearZone(id).ymin = yc - 25
    gearZone(id).ymax = yc + 25
    
    drawZone
    
End Sub

Private Sub Command4_Click()
MsgBox _
"LFS receives a kind of input function other than normal DirectX games, so if you're going to play LFS, check this option." + Chr(13) + _
"由于LFS所接收的输入函数和普通DirectX游戏不同，如果准备玩LFS，勾上此选项。"

End Sub

Private Sub Command5_Click()
MsgBox _
"If you're using a normal joystick which will returen to center after your hand releases, check this option." + Chr(13) + _
"如果用的是普通摇杆，放手后会自动回中，构上这个选项"

End Sub

Private Sub Command6_Click()
MsgBox _
"If you want to send a virtual key pressing while stick is in natural gear, check this option. (for further info, see menu help)" + Chr(13) + _
"欲在空档时输入虚拟按键，勾上此选项。（详情请看菜单帮助）"

End Sub

Private Sub Command7_Click()
MsgBox _
"If you want to use reverse gear, check this option. (for further info, see menu help)" + Chr(13) + _
"若使用倒档，勾上此选项。（详情请看菜单帮助）"

End Sub

Private Sub Command8_Click()
MsgBox Left(JOYINFO.szPname, InStr(JOYINFO.szPname, vbNullChar) - 1)
'saveFile
End Sub

Private Sub Command9_Click()
loadFile "keyfile"

End Sub


  
  
  
  








Private Sub Menu111_Click()
MsgBox strDemo
End Sub

Private Sub Menu222_Click()
MsgBox strDemo
End Sub

Private Sub MenuAbout_Click()
MsgBox vers + Chr(13) + Chr(13) + "by Kane Xu, shikai.xu@hotmail.com" + Chr(13) + "2008.11.25"
End Sub

Private Sub MenuAboutCN_Click()
MsgBox strDemo
End Sub

Private Sub MenuAppFolder_Click()
MsgBox strDemo
End Sub

Private Sub MenuCrtDefault_Click()
genDefaultProfile
End Sub

Private Sub MenuCrtHelp_Click()
MsgBox strDemo
End Sub

Private Sub MenuCrtKey_Click()
genDefaultKeyfile
End Sub

Private Sub MenuMore_Click()
MsgBox strDemo
End Sub

Private Sub MenuNotepad_Click()
MsgBox strDemo
End Sub

Private Sub MenuOpen_Click()
    openFile "profile"
    drawZone

End Sub

Private Sub MenuOpenKey_Click()
  openFile "keyfile"
        

End Sub

Private Sub MenuReload_Click()
    loadFile "keyfile"

End Sub

Private Sub MenuReset_Click()
If MsgBox("All interface option and keys will be reset to default, are you sure?" + Chr(13) + _
          "(default.cfg and key.cfg won't be modified) " + Chr(13) + Chr(13) + _
          "所有界面选项和按键都将被重置，是否确定？" + Chr(13) + _
          "（此操作不会影响default.cfg 和 key.cfg）", vbYesNo + vbDefaultButton2) = vbYes Then
          
          resetAll
          
          Else
          End If
          
          
End Sub

Private Sub MenuSave_Click()
saveFile
End Sub

Private Sub MenuSClist_Click()
MsgBox strDemo
End Sub

Private Sub MenuVClist_Click()
MsgBox strDemo
End Sub

Private Sub OptionG_Click(Index As Integer)
MsgBox "Press and hold joystick button 2 to draw effect zone of gear " + Str(Index) + " ..."
Label3.Caption = "Resize Gear" + Str(Index) + " zone." + Chr(13) + Chr(10) + "重定义" + Str(Index) + "档区域。"
End Sub

Private Sub Text1_Change()
    joyID = Val(Text1.Text)
    Text1.Text = joyID
    
     keybd_event VK_TAB, 0, 1, 0    'jump to next tab to prevent mis-change ID
     keybd_event VK_TAB, 0, 3, 0    '3表示松开
End Sub



Private Sub Timer1_Timer()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim RZid As Integer 'tag for resize zones, value get by checkOptionG function



'get joystick position data:
joyGetPos joyID, joyPos




'calculate cursor pos:
With joyCursor
    .wXpos = joyPos.wXpos / 20 - 100
    .wYpos = joyPos.wYpos / 20 - 100
End With

'print position data:
Label1.Caption = "X: " + Str(joyPos.wXpos) + Chr(13) + "Y: " + Str(joyPos.wYpos)

'drawing cursor:
Labelx.Left = joyCursor.wXpos
Labelx.Top = joyCursor.wYpos

'showing button states:
If (joyPos.wButtons And JOY_BUTTON1) = JOY_BUTTON1 Then Label1.Caption = Label1.Caption + Chr(13) + "Button 1 / 按钮1"
If (joyPos.wButtons And JOY_BUTTON2) = JOY_BUTTON2 Then Label1.Caption = Label1.Caption + Chr(13) + "Button 2 / 按钮2"





If Check1.Value = 0 Then
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'v0.2 pressings:

'N gear:
If ifWithinZone(0, joyPos) = True Then     'cursor in N zone

    If justPressed = False And Check4.Value = 1 Then Pk SK_0    'user checked, then send N gear key pressing
    justCentered = True
    justPressed = True
    
    'NOET: althought in CenterJoystick mode,
    'detecting if the stick has been centered is also necessary
    'so leave the justCentered=True sentence still effective.
    If Check3.Value = 0 Then 'centerJoystick is inactive
    ':::::::::::::::::::::::::::::::::
    '  release keys for holding mode
    ':::::::::::::::::::::::::::::::::
    
    ':::::copyed from v0.2 -- dx game
     If Option1.Value = True And Check2.Value = 0 Then  'release key for "press and holding" mode
        For j = 1 To 7
            PKrelease SK(j)
        Next
 
'        PKrelease SK_9
'        PKrelease SK_8
'        PKrelease SK_7
'        PKrelease SK_6
'        PKrelease SK_5
'        PKrelease SK_4
    
     End If
     
     
     
     'copyed from v0.1 -- LFS S2
      If Option1.Value = True And Check2.Value = 1 Then  'release key for "press and holding" mode
        For j = 1 To 7
            SingleKybdEvent VK(j), &H2
        Next
    '    SingleKybdEvent VK_9, &H2  'key up
    '    SingleKybdEvent VK_8, &H2  'key up
    '    SingleKybdEvent VK_7, &H2  'key up
    '    SingleKybdEvent VK_6, &H2  'key up
    '    SingleKybdEvent VK_5, &H2  'key up
    '    SingleKybdEvent VK_4, &H2  'key up
        End If
     
     
     
 ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
 ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
 ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    End If  'end If check3.value=0

 End If
'end of N gear.
'(end of release keys)


For j = 1 To 7
    If ifWithinZone(j, joyPos) = True Then
    
        'for centerJoystick mode with HOLDING mode, release all keys before push a new key:
        If Check3.Value = 1 And Option1.Value = True And justCentered = True Then
   
         If Check2.Value = 0 Then     'copyed from v0.2 -- dx game
            For k = 1 To 7
                PKrelease SK(k)
            Next
         End If
         If Check2.Value = 1 Then     'copyed from v0.1 -- LFS S2
            For k = 1 To 7
                SingleKybdEvent VK(k), &H2
            Next
         End If
          
        End If
        'end of centerJoystick with HOLDING mode, releasing keys
    
    
        If justPressed = False And justCentered = True Then
            If Check2.Value = 0 Then Pk SK(j)                   'normal dx game
            If Check2.Value = 1 Then SingleKybdEventPro VK(j)   'LFS S2
        End If
        
        justCentered = False
        justPressed = True
        
        'FOR TESTING true "press then release" mode:
        If Option2.Value = True Then
            If Check2.Value = 0 Then PKrelease SK(j)        'normal dx game
            If Check2.Value = 1 Then SingleKybdEvent VK(j), &H2   'LFS S2
        End If
        
    End If
Next



If ifWithoutZone(joyPos) = True Then justPressed = False



'....................................
'....................................
End If  'end of: If check1.value = 0
'end of normal operations.



'=============================
'=============================
'adjust zones:

If Check1.Value = 1 Then

'=============================
'zone operations:
For i = 0 To 7

'     If joyPos.wXpos >= gearZone(i).xmin And joyPos.wXpos <= gearZone(i).xmax _
'     And joyPos.wYpos >= gearZone(i).ymin And joyPos.wYpos <= gearZone(i).ymax Then
     
     If ifWithinZone(i, joyPos) = True Then
             
        Label2.Caption = "Cursor is in zone " + Str(i) + Chr(13) + "光标落在" + Str(i) + "档区内"
       
        
        '================================
        'pick effect zones:
        If (joyPos.wButtons And JOY_BUTTON1) <> JOY_BUTTON1 Then
            temp1 = 0.5 * (gearZone(i).xmax - gearZone(i).xmin)
            temp2 = 0.5 * (gearZone(i).ymax - gearZone(i).ymin)
        Else
            With gearZone(i)
                .xmax = joyPos.wXpos + temp1
                .xmin = joyPos.wXpos - temp1
                .ymax = joyPos.wYpos + temp2
                .ymin = joyPos.wYpos - temp2
            End With

        End If
               
        '--------------------------------
        
        
         '===========================
'        'resize zones:
'        If (joyPos.wButtons And JOY_BUTTON2) = JOY_BUTTON2 Then
'            GoTo RESIZEZONES
'        End If
        '----------------------------
        
        
        drawZone
        
        
    
    Exit For
    Else
        Label2.Caption = "Cursor is out of any zone." + Chr(13) + "光标不在效果区内"
    End If



Next
'end of zone operations
'--------------------------------



         '===========================
        'resize zones:
        If (joyPos.wButtons And JOY_BUTTON2) = JOY_BUTTON2 Then
            RZid = checkOptionG
            
            With gearZone(RZid)
                If joyPos.wXpos < .xmin Then .xmin = joyPos.wXpos
                If joyPos.wXpos > .xmax Then .xmax = joyPos.wXpos
                If joyPos.wYpos < .ymin Then .ymin = joyPos.wYpos
                If joyPos.wYpos > .ymax Then .ymax = joyPos.wYpos
                
            End With

            
        
        End If
        '--------------------------------
        
        
End If  'end of: If check1.value = 1
'end of adjust zones:
'-----------------------------
'-----------------------------


End Sub






















