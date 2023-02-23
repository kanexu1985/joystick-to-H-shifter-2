Attribute VB_Name = "v9module"
Option Explicit

Public Type zone
    xmax As Long
    xmin As Long
    ymax As Long
    ymin As Long
End Type

Public gearZone(8) As zone


Public Function resetAll()
'reset keys:
SK(0) = SK_0
SK(1) = SK_9
SK(2) = SK_8
SK(3) = SK_7
SK(4) = SK_6
SK(5) = SK_5
SK(6) = SK_4
SK(7) = SK_3

VK(0) = VK_0
VK(1) = VK_9
VK(2) = VK_8
VK(3) = VK_7
VK(4) = VK_6
VK(5) = VK_5
VK(6) = VK_4
VK(7) = VK_3

'reset interface:

Form2.Text1.Text = "0"
joyID = Val(Form2.Text1.Text)

Form2.Option1.Value = True

Form2.Check3.Value = 0
Form2.Check2.Value = 0
Form2.Check4.Value = 0
Form2.Check5.Value = 0

'reset Zones
resetZone

drawZone



End Function


Public Function resetZone()
'''''''''''''''''''''''''''''''''''''''''''''''
'Y: MAX                     MIN
' |dz/dzl|....|center|....|dz/dzl|

'''''''''''''''''''''''''''''''''''''''''''''''
'set value for hit point range:
dz = 5000

'set value for netural position range:
dzl = 30000

    '......................
    gearZone(0).xmin = 0
    gearZone(0).xmax = 65535
    
    gearZone(0).ymin = dzl
    gearZone(0).ymax = 65535 - dzl
    '......................
    
    '......................
    gearZone(1).xmin = 0
    gearZone(1).xmax = dz
    
    gearZone(1).ymin = 0
    gearZone(1).ymax = dz
    '......................
    
    '......................
    gearZone(2).xmin = 0
    gearZone(2).xmax = dz
    
    gearZone(2).ymin = 65535 - dz
    gearZone(2).ymax = 65535
    '......................
    
    '......................
    gearZone(3).xmin = 2 * dz
    gearZone(3).xmax = 65535 - 2 * dz
    
    gearZone(3).ymin = 0
    gearZone(3).ymax = dz
    '......................
    
    '......................
    gearZone(4).xmin = 2 * dz
    gearZone(4).xmax = 65535 - 2 * dz
    
    gearZone(4).ymin = 65535 - dz
    gearZone(4).ymax = 65535
    '......................
    
    '......................
    gearZone(5).xmin = 65535 - dz
    gearZone(5).xmax = 65535
    
    gearZone(5).ymin = 0
    gearZone(5).ymax = dz
    '......................
    
    '......................
    gearZone(6).xmin = 65535 - dz
    gearZone(6).xmax = 65535
    
    gearZone(6).ymin = 65535 - dz
    gearZone(6).ymax = 65535
    '......................
    
    '......................
    gearZone(7).xmin = 65535 - 3 * dz
    gearZone(7).xmax = 65535 - 2 * dz
    
    gearZone(7).ymin = 65535 - 2 * dz
    gearZone(7).ymax = 65535 - dz
    '......................


End Function



Public Function drawZone()
    Dim i As Integer

'Form2.Label_G(1).Left = 0
'Form2.Label_G(1).Top = 0
'Form2.Label_G(1).Width = 2000
'Form2.Label_G(1).Height = 20


detectRGear

    For i = 0 To 7
        If gearZone(i).xmin < 0 Then
            Form2.Label_G(i).Visible = False
            Exit For 'don't do the draw if R gear is frozen
        Else
            Form2.Label_G(i).Visible = True 'show GEZ if R is active
        End If
        Form2.Label_G(i).Left = gearZone(i).xmin / 20
        Form2.Label_G(i).Top = gearZone(i).ymin / 20
        Form2.Label_G(i).Width = (gearZone(i).xmax - gearZone(i).xmin) / 20
        Form2.Label_G(i).Height = (gearZone(i).ymax - gearZone(i).ymin) / 20
        
        'MsgBox str(gearZone(i).xmax) + " " + str(gearZone(i).xmin) + " " + str(gearZone(i).ymax) + " " + str(gearZone(i).ymin)
'        MsgBox str(Form2.Label_G(i).Left) + " " + str(Form2.Label_G(i).Top) + " " + str(Form2.Label_G(i).Width) + " " + str(Form2.Label_G(i).Height)
    Next

'ft

'Form2.Label2.Caption = "x:" + Str(Form2.Label_G(7).Left) + " w:" + Str(Form2.Label_G(7).Width) + " y:" + Str(Form2.Label_G(7).Top) + " h:" + Str(Form2.Label_G(7).Height)
'
'Form2.Label2.Caption = "x:" + Str(gearZone(7).xmin) + " w:" + Str(gearZone(7).xmax) + " y:" + Str(gearZone(7).ymin) + " h:" + Str(gearZone(7).ymax)

End Function


Public Function detectRGear()
Dim errMsg As String
errMsg = "R GEZ is invalid, please check the last loaded .cfg." + Chr(13) + Chr(13) + "R档GEZ数据有误，请检查最后次载入的.cfg文件。"
            

    If Form2.Check5.Value = 1 Then
    
        Form2.OptionG(7).Visible = True
        
        If gearZone(7).xmin >= 0 And gearZone(7).xmax >= 0 And gearZone(7).ymin >= 0 And gearZone(7).ymax >= 0 Then
            'N GEZ already shown, do nothing
            '(which is also not possible)
        ElseIf gearZone(7).xmin < 0 And gearZone(7).xmax < 0 And gearZone(7).ymin < 0 And gearZone(7).ymax < 0 Then
            'restore frozen R GEZ to screen:
            With gearZone(7)
                .xmax = (-1) - .xmax
                .xmin = (-1) - .xmin
                .ymax = (-1) - .ymax
                .ymin = (-1) - .ymin
            End With
        Else
            'R GEZ cross zero value
            MsgBox errMsg
        End If
        
        
        
    ElseIf Form2.Check5.Value = 0 Then
    
    
        Form2.OptionG(7).Visible = False
    
        If gearZone(7).xmin >= 0 And gearZone(7).xmax >= 0 And gearZone(7).ymin >= 0 And gearZone(7).ymax >= 0 Then
            'move R GEZ to frozen area.
            With gearZone(7)
                .xmax = (-1) - .xmax
                .xmin = (-1) - .xmin
                .ymax = (-1) - .ymax
                .ymin = (-1) - .ymin
            End With
        ElseIf gearZone(7).xmin < 0 And gearZone(7).xmax < 0 And gearZone(7).ymin < 0 And gearZone(7).ymax < 0 Then
            'R GEZ already frozen, do nothing
            '(which is also not possible)

        Else
            'N GEZ cross zero value
            MsgBox errMsg
        End If
    
    End If
    
    
    
End Function


Public Function checkOptionG() As Integer
    Dim i As Integer
     For i = 0 To 7
        If Form2.OptionG(i).Value = True Then
            checkOptionG = i
            Exit For
        End If
     Next
End Function



Public Function ifWithinZone(gearID As Integer, stick As JOYINFO) As Boolean


     If stick.wXpos >= gearZone(gearID).xmin And stick.wXpos <= gearZone(gearID).xmax _
     And stick.wYpos >= gearZone(gearID).ymin And stick.wYpos <= gearZone(gearID).ymax Then
     
        ifWithinZone = True
        
     Else
     
        ifWithinZone = False
        
     End If
     
        
End Function






Public Function ifWithoutZone(stick As JOYINFO) As Boolean
    Dim gearID As Integer
    Dim temp As Integer
    
    temp = 0
    
    For gearID = 0 To 7

         If stick.wXpos >= gearZone(gearID).xmin And stick.wXpos <= gearZone(gearID).xmax _
         And stick.wYpos >= gearZone(gearID).ymin And stick.wYpos <= gearZone(gearID).ymax Then _

         temp = temp + 1
         
         End If

    Next

    If temp = 0 Then
        ifWithoutZone = True
    Else
        ifWithoutZone = False
    End If

End Function










