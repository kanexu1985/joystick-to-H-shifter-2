Attribute VB_Name = "file"
Option Explicit

Public Function saveFile()
Const ForReading = 1, ForWriting = 2
Dim fso, f
Dim filepath As String
Dim i As Integer

Dim strgez(8) As String
      
      
  
  Form2.CommonDialog2.Flags = Form2.CommonDialog2.Flags Or cdlOFNOverwritePrompt
  
  Form2.CommonDialog2.FileName = ""
  
  '下面语句设置文件过滤方式,可显示扩展名为txt文件
  Form2.CommonDialog2.Filter = "(*.cfg)|*.cfg|(*.*)|*.*"
  
  '建立打开方式的通用对话框,也可使用commondialog1.showopen
  Form2.CommonDialog2.ShowSave
  
  filepath = Form2.CommonDialog2.FileName '得到选择的文件
  
  If filepath = "" Then Exit Function   'CANCEL clicked by user
  
  
   Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.OpenTextFile(filepath, ForWriting, True)
        '************************************
        'start:
        

'f.writeline Replace(("Y2MIN=" + Str(gearZone(2).ymin)), " ", "")



'Instruciont:
f.writeline "//******************************************************************"
f.writeline "//（中文说明请往下滚动）"
f.writeline "//This file is saved by Joystick To H-shifter 2"
f.writeline "//"
f.writeline "//You can change the options and values here if you want."
f.writeline "//"
f.writeline "//"
f.writeline "//Feel free to contact me for help:"
f.writeline "//Kane Xu, shikai.xu@hotmail.com"
f.writeline "//*******************************************************************"

f.writeline ""



'Options:
f.writeline Replace(("DEVID=" + Str(joyID) + Chr(9) + Chr(9) + "//Device ID of Your Joystick In Windows"), " ", "")

f.writeline Replace(("PRSHD=" + Str(Val(Form2.Option1.Value)) + Chr(9) + Chr(9) + "//1=Press And Hold," + Chr(9) + Chr(9) + "0=Press Than Release"), " ", "")

f.writeline Replace(("AUTOC=" + Str(Form2.Check3.Value) + Chr(9) + Chr(9) + "//1=Auto Center Joystick," + Chr(9) + Chr(9) + "0=non-Auto-Center Joystick"), " ", "")

f.writeline Replace(("LFSVK=" + Str(Form2.Check2.Value) + Chr(9) + Chr(9) + "//1=LFS, " + Chr(9) + Chr(9) + Chr(9) + "0=Other DirectX Games"), " ", "")

f.writeline Replace(("NGEAR=" + Str(Form2.Check4.Value) + Chr(9) + Chr(9) + "//1=Press Virtual Key While Joystick Is In Natural Gear," + Chr(9) + "0=Ignore That"), " ", "")

f.writeline Replace(("RGEAR=" + Str(Form2.Check5.Value) + Chr(9) + Chr(9) + "//1=Reverse Gear Active," + Chr(9) + Chr(9) + "0=No Reverse Gear"), " ", "")

f.writeline ""

f.writeline "//0" + Chr(9) + "= N gear"
f.writeline "//1-6" + Chr(9) + "= 1st-6th gear"
f.writeline "//7" + Chr(9) + "= R or 7th gear"

'GEZs:
For i = 0 To 7
strgez(i) = "X" + Str(i) + "MIN=" + Str(gearZone(i).xmin) + Chr(13) + Chr(10) + _
            "X" + Str(i) + "MAX=" + Str(gearZone(i).xmax) + Chr(13) + Chr(10) + _
            "Y" + Str(i) + "MIN=" + Str(gearZone(i).ymin) + Chr(13) + Chr(10) + _
            "Y" + Str(i) + "MAX=" + Str(gearZone(i).ymax) + Chr(13)


f.writeline Replace(strgez(i), " ", "")


Next

       
       
       
       
'end:
f.writeline ""
f.writeline "//******************************************************************"
f.writeline "//本文件由 摇杆变卅档2 生成"
f.writeline "//"
f.writeline "//您可以在这里修改选项和参数"
f.writeline "//"
f.writeline "//若有任何问题或建议，欢迎与我联系"
f.writeline "//徐时开 shikai.xu@hotmail.com"
f.writeline "//*******************************************************************"
f.writeline "//*******************************************************************"
f.writeline "//This is the end of file. 文件到此结束"
f.writeline ""
f.writeline ""
       
       
       
       
       
       
        
        'end
        '-------------------------------------
        f.Close
        
        If FileLen(filepath) = 1506 Then MsgBox "default.cfg successfully created." + Chr(13) + Chr(13) + "default.cfg已生成。"


  
  
      
End Function


Public Function openFile(cfg As String)
  Dim filepath As String
  Dim strLine As String
  
  
  '在未选择文件时,text.text为空字符,,不显示任何内容
  'Text1.Text = ""
  
  Form2.CommonDialog1.FileName = ""
  
  '下面语句设置文件过滤方式,可显示扩展名为txt文件
  Form2.CommonDialog1.Filter = "(*.cfg)|*.cfg|(*.*)|*.*"
  
  '建立打开方式的通用对话框,也可使用commondialog1.showopen
  Form2.CommonDialog1.Action = 1
  
  filepath = Form2.CommonDialog1.FileName '得到选择的文件
  
  If filepath = "" Then Exit Function   'CANCEL clicked by user
  
  
  Open filepath For Input As #1 '打开选择的文件
  
  
  
  
  Do Until EOF(1) '显示打开的文件
    Line Input #1, strLine
    
    
    'write code here
    'to analyze each line read from the file
    '===========================================
    
    If cfg = "profile" Then profileAnalyzer (strLine)
    If cfg = "keyfile" Then keyfileAnalyzer (strLine)
    
    '-------------------------------------------
    
    
    
'    MsgBox strLine
    'Form2.Text1.Text = Form2.Text1.Text + strLine + Chr(13) + Chr(10)
  Loop
  
  Close #1 '关闭打开的文件
  
  If cfg = "profile" Then procfgReply
  If cfg = "keyfile" Then keycfgReply




End Function




Public Function loadFile(cfg As String, Optional startupload As Boolean = False) As Boolean
  Dim filepath As String
  Dim strLine As String
  
  Dim fileexists As Boolean
  
  'asume file doesn't exit:
  fileexists = False
  
  If cfg = "profile" Then filepath = App.Path + "\default.cfg"
  If cfg = "keyfile" Then filepath = App.Path + "\key.cfg"

  If Dir(filepath) <> "" Then
    fileexists = True
  Else
    loadFile = False    'file not opened, returen FALSE
    Exit Function
  End If


  Open filepath For Input As #1 '打开选择的文件
  
  
  
  
  Do Until EOF(1) '显示打开的文件
    Line Input #1, strLine
    
    
    'write code here
    'to analyze each line read from the file
    '===========================================
    
    If cfg = "profile" Then profileAnalyzer (strLine)
    If cfg = "keyfile" Then keyfileAnalyzer (strLine)
    
    '-------------------------------------------
    

  Loop
  
  Close #1 '关闭打开的文件
  
  'file opened, return TRUE
  loadFile = True
  
  If startupload = False Then
    If cfg = "profile" Then procfgReply
    If cfg = "keyfile" Then keycfgReply
  End If
  
  

End Function














Public Function profileAnalyzer(ByVal txt As String) As Integer
    Dim p As Long
    Dim l As Long
    Dim temp As String
    Dim tName As String
    Dim tVal As Long
    

    l = Len(txt)
    
    
    'skip invalid lines (e.g empty)
    If l <= 6 Then Exit Function
    
    'filter text after "//" first
    For p = 1 To l
        temp = Mid(txt, p, 2)
        If temp = "//" Then Exit For
    Next
    txt = Left(txt, p - 1)
    
    'skip invalid lines again (e.g, start with //xxxxxx)
    If txt = "" Then Exit Function
    
    l = Len(txt)
    tName = Left(txt, 5)
    tVal = Val(Right(txt, l - 6))

    'now, assign values:
    '=====================================
    If tName = "DEVID" Then
    'device ID
        joyID = tVal
        Form2.Text1.Text = joyID
    ElseIf tName = "PRSHD" Then
    'Holding mode
        If tVal = 1 Then Form2.Option1.Value = True
        If tVal = 0 Then Form2.Option2.Value = True
    ElseIf tName = "AUTOC" Then
    'auto center joystick
        If tVal = 1 Then Form2.Check3.Value = 1
        If tVal = 0 Then Form2.Check3.Value = 0
    ElseIf tName = "LFSVK" Then
    'LFS S2
        If tVal = 1 Then Form2.Check2.Value = 1
        If tVal = 0 Then Form2.Check2.Value = 0
    ElseIf tName = "NGEAR" Then
    'N gear
        If tVal = 1 Then Form2.Check4.Value = 1
        If tVal = 0 Then Form2.Check4.Value = 0
        ElseIf tName = "RGEAR" Then
    'R gear
        If tVal = 1 Then Form2.Check5.Value = 1
        If tVal = 0 Then Form2.Check5.Value = 0
    Else
    'gear effect zones:
        If tName = "X0MIN" Then gearZone(0).xmin = Val(tVal)
        If tName = "X0MAX" Then gearZone(0).xmax = Val(tVal)
        If tName = "Y0MIN" Then gearZone(0).ymin = Val(tVal)
        If tName = "Y0MAX" Then gearZone(0).ymax = Val(tVal)
                                    
        If tName = "X1MIN" Then gearZone(1).xmin = Val(tVal)
        If tName = "X1MAX" Then gearZone(1).xmax = Val(tVal)
        If tName = "Y1MIN" Then gearZone(1).ymin = Val(tVal)
        If tName = "Y1MAX" Then gearZone(1).ymax = Val(tVal)
                                    
        If tName = "X2MIN" Then gearZone(2).xmin = Val(tVal)
        If tName = "X2MAX" Then gearZone(2).xmax = Val(tVal)
        If tName = "Y2MIN" Then gearZone(2).ymin = Val(tVal)
        If tName = "Y2MAX" Then gearZone(2).ymax = Val(tVal)
                                    
        If tName = "X3MIN" Then gearZone(3).xmin = Val(tVal)
        If tName = "X3MAX" Then gearZone(3).xmax = Val(tVal)
        If tName = "Y3MIN" Then gearZone(3).ymin = Val(tVal)
        If tName = "Y3MAX" Then gearZone(3).ymax = Val(tVal)
                                    
        If tName = "X4MIN" Then gearZone(4).xmin = Val(tVal)
        If tName = "X4MAX" Then gearZone(4).xmax = Val(tVal)
        If tName = "Y4MIN" Then gearZone(4).ymin = Val(tVal)
        If tName = "Y4MAX" Then gearZone(4).ymax = Val(tVal)
                                    
        If tName = "X5MIN" Then gearZone(5).xmin = Val(tVal)
        If tName = "X5MAX" Then gearZone(5).xmax = Val(tVal)
        If tName = "Y5MIN" Then gearZone(5).ymin = Val(tVal)
        If tName = "Y5MAX" Then gearZone(5).ymax = Val(tVal)
                                    
        If tName = "X6MIN" Then gearZone(6).xmin = Val(tVal)
        If tName = "X6MAX" Then gearZone(6).xmax = Val(tVal)
        If tName = "Y6MIN" Then gearZone(6).ymin = Val(tVal)
        If tName = "Y6MAX" Then gearZone(6).ymax = Val(tVal)
                                    
        If tName = "X7MIN" Then gearZone(7).xmin = Val(tVal)
        If tName = "X7MAX" Then gearZone(7).xmax = Val(tVal)
        If tName = "Y7MIN" Then gearZone(7).ymin = Val(tVal)
        If tName = "Y7MAX" Then gearZone(7).ymax = Val(tVal)


        
    
    End If
    
    
    '-------------------------------------


End Function





Public Function keyfileAnalyzer(ByVal txt As String) As Integer
    Dim p As Long
    Dim l As Long
    Dim temp As String
    Dim tName As String
    Dim tVal As Long
    

    l = Len(txt)
    
    
    'skip invalid lines (e.g empty)
    If l <= 6 Then Exit Function
    
    'filter text after "//" first
    For p = 1 To l
        temp = Mid(txt, p, 2)
        If temp = "//" Then Exit For
    Next
    txt = Left(txt, p - 1)
    
    'skip invalid lines again (e.g, start with //xxxxxx)
    If txt = "" Then Exit Function
    
    l = Len(txt)
    tName = Left(txt, 5)
    tVal = Val(Right(txt, l - 6))

    'now, assign values:
    '=====================================
    If tVal = 0 Then
    'do nothing.
    Else
    'assign values:
        If tName = "DXGG0" Then SK(0) = Val(tVal)
        If tName = "DXGG1" Then SK(1) = Val(tVal)
        If tName = "DXGG2" Then SK(2) = Val(tVal)
        If tName = "DXGG3" Then SK(3) = Val(tVal)
        If tName = "DXGG4" Then SK(4) = Val(tVal)
        If tName = "DXGG5" Then SK(5) = Val(tVal)
        If tName = "DXGG6" Then SK(6) = Val(tVal)
        If tName = "DXGG7" Then SK(7) = Val(tVal)
        
        If tName = "LFSG0" Then VK(0) = Val(tVal)
        If tName = "LFSG1" Then VK(1) = Val(tVal)
        If tName = "LFSG2" Then VK(2) = Val(tVal)
        If tName = "LFSG3" Then VK(3) = Val(tVal)
        If tName = "LFSG4" Then VK(4) = Val(tVal)
        If tName = "LFSG5" Then VK(5) = Val(tVal)
        If tName = "LFSG6" Then VK(6) = Val(tVal)
        If tName = "LFSG7" Then VK(7) = Val(tVal)


    
    End If
    
    
    '-------------------------------------

    

End Function


Public Function keycfgReply()
    Dim i As Integer
    Dim vkdscp As String
    Dim skdscp As String
    
    For i = 0 To 7
        vkdscp = vkdscp + "0x" + Format(Hex(VK(i)), "00") + "," + Chr(9)
        skdscp = skdscp + "0x" + Format(Hex(SK(i)), "00") + "," + Chr(9)
    Next
    
    MsgBox "Keys updated/按键已更新:" + Chr(13) + _
            "        ( N," + Chr(9) + "1," + Chr(9) + "2," + Chr(9) + "3," + Chr(9) + "4," + Chr(9) + "5," + Chr(9) + "6," + Chr(9) + "R )" + Chr(13) + _
            "LFS    :" + vkdscp + Chr(13) + _
            "DX Game:" + skdscp

End Function




Public Function procfgReply()
   MsgBox "Interface options updated." + Chr(13) + _
          "界面选项已更新。"
End Function

Public Function WriteLineToFile()
        Const ForReading = 1, ForWriting = 2
        Dim fso, f
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.OpenTextFile("c:\testfile.txt", ForWriting, True)
        f.writeline "Hello   world!"
        f.writeline "VBScript   is   fun!"
        Set f = fso.OpenTextFile("c:\testfile.txt", ForReading)
        WriteLineToFile = f.ReadAll
End Function


