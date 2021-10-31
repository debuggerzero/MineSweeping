VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Amain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Option Explicit
    Dim Page As GPage
    Dim number_hit As Integer
    Dim set_x As Integer, set_y As Integer
    Dim Failure As Boolean, GWin As Boolean
    Dim flag(0 To 1) As Integer
    Dim amount As Integer
    Dim starttime As Long
'======================================

'=================================================================
'x 鼠标点击的x坐标
'y 鼠标点击的y坐标
'以下子过程和自定义函数中出现的参数均为此除非特殊注释
'=================================================================
Public Function Numberofmines(X As Integer, y As Integer) As Integer '自定义函数计算周围雷的个数
'======================计算点击角落时周围的雷数============================
    If X = 1 And y = 1 Then
        Numberofmines = Abs(CInt(reseau(X + 1, y).bomb) + CInt(reseau(X + 1, y + 1).bomb) + CInt(reseau(X, y + 1).bomb))
    ElseIf X = 1 And y = UBound(reseau, 2) Then
        Numberofmines = Abs(CInt(reseau(X, y - 1).bomb) + CInt(reseau(X + 1, y - 1).bomb) + CInt(reseau(X + 1, y).bomb))
    ElseIf X = UBound(reseau, 1) And y = 1 Then
        Numberofmines = Abs(CInt(reseau(X - 1, y).bomb) + CInt(reseau(X - 1, y + 1).bomb) + CInt(reseau(X, y + 1).bomb))
    ElseIf X = UBound(reseau, 1) And y = UBound(reseau, 2) Then
        Numberofmines = Abs(CInt(reseau(X - 1, y).bomb) + CInt(reseau(X - 1, y - 1).bomb) + CInt(reseau(X, y - 1).bomb))
'==================================================================
'======================计算非角落时在边界的周围雷数==========================
    ElseIf y = 1 And X > 1 And X < UBound(reseau, 1) Then
        Numberofmines = Abs(CInt(reseau(X + 1, y).bomb) + CInt(reseau(X - 1, y).bomb) + CInt(reseau(X, y + 1).bomb) + _
        CInt(reseau(X - 1, y + 1).bomb) + CInt(reseau(X + 1, y + 1).bomb))
    ElseIf y = UBound(reseau, 2) And X > 1 And X < UBound(reseau, 1) Then
        Numberofmines = Abs(CInt(reseau(X + 1, y).bomb) + CInt(reseau(X - 1, y).bomb) + CInt(reseau(X, y - 1).bomb) + _
        CInt(reseau(X - 1, y - 1).bomb) + CInt(reseau(X + 1, y - 1).bomb))
    ElseIf X = 1 And y > 1 And y < UBound(reseau, 2) Then
        Numberofmines = Abs(CInt(reseau(X, y - 1).bomb) + CInt(reseau(X, y + 1).bomb) + CInt(reseau(X + 1, y + 1).bomb) _
        + CInt(reseau(X + 1, y - 1).bomb) + CInt(reseau(X + 1, y).bomb))
    ElseIf X = UBound(reseau, 1) And y > 1 And y < UBound(reseau, 2) Then
        Numberofmines = Abs(CInt(reseau(X, y - 1).bomb) + CInt(reseau(X, y + 1).bomb) + CInt(reseau(X - 1, y).bomb) + _
        CInt(reseau(X - 1, y - 1).bomb) + CInt(reseau(X - 1, y + 1).bomb))
'==================================================================
'======================计算在中间时计算雷的个数============================
    Else
        Numberofmines = Abs(CInt(reseau(X, y - 1).bomb) + CInt(reseau(X, y + 1).bomb) + CInt(reseau(X - 1, y).bomb) + _
        CInt(reseau(X + 1, y).bomb) + CInt(reseau(X - 1, y + 1).bomb) + CInt(reseau(X - 1, y - 1).bomb) + CInt(reseau(X + _
        1, y - 1).bomb) + CInt(reseau(X + 1, y + 1).bomb))
    End If
'==================================================================
End Function

Public Function Number() As Boolean
    Number = False
    number_hit = number_hit + 1
    If number_hit = 1 Then
        Number = True
    End If
End Function

Public Sub Initial() '游戏初始化
    ReDim reseau(1 To reseaux, 1 To reseauy)
    ReDim minex(1 To mine)
    ReDim miney(1 To mine)
    ReDim BGMfocus(1 To reseaux, 1 To reseauy)
    For roundi = LBound(reseau, 1) To UBound(reseau, 1)
            For roundj = LBound(reseau, 2) To UBound(reseau, 2)
                With reseau(roundi, roundj)
                    .named = "游戏网格.png"
                    .bomb = False
                    .Rclick = True
                    .Lclick = True
                    .Number = 0
                    .transparet = 0.7
                End With
            Next roundj
    Next roundi
End Sub

Public Sub BGMPlay(a As Integer, b As Integer)
    If BGMfocus(a, b) <= 1 Then BGMfocus(a, b) = BGMfocus(a, b) + 1
    If BGMfocus(a, b) = 1 Then MusicList.Play "focus.wav"
End Sub

Public Sub Creat_mine() '生成地雷
    Randomize
    For roundz = 1 To mine
        minex(roundz) = Int(Rnd * reseaux + 1)
        miney(roundz) = Int(Rnd * reseauy + 1)
        mine_again = True
        Do While mine_again
            If reseau(minex(roundz), miney(roundz)).bomb Or Not reseau(minex(roundz), miney(roundz)).Rclick Then
                minex(roundz) = Int(Rnd * reseaux + 1)
                miney(roundz) = Int(Rnd * reseauy + 1)
            Else
                mine_again = False
            End If
        Loop
        reseau(minex(roundz), miney(roundz)).bomb = True
        'reseau(minex(roundz), miney(roundz)).named = "炸弹.png" '开发者测试使用
    Next
End Sub

Public Sub Lbutton(X As Integer, y As Integer) '鼠标左键点击过程
    If X >= 1 And X <= reseaux And y >= 1 And y <= reseauy Then
        If Not reseau(X, y).bomb And reseau(X, y).Rclick And reseau(X, y).Lclick Then
            reseau(X, y).Rclick = False: reseau(X, y).Lclick = False
            If Number() Then Call Creat_mine
            If Numberofmines(X, y) = 0 Then
                Lbutton X, y - 1: Lbutton X, y + 1: Lbutton X + 1, y: Lbutton X - 1, y
                Lbutton X - 1, y - 1: Lbutton X + 1, y + 1: Lbutton X + 1, y - 1: Lbutton X - 1, y + 1
            Else
                reseau(X, y).named = CStr(Numberofmines(X, y)) & ".png"
            End If
        ElseIf reseau(X, y).bomb And reseau(X, y).Rclick And reseau(X, y).Lclick Then
            reseau(X, y).named = "x.png"
            MusicList.Play "boom.wav"
            Failure = True
        End If
    End If
End Sub

Public Sub Rbutton(X As Integer, y As Integer) '鼠标右键点击过程
    If reseau(X, y).Rclick Then
        reseau(X, y).Number = reseau(X, y).Number + 1
        If reseau(X, y).Number > 2 Then reseau(X, y).Number = 0
        reseau(X, y).Lclick = False
        Select Case reseau(X, y).Number
            Case 0
                reseau(X, y).named = "游戏网格.png"
                reseau(X, y).Lclick = True
            Case 1
                reseau(X, y).named = "红旗.png"
            Case 2
                reseau(X, y).named = "问号.png"
        End Select
    End If
End Sub

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布
    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    If CheckMouse(GW / 15 - 25, GH / 20 - 25, 50, 50) = mMouseDown And Mouse.button = 1 Then
        Mouse.State = 0: Mouse.button = 0
        ECore.NewTransform 1, 500, "Astart"
        CheckBegin = True
    End If

    StartScrollArea Page, junior '打开卷轴
        Page.DrawImage "游戏背景.png", GW / 2 + reelx, GH / 2 + reely, alpha:=0.5, pos:=posOnCenter
        For roundi = LBound(reseau, 1) To UBound(reseau, 1)
            For roundj = LBound(reseau, 2) To UBound(reseau, 2)
                reseau(roundi, roundj).X = GW / 2 - 180 + (roundi - 1) * 40
                reseau(roundi, roundj).y = GH / 2 - 180 + (roundj - 1) * 40
                Page.DrawImage reseau(roundi, roundj).named, reseau(roundi, roundj).X, reseau(roundi, roundj).y, alpha:=reseau(roundi, roundj).transparet
                If Failure Then
                    If reseau(roundi, roundj).bomb And reseau(roundi, roundj).named <> "x.png" Then reseau(roundi, roundj).named = "炸弹.png"
                    reseau(roundi, roundj).transparet = 0.3
                    set_x = 0: set_y = 0
                    GoTo Finish
                ElseIf GWin Then
                    If reseau(roundi, roundj).bomb Then reseau(roundi, roundj).named = "pass.png"
                    reseau(roundi, roundj).transparet = 0.3
                    set_x = 0: set_y = 0
                    GoTo Finish
                End If
                
                If CheckMouse(reseau(roundi, roundj).X - reelx, reseau(roundi, roundj).y - reely, 40 - 1, 40 - 1) = mMouseDown And Mouse.button = 1 Then '鼠标左键按下事件
                    Mouse.button = 0: Mouse.State = 0
                    Lbutton roundi, roundj
                ElseIf CheckMouse(reseau(roundi, roundj).X - reelx, reseau(roundi, roundj).y - reely, 40 - 1, 40 - 1) = mMouseDown And Mouse.button = 2 Then '鼠标右键按下事件
                    Mouse.button = 0: Mouse.State = 0
                    Rbutton roundi, roundj
                ElseIf CheckMouse(reseau(roundi, roundj).X - reelx, reseau(roundi, roundj).y - reely, 40 - 1, 40 - 1) = mMouseIn Then '指针经过事件
                    Call BGMPlay(roundi, roundj)
                    set_x = roundi: set_y = roundj
                    If reseau(roundi, roundj).Rclick Then
                        reseau(roundi, roundj).transparet = 1
                    Else
                        reseau(roundi, roundj).transparet = 0.7
                    End If
                ElseIf CheckMouse(reseau(roundi, roundj).X - reelx, reseau(roundi, roundj).y - reely, 40 - 1, 40 - 1) = mMouseOut Then '指针离开事件
                    BGMfocus(roundi, roundj) = 0
                    If reseau(roundi, roundj).Rclick Then
                        reseau(roundi, roundj).transparet = 0.8
                    Else
                        reseau(roundi, roundj).transparet = 0.6
                    End If
                End If
                If reseau(roundi, roundj).named = "红旗.png" Or reseau(roundi, roundj).named = "问号.png" Then
                    flag(0) = flag(0) + 1
                End If
                If reseau(roundi, roundj).named = "红旗.png" And reseau(roundi, roundj).bomb Then
                    flag(1) = flag(1) + 1
                End If
                If flag(0) = mine And flag(1) = mine Then
                    GWin = True
                End If
                If Not reseau(roundi, roundj).Rclick And Not reseau(roundi, roundj).bomb Then
                    amount = amount + 1
                    If amount = UBound(reseau, 1) * UBound(reseau, 2) - mine Then
                        GWin = True
                    End If
                End If
                
                If GWin Then
                    If Hillsecond > Ghour * 3600 + Gminute * 60 + Gsecond Or Hillsecond = 0 Then
                    ESave.PutData record, Ghour * 3600 + Gminute * 60 + Gsecond
                    ESave.Save
                    ECore.SimpleMsg "你破纪录啦！！" + vbCrLf + "新纪录：" + format((Ghour * 3600 + Gminute * 60 + Gsecond) \ 3600, "00") + ":" + format((Ghour * 3600 + Gminute * 60 _
                                                + Gsecond) \ 60 Mod 60, "00") + ":" + format((Ghour * 3600 + Gminute * 60 + Gsecond) Mod 60, "00"), "恭喜", StrArray("好的")
                    End If
                End If
Finish:
            Next roundj
        Next roundi
        flag(0) = 0: flag(1) = 0: amount = 0
        
        If Not Failure And Not GWin Then
            If GetTickCount - starttime >= 1000 Then
                Gsecond = Gsecond + 1
                If Gsecond >= 60 Then
                    Gminute = Gminute + 1
                    Gsecond = 0
                End If
                If Gminute >= 60 Then
                    Ghour = Ghour + 1
                End If
                starttime = GetTickCount
            End If
        End If
        
        If Failure Then
            Page.Paint 0, GW / 2 + reelx, GH / 2 + reely, GW, GH, argb(255 / 3, 0, 0, 255 / 2), pos:=posOnCenter
            Page.Writes "Game Over", GW / 2 + reelx, GH / 2 + reely, 35, argb(Cubic((GetTickCount Mod 2000) / 2000, 1, 0.4, 0.4, 1) * 255, 255, 255, 255), align:=StringAlignmentCenter
        ElseIf GWin Then
            Page.Paint 0, GW / 2 + reelx, GH / 2 + reely, GW, GH, argb(255 / 3, 0, 255 / 2, 0), pos:=posOnCenter
            Page.Writes "YOU WIN", GW / 2 + reelx, GH / 2 + reely, 35, argb(Cubic((GetTickCount Mod 2000) / 2000, 1, 0.4, 0.4, 1) * 255, 255, 255, 255), align:=StringAlignmentCenter
        End If
        
    EndScrollArea Page, 0, 0, reelx, reely, GW, GH, 1 '关闭卷轴
    
    Page.Paint 0, 0, 0, GW, GH / 10, argb(2 * 255 / 3, 255 / 2, 255 / 2, 255 / 2)
    Page.DrawImage "return.png", GW / 15, GH / 20, alpha:=1, pos:=posOnCenter
    If GWin Or Failure Then
        Page.DrawImage "begin.png", GW / 5, GH / 20, alpha:=1, pos:=posOnCenter
            If CheckMouse(GW / 5 - 20, GH / 20 - 20, 40, 40) = mMouseDown Then
                Mouse.button = 0: Mouse.State = 0
                ECore.NewTransform 1, 500, "Amain"
            End If
    End If
    Page.Writes "位置：" & Str(set_x) & "," & Str(set_y), 4 * GW / 5 - 20, GH / 20, 15, argb(255, 255, 255, 255)
    Page.Writes "记录：" + format(Hillsecond \ 3600, "00") + ":" + format(Hillsecond \ 60 Mod 60, "00") + ":" + format(Hillsecond Mod 60, "00"), _
                        4 * GW / 5 - 20, GH / 20 - 25, 15, argb(255, 255, 255, 255)
    Page.Writes format(Ghour, "00") + ":" + format(Gminute, "00") + ":" + format(Gsecond, "00"), GW / 2, GH / 20 - 20, 25, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    If IsKeyPress(VK_LEFT) Or IsKeyPress(VK_A) Then reelx = reelx - 10
    If IsKeyPress(VK_RIGHT) Or IsKeyPress(VK_D) Then reelx = reelx + 10
    If IsKeyPress(VK_UP) Or IsKeyPress(VK_W) Then reely = reely - 10
    If IsKeyPress(VK_DOWN) Or IsKeyPress(VK_S) Then reely = reely + 10
    
    If reelx < 0 Then reelx = 0
    If reely < 0 Then reely = 0
    If reelx > reelw - GW Then reelx = reelw - GW
    If reely > reelh - GH Then reely = reelh - GH
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
    Erase reseau: Erase minex: Erase miney: Erase BGMfocus
    Call Initial
    junior = CreateScrollArea(reelw, reelh)
    reelx = 0: reely = 0
    number_hit = 0
    Failure = False: GWin = False
    flag(0) = 0: flag(1) = 0
    amount = 0
    Gsecond = 0: Gminute = 0: Ghour = 0
    
    Hillsecond = Val(ESave.GetData(record))
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.path & "\assets"
    '创建页面
    ECore.Add Page, "Amain"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
