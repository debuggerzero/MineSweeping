VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Astart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
Option Explicit
'   页面绘制器
    Dim Page As GPage
'======================================

Public Function Length(a As Integer, b As Single) As Integer
    Length = 2 * (b \ 2 - 180) + 40 * a
End Function

Public Function format(a As String) As Integer
    format = 0
    For roundi = 1 To Len(a)
        If Mid(a, roundi, 1) >= "0" And Mid(a, roundi, 1) <= "9" Then
            format = format + 1
        End If
    Next
End Function

Public Sub BGMPlay(a As Integer)
        If BGMflag(a) <= 1 Then BGMflag(a) = BGMflag(a) + 1
        If BGMflag(a) = 1 Then MusicList.Play "flag.wav"
End Sub

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布
    
    If Not GameBegin Then
        Page.Paint 1, GW / 2 + 70, GH / 2 + 30, 250, 250, argb(pellucidity_1 * 255 / 2, Cubic((GetTickCount Mod 2500) / 2500, 1, 0, 0, 1) * 255, 0, 0), pos:=posOnCenter
        Page.Paint 1, GW / 2 - 30, GH / 2 - 60, 250, 250, argb(pellucidity_1 * 255 / 2, 0, Cubic((GetTickCount Mod 2900) / 2900, 1, 0, 0, 1) * 255, 0), pos:=posOnCenter
        Page.Paint 1, GW / 2 - 60, GH / 2 + 90, 250, 250, argb(pellucidity_1 * 255 / 2, 0, 0, Cubic((GetTickCount Mod 2700) / 2700, 1, 0, 0, 1) * 255), pos:=posOnCenter
        Page.Writes "-点击任意处开始游戏-", GW / 2, 6 * GH / 7, 20, argb(Cubic((GetTickCount Mod 2000) / 2000, 1, 0.4, 0.4, 1) * 255, 0, 0, 0), align:=StringAlignmentCenter
        Page.DrawImage "标题.png", GW / 2, GH / 2, alpha:=pellucidity_1, pos:=posOnCenter
        If CheckMouse(0, 0, GW, GH) = mMouseDown Then
            Mouse.State = 0: Mouse.button = 0
            CheckBegin = True '表示游戏开始
        End If
     End If
     
     Page.DrawImage "背景图片.png", GW / 2, GH / 2, alpha:=0.3, pos:=posOnCenter
     
     If CheckBegin Then '按钮点击初始界面退场开始
        pellucidity_1 = pellucidity_1 - 0.1
        If pellucidity_1 <= 0 Then
            pellucidity_1 = 0: GameBegin = True: classics = True: CheckBegin = False
        End If
    End If
    
    If GameBegin Then '游戏开始进入初始界面并加载选择界面
        If by Then
            If classics Then interface_write = "经典模式"
            If custom Then interface_write = "自定义模式"
            If help Then interface_write = "帮助"
        End If
        '选择按钮1的载入
        Page.DrawImage "选择按钮1.png", GW / 10 + 20, GH / 15, alpha:=pellucidity_2(1), pos:=posOnCenter
        If CheckMouse2 = mMouseDown Then
            Mouse.State = 0: Mouse.button = 0
            classics = True: custom = False: help = False
        ElseIf CheckMouse2 = mMouseIn Then
            BGMPlay (1)
            pellucidity_2(1) = 1: by = False: interface_write = "经典模式"
        ElseIf CheckMouse2 = mMouseOut Then
            BGMflag(1) = 0
            pellucidity_2(1) = 0.65: by = True
        End If
        '选择按钮2的载入
        Page.DrawImage "选择按钮2.png", GW / 10 + 60, GH / 15, alpha:=pellucidity_2(2), pos:=posOnCenter
        If CheckMouse2 = mMouseDown Then
            Mouse.State = 0: Mouse.button = 0
            classics = False: custom = True: help = False
        ElseIf CheckMouse2 = mMouseIn Then
            BGMPlay (2)
            pellucidity_2(2) = 1: by = False: interface_write = "自定义模式"
        ElseIf CheckMouse2 = mMouseOut Then
            BGMflag(2) = 0
            pellucidity_2(2) = 0.65: by = True
        End If
        '选择按钮3的载入
        Page.DrawImage "选择按钮3.png", GW / 10 + 100, GH / 15, alpha:=pellucidity_2(3), pos:=posOnCenter
        If CheckMouse2 = mMouseDown Then
            Mouse.State = 0: Mouse.button = 0
            classics = False: custom = False: help = True
        ElseIf CheckMouse2 = mMouseIn Then
            BGMPlay (3)
            pellucidity_2(3) = 1: by = False: interface_write = "帮助"
        ElseIf CheckMouse2 = mMouseOut Then
            BGMflag(3) = 0
            pellucidity_2(3) = 0.65: by = True
        End If
        '绘制主页文字栏
        Page.Writes interface_write, 4 * GW / 5, GH / 15 - 10, 20, argb(255, 255, 255, 255), align:=StringAlignmentCenter
        '绘制经典模式各类按钮
        If classics Then
            Page.DrawImage "简单模式.png", GW / 2, GH / 4, alpha:=pellucidity_3(1), pos:=posOnCenter
            If CheckMouse2 = mMouseDown Then
                reseaux = 9: reseauy = 9
                mine = 10
                reelw = Length(reseaux, GW)
                reelh = Length(reseauy, GH)
                record = "简单模式"
                ECore.SimpleMsg cue, "提示", StrArray("好的")
                Mouse.State = 0: Mouse.button = 0
                ECore.NewTransform 1, 500, "Amain"
            ElseIf CheckMouse2 = mMouseIn Then
                BGMPlay (4)
                pellucidity_3(1) = 1
            Else
                BGMflag(4) = 0
                pellucidity_3(1) = 0.65
            End If
            Page.DrawImage "中级模式.png", GW / 2, GH / 2, alpha:=pellucidity_3(2), pos:=posOnCenter
            If CheckMouse2 = mMouseDown Then
                reseaux = 16: reseauy = 16
                mine = 48
                reelw = Length(reseaux, GW)
                reelh = Length(reseauy, GH)
                record = "中级模式"
                ECore.SimpleMsg cue, "提示", StrArray("好的")
                Mouse.State = 0: Mouse.button = 0
                ECore.NewTransform 1, 500, "Amain"
            ElseIf CheckMouse2 = mMouseIn Then
                BGMPlay (5)
                pellucidity_3(2) = 1
            Else
                BGMflag(5) = 0
                pellucidity_3(2) = 0.65
            End If
            Page.DrawImage "高级模式.png", GW / 2, 3 * GH / 4, alpha:=pellucidity_3(3), pos:=posOnCenter
            If CheckMouse2 = mMouseDown Then
                reseaux = 30: reseauy = 16
                mine = 99
                reelw = Length(reseaux, GW)
                reelh = Length(reseauy, GH)
                record = "高级模式"
                ECore.SimpleMsg cue, "提示", StrArray("好的")
                Mouse.State = 0: Mouse.button = 0
                ECore.NewTransform 1, 500, "Amain"
            ElseIf CheckMouse2 = mMouseIn Then
                BGMPlay (6)
                pellucidity_3(3) = 1
            Else
                BGMflag(6) = 0
                pellucidity_3(3) = 0.65
            End If
        ElseIf custom Then
            'Page.DrawImage "自定义模式.png", GW / 2, GH / 2, alpha:=1, pos:=posOnCenter
            '绘制文本框与文字======================================================================================================
            Page.ShowEdit Line, 2, GW / 2 - 30, GH / 3 - 50, 200, 50, argb(255, 0, 0, 0), argb(255, 255, 255, 255), argb(255, 255, 255, 255), argb(255, 255, 255, 255), 20, size:=16
            Page.Writes "请输入行数：", GW / 2 - 120, GH / 3 - 50, 16, argb(255, 0, 0, 0), align:=StringAlignmentCenter
            Page.Writes "(9<=Line<=30)", GW / 2 - 120, GH / 3 - 20, 12, argb(255 / 2, 255, 0, 0), align:=StringAlignmentCenter
            Page.ShowEdit Column, 2, GW / 2 - 30, GH / 2 - 50, 200, 50, argb(255, 0, 0, 0), argb(255, 255, 255, 255), argb(255, 255, 255, 255), argb(255, 255, 255, 255), 20, size:=16
            Page.Writes "请输入列数：", GW / 2 - 120, GH / 2 - 50, 16, argb(255, 0, 0, 0), align:=StringAlignmentCenter
            Page.Writes "(9<=Column<=30)", GW / 2 - 120, GH / 2 - 20, 12, argb(255 / 2, 255, 0, 0), align:=StringAlignmentCenter
            Page.ShowEdit mine_number, 2, GW / 2 - 30, 2 * GH / 3 - 50, 200, 50, argb(255, 0, 0, 0), argb(255, 255, 255, 255), argb(255, 255, 255, 255), argb(255, 255, 255, 255), 20, size:=16
            Page.Writes "请输入雷数：", GW / 2 - 120, 2 * GH / 3 - 50, 16, argb(255, 0, 0, 0), align:=StringAlignmentCenter
            Page.Writes "(10<=Mine<=Line*Column-1)", GW / 2 - 120, 2 * GH / 3 - 20, 12, argb(255 / 2, 255, 0, 0), align:=StringAlignmentCenter
            '=================================================================================================================
            If Page.ShowColorButton(2, 130 - 50, 555, 115, 40, "清空", argb(200, 0, 0, 0), argb(120, 0, 183, 195), argb(220, 0, 183, 195), 20, size:=18) = mMouseUp Then
                Line = "": Column = "": mine_number = ""
            End If
            If Page.ShowColorButton(2, 130 + 155, 555, 115, 40, "确定", argb(200, 0, 0, 0), argb(120, 0, 183, 195), argb(220, 0, 183, 195), 20, size:=18) = mMouseUp Then
                Message = Chr$(10) + "请重新输入"
                If Line = "" Or Column = "" Or mine_number = "" Then Message = Chr$(10) + "输入部分不得为空" + Message
                If Line <> "" And format(Line) = Len(Line) Then
                    If Val(Line) < 9 Then Message = Chr$(10) + "行中输入值不得小于9" + Message
                    If Val(Line) > 30 Then Message = Chr$(10) + "行中输入值不得大于30" + Message
                ElseIf Line <> "" And format(Line) <> Len(Line) Then
                    Message = Chr$(10) + "行中输入格式有误" + Message
                End If
                If Column <> "" And format(Column) = Len(Column) Then
                    If Val(Column) < 9 Then Message = Chr$(10) + "列中输入值不得小于9" + Message
                    If Val(Column) > 30 Then Message = Chr$(10) + "列中输入值不得大于30" + Message
                ElseIf Column <> "" And format(Column) <> Len(Column) Then
                    Message = Chr$(10) + "列中输入格式有误" + Message
                End If
                If mine_number <> "" And format(mine_number) = Len(mine_number) Then
                    If Val(mine_number) < 10 Then Message = Chr$(10) + "雷数输入值不得小于10" + Message
                    If Line <> "" And Column <> "" And Val(mine_number) > Val(Line) * Val(Column) - 1 Then Message = Chr$(10) + "雷数输入值不得大于" + CStr(Val(Line) * Val(Column) - 1) + Message
                ElseIf mine_number <> "" And format(mine_number) <> Len(mine_number) Then
                    Message = Chr$(10) + "雷数输入格式有误" + Message
                End If
                If Message <> Chr$(10) + "请重新输入" Then
                    Line = "": Column = "": mine_number = ""
                    ECore.SimpleMsg Message, "提示", StrArray("好的")
                Else
                    reseaux = Val(Column): reseauy = Val(Line)
                    mine = Val(mine_number)
                    reelw = Length(reseaux, GW)
                    reelh = Length(reseauy, GH)
                    Line = "": Column = "": mine_number = ""
                    record = "Line=" + CStr(reseauy) + ":Column=" + CStr(reseaux) + ":mine=" + CStr(mine)
                    ECore.SimpleMsg cue, "提示", StrArray("好的")
                    Mouse.State = 0: Mouse.button = 0
                    ECore.NewTransform 1, 500, "Amain"
                End If
            End If
        ElseIf help Then '帮助的载入
            Page.DrawImage "help.png", GW / 2, GH / 2, alpha:=1, pos:=posOnCenter
        End If
    End If
    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
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
    reseaux = 0: reseauy = 0
    record = ""
    GameBegin = False '表示游戏还未开始
    CheckBegin = False '表示游戏动画还未开始
    pellucidity_1 = 1 '初始透明度
    For ti = 1 To 3
        pellucidity_2(ti) = 1
        pellucidity_3(ti) = 1
    Next
    by = False
    classics = False
    custom = False
    help = False
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
    ECore.Add Page, "Astart"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
