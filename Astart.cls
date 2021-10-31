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
'   ҳ�������
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
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
    Page.Clear '��ջ���
    
    If Not GameBegin Then
        Page.Paint 1, GW / 2 + 70, GH / 2 + 30, 250, 250, argb(pellucidity_1 * 255 / 2, Cubic((GetTickCount Mod 2500) / 2500, 1, 0, 0, 1) * 255, 0, 0), pos:=posOnCenter
        Page.Paint 1, GW / 2 - 30, GH / 2 - 60, 250, 250, argb(pellucidity_1 * 255 / 2, 0, Cubic((GetTickCount Mod 2900) / 2900, 1, 0, 0, 1) * 255, 0), pos:=posOnCenter
        Page.Paint 1, GW / 2 - 60, GH / 2 + 90, 250, 250, argb(pellucidity_1 * 255 / 2, 0, 0, Cubic((GetTickCount Mod 2700) / 2700, 1, 0, 0, 1) * 255), pos:=posOnCenter
        Page.Writes "-������⴦��ʼ��Ϸ-", GW / 2, 6 * GH / 7, 20, argb(Cubic((GetTickCount Mod 2000) / 2000, 1, 0.4, 0.4, 1) * 255, 0, 0, 0), align:=StringAlignmentCenter
        Page.DrawImage "����.png", GW / 2, GH / 2, alpha:=pellucidity_1, pos:=posOnCenter
        If CheckMouse(0, 0, GW, GH) = mMouseDown Then
            Mouse.State = 0: Mouse.button = 0
            CheckBegin = True '��ʾ��Ϸ��ʼ
        End If
     End If
     
     Page.DrawImage "����ͼƬ.png", GW / 2, GH / 2, alpha:=0.3, pos:=posOnCenter
     
     If CheckBegin Then '��ť�����ʼ�����˳���ʼ
        pellucidity_1 = pellucidity_1 - 0.1
        If pellucidity_1 <= 0 Then
            pellucidity_1 = 0: GameBegin = True: classics = True: CheckBegin = False
        End If
    End If
    
    If GameBegin Then '��Ϸ��ʼ�����ʼ���沢����ѡ�����
        If by Then
            If classics Then interface_write = "����ģʽ"
            If custom Then interface_write = "�Զ���ģʽ"
            If help Then interface_write = "����"
        End If
        'ѡ��ť1������
        Page.DrawImage "ѡ��ť1.png", GW / 10 + 20, GH / 15, alpha:=pellucidity_2(1), pos:=posOnCenter
        If CheckMouse2 = mMouseDown Then
            Mouse.State = 0: Mouse.button = 0
            classics = True: custom = False: help = False
        ElseIf CheckMouse2 = mMouseIn Then
            BGMPlay (1)
            pellucidity_2(1) = 1: by = False: interface_write = "����ģʽ"
        ElseIf CheckMouse2 = mMouseOut Then
            BGMflag(1) = 0
            pellucidity_2(1) = 0.65: by = True
        End If
        'ѡ��ť2������
        Page.DrawImage "ѡ��ť2.png", GW / 10 + 60, GH / 15, alpha:=pellucidity_2(2), pos:=posOnCenter
        If CheckMouse2 = mMouseDown Then
            Mouse.State = 0: Mouse.button = 0
            classics = False: custom = True: help = False
        ElseIf CheckMouse2 = mMouseIn Then
            BGMPlay (2)
            pellucidity_2(2) = 1: by = False: interface_write = "�Զ���ģʽ"
        ElseIf CheckMouse2 = mMouseOut Then
            BGMflag(2) = 0
            pellucidity_2(2) = 0.65: by = True
        End If
        'ѡ��ť3������
        Page.DrawImage "ѡ��ť3.png", GW / 10 + 100, GH / 15, alpha:=pellucidity_2(3), pos:=posOnCenter
        If CheckMouse2 = mMouseDown Then
            Mouse.State = 0: Mouse.button = 0
            classics = False: custom = False: help = True
        ElseIf CheckMouse2 = mMouseIn Then
            BGMPlay (3)
            pellucidity_2(3) = 1: by = False: interface_write = "����"
        ElseIf CheckMouse2 = mMouseOut Then
            BGMflag(3) = 0
            pellucidity_2(3) = 0.65: by = True
        End If
        '������ҳ������
        Page.Writes interface_write, 4 * GW / 5, GH / 15 - 10, 20, argb(255, 255, 255, 255), align:=StringAlignmentCenter
        '���ƾ���ģʽ���ఴť
        If classics Then
            Page.DrawImage "��ģʽ.png", GW / 2, GH / 4, alpha:=pellucidity_3(1), pos:=posOnCenter
            If CheckMouse2 = mMouseDown Then
                reseaux = 9: reseauy = 9
                mine = 10
                reelw = Length(reseaux, GW)
                reelh = Length(reseauy, GH)
                record = "��ģʽ"
                ECore.SimpleMsg cue, "��ʾ", StrArray("�õ�")
                Mouse.State = 0: Mouse.button = 0
                ECore.NewTransform 1, 500, "Amain"
            ElseIf CheckMouse2 = mMouseIn Then
                BGMPlay (4)
                pellucidity_3(1) = 1
            Else
                BGMflag(4) = 0
                pellucidity_3(1) = 0.65
            End If
            Page.DrawImage "�м�ģʽ.png", GW / 2, GH / 2, alpha:=pellucidity_3(2), pos:=posOnCenter
            If CheckMouse2 = mMouseDown Then
                reseaux = 16: reseauy = 16
                mine = 48
                reelw = Length(reseaux, GW)
                reelh = Length(reseauy, GH)
                record = "�м�ģʽ"
                ECore.SimpleMsg cue, "��ʾ", StrArray("�õ�")
                Mouse.State = 0: Mouse.button = 0
                ECore.NewTransform 1, 500, "Amain"
            ElseIf CheckMouse2 = mMouseIn Then
                BGMPlay (5)
                pellucidity_3(2) = 1
            Else
                BGMflag(5) = 0
                pellucidity_3(2) = 0.65
            End If
            Page.DrawImage "�߼�ģʽ.png", GW / 2, 3 * GH / 4, alpha:=pellucidity_3(3), pos:=posOnCenter
            If CheckMouse2 = mMouseDown Then
                reseaux = 30: reseauy = 16
                mine = 99
                reelw = Length(reseaux, GW)
                reelh = Length(reseauy, GH)
                record = "�߼�ģʽ"
                ECore.SimpleMsg cue, "��ʾ", StrArray("�õ�")
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
            'Page.DrawImage "�Զ���ģʽ.png", GW / 2, GH / 2, alpha:=1, pos:=posOnCenter
            '�����ı���������======================================================================================================
            Page.ShowEdit Line, 2, GW / 2 - 30, GH / 3 - 50, 200, 50, argb(255, 0, 0, 0), argb(255, 255, 255, 255), argb(255, 255, 255, 255), argb(255, 255, 255, 255), 20, size:=16
            Page.Writes "������������", GW / 2 - 120, GH / 3 - 50, 16, argb(255, 0, 0, 0), align:=StringAlignmentCenter
            Page.Writes "(9<=Line<=30)", GW / 2 - 120, GH / 3 - 20, 12, argb(255 / 2, 255, 0, 0), align:=StringAlignmentCenter
            Page.ShowEdit Column, 2, GW / 2 - 30, GH / 2 - 50, 200, 50, argb(255, 0, 0, 0), argb(255, 255, 255, 255), argb(255, 255, 255, 255), argb(255, 255, 255, 255), 20, size:=16
            Page.Writes "������������", GW / 2 - 120, GH / 2 - 50, 16, argb(255, 0, 0, 0), align:=StringAlignmentCenter
            Page.Writes "(9<=Column<=30)", GW / 2 - 120, GH / 2 - 20, 12, argb(255 / 2, 255, 0, 0), align:=StringAlignmentCenter
            Page.ShowEdit mine_number, 2, GW / 2 - 30, 2 * GH / 3 - 50, 200, 50, argb(255, 0, 0, 0), argb(255, 255, 255, 255), argb(255, 255, 255, 255), argb(255, 255, 255, 255), 20, size:=16
            Page.Writes "������������", GW / 2 - 120, 2 * GH / 3 - 50, 16, argb(255, 0, 0, 0), align:=StringAlignmentCenter
            Page.Writes "(10<=Mine<=Line*Column-1)", GW / 2 - 120, 2 * GH / 3 - 20, 12, argb(255 / 2, 255, 0, 0), align:=StringAlignmentCenter
            '=================================================================================================================
            If Page.ShowColorButton(2, 130 - 50, 555, 115, 40, "���", argb(200, 0, 0, 0), argb(120, 0, 183, 195), argb(220, 0, 183, 195), 20, size:=18) = mMouseUp Then
                Line = "": Column = "": mine_number = ""
            End If
            If Page.ShowColorButton(2, 130 + 155, 555, 115, 40, "ȷ��", argb(200, 0, 0, 0), argb(120, 0, 183, 195), argb(220, 0, 183, 195), 20, size:=18) = mMouseUp Then
                Message = Chr$(10) + "����������"
                If Line = "" Or Column = "" Or mine_number = "" Then Message = Chr$(10) + "���벿�ֲ���Ϊ��" + Message
                If Line <> "" And format(Line) = Len(Line) Then
                    If Val(Line) < 9 Then Message = Chr$(10) + "��������ֵ����С��9" + Message
                    If Val(Line) > 30 Then Message = Chr$(10) + "��������ֵ���ô���30" + Message
                ElseIf Line <> "" And format(Line) <> Len(Line) Then
                    Message = Chr$(10) + "���������ʽ����" + Message
                End If
                If Column <> "" And format(Column) = Len(Column) Then
                    If Val(Column) < 9 Then Message = Chr$(10) + "��������ֵ����С��9" + Message
                    If Val(Column) > 30 Then Message = Chr$(10) + "��������ֵ���ô���30" + Message
                ElseIf Column <> "" And format(Column) <> Len(Column) Then
                    Message = Chr$(10) + "���������ʽ����" + Message
                End If
                If mine_number <> "" And format(mine_number) = Len(mine_number) Then
                    If Val(mine_number) < 10 Then Message = Chr$(10) + "��������ֵ����С��10" + Message
                    If Line <> "" And Column <> "" And Val(mine_number) > Val(Line) * Val(Column) - 1 Then Message = Chr$(10) + "��������ֵ���ô���" + CStr(Val(Line) * Val(Column) - 1) + Message
                ElseIf mine_number <> "" And format(mine_number) <> Len(mine_number) Then
                    Message = Chr$(10) + "���������ʽ����" + Message
                End If
                If Message <> Chr$(10) + "����������" Then
                    Line = "": Column = "": mine_number = ""
                    ECore.SimpleMsg Message, "��ʾ", StrArray("�õ�")
                Else
                    reseaux = Val(Column): reseauy = Val(Line)
                    mine = Val(mine_number)
                    reelw = Length(reseaux, GW)
                    reelh = Length(reseauy, GH)
                    Line = "": Column = "": mine_number = ""
                    record = "Line=" + CStr(reseauy) + ":Column=" + CStr(reseaux) + ":mine=" + CStr(mine)
                    ECore.SimpleMsg cue, "��ʾ", StrArray("�õ�")
                    Mouse.State = 0: Mouse.button = 0
                    ECore.NewTransform 1, 500, "Amain"
                End If
            End If
        ElseIf help Then '����������
            Page.DrawImage "help.png", GW / 2, GH / 2, alpha:=1, pos:=posOnCenter
        End If
    End If
    '================================================================
    '   ʾ������������
    '   Page.Writes "��������", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '�������¼�
    'Direction:����Depth�����
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '������Ϣ����
End Sub
Public Sub Enter()
    'ҳ������¼�
    reseaux = 0: reseauy = 0
    record = ""
    GameBegin = False '��ʾ��Ϸ��δ��ʼ
    CheckBegin = False '��ʾ��Ϸ������δ��ʼ
    pellucidity_1 = 1 '��ʼ͸����
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
    'ҳ���뿪�¼�
End Sub
Public Sub Gotfocus()
    '����ȡ�ý����¼�
End Sub
Public Sub Lostfocus()
    '����ʧȥ�����¼�
End Sub
Public Sub AnimationDone(id As Integer)
    '�������Ž����¼�
    'id�������Ķ���ID
End Sub
Private Sub Class_Initialize()
    '����ҳ�������
    Set Page = New GPage
    Page.Create Me
    '������Ϸ��Դ
    Page.Res.NewImages App.path & "\assets"
    '����ҳ��
    ECore.Add Page, "Astart"
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub