Attribute VB_Name = "Data"
Option Explicit
'=========================================
Public BGMboom As Integer
Public BGMflag(1 To 6) As Integer
Public BGMfocus() As Integer
'=========================================
Public GameBegin As Boolean '判断游戏是否开始
Public pellucidity_1 As Single '定义开始图片透明度
Public pellucidity_2(1 To 3) As Single '三个选择按钮透明度
Public pellucidity_3(1 To 3) As Single '三个经典选项的透明度
Public CheckBegin As Boolean '判断动画是否开始
Public classics As Boolean '是否进入经典模式界面
Public custom As Boolean '是否进入自定义模式界面
Public help As Boolean '是否进入游戏向导界面
Public ti As Integer
Public roundi As Integer, roundj As Integer, roundz As Integer
Public by As Boolean '是否经过选择按钮
Public interface_write As String '提示文字
'==============Amain有关变量=================
Public Type reseau '定义网格类
    X As Long '网格坐标x
    y As Long '网格坐标y
    Number As Integer '网格内雷数量
    transparet As Single '网格透明度
    bomb As Boolean '判断网格中是否有雷
    Rclick As Boolean '判断网格是否被右击
    Lclick As Boolean '判断网格是否被左击
    named As String '网格输出图片的名字
End Type
Public reseau() As reseau '创建网格类动态数组
Public reseaux As Integer, reseauy As Integer '定义数组下标x，y
Public mine As Integer, minex() As Integer, miney() As Integer
'分别定义地雷，地雷坐标x，y
Public mine_again As Boolean '判断是否生成重复地雷
Public reelw As Integer, reelh As Integer
Public reelx As Integer, reely As Integer
Public junior As ScrollArea
Public Line As String, Column As String, mine_number As String
Public Message As String
Public Const cue = "地图不够大？？" + vbCrLf + "试试用WSAD或↑←↓→来移动地图吧!!" _
                                + vbCrLf + "选择界面帮助中可查看更多详细信息!!!"
Public record As String, Gsecond As Long, Gminute As Long, Ghour As Long
Public Hillsecond As Long
