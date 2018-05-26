VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "timerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' 模块名:   modTimers
' 生成日期: 2006-10-12 22:45
' 作者:     龙堂
' QQ:       38284035 如果有BUG或有更好的建议不妨Q我共同探讨

' 模块作用: 有感于VB自带的Timer控件功能还是相对较弱，如不能处理超过60秒的时间事件，时间
'           不准确（如果TIMER过程内执行了某些代码，则会影响到下次执行的准确性），因此封
'           装到一个类中，以方便调用。

'           其功能有：
'           1.定日期时间激活
'             设定未来的某一个时间，如：2006-10-12 33:45:05，则到达该时间时激活；
'           2.定间隔时间激活
'             设定某个激活间隔，如每秒、每小时、每天
'           3.允许设定只激活一次
'             设置后无须编写代码停止该TIMER，即只激活一遍

'---------------------------------------------------------------------------------------
Option Explicit

Public Event TimerEvent()
Private m_TimerID As Long

Friend Property Let ID(ByVal plngValue As Long)
    m_TimerID = plngValue
End Property

Friend Property Get ID() As Long
    ID = m_TimerID
End Property

Private Sub Class_Terminate()
    timerClasstop
End Sub

'激活事件
Friend Sub RaiseTimerEvent(ByVal bRunOnce As Boolean)
    If bRunOnce Then
        '只要求运行一次，
        timerClasstop
    End If
    
    RaiseEvent TimerEvent
End Sub

'函数名称:timerClassetAlert
'作用：设置闹钟
'===========================================================
'整理时间：2006-9-18 22:55
'  参数：
'===========================================================
Public Sub timerClassetAlert(ByVal AlertDateTime As Date)
    timerClasstart DateDiff("s", Now, AlertDateTime), True
    
End Sub

'函数名称:timerClasstart
'作用：开始定时器
'===========================================================
'整理时间：2006-9-18 22:55
'  参数：
'===========================================================
Public Sub timerClasstart(ByVal sinSecond As Single, Optional ByVal bRunOnce As Boolean = False)
        AddTimer Me, sinSecond * (10 ^ 3), bRunOnce

End Sub

'函数名称:timerClasstop
'作用：停止定时器
'===========================================================
'整理时间：2006-9-18 22:55
'  参数：
'===========================================================
Public Sub timerClasstop()
    If m_TimerID > 0 Then
        RemoveTimer Me
    End If
End Sub

Public Sub cyDayClock(Optional ByVal bRunOnce As Boolean)
    timerClasstart 60 ^ 2 * 24, bRunOnce

End Sub

Public Sub cyHourClock(Optional ByVal bRunOnce As Boolean)
    timerClasstart 3600, bRunOnce

End Sub

Public Sub cyMinuteClock(Optional ByVal bRunOnce As Boolean)
    timerClasstart 60, bRunOnce

End Sub

Public Sub cySecondClock(Optional ByVal bRunOnce As Boolean)
    timerClasstart 1, bRunOnce

End Sub

