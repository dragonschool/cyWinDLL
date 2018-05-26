Attribute VB_Name = "modTimers"
'=======================================================================================
'版权:
'    此源码的出处已无从考究，其中90％的代码经本人改写及优化，但本人不拥有此源码的任何版
'    权，纯属学术研究
'=======================================================================================


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

'存放Timer计数器的集合
Private m_TimerColection   As Collection

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nEventID As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nEventID As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'添加一个Timer
Public Sub AddTimer(ByRef Timer As timerClass, ByVal iInterval As Long, ByVal bRunOnce As Boolean)
    '如果集合未实例化则先实例化
    If m_TimerColection Is Nothing Then
        Set m_TimerColection = New Collection
    End If
    
    Timer.ID = SetTimer(0, 0, iInterval, AddressOf TimerProc)
    m_TimerColection.Add ObjPtr(Timer) & ";" & IIf(bRunOnce, 1, 0), Timer.ID & ""
    
End Sub

'删除一个Timer
Public Sub RemoveTimer(ByRef Timer As timerClass)
On Error GoTo ErrHandler

    m_TimerColection.Remove Timer.ID & ""
    KillTimer 0, Timer.ID
    Timer.ID = 0
    If m_TimerColection.Count = 0 Then
        Set m_TimerColection = Nothing
        
    End If

ErrHandler:
    
End Sub

'当到达时间时则行的过程
Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal EventID As Long, ByVal SysTime As Long)

Dim lPointer As Long
Dim objTimer As timerClass

On Error GoTo ErrHandler

    Dim sTmp As String
    sTmp = m_TimerColection.Item(EventID & "")
    lPointer = Split(sTmp, ";")(0)
    Set objTimer = PtrObj(lPointer)
    objTimer.RaiseTimerEvent Split(sTmp, ";")(1)
    Set objTimer = Nothing
    Exit Sub
ErrHandler:

End Sub

Private Function PtrObj(ByVal lPointer As Long) As Object
Dim objTimer   As Object
    CopyMemory objTimer, lPointer, 4&
    Set PtrObj = objTimer
    CopyMemory objTimer, 0&, 4&
    
End Function
