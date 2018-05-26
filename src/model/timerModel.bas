Attribute VB_Name = "modTimers"
'=======================================================================================
'��Ȩ:
'    ��Դ��ĳ������޴ӿ���������90���Ĵ��뾭���˸�д���Ż��������˲�ӵ�д�Դ����κΰ�
'    Ȩ������ѧ���о�
'=======================================================================================


'---------------------------------------------------------------------------------------
' ģ����:   modTimers
' ��������: 2006-10-12 22:45
' ����:     ����
' QQ:       38284035 �����BUG���и��õĽ��鲻��Q�ҹ�ͬ̽��

' ģ������: �и���VB�Դ���Timer�ؼ����ܻ�����Խ������粻�ܴ�����60���ʱ���¼���ʱ��
'           ��׼ȷ�����TIMER������ִ����ĳЩ���룬���Ӱ�쵽�´�ִ�е�׼ȷ�ԣ�����˷�
'           װ��һ�����У��Է�����á�

'           �书���У�
'           1.������ʱ�伤��
'             �趨δ����ĳһ��ʱ�䣬�磺2006-10-12 33:45:05���򵽴��ʱ��ʱ���
'           2.�����ʱ�伤��
'             �趨ĳ������������ÿ�롢ÿСʱ��ÿ��
'           3.�����趨ֻ����һ��
'             ���ú������д����ֹͣ��TIMER����ֻ����һ��

'---------------------------------------------------------------------------------------

Option Explicit

'���Timer�������ļ���
Private m_TimerColection   As Collection

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nEventID As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nEventID As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'���һ��Timer
Public Sub AddTimer(ByRef Timer As timerClass, ByVal iInterval As Long, ByVal bRunOnce As Boolean)
    '�������δʵ��������ʵ����
    If m_TimerColection Is Nothing Then
        Set m_TimerColection = New Collection
    End If
    
    Timer.ID = SetTimer(0, 0, iInterval, AddressOf TimerProc)
    m_TimerColection.Add ObjPtr(Timer) & ";" & IIf(bRunOnce, 1, 0), Timer.ID & ""
    
End Sub

'ɾ��һ��Timer
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

'������ʱ��ʱ���еĹ���
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
