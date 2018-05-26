VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cyTrayEx"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event cyLeftButtonUp()
Public Event cyRightButtonUp()
Public Event cyLeftDblClick()
Public Event cyBallShow()
Public Event cyBallHide()
Public Event cyBallTimeOut()
Public Event cyBallUserClick()

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "SHELL32.DLL" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'Private Type NOTIFYICONDATA
'
'    cbSize As Long               '������NOTIFYICONDATA���ݽṹ�ĳ���
'    HWnd As Long                 '���óɴ��ڵľ��
'    Uid As Long                  'Ϊͼ�������õ�IDֵ
'    UFlags As Long               '����uCallbackMessage,hIcon,szTip�Ƿ���Ч
'    UCallbackMessage As Long     '��Ϣ���
'    HIcon As Long                '��ʾ��״̬���ϵ�ͼ��
'    SzTip As String * 64         '��ʾ��Ϣ
'
'End Type


Private Type NOTIFYICONDATA
    cbSize As Long              ' �ṹ��С(�ֽ�)
    hWnd As Long                ' ������Ϣ�Ĵ��ڵľ��
    Uid As Long                 ' Ψһ�ı�ʶ��
    UFlags As Long              ' Flags
    UCallbackMessage As Long    ' ������Ϣ�Ĵ��ڽ��յ���Ϣ
    HIcon As Long               ' ����ͼ����
    SzTip As String * 128       ' Tooltip ��ʾ�ı�
    dwState As Long             ' ����ͼ��״̬
    dwStateMask As Long         ' ״̬����
    szInfo As String * 256      ' ������ʾ�ı�
    uTimeoutOrVersion As Long   ' ������ʾ��ʧʱ���汾
                                ' uTimeout - ������ʾ��ʧʱ��(��λ:ms, 10000 -- 30000)
                                ' uVersion - �汾(0 for V4, 3 for V5)
    szInfoTitle As String * 64  ' ������ʾ����
    dwInfoFlags As Long         ' ������ʾͼ��
End Type


' dwState to NOTIFYICONDATA structure
' dwInfoFlags to NOTIFIICONDATA structure
' uFlags to NOTIFYICONDATA structure
Private Const NIF_ICON As Long = &H2
Private Const NIF_INFO As Long = &H10
Private Const NIF_MESSAGE As Long = &H1
Private Const NIF_STATE As Long = &H8
Private Const NIF_TIP As Long = &H4
' dwMessage to Shell_NotifyIcon
Private Const NIM_ADD As Long = &H0
Private Const NIM_DELETE As Long = &H2
Private Const NIM_MODIFY As Long = &H1

Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Const WM_USER = &H400
Private Const WM_NOTIFYICON = WM_USER + 1            ' �Զ�����Ϣ
Private Const WM_LBUTTONDBLCLK = &H203
Private Const GWL_WNDPROC = (-4)

' ����������ʾ���Զ�����Ϣ, 2000�²�������Щ��Ϣ
Private Const NIN_BALLOONSHOW = (WM_USER + &H2)         ' �� Balloon Tips ����ʱִ��
Private Const NIN_BALLOONHIDE = (WM_USER + &H3)         ' �� Balloon Tips ��ʧʱִ�У��� SysTrayIcon ��ɾ������
                                                        ' ��ָ���� TimeOut ʱ�䵽������� Balloon Tips �����ʧ�����ʹ���Ϣ
Private Const NIN_BALLOONTIMEOUT = (WM_USER + &H4)      ' �� Balloon Tips �� TimeOut ʱ�䵽ʱִ��
Private Const NIN_BALLOONUSERCLICK = (WM_USER + &H5)    ' ������� Balloon Tips ʱִ�С�
Private Const TRAY_CALLBACK = (WM_USER + 1001&)                                                         ' ע��:��XP��ִ��ʱ Balloon Tips ���и��رհ�ť,
                                                        ' ��������ڰ�ť�Ͻ����յ� NIN_BALLOONTIMEOUT ��Ϣ��
'������������ͼ��
Dim IconData As NOTIFYICONDATA

Dim m_iTrayHwnd As Long
Dim m_iIconHandle As Long

Public Sub RemoveTray()
    Class_Terminate
    
End Sub

Public Sub SetTray(Optional ByVal iTrayHwnd As Long, Optional ByVal IconHandle As Long, Optional ByVal sTitle As String, Optional ByVal sMsg As String, Optional iMsgType As Long = 1, Optional iSecond As Long = 10)
    
    '�����𴰿��Ա��ر�
    m_iTrayHwnd = iTrayHwnd
    m_iIconHandle = IconHandle
    
    sTitle = sTitle & vbNullChar
    sMsg = sMsg & vbNullChar
    
    modTray.objMouseWheel = ObjPtr(Me)
        
'        With IconData
'            .Uid = 0
'            .HWnd = iTrayHwnd            'frm.HWnd�ǳ���������ľ��
'            .cbSize = Len(IconData)
'            .HIcon = IconHandle    'frm.Icon.Handleָ���������ͼ��
'            .UFlags = NIF_ICON
'            .UFlags = .UFlags Or NIF_MESSAGE
'            .UCallbackMessage = TRAY_CALLBACK
'             '����������������Ϣ������һ���л�����ϸ���͡�
'            .cbSize = Len(IconData)
'        End With
'        Shell_NotifyIcon NIM_ADD, IconData
        
    
    
    With IconData
        .cbSize = Len(IconData)
        .hWnd = iTrayHwnd
        .Uid = 0
        .UFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE Or NIF_INFO Or NIF_STATE Or TRAY_CALLBACK
        .UCallbackMessage = WM_NOTIFYICON '
        .SzTip = sTitle
        .HIcon = IconHandle
        .dwState = 0
        .dwStateMask = 0
        .szInfo = sMsg
        .szInfoTitle = sTitle
        .dwInfoFlags = iMsgType
        .uTimeoutOrVersion = iSecond * 1000
    End With
    
    If preTrayProc = 0 Then    'û�г�ʼ��
       Shell_NotifyIcon NIM_ADD, IconData
       preTrayProc = SetWindowLong(iTrayHwnd, GWL_WNDPROC, AddressOf WndProcTray)
    Else                      '�ѳ�ʼ��
       Shell_NotifyIcon NIM_MODIFY, IconData
    End If
    
End Sub

Friend Function FireEvent(ByVal iAction As Long)
    'Const WM_LBUTTONUP = &H202
    'Const WM_RBUTTONUP = &H205
    Const WM_USER = &H400
    Const WM_LBUTTONDBLCLK = &H203
    
    '����������ʾ���Զ�����Ϣ, 2000�²�������Щ��Ϣ
    Const NIN_BALLOONSHOW = (WM_USER + &H2)         ' �� Balloon Tips ����ʱִ��
    Const NIN_BALLOONHIDE = (WM_USER + &H3)         ' �� Balloon Tips ��ʧʱִ�У��� SysTrayIcon ��ɾ������
                                                            ' ��ָ���� TimeOut ʱ�䵽������� Balloon Tips �����ʧ�����ʹ���Ϣ
    Const NIN_BALLOONTIMEOUT = (WM_USER + &H4)      ' �� Balloon Tips �� TimeOut ʱ�䵽ʱִ��
    Const NIN_BALLOONUSERCLICK = (WM_USER + &H5)    ' ������� Balloon Tips ʱִ�С�

    Select Case iAction
            Case &H202
                RaiseEvent cyLeftButtonUp
                
            Case &H205
                RaiseEvent cyRightButtonUp
                
            Case WM_LBUTTONDBLCLK
                RaiseEvent cyLeftDblClick
                
            Case NIN_BALLOONSHOW
                RaiseEvent cyBallShow
                
            Case NIN_BALLOONHIDE
                RaiseEvent cyBallHide
                
            Case NIN_BALLOONTIMEOUT
                RaiseEvent cyBallTimeOut
                
            Case NIN_BALLOONUSERCLICK
                RaiseEvent cyBallUserClick
                
    End Select
    
End Function

Private Sub Class_Terminate()
       
On Error Resume Next

    ' ɾ��������ͼ��
    Dim IconData As NOTIFYICONDATA
    With IconData
        .cbSize = Len(IconData)
        .hWnd = m_iTrayHwnd
        .Uid = 0
        .UFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
        .UCallbackMessage = WM_NOTIFYICON
        .SzTip = ""
        .HIcon = m_iIconHandle
    End With
    Shell_NotifyIcon NIM_DELETE, IconData
    SetWindowLong m_iTrayHwnd, GWL_WNDPROC, preTrayProc
    preTrayProc = 0

End Sub

' *********************************************
' �������̵�ͼ��
' *********************************************
Public Sub SetIcon(ByVal hIconHandle As Long)
    '�ж�һ��picName�д�ŵ��ǲ���ͼ��

    '����ͼ��ΪpicName�д�ŵ�ͼ��
    With IconData
        .HIcon = hIconHandle
        .UFlags = NIF_ICON
    End With
    Shell_NotifyIcon NIM_MODIFY, IconData
End Sub

Public Sub SetTip(Optional ByVal sTitle As String, Optional ByVal sMsg As String, Optional iMsgType As Long = 1, Optional iSecond As Long = 10)
    sTitle = sTitle & vbNullChar
    sMsg = sMsg & vbNullChar
    
    With IconData
        .cbSize = Len(IconData)
        .Uid = 0
        .UFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE Or NIF_INFO Or NIF_STATE
        .UCallbackMessage = WM_NOTIFYICON
        .SzTip = sTitle
        .dwState = 0
        .dwStateMask = 0
        .szInfo = sMsg
        .szInfoTitle = sTitle
        .dwInfoFlags = iMsgType
        .uTimeoutOrVersion = iSecond * 1000
    End With
    
    Shell_NotifyIcon NIM_MODIFY, IconData

End Sub