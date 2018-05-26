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
'    cbSize As Long               '需填入NOTIFYICONDATA数据结构的长度
'    HWnd As Long                 '设置成窗口的句柄
'    Uid As Long                  '为图标所设置的ID值
'    UFlags As Long               '设置uCallbackMessage,hIcon,szTip是否有效
'    UCallbackMessage As Long     '消息编号
'    HIcon As Long                '显示在状态栏上的图标
'    SzTip As String * 64         '提示信息
'
'End Type


Private Type NOTIFYICONDATA
    cbSize As Long              ' 结构大小(字节)
    hWnd As Long                ' 处理消息的窗口的句柄
    Uid As Long                 ' 唯一的标识符
    UFlags As Long              ' Flags
    UCallbackMessage As Long    ' 处理消息的窗口接收的消息
    HIcon As Long               ' 托盘图标句柄
    SzTip As String * 128       ' Tooltip 提示文本
    dwState As Long             ' 托盘图标状态
    dwStateMask As Long         ' 状态掩码
    szInfo As String * 256      ' 气球提示文本
    uTimeoutOrVersion As Long   ' 气球提示消失时间或版本
                                ' uTimeout - 气球提示消失时间(单位:ms, 10000 -- 30000)
                                ' uVersion - 版本(0 for V4, 3 for V5)
    szInfoTitle As String * 64  ' 气球提示标题
    dwInfoFlags As Long         ' 气球提示图标
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
Private Const WM_NOTIFYICON = WM_USER + 1            ' 自定义消息
Private Const WM_LBUTTONDBLCLK = &H203
Private Const GWL_WNDPROC = (-4)

' 关于气球提示的自定义消息, 2000下不产生这些消息
Private Const NIN_BALLOONSHOW = (WM_USER + &H2)         ' 当 Balloon Tips 弹出时执行
Private Const NIN_BALLOONHIDE = (WM_USER + &H3)         ' 当 Balloon Tips 消失时执行（如 SysTrayIcon 被删除），
                                                        ' 但指定的 TimeOut 时间到或鼠标点击 Balloon Tips 后的消失不发送此消息
Private Const NIN_BALLOONTIMEOUT = (WM_USER + &H4)      ' 当 Balloon Tips 的 TimeOut 时间到时执行
Private Const NIN_BALLOONUSERCLICK = (WM_USER + &H5)    ' 当鼠标点击 Balloon Tips 时执行。
Private Const TRAY_CALLBACK = (WM_USER + 1001&)                                                         ' 注意:在XP下执行时 Balloon Tips 上有个关闭按钮,
                                                        ' 如果鼠标点在按钮上将接收到 NIN_BALLOONTIMEOUT 消息。
'向托盘区添加图标
Dim IconData As NOTIFYICONDATA

Dim m_iTrayHwnd As Long
Dim m_iIconHandle As Long

Public Sub RemoveTray()
    Class_Terminate
    
End Sub

Public Sub SetTray(Optional ByVal iTrayHwnd As Long, Optional ByVal IconHandle As Long, Optional ByVal sTitle As String, Optional ByVal sMsg As String, Optional iMsgType As Long = 1, Optional iSecond As Long = 10)
    
    '保存起窗口以备关闭
    m_iTrayHwnd = iTrayHwnd
    m_iIconHandle = IconHandle
    
    sTitle = sTitle & vbNullChar
    sMsg = sMsg & vbNullChar
    
    modTray.objMouseWheel = ObjPtr(Me)
        
'        With IconData
'            .Uid = 0
'            .HWnd = iTrayHwnd            'frm.HWnd是程序主窗体的句柄
'            .cbSize = Len(IconData)
'            .HIcon = IconHandle    'frm.Icon.Handle指向主窗体的图标
'            .UFlags = NIF_ICON
'            .UFlags = .UFlags Or NIF_MESSAGE
'            .UCallbackMessage = TRAY_CALLBACK
'             '作用是允许返回消息，在下一节中会有详细解释。
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
    
    If preTrayProc = 0 Then    '没有初始化
       Shell_NotifyIcon NIM_ADD, IconData
       preTrayProc = SetWindowLong(iTrayHwnd, GWL_WNDPROC, AddressOf WndProcTray)
    Else                      '已初始化
       Shell_NotifyIcon NIM_MODIFY, IconData
    End If
    
End Sub

Friend Function FireEvent(ByVal iAction As Long)
    'Const WM_LBUTTONUP = &H202
    'Const WM_RBUTTONUP = &H205
    Const WM_USER = &H400
    Const WM_LBUTTONDBLCLK = &H203
    
    '关于气球提示的自定义消息, 2000下不产生这些消息
    Const NIN_BALLOONSHOW = (WM_USER + &H2)         ' 当 Balloon Tips 弹出时执行
    Const NIN_BALLOONHIDE = (WM_USER + &H3)         ' 当 Balloon Tips 消失时执行（如 SysTrayIcon 被删除），
                                                            ' 但指定的 TimeOut 时间到或鼠标点击 Balloon Tips 后的消失不发送此消息
    Const NIN_BALLOONTIMEOUT = (WM_USER + &H4)      ' 当 Balloon Tips 的 TimeOut 时间到时执行
    Const NIN_BALLOONUSERCLICK = (WM_USER + &H5)    ' 当鼠标点击 Balloon Tips 时执行。

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

    ' 删除托盘区图标
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
' 设置托盘的图标
' *********************************************
Public Sub SetIcon(ByVal hIconHandle As Long)
    '判断一下picName中存放的是不是图标

    '更换图标为picName中存放的图标
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
