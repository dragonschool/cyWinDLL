Attribute VB_Name = "modTray"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'保存热键对象句柄
Public objMouseWheel As Long

Public preTrayProc As Long

Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Const WM_USER = &H400
Private Const WM_NOTIFYICON = WM_USER + 1            ' 自定义消息
Private Const WM_LBUTTONDBLCLK = &H203

' 关于气球提示的自定义消息, 2000下不产生这些消息
Private Const NIN_BALLOONSHOW = (WM_USER + &H2)         ' 当 Balloon Tips 弹出时执行
Private Const NIN_BALLOONHIDE = (WM_USER + &H3)         ' 当 Balloon Tips 消失时执行（如 SysTrayIcon 被删除），
                                                        ' 但指定的 TimeOut 时间到或鼠标点击 Balloon Tips 后的消失不发送此消息
Private Const NIN_BALLOONTIMEOUT = (WM_USER + &H4)      ' 当 Balloon Tips 的 TimeOut 时间到时执行
Private Const NIN_BALLOONUSERCLICK = (WM_USER + &H5)    ' 当鼠标点击 Balloon Tips 时执行。
                                                        ' 注意:在XP下执行时 Balloon Tips 上有个关闭按钮,
                                                        ' 如果鼠标点在按钮上将接收到 NIN_BALLOONTIMEOUT 消息。

Function WndProcTray(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    '拦截 WM_NOTIFYICON 消息
    If Msg = WM_NOTIFYICON Then
        Select Case lParam
            Case WM_LBUTTONUP
                ptrTray.FireEvent lParam
    
            Case WM_RBUTTONUP
                ptrTray.FireEvent lParam
    
            Case WM_LBUTTONDBLCLK
                ptrTray.FireEvent lParam
    
            Case NIN_BALLOONSHOW
                ptrTray.FireEvent lParam
    
            Case NIN_BALLOONHIDE
                ptrTray.FireEvent lParam
    
            Case NIN_BALLOONTIMEOUT
                ptrTray.FireEvent lParam
    
            Case NIN_BALLOONUSERCLICK
                ptrTray.FireEvent lParam
    
        End Select
    
    End If
    
    WndProcTray = CallWindowProc(preTrayProc, hWnd, Msg, wParam, lParam)
    
End Function

Private Function ptrTray() As cyTrayEx
    Dim Tray As cyTrayEx
    CopyMemory Tray, objMouseWheel, 4&
    Set ptrTray = Tray
    CopyMemory Tray, 0&, 4&
    
End Function

