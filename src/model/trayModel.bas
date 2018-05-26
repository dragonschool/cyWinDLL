Attribute VB_Name = "modTray"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'�����ȼ�������
Public objMouseWheel As Long

Public preTrayProc As Long

Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Const WM_USER = &H400
Private Const WM_NOTIFYICON = WM_USER + 1            ' �Զ�����Ϣ
Private Const WM_LBUTTONDBLCLK = &H203

' ����������ʾ���Զ�����Ϣ, 2000�²�������Щ��Ϣ
Private Const NIN_BALLOONSHOW = (WM_USER + &H2)         ' �� Balloon Tips ����ʱִ��
Private Const NIN_BALLOONHIDE = (WM_USER + &H3)         ' �� Balloon Tips ��ʧʱִ�У��� SysTrayIcon ��ɾ������
                                                        ' ��ָ���� TimeOut ʱ�䵽������� Balloon Tips �����ʧ�����ʹ���Ϣ
Private Const NIN_BALLOONTIMEOUT = (WM_USER + &H4)      ' �� Balloon Tips �� TimeOut ʱ�䵽ʱִ��
Private Const NIN_BALLOONUSERCLICK = (WM_USER + &H5)    ' ������� Balloon Tips ʱִ�С�
                                                        ' ע��:��XP��ִ��ʱ Balloon Tips ���и��رհ�ť,
                                                        ' ��������ڰ�ť�Ͻ����յ� NIN_BALLOONTIMEOUT ��Ϣ��

Function WndProcTray(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    '���� WM_NOTIFYICON ��Ϣ
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

