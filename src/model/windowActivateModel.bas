Attribute VB_Name = "modCheckWndActivate"
Option Explicit


'����ӹ�֮ǰ��ֵ
Public preCheckWndActivateProc As Long

'�����ȼ�������
Public objCheckWndActivate As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)

Public Function WndProcCheckWndActivate(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Const WM_ACTIVATE = &H6
    Const WA_ACTIVE = 1
    Const WA_CLICKACTIVE = 2
    
    If Msg = WM_ACTIVATE Then
        If (wParam = WA_ACTIVE Or wParam = WA_CLICKACTIVE) Then
            '�
            ptrCheckWndActivate.FireEvent (True)

        Else
            '�ǻ
            ptrCheckWndActivate.FireEvent (False)
            
        End If
    End If

     WndProcCheckWndActivate = CallWindowProc(preCheckWndActivateProc, hWnd, Msg, wParam, lParam)
     
End Function

Private Function ptrCheckWndActivate() As checkWindowActivityClass
    Dim CheckWndActivate As checkWindowActivityClass
    CopyMemory CheckWndActivate, objCheckWndActivate, 4&
    Set ptrCheckWndActivate = CheckWndActivate
    CopyMemory CheckWndActivate, 0&, 4&
    
End Function


