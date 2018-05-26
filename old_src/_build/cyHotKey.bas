Attribute VB_Name = "modHotKey"
Option Explicit

'����ӹ�֮ǰ��ֵ
Public preWinProc As Long

'�����ȼ�������
Public objHotKey As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)

'�����ȼ�ָ������
Private Function ptrHotKey() As cyHotKeyEx
    
    Dim HK As cyHotKeyEx
    CopyMemory HK, objHotKey, 4&
    Set ptrHotKey = HK
    CopyMemory HK, 0&, 4&
    
End Function

Public Function WndProcHotKey(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
     
    '��ȡ�ȼ���Ϣ
    Const WM_HOTKEY = &H312
     
    If uMsg = WM_HOTKEY Then
        '�����ȼ���ID
        ptrHotKey.FireEvent (wParam)
    
    End If
    
    '����������Ϣ
    WndProcHotKey = CallWindowProc(preWinProc, hWnd, uMsg, wParam, lParam)

End Function
