Attribute VB_Name = "modHotKey"
Option Explicit

'保存接管之前的值
Public preWinProc As Long

'保存热键对象句柄
Public objHotKey As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)

'处理热键指针问题
Private Function ptrHotKey() As cyHotKeyEx
    
    Dim HK As cyHotKeyEx
    CopyMemory HK, objHotKey, 4&
    Set ptrHotKey = HK
    CopyMemory HK, 0&, 4&
    
End Function

Public Function WndProcHotKey(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
     
    '钩取热键消息
    Const WM_HOTKEY = &H312
     
    If uMsg = WM_HOTKEY Then
        '返回热键的ID
        ptrHotKey.FireEvent (wParam)
    
    End If
    
    '继续接收消息
    WndProcHotKey = CallWindowProc(preWinProc, hWnd, uMsg, wParam, lParam)

End Function
