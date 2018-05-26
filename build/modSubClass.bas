Attribute VB_Name = "modSubClass"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'保存热键对象句柄
Public objSubClass As Long
Public preSubClassProc As Long
Public WinMsg1 As Long
Public WinMsg2 As Long
Public WinMsg3 As Long

'仅接收指定消息
Public Function WndProcSubClass(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next

If uMsg = WinMsg1 Or uMsg = WinMsg2 Or uMsg = WinMsg3 Then
    ptrSubClass.FireEvent uMsg, wParam, lParam

End If

'If uMsg = WinMsg1 Or uMsg = WinMsg2 Or uMsg = WinMsg3 Then
'    ptrSubClass.FireEvent uMsg, wParam, lParam
'
'Else
'    ptrSubClass.FireEvent uMsg, wParam, lParam
'
'End If

WndProcSubClass = CallWindowProc(preSubClassProc, hWnd, uMsg, wParam, lParam)
    
End Function

'接收所有消息
Public Function WndProcSubClassAllMsg(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next

    ptrSubClass.FireEvent uMsg, wParam, lParam

'If uMsg = WinMsg1 Or uMsg = WinMsg2 Or uMsg = WinMsg3 Then
'    ptrSubClass.FireEvent uMsg, wParam, lParam
'
'Else
'    ptrSubClass.FireEvent uMsg, wParam, lParam
'
'End If

WndProcSubClassAllMsg = CallWindowProc(preSubClassProc, hWnd, uMsg, wParam, lParam)
    
End Function

Private Function ptrSubClass() As cySubClassEx
    Dim SC As cySubClassEx
    CopyMemory SC, objSubClass, 4&
    Set ptrSubClass = SC
    CopyMemory SC, 0&, 4&
  
End Function
