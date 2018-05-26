Attribute VB_Name = "modHook"
Option Explicit

'保存接管之前的值
Public preHookProc As Long

'保存热键对象句柄
Public objHook As Long

'保存要HOOK的动作(由于要在模块中检测，所以存放于模块中)
Public HookActions() As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public hJournalHook As Long, hAppHook As Long
Public SHptr As Long
Public Const WM_CANCELJOURNAL = &H4B

Private Type EVENTMSG
     wMsg As Long
     lParamLow As Long
     lParamHigh As Long
     msgTime As Long
     hWndMsg As Long
End Type

Public Function HookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim eMsg As EVENTMSG
    CopyMemory eMsg, ByVal lParam, Len(eMsg)
    
    Dim i As Long
    For i = 0 To UBound(HookActions)
        If HookActions(i) = eMsg.wMsg Then
            ptrHook.FireEvent eMsg.wMsg, eMsg.lParamLow, eMsg.lParamHigh, lParam
        End If
    Next
    
    Call CallNextHookEx(preHookProc, nCode, wParam, lParam)
End Function

Private Function ptrHook() As cyHookEx
    Dim Hook As cyHookEx
    CopyMemory Hook, objHook, 4&
    Set ptrHook = Hook
    CopyMemory Hook, 0&, 4&
    
End Function


