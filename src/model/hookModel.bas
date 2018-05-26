Attribute VB_Name = "modHook"
Option Explicit

'����ӹ�֮ǰ��ֵ
Public preHookProc As Long

'�����ȼ�������
Public objHook As Long

'����ҪHOOK�Ķ���(����Ҫ��ģ���м�⣬���Դ����ģ����)
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


