Attribute VB_Name = "modMouseWheelEx"
'调用例子
'=======================================================================

Option Explicit

'保存接管之前的值
Public preMouseWheelProc As Long

'保存热键对象句柄
Public objMouseWheel As Long

'一定要放在模块内不能在类模块内，且类型为PUBLIC
Declare Function SetProp Lib "user32" Alias "SetPropA" _
     (ByVal hWnd As Long, ByVal lpString As String, _
     ByVal hData As Long) As Long
Declare Function GetProp Lib "user32" Alias "GetPropA" _
     (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" _
     (ByVal hWnd As Long, ByVal lpString As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
     (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
     ByVal Msg As Long, ByVal wParam As Long, _
     ByVal lParam As Long) As Long

Public Function WndProcMouseWheel(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Const WM_MOUSEWHEEL = &H20A
    If wMsg = WM_MOUSEWHEEL Then
        ptrMouseWheel.FireEvent wParam
    
    End If
    
    WndProcMouseWheel = CallWindowProc(preMouseWheelProc, hWnd, wMsg, wParam, lParam)
    
End Function

Private Function ptrMouseWheel() As mouseWheelClass
    Dim MW As mouseWheelClass
    CopyMemory MW, objMouseWheel, 4&
    Set ptrMouseWheel = MW
    CopyMemory MW, 0&, 4&
  
End Function


