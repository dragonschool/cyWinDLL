Attribute VB_Name = "modMouseWheelEx"
'��������
'=======================================================================

Option Explicit

'����ӹ�֮ǰ��ֵ
Public preMouseWheelProc As Long

'�����ȼ�������
Public objMouseWheel As Long

'һ��Ҫ����ģ���ڲ�������ģ���ڣ�������ΪPUBLIC
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


