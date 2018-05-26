VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "checkWindowActivityClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'窗口得到焦点
Public Event cyWndActivate()

'窗口失去焦点
Public Event cyWndDeActivate()

'保存接收MouseWheel的窗口句柄
Dim m_iCheckActivateHwnd As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Sub cySetCheckWndActivate(ByVal iCheckActivateHwnd As Long)
    
    On Error Resume Next
    
    If modCheckWndActivate.preCheckWndActivateProc = 0 Then
        '获得本身的objprt
        modCheckWndActivate.objCheckWndActivate = ObjPtr(Me)
        
        '记录接管前的地址信息
        modCheckWndActivate.preCheckWndActivateProc = SetWindowLong(iCheckActivateHwnd, (-4), AddressOf WndProcCheckWndActivate)
    End If

End Sub

Private Sub Class_Terminate()
    
    '所有事件交回窗口处理
     SetWindowLong m_iCheckActivateHwnd, -4, preCheckWndActivateProc

End Sub

Friend Function FireEvent(ByVal bActivate As Boolean)
    
    If bActivate Then
        RaiseEvent cyWndActivate

    Else
        RaiseEvent cyWndDeActivate

    End If
    
    
End Function


