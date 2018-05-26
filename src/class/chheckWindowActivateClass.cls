VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cyCheckWndActivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'���ڵõ�����
Public Event cyWndActivate()
'����ʧȥ����
Public Event cyWndDeActivate()

'�������MouseWheel�Ĵ��ھ��
Dim m_iCheckActivateHwnd As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
   (ByVal hWnd As Long, ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long

Public Sub cySetCheckWndActivate(ByVal iCheckActivateHwnd As Long)
    On Error Resume Next
    
    If modCheckWndActivate.preCheckWndActivateProc = 0 Then
        '��ñ�����objprt
        modCheckWndActivate.objCheckWndActivate = ObjPtr(Me)
        
        '��¼�ӹ�ǰ�ĵ�ַ��Ϣ
        modCheckWndActivate.preCheckWndActivateProc = SetWindowLong(iCheckActivateHwnd, (-4), AddressOf WndProcCheckWndActivate)
    
    End If

End Sub

Private Sub Class_Terminate()
    
    '�����¼����ش��ڴ���
     SetWindowLong m_iCheckActivateHwnd, -4, preCheckWndActivateProc

End Sub

Friend Function FireEvent(ByVal bActivate As Boolean)
    If bActivate Then
        RaiseEvent cyWndActivate

    Else
        RaiseEvent cyWndDeActivate

    End If
    
    
End Function

