VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cyMouseWheelEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'������¼�
Public Event cyMouseWheelUp()
Public Event cyMouseWheelDown()

'�������MouseWheel�Ĵ��ھ��
Dim m_iMouseWheelHwnd As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Sub RemoveMouseWheel()
    Class_Terminate
    
End Sub

Public Sub SetMouseWheel(ByVal iMouseWheelHwnd As Long)
    '���洰�ھ��
    m_iMouseWheelHwnd = iMouseWheelHwnd
    
    '��ñ�����objprt
    modMouseWheelEx.objMouseWheel = ObjPtr(Me)
    
    preMouseWheelProc = SetWindowLong(iMouseWheelHwnd, -4, AddressOf WndProcMouseWheel)
    
End Sub

Private Sub Class_Terminate()
    '�����¼����ش��ڴ���
     SetWindowLong m_iMouseWheelHwnd, -4, preMouseWheelProc
     
End Sub

Friend Function FireEvent(ByVal i As Long)
    If i > 0 Then
      RaiseEvent cyMouseWheelUp
      
    Else
      RaiseEvent cyMouseWheelDown
      
    End If
    
End Function
