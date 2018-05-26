VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cyHotKeyEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'������¼�
Public Event cyHotKeyEventEx(ByVal IDHotKey As Long)

'�ȼ���������
Dim IDHotKeys() As Long

'�����ȼ����մ���
Dim m_iHotKeyHwnd As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
   (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
   (ByVal hWnd As Long, ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long

Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal VK As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long

'������е��ȼ�
Private Sub Class_Terminate()

    Dim i As Long
    
    '��Ȩ������ϵͳ
    SetWindowLong m_iHotKeyHwnd, -4, preWinProc
    preWinProc = 0
    
    '����ȼ��б�
    For i = 0 To UBound(IDHotKeys)
    
        '��ע�����ȼ�
        Call UnregisterHotKey(m_iHotKeyHwnd, IDHotKeys(i))
        
    Next
    
    'ȥ�����ȼ���ID
    ReDim IDHotKeys(0)
    
End Sub

'��ñ�����objprt
Friend Function GetObjptr() As Long
    GetObjptr = ObjPtr(Me)

End Function

'�����¼�
Friend Function FireEvent(ByVal IDHotKey As Long)
    RaiseEvent cyHotKeyEventEx(IDHotKey)
    
End Function

'�����ȼ�
Sub cySetHotKeyEx(ByVal IDHotKey As Long, ByVal iHotKeyHwnd As Long, Optional ByVal bControlKey As Boolean, Optional ByVal bShiftKey As Boolean, Optional ByVal bAltKey As Boolean, Optional ByVal iVirualKey As Long)
    
On Error Resume Next

'���ȼ�ID��¼�����������У���������ʱҪ�ر�
    Dim i As Long
    '�õ��ϱ�
    i = UBound(IDHotKeys)
    If Err.Number = 9 Then
    'û����0
        ReDim Preserve IDHotKeys(0)
        IDHotKeys(0) = IDHotKey
    Else
        ReDim Preserve IDHotKeys(i + 1)
        IDHotKeys(i + 1) = IDHotKey
    End If
    
    Const MOD_ALT = &H1
    Const MOD_CONTROL = &H2
    Const MOD_SHIFT = &H4
    
    If m_iHotKeyHwnd = 0 Then
    
        '��ñ�����objprt
        modHotKey.objHotKey = ObjPtr(Me)

        '��¼��ݼ����ھ��
        m_iHotKeyHwnd = iHotKeyHwnd
        
        '��¼�ӹ�ǰ�ĵ�ַ��Ϣ
        preWinProc = GetWindowLong(m_iHotKeyHwnd, -4)
        
        '�ӹܸù���
        SetWindowLong m_iHotKeyHwnd, -4, AddressOf WndProcHotKey
        
    End If

    '��ȡ���ܼ���ֵ
    Dim iFunction As Long
    If bControlKey Then iFunction = MOD_CONTROL
    If bShiftKey Then iFunction = iFunction + MOD_SHIFT
    If bAltKey Then iFunction = iFunction + MOD_ALT
    
    '���ÿ�ݼ�
    RegisterHotKey m_iHotKeyHwnd, IDHotKey, iFunction, iVirualKey
    
End Sub

'���ȼ��б���ע�����ȼ�
Sub cyRemoveHotKeyEx(ByVal IDHotKey As Long)

    Dim i As Long
    '����ȼ��б�
    For i = 0 To UBound(IDHotKeys)
    
        '����鵽���ȼ�ID
        If IDHotKeys(i) = IDHotKey Then
            '��ע�����ȼ�
            Call UnregisterHotKey(m_iHotKeyHwnd, IDHotKeys(i))
            
            'ȥ�����ȼ���ID
            IDHotKeys(i) = 0
            
        End If
        
    Next
    
End Sub

'��ͣ�ȼ�
Sub cyPauseHotKeyEx()

On Error Resume Next
    '��ʱ��Ȩ������ϵͳ
    SetWindowLong m_iHotKeyHwnd, -4, preWinProc

End Sub

'�����ȼ�
Sub cyRestartHotKeyEx()

On Error Resume Next

    '��¼�ӹ�ǰ�ĵ�ַ��Ϣ
    preWinProc = GetWindowLong(m_iHotKeyHwnd, -4)
    
    '�ӹܸù���
    SetWindowLong m_iHotKeyHwnd, -4, AddressOf WndProcHotKey

End Sub
