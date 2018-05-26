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

'激活的事件
Public Event cyHotKeyEventEx(ByVal IDHotKey As Long)

'热键队列数组
Dim IDHotKeys() As Long

'保存热键接收窗口
Dim m_iHotKeyHwnd As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
   (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
   (ByVal hWnd As Long, ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long

Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal VK As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long

'清除所有的热键
Private Sub Class_Terminate()

    Dim i As Long
    
    '将权力交回系统
    SetWindowLong m_iHotKeyHwnd, -4, preWinProc
    preWinProc = 0
    
    '检查热键列表
    For i = 0 To UBound(IDHotKeys)
    
        '则注销该热键
        Call UnregisterHotKey(m_iHotKeyHwnd, IDHotKeys(i))
        
    Next
    
    '去掉该热键的ID
    ReDim IDHotKeys(0)
    
End Sub

'获得本身的objprt
Friend Function GetObjptr() As Long
    GetObjptr = ObjPtr(Me)

End Function

'激活事件
Friend Function FireEvent(ByVal IDHotKey As Long)
    RaiseEvent cyHotKeyEventEx(IDHotKey)
    
End Function

'激活热键
Sub cySetHotKeyEx(ByVal IDHotKey As Long, ByVal iHotKeyHwnd As Long, Optional ByVal bControlKey As Boolean, Optional ByVal bShiftKey As Boolean, Optional ByVal bAltKey As Boolean, Optional ByVal iVirualKey As Long)
    
On Error Resume Next

'将热键ID记录到公共数组中，当销毁类时要关闭
    Dim i As Long
    '得到上标
    i = UBound(IDHotKeys)
    If Err.Number = 9 Then
    '没有则赋0
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
    
        '获得本身的objprt
        modHotKey.objHotKey = ObjPtr(Me)

        '记录快捷键窗口句柄
        m_iHotKeyHwnd = iHotKeyHwnd
        
        '记录接管前的地址信息
        preWinProc = GetWindowLong(m_iHotKeyHwnd, -4)
        
        '接管该过程
        SetWindowLong m_iHotKeyHwnd, -4, AddressOf WndProcHotKey
        
    End If

    '读取功能键的值
    Dim iFunction As Long
    If bControlKey Then iFunction = MOD_CONTROL
    If bShiftKey Then iFunction = iFunction + MOD_SHIFT
    If bAltKey Then iFunction = iFunction + MOD_ALT
    
    '设置快捷键
    RegisterHotKey m_iHotKeyHwnd, IDHotKey, iFunction, iVirualKey
    
End Sub

'从热键列表中注销该热键
Sub cyRemoveHotKeyEx(ByVal IDHotKey As Long)

    Dim i As Long
    '检查热键列表
    For i = 0 To UBound(IDHotKeys)
    
        '如果查到该热键ID
        If IDHotKeys(i) = IDHotKey Then
            '则注销该热键
            Call UnregisterHotKey(m_iHotKeyHwnd, IDHotKeys(i))
            
            '去掉该热键的ID
            IDHotKeys(i) = 0
            
        End If
        
    Next
    
End Sub

'暂停热键
Sub cyPauseHotKeyEx()

On Error Resume Next
    '暂时将权力交回系统
    SetWindowLong m_iHotKeyHwnd, -4, preWinProc

End Sub

'重启热键
Sub cyRestartHotKeyEx()

On Error Resume Next

    '记录接管前的地址信息
    preWinProc = GetWindowLong(m_iHotKeyHwnd, -4)
    
    '接管该过程
    SetWindowLong m_iHotKeyHwnd, -4, AddressOf WndProcHotKey

End Sub

