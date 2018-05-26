VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hookClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'激活的事件
Public Event cyMouseDown(Button As Integer, Shift As Integer, X As Integer, Y As Integer)
Public Event cyMouseUp(Button As Integer, Shift As Integer, X As Integer, Y As Integer)
Public Event cyMouseMove(Button As Integer, Shift As Integer, X As Integer, Y As Integer)
Public Event cyKeyDown(KeyCode As Integer, Shift As Integer)
Public Event cyKeyUp(KeyCode As Integer, Shift As Integer)

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetAsyncKeyState% Lib "user32" (ByVal vKey As Long)

Public Enum Events
    KeyDown = &H100
    KeyUp = &H101
    MouseMove = &H200
    LBUTTONDOWN = &H201
    LBUTTONUP = &H202
    RBUTTONDOWN = &H204
    RBUTTONUP = &H205
    MBUTTONDOWN = &H207
    MBUTTONUP = &H208
End Enum

Public Sub cySetHookEx(ByVal Events As Events)
    On Error Resume Next
    Dim i As Long
    Const WH_JOURNALRECORD = 0
    
    '得到上标
    i = UBound(modHook.HookActions)
    If Err.Number = 9 Then
    '没有则赋0
        ReDim Preserve modHook.HookActions(0)
        modHook.HookActions(0) = Events
    Else
        ReDim Preserve modHook.HookActions(i + 1)
        modHook.HookActions(i + 1) = Events
    End If
    
    If modHook.preHookProc = 0 Then
        '获得本身的objprt
        modHook.objHook = ObjPtr(Me)
        
        '记录接管前的地址信息
        modHook.preHookProc = SetWindowsHookEx(WH_JOURNALRECORD, AddressOf HookProc, App.hInstance, 0)
    
    End If

End Sub

Public Sub cyRestartHookEx()
    Const WH_JOURNALRECORD = 0
    
    '记录接管前的地址信息
    modHook.preHookProc = SetWindowsHookEx(WH_JOURNALRECORD, AddressOf HookProc, App.hInstance, 0)

End Sub
Public Sub cyPauseHookEx()
    
    UnhookWindowsHookEx modHook.preHookProc

End Sub
Public Sub cyRemoveHookEx(ByVal Events As Events)
    Dim i As Long
    
    '检查热键列表
    For i = 0 To UBound(modHook.HookActions)
    
        '如果查到该HOOK
        If modHook.HookActions(i) = Events Then
        
            '去掉该HOOK的值
             modHook.HookActions(i) = -1
            
        End If
        
    Next
    
End Sub


Private Sub Class_Terminate()
    
    UnhookWindowsHookEx modHook.preHookProc
    '清空HOOK列表
    ReDim modHook.HookActions(0)
    
    modHook.preHookProc = 0

End Sub

Friend Function FireEvent(ByVal iMsg As Long, ByVal iMsgLow As Long, ByVal iMsgHigh As Long, ByVal lParam As Long)

    Dim i%, j%, k%

    Select Case iMsg
      Case &H100
           j = 0
           If GetAsyncKeyState(vbKeyShift) Then j = (j Or 1)
           If GetAsyncKeyState(vbKeyControl) Then j = (j Or 2)
           If GetAsyncKeyState(vbKeyMenu) Then j = (j Or 4)
           k = (iMsgLow And &HFF)
           RaiseEvent cyKeyDown(k, j)
           
      Case &H101
           j = 0
           If GetAsyncKeyState(vbKeyShift) Then j = (j Or 1)
           If GetAsyncKeyState(vbKeyControl) Then j = (j Or 2)
           If GetAsyncKeyState(vbKeyMenu) Then j = (j Or 4)
           k = (iMsgLow And &HFF)
           RaiseEvent cyKeyUp(k, j)
           
      Case &H200
           i = 0
           If GetAsyncKeyState(vbKeyLButton) Then i = (i Or 1)
           If GetAsyncKeyState(vbKeyRButton) Then i = (i Or 2)
           If GetAsyncKeyState(vbKeyMButton) Then i = (i Or 4)
           j = 0
           If GetAsyncKeyState(vbKeyShift) Then j = (j Or 1)
           If GetAsyncKeyState(vbKeyControl) Then j = (j Or 2)
           If GetAsyncKeyState(vbKeyMenu) Then j = (j Or 4)
           RaiseEvent cyMouseMove(i, j, CSng(iMsgLow), CSng(iMsgHigh))
           
      Case &H201, &H204, &H207
           i = 0
           If GetAsyncKeyState(vbKeyShift) Then i = (i Or 1)
           If GetAsyncKeyState(vbKeyControl) Then i = (i Or 2)
           If GetAsyncKeyState(vbKeyMenu) Then i = (i Or 4)
           RaiseEvent cyMouseDown(2 ^ ((iMsg - 513) / 3), i, CSng(iMsgLow), CSng(iMsgHigh))
           
      Case &H202, &H205, &H208
           i = 0
           If GetAsyncKeyState(vbKeyShift) Then i = (i Or 1)
           If GetAsyncKeyState(vbKeyControl) Then i = (i Or 2)
           If GetAsyncKeyState(vbKeyMenu) Then i = (i Or 4)
           RaiseEvent cyMouseUp(2 ^ ((iMsg - 514) / 3), i, CSng(iMsgLow), CSng(iMsgHigh))
           
      Case Else
      
    End Select
    
End Function


