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

'激活的事件
Public Event cyMouseWheelUp()
Public Event cyMouseWheelDown()

'待捕捉的窗体句柄
Dim m_iMainWnd As Long
Dim m_iCaptureWnd As Long
Dim m_bSubClassed As Boolean

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Sub RemoveMouseWheel()
    Class_Terminate
End Sub

Public Sub SetMouseWheel(ByVal IDMouseWheel As Long, ByVal iMainWndHwnd As Long)
    
On Error Resume Next

''将MouseWheel ID记录到公共数组中
'    Dim i As Long
'    '得到上标
'    i = UBound(IDMouseWheels)
'    If Err.Number = 9 Then
'    '没有则赋0
'        ReDim Preserve IDMouseWheels(0, 1)
'        IDMouseWheels(0, 0) = IDMouseWheel
'        IDMouseWheels(0, 1) = iCaptureWndHwnd
'
'    Else
'        ReDim Preserve IDMouseWheels(i + 1, 1)
'        IDMouseWheels(i + 1, 0) = IDMouseWheel
'        IDMouseWheels(i + 1, 1) = iCaptureWndHwnd
'    End If
    
    If m_iMainWnd = 0 Then
        
        m_iMainWnd = iMainWndHwnd
'        m_iCaptureWnd = iCaptureWndHwnd
'        m_bSubClassed = True
        
        '获得本身的objprt
        modMouseWheel.objMouseWheel = ObjPtr(Me)
        
'        SetProp iMainWndHwnd, "OldWheelProc", GetWindowLong(iCaptureWndHwnd, -4)
        SetProp iMainWndHwnd, "WheelPtr", ObjPtr(Me)
'        SetProp iCaptureWndHwnd, "WheelWnd", iMainWndHwnd
        'SetWindowLong iCaptureWndHwnd, -4, AddressOf WndProcMouseWheel
        preMouseWheelProc = SetWindowLong(m_iMainWnd, -4, AddressOf WndProcMou