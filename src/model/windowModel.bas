Attribute VB_Name = "modWnd"
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public PidhWnd As Long
Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    
    Dim Tid As Long, Pid As Long
    If GetParent(hWnd) = 0 Then
        Tid = GetWindowThreadProcessId(hWnd, Pid)
        If Pid = lParam Then
            PidhWnd = hWnd
            EnumWindowsProc = False
            Exit Function '��ʾֹͣ�о� hWnd
            
        End If
        
    End If
    
    EnumWindowsProc = True '��ʾ�����о� hWnd
    
End Function
 



