VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cyMonitorFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'����Ŀ¼���ӽ��մ���
Dim m_iMonitorFolderHwnd As Long

'������¼�
Public Event cyFolderChangeEvent(ByVal sEventString As String, ByVal sFile1 As String, ByVal sFile2 As String, ByVal sType1 As String, ByVal sType2 As String)

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
        pSource As Any, ByVal dwLength As Long)

Private Type SHNOTIFYSTRUCT
    dwItem1 As Long
    dwItem2 As Long
End Type

Public Sub cySetMonitorFolder(ByVal hWnd As Long, ByVal sMonitorPath As String, Optional ByVal bWatchSubFolder As Boolean = False)
    '����Ŀ¼���ӽ��մ���
    m_iMonitorFolderHwnd = hWnd
    
    '��ñ�����objprt
    modMonitorFolder.objMonitorFolder = ObjPtr(Me)

    If SubClass(hWnd) Then  '�ı�Form1����Ϣ��������
        Call SHNotify_Register(m_iMonitorFolderHwnd, sMonitorPath, bWatchSubFolder)
        
    End If

End Sub

Private Sub Class_Initialize()
    Dim F As New cyFileEx
    If Not F.cyFileExist(App.Path & "ISHF_Ex.tlb") Then
    
    End If
    
End Sub

Private Sub Class_Terminate()
    Call SHNotify_Unregister
    Call UnSubClass(m_iMonitorFolderHwnd)

End Sub

'�����¼�
Friend Function FireEvent(ByVal wParam As Long, ByVal lParam As Long)

    Static i1 As Long
    Static s1 As String
    Static s2 As String

    Dim SHNS As SHNOTIFYSTRUCT
    Dim sDisplayName1 As String
    Dim sDisplayName2 As String
    
    Const SHCNE_RENAMEITEM = &H1
    Const SHCNE_CREATE = &H2
    Const SHCNE_DELETE = &H4
    Const SHCNE_MKDIR = &H8
    Const SHCNE_RMDIR = &H10
    Const SHCNE_MEDIAINSERTED = &H20
    Const SHCNE_MEDIAREMOVED = &H40
    Const SHCNE_DRIVEREMOVED = &H80
    Const SHCNE_DRIVEADD = &H100
    Const SHCNE_NETSHARE = &H200
    Const SHCNE_UPDATEDIR = &H1000
    Const SHCNE_UPDATEITEM = &H2000
    Const SHCNE_SERVERDISCONNECT = &H4000
    Const SHCNE_UPDATEIMAGE = &H8000&
    Const SHCNE_DRIVEADDGUI = &H10000
    Const SHCNE_RENAMEFOLDER = &H20000
    Const SHCNE_FREESPACE = &H40000
    Const SHCNE_ASSOCCHANGED = &H8000000
    
    MoveMemory SHNS, ByVal wParam, Len(SHNS)
      
    Dim sType1 As String
    Dim sType2 As String
    If SHNS.dwItem1 Then
        sDisplayName1 = GetDisplayNameFromPIDL(SHNS.dwItem1, sType1)
    End If
    
    If SHNS.dwItem2 Then
        sDisplayName2 = GetDisplayNameFromPIDL(SHNS.dwItem2, sType2)
    End If
    
    '�ظ���Ϣ����ʾ
    If lParam = i1 And sDisplayName1 = s1 And sDisplayName2 = s2 Then Exit Function
    
    i1 = lParam
    s1 = sDisplayName1
    s2 = sDisplayName2
    
    Select Case lParam
        Case SHCNE_RENAMEITEM: RaiseEvent cyFolderChangeEvent("�������ļ�", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_CREATE: RaiseEvent cyFolderChangeEvent("�½��ļ�", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_DELETE: RaiseEvent cyFolderChangeEvent("ɾ���ļ�", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_MKDIR: RaiseEvent cyFolderChangeEvent("�½�Ŀ¼", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_RMDIR: RaiseEvent cyFolderChangeEvent("ɾ��Ŀ¼", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_MEDIAINSERTED: RaiseEvent cyFolderChangeEvent("������ƶ��洢����", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_MEDIAREMOVED: RaiseEvent cyFolderChangeEvent("��ȥ���ƶ��洢����", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_DRIVEREMOVED: RaiseEvent cyFolderChangeEvent("��ȥ������", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_DRIVEADD: RaiseEvent cyFolderChangeEvent("����������", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_NETSHARE: RaiseEvent cyFolderChangeEvent("�ı�Ŀ¼��������", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_UPDATEDIR: RaiseEvent cyFolderChangeEvent("����Ŀ¼", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_UPDATEITEM: RaiseEvent cyFolderChangeEvent("�����ļ�", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_SERVERDISCONNECT: RaiseEvent cyFolderChangeEvent("�Ͽ��������������", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_UPDATEIMAGE: RaiseEvent cyFolderChangeEvent("SHCNE_UPDATEIMAGE", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_DRIVEADDGUI: RaiseEvent cyFolderChangeEvent("SHCNE_DRIVEADDGUI", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_RENAMEFOLDER: RaiseEvent cyFolderChangeEvent("�������ļ���", sDisplayName1, sDisplayName2, sType1, sType2)
        'Case SHCNE_FREESPACE: RaiseEvent cyFolderChangeEvent("���̿ռ��С�ı�", sDisplayName1, sDisplayName2, sType1, sType2)
        Case SHCNE_ASSOCCHANGED: RaiseEvent cyFolderChangeEvent("�ı��ļ�����", sDisplayName1, sDisplayName2, sType1, sType2)
        
    End Select
    
End Function