Attribute VB_Name = "modMonitorFolder"

'保存接管之前的值
Public preMonitorFolderProc As Long

'保存热键对象句柄
Public objMonitorFolder As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function SHGetDesktopFolder Lib "shell32.dll" (ppshf As Folder) As Long

Private Const WM_NCDESTROY = &H82
Private Const GWL_WNDPROC = (-4)
Private Const OLDWNDPROC = "OldWndProc"

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal _
        hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal _
        hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal _
        hWnd As Long, ByVal lpString As String) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Const MAX_PATH = 260

Private Enum SHSpecialFolderIDs      '列出所有Windows下特殊文件夹的ID
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
End Enum

Private Declare Function SHGetFileInfoPidl Lib "shell32" Alias "SHGetFileInfoA" _
                              (ByVal pidl As Long, _
                              ByVal dwFileAttributes As Long, _
                              psfib As SHFILEINFOBYTE, _
                              ByVal cbFileInfo As Long, _
                              ByVal uFlags As SHGFI_flags) As Long

Private Type SHFILEINFOBYTE
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName(1 To MAX_PATH) As Byte
    szTypeName(1 To 80) As Byte
End Type

Enum SHGFI_flags
    SHGFI_LARGEICON = &H0
    SHGFI_SMALLICON = &H1
    SHGFI_OPENICON = &H2
    SHGFI_SHELLICONSIZE = &H4
    SHGFI_PIDL = &H8
    SHGFI_USEFILEATTRIBUTES = &H10
    SHGFI_ICON = &H100
    SHGFI_DISPLAYNAME = &H200
    SHGFI_TYPENAME = &H400
    SHGFI_ATTRIBUTES = &H800
    SHGFI_ICONLOCATION = &H1000
    SHGFI_EXETYPE = &H2000
    SHGFI_SYSICONINDEX = &H4000
    SHGFI_LINKOVERLAY = &H8000
    SHGFI_SELECTED = &H10000
End Enum

Private m_hSHNotify As Long     '系统消息通告句柄
Private m_PathPIDL As Long      '被监视目录的PIDL

'定义系统通告的消息值
Private Const WM_SHNOTIFY = &H401

Private Type PIDLSTRUCT
    pidl As Long
    bWatchSubFolders As Long
End Type

Private Declare Function SHChangeNotifyRegister Lib "shell32" Alias "#2" _
                              (ByVal hWnd As Long, _
                              ByVal uFlags As SHCN_ItemFlags, _
                              ByVal dwEventID As SHCN_EventIDs, _
                              ByVal uMsg As Long, _
                              ByVal cItems As Long, _
                              lpps As PIDLSTRUCT) As Long

Private Declare Function SHChangeNotifyDeregister Lib "shell32" Alias "#4" _
        (ByVal hNotify As Long) As Boolean

Private Enum SHCN_EventIDs
    SHCNE_RENAMEITEM = &H1
    SHCNE_CREATE = &H2
    SHCNE_DELETE = &H4
    SHCNE_MKDIR = &H8
    SHCNE_RMDIR = &H10
    SHCNE_MEDIAINSERTED = &H20
    SHCNE_MEDIAREMOVED = &H40
    SHCNE_DRIVEREMOVED = &H80
    SHCNE_DRIVEADD = &H100
    SHCNE_NETSHARE = &H200
    SHCNE_NETUNSHARE = &H400
    SHCNE_ATTRIBUTES = &H800
    SHCNE_UPDATEDIR = &H1000
    SHCNE_UPDATEITEM = &H2000
    SHCNE_SERVERDISCONNECT = &H4000
    SHCNE_UPDATEIMAGE = &H8000&
    SHCNE_DRIVEADDGUI = &H10000
    SHCNE_RENAMEFOLDER = &H20000
    SHCNE_FREESPACE = &H40000
    SHCNE_ASSOCCHANGED = &H8000000

    SHCNE_DISKEVENTS = &H2381F
    SHCNE_GLOBALEVENTS = &HC0581E0
    SHCNE_ALLEVENTS = &H7FFFFFFF
    SHCNE_INTERRUPT = &H80000000
End Enum

Private Enum SHCN_ItemFlags
    SHCNF_IDLIST = &H0
    SHCNF_PATHA = &H1
    SHCNF_PRINTERA = &H2
    SHCNF_DWORD = &H3
    SHCNF_PATHW = &H5
    SHCNF_PRINTERW = &H6
    SHCNF_TYPE = &HFF
    SHCNF_FLUSH = &H1000
    SHCNF_FLUSHNOWAIT = &H2000

    #If UNICODE Then
        SHCNF_PATH = SHCNF_PATHW
        SHCNF_PRINTER = SHCNF_PRINTERW
    #Else
        SHCNF_PATH = SHCNF_PATHA
        SHCNF_PRINTER = SHCNF_PRINTERA
    #End If
End Enum

Function SHNotify_Register(ByVal hWnd As Long, ByVal sMonitorPath As String, Optional ByVal bWatchSubFolder As Boolean = False) As Boolean
    Dim PS As PIDLSTRUCT
  
    If (m_hSHNotify = 0) Then
          
        '获得被监视目录的PIDL
        m_PathPIDL = GetPIDLFromPath(sMonitorPath)
        If m_PathPIDL Then
      
            PS.pidl = m_PathPIDL
            PS.bWatchSubFolders = bWatchSubFolder
      
            '注册Windows监视,将获得的句柄保存到m_hSHNotify中
            m_hSHNotify = SHChangeNotifyRegister(hWnd, SHCNF_TYPE Or SHCNF_IDLIST, _
                                            SHCNE_ALLEVENTS Or SHCNE_INTERRUPT, _
                                            WM_SHNOTIFY, 1, PS)
                                            
            SHNotify_Register = CBool(m_hSHNotify)
    
        Else
            Call CoTaskMemFree(m_PathPIDL)
        End If
        
    End If
    
End Function

Function SHNotify_Unregister() As Boolean
    If m_hSHNotify Then
        If SHChangeNotifyDeregister(m_hSHNotify) Then
            m_hSHNotify = 0
            Call CoTaskMemFree(m_PathPIDL)
            m_PathPIDL = 0
            SHNotify_Unregister = True
        End If
        
    End If
    
End Function

Private Sub NotificationReceipt(wParam As Long, lParam As Long)
End Sub

Private Function GetPIDLFromPath(ByVal sPath As String) As Long
    Dim ISF As IShellFolder
    Dim pidlMain     As Long
    Dim cParsed     As Long
    Dim afItem     As Long
    Dim lFilePos     As Long
    Dim lR     As Long
    Dim sRet     As String * 255
      
    lR = GetFullPathName(sPath, MAX_PATH, sRet, lFilePos)
    sPath = Left$(sRet, lR)
    
    '将路径名称转换成PIDL
    Set ISF = GetDesktopFolder
    
    Call ISF.ParseDisplayName(0&, 0&, StrConv(sPath, vbUnicode), cParsed, pidlMain, afItem)
    GetPIDLFromPath = pidlMain
                      
End Function

Private Function GetDesktopFolder() As IShellFolder
    SHGetDesktopFolder GetDesktopFolder
    
End Function
     
Function GetDisplayNameFromPIDL(pidl As Long, sType As String) As String
    Dim sfib As SHFILEINFOBYTE
    If SHGetFileInfoPidl(pidl, 0, sfib, Len(sfib), SHGFI_PIDL Or SHGFI_DISPLAYNAME Or SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES Or SHGFI_ICON Or SHGFI_DISPLAYNAME Or SHGFI_TYPENAME Or SHGFI_ATTRIBUTES Or SHGFI_EXETYPE) Then
        GetDisplayNameFromPIDL = GetStrFromBufferA(StrConv(sfib.szDisplayName, vbUnicode))
        sType = GetStrFromBufferA(StrConv(sfib.szTypeName, vbUnicode))
    
    End If

End Function

Private Function GetStrFromBufferA(sz As String) As String
    If InStr(sz, vbNullChar) Then
        GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
    Else
        GetStrFromBufferA = sz
    End If
    
End Function

Public Function SubClass(hWnd As Long) As Boolean
    Dim lpfnOld As Long
    Dim fSuccess As Boolean
  
    If (GetProp(hWnd, OLDWNDPROC) = 0) Then
        lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc)
        If lpfnOld Then
            fSuccess = SetProp(hWnd, OLDWNDPROC, lpfnOld)
        End If
    End If
  
    If fSuccess Then
        SubClass = True
    Else
        If lpfnOld Then Call UnSubClass(hWnd)
        MsgBox "Unable to successfully subclass &H" & Hex(hWnd), vbCritical
    End If
    
End Function

Public Function UnSubClass(hWnd As Long) As Boolean
    Dim lpfnOld As Long
  
    lpfnOld = GetProp(hWnd, OLDWNDPROC)
    If lpfnOld Then
        If RemoveProp(hWnd, OLDWNDPROC) Then
            UnSubClass = SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld)
        End If
    End If
End Function

Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As _
        Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_SHNOTIFY        '处理系统消息通告函数
            '返回热键的ID
            ptrMonitorFolder.FireEvent wParam, lParam

'            Call NotificationReceipt(wParam, lParam)
        
        Case WM_NCDESTROY
            Call UnSubClass(hWnd)
    
    End Select
    
    WndProc = CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
    
End Function

'处理目录监视指针问题
Private Function ptrMonitorFolder() As cyMonitorFolder
    
    Dim MF As cyMonitorFolder
    CopyMemory MF, objMonitorFolder, 4&
    Set ptrMonitorFolder = MF
    CopyMemory MF, 0&, 4&
    
End Function

