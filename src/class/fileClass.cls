VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "fileClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'将字节转换为MB/KB等形式
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias _
        "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef _
        cchBuf As Long) As String

Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

'-----------------------------------------------------------------------------
'重设文件时间
'-----------------------------------------------------------------------------
Private Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime   As Long
End Type

Private Type SYSTEMTIME
    wYear   As Integer
    wMonth   As Integer
    wDayOfWeek   As Integer
    wDay   As Integer
    wHour   As Integer
    wMinute   As Integer
    wSecond   As Integer
    wMilliseconds   As Integer
End Type

Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
'-----------------------------------------------------------------------------
'重设文件时间
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'获得文件版本，及对比版本新旧
'-----------------------------------------------------------------------------
Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

'专用于获得文件版号，因与其它MOVEMEMORY函数相撞，因此将此函数后＋1，对应
Private Declare Sub MoveMemory1 Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
'Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)


'-----------------------------------------------------------------------------
'获得文件版本，及对比版本新旧
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'从剪贴板中复制文件使用
'-----------------------------------------------------------------------------

'为真时停止复制
Public bStopFilePaste As Boolean

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "SHELL32.DLL" Alias _
        "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

'剪贴版处理函数
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd _
        As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat _
        As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat _
        As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" _
        (ByVal wFormat As Long) As Long

Private Declare Function DragQueryFile Lib "SHELL32.DLL" Alias _
        "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, _
        ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags _
        As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As _
        Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As _
        Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)

'剪贴版数据格式定义
Private Const CF_HDROP = 15

' 内存操作定义
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Private Const FO_COPY = &H2

Private Type DROPFILES
   pFiles As Long
   Pt As POINTAPI
   fNC As Long
   fWide As Long
End Type
'-----------------------------------------------------------------------------
'从剪贴板中复制文件使用
'-----------------------------------------------------------------------------


Public cyDialogSelectFolderInitDir As String

Dim iFolderCounter As Long

Dim bStopSearch As Boolean          '是否结束查找的标志
Public bSearchStoped As Boolean     '搜索是否已结束

Public Event cyFilePasting(ByVal sFileName As String, ByVal iFileCount As Long, ByVal iPastedPos As Long)
Public Event cyFoundFile(sPath As String, sFileName As String)
Public Event cyFoundStoped()

'查找第一个文件的API
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'查找下一个文件的API
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'获取文件属性的API
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
'关闭查找文件的API
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

'Private Type FILETIME
'    dwLowDateTime As Long
'    dwHighDateTime As Long
'End Type

'定义类(用于查找文件)
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

'-----------------------------------------------------------------------------
'文件操作函数
'-----------------------------------------------------------------------------
'    Private Type SHFILEOPSTRUCT
'         hwnd As Long
'         wFunc As Long
'         pFrom As String
'         pTo As String
'         fFlags As Integer
'         fAnyOperationsAborted As Long
'         hNameMappings As Long
'         lpszProgressTitle As String '只有在 FOF_SIMPLEPROGRESS 时用
'     End Type
     
'    Private Declare Function SHFileOperation Lib _
'     "shell32.dll" Alias "SHFileOperationA" (lpFileOp _
'     As SHFILEOPSTRUCT) As Long
    
'-----------------------------------------------------------------------------
'文件操作函数
'-----------------------------------------------------------------------------



Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

'###########################################取得特殊目录
Private Type SHITEMID
  cb As Long
  abID As Byte
End Type

Private Type ITEMIDLIST
  mkid As SHITEMID
End Type

Enum SystemFolders
  
  ' Standard Windows 95 SDK Documented Folder Flags
  '   1 = Hidden Folder
  '   2 = Internet Explorer required
  '   3 = Virtual Folder
  '   4 = Windows NT and Windows 9x
  '   5 = Windows NT Only
  '   6 = Windows 2000 Only
  '   7 = Alpha Systems Only
  Recycle = 10                   ' 1 / Recycle Bin
  ControlsPanel = 3              ' 3 / Control Panel
  DeskTop = 0                    ' 4 / Windows Desktop
  DeskTopDirectory = 16          ' 5 / Directory used to physically store file objects on the desktop (not to be confused with the desktop folder itself).
  MyComputer = 17                ' 3 / My Computer
  Fonts = 20                     ' 3 / Virtual folder containing fonts.
  NetHood = 19                   ' 5 / Directory containing objects that appear in the network neighborhood.
  NetWork = 18                   ' 3 / Network Neighborhood
  Personal = 5                   ' 4 / Directory that serves as a common respository for documents.
  Printers = 4                   ' 3 / Printers folder
  Programs = 2                   ' 5 / Directory that contains the user抯 program groups
  Recent = 8                     ' 4 / Directory that contains the user抯 most recently used documents.
  Sendto = 9                      ' 4 / Directory that contains Send To menu items.
  StartMenu = 11                 ' 5 / Directory containing Start menu items.
  StartUp = 7                    ' 5 / Directory that corresponds to the user抯 Startup program group.
  Templates = 21                 ' 4 / Directory that serves as a common repository for document templates.
    
  
  ' Folder Flags Not Documented In The Windows 95 SDK (Found info on Microsoft's web site)
  UserStartupGroup = 29          ' 2 / User's Startup Group
  ApplicationData = 26           ' 5 / Application data directory for the current user
  'COMMON_StartUp = 30           ' 2 / Common Startup Group
  COMMON_DeskTopDirectory = 25   ' 4 / Desktop Directory - Common
  COMMON_Favorites = 31          ' 5 / Directory that serves as a common repository for all users' favorite items
  COMMON_Programs = 23           ' 4 / Program Menu - All Users
  COMMON_StartMenu = 22          ' 4 / Start Menu - Common
  COMMON_StartUp = 24            ' 4 / Start Up Folder - Common
  Cookies = 33                   ' 2 / Cookies directory for the current user
  Favorites = 6                  ' 4 / Windows Favorites directory
  History = 34                   ' 2 / History directory for the current user
  Internet = 1                   ' 2 / Virtual folder representing the Internet.
  Internet_Cache = 32             ' 2 / Internet Explorer cache directory for the current user
  PrintHood = 27                 ' 5 / Directory containing objects that appear in the Printer neighborhood.
  
  ' Folder flags that are only applicable for Windows 2000, or on RISC systems.
  AdminTools = &H30               ' 6 / Directory that is used to store administrative tools for an individual user
  COMMON_AdminTools = &H2F        ' 6 / Directory containing containing administrative tools for all users of the computer.
  MyPictures = &H27               ' 6 / My Pictures folder
  Program_Files = &H26            ' 6 / Program Files folder.  A typical path is C:\Program Files.
  COMMON_Program_Files = &H2B     ' 6 / A folder for components that are shared across applications.  A typical path is C:\Program Files\Common.
  System = &H25                  ' 6 / System folder. A typical path is C:\WINNT\SYSTEM32.
  Windows = &H24                 ' 6 / Windows directory or SYSROOT. This corresponds to the %windir% or %SYSTEMROOT% environment variables.
  TempPath = &H40
End Enum

Private Declare Function SHGetSpecialFolderLocation Lib "SHELL32.DLL" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "SHELL32.DLL" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'###########################################取得特殊目录



'目标监视使用－－－－－－－－－－
Const INFINITE = &HFFFF
Const FILE_NOTIFY_CHANGE_FILE_NAME As Long = &H1
Const FILE_NOTIFY_CHANGE_DIR_NAME As Long = &H2
Const FILE_NOTIFY_CHANGE_ATTRIBUTES As Long = &H4
Const FILE_NOTIFY_CHANGE_SIZE As Long = &H8
Const FILE_NOTIFY_CHANGE_LAST_WRITE As Long = &H10
Const FILE_NOTIFY_CHANGE_LAST_ACCESS As Long = &H20
Const FILE_NOTIFY_CHANGE_CREATION As Long = &H40
Const FILE_NOTIFY_CHANGE_SECURITY As Long = &H100
Const FILE_NOTIFY_FLAGS = FILE_NOTIFY_CHANGE_ATTRIBUTES Or _
                                 FILE_NOTIFY_CHANGE_FILE_NAME Or _
                                 FILE_NOTIFY_CHANGE_LAST_WRITE

Private Declare Function FindFirstChangeNotification Lib "kernel32" _
    Alias "FindFirstChangeNotificationA" _
   (ByVal lpPathName As String, _
    ByVal bWatchSubtree As Long, _
    ByVal dwNotifyFilter As Long) As Long

Private Declare Function FindCloseChangeNotification Lib "kernel32" _
   (ByVal hChangeHandle As Long) As Long

Private Declare Function FindNextChangeNotification Lib "kernel32" _
   (ByVal hChangeHandle As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Const WAIT_OBJECT_0 = &H0
Const WAIT_ABANDONED = &H80
Const WAIT_IO_COMPLETION = &HC0
Const WAIT_TIMEOUT = &H102
Const STATUS_PENDING = &H103
Dim hWatch As Long
'目标监视使用－－－－－－－－－－


















Dim m_wCurOptIdx As Integer

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXPLORER = &H80000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Declare Function LocalAlloc Lib "kernel32" _
     (ByVal UFlags As Long, _
     ByVal uBytes As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" _
    Alias "RtlMoveMemory" _
     (pDest As Any, _
     pSource As Any, _
     ByVal dwLength As Long)





''''Private Type SHITEMID   ' mkid
''''    cb As Long       ' Size of the ID (including cb itself)
''''    abID() As Byte  ' The item ID (variable length)
''''End Type
''''
''''' The ITEMIDLIST structure defines an element in an item identifier list
''''' (the only member of this structure is an SHITEMID structure). An item
''''' identifier list consists of one or more consecutive ITEMIDLIST structures
''''' packed on byte boundaries, followed by a 16-bit zero value. An application
''''' can walk a list of item identifiers by examining the size specified in each
''''' SHITEMID structure and stopping when it finds a size of zero. A pointer
''''' to an item identifier list, is sometimes called a PIDL (pronounced piddle)
''''Private Type ITEMIDLIST   ' idl
''''    mkID As SHITEMID
''''End Type
''''
''''
''''' SHGetSpecialFolderLocation successful rtn val
''''Const NOERROR = 0
''''
''''' SHGetSpecialFolderLocation nFolder params:
''''' Most folder locations are stored in:
''''' [HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders]
''''' Value specifying the types of folders to be listed in the dialog box as well as other options.
''''' This member can be 0 or one of the following values:
''''
''''' Windows desktop, virtual folder at the root of the name space.
''''Const CSIDL_DESKTOP = &H0
''''
''''' File system directory that contains the user's program groups
''''' (which are also file system directories).
''''Const CSIDL_PROGRAMS = &H2
''''
''''' Control Panel, virtual folder containing icons for the control panel applications.
''''Const CSIDL_CONTROLS = &H3
''''
''''' Printers folder, virtual folder containing installed printers.
''''Const CSIDL_PRINTERS = &H4
''''
''''' File system directory that serves as a common respository for documents.
''''Const CSIDL_PERSONAL = &H5   ' (Documents folder)
''''
''''' File system directory that contains the user's favorite Internet Explorer URLs.
''''Const CSIDL_FAVORITES = &H6
''''
''''' File system directory that corresponds to the user's Startup program group.
''''Const CSIDL_STARTUP = &H7
''''
''''' File system directory that contains the user's most recently used documents.
''''Const CSIDL_RECENT = &H8   ' (Recent folder)
''''
''''' File system directory that contains Send To menu items.
''''Const CSIDL_SENDTO = &H9
''''
''''' Recycle bin, file system directory containing file objects in the user's recycle bin.
''''' The location of this directory is not in the registry; it is marked with the hidden and
''''' system attributes to prevent the user from moving or deleting it.
''''Const CSIDL_BITBUCKET = &HA
''''
''''' File system directory containing Start menu items.
''''Const CSIDL_STARTMENU = &HB
''''
''''' File system directory used to physically store file objects on the desktop
''''' (not to be confused with the desktop folder itself).
''''Const CSIDL_DESKTOPDIRECTORY = &H10
''''
''''' My Computer, virtual folder containing everything on the local computer: storage
''''' devices, printers, and Control Panel. The folder may also contain mapped network drives.
''''Const CSIDL_DRIVES = &H11
''''
''''' Network Neighborhood, virtual folder representing the top level of the network hierarchy.
''''Const CSIDL_NETWORK = &H12
''''
''''' File system directory containing objects that appear in the network neighborhood.
''''Const CSIDL_NETHOOD = &H13
''''
''''' Virtual folder containing fonts.
''''Const CSIDL_FONTS = &H14
''''
''''' File system directory that serves as a common repository for document templates.
''''Const CSIDL_TEMPLATES = &H15   ' (ShellNew folder)

'========================================================

' Frees memory allocated by SHBrowseForFolder()

' Displays a dialog box that enables the user to select a shell folder.
' Returns a pointer to an item identifier list that specifies the location
' of the selected folder relative to the root of the name space. If the user
' chooses the Cancel button in the dialog box, the return value is NULL.
Private Declare Function SHBrowseForFolder Lib "SHELL32.DLL" Alias "SHBrowseForFolderA" _
                              (lpBrowseInfo As BROWSEINFO) As Long ' ITEMIDLIST

' Contains parameters for the the SHBrowseForFolder function and receives
' information about the folder selected by the user.
Private Type BROWSEINFO   ' bi
    
    ' Handle of the owner window for the dialog box.
    hOwner As Long
    
    ' Pointer to an item identifier list (an ITEMIDLIST structure) specifying the location
    ' of the "root" folder to browse from. Only the specified folder and its subfolders
    ' appear in the dialog box. This member can be NULL, and in that case, the
    ' name space root (the desktop folder) is used.
    pidlRoot As Long
    
    ' Pointer to a buffer that receives the display name of the folder selected by the
    ' user. The size of this buffer is assumed to be MAX_PATH bytes.
    pszDisplayName As String
    
    ' Pointer to a null-terminated string that is displayed above the tree view control
    ' in the dialog box. This string can be used to specify instructions to the user.
    lpszTitle As String
    
    ' Value specifying the types of folders to be listed in the dialog box as well as
    ' other options. This member can include zero or more of the following values below.
    ulFlags As Long
    
    ' Address an application-defined function that the dialog box calls when events
    ' occur. For more information, see the description of the BrowseCallbackProc
    ' function. This member can be NULL.
    lpfn As Long
    
    ' Application-defined value that the dialog box passes to the callback function
    ' (if one is specified).
    lParam As Long
    
    ' Variable that receives the image associated with the selected folder. The image
    ' is specified as an index to the system image list.
    iImage As Long

End Type

' BROWSEINFO ulFlags values:
' Value specifying the types of folders to be listed in the dialog box as well as
' other options. This member can include zero or more of the following values:

' Only returns file system directories. If the user selects folders
' that are not part of the file system, the OK button is grayed.
Const BIF_RETURNONLYFSDIRS = &H1

' Does not include network folders below the domain level in the tree view control.
' For starting the Find Computer
Const BIF_DONTGOBELOWDOMAIN = &H2

' Includes a status area in the dialog box. The callback function can set
' the status text by sending messages to the dialog box.
Const BIF_STATUSTEXT = &H4

' Only returns file system ancestors. If the user selects anything other
' than a file system ancestor, the OK button is grayed.
Const BIF_RETURNFSANCESTORS = &H8

' Only returns computers. If the user selects anything other
' than a computer, the OK button is grayed.
Const BIF_BROWSEFORCOMPUTER = &H1000

' Only returns (network) printers. If the user selects anything other
' than a printer, the OK button is grayed.
Const BIF_BROWSEFORPRINTER = &H2000

'济览文件夹------------

Enum cyFileNameCut
    CutFileNameOnly = 0
    CutFileNameWithOutExt = 1
    CutPath = 2
    CutExtionOnly = 3
End Enum

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hWnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long

'函数名称:cyFileDelete

'作用：删除文件/目录

'##################################################
'整理时间：2004-2-11 16:14:57
'  参数：
'   sDirFileName As String  --文件/目录
'   bPromptDelete As Boolean    --是否提示删除
'   bAllowUndo As Boolean   --是否允许恢复
'##################################################
'  例：
Public Function cyFileDelete(ByVal sDirFileName As String, Optional bPromptDelete As Boolean, Optional bAllowUndo As Boolean) As Boolean
    Const FO_DELETE = &H3
    Const FOF_NOCONFIRMATION = &H10
    Const FOF_ALLOWUNDO = &H40
    
    Dim SHFileOp As SHFILEOPSTRUCT
    SHFileOp.wFunc = FO_DELETE
    SHFileOp.pFrom = sDirFileName
    If bPromptDelete = False Then '不提示
        SHFileOp.fFlags = FOF_NOCONFIRMATION
    End If
    If bAllowUndo = True Then
        SHFileOp.fFlags = FOF_NOCONFIRMATION + FOF_ALLOWUNDO
    End If
    If SHFileOperation(SHFileOp) = 0 Then cyFileDelete = True
End Function

'函数名称:cyFileCopy

'作用：目录/文件复制(带动画)

'##################################################
'整理时间：2004-2-11 16:16:20
'  参数：
'   sDirFileName As String  --要复制的目录/文件名
'   sDestPath As String     --目录目标(不带文件名)
'##################################################
'  例：
Public Function cyFileCopy(ByVal sDirFileName As String, ByVal sDestPath As String) As Boolean
    Const FOF_ALLOWUNDO = &H40
    Const FOF_NOCONFIRMMKDIR = &H200
    Const FO_COPY = &H2
    
    Dim SHFileOp As SHFILEOPSTRUCT
    SHFileOp.wFunc = FO_COPY
    SHFileOp.pFrom = sDirFileName
    SHFileOp.pTo = sDestPath
    SHFileOp.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMMKDIR
    If SHFileOperation(SHFileOp) = 0 Then cyFileCopy = True
End Function

'函数名称:cyFileMove

'作用：移动目录/文件

'##################################################
'整理时间：2004-2-11 16:17:09
'  参数：
'   sDirFileName As String  --要移动的目录/文件名
'   sDestPath As String     --目标路径(不带文件名)
'##################################################
'  例：
Public Function cyFileMove(ByVal sDirFileName As String, ByVal sDestPath As String) As Boolean
    Const FO_MOVE = &H1
    Const FOF_NOCONFIRMATION = &H10
    Const FOF_NOCONFIRMMKDIR = &H200
    
    Dim SHFileOp As SHFILEOPSTRUCT
    SHFileOp.wFunc = FO_MOVE
    SHFileOp.pFrom = sDirFileName
    SHFileOp.pTo = sDestPath
    SHFileOp.fFlags = FOF_NOCONFIRMATION + FOF_NOCONFIRMMKDIR
    If SHFileOperation(SHFileOp) = 0 Then cyFileMove = True
End Function

'函数名称:cyCutFileName

'作用：截取一个完整文件名的某部分

'##################################################
'整理时间：2004-2-11 16:17:50
'  参数：
'   sFullFileName As String --源文件名全路径连文件名
'   Action As cyFileNameCut --要截取的部分(路径/只文件名/不带扩展名的文件名/扩展名..)
'##################################################
'  例：
Public Function cyCutFileName(ByVal sFullFileName As String, Action As cyFileNameCut) As String
    Dim Pos As Integer
    Dim sArray() As String
    If sFullFileName = "" Then Exit Function
    Select Case Action
        Case CutFileNameOnly
            sArray = Split(sFullFileName, "\")
            If UBound(sArray) > 0 Then '有多层路径 \.\.\,,,
                cyCutFileName = sArray(UBound(sArray))
            Else '只有一层   c:\..   aaa.exe
                If InStr(1, sArray(0), ":") > 0 Then '有:
                    sArray = Split(sFullFileName, ":")
                    cyCutFileName = sArray(UBound(sArray))
                Else '没有,只是文件名
                    cyCutFileName = sFullFileName
                End If
            End If
        Case CutFileNameWithOutExt
            sArray = Split(sFullFileName, "\")
            If UBound(sArray) > 0 Then  '有多个\
                cyCutFileName = sArray(UBound(sArray))
            Else    '没有\,但可能是c:abc.exe
                If InStr(1, sArray(0), ":") > 0 Then    '如果有:
                    sArray = Split(sFullFileName, ":")
                    cyCutFileName = sArray(UBound(sArray))
                Else '没有,只是文件名
                    cyCutFileName = sFullFileName
                End If
            End If
            
                Dim i As Long
                Dim j As Long
                i = InStr(1, cyCutFileName, ".")
                Do While i > 0
                    j = i
                    i = InStr(i + 1, cyCutFileName, ".")
                Loop
                If j > 0 Then
                    cyCutFileName = Left(cyCutFileName, j - 1)
                End If
                
        Case CutExtionOnly
            sArray = Split(sFullFileName, "\")
            If UBound(sArray) > 0 Then
                cyCutFileName = sArray(UBound(sArray))
            Else
                If InStr(1, sArray(0), ":") > 0 Then
                    sArray = Split(sFullFileName, ":")
                    cyCutFileName = sArray(UBound(sArray))
                Else '没有,只是文件名
                    cyCutFileName = sFullFileName
                End If
            End If
                If InStr(1, cyCutFileName, ".") > 0 Then
                    sArray = Split(cyCutFileName, ".")
                    cyCutFileName = sArray(UBound(sArray))
                Else
                    cyCutFileName = ""
                End If
        Case CutPath
                Pos = 1
                Do While InStr(Pos, sFullFileName, "\") <> 0
                    Pos = InStr(Pos, sFullFileName, "\") + 1
                Loop
                Do While InStr(Pos, sFullFileName, ":") <> 0
                    Pos = InStr(Pos, sFullFileName, ":") + 1
                Loop
            cyCutFileName = Left(sFullFileName, Pos - 1)
            If cyCutFileName = "" Then cyCutFileName = ".\"
    End Select

End Function

'函数名称:cyDialogSelectFolder

'作用：弹出目录选择窗口

'##################################################
'整理时间：2004-2-11 16:19:21
'  参数：
'   iHwnd As Long           --父窗口的名柄或设为0
'   sTitle As String        --该窗口的标题
'##################################################
'  例：
Function cyDialogSelectFolder(ByVal iHwnd As Long, Optional sTitle As String) As String


  Dim bi As BROWSEINFO
  Dim nFolder As Long
  Dim IDL As ITEMIDLIST
  Dim pidl As Long
  Dim sPath As String
  
  With bi
    ' The dialog's owner window...
    .hOwner = iHwnd
    
    ' Set the Browse dialog root folder
    'nFolder = GetFolderValue(0)
    
    ' Fill the item id list with the pointer of the selected folder item, rtns 0 on success
    ' ==================================================
    ' If this function fails because the selected folder doesn't exist,
    ' .pidlRoot will be uninitialized & will equal 0 (CSIDL_DESKTOP)
    ' and the root will be the Desktop.
    ' DO NOT specify the CSIDL_ constants for .pidlRoot !!!!
    ' The SHBrowseForFolder() call below will generate a fatal exception
    ' (GPF) if the folder indicated by the CSIDL_ constant does not exist!!
    ' ==================================================
    Dim NOERROR As Long
    If SHGetSpecialFolderLocation(ByVal iHwnd, ByVal nFolder, IDL) = NOERROR Then
      .pidlRoot = IDL.mkid.cb
    End If
    
    
    .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
    
    
    Dim sSelPath As String
    sSelPath = cyDialogSelectFolderInitDir
    Dim lpSelPath As Long
     lpSelPath = LocalAlloc(&H0 Or &H40, Len(sSelPath))
     MoveMemory ByVal lpSelPath, ByVal sSelPath, Len(sSelPath)
     .lParam = lpSelPath

    
    
    ' Initialize the buffer that rtns the display name of the selected folder
    .pszDisplayName = String$(256, 0)
    
    ' Set the dialog's banner text
    .lpszTitle = sTitle
    
    ' Set the type of folders to display & return
    ' -play with these option constants to see what can be returned
    .ulFlags = 4
    
  End With
  
  ' Clear previous return vals before the
  ' dialog is shown (it might be cancelled)
  'txtPath = ""
  'txtDisplayName = ""
  'pic16Icon.Picture = LoadPicture()   ' clears prev icon
  'pic32Icon.Picture = LoadPicture()
showBrowser:
  ' Show the Browse dialog
  pidl = SHBrowseForFolder(bi)
  
  ' If the dialog was cancelled...
  If pidl = 0 Then Exit Function
    
  ' Fill sPath w/ the selected path from the id list
  ' (will rtn False if the id list can't be converted)
  sPath = String$(256, 0)
  SHGetPathFromIDList ByVal pidl, ByVal sPath
    sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
    If sPath = "" Then
        GoTo showBrowser
    End If
  ' Display the path and the name of the selected folder
  cyDialogSelectFolder = IIf(Right(sPath, 1) = "\", sPath, sPath & "\")
  ' Get the 16x16 icon info from the id list using the pidl
 ' SHGetFileInfo ByVal pIdl, 0&, SHFI, Len(SHFI), _
                       SHGFI_PIDL Or SHGFI_ICON Or SHGFI_SMALLICON
  ' The 16x16 icon handle rtnd in SHFI.hIcon is stretched to 32x32.
  ' DrawIconEx() will shrink (or stretch) the icon per it's cxWidth & cyWidth params
 ' DrawIconEx pic16Icon.hdc, 0, 0, SHFI.hIcon, 16, 16, 0, 0, DI_NORMAL
  'pic16Icon.Refresh
  
  ' Get the 32x32 icon info from the id list
  'SHGetFileInfo ByVal pIdl, 0&, SHFI, Len(SHFI), _
                       SHGFI_PIDL Or SHGFI_ICON
  ' SHFI.hIcon is OK here so DrawIcon() can be used
  'DrawIcon pic32Icon.hdc, 0, 0, SHFI.hIcon
  'pic32Icon.Refresh
  
  ' Frees the memory SHBrowseForFolder()
  ' allocated for the pointer to the item id list
  'CoTaskMemFree pIdl
  
End Function

Public Function cyDialogSave(ByVal hWnd As Long, ByVal sTitle As String, Optional ByVal sFilter As String, Optional ByVal sInitDir As String) As String
Const iBuffLen As Long = 255
Dim OFName As OPENFILENAME, sBuffer As String * iBuffLen
    If sFilter = "" Then sFilter = "所有文件|*.*"
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = Replace(sFilter, "|", Chr(0))
    OFName.lpstrFile = sBuffer
    OFName.nMaxFile = iBuffLen
    OFName.lpstrFileTitle = sTitle
    OFName.nMaxFileTitle = iBuffLen
    OFName.lpstrInitialDir = sInitDir
    
    If Len(sTitle) Then
        OFName.lpstrTitle = sTitle
    Else
        OFName.lpstrTitle = sTitle
    End If
    
    If GetSaveFileName(OFName) Then
        cyDialogSave = Left$(OFName.lpstrFile, InStr(1, OFName.lpstrFile, Chr(0)) - 1)
        
        '扩展名检查
        Dim sA() As String
        sA = Split(OFName.lpstrFilter, Chr(0))
        Dim sTemp As String
        If OFName.nFilterIndex = 0 Then
            sTemp = "*"
        Else
            sTemp = Replace(sA(OFName.nFilterIndex * 2 - 1), "*.", "")
        End If
        
        '如果是扩展名是*则不处理扩展名
        If sTemp = "*" Then Exit Function
        
        '已输入扩展名，不用再加
        If Right(UCase(cyDialogSave), Len(sTemp) + 1) = "." & UCase(sTemp) Then
        Else
        '未输入则在后面添加
            If sA(UBound(sA)) <> "*" Then cyDialogSave = cyDialogSave & "." & sTemp
        End If
        
    End If

End Function

Public Function cyDialogOpenFileNameToArray(ByVal sFileNames As String)
On Error GoTo Err
Dim sA() As String
    sA = Split(sFileNames, Chr(0))
    sA(0) = IIf(Right(sA(0), 1) = "\", sA(0), sA(0) & "\")

Dim i As Long
    For i = 1 To UBound(sA)
        sA(i) = sA(0) & sA(i)
    Next
    cyDialogOpenFileNameToArray = sA
    Exit Function

Err:
    '没有文件名
    If Err.Number = 9 Then
        ReDim sA(0)
        cyDialogOpenFileNameToArray = sA
        
    End If


End Function

'函数名称:cyDialogOpen

'作用：弹出一个通用文件选择窗

'##################################################
'整理时间：2004-2-11 16:20:14
'  参数：
'   sTitle As String        --其标题
'   sFilter As String       --过滤器(如: 文本文件|*.txt|数据库文件|*.mdb)
'   sInitDir As String      --初始目录
'   bMultiSelect As Boolean --是否多选
'##################################################
'  例：
Public Function cyDialogOpen(ByVal hWnd As Long, ByVal sTitle As String, Optional ByVal sFilter As String, Optional ByVal sInitDir As String, Optional bMultiSelect As Boolean) As String
    Dim lpOpenFile As OPENFILENAME
    With lpOpenFile
        .lStructSize = Len(lpOpenFile)
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrTitle = sTitle
        .lpstrFile = String$(1024 * 10, 0)
        .lpstrFilter = Replace(sFilter, "|", Chr(0))
        .lpstrInitialDir = sInitDir
        '.lpstrFileTitle = "*.exe"
        '.lpstrCustomFilter = "*.exe"
        
        If bMultiSelect Then
            .flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT Or OFN_EXPLORER
        Else
            .flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
        End If

        .nMaxFile = 2048
        '.lpstrFileTitle = String$(255, 0)
        '.nMaxFileTitle = 255
    End With
    
    If GetOpenFileName(lpOpenFile) <> 0 Then
        If bMultiSelect Then
            cyDialogOpen = lpOpenFile.lpstrFile
        Else
            cyDialogOpen = Left$(lpOpenFile.lpstrFile, InStr(lpOpenFile.lpstrFile, vbNullChar) - 1)
        End If
        If InStr(1, cyDialogOpen, Chr(0)) > 0 Then
            cyDialogOpen = Left(cyDialogOpen, InStr(1, cyDialogOpen, Chr(0) & Chr(0)) - 1)
        End If
        Exit Function
    Else
        Exit Function
    End If
End Function


Function cyUnitePathFileName(ByVal sPath As String, ByVal sFileName As String) As String
    cyUnitePathFileName = IIf(Left(sPath, 2) = "\\", "\" + Replace(sPath + "\" + sFileName, "\\", "\", 1), Replace(sPath + "\" + sFileName, "\\", "\", 1))
End Function


'过程名称:cyOpenFolder

'作用：使用我的电脑或资源管理器打开一个目录

'##################################################
'整理时间：2004-2-11 16:24:11
'  参数：
'   sFolder As String       --要打开的目录
'   bOpenWithMyComputer As Boolean  --是否用我的电脑来打开
'   bOpenAsRoot As Boolean  --是否作为根目录来打开
'##################################################
'  例：
Public Sub cyOpenFolder(sFolder As String, Optional bOpenWithMyComputer As Boolean, Optional bOpenAsRoot As Boolean)
    Dim sTag As String
        If bOpenWithMyComputer = True Then
            sTag = "explorer /n"
        Else
            sTag = "explorer /e"
        End If
        If bOpenAsRoot = True Then
            Shell sTag + ",/root," + sFolder, vbNormalFocus
        
        Else
            Shell sTag + "," + sFolder, vbNormalFocus
        
        End If
End Sub

'函数名称:cyFileStrGetLineStr

'作用：取得文本文件中某行的字符串

'##################################################
'整理时间：2004-2-11 16:27:39
'  参数：
'   sFileName As String     --文件名
'   iLineNum As Long        --要取得的行号
'##################################################
'  例：
'        Dim F As New fileClass
'        Debug.Print F.cyFileStrGetLineStr("c:\test.txt", 2)
Public Function cyFileStrGetLineStr(ByVal sFileName As String, ByVal iLineNum As Long) As String
    Dim iArray() As Byte
    ReDim iArray(FileLen(sFileName) - 1)
    Open sFileName For Binary As #1
        Get #1, , iArray
    Close #1
    Dim sArray() As String
    sArray = Split(StrConv(iArray, vbUnicode), Chr(13) + Chr(10))
    If iLineNum > UBound(sArray) + 1 Then '取的行数超出了文件的行数
        
    Else
        cyFileStrGetLineStr = sArray(iLineNum - 1)
    End If
End Function

'函数名称:cyFileStrInput

'作用：将文本文件内的字符串读到一个变量中

'##################################################
'整理时间：2004-2-11 16:28:39
'  参数：
'   sFileName As String     --要读出的文件名
'##################################################
'  例：
Public Function cyFileStrInput(ByVal sFileName As String) As String
On Error Resume Next
    Dim iArray() As Byte
    ReDim iArray(FileLen(sFileName) - 1)
    Open sFileName For Binary As #1
        Get #1, , iArray
    Close #1
    cyFileStrInput = StrConv(iArray, vbUnicode)

End Function

'过程名称:cyFileStrOutput

'作用：将字符串变量写到文本文件中

'##################################################
'整理时间：2004-2-11 16:29:19
'  参数：
'   sFileName As String     --要写的文本文件
'   sStr As String          --字符串变量
'##################################################
'  例：
Public Sub cyFileStrOutput(ByVal sFileName As String, sStr As String)
    Open sFileName For Output As #1
        Print #1, sStr
    Close #1
End Sub

'过程名称:cyFileStrAppend

'作用：将字符串添加到文件中

'##################################################
'整理时间：2004-2-11 16:29:53
'  参数：
'   sFileName As String     --要添加的文件名
'   sStr As String          --要增加的字符串
'##################################################
'  例：
Public Sub cyFileStrAppend(ByVal sFileName As String, sStr As String)
    Open sFileName For Append As #1
        Print #1, sStr
    Close #1
End Sub

'函数名称:cyFileToByteArray

'作用：将文件以二进制方式读出并保存到一个字节型数组中

'##################################################
'整理时间：2004-2-11 16:30:49
'  参数：
'   sFileName As String     --要读的文件
'##################################################
'  例：
Public Function cyFileToByteArray(ByVal sFileName As String)
    Dim iArray() As Byte
    ReDim iArray(FileLen(sFileName) - 1)
    Open sFileName For Binary As #1
        Get #1, , iArray
    Close #1
    cyFileToByteArray = iArray
End Function

'函数名称:cyFolderExist
'作用：检测目录是否存在
'===========================================================
'整理时间：2006-8-26 14:55
'  参数：
'           sFolder     ---待检测的目录名称
'===========================================================
Public Function cyFolderExist(sFolder As String) As Boolean
    On Error Resume Next
    cyFolderExist = IIf(Len(Dir(sFolder, vbDirectory + vbHidden + vbSystem + vbVolume)) > 0, True, False)
    
End Function
Public Function fileClassist(sFileName As String) As Boolean
    On Error Resume Next
    FileLen sFileName
    fileClassist = Not (Err.Number > 0)
End Function

'过程名称:cyMkDirEx

'作用：可以生成多层的子目录(如:c:\1\23\3\5\6\7)

'##################################################
'整理时间：2004-2-11 16:31:35
'  参数：
'   sFolder As String       --要生成的目录
'##################################################
'  例：
Public Sub cyMkDirEx(ByVal sFolder As String)
    SHCreateDirectoryEx 0, sFolder, ByVal 0&
End Sub

Public Function cyGetSpecialFolder(Folder As SystemFolders) As String
On Error Resume Next
  
  Dim ReturnCode As Long
  Dim ReturnPath As String
  Dim ID_List As ITEMIDLIST
  
    If Folder = TempPath Then
    '临时目录
        Dim bStr1 As String * 255
        cyGetSpecialFolder = Mid(bStr1, 1, GetTempPath(255, bStr1))
    
    Else
    
        ' Get the special folder
        ReturnCode = SHGetSpecialFolderLocation(0, Folder, ID_List)
        
        If ReturnCode = 0 Then
          ReturnPath = Space(512)                                                   ' Create a buffer
          ReturnCode = SHGetPathFromIDList(ByVal ID_List.mkid.cb, ByVal ReturnPath) ' Get the path from the IDList
          cyGetSpecialFolder = Left(ReturnPath, InStr(ReturnPath, Chr(0)) - 1)        ' Remove the unnecesarry chr(0)'s
        End If
        
    End If
    If cyGetSpecialFolder <> "" Then cyGetSpecialFolder = IIf(Right(cyGetSpecialFolder, 1) = "\", cyGetSpecialFolder, cyGetSpecialFolder & "\")
    
End Function

'打开控制面板
Sub cyOpenControlPanel()
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL", 9)
End Sub

'sFileOrFolder 是要创建快捷方式的文件，全路径！
'sShortCutName　快捷方式文件的名字，如果忽略就是原本文件的名称
'sShortCutPath 目标路径，要创建快捷方式的路径（不带文件名）,如果忽略就用桌面的路径
'IconIndex 要创建方式文件的图标索引值，如果忽略就为0(0为文件本身的图标)，如果超过文件本身的图标索引值，就为0
'如果要为文件夹创建快捷方式，就要在sFileOrFolder加上一个"\" 例如：如果要创建C盘根目录就不应该用C: 而是用C:\
Public Function cyCreateShortCut(ByVal sFileOrFolder As String, Optional ByVal sShortCutName As String, Optional ByVal sShortCutPath As String, Optional ByVal IconIndex As Integer) As Boolean

On Error Resume Next
Dim WSHShell As Object, MyShortcut As Object, sDesktopPath As String
Dim sFileName As String
Dim sPath As String

'文件所执行的路径（有些程序可以忽略，但注意这可能使用程序不能正常运行）
Set WSHShell = CreateObject("WScript.Shell")
sFileName = Dir(sFileOrFolder)

'如果文件不存在就返回0并退出函数
If sFileName = "" And Right(sFileOrFolder, 1) <> "\" Then cyCreateShortCut = 0: Exit Function

'如果没有指定快捷方式的名字就用原来的文件名
If sShortCutName = "" Then sShortCutName = sFileName

'如果没有指定目标的路径
If sShortCutPath = "" Then
   sShortCutPath = WSHShell.SpecialFolders("Desktop")
   '就用桌面为目标的路径
End If

    If Right(sShortCutPath, 1) <> "\" Then sShortCutPath = sShortCutPath & "\"
    sPath = Left(sFileOrFolder, InStrRev(sFileOrFolder, "\", , vbTextCompare) - 1)
    Set MyShortcut = WSHShell.CreateShortcut(sShortCutPath & sShortCutName & ".lnk")
    MyShortcut.TargetPath = WSHShell.ExpandEnvironmentStrings(sFileOrFolder)
    MyShortcut.WorkingDirectory = WSHShell.ExpandEnvironmentStrings(sPath)
    MyShortcut.WindowStyle = 4
    MyShortcut.IconLocation = WSHShell.ExpandEnvironmentStrings(sFileOrFolder & " , " & IconIndex)
    MyShortcut.Save
    cyCreateShortCut = True
    '返回真

If Err Then
    '返回假
    cyCreateShortCut = False
End If

End Function


'将文件以base64编码方式转换为XML文件
Function cyXmlCreateFromFile(sFile As String, sXml As String) As Boolean
On Error GoTo Err
    Dim oEle As IXMLDOMElement
    Dim oRoot As IXMLDOMElement
    Dim oNode As IXMLDOMNode
    Dim oDoc As DOMDocument
    Set oDoc = Nothing
    Set oDoc = New DOMDocument
    oDoc.resolveExternals = True
    
    Set oNode = oDoc.createProcessingInstruction("xml", "version='1.0'")
    Set oNode = oDoc.insertBefore(oNode, oDoc.childNodes.Item(0))
    
    '生成XML的根元素
    Set oRoot = oDoc.createElement("Root")
    Set oDoc.documentElement = oRoot
    oRoot.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"
    
    ' 使用不同的数据类型生成一些简单的节点
    '文本型
    Set oNode = oDoc.createElement("Creator")
    oNode.Text = "DragonSchool"
    oRoot.appendChild oNode
    
    '日期型
    Set oNode = oDoc.createElement("CreateDate")
    oRoot.appendChild oNode
    Set oEle = oNode
    oEle.dataType = "date"  ' 必须设定节点的数据类型
    oEle.nodeTypedValue = Now
    
    Set oNode = oDoc.createElement("bgColor")
    oRoot.appendChild oNode
    Set oEle = oNode
    oEle.dataType = "bin.hex"
    oEle.Text = &HFFCCCC
    
    Set oNode = oDoc.createElement("Data")
    oRoot.appendChild oNode
    Set oEle = oNode
    oEle.dataType = "bin.base64"
    
    Dim iFileNum As Integer
    Dim byteArray() As Byte
    Dim lLen As Long
    
    '将文件读入二进制数组
    iFileNum = FreeFile()
    Open sFile For Binary Access Read As iFileNum
        lLen = FileLen(sFile)
        ReDim byteArray(lLen - 1)
        Get iFileNum, , byteArray
    Close iFileNum
    
    '将数组内容赋予对象
    oEle.nodeTypedValue = byteArray
    
    '保存到XML文件中
    oDoc.Save sXml
    
    cyXmlCreateFromFile = True
    Exit Function
Err:
End Function

'将XML文件还原为文件
Function cyXmlStoreToFile(sXml As String, sFile As String) As Boolean
On Error GoTo Err
    Dim oDoc As DOMDocument
    Set oDoc = New DOMDocument
    
    If oDoc.Load(sXml) = True Then
        '还原文件
        Dim iFileNum As Integer
        Dim byteArray() As Byte
        Dim oNode As IXMLDOMNode
        
        If Not (oDoc Is Nothing) Then
            '取得文件内容
            Set oNode = oDoc.documentElement.selectSingleNode("/Root/Data")
            
            '确保变量类型为数组型
            byteArray = oNode.nodeTypedValue
            
            '保存到文件
            iFileNum = FreeFile()
            Open sFile For Binary Access Write As iFileNum
                Put iFileNum, , byteArray
            Close iFileNum
        End If
        cyXmlStoreToFile = True
    End If
Err:
End Function

Public Function cyGetShortPath(ByVal sFileName As String) As String
    Dim sPath As String * 255
    GetShortPathName sFileName, sPath, 255
    cyGetShortPath = Left(sPath, InStr(1, sPath, Chr(0)) - 1)
    cyGetShortPath = IIf(Right(cyGetShortPath, 1) = "\", cyGetShortPath, cyGetShortPath & "\")
End Function

Public Function cyGetShortFile(ByVal sFileName As String) As String
    Dim sPath As String * 255
    GetShortPathName sFileName, sPath, 255
    cyGetShortFile = Left(sPath, InStr(1, sPath, Chr(0)) - 1)
End Function

'搜索指定路径下的文件
Sub cySearchFile(ByRef sSearchPath As String, ByRef sSearchFile As String, Optional ByVal bSearchSubFolder As Boolean)
    '包括子目录
    If bSearchSubFolder Then
        tmpSearchIncludeSubDir sSearchPath, sSearchFile
    Else
        tmpSearchNoSubDir sSearchPath, sSearchFile
    End If
End Sub

Sub cySearchStop()
    '终止搜索
     bStopSearch = True
End Sub

'不搜索子目录
Private Sub tmpSearchNoSubDir(ByRef sSearchPath As String, ByRef sSearchFile As String)
    '常量
    Const MAX_PATH = 260
    Const INVALID_HANDLE_VALUE = -1
    Const FILE_ATTRIBUTE_DIRECTORY = &H10
    
    
    Dim sFileName As String ' 文件名
    Dim sFolderName As String ' 子目录名
    Dim sFolderNames() As String ' 目录数组
    Dim iFolderCount As Integer ' 当前路径的目录数
    Dim i As Integer ' 循环计数器变量
    Dim hSearch As Long ' 搜索句柄变量
    Dim WFD As WIN32_FIND_DATA
    Dim iCount As Integer

    If Right(sSearchPath, 1) <> "\" Then sSearchPath = sSearchPath & "\"
    
        ' 遍历目录并累计文件总数
        hSearch = FindFirstFile(sSearchPath & sSearchFile, WFD)
        iCount = True
        If hSearch <> INVALID_HANDLE_VALUE Then
        While iCount
            DoEvents
            '终止搜索
            If bStopSearch = True Then Exit Sub
            
            sFileName = Left(WFD.cFileName, InStr(1, WFD.cFileName, Chr(0)) - 1)
            If (sFileName <> ".") And (sFileName <> "..") Then
                RaiseEvent cyFoundFile(sSearchPath, sFileName)
            End If
            iCount = FindNextFile(hSearch, WFD) ' 获取下一个文件
        Wend
        iCount = FindClose(hSearch)
        End If
            
        RaiseEvent cyFoundStoped

End Sub

'包括子目录
Private Sub tmpSearchIncludeSubDir(ByRef sSearchPath As String, ByRef sSearchFile As String)
    '常量
    Const INVALID_HANDLE_VALUE = -1
    Const FILE_ATTRIBUTE_DIRECTORY = &H10
    
    
    Dim sFileName As String ' 文件名
    Dim sFolderName As String ' 子目录名
    Dim sFolderNames() As String ' 目录数组
    Dim iFolderCount As Integer ' 当前路径的目录数
    Dim i As Integer ' 循环计数器变量
    Dim hSearch As Long ' 搜索句柄变量
    Dim WFD As WIN32_FIND_DATA
    Dim iCount As Integer

    If Right(sSearchPath, 1) <> "\" Then sSearchPath = sSearchPath & "\"
    
            '不搜索子目录
            '规范路径
             '搜索子目录
            iFolderCount = 0
            ReDim sFolderNames(iFolderCount)
            iCount = True
            
            hSearch = FindFirstFile(sSearchPath & "*", WFD)
            If hSearch <> INVALID_HANDLE_VALUE Then
                Do While iCount
                    DoEvents
                    '终止搜索
                    If bStopSearch = True Then Exit Sub
                    
                    sFolderName = Left(WFD.cFileName, InStr(1, WFD.cFileName, Chr(0)) - 1)
                    If (sFolderName <> ".") And (sFolderName <> "..") Then
                        If GetFileAttributes(sSearchPath & sFolderName) And FILE_ATTRIBUTE_DIRECTORY Then
                            sFolderNames(iFolderCount) = sFolderName
                            iFolderCount = iFolderCount + 1
                            ReDim Preserve sFolderNames(iFolderCount)
                        End If
                    End If
                    iCount = FindNextFile(hSearch, WFD) '获取下一个子目录
                Loop
                iCount = FindClose(hSearch)
            End If
            
        
        ' 遍历目录并累计文件总数
        hSearch = FindFirstFile(sSearchPath & sSearchFile, WFD)
        iCount = True
        If hSearch <> INVALID_HANDLE_VALUE Then
        While iCount
            DoEvents
            '终止搜索
            If bStopSearch = True Then Exit Sub
            
            sFileName = Left(WFD.cFileName, InStr(1, WFD.cFileName, Chr(0)) - 1)
            If (sFileName <> ".") And (sFileName <> "..") Then
                RaiseEvent cyFoundFile(sSearchPath, sFileName)
            End If
            iCount = FindNextFile(hSearch, WFD) ' 获取下一个文件
        Wend
        iCount = FindClose(hSearch)
        End If
        
        iFolderCounter = iFolderCounter + 1
        
        '如果子目录存在则遍历之
        If iFolderCount > 0 Then
            For i = 0 To iFolderCount - 1
                Call cySearchFile(sSearchPath & sFolderNames(i) & "\", sSearchFile, True)
            Next i
        End If
        
        iFolderCounter = iFolderCounter - 1
        If iFolderCounter = 0 Then
            RaiseEvent cyFoundStoped

        End If
End Sub


Public Function cyFilesCopyToClipBoard(sFiles() As String) As Boolean
   On Error Resume Next
   Dim sFileStr As String
   Dim DF As DROPFILES
   Dim hGlobal As Long
   Dim lpGlobal As Long
   Dim i As Long
   
   '清除剪贴版中现存的数据
   If OpenClipboard(0&) Then
        Call EmptyClipboard
      
        For i = LBound(sFiles) To UBound(sFiles)
            If Trim(sFiles(i)) <> "" Then sFileStr = sFileStr & sFiles(i) & vbNullChar
        Next i
        sFileStr = sFileStr & vbNullChar
      
        hGlobal = GlobalAlloc(GHND, Len(DF) + Len(sFileStr))
        If hGlobal Then
            lpGlobal = GlobalLock(hGlobal)
         
            DF.pFiles = Len(DF)
            Call CopyMem(ByVal lpGlobal, DF, Len(DF))
            Call CopyMem(ByVal (lpGlobal + Len(DF)), ByVal sFileStr, _
                Len(sFileStr))
            Call GlobalUnlock(hGlobal)
         
            If SetClipboardData(CF_HDROP, hGlobal) Then
                cyFilesCopyToClipBoard = True
            End If
        End If
        Call CloseClipboard
    End If
End Function

Public Function cyFilesPasteFromClipBoard(ByVal sDestinationPath As String) As Long
   On Error GoTo Err
   Dim hDrop As Long
   Dim nFiles As Long
   Dim i As Long
   Dim desc As String
   Dim sTemp As String * 255
   Dim sFileName As String
   
'   Dim pt As POINTAPI
   Dim tfStr As SHFILEOPSTRUCT
   
   '确定剪贴版的数据格式是文件，并打开剪贴版
   If IsClipboardFormatAvailable(CF_HDROP) Then
        If OpenClipboard(0&) Then
            hDrop = GetClipboardData(CF_HDROP)
            '获得剪贴板中文件数量
            nFiles = DragQueryFile(hDrop, -1&, "", 0)
            tfStr.wFunc = FO_COPY
            tfStr.pTo = sDestinationPath
         
            For i = 0 To nFiles - 1
                DoEvents
                If bStopFilePaste Then
                    bStopFilePaste = False
                    GoTo Err
                End If
                RaiseEvent cyFilePasting(sFileName, nFiles, i + 1)
               
                Call DragQueryFile(hDrop, i, sTemp, Len(sTemp))
                sFileName = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
                tfStr.pFrom = sFileName
                SHFileOperation tfStr
            Next i
        End If
        cyFilesPasteFromClipBoard = nFiles
    End If
Err:
    Call CloseClipboard
End Function

Sub cyXCOPY(ByVal sSourceFolder As String, ByVal sTargetPath As String)
    Dim sStr1 As String
    Dim sStr2 As String
    
    '得到短的源目录
    sStr1 = IIf(Right(sSourceFolder, 1) = "\", Left(sSourceFolder, Len(sSourceFolder) - 1), sSourceFolder)
    
    '生成短的目标目录
    sStr2 = IIf(Right(sTargetPath, 1) = "\", sTargetPath, sTargetPath & "\") & cyCutFileName(sStr1, CutFileNameWithOutExt)
    cyMkDirEx sStr2
    
    Shell "xcopy """ & sStr1 & """ """ & sStr2 & """ /s/e/c/h", vbNormalFocus

End Sub

Sub cyDeltree(ByVal sSourceFolder As String)
    '得到短的源目录
    Dim sStr As String
    sStr = cyGetShortPath(sSourceFolder)
    
    Shell "cmd /c rmdir /s/q " & Left(sStr, Len(sStr) - 1), vbHide
End Sub

Function cyWinRarAdd(ByVal sSource As String, ByVal sTargetRar As String, Optional ByVal sOtherCommand As String, Optional ByVal sOtherSwitch As String, Optional ByVal sPassWord As String, Optional ByVal iTimeout As Long = 180) As Boolean
    On Error Resume Next
    
    Dim F As New fileClass
    Dim S As New systemClass
    
    Dim bStr As String * 255
    Dim sStr As String
    
    '得到临时目录
    Dim sTemp As String
    Dim sTemp2 As String
    Dim sTemp3 As String
    Dim i As Long
    
    '删除还在运行的RAR实体，避免出错
    Call S.cyKillProcess(KillFromExeName, "RAR")
    
    '删除可能存在的目标文件
    Kill sTargetRar
    
    sTemp = Mid(bStr, 1, GetTempPath(255, bStr))
    sTemp = IIf(Right(sTemp, 1) = "\", sTemp, sTemp & "\")
    
    '如果Rar文件不存在,则生成
    If Not fileClassist(sTemp & "Rar.exe") Then
        Dim bArray() As Byte
        bArray = LoadResData(103, "CUSTOM")          '将自定义资源中103号资源读入数组OCX
        Open sTemp & "Rar.Exe" For Binary As #1
            Seek #1, 1
            Put #1, , bArray
        Close #1
    
    End If
    
    Randomize
    sTemp2 = CStr(CLng(Rnd * 10000))
    
    '得到RAR文件的路径，生成临时RAR文件
    sTemp3 = F.cyCutFileName(sTargetRar, CutPath)
    sStr = sTemp & "Rar.Exe a " & sOtherCommand & IIf(sPassWord = "", " ", " -hp" & sPassWord) & " " & sOtherSwitch & " """ & sTemp3 & sTemp2 & """ """ & sSource & """"
    Shell sStr, vbHide

    '检测是否可以改文件名
    For i = 0 To iTimeout
        '改名成功
        S.cyDelay 1
        If MoveFile(sTemp3 & sTemp2 & ".rar", sTargetRar) > 0 Then
            cyWinRarAdd = True
            Exit Function
        
        End If
    Next

    '失败时删除压缩失败的文件
    Call S.cyKillProcess(KillFromExeName, "RAR")
On Error Resume Next
    Kill sTemp3 & sTemp2 & ".Rar"
    Kill sTemp3 & sTemp2

End Function

'返回文件版本号
Function cyGetFileVersion(ByVal sFileName As String) As String

   Dim rc As Long, lDummy As Long, sBuffer() As Byte
   Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
   Dim lVerbufferLen As Long

   '*** Get size ****
   lBufferLen = GetFileVersionInfoSize(sFileName, lDummy)
   If lBufferLen < 1 Then
      Exit Function
   End If

   '**** Store info to udtVerBuffer struct ****
   ReDim sBuffer(lBufferLen)
   rc = GetFileVersionInfo(sFileName, 0&, lBufferLen, sBuffer(0))
   rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
   
   '函数为加1，详见上面的声明
   MoveMemory1 udtVerBuffer, lVerPointer, Len(udtVerBuffer)
   cyGetFileVersion = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)


'    Dim rc As Long
'    Dim lDummy As Long
'    Dim sBuffer() As Byte
'    Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
'    Dim lVerbufferLen As Long
'
'    lBufferLen = GetFileVersionInfoSize(sFileName, lDummy)
'    If lBufferLen < 1 Then
'         cyGetFileVersion = "0.0.0.0"
'         Exit Function
'    End If
'
'    ReDim sBuffer(lBufferLen)
'    rc = GetFileVersionInfo(sFileName, 0&, lBufferLen, sBuffer(0))
'    rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
'    MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)

    cyGetFileVersion = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)
End Function

'根据文件版本号检查文件是否更新
Function cyFileIsNewVersion(ByVal sOldFileVersion As String, ByVal sNewFileVersion As String) As Boolean
    Dim sA1() As String
    Dim sA2() As String
    
    '文件版本号相同
    If (sOldFileVersion) = (sNewFileVersion) Then Exit Function
    '文件版本号不完整
    If Len(sOldFileVersion) = 0 Or Len(sNewFileVersion) = 0 Then Exit Function
    sA1 = Split(sOldFileVersion, ".")
    sA2 = Split(sNewFileVersion, ".")
    
    If UBound(sA1) <> UBound(sA2) Then Exit Function
    
    If CLng(sA1(0)) < CLng(sA2(0)) Then
        cyFileIsNewVersion = True
    
    ElseIf CLng(sA1(0)) > CLng(sA2(0)) Then
    
    Else
    
        If CLng(sA1(1)) < CLng(sA2(1)) Then
        '第一级后者大
            cyFileIsNewVersion = True
        
    
        ElseIf CLng(sA1(1)) > CLng(sA2(1)) Then
        
        Else
            If CLng(sA1(2)) < CLng(sA2(2)) Then
            '第一级后者大
                cyFileIsNewVersion = True
            
            ElseIf CLng(sA1(2)) > CLng(sA2(2)) Then
            Else
                If CLng(sA1(3)) < CLng(sA2(3)) Then
                '第一级后者大
                    cyFileIsNewVersion = True
                    
                End If
            
            End If
        End If
    End If
    
End Function

'函数名称:cyFileSplit
'作用：将文件分割为适当大小的多个小文件
'===========================================================
'整理时间：2004-2-11 16:25:29
'  参数：
'   sSourceFile As String           --源文件
'   iKiloByte As Long               --分割的大小(以K为单位)
'   sTargetPath As String           --目标文件
'===========================================================

Public Function cyFileSplit(ByVal sSourceFile As String, ByVal iKiloByte As Long, Optional sTargetPath As String) As Boolean
On Error GoTo Err
    
    Dim Buff() As Byte
    Dim i As Long                       '记录文件数量
    Dim iFileLen As Long                '记录文件的字节长度
    Dim sFileName As String
    
    iKiloByte = iKiloByte * 2 ^ 10
    ReDim Buff(iKiloByte - 1)
    
    sFileName = cyCutFileName(sSourceFile, CutFileNameOnly)
    
    '生成目标目录
    cyMkDirEx sTargetPath
    
    i = 1
    '得到文件大小
    iFileLen = FileLen(sSourceFile)
    sTargetPath = IIf(Right(sTargetPath, 1) = "\", sTargetPath, sTargetPath & "\")
    sTargetPath = IIf(sTargetPath = "", sSourceFile, sTargetPath)
    
    '批处理文件名
    Dim sBatName As String
    
    '打开源文件
    Open sSourceFile For Binary Access Read As #1
        
        Do While iFileLen > i * iKiloByte
            Seek #1, (i - 1) * iKiloByte + 1
            Get #1, , Buff
            
            '写入文件名列表
            sBatName = sBatName & " + """ & sFileName & "." & Format(i, "000") & """"
            
            Open sTargetPath & sFileName & "." & Format(i, "000") For Binary Access Write As #2
                Put #2, , Buff
            Close #2
            i = i + 1
        Loop
        
        '写入文件名列表
        sBatName = sBatName & " + """ & sFileName & "." & Format(i, "000") & """"
        
        '剩下不够一个文件的尺寸，则以剩下的尺寸建一个文件
        Open sTargetPath & sFileName & "." & Format(i, "000") For Binary Access Write As #2
            ReDim Buff(iFileLen - (i - 1) * (iKiloByte) - 1)
            Seek #1, (i - 1) * iKiloByte + 1
            Get #1, , Buff  '每次读入个iKiloByte个字节进来
            Put #2, , Buff
            
        Close #2
        
    Close #1

    Mid(sBatName, 1, 3) = "   "
    cyFileStrOutput sTargetPath & sFileName & ".Bat", "copy /b " & sBatName & " """ & sFileName & """"
    cyFileSplit = True
    Exit Function
    
Err:
    Close #1
    Close #2
    
End Function

'函数名称:cyGetSerialFileName
'作用：根据当前时间返回流水文件名称（如：20060826145348），可作为目录或随机文件名
'===========================================================
'整理时间：2006-8-26 14:55
'  参数：
'===========================================================
Function cyGetSerialFileName() As String
    cyGetSerialFileName = Replace(Format(Now, "YYYY-MM-DD HH:MM:SS"), "-", "")
    cyGetSerialFileName = Replace(cyGetSerialFileName, ":", "")
    cyGetSerialFileName = Replace(cyGetSerialFileName, " ", "")
End Function
'
''函数名称:cyMoveFileEx
''作用：允许称动或重命名正在运行的程序（在同一盘符下允许移动，不同盘符不行）
''===========================================================
''整理时间：2006-10-27 1:22
''  参数：
''           sSourceFile --  欲移动的文件名
''           sTargetFile --  移动后的文件名
''===========================================================
'Sub cyFileMoveEx(ByVal sSourceFile As String, ByVal sTargetFile)
'
'    Const MOVEFILE_COPY_ALLOWED = &H2                   '如目标文件存在，则将其替换
'    Const MOVEFILE_REPLACE_EXISTING = &H1               '如移动到一个不同的卷，则复制文件并删除原来的文件
'    Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4             '移动操作在系统下次重新?动时正式进行。这样便可在Windows NT中改换系统文件
'    Call MoveFileEx(sSourceFile, sTargetFile, MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED Or MOVEFILE_DELAY_UNTIL_REBOOT)
'
'End Sub

'函数名称:cyGetFilesCount
'作用：获得目录下文件数量
'===========================================================
'整理时间：2006-12-31 11:29
'  参数：
'           sPath --  欲获得文件数量的路径

'  例子：    Debug.Print F.cyGetFilesCount("c:\windows")
'===========================================================
Function cyGetFilesCount(ByVal sPath As String) As Long
    Dim sFile As String
    
    '规范路径
    sPath = IIf(Right(sPath, 1) = "\", sPath, sPath & "\")
    
    '找出源目录内的第一个文件
    sFile = Dir(sPath)
    Do While sFile <> ""
        cyGetFilesCount = cyGetFilesCount + 1
        sFile = Dir()
    Loop

End Function

'函数名称:cyReSetFileDateTime
'作用：重设文件时间（1980-2099年）
'===========================================================
'整理时间：2006-12-31 13:35
'  参数：
'           sFileName --  欲设置文件时间的文件名
'           dDateTime --  欲设置的文件时间
'  例子：        Debug.Print F.cyReSetFileDateTime("E:\localver.ini", "1980-1-3 9:0:2")

'===========================================================
Function cyReSetFileDateTime(ByVal sFileName As String, ByVal dDateTime As Date) As Boolean

    Dim lngHandle       As Long
    Dim udtFileTime     As FILETIME
    Dim udtLocalTime     As FILETIME
    Dim udtSystemTime     As SYSTEMTIME
    
    '规范时间格式
    dDateTime = Format(dDateTime, "YYYY-MM-DD HH:MM:SS")
    
    '将文件时间写入结构
    udtSystemTime.wYear = Year(dDateTime)
    udtSystemTime.wMonth = Month(dDateTime)
    udtSystemTime.wDay = Day(dDateTime)
    udtSystemTime.wDayOfWeek = Weekday(dDateTime) - 1
    udtSystemTime.wHour = Hour(dDateTime)
    udtSystemTime.wMinute = Minute(dDateTime)
    udtSystemTime.wSecond = Second(dDateTime)
    udtSystemTime.wMilliseconds = 0
    
    '系统时间转换为本地时间
    SystemTimeToFileTime udtSystemTime, udtLocalTime
    '本地时间转换为文件时间
    LocalFileTimeToFileTime udtLocalTime, udtFileTime
    '得到文件句柄
    lngHandle = CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    '设置文件时间
    cyReSetFileDateTime = SetFileTime(lngHandle, udtFileTime, udtFileTime, udtFileTime)
    '关闭文件句柄
    CloseHandle lngHandle
    
End Function


'函数名称:cyFileSizeToMB
'作用：'将字节转换为MB/KB等形式
'===========================================================
'整理时间：2007-7-28 13:12
'  参数：
'           lFileSize --  文件的字节数

'  例子：            Debug.Print F.cyFileSizeToMB(59676896)
'                    5.68 MB

'===========================================================

Public Function cyFileSizeToMB(ByVal lFileSize As Long) As String
    Dim strSize As String * 128
    StrFormatByteSize lFileSize, strSize, 128
    cyFileSizeToMB = Left$(strSize, InStr(1, strSize, Chr$(0)) - 1)

End Function

