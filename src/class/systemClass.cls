VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "systemClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
 

'-----------------------------------------------------------------------------
'快速关机
'-----------------------------------------------------------------------------
Private Declare Function RtlAdjustPrivilege& Lib "ntdll" (ByVal Privilege&, ByVal Newvalue&, ByVal NewThread&, Oldvalue&)
Private Declare Function NtShutdownSystem& Lib "ntdll" (ByVal ShutdownAction&)
'-----------------------------------------------------------------------------
'快速关机
'-----------------------------------------------------------------------------


Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'-----------------------------------------------------------------------------
'设置进程为高优先级
'-----------------------------------------------------------------------------



Private Declare Function getusername Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'-----------------------------------------------------------------------------
'取得磁盘序列号
'-----------------------------------------------------------------------------
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const CREATE_NEW = 1
Private Const INVALID_HANDLE_VALUE = -1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const IDENTIFY_BUFFER_SIZE = 512
Private Const OUTPUT_DATA_SIZE = IDENTIFY_BUFFER_SIZE + 16

'GETVERSIONOUTPARAMS contains the data returned
'from the Get Driver Version function
Private Type GETVERSIONOUTPARAMS
   bVersion       As Byte 'Binary driver version.
   bRevision      As Byte 'Binary driver revision
   bReserved      As Byte 'Not used
   bIDEDeviceMap  As Byte 'Bit map of IDE devices
   fCapabilities  As Long 'Bit mask of driver capabilities
   dwReserved(3)  As Long 'For future use
End Type

'IDE registers
Private Type IDEREGS
   bFeaturesReg     As Byte 'Used for specifying SMART "commands"
   bSectorCountReg  As Byte 'IDE sector count register
   bSectorNumberReg As Byte 'IDE sector number register
   bCylLowReg       As Byte 'IDE low order cylinder value
   bCylHighReg      As Byte 'IDE high order cylinder value
   bDriveHeadReg    As Byte 'IDE drive/head register
   bCommandReg      As Byte 'Actual IDE command
   bReserved        As Byte 'reserved for future use - must be zero
End Type

'SENDCMDINPARAMS contains the input parameters for the
'Send Command to Drive function
Private Type SENDCMDINPARAMS
   cBufferSize     As Long     'Buffer size in bytes
   irDriveRegs     As IDEREGS  'Structure with drive register values.
   bDriveNumber    As Byte     'Physical drive number to send command to (0,1,2,3).
   bReserved(2)    As Byte     'Bytes reserved
   dwReserved(3)   As Long     'DWORDS reserved
   bBuffer()      As Byte      'Input buffer.
End Type

'Valid values for the bCommandReg member of IDEREGS.
Private Const IDE_ID_FUNCTION = &HEC            'Returns ID sector for ATA.
Private Const IDE_EXECUTE_SMART_FUNCTION = &HB0 'Performs SMART cmd.
                                                'Requires valid bFeaturesReg,
                                                'bCylLowReg, and bCylHighReg

'Cylinder register values required when issuing SMART command
Private Const SMART_CYL_LOW = &H4F
Private Const SMART_CYL_HI = &HC2

'Status returned from driver
Private Type DRIVERSTATUS
   bDriverError  As Byte          'Error code from driver, or 0 if no error
   bIDEStatus    As Byte          'Contents of IDE Error register
                                  'Only valid when bDriverError is SMART_IDE_ERROR
   bReserved(1)  As Byte
   dwReserved(1) As Long
 End Type

Private Type IDSECTOR
   wGenConfig                 As Integer
   wNumCyls                   As Integer
   wReserved                  As Integer
   wNumHeads                  As Integer
   wBytesPerTrack             As Integer
   wBytesPerSector            As Integer
   wSectorsPerTrack           As Integer
   wVendorUnique(2)           As Integer
   sSerialNumber(19)          As Byte
   wBufferType                As Integer
   wBufferSize                As Integer
   wECCSize                   As Integer
   sFirmwareRev(7)            As Byte
   sModelNumber(39)           As Byte
   wMoreVendorUnique          As Integer
   wDoubleWordIO              As Integer
   wCapabilities              As Integer
   wReserved1                 As Integer
   wPIOTiming                 As Integer
   wDMATiming                 As Integer
   wBS                        As Integer
   wNumCurrentCyls            As Integer
   wNumCurrentHeads           As Integer
   wNumCurrentSectorsPerTrack As Integer
   ulCurrentSectorCapacity    As Long
   wMultSectorStuff           As Integer
   ulTotalAddressableSectors  As Long
   wSingleWordDMA             As Integer
   wMultiWordDMA              As Integer
   bReserved(127)             As Byte
End Type

'Structure returned by SMART IOCTL commands
Private Type SENDCMDOUTPARAMS
  cBufferSize   As Long         'Size of Buffer in bytes
  DRIVERSTATUS  As DRIVERSTATUS 'Driver status structure
  bBuffer()    As Byte          'Buffer of arbitrary length for data read from drive
End Type

'Vendor specific feature register defines
'for SMART "sub commands"
Private Const SMART_ENABLE_SMART_OPERATIONS = &HD8

'Status Flags Values
Private Enum STATUS_FLAGS
   PRE_FAILURE_WARRANTY = &H1
   ON_LINE_COLLECTION = &H2
   PERFORMANCE_ATTRIBUTE = &H4
   ERROR_RATE_ATTRIBUTE = &H8
   EVENT_COUNT_ATTRIBUTE = &H10
   SELF_PRESERVING_ATTRIBUTE = &H20
End Enum

'IOCTL commands
Private Const DFP_GET_VERSION = &H74080
Private Const DFP_SEND_DRIVE_COMMAND = &H7C084
Private Const DFP_RECEIVE_DRIVE_DATA = &H7C088

Private Type ATTR_DATA
   AttrID As Byte
   AttrName As String
   AttrValue As Byte
   ThresholdValue As Byte
   WorstValue As Byte
   StatusFlags As STATUS_FLAGS
End Type

Private Type DRIVE_INFO
   bDriveType As Byte
   SerialNumber As String
   Model As String
   FirmWare As String
   Cilinders As Long
   Heads As Long
   SecPerTrack As Long
   BytesPerSector As Long
   BytesperTrack As Long
   NumAttributes As Byte
   Attributes() As ATTR_DATA
End Type

Private Enum IDE_DRIVE_NUMBER
   PRIMARY_MASTER
   PRIMARY_SLAVE
   SECONDARY_MASTER
   SECONDARY_SLAVE
   TERTIARY_MASTER
   TERTIARY_SLAVE
   QUARTIARY_MASTER
   QUARTIARY_SLAVE
End Enum

Private Declare Function CreateFile Lib "kernel32" _
   Alias "CreateFileA" _
  (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   lpSecurityAttributes As Any, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long
  
Private Declare Function DeviceIoControl Lib "kernel32" _
  (ByVal hDevice As Long, _
   ByVal dwIoControlCode As Long, _
   lpInBuffer As Any, _
   ByVal nInBufferSize As Long, _
   lpOutBuffer As Any, _
   ByVal nOutBufferSize As Long, _
   lpBytesReturned As Long, _
   lpOverlapped As Any) As Long
  
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (hpvDest As Any, _
   hpvSource As Any, _
   ByVal cbCopy As Long)
  
Private Type OSVERSIONINFO
   OSVSize As Long
   dwVerMajor As Long
   dwVerMinor As Long
   dwBuildNumber As Long
   PlatformID As Long
   szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" _
   Alias "GetVersionExA" _
  (LpVersionInformation As OSVERSIONINFO) As Long
'-----------------------------------------------------------------------------
'取得磁盘序列号
'-----------------------------------------------------------------------------



'-----------------------------------------------------------------------------
'运行为函数
'-----------------------------------------------------------------------------
Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Const CREATE_DEFAULT_ERROR_MODE = &H4000000
Private Const LOGON_WITH_PROFILE = &H1

Private Declare Function CreateProcessWithLogonW Lib "advapi32.dll" _
        (ByVal lpUserName As String, _
        ByVal lpDomain As String, _
        ByVal lpPassword As String, _
        ByVal dwLogonFlags As Long, _
        ByVal lpApplicationName As Long, _
        ByVal lpCommandLine As String, _
        ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, _
        ByVal lpCurrentDirectory As String, _
        ByRef lpStartupInfo As STARTUPINFO, _
        ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'-----------------------------------------------------------------------------
'运行为函数
'-----------------------------------------------------------------------------

Private Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'-----------------------------------------------------------------------------
'获得当前CPU利用率
'-----------------------------------------------------------------------------

Private Const SYSTEM_BASICINFORMATION = 0&
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&
Private Type LARGE_INTEGER
    dwLow As Long
    dwHigh As Long
End Type
Private Type SYSTEM_BASIC_INFORMATION
    dwUnknown1 As Long
    uKeMaximumIncrement As Long
    uPageSize As Long
    uMmNumberOfPhysicalPages As Long
    uMmLowestPhysicalPage As Long
    uMmHighestPhysicalPage As Long
    uAllocationGranularity As Long
    pLowestUserAddress As Long
    pMmHighestUserAddress As Long
    uKeActiveProcessors As Long
    bKeNumberProcessors As Byte
    bUnknown2 As Byte
    wUnknown3 As Integer
End Type
Private Type SYSTEM_PERFORMANCE_INFORMATION
    liIdleTime As LARGE_INTEGER
    dwSpare(0 To 75) As Long
End Type
Private Type SYSTEM_TIME_INFORMATION
    liKeBootTime As LARGE_INTEGER
    liKeSystemTime As LARGE_INTEGER
    liExpTimeZoneBias  As LARGE_INTEGER
    uCurrentTimeZoneId As Long
    dwReserved As Long
End Type
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, ByVal lpStructure As Long, ByVal dwSize As Long, ByVal dwReserved As Long) As Long
Private liOldIdleTime As LARGE_INTEGER
Private liOldSystemTime As LARGE_INTEGER
'-----------------------------------------------------------------------------
'获得当前CPU利用率
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'根据句柄得到其执行文件全路径
'-----------------------------------------------------------------------------
Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Byte
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 1024
End Type

Const TH32CS_SNAPmodule = &H8

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, lpme As Any) As Boolean
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Any) As Long
'-----------------------------------------------------------------------------
'根据句柄得到其执行文件全路径
'-----------------------------------------------------------------------------

'-----------------------------------------------------------------------------
'得到DOS输出
'-----------------------------------------------------------------------------

Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'Private Type PROCESS_INFORMATION
'    hProcess As Long
'    hThread As Long
'    dwProcessId As Long
'    dwThreadId As Long
'End Type

'Private Type STARTUPINFO
'    cb As Long
'    lpReserved As Long
'    lpDesktop As Long
'    lpTitle As Long
'    dwX As Long
'    dwY As Long
'    dwXSize As Long
'    dwYSize As Long
'    dwXCountChars As Long
'    dwYCountChars As Long
'    dwFillAttribute As Long
'    dwFlags As Long
'    wShowWindow As Integer
'    cbReserved2 As Integer
'    lpReserved2 As Byte
'    hStdInput As Long
'    hStdOutput As Long
'    hStdError As Long
'End Type

Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESTDHANDLES = &H100
Private Const SW_HIDE = 0

'-----------------------------------------------------------------------------
'得到DOS输出
'-----------------------------------------------------------------------------


Dim bStopTag As Boolean                                 '决定是否停止延时的标志
'Public cyUpdateMyself_ShowRunbar As Boolean          '决定更新时是否显示标志

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long

Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetMessageExtraInfo Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal ID As Long) As Long

Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As String, ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long) As Long


'-----------------------------------------------------------------------------
'更改当前屏幕大小,颜色值及分辨率
'-----------------------------------------------------------------------------
Private Declare Function EnumDisplaySettings _
    Lib "user32" Alias "EnumDisplaySettingsA" ( _
    ByVal lpszDeviceName As Long, _
    ByVal iModeNum As Long, _
    lpDevMode As Any) As Boolean
    
Private Declare Function ChangeDisplaySettings _
    Lib "user32" Alias "ChangeDisplaySettingsA" ( _
    lpDevMode As Any, _
    ByVal dwFlags As Long) As Long

Private Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const DM_DISPLAYFREQUENCY = &H400000
'-----------------------------------------------------------------------------
'更改当前屏幕大小,颜色值及分辨率
'-----------------------------------------------------------------------------

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
  (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
    ByVal lpProcName As String) As Long
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, _
   ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, _
   ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, _
    lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Dim mresult

Enum bMethod
    RegOcxDLL = 0
    unRegOcxDLL = 1
End Enum

'-----------------------------------------------------------------------------
'返回进程列表
'-----------------------------------------------------------------------------
'Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 1024
End Type

Const TH32CS_SNAPPROCESS = &H2
'-----------------------------------------------------------------------------
'返回进程列表
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'杀死进程
'-----------------------------------------------------------------------------
'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long


Private Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" _
                         (ByVal hWnd As Long, ByVal lpOperation As String, _
                          ByVal lpFile As String, ByVal lpParameters As String, _
                          ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_NORMAL = 1

'-----------------------------------------------------------------------------
'获取硬盘序列号
'-----------------------------------------------------------------------------
Private Declare Function HDSerialNumRead Lib "GetHDSN.dat" () As String



'-----------------------------------------------------------------------------
'关机及重启
'-----------------------------------------------------------------------------
Private Declare Function ExitWindowsEx Lib "user32" (ByVal UFlags As Long, ByVal dwReserved As Long) As Long

 Const EWX_SHUTDOWN = 1 '关机
 Const EWX_REBOOT = 2 '重启动
 Const EWX_POWEROFF = 8 '物理关闭电源
 Const EWX_FORCE = 4 '强制关机，即不通知现在活动应用程序让其先自我关闭

 Const TOKEN_ADJUST_PRIVILEGES = &H20
 Const TOKEN_QUERY = &H8
 Const SE_PRIVILEGE_ENABLED = &H2
 Const ANYSIZE_ARRAY = 1

Private Type LUID
    lowpart As Long
    highpart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
'-----------------------------------------------------------------------------
'关机及重启
'-----------------------------------------------------------------------------


Dim iStartTime As Long  '计算时间间隔时存放开始的时间

Enum cyKeybAction
    KB_MINWND = 1
    KB_PrintScreenToClipBoard = 2
    KB_PrintFormToClipBoard = 3
    KB_MiniAllForm = 4
    KB_PopupStartMenu = 5
    KB_PopupRun = 6
    KB_DisplayDeskTop = 7
    KB_PopupSearch = 8
    KB_LockComputer = 9
    KB_LostFocus = 10
    KB_MenuAltPress = 11
End Enum

Enum cyCDRomAction
    DoorOpen = 0
    DoorClose = 1
End Enum

Enum cyTimeAction
    TA_TimeSetStart = 0
    TA_TimeGetStart = 1
    TA_TimeComputingInterVal = 2
End Enum

Enum cyWindowVersion
    Win9x = -1073739260
    WinMe = -1073718780
    WinNT = 90505220
    Win2K = 143851525
    WinXp = 170393861
    Win2003 = 248381957
End Enum

Enum cyMouseAction
    LeftClick = 8
    LeftDblClick = 1
    LeftDown = 2
    LeftUp = 3
    RightClick = 4
    RightDblClick = 5
    RightDown = 6
    RightUp = 7
    GetPointColor = 9
    GetPointHwnd = 10
    CursorPosGet = 11
    CursorPosSet = 12
    CursorCapture = 13
    CursorRelease = 14
End Enum

Enum cyKillProcessMethod
    KillFromPID = 0
    KillFromExeName = 1
End Enum

'这个函数就是用于NT/2000系统中取得关机句柄,
'否则会出现关机选项而不会自动关机。
Private Sub AdjustTokenPrivilegesForNT()
    Dim hdlProcessHandle As Long
    Dim hdlTokenHandle As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    hdlProcessHandle = GetCurrentProcess()
    OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
    TOKEN_QUERY), hdlTokenHandle
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    tkp.PrivilegeCount = 1
    tkp.Privileges(0).pLuid = tmpLuid
    tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    
    AdjustTokenPrivileges hdlTokenHandle, False, tkp, _
    Len(tkpNewButIgnored), tkpNewButIgnored, _
    lBufferNeeded
End Sub

'过程名称:cyDelay

'作用：精确延时[准确到千分之十秒]

'##################################################
'整理时间：2004-2-13 18:22:00
'  参数：
'   singleSecond As Single  --要延时的秒数[如:5.3]
'##################################################
'  例：
Public Sub cyDelay(ByVal singleSecond As Single)
    Dim iStopTime As Long
    iStopTime = GetTickCount + singleSecond * 1000
    bStopTag = False
    Do While GetTickCount < iStopTime '一直运行直至到达时间
        '设置为停止
        If bStopTag = True Then GoTo StopTag
        DoEvents
    Loop
StopTag:
    bStopTag = False
End Sub

'设置为终止延时
Sub cyDelayTerminate()
    bStopTag = True
End Sub

'函数名称:cyTimeAction

'作用：精确计算时间差

'##################################################
'整理时间：2004-2-13 18:26:58
'  参数：
'   Action As cyTimeAction  --
'##################################################
'  例：
Function cyTimeAction(Optional ByVal Action As cyTimeAction) As Long
    Select Case Action
        Case TA_TimeSetStart
            iStartTime = GetTickCount
        Case TA_TimeGetStart
            cyTimeAction = iStartTime
        Case TA_TimeComputingInterVal
            cyTimeAction = GetTickCount - iStartTime
    End Select
End Function

Function cyMouseAction(Optional Action As cyMouseAction, Optional ByRef X As Long, Optional ByRef Y As Long, Optional bReturnSourcePoint As Boolean, Optional ByVal hWnd As Long, Optional ByVal sString As String)
    Dim cyRect As RECT
    
    Const MOUSEEVENTF_LEFTDOWN = &H2
    Const MOUSEEVENTF_LEFTUP = &H4
    Const MOUSEEVENTF_RIGHTDOWN = &H8
    Const MOUSEEVENTF_RIGHTUP = &H10
    
    Dim oldPos As POINTAPI
    
    '保存旧的光标位置
    GetCursorPos oldPos
    
Select Case Action
    Case 8
        '如果给定了光标位置则先将光标锁定在该位置
        If X <> 0 Or Y <> 0 Then
            cyRect.Left = X
            cyRect.Top = Y
            cyRect.Bottom = Y
            cyRect.Right = X
            
            Call GetWindowRect(hWnd, cyRect)
            Call ClipCursor(cyRect)
            
        End If
    
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, GetMessageExtraInfo())
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, GetMessageExtraInfo())
    
        '释放鼠标锁定范围
        Call ClipCursor(ByVal vbNullString)
            
    Case 1
        '如果给定了光标位置则先将光标锁定在该位置
        If X <> 0 Or Y <> 0 Then
            cyRect.Left = X
            cyRect.Top = Y
            cyRect.Bottom = Y
            cyRect.Right = X
            
            Call GetWindowRect(hWnd, cyRect)
            Call ClipCursor(cyRect)
            
        End If
        
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, GetMessageExtraInfo())
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, GetMessageExtraInfo())

        Call GetWindowRect(hWnd, cyRect)
        Call ClipCursor(cyRect)

        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, GetMessageExtraInfo())
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, GetMessageExtraInfo())
    
        '释放鼠标锁定范围
        Call ClipCursor(ByVal vbNullString)
            
    Case 2
        '如果给定了光标位置则先将光标锁定在该位置
        If X <> 0 Or Y <> 0 Then
            cyRect.Left = X
            cyRect.Top = Y
            cyRect.Bottom = Y
            cyRect.Right = X
            
            Call GetWindowRect(hWnd, cyRect)
            Call ClipCursor(cyRect)
            
        End If
        
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, GetMessageExtraInfo())
    
        '释放鼠标锁定范围
        Call ClipCursor(ByVal vbNullString)
            
    Case 3
        '如果给定了光标位置则先将光标锁定在该位置
        If X <> 0 Or Y <> 0 Then
            cyRect.Left = X
            cyRect.Top = Y
            cyRect.Bottom = Y
            cyRect.Right = X
            
            Call GetWindowRect(hWnd, cyRect)
            Call ClipCursor(cyRect)
            
        End If
        
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, GetMessageExtraInfo())
    
        '释放鼠标锁定范围
        Call ClipCursor(ByVal vbNullString)
            
    Case 4
        '如果给定了光标位置则先将光标锁定在该位置
        If X <> 0 Or Y <> 0 Then
            cyRect.Left = X
            cyRect.Top = Y
            cyRect.Bottom = Y
            cyRect.Right = X
            
            Call GetWindowRect(hWnd, cyRect)
            Call ClipCursor(cyRect)
            
        End If
        
        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, GetMessageExtraInfo())
        Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, GetMessageExtraInfo())
        
        '释放鼠标锁定范围
        Call ClipCursor(ByVal vbNullString)
            
    Case 5
        '如果给定了光标位置则先将光标锁定在该位置
        If X <> 0 Or Y <> 0 Then
            cyRect.Left = X
            cyRect.Top = Y
            cyRect.Bottom = Y
            cyRect.Right = X
            
            Call GetWindowRect(hWnd, cyRect)
            Call ClipCursor(cyRect)
            
        End If
        
        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, GetMessageExtraInfo())
        Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, GetMessageExtraInfo())
        
        Call GetWindowRect(hWnd, cyRect)
        Call ClipCursor(cyRect)
            
        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, GetMessageExtraInfo())
        Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, GetMessageExtraInfo())
    
        '释放鼠标锁定范围
        Call ClipCursor(ByVal vbNullString)
            
    Case 6
        '如果给定了光标位置则先将光标锁定在该位置
        If X <> 0 Or Y <> 0 Then
            cyRect.Left = X
            cyRect.Top = Y
            cyRect.Bottom = Y
            cyRect.Right = X
            
            Call GetWindowRect(hWnd, cyRect)
            Call ClipCursor(cyRect)
            
        End If
        
        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, GetMessageExtraInfo())
        
        '释放鼠标锁定范围
        Call ClipCursor(ByVal vbNullString)
            
    Case 7
        '如果给定了光标位置则先将光标锁定在该位置
        If X <> 0 Or Y <> 0 Then
            cyRect.Left = X
            cyRect.Top = Y
            cyRect.Bottom = Y
            cyRect.Right = X
            
            Call GetWindowRect(hWnd, cyRect)
            Call ClipCursor(cyRect)
            
        End If
        
        Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, GetMessageExtraInfo())
    
        '释放鼠标锁定范围
        Call ClipCursor(ByVal vbNullString)
    
    Case GetPointColor
        '如果给定了光标位置则先设置其位置
        If X > 0 Or Y > 0 Then
            cyMouseAction = GetPixel(GetWindowDC(0), X, Y)
        Else
            cyMouseAction = GetPixel(GetWindowDC(0), oldPos.X, oldPos.Y)
        End If
        Exit Function
        
    Case GetPointHwnd
        '如果给定了光标位置则先设置其位置
        If X > 0 Or Y > 0 Then
            cyMouseAction = WindowFromPoint(X, Y)
        Else
            cyMouseAction = WindowFromPoint(oldPos.X, oldPos.Y)
        End If
        Exit Function

    Case CursorPosGet
        X = oldPos.X
        Y = oldPos.Y
    
    Case CursorPosSet
        SetCursorPos X, Y
    
    Case CursorCapture
        On Error Resume Next
        
        '如果给定了光标位置则先将光标锁定在该位置
        If X <> 0 Or Y <> 0 Then
            cyRect.Left = X
            cyRect.Top = Y
            cyRect.Bottom = Y
            cyRect.Right = X
            
            Call GetWindowRect(hWnd, cyRect)
            Call ClipCursor(cyRect)
            
        End If
    
    Case CursorRelease
        '释放鼠标锁定范围
        Call ClipCursor(ByVal vbNullString)

End Select
        
    '完成操作后是否返回之前的鼠标位置
    If bReturnSourcePoint Then SetCursorPos oldPos.X, oldPos.Y
End Function

'函数名称:cyExpressionComputing

'作用：根据文本表达式计算数值

'##################################################
'整理时间：2004-2-13 18:29:25
'  参数：
'   sExpression As String    --表达式如:5+3+2*15
'   iFloatBit As Byte       --保留小数点位数
'##################################################
'  例：
Public Function cyExpressionComputing(ByVal sExpression As String, Optional ByVal iFloatBit As Byte)
    Dim Vbs
    Set Vbs = CreateObject("MSScriptControl.ScriptControl")
    Vbs.Language = "vbs"
    Dim i As Long
    Dim sFormat As String
    
    If iFloatBit > 0 Then
        For i = 1 To iFloatBit
            sFormat = sFormat + "0"
        Next i
        cyExpressionComputing = Format(Vbs.Eval(sExpression), "0." + sFormat)
    Else
        cyExpressionComputing = CLng(Vbs.Eval(sExpression))
    End If
    
End Function

'函数名称:cyRnd

'作用：在某个数值范围内取得随机值

'##################################################
'整理时间：2004-2-13 18:30:19
'  参数：
'   iMin                    --下限
'   iMax                    --上限
'   iFloatBit As Byte       --保留小数位数
'##################################################
'  例：
Public Function cyRnd(ByVal iMin, ByVal iMax, Optional iFloatBit As Byte)
    Dim sFormat As String
    
    Randomize
    
    If TypeName(iMin) = "Double" Or TypeName(iMax) = "Double" Then '如果有小数
        If iFloatBit > 0 Then '有小数要求
            Dim i As Byte
            For i = 1 To iFloatBit
              sFormat = sFormat + "0"
            Next i
            If iFloatBit > 0 Then cyRnd = Format((iMax - iMin) * Rnd + iMin, "0." + sFormat)
        Else
            cyRnd = CDbl((iMax - iMin) * Rnd + iMin)
        End If
    
    Else '整型
        cyRnd = (CLng((iMax - iMin) * Rnd + iMin))
    End If
End Function

'过程名称:cyKeyBoardAction

'作用：控制键盘的一系列操作

'##################################################
'整理时间：2004-2-13 18:31:03
'  参数：
'   Action As cyKeybAction = -1 --要完成的动作
'   Key1 As KeyCodeConstants = -1   --第一个按键
'   Key2 As KeyCodeConstants = -1   --第二个按键
'   Key3 As KeyCodeConstants = -1   --第三个按键
'##################################################
'  例：
Sub cyKeyBoardAction(Optional Action As cyKeybAction = -1, Optional ByVal Key1 As KeyCodeConstants = -1, Optional ByVal Key2 As KeyCodeConstants = -1, Optional ByVal Key3 As KeyCodeConstants = -1)
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
    If Action <> -1 Then        '有特殊动作
        Select Case Action
            Case KB_MINWND
                ' Hold the Alt key while typing Space.
                keybd_event vbKeyMenu, 0, 0, 0  ' press Alt
                keybd_event vbKeySpace, 0, 0, 0  ' press Space
                keybd_event vbKeySpace, 0, KEYEVENTF_KEYUP, 0  ' release Space
                keybd_event vbKeyMenu, 0, KEYEVENTF_KEYUP, 0  ' release Alt
                ' Type the N key.
                keybd_event vbKeyN, 0, 0, 0  ' press N
                keybd_event vbKeyN, 0, KEYEVENTF_KEYUP, 0  ' release N
                
            Case KB_PrintScreen2ClipBoard
                keybd_event vbKeySnapshot, 0, 0&, 0&
                
            Case KB_PrintActiveForm2ClipBoard
                keybd_event vbKeySnapshot, 1, 0&, 0&
            
            Case KB_MiniAllForm
                keybd_event &H5B, 0, 0, 0  ' press Left Wnd
                keybd_event vbKeyM, 0, 0, 0  ' press m
                keybd_event vbKeyM, 0, KEYEVENTF_KEYUP, 0  ' release m
                keybd_event &H5B, 0, KEYEVENTF_KEYUP, 0  ' release Left Wnd
            Case KB_PopupStartMenu
                keybd_event &H5B, 0, 0, 0  ' press Left Wnd
                keybd_event &H5B, 0, KEYEVENTF_KEYUP, 0  ' release Left Wnd
                
            Case KB_PopupRun
                keybd_event &H5B, 0, 0, 0  ' press Left Wnd
                keybd_event vbKeyR, 0, 0, 0  ' press m
                keybd_event vbKeyR, 0, KEYEVENTF_KEYUP, 0  ' release m
                keybd_event &H5B, 0, KEYEVENTF_KEYUP, 0  ' release Left Wnd
            Case KB_DisplayDeskTop
                keybd_event &H5B, 0, 0, 0  ' press Left Wnd
                keybd_event vbKeyD, 0, 0, 0  ' press m
                keybd_event vbKeyD, 0, KEYEVENTF_KEYUP, 0  ' release m
                keybd_event &H5B, 0, KEYEVENTF_KEYUP, 0  ' release Left Wnd
            Case KB_PopupSearch
                keybd_event &H5B, 0, 0, 0  ' press Left Wnd
                keybd_event vbKeyF, 0, 0, 0  ' press m
                keybd_event vbKeyF, 0, KEYEVENTF_KEYUP, 0  ' release m
                keybd_event &H5B, 0, KEYEVENTF_KEYUP, 0  ' release Left Wnd
            Case KB_LockComputer
                keybd_event &H5B, 0, 0, 0  ' press Left Wnd
                keybd_event vbKeyL, 0, 0, 0  ' press m
                keybd_event vbKeyL, 0, KEYEVENTF_KEYUP, 0  ' release m
                keybd_event &H5B, 0, KEYEVENTF_KEYUP, 0  ' release Left Wnd
            Case KB_LostFocus
                keybd_event &H5B, 0, 0, 0  ' press Left Wnd
                keybd_event vbKeyB, 0, 0, 0  ' press m
                keybd_event vbKeyB, 0, KEYEVENTF_KEYUP, 0  ' release m
                keybd_event &H5B, 0, KEYEVENTF_KEYUP, 0  ' release Left Wnd
            Case KB_MenuAltPress
                '按下alt键
                keybd_event 18, 0, 0, 0  ' press Left Wnd
                
                If Key1 <> -1 Then
                    keybd_event Key1, 0, 0, 0  ' press Left Wnd
                    keybd_event Key1, 0, KEYEVENTF_KEYUP, 0  ' press Left Wnd
                End If
                If Key2 <> -1 Then
                    keybd_event Key2, 0, 0, 0  ' press Left Wnd
                    keybd_event Key2, 0, KEYEVENTF_KEYUP, 0  ' press Left Wnd
                End If
                If Key3 <> -1 Then
                    keybd_event Key3, 0, 0, 0  ' press Left Wnd
                    keybd_event Key3, 0, KEYEVENTF_KEYUP, 0  ' release m
                End If
                keybd_event 18, 0, KEYEVENTF_KEYUP, 0  ' press Left Wnd
        End Select
    Else            '没特殊动作,用组合键
    
        Dim iKey1 As Byte
        Dim iKey2 As Byte
        Dim iKey3 As Byte
        
        If Key1 <> -1 Then
            iKey1 = Key1
            keybd_event iKey1, 0, 0, 0  ' press Left Wnd
        End If
        If Key2 <> -1 Then
            iKey2 = Key2
            keybd_event iKey2, 0, 0, 0  ' press Left Wnd
        End If
        If Key3 <> -1 Then
            iKey3 = Key3
            keybd_event iKey3, 0, 0, 0  ' press Left Wnd
            keybd_event iKey3, 0, KEYEVENTF_KEYUP, 0  ' release m
        End If
        If Key2 <> -1 Then
            iKey2 = Key2
            keybd_event iKey2, 0, KEYEVENTF_KEYUP, 0  ' press Left Wnd
        End If
        If Key1 <> -1 Then
            iKey1 = Key1
            keybd_event iKey1, 0, KEYEVENTF_KEYUP, 0  ' press Left Wnd
        End If

    End If
End Sub

'函数名称:cyCheckVersion

'作用：查找当前使用的系统

'##################################################
'整理时间：2004-2-13 18:32:07
'  参数：
'   isVersion As cyWindowVersion    --系统
'##################################################
'  例：
Function cyCheckVersion(isVersion As cyWindowVersion) As Boolean
    Dim Ver As Long
    Ver = GetVersion()
    Select Case Ver
        Case Win9x
            If Ver = isVersion Then cyCheckVersion = True
        Case WinMe
            If Ver = isVersion Then cyCheckVersion = True
        Case WinNT
            If Ver = isVersion Then cyCheckVersion = True
        Case Win2K
            If Ver = isVersion Then cyCheckVersion = True
        Case WinXp
            If Ver = isVersion Then cyCheckVersion = True
        Case Win2003
            If Ver = isVersion Then cyCheckVersion = True
    End Select
End Function


Function cyCDRomAction(Action As cyCDRomAction) As Boolean
    Select Case Action
        Case 0
            Call mciSendString("set CDAudio door open", vbNull, 127, 0)
        Case 1
            Call mciSendString("set CDAudio door closed", vbNull, 127, 0)
    End Select
End Function
Sub cyShutDown()
    AdjustTokenPrivilegesForNT
    
    '在95/98中调用没作用，取得关机句柄
    If GetVersion = 90505220 Then ' NT
        ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
    Else
        ExitWindowsEx EWX_SHUTDOWN Or EWX_POWEROFF Or EWX_FORCE, 0
    End If
    '如果只用 "ExitWindowsEx EWX_SHUTDOWN , 0",则不会物理关机而只会出现“您可以安全关机”之类的提示，呵呵!
End Sub

'函数名称:cyShutDownQuickMode
'作用：快速关机
'===========================================================
'整理时间：2007-06-02 18:35
'  参数：
'  例子：        call cyShutDownQuickMode

'===========================================================
Sub cyShutDownQuickMode()
    Const SE_SHUTDOWN_PRIVILEGE& = 19
    Const SHUTDOWN& = 0
    RtlAdjustPrivilege SE_SHUTDOWN_PRIVILEGE, 1, 0, 0
    NtShutdownSystem SHUTDOWN

End Sub

Sub cyReBoot()
    AdjustTokenPrivilegesForNT
    ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End Sub

Sub cyShell(sCommand As String, iWndStyle As VbAppWinStyle)
On Error GoTo errXp
    Shell sCommand, iWndStyle
    Exit Sub
errXp:
On Error GoTo err98
    Shell "cmd /c " & sCommand, iWndStyle
    Exit Sub
err98:
    Shell "Command /c " & sCommand, iWndStyle
End Sub

Sub cyShellExecute(sCommand As String)
    ShellExecute 0, vbNullString, _
                 sCommand, vbNullString, _
                 "C:\", SW_NORMAL
End Sub

 '函数名称:cyKillProcess

'作用：根据进程ID(Pid)杀死进程

'##################################################
'整理时间：2004-2-13 15:23:28
'  参数：
'   Pid As Long             --进程ID
'##################################################
'  例： Debug.Print cyKillProcess(336)
Function cyKillProcess(ByVal KillMethod As cyKillProcessMethod, Optional ByVal vParameter) As Boolean
    Dim sTemp As String
    Dim bStr As String * 255
    Dim Pid As Long
    Dim bArray() As Byte
    Dim F As New fileClass

    '防止传入错误的参数
    If vParameter = 0 Or vParameter = "" Then Exit Function

    If KillMethod = KillFromPID Then
    '根据PID关闭程序
        Dim hProcess As Long
        Const PROCESS_TERMINATE = &H1
        hProcess = OpenProcess(PROCESS_TERMINATE, 0, vParameter)
        
        If TerminateProcess(hProcess, 0) <> 0 Then
            cyKillProcess = True
            
        Else
            '得到临时目录
            sTemp = Mid(bStr, 1, GetTempPath(255, bStr))
            sTemp = IIf(Right(sTemp, 1) = "\", sTemp, sTemp & "\")
            
            On Error Resume Next
            '获放该主文件为临时文件
            bArray = LoadResData(105, "CUSTOM") '将自定义资源中103号资源读入数组OCX
            Open sTemp & "KillProcessAgent.Exe" For Binary As #1
                Seek #1, 1
                Put #1, , bArray
            Close #1
            WinExec sTemp & "KillProcessAgent.exe """ & F.cyCutFileName(cyGetAppNameFromPid(vParameter), CutFileNameOnly) & """", vbHide
            
            cyDelay 0.3
            If CreateToolhelp32Snapshot(&H8, vParameter) = -1 Then cyKillProcess = True
                    
        End If
        
        Call CloseHandle(hProcess)
    Else
    '根据EXE关闭程序
        Dim sA() As String
        Dim i As Long
        sA = cyGetProcessToArray
        For i = 0 To UBound(sA)
            If UCase((Split(sA(i), "#")(1))) = UCase(vParameter) Then
                hProcess = OpenProcess(PROCESS_TERMINATE, 0, CLng(Split(sA(i), vbTab)(0)))
                If TerminateProcess(hProcess, 0) <> 0 Then
                    cyKillProcess = True
                    Exit Function
                    
                Else
                    '得到临时目录
                    sTemp = Mid(bStr, 1, GetTempPath(255, bStr))
                    sTemp = IIf(Right(sTemp, 1) = "\", sTemp, sTemp & "\")
On Error Resume Next
                    bArray = LoadResData(105, "CUSTOM") '将自定义资源中103号资源读入数组OCX
                    Open sTemp & "KillProcessAgent.Exe" For Binary As #1
                        Seek #1, 1
                        Put #1, , bArray
                    Close #1
                    
                    WinExec sTemp & "KillProcessAgent.exe """ & F.cyCutFileName(vParameter, CutFileNameOnly) & """", vbHide
            
                    cyDelay 0.3
                    If CreateToolhelp32Snapshot(&H8, vParameter) = -1 Then
                        cyKillProcess = True
                        Exit Function
                        
                    End If
                
                End If
                
            End If
        
        Next

    End If
    
End Function

Sub cyClearProcess(Optional ByVal sKeepList As String)
    Dim sTemp As String
    Dim bStr As String * 255
    sTemp = Mid(bStr, 1, GetTempPath(255, bStr))
    sTemp = IIf(Right(sTemp, 1) = "\", sTemp, sTemp & "\")
    
    On Error Resume Next
    '获放该主文件为临时文件
    Dim bArray() As Byte
    bArray = LoadResData(105, "CUSTOM") '将自定义资源中103号资源读入数组OCX
    Open sTemp & "Temp.exe" For Binary As #1
        Seek #1, 1
        Put #1, , bArray
    Close #1

    Dim sA() As String
    Dim sA1() As String
    Dim sA2() As String
    Dim sStr As String
    
    '可信进程列表
    sStr = ";[SYSTEM PROCESS];ALG.EXE;CSRSS.EXE;DDHELP.EXE;DLLHOST.EXE;INETINFO.EXE;INTERNAT.EXE;KERNEL32.DLL;LSASS.EXE;MDM.EXE;MMTASK.TSK;MPREXE.EXE;MSGSRV32.EXE;MSTASK.EXE;REGSVC.EXE;RPCSS.EXE;SERVICES.EXE;SMSS.EXE;SNMP.EXE;SPOOL32.EXE;SPOOLSV.EXE;STISVC.EXE;SVCHOST.EXE;SYSTEM;TASKMON.EXE;TCPSVCS.EXE;WINLOGON.EXE;WINMGMT.EXE;SURFER.EXE;EXPLORER.EXE;" & UCase(sKeepList) & ";"
    
    Dim Process As PROCESSENTRY32
    Dim j As Long
    j = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    Dim i As Long
    If j Then
        Process.dwSize = 1060
        If (Process32First(j, Process)) Then '遍历第一个进程
            Do
            
                Dim hProcess As Long
                Const PROCESS_TERMINATE = &H1
            
                '如果找到的进程不在必要保留的进程之内时
                If InStr(1, sStr, UCase(Left(Process.szExeFile, InStr(1, Process.szExeFile, Chr(0)) - 1))) = 0 Then
                    hProcess = OpenProcess(PROCESS_TERMINATE, 0, Process.th32ProcessID)
                    TerminateProcess hProcess, 0
                    WinExec sTemp & "Temp.exe " & UCase(Left(Process.szExeFile, InStr(1, Process.szExeFile, Chr(0)) - 1)), vbHide
                    Call CloseHandle(hProcess)
                End If
            Loop Until (Process32Next(j, Process) < 1)
        End If
        CloseHandle j
    End If

End Sub
Public Function cyGetProcessToArray()
    Dim sArray() As String
    Dim Process As PROCESSENTRY32
    Dim j As Long
    j = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    Dim i As Long
    If j Then
        Process.dwSize = 1060
        If (Process32First(j, Process)) Then '遍历第一个进程
            Do
                ReDim Preserve sArray(i)
                sArray(i) = Process.th32ProcessID & vbTab & "#" & Left(Process.szExeFile, InStr(1, Process.szExeFile, Chr(0)) - 1)
                i = i + 1
            Loop Until (Process32Next(j, Process) < 1)
        End If
        CloseHandle j
        cyGetProcessToArray = sArray
    End If
End Function

'函数名称:cyGetPIDFromExeName
'作用：根据程序的文件名取得PID
'===========================================================
'整理时间：2006-8-3 23:33
'  参数：
    'ByVal sExeName As String                 --程序的文件名
'===========================================================
Public Function cyGetPidFromExeName(ByVal sExeName As String) As Long
    Dim Process As PROCESSENTRY32
    Dim j As Long
    
    j = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If j Then
    Process.dwSize = 1060
        If (Process32First(j, Process)) Then '遍历第一个进程
          Do
            If UCase(sExeName) = UCase(Left(Process.szExeFile, InStr(1, Process.szExeFile, Chr(0)) - 1)) Then
                cyGetPidFromExeName = Process.th32ProcessID
                CloseHandle j
                Exit Function
            End If
          Loop Until (Process32Next(j, Process) < 1)
        End If
    CloseHandle j
    End If
End Function

Public Sub cySetMousePointer(ByVal sFileName As String)
    Dim Cursor As Long
    Cursor = LoadCursorFromFile(sFileName)
    Screen.MousePointer = 6
    SetSystemCursor Cursor, 32643
End Sub

Public Function cyRegSvr32Ex(ByVal sFileName As String, Optional Method As bMethod) As Boolean
    On Error Resume Next
    Dim lLib As Long
    Dim lpDLLEntryPoint As Long
    Dim lpThreadID As Long
    Dim lpExitCode As Long
    Dim mThread As Long
    
    lLib = LoadLibrary(sFileName)
    '没有找到文件或文件被占用
    If lLib = 0 Then
        Exit Function
    End If
    
    If Method = 0 Then
        lpDLLEntryPoint = GetProcAddress(lLib, "DllRegisterServer")
    ElseIf Method = 1 Then
        lpDLLEntryPoint = GetProcAddress(lLib, "DllUnregisterServer")
    Else
        '找不到接入点则退出
        Exit Function
    End If
    
    If lpDLLEntryPoint = vbNull Or lpDLLEntryPoint = 0 Then
        GoTo earlyExit1
    End If
      
    mThread = CreateThread(ByVal 0, 0, ByVal lpDLLEntryPoint, ByVal 0, 0, lpThreadID)
    If mThread = 0 Then
        GoTo earlyExit1
    End If
    
    mresult = WaitForSingleObject(mThread, 10000)
    If mresult <> 0 Then
        GoTo earlyExit2
    End If
    
    CloseHandle mThread
    FreeLibrary lLib
    cyRegSvr32Ex = True
    Exit Function

earlyExit1:
    FreeLibrary lLib
    Exit Function
earlyExit2:
    FreeLibrary lLib
    lpExitCode = GetExitCodeThread(mThread, lpExitCode)
    ExitThread lpExitCode

End Function

'自删除
Sub cyKillMySelf()
    
    '保存当前程序的PID
    SaveSetting "cyDLL", "KillMySelf", "Pid", GetCurrentProcessId
    SaveSetting "cyDLL", "KillMySelf", "ExeName", cyGetAppName
        
    Dim sTemp As String
    Dim bStr As String * 255
    sTemp = Mid(bStr, 1, GetTempPath(255, bStr))
    sTemp = IIf(Right(sTemp, 1) = "\", sTemp, sTemp & "\")
    
    On Error Resume Next
    '获放该主文件为临时文件
    Dim bArray() As Byte
    bArray = LoadResData(104, "CUSTOM") '将自定义资源中103号资源读入数组OCX
    Open sTemp & "EndMySelf.exe" For Binary As #104
        Seek #104, 1
        Put #104, , bArray
    Close #104
    WinExec sTemp & "EndMySelf.exe", 0
    
End Sub

'改变分辨率
'Call cyChangeRes(1024, 768, 32, 85) '85是刷新率
Public Sub cyChangeRes(iWidthPixel As Long, _
                      iHeightPixel As Long, _
                      Optional ByVal iColorBit As Integer, _
                      Optional ByVal iRefreshRate As Long)
    Dim dm As DEVMODE
    
    If iColorBit = 0 Then iColorBit = 16
    If iRefreshRate = 0 Then iRefreshRate = 75
    
    Call EnumDisplaySettings(0&, -1, dm)
    
    dm.dmFields = DM_PELSWIDTH Or _
                  DM_PELSHEIGHT Or _
                  DM_BITSPERPEL Or _
                  DM_DISPLAYFREQUENCY
    dm.dmPelsWidth = iWidthPixel
    dm.dmPelsHeight = iHeightPixel
    dm.dmBitsPerPel = iColorBit
    dm.dmDisplayFrequency = iRefreshRate
    Call ChangeDisplaySettings(dm, 1)
End Sub

'将公历转换为农历
Function cyYearConvert(dDateToConvert As Date) As String

Dim WeekName(7), MonthAdd(11), NongliData(99), TianGan(9), DiZhi(11), ShuXiang(11), DayName(30), MonName(12)
Dim curTime, curYear, curMonth, curDay, curWeekday
Dim GongliStr, WeekdayStr, NongliStr, NongliDayStr
Dim i, m, n, k, isEnd, bit, TheDate

'获取当前系统时间
curTime = dDateToConvert

'星期名
WeekName(0) = " * "
WeekName(1) = "星期日"
WeekName(2) = "星期一"
WeekName(3) = "星期二"
WeekName(4) = "星期三"
WeekName(5) = "星期四"
WeekName(6) = "星期五"
WeekName(7) = "星期六"

'天干名称
TianGan(0) = "甲"
TianGan(1) = "乙"
TianGan(2) = "丙"
TianGan(3) = "丁"
TianGan(4) = "戊"
TianGan(5) = "己"
TianGan(6) = "庚"
TianGan(7) = "辛"
TianGan(8) = "壬"
TianGan(9) = "癸"

'地支名称
DiZhi(0) = "子"
DiZhi(1) = "丑"
DiZhi(2) = "寅"
DiZhi(3) = "卯"
DiZhi(4) = "辰"
DiZhi(5) = "巳"
DiZhi(6) = "午"
DiZhi(7) = "未"
DiZhi(8) = "申"
DiZhi(9) = "酉"
DiZhi(10) = "戌"
DiZhi(11) = "亥"

'属相名称
ShuXiang(0) = "鼠"
ShuXiang(1) = "牛"
ShuXiang(2) = "虎"
ShuXiang(3) = "兔"
ShuXiang(4) = "龙"
ShuXiang(5) = "蛇"
ShuXiang(6) = "马"
ShuXiang(7) = "羊"
ShuXiang(8) = "猴"
ShuXiang(9) = "鸡"
ShuXiang(10) = "狗"
ShuXiang(11) = "猪"

'农历日期名
DayName(0) = "*"
DayName(1) = "初一"
DayName(2) = "初二"
DayName(3) = "初三"
DayName(4) = "初四"
DayName(5) = "初五"
DayName(6) = "初六"
DayName(7) = "初七"
DayName(8) = "初八"
DayName(9) = "初九"
DayName(10) = "初十"
DayName(11) = "十一"
DayName(12) = "十二"
DayName(13) = "十三"
DayName(14) = "十四"
DayName(15) = "十五"
DayName(16) = "十六"
DayName(17) = "十七"
DayName(18) = "十八"
DayName(19) = "十九"
DayName(20) = "二十"
DayName(21) = "廿一"
DayName(22) = "廿二"
DayName(23) = "廿三"
DayName(24) = "廿四"
DayName(25) = "廿五"
DayName(26) = "廿六"
DayName(27) = "廿七"
DayName(28) = "廿八"
DayName(29) = "廿九"
DayName(30) = "三十"

'农历月份名
MonName(0) = "*"
MonName(1) = "正"
MonName(2) = "二"
MonName(3) = "三"
MonName(4) = "四"
MonName(5) = "五"
MonName(6) = "六"
MonName(7) = "七"
MonName(8) = "八"
MonName(9) = "九"
MonName(10) = "十"
MonName(11) = "十一"
MonName(12) = "腊"

'公历每月前面的天数
MonthAdd(0) = 0
MonthAdd(1) = 31
MonthAdd(2) = 59
MonthAdd(3) = 90
MonthAdd(4) = 120
MonthAdd(5) = 151
MonthAdd(6) = 181
MonthAdd(7) = 212
MonthAdd(8) = 243
MonthAdd(9) = 273
MonthAdd(10) = 304
MonthAdd(11) = 334

'农历数据
NongliData(0) = 2635
NongliData(1) = 333387
NongliData(2) = 1701
NongliData(3) = 1748
NongliData(4) = 267701
NongliData(5) = 694
NongliData(6) = 2391
NongliData(7) = 133423
NongliData(8) = 1175
NongliData(9) = 396438
NongliData(10) = 3402
NongliData(11) = 3749
NongliData(12) = 331177
NongliData(13) = 1453
NongliData(14) = 694
NongliData(15) = 201326
NongliData(16) = 2350
NongliData(17) = 465197
NongliData(18) = 3221
NongliData(19) = 3402
NongliData(20) = 400202
NongliData(21) = 2901
NongliData(22) = 1386
NongliData(23) = 267611
NongliData(24) = 605
NongliData(25) = 2349
NongliData(26) = 137515
NongliData(27) = 2709
NongliData(28) = 464533
NongliData(29) = 1738
NongliData(30) = 2901
NongliData(31) = 330421
NongliData(32) = 1242
NongliData(33) = 2651
NongliData(34) = 199255
NongliData(35) = 1323
NongliData(36) = 529706
NongliData(37) = 3733
NongliData(38) = 1706
NongliData(39) = 398762
NongliData(40) = 2741
NongliData(41) = 1206
NongliData(42) = 267438
NongliData(43) = 2647
NongliData(44) = 1318
NongliData(45) = 204070
NongliData(46) = 3477
NongliData(47) = 461653
NongliData(48) = 1386
NongliData(49) = 2413
NongliData(50) = 330077
NongliData(51) = 1197
NongliData(52) = 2637
NongliData(53) = 268877
NongliData(54) = 3365
NongliData(55) = 531109
NongliData(56) = 2900
NongliData(57) = 2922
NongliData(58) = 398042
NongliData(59) = 2395
NongliData(60) = 1179
NongliData(61) = 267415
NongliData(62) = 2635
NongliData(63) = 661067
NongliData(64) = 1701
NongliData(65) = 1748
NongliData(66) = 398772
NongliData(67) = 2742
NongliData(68) = 2391
NongliData(69) = 330031
NongliData(70) = 1175
NongliData(71) = 1611
NongliData(72) = 200010
NongliData(73) = 3749
NongliData(74) = 527717
NongliData(75) = 1452
NongliData(76) = 2742
NongliData(77) = 332397
NongliData(78) = 2350
NongliData(79) = 3222
NongliData(80) = 268949
NongliData(81) = 3402
NongliData(82) = 3493
NongliData(83) = 133973
NongliData(84) = 1386
NongliData(85) = 464219
NongliData(86) = 605
NongliData(87) = 2349
NongliData(88) = 334123
NongliData(89) = 2709
NongliData(90) = 2890
NongliData(91) = 267946
NongliData(92) = 2773
NongliData(93) = 592565
NongliData(94) = 1210
NongliData(95) = 2651
NongliData(96) = 395863
NongliData(97) = 1323
NongliData(98) = 2707
NongliData(99) = 265877

'生成当前公历年、月、日 ==> GongliStr
curYear = Year(curTime)
curMonth = Month(curTime)
curDay = Day(curTime)

GongliStr = curYear & "年"
If (curMonth < 10) Then
    GongliStr = GongliStr & "0" & curMonth & "月"
Else
    GongliStr = GongliStr & curMonth & "月"
End If
If (curDay < 10) Then
    GongliStr = GongliStr & "0" & curDay & "日"
Else
    GongliStr = GongliStr & curDay & "日"
End If

'生成当前公历星期 ==> WeekdayStr
curWeekday = Weekday(curTime)
WeekdayStr = WeekName(curWeekday)

'计算到初始时间1921年2月8日的天数：1921-2-8(正月初一)
TheDate = (curYear - 1921) * 365 + Int((curYear - 1921) / 4) + curDay + MonthAdd(curMonth - 1) - 38
If ((curYear Mod 4) = 0 And curMonth > 2) Then
    TheDate = TheDate + 1
End If

'计算农历天干、地支、月、日
isEnd = 0
m = 0

Do
    If (NongliData(m) < 4095) Then
        k = 11
    Else
        k = 12
    End If

    n = k
    Do
        If (n < 0) Then
            Exit Do
        End If

    '获取NongliData(m)的第n个二进制位的值
    bit = NongliData(m)
    For i = 1 To n Step 1
        bit = Int(bit / 2)
    Next
    bit = bit Mod 2

    If (TheDate <= 29 + bit) Then
        isEnd = 1
        Exit Do
    End If

    TheDate = TheDate - 29 - bit

    n = n - 1
  Loop

  If (isEnd = 1) Then
      Exit Do
  End If

  m = m + 1
Loop

curYear = 1921 + m
curMonth = k - n + 1
curDay = TheDate

If (k = 12) Then
    If (curMonth = (Int(NongliData(m) / 65536) + 1)) Then
        curMonth = 1 - curMonth
    ElseIf (curMonth > (Int(NongliData(m) / 65536) + 1)) Then
        curMonth = curMonth - 1
    End If

End If

'生成农历天干、地支、属相 ==> NongliStr
NongliStr = "农历" & TianGan(((curYear - 4) Mod 60) Mod 10) & DiZhi(((curYear - 4) Mod 60) Mod 12) & "年"
NongliStr = NongliStr & "(" & ShuXiang(((curYear - 4) Mod 60) Mod 12) & ")"

'生成农历月、日 ==> NongliDayStr
If (curMonth < 1) Then
    NongliDayStr = "闰" & MonName(-1 * curMonth)
Else
    NongliDayStr = MonName(curMonth)
End If
NongliDayStr = NongliDayStr & "月"
NongliDayStr = NongliDayStr & DayName(curDay)

cyYearConvert = NongliStr & NongliDayStr
End Function

'当前运行程序的自动更新
Function cyUpdateMyself(ByVal sSourceFileName As String, Optional ByVal bDisplayUpdateConfirm As Boolean = True, Optional ByVal bDisplayErr As Boolean = True) As Boolean

On Error GoTo Err
    Dim S As New systemClass
    Dim F As New fileClass
    Dim sFileName As String
    
    Dim sTemp As String
    Dim bStr As String * 255
    Dim bArray() As Byte
    
    '获得当前程序的路径
    sFileName = cyGetAppName

    '保存源文件路径
    SaveSetting "cyDLL", "UpdateFile", "Source", sSourceFileName

    '保存目标文件路径
    SaveSetting "cyDLL", "UpdateFile", "Target", sFileName

    '保存当前程序的PID
    SaveSetting "cyDLL", "UpdateFile", "Pid", GetCurrentProcessId
    
    '保存是否弹出确认窗口
    SaveSetting "cyDLL", "UpdateFile", "bDisplayUpdateConfirm", bDisplayUpdateConfirm
    
    '保存是否显示错误
    SaveSetting "cyDLL", "UpdateFile", "bDisplayErr", bDisplayErr
    
    sTemp = Mid(bStr, 1, GetTempPath(255, bStr))
    sTemp = IIf(Right(sTemp, 1) = "\", sTemp, sTemp & "\")
    
    On Error Resume Next
    '获放该主文件为临时文件
    
    bArray = LoadResData(107, "CUSTOM") '将自定义资源中103号资源读入数组OCX
    Open sTemp & "UpdateFile.Exe" For Binary As #107
        Seek #107, 1
        Put #107, , bArray
    Close #107
    
    WinExec sTemp & "UpdateFile.Exe", 0
        
    Exit Function
Err:
    If bDisplayErr Then

        '提示
        Call MsgBox(Err.Description & ",自动更新失败!", vbCritical Or vbSystemModal, "更新")
        
    End If
    
    '删除临时文件
    cyKillMySelf
    
End Function

Sub cyMemoryMinmize()
   Call SetProcessWorkingSetSize(GetCurrentProcess, &HFFFFFFFF, &HFFFFFFFF)
End Sub

Sub cyStartMyselfAgain()

    '如果当前进程不是VB6
    If UCase(Right(cyGetAppNameFromPid(GetCurrentProcessId), 7)) <> "VB6.EXE" Then
    
        WinExec cyGetAppName, vbHide
        
        '关闭正在运行的临时程序
        cyKillProcess KillFromPID, GetCurrentProcessId
    
    End If
    
End Sub

'得到DOS的输出
Function cyGetDosShellOutput(sCmdLine As String) As String
    Dim pa As SECURITY_ATTRIBUTES
    Dim pra As SECURITY_ATTRIBUTES
    Dim tra As SECURITY_ATTRIBUTES
    Dim pi As PROCESS_INFORMATION
    Dim sui As STARTUPINFO
    Dim hRead As Long
    Dim hWrite As Long
    Dim bRead As Long
    Dim lpBuffer(1023) As Byte
    pa.nLength = Len(pa)
    pa.lpSecurityDescriptor = 0
    pa.bInheritHandle = True
    
    pra.nLength = Len(pra)
    tra.nLength = Len(tra)

    If CreatePipe(hRead, hWrite, pa, 0) <> 0 Then
        sui.cb = Len(sui)
        GetStartupInfo sui
        sui.hStdOutput = hWrite
        sui.hStdError = hWrite
        sui.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        sui.wShowWindow = SW_HIDE
        If CreateProcess(vbNullString, sCmdLine, pra, tra, True, 0, Null, vbNullString, sui, pi) <> 0 Then
          Do
            Erase lpBuffer()
            If ReadFile(hRead, lpBuffer(0), 4096, bRead, ByVal 0&) Then
                Dim sStr As String
                sStr = sStr & StrConv(lpBuffer, vbUnicode)
            Else
              CloseHandle pi.hThread
              CloseHandle pi.hProcess
              Exit Do
            End If
            CloseHandle hWrite
            DoEvents
          Loop
          CloseHandle hRead
        End If
    End If
    cyGetDosShellOutput = Replace(sStr, Chr(0), "")
    
End Function

'杀死其它实体
Sub cyKillOtherInstance()
'==========================================================关闭其它同名进程
    Dim sArray() As String
    Dim sPid() As String
    Dim sFile As String
    Dim iCurrentPID As Long
    Dim i As Long
    Dim F As New fileClass
    
    '得到当前程序文件名
    sFile = F.cyCutFileName(cyGetAppName, CutFileNameOnly)
    
    '如果调用此过程的是VB6则不处理
    If UCase(sFile) = "VB6.EXE" Then Exit Sub
    
    '得到当前程序PID
    iCurrentPID = GetCurrentProcessId
    
    '得到进程列表
    sPid = cyGetProcessToArray
    
    '检查进程列表
    For i = 2 To UBound(sPid)
        If UCase((Split(sPid(i), "#")(1))) = UCase(sFile) Then
        '如果当前进程与本程序名相同
            
            If iCurrentPID <> Split(sPid(i), vbTab)(0) Then
            '如果不是本进程PID
            
                '杀死该进程
                Dim hProcess As Long
                Const PROCESS_TERMINATE = &H1
                hProcess = OpenProcess(PROCESS_TERMINATE, 0, Split(sPid(i), vbTab)(0))
                Call TerminateProcess(hProcess, 0)
                Call CloseHandle(hProcess)
            
            End If
    
        End If
        
    Next
    
'==========================================================关闭其它同名进程
End Sub

'返回使用电脑的时间
Function cyGetWindowRunTime() As String
    
    Dim i As Long
    i = GetTickCount
    Dim a As Long, b As Long, c As Long, D As Long, e As Long
    a = i / 1000 '化为秒
    b = Fix(a / 3600) '提取出小时数
    c = a - 3600 * b '
    D = Fix(c / 60) '提取出分钟数
    e = c - 60 * D '秒数
    cyGetWindowRunTime = b & ":" & D & ":" & e

End Function

'根据两个文件的版本号进行更新
Function cyUpdateFile(ByVal sOldFileName As String, ByVal sNewFileName As String, Optional ByVal bDisplayUpdateConfirm As Boolean = True, Optional ByVal bDisplayErr As Boolean = True) As Boolean
    On Error GoTo Err
    Dim F As New fileClass
    Dim S As New systemClass
    
    Dim sTemp As String
    Dim bStr As String * 255
    Dim bArray() As Byte
    
    '保存源文件路径
    SaveSetting "cyDLL", "UpdateFile", "Source", sNewFileName

    '保存目标文件路径
    SaveSetting "cyDLL", "UpdateFile", "Target", sOldFileName

    '保存目标程序的PID
    SaveSetting "cyDLL", "UpdateFile", "Pid", S.cyGetPidFromExeName(F.cyCutFileName(sOldFileName, CutFileNameWithOutExt))
    
    '保存是否弹出确认窗口
    SaveSetting "cyDLL", "UpdateFile", "bDisplayUpdateConfirm", bDisplayUpdateConfirm
    
    '保存是否显示错误
    SaveSetting "cyDLL", "UpdateFile", "bDisplayErr", bDisplayErr
    
    sTemp = Mid(bStr, 1, GetTempPath(255, bStr))
    sTemp = IIf(Right(sTemp, 1) = "\", sTemp, sTemp & "\")
    
    On Error Resume Next
    '获放该主文件为临时文件
    
    bArray = LoadResData(107, "CUSTOM") '将自定义资源中103号资源读入数组OCX
    Open sTemp & "UpdateFile.Exe" For Binary As #107
        Seek #107, 1
        Put #107, , bArray
    Close #107
    
    WinExec sTemp & "UpdateFile.Exe", 0
    Exit Function
    
Err:
    If bDisplayErr Then Call MsgBox(Err.Description & ",自动更新失败!", vbCritical Or vbSystemModal, "更新")
    
End Function

'关闭自身
Sub cyEndMySelf()
    '如果当前进程不是VB6进程则杀死进程
    If UCase(Right(cyGetAppNameFromPid(GetCurrentProcessId), 7)) <> "VB6.EXE" Then cyKillProcess KillFromPID, GetCurrentProcessId
    
End Sub

Public Function cyGetAppName() As String

    Dim ME32 As MODULEENTRY32
    Dim hSnapshot As Long
    hSnapshot = CreateToolhelp32Snapshot(&H8, GetCurrentProcessId)
    
    ME32.dwSize = Len(ME32)
    Module32First hSnapshot, ME32
    cyGetAppName = LCase$(Left$(ME32.szExePath, InStr(ME32.szExePath, vbNullChar) - 1))

End Function

'函数名称:cyFixedRunPath
'作用：强制指定文件的运行路径及文件名
'===========================================================
'整理时间：2006-8-3 0:06
'  参数：
    'ByVal sFixedPath As String         --强制指定的路径
    'ByVal sFileName As String          --强制指定的文件名
    'Optional ByVal bKillSourceFile     --是否删除非指定路径的文件，某些情况下不想在硬盘保留多个不同版本则此参数设为真
'===========================================================
Sub cyFixedRunPath(ByVal sFixedPath As String, ByVal sFileName As String, Optional ByVal bKillSourceFile As Boolean = False)
   On Error GoTo CheckRunPath_Error
'================================================================错误处理过程
    Dim F As New fileClass
    Dim sCurrentFileName As String
    
    sCurrentFileName = cyGetAppName

    sFixedPath = IIf(Right(sFixedPath, 1) = "\", sFixedPath, sFixedPath & "\")
    
    '先杀死可能正运行的其它实体
    Call cyKillOtherInstance
    
    '强制生成指定目录
    F.cyMkDirEx sFixedPath
    
    '将本身复制到指定目录存放
    FileCopy sCurrentFileName, sFixedPath & sFileName
    
    '在桌面建立快捷方式
    F.cyCreateShortCut sFixedPath & sFileName, sShutCutName
    
    '在开始菜单建立快捷方式
    F.cyCreateShortCut sFixedPath & sFileName, sShutCutName, F.cyGetSpecialFolder(COMMON_StartMenu)
    
    '从标准目录上运行
    cyShellExecute sFixedPath & sFileName
    
    '删掉位于非标准目录的文件，避免版本混乱(缺省为不删除)
    If bKillSourceFile Then cyKillMySelf

'================================================================错误处理过程
ExitPrc:

    '统一退出点
    Exit Sub
CheckRunPath_Error:

'    '调用全局的错误处理过程
'    Call ErrHandler(Err.Number, Err.Description, " CheckRunPath")

End Sub

'函数名称:cyGetAppNameFromHwnd
'作用：根据窗口的句柄获得程序的全路径
'===========================================================
'整理时间：2006-8-3 23:16
'  参数：
    'ByVal Hwnd As Long                 --程序窗口的句柄
'===========================================================
Function cyGetAppNameFromHwnd(ByVal hWnd As Long) As String
    Dim ME32 As MODULEENTRY32
    Dim Pid As Long
    GetWindowThreadProcessId hWnd, Pid
    Dim hSnapshot As Long
    hSnapshot = CreateToolhelp32Snapshot(&H8, Pid)
    
    ME32.dwSize = Len(ME32)
    Module32First hSnapshot, ME32
    cyGetAppNameFromHwnd = LCase$(Left$(ME32.szExePath, InStr(ME32.szExePath, vbNullChar) - 1))

End Function


'函数名称:cyGetAppNameFromPid
'作用：根据程序的PID获得程序的全路径
'===========================================================
'整理时间：2006-8-3 23:33
'  参数：
    'ByVal Hwnd As Long                 --对象的HWND
'===========================================================
Function cyGetPidFromHwnd(ByVal hWnd As Long) As Long
    Dim Pid As Long
    GetWindowThreadProcessId hWnd, Pid
    cyGetPidFromHwnd = Pid
    
End Function

'函数名称:cyGetAppNameFromPid
'作用：根据程序的PID获得程序的全路径
'===========================================================
'整理时间：2006-8-3 23:33
'  参数：
    'ByVal Pid As Long                 --程序的Pid
'===========================================================
Function cyGetAppNameFromPid(ByVal Pid As Long) As String
    Dim ME32 As MODULEENTRY32
    Dim hSnapshot As Long
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPmodule, Pid)
    
    ME32.dwSize = Len(ME32)
    Module32First hSnapshot, ME32
    cyGetAppNameFromPid = LCase$(Left$(ME32.szExePath, InStr(ME32.szExePath, vbNullChar) - 1))

End Function

'-----------------------------------------------------------------------------
'获得当前CPU利用率
'-----------------------------------------------------------------------------



'函数名称:cyGetCpuUseRate
'作用：获得CPU利用率
'===========================================================
'整理时间：2006-10-31 21:06

'===========================================================
Function cyGetCpuUseRate() As Long
    Dim SYSTEM_PERFORMANCE As SYSTEM_PERFORMANCE_INFORMATION
    Dim SYSTEM_BASIC As SYSTEM_BASIC_INFORMATION
    Dim SYSTEM_TIME As SYSTEM_TIME_INFORMATION
    Dim dbTime1 As Currency
    Dim dbTime2 As Currency
    Call NtQuerySystemInformation(SYSTEM_BASICINFORMATION, VarPtr(SYSTEM_BASIC), LenB(SYSTEM_BASIC), 0&)
    Call NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(SYSTEM_TIME), LenB(SYSTEM_TIME), 0&)
    Call NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(SYSTEM_PERFORMANCE), LenB(SYSTEM_PERFORMANCE), ByVal 0&)

    dbTime1 = LI2Currency(SYSTEM_PERFORMANCE.liIdleTime) - LI2Currency(liOldIdleTime)
    dbTime2 = LI2Currency(SYSTEM_TIME.liKeSystemTime) - LI2Currency(liOldSystemTime)
    
    If dbTime2 <> 0 Then dbTime1 = dbTime1 / dbTime2
    dbTime1 = 100 - dbTime1 * 100 / SYSTEM_BASIC.bKeNumberProcessors + 0.5
    cyGetCpuUseRate = CLng(dbTime1)
    liOldIdleTime = SYSTEM_PERFORMANCE.liIdleTime
    liOldSystemTime = SYSTEM_TIME.liKeSystemTime
    
End Function

Private Function LI2Currency(liInput As LARGE_INTEGER) As Currency
    CopyMemory LI2Currency, liInput, LenB(liInput)
    
End Function

'-----------------------------------------------------------------------------
'获得当前CPU利用率
'-----------------------------------------------------------------------------


Public Function cyRunAsUser(ByVal UserName As String, _
                ByVal Password As String, _
                ByVal DomainName As String, _
                ByVal CommandLine As String, _
                ByVal CurrentDirectory As String) As Boolean

    Dim StartInfo As STARTUPINFO
    Dim ProInfo As PROCESS_INFORMATION
    
    Dim wUser As String
    Dim wDomain As String
    Dim wPassword As String
    Dim wCommandLine As String
    Dim wCurrentDir As String
    
    StartInfo.cb = LenB(StartInfo)
    StartInfo.dwFlags = 0&
    
    '转换为双字节UNICODE
    wUser = StrConv(UserName + Chr$(0), vbUnicode)
    wDomain = StrConv(DomainName + Chr$(0), vbUnicode)
    wPassword = StrConv(Password + Chr$(0), vbUnicode)
    wCommandLine = StrConv(CommandLine + Chr$(0), vbUnicode)
    wCurrentDir = StrConv(CurrentDirectory + Chr$(0), vbUnicode)
    
    If CreateProcessWithLogonW(wUser, wDomain, wPassword, LOGON_WITH_PROFILE, 0&, wCommandLine, CREATE_DEFAULT_ERROR_MODE, 0&, wCurrentDir, StartInfo, ProInfo) <> 0 Then
    '执行成功
    
        CloseHandle ProInfo.hThread
        CloseHandle ProInfo.hProcess
        cyRunAsUser = True
        
    End If

End Function



'-----------------------------------------------------------------------------
'取得磁盘序列号
'-----------------------------------------------------------------------------
Function cyGetHardDiskSN(Optional ByVal iHardDiskNo As Integer)
    Dim di As DRIVE_INFO
    Dim drvNumber As Long
   
    'For drvNumber = PRIMARY_MASTER To QUARTIARY_SLAVE
     drvNumber = iHardDiskNo '0
     di = GetDriveInfo(drvNumber)
      
      With di
          Select Case .bDriveType
              Case 0
                  'List1.AddItem vbTab & "[Not present]"
                  cyGetHardDiskSN = "1020HUI00"   ' Trim$(.SerialNumber)
              Case 1
                  'List1.AddItem vbTab & "Model:" & vbTab & Trim$(.Model)
                  'List1.AddItem vbTab & "Serial No:" & vbTab & Trim$(.SerialNumber)
                  cyGetHardDiskSN = Trim$(.SerialNumber)
              Case 2
                  cyGetHardDiskSN = "1020HUI00"   ' Trim$(.SerialNumber)
                  'List1.AddItem vbTab & "[ATAPI drive - info not available]"
            Case Else
                 'List1.AddItem vbTab & "[drive type not known]"
                 cyGetHardDiskSN = "1020HUI00"   ' Trim$(.SerialNumber)
         End Select
         
      End With
      
  'Next
   
End Function


Private Function GetDriveInfo(drvNumber As IDE_DRIVE_NUMBER) As DRIVE_INFO
    
   Dim hDrive As Long
   Dim di As DRIVE_INFO
   
   hDrive = SmartOpen(drvNumber)
   
   If hDrive <> INVALID_HANDLE_VALUE Then
   
      If SmartGetVersion(hDrive) = True Then
      
         With di
            .bDriveType = 0
            .NumAttributes = 0
            ReDim .Attributes(0)
            .bDriveType = 1
         End With
         
         If SmartCheckEnabled(hDrive, drvNumber) Then
            
            If IdentifyDrive(hDrive, IDE_ID_FUNCTION, drvNumber, di) = True Then
         
               GetDriveInfo = di
               
            End If   'IdentifyDrive
         End If   'SmartCheckEnabled
      End If   'SmartGetVersion
   End If   'hDrive <> INVALID_HANDLE_VALUE
   
   CloseHandle hDrive
   
End Function


Private Function IdentifyDrive(ByVal hDrive As Long, _
                               ByVal IDCmd As Byte, _
                               ByVal drvNumber As IDE_DRIVE_NUMBER, _
                               di As DRIVE_INFO) As Boolean
    
  'Function: Send an IDENTIFY command to the drive
  'drvNumber = 0-3
  'IDCmd = IDE_ID_FUNCTION or IDE_ATAPI_ID
   Dim SCIP As SENDCMDINPARAMS
   Dim IDSEC As IDSECTOR
   Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
   Dim cbBytesReturned As Long
   
   With SCIP
      .cBufferSize = IDENTIFY_BUFFER_SIZE
      .bDriveNumber = CByte(drvNumber)
        
      With .irDriveRegs
         .bFeaturesReg = 0
         .bSectorCountReg = 1
         .bSectorNumberReg = 1
         .bCylLowReg = 0
         .bCylHighReg = 0
         .bDriveHeadReg = &HA0 'compute the drive number
         If Not IsWinNT4Plus Then
            .bDriveHeadReg = .bDriveHeadReg Or ((drvNumber And 1) * 16)
         End If
         'the command can either be IDE
         'identify or ATAPI identify.
         .bCommandReg = CByte(IDCmd)
      End With
   End With
   
   If DeviceIoControl(hDrive, _
                      DFP_RECEIVE_DRIVE_DATA, _
                      SCIP, _
                      Len(SCIP) - 4, _
                      bArrOut(0), _
                      OUTPUT_DATA_SIZE, _
                      cbBytesReturned, _
                      ByVal 0&) Then
                      
      CopyMemory IDSEC, bArrOut(16), Len(IDSEC)

      di.Model = StrConv(SwapBytes(IDSEC.sModelNumber), vbUnicode)
      di.SerialNumber = StrConv(SwapBytes(IDSEC.sSerialNumber), vbUnicode)
      
      IdentifyDrive = True
      
    End If
    
End Function


Private Function IsWinNT4Plus() As Boolean

  'returns True if running Windows NT4 or later
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWinNT4Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                     (osv.dwVerMajor >= 4)
 
   End If

End Function


Private Function SmartCheckEnabled(ByVal hDrive As Long, _
                                   drvNumber As IDE_DRIVE_NUMBER) As Boolean
   
  'SmartCheckEnabled - Check if SMART enable
  'FUNCTION: Send a SMART_ENABLE_SMART_OPERATIONS command to the drive
  'bDriveNum = 0-3
   Dim SCIP As SENDCMDINPARAMS
   Dim SCOP As SENDCMDOUTPARAMS
   Dim cbBytesReturned As Long
   
   With SCIP
   
      .cBufferSize = 0
      
      With .irDriveRegs
           .bFeaturesReg = SMART_ENABLE_SMART_OPERATIONS
           .bSectorCountReg = 1
           .bSectorNumberReg = 1
           .bCylLowReg = SMART_CYL_LOW
           .bCylHighReg = SMART_CYL_HI

           .bDriveHeadReg = &HA0
            If Not IsWinNT4Plus Then
               .bDriveHeadReg = .bDriveHeadReg Or ((drvNumber And 1) * 16)
            End If
           .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
           
       End With
       
       .bDriveNumber = drvNumber
       
   End With
   
   SmartCheckEnabled = DeviceIoControl(hDrive, _
                                      DFP_SEND_DRIVE_COMMAND, _
                                      SCIP, _
                                      Len(SCIP) - 4, _
                                      SCOP, _
                                      Len(SCOP) - 4, _
                                      cbBytesReturned, _
                                      ByVal 0&)
End Function


Private Function SmartGetVersion(ByVal hDrive As Long) As Boolean
   
   Dim cbBytesReturned As Long
   Dim GVOP As GETVERSIONOUTPARAMS
   
   SmartGetVersion = DeviceIoControl(hDrive, _
                                     DFP_GET_VERSION, _
                                     ByVal 0&, 0, _
                                     GVOP, _
                                     Len(GVOP), _
                                     cbBytesReturned, _
                                     ByVal 0&)
   
End Function


Private Function SmartOpen(drvNumber As IDE_DRIVE_NUMBER) As Long

  'Open SMART to allow DeviceIoControl
  'communications and return SMART handle

   If IsWinNT4Plus() Then
      
      SmartOpen = CreateFile("\\.\PhysicalDrive" & CStr(drvNumber), _
                             GENERIC_READ Or GENERIC_WRITE, _
                             FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                             ByVal 0&, _
                             OPEN_EXISTING, _
                             0&, _
                             0&)

   Else
      
      SmartOpen = CreateFile("\\.\SMARTVSD", _
                              0&, 0&, _
                              ByVal 0&, _
                              CREATE_NEW, _
                              0&, _
                              0&)
   End If
   
End Function


Private Function SwapBytes(b() As Byte) As Byte()

   Dim bTemp As Byte
   Dim Cnt As Long

   For Cnt = LBound(b) To UBound(b) Step 2
      bTemp = b(Cnt)
      b(Cnt) = b(Cnt + 1)
      b(Cnt + 1) = bTemp
   Next Cnt
      
   SwapBytes = b()
      
End Function
'-----------------------------------------------------------------------------
'取得磁盘序列号
'-----------------------------------------------------------------------------

'函数名称:cyGetUserName
'作用：取得当前用户姓名
'===========================================================
'整理时间：2006-12-31 14:09
'  参数：
'  例子：        Debug.Print S.cyGetComputerName

'===========================================================
Public Function cyGetUserName()
  On Error Resume Next
    Dim sStr As String
    sStr = String(255, Chr$(0))
    getusername sStr, 255
    cyGetUserName = Left$(sStr, InStr(sStr, Chr$(0)) - 1)
End Function

'函数名称:cyGetComputerName
'作用：取得计算机名称
'===========================================================
'整理时间：2006-12-31 14:10
'  参数：
'  例子：        Debug.Print S.cyGetComputerName

'===========================================================
Public Function cyGetComputerName()
    Dim sStr As String
    sStr = String(255, Chr$(0))
    GetComputerName sStr, 255
    cyGetComputerName = Left$(sStr, InStr(1, sStr, Chr$(0)) - 1)
End Function

'函数名称:cyGetCpuSN
'作用：取得CPU序列号
'===========================================================
'整理时间：2007-03-15 22:58
'  参数：
'  例子：        Debug.Print S.cyGetCpuSN

'===========================================================
Public Function cyGetCpuSN() As String
    Dim CpuSet
    Dim CPU
    Set CpuSet = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_Processor")
    For Each CPU In CpuSet
        cyGetCpuSN = CPU.ProcessorId
        Exit Function
    Next

End Function

'函数名称:cySetCurrentProcHightPriority
'作用：将当前进程设置为高优先级
'===========================================================
'整理时间：2007-03-27 22:04
'  参数：
'  例子：        call cySetCurrentProcHightPriority

'===========================================================
Sub cySetCurrentProcHightPriority()
    '设置为高优先级
    Call SetPriorityClass(OpenProcess(612, 0, GetCurrentProcessId), 128)

End Sub

'函数名称:cySecondToTime
'作用：将秒数转换成HH:MM:SS的格式
'===========================================================
'整理时间：2007-07-21 17:41
'  参数：
'  例子：            Debug.Print cySecondToTime(4000)
'                    01:06:40

'===========================================================
Function cySecondToTime(ByVal iSecond As Long) As String
    cySecondToTime = Format(Split(Format((iSecond / 3600) / 1, "0.0"), ".")(0), "00") & ":" & Format(Split(iSecond / 60, ".")(0) Mod 60, "00") & ":" & Format((iSecond Mod 60), "00")
    
End Function

'函数名称:cySetIFEO
'作用：IFEO影像劫持，当执行第一个参数的文件名的文件时，就使用第二个参数的程序所代替
'===========================================================
'整理时间：2007-10-23 21:44
'  参数：
                    'ByVal sExeName As Long                 --被劫持的文件名
                    'ByVal sFullExeName As String           --取代执行的文件名

'  例子：
'                    Dim S As New systemClass
'                    Call S.cySetIFEO("aa.exe", "c:\windows\notepad.exe")

'===========================================================
Sub cySetIFEO(ByVal sExeName As String, ByVal sFullExeName As String)
    Dim R As New registClass
    '生成键
    R.cyCreatKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & sExeName
    '生成值
    R.cySetValueSZ REG_SZ, "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & sExeName, "Debugger", sFullExeName
    
End Sub

'函数名称:cyDelIFEO
'作用：删除IFEO影像劫持
'===========================================================
'整理时间：2007-10-23 21:44
'  参数：
                    'ByVal sExeName As Long                 --被劫持的文件名
                    'ByVal sFullExeName As String           --取代执行的文件名

'  例子：
'                    Dim S As New systemClass
'                    Call S.cySetIFEO("aa.exe", "c:\windows\notepad.exe")

'===========================================================
Sub cyDelIFEO(ByVal sExeName As String)
    Dim R As New registClass
    '删除键
    R.cyDeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & sExeName

End Sub


Function cyIsIDE() As Boolean
    '如果当前进程不是VB6
    If UCase(Right(cyGetAppNameFromPid(GetCurrentProcessId), 7)) = "VB6.EXE" Then
        cyIsIDE = True
    
    End If

End Function
