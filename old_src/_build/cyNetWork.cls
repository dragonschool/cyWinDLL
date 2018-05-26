VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cyNetWorkEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------
'检查域用户
'-----------------------------------------------------------------------------
Private Declare Function LogonUser Lib "advapi32.dll" Alias _
                "LogonUserA" _
                (ByVal lpszUsername As String, _
                ByVal lpszDomain As String, _
                ByVal lpszPassword As String, _
                ByVal dwLogonType As Long, _
                ByVal dwLogonProvider As Long, _
                phToken As Long) As Long
  
Private Declare Function CloseHandle Lib "kernel32.dll" _
                (ByVal hObject As Long) As Long

'-----------------------------------------------------------------------------
'检查域用户
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'发邮件
'-----------------------------------------------------------------------------
Dim sMailSmtpServer As String
Dim sMailID As String
Dim sMailPWD As String
'-----------------------------------------------------------------------------
'发邮件
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'取得MAC
'-----------------------------------------------------------------------------
Private Const NCBASTAT = &H33
Private Const NCBNAMSZ = 16
Private Const HEAP_ZERO_MEMORY = &H8
Private Const HEAP_GENERATE_EXCEPTIONS = &H4
Private Const NCBRESET = &H32

Private Type NCB
    ncb_command As Byte 'Integer
    ncb_retcode As Byte 'Integer
    ncb_lsn As Byte 'Integer
    ncb_num As Byte ' Integer
    ncb_buffer As Long 'String
    ncb_length As Integer
    ncb_callname As String * NCBNAMSZ
    ncb_name As String * NCBNAMSZ
    ncb_rto As Byte 'Integer
    ncb_sto As Byte ' Integer
    ncb_post As Long
    ncb_lana_num As Byte 'Integer
    ncb_cmd_cplt As Byte 'Integer
    ncb_reserve(9) As Byte ' Reserved, must be 0
    ncb_event As Long
End Type

Private Type ADAPTER_STATUS
    adapter_address(5) As Byte 'As String * 6
    rev_major As Byte 'Integer
    reserved0 As Byte 'Integer
    adapter_type As Byte 'Integer
    rev_minor As Byte 'Integer
    duration As Integer
    frmr_recv As Integer
    frmr_xmit As Integer
    iframe_recv_err As Integer
    xmit_aborts As Integer
    xmit_success As Long
    recv_success As Long
    iframe_xmit_err As Integer
    recv_buff_unavail As Integer
    t1_timeouts As Integer
    ti_timeouts As Integer
    Reserved1 As Long
    free_ncbs As Integer
    max_cfg_ncbs As Integer
    max_ncbs As Integer
    xmit_buf_unavail As Integer
    max_dgram_size As Integer
    pending_sess As Integer
    max_cfg_sess As Integer
    max_sess As Integer
    max_sess_pkt_size As Integer
    name_count As Integer
End Type

Private Type NAME_BUFFER
    name As String * NCBNAMSZ
    name_num As Integer
    name_flags As Integer
End Type

Private Type ASTAT
    adapt As ADAPTER_STATUS
    NameBuff(30) As NAME_BUFFER
End Type

Private Declare Function Netbios Lib "Netapi32.dll" (pncb As NCB) As Byte
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

'-----------------------------------------------------------------------------
'取得MAC
'-----------------------------------------------------------------------------



'-----------------------------------------------------------------------------
'FTP
'-----------------------------------------------------------------------------
Dim bPASSIVE As Boolean         '设置是否使用PASSIVE模式
Private hInternet As Long       ' The handle to the Internet
Private hConnect As Long        ' The handle to the Internet Connection
Private sFtpSite1 As String         ' The site to which we connect to
Private sUserName1 As String         ' The user name used to connect to the site
Private sPassWord1 As String     ' The password for the user
Private iPort1 As Long


Private Const MAX_PATH  As Long = 260
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const INTERNET_OPEN_TYPE_DIRECT = 1                         ' direct to net
Private Const INTERNET_FLAG_PASSIVE = &H8000000             ' used for FTP connections
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000      ' don't write this item to the cache
Private Const INTERNET_DEFAULT_FTP_PORT = 21                ' default for FTP servers
Private Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000   ' FTP: use existing InternetConnect handle for server if possible
Private Const INTERNET_SERVICE_FTP = 1
Private Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Private Const FTP_TRANSFER_TYPE_BINARY = &H2

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Declare Function InternetOpen Lib "wininet.dll" _
   Alias "InternetOpenA" _
  (ByVal lpszAgent As String, _
   ByVal dwAccessType As Long, _
   ByVal lpszProxyName As String, _
   ByVal lpszProxyBypass As String, _
   ByVal dwFlags As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
   (ByVal hEnumHandle As Long) As Long

Private Declare Function InternetConnect Lib "wininet.dll" _
   Alias "InternetConnectA" _
  (ByVal hInternet As Long, _
   ByVal lpszServerName As String, _
   ByVal nServerPort As Long, _
   ByVal lpszUsername As String, _
   ByVal lpszPassword As String, _
   ByVal dwService As Long, _
   ByVal dwFlags As Long, _
   ByVal dwContext As Long) As Long

Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long

Private Declare Function FtpFindFirstFile Lib "wininet.dll" _
   Alias "FtpFindFirstFileA" _
  (ByVal hConnect As Long, _
   ByVal lpszSearchFile As String, _
   lpFindFileData As Any, _
   ByVal dwFlags As Long, _
   ByVal dwContext As Long) As Long

Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" _
   Alias "FtpGetCurrentDirectoryA" _
  (ByVal hConnect As Long, _
   ByVal lpszCurrentDirectory As String, _
    lpdwCurrentDirectory As Long) As Long

Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" _
   Alias "FtpSetCurrentDirectoryA" _
  (ByVal hConnect As Long, _
   ByVal lpszDirectory As String) As Long
   
Private Declare Function FtpGetFile Lib "wininet.dll" _
   Alias "FtpGetFileA" _
  (ByVal hConnect As Long, _
   ByVal lpszRemoteFile As String, _
   ByVal lpszNewFile As String, _
   ByVal fFailIfExists As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal dwFlags As Long, _
   ByVal dwContext As Long) As Long

Private Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" _
    (ByVal hConnect As Long, _
     ByVal lpszFileName As String) As Long
     
Private Declare Function FtpPutFile Lib "wininet.dll" _
    Alias "FtpPutFileA" _
    (ByVal hFtpSession As Long, _
     ByVal lpszLocalFile As String, _
     ByVal lpszRemoteFile As String, _
     ByVal dwFlags As Long, _
     ByVal dwContext As Long) As Boolean

Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean

'-----------------------------------------------------------------------------
'FTP
'-----------------------------------------------------------------------------




Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long


'-----------------------------------------------------------------------------
'断开所有网络设备
'-----------------------------------------------------------------------------

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Private Declare Function SetupDiGetClassDevs Lib "setupapi" Alias "SetupDiGetClassDevsA" (ByVal ClassGuid As Long, ByVal Enumerator As Long, ByVal hWndParent As Long, ByVal flags As Long) As Long

Private Const DIGCF_PRESENT            As Long = &H2
Private Const DIGCF_ALLCLASSES         As Long = &H4

Private Declare Function SetupDiDestroyDeviceInfoList Lib "setupapi" (ByVal DeviceInfoSet As Long) As Long

Private Declare Function SetupDiEnumDeviceInfo Lib "setupapi" (ByVal DeviceInfoSet As Long, ByVal MemberIndex As Long, DeviceInfoData As SP_DEVINFO_DATA) As Long

Private Type SP_DEVINFO_DATA
    cbSize As Long
    ClassGuid As GUID
    DevInst As Long
    Reserved As Long
End Type

Private Declare Function SetupDiGetDeviceRegistryProperty Lib "setupapi" Alias "SetupDiGetDeviceRegistryPropertyA" (ByVal DeviceInfoSet As Long, DeviceInfoData As SP_DEVINFO_DATA, ByVal Property As Long, ByRef PropertyRegDataType As Long, ByVal PropertyBuffer As Long, ByVal PropertyBufferSize As Long, RequiredSize As Long) As Long

Private Const SPDRP_CLASS                        As Long = (&H7)        '// Class (R--tied to ClassGUID)
Private Const SPDRP_CLASSGUID                    As Long = (&H8)        '// ClassGUID (R/W)

Private Declare Function SetupDiSetClassInstallParams Lib "setupapi" Alias "SetupDiSetClassInstallParamsA" (ByVal DeviceInfoSet As Long, DeviceInfoData As SP_DEVINFO_DATA, ClassInstallParams As SP_CLASSINSTALL_HEADER, ByVal ClassInstallParamsSize As Long) As Long

Private Declare Function SetupDiCallClassInstaller Lib "setupapi" (ByVal InstallFunction As Long, ByVal DeviceInfoSet As Long, DeviceInfoData As SP_DEVINFO_DATA) As Long

Private Type SP_CLASSINSTALL_HEADER
    cbSize As Long
    InstallFunction As Long
End Type

Private Const DIF_PROPERTYCHANGE                    As Long = &H12

Private Type SP_PROPCHANGE_PARAMS
    ClassInstallHeader As SP_CLASSINSTALL_HEADER
    StateChange As Long
    Scope As Long
    HwProfile As Long
End Type

Private Const DICS_ENABLE        As Long = &H1
Private Const DICS_DISABLE       As Long = &H2

Private Const DICS_FLAG_GLOBAL           As Long = &H1         '// make change in all hardware profiles
Private Const DICS_FLAG_CONFIGSPECIFIC   As Long = &H2         '// make change in specified profile only

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122
'-----------------------------------------------------------------------------
'断开所有网络设备
'-----------------------------------------------------------------------------



'-----------------------------------------------------------------------------
'检测是否连接INTERNET
'-----------------------------------------------------------------------------'***********************************************************************************************
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function InternetGetConnectedState _
    Lib "wininet.dll" (ByRef lpdwFlags As Long, _
    ByVal dwReserved As Long) As Long
    'Local system uses a modem to connect to
    '     the Internet.
Private Const INTERNET_CONNECTION_MODEM As Long = &H1
'Local system uses a LAN to connect to t
'     he Internet.
Private Const INTERNET_CONNECTION_LAN As Long = &H2
'Local system uses a proxy server to con
'     nect to the Internet.
Private Const INTERNET_CONNECTION_PROXY As Long = &H4
'No longer used.
Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
Private Const INTERNET_RAS_INSTALLED As Long = &H10
Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
    'InternetGetConnectedState wrapper funct
    '     ions
'-----------------------------------------------------------------------------
'检测是否连接INTERNET
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'查出IP地址
'-----------------------------------------------------------------------------
Private Const MAX_ADAPTER_NAME_LENGTH           As Long = 256
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH    As Long = 128
Private Const MAX_ADAPTER_ADDRESS_LENGTH        As Long = 8
Private Const ERROR_SUCCESS                     As Long = 0

Private Type IP_ADDRESS_STRING
     IpAddr(0 To 15)                    As Byte
End Type

Private Type IP_MASK_STRING
     IpMask(0 To 15)                    As Byte
End Type

Private Type IP_ADDR_STRING
     dwNext                             As Long
     IpAddress                          As IP_ADDRESS_STRING
     IpMask                             As IP_MASK_STRING
     dwContext                          As Long
End Type

Private Type IP_ADAPTER_INFO
  dwNext                                As Long
  ComboIndex                            As Long
  sAdapterName(0 To (MAX_ADAPTER_NAME_LENGTH + 3))          As Byte
  sDescription(0 To (MAX_ADAPTER_DESCRIPTION_LENGTH + 3))   As Byte
  dwAddressLength                                           As Long
  sIPAddress(0 To (MAX_ADAPTER_ADDRESS_LENGTH - 1))         As Byte
  dwIndex                                                   As Long
  uType                                                     As Long
  uDhcpEnabled                                              As Long
  CurrentIpAddress                                          As Long
  IpAddressList                                             As IP_ADDR_STRING
  GatewayList                                               As IP_ADDR_STRING
  DhcpServer                                                As IP_ADDR_STRING
  bHaveWins                                                 As Long
  PrimaryWinsServer                                         As IP_ADDR_STRING
  SecondaryWinsServer                                       As IP_ADDR_STRING
  LeaseObtained                                             As Long
  LeaseExpires                                              As Long
End Type

Private Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
'-----------------------------------------------------------------------------
'查出IP地址
'-----------------------------------------------------------------------------



'-----------------------------------------------------------------------------
'检查是否已连上INTERNET
'-----------------------------------------------------------------------------
Private Const FLAG_ICC_FORCE_CONNECTION = &H1
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
'-----------------------------------------------------------------------------
'检查是否已连上INTERNET
'-----------------------------------------------------------------------------


'Add this code to a module and set the Project's Startup Object to 'Sub Main'
'    (-> Project Menu -> Project Properties -> General Tab)
Private Const RESOURCE_CONNECTED As Long = &H1&
Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCE_REMEMBERED As Long = &H3&
Private Const RESOURCEDISPLAYTYPE_DIRECTORY& = &H9
Private Const RESOURCEDISPLAYTYPE_DOMAIN& = &H1
Private Const RESOURCEDISPLAYTYPE_FILE& = &H4
Private Const RESOURCEDISPLAYTYPE_GENERIC& = &H0
Private Const RESOURCEDISPLAYTYPE_GROUP& = &H5
Private Const RESOURCEDISPLAYTYPE_NETWORK& = &H6
Private Const RESOURCEDISPLAYTYPE_ROOT& = &H7
Private Const RESOURCEDISPLAYTYPE_SERVER& = &H2
Private Const RESOURCEDISPLAYTYPE_SHARE& = &H3
Private Const RESOURCEDISPLAYTYPE_SHAREADMIN& = &H8
Private Const RESOURCETYPE_ANY As Long = &H0&
Private Const RESOURCETYPE_DISK As Long = &H1&
Private Const RESOURCETYPE_PRINT As Long = &H2&
Private Const RESOURCETYPE_UNKNOWN As Long = &HFFFF&
Private Const RESOURCEUSAGE_ALL As Long = &H0&
Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&
Private Const RESOURCEUSAGE_CONTAINER As Long = &H2&
Private Const RESOURCEUSAGE_RESERVED As Long = &H80000000
Private Const NO_ERROR = 0
Private Const ERROR_MORE_DATA = 234                        'L    // dderror
Private Const RESOURCE_ENUM_ALL As Long = &HFFFF

Private Type NETRESOURCE
dwScope As Long
dwType As Long
dwDisplayType As Long
dwUsage As Long
lpLocalName As String
lpRemoteName As String
lpComment As String
lpProvider As String
End Type

Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Private Declare Function getusername Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hWnd As Long, ByVal dwType As Long) As Long


'-----------------------------------------------------------------------------
'设为脱机模式
'-----------------------------------------------------------------------------
Private Type INTERNET_CONNECTED_INFO

    dwConnectedState As Long
    dwFlags As Long
End Type

Private Const INTERNET_OPTION_CONNECTED_STATE = 50
Private Const INTERNET_STATE_CONNECTED = 1
Private Const INTERNET_STATE_DISCONNECTED = 2
Private Const INTERNET_STATE_DISCONNECTED_BY_USER = &H10
Private Const INTERNET_STATE_IDLE = &H100
Private Const INTERNET_STATE_BUSY = &H200
Private Const ISO_FORCE_DISCONNECTED = 1

Private Declare Function InternetSetOption Lib "wininet.dll" Alias _
    "InternetSetOptionA" (ByVal hInternet As Long, ByVal dwOption As Long, _
    lpBuffer As INTERNET_CONNECTED_INFO, ByVal dwBufferLength As Long) As _
    Boolean
'-----------------------------------------------------------------------------
'设为脱机模式
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'断开Internet连接
'-----------------------------------------------------------------------------
'Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
'Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Const IF_FROM_CACHE = &H1000000
Private Const IF_MAKE_PERSISTENT = &H2000000
Private Const IF_NO_CACHE_WRITE = &H4000000
Private Const BUFFER_LEN = 256

Private Const RAS_MAXENTRYNAME As Integer = 256
Private Const RAS_MAXDEVICETYPE As Integer = 16
Private Const RAS_MAXDEVICENAME As Integer = 128
Private Const RAS_RASCONNSIZE As Integer = 412
'Private Const ERROR_SUCCESS = 0&

Private Type RasEntryName
    dwSize As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
End Type

Private Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

Private Declare Function RasEnumConnections Lib _
    "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As _
    Any, lpcb As Long, lpcConnections As Long) As Long

Private Declare Function RasHangUp Lib "rasapi32.dll" Alias _
    "RasHangUpA" (ByVal hRasConn As Long) As Long
Private gstrISPName As String
Private ReturnCode As Long
'-----------------------------------------------------------------------------
'断开Internet连接
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'获取DNS及服务器时间
'-----------------------------------------------------------------------------
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(32) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(32) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Type TIME_OF_DAY_INFO
    tod_elapsedt As Long
    tod_msecs As Long
    tod_hours As Long
    tod_mins As Long
    tod_secs As Long
    tod_hunds As Long
    tod_timezone As Long
    tod_tinterval As Long
    tod_day As Long
    tod_month As Long
    tod_year As Long
    tod_weekday As Long
End Type
'
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function NetRemoteTOD Lib "Netapi32.dll" (byteServer As Any, pBuffer As Long) As Long
'Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function NetApiBufferFree Lib "Netapi32.dll" (ByVal lpBuffer As Long) As Long
'
'
Private Const MAX_HOSTNAME_LEN = 132
Private Const MAX_DOMAIN_NAME_LEN = 132
Private Const MAX_SCOPE_ID_LEN = 260
'Private Const MAX_ADAPTER_NAME_LENGTH = 260
'Private Const MAX_ADAPTER_ADDRESS_LENGTH = 8
'Private Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Private Const ERROR_BUFFER_OVERFLOW = 111
Private Const MIB_IF_TYPE_ETHERNET = 1
Private Const MIB_IF_TYPE_TOKENRING = 2
Private Const MIB_IF_TYPE_FDDI = 3
Private Const MIB_IF_TYPE_PPP = 4
Private Const MIB_IF_TYPE_LOOPBACK = 5
Private Const MIB_IF_TYPE_SLIP = 6
'
'Private Type IP_ADDR_STRING
'    Next As Long
'    IpAddress As String * 16
'    IpMask As String * 16
'    Context As Long
'End Type
'
'Private Type IP_ADAPTER_INFO
'    Next As Long
'    ComboIndex As Long
'    AdapterName As String * MAX_ADAPTER_NAME_LENGTH
'    Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
'    AddressLength As Long
'    Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
'    Index As Long
'    type As Long
'    DhcpEnabled As Long
'    CurrentIpAddress As Long
'    IpAddressList As IP_ADDR_STRING
'    GatewayList As IP_ADDR_STRING
'    DhcpServer As IP_ADDR_STRING
'    HaveWins As Boolean
'    PrimaryWinsServer As IP_ADDR_STRING
'    SecondaryWinsServer As IP_ADDR_STRING
'    LeaseObtained As Long
'    LeaseExpires As Long
'End Type
'
Private Type FIXED_INFO
    HostName As String * MAX_HOSTNAME_LEN
    DomainName As String * MAX_DOMAIN_NAME_LEN
    CurrentDnsServer As Long
    DnsServerList As IP_ADDR_STRING
    NodeType As Long
    ScopeId  As String * MAX_SCOPE_ID_LEN
    EnableRouting As Long
    EnableProxy As Long
    EnableDns As Long
End Type
'
Private Declare Function GetNetworkParams Lib "IPHlpApi" (FixedInfo As Any, pOutBufLen As Long) As Long
'Private Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
'-----------------------------------------------------------------------------
'获取DNS及服务器时间
'-----------------------------------------------------------------------------
'
'
'
''-----------------------------------------------------------------------------
''发邮件
''-----------------------------------------------------------------------------
'Dim sMailSmtpServer As String
'Dim sMailID As String
'Dim sMailPWD As String
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'
''-----------------------------------------------------------------------------
''发邮件
''-----------------------------------------------------------------------------
'
'
'
''-----------------------------------------------------------------------------
''取得MAC
''-----------------------------------------------------------------------------
'Private Const NCBASTAT = &H33
'Private Const NCBNAMSZ = 16
'Private Const HEAP_ZERO_MEMORY = &H8
'Private Const HEAP_GENERATE_EXCEPTIONS = &H4
'Private Const NCBRESET = &H32
'
'Private Type NCB
'    ncb_command As Byte 'Integer
'    ncb_retcode As Byte 'Integer
'    ncb_lsn As Byte 'Integer
'    ncb_num As Byte ' Integer
'    ncb_buffer As Long 'String
'    ncb_length As Integer
'    ncb_callname As String * NCBNAMSZ
'    ncb_name As String * NCBNAMSZ
'    ncb_rto As Byte 'Integer
'    ncb_sto As Byte ' Integer
'    ncb_post As Long
'    ncb_lana_num As Byte 'Integer
'    ncb_cmd_cplt As Byte 'Integer
'    ncb_reserve(9) As Byte ' Reserved, must be 0
'    ncb_event As Long
'End Type
'
'Private Type ADAPTER_STATUS
'    adapter_address(5) As Byte 'As String * 6
'    rev_major As Byte 'Integer
'    reserved0 As Byte 'Integer
'    adapter_type As Byte 'Integer
'    rev_minor As Byte 'Integer
'    duration As Integer
'    frmr_recv As Integer
'    frmr_xmit As Integer
'    iframe_recv_err As Integer
'    xmit_aborts As Integer
'    xmit_success As Long
'    recv_success As Long
'    iframe_xmit_err As Integer
'    recv_buff_unavail As Integer
'    t1_timeouts As Integer
'    ti_timeouts As Integer
'    Reserved1 As Long
'    free_ncbs As Integer
'    max_cfg_ncbs As Integer
'    max_ncbs As Integer
'    xmit_buf_unavail As Integer
'    max_dgram_size As Integer
'    pending_sess As Integer
'    max_cfg_sess As Integer
'    max_sess As Integer
'    max_sess_pkt_size As Integer
'    name_count As Integer
'End Type
'
'Private Type NAME_BUFFER
'    Name As String * NCBNAMSZ
'    name_num As Integer
'    name_flags As Integer
'End Type
'
'Private Type ASTAT
'    adapt As ADAPTER_STATUS
'    NameBuff(30) As NAME_BUFFER
'End Type
'
'Private Declare Function Netbios Lib "Netapi32.dll" (pncb As NCB) As Byte
''Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
'Private Declare Function GetProcessHeap Lib "kernel32" () As Long
'Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
'
''-----------------------------------------------------------------------------
''取得MAC
''-----------------------------------------------------------------------------
'
'
'
''-----------------------------------------------------------------------------
''FTP
''-----------------------------------------------------------------------------
'Dim bPASSIVE As Boolean         '设置是否使用PASSIVE模式
'Private hInternet As Long       ' The handle to the Internet
'Private hConnect As Long        ' The handle to the Internet Connection
'Private sFtpSite1 As String         ' The site to which we connect to
'Private sUserName1 As String         ' The user name used to connect to the site
'Private sPassWord1 As String     ' The password for the user
'Private iPort1 As Long
'
'
'Private Const MAX_PATH  As Long = 260
'Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
'Private Const INTERNET_OPEN_TYPE_DIRECT = 1                         ' direct to net
'Private Const INTERNET_FLAG_PASSIVE = &H8000000             ' used for FTP connections
'Private Const INTERNET_FLAG_RELOAD = &H80000000
'Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000      ' don't write this item to the cache
'Private Const INTERNET_DEFAULT_FTP_PORT = 21                ' default for FTP servers
'Private Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000   ' FTP: use existing InternetConnect handle for server if possible
'Private Const INTERNET_SERVICE_FTP = 1
'Private Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
'Private Const FTP_TRANSFER_TYPE_BINARY = &H2
'
'Private Type FILETIME
'   dwLowDateTime As Long
'   dwHighDateTime As Long
'End Type
'
'Private Type WIN32_FIND_DATA
'   dwFileAttributes As Long
'   ftCreationTime As FILETIME
'   ftLastAccessTime As FILETIME
'   ftLastWriteTime As FILETIME
'   nFileSizeHigh As Long
'   nFileSizeLow As Long
'   dwReserved0 As Long
'   dwReserved1 As Long
'   cFileName As String * MAX_PATH
'   cAlternate As String * 14
'End Type
'
'Private Declare Function InternetOpen Lib "wininet.dll" _
'   Alias "InternetOpenA" _
'  (ByVal lpszAgent As String, _
'   ByVal dwAccessType As Long, _
'   ByVal lpszProxyName As String, _
'   ByVal lpszProxyBypass As String, _
'   ByVal dwFlags As Long) As Long
'
'Private Declare Function InternetCloseHandle Lib "wininet.dll" _
'   (ByVal hEnumHandle As Long) As Long
'
'Private Declare Function InternetConnect Lib "wininet.dll" _
'   Alias "InternetConnectA" _
'  (ByVal hInternet As Long, _
'   ByVal lpszServerName As String, _
'   ByVal nServerPort As Long, _
'   ByVal lpszUserName As String, _
'   ByVal lpszPassword As String, _
'   ByVal dwService As Long, _
'   ByVal dwFlags As Long, _
'   ByVal dwContext As Long) As Long
'
'Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
'Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
'Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
'
'Private Declare Function FtpFindFirstFile Lib "wininet.dll" _
'   Alias "FtpFindFirstFileA" _
'  (ByVal hConnect As Long, _
'   ByVal lpszSearchFile As String, _
'   lpFindFileData As Any, _
'   ByVal dwFlags As Long, _
'   ByVal dwContext As Long) As Long
'
'Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" _
'   Alias "FtpGetCurrentDirectoryA" _
'  (ByVal hConnect As Long, _
'   ByVal lpszCurrentDirectory As String, _
'    lpdwCurrentDirectory As Long) As Long
'
'Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" _
'   Alias "FtpSetCurrentDirectoryA" _
'  (ByVal hConnect As Long, _
'   ByVal lpszDirectory As String) As Long
'
'Private Declare Function FtpGetFile Lib "wininet.dll" _
'   Alias "FtpGetFileA" _
'  (ByVal hConnect As Long, _
'   ByVal lpszRemoteFile As String, _
'   ByVal lpszNewFile As String, _
'   ByVal fFailIfExists As Long, _
'   ByVal dwFlagsAndAttributes As Long, _
'   ByVal dwFlags As Long, _
'   ByVal dwContext As Long) As Long
'
'Private Declare Function FtpDeleteFile Lib "wininet.dll" _
'    Alias "FtpDeleteFileA" _
'    (ByVal hConnect As Long, _
'     ByVal lpszFileName As String) As Long
'
'Private Declare Function FtpPutFile Lib "wininet.dll" _
'    Alias "FtpPutFileA" _
'    (ByVal hFtpSession As Long, _
'     ByVal lpszLocalFile As String, _
'     ByVal lpszRemoteFile As String, _
'     ByVal dwFlags As Long, _
'     ByVal dwContext As Long) As Boolean
'
'Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
'
''-----------------------------------------------------------------------------
''FTP
''-----------------------------------------------------------------------------
'
'
'
'
'Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
'
'
''-----------------------------------------------------------------------------
''断开所有网络设备
''-----------------------------------------------------------------------------
'
'Private Type GUID
'  Data1 As Long
'  Data2 As Integer
'  Data3 As Integer
'  Data4(7) As Byte
'End Type
'
'Private Declare Function SetupDiGetClassDevs Lib "setupapi" Alias "SetupDiGetClassDevsA" (ByVal ClassGuid As Long, ByVal Enumerator As Long, ByVal hWndParent As Long, ByVal flags As Long) As Long
'
'Private Const DIGCF_PRESENT            As Long = &H2
'Private Const DIGCF_ALLCLASSES         As Long = &H4
'
'Private Declare Function SetupDiDestroyDeviceInfoList Lib "setupapi" (ByVal DeviceInfoSet As Long) As Long
'
'Private Declare Function SetupDiEnumDeviceInfo Lib "setupapi" (ByVal DeviceInfoSet As Long, ByVal MemberIndex As Long, DeviceInfoData As SP_DEVINFO_DATA) As Long
'
'Private Type SP_DEVINFO_DATA
'    cbSize As Long
'    ClassGuid As GUID
'    DevInst As Long
'    Reserved As Long
'End Type
'
'Private Declare Function SetupDiGetDeviceRegistryProperty Lib "setupapi" Alias "SetupDiGetDeviceRegistryPropertyA" (ByVal DeviceInfoSet As Long, DeviceInfoData As SP_DEVINFO_DATA, ByVal Property As Long, ByRef PropertyRegDataType As Long, ByVal PropertyBuffer As Long, ByVal PropertyBufferSize As Long, RequiredSize As Long) As Long
'
'Private Const SPDRP_CLASS                        As Long = (&H7)        '// Class (R--tied to ClassGUID)
'Private Const SPDRP_CLASSGUID                    As Long = (&H8)        '// ClassGUID (R/W)
'
'Private Declare Function SetupDiSetClassInstallParams Lib "setupapi" Alias "SetupDiSetClassInstallParamsA" (ByVal DeviceInfoSet As Long, DeviceInfoData As SP_DEVINFO_DATA, ClassInstallParams As SP_CLASSINSTALL_HEADER, ByVal ClassInstallParamsSize As Long) As Long
'
'Private Declare Function SetupDiCallClassInstaller Lib "setupapi" (ByVal InstallFunction As Long, ByVal DeviceInfoSet As Long, DeviceInfoData As SP_DEVINFO_DATA) As Long
'
'Private Type SP_CLASSINSTALL_HEADER
'    cbSize As Long
'    InstallFunction As Long
'End Type
'
'Private Const DIF_PROPERTYCHANGE                    As Long = &H12
'
'Private Type SP_PROPCHANGE_PARAMS
'    ClassInstallHeader As SP_CLASSINSTALL_HEADER
'    StateChange As Long
'    Scope As Long
'    HwProfile As Long
'End Type
'
'Private Const DICS_ENABLE        As Long = &H1
'Private Const DICS_DISABLE       As Long = &H2
'
'Private Const DICS_FLAG_GLOBAL           As Long = &H1         '// make change in all hardware profiles
'Private Const DICS_FLAG_CONFIGSPECIFIC   As Long = &H2         '// make change in specified profile only
'
'Private Const INVALID_HANDLE_VALUE As Long = -1
'Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122
''-----------------------------------------------------------------------------
''断开所有网络设备
''-----------------------------------------------------------------------------
'
'
'
''-----------------------------------------------------------------------------
''检测是否连接INTERNET
''-----------------------------------------------------------------------------'***********************************************************************************************
'Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'
'Private Declare Function InternetGetConnectedState _
'    Lib "wininet.dll" (ByRef lpdwFlags As Long, _
'    ByVal dwReserved As Long) As Long
'    'Local system uses a modem to connect to
'    '     the Internet.
'Private Const INTERNET_CONNECTION_MODEM As Long = &H1
''Local system uses a LAN to connect to t
''     he Internet.
'Private Const INTERNET_CONNECTION_LAN As Long = &H2
''Local system uses a proxy server to con
''     nect to the Internet.
'Private Const INTERNET_CONNECTION_PROXY As Long = &H4
''No longer used.
'Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
'Private Const INTERNET_RAS_INSTALLED As Long = &H10
'Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
'Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
'    'InternetGetConnectedState wrapper funct
'    '     ions
''-----------------------------------------------------------------------------
''检测是否连接INTERNET
''-----------------------------------------------------------------------------
'
'
''-----------------------------------------------------------------------------
''查出IP地址
''-----------------------------------------------------------------------------
''Private Const MAX_ADAPTER_NAME_LENGTH           As Long = 256
''Private Const MAX_ADAPTER_DESCRIPTION_LENGTH    As Long = 128
''Private Const MAX_ADAPTER_ADDRESS_LENGTH        As Long = 8
'Private Const ERROR_SUCCESS                     As Long = 0
'
'Private Type IP_ADDRESS_STRING
'     IpAddr(0 To 15)                    As Byte
'End Type
'
'Private Type IP_MASK_STRING
'     IpMask(0 To 15)                    As Byte
'End Type
'
''Private Type IP_ADDR_STRING
''     dwNext                             As Long
''     IpAddress                          As IP_ADDRESS_STRING
''     IpMask                             As IP_MASK_STRING
''     dwContext                          As Long
''End Type
''
''Private Type IP_ADAPTER_INFO
''  dwNext                                As Long
''  ComboIndex                            As Long
''  sAdapterName(0 To (MAX_ADAPTER_NAME_LENGTH + 3))          As Byte
''  sDescription(0 To (MAX_ADAPTER_DESCRIPTION_LENGTH + 3))   As Byte
''  dwAddressLength                                           As Long
''  sIPAddress(0 To (MAX_ADAPTER_ADDRESS_LENGTH - 1))         As Byte
''  dwIndex                                                   As Long
''  uType                                                     As Long
''  uDhcpEnabled                                              As Long
''  CurrentIpAddress                                          As Long
''  IpAddressList                                             As IP_ADDR_STRING
''  GatewayList                                               As IP_ADDR_STRING
''  DhcpServer                                                As IP_ADDR_STRING
''  bHaveWins                                                 As Long
''  PrimaryWinsServer                                         As IP_ADDR_STRING
''  SecondaryWinsServer                                       As IP_ADDR_STRING
''  LeaseObtained                                             As Long
''  LeaseExpires                                              As Long
''End Type
'
''Private Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
''-----------------------------------------------------------------------------
''查出IP地址
''-----------------------------------------------------------------------------
'
'
'
''-----------------------------------------------------------------------------
''检查是否已连上INTERNET
''-----------------------------------------------------------------------------
'Private Const FLAG_ICC_FORCE_CONNECTION = &H1
'Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
''-----------------------------------------------------------------------------
''检查是否已连上INTERNET
''-----------------------------------------------------------------------------
'
'
''Add this code to a module and set the Project's Startup Object to 'Sub Main'
''    (-> Project Menu -> Project Properties -> General Tab)
'Private Const RESOURCE_CONNECTED As Long = &H1&
'Private Const RESOURCE_GLOBALNET As Long = &H2&
'Private Const RESOURCE_REMEMBERED As Long = &H3&
'Private Const RESOURCEDISPLAYTYPE_DIRECTORY& = &H9
'Private Const RESOURCEDISPLAYTYPE_DOMAIN& = &H1
'Private Const RESOURCEDISPLAYTYPE_FILE& = &H4
'Private Const RESOURCEDISPLAYTYPE_GENERIC& = &H0
'Private Const RESOURCEDISPLAYTYPE_GROUP& = &H5
'Private Const RESOURCEDISPLAYTYPE_NETWORK& = &H6
'Private Const RESOURCEDISPLAYTYPE_ROOT& = &H7
'Private Const RESOURCEDISPLAYTYPE_SERVER& = &H2
'Private Const RESOURCEDISPLAYTYPE_SHARE& = &H3
'Private Const RESOURCEDISPLAYTYPE_SHAREADMIN& = &H8
'Private Const RESOURCETYPE_ANY As Long = &H0&
'Private Const RESOURCETYPE_DISK As Long = &H1&
'Private Const RESOURCETYPE_PRINT As Long = &H2&
'Private Const RESOURCETYPE_UNKNOWN As Long = &HFFFF&
'Private Const RESOURCEUSAGE_ALL As Long = &H0&
'Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&
'Private Const RESOURCEUSAGE_CONTAINER As Long = &H2&
'Private Const RESOURCEUSAGE_RESERVED As Long = &H80000000
'Private Const NO_ERROR = 0
'Private Const ERROR_MORE_DATA = 234                        'L    // dderror
'Private Const RESOURCE_ENUM_ALL As Long = &HFFFF
'
'Private Type NETRESOURCE
'dwScope As Long
'dwType As Long
'dwDisplayType As Long
'dwUsage As Long
'lpLocalName As String
'lpRemoteName As String
'lpComment As String
'lpProvider As String
'End Type
'
'Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
'Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
'Private Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hWnd As Long, ByVal dwType As Long) As Long
'
'
''Get Nic Address-=-----------------------------
''Private Const NCBASTAT = &H33
''Private Const NCBNAMSZ = 16
''Private Const HEAP_ZERO_MEMORY = &H8
''Private Const HEAP_GENERATE_EXCEPTIONS = &H4
''Private Const NCBRESET = &H32
'
''Private Type NCB
''  ncb_command As Byte
''  ncb_retcode As Byte
''  ncb_lsn As Byte
''  ncb_num As Byte
''  ncb_buffer As Long
''  ncb_length As Integer
''  ncb_callname As String * NCBNAMSZ
''  ncb_name As String * NCBNAMSZ
''  ncb_rto As Byte
''  ncb_sto As Byte
''  ncb_post As Long
''  ncb_lana_num As Byte
''  ncb_cmd_cplt As Byte
''  ncb_reserve(9) As Byte ' Reserved, must be 0
''  ncb_event As Long
''End Type
'
''Private Type ADAPTER_STATUS
''    adapter_address(5) As Byte
''    rev_major As Byte
''    reserved0 As Byte
''    adapter_type As Byte
''    rev_minor As Byte
''    duration As Integer
''    frmr_recv As Integer
''    frmr_xmit As Integer
''    iframe_recv_err As Integer
''    xmit_aborts As Integer
''    xmit_success As Long
''    recv_success As Long
''    iframe_xmit_err As Integer
''    recv_buff_unavail As Integer
''    t1_timeouts As Integer
''    ti_timeouts As Integer
''    Reserved1 As Long
''    free_ncbs As Integer
''    max_cfg_ncbs As Integer
''    max_ncbs As Integer
''    xmit_buf_unavail As Integer
''    max_dgram_size As Integer
''    pending_sess As Integer
''    max_cfg_sess As Integer
''    max_sess As Integer
''    max_sess_pkt_size As Integer
''    name_count As Integer
''End Type
'
''Private Type NAME_BUFFER
''    name As String * NCBNAMSZ
''    name_num As Integer
''    name_flags As Integer
''End Type
'
''Private Type ASTAT
''    adapt As ADAPTER_STATUS
''    NameBuff(30) As NAME_BUFFER
''End Type
'
''Private Declare Function Netbios Lib "netapi32.dll" _
'(pncb As NCB) As Byte
'
''Private Declare Sub CopyMemory Lib "kernel32" Alias _
'"RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, _
'ByVal cbCopy As Long)
'
''Private Declare Function GetProcessHeap Lib "kernel32" () _
'As Long
'
''Private Declare Function HeapAlloc Lib "kernel32" _
'(ByVal hHeap As Long, ByVal dwFlags As Long, _
'ByVal dwBytes As Long) As Long
'
''Private Declare Function HeapFree Lib "kernel32" _
'(ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) _
'As Long
''Get Nic-------------------
'
'
'
''###################################设为脱机模式
'Private Type INTERNET_CONNECTED_INFO
'
'    dwConnectedState As Long
'    dwFlags As Long
'End Type
'
'Private Const INTERNET_OPTION_CONNECTED_STATE = 50
'Private Const INTERNET_STATE_CONNECTED = 1
'Private Const INTERNET_STATE_DISCONNECTED = 2
'Private Const INTERNET_STATE_DISCONNECTED_BY_USER = &H10
'Private Const INTERNET_STATE_IDLE = &H100
'Private Const INTERNET_STATE_BUSY = &H200
'Private Const ISO_FORCE_DISCONNECTED = 1
'
'Private Declare Function InternetSetOption Lib "wininet.dll" Alias _
'    "InternetSetOptionA" (ByVal hInternet As Long, ByVal dwOption As Long, _
'    lpBuffer As INTERNET_CONNECTED_INFO, ByVal dwBufferLength As Long) As _
'    Boolean
''###################################设为脱机模式
'
'
''-----------------------------------------------------------------------------
''断开Internet连接
''-----------------------------------------------------------------------------
''Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
''Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'Private Const IF_FROM_CACHE = &H1000000
'Private Const IF_MAKE_PERSISTENT = &H2000000
'Private Const IF_NO_CACHE_WRITE = &H4000000
'Private Const BUFFER_LEN = 256
'
'Private Const RAS_MAXENTRYNAME As Integer = 256
'Private Const RAS_MAXDEVICETYPE As Integer = 16
'Private Const RAS_MAXDEVICENAME As Integer = 128
'Private Const RAS_RASCONNSIZE As Integer = 412
''Private Const ERROR_SUCCESS = 0&
'
'Private Type RasEntryName
'    dwSize As Long
'    szEntryName(RAS_MAXENTRYNAME) As Byte
'End Type
'
'Private Type RasConn
'    dwSize As Long
'    hRasConn As Long
'    szEntryName(RAS_MAXENTRYNAME) As Byte
'    szDeviceType(RAS_MAXDEVICETYPE) As Byte
'    szDeviceName(RAS_MAXDEVICENAME) As Byte
'End Type
'
'Private Declare Function RasEnumConnections Lib _
'    "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As _
'    Any, lpcb As Long, lpcConnections As Long) As Long
'
'Private Declare Function RasHangUp Lib "rasapi32.dll" Alias _
'    "RasHangUpA" (ByVal hRasConn As Long) As Long
'Private gstrISPName As String
'Private ReturnCode As Long
''-----------------------------------------------------------------------------
''断开Internet连接
''-----------------------------------------------------------------------------
'
'-----------------------------------------------------------------------------
'取得当前域名
'-----------------------------------------------------------------------------
Private Type WKSTA_INFO_102
   wki100_platform_id As Long
   pwki100_computername As Long
   pwki100_langroup As Long
   wki100_ver_major As Long
   wki100_ver_minor As Long
   pwki102_lanroot As Long
   wki102_logged_on_users As Long
End Type

Private Declare Function NetWkstaGetInfo Lib "netapi32" (ByVal servername As String, ByVal level As Long, lpBuf As Any) As Long
'Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
'Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'-----------------------------------------------------------------------------
'取得当前域名
'-----------------------------------------------------------------------------
 
Private Type IP_ADDR_STRING1
    Next As Long
    IpAddress As String * 16
    IpMask As String * 16
    Context As Long
End Type

Private Type FIXED_INFO1
    HostName As String * MAX_HOSTNAME_LEN
    DomainName As String * MAX_DOMAIN_NAME_LEN
    CurrentDnsServer As Long
    DnsServerList As IP_ADDR_STRING1
    NodeType As Long
    ScopeId  As String * MAX_SCOPE_ID_LEN
    EnableRouting As Long
    EnableProxy As Long
    EnableDns As Long
End Type



'函数名称:cyGetDomainName
'作用：取得计算机的域名（暂未能确定如果计算机加入域时使用本机帐号登录得到的结果）
'===========================================================
'整理时间：2007-03-13 22:58
'  参数：
'  例子：           Debug.Print cyGetDomainName

'                   WORKGROUP

'===========================================================
Function cyGetDomainName()

    Dim pWrkInfo As Long
    Dim WrkInfo(0) As WKSTA_INFO_102
    Dim sComputerName As String * 255
    Dim sDomain As String * 255
    
    '取得计算机名
    GetComputerName sComputerName, 255
    NetWkstaGetInfo StrConv("\\" & sComputerName, vbUnicode), 102, pWrkInfo
    CopyMemory WrkInfo(0), ByVal pWrkInfo, ByVal Len(WrkInfo(0))
    CopyMemory ByVal sDomain, ByVal WrkInfo(0).pwki100_langroup, ByVal 255
    cyGetDomainName = Split(StrConv(sDomain, vbFromUnicode), Chr(0))(0)
    NetApiBufferFree ByVal pWrkInfo
   
End Function

'断开INTERNET连接
Public Sub cyInternetDisConnect()
    Dim i As Long
    Dim lpRasConn(255) As RasConn
    Dim lpcb As Long
    Dim lpcConnections As Long
    Dim hRasConn As Long
    lpRasConn(0).dwSize = RAS_RASCONNSIZE
    lpcb = RAS_MAXENTRYNAME * lpRasConn(0).dwSize
    lpcConnections = 0
    
    ReturnCode = RasEnumConnections(lpRasConn(0), lpcb, lpcConnections)
    
    If ReturnCode = ERROR_SUCCESS Then
        For i = 0 To lpcConnections - 1
            If Trim(ByteToString(lpRasConn(i).szEntryName)) = Trim(gstrISPName) Then
                hRasConn = lpRasConn(i).hRasConn
                ReturnCode = RasHangUp(ByVal hRasConn)
            End If
        Next i
    End If
End Sub

Private Function ByteToString(bytString() As Byte) As String
    Dim i As Integer
    ByteToString = ""
    i = 0
    While bytString(i) = 0&
        ByteToString = ByteToString & Chr(bytString(i))
        i = i + 1
    Wend
End Function






'过程名称:cyDriveMapDialog

'作用：弹出映射对话框

'##################################################
'整理时间：2004-2-11 16:41:43
'  参数：
'##################################################
'  例：
Sub cyDriveMapDialog()
    WNetConnectionDialog 0, 1
End Sub

'函数名称:cyDriveMap

'作用：将网络路径映射为逻辑盘符

'##################################################
'整理时间：2004-2-11 16:39:47
'  参数：
'   sShareName As String    --共享路径如:\\PC\ShareName
'   sDrive As String        --S:
'   sUserName As String = vbNullString  --用户名(如需要)
'   sPassWord As String = vbNullString  --密码
'   bPersistend As Boolean  --是否常驻
'##################################################
'  例：
Function cyDriveMap(sShareName As String, sDrive As String, Optional ByVal sUserName As String = vbNullString, Optional ByVal sPassWord As String = vbNullString, Optional ByVal bPersistend As Boolean) As Boolean
    Dim Ns As NETRESOURCE
    Dim i As Long
    If Len(sUserName) = 0 Then sUserName = vbNullString
    If Len(sPassWord) = 0 Then sPassWord = vbNullString
    WNetCancelConnection2 sDrive, 1, 1
    sDrive = Replace(Trim(sDrive), ":", "") + ":"
    Ns.lpRemoteName = sShareName
    Ns.lpLocalName = sDrive
    i = WNetAddConnection2(Ns, sPassWord, sUserName, 0)
    If i = 0 Then cyDriveMap = True
End Function

'函数名称:cyCancleDriveMap

'作用：取消逻辑盘的网络连接

'##################################################
'整理时间：2004-2-13 17:54:42
'  参数：
'   sDrive As String        --要取消的驱动盘符
'##################################################
'  例：
Function cyCancleDriveMap(sDrive As String) As Boolean
    sDrive = Replace(Trim(sDrive), ":", "") + ":"
    Dim i As Long
    i = WNetCancelConnection2(sDrive, 1, 1)
    If i = 0 Then cyCancleDriveMap = True
End Function

'函数名称:cyGetNetCardAddressToArray
'作用：返回当前计算机的网卡到字符串数组中（可能有多个网卡）
'===========================================================
'整理时间：2006-9-18 22:55
'  参数：
'===========================================================
Public Function cyGetNetCardAddressToArray()
On Error Resume Next
    
    Dim S As New cySystemEx
    Dim sTemp As String
    Dim bStr As String * 255
    sTemp = Mid(bStr, 1, GetTempPath(255, bStr))
    sTemp = IIf(Right(sTemp, 1) = "\", sTemp, sTemp & "\")
    Shell "cmd /c ipconfig /all>" & sTemp & "NetCard.Txt", vbHide
    Dim F As New cyFileEx
    Dim sStr As String
    Dim sA() As String
    Dim sA1() As String
    Dim sB() As String
    Dim i As Long
    
    '在一秒内检测文件是否生成
    For i = 1 To 5
        S.cyDelay 0.2
        sStr = F.cyFileStrInput(sTemp & "NetCard.Txt")
        If sStr <> "" Then Exit For
    Next
    
    sA = Split(UCase(sStr), "PHYSICAL")
    ReDim sB(UBound(sA) - 1) As String
    
    For i = 1 To UBound(sA)
        Dim sS As New cyStringEx
        sB(i - 1) = Replace(Trim(sS.cyMidEx(sA(i), ":", vbCrLf)), Chr(13), "")
    Next
    
    cyGetNetCardAddressToArray = sB

End Function


'过程名称:cySetOffLineMode

'作用：将IE浏览器设为脱机模式

'##################################################
'整理时间：2004-2-11 13:51:29
'  参数：
'   offLineMode As Boolean  --设置为脱机/联机模式
'##################################################
'  例：
Sub cySetOffLineMode(ByVal offLineMode As Boolean)
    Dim ci As INTERNET_CONNECTED_INFO
    Dim retvalue As Boolean

    If offLineMode Then
        ci.dwConnectedState = INTERNET_STATE_DISCONNECTED_BY_USER
        ci.dwFlags = ISO_FORCE_DISCONNECTED
    Else
        ci.dwConnectedState = INTERNET_STATE_CONNECTED
    End If

    retvalue = InternetSetOption(0&, INTERNET_OPTION_CONNECTED_STATE, ci, _
        LenB(ci))
    If retvalue = False Then Err.Raise vbObjectError + 1000, , _
        "设置脱机模式时出错！"
End Sub

'函数名称:cyGetIP

'作用：获得IP地址[如果包含公网IP则用；隔开]

'##################################################
'整理时间：2004-5-11 上午 09:07:47
'  参数：
'##################################################
'  例：

'利用MODEM拨电话
''    Dim N As New cyNetWork
''    N.cyDialPhoneNumber MSComm1, 4, "13502413547"
Sub cyDialPhoneNumber(MSCOMM As Object, iCommPort As Long, sPhoneNumber As String)
    '清除接收缓冲区
    MSCOMM.InBufferCount = 0
    
    
    '打开端口
    If MSCOMM.PortOpen = False Then
        MSCOMM.CommPort = iCommPort
        MSCOMM.PortOpen = True
    End If
    '拨电话号码
    MSCOMM.Output = "ATDT" & sPhoneNumber & ";" + vbCr
End Sub

Private Function Trans2Tel(sTelNum As String) As String
Dim iTelLen As Integer
Dim sLeftCode, sRightCode, sAllCode As String
Dim i As Integer
sAllCode = ""
iTelLen = 13
    sTelNum = "86" & sTelNum
For i = 1 To 13 Step 2
  sLeftCode = Mid(sTelNum, i, 1)
  sRightCode = Mid(sTelNum, i + 1, 1)
  If i = iTelLen Then sRightCode = "F"
  sAllCode = sAllCode & sRightCode & sLeftCode
Next i
Trans2Tel = sAllCode
End Function

Private Function SendSMS(MSCOMM As Object, iCommPort As Long, sSMSCenterNum As String, sReceiverNum As String, sMsg As String, Optional ByRef sErr As String) As Boolean
On Error GoTo Err
    Dim sS As New cyStringEx
    Dim S As New cySystemEx
    
    '格式化短信中心号码
    sSMSCenterNum = Replace(sSMSCenterNum, "+86", "")
    sSMSCenterNum = Replace(sSMSCenterNum, "86", "")
    If Len(sSMSCenterNum) = 12 And Left(sSMSCenterNum, 1) = "0" Then
        sSMSCenterNum = Replace(sSMSCenterNum, "0", "", , 1)
    End If
    
    sReceiverNum = Replace(sReceiverNum, "+86", "")
    sReceiverNum = Replace(sReceiverNum, "86", "")
    If Len(sReceiverNum) = 12 And Left(sReceiverNum, 1) = "0" Then
        sReceiverNum = Replace(sReceiverNum, "0", "", , 1)
    End If
    
    Dim iLenMsg As String
    Dim sPDU, sSMSCenterNumCode, sRecNumCode, psMsg
'将短信号码转换为特定的码
    sSMSCenterNumCode = Trim(Trans2Tel(sSMSCenterNum))
'将手机号码转换为特定的码
    sRecNumCode = Trim(Trans2Tel(sReceiverNum))
'打开指定的端口
    If MSCOMM.PortOpen = False Then
        MSCOMM.CommPort = iCommPort
    ElseIf MSCOMM.PortOpen = True Then
    '如果本来已打开则先关闭
        MSCOMM.PortOpen = False
    End If
    '重新打开
    MSCOMM.PortOpen = True
    MSCOMM.Output = "AT+CMGF=0" + vbCr
    
    psMsg = Trim(sS.cyFormatString2HexString(sMsg))
    iLenMsg = Hex(Len(sMsg) * 2)
    iLenMsg = IIf(Len(iLenMsg) = 1, "0" & iLenMsg, iLenMsg)

                    '短信中心号码                    '目标手机号码            '字符串长度的两倍(中英文各当1个)
    sPDU = "0891" & sSMSCenterNumCode & "11000D91" & sRecNumCode & "000800" & iLenMsg & psMsg
    'MSCOMM.Output = "AT+CMGS=" & Trim(Str(15 + Len(sMsg) * 2)) + vbCr
    MSCOMM.Output = "AT+CMGS=" & Trim(Len("11000D91" & sRecNumCode & "000800" & "  " & psMsg) / 2) + vbCr
    S.cyDelay 1
    MSCOMM.Output = sPDU & Chr$(26)
    SendSMS = True
    Exit Function
Err:
    sErr = Err.Description
    SendSMS = False
'''    psMsg = Trim(sS.cyFormatString2HexString(sMsg))
'''    iLenMsg = Hex(Len(sMsg) * 2)
'''    iLenMsg = IIf(Len(iLenMsg) = 1, "0" & iLenMsg, iLenMsg)
'''
'''                    '短信中心号码                    '目标手机号码            '字符串长度的两倍(中英文各当1个)
'''    sPDU = "0891" & sSMSCenterNumCode & "11000D91" & sRecNumCode & "000800" & iLenMsg & psMsg
'''    S.cyDelay 0.1
'''    MSCOMM.Output = "AT+CMGS=" & Trim(Str(15 + Len(sMsg) * 2)) + vbCr
'''    S.cyDelay 0.1
'''    MSCOMM.Output = sPDU & Chr$(26)
End Function



Public Function cySendSMS(MSCOMM As Object, ByVal iCommPort As Long, ByVal sSMSCenterNum As String, ByVal sReceiverNum As String, ByVal sMsg As String, Optional ByRef sErr As String) As Boolean
On Error GoTo Err
    Dim sS As New cyStringEx
    Dim S As New cySystemEx
    '将它格式为某个长度的段再进行发送
    Dim sTemp As String
    Dim i As Long
    Dim sArray() As String
    Dim sA() As String
    
    '将回车替换为两个空格
    sMsg = Replace(sMsg, Chr(13) & Chr(10), "  ")
    
If Len(sMsg) = sS.cyLenEx(sMsg) Then
'长度与实际长度相同，则表示是全英文字符串

Else
'不相同则表示中英文混合
    If Len(sMsg) > 70 Then   '号码超过70个中文字
    '截断后发送,末尾+[待续]
        ReDim Preserve sA(0)
        '将其格式为等长的字段,减三个字留作"[待续]"
        sA(0) = Mid(sMsg, 1, 63)
        sMsg = Right(sMsg, Len(sMsg) - 63)
        
        Do While Len(sMsg) > 63
            ReDim Preserve sA(UBound(sA) + 1)
            sA(UBound(sA)) = Mid(sMsg, 1, 63)
            sMsg = Right(sMsg, Len(sMsg) - 63)
        Loop
        '当剩下的字符串少于等于70字时则直接写进去不加入待续
        ReDim Preserve sA(UBound(sA) + 1)
        sA(UBound(sA)) = sMsg
        
        
        
        
        '末尾加上[1/2]
        For i = 0 To UBound(sA)
            sA(i) = sA(i) & "[" & Trim(CStr(i + 1)) & "/" & Trim(CStr(UBound(sA) + 1)) & "]"
            Debug.Print sA(i)
            
            
            '如果发送不成功则返回假值
            If Not SendSMS(MSCOMM, iCommPort, sSMSCenterNum, sReceiverNum, sA(i), sErr) Then
                Exit Function
            End If
            S.cyDelay 10
        Next
        '成功发送
        cySendSMS = True
        
'''        Exit Function
'''
'''        sTemp = sS.cyFixTxtLength(sMsg, 63)
'''        sA = Split(sTemp, Chr(13) & Chr(10))
'''        For i = 0 To UBound(sA) - 1
'''            '将~~转回回车值
'''            sTemp = Replace(sA(i), "~~", Chr(13) & Chr(10)) & "[待续]"
'''
'''        Next
'''        '上面是每条信息后加[待续],最后那一条就不用了
'''
'''            sTemp = Replace(sA(i), "~~", Chr(13) & Chr(10))
'''            If Not SendSMS(MSCOMM, iCommPort, sSMSCenterNum, sReceiverNum, sTemp, sErr) Then
'''                Exit Function
'''            End If
            
    
    Else    '小于70字,直接发送
        If SendSMS(MSCOMM, iCommPort, sSMSCenterNum, sReceiverNum, sMsg, sErr) Then
        '成功发送
            cySendSMS = True
            Exit Function
        End If
    End If
End If
    
If sS.cyLenEx(sMsg) > 140 Then   '号码超过70个中文字
'截断后发送,末尾+[待续]

    '将其格式为等长的字段,减三个字留作"[待续]"
    sTemp = sS.cyFixTxtLength(sMsg, 64)
    sA = Split(sTemp, Chr(13) & Chr(10))
    For i = 0 To UBound(sA) - 1
        '将~~转回回车值
        sTemp = Replace(sA(i), "~~", Chr(13) & Chr(10)) & "[待续]"
        
        '如果发送不成功则返回假值
        If Not SendSMS(MSCOMM, iCommPort, sSMSCenterNum, sReceiverNum, sTemp, sErr) Then
            Exit Function
        End If
        S.cyDelay 10
    Next
    '上面是每条信息后加[待续],最后那一条就不用了
    
        sTemp = Replace(sA(i), "~~", Chr(13) & Chr(10))
        If Not SendSMS(MSCOMM, iCommPort, sSMSCenterNum, sReceiverNum, sTemp, sErr) Then
            Exit Function
        End If
        
        '成功发送
        cySendSMS = True

Else    '小于70字,直接发送
    If SendSMS(MSCOMM, iCommPort, sSMSCenterNum, sReceiverNum, sMsg, sErr) Then
    '成功发送
        cySendSMS = True
        Exit Function
    End If
End If

    
Err:
    sErr = Err.Description
    cySendSMS = False
End Function

Public Function cyInternetOnline() As Boolean
    cyInternetOnline = InternetGetConnectedState(0&, 0&)
End Function

Public Function cyNetworkEnable(Optional bEnabledNetCard As Boolean = True) As Boolean

    On Error GoTo Error_Handler
    Dim hDevInfo As Long
    Dim W As New cyWndEx
    Dim S As New cySystemEx
    hDevInfo = SetupDiGetClassDevs(0, 0, 0, DIGCF_PRESENT + DIGCF_ALLCLASSES)
    If hDevInfo <> INVALID_HANDLE_VALUE Then
        Dim DeviceInfoData  As SP_DEVINFO_DATA
        Dim i As Long
        Dim bDevInfo() As Byte
        Dim sDeviceDescription As String, sDeviceClassName As String, sDeviceEnumeratorName As String, sDeviceFriendlyName As String, sDeviceClassGuid As String
        DeviceInfoData.cbSize = Len(DeviceInfoData)
        i = 0
        Do While (SetupDiEnumDeviceInfo(hDevInfo, i, DeviceInfoData) <> 0)
            DoEvents
            sDeviceClassName = ""
            If GetDeviceRegInfo(hDevInfo, DeviceInfoData, SPDRP_CLASS, bDevInfo()) Then
                sDeviceClassName = StrConv(bDevInfo, vbUnicode)
            End If
            
            '是网络设备
            If InStr(1, sDeviceClassName, "Net") > 0 Then
                '获得类ID
                If GetDeviceRegInfo(hDevInfo, DeviceInfoData, SPDRP_CLASSGUID, bDevInfo()) Then
                    sDeviceClassGuid = StrConv(bDevInfo, vbUnicode)
                End If
                DoEvents
                SetDeviceState i, sDeviceClassGuid, bEnabledNetCard
                DoEvents
            End If
            i = i + 1
        Loop
                    
        Call SetupDiDestroyDeviceInfoList(hDevInfo)
    End If

    Exit Function
Error_Handler:
End Function

Private Function SetDeviceState(ByVal lDeviceIndex As Long, ByVal sDeviceClassGuid As String, Optional ByVal fEnabled As Boolean = True) As Boolean
    Dim hDevInfo As Long
    hDevInfo = SetupDiGetClassDevs(0, 0, 0, DIGCF_PRESENT + DIGCF_ALLCLASSES)
    If hDevInfo <> INVALID_HANDLE_VALUE Then
        Dim DeviceInfoData  As SP_DEVINFO_DATA
        Dim bDevInfo() As Byte
        
        DeviceInfoData.cbSize = Len(DeviceInfoData)
        If SetupDiEnumDeviceInfo(hDevInfo, lDeviceIndex, DeviceInfoData) <> 0 Then
            If GetDeviceRegInfo(hDevInfo, DeviceInfoData, SPDRP_CLASSGUID, bDevInfo()) Then
                If sDeviceClassGuid = StrConv(bDevInfo, vbUnicode) Then
                    Dim PropChangeParams  As SP_PROPCHANGE_PARAMS
                    
                    With PropChangeParams
                        .ClassInstallHeader.cbSize = Len(.ClassInstallHeader)
                        .ClassInstallHeader.InstallFunction = DIF_PROPERTYCHANGE
                        .Scope = DICS_FLAG_GLOBAL
                        .StateChange = IIf(fEnabled, DICS_ENABLE, DICS_DISABLE)
                        .HwProfile = 0
                    End With
                    
                    If SetupDiSetClassInstallParams(hDevInfo, DeviceInfoData, PropChangeParams.ClassInstallHeader, Len(PropChangeParams)) <> 0 Then
                        With PropChangeParams
                            .ClassInstallHeader.cbSize = Len(.ClassInstallHeader)
                            .ClassInstallHeader.InstallFunction = DIF_PROPERTYCHANGE
                            .Scope = DICS_FLAG_CONFIGSPECIFIC
                            .StateChange = IIf(fEnabled, DICS_ENABLE, DICS_DISABLE)
                            .HwProfile = 0
                        End With
                        If SetupDiSetClassInstallParams(hDevInfo, DeviceInfoData, PropChangeParams.ClassInstallHeader, Len(PropChangeParams)) <> 0 Then
                            If SetupDiCallClassInstaller(DIF_PROPERTYCHANGE, hDevInfo, DeviceInfoData) <> 0 Then
                                SetDeviceState = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
                    
        Call SetupDiDestroyDeviceInfoList(hDevInfo)
    End If
    Exit Function
Error_Handler:
End Function

Private Function GetDeviceRegInfo(ByVal hDevInfo As Long, DeviceInfoData As SP_DEVINFO_DATA, ByVal lPropertyName As Long, bDevInfo() As Byte) As Boolean
    Dim lBufferSize As Long
    Dim lRegDataType As Long
    Call SetupDiGetDeviceRegistryProperty(hDevInfo, DeviceInfoData, lPropertyName, lRegDataType, 0, 0, lBufferSize)
    If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
        ReDim bDevInfo(lBufferSize - 1)
        Call SetupDiGetDeviceRegistryProperty(hDevInfo, DeviceInfoData, lPropertyName, lRegDataType, VarPtr(bDevInfo(0)), lBufferSize, ByVal 0)
        GetDeviceRegInfo = True
    End If
End Function

'根据句柄返回URL
Function cyIeGetUrl(ByVal hWnd As Long) As String
On Error Resume Next
    Dim browser
    Dim IEList As New ShellWindows
    '遍历当前地浏览器窗口
    For Each browser In IEList
        If browser.hWnd = hWnd Then
        '找到要找的浏览器，则返回其URL
            cyIeGetUrl = browser.LocationURL
            
        End If
    Next
End Function

'根据句柄返回IE的字符
Function cyIeGetText(ByVal hWnd As Long) As String
On Error Resume Next
    Dim browser
    Dim IEList As New ShellWindows
    '遍历当前地浏览器窗口
    For Each browser In IEList
        If browser.hWnd = hWnd Then
        '找到要找的浏览器，则返回其URL
            cyIeGetText = browser.Document.Body.innertext
            
        End If
    Next
End Function

'根据句柄返回IE的HTML
Function cyIeGetHtml(ByVal hWnd As Long) As String
On Error Resume Next
    Dim browser
    Dim IEList As New ShellWindows
    '遍历当前地浏览器窗口
    For Each browser In IEList
        If browser.hWnd = hWnd Then
        '找到要找的浏览器，则返回其URL
            cyIeGetHtml = browser.Document.Body.innerhtml
            
        End If
    Next
End Function

'将某个IE指向某个URL
Sub cyIeSetURL(ByVal hWnd As Long, ByVal sURL As String)
On Error Resume Next
    Dim browser
    Dim IEList As New ShellWindows
    '遍历当前地浏览器窗口
    For Each browser In IEList
        If browser.hWnd = hWnd Then
        '找到要找的浏览器，则返回其URL
            browser.Navigate2 sURL
        End If
    Next
End Sub

'向IE发送文本框发送内容
'Call cyIePutText(1442364, "LoginName", "cy")
Sub cyIePutText(ByVal hWnd As Long, ByVal sTextBoxName As String, ByVal sStr As String)
On Error Resume Next
    Dim browser
    Dim vDoc
    Dim vTag
    Dim i As Long
    Dim IEList As New ShellWindows
    Dim iHwnd  As Long
    '遍历当前地浏览器窗口
    For Each browser In IEList
        iHwnd = browser.hWnd
        
        If iHwnd = hWnd Then
        '找到要找的浏览器
            Set vDoc = browser.Document
            vDoc.Body.All(sTextBoxName).Value = sStr

            For i = 0 To vDoc.All.Length - 1 '检测所有标签
                Set vTag = vDoc.All(i)
                 
                If UCase(vDoc.All(i).tagName) = "INPUT" Then '找到input标签
                    Set vTag = vDoc.All(i)
                    '如果控件的名称符合则
                    If vTag.name = sTextBoxName Then
                        vTag.Value = sStr
                        Exit Sub
                    End If
                    
                ElseIf UCase(vDoc.All(i).tagName) = "TEXTAREA" Then '找到input标签
                    Set vTag = vDoc.All(i)
                    '如果控件的名称符合则
                    If vTag.name = sTextBoxName Then
                        vTag.Value = sStr
                        Exit Sub
                    End If
                
                End If
                
            Next
        End If
    Next
    
End Sub

Sub cyIeClickButton(ByVal hWnd As Long, Optional ByVal sValue As String)
On Error Resume Next
    Dim browser
    Dim vDoc
    Dim vTag
    Dim i As Long
    Dim IEList As New ShellWindows
    '遍历当前地浏览器窗口
    For Each browser In IEList
        If browser.hWnd = hWnd Then
        '找到要找的浏览器
            Set vDoc = browser.Document
            For i = 0 To vDoc.All.Length - 1 '检测所有标签
                    Set vTag = vDoc.All(i)
                If UCase(vDoc.All(i).tagName) = "INPUT" Then '找到input标签
                    Set vTag = vDoc.All(i)
                    '如果控件的名称符合则
                    If InStr(1, vTag.Value, sValue) > 0 Then
                        vTag.Select
                        vTag.Click
                        Exit Sub
                    End If
                End If
            Next
        End If
    Next
End Sub

Public Function cyDownloadFile(ByVal sURL As String, ByVal sFileName As String) As Boolean
    Dim lngRetval As Long
    DeleteUrlCacheEntry sURL
    lngRetval = URLDownloadToFile(0, sURL, sFileName, 0, 0)
    If lngRetval = 0 Then cyDownloadFile = True
End Function

'-----------------------------------------------------------------------------
'将数据PUT到网页上
'-----------------------------------------------------------------------------
Private Sub PackBytes(byteArray() As Byte, ByVal PostData As String)
    Dim iNewBytes
    Dim i As Long
    iNewBytes = Len(PostData) - 1
    If iNewBytes < 0 Then Exit Sub
    ReDim byteArray(iNewBytes)
    For i = 0 To iNewBytes
       ch = Mid(PostData, i + 1, 1)
       If ch = Space(1) Then
          ch = "+"
       End If
       byteArray(i) = Asc(ch)
    Next
End Sub
 
Sub cyPostDataToURL(ByVal sURL As String, ByVal sData As String)
    Dim g_oIE As InternetExplorer
    Dim vPost As Variant
    Dim vHeaders As Variant
    Set g_oIE = New InternetExplorer
    g_oIE.Visible = True
    ReDim aByte(0) As Byte
    cPostData = sData
    PackBytes aByte(), cPostData
    vPost = aByte
    vHeaders = "Content-Type: application/x-www-form-urlencoded" + Chr(10) + Chr(13)
    g_oIE.Navigate sURL, , , vPost, vHeaders
 
End Sub
'-----------------------------------------------------------------------------
'将数据PUT到网页上
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'FTP
'-----------------------------------------------------------------------------

Public Function cyFtpConnect(ByVal sFtpSite As String, ByVal sUserName As String, ByVal sPassWord As String, Optional ByVal iPort As Long = 21, Optional ByVal Directory As String, Optional ByVal bPassiveMode As Boolean = True) As Boolean
    bPASSIVE = bPassiveMode
    sFtpSite1 = sFtpSite
    sUserName1 = sUserName
    sPassWord1 = sPassWord
    iPort1 = iPort
    cyFtpConnect = False
    
    hInternet = InternetOpen("Media Guide Import App", INTERNET_OPEN_TYPE_DIRECT, _
        vbNullString, vbNullString, INTERNET_FLAG_NO_CACHE_WRITE)
    
    If hInternet Then
        'Get a handle to the connection
        hConnect = GetInternetConnectionHandle
    
        If hConnect <> 0 Then
            cyFtpConnect = True
            
            ' Set the intial path
            FtpSetCurrentDirectory hConnect, Directory
            
            ' Check if the initial path was set
            If Right(Replace(Directory, "\", "/"), 1) <> "/" Then Directory = Directory & "/"
            If UCase(GetFTPDirectory(hConnect)) <> UCase(Directory) Then
            End If
        End If
    End If
End Function
Public Function cyFtpRemoveDirectory(ByVal sDirectory As String) As Boolean
    Dim sA() As String
    sA = cyFtpEnumFiles(sDirectory)
    Dim i As Long
    If UBound(sA) > 1 Then
        For i = 2 To UBound(sA)
            '不能删除,则可能是目录
            If Not cyFtpDeleteFile(sDirectory & "/" & sA(i)) Then
                Call cyFtpRemoveDirectory(sDirectory & "/" & sA(i))
                Call FtpRemoveDirectory(hConnect, sDirectory & "/" & sA(i))
            End If
        Next
    End If
    cyFtpRemoveDirectory = FtpRemoveDirectory(hConnect, sDirectory)
End Function

Public Function cyFtpDisConnect() As Boolean
    
    ' Close the handles
    InternetCloseHandle hConnect
    InternetCloseHandle hInternet
    
    ' Set the handles to 0
    hInternet = 0
    hConnect = 0
    cyFtpDisConnect = True

End Function

Public Function cyFtpCreateDirectory(ByVal sDirectory As String) As Boolean
    
    '格式化路径
    sDirectory = Replace(sDirectory, "\", "/")
    If Left(sDirectory, 1) = "/" Then sDirectory = Right(sDirectory, Len(sDirectory) - 1)
    
    '生成多层目录
    Dim sA() As String
    Dim i As Long
    Dim sTemp As String
    
    sA = Split(sDirectory, "/")
    For i = 0 To UBound(sA) - 1
        sTemp = sTemp & "/" & sA(i)
        Call FtpCreateDirectory(hConnect, sTemp)
        
    Next
    
    cyFtpCreateDirectory = FtpCreateDirectory(hConnect, sDirectory)
End Function

Public Function cyFtpEnumFiles(Optional ByVal sDirectory As String = "/")
    Dim sFileName() As String
    Dim i As Long
    Dim pData As WIN32_FIND_DATA, hFind As Long, lRet As Long
    Dim tmrSystem As SYSTEMTIME
    
    pData.cFileName = String(256, 0)
    sDirectory = Replace(sDirectory, "\", "/")
    
    '枚举第一个文件
    hFind = FtpFindFirstFile(hConnect, Replace(sDirectory & "/*.*", "//", "/"), pData, 0, 0)
    
    '如果句柄不存在则退出函数
    If hFind = 0 Then
        '保存文件名到数组
        ReDim Preserve sFileName(0)
        sFileName(0) = ""
        GoTo ExitProc
    End If
    
    '保存文件名到数组
    ReDim Preserve sFileName(i)
    
    sFileName(i) = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
    
    Do
        '清空原内容
        pData.cFileName = ""
        '查找下一个文件
        lRet = InternetFindNextFile(hFind, pData)
        '如果句柄不存在则退出函数
        If lRet = 0 Then Exit Do
        i = i + 1
        ReDim Preserve sFileName(i)
        FileTimeToSystemTime pData.ftLastWriteTime, tmrSystem
        
        sFileName(i) = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        
    Loop
    
ExitProc:
    cyFtpEnumFiles = sFileName
    '关闭句柄
    InternetCloseHandle hFind
End Function

'函数名称:cyFtpEnumFilesInfoToRs
'作用：枚举当前目录下的所有目录及文件，并返回一个数据集
'===========================================================
'整理时间：2007-03-19 11:51
'  参数：
'  例子：           Debug.Print N.cyGetTimeFromServer("127.0.0.1")
'                   2007-1-16 21:46:46

'===========================================================
Public Function cyFtpEnumFilesInfoToRs(Optional ByVal sDirectory As String = "/") As Recordset
    Dim i As Long
    Dim pData As WIN32_FIND_DATA, hFind As Long, lRet As Long
    Dim tmrSystem As SYSTEMTIME
    
    Dim FilesRs As New Recordset
    FilesRs.CursorLocation = adUseClient
    FilesRs.Fields.Append "名称", adVarChar, 255
    FilesRs.Fields.Append "大小", adUnsignedBigInt, 255
    FilesRs.Fields.Append "更改日期", adFileTime
    FilesRs.Fields.Append "属性", adSmallInt
    FilesRs.Open
    
    pData.cFileName = String(256, 0)
    
    sDirectory = Replace(sDirectory, "\", "/")
    
    '枚举第一个文件
    hFind = FtpFindFirstFile(hConnect, Replace(sDirectory & "/*.*", "//", "/"), pData, 0, 0)
    
    '如果句柄不存在则退出函数
    If hFind = 0 Then
        GoTo ExitProc
    End If
    
    On Error Resume Next    '忽略根目录的时间错误
    FilesRs.AddNew
    FilesRs("名称") = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
    FileTimeToSystemTime pData.ftLastWriteTime, tmrSystem
    FilesRs("更改日期") = tmrSystem.wYear & "-" & tmrSystem.wMonth & "-" & tmrSystem.wDay & " " & tmrSystem.wHour & ":" & tmrSystem.wMinute & ":" & tmrSystem.wSecond
    FilesRs("属性") = pData.dwFileAttributes
    FilesRs("大小") = pData.nFileSizeLow
    
    Do
        '清空原内容
        pData.cFileName = ""
        '查找下一个文件
        lRet = InternetFindNextFile(hFind, pData)
        '如果句柄不存在则退出函数
        If lRet = 0 Then Exit Do
        i = i + 1
        
        FilesRs.AddNew
        FilesRs("名称") = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        FileTimeToSystemTime pData.ftLastWriteTime, tmrSystem
        FilesRs("更改日期") = tmrSystem.wYear & "-" & tmrSystem.wMonth & "-" & tmrSystem.wDay & " " & tmrSystem.wHour & ":" & tmrSystem.wMinute & ":" & tmrSystem.wSecond
        FilesRs("属性") = pData.dwFileAttributes
        FilesRs("大小") = pData.nFileSizeLow
        
    Loop
    FilesRs.MoveFirst
    Set cyFtpEnumFilesInfoToRs = FilesRs
ExitProc:
    
    '关闭句柄
    InternetCloseHandle hFind
End Function

Public Function cyFtpGetCurrentDirectory() As String
    Dim sDirectory As String * 255
    FtpGetCurrentDirectory hConnect, sDirectory, 255
    cyFtpGetCurrentDirectory = Left(sDirectory, InStr(1, sDirectory, Chr(0)) - 1)
End Function


Public Function cyFtpSetCurrentDirectory(ByVal sDirectory As String) As Boolean
    cyFtpSetCurrentDirectory = FtpSetCurrentDirectory(hConnect, sDirectory)
End Function

Public Function cyFtpDownloadFile(sRemoteFile As String, sLocalFileName As String) As Boolean
    Dim sCurrentDir As String
    
    'Only if a valid connection...
    If hConnect Then
        ' Get the Current Internet directory
        sCurrentDir = GetFTPDirectory(hConnect)
        
        ' Setup the file name to retrieve and the destination
        sRemoteFile = sCurrentDir & sRemoteFile
        
        ' Download file
        If FtpGetFile(hConnect, sRemoteFile, sLocalFileName, False, FILE_ATTRIBUTE_ARCHIVE, _
            FTP_TRANSFER_TYPE_UNKNOWN, 0&) Then
            ' Success
            cyFtpDownloadFile = True
        End If
    End If
End Function

Public Function cyFtpUploadFile(ByVal LocalFile As String, ByVal RemoteFile As String) As Boolean
    
    '先尝试上传文件
    If FtpPutFile(hConnect, LocalFile, RemoteFile, FTP_TRANSFER_TYPE_BINARY, 0) Then
        cyFtpUploadFile = True
        
    '上传失败则生成路径
    Else
        '格式化路径
        RemoteFile = Replace(RemoteFile, "\", "/")
        If Left(RemoteFile, 1) = "/" Then RemoteFile = Right(RemoteFile, Len(RemoteFile) - 1)
        
        '生成多层目录
        Dim sA() As String
        Dim i As Long
        Dim sTemp As String
        
        sA = Split(RemoteFile, "/")
        For i = 0 To UBound(sA) - 1
            sTemp = sTemp & "/" & sA(i)
            Call cyFtpCreateDirectory(sTemp)
            
        Next
        
        '再次尝试上传文件
        If FtpPutFile(hConnect, LocalFile, RemoteFile, FTP_TRANSFER_TYPE_BINARY, 0) Then
            cyFtpUploadFile = True
        End If
    End If
    
End Function

Public Function cyFtpDeleteFile(sRemoteFile As String) As Boolean
    cyFtpDeleteFile = FtpDeleteFile(hConnect, sRemoteFile)
End Function

Public Function cyFtpRenameFile(sRemoteName As String, sNewRemoteName As String) As Boolean
    cyFtpRenameFile = FtpRenameFile(hConnect, sRemoteName, sNewRemoteName)
End Function

'得到该文件的文件大小
Public Function cyFtpGetFileSize(Optional ByVal sFileName As String) As Long
    
    Dim sTemp As String
    Dim pData As WIN32_FIND_DATA, hFind As Long, lRet As Long
    
    pData.cFileName = String(256, 0)
    
    sTemp = UCase(Right(sFileName, Len(sFileName) - InStrRev(sFileName, "/")))
    
    hConnect = GetInternetConnectionHandle
    
    '枚举第一个文件
    hFind = FtpFindFirstFile(hConnect, Left(sFileName, InStrRev(sFileName, "/")) & "*.*", pData, 0, 0)
    
    '如果句柄不存在则退出函数
    If hFind = 0 Then
        '保存文件名到数组
        GoTo ExitProc
    End If
    
    Do
        '清空原内容
        pData.cFileName = ""
        '查找下一个文件
        lRet = InternetFindNextFile(hFind, pData)
        '如果句柄不存在则退出函数
        If lRet = 0 Then Exit Do
        If UCase(Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)) = sTemp Then
              cyFtpGetFileSize = (pData.nFileSizeHigh * MAXDWORD) + pData.nFileSizeLow
              Exit Function
        End If
    Loop
    
ExitProc:
    '关闭句柄
    InternetCloseHandle hFind
End Function

Public Function cyFtpFileExist(sRemoteFile As String) As Boolean
    Dim WFD As WIN32_FIND_DATA
    Dim sPath As String
    Dim hFindConnect As Long
    Dim hFind As Long
    
    ' Get the current path
    sPath = GetFTPDirectory(hConnect)
     
    ' Set the path to the file we are searching for
    sPath = sPath & sRemoteFile
        
    ' Get a new connection to the site
    hFindConnect = GetInternetConnectionHandle()
     
    If hFindConnect Then
     
        ' Get a handle to the file we are trying to open
        hFind = FtpFindFirstFile(hFindConnect, sPath, WFD, _
            INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0&)
     
        ' If we retrive a handle then the file is found
        If hFind Then cyFtpFileExist = True
        
    End If
     
    ' Close handles
    InternetCloseHandle hFind
    InternetCloseHandle hFindConnect
End Function

Private Function GetInternetConnectionHandle() As Long
    Dim lConn As Long
    
    If hInternet Then
        lConn = InternetConnect(hInternet, sFtpSite1, iPort1, _
            sUserName1, sPassWord1, INTERNET_SERVICE_FTP, _
            INTERNET_FLAG_EXISTING_CONNECT Or IIf(bPASSIVE, INTERNET_FLAG_PASSIVE, 0), &H0)
    End If
  
    ' Return the handle
    GetInternetConnectionHandle = lConn
End Function

Private Function GetFTPDirectory(hConnect As Long) As String
    Dim nCurrDir As Long
    Dim sCurrDir As String
    
    'pad the requisite buffers
    sCurrDir = Space$(256)
    nCurrDir = Len(sCurrDir)
      
    'WinInetAPI call returns 1 if successful
    If FtpGetCurrentDirectory(hConnect, sCurrDir, nCurrDir) = 1 Then
        'return a properly qualified path
        sCurrDir = Left(sCurrDir, InStr(1, sCurrDir, Chr(0)) - 1)
        GetFTPDirectory = IIf(Right(sCurrDir, 1) = "/", sCurrDir, sCurrDir & "/")
    End If
    
End Function


'-----------------------------------------------------------------------------
'FTP
'-----------------------------------------------------------------------------

Sub cyPopupIE(ByVal sURL As String, Optional ByVal iWidth As Long, Optional ByVal iHeight As Long, Optional ByVal iLeft As Long, Optional ByVal iTop As Long, Optional ByVal bMenuBar As Boolean = True, Optional ByVal bToolBar As Boolean = True)
    
    Dim IE As InternetExplorer
    Set IE = New InternetExplorer
    
    IE.Navigate sURL
    IE.MenuBar = bMenuBar
    IE.ToolBar = bToolBar
    
    '自定义则按自定义
    If iWidth > 0 And iHeight > 0 Then
        IE.Width = iWidth
        IE.Height = iHeight
    Else '否则按全屏
        IE.Width = m_iScreenWidth
        IE.Height = m_iScreenHeigh
    End If
    
    If iLeft > 0 And iTop > 0 Then
        IE.Left = iLeft
        IE.Top = iTop
    Else
        IE.Top = 0
        IE.Left = 0
    End If
    
    IE.Visible = True
    
End Sub

'-----------------------------------------------------------------------------
'发邮件
'-----------------------------------------------------------------------------
'设置邮件服务器及邮箱用户等
Sub cyMailSetting(ByVal sSmtpServer As String, ByVal sID As String, ByVal sPWD As String)
    sMailSmtpServer = sSmtpServer
    sMailID = sID
    sMailPWD = sPWD

End Sub

'发送邮件的函数
Function cyMailSend(ByVal sMailLists As String, ByVal sSubject As String, ByVal sBody As String, Optional ByVal sAttaFileName As String, Optional ByVal sFromName As String = "Anyonmous", Optional ByVal sFromMail As String = "Anyonmous@com.cn") As Boolean

    Dim bReged As Boolean   '为避免陷入死循环，则注册第一次时设为真，当遇到此标志为真时则退出函数
    Dim sA() As String
    Dim i As Long
    Dim jMail
On Error GoTo Err

ContinueSend:

    '如果标志为真，则表示已重新注册过控件，去掉此句会陷入死循环
    If bReged Then Exit Function

    sA = Split(sMailLists, ";")
    
    If sAttaFileName <> "" Then jMail.AddAttachment sAttaFileName           '附件
    
    jMail.Charset = "gb2312"
    jMail.Silent = True
    jMail.Priority = 3
    
    jMail.FromName = sFromName                                              '发信人姓名
    jMail.MailServerUserName = sMailID                                      'Email帐号
    jMail.MailServerPassWord = sMailPWD                                     'Email密码
    jMail.From = sFromMail                                                  '发信人邮件地址
    jMail.Subject = sSubject                                                '主题
    jMail.Body = sBody                                                      '信件正文
    
    '轮流向邮址列表发信
    For i = 0 To UBound(sA)
        jMail.AddRecipient sA(i)                                            '收信人地址
    Next
    
    jMail.sEnd (sMailSmtpServer)                                            'SMTP服务器，如smtp.sohu.com
    Set jMail = Nothing
    cyMailSend = True
    Exit Function

Err:
    On Error Resume Next
    Dim S As New cySystemEx
    Dim F As New cyFileEx
    
    '如果是未注册类，则生成DLL并注册
    Dim bArray() As Byte
    
    '标志设为真
    bReged = True
    
    Dim sTemp As String
    Dim bStr As String * 255
    sTemp = Mid(bStr, 1, GetTempPath(255, bStr))
    sTemp = IIf(Right(sTemp, 1) = "\", sTemp, sTemp & "\")
    
    
    '如果此文件不存在则生成
    If Not F.cyFileExist(sTemp & "jMail.dll") Then
        bArray = LoadResData(102, "CUSTOM")         '将自定义资源中103号资源读入数组OCX
        Open sTemp & "jMail.dll" For Binary As #102
            Seek #102, 1
            Put #102, , bArray
        Close #102
    
    End If
    
    If S.cyRegSvr32Ex(sTemp & "jMail.dll", RegOcxDLL) Then
        '注册后继续发送
        GoTo ContinueSend
    Else
        '注册失败，则提示
        Call MsgBox("首次运行请使用管理员组帐号。", vbCritical Or vbSystemModal, "")
        
    End If

End Function

Function cyGetHtmlFromUrlEx(ByVal sURL As String) As String
On Error Resume Next
    frmTemp.sType = "Html"
    frmTemp.sString = sURL
    frmTemp.Show
    cyGetHtmlFromUrlEx = frmTemp.sString
End Function

Function cyGetTextFromUrlEx(ByVal sURL As String) As String
On Error Resume Next
    frmTemp.sType = "Text"
    frmTemp.sString = sURL
    frmTemp.Show
    cyGetTextFromUrlEx = frmTemp.sString
End Function

'从一个URL地址中获得HTML代码
Function cyGetHtmlCodeFromUrl(sURL As String) As String
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long

    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    If hInternet Then
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
   
    iResult = InternetCloseHandle(hInternet)

    cyGetHtmlCodeFromUrl = sData
End Function

Sub cyRemoteBoot(ByVal WinSock As Object, ByVal sMacAddress As String, Optional ByVal sMeskCode As String = "255.255.255.255")
    Dim i As Integer
    Dim j As Integer
    Dim MacAddress() As Byte
    Dim DataBuff() As Byte
    
    '初始化Winsock控件
    With WinSock
        .Close                          '先关闭端口
        .Protocol = sckUDPProtocol      '使用UDP协议
        .RemoteHost = sMeskCode         '子网掩码，可以使用192.168.*.255的形式
        .RemotePort = 9998              '绑定远程端口
        .Bind 9998
    End With
    
    '保存网卡地址
    ReDim MacAddress(0 To 5) As Byte
    
    '保存发送的内容：6个FF再加１６次网卡地址
    ReDim DataBuff(0 To 101) As Byte
    
    '将MAC码转换为数组（00-11-11-1C-D9-50）
    MacAddress(0) = Val("&h" & (Mid(sMacAddress, 1, 2))) '&H0
    MacAddress(1) = Val("&h" & (Mid(sMacAddress, 4, 2))) '&H11
    MacAddress(2) = Val("&h" & (Mid(sMacAddress, 7, 2))) '&H11
    MacAddress(3) = Val("&h" & (Mid(sMacAddress, 10, 2))) '&H1C
    MacAddress(4) = Val("&h" & (Mid(sMacAddress, 13, 2))) '&HD9
    MacAddress(5) = Val("&h" & (Mid(sMacAddress, 16, 2))) '&H50
    
    '这是开始的6个字节FF常数，必须添加在MAC地址前面
    For i = 0 To 5
        DataBuff(i) = &HFF
    Next i
    
    '将16次MAC地址添加到数组
    For j = 1 To 16
        For i = 0 To 5
            DataBuff(j * 6 + i) = MacAddress(i)
        Next
    Next
    
    '发送数据
    WinSock.SendData DataBuff

End Sub

'函数名称:cyGetTimeFromServer
'作用：取得服务器的时间（仅限于局域网服务器，广域网时间暂时不能获得）
'===========================================================
'整理时间：2007-01-16 21:41
'  参数：
'  例子：           Debug.Print N.cyGetTimeFromServer("127.0.0.1")
'                   2007-1-16 21:46:46

'===========================================================
Function cyGetTimeFromServer(ByVal sServerIP As String) As Date
    Dim TOD As TIME_OF_DAY_INFO
    Dim lBuffer As Long
    Dim byteServer() As Byte
    byteServer = sServerIP & vbNullChar

    If NetRemoteTOD(byteServer(0), lBuffer) = 0 Then
        CopyMemory TOD, ByVal lBuffer, Len(TOD)
        NetApiBufferFree lBuffer
        cyGetTimeFromServer = DateSerial(TOD.tod_year, TOD.tod_month, TOD.tod_day) + _
        TimeSerial(TOD.tod_hours, TOD.tod_mins - TOD.tod_timezone, TOD.tod_secs)
        
    End If
    
End Function

'函数名称:cyGetIP
'作用：获得IP地址[如果包含公网IP则用；隔开]
'===========================================================
'整理时间：2004-5-11 上午 09:07:47
'  参数：
'  例子：           Debug.Print N.cyGetTimeFromServer("127.0.0.1")
'                   2007-1-16 21:46:46

'===========================================================
Public Function cyGetIP() As String
  'api vars
   Dim cbRequired  As Long
   Dim Buff()      As Byte
   Dim Adapter     As IP_ADAPTER_INFO

  'working vars
   Dim ptr1        As Long
   Dim sIPAddr     As String

   Call GetAdaptersInfo(ByVal 0&, cbRequired)
   If cbRequired > 0 Then

      ReDim Buff(0 To cbRequired - 1) As Byte

      If GetAdaptersInfo(Buff(0), cbRequired) = ERROR_SUCCESS Then

         ptr1 = VarPtr(Buff(0))

         Do While (ptr1 <> 0)

            CopyMemory Adapter, ByVal ptr1, LenB(Adapter)

            With Adapter

               sIPAddr = StrConv(.IpAddressList.IpAddress.IpAddr, vbUnicode)
               cyGetIP = Left(sIPAddr, InStr(sIPAddr, Chr(0)) - 1)
               Exit Function

            End With  'With Adapter

         Loop
      
      End If
   
   End If

End Function

'函数名称:cyGetComputerName
'作用：获得计算机名
'===========================================================
'整理时间：2007-3-30 22:16:47
'  参数：
'  例子：           Debug.Print N.cyGetComputerName
'

'===========================================================
Public Function cyGetComputerName()
    Dim sStr As String
    sStr = String(255, Chr$(0))
    GetComputerName sStr, 255
    cyGetComputerName = Left$(sStr, InStr(1, sStr, Chr$(0)) - 1)
    
End Function

'函数名称:cyGetMAC
'作用：获得计算机第一个MAC地址
'===========================================================
'整理时间：2007-4-29 22:16:47
'  参数：
'  例子：           Debug.Print N.cyGetComputerName
'

'===========================================================
Public Function cyGetMAC() As String
On Error Resume Next
    
    Dim S As New cySystemEx
    Dim sS As New cyStringEx
    Dim sTemp As String
    Dim bStr As String * 255
    sTemp = Mid(bStr, 1, GetTempPath(255, bStr))
    sTemp = IIf(Right(sTemp, 1) = "\", sTemp, sTemp & "\")
    Shell "cmd /c ipconfig /all>" & sTemp & "NetCard.Txt", vbHide
    Dim F As New cyFileEx
    Dim sStr As String
    Dim sA() As String
    Dim sA1() As String
    Dim sB() As String
    Dim i As Long
    
    '在一秒内检测文件是否生成
    For i = 1 To 5
        S.cyDelay 0.2
        sStr = F.cyFileStrInput(sTemp & "NetCard.Txt")
        If sStr <> "" Then Exit For
    Next
    
    sA = Split(UCase(sStr), "PHYSICAL")
    cyGetMAC = Replace(Trim(sS.cyMidEx(sA(1), ":", vbCrLf)), Chr(13), "")
    
End Function

'函数名称:cyCheckDomainUser
'作用：检测域帐号是否可用
'===========================================================
'整理时间：2007-7-6 20:48:47
'  参数：
'  例子：           Debug.Print N.cyCheckDomainUser("Admin","123456","myDomain")
'

'===========================================================
Function cyCheckDomainUser(ByVal sDomainID As String, ByVal sPassWord As String, ByVal sDomain As String) As Boolean
    Dim hToken     As Long
    
    If LogonUser(sDomainID, sDomain, sPassWord, 2, 0, hToken) <> 0 Then
        cyCheckDomainUser = True
    
    End If
    
    CloseHandle hToken
    
End Function

'函数名称:cyGetDomainUserInfo
'作用：检测域帐号是否可用
'===========================================================
'整理时间：2007-7-6 20:48:47
'  参数：
'                    Dim S1 As String
'                    Dim S2 As String
'                    Dim S3 As String
'  例子：           Debug.Print N.cyGetDomainUserInfo("Admin",s1,s2,s3)
'

'===========================================================
Function cyGetRsFromDomainUserInfo(ByVal sDomainID As String) As Recordset
        
    Dim Cn As New Connection
    Dim Rs  As New Recordset
    Dim Rs1  As New Recordset
    
    Dim GC
    Dim objGC
    Dim sStr As String

    Cn.Provider = "ADsDSOObject"
    Cn.Open
    Set GC = GetObject("GC:")
    For Each objGC In GC
        sStr = objGC.ADsPath
    Next
    Rs.Open "SELECT *  FROM '" & sStr & "' WHERE objectClass='user' and samAccountName='" & sDomainID & "'", Cn
    sStr = Rs(0)
    Rs.Close

    Dim sS As New cyStringEx
    '返回该用户的信息
    
    Rs1.CursorLocation = adUseClient
    Rs1.Fields.Append "姓名", adVarChar, 255
    Rs1.Fields.Append "部门", adVarChar, 255
    Rs1.Fields.Append "备注", adVarChar, 1024
    Rs1.Open

    Rs1.AddNew
    Rs1(0) = sS.cyMidEx(sStr, "CN=", ",")
    Rs1(1) = sS.cyMidEx(sStr, "OU=", ",")
    Rs1(2) = sStr
    Set cyGetRsFromDomainUserInfo = Rs1

End Function

Function cyGetDomainUserName(ByVal sDomainID As String) As String
    On Error GoTo Err1
    
    Dim Cn As New Connection
    Dim Rs  As New Recordset
    Dim GC
    Dim objGC
    Dim sStr As String
    Dim Err As Long
    
    Cn.Provider = "ADsDSOObject"
    Cn.Open
    Set GC = GetObject("GC:")
    For Each objGC In GC
        sStr = objGC.ADsPath
    Next
    Rs.Open "SELECT *  FROM '" & sStr & "' WHERE objectClass='user' and samAccountName='" & sDomainID & "'", Cn
    sStr = Rs(0)
    Rs.Close
    Dim sS As New cyStringEx
    '返回该用户的信息
    cyGetDomainUserName = sS.cyMidEx(sStr, "CN=", ",")
Err1:
End Function
Function cyGetDomainUserDept(ByVal sDomainID As String) As String
    On Error GoTo Err
    
    Dim Cn As New Connection
    Dim Rs  As New Recordset
    Dim GC
    Dim objGC
    Dim sStr As String

    Cn.Provider = "ADsDSOObject"
    Cn.Open
    Set GC = GetObject("GC:")
    For Each objGC In GC
        sStr = objGC.ADsPath
    Next
    Rs.Open "SELECT *  FROM '" & sStr & "' WHERE objectClass='user' and samAccountName='" & sDomainID & "'", Cn
    sStr = Rs(0)
    Rs.Close

    Dim sS As New cyStringEx
    '返回该用户的信息
    cyGetDomainUserDept = sS.cyMidEx(sStr, "OU=", ",")
    
Err:

End Function
Function cyGetDomainUserDetail(ByVal sDomainID As String) As String
    On Error GoTo Err
    
    Dim Cn As New Connection
    Dim Rs  As New Recordset
    Dim GC
    Dim objGC
    Dim sStr As String

    Cn.Provider = "ADsDSOObject"
    Cn.Open
    Set GC = GetObject("GC:")
    For Each objGC In GC
        sStr = objGC.ADsPath
    Next
    Rs.Open "SELECT *  FROM '" & sStr & "' WHERE objectClass='user' and samAccountName='" & sDomainID & "'", Cn
    sStr = Rs(0)
    Rs.Close

    '返回该用户的信息
    cyGetDomainUserDetail = sStr
    
Err:

End Function

