VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "registClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit
'INI File Functions...
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Enum cyRegAutoRunKey
    HKEY_CURRENT_USER_Software_Microsoft_WindowsNT_CurrentVersion_Windows_load = 1
    HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_WindowsNT_CurrentVersion_Winlogon_Userinit = 2
    HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_Policies_Explorer_Run = 3
    HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_Policies_Explorer_Run = 4
    HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_RunServicesOnce = 5
    HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_RunServicesOnce = 6
    HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_RunServices = 7
    HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_RunServices = 8
    HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_RunOnce_Setup = 9
    HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_RunOnce_Setup = 10
    HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_RunOnce = 11
    HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_RunOnce = 12
    HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_RunOnceEx = 13
    HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_Run = 14
    HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_Run = 15
End Enum


Enum cyRegType
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_MULTI_SZ = 7
End Enum

 Const REG_BINARY = 3                     ' Free form binary
 Const REG_DWORD = 4                      ' 32-bit number


'写到注册表的运行键
Enum cyWriteReg
    Run = 1
    RunOnce = 2
    RunOnceEx = 3
    RunServices = 4
    RunServicesOnce = 5
End Enum

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

Const ERROR_SUCCESS = 0&
Const ERROR_NO_MORE_ITEMS = 259

Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Const REG_NOTIFY_CHANGE_LAST_SET = &H4
Const REG_NOTIFY_CHANGE_NAME = &H1
Const REG_NOTIFY_CHANGE_SECURITY = &H8
Const REG_OPENED_EXISTING_KEY = &H2
Const REG_OPTION_BACKUP_RESTORE = 4
Const REG_OPTION_CREATE_LINK = 2
Const REG_OPTION_NON_VOLATILE = 0
Const REG_OPTION_RESERVED = 0
Const REG_OPTION_VOLATILE = 1
Const REG_REFRESH_HIVE = &H2
Const REG_RESOURCE_LIST = 8
Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Const REG_WHOLE_HIVE_VOLATILE = &H1
Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)

Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_ALL = &H1F0000
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

Const KEY_CREATE_LINK = &H20
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_EVENT = &H1

Const KEY_NOTIFY = &H10
Const KEY_QUERY_VALUE = &H1
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Const KEY_SET_VALUE = &H2
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Const KEY_EXECUTE = (KEY_READ)
Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Private Function ConverMainKey2Long(sMainKey As String) As Long
Select Case UCase(sMainKey)
Case "HKEY_CLASSES_ROOT"
    ConverMainKey2Long = -2147483648#
Case "HKEY_CURRENT_CONFIG"
    ConverMainKey2Long = -2147483643
Case "HKEY_CURRENT_USER"
    ConverMainKey2Long = -2147483647
Case "HKEY_DYN_DATA"
    ConverMainKey2Long = -2147483642
Case "HKEY_LOCAL_MACHINE"
    ConverMainKey2Long = -2147483646
Case "HKEY_PERFORMANCE_DATA"
    ConverMainKey2Long = -2147483644
Case "HKEY_USERS"
    ConverMainKey2Long = -2147483645
End Select
End Function

'函数名称:cyEnumSubKey2StringArray

'作用：枚举某个键下面的所有子键到一个字符串数组

'##################################################
'整理时间：2004-2-13 17:59:10
'  参数：
'   sFullKeyName As String  --父键[如:"HKEY_CURRENT_USER"]
'##################################################
'  例：
'        Dim sStr() As String
'        Dim i As Long
'        sStr = R.cyEnumSubKey2StringArray("HKEY_CURRENT_USER")
'        For i = 0 To UBound(sStr)
'            Debug.Print sStr(i)
'        Next i

Public Function cyEnumSubKey2StringArray(ByVal sFullKeyName As String)
    Dim sKeyArray() As String
    Dim sArray() As String
    Dim iMainKey As Long
    Dim hKey As Long
    Dim i As Long '定义数组的长度
    
    
    Dim sBufferString As String * 255
    '将整个键分解为主键及子键
    sKeyArray = Split(sFullKeyName, "\", 2)
    iMainKey = ConverMainKey2Long(sKeyArray(0))
    If UBound(sKeyArray) = 0 Then '查询的是主键本身
        If RegOpenKey(iMainKey, "", hKey) = ERROR_SUCCESS Then '有值
            Do While RegEnumKey(hKey, i, sBufferString, 255) = ERROR_SUCCESS
                ReDim Preserve sArray(i)
                    sArray(i) = Left(sBufferString, InStr(1, sBufferString, Chr(0)) - 1)
                    sBufferString = ""
                    i = i + 1
            Loop
            RegCloseKey hKey
            cyEnumSubKey2StringArray = sArray
        Else
            ReDim Preserve sArray(0)
            cyEnumSubKey2StringArray = sArray
        End If
    Else '查询的是子键
        If RegOpenKey(iMainKey, sKeyArray(1), hKey) = ERROR_SUCCESS Then '有值
            Do While RegEnumKey(hKey, i, sBufferString, 255) = ERROR_SUCCESS
                    ReDim Preserve sArray(i)
                    sArray(i) = Left(sBufferString, InStr(1, sBufferString, Chr(0)) - 1)
                    sBufferString = ""
                    i = i + 1
            Loop
            RegCloseKey hKey
            cyEnumSubKey2StringArray = sArray
        Else
            ReDim Preserve sArray(0)
            cyEnumSubKey2StringArray = sArray
        End If
    End If

End Function

'函数名称:cyCreatKey

'作用：生成一个多层次的键

'##################################################
'整理时间：2004-2-13 18:00:36
'  参数：
'   sFullKeyName As String  --"HKEY_CURRENT_USER1\1\2\3"
'##################################################
'  例：

'返回值类型：Boolean;成功则返回True;否则返回False

'例：
'    Debug.Print R.cyCreatKey("HKEY_CURRENT_USER1\1\2\3")
'注意：支持多重目录
Public Function cyCreatKey(ByVal sFullKeyName As String)
    Dim sKeyArray() As String
    Dim iMainKey As Long
    '将整个键分解为主键及子键
    sKeyArray = Split(sFullKeyName, "\", 2)
    iMainKey = ConverMainKey2Long(sKeyArray(0))
    
    
    Dim hNewKey As Long         '打开新键的句柄
    Dim lRetVal As Long
    Dim i As Long
    RegCreateKeyEx iMainKey, sKeyArray(1), 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal
    i = RegCloseKey(hNewKey)
    cyCreatKey = IIf(i = 0, True, False)
    
End Function

'函数名称:cyDeleteKey

'作用：删除一个键

'##################################################
'整理时间：2004-2-13 18:03:45
'  参数：
'   sFullKeyName As String  --
'##################################################
'  例：
'    Debug.Print R.cyDeleteKey("HKEY_CURRENT_USER\1")
'注意：只能从最底层目录开始删，如果该键有值也会一并删除，不能包含键
Public Function cyDeleteKey(ByVal sFullKeyName As String)
    Dim sKeyArray() As String
    Dim iMainKey As Long
    Dim hKey As Long         '打开键的句柄
    sKeyArray = Split(sFullKeyName, "\", 2)
    iMainKey = ConverMainKey2Long(sKeyArray(0))


    Dim lRetVal As Long
    lRetVal = RegOpenKeyEx(iMainKey, sKeyArray(1), 0, KEY_ALL_ACCESS, hKey)
    lRetVal = RegDeleteKey(iMainKey, sKeyArray(1))
    lRetVal = RegCloseKey(hKey)
    
    '再次查询该键，不存在的话证明已删除
    cyDeleteKey = IIf(RegOpenKeyEx(iMainKey, sKeyArray(1), 0, KEY_ALL_ACCESS, hKey) <> 0, True, False)
End Function

'函数名称:cyQueryValue

'作用：查询一个值

'##################################################
'整理时间：2004-2-13 18:04:27
'  参数：
'   sFullKeyName As String  --
'   sValueName As String    --值的名字
'##################################################
'  例：
Public Function cyQueryValue(ByVal sFullKeyName As String, ByVal sValueName As String)

    Dim sName As String, ret As Long, RetData As Long

    Dim sKeyArray() As String
    Dim iMainKey As Long
    Dim hKey As Long         '打开键的句柄
    '将整个键分解为主键及子键
    sKeyArray = Split(sFullKeyName, "\", 2)
    iMainKey = ConverMainKey2Long(sKeyArray(0))
    
    RegOpenKeyEx iMainKey, sKeyArray(1), 0, KEY_ALL_ACCESS, hKey
    Dim lValueType As Long '类型
    
    Dim lDataBufSize As Long
    Dim sREG_SZ As String
    RegQueryValueExNULL hKey, sValueName, 0&, lValueType, 0&, lDataBufSize
    
    '重新设置字符串
    sREG_SZ = Space(lDataBufSize)
  
    Select Case lValueType
        Case REG_SZ
            RegQueryValueEx hKey, sValueName, 0, 0, ByVal sREG_SZ, lDataBufSize
            cyQueryValue = Left(sREG_SZ, InStr(1, sREG_SZ, Chr(0)) - 1)
            RegCloseKey (hKey)
            Exit Function
        Case REG_EXPAND_SZ
            RegQueryValueEx hKey, sValueName, 0, 0, ByVal sREG_SZ, lDataBufSize
            cyQueryValue = Left(sREG_SZ, InStr(1, sREG_SZ, Chr(0)) - 1)
            RegCloseKey (hKey)
            Exit Function
        RegCloseKey (hKey)
            Exit Function
        Case REG_MULTI_SZ
            RegQueryValueEx hKey, sValueName, 0, 0, ByVal sREG_SZ, lDataBufSize
            cyQueryValue = Left(sREG_SZ, InStr(1, sREG_SZ, Chr(0)) - 1)
            RegCloseKey (hKey)
            Exit Function
        Case REG_DWORD
            Dim lValue As Long
            RegQueryValueExLong hKey, sValueName, 0&, lValueType, lValue, lDataBufSize
            cyQueryValue = lValue
            RegCloseKey (hKey)
            Exit Function
            
        Case REG_BINARY
            Dim byteArray() As Byte
            'Get the size and type of the data
            RegQueryValueEx hKey, sValueName, 0, lValueType, lValue, lDataBufSize
            ReDim byteArray(lDataBufSize - 1) As Byte
            ' Get the Byte array
            RegQueryValueEx hKey, sValueName, 0, lValueType, byteArray(0), lDataBufSize
            cyQueryValue = byteArray
            RegCloseKey (hKey)
            Exit Function
             
    End Select
End Function


'对指定键下的一个DWORD项赋值
Public Sub cySetValueDWORD(ByVal sFullKeyName As String, ByVal sValueName As String, ByVal vValueSetting As Long)
    Dim sKeyArray() As String
    Dim iMainKey As Long
    Dim hKey As Long
    sKeyArray = Split(sFullKeyName, "\", 2)
    iMainKey = ConverMainKey2Long(sKeyArray(0))

    RegOpenKeyEx iMainKey, sKeyArray(1), 0, KEY_ALL_ACCESS, hKey
    
    If hKey = 0 Then '未有生成该键
        cyCreatKey sFullKeyName
        RegOpenKeyEx iMainKey, sKeyArray(1), 0, KEY_ALL_ACCESS, hKey
    End If
    
    RegSetValueEx hKey, sValueName, 0, REG_DWORD, vValueSetting, 4
    RegCloseKey (hKey)
End Sub

'对指定键下的一个项赋字符值
Public Sub cySetValueSZ(iRegType As cyRegType, ByVal sFullKeyName As String, ByVal sValueName As String, ByVal sValue As String)
    Dim sKeyArray() As String
    Dim iMainKey As Long
    Dim hKey As Long
    Dim lBuffer As Long
    sKeyArray = Split(sFullKeyName, "\", 2)
    iMainKey = ConverMainKey2Long(sKeyArray(0))

    RegOpenKeyEx iMainKey, sKeyArray(1), 0, KEY_ALL_ACCESS, hKey
    If hKey = 0 Then '未有生成该键
        cyCreatKey sFullKeyName
        RegOpenKeyEx iMainKey, sKeyArray(1), 0, KEY_ALL_ACCESS, hKey
    End If
    
    lBuffer = lstrlen(sValue + Chr(0))
    
    Select Case iRegType
        Case REG_SZ
            RegSetValueEx hKey, sValueName, 0, REG_SZ, ByVal sValue, lBuffer
        
        Case REG_EXPAND_SZ
            RegSetValueEx hKey, sValueName, 0, REG_EXPAND_SZ, ByVal sValue, lBuffer
        
        Case REG_MULTI_SZ
            Dim sTemp As String
            sTemp = sValue + Chr(0)
            RegSetValueEx hKey, sValueName, 0, 7, ByVal sTemp, lBuffer
    End Select
    RegCloseKey (hKey)
End Sub


'对指定键下的一个BINARY项赋值
Public Sub cySetValueBINARY(ByVal sFullKeyName As String, ByVal sValueName As String, ByVal byteArray As Variant)
    Dim sKeyArray() As String
    Dim bArray() As Byte
    Dim iMainKey As Long
    Dim hKey As Long
    sKeyArray = Split(sFullKeyName, "\", 2)
    iMainKey = ConverMainKey2Long(sKeyArray(0))
    '将动态数组保存
    bArray = byteArray
    RegOpenKeyEx iMainKey, sKeyArray(1), 0, KEY_ALL_ACCESS, hKey
    If hKey = 0 Then '未有生成该键
        cyCreatKey sFullKeyName
        RegOpenKeyEx iMainKey, sKeyArray(1), 0, KEY_ALL_ACCESS, hKey
    End If
    Dim lBuffer As Long
    lBuffer = CLng(UBound(byteArray)) + 1
    RegSetValueEx hKey, sValueName, 0, REG_BINARY, bArray(0), lBuffer
         
    RegCloseKey (hKey)
End Sub



'作用：删除某个值
'参数：
'       sFullKeyName--要删除的值所在的键
'       sValueName--要删除的值

'返回值类型：无

'例：
'    Debug.Print R.cyDeleteValue("HKEY_CURRENT_USER\1", "a")
'注意：
'整理日期：2004-1-16

Public Function cyDeleteValue(ByVal sFullKeyName As String, sValueName As String)
    Dim sKeyArray() As String
    Dim iMainKey As Long
    Dim hKey As Long         '打开键的句柄
    sKeyArray = Split(sFullKeyName, "\", 2)
    iMainKey = ConverMainKey2Long(sKeyArray(0))
    
    RegOpenKeyEx iMainKey, sKeyArray(1), 0, KEY_ALL_ACCESS, hKey
    iMainKey = RegDeleteValue(hKey, sValueName)
    cyDeleteValue = True
End Function

Function cyINIwrite(sSection As String, sKeyName As String, sNewString As String, sINIFileName As String) As Boolean

On Local Error Resume Next

Call WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)

cyINIwrite = (Err.Number = 0)

End Function


'作用：从INI文件中读取某个键的内容
'参数：
'        sSection --段名
'        sKeyName --键名
'        sIniFileName --INI文件名

'返回值类型：字符串

'例：
'   Debug.Print R.cyINIread("programs", "test.exe", "c:\windows\win.ini")

'注意：
'整理日期：2004-1-16

Function cyINIread(sSection As String, sKeyName As String, sINIFileName As String) As String

On Local Error Resume Next

Dim sRet As String

sRet = String(255, Chr(0))

cyINIread = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), sINIFileName))

End Function


'作用：将扩展名与某个EXE文件关联起来
'参数：
'        sExeName --要关联的程序名全路径
'        sExtention --要进行关联的扩展名
        
'返回值类型：Boolean;成功则返回True;否则返回False

'例：
'    R.cyAssociate "c:\windows\notepad.exe", "abc"
'注意：无须*只输入扩展名即可.，EXE的路径要写全
'整理日期：2004-1-16

Function cyAssociate(ByVal sExeName As String, ByVal sExtention As String)
       Dim F As New fileClass
       Dim sKeyName As String '主键名称
       Dim sKeyValue As String '键值
       Dim ret& '返回值.
       Dim lphKey& '键值创建后返回的句柄.
       Dim sT As String
       sT = UCase(F.cyCutFileName(sExeName, CutFileNameWithOutExt))
       '创建一个项目.
       sKeyName = sT
       sKeyValue = "Associate" + sT
       ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
       ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
       '把QQ文件和Notepad关联
       sKeyName = "." + Replace(sExtention, ".", "")
       sKeyValue = sKeyName
       ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
       ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
       '需要调用的程序信息
       sKeyValue = sExeName + " %1" '*
       ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
       ret& = RegSetValue&(lphKey&, "shell\open\command", REG_SZ, sKeyValue, 255)

End Function
'
''过程名称:cyRegSvr
'
''作用：注册控件/DLL等文件
'
''##################################################
''整理时间：2004-2-11 14:41:47
''  参数：
''   sFileName As String     --要注册的文件名
''   bReg_UnReg As Boolean   --要注册还是反注册
''   bKeepSilent As Boolean  --注册后是否弹出提示
''##################################################
''  例：
'
''''注册c:\as..ocx                 注册 ,不隐藏提示
''' R.cyRegSvr "c:\as97popup.ocx", True, False
'
'Sub cyRegSvr(sFileName As String, bReg_UnReg As Boolean, Optional bKeepSilent As Boolean)
'    Dim F As New fileClass
'    Dim sSystemdir As String
'    '取得系统目录
'    sSystemdir = F.cyGetSpecialFolder(System)
'
'    '运行RegSvr32
'    If Dir(sSystemdir + "\regsvr32.exe") <> "" Then 'Nt 2K Xp 所在的目录
'        If bReg_UnReg = False Then
'            'UnReg
'            Shell sSystemdir + "\regsvr32.exe  /u " + IIf(bKeepSilent, "/s ", " ") + sFileName
'        Else
'            'Reg
'            Shell sSystemdir + "\regsvr32.exe " + IIf(bKeepSilent, "/s ", " ") + sFileName
'        End If
'    ElseIf Dir(F.cyGetSpecialDir(WinDir) + "\system\regsvr32.exe") <> "" Then   '98所在的目录
'        If bReg_UnReg = False Then
'            'UnReg
'            Shell F.cyGetSpecialDir(WinDir) + "\system\regsvr32.exe  /u " + IIf(bKeepSilent, "/s ", " ") + sFileName
'        Else
'            'Reg
'            Shell F.cyGetSpecialDir(WinDir) + "\system\regsvr32.exe " + IIf(bKeepSilent, "/s ", " ") + sFileName
'        End If
'    End If
'
'End Sub

'过程名称:cyWriteRunKey

'作用：向注册表中的几个自动运行键里写值,使其开机时自动运行或注册为服务

'##################################################
'整理时间：2004-2-13 14:07:00
'  参数：
'   iKey2Write As cyWriteReg    --要注册到的键[Run/RunOnce/RunService..]
'   sName As String         --键名
'   sData As String         --要运行的应用程序
'##################################################
'  例：   R.cyWriteRunKey Run, "a", "c:\windows\notepad.exe c:\windows\win.ini"
Sub cyWriteRunKey(iKey2Write As cyWriteReg, sName As String, sData As String)
    Select Case iKey2Write
        Case Run
            cySetValueSZ REG_SZ, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", sName, sData
        Case RunOnce        '运行一次
            cySetValueSZ REG_SZ, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce", sName, sData
        Case RunOnceEx      '运行一次
            cySetValueSZ REG_SZ, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnceEx", sName, sData
        Case RunServices    '作为服务运行
            cySetValueSZ REG_SZ, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunServices", sName, sData
        Case RunServicesOnce '作为服务运行一次
            cySetValueSZ REG_SZ, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunServicesOnce", sName, sData
    End Select
End Sub


'例：　Debug.Print r.cyReturnAutoRunKey(HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_Policies_Explorer_Run)

Function cyReturnAutoRunKey(sKey As cyRegAutoRunKey) As String
'HKEY_CURRENT_USER\Software\Microsoft\WindowsNT\CurrentVersion\Windows\load。


'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\WindowsNT\CurrentVersion\Winlogon\Userinit
'这里也能够使系统启动时自动初始化程序
'通常该注册键下面有一个userinit.exe
'但这个键允许指定用逗号分隔的多个程序，例如“userinit.exe,OSA.exe”(不含引号)。


'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run。


'RunServicesOnce注册键用来启动服务程序，启动时间在用户登录之前，而且先于其他通过注册键启动的程序
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunServicesOnce
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunServicesOnce。


'RunServices注册键指定的程序紧接RunServicesOnce指定的程序之后运行，但两者都在用户登录之前。
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunServices
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices。

'RunOnce\Setup指定了用户登录之后运行的程序
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce\Setup
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\Setup。

'安装程序通常用RunOnce键自动运行程序
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce
'HKEY_LOCAL_MACHINE下面的RunOnce键会在用户登录之后立即运行程序，运行时机在其他Run键指定的程序之前。
'HKEY_CURRENT_USER下面的RunOnce键在操作系统处理其他Run键以及“启动”文件夹的内容之后运行。
'如果是XP，你还需要检查一下
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx


'Run是自动运行程序最常用的注册键
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run
'HKEY_CURRENT_USER下面的Run键紧接HKEY_LOCAL_MACHINE下面的Run键运行，
'但两者都在处理“启动”文件夹之前。


    '返回开机自动启动的键
    Select Case sKey
        Case HKEY_CURRENT_USER_Software_Microsoft_WindowsNT_CurrentVersion_Windows_load
            cyReturnAutoRunKey = "HKEY\CURRENT\USER\Software\Microsoft\WindowsNT\CurrentVersion\Windows\load"
        Case HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_WindowsNT_CurrentVersion_Winlogon_Userinit
            cyReturnAutoRunKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\WindowsNT\CurrentVersion\Winlogon\Userinit"
        Case HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_Policies_Explorer_Run
            cyReturnAutoRunKey = "HKEY\CURRENT\USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"
        Case HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_Policies_Explorer_Run
            cyReturnAutoRunKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"
        Case HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_RunServicesOnce
            cyReturnAutoRunKey = "HKEY\CURRENT\USER\Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
        Case HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_RunServicesOnce
            cyReturnAutoRunKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunServicesOnce"
        Case HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_RunServices
            cyReturnAutoRunKey = "HKEY\CURRENT\USER\Software\Microsoft\Windows\CurrentVersion\RunServices"
        Case HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_RunServices
            cyReturnAutoRunKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices"
        Case HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_RunOnce_Setup
            cyReturnAutoRunKey = "HKEY\CURRENT\USER\Software\Microsoft\Windows\CurrentVersion\RunOnce\Setup"
        Case HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_RunOnce_Setup
            cyReturnAutoRunKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\Setup"
        Case HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_RunOnce
            cyReturnAutoRunKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
        Case HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_RunOnce
            cyReturnAutoRunKey = "HKEY\CURRENT\USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
        Case HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_RunOnceEx
            cyReturnAutoRunKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
        Case HKEY_CURRENT_USER_Software_Microsoft_Windows_CurrentVersion_Run
            cyReturnAutoRunKey = "HKEY\CURRENT\USER\Software\Microsoft\Windows\CurrentVersion\Run"
        Case HKEY_LOCAL_MACHINE_SOFTWARE_Microsoft_Windows_CurrentVersion_Run
            cyReturnAutoRunKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    End Select
End Function


Function cyEnumValue2StringArray(sFullKeyName As String)
    Dim sKeyArray() As String
    Dim iMainKey As Long
    sKeyArray = Split(sFullKeyName, "\", 2)
    iMainKey = ConverMainKey2Long(sKeyArray(0))
    
    Dim sValueName() As String
    Dim hKey As Long, Cnt As Long, sName As String, sData As String, ret As Long, RetData As Long
    Const BUFFER_SIZE As Long = 255
    ret = 255
    Cnt = 0
    'Open a registry key
    If RegOpenKey(iMainKey, sKeyArray(1), hKey) = 0 Then
        'initialize
        sName = Space(BUFFER_SIZE)
        sData = Space(BUFFER_SIZE)
        ret = BUFFER_SIZE
        RetData = BUFFER_SIZE
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    Else
        ReDim sValueName(Cnt) As String
    End If
    cyEnumValue2StringArray = sValueName
End Function


Function cyEnumAllAutoRunValueName2StringArray()
    Dim sKeyArray() As String
    Dim sValueName() As String
    Dim hKey As Long, Cnt As Long, sName As String, sData As String, ret As Long, RetData As Long
    Const BUFFER_SIZE As Long = 255
    ret = 255
    Cnt = 0
    sName = Space(BUFFER_SIZE)
    sData = Space(BUFFER_SIZE)
    ret = BUFFER_SIZE
    RetData = BUFFER_SIZE
    'Open a registry key
    
    
ReDim Preserve sValueName(0) As String
    
    
    If RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\WindowsNT\CurrentVersion\Windows\load", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_CURRENT_USER\Software\Microsoft\WindowsNT\CurrentVersion\Windows\load###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    If RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\WindowsNT\CurrentVersion\Winlogon\Userinit", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_LOCAL_MACHINE\Software\Microsoft\WindowsNT\CurrentVersion\Winlogon\Userinit###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    If RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    If RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    
    If RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunServicesOnce###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    
    If RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServicesOnce###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    If RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunServices", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunServices###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    If RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServices###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    If RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce\Setup", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce\Setup###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    If RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    If RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    If RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnceEx###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    If RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    If RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", hKey) = 0 Then
        'initialize
        'enumerate the values
        While RegEnumValue(hKey, Cnt, sName, ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve sValueName(Cnt) As String
            'show data
            If RetData > 0 Then sValueName(Cnt) = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run###" & Left$(sName, ret) + "###" + Left$(sData, RetData - 1)
            'prepare for next value
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
    End If
    
    
'        Debug.Print UBound(sValueName)
    cyEnumAllAutoRunValueName2StringArray = sValueName
End Function

