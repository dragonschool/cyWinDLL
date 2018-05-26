Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

'-----------------------------------------------------------------------------
'返回进程列表
'-----------------------------------------------------------------------------
'Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

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
'根据句柄得到其执行文件全路径
'-----------------------------------------------------------------------------

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
'-----------------------------------------------------------------------------
'根据句柄得到其执行文件全路径
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
Private Declare Sub MoveMemory1 Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
'Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)


'-----------------------------------------------------------------------------
'获得文件版本，及对比版本新旧
'-----------------------------------------------------------------------------

Sub Main()

On Error Resume Next

    Dim sTemp As String
    Dim bStr As String * 255
    Dim bArray() As Byte

    Dim sSource As String
    Dim sTarget As String
    Dim sFileName As String
    Dim PID As Long
    Dim bDisplayUpdateConfirm As Boolean
    Dim bDisplayErr As Boolean
    
    '读出文件
    sSource = GetSetting("cyDLL", "UpdateFile", "Source", "")
    sTarget = GetSetting("cyDLL", "UpdateFile", "Target", "")
    sFileName = cyCutFileName(sTarget)
    '如果是调式模式则不处理
    If InStr(1, UCase(sTarget), "VB6.EXE") > 0 Then End
    
    '保存当前程序的PID
    PID = GetSetting("cyDLL", "UpdateFile", "Pid", 0)
    
    '保存是否弹出确认窗口
    bDisplayUpdateConfirm = GetSetting("cyDLL", "UpdateFile", "bDisplayUpdateConfirm", True)

    '保存是否显示错误
    bDisplayErr = GetSetting("cyDLL", "UpdateFile", "bDisplayErr", True)
    
        '有新版本
        If cyFileIsNewVersion(cyGetFileVersion(sTarget), cyGetFileVersion(sSource)) Then
        
            '是否显示确认窗口
            If bDisplayUpdateConfirm Then
                
                Select Case MsgBox("发现有新版本,是否现在更新 ?", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton1, "更新")
                
                    Case vbYes
'开始捕捉错误
On Error GoTo Err

                        Dim sA() As String
                        Dim i As Long
                        Dim hProcess As Long
                        Const PROCESS_TERMINATE = &H1
                        
                        sA = cyGetProcessToArray
                        For i = 2 To UBound(sA)
                            If UCase((Split(sA(i), "#")(1))) = UCase(sFileName) Then
                                hProcess = OpenProcess(PROCESS_TERMINATE, 0, Split(sA(i), vbTab)(0))
                                Call TerminateProcess(hProcess, 0)
                                CloseHandle hProcess
                                
                            End If
                        
                        Next

                        '将目标文件覆盖源文件
                        FileCopy sSource, sTarget
                    
                        '如果已经不存在版本问题，则表示更新已成功，则重启源程序
                        If Not cyFileIsNewVersion(cyGetFileVersion(sTarget), cyGetFileVersion(sSource)) Then
                            WinExec sTarget, 0
                        
                        Else
                            Call MsgBox("程序更新失败!", vbCritical Or vbSystemModal, "")
                        
                        End If
                        
                    Case vbNo
                    
                        '不更新则关闭程序
                        End
                        
                End Select
                
            Else
            '不提示直接更新
            
'开始捕捉错误
On Error GoTo Err
                
                '读出当前程序的PID并杀死
                TerminateProcess OpenProcess(&H1, 0, PID), 0
                DoEvents
                '将目标文件覆盖源文件
                FileCopy sSource, sTarget
                
            End If
        
        End If
        Exit Sub
        
Err:
    If bDisplayErr Then

        '提示
        Call MsgBox(Err.Description & ",自动更新失败!", vbCritical Or vbSystemModal, "更新")
        
    End If

End Sub

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

End Function

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

Public Function cyCutFileName(ByVal sFullFileName As String) As String
    Dim sArray() As String
    If sFullFileName = "" Then Exit Function
    
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

End Function

