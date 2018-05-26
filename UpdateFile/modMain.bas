Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

'-----------------------------------------------------------------------------
'���ؽ����б�
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
'���ؽ����б�
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'���ݾ���õ���ִ���ļ�ȫ·��
'-----------------------------------------------------------------------------

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
'-----------------------------------------------------------------------------
'���ݾ���õ���ִ���ļ�ȫ·��
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'����ļ��汾�����ԱȰ汾�¾�
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

'ר���ڻ���ļ���ţ���������MOVEMEMORY������ײ����˽��˺�����1����Ӧ
Private Declare Sub MoveMemory1 Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
'Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)


'-----------------------------------------------------------------------------
'����ļ��汾�����ԱȰ汾�¾�
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
    
    '�����ļ�
    sSource = GetSetting("cyDLL", "UpdateFile", "Source", "")
    sTarget = GetSetting("cyDLL", "UpdateFile", "Target", "")
    sFileName = cyCutFileName(sTarget)
    '����ǵ�ʽģʽ�򲻴���
    If InStr(1, UCase(sTarget), "VB6.EXE") > 0 Then End
    
    '���浱ǰ�����PID
    PID = GetSetting("cyDLL", "UpdateFile", "Pid", 0)
    
    '�����Ƿ񵯳�ȷ�ϴ���
    bDisplayUpdateConfirm = GetSetting("cyDLL", "UpdateFile", "bDisplayUpdateConfirm", True)

    '�����Ƿ���ʾ����
    bDisplayErr = GetSetting("cyDLL", "UpdateFile", "bDisplayErr", True)
    
        '���°汾
        If cyFileIsNewVersion(cyGetFileVersion(sTarget), cyGetFileVersion(sSource)) Then
        
            '�Ƿ���ʾȷ�ϴ���
            If bDisplayUpdateConfirm Then
                
                Select Case MsgBox("�������°汾,�Ƿ����ڸ��� ?", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton1, "����")
                
                    Case vbYes
'��ʼ��׽����
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

                        '��Ŀ���ļ�����Դ�ļ�
                        FileCopy sSource, sTarget
                    
                        '����Ѿ������ڰ汾���⣬���ʾ�����ѳɹ���������Դ����
                        If Not cyFileIsNewVersion(cyGetFileVersion(sTarget), cyGetFileVersion(sSource)) Then
                            WinExec sTarget, 0
                        
                        Else
                            Call MsgBox("�������ʧ��!", vbCritical Or vbSystemModal, "")
                        
                        End If
                        
                    Case vbNo
                    
                        '��������رճ���
                        End
                        
                End Select
                
            Else
            '����ʾֱ�Ӹ���
            
'��ʼ��׽����
On Error GoTo Err
                
                '������ǰ�����PID��ɱ��
                TerminateProcess OpenProcess(&H1, 0, PID), 0
                DoEvents
                '��Ŀ���ļ�����Դ�ļ�
                FileCopy sSource, sTarget
                
            End If
        
        End If
        Exit Sub
        
Err:
    If bDisplayErr Then

        '��ʾ
        Call MsgBox(Err.Description & ",�Զ�����ʧ��!", vbCritical Or vbSystemModal, "����")
        
    End If

End Sub

'�����ļ��汾�ż���ļ��Ƿ����
Function cyFileIsNewVersion(ByVal sOldFileVersion As String, ByVal sNewFileVersion As String) As Boolean
    Dim sA1() As String
    Dim sA2() As String
    
    '�ļ��汾����ͬ
    If (sOldFileVersion) = (sNewFileVersion) Then Exit Function
    '�ļ��汾�Ų�����
    If Len(sOldFileVersion) = 0 Or Len(sNewFileVersion) = 0 Then Exit Function
    sA1 = Split(sOldFileVersion, ".")
    sA2 = Split(sNewFileVersion, ".")
    
    If UBound(sA1) <> UBound(sA2) Then Exit Function
    
    If CLng(sA1(0)) < CLng(sA2(0)) Then
        cyFileIsNewVersion = True
    
    ElseIf CLng(sA1(0)) > CLng(sA2(0)) Then
    
    Else
    
        If CLng(sA1(1)) < CLng(sA2(1)) Then
        '��һ�����ߴ�
            cyFileIsNewVersion = True
        
    
        ElseIf CLng(sA1(1)) > CLng(sA2(1)) Then
        
        Else
            If CLng(sA1(2)) < CLng(sA2(2)) Then
            '��һ�����ߴ�
                cyFileIsNewVersion = True
            
            ElseIf CLng(sA1(2)) > CLng(sA2(2)) Then
            Else
                If CLng(sA1(3)) < CLng(sA2(3)) Then
                '��һ�����ߴ�
                    cyFileIsNewVersion = True
                    
                End If
            
            End If
            
        End If
        
    End If
    
End Function

'�����ļ��汾��
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
    
    '����Ϊ��1��������������
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
        If (Process32First(j, Process)) Then '������һ������
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
    If UBound(sArray) > 0 Then '�ж��·�� \.\.\,,,
        cyCutFileName = sArray(UBound(sArray))
    Else 'ֻ��һ��   c:\..   aaa.exe
        If InStr(1, sArray(0), ":") > 0 Then '��:
            sArray = Split(sFullFileName, ":")
            cyCutFileName = sArray(UBound(sArray))
        Else 'û��,ֻ���ļ���
            cyCutFileName = sFullFileName
        End If
    End If

End Function

