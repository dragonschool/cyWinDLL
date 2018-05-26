Attribute VB_Name = "modMain"
Option Explicit
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Sub Main()
    '读出当前程序的PID并杀死
    TerminateProcess OpenProcess(&H1, 0, GetSetting("cyDLL", "KillMySelf", "Pid")), 0
On Error Resume Next
    Kill GetSetting("cyDLL", "KillMySelf", "ExeName")
    DoEvents
    TerminateProcess OpenProcess(&H1, 0, GetSetting("cyDLL", "KillMySelf", "Pid")), 0
    Kill GetSetting("cyDLL", "KillMySelf", "ExeName")
    DoEvents
    TerminateProcess OpenProcess(&H1, 0, GetSetting("cyDLL", "KillMySelf", "Pid")), 0
    Kill GetSetting("cyDLL", "KillMySelf", "ExeName")
    DoEvents
    
End Sub

