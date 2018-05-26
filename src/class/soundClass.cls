VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "soundClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''Private Sub Command1_Click()
''''当出错时弹出错误显示
'''If sSound.cyMP3Play("L:\DBXY\aa\不存在的歌曲") = False Then MsgBox sSound.cyMP3Err
'''
'''sSound.MP3Pause
'''
''''跳到某一个地方开始播放
'''sSound.MP3Seek 11000
'''
'''
''''播放一个WAV音乐文件
'''sSound.cyPlayWave "C:\WINDOWS\Media\tada.wav"
'''
''''返回已经播放的长度[单位:毫秒]
'''Debug.Print "现在正播放到:" & sSound.MP3PlayPos & "毫秒"
'''
''''停止当前的MP3播放
'''sSound.MP3Stop
'''
''''返回当前歌曲的长度[单位:毫秒]
'''Debug.Print "此歌的长度:" & sSound.MP3Length & "毫秒"
'''
''''暂停播放MP3
'''sSound.MP3Pause
'''
''''现在播放的状态[playing/stopped/paused等]
'''Debug.Print sSound.MP3Status
'''
''''返回出错的描述
'''Debug.Print sSound.cyMP3Err
'''
''''以时间的形式返回正在播放的时间
'''Debug.Print sSound.MP3PosFormat2Time
'''End Sub
'''

Option Explicit
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private mmOpen As String, sec As Integer, mins As Integer
Public MP3Err As String
Enum EnumSoundFlags
 SND_APPLICATION = &H80
 SND_ALIAS = &H10000
 SND_ALIAS_ID = &H110000
 SND_ASYNC = &H1
 SND_FILENAME = &H20000
 SND_LOOP = &H8
 SND_MEMORY = &H4
 SND_NODEFAULT = &H2
 SND_NOSTOP = &H10
 SND_NOWAIT = &H2000
 SND_PURGE = &H40
 SND_RESOURCE = &H40004
 SND_SYNC = &H0
End Enum
Public MP3isPlaying As Boolean
Public MP3isPause As Boolean
Public Function MP3Status() As String
    Static S As String * 30
    mciSendString "status MP3Play mode", S, Len(S), 0
    MP3Status = S
End Function

Public Function cyMP3Err() As String
    cyMP3Err = MP3Err
End Function

Public Function cyMP3Play(ByVal sFileName As String) As Boolean
Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128

If MP3isPause = True Then '是继续播放
    mciSendString "play MP3Play", 0, 0, 0
    cyMP3Play = True
    MP3isPlaying = True
    MP3isPause = False
    Exit Function
End If

Dim tmp As String * 255
Dim lenShort As Long
Dim ShortPathAndFie As String, glo_HWND As Long
    If InStr(1, sFileName, " ") > 0 Then '有空格不行
        lenShort = GetShortPathName(sFileName, tmp, 255)
        sFileName = Mid(tmp, 1, InStr(1, tmp, Chr(0)) - 1)
    End If
    
    glo_HWND = 0
    cmdToDo = "open " & sFileName & " type MPEGVideo Alias MP3Play"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn = 0 Then    '成功
        mciSendString "play MP3Play", 0, 0, 0
        cyMP3Play = True
        
        MP3isPlaying = True
        MP3isPause = False
        
        Exit Function
    ElseIf dwReturn = 289 Then  '已打开
        '关闭后重新打开
        mciSendString "stop MP3Play", 0, 0, 0
        mciSendString "close MP3Play", 0, 0, 0
        dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
        If dwReturn = 0 Then '成功
            mciSendString "play MP3Play", 0, 0, 0
            cyMP3Play = True
            Exit Function
        ElseIf dwReturn <> 0 Then  '还是众所
            mciGetErrorString dwReturn, ret, 128
            mmOpen = ret
            'MsgBox ret, vbCritical
            Exit Function
        End If
    Else
        mciGetErrorString dwReturn, ret, 128
        MP3Err = ret
        Exit Function
    End If
End Function
Public Function MP3Pause()                      '暂停
    mciSendString "pause MP3Play", 0, 0, 0
    MP3isPlaying = False
    MP3isPause = True
End Function
  
'函数名称:MP3Stop

'作用：停止播放正在播放的MP3

'##################################################
'整理时间：2004-2-16 12:04:41
'  参数：
'##################################################
'  例：
''  '停止播放正在播放的MP3
''  sSound.MP3Stop
Public Function MP3Stop() As String             '停止
    mciSendString "stop MP3Play", 0, 0, 0
    mciSendString "close MP3Play", 0, 0, 0
End Function
Public Function MP3Length() As Long             '长度,毫秒
Static S As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play length", S, Len(S), 0
    If IsNumeric(S) = False Then '不是有效的时间
        Exit Function
    End If
    MP3Length = CLng(S)
End Function
Public Function MP3PlayPos() As Long            '现在播放到的位置
Static S As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play position", S, Len(S), 0
    If IsNumeric(S) = False Then '不是有效的时间
        Exit Function
    End If
    MP3PlayPos = S
End Function
Public Function MP3PosFormat2Time() As String   '将正在播放的时间转换为时间格式
Static S As String * 30
Dim sec As Integer
Dim mins As Integer
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play position", S, Len(S), 0
    If IsNumeric(S) = False Then '不是有效的时间
        Exit Function
    End If
    sec = Round(Mid$(S, 1, Len(S)) / 1000)
    If sec < 60 Then MP3PosFormat2Time = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        MP3PosFormat2Time = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End Function
Public Function MP3Seek(MilliSecond As Long)    '跳到指定的地方播放
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    If MP3isPlaying = True Then
        mciSendString "play MP3Play from " & MilliSecond, 0, 0, 0
    Else
        MP3Pause
        mciSendString "seek MP3Play to " & MilliSecond, 0, 0, 0
    End If
    
    mciSendString "play MP3Play", 0, 0, 0
End Function
Function cyPlayWave(sFileName As String)
    PlaySound sFileName, ByVal 0&, SND_FILENAME Or SND_ASYNC
End Function

