VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "formClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    
Private Type COPYDATASTRUCT
        dwData   As Long
        cbData   As Long
        lpData   As Long
End Type

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

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Dim m_Hwnd() As Long     '保存所有子句柄，由于不能直接返回，因此放在全局窗

Enum cyFromWnd
    FromTopWnd = 0
    FromSubWnd = 1
End Enum

Enum cyWinPosAction
    SaveWinPos = 0
    GetWinPos = 1
End Enum

Enum cyGridAction
    SetGridWidth = 0
    GetGridWidth = 1
End Enum

Enum cyWndAction
    Wnd_AllowClose = &O1
    Wnd_AllowMaximize = &O2
    Wnd_AllowMinimize = &O4
    Chk_SetValue = 11
    Chk_GetValue = 13
    Ctl_Click = 19
    Wnd_TOPMOST = -2
    Wnd_SetActive = 15
    Wnd_SetWindowText = 16
    Wnd_GetWindowText = 17
    Wnd_TranSparent = 18
    Wnd_Close = 14
    Wnd_Move = 20
    Wnd_SetWindowSize = 21
    Wnd_Show = 23
    Wnd_Maxmize = 24
    Wnd_Minmize = 25
    Wnd_MOVE_LeftTop = 30
    Wnd_MOVE_CenterTop = 31
    Wnd_MOVE_RightTop = 32
    Wnd_MOVE_LeftCenter = 33
    Wnd_MOVE_Center = 34
    Wnd_MOVE_RightCenter = 35
    Wnd_MOVE_LeftBottom = 36
    Wnd_MOVE_CenterBottom = 37
    Wnd_MOVE_RightBottom = 38
    Wnd_Destory = 39
    Wnd_Enable = 40
    Wnd_SetFocus = 42
    Lst_AddHorizentialBar = 43
    Txt_Append = 44
    Txt_AllSelect = 45
    Txt_AllClear = 46
    Txt_Copy = 47
    Txt_Paste = 48
    Txt_Cut = 53
    Txt_Undo = 54
    Wnd_IsWindow = 49
    Wnd_IsVisible = 50
    Wnd_IsEnable = 51
    LV_GirdLine = 55
    LV_FullRowSel = 56
    lst_GetListCount = 57
    Txt_GetPassWord = 58
    Wnd_DropToMove = 59
    Cbo_GetItemToArray = 60
    Lst_GetIndexFromString = 61
    Cbo_GetIndexFromString = 62
    lst_GetItemToArray = 63
    Txt_AutoType = 64
    Lst_SetListIndex = 65
    Cbo_SetListIndex = 66
    Cbo_PopupList = 67
    Txt_SetReadOnly = 68
    Txt_ScroolText = 69
    LV_AutoFitByItem = 70
    Wnd_GetWindowWidth = 71
    Wnd_GetWindowHeight = 72
    Wnd_ShowFrame = 73
    Wnd_GetWindowTop = 74
    Wnd_GetWindowLeft = 75
    Wnd_GetWindowPosSize = 76
    Wnd_SaveWindowPosSize = 77
    Ctl_HasHorScrollBar = 78
    Ctl_HasVerScrollBar = 79
    Wnd_SetParent = 80
    Wnd_GetParent = 81
    Wnd_Flash = 82
    Txt_SetReadWrite = 83
    Txt_Redo = 84
    Wnd_IsCursorOver = 85
    Wnd_CannotGetFocus = 86
    Lst_GetListIndex = 87
    Cbo_GetListIndex = 88
    
End Enum

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Enum cyWindowVersion
    Win9x = -1073739260
    WinMe = -1073718780
    WinNT = 90505220
    Win2K = 143851525
    WinXp = 170393861
    Win2003 = 248381957
End Enum


Private Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

'-----------------------------------------------------------------------------
'获得及设置窗口的STYLE
'-----------------------------------------------------------------------------
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)

Enum cyStyleAction
    SetStyle = 0
    RemoveStyle = 1
    GetStyleToArray = 2
End Enum

Enum cyWndStyle
    WS_CLOSE = &O101
    WS_BORDER = &H800000
    WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
    WS_CHILD = &H40000000
    WS_CHILDWINDOW = (WS_CHILD)
    WS_CLIPCHILDREN = &H2000000
    WS_CLIPSIBLINGS = &H4000000
    WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    WS_GROUP = &H20000
    WS_HSCROLL = &H100000
    WS_MINIMIZE = &H20000000
    WS_ICONIC = WS_MINIMIZE
    WS_MAXIMIZE = &H1000000
    WS_MAXIMIZEBOX = &H10000
    WS_SYSMENU = &H80000
    WS_THICKFRAME = &H40000
    WS_MINIMIZEBOX = &H20000
    WS_OVERLAPPED = &H0&
    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    WS_POPUP = &H80000000
    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
    WS_SIZEBOX = WS_THICKFRAME
    WS_TABSTOP = &H10000
    WS_TILED = WS_OVERLAPPED
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
    WS_VISIBLE = &H10000000
    WS_VSCROLL = &H200000
    
    
    WS_EX_ACCEPTFILES = &H10&
    WS_EX_DLGMODALFRAME = &H1&
    WS_EX_NOPARENTNOTIFY = &H4&
    WS_EX_TOPMOST = &H8&
    WS_EX_TRANSPARENT = &H20&
    WS_EX_TOOLWINDOW = &H80&
    WS_EX_MDICHILD = &H40
    WS_EX_WINDOWEDGE = &H100
    WS_EX_CLIENTEDGE = &H200
    WS_EX_CONTEXTHELP = &H400
    WS_EX_RIGHT = &H1000
    WS_EX_LEFT = &H0
    WS_EX_RTLREADING = &H2000
    WS_EX_LTRREADING = &H0
    WS_EX_LEFTSCROLLBAR = &H4000
    WS_EX_RIGHTSCROLLBAR = &H0
    WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    WS_EX_APPWINDOW = &H40000
    WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
    WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)

End Enum
'-----------------------------------------------------------------------------
'获得及设置窗口的STYLE
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'上下移动LISTBOX项
'-----------------------------------------------------------------------------
Enum cyListItemMove
    ListItemUp = 0
    ListItemDown = 1
End Enum
'-----------------------------------------------------------------------------
'上下移动LISTBOX项
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'自动从ListBox or ComboBox中查找相应的项
'-----------------------------------------------------------------------------

Enum MatchMode
    FirstWordMatch = 0
    FullWordMatch = 1
End Enum
'-----------------------------------------------------------------------------
'自动从ListBox or ComboBox中查找相应的项
'-----------------------------------------------------------------------------



'-----------------------------------------------------------------------------
'文本框自动完成
'-----------------------------------------------------------------------------
Private Declare Function SHAutoComplete Lib "Shlwapi.dll" (ByVal hwndEdit As Long, ByVal dwFlags As Long) As Long
Private Const SHACF_DEFAULT = &H0
'-----------------------------------------------------------------------------
'文本框自动完成
'-----------------------------------------------------------------------------


Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Enum cySpecialWnd
    XpDeskTop = 0
    XpTaskBar = 2
    XpTaskBar_StartButton = 1
    XpTaskBar_QuickLaunch = 3
    XpTaskBar_RunningBar = 4
    XpTaskBar_TrayNotify = 5
End Enum


Enum cyCheckStatus
    Check = 2
    unCheck = 1
    ReturnStatus = 0
End Enum

Enum cyRunningBar
    ShowBar = 0
    ShowMsg = 1
    ShowPercent = 2
    CloseBar = 3
    ReturnHwnd = 4
    BarAutoRunning = 5
End Enum

Enum cyGetParentAction
    '仅上级窗口
    FatherOnly = 0
    '第一个顶级窗口
    FatherAndTop = 1
    '所有的顶级窗口
    GrandFather = 2
End Enum

Dim sCaptionArray() As String      '全局的标题记录数组


'-----------------------------------------------------------------------------
'显示窗口边框
'-----------------------------------------------------------------------------
'当前窗口句柄
Dim hCurrent As Long

'窗口设备上下文
Dim hDCScreen As Long

Dim hPrevious As Long

Const PEN_WIDTH = 3

'范围
Dim rectCurrent As RECT
Dim rectPrev As RECT
Dim rectTemp As RECT

Dim lretSpy As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

'-----------------------------------------------------------------------------
'显示窗口边框
'-----------------------------------------------------------------------------

Dim iHwndMsgbox     '保存Msgbox的句柄


'取得该窗口是父窗口的第几个子窗口
Public Function cyGetWhickChildAmI(ByVal hWnd As Long) As Long
    Const GW_CHILD = 5
    Const GW_HWNDNEXT = 2

'获得文本框及按钮的句柄

Dim Child_hwnd As Long
Dim num_Children As Integer
Dim iFatherHwnd As Long

iFatherHwnd = GetParent(hWnd) '取得父窗口的句柄
    num_Children = 0
    Child_hwnd = GetWindow(iFatherHwnd, GW_CHILD)
    Do While Child_hwnd <> 0
        If Child_hwnd = hWnd Then
            cyGetWhickChildAmI = num_Children
            Exit Function
        End If
        
        num_Children = num_Children + 1
        Child_hwnd = GetWindow(Child_hwnd, GW_HWNDNEXT)
        
    Loop
    Exit Function

End Function

Public Function cyGetHwndToArrary(Optional ByVal Wnd As cyFromWnd, Optional ByVal hWnd As Long)
    Dim i As Long
    Const GW_HWNDNEXT = 2
    Dim iHwnd As Long
    
    If Wnd = FromTopWnd Then
        Erase m_Hwnd
        iHwnd = FindWindow(vbNullString, vbNullString)
        Do While iHwnd <> 0
        ReDim Preserve m_Hwnd(i)
        m_Hwnd(i) = iHwnd
            i = i + 1
            iHwnd = GetWindow(iHwnd, GW_HWNDNEXT)
        Loop
        cyGetHwndToArrary = m_Hwnd
        
    ElseIf Wnd = FromSubWnd Then
        Const GW_CHILD = 5
        
        '取得第一个子对象的句柄
        hWnd = GetWindow(hWnd, GW_CHILD)
        Do While hWnd <> 0
            '将句柄写到数组中保存起来
            ReDim Preserve m_Hwnd(i)
            m_Hwnd(i) = hWnd
            hWnd = GetWindow(hWnd, GW_HWNDNEXT)
            i = i + 1
        Loop
        
'        '如果数组有内容则找其下级
'        On Error Resume Next
'
'        For i = 0 To i - 1
'            cyGetHwndToArrary Wnd, m_Hwnd(i)
'        Next
        cyGetHwndToArrary = m_Hwnd
    
    End If
    
End Function

'根据标题从窗体中取得子对象的名柄
Function cyGetSubObjHwndFromCaption(ByVal hWnd As Long, sCaption As String) As Long
    Dim iA() As Long
    Dim i As Long
    iA = cyGetHwndToArrary(FromSubWnd, hWnd)
    For i = 0 To UBound(iA)
        If InStr(1, UCase(cyGetCaption(iA(i))), UCase(sCaption)) Then
            cyGetSubObjHwndFromCaption = iA(i)
            Exit Function
        End If
    Next
End Function

'取得窗体的所有子对象的标题列表
Function cyGetAllSubCaptionToArray(ByVal hWnd As Long)
    Dim sTemp As String * 256
    Dim sCaption() As String
    Dim i As Long
    ReDim Preserve m_Hwnd(0)
    '枚举所有窗口
    tmpGetHwnd hWnd
    ReDim Preserve sCaption(0)
    
    For i = 0 To UBound(m_Hwnd) - 1
        GetWindowText m_Hwnd(i), sTemp, 256      '得到窗口的标题
        sCaption(i) = Replace(sTemp, Chr(0), "")
        ReDim Preserve sCaption(i + 1)
    Next
        ReDim Preserve sCaption(i - 1)
    cyGetAllSubCaptionToArray = sCaption
End Function


Public Function cyMsgBox(Optional ByVal sMsg As String, Optional ByVal iDelaySecond As Long, Optional bReturnHwnd As Boolean)
    If sMsg = "" Then Unload frmMsg
    If bReturnHwnd = True Then
        cyMsgBox = frmMsg.hWnd
        Exit Function
    End If
    Unload frmMsg
    
    With frmMsg
    
        '写入内容
        .lblMsg.Left = (.Width - .lblMsg.Width) / 2
        
        If .lblMsg.Height = 210 Then
            .lblMsg = Trim(sMsg)
            .lblMsg.Alignment = 2 '信息只有一行,居中
        End If
        .lblMsg.Top = .Height / 2 - .lblMsg.Height / 2
        
        
        
        If iDelaySecond = 0 Then
            .tmrRemain.Enabled = False
        ElseIf iDelaySecond > 0 And iDelaySecond < 61 Then .tmrDelay.Interval = iDelaySecond * 1000
            .lblRemain = iDelaySecond
            .tmrDelay.Enabled = True
            .Label1.Visible = True                   '显示秒关闭
    
        End If
        
        .sMsg = sMsg
        .lblMsg.AutoSize = True
        
        .Show
        
    End With
End Function

Public Function cyRunningBar(Optional Action As cyRunningBar, Optional varParamater As Variant)
    
    Select Case Action
        Case ShowBar
            Dim sTemp As String
            Dim bStr As String * 255
            sTemp = Mid(bStr, 1, GetTempPath(255, bStr))
            
            SaveSetting "cyDLLsp2", "RunningBar", "Value", 0
            SaveSetting "cyDLLsp2", "RunningBar", "CommandStr", ""
        
            On Error Resume Next
            
            Dim F As New fileClass
            '如果不存在则释放文件
            If Not F.fileClassist(sTemp & "RunningBar.Bin") Then
            
                '获放该主文件为临时文件
                Dim bArray() As Byte
                bArray = LoadResData(106, "CUSTOM") '将自定义资源中106号资源读入数组OCX
                Open sTemp & "RunningBar.Bin" For Binary As #1
                    Seek #1, 1
                    Put #1, , bArray
                Close #1
            
            End If
            
            WinExec sTemp & "RunningBar.Bin", vbNormal
           
        Case ShowMsg
            SaveSetting "cyDLLsp2", "RunningBar", "CommandStr", varParamater
        
        Case ShowPercent
            If Not IsNumeric(varParamater) Then Exit Function
            If varParamater < 0 Then
                varParamater = 0
            ElseIf varParamater > 100 Then
                varParamater = 100
            End If
            SaveSetting "cyDLLsp2", "RunningBar", "Value", varParamater
        
        Case CloseBar
            SaveSetting "cyDLLsp2", "RunningBar", "CommandStr", "CloseRunningBar"
            
        Case ReturnHwnd
            cyRunningBar = GetSetting("cyDLLsp2", "RunningBar", "Hwnd")

        Case BarAutoRunning
            SaveSetting "cyDLLsp2", "RunningBar", "Value", -1
    
    End Select
End Function

Public Function cyGetClassName(ByVal hWnd As Long) As String
    Dim sTemp As String * 256
    GetClassName hWnd, sTemp, 255       '得到窗口的类
    cyGetClassName = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
End Function

'根据窗口的句柄取得窗口的标题
Public Function cyGetCaption(ByVal hWnd As Long) As String
    Dim sTemp As String * 256
    GetWindowText hWnd, sTemp, 256      '得到窗口的标题
    cyGetCaption = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
End Function

'根据窗口的标题和类取得窗口的句柄
'参数：
'sCaption--窗口的标题（部分即可）
'sClassName--窗口的类名（部分即可）
'iTimeOutSecond--延时搜索的秒数，将在此时间内循环搜索
Public Function cyGetHwndFromCaptionAndClass(Optional ByVal sCaption As String, Optional ByVal sClassName As String, Optional iTimeOutSecond As Long) As Long
    Dim iHwnd As Long
    Dim sTemp As String * 256
    Dim sTemp1 As String * 256
    Const GW_HWNDNEXT = 2
    
    '计算结束时间
    Dim timeStopTime As Date
    timeStopTime = DateAdd("s", iTimeOutSecond, Now)
    
    '得到第一个顶级窗口的句柄
    iHwnd = FindWindow(vbNullString, vbNullString)
    
    '循环检索得到所有顶级窗口
    Do While iHwnd <> 0
        '取得顶级窗口的类名
        GetClassName iHwnd, sTemp, 255
            
        '取得顶级窗口的标题
        GetWindowText iHwnd, sTemp1, 255
        
        '如果参数提供的标题及类名包含在当前顶级窗口的标题和类名中（部分匹配即可）时，则返回该窗口的句柄
        If InStr(1, UCase(sTemp), UCase(sClassName)) > 0 And InStr(1, UCase(sTemp1), UCase(sCaption)) > 0 Then
            cyGetHwndFromCaptionAndClass = iHwnd
            Exit Function
        End If
        '搜索下一个顶级窗口
        iHwnd = GetWindow(iHwnd, GW_HWNDNEXT)
        DoEvents
    Loop
    
    '当第一次搜索没有值返回则在规定时间内继续查找
    Do While Now < timeStopTime     '未到时间则继续查找
        iHwnd = FindWindow(vbNullString, vbNullString)
        Do While iHwnd <> 0
            GetClassName iHwnd, sTemp, 255
            GetWindowText iHwnd, sTemp1, 255
            If InStr(1, UCase(sTemp), UCase(sClassName)) > 0 And InStr(1, UCase(sTemp1), UCase(sCaption)) > 0 Then
                cyGetHwndFromCaptionAndClass = iHwnd
                Exit Function
            End If
            iHwnd = GetWindow(iHwnd, GW_HWNDNEXT)
            DoEvents
        Loop
    Loop
    '在指定时间内还未能找到窗口时则退出函数
End Function

'取得子对象的句柄
Public Function cyGetSubObjHwndFromHwnd(ByVal hWnd As Long, ByVal iSubObjectNo As Long) As Long
    Const GW_CHILD = 5
    Const GW_HWNDNEXT = 2

    Dim i As Long
    
    '取得第一个子对象的句柄
    hWnd = GetWindow(hWnd, GW_CHILD)
    
    Do While hWnd <> 0
        If i = iSubObjectNo Then '与子对象序列相同
            cyGetSubObjHwndFromHwnd = hWnd
            Exit Function
        Else
            hWnd = GetWindow(hWnd, GW_HWNDNEXT)
        End If
        i = i + 1
    Loop
End Function

Function cyGetParentHwnd(ByVal hWnd As Long, Optional Action As cyGetParentAction) As Long
    Dim i As Long
    
    If Action = FatherOnly Then
        cyGetParentHwnd = GetParent(hWnd)
    
    ElseIf Action = FatherAndTop Then
        i = GetParent(hWnd)
        
        '当父窗口句柄有效且该句柄不是顶级窗口时，则继续执行
        Do While i <> 0 And Not cyIsTopWnd(i)
            i = GetParent(i)
        Loop
        cyGetParentHwnd = i
        
    ElseIf Action = GrandFather Then
        Dim iTmp As Long
        iTmp = GetParent(hWnd)
        i = iTmp
        '当父窗口句柄有效且该句柄不是顶级窗口时，则继续执行
        Do While iTmp <> 0
            i = iTmp
            iTmp = GetParent(i)
            
        Loop
        cyGetParentHwnd = i
    
    End If
    
End Function
''
'''缺省是查找其最顶级窗口
'''如果bOwnerParentOnly设为真则只查找其父窗口,而不查找最顶级窗口
''Public Function cyGetParentHwndFromHwnd(ByVal iHwnd As Long, Optional bOwnerParentOnly As cyGetParentAction) As Long
''
''If bOwnerParentOnly = 1 Then '只找其第一父窗
''    Dim iArray() As Long
''    iArray = cyGetHwndArraryFromTopWnd '取得所有顶级窗
''    Dim j As Long
''    '查找其父是否其中一个顶级窗
''    Do While i <> 0
''        For j = 0 To UBound(iArray)
''            If i = iArray(j) Then '是顶级
''                cyGetParentHwndFromHwnd = i
''                Exit Function
''            End If
''        Next
''        iHwnd = i
''        i = GetParent(iHwnd)
''    Loop
''        cyGetParentHwndFromHwnd = i
''
''ElseIf bOwnerParentOnly = 0 Then  '找顶级父窗
''        Do While i <> 0
''            iHwnd = i
''            i = GetParent(iHwnd)
''        Loop
''        cyGetParentHwndFromHwnd = iHwnd
''
''ElseIf bOwnerParentOnly = 2 Then '只找其上级窗口
''        cyGetParentHwndFromHwnd = GetParent(iHwnd)
''
''End If
''End Function

'返回值类型：无

'例：
'注意：
'整理日期：2004-1-19

Function cyWndAction(ByVal hWnd As Long, ByVal Action As cyWndAction, Optional ByVal lPara1 As Long, Optional ByVal lPara2 As Long, Optional ByVal sPara3 As String, Optional ByVal sPara4 As String)

Const SWP_NOMOVE = &H2
Const SWP_NOOWNERZORDER = &H200
Const SWP_NOSIZE = &H1

Const MF_REMOVE = &H2
Const MF_APPEND = &H100&
Const SC_CLOSE = &HF060&
Const SC_MAXIMIZE = &HF030&
Const SC_MINIMIZE = &HF020&
Const SC_MOVE = &HF010&
Const BM_SETCHECK = &HF1
Const BM_GETCHECK = &HF0
Const SW_HIDE = 0
Const SW_SHOW = 5
Const SW_MAXIMIZE = 3
Const SW_MINIMIZE = 6

Const WM_GETTEXTLENGTH = &HE  '文字バイト数取得
Const WM_GETTEXT = &HD  '文字列を取得する
Const WM_CUT = &H300
Const WM_COPY = &H301  'コピー
Const WM_PASTE = &H302  '貼り付け
Const WM_CLEAR = &H303  '選択文字列を削除
Const WM_UNDO = &H304
Const EM_SETSEL = &HB1  '選択する
Const EM_REPLACESEL = &HC2  '選択文字列を置換する

Const LVM_FIRST = &H1000
Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Const LVS_EX_FULLROWSELECT = &H20
Const LVS_EX_GRIDLINES = &H1

Dim iTmp As Long    '临时变量

    Dim wSize As RECT
    Select Case Action
        Case Wnd_AllowClose
            If lPara1 = 1 Then
                AppendMenu GetSystemMenu(hWnd, 0), MF_APPEND, SC_CLOSE, ByVal 0&
            Else
                RemoveMenu GetSystemMenu(hWnd, 0), SC_CLOSE, MF_REMOVE
            End If
            
        Case Wnd_AllowMaximize
            If lPara1 = 1 Then
                AppendMenu GetSystemMenu(hWnd, 0), MF_APPEND, SC_MAXIMIZE, ByVal 0&
            Else
                RemoveMenu GetSystemMenu(hWnd, 0), SC_MAXIMIZE, MF_REMOVE
            End If
        
        Case Wnd_AllowMinimize
            If lPara1 = 1 Then
                AppendMenu GetSystemMenu(hWnd, 0), MF_APPEND, SC_MINIMIZE, ByVal 0&
            Else
                RemoveMenu GetSystemMenu(hWnd, 0), SC_MINIMIZE, MF_REMOVE
            End If
            
        Case Wnd_TOPMOST
            If lPara1 = 1 Then
                SetWindowPos hWnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            Else
                SetWindowPos hWnd, -2, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            End If
            
        Case Chk_SetValue
            If lPara1 = 1 Then
                Call SendMessageByNum(hWnd, BM_SETCHECK, 1, 0)
            Else
                Call SendMessageByNum(hWnd, BM_SETCHECK, 0, 0)
            End If
        Case Chk_GetValue
            cyWndAction = SendMessageByNum(hWnd, BM_GETCHECK, 0, 0)
        Case Wnd_Close
            Call PostMessage(hWnd, &H10, 0, ByVal 0&)
        Case Wnd_SetActive
            OpenIcon hWnd
            SetActiveWindow hWnd
        Case Wnd_SetWindowText
            SendMessageStr hWnd, &HC, 0, sPara3
        Case Wnd_GetWindowText
            Dim sBuf As String
            sBuf = String(255, Chr(0))
            GetWindowText hWnd, sBuf, 255
            cyWndAction = Left(sBuf, InStr(1, sBuf, Chr(0)) - 1)
        Case Wnd_TranSparent
            '如果系统支持才使用透明
            If lPara1 > 255 Then lPara1 = 255
            If cyCheckVersion(Win2K) Or cyCheckVersion(WinXp) Or cyCheckVersion(Win2003) Then cyTranWnd hWnd, lPara1
            
        Case Ctl_Click
            Dim iTemp As Long
            iTemp = GetParent(hWnd)
            Do While iTemp <> 0
                If cyGetClassName(iTemp) = "#32770" Then
                    OpenIcon iTemp
                    SetActiveWindow iTemp
                    Exit Do
                End If
                iTemp = GetParent(iTemp)
            Loop
                Const WM_LBUTTONDOWN = &H201
                Const WM_LBUTTONUP = &H202
                Call PostMessage(hWnd, WM_LBUTTONDOWN, 0, &H20002)
                Call PostMessage(hWnd, WM_LBUTTONUP, 0, &H20002)
        Case Wnd_Move
            GetWindowRect hWnd, wSize
            MoveWindow hWnd, lPara1, lPara2, wSize.Right - wSize.Left, wSize.Bottom - wSize.Top, 1
        Case Wnd_SetWindowSize
            GetWindowRect hWnd, wSize
            MoveWindow hWnd, wSize.Top, wSize.Left, lPara1, lPara2, 1
        Case Wnd_Show
            If lPara1 = 1 Then
                ShowWindow hWnd, SW_SHOW
            Else
                ShowWindow hWnd, SW_HIDE
            End If
        Case Wnd_Maxmize
            ShowWindow hWnd, SW_MAXIMIZE
        Case Wnd_Minmize
            ShowWindow hWnd, SW_MINIMIZE
        Case Wnd_MOVE_LeftTop
            GetWindowRect hWnd, wSize
            MoveWindow hWnd, 0, 0, wSize.Right - wSize.Left, wSize.Bottom - wSize.Top, 1
        
        Case Wnd_MOVE_CenterTop
            GetWindowRect hWnd, wSize
                            '左偏移      上偏移
            MoveWindow hWnd, (Screen.Width / Screen.TwipsPerPixelX - (wSize.Right - wSize.Left)) / 2 + lPara1, 0 + lPara2, wSize.Right - wSize.Left, wSize.Bottom - wSize.Top, 1
        
        Case Wnd_MOVE_RightTop
            GetWindowRect hWnd, wSize
                            '左偏移      上偏移
            MoveWindow hWnd, (Screen.Width / Screen.TwipsPerPixelX) - (wSize.Right - wSize.Left) + lPara1, 0 + lPara2, wSize.Right - wSize.Left, wSize.Bottom - wSize.Top, 1
        
        Case Wnd_MOVE_LeftCenter
            GetWindowRect hWnd, wSize
                            '左偏移      上偏移
            MoveWindow hWnd, 0 + lPara1, (Screen.Height / Screen.TwipsPerPixelY) / 2 - (wSize.Bottom - wSize.Top) / 2 + lPara2, wSize.Right - wSize.Left, wSize.Bottom - wSize.Top, 1
        
        Case Wnd_MOVE_Center
            GetWindowRect hWnd, wSize
                            '左偏移      上偏移
            MoveWindow hWnd, (Screen.Width / Screen.TwipsPerPixelX) / 2 - (wSize.Right - wSize.Left) / 2 + lPara1, (Screen.Height / Screen.TwipsPerPixelY) / 2 - (wSize.Bottom - wSize.Top) / 2 + lPara2, wSize.Right - wSize.Left, wSize.Bottom - wSize.Top, 1
        
        Case Wnd_MOVE_RightCenter
            GetWindowRect hWnd, wSize
                            '左偏移      上偏移
            MoveWindow hWnd, (Screen.Width / Screen.TwipsPerPixelX) - (wSize.Right - wSize.Left) + lPara1, (Screen.Height / Screen.TwipsPerPixelY) / 2 - (wSize.Bottom - wSize.Top) / 2 + lPara2, wSize.Right - wSize.Left, wSize.Bottom - wSize.Top, 1
        
        Case Wnd_MOVE_LeftBottom
            GetWindowRect hWnd, wSize
                            '左偏移      上偏移
            MoveWindow hWnd, 0 + lPara1, (Screen.Height / Screen.TwipsPerPixelY) - (wSize.Bottom - wSize.Top) + lPara2, wSize.Right - wSize.Left, wSize.Bottom - wSize.Top, 1
        
        Case Wnd_MOVE_CenterBottom
            GetWindowRect hWnd, wSize
                            '左偏移      上偏移
            MoveWindow hWnd, (Screen.Width / Screen.TwipsPerPixelX) / 2 - (wSize.Right - wSize.Left) / 2 + lPara1, (Screen.Height / Screen.TwipsPerPixelY) - (wSize.Bottom - wSize.Top) + lPara2, wSize.Right - wSize.Left, wSize.Bottom - wSize.Top, 1
        
        Case Wnd_MOVE_RightBottom
            GetWindowRect hWnd, wSize
                            '左偏移      上偏移
            MoveWindow hWnd, (Screen.Width / Screen.TwipsPerPixelX) - (wSize.Right - wSize.Left) + lPara1, (Screen.Height / Screen.TwipsPerPixelY) - (wSize.Bottom - wSize.Top) + lPara2, wSize.Right - wSize.Left, wSize.Bottom - wSize.Top, 1

        Case Wnd_Destory
            Call PostMessage(hWnd, &H2, 0, ByVal 0&) '销毁,不提示
            
        Case Wnd_Enable
            If lPara1 = 1 Then
                EnableWindow hWnd, 1
            Else
                EnableWindow hWnd, 0
            End If
        
        Case Wnd_SetFocus
            Putfocus hWnd
        Case Lst_AddHorizentialBar
            'ListBox添加滚动条
            Call SendMessageByNum(hWnd, &H194, IIf(lPara1 = 0, 300, lPara1), ByVal 0&)
        Case Txt_Append   '追加字符串到文本框
            '取得文本框的长度
            iTmp = SendMessageStr(hWnd, WM_GETTEXTLENGTH, 0&, ByVal 0&)
            '设置当前位置
            SendMessageStr hWnd, EM_SETSEL, iTmp, ByVal iTmp
            '添加到当前位置
            SendMessageStr hWnd, EM_REPLACESEL, 0, ByVal sPara3
        Case Txt_AllSelect
            '取得文本框的长度
            iTmp = SendMessageStr(hWnd, WM_GETTEXTLENGTH, 0&, ByVal 0&)
            SendMessageStr hWnd, EM_SETSEL, 0, iTmp
        Case Txt_AllClear
            '取得文本框的长度
            iTmp = SendMessageStr(hWnd, WM_GETTEXTLENGTH, 0&, ByVal 0&)
            SendMessageStr hWnd, EM_SETSEL, 0, iTmp
            SendMessageStr hWnd, WM_CLEAR, 0, 0
        Case Txt_Copy   '将文本框的字符串复制到剪贴板
            '复制当前字符串
            SendMessageStr hWnd, WM_COPY, 0&, ByVal 0&
        Case Txt_Paste   '粘贴
            SendMessageStr hWnd, WM_PASTE, 0&, ByVal 0&
        Case Txt_Cut      '剪切
            SendMessageStr hWnd, WM_CUT, 0&, ByVal 0&
        Case Txt_Undo
            SendMessageStr hWnd, WM_UNDO, 0&, ByVal 0&
        Case Wnd_IsWindow   '窗口有效
            cyWndAction = IsWindow(hWnd)
        Case Wnd_IsVisible    '窗口是否可见
            cyWndAction = IsWindowVisible(hWnd)
        Case Wnd_IsEnable    '窗口是否可用
            cyWndAction = IsWindowEnabled(hWnd)
        Case LV_GirdLine
            '取得ListView的风格
            iTmp = SendMessageByNum(hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
            If lPara1 = 1 Then
            '设置全选
                iTmp = iTmp Or LVS_EX_GRIDLINES
            Else
            '只选一格
                iTmp = iTmp Xor LVS_EX_GRIDLINES
            End If
            SendMessageLong hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, iTmp
        
        Case LV_FullRowSel
   
            '取得ListView的风格
            iTmp = SendMessageByNum(hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
            If lPara1 = 1 Then
            '设置全选
                iTmp = iTmp Or LVS_EX_FULLROWSELECT
            Else
            '只选一格
                iTmp = iTmp Xor LVS_EX_FULLROWSELECT
            End If
            SendMessageLong hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, iTmp
        Case lst_GetListCount
            cyWndAction = SendMessage(hWnd, &H18B, 0, vbNullString)
        Case Txt_GetPassWord
            Const EM_GETPASSWORDCHAR = &HD2
            Const EM_SETPASSWORDCHAR = &HCC
            Dim byteArray As Byte, sPassWord As String
            byteArray = SendMessage(hWnd, EM_GETPASSWORDCHAR, 0, 0)
            Call PostMessage(hWnd, EM_SETPASSWORDCHAR, 0, 0)
            DoEvents
            sPassWord = String(255, vbNullChar)
            Call SendMessage(hWnd, WM_GETTEXT, 260, ByVal sPassWord)
            Call PostMessage(hWnd, EM_SETPASSWORDCHAR, byteArray, 0)
            cyWndAction = Replace(sPassWord, Chr(0), "")
        Case Wnd_DropToMove
            Const WM_SYSCOMMAND = &H112
            Const HTCAPTION = 2
            ReleaseCapture
            PostMessage hWnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
        Case Cbo_GetItemToArray
    
            Const CB_GETCOUNT = &H146
            Const CB_GETLBTEXT = &H148
            Const CB_GETLBTEXTLEN = &H149
        
            Dim iCount As Long
            Dim sItem As String
            Dim sStr As String
            Dim j As Long
            
            '得到Combo的项数
            iCount = SendMessage(hWnd, CB_GETCOUNT, ByVal CLng(0), ByVal CLng(0))
            
            '没有项则退出
            If iCount = 0 Then Exit Function
            
            For j = 0 To iCount - 1
                sItem = Space(SendMessage(hWnd, CB_GETLBTEXTLEN, ByVal j, ByVal CLng(0)))
                Call SendMessage(hWnd, CB_GETLBTEXT, ByVal j, ByVal sItem)
                sStr = sStr & ";" & sItem
            Next
            cyWndAction = Split(Right(sStr, Len(sStr) - 1), ";")
            
        Case lst_GetItemToArray
            Const LB_GETCOUNT = &H18B
            Const LB_GETTEXT = &H189
            Const LB_GETTEXTLEN = &H18A
            
            iCount = SendMessage(hWnd, LB_GETCOUNT, ByVal CLng(0), ByVal CLng(0))
            If iCount > 0 Then
                For j = 0 To iCount - 1
                    sItem = Space(SendMessage(hWnd, LB_GETTEXTLEN, ByVal j, ByVal CLng(0)))
                    Call SendMessage(hWnd, LB_GETTEXT, ByVal j, ByVal sItem)
                    sStr = sStr & ";" & sItem
                Next
            End If
            cyWndAction = Split(Right(sStr, Len(sStr) - 1), ";")
        
        Case Lst_GetIndexFromString
            Const LB_FINDSTRING = &H18F

            cyWndAction = SendMessage(hWnd, LB_FINDSTRING, -1, ByVal sPara3)
        Case Cbo_GetIndexFromString
            Const CB_FINDSTRING = &H14C
            cyWndAction = SendMessage(hWnd, CB_FINDSTRING, -1, ByVal sPara3)
        
        Case Txt_AutoType
            SHAutoComplete hWnd, SHACF_DEFAULT
            
        Case Lst_SetListIndex
            Const LB_SETCURSEL = &H186
            SendMessage hWnd, LB_SETCURSEL, lPara1, 0&
        
        Case Cbo_SetListIndex
            Const CB_SETCURSEL = &H14E
            SendMessage hWnd, CB_SETCURSEL, lPara1, 0&

        Case Cbo_PopupList
            Const CB_SHOWDROPDOWN = &H14F
            SendMessage hWnd, CB_SHOWDROPDOWN, True, 0
            
        Case Txt_SetReadOnly
            Const EM_SETREADONLY = &HCF
            SendMessage hWnd, EM_SETREADONLY, True, 0
            
        Case Txt_ScroolText
            Const EM_LINESCROLL = &HB6
            SendMessage hWnd, EM_LINESCROLL, 1, ByVal lPara1
        
        Case LV_AutoFitByItem
            Const LVM_SETCOLUMNWIDTH As Long = (&H1000 + 30)
            Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
            Dim iCounter As Long
            lPara1 = IIf(lPara1 = 0, 10, lPara1)
            For iCounter = 1 To lPara1
                Call SendMessage(hWnd, LVM_SETCOLUMNWIDTH, iCounter - 1, ByVal LVSCW_AUTOSIZE_USEHEADER)
            Next
            
        Case Wnd_GetWindowTop
            GetWindowRect hWnd, wSize
            cyWndAction = wSize.Top
        Case Wnd_GetWindowLeft
            GetWindowRect hWnd, wSize
            cyWndAction = wSize.Left
        Case Wnd_GetWindowWidth
            GetWindowRect hWnd, wSize
            cyWndAction = wSize.Right - wSize.Left
        Case Wnd_GetWindowHeight
            GetWindowRect hWnd, wSize
            cyWndAction = wSize.Bottom - wSize.Top
        
        Case Wnd_ShowFrame
            If lPara1 = 1 Then
                ShowWndFrame hWnd

            ElseIf lPara1 = 0 Then
                ClearWndFrame
            End If
        Case Wnd_GetWindowPosSize
            On Error Resume Next
            Dim iL As Long
            Dim iT As Long
            Dim iW As Long
            Dim iH As Long
            
            If sPara3 = "" Then
                Call MsgBox("请输入参数sPara3(程序名称)!", vbCritical Or vbSystemModal, "")
                Exit Function
                
            End If
            
            iL = GetSetting(sPara3, "StorePosWnd", "Left", 0)
            iT = GetSetting(sPara3, "StorePosWnd", "Top", 0)
            iW = GetSetting(sPara3, "StorePosWnd", "Width", 0)
            iH = GetSetting(sPara3, "StorePosWnd", "Height", 0)
            If iL < 0 Or iT < 0 Then iL = 0: iT = 0
            '宽和高未初始化，则按缺少
            If iW = 0 Or iH = 0 Then
                MoveWindow hWnd, iL, iT, 0, 0, 1
            
            Else
                MoveWindow hWnd, iL, iT, iW, iH, 1
                
            End If
            
        
        Case Wnd_SaveWindowPosSize
'必须输入第三四个参数
            If sPara3 = "" Then
                Call MsgBox("请输入参数sPara3(程序名称)!", vbCritical Or vbSystemModal, "")
                Exit Function
                
            End If
            
            GetWindowRect hWnd, wSize
            Call SaveSetting(sPara3, "StorePosWnd", "Top", IIf(wSize.Top < 0, 0, wSize.Top))
            Call SaveSetting(sPara3, "StorePosWnd", "Left", IIf(wSize.Left < 0, 0, wSize.Left))
            Call SaveSetting(sPara3, "StorePosWnd", "Width", wSize.Right - wSize.Left)
            Call SaveSetting(sPara3, "StorePosWnd", "Height", wSize.Bottom - wSize.Top)

        Case Ctl_HasHorScrollBar
            Dim lpMinPos As Long
            Dim lpMaxPos As Long
            '获得滚动条的位置
            GetScrollRange hWnd, 0, lpMinPos, lpMaxPos
            cyWndAction = IIf(lpMaxPos <> lpMinPos, 1, 0)
        
        Case Ctl_HasVerScrollBar
            '获得滚动条的位置
            GetScrollRange hWnd, &H1, lpMinPos, lpMaxPos
            cyWndAction = IIf(lpMaxPos <> lpMinPos, 1, 0)

        Case Wnd_SetParent
        '设置新父窗口
            SetParent hWnd, lPara1
        
        Case Wnd_GetParent
        '返回父窗口句柄
            cyWndAction = GetParent(hWnd)
            
        Case Wnd_Flash
            '闪烁提示窗口已被锁定
            ShowWndFrame hWnd
            cyDelay 0.2
            ClearWndFrame
    
        Case Txt_SetReadWrite
            SendMessage hWnd, EM_SETREADONLY, False, 0
    
        Case Txt_Redo
            '与UNDO功能相同，再使用一次UNDO即REDO
            SendMessageStr hWnd, &H304, 0&, ByVal 0&
    
        Case Wnd_IsCursorOver
        '当前鼠标移动到对象上
            Dim tmpRect As RECT
            Dim tmpPoint As POINTAPI
            GetCursorPos tmpPoint
            GetWindowRect hWnd, tmpRect
            cyWndAction = PtInRect(tmpRect, tmpPoint.X, tmpPoint.Y)
        
        Case Wnd_CannotGetFocus
            Const LWA_COLORKEY = &H1
            Const LWA_ALPHA = &H2
            Const GWL_EXSTYLE = (-20)
            Const WS_EX_LAYERED = &H80000
            Const WS_EX_TRANSPARENT    As Long = &H20&
            Dim iOldWndStyle   As Long
            iOldWndStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
            
            If lPara1 = 1 Then
            '设置为窗口穿透
                
                If Not ((iOldWndStyle Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT) = 264 Or (iOldWndStyle Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT) = 256) Then
                    iOldWndStyle = iOldWndStyle Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
                    SetWindowLong hWnd, GWL_EXSTYLE, iOldWndStyle
                    SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA

                End If
                
            ElseIf lPara1 = 0 Then
            '取消窗口穿透
                
                If (iOldWndStyle Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT) = 264 Or (iOldWndStyle Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT) = 256 Then
                '是透明状态
                    iOldWndStyle = iOldWndStyle Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT
                    SetWindowLong hWnd, GWL_EXSTYLE, iOldWndStyle
                    SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
                
                End If
            
            ElseIf lPara1 = 2 Then


                If (iOldWndStyle Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT) = 264 Or (iOldWndStyle Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT) = 256 Then
                    cyWndAction = 1
                
                Else
                    cyWndAction = 0
                
                End If
            
            End If

        Case Lst_GetListIndex
            cyWndAction = SendMessage(hWnd, &H188, 0, 0&)
        
        Case Cbo_GetListIndex
            cyWndAction = SendMessage(hWnd, &H147, 0, 0&)
    
    End Select
    
End Function

Function cyGetSpecialHwnd(ByVal Wnd As cySpecialWnd)
    
    Select Case Wnd
        Case XpDeskTop
            cyGetSpecialHwnd = cyGetSubObjHwndFromHwnd(cyGetSubObjHwndFromHwnd(cyGetHwndFromCaptionAndClass(, "Progman"), 0), 0)
        Case XpTaskBar
            cyGetSpecialHwnd = cyGetHwndFromCaptionAndClass(, "Shell_TrayWnd")
        Case XpTaskBar_StartButton
            cyGetSpecialHwnd = cyGetSubObjHwndFromHwnd(cyGetHwndFromCaptionAndClass(, "Shell_TrayWnd"), 0)
        Case XpTaskBar_QuickLaunch
            cyGetSpecialHwnd = cyGetSubObjHwndFromHwnd(cyGetSubObjHwndFromHwnd(cyGetHwndFromCaptionAndClass(, "Shell_TrayWnd"), 2), 2)
        Case XpTaskBar_TrayNotify
            cyGetSpecialHwnd = cyGetSubObjHwndFromHwnd(cyGetHwndFromCaptionAndClass(, "Shell_TrayWnd"), 1)
        Case XpTaskBar_RunningBar
            cyGetSpecialHwnd = cyGetSubObjHwndFromHwnd(cyGetSubObjHwndFromHwnd(cyGetHwndFromCaptionAndClass(, "Shell_TrayWnd"), 2), 1)
    End Select
    
End Function

'将窗体的子对象的标题保存到字符串数组中
Function cyGetSubObjCaptionToArray(ByVal hWnd As Long)
    Dim sArray() As String
    Dim i As Long        '暂时存放窗口子对象的句柄
    Dim iCount As Long   '存放窗口子对象的个数
    
    '取得第一个子对象的句柄
    i = cyGetSubObjHwndFromHwnd(hWnd, 0)
    '设置停止条件
    Do While i > 0
        '重定义数组大小
        ReDim Preserve sArray(iCount)
        '对数组赋值
        sArray(iCount) = cyGetCaption(i)
        '将主对象数目+1
        iCount = iCount + 1
        '取得下一个子对象的句柄,如果句柄为0则表示没有找到下一个子对象,则停止搜索
        i = cyGetSubObjHwndFromHwnd(hWnd, iCount)
    Loop
    '返回数组
    cyGetSubObjCaptionToArray = sArray
End Function

'将窗体的子对象的类名保存到字符串数组中
Function cyGetSubObjClassNameToArray(ByVal hWnd As Long)
    Dim sArray() As String
    Dim i As Long        '暂时存放窗口子对象的类名
    Dim iCount As Long   '存放窗口子对象的个数
    
    '取得第一个子对象的句柄
    i = cyGetSubObjHwndFromHwnd(hWnd, 0)
    '设置停止条件
    Do While i > 0
        '重定义数组大小
        ReDim Preserve sArray(iCount)
        '对数组赋值
        sArray(iCount) = cyGetClassName(i)
        '将主对象数目+1
        iCount = iCount + 1
        '取得下一个子对象的句柄,如果句柄为0则表示没有找到下一个子对象,则停止搜索
        i = cyGetSubObjHwndFromHwnd(hWnd, iCount)
    Loop
    '返回数组
    cyGetSubObjClassNameToArray = sArray
End Function

'取得子对象的父对象列表 , 如:   0|5|3|3
'列表就是某子对象是其父对象的第几个子对象,而它本身又是其父的第几个子对象
'最上的是该子对象的顶级窗口,第二级是该子对象的顶级窗口的第几个子对象,如此类推
'中间以|分隔
'*由于部分#32770窗并不一定是最顶窗,因此那则就要设置该弹出窗为
Function cyGetSubObjPosList(hWnd As Long) As String
    Dim i As Long
    Dim iTmp As Long
    Dim sList As String
    
    iTmp = GetParent(hWnd)
    sList = ";" & cyGetWhickChildAmI(hWnd)
    i = iTmp
    '当父窗口句柄有效且该句柄不是顶级窗口时，则继续执行
    Do While iTmp <> 0 And Not cyIsTopWnd(iTmp)
        sList = ";" & cyGetWhickChildAmI(iTmp) & sList
        i = iTmp
        iTmp = GetParent(i)
        If iTmp > 0 Then
        Else
            Exit Do
        End If
        
    Loop
    cyGetSubObjPosList = Replace(sList, ";", "", , 1)
    
    'cyGetParentHwnd = i
    
    
'
'''无设置顶级窗,则以其最顶级窗口为结束
''If TopWndHwnd = 0 Then
'    '循环查找直到其顶级
'    Do While Hwnd <> iTmp
'
'        sList = ";" & cyGetWhickChildAmI(Hwnd) & sList
'        iTmp = Hwnd '直到两次查找都是同一个窗体时认为已找到顶级窗
'        Hwnd = cyGetParentHwnd(Hwnd, FatherOnly)
'    Loop
'    On Error Resume Next
'    cyGetSubObjPosList = Right(sList, Len(sList) - 5)
'
'Else    '如果是#32770窗则以其为顶级
'    Do While Hwnd <> iTmp
'        sList = "|" & cyGetWhickChildAmI(Hwnd) & sList
'        iTmp = Hwnd
'        Hwnd = cyGetParentHwndFromHwnd(Hwnd, FatherOnly)
'        If Hwnd = TopWndHwnd Then
'        '如果其父窗口就是设置的顶级窗体则返回列表
'            cyGetSubObjPosList = Right(sList, Len(sList) - 1)
'            Exit Function
'        End If
'    Loop
'
'End If
End Function

'根据子对象列表返回句柄
'对于类名为:#32770的弹出式窗口,应提供该窗口的句柄为父窗口句柄
Function cyGetHwndFromSubObjPosList(hWnd As Long, sSubObjPosList As String) As Long
    Dim sArray() As String
    Dim i As Long
    Dim iTemp As Long
    iTemp = hWnd
    sSubObjPosList = Replace(sSubObjPosList, ";", "|")
    sArray = Split(sSubObjPosList, "|")
    For i = 0 To UBound(sArray)
        iTemp = cyGetSubObjHwndFromHwnd(iTemp, sArray(i))
    Next
    cyGetHwndFromSubObjPosList = iTemp
End Function

'智能间隔
Function cyGridsWidth(Grid As Object, Optional Action As cyGridAction, Optional ByVal sWidthList As String) As String
On Error GoTo Err
    Dim iWidth As Long     '存储需要减去的宽度,如果有滚动条则减567,没有则减338
    Dim sList As String
    Dim iCol() As String
    Dim i As Long
    Dim iSum As Long
    iCol = Split(sWidthList, ";")
    With Grid
        If TypeName(Grid) = "DataGrid" Then
            If Action = SetGridWidth Then
        
                iWidth = .Width - IIf((.RowHeight * .ApproxCount - .Height) > 0, 567, 338)
                For i = 0 To UBound(iCol)
                    If iCol(i) = "*" Then   '要求占用剩余的长度
                        .Columns(i).Width = iWidth - iSum
                        'Exit Sub
                    Else                    '可用总长/10 * 要求的宽度
                        If IsNumeric(iCol(i)) Then  '是数字,则计算
                            .Columns(i).Width = (iWidth / 10) * iCol(i)
                        Else
                        '不是有效数字则忽略其长度
                            If Left(iCol(i), 1) = "[" Then  '如果左边第一个字符是[则认为是以实际长度为准
                                .Columns(i).Width = Mid(iCol(i), 2, Len(iCol(i)) - 2)
                            End If
                        End If
                    End If
                    iSum = iSum + .Columns(i).Width
                Next
            ElseIf Action = GetGridWidth Then
                On Error GoTo ErrGetGrid
                For i = 0 To 100
                    sList = sList & ";" & CLng(Grid.Columns(i).Width)
                Next
                Exit Function
            End If
        ElseIf TypeName(Grid) = "MSHFlexGrid" Then
            If Action = SetGridWidth Then
                iWidth = .Width - IIf((.RowHeight(1) * .Rows - .Height) > 0, 330, 75)
                For i = 0 To UBound(iCol)
                    If iCol(i) = "*" Then   '要求占用剩余的长度
                        .ColWidth(i) = iWidth - iSum
                        'Exit Sub
                    Else                    '可用总长/10 * 要求的宽度
                        If IsNumeric(iCol(i)) Then  '是数字,则计算
                            .ColWidth(i) = (iWidth / 10) * iCol(i)
                        Else
                        '不是有效数字则忽略其长度
                            If Left(iCol(i), 1) = "[" Then  '如果左边第一个字符是[则认为是以实际长度为准
                                .ColWidth(i) = Mid(iCol(i), 2, Len(iCol(i)) - 2)
                                
                            End If
                        End If
                    End If
                    iSum = iSum + .ColWidth(i)
                Next
            ElseIf Action = GetGridWidth Then
                On Error GoTo ErrGetGrid
                For i = 0 To 100
                    sList = sList & ";" & Format((.ColWidth(i).Width / Grid.Width) * 10, "0.0")
                Next
                Exit Function
            End If
        
        ElseIf TypeName(Grid) = "ListView" Then
            If Action = SetGridWidth Then
    
                iWidth = .Width
                For i = 0 To UBound(iCol)
                    If iCol(i) = "*" Then   '要求占用剩余的长度
                        .ColumnHeaders(i + 1).Width = iWidth - iSum - 350
                        'Exit Sub
                    Else                    '可用总长/10 * 要求的宽度
                        If IsNumeric(iCol(i)) Then  '是数字,则计算
                            .ColumnHeaders(i + 1).Width = (iWidth / 10) * iCol(i)
                        Else
                        '不是有效数字则忽略其长度
                            If Left(iCol(i), 1) = "[" Then  '如果左边第一个字符是[则认为是以实际长度为准
                                .ColumnHeaders(i + 1).Width = Mid(iCol(i), 2, Len(iCol(i)) - 2)
                                
                            End If
                        End If
                    End If
                    iSum = iSum + .ColumnHeaders(i + 1).Width
                Next
            ElseIf Action = GetGridWidth Then
                On Error GoTo ErrGetGrid
                For i = 1 To 100
                    sList = sList & ";" & CLng(Grid.ColumnHeaders(i).Width)
                Next
                Exit Function
    
            End If
        End If
    End With
Err:
    Exit Function
ErrGetGrid:
    If Err.Number = 9 Or Err.Number = 35600 Then
        cyGridsWidth = Replace(sList, ";", "", , 1)
        
    End If
End Function

Function cySendMessage(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Variant) As Long
    cySendMessage = SendMessage(hWnd, wMsg, wParam, lParam)
End Function

Sub cyMoveListItem(ByVal ListBoxObject As Object, ByVal ListItemAction As cyListItemMove)
    Dim i As Long
    Dim iItemData As Long
    Dim sStr As String
    i = ListBoxObject.ListIndex
    If i = -1 Then Exit Sub
        
    If ListItemAction = ListItemUp Then
        '在第一行则不再向上
        If i = 0 Then Exit Sub
        sStr = ListBoxObject.List(ListBoxObject.ListIndex)
        iItemData = ListBoxObject.itemData(ListBoxObject.ListIndex)
        ListBoxObject.RemoveItem i
        ListBoxObject.AddItem sStr, i - 1
        ListBoxObject.itemData(i - 1) = iItemData
        ListBoxObject.ListIndex = i - 1
    
    ElseIf ListItemAction = ListItemDown Then
        '在最后一行则不再向下
        If i = ListBoxObject.ListCount - 1 Then Exit Sub
        sStr = ListBoxObject.List(ListBoxObject.ListIndex)
        iItemData = ListBoxObject.itemData(ListBoxObject.ListIndex)
        ListBoxObject.RemoveItem i
        ListBoxObject.AddItem sStr, i + 1
        ListBoxObject.itemData(i + 1) = iItemData
        ListBoxObject.ListIndex = i + 1

    End If
End Sub
 
Function cyWndStyleAction(ByVal hWnd As Long, ByVal Action As cyStyleAction, Optional ByVal Style As cyWndStyle)

Const SWP_NOMOVE = &H2
Const SWP_NOOWNERZORDER = &H200
Const SWP_NOSIZE = &H1

Const MF_REMOVE = &H2
Const MF_APPEND = &H100&
Const SC_CLOSE = &HF060&
Const SC_MAXIMIZE = &HF030&
Const SC_MINIMIZE = &HF020&
Const SC_MOVE = &HF010&
Const BM_SETCHECK = &HF1
Const BM_GETCHECK = &HF0
Const SW_HIDE = 0
Const SW_SHOW = 5
Const SW_MAXIMIZE = 3
Const SW_MINIMIZE = 6

Const WM_GETTEXTLENGTH = &HE  '文字バイト数取得
Const WM_GETTEXT = &HD  '文字列を取得する
Const WM_CUT = &H300
Const WM_COPY = &H301  'コピー
Const WM_PASTE = &H302  '貼り付け
Const WM_CLEAR = &H303  '選択文字列を削除
Const WM_UNDO = &H304
Const EM_SETSEL = &HB1  '選択する
Const EM_REPLACESEL = &HC2  '選択文字列を置換する


Dim iStyle As Long
iStyle = GetWindowLong(hWnd, GWL_STYLE)


If Action = SetStyle Then
    If Action = WS_CLOSE Then
        AppendMenu GetSystemMenu(hWnd, 0), MF_APPEND, SC_CLOSE, ByVal 0&
        
    Else
        SetWindowLong hWnd, GWL_STYLE, iStyle Or Style
    End If
ElseIf Action = RemoveStyle Then
    If Action = WS_CLOSE Then
        RemoveMenu GetSystemMenu(hWnd, 0), SC_CLOSE, MF_REMOVE
        
    Else
        SetWindowLong hWnd, GWL_STYLE, iStyle Xor Style
    End If
    
ElseIf Action = GetStyleToArray Then
    Dim sArray() As String
    ReDim Preserve sArray(0)
    
    If (iStyle And WS_BORDER) Then
        sArray(UBound(sArray)) = "WS_BORDER"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_CAPTION) Then
        sArray(UBound(sArray)) = "WS_CAPTION"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_CHILD) Then
        sArray(UBound(sArray)) = "WS_CHILD"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_CHILDWINDOW) Then
        sArray(UBound(sArray)) = "WS_CHILDWINDOW"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_CLIPCHILDREN) Then
        sArray(UBound(sArray)) = "WS_CLIPCHILDREN"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_CLIPSIBLINGS) Then
        sArray(UBound(sArray)) = "WS_CLIPSIBLINGS"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_CLIPSIBLINGS) Then
        sArray(UBound(sArray)) = "WS_CLIPSIBLINGS"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_DISABLED) Then
        sArray(UBound(sArray)) = "WS_DISABLED"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_DLGFRAME) Then
        sArray(UBound(sArray)) = "WS_DLGFRAME"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_GROUP) Then
        sArray(UBound(sArray)) = "WS_GROUP"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_HSCROLL) Then
        sArray(UBound(sArray)) = "WS_HSCROLL"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_MINIMIZE) Then
        sArray(UBound(sArray)) = "WS_MINIMIZE"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_ICONIC) Then
        sArray(UBound(sArray)) = "WS_ICONIC"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_MAXIMIZE) Then
        sArray(UBound(sArray)) = "WS_MAXIMIZE"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_MAXIMIZEBOX) Then
        sArray(UBound(sArray)) = "WS_MAXIMIZEBOX"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_SYSMENU) Then
        sArray(UBound(sArray)) = "WS_SYSMENU"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_THICKFRAME) Then
        sArray(UBound(sArray)) = "WS_THICKFRAME"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_MINIMIZEBOX) Then
        sArray(UBound(sArray)) = "WS_MINIMIZEBOX"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_OVERLAPPED) Then
        sArray(UBound(sArray)) = "WS_OVERLAPPED"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_OVERLAPPEDWINDOW) Then
        sArray(UBound(sArray)) = "WS_OVERLAPPEDWINDOW"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_POPUP) Then
        sArray(UBound(sArray)) = "WS_POPUP"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_POPUPWINDOW) Then
        sArray(UBound(sArray)) = "WS_POPUPWINDOW"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_SIZEBOX) Then
        sArray(UBound(sArray)) = "WS_SIZEBOX"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_TABSTOP) Then
        sArray(UBound(sArray)) = "WS_TABSTOP"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_TILED) Then
        sArray(UBound(sArray)) = "WS_TILED"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_TILEDWINDOW) Then
        sArray(UBound(sArray)) = "WS_TILEDWINDOW"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_VISIBLE) Then
        sArray(UBound(sArray)) = "WS_VISIBLE"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_VSCROLL) Then
        sArray(UBound(sArray)) = "WS_VSCROLL"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    
    
    
    If (iStyle And WS_EX_ACCEPTFILES) Then
        sArray(UBound(sArray)) = "WS_EX_ACCEPTFILES"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_DLGMODALFRAME) Then
        sArray(UBound(sArray)) = "WS_EX_DLGMODALFRAME"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_NOPARENTNOTIFY) Then
        sArray(UBound(sArray)) = "WS_EX_NOPARENTNOTIFY"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_TOPMOST) Then
        sArray(UBound(sArray)) = "WS_EX_TOPMOST"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_TRANSPARENT) Then
        sArray(UBound(sArray)) = "WS_EX_TRANSPARENT"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_TOOLWINDOW) Then
        sArray(UBound(sArray)) = "WS_EX_TOOLWINDOW"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_MDICHILD) Then
        sArray(UBound(sArray)) = "WS_EX_MDICHILD"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_WINDOWEDGE) Then
        sArray(UBound(sArray)) = "WS_EX_WINDOWEDGE"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_CLIENTEDGE) Then
        sArray(UBound(sArray)) = "WS_EX_CLIENTEDGE"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_CONTEXTHELP) Then
        sArray(UBound(sArray)) = "WS_EX_CONTEXTHELP"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_RIGHT) Then
        sArray(UBound(sArray)) = "WS_EX_RIGHT"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_LEFT) Then
        sArray(UBound(sArray)) = "WS_EX_LEFT"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_RTLREADING) Then
        sArray(UBound(sArray)) = "WS_EX_RTLREADING"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_LTRREADING) Then
        sArray(UBound(sArray)) = "WS_EX_LTRREADING"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_CLIENTEDGE) Then
        sArray(UBound(sArray)) = "WS_EX_CLIENTEDGE"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_RIGHTSCROLLBAR) Then
        sArray(UBound(sArray)) = "WS_EX_RIGHTSCROLLBAR"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_CONTROLPARENT) Then
        sArray(UBound(sArray)) = "WS_EX_CONTROLPARENT"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_STATICEDGE) Then
        sArray(UBound(sArray)) = "WS_EX_STATICEDGE"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_APPWINDOW) Then
        sArray(UBound(sArray)) = "WS_EX_APPWINDOW"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_OVERLAPPEDWINDOW) Then
        sArray(UBound(sArray)) = "WS_EX_OVERLAPPEDWINDOW"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    If (iStyle And WS_EX_PALETTEWINDOW) Then
        sArray(UBound(sArray)) = "WS_EX_PALETTEWINDOW"
        ReDim Preserve sArray(UBound(sArray) + 1)
    End If
    
    ReDim Preserve sArray(UBound(sArray) - 1)
    cyWndStyleAction = sArray
End If

End Function

Private Function cyCheckVersion(isVersion As cyWindowVersion) As Boolean
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


Private Sub cyTranWnd(ByVal hWnd As Long, Optional ByVal lPara1 As Byte)
    
    Dim iOldWndStyle As Long
    If IsWindow(hWnd) = 0 Then Exit Sub
    iOldWndStyle = GetWindowLong(hWnd, -20)
    Debug.Print iOldWndStyle
    If lPara1 >= 255 Then
        iOldWndStyle = iOldWndStyle Xor &H80000
        SetWindowLong hWnd, -20, iOldWndStyle
        SetLayeredWindowAttributes hWnd, 0, lPara1, &H2
    
    Else
        iOldWndStyle = iOldWndStyle Or &H80000
        SetWindowLong hWnd, -20, iOldWndStyle
        SetLayeredWindowAttributes hWnd, 0, lPara1, &H2
        
    End If
        
End Sub


Function cyGetAllSubHwndToArrary(ByVal hWnd As Long)
    ReDim Preserve m_Hwnd(0)
    '枚举所有窗口
    tmpGetHwnd hWnd
    ReDim Preserve m_Hwnd(UBound(m_Hwnd) - 1)
    cyGetAllSubHwndToArrary = m_Hwnd
End Function

'取得子对象的句柄列表
Private Sub tmpGetHwnd(ByVal iHwnd As Long)

    Const GW_CHILD = 5
    Const GW_HWNDNEXT = 2
    
    Dim hWnd() As Long
    Dim i As Long
    
    '取得第一个子对象的句柄
    iHwnd = GetWindow(iHwnd, GW_CHILD)
    Do While iHwnd <> 0
        '将句柄写到数组中保存起来
        ReDim Preserve hWnd(i)
        hWnd(i) = iHwnd
        m_Hwnd(UBound(m_Hwnd)) = iHwnd
        ReDim Preserve m_Hwnd(UBound(m_Hwnd) + 1)
        iHwnd = GetWindow(iHwnd, GW_HWNDNEXT)
        i = i + 1
    Loop
    
    For i = 0 To i - 1
        tmpGetHwnd hWnd(i)
    Next
End Sub

'检查是否在顶級窗口内
Public Function cyIsTopWnd(ByVal hWnd As Long) As Boolean
    Dim iHwnd As Long
    Const GW_HWNDNEXT = 2
    
    Erase m_Hwnd
    iHwnd = FindWindow(vbNullString, vbNullString)
    Do While iHwnd <> 0
        If iHwnd = hWnd Then
            cyIsTopWnd = True
            Exit Function
        End If
                
        iHwnd = GetWindow(iHwnd, GW_HWNDNEXT)
    Loop

End Function

'从类名及标题名得到子窗口的句柄
Public Function cyGetHwndFromSubWndCaptionAndClass(ByVal hWnd As Long, Optional ByVal sCaptionName As String, Optional ByVal sClassName As String) As Long
    Dim sTemp As String * 256
    Dim sTemp1 As String * 256
    
    ReDim Preserve m_Hwnd(0)
    '枚举所有子窗口句柄
    tmpGetHwnd hWnd
    Dim i As Long
    For i = 0 To UBound(m_Hwnd) - 1
        GetClassName m_Hwnd(i), sTemp, 255
        GetWindowText m_Hwnd(i), sTemp1, 255
        If InStr(1, UCase(sTemp), UCase(sClassName)) > 0 And InStr(1, UCase(sTemp1), UCase(sCaptionName)) > 0 Then
            cyGetHwndFromSubWndCaptionAndClass = m_Hwnd(i)
            Exit Function
        End If
    Next
End Function

Function cyGetPopupHwndByHotKey(ByVal sMainCaption As String, ByVal sPopupCaption As String, ByVal FunKey As KeyCodeConstants, ByVal Key1 As KeyCodeConstants, Optional ByVal Key2 As KeyCodeConstants) As Long
    Dim i As Long

    i = cyGetPopupHwnd(sMainCaption, sPopupCaption, FunKey, Key1, Key2)
    If i > 0 Then
        '有值返回
        cyGetPopupHwndByHotKey = i
        Exit Function
    End If
    
    i = cyGetPopupHwnd(sMainCaption, sPopupCaption, FunKey, Key1, Key2)
    If i > 0 Then
        '有值返回
        cyGetPopupHwndByHotKey = i
        Exit Function
    End If
    
    i = cyGetPopupHwnd(sMainCaption, sPopupCaption, FunKey, Key1, Key2)
    If i > 0 Then
        '有值返回
        cyGetPopupHwndByHotKey = i
        Exit Function
    End If
    
End Function

'获得弹出窗口的句柄
Private Function cyGetPopupHwnd(ByVal sMainCaption As String, ByVal sPopupCaption As String, ByVal FunKey As KeyCodeConstants, ByVal Key1 As KeyCodeConstants, Optional ByVal Key2 As KeyCodeConstants) As Long
    Dim S As New systemClass
    Dim i As Long

    '检查句柄是否存在
    i = cyGetHwndFromCaptionAndClass(sPopupCaption)

    '弹出窗口已存在，则直接返回句柄
    If i > 0 Then
        cyGetPopupHwnd = i
        Exit Function
    End If

    '否则先寻找主窗口句柄
    i = cyGetHwndFromCaptionAndClass(sMainCaption)

    If i = 0 Then Exit Function
    cyWndAction i, Wnd_TOPMOST, 1
    SetActiveWindow i
    
    If Key2 = 0 Then '只使用了两个快捷键
        S.cyKeyBoardAction , FunKey, Key1
    Else '使用了三个快捷键如：alt+F+V
        S.cyKeyBoardAction KB_MenuAltPress, FunKey, Key1, Key2
    End If

    cyWndAction i, Wnd_TOPMOST, 0
    cyGetPopupHwnd = Me.cyGetHwndFromCaptionAndClass(sPopupCaption, , 1)
End Function

Private Sub Class_Initialize()
    hDCScreen = GetDC(0&)

End Sub

Private Sub Class_Terminate()
    If hDCScreen <> 0 Then ReleaseDC 0&, hDCScreen
    SaveSetting "cyDLLsp2", "RunningBar", "CommandStr", "CloseRunningBar"

End Sub

'显示窗口虚框
Private Sub ShowWndFrame(ByVal hCurrent As Long)
    
    ClearWndFrame
    
    ' Get rectangle
    lretSpy = GetWindowRect(hCurrent, rectCurrent)
    
    ' Top line
    rectTemp = rectPrev
    rectTemp.Bottom = rectPrev.Top + PEN_WIDTH
    InvertRect hDCScreen, rectTemp
    ' Bottom line
    rectTemp = rectPrev
    rectTemp.Top = rectPrev.Bottom - PEN_WIDTH
    InvertRect hDCScreen, rectTemp
    ' Left side
    rectTemp = rectPrev
    rectTemp.Right = rectPrev.Left + PEN_WIDTH
    InvertRect hDCScreen, rectTemp
    ' Right side
    rectTemp = rectPrev
    rectTemp.Left = rectPrev.Right - PEN_WIDTH
    InvertRect hDCScreen, rectTemp
    
    ' Invert the borders of new rectangle
    ' Top line
    rectTemp = rectCurrent
    rectTemp.Bottom = rectCurrent.Top + PEN_WIDTH
    InvertRect hDCScreen, rectTemp
    ' Bottom line
    rectTemp = rectCurrent
    rectTemp.Top = rectCurrent.Bottom - PEN_WIDTH
    InvertRect hDCScreen, rectTemp
    ' Left side
    rectTemp = rectCurrent
    rectTemp.Right = rectCurrent.Left + PEN_WIDTH
    InvertRect hDCScreen, rectTemp
    ' Right side
    rectTemp = rectCurrent
    rectTemp.Left = rectCurrent.Right - PEN_WIDTH
    InvertRect hDCScreen, rectTemp
    
    ' Update previous
    hPrevious = hCurrent
    rectPrev = rectCurrent

End Sub

Private Sub ClearWndFrame()
   
' Invert the borders of current rectangle
' Top line
rectTemp = rectPrev
rectTemp.Bottom = rectCurrent.Top + PEN_WIDTH
InvertRect hDCScreen, rectTemp
' Bottom line
rectTemp = rectPrev
rectTemp.Top = rectCurrent.Bottom - PEN_WIDTH
InvertRect hDCScreen, rectTemp
' Left side
rectTemp = rectPrev
rectTemp.Right = rectCurrent.Left + PEN_WIDTH
InvertRect hDCScreen, rectTemp
' Right side
rectTemp = rectPrev
rectTemp.Left = rectCurrent.Right - PEN_WIDTH
InvertRect hDCScreen, rectTemp

' Clear all variables
hCurrent = 0
hPrevious = 0
rectPrev.Bottom = 0
rectPrev.Top = 0
rectPrev.Left = 0
rectPrev.Right = 0

End Sub

'获得弹出窗口的句柄
Function cyGetHwndFromPopupWnd(ByVal sMainCaption As String, ByVal sMainClass As String, ByVal sPopupCaption As String, ByVal sPopupClass, ByVal FunKey As KeyCodeConstants, ByVal Key1 As KeyCodeConstants, Optional ByVal Key2 As KeyCodeConstants) As Long
    Dim iHwnd As Long
    
    '找到并最大化主窗口
    iHwnd = cyGetHwndFromCaptionAndClass(sPopupCaption, sPopupClass)
    
    '弹出窗口已存在，则直接返回句柄
    If iHwnd > 0 Then
        cyGetHwndFromPopupWnd = iHwnd
        Exit Function
    
    End If
    
    '找到并最大化主窗口
    iHwnd = cyGetHwndFromCaptionAndClass(sMainCaption, sMainClass)
    OpenIcon iHwnd
        
    '激活窗口
    cyWndAction iHwnd, Wnd_SetActive
    
    Dim S As New systemClass
    If Key2 = 0 Then '只使用了两个快捷键
        S.cyKeyBoardAction , FunKey, Key1
        
    Else '使用了三个快捷键如：alt+F+V
    
        S.cyKeyBoardAction KB_MenuAltPress, 18, Key1, Key2
    End If
    
    '找到并最大化主窗口
    iHwnd = cyGetHwndFromCaptionAndClass(sPopupCaption, sPopupClass, 3)
    
    '弹出窗口已存在，则直接返回句柄
    If iHwnd > 0 Then
        cyGetHwndFromPopupWnd = iHwnd
        Exit Function
    
    End If
    
End Function

Private Sub cyDelay(ByVal singleSecond As Single)
    Dim iStopTime As Long
    iStopTime = GetTickCount + singleSecond * 1000
    Do While GetTickCount < iStopTime '一直运行直至到达时间
        DoEvents
        
    Loop
StopTag:
End Sub

'函数名称:cyCopyData
'作用：向窗体发送WM_COPYDATA消息
'===========================================================
'整理时间：2007-05-18 11:19
'  参数：
'  例子：     debug.print cyCopyData(109987,me.hwnd, 32, "DataForSend")
'             第二个参数是本身的窗口句柄，用于接收程序向本窗口返回必要信息，然后使用子类进行接收
'             第三个参数是消息的标识符，有部分程序只处理固定的标识符的消息

'===========================================================
Function cyCopyData(ByVal hWnd As Long, ByVal OwnerHwnd As Long, Optional ByVal dwData As Long, Optional ByVal sSendText As String)
    Dim SendData     As COPYDATASTRUCT
    SendData.dwData = dwData                '数据标识号，部分程序只接收指定标识号的消息
    SendData.cbData = LenB(sSendText)
    SendData.lpData = StrPtr(sSendText)
    cyCopyData = SendMessage(hWnd, &H4A, OwnerHwnd, SendData)
    
End Function

'函数名称:cyGetCopyDataString
'作用：将别人发过来的WM_COPYDATA消息转换成文本
'===========================================================
'整理时间：2007-05-18 11:19
'  参数：
'  例子：
'                    Private Sub Sub1_WindowMsgEvent(ByVal sMsg As String, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'                        Debug.Print W.cyGetCopyDataString(lParam)
'
'                    End Sub
'===========================================================
Function cyGetCopyDataString(ByVal lParam As Long) As String
    
    Dim CDS As COPYDATASTRUCT
    Dim sStr As String
    CopyMemory CDS, ByVal lParam, Len(CDS)
    sStr = Space(CDS.cbData)
    CopyMemory ByVal sStr, ByVal CDS.lpData, CDS.cbData
    cyGetCopyDataString = StrConv(sStr, vbFromUnicode)

End Function


'函数名称:cyWindowPos
'作用：保存/读出窗口的位置及大小
'===========================================================
'整理时间：2007-06-06 12:45
'  参数：
'  例子：     debug.print cyCopyData(109987,me.hwnd, 32, "DataForSend")
'             第二个参数是本身的窗口句柄，用于接收程序向本窗口返回必要信息，然后使用子类进行接收
'             第三个参数是消息的标识符，有部分程序只处理固定的标识符的消息

'===========================================================
Sub cyWndPosition(ByVal hWnd As Long, ByVal sSystemName As String, ByVal sWinName As String, ByVal sWinPosAction As cyWinPosAction)
    On Error Resume Next
    Dim wSize As RECT
    Dim iL As Long
    Dim iT As Long
    Dim iW As Long
    Dim iH As Long
    
    Select Case sWinPosAction
        Case SaveWinPos
            GetWindowRect hWnd, wSize
            Call SaveSetting(sSystemName, sWinName & "Position", "Top", IIf(wSize.Top < 0, 0, wSize.Top))
            Call SaveSetting(sSystemName, sWinName & "Position", "Left", IIf(wSize.Left < 0, 0, wSize.Left))
            Call SaveSetting(sSystemName, sWinName & "Position", "Width", wSize.Right - wSize.Left)
            Call SaveSetting(sSystemName, sWinName & "Position", "Height", wSize.Bottom - wSize.Top)
        
        Case GetWinPos
            
            iL = GetSetting(sSystemName, sWinName & "Position", "Left", 0)
            iT = GetSetting(sSystemName, sWinName & "Position", "Top", 0)
            iW = GetSetting(sSystemName, sWinName & "Position", "Width", 0)
            iH = GetSetting(sSystemName, sWinName & "Position", "Height", 0)
            If iL < 0 Or iT < 0 Then iL = 0: iT = 0
            '宽和高未初始化，则居中
            If iL = 0 Or iT = 0 Or iW = 0 Or iH = 0 Then
                cyWndAction hWnd, Wnd_MOVE_Center
            
            Else
                MoveWindow hWnd, iL, iT, iW, iH, 1
                
            End If
            
    
    End Select
    
End Sub

'函数名称:cyGetListviewWidths
'作用：读出ListView的列宽
'===========================================================
'整理时间：2007-11-21 21:57
'  参数：
'   Dim sCaptions As String
'   Dim sWidths As String
'   W.cyGetListviewWidths lv, sCaptions, sWidths

'===========================================================
Sub cyGetListviewWidths(ByVal lvListview As Object, ByRef sCaptionList As String, ByRef sWidthList As String)
    Dim i As Long
    '循环读出所有列的文本及宽度
    For i = 1 To lvListview.ColumnHeaders.Count
        sCaptionList = sCaptionList & ";" & lvListview.ColumnHeaders(i).Text
        sWidthList = sWidthList & ";" & CLng(lvListview.ColumnHeaders(i).Width)
    
    Next
    
    sCaptionList = Replace(sCaptionList, ";", 1, 1, 1)
    sWidthList = Replace(sWidthList, ";", 1, 1, 1)

End Sub

'函数名称:cySetListviewWidths
'作用：设置ListView的列宽
'===========================================================
'整理时间：2007-08-23 12:45
'  参数：
'     Dim W As New formClass
'     W.cySetListviewWidths lv, "234;234;234;234", "1000;2000;1320;2343"

'===========================================================
Sub cySetListviewWidths(ByVal lvListview As Object, ByVal sCaptionList As String, ByVal sWidthList As String)

    Dim i As Long
    Dim j As Long
    Dim W As New formClass
    Dim sA() As String
    Dim sB() As String
    
    lvListview.View = 3 'lvmReport
    
    W.cyWndAction lvListview.hWnd, LV_FullRowSel, 1
    W.cyWndAction lvListview.hWnd, LV_GirdLine, 1
    
    sA = Split(sCaptionList, ";")
    sB = Split(sWidthList, ";")
    If UBound(sA) = -1 Then
        j = UBound(sB)
        
    Else
        j = UBound(sA)
        
    End If
    
    On Error Resume Next
    
    If lvListview.ColumnHeaders.Count = 0 Then
        For i = 0 To j
            '设定每格的Title,及表格宽
            lvListview.ColumnHeaders.Add , , sA(i), sB(i)
        
        Next
    
    Else
        For i = 0 To j
            '设定每格的Title,及表格宽
            lvListview.ColumnHeaders(i + 1).Text = sA(i)
            lvListview.ColumnHeaders(i + 1).Width = sB(i)
        
        Next
    
    End If

End Sub

'函数名称:cyGetHwndFromPID
'作用：根据Pid获得Hwnd
'===========================================================
'整理时间：2007-11-03 06:39
'  参数：
'     Dim W As New formClass
'     Debug.print W.cyGetHwndFromPID (2524)

'===========================================================
Function cyGetHwndFromPID(ByVal Pid As Long) As Long
    modWnd.PidhWnd = 0
    Call EnumWindows(AddressOf EnumWindowsProc, Pid)
    cyGetHwndFromPID = modWnd.PidhWnd
    
End Function


