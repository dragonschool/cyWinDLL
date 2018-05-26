VERSION 5.00
Begin VB.Form frmDebugHwnd 
   Caption         =   "句柄跟踪器"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   Icon            =   "frmDebugHwnd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2134.503
   ScaleMode       =   0  'User
   ScaleWidth      =   10184.74
   Begin VB.CommandButton cmd 
      Caption         =   ".."
      Height          =   300
      Index           =   0
      Left            =   6300
      TabIndex        =   29
      ToolTipText     =   "显示程序所在目录"
      Top             =   540
      Width           =   435
   End
   Begin VB.TextBox txtFullPath 
      BackColor       =   &H00C0E0FF&
      Height          =   270
      Left            =   1140
      TabIndex        =   28
      Top             =   555
      Width           =   5115
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "&L锁定"
      Height          =   540
      Left            =   9405
      TabIndex        =   21
      Top             =   1155
      Width           =   705
   End
   Begin VB.PictureBox lblColor 
      Height          =   270
      Left            =   9405
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   810
      Width           =   255
   End
   Begin VB.TextBox txtSetTxt 
      Height          =   270
      Index           =   3
      Left            =   3495
      TabIndex        =   15
      Top             =   1410
      Width           =   3255
   End
   Begin VB.TextBox txtClassName 
      Height          =   270
      Index           =   3
      Left            =   1635
      TabIndex        =   14
      Top             =   1410
      Width           =   1800
   End
   Begin VB.TextBox txtSetTxt 
      Height          =   270
      Index           =   2
      Left            =   3255
      TabIndex        =   13
      Top             =   1125
      Width           =   3255
   End
   Begin VB.TextBox txtClassName 
      Height          =   270
      Index           =   2
      Left            =   1395
      TabIndex        =   12
      Top             =   1125
      Width           =   1800
   End
   Begin VB.TextBox txtSetTxt 
      Height          =   270
      Index           =   1
      Left            =   3000
      TabIndex        =   11
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtClassName 
      Height          =   270
      Index           =   1
      Left            =   1140
      TabIndex        =   10
      Top             =   840
      Width           =   1800
   End
   Begin VB.TextBox txtClassName 
      Height          =   270
      Index           =   0
      Left            =   990
      TabIndex        =   9
      Top             =   270
      Width           =   1800
   End
   Begin VB.TextBox txtSetTxt 
      Height          =   270
      Index           =   0
      Left            =   2850
      TabIndex        =   8
      Top             =   270
      Width           =   2535
   End
   Begin VB.TextBox lblCurrentHwnd 
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   7
      Top             =   270
      Width           =   810
   End
   Begin VB.TextBox lblCurrentHwnd 
      Height          =   270
      Index           =   1
      Left            =   255
      TabIndex        =   6
      Top             =   840
      Width           =   810
   End
   Begin VB.TextBox lblCurrentHwnd 
      Height          =   270
      Index           =   2
      Left            =   510
      TabIndex        =   5
      Top             =   1125
      Width           =   810
   End
   Begin VB.TextBox lblCurrentHwnd 
      Height          =   270
      Index           =   3
      Left            =   735
      TabIndex        =   4
      Top             =   1410
      Width           =   810
   End
   Begin VB.TextBox lblNo 
      Height          =   270
      Left            =   5430
      TabIndex        =   3
      ToolTipText     =   "当前句柄与其上级的层次关系"
      Top             =   270
      Width           =   1305
   End
   Begin VB.ComboBox cboScaleNo 
      Height          =   300
      Left            =   9405
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "放大的倍数"
      Top             =   495
      Width           =   675
   End
   Begin VB.CheckBox chkScale 
      Caption         =   "放大显示"
      Height          =   465
      Left            =   9390
      TabIndex        =   1
      ToolTipText     =   "是否放大显示"
      Top             =   30
      Width           =   660
   End
   Begin VB.PictureBox Picture1 
      Height          =   1035
      Left            =   6795
      ScaleHeight     =   975
      ScaleWidth      =   2430
      TabIndex        =   0
      Top             =   45
      Width           =   2490
   End
   Begin VB.Label lbl 
      Caption         =   "PID："
      Height          =   255
      Index           =   4
      Left            =   135
      TabIndex        =   30
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Pos:"
      Height          =   195
      Left            =   7200
      TabIndex        =   27
      Top             =   1290
      Width           =   495
   End
   Begin VB.Label lblPosition 
      Height          =   195
      Left            =   7800
      TabIndex        =   26
      Top             =   1290
      Width           =   1155
   End
   Begin VB.Label lblSize 
      Height          =   195
      Left            =   7800
      TabIndex        =   25
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Size:"
      Height          =   195
      Left            =   7200
      TabIndex        =   24
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblPos 
      Height          =   195
      Left            =   7800
      TabIndex        =   23
      Top             =   1500
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "MousePos:"
      Height          =   195
      Left            =   6855
      TabIndex        =   22
      Top             =   1500
      Width           =   840
   End
   Begin VB.Label lbl 
      Caption         =   "句柄位置："
      Height          =   255
      Index           =   3
      Left            =   5385
      TabIndex        =   20
      Top             =   75
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "标题/内容："
      Height          =   255
      Index           =   2
      Left            =   2865
      TabIndex        =   19
      Top             =   75
      Width           =   1065
   End
   Begin VB.Label lbl 
      Caption         =   "类名："
      Height          =   255
      Index           =   1
      Left            =   1065
      TabIndex        =   18
      Top             =   75
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "句柄："
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   17
      Top             =   75
      Width           =   615
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuSize 
         Caption         =   "窗口尺寸"
      End
      Begin VB.Menu mnuFindClass 
         Caption         =   "查找类及文本"
      End
      Begin VB.Menu mnuSetTxt 
         Caption         =   "设置文本"
      End
      Begin VB.Menu mnuDisableInput 
         Caption         =   "禁止输入"
      End
   End
End
Attribute VB_Name = "frmDebugHwnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

Dim iScaleX As Integer
Dim iscaleY As Integer

Dim WithEvents HotKey As cyHotKeyEx
Attribute HotKey.VB_VarHelpID = -1
Dim WithEvents tmrMouseMove As cyTimers
Attribute tmrMouseMove.VB_VarHelpID = -1
Dim WithEvents tmrCheckWndActive As cyTimers    '每秒钟检测窗口是否活跃，如果不活跃则暂停时钟
Attribute tmrCheckWndActive.VB_VarHelpID = -1

'放大显示
Private Sub chkScale_Click()
End Sub

Private Sub cmd_Click(Index As Integer)
    If Index = 0 Then
    '打开文件所在的路径
        Dim F As New cyFileEx
        F.cyOpenFolder F.cyCutFileName(txtFullPath, CutPath), True
        Set F = Nothing
        
    End If
    
End Sub

'锁定键盘
Private Sub cmdLock_Click()
    Dim W As New cyWndEx
    
    If cmdLock.Caption = "&L锁定" Then
    '如果当前是非锁定状态则设置为锁定状态
        cmdLock.Caption = "&L解锁"
        
        '鼠标时钟
        Set tmrMouseMove = Nothing
        '移除快捷键
        Set HotKey = Nothing
        
On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText lblCurrentHwnd(Index)
        
        W.cyWndAction lblCurrentHwnd(0), Wnd_Flash
        
    Else
    '如果当前是锁定状态则设置为非锁定状态
        cmdLock.Caption = "&L锁定"
        
        Set tmrMouseMove = New cyTimers
        '鼠标时钟
        tmrMouseMove.cyTimerStart 0.1
        '设置快捷键
        Set HotKey = New cyHotKeyEx
        HotKey.cySetHotKeyEx 100009, Me.hWnd, , , True, vbKeyL
        '设置位置检查时钟
        Set tmrCheckWndActive = New cyTimers
        tmrCheckWndActive.cySecondClock
        
        W.cyWndAction lblCurrentHwnd(0), Wnd_ShowFrame, 0
        
    End If
    
End Sub

Private Sub Form_Load()

Dim W As New cyWndEx

'打开时钟
Set tmrMouseMove = New cyTimers
tmrMouseMove.cyTimerStart 0.1
 
Set tmrCheckWndActive = New cyTimers
tmrCheckWndActive.cySecondClock

'设置快捷键
Set HotKey = New cyHotKeyEx
HotKey.cySetHotKeyEx 100009, Me.hWnd, , , True, vbKeyL

Me.Top = 0
Me.Height = 2295

'使窗口处于最顶层
W.cyWndAction Me.hWnd, Wnd_TOPMOST, 1
    
'设置放大率
cboScaleNo.AddItem "2"
cboScaleNo.AddItem "3"
cboScaleNo.AddItem "4"
cboScaleNo.AddItem "5"
cboScaleNo.AddItem "6"
cboScaleNo.ListIndex = 2

End Sub

Private Sub Form_Resize()
'On Error GoTo Err
'    If Me.WindowState = 1 Then
''        tmrMouseMove.cyTimerStop
''        Set HotKey = Nothing
'        cmdLock.Caption = "&L锁定"
'        cmdLock_Click
'    ElseIf Me.WindowState = 0 Then
''        '鼠标时钟
''        tmrMouseMove.cyTimerStart 0.1
''        '设置快捷键
''        HotKey.cySetHotKeyEx 100009, Me.hWnd, , , True, vbKeyL
'        cmdLock.Caption = "&L解锁"
'        cmdLock_Click
'    End If
'
'    Me.Width = 10245
'    Me.Height = 2295
'
'Err:

End Sub

Private Sub Form_Unload(Cancel As Integer)

    '移除快捷键
    Set HotKey = Nothing
    '移除时钟
    Set tmrMouseMove = Nothing
    Set tmrCheckWndActive = Nothing
    
End Sub


Private Sub HotKey_cyHotKeyEvent(ByVal IDHotKey As Long)

End Sub

Private Sub HotKey_cyHotKeyEventEx(ByVal IDHotKey As Long)
    '锁定
    cmdLock_Click

End Sub

'复制当前句柄
Private Sub lblCurrentHwnd_Click(Index As Integer)
On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText lblCurrentHwnd(Index)

End Sub

Private Sub lblNo_Click()
On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText lblNo

End Sub

Sub DisplayHwnd(ByVal iWndHwnd As Long)
    '枚举所有顶级窗口到数组
    Dim iHwnd() As Long
    iHwnd = W.cyGetHwndArraryFromTopWnd
    Dim i As Long
    For i = 0 To UBound(iHwnd)
        If iHwnd(i) = iWndHwnd Then
        '在窗口列表内
            lblCurrentHwnd(0) = iWndHwnd
            txtClassName(0) = W.cyGetClassFromHwnd(iWndHwnd)
            txtSetTxt(0) = W.cyGetCaptionFromHwnd(iWndHwnd)
            
            lblCurrentHwnd(1) = 0
            txtClassName(1) = ""
            txtSetTxt(1) = ""
            
            lblCurrentHwnd(2) = 0
            txtClassName(2) = ""
            txtSetTxt(2) = ""
            
            lblCurrentHwnd(3) = 0
            txtClassName(3) = ""
            txtSetTxt(3) = ""
        End If
    Next
    '没有找到句柄,则退出
    Exit Sub
End Sub

Private Sub tmrCheckWndActive_TimerEvent()
    If Me.Left < -10000 Then
    '不在当前，则停止捕捉
        cmdLock.Caption = "&L锁定"
    
        '点击锁定按钮
        cmdLock_Click
        Set tmrCheckWndActive = Nothing
    
    End If
    
End Sub

Private Sub tmrMouseMove_TimerEvent()
    Dim S As New cySystemEx
    Dim W As New cyWndEx

'实时取得鼠标的光标位置
    Dim X As Long
    Dim Y As Long
    S.cyMouseAction CursorPosGet, X, Y
    lblPos = X & " x " & Y

'如果窗口是正常化就避开鼠标
If Me.WindowState = 0 Then
    If Y > IIf(Screen.Width / Screen.TwipsPerPixelX = 800, 550, 718) Then
        Me.Top = 0
    ElseIf Y < 50 Then
        Me.Top = IIf(Screen.Width / Screen.TwipsPerPixelX = 800, 6360, 8880)
    End If
End If

Dim iCurrentHwnd
Dim iChildHwnd As Long
Dim curPos As POINTAPI
Dim CurRect As RECT
Dim bOverWnd As Boolean

Const GW_CHILD = 5
Const GW_HWNDNEXT = 2




GetCursorPos curPos                                     '返回鼠标当前位置
iCurrentHwnd = WindowFromPoint(curPos.X, curPos.Y)      '取得鼠标指针处窗口句柄
GetWindowRect iCurrentHwnd, CurRect                     '返回当前句柄的范围

If Not PtInRect(CurRect, curPos.X, curPos.Y) Then

    iChildHwnd = GetWindow(iCurrentHwnd, GW_CHILD)
    
    Do While (iChildHwnd)
        GetWindowRect iChildHwnd, CurRect
        If PtInRect(CurRect, curPos.X, curPos.Y) Then
            bOverWnd = True
            Exit Do
        Else
            iChildHwnd = GetWindow(iChildHwnd, GW_HWNDNEXT)
        End If
        
    Loop
    
    If bOverWnd = True Then
        bOverWnd = False
        iCurrentHwnd = iChildHwnd
        
    End If

End If



'For j = 0 To 100
'    k = W.cyGetSubObjHwndFromHwnd(iCurrentHwnd, j)
'    If k = 0 Then
'    '为０时退出，则检查该句柄
'        Exit For
'
'    Else
'        If W.cyWndAction(k, Wnd_IsCursorOver) Then
'            iCurrentHwnd = k
'            Exit For
'
'        End If
'
'    End If
'
'Next
'
'''    UnHwnd = WindowFromPoint(pnt.X, pnt.Y)     '取得鼠标指针处窗口句柄
'''    grayHwnd = GetWindow(UnHwnd, GW_CHILD)
'''
'''    Do While (grayHwnd)
'''        GetWindowRect grayHwnd, tempRc
'''        If PtInRect(tempRc, pnt.X, pnt.Y) Then
'''            FindIt = True
'''            Exit Do
'''        Else
'''            grayHwnd = GetWindow(grayHwnd, GW_HWNDNEXT)
'''        End If
'''    Loop
'''    If FindIt = True Then
'''        FindIt = False
'''        SnapHwnd = grayHwnd
'''    Else
'''        SnapHwnd = UnHwnd
'''    End If



lblCurrentHwnd(0) = iCurrentHwnd

lblCurrentHwnd(1) = W.cyGetParentHwnd(iCurrentHwnd, FatherOnly)
If lblCurrentHwnd(1) <> "" Then '有句柄
    lblCurrentHwnd(2) = W.cyGetParentHwnd(CLng(lblCurrentHwnd(1)), FatherOnly)
        
    If lblCurrentHwnd(2) <> "" Then '有句柄
        lblCurrentHwnd(3) = W.cyGetParentHwnd(CLng(lblCurrentHwnd(2)), FatherOnly)
    End If
    
End If

'根据句柄返回类名
txtClassName(0) = W.cyGetClassName(iCurrentHwnd)
txtClassName(1) = W.cyGetClassName(CLng(lblCurrentHwnd(1)))
txtClassName(2) = W.cyGetClassName(CLng(lblCurrentHwnd(2)))
txtClassName(3) = W.cyGetClassName(CLng(lblCurrentHwnd(3)))

'根据句柄返回标题
txtSetTxt(0) = W.cyWndAction(iCurrentHwnd, Txt_GetPassWord)
txtSetTxt(1) = W.cyGetCaption(CLng(lblCurrentHwnd(1)))
txtSetTxt(2) = W.cyGetCaption(CLng(lblCurrentHwnd(2)))
txtSetTxt(3) = W.cyGetCaption(CLng(lblCurrentHwnd(3)))

'取得当前的位置
lblNo = W.cyGetSubObjPosList(CLng(iCurrentHwnd))

'取得当前对象的大小
lblSize = W.cyWndAction(CLng(iCurrentHwnd), Wnd_GetWindowWidth) & " x " & W.cyWndAction(CLng(iCurrentHwnd), Wnd_GetWindowHeight)
'取得当前对象的位置
lblPosition = W.cyWndAction(CLng(iCurrentHwnd), Wnd_GetWindowLeft) & " , " & W.cyWndAction(CLng(iCurrentHwnd), Wnd_GetWindowTop)
    
ihdc = GetWindowDC(0)

'取得屏幕的HDC
If chkScale.Value = vbChecked Then
    Const SRCCOPY = &HCC0020
    StretchBlt Picture1.hdc, 0, 0, CLng(cboScaleNo.Text) * 100, CLng(cboScaleNo.Text) * 100, ihdc, X - iScaleX, Y - iscaleY, 100, 100, SRCCOPY '4

End If

'取得当前点的颜色
lblColor.BackColor = GetPixel(ihdc, X, Y)
lblColor.ToolTipText = "HEX:" + CStr(Hex(lblColor.BackColor))

'取得窗口的程序名称
txtFullPath = S.cyGetAppNameFromHwnd(iCurrentHwnd)

'取得程序的PID
lbl(4) = "Pid：" & S.cyGetPidFromHwnd(iCurrentHwnd)
End Sub

'复制类名
Private Sub txtClassName_Click(Index As Integer)
On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText txtClassName(Index)

End Sub

'复制标题
Private Sub txtSetTxt_Click(Index As Integer)
On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText txtSetTxt(Index)

End Sub
