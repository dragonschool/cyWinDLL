Attribute VB_Name = "modMenu"
'**************************************************************************************************************
'* 本模块配合 cyMenu 菜单类模块
'*
'* 版权: LPP软件工作室
'* 作者: 卢培培(goodname008)
'* (******* 复制请保留以上信息 *******)
'**************************************************************************************************************

Option Explicit

' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- API 函数声明 -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ByRef ColorRef As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal HIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- API 常量声明 -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Public Const GWL_WNDPROC = (-4)                     ' SetWindowLong 设置窗口函数入口地址
Public Const SM_CYMENU = 15                         ' GetSystemMetrics 获得系统菜单项高度

Public Const WM_COMMAND = &H111                     ' 消息: 单击菜单项
Public Const WM_DRAWITEM = &H2B                     ' 消息: 绘制菜单项
Public Const WM_EXITMENULOOP = &H212                ' 消息: 退出菜单消息循环
Public Const WM_MEASUREITEM = &H2C                  ' 消息: 处理菜单高度和宽度
Public Const WM_MENUSELECT = &H11F                  ' 消息: 选择菜单项
Public Const WM_MENUCHAR = &H120                    ' 消息: 使用快捷键选择菜单

' ODT
Public Const ODT_MENU = 1                           ' 菜单
Public Const ODT_LISTBOX = 2                        ' 列表框
Public Const ODT_COMBOBOX = 3                       ' 组合框
Public Const ODT_BUTTON = 4                         ' 按钮

' ODS
Public Const ODS_SELECTED = &H1                     ' 菜单被选择
Public Const ODS_GRAYED = &H2                       ' 灰色字
Public Const ODS_DISABLED = &H4                     ' 禁用
Public Const ODS_CHECKED = &H8                      ' 选中
Public Const ODS_FOCUS = &H10                       ' 聚焦

' diFlags to DrawIconEx
Public Const DI_MASK = &H1                          ' 绘图时使用图标的MASK部分 (如单独使用, 可获得图标的掩模)
Public Const DI_IMAGE = &H2                         ' 绘图时使用图标的XOR部分 (即图标没有透明区域)
Public Const DI_NORMAL = DI_MASK Or DI_IMAGE        ' 用常规方式绘图 (合并 DI_IMAGE 和 DI_MASK)

' nBkMode to SetBkMode
Public Const TRANSPARENT = 1                        ' 透明处理, 即不作上述填充
Public Const OPAQUE = 2                             ' 用当前的背景色填充虚线画笔、阴影刷子以及字符的空隙
Public Const NEWTRANSPARENT = 3                     ' 在有颜色的菜单上画透明文字


' MF 菜单相关常数
Public Const MF_BYCOMMAND = &H0&                    ' 菜单条目由菜单的命令ID指定
Public Const MF_BYPOSITION = &H400&                 ' 菜单条目由条目在菜单中的位置决定 (零代表菜单中的第一个条目)

Public Const MF_CHECKED = &H8&                      ' 检查指定的菜单条目 (不能与VB的Checked属性兼容)
Public Const MF_DISABLED = &H2&                     ' 禁止指定的菜单条目 (不与VB的Enabled属性兼容)
Public Const MF_ENABLED = &H0&                      ' 允许指定的菜单条目 (不与VB的Enabled属性兼容)
Public Const MF_GRAYED = &H1&                       ' 禁止指定的菜单条目, 并用浅灰色描述它. (不与VB的Enabled属性兼容)
Public Const MF_HILITE = &H80&
Public Const MF_SEPARATOR = &H800&                  ' 在指定的条目处显示一条分隔线
Public Const MF_STRING = &H0&                       ' 在指定的条目处放置一个字串 (不与VB的Caption属性兼容)
Public Const MF_UNCHECKED = &H0&                    ' 检查指定的条目 (不能与VB的Checked属性兼容)
Public Const MF_UNHILITE = &H0&

Public Const MF_BITMAP = &H4&                       ' 菜单条目是一幅位图. 一旦设入菜单, 这幅位图就绝对不能删除, 所以不应该使用由VB的Image属性返回的值.
Public Const MF_OWNERDRAW = &H100&                  ' 创建一个物主绘图菜单 (由您设计的程序负责描绘每个菜单条目)
Public Const MF_USECHECKBITMAPS = &H200&

Public Const MF_MENUBARBREAK = &H20&                ' 在弹出式菜单中, 将指定的条目放置于一个新列, 并用一条垂直线分隔不同的列.
Public Const MF_MENUBREAK = &H40&                   ' 在弹出式菜单中, 将指定的条目放置于一个新列. 在顶级菜单中, 将条目放置到一个新行.

Public Const MF_POPUP = &H10&                       ' 将一个弹出式菜单置于指定的条目, 可用于创建子菜单及弹出式菜单.
Public Const MF_HELP = &H4000&

Public Const MF_DEFAULT = &H1000
Public Const MF_RIGHTJUSTIFY = &H4000

' fMask To InsertMenuItem                           ' 指定 MENUITEMINFO 中哪些成员有效
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_CHECKMARKS = &H8
Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA = &H20
Public Const MIIM_STRING = &H40
Public Const MIIM_BITMAP = &H80
Public Const MIIM_FTYPE = &H100

' fType To InsertMenuItem                           ' MENUITEMINFO 中菜单项类型
Public Const MFT_BITMAP = &H4&
Public Const MFT_MENUBARBREAK = &H20&
Public Const MFT_MENUBREAK = &H40&
Public Const MFT_OWNERDRAW = &H100&
Public Const MFT_SEPARATOR = &H800&
Public Const MFT_STRING = &H0&

' fState to InsertMenuItem                          ' MENUITEMINFO 中菜单项状态
Public Const MFS_CHECKED = &H8&
Public Const MFS_DISABLED = &H2&
Public Const MFS_ENABLED = &H0&
Public Const MFS_GRAYED = &H1&
Public Const MFS_HILITE = &H80&
Public Const MFS_UNCHECKED = &H0&
Public Const MFS_UNHILITE = &H0&

' nFormat to DrawText
Public Const DT_LEFT = &H0                          ' 水平左对齐
Public Const DT_CENTER = &H1                        ' 水平居中对齐
Public Const DT_RIGHT = &H2                         ' 水平右对齐

Public Const DT_SINGLELINE = &H20                   ' 单行

Public Const DT_TOP = &H0                           ' 垂直上对齐 (仅单行时有效)
Public Const DT_VCENTER = &H4                       ' 垂直居中对齐 (仅单行时有效)
Public Const DT_BOTTOM = &H8                        ' 垂直下对齐 (仅单行时有效)

Public Const DT_CALCRECT = &H400                    ' 多行绘图时矩形的底边根据需要进行延展, 以便容下所有文字; 单行绘图时, 延展矩形的右侧, 不描绘文字, 由lpRect参数指定的矩形会载入计算出来的值.
Public Const DT_WORDBREAK = &H10                    ' 进行自动换行. 如用SetTextAlign函数设置了TA_UPDATECP标志, 这里的设置则无效.

Public Const DT_NOCLIP = &H100                      ' 描绘文字时不剪切到指定的矩形
Public Const DT_NOPREFIX = &H800                    ' 通常, 函数认为 & 字符表示应为下一个字符加上下划线, 该标志禁止这种行为.

Public Const DT_EXPANDTABS = &H40                   ' 描绘文字的时候, 对制表站进行扩展. 默认的制表站间距是8个字符. 但是, 可用DT_TABSTOP标志改变这项设定.
Public Const DT_TABSTOP = &H80                      ' 指定新的制表站间距, 采用这个整数的高 8 位.
Public Const DT_EXTERNALLEADING = &H200             ' 计算文本行高度的时候, 使用当前字体的外部间距属性.

' nIndex to GetSysColor  标准: 0--20
Public Const COLOR_ACTIVEBORDER = 10                ' 活动窗口的边框
Public Const COLOR_ACTIVECAPTION = 2                ' 活动窗口的标题
Public Const COLOR_APPWORKSPACE = 12                ' MDI桌面的背景
Public Const COLOR_BACKGROUND = 1                   ' Windows 桌面
Public Const COLOR_BTNFACE = 15                     ' 按钮
Public Const COLOR_BTNHIGHLIGHT = 20                ' 按钮的3D加亮区
Public Const COLOR_BTNSHADOW = 16                   ' 按钮的3D阴影
Public Const COLOR_BTNTEXT = 18                     ' 按钮文字
Public Const COLOR_CAPTIONTEXT = 9                  ' 窗口标题中的文字
Public Const COLOR_GRAYTEXT = 17                    ' 灰色文字; 如使用了抖动技术则为零
Public Const COLOR_HIGHLIGHT = 13                   ' 选定的项目背景
Public Const COLOR_HIGHLIGHTTEXT = 14               ' 选定的项目文字
Public Const COLOR_INACTIVEBORDER = 11              ' 不活动窗口的边框
Public Const COLOR_INACTIVECAPTION = 3              ' 不活动窗口的标题
Public Const COLOR_INACTIVECAPTIONTEXT = 19         ' 不活动窗口的文字
Public Const COLOR_MENU = 4                         ' 菜单
Public Const COLOR_MENUTEXT = 7                     ' 菜单文字
Public Const COLOR_SCROLLBAR = 0                    ' 滚动条
Public Const COLOR_WINDOW = 5                       ' 窗口背景
Public Const COLOR_WINDOWFRAME = 6                  ' 窗框
Public Const COLOR_WINDOWTEXT = 8                   ' 窗口文字

' un to DrawState
Public Const DST_COMPLEX = &H0                      ' 绘图在由lpDrawStateProc参数指定的回调函数期间执行, lParam和wParam会传递给回调事件.
Public Const DST_TEXT = &H1                         ' lParam代表文字的地址(可使用一个字串别名),wParam代表字串的长度.
Public Const DST_PREFIXTEXT = &H2                   ' 与DST_TEXT类似, 只是 & 字符指出为下各字符加上下划线.
Public Const DST_ICON = &H3                         ' lParam包括图标的句柄
Public Const DST_BITMAP = &H4                       ' lParam包括位图的句柄
Public Const DSS_NORMAL = &H0                       ' 普通图像
Public Const DSS_UNION = &H10                       ' 图像进行抖动处理
Public Const DSS_DISABLED = &H20                    ' 图象具有浮雕效果
Public Const DSS_MONO = &H80                        ' 用hBrush描绘图像
Public Const DSS_RIGHT = &H8000                     ' 无任何作用

' edge to DrawEdge
Public Const BDR_RAISEDOUTER = &H1                  ' 外层凸
Public Const BDR_SUNKENOUTER = &H2                  ' 外层凹
Public Const BDR_RAISEDINNER = &H4                  ' 内层凸
Public Const BDR_SUNKENINNER = &H8                  ' 内层凹
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA
Public Const BDR_INNER = &HC
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

' grfFlags to DrawEdge
Public Const BF_LEFT = &H1                          ' 左边缘
Public Const BF_TOP = &H2                           ' 上边缘
Public Const BF_RIGHT = &H4                         ' 右边缘
Public Const BF_BOTTOM = &H8                        ' 下边缘
Public Const BF_DIAGONAL = &H10                     ' 对角线
Public Const BF_MIDDLE = &H800                      ' 填充矩形内部
Public Const BF_SOFT = &H1000                       ' MSDN: Soft buttons instead of tiles.
Public Const BF_ADJUST = &H2000                     ' 调整矩形, 预留客户区
Public Const BF_FLAT = &H4000                       ' 平面边缘
Public Const BF_MONO = &H8000                       ' 一维边缘

Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

' nPenStyle to CreatePen
Public Const PS_DASH = 1                            ' 画笔类型:虚线 (nWidth必须是1)         -------
Public Const PS_DASHDOT = 3                         ' 画笔类型:点划线 (nWidth必须是1)       _._._._
Public Const PS_DASHDOTDOT = 4                      ' 画笔类型:点-点-划线 (nWidth必须是1)   _.._.._
Public Const PS_DOT = 2                             ' 画笔类型:点线 (nWidth必须是1)         .......
Public Const PS_SOLID = 0                           ' 画笔类型:实线                         _______


' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- API 类型声明 -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-


Public Enum DrawStateStyle
    dssNormal = 0               '正常状态
    dssDisabled = &H20          '无效状态
    dssShadow = &H80            '绘制阴影
    dssSmooth = &H100           '绘制平滑图案
End Enum

'描述一个点的结构
Public Type POINTAPI
   X As Long
   Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    itemData As Long
End Type

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    itemData As Long
End Type

Public Type Size
    cx As Long
    cy As Long
End Type


' 自定义菜单项数据结构
Public Type MyMenuItemInfo
    itemKeyID As Long
    itemIcon As StdPicture
    itemAlias As String
    itemText As String
    itemType As MenuItemType
    itemState As MenuItemState
    itemhSubMenu As Long            '子菜单句柄
    itemShutCutKey As String        '快捷键字母
End Type

' 菜单相关结构
Private MeasureInfo As MEASUREITEMSTRUCT

Public hMenu As Long
Public hWnd As Long

Public preMenuWndProc As Long
Public MyItemInfo() As MyMenuItemInfo

' 菜单类属性
Public BarWidth As Long                             ' 菜单附加条宽度
Public BarStyle As MenuLeftBarStyle                 ' 菜单附加条风格
Public BarImage As StdPicture                       ' 菜单附加条图像
Public BarStartColor As Long                        ' 菜单附加条过渡色起始颜色
Public BarEndColor As Long                          ' 菜单附加条过渡色终止颜色
Public SelectScope As MenuItemSelectScope           ' 菜单项高亮条的范围
Public TextEnabledColor As Long                     ' 菜单项可用时文字颜色
Public TextDisabledColor As Long                    ' 菜单项不可用时文字颜色
Public TextSelectColor As Long                      ' 菜单项选中时文字颜色
Public IconStyle As MenuItemIconStyle               ' 菜单项图标风格
Public EdgeStyle As MenuItemSelectEdgeStyle         ' 菜单项边框风格
Public EdgeColor As Long                            ' 菜单项边框颜色
Public FillStyle As MenuItemSelectFillStyle         ' 菜单项背景填充风格
Public FillStartColor As Long                       ' 菜单项过渡色起始颜色
Public FillEndColor As Long                         ' 菜单项过渡色终止颜色
Public BkColor As Long                              ' 菜单背景颜色
Public SepStyle As MenuSeparatorStyle               ' 菜单分隔条风格
Public SepColor As Long                             ' 菜单分隔条颜色
Public MenuStyle As MenuUserStyle                   ' 菜单总体风格

'保存热键对象句柄
Public objMenu As Long
Public preMenuProc As Long


' 拦截菜单消息 (frmMenu 窗口入口函数)
Function MenuWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
        Case WM_COMMAND                                                 ' 单击菜单项
'            If MyItemInfo(wParam).itemType = MIT_CHECKBOX Then
'                If MyItemInfo(wParam).itemState = MIS_CHECKED Then
'                    MyItemInfo(wParam).itemState = MIS_UNCHECKED
'
'                Else
'                    MyItemInfo(wParam).itemState = MIS_CHECKED
'
'                End If
'            End If
            
            MenuItemSelected wParam
        Case WM_EXITMENULOOP                                            ' 退出菜单消息循环(保留)
            ptrMenu.FireClose
        
        Case WM_MENUCHAR
            Dim lngRetval As Long
            If OnMenuChar(wParam, lngRetval) Then
                MenuWndProc = lngRetval
            
            End If
        
        Case WM_MEASUREITEM                                             ' 处理菜单项高度和宽度
            MeasureItem hWnd, lParam
            
        Case WM_DRAWITEM                                                ' 绘制菜单项
            If OnDrawItem(lParam) Then
                MenuWndProc = 1
            
            End If
       
    End Select
    MenuWndProc = CallWindowProc(preMenuWndProc, hWnd, Msg, wParam, lParam)
    
End Function

' 处理菜单高度和宽度
Private Sub MeasureItem(ByVal hWnd As Long, ByVal lParam As Long)
    Dim hdc As Long
    hdc = GetDC(hWnd)
    CopyMemory MeasureInfo, ByVal lParam, Len(MeasureInfo)
    If MeasureInfo.CtlType And ODT_MENU Then
        MeasureInfo.itemWidth = lstrlen(MyItemInfo(MeasureInfo.itemID).itemText) * (GetSystemMetrics(SM_CYMENU) / 2.5) + BarWidth + 15
        If MyItemInfo(MeasureInfo.itemID).itemType <> MIT_SEPARATOR Then
            MeasureInfo.itemHeight = GetSystemMetrics(SM_CYMENU) + 2
        Else
            MeasureInfo.itemHeight = 6
        End If
    End If
    CopyMemory ByVal lParam, MeasureInfo, Len(MeasureInfo)
    ReleaseDC hWnd, hdc
End Sub

' 菜单项事件响应(单击菜单项)
Private Sub MenuItemSelected(ByVal itemID As Long)
On Error GoTo Err
    If MyItemInfo(itemID).itemShutCutKey = "" Then
    '没有快捷键
        ptrMenu.FireEvent MyItemInfo(itemID).itemKeyID, MyItemInfo(itemID).itemText
        
    Else
    '有快捷键
        ptrMenu.FireEvent MyItemInfo(itemID).itemKeyID, Left(MyItemInfo(itemID).itemText, Len(MyItemInfo(itemID).itemText) - 4)
        
    End If
Err:

End Sub

' 菜单项事件响应(选择菜单项)
Private Sub MenuItemSelecting(ByVal itemID As Long)
    
End Sub

Private Function OnMenuChar(wParam As Long, lngRetval As Long) As Boolean
    Dim sMenuChar As String
    Dim i As Long
    
    sMenuChar = UCase(Chr((wParam And &HFFFF&)))
    
    For i = 1 To UBound(MyItemInfo) + 1
        If MyItemInfo(i).itemShutCutKey = sMenuChar Then
            lngRetval = &H20000 + i
            OnMenuChar = True
            Exit Function
            
        End If
        
    Next

End Function

Private Function ptrMenu() As cyMenu
    Dim Menu As cyMenu
    CopyMemory Menu, objMenu, 4&
    Set ptrMenu = Menu
    CopyMemory Menu, 0&, 4&
    
End Function

Private Function OnDrawItem(ByVal StructPtr As Long) As Boolean
    Dim udtStruct As DRAWITEMSTRUCT
    Dim strText As String
    Dim udtRect As RECT
    Dim lngBrush As Long
    Dim hBrush As Long
    Dim hOldBrush As Long
    Dim hPen As Long
    Dim hOldPen As Long
    Dim arrPoint(1 To 2) As POINTAPI
    Dim blnHighlight As Boolean
    Dim lngOldFont As Long
    
    CopyMemory udtStruct, ByVal StructPtr, Len(udtStruct)

    With udtStruct
        '获取要绘制的对象是菜单
        If .CtlType = ODT_MENU Then
        
            '确定当前是不是高亮显示
            blnHighlight = .itemState And ODS_SELECTED
            
            Select Case MyItemInfo(.itemID).itemType
                Case &H800            '分割线类型
                    '绘制图标区域
                    udtRect.Left = .rcItem.Left
                    udtRect.Top = .rcItem.Top
                    udtRect.Right = udtRect.Left + 21
                    udtRect.Bottom = .rcItem.Bottom + 5
                    lngBrush = CreateSolidBrush(&HD1D8DB)
                    FillRect .hdc, udtRect, lngBrush
                    DeleteObject lngBrush

                    '绘制分隔线
                    hPen = CreatePen(PS_SOLID, 1, &HA6A6A6)
                    hOldPen = SelectObject(.hdc, hPen)
                    arrPoint(1).X = udtRect.Right + 8
                    arrPoint(1).Y = udtRect.Top + 1
                    arrPoint(2).X = .rcItem.Right + 1
                    arrPoint(2).Y = arrPoint(1).Y
                    Polyline .hdc, arrPoint(1), 2
                    SelectObject .hdc, hOldPen
                    DeleteObject hPen
                Case Else
                    '绘制边框，底色
                    If blnHighlight And Not CBool(MyItemInfo(.itemID).itemState And MIS_DISABLED) Then   ' 当菜单项并被选中时
                        hPen = CreatePen(PS_SOLID, 1, &H6A240A)
                        hBrush = CreateSolidBrush(&HD2BDB6)
                        hOldPen = SelectObject(.hdc, hPen)
                        hOldBrush = SelectObject(.hdc, hBrush)
                        Rectangle .hdc, .rcItem.Left, .rcItem.Top, .rcItem.Right, .rcItem.Bottom
                        SelectObject .hdc, hOldPen
                        SelectObject .hdc, hOldBrush
                        DeleteObject hPen
                        DeleteObject hOldBrush

                    Else
                        udtRect.Left = .rcItem.Left
                        udtRect.Top = .rcItem.Top - 1
                        udtRect.Right = udtRect.Left + 21
                        udtRect.Bottom = .rcItem.Bottom + 0

                        hBrush = CreateSolidBrush(&HD1D8DB)
                        FillRect .hdc, udtRect, hBrush
                        DeleteObject hBrush

                        udtRect.Left = udtRect.Right
                        udtRect.Right = .rcItem.Right + 1 + 20

                        hBrush = CreateSolidBrush(&HF7F8F9)
                        FillRect .hdc, udtRect, hBrush
                        DeleteObject hBrush

                    End If
                    
                    '绘制图标
                    If MyItemInfo(.itemID).itemType = MIT_STRING Then
                        If MyItemInfo(.itemID).itemState = MIS_ENABLED Then
                        '菜单状态为可用
                        
                            If blnHighlight Then
                            '如果当前是选中状态，则图标往右下移２个像素
                                Call DrawPictureEx(.hdc, MyItemInfo(.itemID).itemIcon, .rcItem.Left + 2, .rcItem.Top + 2, 16, 16, dssNormal)
                            Else
                            '如果当前是非选中状态，则图标往右下移１个像素
                                Call DrawPictureEx(.hdc, MyItemInfo(.itemID).itemIcon, .rcItem.Left + 1, .rcItem.Top + 1, 16, 16, dssNormal)
                            End If
                            
                        Else
                        '不可用状态则图标变灰
                            Call DrawPictureEx(.hdc, MyItemInfo(.itemID).itemIcon, .rcItem.Left + 1, .rcItem.Top + 1, 16, 16, dssSmooth)
                            
                        End If

                    Else

                    End If

                    SetBkMode .hdc, TRANSPARENT
                    udtRect.Left = .rcItem.Left + 27
                    udtRect.Top = .rcItem.Top + 1
                    udtRect.Right = .rcItem.Right - 0
                    udtRect.Bottom = .rcItem.Bottom
                    
                    If MyItemInfo(.itemID).itemState = MIS_ENABLED Then
                    '菜单状态为可用
                        If blnHighlight Then
                            '当前是选中状态
                            SetTextColor .hdc, &H0
                        Else
                            SetTextColor .hdc, &H0
                             '当前是非选中状态
                       End If
                    Else
                        '不可用状态则文字变灰
                        SetTextColor .hdc, &HA6A6A6
                    End If
                    
                    strText = MyItemInfo(.itemID).itemText
                    lngOldFont = SelectObject(.hdc, 0)
                    DrawText .hdc, strText, lstrlen(strText), udtRect, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
                    SelectObject .hdc, lngOldFont
                    
            End Select
                
        End If
    End With
    OnDrawItem = True
    
End Function

'自定义函数  根据指针获取一个对象
Public Function GetMenuItemFromPtr(ByVal lngPtr As Long) As cyMenu
   Dim oTemp As Object
   '实现技术：通过API函数CopyMemory,直接将对象长整型的值当作指针来用
   CopyMemory oTemp, lngPtr, 4
   Set GetMenuItemFromPtr = oTemp
   CopyMemory oTemp, 0&, 4
   
End Function

Public Sub DrawPictureEx(ByVal HdcDest As Long, ByVal Source As StdPicture, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Style As DrawStateStyle)
    Dim lngFlags As Long
    Dim hBrush As Long
    Dim hMemDc As Long
    Dim DrawX As Long
    Dim DrawY As Long
    Dim lngBitmap As Long
    Dim hBitmap As Long
    Dim lngColor As Long
    
    '如果源为空，放弃退出
    If Source Is Nothing Then Exit Sub
    
    '如果待绘制类型为Smooth
    If Style = dssSmooth Then
        hMemDc = CreateCompatibleDC(HdcDest)
        Select Case Source.type
            Case vbPicTypeBitmap
                lngBitmap = SelectObject(hMemDc, Source.Handle)
            Case vbPicTypeIcon
                hBitmap = CreateCompatibleBitmap(HdcDest, Width, Height)
                lngBitmap = SelectObject(hMemDc, hBitmap)
                DrawState hMemDc, 0, 0, Source.Handle, 0, 0, 0, Width, Height, DST_ICON
            Case Else
                DeleteDC hMemDc
                Exit Sub
        End Select
        
        For DrawX = 0 To Width - 1
            For DrawY = 0 To Height - 1
                lngColor = GetPixel(hMemDc, DrawX, DrawY)
                If lngColor <> 0 Then SetPixel HdcDest, DrawX + X, DrawY + Y, SmoothColor(lngColor)
            Next DrawY
        Next DrawX
    
        SelectObject hMemDc, lngBitmap
        DeleteObject hBitmap
        DeleteDC hMemDc
    Else
        Select Case Source.type
            Case vbPicTypeBitmap
                lngFlags = DST_BITMAP
            Case vbPicTypeIcon
                lngFlags = DST_ICON
            Case Else
                lngFlags = DST_ICON
'                lngFlags = DST_COMPLEX
        End Select
        
        Select Case Style
            Case dssShadow
                DrawState HdcDest, hBrush, 0, Source.Handle, 0, X + 1, Y + 1, Width, Height, dssShadow Or lngFlags
                DeleteObject hBrush
                DrawState HdcDest, 0, 0, Source.Handle, 0, X, Y, Width, Height, lngFlags
                
            Case dssNormal
                DrawState HdcDest, 0, 0, Source.Handle, 0, X, Y, Width, Height, lngFlags
                
            Case dssDisabled
                hBrush = CreateSolidBrush(&H808080)
                DrawState HdcDest, 0, 0, Source.Handle, 0, X, Y, Width, Height, lngFlags Or dssShadow
                DeleteObject hBrush
                
        End Select
    End If
End Sub


Private Function SmoothColor(ByVal Color As Long) As Long
    Dim R As Long
    Dim G As Long
    Dim b As Long
    Dim lngTempColor As Long
    
    ColorToRGB Color, R, G, b
    
    R = R + 76 - Int((R + 32) / 64) * 19
    G = G + 76 - Int((G + 32) / 64) * 19
    b = b + 76 - Int((b + 32) / 64) * 19
    
    lngTempColor = TranslateColor(RGB(R, G, b))
    ColorToRGB lngTempColor, R, G, b
    SmoothColor = RGB(R, G, b)
End Function


Private Sub ColorToRGB(ByVal lngColor As Long, Red As Long, Green As Long, Blue As Long)
   Dim lngHalf As Long

   lngHalf = CLng(lngColor \ 256)
   Blue = Int(lngHalf \ 256)
   Green = lngHalf - Blue * 256
   Red = lngColor - lngHalf * 256
End Sub


Public Function TranslateColor(ByVal Color As OLE_COLOR, Optional hpal As Long = 0) As Long
    Const CLR_INVALID = -1
    If OleTranslateColor(Color, hpal, TranslateColor) Then TranslateColor = CLR_INVALID
End Function

Public Function MaskColor(ByVal lngScale As Long, ByVal lngColor As Long) As Long
    Dim R As Long
    Dim G As Long
    Dim b As Long

    ColorToRGB lngColor, R, G, b

    R = R - Int(R * lngScale / 255)
    G = G - Int(G * lngScale / 255)
    b = b - Int(b * lngScale / 255)

    MaskColor = RGB(R, G, b)
End Function
