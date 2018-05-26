VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cyMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************************************
'* 本类模块是一个菜单类, 提供了各种样式的菜单的制作方案
'*
'* 版权: LPP软件工作室
'* 作者: 卢培培(goodname008)
'* (******* 复制请保留以上信息 *******)
'**************************************************************************************************************

Option Explicit

Public Event cyMenuShow()

Public Event cyMenuHide()

Public Event cyMenuClick(ByVal iMenuKeyID As Long, ByVal sMenuCaption As String)

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As Any) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum MenuUserStyle                                   ' 菜单总体风格
    STYLE_WINDOWS
    STYLE_XP
    STYLE_SHADE
    STYLE_3D
    STYLE_COLORFUL
End Enum

Public Enum MenuSeparatorStyle                              ' 菜单分隔条风格
    MSS_SOLID
    MSS_DASH
    MSS_DOT
    MSS_DASDOT
    MSS_DASHDOTDOT
    MSS_NONE
    MSS_DEFAULT
End Enum

Public Enum MenuItemSelectFillStyle                         ' 菜单项背景填充风格
    ISFS_NONE
    ISFS_SOLIDCOLOR
    ISFS_HORIZONTALCOLOR
    ISFS_VERTICALCOLOR
End Enum

Public Enum MenuItemSelectEdgeStyle                         ' 菜单项边框风格
    ISES_SOLID
    ISES_DASH
    ISES_DOT
    ISES_DASDOT
    ISES_DASHDOTDOT
    ISES_NONE
    ISES_SUNKEN
    ISES_RAISED
End Enum

Public Enum MenuItemIconStyle                               ' 菜单项图标风格
    IIS_NONE
    IIS_SUNKEN
    IIS_RAISED
    IIS_SHADOW
End Enum

Public Enum MenuItemSelectScope                             ' 菜单项高亮条的范围
    ISS_TEXT = &H1
    ISS_ICON_TEXT = &H2
    ISS_LEFTBAR_ICON_TEXT = &H4
End Enum

Public Enum MenuLeftBarStyle                                ' 菜单附加条风格
    LBS_NONE
    LBS_SOLIDCOLOR
    LBS_HORIZONTALCOLOR
    LBS_VERTICALCOLOR
    LBS_IMAGE
End Enum

Public Enum MenuItemType                                    ' 菜单项类型
    MIT_STRING = &H0
    MIT_CHECKBOX = &H200
    MIT_SEPARATOR = &H800
End Enum

Public Enum MenuItemState                                   ' 菜单项状态
    MIS_ENABLED = &H0
    MIS_DISABLED = &H2
    MIS_CHECKED = &H8
    MIS_UNCHECKED = &H0
End Enum

Public Enum PopupAlign                                      ' 菜单弹出对齐方式
    POPUP_LEFTALIGN = &H0&                                  ' 水平左对齐
    POPUP_CENTERALIGN = &H4&                                ' 水平居中对齐
    POPUP_RIGHTALIGN = &H8&                                 ' 水平右对齐
    POPUP_TOPALIGN = &H0&                                   ' 垂直上对齐
    POPUP_VCENTERALIGN = &H10&                              ' 垂直居中对齐
    POPUP_BOTTOMALIGN = &H20&                               ' 垂直下对齐
End Enum

Private Sub Class_Initialize()
    modMenu.objMenu = ObjPtr(Me)
    Erase MyItemInfo
    
End Sub

' 释放类
Private Sub Class_Terminate()
    SetWindowLong hWnd, GWL_WNDPROC, preMenuWndProc
    DestroyMenu hMenu
End Sub

' 插入菜单项并保存自定义菜单项数组, 设置Owner_Draw自绘菜单
Public Sub AddItem(ByVal itemKeyID As Long, ByVal itemText As String, Optional ByVal itemType As MenuItemType = MIT_STRING, Optional ByVal itemShutCutKey As String, Optional ByVal itemIcon As StdPicture, Optional ByVal itemState As MenuItemState, Optional ByVal itemHasSubMenu As Boolean, Optional ByVal ParentItemKeyID As Long)
    Static ID As Long, i As Long
    Dim ItemInfo As MENUITEMINFO
    
    itemShutCutKey = UCase(itemShutCutKey)
    itemText = IIf(itemShutCutKey = "", itemText, itemText & "(&" & itemShutCutKey & ")")
    
    ' 插入菜单项
    With ItemInfo
        .cbSize = LenB(ItemInfo)
        .fMask = MIIM_STRING Or MIIM_FTYPE Or MIIM_STATE Or MIIM_SUBMENU Or MIIM_ID Or MIIM_DATA
        .fType = itemType
        .fState = itemState
        .wid = ID
        .dwItemData = True
        .cch = lstrlen(itemText)
        .dwTypeData = itemText
        
        '如果有子菜单
        If itemHasSubMenu Then
            .hSubMenu = CreatePopupMenu()

        End If
        
    End With
    
    If ParentItemKeyID > 0 Then
    '其有父菜单
        For i = 0 To UBound(MyItemInfo)
            If MyItemInfo(i).itemKeyID = ParentItemKeyID Then GoTo FindParentID
        
        Next i
        Err.Raise -9999, "Menu", "没有找到该菜单的父ID."
        Exit Sub
FindParentID:

        InsertMenuItem MyItemInfo(i).itemhSubMenu, ID, False, ItemInfo
    Else
    '本身已是父菜单
        InsertMenuItem hMenu, ID, False, ItemInfo
    End If
    
    ' 将菜单项数据存入动态数组
    ReDim Preserve MyItemInfo(ID) As MyMenuItemInfo
    
    '检测是否有相同的KeyID
    For i = 0 To UBound(MyItemInfo)
        If MyItemInfo(i).itemKeyID = itemKeyID Then
            Class_Terminate
            Err.Raise -9999, "Menu", "两个菜单的KeyID存在冲突."
        End If
    Next i
    
    With MyItemInfo(ID)
        .itemKeyID = itemKeyID
        If (itemIcon Is Nothing) Then
            If itemState = MIS_CHECKED Then
            '如果状态是被选中状态，则显示缺省的勾号
                Set .itemIcon = frmIcon.Label1.DragIcon
            
            Else
            '否则不使用ICON
                Set .itemIcon = LoadPicture()
            
            End If
            
        Else
            Set .itemIcon = itemIcon
            
        End If
        
        .itemText = itemText
        .itemType = itemType
        .itemState = itemState
        .itemhSubMenu = ItemInfo.hSubMenu
        .itemShutCutKey = itemShutCutKey
    End With
    
    
    ' 获得菜单项数据
    With ItemInfo
        .cbSize = LenB(ItemInfo)
        .fMask = MIIM_DATA Or MIIM_ID Or MIIM_TYPE
    End With
    GetMenuItemInfo hMenu, ID, False, ItemInfo
    
    ' 设置菜单项数据
    With ItemInfo
        .fMask = .fMask Or MIIM_TYPE
        .fType = MFT_OWNERDRAW
    End With
    SetMenuItemInfo hMenu, ID, False, ItemInfo
    
    ' 菜单项ID累加
    ID = ID + 1
    
End Sub
'
'' 删除菜单项
'Public Sub DeleteItem(ByVal itemKeyID As Long)
'    Dim i As Long
'    For i = 0 To UBound(MyItemInfo)
'        If MyItemInfo(i).itemKeyID = itemKeyID Then
'            DeleteMenu hMenu, i, 0
'            Exit For
'        End If
'    Next i
'End Sub

' 弹出菜单
Public Sub PopupMenu()
    Dim Pos As POINTAPI
    GetCursorPos Pos
    SetForegroundWindow hWnd
    TrackPopupMenu hMenu, POPUP_LEFTALIGN Or POPUP_TOPALIGN, Pos.X, Pos.Y, 0, hWnd, ByVal 0

End Sub
'
'' 设置菜单项图标
'Public Sub SetItemIcon(ByVal itemKeyID As Long, ByVal itemIcon As StdPicture)
'    Dim i As Long
'    For i = 0 To UBound(MyItemInfo)
'        If MyItemInfo(i).itemKeyID = itemKeyID Then
'            Set MyItemInfo(i).itemIcon = itemIcon
'            Exit For
'        End If
'    Next i
'End Sub
'
'' 获得菜单项图标
'Public Function GetItemIcon(ByVal itemKeyID As Long) As StdPicture
'    Dim i As Long
'    For i = 0 To UBound(MyItemInfo)
'        If MyItemInfo(i).itemKeyID = itemKeyID Then
'            Set GetItemIcon = MyItemInfo(i).itemIcon
'            Exit For
'        End If
'    Next i
'End Function
'
'' 设置菜单项文字
'Public Sub SetItemText(ByVal itemKeyID As Long, ByVal itemText As String)
'    Dim i As Long
'    For i = 0 To UBound(MyItemInfo)
'        If MyItemInfo(i).itemKeyID = itemKeyID Then
'            MyItemInfo(i).itemText = itemText
'            Exit For
'        End If
'    Next i
'End Sub
'
'' 获得菜单项文字
'Public Function GetItemText(ByVal itemKeyID As Long) As String
'    Dim i As Long
'    For i = 0 To UBound(MyItemInfo)
'        If MyItemInfo(i).itemKeyID = itemKeyID Then
'            GetItemText = MyItemInfo(i).itemText
'            Exit For
'        End If
'    Next i
'End Function
'
'' 设置菜单项状态
'Public Sub SetItemState(ByVal itemKeyID As Long, ByVal itemState As MenuItemState)
'    Dim i As Long
'    For i = 0 To UBound(MyItemInfo)
'        If MyItemInfo(i).itemKeyID = itemKeyID Then
'            MyItemInfo(i).itemState = itemState
'            Dim ItemInfo As MENUITEMINFO
'            With ItemInfo
'                .cbSize = Len(ItemInfo)
'                .fMask = MIIM_STRING Or MIIM_FTYPE Or MIIM_STATE Or MIIM_SUBMENU Or MIIM_ID Or MIIM_DATA
'            End With
'            GetMenuItemInfo hMenu, i, False, ItemInfo
'            With ItemInfo
'                .fState = .fState Or itemState
'            End With
'            SetMenuItemInfo hMenu, i, False, ItemInfo
'            Exit For
'        End If
'    Next i
'End Sub
'
'' 获得菜单项状态
'Public Function GetItemState(ByVal itemKeyID As Long) As MenuItemState
'    Dim i As Long
'    For i = 0 To UBound(MyItemInfo)
'        If MyItemInfo(i).itemKeyID = itemKeyID Then
'            GetItemState = MyItemInfo(i).itemState
'            Exit For
'        End If
'    Next i
'End Function
'
'
'' 属性: 菜单附加条宽度
'Public Property Get LeftBarWidth() As Long
'    LeftBarWidth = BarWidth
'End Property
'
'Public Property Let LeftBarWidth(ByVal nBarWidth As Long)
'    If nBarWidth >= 0 Then
'        BarWidth = nBarWidth
'    End If
'End Property
'
'' 属性: 菜单附加条风格
'Public Property Get LeftBarStyle() As MenuLeftBarStyle
'    LeftBarStyle = BarStyle
'End Property
'
'Public Property Let LeftBarStyle(ByVal nBarStyle As MenuLeftBarStyle)
'    If nBarStyle >= 0 And nBarStyle <= 4 Then
'        BarStyle = nBarStyle
'    End If
'End Property
'
'' 属性: 菜单附加条图像(只有当 LeftBarStyle 设置为 LBS_IMAGE 时才有效)
'Public Property Get LeftBarImage() As StdPicture
'    Set LeftBarImage = BarImage
'End Property
'
'Public Property Let LeftBarImage(ByVal nBarImage As StdPicture)
'    Set BarImage = nBarImage
'End Property
'
'' 属性: 菜单附加条过渡色起始颜色(只有当 LeftBarStyle 设置为 LBS_HORIZONTALCOLOR 或 LBS_VERTICALCOLOR 时才有效)
''       当 LeftBarStyle 设置为 LBS_SOLIDCOLOR (实色填充)时以 LeftBarStartColor 颜色为准
'Public Property Get LeftBarStartColor() As Long
'    LeftBarStartColor = BarStartColor
'End Property
'
'Public Property Let LeftBarStartColor(ByVal nBarStartColor As Long)
'    BarStartColor = nBarStartColor
'End Property
'
'' 属性: 菜单附加条过渡色终止颜色(只有当 LeftBarStyle 设置为 LBS_HORIZONTALCOLOR 或 LBS_VERTICALCOLOR 时才有效)
''       当 LeftBarStyle 设置为 LBS_SOLIDCOLOR (实色填充)时以 LeftBarStartColor 颜色为准
'Public Property Get LeftBarEndColor() As Long
'    LeftBarEndColor = BarEndColor
'End Property
'
'Public Property Let LeftBarEndColor(ByVal nBarEndColor As Long)
'    BarEndColor = nBarEndColor
'End Property
'
'' 属性: 菜单项高亮条的范围
'Public Property Get ItemSelectScope() As MenuItemSelectScope
'    ItemSelectScope = SelectScope
'End Property
'
'Public Property Let ItemSelectScope(ByVal nSelectScope As MenuItemSelectScope)
'    SelectScope = nSelectScope
'End Property
'
'' 属性: 菜单项可用时文字颜色
'Public Property Get ItemTextEnabledColor() As Long
'    ItemTextEnabledColor = TextEnabledColor
'End Property
'
'Public Property Let ItemTextEnabledColor(ByVal nTextEnabledColor As Long)
'    TextEnabledColor = nTextEnabledColor
'End Property
'
'' 属性: 菜单项不可用时文字颜色
'Public Property Get ItemTextDisabledColor() As Long
'    ItemTextDisabledColor = TextDisabledColor
'End Property
'
'Public Property Let ItemTextDisabledColor(ByVal nTextDisabledColor As Long)
'    TextDisabledColor = nTextDisabledColor
'End Property
'
'' 属性: 菜单项选中时文字颜色
'Public Property Get ItemTextSelectColor() As Long
'    ItemTextSelectColor = TextSelectColor
'End Property
'
'Public Property Let ItemTextSelectColor(ByVal nTextSelectColor As Long)
'    TextSelectColor = nTextSelectColor
'End Property
'
'' 属性: 菜单项图标风格
'Public Property Get ItemIconStyle() As MenuItemIconStyle
'    ItemIconStyle = IconStyle
'End Property
'
'Public Property Let ItemIconStyle(ByVal nIconStyle As MenuItemIconStyle)
'    IconStyle = nIconStyle
'End Property
'
'' 属性: 菜单项边框风格
'Public Property Get ItemSelectEdgeStyle() As MenuItemSelectEdgeStyle
'    ItemSelectEdgeStyle = EdgeStyle
'End Property
'
'Public Property Let ItemSelectEdgeStyle(ByVal nEdgeStyle As MenuItemSelectEdgeStyle)
'    EdgeStyle = nEdgeStyle
'End Property
'
'' 属性: 菜单项边框颜色
'Public Property Get ItemSelectEdgeColor() As Long
'    ItemSelectEdgeColor = EdgeColor
'End Property
'
'Public Property Let ItemSelectEdgeColor(ByVal nEdgeColor As Long)
'    EdgeColor = nEdgeColor
'End Property
'
'' 属性: 菜单项背景填充风格
'Public Property Get ItemSelectFillStyle() As MenuItemSelectFillStyle
'    ItemSelectFillStyle = FillStyle
'End Property
'
'Public Property Let ItemSelectFillStyle(ByVal nFillStyle As MenuItemSelectFillStyle)
'    FillStyle = nFillStyle
'End Property
'
' 属性: 菜单项过渡色起始颜色(只有当 ItemSelectFillStyle 设置为 ISFS_HORIZONTALCOLOR 或 ISFS_VERTICALCOLOR 时才有效)
'       当 ItemSelectFillStyle 设置为 ISFS_SOLIDCOLOR (实色填充)时以 ItemSelectFillStartColor 颜色为准
'Public Property Get ItemSelectFillStartColor() As Long
'    ItemSelectFillStartColor = FillStartColor
'End Property
'
'Public Property Let ItemSelectFillStartColor(ByVal nFillStartColor As Long)
'    FillStartColor = nFillStartColor
'End Property
'
'' 属性: 菜单项过渡色终止颜色(只有当 ItemSelectFillStyle 设置为 ISFS_HORIZONTALCOLOR 或 ISFS_VERTICALCOLOR 时才有效)
''       当 ItemSelectFillStyle 设置为 ISFS_SOLIDCOLOR (实色填充)时以 ItemSelectFillStartColor 颜色为准
'Public Property Get ItemSelectFillEndColor() As Long
'    ItemSelectFillEndColor = FillEndColor
'End Property
'
'Public Property Let ItemSelectFillEndColor(ByVal nFillEndColor As Long)
'    FillEndColor = nFillEndColor
'End Property
'
'' 属性: 菜单背景颜色
'Public Property Get BackColor() As Long
'    BackColor = BkColor
'End Property
'
'Public Property Let BackColor(ByVal nBkColor As Long)
'    BkColor = nBkColor
'End Property
'
'' 属性: 菜单分隔条风格
'Public Property Get SeparatorStyle() As MenuSeparatorStyle
'    SeparatorStyle = SepStyle
'End Property
'
'Public Property Let SeparatorStyle(ByVal nSepStyle As MenuSeparatorStyle)
'    SepStyle = nSepStyle
'End Property
'
'' 属性: 菜单分隔条颜色
'Public Property Get SeparatorColor() As Long
'    SeparatorColor = SepColor
'End Property
'
'Public Property Let SeparatorColor(ByVal nSepColor As Long)
'    SepColor = nSepColor
'End Property
'
' 属性: 菜单总体风格

'Private Property Get Style() As MenuUserStyle
'    Style = MenuStyle
'End Property
'
'Private Property Let Style(ByVal nMenuStyle As MenuUserStyle)
'    MenuStyle = nMenuStyle
'    Select Case nMenuStyle
'        Case STYLE_WINDOWS                                              ' Windows 默认风格
'            Set BarImage = LoadPicture()
'            BarWidth = 10
'            BarStyle = LBS_NONE
'            BarStartColor = GetSysColor(COLOR_MENU)
'            BarEndColor = BarStartColor
'            SelectScope = ISS_ICON_TEXT
'            TextEnabledColor = GetSysColor(COLOR_MENUTEXT)
'            TextDisabledColor = GetSysColor(COLOR_GRAYTEXT)
'            TextSelectColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
'            IconStyle = IIS_NONE
'            EdgeStyle = ISES_SOLID
'            EdgeColor = GetSysColor(COLOR_HIGHLIGHT)
'            FillStyle = ISFS_SOLIDCOLOR
'            FillStartColor = EdgeColor
'            FillEndColor = FillStartColor
'            BkColor = GetSysColor(COLOR_MENU)
'            SepColor = TextDisabledColor
'            SepStyle = MSS_DEFAULT
'        Case STYLE_XP                                                   ' XP 风格
'            Set BarImage = LoadPicture()
'            BarWidth = 20
'            BarStyle = LBS_NONE
'            BarStartColor = GetSysColor(COLOR_MENU)
'            BarEndColor = BarStartColor
'            SelectScope = ISS_ICON_TEXT
'            TextEnabledColor = GetSysColor(COLOR_MENUTEXT)
'            TextDisabledColor = GetSysColor(COLOR_GRAYTEXT)
'            TextSelectColor = TextEnabledColor
'            IconStyle = IIS_SHADOW
'            EdgeStyle = ISES_SOLID
'            EdgeColor = RGB(49, 106, 197)
'            FillStyle = ISFS_SOLIDCOLOR
'            FillStartColor = RGB(180, 195, 210)
'            FillEndColor = FillStartColor
'            BkColor = GetSysColor(COLOR_MENU)
'            SepColor = RGB(192, 192, 192)
'            SepStyle = MSS_SOLID
'        Case STYLE_SHADE                                                ' 渐变风格
'            Set BarImage = LoadPicture()
'            BarWidth = 20
'            BarStyle = LBS_VERTICALCOLOR
'            BarStartColor = vbBlack
'            BarEndColor = vbWhite
'            SelectScope = ISS_ICON_TEXT
'            TextEnabledColor = GetSysColor(COLOR_MENUTEXT)
'            TextDisabledColor = GetSysColor(COLOR_GRAYTEXT)
'            TextSelectColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
'            IconStyle = IIS_NONE
'            EdgeStyle = ISES_NONE
'            EdgeColor = GetSysColor(COLOR_HIGHLIGHT)
'            FillStyle = ISFS_HORIZONTALCOLOR
'            FillStartColor = vbBlack
'            FillEndColor = vbWhite
'            BkColor = GetSysColor(COLOR_MENU)
'            SepColor = TextDisabledColor
'            SepStyle = MSS_DEFAULT
'        Case STYLE_3D                                                   ' 3D 立体风格
'            Set BarImage = LoadPicture()
'            BarWidth = 20
'            BarStyle = LBS_NONE
'            BarStartColor = GetSysColor(COLOR_MENU)
'            BarEndColor = BarStartColor
'            SelectScope = ISS_TEXT
'            TextEnabledColor = GetSysColor(COLOR_MENUTEXT)
'            TextDisabledColor = GetSysColor(COLOR_GRAYTEXT)
'            TextSelectColor = vbBlue
'            IconStyle = IIS_RAISED
'            EdgeStyle = ISES_SUNKEN
'            EdgeColor = GetSysColor(COLOR_HIGHLIGHT)
'            FillStyle = ISFS_NONE
'            FillStartColor = EdgeColor
'            FillEndColor = FillStartColor
'            BkColor = GetSysColor(COLOR_MENU)
'            SepColor = TextDisabledColor
'            SepStyle = MSS_DEFAULT
'    End Select
'End Property

Public Sub SetPopupMenu(ByVal hMenuOwner As Long)
    hWnd = hMenuOwner
    preMenuWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf MenuWndProc)
    hMenu = CreatePopupMenu()
'    Me.Style = STYLE_WINDOWS
    RaiseEvent cyMenuShow


End Sub

Friend Function FireEvent(ByVal iMenuKeyID As Long, ByVal sMenuCaption As String)
    RaiseEvent cyMenuClick(iMenuKeyID, sMenuCaption)
    
End Function

Friend Function FireClose()
    RaiseEvent cyMenuHide
    
End Function

