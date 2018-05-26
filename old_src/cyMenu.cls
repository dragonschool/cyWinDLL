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
'* ����ģ����һ���˵���, �ṩ�˸�����ʽ�Ĳ˵�����������
'*
'* ��Ȩ: LPP����������
'* ����: ¬����(goodname008)
'* (******* �����뱣��������Ϣ *******)
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

Public Enum MenuUserStyle                                   ' �˵�������
    STYLE_WINDOWS
    STYLE_XP
    STYLE_SHADE
    STYLE_3D
    STYLE_COLORFUL
End Enum

Public Enum MenuSeparatorStyle                              ' �˵��ָ������
    MSS_SOLID
    MSS_DASH
    MSS_DOT
    MSS_DASDOT
    MSS_DASHDOTDOT
    MSS_NONE
    MSS_DEFAULT
End Enum

Public Enum MenuItemSelectFillStyle                         ' �˵���������
    ISFS_NONE
    ISFS_SOLIDCOLOR
    ISFS_HORIZONTALCOLOR
    ISFS_VERTICALCOLOR
End Enum

Public Enum MenuItemSelectEdgeStyle                         ' �˵���߿���
    ISES_SOLID
    ISES_DASH
    ISES_DOT
    ISES_DASDOT
    ISES_DASHDOTDOT
    ISES_NONE
    ISES_SUNKEN
    ISES_RAISED
End Enum

Public Enum MenuItemIconStyle                               ' �˵���ͼ����
    IIS_NONE
    IIS_SUNKEN
    IIS_RAISED
    IIS_SHADOW
End Enum

Public Enum MenuItemSelectScope                             ' �˵���������ķ�Χ
    ISS_TEXT = &H1
    ISS_ICON_TEXT = &H2
    ISS_LEFTBAR_ICON_TEXT = &H4
End Enum

Public Enum MenuLeftBarStyle                                ' �˵����������
    LBS_NONE
    LBS_SOLIDCOLOR
    LBS_HORIZONTALCOLOR
    LBS_VERTICALCOLOR
    LBS_IMAGE
End Enum

Public Enum MenuItemType                                    ' �˵�������
    MIT_STRING = &H0
    MIT_CHECKBOX = &H200
    MIT_SEPARATOR = &H800
End Enum

Public Enum MenuItemState                                   ' �˵���״̬
    MIS_ENABLED = &H0
    MIS_DISABLED = &H2
    MIS_CHECKED = &H8
    MIS_UNCHECKED = &H0
End Enum

Public Enum PopupAlign                                      ' �˵��������뷽ʽ
    POPUP_LEFTALIGN = &H0&                                  ' ˮƽ�����
    POPUP_CENTERALIGN = &H4&                                ' ˮƽ���ж���
    POPUP_RIGHTALIGN = &H8&                                 ' ˮƽ�Ҷ���
    POPUP_TOPALIGN = &H0&                                   ' ��ֱ�϶���
    POPUP_VCENTERALIGN = &H10&                              ' ��ֱ���ж���
    POPUP_BOTTOMALIGN = &H20&                               ' ��ֱ�¶���
End Enum

Private Sub Class_Initialize()
    modMenu.objMenu = ObjPtr(Me)
    Erase MyItemInfo
    
End Sub

' �ͷ���
Private Sub Class_Terminate()
    SetWindowLong hWnd, GWL_WNDPROC, preMenuWndProc
    DestroyMenu hMenu
End Sub

' ����˵�������Զ���˵�������, ����Owner_Draw�Ի�˵�
Public Sub AddItem(ByVal itemKeyID As Long, ByVal itemText As String, Optional ByVal itemType As MenuItemType = MIT_STRING, Optional ByVal itemShutCutKey As String, Optional ByVal itemIcon As StdPicture, Optional ByVal itemState As MenuItemState, Optional ByVal itemHasSubMenu As Boolean, Optional ByVal ParentItemKeyID As Long)
    Static ID As Long, i As Long
    Dim ItemInfo As MENUITEMINFO
    
    itemShutCutKey = UCase(itemShutCutKey)
    itemText = IIf(itemShutCutKey = "", itemText, itemText & "(&" & itemShutCutKey & ")")
    
    ' ����˵���
    With ItemInfo
        .cbSize = LenB(ItemInfo)
        .fMask = MIIM_STRING Or MIIM_FTYPE Or MIIM_STATE Or MIIM_SUBMENU Or MIIM_ID Or MIIM_DATA
        .fType = itemType
        .fState = itemState
        .wid = ID
        .dwItemData = True
        .cch = lstrlen(itemText)
        .dwTypeData = itemText
        
        '������Ӳ˵�
        If itemHasSubMenu Then
            .hSubMenu = CreatePopupMenu()

        End If
        
    End With
    
    If ParentItemKeyID > 0 Then
    '���и��˵�
        For i = 0 To UBound(MyItemInfo)
            If MyItemInfo(i).itemKeyID = ParentItemKeyID Then GoTo FindParentID
        
        Next i
        Err.Raise -9999, "Menu", "û���ҵ��ò˵��ĸ�ID."
        Exit Sub
FindParentID:

        InsertMenuItem MyItemInfo(i).itemhSubMenu, ID, False, ItemInfo
    Else
    '�������Ǹ��˵�
        InsertMenuItem hMenu, ID, False, ItemInfo
    End If
    
    ' ���˵������ݴ��붯̬����
    ReDim Preserve MyItemInfo(ID) As MyMenuItemInfo
    
    '����Ƿ�����ͬ��KeyID
    For i = 0 To UBound(MyItemInfo)
        If MyItemInfo(i).itemKeyID = itemKeyID Then
            Class_Terminate
            Err.Raise -9999, "Menu", "�����˵���KeyID���ڳ�ͻ."
        End If
    Next i
    
    With MyItemInfo(ID)
        .itemKeyID = itemKeyID
        If (itemIcon Is Nothing) Then
            If itemState = MIS_CHECKED Then
            '���״̬�Ǳ�ѡ��״̬������ʾȱʡ�Ĺ���
                Set .itemIcon = frmIcon.Label1.DragIcon
            
            Else
            '����ʹ��ICON
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
    
    
    ' ��ò˵�������
    With ItemInfo
        .cbSize = LenB(ItemInfo)
        .fMask = MIIM_DATA Or MIIM_ID Or MIIM_TYPE
    End With
    GetMenuItemInfo hMenu, ID, False, ItemInfo
    
    ' ���ò˵�������
    With ItemInfo
        .fMask = .fMask Or MIIM_TYPE
        .fType = MFT_OWNERDRAW
    End With
    SetMenuItemInfo hMenu, ID, False, ItemInfo
    
    ' �˵���ID�ۼ�
    ID = ID + 1
    
End Sub
'
'' ɾ���˵���
'Public Sub DeleteItem(ByVal itemKeyID As Long)
'    Dim i As Long
'    For i = 0 To UBound(MyItemInfo)
'        If MyItemInfo(i).itemKeyID = itemKeyID Then
'            DeleteMenu hMenu, i, 0
'            Exit For
'        End If
'    Next i
'End Sub

' �����˵�
Public Sub PopupMenu()
    Dim Pos As POINTAPI
    GetCursorPos Pos
    SetForegroundWindow hWnd
    TrackPopupMenu hMenu, POPUP_LEFTALIGN Or POPUP_TOPALIGN, Pos.X, Pos.Y, 0, hWnd, ByVal 0

End Sub
'
'' ���ò˵���ͼ��
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
'' ��ò˵���ͼ��
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
'' ���ò˵�������
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
'' ��ò˵�������
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
'' ���ò˵���״̬
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
'' ��ò˵���״̬
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
'' ����: �˵�����������
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
'' ����: �˵����������
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
'' ����: �˵�������ͼ��(ֻ�е� LeftBarStyle ����Ϊ LBS_IMAGE ʱ����Ч)
'Public Property Get LeftBarImage() As StdPicture
'    Set LeftBarImage = BarImage
'End Property
'
'Public Property Let LeftBarImage(ByVal nBarImage As StdPicture)
'    Set BarImage = nBarImage
'End Property
'
'' ����: �˵�����������ɫ��ʼ��ɫ(ֻ�е� LeftBarStyle ����Ϊ LBS_HORIZONTALCOLOR �� LBS_VERTICALCOLOR ʱ����Ч)
''       �� LeftBarStyle ����Ϊ LBS_SOLIDCOLOR (ʵɫ���)ʱ�� LeftBarStartColor ��ɫΪ׼
'Public Property Get LeftBarStartColor() As Long
'    LeftBarStartColor = BarStartColor
'End Property
'
'Public Property Let LeftBarStartColor(ByVal nBarStartColor As Long)
'    BarStartColor = nBarStartColor
'End Property
'
'' ����: �˵�����������ɫ��ֹ��ɫ(ֻ�е� LeftBarStyle ����Ϊ LBS_HORIZONTALCOLOR �� LBS_VERTICALCOLOR ʱ����Ч)
''       �� LeftBarStyle ����Ϊ LBS_SOLIDCOLOR (ʵɫ���)ʱ�� LeftBarStartColor ��ɫΪ׼
'Public Property Get LeftBarEndColor() As Long
'    LeftBarEndColor = BarEndColor
'End Property
'
'Public Property Let LeftBarEndColor(ByVal nBarEndColor As Long)
'    BarEndColor = nBarEndColor
'End Property
'
'' ����: �˵���������ķ�Χ
'Public Property Get ItemSelectScope() As MenuItemSelectScope
'    ItemSelectScope = SelectScope
'End Property
'
'Public Property Let ItemSelectScope(ByVal nSelectScope As MenuItemSelectScope)
'    SelectScope = nSelectScope
'End Property
'
'' ����: �˵������ʱ������ɫ
'Public Property Get ItemTextEnabledColor() As Long
'    ItemTextEnabledColor = TextEnabledColor
'End Property
'
'Public Property Let ItemTextEnabledColor(ByVal nTextEnabledColor As Long)
'    TextEnabledColor = nTextEnabledColor
'End Property
'
'' ����: �˵������ʱ������ɫ
'Public Property Get ItemTextDisabledColor() As Long
'    ItemTextDisabledColor = TextDisabledColor
'End Property
'
'Public Property Let ItemTextDisabledColor(ByVal nTextDisabledColor As Long)
'    TextDisabledColor = nTextDisabledColor
'End Property
'
'' ����: �˵���ѡ��ʱ������ɫ
'Public Property Get ItemTextSelectColor() As Long
'    ItemTextSelectColor = TextSelectColor
'End Property
'
'Public Property Let ItemTextSelectColor(ByVal nTextSelectColor As Long)
'    TextSelectColor = nTextSelectColor
'End Property
'
'' ����: �˵���ͼ����
'Public Property Get ItemIconStyle() As MenuItemIconStyle
'    ItemIconStyle = IconStyle
'End Property
'
'Public Property Let ItemIconStyle(ByVal nIconStyle As MenuItemIconStyle)
'    IconStyle = nIconStyle
'End Property
'
'' ����: �˵���߿���
'Public Property Get ItemSelectEdgeStyle() As MenuItemSelectEdgeStyle
'    ItemSelectEdgeStyle = EdgeStyle
'End Property
'
'Public Property Let ItemSelectEdgeStyle(ByVal nEdgeStyle As MenuItemSelectEdgeStyle)
'    EdgeStyle = nEdgeStyle
'End Property
'
'' ����: �˵���߿���ɫ
'Public Property Get ItemSelectEdgeColor() As Long
'    ItemSelectEdgeColor = EdgeColor
'End Property
'
'Public Property Let ItemSelectEdgeColor(ByVal nEdgeColor As Long)
'    EdgeColor = nEdgeColor
'End Property
'
'' ����: �˵���������
'Public Property Get ItemSelectFillStyle() As MenuItemSelectFillStyle
'    ItemSelectFillStyle = FillStyle
'End Property
'
'Public Property Let ItemSelectFillStyle(ByVal nFillStyle As MenuItemSelectFillStyle)
'    FillStyle = nFillStyle
'End Property
'
' ����: �˵������ɫ��ʼ��ɫ(ֻ�е� ItemSelectFillStyle ����Ϊ ISFS_HORIZONTALCOLOR �� ISFS_VERTICALCOLOR ʱ����Ч)
'       �� ItemSelectFillStyle ����Ϊ ISFS_SOLIDCOLOR (ʵɫ���)ʱ�� ItemSelectFillStartColor ��ɫΪ׼
'Public Property Get ItemSelectFillStartColor() As Long
'    ItemSelectFillStartColor = FillStartColor
'End Property
'
'Public Property Let ItemSelectFillStartColor(ByVal nFillStartColor As Long)
'    FillStartColor = nFillStartColor
'End Property
'
'' ����: �˵������ɫ��ֹ��ɫ(ֻ�е� ItemSelectFillStyle ����Ϊ ISFS_HORIZONTALCOLOR �� ISFS_VERTICALCOLOR ʱ����Ч)
''       �� ItemSelectFillStyle ����Ϊ ISFS_SOLIDCOLOR (ʵɫ���)ʱ�� ItemSelectFillStartColor ��ɫΪ׼
'Public Property Get ItemSelectFillEndColor() As Long
'    ItemSelectFillEndColor = FillEndColor
'End Property
'
'Public Property Let ItemSelectFillEndColor(ByVal nFillEndColor As Long)
'    FillEndColor = nFillEndColor
'End Property
'
'' ����: �˵�������ɫ
'Public Property Get BackColor() As Long
'    BackColor = BkColor
'End Property
'
'Public Property Let BackColor(ByVal nBkColor As Long)
'    BkColor = nBkColor
'End Property
'
'' ����: �˵��ָ������
'Public Property Get SeparatorStyle() As MenuSeparatorStyle
'    SeparatorStyle = SepStyle
'End Property
'
'Public Property Let SeparatorStyle(ByVal nSepStyle As MenuSeparatorStyle)
'    SepStyle = nSepStyle
'End Property
'
'' ����: �˵��ָ�����ɫ
'Public Property Get SeparatorColor() As Long
'    SeparatorColor = SepColor
'End Property
'
'Public Property Let SeparatorColor(ByVal nSepColor As Long)
'    SepColor = nSepColor
'End Property
'
' ����: �˵�������

'Private Property Get Style() As MenuUserStyle
'    Style = MenuStyle
'End Property
'
'Private Property Let Style(ByVal nMenuStyle As MenuUserStyle)
'    MenuStyle = nMenuStyle
'    Select Case nMenuStyle
'        Case STYLE_WINDOWS                                              ' Windows Ĭ�Ϸ��
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
'        Case STYLE_XP                                                   ' XP ���
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
'        Case STYLE_SHADE                                                ' ������
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
'        Case STYLE_3D                                                   ' 3D ������
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
