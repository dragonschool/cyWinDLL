Attribute VB_Name = "modMenu"
'**************************************************************************************************************
'* ��ģ����� cyMenu �˵���ģ��
'*
'* ��Ȩ: LPP���������
'* ����: ¬����(goodname008)
'* (******* �����뱣��������Ϣ *******)
'**************************************************************************************************************

Option Explicit

' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- API �������� -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

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


' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- API �������� -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Public Const GWL_WNDPROC = (-4)                     ' SetWindowLong ���ô��ں�����ڵ�ַ
Public Const SM_CYMENU = 15                         ' GetSystemMetrics ���ϵͳ�˵���߶�

Public Const WM_COMMAND = &H111                     ' ��Ϣ: �����˵���
Public Const WM_DRAWITEM = &H2B                     ' ��Ϣ: ���Ʋ˵���
Public Const WM_EXITMENULOOP = &H212                ' ��Ϣ: �˳��˵���Ϣѭ��
Public Const WM_MEASUREITEM = &H2C                  ' ��Ϣ: ����˵��߶ȺͿ��
Public Const WM_MENUSELECT = &H11F                  ' ��Ϣ: ѡ��˵���
Public Const WM_MENUCHAR = &H120                    ' ��Ϣ: ʹ�ÿ�ݼ�ѡ��˵�

' ODT
Public Const ODT_MENU = 1                           ' �˵�
Public Const ODT_LISTBOX = 2                        ' �б��
Public Const ODT_COMBOBOX = 3                       ' ��Ͽ�
Public Const ODT_BUTTON = 4                         ' ��ť

' ODS
Public Const ODS_SELECTED = &H1                     ' �˵���ѡ��
Public Const ODS_GRAYED = &H2                       ' ��ɫ��
Public Const ODS_DISABLED = &H4                     ' ����
Public Const ODS_CHECKED = &H8                      ' ѡ��
Public Const ODS_FOCUS = &H10                       ' �۽�

' diFlags to DrawIconEx
Public Const DI_MASK = &H1                          ' ��ͼʱʹ��ͼ���MASK���� (�絥��ʹ��, �ɻ��ͼ�����ģ)
Public Const DI_IMAGE = &H2                         ' ��ͼʱʹ��ͼ���XOR���� (��ͼ��û��͸������)
Public Const DI_NORMAL = DI_MASK Or DI_IMAGE        ' �ó��淽ʽ��ͼ (�ϲ� DI_IMAGE �� DI_MASK)

' nBkMode to SetBkMode
Public Const TRANSPARENT = 1                        ' ͸������, �������������
Public Const OPAQUE = 2                             ' �õ�ǰ�ı���ɫ������߻��ʡ���Ӱˢ���Լ��ַ��Ŀ�϶
Public Const NEWTRANSPARENT = 3                     ' ������ɫ�Ĳ˵��ϻ�͸������


' MF �˵���س���
Public Const MF_BYCOMMAND = &H0&                    ' �˵���Ŀ�ɲ˵�������IDָ��
Public Const MF_BYPOSITION = &H400&                 ' �˵���Ŀ����Ŀ�ڲ˵��е�λ�þ��� (�����˵��еĵ�һ����Ŀ)

Public Const MF_CHECKED = &H8&                      ' ���ָ���Ĳ˵���Ŀ (������VB��Checked���Լ���)
Public Const MF_DISABLED = &H2&                     ' ��ָֹ���Ĳ˵���Ŀ (����VB��Enabled���Լ���)
Public Const MF_ENABLED = &H0&                      ' ����ָ���Ĳ˵���Ŀ (����VB��Enabled���Լ���)
Public Const MF_GRAYED = &H1&                       ' ��ָֹ���Ĳ˵���Ŀ, ����ǳ��ɫ������. (����VB��Enabled���Լ���)
Public Const MF_HILITE = &H80&
Public Const MF_SEPARATOR = &H800&                  ' ��ָ������Ŀ����ʾһ���ָ���
Public Const MF_STRING = &H0&                       ' ��ָ������Ŀ������һ���ִ� (����VB��Caption���Լ���)
Public Const MF_UNCHECKED = &H0&                    ' ���ָ������Ŀ (������VB��Checked���Լ���)
Public Const MF_UNHILITE = &H0&

Public Const MF_BITMAP = &H4&                       ' �˵���Ŀ��һ��λͼ. һ������˵�, ���λͼ�;��Բ���ɾ��, ���Բ�Ӧ��ʹ����VB��Image���Է��ص�ֵ.
Public Const MF_OWNERDRAW = &H100&                  ' ����һ��������ͼ�˵� (������Ƶĳ��������ÿ���˵���Ŀ)
Public Const MF_USECHECKBITMAPS = &H200&

Public Const MF_MENUBARBREAK = &H20&                ' �ڵ���ʽ�˵���, ��ָ������Ŀ������һ������, ����һ����ֱ�߷ָ���ͬ����.
Public Const MF_MENUBREAK = &H40&                   ' �ڵ���ʽ�˵���, ��ָ������Ŀ������һ������. �ڶ����˵���, ����Ŀ���õ�һ������.

Public Const MF_POPUP = &H10&                       ' ��һ������ʽ�˵�����ָ������Ŀ, �����ڴ����Ӳ˵�������ʽ�˵�.
Public Const MF_HELP = &H4000&

Public Const MF_DEFAULT = &H1000
Public Const MF_RIGHTJUSTIFY = &H4000

' fMask To InsertMenuItem                           ' ָ�� MENUITEMINFO ����Щ��Ա��Ч
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_CHECKMARKS = &H8
Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA = &H20
Public Const MIIM_STRING = &H40
Public Const MIIM_BITMAP = &H80
Public Const MIIM_FTYPE = &H100

' fType To InsertMenuItem                           ' MENUITEMINFO �в˵�������
Public Const MFT_BITMAP = &H4&
Public Const MFT_MENUBARBREAK = &H20&
Public Const MFT_MENUBREAK = &H40&
Public Const MFT_OWNERDRAW = &H100&
Public Const MFT_SEPARATOR = &H800&
Public Const MFT_STRING = &H0&

' fState to InsertMenuItem                          ' MENUITEMINFO �в˵���״̬
Public Const MFS_CHECKED = &H8&
Public Const MFS_DISABLED = &H2&
Public Const MFS_ENABLED = &H0&
Public Const MFS_GRAYED = &H1&
Public Const MFS_HILITE = &H80&
Public Const MFS_UNCHECKED = &H0&
Public Const MFS_UNHILITE = &H0&

' nFormat to DrawText
Public Const DT_LEFT = &H0                          ' ˮƽ�����
Public Const DT_CENTER = &H1                        ' ˮƽ���ж���
Public Const DT_RIGHT = &H2                         ' ˮƽ�Ҷ���

Public Const DT_SINGLELINE = &H20                   ' ����

Public Const DT_TOP = &H0                           ' ��ֱ�϶��� (������ʱ��Ч)
Public Const DT_VCENTER = &H4                       ' ��ֱ���ж��� (������ʱ��Ч)
Public Const DT_BOTTOM = &H8                        ' ��ֱ�¶��� (������ʱ��Ч)

Public Const DT_CALCRECT = &H400                    ' ���л�ͼʱ���εĵױ߸�����Ҫ������չ, �Ա�������������; ���л�ͼʱ, ��չ���ε��Ҳ�, ���������, ��lpRect����ָ���ľ��λ�������������ֵ.
Public Const DT_WORDBREAK = &H10                    ' �����Զ�����. ����SetTextAlign����������TA_UPDATECP��־, �������������Ч.

Public Const DT_NOCLIP = &H100                      ' �������ʱ�����е�ָ���ľ���
Public Const DT_NOPREFIX = &H800                    ' ͨ��, ������Ϊ & �ַ���ʾӦΪ��һ���ַ������»���, �ñ�־��ֹ������Ϊ.

Public Const DT_EXPANDTABS = &H40                   ' ������ֵ�ʱ��, ���Ʊ�վ������չ. Ĭ�ϵ��Ʊ�վ�����8���ַ�. ����, ����DT_TABSTOP��־�ı������趨.
Public Const DT_TABSTOP = &H80                      ' ָ���µ��Ʊ�վ���, ������������ĸ� 8 λ.
Public Const DT_EXTERNALLEADING = &H200             ' �����ı��и߶ȵ�ʱ��, ʹ�õ�ǰ������ⲿ�������.

' nIndex to GetSysColor  ��׼: 0--20
Public Const COLOR_ACTIVEBORDER = 10                ' ����ڵı߿�
Public Const COLOR_ACTIVECAPTION = 2                ' ����ڵı���
Public Const COLOR_APPWORKSPACE = 12                ' MDI����ı���
Public Const COLOR_BACKGROUND = 1                   ' Windows ����
Public Const COLOR_BTNFACE = 15                     ' ��ť
Public Const COLOR_BTNHIGHLIGHT = 20                ' ��ť��3D������
Public Const COLOR_BTNSHADOW = 16                   ' ��ť��3D��Ӱ
Public Const COLOR_BTNTEXT = 18                     ' ��ť����
Public Const COLOR_CAPTIONTEXT = 9                  ' ���ڱ����е�����
Public Const COLOR_GRAYTEXT = 17                    ' ��ɫ����; ��ʹ���˶���������Ϊ��
Public Const COLOR_HIGHLIGHT = 13                   ' ѡ������Ŀ����
Public Const COLOR_HIGHLIGHTTEXT = 14               ' ѡ������Ŀ����
Public Const COLOR_INACTIVEBORDER = 11              ' ������ڵı߿�
Public Const COLOR_INACTIVECAPTION = 3              ' ������ڵı���
Public Const COLOR_INACTIVECAPTIONTEXT = 19         ' ������ڵ�����
Public Const COLOR_MENU = 4                         ' �˵�
Public Const COLOR_MENUTEXT = 7                     ' �˵�����
Public Const COLOR_SCROLLBAR = 0                    ' ������
Public Const COLOR_WINDOW = 5                       ' ���ڱ���
Public Const COLOR_WINDOWFRAME = 6                  ' ����
Public Const COLOR_WINDOWTEXT = 8                   ' ��������

' un to DrawState
Public Const DST_COMPLEX = &H0                      ' ��ͼ����lpDrawStateProc����ָ���Ļص������ڼ�ִ��, lParam��wParam�ᴫ�ݸ��ص��¼�.
Public Const DST_TEXT = &H1                         ' lParam�������ֵĵ�ַ(��ʹ��һ���ִ�����),wParam�����ִ��ĳ���.
Public Const DST_PREFIXTEXT = &H2                   ' ��DST_TEXT����, ֻ�� & �ַ�ָ��Ϊ�¸��ַ������»���.
Public Const DST_ICON = &H3                         ' lParam����ͼ��ľ��
Public Const DST_BITMAP = &H4                       ' lParam����λͼ�ľ��
Public Const DSS_NORMAL = &H0                       ' ��ͨͼ��
Public Const DSS_UNION = &H10                       ' ͼ����ж�������
Public Const DSS_DISABLED = &H20                    ' ͼ����и���Ч��
Public Const DSS_MONO = &H80                        ' ��hBrush���ͼ��
Public Const DSS_RIGHT = &H8000                     ' ���κ�����

' edge to DrawEdge
Public Const BDR_RAISEDOUTER = &H1                  ' ���͹
Public Const BDR_SUNKENOUTER = &H2                  ' ��㰼
Public Const BDR_RAISEDINNER = &H4                  ' �ڲ�͹
Public Const BDR_SUNKENINNER = &H8                  ' �ڲ㰼
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA
Public Const BDR_INNER = &HC
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

' grfFlags to DrawEdge
Public Const BF_LEFT = &H1                          ' ���Ե
Public Const BF_TOP = &H2                           ' �ϱ�Ե
Public Const BF_RIGHT = &H4                         ' �ұ�Ե
Public Const BF_BOTTOM = &H8                        ' �±�Ե
Public Const BF_DIAGONAL = &H10                     ' �Խ���
Public Const BF_MIDDLE = &H800                      ' �������ڲ�
Public Const BF_SOFT = &H1000                       ' MSDN: Soft buttons instead of tiles.
Public Const BF_ADJUST = &H2000                     ' ��������, Ԥ���ͻ���
Public Const BF_FLAT = &H4000                       ' ƽ���Ե
Public Const BF_MONO = &H8000                       ' һά��Ե

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
Public Const PS_DASH = 1                            ' ��������:���� (nWidth������1)         -------
Public Const PS_DASHDOT = 3                         ' ��������:�㻮�� (nWidth������1)       _._._._
Public Const PS_DASHDOTDOT = 4                      ' ��������:��-��-���� (nWidth������1)   _.._.._
Public Const PS_DOT = 2                             ' ��������:���� (nWidth������1)         .......
Public Const PS_SOLID = 0                           ' ��������:ʵ��                         _______


' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- API �������� -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-


Public Enum DrawStateStyle
    dssNormal = 0               '����״̬
    dssDisabled = &H20          '��Ч״̬
    dssShadow = &H80            '������Ӱ
    dssSmooth = &H100           '����ƽ��ͼ��
End Enum

'����һ����Ľṹ
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


' �Զ���˵������ݽṹ
Public Type MyMenuItemInfo
    itemKeyID As Long
    itemIcon As StdPicture
    itemAlias As String
    itemText As String
    itemType As MenuItemType
    itemState As MenuItemState
    itemhSubMenu As Long            '�Ӳ˵����
    itemShutCutKey As String        '��ݼ���ĸ
End Type

' �˵���ؽṹ
Private MeasureInfo As MEASUREITEMSTRUCT

Public hMenu As Long
Public hWnd As Long

Public preMenuWndProc As Long
Public MyItemInfo() As MyMenuItemInfo

' �˵�������
Public BarWidth As Long                             ' �˵����������
Public BarStyle As MenuLeftBarStyle                 ' �˵����������
Public BarImage As StdPicture                       ' �˵�������ͼ��
Public BarStartColor As Long                        ' �˵�����������ɫ��ʼ��ɫ
Public BarEndColor As Long                          ' �˵�����������ɫ��ֹ��ɫ
Public SelectScope As MenuItemSelectScope           ' �˵���������ķ�Χ
Public TextEnabledColor As Long                     ' �˵������ʱ������ɫ
Public TextDisabledColor As Long                    ' �˵������ʱ������ɫ
Public TextSelectColor As Long                      ' �˵���ѡ��ʱ������ɫ
Public IconStyle As MenuItemIconStyle               ' �˵���ͼ����
Public EdgeStyle As MenuItemSelectEdgeStyle         ' �˵���߿���
Public EdgeColor As Long                            ' �˵���߿���ɫ
Public FillStyle As MenuItemSelectFillStyle         ' �˵���������
Public FillStartColor As Long                       ' �˵������ɫ��ʼ��ɫ
Public FillEndColor As Long                         ' �˵������ɫ��ֹ��ɫ
Public BkColor As Long                              ' �˵�������ɫ
Public SepStyle As MenuSeparatorStyle               ' �˵��ָ������
Public SepColor As Long                             ' �˵��ָ�����ɫ
Public MenuStyle As MenuUserStyle                   ' �˵�������

'�����ȼ�������
Public objMenu As Long
Public preMenuProc As Long


' ���ز˵���Ϣ (frmMenu ������ں���)
Function MenuWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
        Case WM_COMMAND                                                 ' �����˵���
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
        Case WM_EXITMENULOOP                                            ' �˳��˵���Ϣѭ��(����)
            ptrMenu.FireClose
        
        Case WM_MENUCHAR
            Dim lngRetval As Long
            If OnMenuChar(wParam, lngRetval) Then
                MenuWndProc = lngRetval
            
            End If
        
        Case WM_MEASUREITEM                                             ' ����˵���߶ȺͿ��
            MeasureItem hWnd, lParam
            
        Case WM_DRAWITEM                                                ' ���Ʋ˵���
            If OnDrawItem(lParam) Then
                MenuWndProc = 1
            
            End If
       
    End Select
    MenuWndProc = CallWindowProc(preMenuWndProc, hWnd, Msg, wParam, lParam)
    
End Function

' ����˵��߶ȺͿ��
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

' �˵����¼���Ӧ(�����˵���)
Private Sub MenuItemSelected(ByVal itemID As Long)
On Error GoTo Err
    If MyItemInfo(itemID).itemShutCutKey = "" Then
    'û�п�ݼ�
        ptrMenu.FireEvent MyItemInfo(itemID).itemKeyID, MyItemInfo(itemID).itemText
        
    Else
    '�п�ݼ�
        ptrMenu.FireEvent MyItemInfo(itemID).itemKeyID, Left(MyItemInfo(itemID).itemText, Len(MyItemInfo(itemID).itemText) - 4)
        
    End If
Err:

End Sub

' �˵����¼���Ӧ(ѡ��˵���)
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
        '��ȡҪ���ƵĶ����ǲ˵�
        If .CtlType = ODT_MENU Then
        
            'ȷ����ǰ�ǲ��Ǹ�����ʾ
            blnHighlight = .itemState And ODS_SELECTED
            
            Select Case MyItemInfo(.itemID).itemType
                Case &H800            '�ָ�������
                    '����ͼ������
                    udtRect.Left = .rcItem.Left
                    udtRect.Top = .rcItem.Top
                    udtRect.Right = udtRect.Left + 21
                    udtRect.Bottom = .rcItem.Bottom + 5
                    lngBrush = CreateSolidBrush(&HD1D8DB)
                    FillRect .hdc, udtRect, lngBrush
                    DeleteObject lngBrush

                    '���Ʒָ���
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
                    '���Ʊ߿򣬵�ɫ
                    If blnHighlight And Not CBool(MyItemInfo(.itemID).itemState And MIS_DISABLED) Then   ' ���˵����ѡ��ʱ
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
                    
                    '����ͼ��
                    If MyItemInfo(.itemID).itemType = MIT_STRING Then
                        If MyItemInfo(.itemID).itemState = MIS_ENABLED Then
                        '�˵�״̬Ϊ����
                        
                            If blnHighlight Then
                            '�����ǰ��ѡ��״̬����ͼ���������ƣ�������
                                Call DrawPictureEx(.hdc, MyItemInfo(.itemID).itemIcon, .rcItem.Left + 2, .rcItem.Top + 2, 16, 16, dssNormal)
                            Else
                            '�����ǰ�Ƿ�ѡ��״̬����ͼ���������ƣ�������
                                Call DrawPictureEx(.hdc, MyItemInfo(.itemID).itemIcon, .rcItem.Left + 1, .rcItem.Top + 1, 16, 16, dssNormal)
                            End If
                            
                        Else
                        '������״̬��ͼ����
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
                    '�˵�״̬Ϊ����
                        If blnHighlight Then
                            '��ǰ��ѡ��״̬
                            SetTextColor .hdc, &H0
                        Else
                            SetTextColor .hdc, &H0
                             '��ǰ�Ƿ�ѡ��״̬
                       End If
                    Else
                        '������״̬�����ֱ��
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

'�Զ��庯��  ����ָ���ȡһ������
Public Function GetMenuItemFromPtr(ByVal lngPtr As Long) As cyMenu
   Dim oTemp As Object
   'ʵ�ּ�����ͨ��API����CopyMemory,ֱ�ӽ��������͵�ֵ����ָ������
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
    
    '���ԴΪ�գ������˳�
    If Source Is Nothing Then Exit Sub
    
    '�������������ΪSmooth
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
