VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "subClassClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum cyWindowMsg
    WM_NULL = &H0
    WM_CREATE = &H1
    WM_DESTROY = &H2
    WM_MOVE = &H3
    WM_SIZE = &H5
    WM_ACTIVATE = &H6
    WM_SETFOCUS = &H7
    WM_KILLFOCUS = &H8
    WM_ENABLE = &HA
    WM_SETREDRAW = &HB
    WM_SETTEXT = &HC
    WM_GETTEXT = &HD
    WM_GETTEXTLENGTH = &HE
    WM_PAINT = &HF
    WM_CLOSE = &H10
    WM_QUERYENDSESSION = &H11
    WM_QUIT = &H12
    WM_QUERYOPEN = &H13
    WM_ERASEBKGND = &H14
    WM_SYSCOLORCHANGE = &H15
    WM_ENDSESSION = &H16
    WM_SYSTEMERROR = &H17
    WM_SHOWWINDOW = &H18
    WM_ACTIVATEAPP = &H1C
    WM_FONTCHANGE = &H1D
    WM_TIMECHANGE = &H1E
    WM_CANCELMODE = &H1F
    WM_SETCURSOR = &H20
    WM_MOUSEACTIVATE = &H21
    WM_CHILDACTIVATE = &H22
    WM_QUEUESYNC = &H23
    WM_GETMINMAXINFO = &H24
    WM_PAINTICON = &H26
    WM_ICONERASEBKGND = &H27
    WM_NEXTDLGCTL = &H28
    WM_SPOOLERSTATUS = &H2A
    WM_DRAWITEM = &H2B
    WM_MEASUREITEM = &H2C
    WM_DELETEITEM = &H2D
    WM_VKEYTOITEM = &H2E
    WM_CHARTOITEM = &H2F
    WM_SETFONT = &H30
    WM_GETFONT = &H31
    WM_SETHOTKEY = &H32
    WM_GETHOTKEY = &H33
    WM_QUERYDRAGICON = &H37
    WM_COMPAREITEM = &H39
    WM_GETOBJECT = &H3D
    WM_COMPACTING = &H41
    WM_WINDOWPOSCHANGING = &H46
    WM_WINDOWPOSCHANGED = &H47
    WM_POWER = &H48
    WM_COPYDATA = &H4A
    WM_CANCELJOURNAL = &H4B
    WM_NOTIFY = &H4E
    WM_INPUTLANGCHANGEREQUEST = &H50
    WM_INPUTLANGCHANGE = &H51
    WM_TCARD = &H52
    WM_HELP = &H53
    WM_USERCHANGED = &H54
    WM_NOTIFYFORMAT = &H55
    WM_CONTEXTMENU = &H7B
    WM_STYLECHANGING = &H7C
    WM_STYLECHANGED = &H7D
    WM_DISPLAYCHANGE = &H7E
    WM_GETICON = &H7F
    WM_SETICON = &H80
    WM_NCCREATE = &H81
    WM_NCDESTROY = &H82
    WM_NCCALCSIZE = &H83
    WM_NCHITTEST = &H84
    WM_NCPAINT = &H85
    WM_NCACTIVATE = &H86
    WM_GETDLGCODE = &H87
    WM_NCMOUSEMOVE = &HA0
    WM_NCLBUTTONDOWN = &HA1
    WM_NCLBUTTONUP = &HA2
    WM_NCLBUTTONDBLCLK = &HA3
    WM_NCRBUTTONDOWN = &HA4
    WM_NCRBUTTONUP = &HA5
    WM_NCRBUTTONDBLCLK = &HA6
    WM_NCMBUTTONDOWN = &HA7
    WM_NCMBUTTONUP = &HA8
    WM_NCMBUTTONDBLCLK = &HA9
    WM_KEYFIRST = &H100
    WM_KEYDOWN = &H100
    WM_KEYUP = &H101
    WM_CHAR = &H102
    WM_DEADCHAR = &H103
    WM_SYSKEYDOWN = &H104
    WM_SYSKEYUP = &H105
    WM_SYSCHAR = &H106
    WM_SYSDEADCHAR = &H107
    WM_INITDIALOG = &H110
    WM_COMMAND = &H111
    WM_SYSCOMMAND = &H112
    WM_TIMER = &H113
    WM_HSCROLL = &H114
    WM_VSCROLL = &H115
    WM_INITMENU = &H116
    WM_INITMENUPOPUP = &H117
    WM_MENUSELECT = &H11F
    WM_MENUCHAR = &H120
    WM_ENTERIDLE = &H121
    WM_MENURBUTTONUP = &H122
    WM_MENUCOMMAND = &H126
    WM_CTLCOLORMSGBOX = &H132
    WM_CTLCOLOREDIT = &H133
    WM_CTLCOLORLISTBOX = &H134
    WM_CTLCOLORBTN = &H135
    WM_CTLCOLORDLG = &H136
    WM_CTLCOLORSCROLLBAR = &H137
    WM_CTLCOLORSTATIC = &H138
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDOWN1 = &H201
    WM_LBUTTONUP1 = &H202
    WM_LBUTTONDBLCLK1 = &H203
    WM_RBUTTONDOWN1 = &H204
    WM_RBUTTONUP1 = &H205
    WM_RBUTTONDBLCLK1 = &H206
    WM_MBUTTONDOWN1 = &H207
    WM_MBUTTONUP1 = &H208
    WM_MBUTTONDBLCLK1 = &H209
    WM_MOUSEWHEEL = &H20A
    WM_PARENTNOTIFY = &H210
    WM_ENTERMENULOOP = &H211
    WM_EXITMENULOOP = &H212
    WM_NEXTMENU = &H213
    WM_SIZING = 532
    WM_CAPTURECHANGED = 533
    WM_MOVING = 534
    WM_POWERBROADCAST = 536
    WM_DEVICECHANGE = 537
    WM_IME_STARTCOMPOSITION = &H10D
    WM_IME_ENDCOMPOSITION = &H10E
    WM_IME_COMPOSITION = &H10F
    WM_IME_KEYLAST = &H10F
    WM_IME_SETCONTEXT = &H281
    WM_IME_NOTIFY = &H282
    WM_IME_CONTROL = &H283
    WM_IME_COMPOSITIONFULL = &H284
    WM_IME_SELECT = &H285
    WM_IME_CHAR = &H286
    WM_IME_REQUEST = &H288
    WM_IME_KEYDOWN = &H290
    WM_IME_KEYUP = &H291
    WM_MDICREATE = &H220
    WM_MDIDESTROY = &H221
    WM_MDIACTIVATE = &H222
    WM_MDIRESTORE = &H223
    WM_MDINEXT = &H224
    WM_MDIMAXIMIZE = &H225
    WM_MDITILE = &H226
    WM_MDICASCADE = &H227
    WM_MDIICONARRANGE = &H228
    WM_MDIGETACTIVE = &H229
    WM_MDISETMENU = &H230
    WM_ENTERSIZEMOVE = &H231
    WM_EXITSIZEMOVE = &H232
    WM_DROPFILES = &H233
    WM_MDIREFRESHMENU = &H234
    WM_MOUSEHOVER = &H2A1
    WM_MOUSELEAVE = &H2A3
    WM_CUT = &H300
    WM_COPY = &H301
    WM_PASTE = &H302
    WM_CLEAR = &H303
    WM_UNDO = &H304
    WM_RENDERFORMAT = &H305
    WM_RENDERALLFORMATS = &H306
    WM_DESTROYCLIPBOARD = &H307
    WM_DRAWCLIPBOARD = &H308
    WM_PAINTCLIPBOARD = &H309
    WM_VSCROLLCLIPBOARD = &H30A
    WM_SIZECLIPBOARD = &H30B
    WM_ASKCBFORMATNAME = &H30C
    WM_CHANGECBCHAIN = &H30D
    WM_HSCROLLCLIPBOARD = &H30E
    WM_QUERYNEWPALETTE = &H30F
    WM_PALETTEISCHANGING = &H310
    WM_PALETTECHANGED = &H311
    WM_HOTKEY = &H312
    WM_PRINT = 791
    WM_PRINTCLIENT = 792
    WM_HANDHELDFIRST = 856
    WM_HANDHELDLAST = 863
    WM_PENWINFIRST = &H380
    WM_PENWINLAST = &H38F
    WM_COALESCE_FIRST = &H390
    WM_COALESCE_LAST = &H39F
    WM_DDE_FIRST = &H3E0
    WM_DDE_INITIATE = &H3E0
    WM_DDE_TERMINATE = &H3E1
    WM_DDE_ADVISE = &H3E2
    WM_DDE_UNADVISE = &H3E3
    WM_DDE_ACK = &H3E4
    WM_DDE_REQUEST = &H3E6
    WM_DDE_POKE = &H3E0 + 7
    WM_DDE_EXECUTE = &H3E8
    WM_DDE_LAST = &H3E8
    WM_APP = &H8000
    WM_USER = &H400
End Enum

Public Event WindowMsgEvent(ByVal sMsg As String, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Public Event SubClassEvent(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private m_SubClassHwnd As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)

Public Sub SetSubClass(ByVal iSubClassHwnd As Long, Optional ByVal uMsg1 As cyWindowMsg, Optional ByVal uMsg2 As cyWindowMsg, Optional ByVal uMsg3 As cyWindowMsg)
    modSubClass.WinMsg1 = uMsg1
    modSubClass.WinMsg2 = uMsg2
    modSubClass.WinMsg3 = uMsg3
    m_SubClassHwnd = iSubClassHwnd
    modSubClass.objSubClass = ObjPtr(Me)

    If uMsg1 = 0 And uMsg2 = 0 And uMsg3 = 0 Then
    '捕捉所有消息
        preSubClassProc = SetWindowLong(iSubClassHwnd, GWL_WNDPROC, AddressOf WndProcSubClassAllMsg)
    
    Else
    '捕捉指定消息
        preSubClassProc = SetWindowLong(iSubClassHwnd, GWL_WNDPROC, AddressOf WndProcSubClass)
    
    End If
    
End Sub

Private Sub Class_Terminate()
    Call SetWindowLong(m_SubClassHwnd, GWL_WNDPROC, preSubClassProc)

End Sub

Friend Function FireEvent(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    On Error Resume Next
    RaiseEvent WindowMsgEvent(LoadResString(uMsg), uMsg, wParam, lParam)
    
End Function
