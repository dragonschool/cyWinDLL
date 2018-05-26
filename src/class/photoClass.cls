VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "photoClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'捕捉截图事件
Public Event cyCaptureEvent()

'-----------------------------------------------------------------------------
'改变墙纸
'-----------------------------------------------------------------------------
Private Declare Function SystemParametersInfo& Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long)
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPIF_UPDATEINIFILE = &H1

Enum cyWallPaperMode
    LaShen = 0
    JuZhong = 1
    PingPu = 2
End Enum
'-----------------------------------------------------------------------------
'改变墙纸
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
'转换成JPG
'-----------------------------------------------------------------------------
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
   GdiplusVersion As Long
   DebugEventCallback As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
   GUID As GUID
   NumberOfValues As Long
   type As Long
   Value As Long
End Type

Private Type EncoderParameters
   Count As Long
   Parameter As EncoderParameter
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, _
   inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long

Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" ( _
   ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As Long

Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long

Private Declare Function GdipSaveImageToFile Lib "GDIPlus" ( _
   ByVal Image As Long, ByVal filename As Long, _
   clsidEncoder As GUID, encoderParams As Any) As Long

Private Declare Function CLSIDFromString Lib "ole32" ( _
   ByVal str As Long, ID As GUID) As Long

'-----------------------------------------------------------------------------
'转换成JPG
'-----------------------------------------------------------------------------

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim ret As Long

'Private Type RGBTRIPLE
'     rgbRed As Byte
'     rgbGreen As Byte
'     rgbBlue As Byte
'End Type

'Private Type BITMAPINFO256
'     bmiHeader As BITMAPINFOHEADER
'     bmiColors(0 To 255) As RGBQUAD
'End Type

Enum cyPrint
    PrintFullSceen = 0
    PrintActiveForm = 1
End Enum

Enum cyCaptureType
    CaptureFullScreen = 0
    CaptureRangle = 1
End Enum

Private Const SRCCOPY              As Long = &HCC0020
Private Const CF_BITMAP            As Long = 2
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private dsk_Wnd    As Long                                 ' Desktop window
Private dsk_hDC    As Long                                 ' Desktop DC

'-----------------------------------------------------------------------------
'剪贴板到StdPicture
'-----------------------------------------------------------------------------

'Window device
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

'clip board
'Private Const CF_BITMAP = 2
Private Const CF_PALETTE = 9
'Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
'Private Declare Function CloseClipboard Lib "user32.dll" () As Long

'Get a Picture from a handle.
Private Const vbPicTypeBitmap = 1
Private Type TPICTDESC
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    Option1 As Long
    Option2 As Long
End Type

Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(1 To 8) As Byte
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
                                                  (lpPictDesc As TPICTDESC, _
                                                   RefIID As TGUID, _
                                                   ByVal fPictureOwnsHandle As Long, _
                                                   ByRef IPic As IPicture) As Long

'-----------------------------------------------------------------------------
'剪贴板到StdPicture
'-----------------------------------------------------------------------------

Public Function cyPrintScreenToClipBoard(Optional ByVal PrintMode As cyPrint)
        Call keybd_event(vbKeySnapshot, PrintMode, 0, 0)
End Function

Sub cyClipBoardToPictureBox(ByVal sControlName As Object)
    sControlName.Picture = Clipboard.GetData(vbCFBitmap)
End Sub

Sub cyCursorShow()
    Dim curPos As POINTAPI

    '得到当前当标位置
    GetCursorPos curPos
    '画出模拟光标
    Dim ShapeV(9) As POINTAPI
    ShapeV(0).X = curPos.X:     ShapeV(0).Y = curPos.Y
    ShapeV(1).X = curPos.X:      ShapeV(1).Y = curPos.Y + 16
    ShapeV(2).X = curPos.X + 4:    ShapeV(2).Y = curPos.Y + 12
    ShapeV(3).X = curPos.X + 8:    ShapeV(3).Y = curPos.Y + 20
    ShapeV(4).X = curPos.X + 9:    ShapeV(4).Y = curPos.Y + 20
    ShapeV(5).X = curPos.X + 10:    ShapeV(5).Y = curPos.Y + 19
    ShapeV(6).X = curPos.X + 7:    ShapeV(6).Y = curPos.Y + 11
    ShapeV(7).X = curPos.X + 11:    ShapeV(7).Y = curPos.Y + 11
    ShapeV(8).X = curPos.X:    ShapeV(8).Y = curPos.Y
    ret = GetDC(&H0)
    Polygon ret, ShapeV(0), 8
    
End Sub

Sub cyCursorHide()
    Const RDW_ALLCHILDREN = &H80
    Const RDW_INVALIDATE = &H1
    RedrawWindow &H0, ByVal 0&, ByVal 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
End Sub

Function cyRectToClipboard(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Boolean

    Dim clp_hDC     As Long
    Dim clp_hObj    As Long
    Dim clp_hOldObj As Long
    '-- Get desktop DC
    dsk_Wnd = GetDesktopWindow
    dsk_hDC = GetWindowDC(dsk_Wnd)

        '-- Create DC
        clp_hDC = CreateCompatibleDC(dsk_hDC)
        If (clp_hDC <> 0) Then
            clp_hObj = CreateCompatibleBitmap(dsk_hDC, Width, Height)
            If (clp_hObj <> 0) Then

                '-- Set created bitmap as current Clipboard object
                clp_hOldObj = SelectObject(clp_hDC, clp_hObj)
                '-- Paint selected area
                BitBlt clp_hDC, 0, 0, Width, Height, dsk_hDC, X, Y, SRCCOPY
                '-- Restore last Clipboard object
                SelectObject clp_hDC, clp_hOldObj

                If (OpenClipboard(0) <> 0) Then
                    '-- Empty clipboard, set data and close it
                    EmptyClipboard
                    SetClipboardData CF_BITMAP, clp_hObj
                    CloseClipboard
                    '-- Success
                    cyRectToClipboard1 = -1
                End If
            End If
            '-- Delete our temp. DC
            DeleteDC clp_hDC
        End If
End Function

Sub cyShowCaptureTray()
    frmRect.bShowTray = True
End Sub

Sub cyShowCaptureWnd()
    frmRect.Show
End Sub

Sub cyClipBoardSaveToBmp(ByVal sFileName As String)
    SavePicture Clipboard.GetData(vbCFBitmap), sFileName
End Sub

Sub cyBmpSaveToJpg(ByVal sBmpFileName As String, ByVal sJpgFileName As String, Optional ByVal Quality As Byte = 80)
    Dim StdPic As StdPicture
    Set StdPic = LoadPicture(sBmpFileName)
    cyStdPictureSaveToJpg StdPic, sJpgFileName, Quality
End Sub

Sub cyClipBoardSaveToJpg(ByVal sJpgFileName As String, Optional ByVal Quality As Byte = 80)
    cyStdPictureSaveToJpg cyClipBoardToStdPicture, sJpgFileName, Quality
End Sub


'-----------------------------------------------------------------------------
'剪贴板到StdPicture
'-----------------------------------------------------------------------------
Public Function cyClipBoardToStdPicture() As StdPicture
    Dim hBMP As Long
    Dim hPalette As Long
    Dim TPICTDESC As TPICTDESC
    Dim TGUID As TGUID

    Set cyClipBoardToStdPicture = Nothing

    If IsClipboardFormatAvailable(CF_BITMAP) = False Then Exit Function
    If OpenClipboard(CLng(0)) = False Then Exit Function

    hBMP = GetClipboardData(CF_BITMAP)
    hPalette = GetClipboardData(CF_PALETTE)
    Call CloseClipboard
    If hBMP = 0 Then Exit Function

    With TPICTDESC
        .cbSizeofStruct = Len(TPICTDESC)
        .picType = vbPicTypeBitmap
        .hImage = hBMP
        .Option1 = hPalette
    End With
    With TGUID
        .Data1 = &H20400
        .Data4(1) = &HC0
        .Data4(8) = &H46
    End With
    Call OleCreatePictureIndirect(TPICTDESC, TGUID, True, cyClipBoardToStdPicture)
End Function

Public Sub cyStdPictureSaveToJpg(ByVal StdPic As StdPicture, ByVal sFileName As String, Optional ByVal Quality As Byte = 80)
   Dim tSI As GdiplusStartupInput
   Dim lRes As Long
   Dim lGDIP As Long
   Dim lBitmap As Long

   ' 初始化 GDI+
   tSI.GdiplusVersion = 1
   lRes = GdiplusStartup(lGDIP, tSI)
   
   If lRes = 0 Then
   
      ' 从句柄创建 GDI+ 图像
      lRes = GdipCreateBitmapFromHBITMAP(StdPic.Handle, 0, lBitmap)
   
      If lRes = 0 Then
         Dim tJpgEncoder As GUID
         Dim tParams As EncoderParameters
         
         '初始化解码器的GUID标识
         CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), _
                         tJpgEncoder
      
         ' 设置解码器参数
         tParams.Count = 1
         With tParams.Parameter ' Quality
            ' 得到Quality参数的GUID标识
            CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
            .NumberOfValues = 1
            .type = 4
            .Value = VarPtr(Quality)
         End With
         
         '保存图像
         lRes = GdipSaveImageToFile( _
                  lBitmap, _
                  StrPtr(sFileName), _
                  tJpgEncoder, _
                  tParams)
                             
         ' 销毁GDI+图像
         GdipDisposeImage lBitmap
         
      End If
      
      '销毁 GDI+
      GdiplusShutdown lGDIP

   End If
   
   If lRes Then
      Err.Raise 5, , "不能保存这个图像：" & lRes
   End If
   
End Sub

'函数名称:cyChangeWallPaper
'作用：改变墙纸（有３种显示方式,分别为拉伸、居中、平铺）
'===========================================================
'整理时间：2007-03-17 22:40
'  参数：
'  例子：
'    Dim P As New photoClass
'    P.cyChangeWallPaper "E:\My Pictures\04.07家居照片\001.jpg", LaShen

'===========================================================
Sub cyChangeWallPaper(ByVal sPicFileName As String, Optional ByVal Mode As cyWallPaperMode) '改变墙纸

    Dim sTempPath As String * 255
    Dim sTempBmp As String
    Dim R As New registClass
    
    '取得临时目录
    GetTempPath 255, sTempPath
    '取得临时图片文件名
    sTempBmp = Left$(sTempPath, InStr(1, sTempPath, Chr(0)) - 1) & "WallPaper.Bmp"
    
    SavePicture LoadPicture(sPicFileName), sTempBmp
    
    Select Case Mode
        Case 0
            R.cySetValueSZ REG_SZ, "HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", "0"
            R.cySetValueSZ REG_SZ, "HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", "2"
        Case 1
            R.cySetValueSZ REG_SZ, "HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", "0"
            R.cySetValueSZ REG_SZ, "HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", "0"
        Case 2
            R.cySetValueSZ REG_SZ, "HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", "1"
            R.cySetValueSZ REG_SZ, "HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", "0"
        
    End Select
    Set R = Nothing
    
    '执行改变墙纸的行动
    SystemParametersInfo ByVal SPI_SETDESKWALLPAPER, True, ByVal sTempBmp, SPIF_UPDATEINIFILE
    
End Sub

'函数名称:cyPrintBarCode
'作用：在指定位置打印条形码
'===========================================================
'整理时间：2007-03-18 22:05
'  参数：
        'strBarCode - 要打印的条形码字符串
        'intXPos, intYPos - 打印条形码的左上角坐标(缺省为(0,0),坐标刻度为:毫米)
        'intHeight - 打印高度(缺省为一厘米,坐标刻度为:毫米)
        'bPrintText - 是否打印人工识别字符(缺省为true)

'  例子：
'    Dim P As New photoClass
'    P.cyPrintBarCode "980", 0, 60, 10, True, 9
'===========================================================

Sub cyPrintBarCode(ByVal strBarCode As String, _
                         Optional ByVal intXPos As Integer = 0, _
                         Optional ByVal intYPos As Integer = 0, _
                         Optional ByVal intPrintHeight As Integer = 10, _
                         Optional ByVal bPrintText As Boolean = True, _
                         Optional ByVal bTextSize As Long = 9)
    
    If strBarCode = "" Then Exit Sub '不打印空串
    
    '"0-9","A-Z","-","%","$"和"*" 的条码编码格式,总共 40 个字符
    Static strBarTable(39) As String
    '初始化条码编码格式表
    strBarTable(0) = "001100100"     '0
    strBarTable(1) = "100010100"     '1
    strBarTable(2) = "010010100"     '2
    strBarTable(3) = "110000100"     '3
    strBarTable(4) = "001010100"     '4
    strBarTable(5) = "101000100"     '5
    strBarTable(6) = "011000100"     '6
    strBarTable(7) = "000110100"     '7
    strBarTable(8) = "100100100"     '8
    strBarTable(9) = "010100100"     '9
    strBarTable(10) = "100010010"    'A
    strBarTable(11) = "010010010"    'B
    strBarTable(12) = "110000010"    'C
    strBarTable(13) = "001010010"    'D
    strBarTable(14) = "101000010"    'E
    strBarTable(15) = "011000010"    'F
    strBarTable(16) = "000110010"    'G
    strBarTable(17) = "100100010"    'H
    strBarTable(18) = "010100010"    'I
    strBarTable(19) = "001100010"    'J
    strBarTable(20) = "100010001"    'K
    strBarTable(21) = "010010001"    'L
    strBarTable(22) = "110000001"    'M
    strBarTable(23) = "001010001"    'N
    strBarTable(24) = "101000001"    'O
    strBarTable(25) = "011000001"    'P
    strBarTable(26) = "000110001"    'Q
    strBarTable(27) = "100100001"    'R
    strBarTable(28) = "010100001"    'S
    strBarTable(29) = "001100001"    'T
    strBarTable(30) = "100011000"    'U
    strBarTable(31) = "010011000"    'V
    strBarTable(32) = "110001000"    'W
    strBarTable(33) = "001011000"    'X
    strBarTable(34) = "101001000"    'Y
    strBarTable(35) = "011001000"    'Z
    strBarTable(36) = "000111000"    '-
    strBarTable(37) = "100101000"    '%
    strBarTable(38) = "010101000"    '$
    strBarTable(39) = "001101000"    '*

    '保存打印机 ScaleMode
    Dim intOldScaleMode As ScaleModeConstants
    intOldScaleMode = Printer.ScaleMode

    '保存打印机 DrawWidth
    Dim intOldDrawWidth As Integer
    intOldDrawWidth = Printer.DrawWidth

    '保存打印机 Font
    Dim fntOldFont As StdFont
    Set fntOldFont = Printer.Font

    '设置打印用的坐标刻度为缇(twip=1)
    Printer.ScaleMode = vbTwips
    
    '线宽为 1
    Printer.DrawWidth = 1
    
    '打印在条码下方字符的字体和大小
    Printer.FontName = "宋体"
    Printer.FontSize = bTextSize

    Dim strBC As String             '要打印的条码字符串
    strBC = UCase(strBarCode)

    '将以毫米表示的 X 坐标转换为以缇表示
    Dim X As Integer
    X = Printer.ScaleX(intXPos, vbMillimeters, vbTwips)

    '将以毫米表示的 Y 坐标转换为以缇表示
    Dim Y As Integer
    Y = Printer.ScaleY(intYPos, vbMillimeters, vbTwips)

    '将以毫米表示的高度转换为以缇表示
    Dim intHeight As Integer
    intHeight = Printer.ScaleY(intPrintHeight, vbMillimeters, vbTwips)
   
    '是否在条形码下方打印人工识别字符
    If bPrintText = True Then
        '条码打印高度要减去下面的字符显示高度
        intHeight = intHeight - Printer.TextHeight(strBC)
    End If

    '粗线和宽间隙宽度
    Const intWidthCU As Integer = 30
    
    '细线和窄间隙宽度
    Const intWidthXI As Integer = 10
    
    '当前处理的字符串索引
    Dim intIndex As Integer
    
    '循环控制变量
    Dim i As Integer, j As Integer, k As Integer
    
    '循环处理每个要显示的条码字符
    For i = 1 To Len(strBC)
        
        '确定当前字符在 strBarTable 中的索引
        Select Case Mid(strBC, i, 1)
            Case "*": intIndex = 39
            Case "$": intIndex = 38
            Case "%": intIndex = 37
            Case "-": intIndex = 36
            Case "0" To "9": intIndex = CInt(Mid(strBC, i, 1))
            Case "A" To "Z": intIndex = Asc(Mid(strBC, i, 1)) - Asc("A") + 10
            Case Else
                Call MsgBox("打印内容包含非法字符！", vbCritical Or vbSystemModal, "")
                
        End Select
        
        '是否在条形码下方打印人工识别字符
        If bPrintText = True Then
            Printer.CurrentX = X
            Printer.CurrentY = Y + intHeight
            Printer.Print Mid(strBC, i, 1)
        End If

        For j = 1 To 5
            If Mid(strBarTable(intIndex), j, 1) = "0" Then
                '画细线
                For k = 0 To intWidthXI - 1
                    Printer.Line (X + k, Y)-Step(0, intHeight)
                Next k
                X = X + intWidthXI
            Else
                '画宽线
                For k = 0 To intWidthCU - 1
                    Printer.Line (X + k, Y)-Step(0, intHeight)
                Next k
                X = X + intWidthCU
            End If

            '每个字符条码之间为窄间隙
            If j = 5 Then
                X = X + intWidthXI * 3
                Exit For
            End If
            
            If Mid(strBarTable(intIndex), j + 5, 1) = "0" Then
                '窄间隙
                X = X + intWidthXI * 3
            Else
                '宽间隙
                X = X + intWidthCU * 2
            End If
        Next j
    Next i
    
    '开始打印
    Printer.EndDoc
    
    '恢复打印机 ScaleMode
    Printer.ScaleMode = intOldScaleMode
    
    '恢复打印机 DrawWidth
    Printer.DrawWidth = intOldDrawWidth
    
    '恢复打印机 Font
    Set Printer.Font = fntOldFont
    
End Sub

Friend Function FireEvent()
    '捕捉到激活事件
    RaiseEvent cyCaptureEvent
      
End Function

Private Sub Class_Initialize()
    
    '获得本身的objprt
    modPhoto.objPhoto = ObjPtr(Me)
    
End Sub
