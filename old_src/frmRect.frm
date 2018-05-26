VERSION 5.00
Begin VB.Form frmRect 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000003&
   ClientHeight    =   3405
   ClientLeft      =   4935
   ClientTop       =   3735
   ClientWidth     =   4140
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmRect.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   227
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   276
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Image img 
      Height          =   285
      Index           =   0
      Left            =   30
      Picture         =   "frmRect.frx":1232
      ToolTipText     =   "关闭窗口"
      Top             =   30
      Width           =   285
   End
   Begin VB.Image img1 
      Height          =   285
      Index           =   0
      Left            =   30
      Picture         =   "frmRect.frx":390B
      ToolTipText     =   "关闭窗口"
      Top             =   30
      Width           =   285
   End
   Begin VB.Image imgTitle 
      Height          =   360
      Left            =   0
      Picture         =   "frmRect.frx":5FA7
      Top             =   0
      Width           =   15450
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuCapture 
         Caption         =   "Shift+F8范围截图"
         Index           =   0
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "Shift+F9全屏截图"
         Index           =   1
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "Shift+F10存到桌面"
         Index           =   2
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "Shift+F11保存剪贴板"
         Index           =   3
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "关闭"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim WithEvents Tray As cyTrayEx
Attribute Tray.VB_VarHelpID = -1
Dim WithEvents HotKey As cyHotKeyEx
Attribute HotKey.VB_VarHelpID = -1

Dim iClick As Long
Dim m_sPath As String
Dim WithEvents P As cyPhotoEx
Attribute P.VB_VarHelpID = -1

Public bShowTray As Boolean

Private Sub Form_Load()
           
    Set P = New cyPhotoEx
    Dim F As New cyFileEx
    m_sPath = F.cyGetSpecialFolder(DeskTop)
    
    Set HotKey = New cyHotKeyEx
    HotKey.cySetHotKeyEx 100001, Me.hWnd, , True, , vbKeyF8
    HotKey.cySetHotKeyEx 100002, Me.hWnd, , True, , vbKeyF9
    HotKey.cySetHotKeyEx 100003, Me.hWnd, , True, , vbKeyF10
    HotKey.cySetHotKeyEx 100004, Me.hWnd, , True, , vbKeyF11
    
    If bShowTray = True Then
        Set Tray = New cyTrayEx
        Tray.SetTray Me.hWnd, Me.Icon.Handle, "ScreenCapture", "Shift+F8范围截图" & vbTab & "Shift+F9全屏截图" & vbCrLf & "Shift+F10存到桌面" & vbTab & "Shift+F11保存剪贴板", 1, 10
        
    End If
    
    Me.BackColor = &HE9967A
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H2 Or &H1
    img(0).ZOrder 0
    
    On Error Resume Next
    Dim i1 As Long
    Dim i2 As Long
    Dim i3 As Long
    Dim i4 As Long
    '保存窗口位置
    i1 = GetSetting("cyDLL", "ScreenCapture", "Left")
    i2 = GetSetting("cyDLL", "ScreenCapture", "Top")
    i3 = GetSetting("cyDLL", "ScreenCapture", "Width")
    i4 = GetSetting("cyDLL", "ScreenCapture", "Height")
    If i3 <= 0 Or i4 <= 0 Then
    '如果数值不正确则设置一个缺少数值
        i3 = 4000: i4 = 3000
        
    End If
    
    Me.Move i1, i2, i3, i4
    
End Sub

Private Sub Form_Resize()
    Const RGN_DIFF = 4
    
    Dim outer_rgn As Long
    Dim inner_rgn As Long
    Dim combined_rgn As Long
    Dim wid As Single
    Dim hgt As Single
    Dim border_width As Single
    Dim title_height As Single
    
    wid = ScaleX(Width, vbTwips, vbPixels)
    hgt = ScaleY(Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wid, hgt)
    
    border_width = (wid - ScaleWidth) / 2
    title_height = hgt - border_width - ScaleHeight
    
    inner_rgn = CreateRectRgn(4, 28, wid - 4, hgt - 4)

    combined_rgn = CreateRectRgn(0, 0, 0, 0)
    CombineRgn combined_rgn, outer_rgn, _
        inner_rgn, RGN_DIFF
    
    SetWindowRgn hWnd, combined_rgn, True
        
    '保存窗口位置
    SaveSetting "cyDLL", "ScreenCapture", "Left", Me.Left
    SaveSetting "cyDLL", "ScreenCapture", "Top", Me.Top
    SaveSetting "cyDLL", "ScreenCapture", "Width", Me.Width
    SaveSetting "cyDLL", "ScreenCapture", "Height", Me.Height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set HotKey = Nothing
    
    If bShowTray = True Then
        
        Tray.RemoveTray
        Set Tray = Nothing
    
    End If
    Unload Me
    
End Sub


Private Sub HotKey_cyHotKeyEventEx(ByVal IDHotKey As Long)

On Error Resume Next
Static i As Long

'间隔时间小于500则表示
If (GetTickCount - i) < 300 Then
    i = GetTickCount
    Exit Sub
End If

Select Case IDHotKey
    Case 100001
        P.cyCursorShow
        P.cyRectToClipboard Me.Left / Screen.TwipsPerPixelX + 3, Me.Top / Screen.TwipsPerPixelY + 27, Me.Width / Screen.TwipsPerPixelX - 6, Me.Height / Screen.TwipsPerPixelY - 30
        P.cyCursorHide
        modPhoto.CapturePhoto
        
    Case 100002
        Dim iLeft As Long
        iLeft = frmRect.Left
        frmRect.Left = -8000
        DoEvents
        P.cyCursorShow
        P.cyRectToClipboard 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY
        P.cyCursorHide
        DoEvents
        frmRect.Left = iLeft
        modPhoto.CapturePhoto
        
    Case 100003
        Static iFileNum As Long
        P.cyCursorShow
        P.cyRectToClipboard Me.Left / Screen.TwipsPerPixelX + 3, Me.Top / Screen.TwipsPerPixelY + 27, Me.Width / Screen.TwipsPerPixelX - 6, Me.Height / Screen.TwipsPerPixelY - 30
        P.cyCursorHide
        iFileNum = iFileNum + 1
    
        Call P.cyClipBoardSaveToJpg(m_sPath & Format(iFileNum, "000") & ".Jpg")
    
    Case 100004
        Dim sFileName As String
        Dim F As New cyFileEx
        sFileName = F.cyDialogSave(Me.hWnd, "保存", "*.Bmp|*.Bmp|*.Jpg|*.jpg")
        If sFileName = "" Then Exit Sub
        If UCase(Right(sFileName, 3)) = "BMP" Then
            P.cyClipBoardSaveToBmp sFileName
            
        ElseIf UCase(Right(sFileName, 3)) = "JPG" Then
            P.cyClipBoardSaveToJpg sFileName
            
        End If
        Set F = Nothing
        
End Select

i = GetTickCount

End Sub

Private Sub img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    iClick = Index
    img1(Index).ZOrder 0
End Sub

Private Sub img1_Click(Index As Integer)
    If Index = 0 Then
        Unload Me
    End If
End Sub

Private Sub imgTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If iClick > 0 Then
        img(0).ZOrder 0
        iClick = 0
    End If
    
    Const WM_SYSCOMMAND = &H112
    Const SC_MOVE = &HF010&
    Const HTCAPTION = 2
    ReleaseCapture
    PostMessage hWnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0

End Sub

Private Sub mnuCapture_Click(Index As Integer)
    If Index = 0 Then
        HotKey_cyHotKeyEventEx 100001
        
    ElseIf Index = 1 Then
        HotKey_cyHotKeyEventEx 100002
    
    ElseIf Index = 2 Then
        HotKey_cyHotKeyEventEx 100003
        
    ElseIf Index = 3 Then
        HotKey_cyHotKeyEventEx 100004
        
    ElseIf Index = 5 Then
        Form_Unload 0
        
    End If
    
End Sub

Private Sub Tray_cyRightButtonUp()
    PopupMenu mnuPopup
    
End Sub
