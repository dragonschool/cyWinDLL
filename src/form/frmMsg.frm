VERSION 5.00
Begin VB.Form frmMsg 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   Picture         =   "frmMsg.frx":0000
   ScaleHeight     =   1740
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrRemain 
      Interval        =   1000
      Left            =   1005
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1125
      Top             =   210
   End
   Begin VB.Timer tmrDelay 
      Left            =   1635
      Top             =   210
   End
   Begin VB.Timer tmrTransWnd 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   525
      Top             =   225
   End
   Begin VB.Label lblClose 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2415
      TabIndex        =   3
      Top             =   75
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "秒关闭"
      Height          =   195
      Left            =   1980
      TabIndex        =   2
      Top             =   1470
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblRemain 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1695
      TabIndex        =   1
      Top             =   1470
      Width           =   345
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   855
      Width           =   2475
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------
'窗体透明函数
'-----------------------------------------------------------------------------
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WM_CLOSE = &H10
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'-----------------------------------------------------------------------------
'窗体透明函数

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public sMsg As String

Private Sub Form_Load()
    Dim W As New formClass
    W.cyWndAction Me.hWnd, Wnd_TOPMOST, 1
    lblMsg = sMsg
    Me.Top = Screen.Height - Screen.TwipsPerPixelY * 30
    Me.Left = Screen.Width - Me.Width - 30
    
On Error Resume Next

    Dim bStr As String * 255
    
    
    PlaySound Replace(Mid(bStr, 1, GetWindowsDirectory(bStr, 255)) & "\media\notify.wav", "\\", "\"), ByVal 0&, SND_FILENAME Or SND_ASYNC

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim W As New formClass
    W.cyWndAction Me.hWnd, Wnd_DropToMove
End Sub

Private Sub Label2_Click()
    
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
Me.Top = Me.Top - 250
If Me.Top < Screen.Height - 20 * Screen.TwipsPerPixelY - Me.Height + 60 Then
    Timer1.Enabled = False
End If
End Sub

Private Sub tmrRemain_Timer()
    If tmrDelay.Interval = 0 Then
        tmrRemain.Enabled = False
        Unload Me
        Exit Sub
    End If
    lblRemain = lblRemain - 1
End Sub

Private Sub tmrTransWnd_Timer()
    On Error Resume Next
    Static i As Long
    If i > 250 Then
        i = 0
        tmrTransWnd.Enabled = False
        Unload Me
        Exit Sub
    End If
        i = i + 60
        Dim iOldWndStyle As Long
        iOldWndStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
        iOldWndStyle = iOldWndStyle Or WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, iOldWndStyle
        SetLayeredWindowAttributes hWnd, 0, 255 - i, LWA_ALPHA
End Sub
Private Sub tmrDelay_Timer()
    tmrTransWnd.Enabled = True
    tmrDelay.Enabled = False
End Sub

