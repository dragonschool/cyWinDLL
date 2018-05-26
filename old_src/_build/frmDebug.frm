VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "字符串/数组/文件分析"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   10305
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtString 
      Height          =   6015
      Left            =   3780
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   345
      Width           =   4140
   End
   Begin VB.ListBox List1 
      Height          =   6000
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3435
   End
   Begin VB.Label lblCount 
      Caption         =   "内容："
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblCount 
      Caption         =   "数组个数"
      Height          =   240
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public sDebugType As String     '记录调试模式，由dBugEx模块传入

Private Sub Form_Load()
    Call SendMessageByNum(List1.hWnd, &H194, 3000, ByVal 0&)

End Sub

Private Sub Form_Resize()
On Error Resume Next
    List1.Height = Me.Height - 1000
    txtString.Height = List1.Height
    txtString.Width = Me.Width - 4050

End Sub

Private Sub List1_Click()
On Error Resume Next

If frmDebug.sDebugType = "ListArrayString" Then
    txtString = Split(List1.List(List1.ListIndex), vbTab & "|" & vbTab)(1)
    
ElseIf frmDebug.sDebugType = "AppearString" Then
    txtString.SelLength = 1
    txtString.SelStart = List1.ListIndex
    txtString.SetFocus
    
End If

End Sub

Private Sub txtString_Click()
On Error Resume Next
Me.Caption = txtString.SelStart
List1.Selected(txtString.SelStart) = True
End Sub
