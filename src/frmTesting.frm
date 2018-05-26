VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Begin VB.Form frmTesting 
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   1905
   ClientTop       =   1920
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   5955
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   3120
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin ComctlLib.ListView lv 
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3413
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   3000
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   3330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4895
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      DragIcon        =   "frmTesting.frx":0000
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "frmTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents winSockClass As winSockClass
Attribute winSockClass.VB_VarHelpID = -1

Private Sub Form_Load()

Dim fileClass As New fileClass
fileClass.cyXCOPY "E:\old file", "c:\a"



Dim dBug As New debugClass
'dBug.cyShowHwnd

End Sub
