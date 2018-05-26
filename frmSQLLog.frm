VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmSQLLog 
   Caption         =   "SQL语句历史记录"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   Icon            =   "frmSQLLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8070
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd 
      Caption         =   "删除"
      Height          =   380
      Index           =   2
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "删除当前SQL语句"
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Caption         =   "关闭"
      Height          =   380
      Index           =   1
      Left            =   7320
      TabIndex        =   2
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Caption         =   "应用"
      Height          =   380
      Index           =   0
      Left            =   6480
      TabIndex        =   1
      ToolTipText     =   "应用当前SQL语句"
      Top             =   5880
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid dgSQL 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1.3
      RowHeight       =   51
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "SQL语句:"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmSQLLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        frmDebugDB.txtExec = dgSQL.Text
        frmDebugDB.SetFocus
        frmDebugDB.txtExec.SetFocus
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        Select Case MsgBox("是否确定要删除此SQL记录？", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton2, "")
           Case vbYes
                Call frmDebugDB.DeleteSQL
           Case vbNo
        
        End Select
    End If

End Sub

Private Sub dgSQL_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 Then
        Dim S As New cySystemEx
        Call frmDebugDB.DeleteSQL
    End If

End Sub

