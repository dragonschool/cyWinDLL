VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Begin VB.Form frmDebugDB 
   Caption         =   "数据查询器"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmDebugDB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   658
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   Tag             =   "2004-3-1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Height          =   4110
      Left            =   120
      ScaleHeight     =   4050
      ScaleWidth      =   8640
      TabIndex        =   19
      Top             =   120
      Width           =   8700
      Begin VB.CommandButton cmd 
         Caption         =   "连接"
         Height          =   375
         Index           =   2
         Left            =   5760
         TabIndex        =   43
         ToolTipText     =   "连接曾访问的数据库"
         Top             =   3480
         Width           =   690
      End
      Begin VB.CommandButton cmd 
         Caption         =   "删除"
         Height          =   375
         Index           =   3
         Left            =   6960
         TabIndex        =   42
         ToolTipText     =   "删除此记录"
         Top             =   3480
         Width           =   690
      End
      Begin VB.CommandButton cmd 
         Caption         =   "取消"
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   41
         ToolTipText     =   "删除此记录"
         Top             =   3480
         Width           =   690
      End
      Begin VB.PictureBox Picture2 
         Height          =   1005
         Left            =   120
         ScaleHeight     =   945
         ScaleWidth      =   8355
         TabIndex        =   29
         Top             =   120
         Width           =   8415
         Begin VB.CommandButton cmdConnectDB 
            Caption         =   "连接"
            Height          =   380
            Left            =   7680
            TabIndex        =   36
            ToolTipText     =   "连接SQL数据库"
            Top             =   510
            Width           =   615
         End
         Begin VB.TextBox txtID 
            Height          =   300
            Left            =   2445
            TabIndex        =   35
            Top             =   555
            Width           =   1185
         End
         Begin VB.ComboBox cboSQL 
            Height          =   300
            Left            =   120
            TabIndex        =   34
            Top             =   555
            Width           =   1635
         End
         Begin VB.ComboBox cboDb 
            Height          =   300
            Left            =   5145
            TabIndex        =   33
            Top             =   555
            Width           =   1755
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "搜索"
            Height          =   380
            Left            =   1800
            TabIndex        =   32
            ToolTipText     =   "搜索网络上可用的SQL服务器"
            Top             =   510
            Width           =   615
         End
         Begin VB.CommandButton cmdRefreshDB 
            Caption         =   "刷新"
            Height          =   380
            Left            =   6960
            TabIndex        =   31
            ToolTipText     =   "刷新当前数据库的表"
            Top             =   510
            Width           =   615
         End
         Begin VB.TextBox txtPw 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3705
            MultiLine       =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   30
            Top             =   555
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "密码:"
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   40
            Top             =   135
            Width           =   1005
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "用户名:"
            Height          =   255
            Index           =   1
            Left            =   2475
            TabIndex        =   39
            Top             =   135
            Width           =   1005
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "数据库:"
            Height          =   255
            Index           =   0
            Left            =   5160
            TabIndex        =   38
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "可用服务器:"
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   37
            Top             =   135
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   8355
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   8415
         Begin VB.TextBox txtAccessPW 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3360
            PasswordChar    =   "*"
            TabIndex        =   25
            Top             =   570
            Width           =   1455
         End
         Begin VB.TextBox txtAccessID 
            Height          =   300
            Left            =   1320
            TabIndex        =   24
            Top             =   570
            Width           =   1335
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "选择"
            Height          =   380
            Left            =   60
            TabIndex        =   23
            ToolTipText     =   "选择要打开的ACCESS数据库"
            Top             =   510
            Width           =   615
         End
         Begin VB.CommandButton cmdOpenMdb 
            Caption         =   "连接"
            Height          =   380
            Left            =   7680
            TabIndex        =   22
            ToolTipText     =   "连接当前ACCESS数据库"
            Top             =   510
            Width           =   615
         End
         Begin VB.TextBox txtFile 
            Height          =   300
            Left            =   1320
            TabIndex        =   21
            Top             =   120
            Width           =   6960
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "用户名:"
            Height          =   255
            Left            =   480
            TabIndex        =   28
            Top             =   615
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "密码:"
            Height          =   255
            Left            =   2640
            TabIndex        =   27
            Top             =   615
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "数据库文件名:"
            Height          =   225
            Left            =   135
            TabIndex        =   26
            Top             =   165
            Width           =   1500
         End
      End
      Begin ComctlLib.ListView lv 
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.ListBox lstField 
      Height          =   1950
      Left            =   3000
      Style           =   1  'Checkbox
      TabIndex        =   17
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmd 
      Caption         =   ".."
      Height          =   300
      Index           =   0
      Left            =   2595
      TabIndex        =   16
      ToolTipText     =   "新建一个数据访问器实例"
      Top             =   120
      Width           =   300
   End
   Begin VB.CheckBox chk 
      Caption         =   "显示内容"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtSpContent 
      Height          =   2295
      Left            =   8880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Top             =   4320
      Width           =   6375
   End
   Begin VB.PictureBox pic 
      Height          =   300
      Index           =   1
      Left            =   480
      Picture         =   "frmDebugDB.frx":1042
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox pic 
      Height          =   300
      Index           =   0
      Left            =   120
      Picture         =   "frmDebugDB.frx":4481
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   120
      Width           =   300
   End
   Begin VB.OptionButton opt 
      Caption         =   "存储过程"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Top             =   3540
      Width           =   1095
   End
   Begin VB.OptionButton opt 
      Caption         =   "视图"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   3540
      Width           =   855
   End
   Begin VB.OptionButton opt 
      Caption         =   "表"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3540
      Width           =   855
   End
   Begin VB.TextBox txtSql 
      Height          =   1665
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   495
      Width           =   5775
   End
   Begin VB.TextBox txtSqlLog 
      Height          =   3720
      Left            =   8865
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      ToolTipText     =   "在此输入SQL语句"
      Top             =   480
      Width           =   6390
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "(F5)运行"
      Height          =   375
      Left            =   1980
      TabIndex        =   4
      ToolTipText     =   "直接运行SQL语句"
      Top             =   3840
      Width           =   915
   End
   Begin ComctlLib.ListView lv 
      Height          =   3015
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin ComctlLib.ListView lv 
      Height          =   1935
      Index           =   2
      Left            =   4800
      TabIndex        =   8
      Top             =   2280
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin ComctlLib.ListView lv 
      Height          =   2535
      Index           =   3
      Left            =   8880
      TabIndex        =   18
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL历史记录:"
      Height          =   240
      Index           =   3
      Left            =   8850
      TabIndex        =   14
      Top             =   165
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL语句:"
      Height          =   240
      Index           =   2
      Left            =   3000
      TabIndex        =   10
      Top             =   165
      Width           =   855
   End
   Begin VB.Label lblRsCount 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Menu mm1 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu m1 
         Caption         =   "复制SQL语句到变量"
         Index           =   0
      End
   End
   Begin VB.Menu mm2 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu m2 
         Caption         =   "显示所有记录"
         Index           =   0
      End
      Begin VB.Menu m2 
         Caption         =   "显示全部字段"
         Index           =   1
      End
      Begin VB.Menu m2 
         Caption         =   "显示相同数据"
         Index           =   2
      End
   End
   Begin VB.Menu mm3 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu m3 
         Caption         =   "导出到Excel"
         Index           =   0
      End
      Begin VB.Menu m3 
         Caption         =   "导出到Xml"
         Index           =   1
      End
      Begin VB.Menu m3 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu m3 
         Caption         =   "删除当前数据"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmDebugDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W As New cyWndEx
Dim F As New cyFileEx
Dim S As New cySystemEx
Dim D As New cyDataBaseEx
Dim sS As New cyStringEx

Dim Rs As Recordset                 '公用暂存数据集
Dim Rslog As New Recordset          '保存连网记录及SQL语句记录
Dim RsSqlLog As New Recordset       '保存SQL语句

Dim iOpenID As Long                 '记录当前打开数据库的ID
Dim sLogFileName As String          '记录公共文件名，避免反复读取

Dim m_sTempSql As String            '临时记录SQL语句串,用于选择字段时显示字段名称
Dim m_RsDuplate As Recordset        '记录重复的项

Dim iWndCounter As Long             '记录实例的数量

'数据库列表得到焦点
Private Sub cboDb_GotFocus()
   On Error GoTo cboDb_GotFocus_Error

'================================================================错误处理过程
    If cboSQL.Text = "" Or txtID = "" Then Exit Sub
    cboDb.Clear
    
    '读出所有数据库到下拉框中
    D.cyRsToCtl D.cyGetSQLDataBaseNameToRs(cboSQL.Text, txtID, txtPw), cboDb
    
    '弹出数据库下拉框
    W.cyWndAction cboDb.hWnd, Cbo_PopupList
    
    '刷新页面
    Me.Refresh
'================================================================错误处理过程
ExitPrc:

    '统一退出点
    Exit Sub
cboDb_GotFocus_Error:

    '调用全局的错误处理过程
    Call ErrHandler(Err.Number, Err.Description, " cboDb_GotFocus")
    
End Sub

Private Sub chk_Click(Index As Integer)
    If Index = 0 Then
        If chk(Index).Value = vbChecked And chk(Index).Visible = True Then
        '当前选择了显示存储过程的内容,则数据显示框变窄
        
            txtSpContent.Height = dgTable.Height
            dgTable.Width = 577
            txtSpContent.Visible = True
    
        Else
        '其他情况下按最大宽度
        
            dgTable.Width = Me.Width / Screen.TwipsPerPixelX - 25
            dgTable.Height = Me.Height / Screen.TwipsPerPixelY - 345
            txtSpContent.Visible = False
    
        End If
        
    End If
        
End Sub

Private Sub cmd_Click(Index As Integer)
    
   On Error GoTo cmd_Click_Error
    Screen.MousePointer = vbHourglass

'================================================================错误处理过程

    If Index = 0 Then
        '新建一个数据库访问器实例
        Dim a As New frmDebugDB
        
        If Left(Me.Caption, 1) = "[" Then
        '已经是一个实例
        
        Else
        '不是实例
            Me.Caption = "[1] - " & Me.Caption
        End If
        
        a.Caption = Replace(Me.Caption, "[1]", "[2]")
        
        a.Show
    
    ElseIf Index = 2 Then
        '调用数据库列表的双击事件打开数据库
        lv_DblClick 0
        
        '清空当前显示的数据
        Set dgTable.DataSource = Nothing

    ElseIf Index = 3 Then
        If iOpenID = 0 Then
            Call MsgBox("请先选择待删除的数据库连接!", vbInformation Or vbSystemModal, "")
            Exit Sub
        
        End If
        
        Select Case MsgBox("是否确定要删除此数据库连接？", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton2, "警告")
        
            Case vbYes
                Set dgTable.DataSource = Rslog
                
                Rslog.Delete
                
                '保存数据库记录到文件
                D.cyRsStoreToFile Rslog, sLogFileName & ".Bin", BinaryFile
                
                '清空当前显示的数据
                Set dgTable.DataSource = Nothing
                
                '重新打开数据库链接保存库
                Set Rslog = D.cyRsGetFromFile(sLogFileName & ".Bin")
                
                '读出数据库连接集合
                Rslog.Filter = "ID>0"
                
                '显示列表排序
                Rslog.Sort = "服务器,数据库"
                
                If Me.Picture2.Visible Then
                    Rslog.Filter = "方法=2"
                
                Else
                    Rslog.Filter = "方法=1"
                
                End If
                
                '更新表/视图/存储过程列表
                lv(0).ListItems.Clear
                D.cyRsToCtl Rslog, lv(0)
                
            Case vbNo
        
        End Select
    
    ElseIf Index = 4 Then
        '取消显示
        Set dgTable.DataSource = Nothing
        Picture3.Visible = False

    End If
'================================================================错误处理过程
ExitPrc:

    '统一退出点
    Screen.MousePointer = 0
    Exit Sub
cmd_Click_Error:

    '调用全局的错误处理过程
    Screen.MousePointer = 0
    Call ErrHandler(Err.Number, Err.Description, " cmd_Click")

End Sub

Private Sub cmdSearch_Click()
   On Error GoTo cmdSearch_Click_Error
'================================================================错误过程处理
    Dim Rs  As Recordset
    
    '查找当前可用的SQL服务器
    Set Rs = D.cyGetSQLServerListToRs
    
    If Rs.RecordCount = 0 Then
        Call MsgBox("未发现可用的SQL服务器!", vbInformation Or vbSystemModal, "")
        
    Else
        D.cyRsToCtl Rs, cboSQL
        
    End If

'================================================================错误过程处理
   On Error GoTo 0
   Exit Sub
cmdSearch_Click_Error:
    If Err.Number = 91 Then
        Call MsgBox("未发现可用的SQL服务器!", vbInformation Or vbSystemModal, "")
    
    Else
            Call ErrHandler(Err.Number, Err.Description, "‘Form frmDB’中的函数‘cmdSearch_Click’")
        
    End If
    
End Sub

'执行Sql语句
Private Sub ExecSql()

   On Error GoTo ExecSql_Error
'================================================================错误过程处理
    Dim sStr As String
    Dim iMaxID As Long


    If txtSql = "" Then GoTo ExitPrc
    
    '有选择文本，则将其复制到剪贴板
    If txtSql.SelLength > 0 Then W.cyWndAction txtSql.hWnd, Txt_Copy

    '将SQL语句格式化为分行
    sStr = txtSql

    '格式化Sql语句
    sStr = Replace(sStr, vbCrLf, " ")
    sStr = Replace(sStr, "  ", " ")
    sStr = Replace(sStr, "select", "SELECT")
    sStr = Replace(sStr, "insert into", "INSERT INTO")
    sStr = Replace(sStr, "from", vbCrLf & "FROM")
    sStr = Replace(sStr, "where", vbCrLf & "WHERE")
    sStr = Replace(sStr, "order by", vbCrLf & "ORDER BY")
    sStr = Replace(sStr, "group by", vbCrLf & "GROUP BY")
    
    sStr = Replace(sStr, "FROM", vbCrLf & "FROM")
    sStr = Replace(sStr, "WHERE", vbCrLf & "WHERE")
    sStr = Replace(sStr, "ORDER BY", vbCrLf & "ORDER BY")
    sStr = Replace(sStr, "GROUP BY", vbCrLf & "GROUP BY")
    sStr = Replace(sStr, vbCrLf & vbCrLf, vbCrLf)
    
    sStr = Replace(sStr, "＝", "=")
    sStr = Replace(sStr, "‘", "'")
    sStr = Replace(sStr, "’", "'")
    sStr = Replace(sStr, "　", " ")
    
    txtSql = sStr
    
    '执行SQL语句
    Set Rs = D.cyGetRs(sStr)
    D.cyRsToCtl Rs, dgTable

    lblRsCount = Rs.RecordCount & "条记录"
    
'================================================================错误过程处理
   
ExitPrc:
   Exit Sub
   
ExecSql_Error:
    Screen.MousePointer = 0
    
    '避免显示存储过程内部抛出的错误提示
        Call ErrHandler(Err.Number, Err.Description, "‘Form frmDebugDB’中的函数‘ExecSql’")
    
End Sub

Private Sub cmdExec_Click()
    
   On Error GoTo cmdExec_Click_Error
    Screen.MousePointer = vbHourglass

'================================================================错误处理过程
    '去除行首的回车
    Do While Left(txtSql, 1) = Chr(13)
        txtSql = Replace(txtSql, vbCrLf, "")
        
    Loop
    
    '检查是否有选中运行的语句
    If Len(txtSql) = 0 Then
        Call MsgBox("请先选中要执行的SQL语句.", vbInformation Or vbSystemModal, "")
        W.cyWndAction txtSqlLog.hWnd, Wnd_Flash
        txtSqlLog.SetFocus
        Exit Sub
    
    End If
    
    '执行Sql语句
    Call ExecSql
    
    '记录已执行的Sql
    
    If InStr(1, txtSqlLog, txtSql) = 0 Then
    '该语句没保存过，则添加
        txtSqlLog = "----------------------------------" & vbCrLf & vbCrLf & txtSql & vbCrLf & vbCrLf & "----------------------------------" & vbCrLf & txtSqlLog
    
    End If
    
    txtSql.SetFocus
    
    '如果该数据库之前没有SQL记录，则添加一条
    If RsSqlLog.RecordCount = 0 Then
        RsSqlLog.AddNew
        RsSqlLog(0) = iOpenID
        RsSqlLog(1) = txtSqlLog
        RsSqlLog.Update
        
    Else
    '有则更新
    
        '保存SQL记录
        RsSqlLog(1) = txtSqlLog
        
        '读出SQL
        If RsSqlLog.RecordCount > 0 Then txtSqlLog = RsSqlLog(1)
    
    End If
    
    RsSqlLog.Filter = "ID>0"
    D.cyRsStoreToFile RsSqlLog, sLogFileName & cboDb & ".BinLog", BinaryFile
'================================================================错误处理过程
ExitPrc:

    '统一退出点
    Screen.MousePointer = 0
    Exit Sub
cmdExec_Click_Error:

    '调用全局的错误处理过程
    Screen.MousePointer = 0
        Call ErrHandler(Err.Number, Err.Description, " cmdExec_Click")

End Sub

Private Sub cmdOpenMdb_Click()

On Error GoTo cmdOpenMdb_Click_Error
'================================================================错误过程处理

    Dim iMaxID As Long
    
    If Not F.cyFileExist(txtFile) Then
        Call MsgBox("数据库文件不存在!", vbCritical Or vbSystemModal, "")
        GoTo ExitPrc
        
    End If
    
    If txtAccessPW <> "" Then '有输入密码
        If txtAccessID <> "" Then '使用用户组帐号
            D.cyConnectAccess txtFile, , txtAccessID, txtAccessPW
        Else
            D.cyConnectAccess txtFile, txtAccessPW
        End If
        
    Else    '没有密码
        D.cyConnectAccess txtFile
        
    End If
    
    '读出数据库连接集合
    Rslog.Filter = "ID>0"
    
    '显示列表排序
    Rslog.Sort = "服务器,数据库"
    
    Rslog.MoveLast
    
    iMaxID = Rslog("ID")

    '记录数据库连接
    Rslog.Filter = "数据库='" & txtFile & "' and 帐号='" & txtAccessID & "' and 密码='" & sS.cyStrEncrypt("cyDLL", txtAccessPW) & "'"
    
    '如果不存在则添加
    If Rslog.RecordCount = 0 Then
        Rslog.Filter = "ID>0 and 方法=1"
        
        '避免当新建数据集时记录则IMAXID为空
    On Error Resume Next

        Rslog.AddNew
        Rslog("ID") = iMaxID + 1
        Rslog("方法") = 1
        Rslog("数据库") = txtFile
        Rslog("帐号") = txtAccessID
        Rslog("密码") = sS.cyStrEncrypt("cyDLL", txtAccessPW)
        iOpenID = Rslog("ID")
        Rslog.Update
    End If
    
    '清空密码
    txtAccessPW = ""
    lv(1).ListItems.Clear
    
    '读出该数据库的列表
    D.cyRsToCtl D.cyGetTableNameToRs, lstTable

    Rslog.Filter = "方法 = 3 and 关联ID = " & iOpenID

    '保存数据库列表
    Rslog.Filter = "ID<100000"
    D.cyRsStoreToFile Rslog, sLogFileName & ".Bin", BinaryFile
    Call cmd_Click(4)
    
    '读出SQL语句记录
    
    '打开数据库链接保存库
    Set RsSqlLog = D.cyRsGetFromFile(sLogFileName & ".BinLog")

    '如果没有SQL记录则添加一条
    If RsSqlLog.RecordCount = 0 Then
        RsSqlLog.AddNew
        RsSqlLog(0) = iOpenID
        RsSqlLog.Update
    Else
        '读出SQL记录
        RsSqlLog.Filter = "ID=" & iOpenID
        
        '读出SQL
        If RsSqlLog.RecordCount > 0 Then txtSqlLog = RsSqlLog(1)
    
    End If

    Me.Caption = txtFile & " - 数据访问器"

'================================================================错误过程处理
   On Error GoTo 0
ExitPrc:
   Exit Sub
   
cmdOpenMdb_Click_Error:
        Call ErrHandler(Err.Number, Err.Description, "‘Form frmDebugDB’中的函数‘cmdOpenMdb_Click’")
        
End Sub

Private Sub cmdConnectDB_Click()
    Dim Rs As New Recordset
    Dim iMaxID As Long
   
   On Error GoTo cmdConnectDB_Click_Error
'================================================================错误过程处理

    opt(0).Value = True
    lv(1).ListItems.Clear
    D.cyConnectSqlServer cboSQL, cboDb, txtID, txtPw
    
    '显示表
    D.cyRsToCtl D.cyGetTableNameToRs, lv(1)
    
    '用户选择了不记录数据库连接
    If Not F.cyFileExist(sLogFileName & ".Bin") Then
        Picture3.Visible = False
    
    Else
    
        Rslog.Filter = "ID>0"
        Rslog.Sort = "ID"
        Rslog.MoveLast
    
        iMaxID = Rslog("ID")
        
        Rslog.Filter = "服务器='" & cboSQL & "' and 数据库='" & cboDb & "' and 帐号='" & txtID & "' and 密码='" & sS.cyStrEncrypt("cyDLL", txtPw) & "'"
        
        '如果不存在则添加
        If Rslog.RecordCount = 0 Then
            Rslog.Filter = "ID>0 and 方法=2"
            
            '避免当新建数据集时记录则IMAXID为空
        On Error Resume Next
            
            Rslog.AddNew
            Rslog("ID") = iMaxID + 1
            Rslog("方法") = 2
            Rslog("服务器") = cboSQL
            Rslog("数据库") = cboDb
            Rslog("帐号") = txtID
            Rslog("密码") = sS.cyStrEncrypt("cyDLL", txtPw)
            Rslog.Update
        End If
        txtPw = ""
        
        Rslog.Filter = "方法 = 3 and 关联ID = " & iOpenID
        
        '保存数据库列表
        Rslog.Filter = "ID<100000"
        D.cyRsStoreToFile Rslog, sLogFileName & ".Bin", BinaryFile
        
        Call cmd_Click(4)
        
        '读出SQL语句记录
        '打开数据库链接保存库
        Set RsSqlLog = D.cyRsGetFromFile(sLogFileName & ".BinLog")
    
        If F.cyFileExist(sLogFileName & cboDb & ".BinLog") Then
        '该数据库记录已存在
            Set Rs = D.cyRsGetFromFile(sLogFileName & cboDb & ".BinLog")
            If Rs.RecordCount > 0 Then txtSqlLog = Rs(1)
        
        End If
    
    End If
    
    Me.Caption = cboDb & "/" & cboSQL & " - 数据访问器"
    
'================================================================错误过程处理
   On Error GoTo 0
   Exit Sub
   
cmdConnectDB_Click_Error:
    Call ErrHandler(Err.Number, Err.Description, "‘Form frmDebugDB’中的函数‘cmdConnectDB_Click’")
    
End Sub

Private Sub cmdSel_Click()
        Dim sStr As String
   On Error GoTo cmdSel_Click_Error
'================================================================错误过程处理

    sStr = F.cyDialogOpen(Me.hWnd, "请选择要打开的数据库", "*.mdb|*.mdb")
    txtFile = sStr
    txtAccessID = ""
    txtAccessPW = ""
    cmdOpenMdb.SetFocus
    
'================================================================错误过程处理
   Exit Sub
   
cmdSel_Click_Error:
    Call ErrHandler(Err.Number, Err.Description, "‘Form frmDebugDB’中的函数‘cmdSel_Click’")
        
End Sub

Private Sub Command1_Click()
End Sub

Private Sub dgDuplate_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

End Sub

Private Sub dgTable_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mm3
        
    End If

End Sub

Private Sub Form_Load()

    '初始化显示列表
    W.cySetListviewWidths lv(1), "数据表/视图", "181"
    W.cySetListviewWidths lv(2), "参数名;参数类型", "150;140"
    
    '则生成一个
    sLogFileName = F.cyGetSpecialFolder(Personal) & sS.cyMD5(S.cyGetComputerName & S.cyGetUserName)
    
    If Not F.cyFileExist(sLogFileName & ".Bin") Then
    '数据库记录文件不存在
        
        Select Case MsgBox("是否在本机保存数据库连接信息?", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton2, "")
        
            Case vbYes
            
            
                '新建一条记录
                Set Rslog = Nothing
                Rslog.Fields.Append "ID", adInteger, 2
                Rslog.Fields.Append "方法", adInteger, 2
                Rslog.Fields.Append "关联ID", adInteger, 2
                Rslog.Fields.Append "服务器", adVarChar, 20
                Rslog.Fields.Append "数据库", adVarChar, 255
                Rslog.Fields.Append "帐号", adVarChar, 50
                Rslog.Fields.Append "密码", adVarChar, 50
                Rslog.Fields.Append "其它", adLongVarChar, 2096
                Rslog.CursorLocation = adUseClient
                Rslog.Open
                Rslog.AddNew
                Rslog("ID") = 1
                Rslog.Update
                D.cyRsStoreToFile Rslog, sLogFileName & ".Bin", BinaryFile
        
            Case vbNo
        
        End Select
    
    End If
        
    '缺省显示SQL连接
    Call pic_Click(0)
    
    '计数器+1
    iWndCounter = iWndCounter + 1
    
    Debug.Print iWndCounter
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
    If chk(0).Value = vbChecked Then
    '当前选择了显示存储过程的内容
        txtSpContent.Height = dgTable.Height
        dgTable.Width = 577
        txtSpContent.Visible = True
    
    Else
        dgTable.Width = Me.Width / Screen.TwipsPerPixelX - 25
        dgTable.Height = Me.Height / Screen.TwipsPerPixelY - 345
        txtSpContent.Visible = False
    
    End If
        
    lblRsCount.Top = dgTable.Top + dgTable.Height + 5
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err
    
    '保存Sql记录
    RsSqlLog(1) = txtSqlLog
    D.cyRsStoreToFile RsSqlLog, sLogFileName & ".BinLog", BinaryFile

    Set Rs = Nothing
    D.cyCnClose
    
Err:

End Sub

Private Sub lstField_Click()
    Dim i As Long
    Dim sStr As String
    
'    If lstField.Selected(0) = True Then
'
'        '选中全部
'
'        lstField.Selected(0) = False
'
'        For i = 1 To lstField.ListCount - 1
'
'            '如果该项被选中
'            lstField.Selected(i) = True
'
'        Next
'
'    End If
    
    For i = 0 To lstField.ListCount - 1
    
        '如果该项被选中
        If lstField.Selected(i) Then sStr = lstField.List(i) & " , " & sStr
    
    Next
    
    If sStr = "" Then
        txtSql = m_sTempSql
        
    Else
        sStr = " " & Left(sStr, Len(sStr) - 2)
        txtSql = Replace(m_sTempSql, "10 * ", "10 " & sStr)
    
    End If
    
End Sub

Private Sub lv_Click(Index As Integer)

   On Error GoTo lv_Click_Error

'================================================================错误处理过程
If Index = 0 Then

    If lv(0).ListItems.Count = 0 Then GoTo ExitPrc
    
    '打开当前选择的数据库
    Rslog.Find "ID=" & lv(0).SelectedItem.Text
    
    If Picture1.Visible = True Then
    '当前的是ACCESS
        
        '读出当前选择的数据库信息
        txtFile = lv(0).SelectedItem.SubItems(4)
        txtAccessID = lv(0).SelectedItem.SubItems(5)
        txtAccessPW = sS.cyStrDecrypt("cyDLL", lv(0).SelectedItem.SubItems(5))
    
    ElseIf Picture2.Visible = True Then
    '当前的是SQL
    
        '读出当前选择的数据库信息
        cboSQL = lv(0).SelectedItem.SubItems(3)
        cboDb.Text = lv(0).SelectedItem.SubItems(4)
        txtID = lv(0).SelectedItem.SubItems(5)
        txtPw = sS.cyStrDecrypt("cyDLL", lv(0).SelectedItem.SubItems(6))
    
    End If
    
    '记录当前打开的记录ID
    iOpenID = lv(0).SelectedItem.Text

ElseIf Index = 3 Then
    
    Dim i As Long
    Dim sStr As String
    
    For i = 0 To lv(Index).ListItems.Count - 2
    
        '如果该项被选中
        sStr = sStr & " , " & lv(Index).ColumnHeaders(1) & " ='" & lv(Index).SelectedItem.Text & "'"
    
    Next
    
    If sStr = "" Then
        txtSql = m_sTempSql
        
    Else
        sStr = " " & Right(sStr, Len(sStr) - 2)

    End If

    Debug.Print sS.cyMidEx(m_sTempSql, vbCrLf, vbCrLf)
    txtSql = "SELECT TOP 10 * " & vbCrLf & sS.cyMidEx(m_sTempSql, vbCrLf, vbCrLf) & vbCrLf & _
            " WHERE " & Replace(sStr, " , ", " AND ")

    Call cmdExec_Click

        
End If
'================================================================错误处理过程
ExitPrc:

    '统一退出点
    Screen.MousePointer = 0
    Exit Sub
lv_Click_Error:

    '调用全局的错误处理过程
    Screen.MousePointer = 0
    Call ErrHandler(Err.Number, Err.Description, " lv_Click")

End Sub

Private Sub lv_DblClick(Index As Integer)

'如果没有保存数据库记录，则双击时忽略该动作
   On Error GoTo lv_DblClick_Error
    Screen.MousePointer = vbHourglass

'================================================================错误处理过程
    If lv(0).ListItems.Count = 0 Then GoTo ExitPrc
    
    If Picture1.Visible = True Then
    '当前的是ACCESS
    
        '先读出当前的数据库
        Call lv_Click(0)
        '再打开
        Call cmdOpenMdb_Click
        
    ElseIf Picture2.Visible = True Then
    '当前的是SQL
        
        '先读出当前的数据库
        Call lv_Click(0)
        '再打开
        Call cmdConnectDB_Click
        
    End If
    
'================================================================错误处理过程
ExitPrc:

    '统一退出点
    Screen.MousePointer = 0
    Exit Sub
lv_DblClick_Error:

    '调用全局的错误处理过程
    Screen.MousePointer = 0
    Call ErrHandler(Err.Number, Err.Description, " lv_DblClick")

End Sub

Private Sub lv_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo lv_MouseUp_Error

'================================================================错误处理过程
If Index = 1 Then

    If lv(1).ListItems.Count = 0 Then GoTo ExitPrc
    If lv(1).SelectedItem = "" Then GoTo ExitPrc
    
       
    If Button = 1 And Shift = 2 Then
    'Ctl+左键侧显示该表所有记录
        txtSql = "SELECT  * FROM [" & lv(1).SelectedItem & "]"
            
    ElseIf Button = 2 And Shift = 0 Then
        PopupMenu mm2

    Else
        If opt(2).Value = True Then
        '存储过程，则读出参数列表
            
            Dim i As Long
            Dim sSQL As String
            Dim sStr As String
            sSQL = "select  参数名称=b.name " + _
                   "  ,参数类型=c.name " + _
                   "  +case   when   c.name   in   ('binary','char','nchar','nvarchar','varbinary','varchar','float','real') " + _
                   "    then   '('+cast(b.prec   as   varchar)+')' " + _
                   "    when   c.name   in   ('decimal','numeric') " + _
                   "    then   '('+cast(b.prec   as   varchar)+','+cast(b.scale   as   varchar)+')' " + _
                   "  Else   ''   end " + _
                   "  from   sysobjects   a " + _
                   "  join   syscolumns   b   on   a.id=b.id " + _
                   "  join   systypes   c   on   b.xusertype=c.xusertype " + _
                   "  where   a.xtype='P'   and   a.status>0    and a.name='" & lv(1).SelectedItem & "' " + _
                   "  order   by   a.name,b.colid"
            lv(2).ListItems.Clear
            D.cyRsToCtl D.cyGetRs(sSQL), lv(2)
            
            '读出存储过程内容
            sSQL = "select c.text, c.encrypted, c.number,  " + _
                   "xtype=convert(nchar(2), o.xtype),  " + _
                   "datalength(c.text), convert(varbinary(8000),  " + _
                   "c.text), 0 from dbo.syscomments c, dbo.sysobjects o  " + _
                   "where o.id = c.id and c.id = object_id('" & lv(1).SelectedItem & "')  " + _
                   "order by c.number, c.colid option(robust plan) "
                   
            txtSpContent = D.cyGetRsOneField(sSQL, "")
            
            '自动完成参数到文本
            For i = 1 To lv(2).ListItems.Count
                sStr = "''" & "," & sStr
            Next
            
            If sStr <> "" Then
            '有参数,则去掉最后的,
                sStr = Left(sStr, Len(sStr) - 1)
                
            End If
            
            '如果存储过程名字带空格则额头[]括起来
            txtSql = IIf(InStr(1, lv(1).SelectedItem, " ") > 0, "[" & lv(1).SelectedItem & "]", lv(1).SelectedItem) & " " & sStr
        
        Else
        '表或视图
        
            '先拿到表的第一个字段名
            Set Rs = D.cyGetRs(Replace(Replace("SELECT TOP 1 * FROM [" & lv(1).SelectedItem & "]", "[[", "["), "]]", "]"))
            
            '显示该表的所有字段
            lstField.Clear
            For i = 1 To Rs.Fields.Count
                lstField.AddItem Rs.Fields(i - 1).name
            Next
            lstField.AddItem "*", 0
            
            '按倒序找到最后十条记录
            txtSql = "SELECT TOP 10 * FROM [" & lv(1).SelectedItem & "]" & vbCrLf & vbCrLf & "ORDER BY  " & Rs(0).name & " DESC"
            
        End If
        
    End If
    
    txtSql = Replace(txtSql, "[[", "[")
    txtSql = Replace(txtSql, "]]", "]")
        
    '临时存放Sql语句,选择字段时进行替换
    m_sTempSql = txtSql
        
    If Left(UCase(txtSql), 6) = "SELECT" Then
    '返回RS
        Set Rs = D.cyGetRs(txtSql)
        D.cyRsToCtl Rs, dgTable
    
        If Button = 1 And Shift = 2 Then
            lblRsCount = Rs.RecordCount & "条记录"
         
        Else
            lblRsCount = D.cyGetRsOneField(Replace(Split(txtSql, "ORDER BY ")(0), "*", "Count(*)")) & " 条记录"
            
        End If
    
    End If
    
End If
'================================================================错误处理过程
ExitPrc:
    '统一退出点
    Exit Sub
    
lv_MouseUp_Error:

'    '调用全局的错误处理过程
'    Call ErrHandler(Err.Number, Err.Description, " lv_MouseUp")
    
End Sub

Private Sub Menu_cyMenuClick(ByVal iMenuKeyID As Long, ByVal sMenuCaption As String)
    
   On Error GoTo Menu_cyMenuClick_Error
    
    Select Case iMenuKeyID

        Case 21
            txtSql = "SELECT  * FROM [" & lv(1).SelectedItem & "]"
            txtSql = Replace(txtSql, "[[", "[")
            txtSql = Replace(txtSql, "]]", "]")
            
            If Left(UCase(txtSql), 6) = "SELECT" Then
            '返回RS
                Set Rs = D.cyGetRs(txtSql)
                D.cyRsToCtl Rs, dgTable
            
                If Button = 1 And Shift = 2 Then
                    lblRsCount = Rs.RecordCount & "条记录"
                 
                Else
                    lblRsCount = D.cyGetRsOneField(Replace(txtSql, "*", "Count(*)")) & "条记录"
                    
                End If
            
            Else
            '执行SQL语句
                D.cyExeCute txtSql
                
            End If
        
        Case 22
            If Left(UCase(txtSql), 6) = "SELECT" Then
            '返回RS
                
                For i = 0 To Rs.Fields.Count - 1
                    sTemp = sTemp & " ," & IIf(InStr(1, Rs(i).name, " ") > 0, "[" & Rs(i).name & "]", Rs(i).name)
                Next
                sTemp = Replace(sTemp, " ,", "", , 1)
                
                txtSql = Replace(txtSql, "*", sTemp & " ", , 1)
                
            End If
    
            
    End Select
'================================================================错误处理过程
Exit Sub:

    '统一退出点
    Screen.MousePointer = 0
    Exit Sub
    
Menu_cyMenuClick_Error:

    '调用全局的错误处理过程
    Screen.MousePointer = 0
    Call ErrHandler(Err.Number, Err.Description, " Menu_cyMenuClick")

End Sub

Private Sub m1_Click(Index As Integer)
    If Index = 0 Then
        sStr = txtSql
        '变换为数组并添加头尾字符
        sA = Split(sStr, vbCrLf)
        For i = 0 To UBound(sA) - 1
            If i = 0 Then
                sA(i) = "sSQL=" & Chr(34) & sA(i) & " " & Chr(34) & " + _"
            Else
                sA(i) = "           " & Chr(34) & sA(i) & " " & Chr(34) & " + _"
            End If
        Next
        sA(UBound(sA)) = "           " & Chr(34) & sA(UBound(sA)) & " " & Chr(34)
        sStr = Join(sA, vbCrLf)
        
        '添加变量名
        sStr = vbTab & "dim sSQL as string " & vbCrLf & vbTab & sStr
        
        '存放到剪贴板中
        Clipboard.Clear
        Clipboard.SetText sStr, 1
    
    End If
    
End Sub

Private Sub m2_Click(Index As Integer)
    Dim sTemp As String

    If Index = 0 Then
        txtSql = "SELECT  * FROM [" & lv(1).SelectedItem & "]"
        txtSql = Replace(txtSql, "[[", "[")
        txtSql = Replace(txtSql, "]]", "]")
        
        If Left(UCase(txtSql), 6) = "SELECT" Then
        '返回RS
            Set Rs = D.cyGetRs(txtSql)
            D.cyRsToCtl Rs, dgTable
        
            If Button = 1 And Shift = 2 Then
                lblRsCount = Rs.RecordCount & "条记录"
             
            Else
                lblRsCount = D.cyGetRsOneField(Replace(txtSql, "*", "Count(*)")) & "条记录"
                
            End If
        
        Else
        '执行SQL语句
            D.cyExeCute txtSql
            
        End If
        
    ElseIf Index = 1 Then
        If Left(UCase(txtSql), 6) = "SELECT" Then
        '返回RS
            
            For i = 0 To Rs.Fields.Count - 1
                sTemp = sTemp & " ," & IIf(InStr(1, Rs(i).name, " ") > 0, "[" & Rs(i).name & "]", Rs(i).name)
            Next
            sTemp = Replace(sTemp, " ,", "", , 1)
            
            txtSql = Replace(txtSql, "*", sTemp & " ", , 1)
            
        End If

    ElseIf Index = 2 Then
        lv(3).Width = txtSpContent.Width
        lv(3).Height = txtSpContent.Height
        lv(3).Visible = True
        lv(3).ZOrder 0
        
    Dim sStr As String
        
        For i = 0 To lstField.ListCount - 1
        
            '如果该项被选中
            If lstField.Selected(i) Then sStr = sStr & " , " & lstField.List(i)
        
        Next
        
        If sStr = "" Then
            txtSql = m_sTempSql
            
        Else
            sStr = " " & Right(sStr, Len(sStr) - 3)
            txtSql = "SELECT " & sStr & " , COUNT(" & Split(sStr, " ,")(0) & " ) AS 重复数量 " & vbCrLf & _
                    "FROM " & lv(1).SelectedItem & " " & vbCrLf & _
                    "GROUP BY" & sStr & " HAVING COUNT(" & Split(sStr, ",")(0) & " )>1"
        
            Set m_RsDuplate = D.cyGetRs(txtSql)
            lv(3).ListItems.Clear
            
            If m_RsDuplate.Fields.Count > lv(3).ColumnHeaders.Count Then
            '如果列数不相同
            
                For i = lv(3).ColumnHeaders.Count To m_RsDuplate.Fields.Count
                    lv(3).ColumnHeaders.Add , , " "
                    
                Next
            
            End If
            
            D.cyRsToCtl m_RsDuplate, lv(3)
            For i = 0 To m_RsDuplate.Fields.Count - 1
                lv(3).ColumnHeaders(i + 1).Text = m_RsDuplate(i).name
                
            Next
        
        End If
            
        
    End If
    
End Sub

Private Sub m3_Click(Index As Integer)

Dim i As Long
Dim sA() As String
Dim sStr As String
Dim sTemp As String
Dim RsTemp As Recordset

   On Error GoTo m3_Click_Error
    Screen.MousePointer = vbHourglass

'================================================================错误处理过程
If Index = 0 Then
    On Error Resume Next
    
        If Rs.RecordCount = 0 Then
            Call MsgBox("没有数据可供导出!", vbCritical Or vbSystemModal, "")
            Exit Sub
        
        End If
        
        Set RsTemp = Rs
        
        '文件名
        sStr = F.cyDialogSave(Me.hWnd, "保存XLS", "*.XLS|*.XLS")
        If sStr = "" Then Exit Sub
        
        Screen.MousePointer = vbHourglass

        '断开与数据库的连接
        Set RsTemp.ActiveConnection = Nothing
        RsTemp.MoveFirst
        D.cyRsToExcel RsTemp, sStr
        Screen.MousePointer = 0
        
        Call MsgBox("已成功导出当前数据到Excel!", vbInformation Or vbSystemModal, "")

ElseIf Index = 1 Then
    On Error Resume Next
        
        If Rs.RecordCount = 0 Then
            Call MsgBox("没有数据可供导出!", vbCritical Or vbSystemModal, "")
            Exit Sub
        
        End If
        
        Set RsTemp = Rs
        
        '文件名
        sStr = F.cyDialogSave(Me.hWnd, "保存Xml", "*.Xml|*.Xml")
        If sStr = "" Then Exit Sub
        
        
        '断开与数据库的连接
        Set RsTemp.ActiveConnection = Nothing
        RsTemp.MoveFirst
        D.cyRsStoreToFile RsTemp, sStr, XmlFile
        Screen.MousePointer = 0
        
        Call MsgBox("已成功导出当前数据到Xml!", vbInformation Or vbSystemModal, "")


ElseIf Index = 3 Then

        '逐条删除当前数据集中的内容
        If Rs.RecordCount = 0 Then Exit Sub
        Select Case MsgBox("此操作将删除当前显示的所有数据,是否确定删除?", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton2, "")
        
            Case vbYes
                
                Screen.MousePointer = vbHourglass
                Do While Rs.RecordCount <> 0
                    Rs.MoveFirst
                    Rs.Delete
                  
                Loop
                Screen.MousePointer = 0
                    
            Case vbNo
        
        End Select

End If
'================================================================错误处理过程
ExitPrc:

    '统一退出点
    Screen.MousePointer = 0
    Exit Sub
m3_Click_Error:

    '调用全局的错误处理过程
    Screen.MousePointer = 0
    Call ErrHandler(Err.Number, Err.Description, " m3_Click")

End Sub

Private Sub opt_Click(Index As Integer)
On Error GoTo Pass

    '清空表列表
    lv(1).ListItems.Clear
    
    '清空参数列表
    lv(2).ListItems.Clear
    
    '清空SQL语句
    txtSql = ""
    
    If Index = 0 Then
    '读出表列表
        D.cyRsToCtl D.cyGetTableNameToRs, lv(1)
        chk(0).Visible = False
        
    ElseIf Index = 1 Then
    '读出视图列表
        D.cyRsToCtl D.cyGetRs("Select name  From sysobjects WHERE XTYPE='V' AND category='0' order by name"), lv(1)
        chk(0).Visible = False
        
    ElseIf Index = 2 Then
    '读出存储过程列表
        D.cyRsToCtl D.cyGetRs("select name from sysobjects where xtype='p' and status>0 order by name"), lv(1)
        chk(0).Visible = True
    
    End If
    
    chk_Click 0
    
Pass:
'避免未连接数据库时出错
End Sub

Private Sub pic_Click(Index As Integer)

    '切换数据库前清空记录
    txtSqlLog = ""
    txtSql = ""

    '用户选择了不记录数据库连接
    If Not F.cyFileExist(sLogFileName & ".Bin") Then
    
        If Index = 1 Then
            '设置适当的列宽
            W.cySetListviewWidths lv(0), ";;;;数据库;账户;密码;;", "0;0;0;0;6500;2520;0;0"
            Picture1.Visible = True
            Picture2.Visible = False
            Picture3.Visible = True
            Picture3.ZOrder 0
            
        ElseIf Index = 0 Then
            '设置适当的列宽
            W.cySetListviewWidths lv(0), ";;;数据库;账户;密码;;", "0;0;0;3075.024;2564.788;2520;0;0"
            Picture1.Visible = False
            Picture2.Visible = True
            Picture3.Visible = True
            Picture3.ZOrder 0
        
        End If
    
        Exit Sub
    
    End If


    '打开数据库链接保存库
    Set Rslog = D.cyRsGetFromFile(sLogFileName & ".Bin")
    
    If Index = 1 Then
    '显示ACCESS的已保存连接字
    
        '清空SQL语句显示
        txtSqlLog = ""
        Picture1.Visible = True
        Picture2.Visible = False
        Picture3.Visible = True
        Picture3.ZOrder 0
       
        '如果没有连接记录则忽略
        If Rslog.RecordCount = 0 Then Exit Sub
        
        '打开数据库链接保存库
        Set Rslog = D.cyRsGetFromFile(sLogFileName & ".Bin")
        
        '读出ACCESS的数据库集合
        Rslog.Filter = "ID>0"
        Rslog.Filter = "方法=1"
        lv(0).ListItems.Clear
        D.cyRsToCtl Rslog, lv(0)
        
   
    ElseIf Index = 0 Then
    '显示SQL的已保存连接字
        
        Picture1.Visible = False
        Picture2.Visible = True
        Picture3.Visible = True
        Picture3.ZOrder 0
        
        '如果只有一条记录，则表示未有记录连接，则
        If Rslog.RecordCount = 0 Then
        
            '新建一条记录
            Set Rslog = Nothing
            Rslog.Fields.Append "ID", adInteger, 2
            Rslog.Fields.Append "方法", adInteger, 2
            Rslog.Fields.Append "关联ID", adInteger, 2
            Rslog.Fields.Append "服务器", adVarChar, 20
            Rslog.Fields.Append "数据库", adVarChar, 255
            Rslog.Fields.Append "帐号", adVarChar, 50
            Rslog.Fields.Append "密码", adVarChar, 50
            Rslog.Fields.Append "其它", adLongVarChar, 2096
            Rslog.CursorLocation = adUseClient
            Rslog.Open
            Rslog.AddNew
            Rslog("ID") = 1
            Rslog.Update
            D.cyRsStoreToFile Rslog, sLogFileName & ".Bin", BinaryFile

        End If
        
        '读出SQL的数据库集合
        Rslog.Filter = "ID>0"
        Rslog.Filter = "方法=2"
        
        '显示列表排序
        Rslog.Sort = "服务器,数据库"
        
        lv(0).ListItems.Clear
        D.cyRsToCtl Rslog, lv(0)
        
        '设置适当的列宽
        lv(0).ColumnHeaders(1).Width = 0
        lv(0).ColumnHeaders(2).Width = 0
        lv(0).ColumnHeaders(3).Width = 0
        lv(0).ColumnHeaders(4).Width = 3075.024
        lv(0).ColumnHeaders(5).Width = 2564.788
        lv(0).ColumnHeaders(6).Width = 2520
        lv(0).ColumnHeaders(7).Width = 0
        lv(0).ColumnHeaders(8).Width = 0
    
    End If
    
End Sub

Private Sub txtSql_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 2 Then
            
            txtSql.Enabled = False
            txtSql.Enabled = True
            
            PopupMenu mm1
            
        End If

End Sub

Private Sub txtSql_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtSql.SetFocus

End Sub

Private Sub txtSqlLog_Click()
On Error Resume Next
    Dim i As Long
    Dim j As Long
    i = InStrRev(txtSqlLog, "----------------------------------", txtSqlLog.SelStart) + 34
    j = InStr(i, txtSqlLog, "----------------------------------")
    txtSql = Mid(txtSqlLog, i, j - i)
    
End Sub

Private Sub txtSqlLog_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
    'F5执行
        Call cmdExec_Click
        
    ElseIf KeyCode = 65 And Shift = 2 Then
    'Ctl+A,全选文字
        S.cyKeyBoardAction , vbKeyControl, vbKeyHome
        S.cyKeyBoardAction , vbKeyShift, vbKeyControl, vbKeyEnd
    
    End If
    
End Sub

'全局错误处理过程
Public Sub ErrHandler(ByVal ErrorNumber As Long, ByVal ErrorMessage As String, Optional ByVal ErrorModule As String)
On Error Resume Next
    Screen.MousePointer = 0
    
    If ErrorNumber = -2147467259 Then
        Call MsgBox("数据库已断开连接!", vbCritical Or vbSystemModal, "错误")
       
    Else
        Call MsgBox("错误位置：" & ErrorModule & vbCrLf & "错误代码：" & ErrorNumber & vbCrLf & "错误描述：" & ErrorMessage, vbCritical Or vbSystemModal, "错误")
        
    End If
    
End Sub
'
'Private Sub txtSqlLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
'    If Button = 2 Then
'        txtSqlLog.Enabled = False
'
'        txtSqlLog.Enabled = True
'
'    End If
'
'End Sub

Private Sub txtSql_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
    'F5执行
        Call cmdExec_Click
    
    ElseIf KeyCode = 65 And Shift = 2 Then
    'Ctl+A
        S.cyKeyBoardAction , vbKeyControl, vbKeyHome
        S.cyKeyBoardAction , vbKeyShift, vbKeyControl, vbKeyEnd
    
    End If

End Sub

