VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Begin VB.Form frmDebugDB 
   Caption         =   "���ݲ�ѯ��"
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
         Caption         =   "����"
         Height          =   375
         Index           =   2
         Left            =   5760
         TabIndex        =   43
         ToolTipText     =   "���������ʵ����ݿ�"
         Top             =   3480
         Width           =   690
      End
      Begin VB.CommandButton cmd 
         Caption         =   "ɾ��"
         Height          =   375
         Index           =   3
         Left            =   6960
         TabIndex        =   42
         ToolTipText     =   "ɾ���˼�¼"
         Top             =   3480
         Width           =   690
      End
      Begin VB.CommandButton cmd 
         Caption         =   "ȡ��"
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   41
         ToolTipText     =   "ɾ���˼�¼"
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
            Caption         =   "����"
            Height          =   380
            Left            =   7680
            TabIndex        =   36
            ToolTipText     =   "����SQL���ݿ�"
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
            Caption         =   "����"
            Height          =   380
            Left            =   1800
            TabIndex        =   32
            ToolTipText     =   "���������Ͽ��õ�SQL������"
            Top             =   510
            Width           =   615
         End
         Begin VB.CommandButton cmdRefreshDB 
            Caption         =   "ˢ��"
            Height          =   380
            Left            =   6960
            TabIndex        =   31
            ToolTipText     =   "ˢ�µ�ǰ���ݿ�ı�"
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
            Caption         =   "����:"
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   40
            Top             =   135
            Width           =   1005
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "�û���:"
            Height          =   255
            Index           =   1
            Left            =   2475
            TabIndex        =   39
            Top             =   135
            Width           =   1005
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "���ݿ�:"
            Height          =   255
            Index           =   0
            Left            =   5160
            TabIndex        =   38
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "���÷�����:"
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
            Caption         =   "ѡ��"
            Height          =   380
            Left            =   60
            TabIndex        =   23
            ToolTipText     =   "ѡ��Ҫ�򿪵�ACCESS���ݿ�"
            Top             =   510
            Width           =   615
         End
         Begin VB.CommandButton cmdOpenMdb 
            Caption         =   "����"
            Height          =   380
            Left            =   7680
            TabIndex        =   22
            ToolTipText     =   "���ӵ�ǰACCESS���ݿ�"
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
            Caption         =   "�û���:"
            Height          =   255
            Left            =   480
            TabIndex        =   28
            Top             =   615
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "����:"
            Height          =   255
            Left            =   2640
            TabIndex        =   27
            Top             =   615
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "���ݿ��ļ���:"
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
            Name            =   "����"
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
      ToolTipText     =   "�½�һ�����ݷ�����ʵ��"
      Top             =   120
      Width           =   300
   End
   Begin VB.CheckBox chk 
      Caption         =   "��ʾ����"
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
      Caption         =   "�洢����"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Top             =   3540
      Width           =   1095
   End
   Begin VB.OptionButton opt 
      Caption         =   "��ͼ"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   3540
      Width           =   855
   End
   Begin VB.OptionButton opt 
      Caption         =   "��"
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
      ToolTipText     =   "�ڴ�����SQL���"
      Top             =   480
      Width           =   6390
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "(F5)����"
      Height          =   375
      Left            =   1980
      TabIndex        =   4
      ToolTipText     =   "ֱ������SQL���"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "SQL��ʷ��¼:"
      Height          =   240
      Index           =   3
      Left            =   8850
      TabIndex        =   14
      Top             =   165
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL���:"
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
         Caption         =   "����SQL��䵽����"
         Index           =   0
      End
   End
   Begin VB.Menu mm2 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu m2 
         Caption         =   "��ʾ���м�¼"
         Index           =   0
      End
      Begin VB.Menu m2 
         Caption         =   "��ʾȫ���ֶ�"
         Index           =   1
      End
      Begin VB.Menu m2 
         Caption         =   "��ʾ��ͬ����"
         Index           =   2
      End
   End
   Begin VB.Menu mm3 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu m3 
         Caption         =   "������Excel"
         Index           =   0
      End
      Begin VB.Menu m3 
         Caption         =   "������Xml"
         Index           =   1
      End
      Begin VB.Menu m3 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu m3 
         Caption         =   "ɾ����ǰ����"
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

Dim Rs As Recordset                 '�����ݴ����ݼ�
Dim Rslog As New Recordset          '����������¼��SQL����¼
Dim RsSqlLog As New Recordset       '����SQL���

Dim iOpenID As Long                 '��¼��ǰ�����ݿ��ID
Dim sLogFileName As String          '��¼�����ļ��������ⷴ����ȡ

Dim m_sTempSql As String            '��ʱ��¼SQL��䴮,����ѡ���ֶ�ʱ��ʾ�ֶ�����
Dim m_RsDuplate As Recordset        '��¼�ظ�����

Dim iWndCounter As Long             '��¼ʵ��������

'���ݿ��б�õ�����
Private Sub cboDb_GotFocus()
   On Error GoTo cboDb_GotFocus_Error

'================================================================���������
    If cboSQL.Text = "" Or txtID = "" Then Exit Sub
    cboDb.Clear
    
    '�����������ݿ⵽��������
    D.cyRsToCtl D.cyGetSQLDataBaseNameToRs(cboSQL.Text, txtID, txtPw), cboDb
    
    '�������ݿ�������
    W.cyWndAction cboDb.hWnd, Cbo_PopupList
    
    'ˢ��ҳ��
    Me.Refresh
'================================================================���������
ExitPrc:

    'ͳһ�˳���
    Exit Sub
cboDb_GotFocus_Error:

    '����ȫ�ֵĴ��������
    Call ErrHandler(Err.Number, Err.Description, " cboDb_GotFocus")
    
End Sub

Private Sub chk_Click(Index As Integer)
    If Index = 0 Then
        If chk(Index).Value = vbChecked And chk(Index).Visible = True Then
        '��ǰѡ������ʾ�洢���̵�����,��������ʾ���խ
        
            txtSpContent.Height = dgTable.Height
            dgTable.Width = 577
            txtSpContent.Visible = True
    
        Else
        '��������°������
        
            dgTable.Width = Me.Width / Screen.TwipsPerPixelX - 25
            dgTable.Height = Me.Height / Screen.TwipsPerPixelY - 345
            txtSpContent.Visible = False
    
        End If
        
    End If
        
End Sub

Private Sub cmd_Click(Index As Integer)
    
   On Error GoTo cmd_Click_Error
    Screen.MousePointer = vbHourglass

'================================================================���������

    If Index = 0 Then
        '�½�һ�����ݿ������ʵ��
        Dim a As New frmDebugDB
        
        If Left(Me.Caption, 1) = "[" Then
        '�Ѿ���һ��ʵ��
        
        Else
        '����ʵ��
            Me.Caption = "[1] - " & Me.Caption
        End If
        
        a.Caption = Replace(Me.Caption, "[1]", "[2]")
        
        a.Show
    
    ElseIf Index = 2 Then
        '�������ݿ��б��˫���¼������ݿ�
        lv_DblClick 0
        
        '��յ�ǰ��ʾ������
        Set dgTable.DataSource = Nothing

    ElseIf Index = 3 Then
        If iOpenID = 0 Then
            Call MsgBox("����ѡ���ɾ�������ݿ�����!", vbInformation Or vbSystemModal, "")
            Exit Sub
        
        End If
        
        Select Case MsgBox("�Ƿ�ȷ��Ҫɾ�������ݿ����ӣ�", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton2, "����")
        
            Case vbYes
                Set dgTable.DataSource = Rslog
                
                Rslog.Delete
                
                '�������ݿ��¼���ļ�
                D.cyRsStoreToFile Rslog, sLogFileName & ".Bin", BinaryFile
                
                '��յ�ǰ��ʾ������
                Set dgTable.DataSource = Nothing
                
                '���´����ݿ����ӱ����
                Set Rslog = D.cyRsGetFromFile(sLogFileName & ".Bin")
                
                '�������ݿ����Ӽ���
                Rslog.Filter = "ID>0"
                
                '��ʾ�б�����
                Rslog.Sort = "������,���ݿ�"
                
                If Me.Picture2.Visible Then
                    Rslog.Filter = "����=2"
                
                Else
                    Rslog.Filter = "����=1"
                
                End If
                
                '���±�/��ͼ/�洢�����б�
                lv(0).ListItems.Clear
                D.cyRsToCtl Rslog, lv(0)
                
            Case vbNo
        
        End Select
    
    ElseIf Index = 4 Then
        'ȡ����ʾ
        Set dgTable.DataSource = Nothing
        Picture3.Visible = False

    End If
'================================================================���������
ExitPrc:

    'ͳһ�˳���
    Screen.MousePointer = 0
    Exit Sub
cmd_Click_Error:

    '����ȫ�ֵĴ��������
    Screen.MousePointer = 0
    Call ErrHandler(Err.Number, Err.Description, " cmd_Click")

End Sub

Private Sub cmdSearch_Click()
   On Error GoTo cmdSearch_Click_Error
'================================================================������̴���
    Dim Rs  As Recordset
    
    '���ҵ�ǰ���õ�SQL������
    Set Rs = D.cyGetSQLServerListToRs
    
    If Rs.RecordCount = 0 Then
        Call MsgBox("δ���ֿ��õ�SQL������!", vbInformation Or vbSystemModal, "")
        
    Else
        D.cyRsToCtl Rs, cboSQL
        
    End If

'================================================================������̴���
   On Error GoTo 0
   Exit Sub
cmdSearch_Click_Error:
    If Err.Number = 91 Then
        Call MsgBox("δ���ֿ��õ�SQL������!", vbInformation Or vbSystemModal, "")
    
    Else
            Call ErrHandler(Err.Number, Err.Description, "��Form frmDB���еĺ�����cmdSearch_Click��")
        
    End If
    
End Sub

'ִ��Sql���
Private Sub ExecSql()

   On Error GoTo ExecSql_Error
'================================================================������̴���
    Dim sStr As String
    Dim iMaxID As Long


    If txtSql = "" Then GoTo ExitPrc
    
    '��ѡ���ı������临�Ƶ�������
    If txtSql.SelLength > 0 Then W.cyWndAction txtSql.hWnd, Txt_Copy

    '��SQL����ʽ��Ϊ����
    sStr = txtSql

    '��ʽ��Sql���
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
    
    sStr = Replace(sStr, "��", "=")
    sStr = Replace(sStr, "��", "'")
    sStr = Replace(sStr, "��", "'")
    sStr = Replace(sStr, "��", " ")
    
    txtSql = sStr
    
    'ִ��SQL���
    Set Rs = D.cyGetRs(sStr)
    D.cyRsToCtl Rs, dgTable

    lblRsCount = Rs.RecordCount & "����¼"
    
'================================================================������̴���
   
ExitPrc:
   Exit Sub
   
ExecSql_Error:
    Screen.MousePointer = 0
    
    '������ʾ�洢�����ڲ��׳��Ĵ�����ʾ
        Call ErrHandler(Err.Number, Err.Description, "��Form frmDebugDB���еĺ�����ExecSql��")
    
End Sub

Private Sub cmdExec_Click()
    
   On Error GoTo cmdExec_Click_Error
    Screen.MousePointer = vbHourglass

'================================================================���������
    'ȥ�����׵Ļس�
    Do While Left(txtSql, 1) = Chr(13)
        txtSql = Replace(txtSql, vbCrLf, "")
        
    Loop
    
    '����Ƿ���ѡ�����е����
    If Len(txtSql) = 0 Then
        Call MsgBox("����ѡ��Ҫִ�е�SQL���.", vbInformation Or vbSystemModal, "")
        W.cyWndAction txtSqlLog.hWnd, Wnd_Flash
        txtSqlLog.SetFocus
        Exit Sub
    
    End If
    
    'ִ��Sql���
    Call ExecSql
    
    '��¼��ִ�е�Sql
    
    If InStr(1, txtSqlLog, txtSql) = 0 Then
    '�����û������������
        txtSqlLog = "----------------------------------" & vbCrLf & vbCrLf & txtSql & vbCrLf & vbCrLf & "----------------------------------" & vbCrLf & txtSqlLog
    
    End If
    
    txtSql.SetFocus
    
    '��������ݿ�֮ǰû��SQL��¼�������һ��
    If RsSqlLog.RecordCount = 0 Then
        RsSqlLog.AddNew
        RsSqlLog(0) = iOpenID
        RsSqlLog(1) = txtSqlLog
        RsSqlLog.Update
        
    Else
    '�������
    
        '����SQL��¼
        RsSqlLog(1) = txtSqlLog
        
        '����SQL
        If RsSqlLog.RecordCount > 0 Then txtSqlLog = RsSqlLog(1)
    
    End If
    
    RsSqlLog.Filter = "ID>0"
    D.cyRsStoreToFile RsSqlLog, sLogFileName & cboDb & ".BinLog", BinaryFile
'================================================================���������
ExitPrc:

    'ͳһ�˳���
    Screen.MousePointer = 0
    Exit Sub
cmdExec_Click_Error:

    '����ȫ�ֵĴ��������
    Screen.MousePointer = 0
        Call ErrHandler(Err.Number, Err.Description, " cmdExec_Click")

End Sub

Private Sub cmdOpenMdb_Click()

On Error GoTo cmdOpenMdb_Click_Error
'================================================================������̴���

    Dim iMaxID As Long
    
    If Not F.cyFileExist(txtFile) Then
        Call MsgBox("���ݿ��ļ�������!", vbCritical Or vbSystemModal, "")
        GoTo ExitPrc
        
    End If
    
    If txtAccessPW <> "" Then '����������
        If txtAccessID <> "" Then 'ʹ���û����ʺ�
            D.cyConnectAccess txtFile, , txtAccessID, txtAccessPW
        Else
            D.cyConnectAccess txtFile, txtAccessPW
        End If
        
    Else    'û������
        D.cyConnectAccess txtFile
        
    End If
    
    '�������ݿ����Ӽ���
    Rslog.Filter = "ID>0"
    
    '��ʾ�б�����
    Rslog.Sort = "������,���ݿ�"
    
    Rslog.MoveLast
    
    iMaxID = Rslog("ID")

    '��¼���ݿ�����
    Rslog.Filter = "���ݿ�='" & txtFile & "' and �ʺ�='" & txtAccessID & "' and ����='" & sS.cyStrEncrypt("cyDLL", txtAccessPW) & "'"
    
    '��������������
    If Rslog.RecordCount = 0 Then
        Rslog.Filter = "ID>0 and ����=1"
        
        '���⵱�½����ݼ�ʱ��¼��IMAXIDΪ��
    On Error Resume Next

        Rslog.AddNew
        Rslog("ID") = iMaxID + 1
        Rslog("����") = 1
        Rslog("���ݿ�") = txtFile
        Rslog("�ʺ�") = txtAccessID
        Rslog("����") = sS.cyStrEncrypt("cyDLL", txtAccessPW)
        iOpenID = Rslog("ID")
        Rslog.Update
    End If
    
    '�������
    txtAccessPW = ""
    lv(1).ListItems.Clear
    
    '���������ݿ���б�
    D.cyRsToCtl D.cyGetTableNameToRs, lstTable

    Rslog.Filter = "���� = 3 and ����ID = " & iOpenID

    '�������ݿ��б�
    Rslog.Filter = "ID<100000"
    D.cyRsStoreToFile Rslog, sLogFileName & ".Bin", BinaryFile
    Call cmd_Click(4)
    
    '����SQL����¼
    
    '�����ݿ����ӱ����
    Set RsSqlLog = D.cyRsGetFromFile(sLogFileName & ".BinLog")

    '���û��SQL��¼�����һ��
    If RsSqlLog.RecordCount = 0 Then
        RsSqlLog.AddNew
        RsSqlLog(0) = iOpenID
        RsSqlLog.Update
    Else
        '����SQL��¼
        RsSqlLog.Filter = "ID=" & iOpenID
        
        '����SQL
        If RsSqlLog.RecordCount > 0 Then txtSqlLog = RsSqlLog(1)
    
    End If

    Me.Caption = txtFile & " - ���ݷ�����"

'================================================================������̴���
   On Error GoTo 0
ExitPrc:
   Exit Sub
   
cmdOpenMdb_Click_Error:
        Call ErrHandler(Err.Number, Err.Description, "��Form frmDebugDB���еĺ�����cmdOpenMdb_Click��")
        
End Sub

Private Sub cmdConnectDB_Click()
    Dim Rs As New Recordset
    Dim iMaxID As Long
   
   On Error GoTo cmdConnectDB_Click_Error
'================================================================������̴���

    opt(0).Value = True
    lv(1).ListItems.Clear
    D.cyConnectSqlServer cboSQL, cboDb, txtID, txtPw
    
    '��ʾ��
    D.cyRsToCtl D.cyGetTableNameToRs, lv(1)
    
    '�û�ѡ���˲���¼���ݿ�����
    If Not F.cyFileExist(sLogFileName & ".Bin") Then
        Picture3.Visible = False
    
    Else
    
        Rslog.Filter = "ID>0"
        Rslog.Sort = "ID"
        Rslog.MoveLast
    
        iMaxID = Rslog("ID")
        
        Rslog.Filter = "������='" & cboSQL & "' and ���ݿ�='" & cboDb & "' and �ʺ�='" & txtID & "' and ����='" & sS.cyStrEncrypt("cyDLL", txtPw) & "'"
        
        '��������������
        If Rslog.RecordCount = 0 Then
            Rslog.Filter = "ID>0 and ����=2"
            
            '���⵱�½����ݼ�ʱ��¼��IMAXIDΪ��
        On Error Resume Next
            
            Rslog.AddNew
            Rslog("ID") = iMaxID + 1
            Rslog("����") = 2
            Rslog("������") = cboSQL
            Rslog("���ݿ�") = cboDb
            Rslog("�ʺ�") = txtID
            Rslog("����") = sS.cyStrEncrypt("cyDLL", txtPw)
            Rslog.Update
        End If
        txtPw = ""
        
        Rslog.Filter = "���� = 3 and ����ID = " & iOpenID
        
        '�������ݿ��б�
        Rslog.Filter = "ID<100000"
        D.cyRsStoreToFile Rslog, sLogFileName & ".Bin", BinaryFile
        
        Call cmd_Click(4)
        
        '����SQL����¼
        '�����ݿ����ӱ����
        Set RsSqlLog = D.cyRsGetFromFile(sLogFileName & ".BinLog")
    
        If F.cyFileExist(sLogFileName & cboDb & ".BinLog") Then
        '�����ݿ��¼�Ѵ���
            Set Rs = D.cyRsGetFromFile(sLogFileName & cboDb & ".BinLog")
            If Rs.RecordCount > 0 Then txtSqlLog = Rs(1)
        
        End If
    
    End If
    
    Me.Caption = cboDb & "/" & cboSQL & " - ���ݷ�����"
    
'================================================================������̴���
   On Error GoTo 0
   Exit Sub
   
cmdConnectDB_Click_Error:
    Call ErrHandler(Err.Number, Err.Description, "��Form frmDebugDB���еĺ�����cmdConnectDB_Click��")
    
End Sub

Private Sub cmdSel_Click()
        Dim sStr As String
   On Error GoTo cmdSel_Click_Error
'================================================================������̴���

    sStr = F.cyDialogOpen(Me.hWnd, "��ѡ��Ҫ�򿪵����ݿ�", "*.mdb|*.mdb")
    txtFile = sStr
    txtAccessID = ""
    txtAccessPW = ""
    cmdOpenMdb.SetFocus
    
'================================================================������̴���
   Exit Sub
   
cmdSel_Click_Error:
    Call ErrHandler(Err.Number, Err.Description, "��Form frmDebugDB���еĺ�����cmdSel_Click��")
        
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

    '��ʼ����ʾ�б�
    W.cySetListviewWidths lv(1), "���ݱ�/��ͼ", "181"
    W.cySetListviewWidths lv(2), "������;��������", "150;140"
    
    '������һ��
    sLogFileName = F.cyGetSpecialFolder(Personal) & sS.cyMD5(S.cyGetComputerName & S.cyGetUserName)
    
    If Not F.cyFileExist(sLogFileName & ".Bin") Then
    '���ݿ��¼�ļ�������
        
        Select Case MsgBox("�Ƿ��ڱ����������ݿ�������Ϣ?", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton2, "")
        
            Case vbYes
            
            
                '�½�һ����¼
                Set Rslog = Nothing
                Rslog.Fields.Append "ID", adInteger, 2
                Rslog.Fields.Append "����", adInteger, 2
                Rslog.Fields.Append "����ID", adInteger, 2
                Rslog.Fields.Append "������", adVarChar, 20
                Rslog.Fields.Append "���ݿ�", adVarChar, 255
                Rslog.Fields.Append "�ʺ�", adVarChar, 50
                Rslog.Fields.Append "����", adVarChar, 50
                Rslog.Fields.Append "����", adLongVarChar, 2096
                Rslog.CursorLocation = adUseClient
                Rslog.Open
                Rslog.AddNew
                Rslog("ID") = 1
                Rslog.Update
                D.cyRsStoreToFile Rslog, sLogFileName & ".Bin", BinaryFile
        
            Case vbNo
        
        End Select
    
    End If
        
    'ȱʡ��ʾSQL����
    Call pic_Click(0)
    
    '������+1
    iWndCounter = iWndCounter + 1
    
    Debug.Print iWndCounter
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
    If chk(0).Value = vbChecked Then
    '��ǰѡ������ʾ�洢���̵�����
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
    
    '����Sql��¼
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
'        'ѡ��ȫ��
'
'        lstField.Selected(0) = False
'
'        For i = 1 To lstField.ListCount - 1
'
'            '������ѡ��
'            lstField.Selected(i) = True
'
'        Next
'
'    End If
    
    For i = 0 To lstField.ListCount - 1
    
        '������ѡ��
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

'================================================================���������
If Index = 0 Then

    If lv(0).ListItems.Count = 0 Then GoTo ExitPrc
    
    '�򿪵�ǰѡ������ݿ�
    Rslog.Find "ID=" & lv(0).SelectedItem.Text
    
    If Picture1.Visible = True Then
    '��ǰ����ACCESS
        
        '������ǰѡ������ݿ���Ϣ
        txtFile = lv(0).SelectedItem.SubItems(4)
        txtAccessID = lv(0).SelectedItem.SubItems(5)
        txtAccessPW = sS.cyStrDecrypt("cyDLL", lv(0).SelectedItem.SubItems(5))
    
    ElseIf Picture2.Visible = True Then
    '��ǰ����SQL
    
        '������ǰѡ������ݿ���Ϣ
        cboSQL = lv(0).SelectedItem.SubItems(3)
        cboDb.Text = lv(0).SelectedItem.SubItems(4)
        txtID = lv(0).SelectedItem.SubItems(5)
        txtPw = sS.cyStrDecrypt("cyDLL", lv(0).SelectedItem.SubItems(6))
    
    End If
    
    '��¼��ǰ�򿪵ļ�¼ID
    iOpenID = lv(0).SelectedItem.Text

ElseIf Index = 3 Then
    
    Dim i As Long
    Dim sStr As String
    
    For i = 0 To lv(Index).ListItems.Count - 2
    
        '������ѡ��
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
'================================================================���������
ExitPrc:

    'ͳһ�˳���
    Screen.MousePointer = 0
    Exit Sub
lv_Click_Error:

    '����ȫ�ֵĴ��������
    Screen.MousePointer = 0
    Call ErrHandler(Err.Number, Err.Description, " lv_Click")

End Sub

Private Sub lv_DblClick(Index As Integer)

'���û�б������ݿ��¼����˫��ʱ���Ըö���
   On Error GoTo lv_DblClick_Error
    Screen.MousePointer = vbHourglass

'================================================================���������
    If lv(0).ListItems.Count = 0 Then GoTo ExitPrc
    
    If Picture1.Visible = True Then
    '��ǰ����ACCESS
    
        '�ȶ�����ǰ�����ݿ�
        Call lv_Click(0)
        '�ٴ�
        Call cmdOpenMdb_Click
        
    ElseIf Picture2.Visible = True Then
    '��ǰ����SQL
        
        '�ȶ�����ǰ�����ݿ�
        Call lv_Click(0)
        '�ٴ�
        Call cmdConnectDB_Click
        
    End If
    
'================================================================���������
ExitPrc:

    'ͳһ�˳���
    Screen.MousePointer = 0
    Exit Sub
lv_DblClick_Error:

    '����ȫ�ֵĴ��������
    Screen.MousePointer = 0
    Call ErrHandler(Err.Number, Err.Description, " lv_DblClick")

End Sub

Private Sub lv_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo lv_MouseUp_Error

'================================================================���������
If Index = 1 Then

    If lv(1).ListItems.Count = 0 Then GoTo ExitPrc
    If lv(1).SelectedItem = "" Then GoTo ExitPrc
    
       
    If Button = 1 And Shift = 2 Then
    'Ctl+�������ʾ�ñ����м�¼
        txtSql = "SELECT  * FROM [" & lv(1).SelectedItem & "]"
            
    ElseIf Button = 2 And Shift = 0 Then
        PopupMenu mm2

    Else
        If opt(2).Value = True Then
        '�洢���̣�����������б�
            
            Dim i As Long
            Dim sSQL As String
            Dim sStr As String
            sSQL = "select  ��������=b.name " + _
                   "  ,��������=c.name " + _
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
            
            '�����洢��������
            sSQL = "select c.text, c.encrypted, c.number,  " + _
                   "xtype=convert(nchar(2), o.xtype),  " + _
                   "datalength(c.text), convert(varbinary(8000),  " + _
                   "c.text), 0 from dbo.syscomments c, dbo.sysobjects o  " + _
                   "where o.id = c.id and c.id = object_id('" & lv(1).SelectedItem & "')  " + _
                   "order by c.number, c.colid option(robust plan) "
                   
            txtSpContent = D.cyGetRsOneField(sSQL, "")
            
            '�Զ���ɲ������ı�
            For i = 1 To lv(2).ListItems.Count
                sStr = "''" & "," & sStr
            Next
            
            If sStr <> "" Then
            '�в���,��ȥ������,
                sStr = Left(sStr, Len(sStr) - 1)
                
            End If
            
            '����洢�������ִ��ո����ͷ[]������
            txtSql = IIf(InStr(1, lv(1).SelectedItem, " ") > 0, "[" & lv(1).SelectedItem & "]", lv(1).SelectedItem) & " " & sStr
        
        Else
        '�����ͼ
        
            '���õ���ĵ�һ���ֶ���
            Set Rs = D.cyGetRs(Replace(Replace("SELECT TOP 1 * FROM [" & lv(1).SelectedItem & "]", "[[", "["), "]]", "]"))
            
            '��ʾ�ñ�������ֶ�
            lstField.Clear
            For i = 1 To Rs.Fields.Count
                lstField.AddItem Rs.Fields(i - 1).name
            Next
            lstField.AddItem "*", 0
            
            '�������ҵ����ʮ����¼
            txtSql = "SELECT TOP 10 * FROM [" & lv(1).SelectedItem & "]" & vbCrLf & vbCrLf & "ORDER BY  " & Rs(0).name & " DESC"
            
        End If
        
    End If
    
    txtSql = Replace(txtSql, "[[", "[")
    txtSql = Replace(txtSql, "]]", "]")
        
    '��ʱ���Sql���,ѡ���ֶ�ʱ�����滻
    m_sTempSql = txtSql
        
    If Left(UCase(txtSql), 6) = "SELECT" Then
    '����RS
        Set Rs = D.cyGetRs(txtSql)
        D.cyRsToCtl Rs, dgTable
    
        If Button = 1 And Shift = 2 Then
            lblRsCount = Rs.RecordCount & "����¼"
         
        Else
            lblRsCount = D.cyGetRsOneField(Replace(Split(txtSql, "ORDER BY ")(0), "*", "Count(*)")) & " ����¼"
            
        End If
    
    End If
    
End If
'================================================================���������
ExitPrc:
    'ͳһ�˳���
    Exit Sub
    
lv_MouseUp_Error:

'    '����ȫ�ֵĴ��������
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
            '����RS
                Set Rs = D.cyGetRs(txtSql)
                D.cyRsToCtl Rs, dgTable
            
                If Button = 1 And Shift = 2 Then
                    lblRsCount = Rs.RecordCount & "����¼"
                 
                Else
                    lblRsCount = D.cyGetRsOneField(Replace(txtSql, "*", "Count(*)")) & "����¼"
                    
                End If
            
            Else
            'ִ��SQL���
                D.cyExeCute txtSql
                
            End If
        
        Case 22
            If Left(UCase(txtSql), 6) = "SELECT" Then
            '����RS
                
                For i = 0 To Rs.Fields.Count - 1
                    sTemp = sTemp & " ," & IIf(InStr(1, Rs(i).name, " ") > 0, "[" & Rs(i).name & "]", Rs(i).name)
                Next
                sTemp = Replace(sTemp, " ,", "", , 1)
                
                txtSql = Replace(txtSql, "*", sTemp & " ", , 1)
                
            End If
    
            
    End Select
'================================================================���������
Exit Sub:

    'ͳһ�˳���
    Screen.MousePointer = 0
    Exit Sub
    
Menu_cyMenuClick_Error:

    '����ȫ�ֵĴ��������
    Screen.MousePointer = 0
    Call ErrHandler(Err.Number, Err.Description, " Menu_cyMenuClick")

End Sub

Private Sub m1_Click(Index As Integer)
    If Index = 0 Then
        sStr = txtSql
        '�任Ϊ���鲢���ͷβ�ַ�
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
        
        '��ӱ�����
        sStr = vbTab & "dim sSQL as string " & vbCrLf & vbTab & sStr
        
        '��ŵ���������
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
        '����RS
            Set Rs = D.cyGetRs(txtSql)
            D.cyRsToCtl Rs, dgTable
        
            If Button = 1 And Shift = 2 Then
                lblRsCount = Rs.RecordCount & "����¼"
             
            Else
                lblRsCount = D.cyGetRsOneField(Replace(txtSql, "*", "Count(*)")) & "����¼"
                
            End If
        
        Else
        'ִ��SQL���
            D.cyExeCute txtSql
            
        End If
        
    ElseIf Index = 1 Then
        If Left(UCase(txtSql), 6) = "SELECT" Then
        '����RS
            
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
        
            '������ѡ��
            If lstField.Selected(i) Then sStr = sStr & " , " & lstField.List(i)
        
        Next
        
        If sStr = "" Then
            txtSql = m_sTempSql
            
        Else
            sStr = " " & Right(sStr, Len(sStr) - 3)
            txtSql = "SELECT " & sStr & " , COUNT(" & Split(sStr, " ,")(0) & " ) AS �ظ����� " & vbCrLf & _
                    "FROM " & lv(1).SelectedItem & " " & vbCrLf & _
                    "GROUP BY" & sStr & " HAVING COUNT(" & Split(sStr, ",")(0) & " )>1"
        
            Set m_RsDuplate = D.cyGetRs(txtSql)
            lv(3).ListItems.Clear
            
            If m_RsDuplate.Fields.Count > lv(3).ColumnHeaders.Count Then
            '�����������ͬ
            
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

'================================================================���������
If Index = 0 Then
    On Error Resume Next
    
        If Rs.RecordCount = 0 Then
            Call MsgBox("û�����ݿɹ�����!", vbCritical Or vbSystemModal, "")
            Exit Sub
        
        End If
        
        Set RsTemp = Rs
        
        '�ļ���
        sStr = F.cyDialogSave(Me.hWnd, "����XLS", "*.XLS|*.XLS")
        If sStr = "" Then Exit Sub
        
        Screen.MousePointer = vbHourglass

        '�Ͽ������ݿ������
        Set RsTemp.ActiveConnection = Nothing
        RsTemp.MoveFirst
        D.cyRsToExcel RsTemp, sStr
        Screen.MousePointer = 0
        
        Call MsgBox("�ѳɹ�������ǰ���ݵ�Excel!", vbInformation Or vbSystemModal, "")

ElseIf Index = 1 Then
    On Error Resume Next
        
        If Rs.RecordCount = 0 Then
            Call MsgBox("û�����ݿɹ�����!", vbCritical Or vbSystemModal, "")
            Exit Sub
        
        End If
        
        Set RsTemp = Rs
        
        '�ļ���
        sStr = F.cyDialogSave(Me.hWnd, "����Xml", "*.Xml|*.Xml")
        If sStr = "" Then Exit Sub
        
        
        '�Ͽ������ݿ������
        Set RsTemp.ActiveConnection = Nothing
        RsTemp.MoveFirst
        D.cyRsStoreToFile RsTemp, sStr, XmlFile
        Screen.MousePointer = 0
        
        Call MsgBox("�ѳɹ�������ǰ���ݵ�Xml!", vbInformation Or vbSystemModal, "")


ElseIf Index = 3 Then

        '����ɾ����ǰ���ݼ��е�����
        If Rs.RecordCount = 0 Then Exit Sub
        Select Case MsgBox("�˲�����ɾ����ǰ��ʾ����������,�Ƿ�ȷ��ɾ��?", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton2, "")
        
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
'================================================================���������
ExitPrc:

    'ͳһ�˳���
    Screen.MousePointer = 0
    Exit Sub
m3_Click_Error:

    '����ȫ�ֵĴ��������
    Screen.MousePointer = 0
    Call ErrHandler(Err.Number, Err.Description, " m3_Click")

End Sub

Private Sub opt_Click(Index As Integer)
On Error GoTo Pass

    '��ձ��б�
    lv(1).ListItems.Clear
    
    '��ղ����б�
    lv(2).ListItems.Clear
    
    '���SQL���
    txtSql = ""
    
    If Index = 0 Then
    '�������б�
        D.cyRsToCtl D.cyGetTableNameToRs, lv(1)
        chk(0).Visible = False
        
    ElseIf Index = 1 Then
    '������ͼ�б�
        D.cyRsToCtl D.cyGetRs("Select name  From sysobjects WHERE XTYPE='V' AND category='0' order by name"), lv(1)
        chk(0).Visible = False
        
    ElseIf Index = 2 Then
    '�����洢�����б�
        D.cyRsToCtl D.cyGetRs("select name from sysobjects where xtype='p' and status>0 order by name"), lv(1)
        chk(0).Visible = True
    
    End If
    
    chk_Click 0
    
Pass:
'����δ�������ݿ�ʱ����
End Sub

Private Sub pic_Click(Index As Integer)

    '�л����ݿ�ǰ��ռ�¼
    txtSqlLog = ""
    txtSql = ""

    '�û�ѡ���˲���¼���ݿ�����
    If Not F.cyFileExist(sLogFileName & ".Bin") Then
    
        If Index = 1 Then
            '�����ʵ����п�
            W.cySetListviewWidths lv(0), ";;;;���ݿ�;�˻�;����;;", "0;0;0;0;6500;2520;0;0"
            Picture1.Visible = True
            Picture2.Visible = False
            Picture3.Visible = True
            Picture3.ZOrder 0
            
        ElseIf Index = 0 Then
            '�����ʵ����п�
            W.cySetListviewWidths lv(0), ";;;���ݿ�;�˻�;����;;", "0;0;0;3075.024;2564.788;2520;0;0"
            Picture1.Visible = False
            Picture2.Visible = True
            Picture3.Visible = True
            Picture3.ZOrder 0
        
        End If
    
        Exit Sub
    
    End If


    '�����ݿ����ӱ����
    Set Rslog = D.cyRsGetFromFile(sLogFileName & ".Bin")
    
    If Index = 1 Then
    '��ʾACCESS���ѱ���������
    
        '���SQL�����ʾ
        txtSqlLog = ""
        Picture1.Visible = True
        Picture2.Visible = False
        Picture3.Visible = True
        Picture3.ZOrder 0
       
        '���û�����Ӽ�¼�����
        If Rslog.RecordCount = 0 Then Exit Sub
        
        '�����ݿ����ӱ����
        Set Rslog = D.cyRsGetFromFile(sLogFileName & ".Bin")
        
        '����ACCESS�����ݿ⼯��
        Rslog.Filter = "ID>0"
        Rslog.Filter = "����=1"
        lv(0).ListItems.Clear
        D.cyRsToCtl Rslog, lv(0)
        
   
    ElseIf Index = 0 Then
    '��ʾSQL���ѱ���������
        
        Picture1.Visible = False
        Picture2.Visible = True
        Picture3.Visible = True
        Picture3.ZOrder 0
        
        '���ֻ��һ����¼�����ʾδ�м�¼���ӣ���
        If Rslog.RecordCount = 0 Then
        
            '�½�һ����¼
            Set Rslog = Nothing
            Rslog.Fields.Append "ID", adInteger, 2
            Rslog.Fields.Append "����", adInteger, 2
            Rslog.Fields.Append "����ID", adInteger, 2
            Rslog.Fields.Append "������", adVarChar, 20
            Rslog.Fields.Append "���ݿ�", adVarChar, 255
            Rslog.Fields.Append "�ʺ�", adVarChar, 50
            Rslog.Fields.Append "����", adVarChar, 50
            Rslog.Fields.Append "����", adLongVarChar, 2096
            Rslog.CursorLocation = adUseClient
            Rslog.Open
            Rslog.AddNew
            Rslog("ID") = 1
            Rslog.Update
            D.cyRsStoreToFile Rslog, sLogFileName & ".Bin", BinaryFile

        End If
        
        '����SQL�����ݿ⼯��
        Rslog.Filter = "ID>0"
        Rslog.Filter = "����=2"
        
        '��ʾ�б�����
        Rslog.Sort = "������,���ݿ�"
        
        lv(0).ListItems.Clear
        D.cyRsToCtl Rslog, lv(0)
        
        '�����ʵ����п�
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
    'F5ִ��
        Call cmdExec_Click
        
    ElseIf KeyCode = 65 And Shift = 2 Then
    'Ctl+A,ȫѡ����
        S.cyKeyBoardAction , vbKeyControl, vbKeyHome
        S.cyKeyBoardAction , vbKeyShift, vbKeyControl, vbKeyEnd
    
    End If
    
End Sub

'ȫ�ִ��������
Public Sub ErrHandler(ByVal ErrorNumber As Long, ByVal ErrorMessage As String, Optional ByVal ErrorModule As String)
On Error Resume Next
    Screen.MousePointer = 0
    
    If ErrorNumber = -2147467259 Then
        Call MsgBox("���ݿ��ѶϿ�����!", vbCritical Or vbSystemModal, "����")
       
    Else
        Call MsgBox("����λ�ã�" & ErrorModule & vbCrLf & "������룺" & ErrorNumber & vbCrLf & "����������" & ErrorMessage, vbCritical Or vbSystemModal, "����")
        
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
    'F5ִ��
        Call cmdExec_Click
    
    ElseIf KeyCode = 65 And Shift = 2 Then
    'Ctl+A
        S.cyKeyBoardAction , vbKeyControl, vbKeyHome
        S.cyKeyBoardAction , vbKeyShift, vbKeyControl, vbKeyEnd
    
    End If

End Sub

