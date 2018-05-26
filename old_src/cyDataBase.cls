VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cyDataBaseEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim cmd As New Command

Enum ExportType
    Export2FoxPro = 1
    Export2Excel = 2
    Export2HTML = 4
    Export2Text = 5
End Enum

Enum DefaultReturnType
    ReturnString = 0
    ReturnLong = 1
    ReturnCurency = 2
End Enum

'-----------------------------------------------------------------------------
'获取当前可用的服务器列表
'-----------------------------------------------------------------------------
Private Const SQL_NEED_DATA As Long = 99&
Private Const SQL_SUCCESS As Long = 0&
Private Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv As Long, phdbc As Long) As Integer
Private Declare Function SQLAllocEnv Lib "odbc32.dll" (phenv As Long) As Integer
Private Declare Function SQLBrowseConnect Lib "odbc32.dll" (ByVal hdbc As Long, ByVal szConnStrIn As String, ByVal cbConnStrIn As Integer, ByVal szConnStrOut As String, ByVal cbConnStrOutMax As Integer, pcbconnstrout As Integer) As Integer
Private Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hdbc As Long) As Integer
Private Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc As Long) As Integer
Private Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv As Long) As Integer
'-----------------------------------------------------------------------------
'获取当前可用的服务器列表
'-----------------------------------------------------------------------------

'-----------------------------------------------------------------------------
'将数据集保存成文件的类型
'-----------------------------------------------------------------------------
Enum RsToFileType
    BinaryFile = 0
    XmlFile = 1
End Enum
'-----------------------------------------------------------------------------
'将数据集保存成文件的类型
'-----------------------------------------------------------------------------

'定义一个过程级的数据库连接
Dim Cn As New ADODB.Connection

'用作存放数据库的连接字
Dim sConnectString As String

'获得当前连接
Public Function cyGetConnection() As Connection
    Set cyGetConnection = Cn
    
End Function

'设置连接
Public Function cySetConnection(ByRef Conn As Connection)
    Set Cn = Conn

End Function

'取得连接字
Public Static Property Get ConnectString() As String
    ConnectString = sConnectString
End Property

'写入连接字
Public Static Property Let ConnectString(ByVal vNewValue As String)
    sConnectString = vNewValue
End Property

Public Sub cyConnectSqlServer(ByVal sServer As String, ByVal sDataBase As String, Optional ByVal sID As String, Optional ByVal sPassWord As String, Optional ByVal iConnectTimeOut As Byte = 5, Optional ByVal iCommandTimeOut As Byte = 5)
    Screen.MousePointer = 13
    '如果连接已打开，则先关闭当前连接
    If Cn.State = adStateOpen Then Cn.Close
    If sID = "" Then
    '如果连接时不提供用户名则以WINDOWS认证方式进行连接
        sConnectString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & sDataBase & ";Data Source=" & sServer
    Else
    '否则以标准SQL认证方式进行连接
        sConnectString = "driver={SQL server};server=" & sServer & ";uid=" & sID & ";pwd=" & sPassWord & ";database=" & sDataBase & ";timeout=5;"
    End If
    '打开与数据库的连接
    Cn.ConnectionTimeout = iConnectTimeOut
    Cn.CommandTimeout = iCommandTimeOut
    Cn.Open sConnectString
    Screen.MousePointer = 0
    
End Sub

'    D.cyConnectAccess "C:\Documents and Settings\cy\桌面\零售[统计部端]\零售[统计部端]\dbsalesta.mdb", "DbRetail0423"
Public Sub cyConnectAccess(ByVal sDataBaseFileName As String, Optional ByVal sDataBasePassWord As String, Optional ByVal sGroupUserName As String, Optional ByVal sGroupPassWord As String)
Screen.MousePointer = 13
    '如果连接已打开，则先关闭当前连接
    If Cn.State = adStateOpen Then Cn.Close
    If sGroupUserName <> "" Then
    '如果提供了工作组帐号则证明是以工作组方式进行连接
        sConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=" + sGroupUserName + ";Password=" & sGroupPassWord & ";Data Source=" + sDataBaseFileName + ";Persist Security Info=False"
    Else
    '否则以普通方式连接数据库（如果提供了数据库密码则根据数据库密码进行连接）
        sConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sDataBaseFileName + ";" + IIf(sDataBasePassWord <> "", "Jet OLEDB:Database Password=" + sDataBasePassWord, "")
    End If
    '打开与数据库的连接
    Cn.Open sConnectString
Screen.MousePointer = 0
End Sub

'重新打开连接
Sub cyCnOpen()
    '如果连接已关闭则重新打开连接
    If Not Cn.State = adStateOpen Then Cn.Open sConnectString
End Sub

'关闭连接
Sub cyCnClose()
    '如果连接已打开则关闭此连接并释放所占用的资源
    If Cn.State = adStateOpen Then
        Cn.Close
        Set Cn = Nothing
    End If
End Sub

'    Set Rs = D.cyGetRs("select * from tblUser")
'根据所提供的SQL语句返回数据集
Function cyGetRs(ByVal sSQL As String) As Recordset
Screen.MousePointer = 13
    Dim Rs As Recordset
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open sSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Set cyGetRs = Rs
Screen.MousePointer = 0
End Function

'Set Rs = D.cyGetSQLUsersNameToRs("(local)", "sa")
Function cyGetSQLUsersNameToRs(ByVal sServer As String, Optional ByVal sID As String, Optional ByVal sPassWord As String) As Recordset
Screen.MousePointer = 13
    
    '如果连接已打开，则先关闭当前连接
    If Cn.State = adStateOpen Then Cn.Close
    If sID = "" Then
    '如果连接时不提供用户名则以WINDOWS认证方式进行连接
        sConnectString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=master;Data Source=" & sServer
    Else
    '否则以标准SQL认证方式进行连接
        sConnectString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & sID & ";PassWord=" & sPassWord & ";Initial Catalog=master;Data Source=" & sServer
    End If
    '打开与数据库的连接
    Cn.Open sConnectString
    '返回数据库名
    Set cyGetSQLUsersNameToRs = cyGetRs("SELECT name FROM sysxlogins")
    
Screen.MousePointer = 0
End Function

'取得SQL服务器的所有数据库
Function cyGetSQLDataBaseNameToRs(ByVal sServer As String, Optional ByVal sID As String, Optional ByVal sPassWord As String) As Recordset
Screen.MousePointer = 13
    '如果连接已打开，则先关闭当前连接
    If Cn.State = adStateOpen Then Cn.Close
    If sID = "" Then
    '如果连接时不提供用户名则以WINDOWS认证方式进行连接
        sConnectString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=master;Data Source=" & sServer
    Else
    '否则以标准SQL认证方式进行连接
        sConnectString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & sID & ";PassWord=" & sPassWord & ";Initial Catalog=master;Data Source=" & sServer
    End If
    '打开与数据库的连接
    Cn.Open sConnectString
    '返回数据库名
    Set cyGetSQLDataBaseNameToRs = cyGetRs("SELECT name FROM sysdatabases")
Screen.MousePointer = 0
End Function

'返回所有可检测到的SQL服务器
Function cyGetSQLServerListToRs() As Recordset
Screen.MousePointer = 13
    Dim Rs As New Recordset
    Dim rc As Integer
    Dim henv As Long
    Dim hdbc As Long
    Dim stCon As String
    Dim stConOut As String
    Dim pcbConOut As Integer
    Dim ichBegin As Integer
    Dim ichEnd As Integer
    Dim stOut As String
    Dim sArray() As String
    rc = SQLAllocEnv(henv)
    rc = SQLAllocConnect(ByVal henv, hdbc)
    stCon = "DRIVER=SQL Server"
    
    ' Get the size of the buffer to create and create it
    rc = SQLBrowseConnect(ByVal hdbc, stCon, Len(stCon), stConOut, Len(stConOut) + 2, pcbConOut)
    stConOut = String$(pcbConOut + 2, vbNullChar)
    
    ' Get the actual server list
    rc = SQLBrowseConnect(ByVal hdbc, stCon, Len(stCon), stConOut, Len(stConOut) + 2, pcbConOut)
    
    If (rc <> SQL_SUCCESS) And (rc <> SQL_NEED_DATA) Then
        ' ignore the errors if any occur
    Else
        ' Parse out the server list
        ichBegin = InStr(InStr(1, UCase(stConOut), "SERVER="), stConOut, "{", vbBinaryCompare)
        stOut = Mid$(stConOut, ichBegin + 1)
        ichEnd = InStr(1, stOut, "}", vbBinaryCompare)
        
        sArray = Split(Left$(stOut, ichEnd - 1), ",")
If UBound(sArray) = -1 Then GoTo Pass
        Dim i As Long
        With Rs
            .CursorLocation = adUseClient
            .Fields.Append "Server", adLongVarWChar, 1
            .Open
                For i = 0 To UBound(sArray)
                    .AddNew 0, sArray(i)
                Next
                If Rs.RecordCount > 0 Then Rs.MoveFirst
            Set cyGetSQLServerListToRs = Rs
        End With
        
    End If

Pass:

    ' Disconnect, free the connection handle, then
    ' free the environment handle.
    rc = SQLDisconnect(hdbc)
    rc = SQLFreeConnect(hdbc)
    rc = SQLFreeEnv(henv)
Screen.MousePointer = 0
End Function

'返回已打开的数据库的所有表名
Function cyGetTableNameToRs() As Recordset
Screen.MousePointer = 13
    Dim Rs As Recordset
    
    '状态为未打开,则不返回内容
    If Cn.State = 0 Then Exit Function
    
    '返回的表名的记录集,由于返回的记录集有3列 ,只有第3列有内容,因此建立一个新的记录集,将表名写进去
    Set Rs = Cn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
    
    Dim RsTmp As New Recordset
    With RsTmp
        .CursorLocation = adUseClient
        '新建一个Field用作存放表名
        .Fields.Append "Table", adLongVarWChar, 1
        '打开这个数据集
        .Open
            Do While Not Rs.EOF
                '添加表名到数据集中
                .AddNew 0, IIf(InStr(1, Rs(2).Value, " ") > 0, "[" + Rs(2).Value + "]", Rs(2).Value)
                Rs.MoveNext
            Loop
            '将指针移到第一条记录
            If RsTmp.RecordCount > 0 Then RsTmp.MoveFirst
        Set cyGetTableNameToRs = RsTmp
    End With
Screen.MousePointer = 0
End Function

Function cyExeCute(ByVal sSQL As String) As Recordset

    If sSQL = "" Then Exit Function
    Dim tmpRs As New Recordset
    Dim cmd As New Command
    Set cmd.ActiveConnection = Cn
    cmd.CommandText = sSQL
    tmpRs.CursorLocation = adUseClient
    Set tmpRs = cmd.Execute(, , adCmdText)
    Set cyExeCute = tmpRs
    
End Function

Public Sub cyRsStoreToFile(Rs As Recordset, sFileName As String, Optional FileType As RsToFileType)
On Error GoTo Err
Screen.MousePointer = 13

     '将数据集保存到不同的文件类型中(二进制类型文件或XML文件)
    Rs.Save sFileName, FileType
    Screen.MousePointer = 0
    Exit Sub
Err:
'该文件已经保存过
If Err.Number = 58 Then
    On Error Resume Next
    Kill sFileName
    '将数据集保存到不同的文件类型中(二进制类型文件或XML文件)
    Rs.Save sFileName, FileType
End If
Screen.MousePointer = 0
End Sub

Public Function cyRsGetFromFile(sFileName As String) As Recordset
Screen.MousePointer = 13
    
    Dim Rs As New Recordset '注意此处一定要加关键字New
    '将数据集从文件中读入
    Rs.Open sFileName, "Provider=MSPersist"
    '返回数据集
    Set cyRsGetFromFile = Rs

Screen.MousePointer = 0
End Function

Sub cyFileStoreToDB(sSQL As String, sFileName As String)
Screen.MousePointer = 13

    Dim Rs As Recordset
    '将文件读入二进制数组
    Dim byteArray() As Byte
    ReDim byteArray(FileLen(sFileName) - 1) As Byte
    Open sFileName For Binary As #1
      Get #1, , byteArray
    Close #1
    
    '将二进制数组添加到数据集中
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open sSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Rs.AddNew Rs(0).name, byteArray
    '更新并关闭数据集
    Rs.Update
    Rs.Close
    '释放资源
    Set Rs = Nothing

Screen.MousePointer = 0
End Sub

Sub cyFileGetFromDB(sSQL As String, sFileName As String)
Screen.MousePointer = 13

    Dim Rs As Recordset
    '定义一个二进制数组
    Dim byteArray() As Byte
    '将数据库二进制字段中的数据写入二进制组组中
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open sSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    byteArray = Rs(0).Value
    '关闭并释放资源
    Rs.Close
    Set Rs = Nothing
        
    '将二进制数组写入到文件中
    Open sFileName For Binary As #1
      Put #1, , byteArray
    Close #1

Screen.MousePointer = 0
End Sub

Sub cyMdbBackUp(ByVal sTargetName As String, Optional ByVal sPassWord As String)
On Error GoTo Err
Screen.MousePointer = 13

    '保存目标文件名
    Dim sTarget As String
    Dim JRO As JRO.JetEngine
    Set JRO = New JRO.JetEngine
    '目标数据库的连接字
    sTarget = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sTargetName & ";Jet OLEDB:Database Password=" & sPassWord
    '关闭原连接,不关闭则不能转换
    Cn.Close
    Set Cn = Nothing
    '执行转換(压缩)
    JRO.CompactDatabase sConnectString, sTarget
    '重新连接
    Call cyCnOpen
    Screen.MousePointer = 0
    Exit Sub
Err:
'该文件已经保存过
If Err.Number = -2147217897 Then
    On Error Resume Next
    Kill sTargetName
    '执行转換(压缩)
    JRO.CompactDatabase sConnectString, sTarget
    '重新连接
    Call cyCnOpen
End If

Screen.MousePointer = 0
End Sub

'函数名称:cyRsToTreeView
'作用：将数据集绑定到TreeView
'===========================================================
'整理时间：2006-12-31 12:59
'  参数：
'           tvTreeView         --  待绑定的组件
'           Rs                 --  须绑定的数据集
'           OwnerIDField       --  本身记录的ID字段名称
'           ParentIDField      --  父记录的ID字段名称
'           sNodeTextField     --  节点字符显示的字段名称
'           bNodeExpanded      --  是否展开


'  例子：            Set Rs = D.cyRsGetFromFile("c:\software.xml")
'                    D.cyRsToTreeView Me.TreeView1, Rs, "ID", "FATHERID", "TEXT", True

'===========================================================
Sub cyRsToTreeView(ByVal tvTreeView As Object, ByVal Rs As Recordset, ByVal OwnerIDField As String, ByVal ParentIDField As String, ByVal sNodeTextField As String, Optional ByVal bNodeExpanded As Boolean)
    Dim i As Long
    tvTreeView.Nodes.Clear
    Rs.MoveFirst
    For i = 1 To Rs.RecordCount
        If Rs(ParentIDField) = 0 Then
            tvTreeView.Nodes.Add , , "_" & Rs(OwnerIDField), Rs(sNodeTextField)
        
        Else
            tvTreeView.Nodes.Add "_" & Rs(ParentIDField), tvwChild, "_" & Rs(OwnerIDField), Rs(sNodeTextField)

        End If
        
        '是否展开该项
        tvTreeView.Nodes(i).Expanded = bNodeExpanded
        
        Rs.MoveNext
        
    Next

End Sub

Sub cyRsToCtl(ByVal Rs As Recordset, ByVal Ctl As Object, Optional ByVal sCellWidthLists As String)
Dim i As Long
Dim j As Long
Dim sType As String     '保存类型名

'得到每格的宽度

Dim iCellWidth() As String
On Error Resume Next
iCellWidth = Split(sCellWidthLists, ";")

'如果返回的数据集没有游标则复制一个RS
If Rs.RecordCount = -1 Then
    Set Rs = cyRsToRsEx(Rs)
End If

'将Rs游标移到最前面
If Rs.RecordCount > 0 Then
    Rs.MoveFirst
End If

'得到控件的类型
sType = TypeName(Ctl)
If sType = "ComboBox" Or sType = "ListBox" Then
        Dim iFieldCount As Long '记录数据集的字段数,如果字段数超过1则检查其是否数字,如果是则添加到ITEMDATA里
        iFieldCount = Rs.Fields.Count
        If iFieldCount = 1 Then '如果只有一个字段则直接写入控件
            '循环写入数据
            Do While Not Rs.EOF
                Ctl.AddItem CStr(Rs(0).Value)
                Rs.MoveNext
            Loop
        
        Else '如果有一个以上字段就检查其第二个字段是否数字,如果是则写进去
            '循环写入数据
            Do While Not Rs.EOF
                Ctl.AddItem CStr(Rs.Fields(0))
                '第一列显示在控件,如果第二列是数字则将其绑定到ItemData
                If IsNumeric(Rs.Fields(1)) Then Ctl.itemData(i) = Rs.Fields(1)
                i = i + 1
                Rs.MoveNext
            Loop
        End If
        Exit Sub
ElseIf sType = "DataCombo" Or sType = "DataList" Then
        Ctl.ListField = Rs.Fields(0).name
        Set Ctl.RowSource = Rs
        Exit Sub
ElseIf sType = "DataGrid" Then
    
    Set Ctl.DataSource = Rs
On Error Resume Next
    For i = 0 To UBound(iCellWidth)
        Ctl.Columns(i).Width = iCellWidth(i)
    Next
ElseIf sType = "MSHFlexGrid" Then
    Set Ctl.DataSource = Rs
On Error Resume Next
    For i = 0 To UBound(iCellWidth)
        Ctl.ColWidth(i) = CLng(iCellWidth(i))
    Next
 
ElseIf sType = "MSFlexGrid" Then
    '循环写入数据
    
    Do While Not Rs.EOF
        Ctl.cols = Rs.Fields.Count + 1
        For i = 0 To Rs.Fields.Count - 1
            Ctl.TextMatrix(Ctl.Rows - 1, i + 1) = Rs.Fields(i)
        Next
        Ctl.Rows = Ctl.Rows + 1
        Rs.MoveNext
    Loop
    Exit Sub
'        For j = 0 To Rs.RecordCount - 1
'        Rs.MoveNext
'        Next
'        Ctl.Rows = Ctl.Rows - 1
ElseIf sType = "ListView" Then
        
        Dim Fld As Field
        Dim FirstRow As Boolean
        Dim iWidth As Long
        Ctl.View = 3 'lvmReport
        Ctl.LabelEdit = 1 'lvwManual
        
        If Ctl.ColumnHeaders.Count = 0 Then
            For j = 0 To Rs.Fields.Count - 1
                If j > UBound(iCellWidth) Then
                    Ctl.ColumnHeaders.Add , , Rs.Fields(j).name            'CREATE COLUMN HEADERS
                
                Else
                    '设定每格的Title,及表格宽
                    Ctl.ColumnHeaders.Add , , Rs.Fields(j).name, iCellWidth(j)           'CREATE COLUMN HEADERS
                
                End If
            Next
            
        End If
        
On Error Resume Next
        For i = 0 To Rs.RecordCount - 1
            FirstRow = True
            For j = 0 To Rs.Fields.Count - 1
                If FirstRow Then                            'FIRST ROW IS A LIST ITEM
                    If Rs(j) <> "" Then
                        Ctl.ListItems.Add , , Rs(j)
                    Else
                        Ctl.ListItems.Add , , ""
                    End If
                    FirstRow = False
                Else
                    If Rs(j) <> "" Then
                        Ctl.ListItems.Item(Ctl.ListItems.Count).SubItems(j) = Rs(j)
                        Ctl.ListItems(Ctl.ListItems.Count).ListSubItems.Add , , Fld.Value   'REMAINING ROWS ARE SUBITEMS
                    Else
                        Ctl.ListItems.Item(Ctl.ListItems.Count).SubItems(j) = ""
                        Ctl.ListItems(Ctl.ListItems.Count).ListSubItems.Add , , ""      'INSERT BLANKS FOR NULLS SO DATA
                    End If                                      'DOESN'T COALESCE TO THE LEFT
                End If
            Next
            Rs.MoveNext
        Next
        Dim W As New cyWndEx
        W.cyWndAction Ctl.hWnd, LV_FullRowSel, 1
        W.cyWndAction Ctl.hWnd, LV_GirdLine, 1
End If
End Sub

Public Sub cySqlBackup(ByVal sDataBase As String, ByVal sBackUpFile As String)
Screen.MousePointer = 13

    Dim sBackup As String
    Screen.MousePointer = 11
    sBackup = "backup database " & sDataBase & " to disk='" + sBackUpFile + "'"
    Cn.Execute sBackup
    
Screen.MousePointer = 0
End Sub

Public Sub cySqlRestore(ByVal sDataBase As String, ByVal sBackUpFile As String)
Screen.MousePointer = 13

    Dim sRestore As String
    Dim Rs As New Recordset
        
    '删除原数据库在服务器的记录
    Rs.Open "select  spid  from  sysprocesses  where  dbid=db_id('" & sDataBase & "')", Cn
    Do While Not Rs.EOF
        Cn.Execute "kill " & Rs("spid")
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
    
    '使用覆盖模式恢复数据库
    sRestore = "restore database " & sDataBase & " from disk='" + sBackUpFile + "'  with REPLACE"
    Cn.Execute sRestore
    
Screen.MousePointer = 0
End Sub

'增加参数
Sub cySP_Append(ByVal sParaName As String, ByVal varParaValue, Optional ByVal ParaType As DataTypeEnum, Optional ByVal ParaDirtion As ParameterDirectionEnum, Optional ByVal iSize As Integer)
    If ParaDirtion = 0 Then
        ParaDirtion = adParamInput
    End If
    
    If iSize = 0 And TypeName(varParaValue) = "String" Then
        Dim sStr As String
        sStr = varParaValue
        iSize = LenB(StrConv(varParaValue, vbFromUnicode))
    Else
        iSize = 20
    End If
    
    If ParaType = 0 Then
        Select Case TypeName(varParaValue)
            Case "Integer"
                ParaType = adInteger
            Case "Long"
                ParaType = adBigInt
            Case "String"
                ParaType = adChar
            Case "Date"
                ParaType = adDate
        End Select
    End If
    cmd.Parameters.Append cmd.CreateParameter(sParaName, ParaType, ParaDirtion, iSize, varParaValue)
End Sub

'执行存储过程
Function cySP_Execute(ByVal SP_Name As String) As Recordset
    
    Set cmd.ActiveConnection = Cn '指示指定的 Command对象当前所属的 Connection对象。
    cmd.CommandText = "sp_Fold_Update" '设置Command对象源。
    cmd.CommandType = adCmdStoredProc '通知提供者CommandText属性有什么，它可能包括Command对象的源类型。设置这个属性优化了该命令的执行。
    cmd.Parameters.Append cmd.CreateParameter("@ID", adInteger, adParamInput, 255, 43788)
    cmd.Parameters.Append cmd.CreateParameter("@OldPage", adInteger, adParamInput, 255, 16)
    cmd.Parameters.Append cmd.CreateParameter("@NewPage", adInteger, adParamInput, 255, 8)
    cmd.Parameters.Append cmd.CreateParameter("@oldPageList", adVarChar, adParamInput, 255, "1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16")
    cmd.Parameters.Append cmd.CreateParameter("@NewPageList", adVarChar, adParamInput, 255, "1;2;3;4;5;6;7;8")
    
    Dim Rs As Recordset
    Set Rs = cmd.Execute



    cmd.ActiveConnection = Cn
    Cn.CursorLocation = adUseClient
    cmd.CommandText = SP_Name
    cmd.CommandType = adCmdStoredProc
    
    Rs.CursorType = 3
    Rs.Open cmd, , 1, 3
    
    '清除参数
    Dim i As Long
    For i = 1 To cmd.Parameters.Count
        cmd.Parameters.Delete 0
    Next
End Function

'开始事务
Sub cyCnBegintrans()
    Cn.BeginTrans
End Sub

'事务回滚
Sub cyCnRollbackTrans()
    Cn.RollbackTrans
End Sub

'事务结束
Sub cyCnCommitTrans()
    Cn.CommitTrans
End Sub


'函数名称:cyMdbRsExport

'作用：将当前的RecordSet导出到其它格式

'##################################################
'整理时间：2004-2-11 15:35:36
'  参数：
'   sSQL As String    --执行的SQL语句
'   iType As ExportType     --类型(FoxBase/Html/Text/Excel)
'   sFileName As String  --目标文件名(txt/xls..)
'##################################################
'  例：
Function cyMdbRsExport(sSQL As String, iType As ExportType, ByVal sFileName As String, Optional ByVal sXlsSheetName As String)

    On Error Resume Next
    Kill sFileName
    On Error GoTo 0
    
'分离文件路径及文件名并组成导出语句
    Dim i As Integer
    Dim sPath As String
    Dim sFile As String
    Dim sTemp As String '存放into等语句
    For i = Len(sFileName) To 1 Step -1
        If Mid(sFileName, i, 1) = "\" Then Exit For
    Next
    sPath = Left(sFileName, i - 1)
    sFile = Mid$(sFileName, i + 1, Len(sFileName) - i)

'分离出表名
    Dim sTableName As String
    sSQL = UCase(Replace(sSQL, "  ", " "))
    i = InStr(1, sSQL, " FROM ")
    i = InStr(i + 5, sSQL, " ")
    If InStr(i + 1, sSQL, " ") = 0 Then
        '表名后没有条件表达式
        sTableName = Right(sSQL, Len(sSQL) - i)
    Else
        '表名后还有 Where 等条件表达式
        sTableName = Mid(sSQL, i + 1, InStr(i + 1, sSQL, " ") - i - 1)
    End If

    '如果有定制表名则替代
    If (sXlsSheetName <> "" And iType = Export2Excel) Then sTableName = sXlsSheetName


Select Case iType
    Case 1
        'Visual FoxPro 3.0          FoxPro 3.0; drive:\path
        sTemp = " INTO [FoxPro 3.0;DATABASE=" + sPath + "].[" + sFile + "] "
    Case 2
        'Microsoft  97         Excel 8.0; drive:\path\filename.xls
        sTemp = " INTO [Excel 8.0;DATABASE=" + sPath + "\" + sFile + "].[" + sTableName + "]"
    Case 4
        'HTML                       Import HTML Import; drive:\path\filename
        'HTML                       Export HTML Export; drive:\path
        sTemp = " INTO [HTML Export;DATABASE=" + IIf(Right(sPath, 1) = "\", sPath, sPath & "\") + "].[" + sFile + "] "
    Case 5
        'Text                       Text; drive:\path
        sTemp = " INTO [text;DATABASE=" + IIf(Right(sPath, 1) = "\", sPath, sPath & "\") + "].[" + sFile + "] "
        'sTemp = " INTO [ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=myODBC;] "
    Case 20
        'ODBC                       [LOGINTIMEOUT=seconds;]        NonePath
    Case 21
        'Microsoft Exchange         Exchange 4.0;MAPILEVEL=folderpath; [TABLETYPE={ 0 | 1 }];[PROFILE=profile;][PWD=password;][DATABASE=database;]

End Select


    i = InStr(1, sSQL, " FROM ")
    sPath = Left(sSQL, i)
    sFile = Right(sSQL, Len(sSQL) - i)
    sSQL = sPath + sTemp + sFile
    '建立最后的连接字


    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = Cn
    cmd.CommandType = adCmdText
    cmd.CommandText = sSQL
    cmd.Execute
    cyMdbRsExport = True
End Function

'复制一个数据集的结构
Function cyRsToRs(ByVal Rs As Recordset) As Recordset
    Set cyRsToRs = Rs.Clone(adLockOptimistic)
    Set cyRsToRs.ActiveConnection = Nothing
End Function

'复制一个没有RECORDCOUNT的数据集结构
Function cyRsToRsEx(ByVal Rs As Recordset) As Recordset
    Dim tmpRs As New Recordset
    Dim iFieldCount As Long
    iFieldCount = Rs.Fields.Count
    Dim i As Long
    
    '新建一个数据集
    tmpRs.CursorLocation = adUseClient
    For i = 0 To iFieldCount - 1
        tmpRs.Fields.Append Rs(i).name, Rs(i).type, Rs(i).DefinedSize, Rs(i).Attributes
    Next
    tmpRs.Open
    
    '循环写入数据
    Do While Not Rs.EOF
        tmpRs.AddNew
        For i = 0 To iFieldCount - 1
            tmpRs(i) = Rs(i)
        Next
        Rs.MoveNext
    Loop
    tmpRs.MoveFirst
    Set cyRsToRsEx = tmpRs
End Function

'计算数据集中某字段的总和
Function cyRsFieldSum(ByVal Rs As Recordset, sField As String)
Screen.MousePointer = 13

    Dim tmpRs As Recordset
    Dim iCount
    Dim i As Long
    Set tmpRs = Rs.Clone(adLockReadOnly)
    If tmpRs.RecordCount = 0 Then Exit Function
    tmpRs.MoveFirst
    For i = 0 To tmpRs.RecordCount - 1
        If Not IsNull(tmpRs(sField).Value) Then iCount = tmpRs(sField).Value + iCount
        tmpRs.MoveNext
    Next
    cyRsFieldSum = iCount

End Function

'函数名称:cyGetRsOneField
'作用：根据SQL语句返回某条记录的某个字段的值(如果没有值则给缺省值)
'  例：
Function cyGetRsOneField(ByVal sSQL As String, Optional ByVal DefaultValue As Variant)
    Dim tmpRs As Recordset
    Set tmpRs = New ADODB.Recordset
    tmpRs.CursorLocation = adUseClient
    tmpRs.CursorType = adOpenKeyset
    tmpRs.LockType = adLockOptimistic
    tmpRs.Open sSQL, Cn, , , adCmdText
    If tmpRs.RecordCount > 0 Then
    '有数据记录返回
        If IsNull(tmpRs(0).Value) Then
        '如果值为NULL,则返回缺省值
            cyGetRsOneField = DefaultValue
        Else
        '不为NULL，则返回该值
            cyGetRsOneField = tmpRs(0).Value
        End If
    Else
        '没有则给个缺省值
        cyGetRsOneField = DefaultValue
        
    End If
End Function

'函数名称:cyExcelToRs
'作用：将Excel文件读入到数据集
'===========================================================
'整理时间：2008-12-20 16:29
'  参数：
'           sFileName       --  Excel文件名
'           sSheetName      --  Excel文件的表名(如不提供,则使用的Sheet1)

'  例子：   Set Rs = D.cyExcelToRs("c:\1.xls", "Sheet1")

'===========================================================
Public Function cyExcelToRs(ByVal sFileName As String, Optional ByVal sSheetName As String) As Recordset
    
On Error GoTo Err
    Const adOpenStatic = 3
    Dim Rs As Recordset
    Dim Conn As Connection
    
    '如果未提供表名则使用缺省表名
    If sSheetName = "" Then sSheetName = "Sheet1"
    
    '将Excel读到数据集
    Set Conn = CreateObject("ADODB.Connection")
    
    Conn.Open "driver={Microsoft Excel Driver (*.xls)};dbq=" & sFileName
    Set Rs = CreateObject("ADODB.Recordset")
    
    '打开Excel的某个表
    Rs.Open "[" & sSheetName & "$]", Conn, adOpenStatic
    
    '返回数据集
    Set cyExcelToRs = Rs
    
    '清空数据
    Set Rs = Nothing
    Set Conn = Nothing
    Exit Function
Err:
    
    If Err.Number = -2147217900 Then
    '表名不正确
        Err.Raise -999, "", "请提供正确的Sheet名称!"
        
    Else
    '其它错误
        Err.Raise Err.Number, Err.Source, Err.Description
        
    End If
    
End Function

'sWidthsList格式如下:为1则打印,为0或为空则不打印.
' 1;;1; 只打印打1/3列
Public Function cyRsToExcel(ByVal Rs As Recordset, ByVal sFileName As String, Optional ByVal sSheetName As String, Optional ByVal sWidthsList As String) As Boolean
Dim tmpRs As Recordset
Dim sA() As String  '保存是否打印列表
Dim exl
Dim xlsWorkBook
Dim xlsWorkSheet
Set tmpRs = Rs.Clone(adLockReadOnly)
Set exl = CreateObject("Excel.Application")

Dim S As New cySystemEx
Dim W As New cyWndEx


Set xlsWorkBook = exl.Workbooks.Add
Set xlsWorkSheet = xlsWorkBook.Worksheets.Add

Dim i As Long
Dim j As Long

    If sWidthsList <> "" Then
    '不为空则表示根据设置
        sA = Split(sWidthsList, ";")
        '所给列表长度与tmprs的字段数不相等,则重定义使其相等
        If UBound(sA) <> tmpRs.Fields.Count Then
            ReDim Preserve sA(tmpRs.Fields.Count - 1)
            
        End If
    
        '先写入字段名
        For j = 0 To tmpRs.Fields.Count - 1
            '不需要检查则直接打印
            xlsWorkSheet.Cells(i + 1, j + 1) = tmpRs.Fields(j).name
            xlsWorkSheet.Columns(Chr(65 + j)).ColumnWidth = IIf(IsNumeric(sA(j)), sA(j), 0)
    
        Next
    
    Else
    '为空则按照原来的列宽
    
        '先写入字段名
        For j = 0 To tmpRs.Fields.Count - 1
            '不需要检查则直接打印
            xlsWorkSheet.Cells(i + 1, j + 1) = tmpRs.Fields(j).name
    
        Next
    
    End If
    
    tmpRs.MoveFirst
    
    If tmpRs.RecordCount > 1000 Then    '大于一千条则显示运行条
        W.cyRunningBar ShowBar
        W.cyRunningBar ShowMsg, "正在导出..."
        For i = 0 To tmpRs.RecordCount - 1
            If i Mod 1000 = 0 Then
                W.cyRunningBar ShowPercent, (i / tmpRs.RecordCount) * 100
            End If
        
        
            '写入字段内容
            For j = 0 To tmpRs.Fields.Count - 1
                '不需要检查则直接打印
                xlsWorkSheet.Cells(i + 2, j + 1) = tmpRs(j).Value
                
            Next
            tmpRs.MoveNext
            
        Next
        W.cyRunningBar CloseBar
    Else
        For i = 0 To tmpRs.RecordCount - 1
            '写入字段内容
            For j = 0 To tmpRs.Fields.Count - 1
                '不需要检查则直接打印
                xlsWorkSheet.Cells(i + 2, j + 1) = tmpRs(j).Value
                
            Next
            tmpRs.MoveNext
            
        Next
    
    End If
    xlsWorkSheet.Range("A1", "Z1").Font.Bold = True
    xlsWorkSheet.Cells.Font.Size = 10
    'Excel.Worksheets.PrintOut
    
    If sSheetName = "" Then
    '如果表名为空则使用缺省的表名
        xlsWorkBook.ActiveSheet.name = "数据导出"
    Else
    '有提供表名则使用提供的表名
        xlsWorkBook.ActiveSheet.name = sSheetName
    End If
    
    On Error Resume Next

    xlsWorkBook.SaveAs sFileName
    xlsWorkBook.Close
    
    Set xlsWorkBook = Nothing
    Set xlsWorkSheet = Nothing
    
    cyRsToExcel = True

    W.cyRunningBar CloseBar

End Function

'函数名称:cyRsAppend
'作用：将两个结构完全一样的数据集合并
'===========================================================
'整理时间：2006-12-31 12:59
'  参数：
'           RsDistination --  合并后的数据集(Rs1)
'           RsSource      --  须合并的数据集(Rs2)

'  例子：        Call D.cyRsAppend(Rs1, Rs2)
'===========================================================
Sub cyRsAppend(ByRef RsDistination As Recordset, ByVal RsSource As Recordset)
    Dim i As Long
    Dim j As Long
    '循环读取待添加的Rs
    
    RsSource.MoveFirst
    
    For i = 1 To RsSource.RecordCount
        '逐条添加
        RsDistination.AddNew
            '逐个字段追加
            For j = 0 To RsDistination.Fields.Count - 1
                RsDistination(j).Value = RsSource(j).Value
            Next
        RsSource.MoveNext
    Next
    '批量更新
    RsDistination.UpdateBatch
    
End Sub

'函数名称:cyDelDoubleDatas
'作用：删除表内某几个字段内容相同的其它行,只保留第一行
'===========================================================
'整理时间：2006-12-31 12:59
'  参数：
'           sTable              --表名
'           sFieldID            --唯一标识的ID字段名，主键
'           sFieldLists         --字段列表,使用,分隔

'  例子：            D.cyDelDoubleDatas "(local)", "pubs", "sa", "123456", "discounts", "ID", "stor_id,lowqty,highqty,discount"

'===========================================================
Sub cyDelDoubleDatas(ByVal sServer As String, ByVal sDataBase As String, ByVal sID As String, ByVal sPassWord As String, ByVal sTable As String, ByVal sFieldID As String, ByVal sFieldLists As String)
    Dim i As Long
    Dim D As New cyDataBaseEx
    Dim Rs As Recordset
    Dim sA() As String
    Dim sStr As String
    Dim sSQL As String
    sFieldLists = Replace(sFieldLists, ";", ",")
    sA = Split(sFieldLists, ",")
    For i = 0 To UBound(sA)
        sStr = sStr & "t1." & sA(i) & " = a." & sA(i) & " AND "
        
    Next
    sStr = Left(sStr, Len(sStr) - 4)
    
    Select Case MsgBox("是否确定删除(字段:" & sFieldLists & ")重复的数据?" & "请先确定已备份数据.", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton2, "")
    
        Case vbYes
    
        Case vbNo
            Exit Sub
            
    End Select
    
    sSQL = "SELECT * " + _
       "FROM " & sTable & " " + _
       "WHERE (" & sFieldID & " NOT IN (SELECT " & sFieldID & " " + _
       "FROM " & sTable & " a LEFT OUTER JOIN (SELECT " & sFieldLists & " " + _
       "FROM " & sTable & " " + _
       "GROUP BY " & sFieldLists & " HAVING COUNT(*) > 1) t1 ON " & sStr + _
       "WHERE a." & sFieldID & " NOT IN (SELECT MAX(" & sFieldID & ") " + _
       "FROM " & sTable & " " + _
       "GROUP BY " & sFieldLists & " HAVING COUNT(*) > 1))) "

    D.cyConnectSqlServer sServer, sDataBase, sID, sPassWord

    Set Rs = D.cyGetRs(sSQL)
    
    Do While Rs.RecordCount > 0
        D.cyExeCute "DELETE " & Right(sSQL, Len(sSQL) - 8)
    
        sSQL = "SELECT * " + _
           "FROM " & sTable & " " + _
           "WHERE (" & sFieldID & " NOT IN (SELECT " & sFieldID & " " + _
           "FROM " & sTable & " a LEFT OUTER JOIN (SELECT " & sFieldLists & " " + _
           "FROM " & sTable & " " + _
           "GROUP BY " & sFieldLists & " HAVING COUNT(*) > 1) t1 ON " & sStr + _
           "WHERE a." & sFieldID & " NOT IN (SELECT MAX(" & sFieldID & ") " + _
           "FROM " & sTable & " " + _
           "GROUP BY " & sFieldLists & " HAVING COUNT(*) > 1))) "
    
        Set Rs = D.cyGetRs(sSQL)
    
    Loop

End Sub

