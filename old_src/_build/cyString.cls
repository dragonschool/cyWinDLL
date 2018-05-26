VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cyStringEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetTickCount Lib "kernel32" () As Long

'-----------------------------------------------------------------------------
'范式检测关键字
'-----------------------------------------------------------------------------
Public KeywordFoundWord As String     '出现的关键词
Public KeywordFoundPos As Long            '出现的位置
Public KeywordFoundLen As Long            '可疑字符的长度
'-----------------------------------------------------------------------------
'范式检测关键字
'-----------------------------------------------------------------------------



'-----------------------------------------------------------------------------
'32位MD5
'-----------------------------------------------------------------------------
Private Const BITS_TO_A_BYTE_NET = 8
Private Const BYTES_TO_A_WORD_NET = 4
Private Const BITS_TO_A_WORD_NET = 32
Private m_lOnBits_NET(30)
Private m_l2Power_NET(30)
'-----------------------------------------------------------------------------
'32位MD5
'-----------------------------------------------------------------------------

'设定左/右对齐的枚举类型
Enum cyAlignAction
    LeftAlign = 0
    RightAlign = 1
End Enum


'设定需要转换的枚举类型
Enum cyChangeAction
    GBToGBK = 0
    GBToBig5 = 1
    GBKToGB = 2
    GBKToBig5 = 3
    Big5ToGB = 4
    Big5ToGBK = 5
End Enum


'MD5###################################
Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32

Private m_lOnBits(30)
Private m_l2Power(30)
'MD5###################################


'生成UID###############################
Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long
'生成UID###############################

Private Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long


Enum cyFilterAction
    AllowAllChar = 0
    AllowNumberOnly = 1
    AllowUcaseEngCharOnly = 2
    AllowLcaseEngCharOnly = 3
    AllowChineseCharOnly = 4
    AllowEngCharOnly = 9
    denyAllChar = 11
    denyNumber = 5
    DenyUcaseEngChar = 6
    DenyLcaseEngChar = 7
    denyChineseChar = 8
    denyEngChar = 10
End Enum

'---------------------------------------------------------------------------------------
' 过程名: cyMidEx
' 生成日期: 2006-3-3 12:09
' 作者: chenyon
' 过程作用:取得字符串中某两段字符之间的内容
'---------------------------------------------------------------------------------------
Public Function cyMidEx(sString As String, sStart As String, sEnd As String, Optional ByVal iStartPos As Long = 1) As String
    Dim i As Long
    Dim j As Long
    '取得前面字符的位置
    i = InStr(iStartPos, sString, sStart) + Len(sStart)
    '如果位置位于字符串的末尾则退出函数
    If i = Len(sStart) Then Exit Function
    '从后段字符串中查找后面的字符
    j = InStr(i, sString, sEnd)
    If j = 0 Then Exit Function
    '根据前后位置取得中间内容
    cyMidEx = Mid(sString, i, j - i)
End Function

'函数名称:cyTrimEx

'作用：取得定长字符串的有用字符串

'##################################################
'整理时间：2004-2-13 18:10:58
'  参数：
'   sString As String       --
'##################################################
'  例：
Public Function cyTrimEx(sString As String) As String
    sString = Trim(sString)
    cyTrimEx = Mid(sString, 1, InStr(1, sString, Chr(0)) - 1)
End Function

'函数名称:cyLenEx

'作用：取得中英文混合字符串的真正长度[中文算两个长度]

'##################################################
'整理时间：2004-2-13 18:12:19
'  参数：
'   sEngAndChiString As String  --要取得字数的字符串
'##################################################
'  例：
Public Function cyLenEx(sEngAndChiString As String) As Long
    cyLenEx = LenB(StrConv(sEngAndChiString, vbFromUnicode))
End Function

'函数名称:cyFixTxtLength

'作用：将一段文本格式化为等长度的段落文字

'##################################################
'整理时间：2004-2-13 18:13:57
'  参数：
'   sSourceTxt As String    --源文字
'   iLengthOfChiChar As Integer --中文字长度
'##################################################
'  例：
Public Function cyFixTxtLength(sSourceTxt As String, iLengthOfChiChar As Integer) As String
Dim iStart As Long
Dim iStop As Long
Dim sTarget As String
Dim iLength As Long
Dim i As Long '此变量记录增加长度前的字符串长度，如果一样证明已到尾了
iStart = 1
iStop = iLengthOfChiChar '设置截取的步距
    
'再进入循环
cyTag:
Do While Len(sSourceTxt) > 0
    iLength = lstrlen(Mid(sSourceTxt, iStart, iStop) + Chr(0))
    If iLength = iLengthOfChiChar * 2 Then '刚才符合长度
        iStop = iStop + 1
    Else
        Do While iLength <= iLengthOfChiChar * 2  '如果长度没到就继续增加字符
            i = iLength
            iStop = iStop + 1
            iLength = lstrlen(Mid(sSourceTxt, iStart, iStop) + Chr(0))
            If iLength = i Then '如果前后无改变证明到尾了
                sTarget = sTarget + Chr(13) + Chr(10) + Mid(sSourceTxt, iStart, Len(sSourceTxt))
                GoTo cyFinish
            End If
        Loop
    End If
    
    '长度足够，开始将分好的行赋值给sTarget
    sTemp = Mid(sSourceTxt, iStart, iStop - 1)
    If InStr(1, sTemp, Chr(13)) > 0 Then
        If Asc(Left(sTemp, 1)) = 13 Then
            sTarget = sTarget + Left(sTemp, InStr(1, sTemp, Chr(13)) - 1)
        Else
            sTarget = sTarget + Chr(13) + Chr(10) + Left(sTemp, InStr(1, sTemp, Chr(13)) - 1)
        
        End If
        
        '如果下一个字符是回车就把开始号加2
        sSourceTxt = Right(sSourceTxt, Len(sSourceTxt) - InStr(1, sTemp, Chr(13)) - 1)
        'iStart = iStart +  '将结束的长度赋给下一次开始
        iStop = iLengthOfChiChar '设置截取的步距
    Else
        sTarget = sTarget + Chr(13) + Chr(10) + Mid(sSourceTxt, iStart, iStop - 1)
        sSourceTxt = Right(sSourceTxt, Len(sSourceTxt) - iStop + 1)
        If Len(sSourceTxt) = 0 Then
            GoTo cyFinish
        ElseIf lstrlen(sSourceTxt + Chr(0)) < iLengthOfChiChar * 2 Then '如果剩下的长度少于0则将结束
            sTarget = sTarget + Chr(13) + Chr(10) + sSourceTxt
            GoTo cyFinish
        End If
        iStop = iLengthOfChiChar '设置截取的步距
    End If
    

Loop



cyFinish:
    cyFixTxtLength = Right(sTarget, Len(sTarget) - 2)
End Function

'---------------------------------------------------------------------------------------
' 过程名: cyFixCharsLength
' 生成日期: 2006-3-6 22:42
' 作者: Dragonschool
' 过程作用:将一串文字格式化为字长的字符串,字数不够用空格代替
' 参数：
'   sChars As String            --源字符串
'   iCharsLength As Integer     --格式化后的长度
'   Align As cyAlignAction      --左/右对齐
'---------------------------------------------------------------------------------------
Public Function cyFixCharsLength(ByVal sChars As String, ByVal iCharsLength As Integer, ByVal Action As cyAlignAction) As String
    Dim i As Long   '得到字符串的混合长度，中文占2个字位，英文占1个字位
    i = lstrlen(sChars + Chr(0))
    
    If i > iCharsLength Then    '如果源串比目标长度长则不变直接输出
        cyFixCharsLength = sChars
        Exit Function
    End If
    
    '否则根据左或右对齐填充空格
    cyFixCharsLength = IIf(Action = LeftAlign, sChars & Space(iCharsLength - i), Space(iCharsLength - i) & sChars)
        
End Function

'---------------------------------------------------------------------------------------
' 过程名: cyFiltKeyAscii
' 生成日期: 2006-3-4 07:33
' 作者: Dragonschool
' 过程作用:字符输入过滤,在文本框中把不希望输入的字符过滤掉(如不允许输入中文，仅允许输入数字等..)
' 参数：
'   iKeyAscii As Long           --Keyascii值
'   Action As cyFilterAction    --过滤的规则[只允许中文/英文...]
'   sOtherAllowList As String   --其它允许输入的字符
'   sOtherDenyList As String    --不允许输入的字符
'   sReplaceList As String      --替换列表,如将A替换为a,用;分隔，如ABC;abc

'---------------------------------------------------------------------------------------
Function cyFiltKeyAscii(ByVal iKeyAscii As Long, Optional ByVal Action As cyFilterAction, Optional ByVal sOtherAllowList As String, Optional ByVal sOtherDenyList As String, Optional ByVal sReplaceList As String)

    If iKeyAscii = 8 Then: cyFiltKeyAscii = iKeyAscii: Exit Function
    Select Case Action
        Case 0
                    '无过滤规则
            GoTo PassCheck
        Case 1      '仅允许数字
            If iKeyAscii > 47 And iKeyAscii < 58 Then
                '符合规则,则直接检测是否包含在拒绝列表内
                GoTo PassCheck
            ElseIf InStr(1, sOtherAllowList, Chr(iKeyAscii)) > 0 Then
                '不符合规则,但是在允许列表内
                GoTo PassCheck
            Else
                 '不符合规则,过滤掉
                cyFiltKeyAscii = 0
                Beep
                Exit Function
            End If
            
        Case 2      '仅允许大写英文
            If iKeyAscii > 64 And iKeyAscii < 91 Then
                '符合规则,则直接检测是否包含在拒绝列表内
                GoTo PassCheck
            ElseIf InStr(1, sOtherAllowList, Chr(iKeyAscii)) > 0 Then
                '不符合规则,但是在允许列表内
                GoTo PassCheck
            Else
                 '不符合规则,过滤掉
                cyFiltKeyAscii = 0
                Beep
                Exit Function
            End If
            
        Case 3      '仅允许小写英文
            If iKeyAscii > 96 And iKeyAscii < 123 Then
                '符合规则,则直接检测是否包含在拒绝列表内
                GoTo PassCheck
            ElseIf InStr(1, sOtherAllowList, Chr(iKeyAscii)) > 0 Then
                '不符合规则,但是在允许列表内
                GoTo PassCheck
            Else
                 '不符合规则,过滤掉
                cyFiltKeyAscii = 0
                Beep
                Exit Function
            End If
            
        Case 4      '仅允许中文字符
            If iKeyAscii < 0 Then
                '符合规则,则直接检测是否包含在拒绝列表内
                GoTo PassCheck
            ElseIf InStr(1, sOtherAllowList, Chr(iKeyAscii)) > 0 Then
                '不符合规则,但是在允许列表内
                GoTo PassCheck
            Else
                 '不符合规则,过滤掉
                cyFiltKeyAscii = 0
                Beep
                Exit Function
            End If
        
        Case 5      '过滤掉数字
            If Not (iKeyAscii > 47 And iKeyAscii < 58) Then
                '符合规则,则直接检测是否包含在拒绝列表内
                GoTo PassCheck
            ElseIf InStr(1, sOtherAllowList, Chr(iKeyAscii)) > 0 Then
                '不符合规则,但是在允许列表内
                GoTo PassCheck
            Else
                 '不符合规则,过滤掉
                cyFiltKeyAscii = 0
                Beep
                Exit Function
            End If
        Case 6      '过滤掉大写英文
            If Not (iKeyAscii > 64 And iKeyAscii < 91) Then
                '符合规则,则直接检测是否包含在拒绝列表内
                GoTo PassCheck
            ElseIf InStr(1, sOtherAllowList, Chr(iKeyAscii)) > 0 Then
                '不符合规则,但是在允许列表内
                GoTo PassCheck
            Else
                 '不符合规则,过滤掉
                cyFiltKeyAscii = 0
                Beep
                Exit Function
            End If
        Case 7      '过滤掉小写英文
            If Not (iKeyAscii > 96 And iKeyAscii < 123) Then
                '符合规则,则直接检测是否包含在拒绝列表内
                GoTo PassCheck
            ElseIf InStr(1, sOtherAllowList, Chr(iKeyAscii)) > 0 Then
                '不符合规则,但是在允许列表内
                GoTo PassCheck
            Else
                 '不符合规则,过滤掉
                cyFiltKeyAscii = 0
                Beep
                Exit Function
            End If
        Case 8      '过滤掉中文字符
            If Not (iKeyAscii < 0) Then
                '符合规则,则直接检测是否包含在拒绝列表内
                GoTo PassCheck
            ElseIf InStr(1, sOtherAllowList, Chr(iKeyAscii)) > 0 Then
                '不符合规则,但是在允许列表内
                GoTo PassCheck
            Else
                 '不符合规则,过滤掉
                cyFiltKeyAscii = 0
                Beep
                Exit Function
            End If
        Case 9      '允许大小写英文字符
            If (iKeyAscii > 64 And iKeyAscii < 91) Or (iKeyAscii > 96 And iKeyAscii < 123) Then
                '符合规则,则直接检测是否包含在拒绝列表内
                GoTo PassCheck
            ElseIf InStr(1, sOtherAllowList, Chr(iKeyAscii)) > 0 Then
                '不符合规则,但是在允许列表内
                GoTo PassCheck
            Else
                 '不符合规则,过滤掉
                cyFiltKeyAscii = 0
                Beep
                Exit Function
            End If
        Case 10      '拒绝大小写英文字符
            If Not ((iKeyAscii > 64 And iKeyAscii < 91) Or (iKeyAscii > 96 And iKeyAscii < 123)) Then
                '符合规则,则直接检测是否包含在拒绝列表内
                GoTo PassCheck
            ElseIf InStr(1, sOtherAllowList, Chr(iKeyAscii)) > 0 Then
                '不符合规则,但是在允许列表内
                GoTo PassCheck
            Else
                 '不符合规则,过滤掉
                cyFiltKeyAscii = 0
                Beep
                Exit Function
            End If
        Case 11      '拒绝所有输入
                cyFiltKeyAscii = 0
                Beep
    End Select
PassCheck: '注意:拒绝列表优先于其它所有规则及允许列表
            If InStr(1, sOtherDenyList, Chr(iKeyAscii)) > 0 Then '输入的字符串存在于拒绝输入列表内
                 '过滤掉
                cyFiltKeyAscii = 0
                Beep
            Else '不存在于拒绝列表
                '不过滤,允许通过
                Dim sReplace() As String
                '查找字符是否位于替换列表
                If Len(sReplaceList) > 0 Then
                    sReplace = Split(sReplaceList, ";")
                    Dim iPos As Long
                    iPos = InStr(1, sReplace(0), Chr(iKeyAscii))
                    If iPos > 0 Then
                        cyFiltKeyAscii = Asc(Mid(sReplace(1), iPos, 1))
                        Exit Function
                    End If
                End If
                cyFiltKeyAscii = iKeyAscii
            End If
            Exit Function
    
End Function

'MD5字加密########################
'加密子程序
Function cyStrEncrypt(ByVal sKey As String, ByVal strToEncrypt As String) As String
Const MIN_ASC = 32 ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Dim offset As Long
Dim Str_len As Integer
Dim i As Integer
Dim ch As Integer

'得到了加密的数字
offset = NumericPassword(sKey)

Rnd -1
'对随机数生成器做初始化的动作
Randomize offset

Str_len = Len(strToEncrypt)
For i = 1 To Str_len
    ch = Asc(Mid$(strToEncrypt, i, 1))
    If ch >= MIN_ASC And ch <= MAX_ASC Then
        ch = ch - MIN_ASC
        offset = Int((NUM_ASC + 1) * Rnd)
        ch = ((ch + offset) Mod NUM_ASC)
        ch = ch + MIN_ASC
        cyStrEncrypt = cyStrEncrypt & Chr$(ch)
    End If
Next i

'处理特殊字符
cyStrEncrypt = Me.cyURLEncode(cyStrEncrypt)

End Function

'使用流水时间生成的KEY进行加密再经过URL转换，避免在INTERNET传输失败
Function cyStrEncryptEx(ByVal sString As String) As String
Const MIN_ASC = 32 ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

'记录流水KEY
Dim sKey As String
Dim sStr As String
Dim iOffSet As Long
Dim iStrLen As Integer
Dim i As Integer
Dim iChrAsc As Long

sKey = Format(Right(GetTickCount, 4), "0000")

'sString = Me.cyURLEncode(sKey & sString)
'sString = Right(sString, Len(sString) - 12)

'得到了加密的数字
iOffSet = NumericPassword(sKey)

Rnd -1

'对随机数生成器做初始化的动作
Randomize iOffSet

iStrLen = Len(sString)
For i = 1 To iStrLen
    iChrAsc = Asc(Mid$(sString, i, 1))
    If iChrAsc >= MIN_ASC And iChrAsc <= MAX_ASC Then
        iChrAsc = iChrAsc - MIN_ASC
        iOffSet = Int((NUM_ASC + 1) * Rnd)
        iChrAsc = ((iChrAsc + iOffSet) Mod NUM_ASC)
        iChrAsc = iChrAsc + MIN_ASC
        
        If (iChrAsc >= 65 And iChrAsc <= 90) Or (iChrAsc >= 97 And iChrAsc <= 122) Then
             cyStrEncryptEx = cyStrEncryptEx & Chr$(iChrAsc)
        Else
             '如果不是中、英文，则使用HEX转换其ASC码
             sStr = Hex(iChrAsc)
             '使用IIf是为了确保ASC码<16时返回的是两位长度的字符串，否则解码时会出错
             cyStrEncryptEx = cyStrEncryptEx & "#" & IIf(Len(sStr) = 2, sStr, "0" & sStr)
        End If
        
'        cyStrEncryptEx = cyStrEncryptEx & Chr$(iChrAsc)
    End If
Next i
'
' sStr = Asc(Mid(sURL, i, 1))
'
'处理特殊字符
cyStrEncryptEx = sKey & cyStrEncryptEx

End Function


'解密子程序
Function cyStrDecrypt(ByVal sKey As String, ByVal strToUnEncrypt As String) As String
Dim offset As Long
Dim iLen As Integer
Dim i As Integer
Dim ch As Integer


'解密
Const MIN_ASC = 32 ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

strToUnEncrypt = cyURLDecode(strToUnEncrypt)

offset = NumericPassword(sKey)
Rnd -1
Randomize offset

iLen = Len(strToUnEncrypt)
For i = 1 To iLen
    ch = Asc(Mid$(strToUnEncrypt, i, 1))
    If ch >= MIN_ASC And ch <= MAX_ASC Then
        ch = ch - MIN_ASC
        offset = Int((NUM_ASC + 1) * Rnd)
        ch = ((ch - offset) Mod NUM_ASC)
        If ch < 0 Then ch = ch + NUM_ASC
        ch = ch + MIN_ASC
        cyStrDecrypt = cyStrDecrypt & Chr$(ch)
    End If
Next i
End Function


'解密子程序
Function cyStrDecryptEx(ByVal sString As String) As String
Dim iOffSet As Long
Dim iStrLen As Integer
Dim i As Integer
Dim iChrAsc As Integer

If sString = "" Then Exit Function

'解密
Const MIN_ASC = 32 ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1
'
'sString = Me.cyURLDecode(sString)

'取出KEY
iOffSet = NumericPassword(Left(sString, 4))
Rnd -1
Randomize iOffSet

'返回真正的字符串
Dim sStr As String
Dim sStr1 As String
sStr = Right(sString, Len(sString) - 4)

iStrLen = Len(sString)
For i = 1 To iStrLen

    '如果字符串的当前字符是#则认为是编码的一部分，则继续分析
    If Mid(sStr, i, 1) = "#" Then
'        If Val("&H" & Mid(sStr, i + 1, 2)) > 127 Then
'            '紧跟着#后的两个字符的十六进制如果大于127就认为是中文字符，则要将其后两位的字符加起来再通过CHR函数还原成中文字符
'            sString = sString & Chr(Val("&H" & Mid(sStr, i + 1, 2) & Mid(sStr, i + 4, 2)))
'            i = i + 5
'        Else
            '带#号且不属于中文字的就是数字或其它除英文字符外的字符
            sStr1 = sStr1 & Chr(Val("&H" & Mid(sStr, i + 1, 2)))
            i = i + 2
'        End If
    Else
        '否则就是英文字符
        sStr1 = sStr1 & Mid(sStr, i, 1)
    End If

Next i
    
For i = 1 To Len(sStr1)
    iChrAsc = Asc(Mid$(sStr1, i, 1))
    If iChrAsc >= MIN_ASC And iChrAsc <= MAX_ASC Then
        iChrAsc = iChrAsc - MIN_ASC
        iOffSet = Int((NUM_ASC + 1) * Rnd)
        iChrAsc = ((iChrAsc - iOffSet) Mod NUM_ASC)
        If iChrAsc < 0 Then iChrAsc = iChrAsc + NUM_ASC
        iChrAsc = iChrAsc + MIN_ASC
        cyStrDecryptEx = cyStrDecryptEx & Chr$(iChrAsc)
    End If
Next

End Function

'将你输入的每个字符转换成密码数字
Private Function NumericPassword(ByVal Password As String) As Long
Dim iChr As Long
Dim Shift1 As Long
Dim Shift2 As Long
Dim i As Integer
Dim iLen As Integer

'得到字符串内字符的数目
iLen = Len(Password)

'给每个字符转换成密码数字
For i = 1 To iLen
    iChr = Asc(Mid$(Password, i, 1))
    NumericPassword = Value Xor (iChr * 2 ^ Shift1)
    NumericPassword = Value Xor (iChr * 2 ^ Shift2)
    Shift1 = (Shift1 + 7) Mod 19
    Shift2 = (Shift2 + 13) Mod 23
Next i

End Function

Private Function NumChange(ByVal Num As Double) As String
    On Error Resume Next
    Dim Str1 As String, Str2 As String, Strbak As String
    Dim Numi As Double
    Str1 = "壹贰叁肆伍陆柒捌玖拾零"
    Str2 = "仟佰拾元"
    Numi = Int(Num)
    Do While Numi > 0
        If Numi Mod 10 = 0 Then
            If Len(Str2) < 4 And Left(NumChange, 1) <> "零" And Left(NumChange, 1) <> "元" And Left(NumChange, 1) <> "万" Then
                NumChange = Mid(Str1, 11, 1) & NumChange
            Else
                If Len(Str2) = 4 Then NumChange = Mid(Str2, 4, 1) & NumChange
            End If
        Else
            NumChange = Mid(Str1, Numi Mod 10, 1) & Mid(Str2, Len(Str2), 1) & NumChange
        End If
        Numi = Numi \ 10
        If Len(Str2) = 1 Then
            Str2 = "仟佰拾万"
        Else
            Str2 = Mid(Str2, 1, Len(Str2) - 1)
        End If
    Loop
    If InStr(Trim(str(Num)), ".") > 0 Then
        Numi = Val(Mid(Trim(str(Num)), InStr(Trim(str(Num)), ".") + 1, Len(Trim(str(Num))) - InStr(Trim(str(Num)), ".")))
        If Left(Mid(Trim(str(Num)), InStr(Trim(str(Num)), ".") + 1, Len(Trim(str(Num))) - InStr(Trim(str(Num)), ".")), 1) = "0" Then
            Str2 = "分"
        Else
            Str2 = Left("角分", Len(Trim(str(Numi))))
        
        End If
        Do While Numi > 0
            If Numi Mod 10 <> 0 Then
                Strbak = Mid(Str1, Numi Mod 10, 1) & Mid(Str2, Len(Str2), 1) & Strbak
            End If
            Numi = Numi \ 10
            Str2 = Mid(Str2, 1, Len(Str2) - 1)
        Loop
    End If
    NumChange = NumChange & Strbak
End Function


Public Function cyNum2ChiNum(ByVal iNum As Double) As String
    On Error Resume Next
    If InStr(1, CStr(iNum), ".") = 0 Then
        Num = CLng(iNum)
    Else
        Num = iNum
    End If
    
    Dim i As Long
    i = CLng(iNum)
    
    If Len(Trim(CStr(i))) > 8 Then  '超亿元
        Dim sTemp As String
        Dim sTemp1 As String
        sTemp = Trim(CStr(i))
        sTemp = NumChange(CInt(Left(sTemp, Len(sTemp) - 8)))
        sTemp = Left(sTemp, Len(sTemp) - 1) + "亿"
        
        sTemp1 = Right(Left(i, Len(CStr(i)) - 4), 4)
        If Left(sTemp1, 1) = "0" Or Left(sTemp1, 2) = "00" Or Left(sTemp1, 3) = "000" Then sTemp = sTemp + "零"
        
        cyNum2ChiNum = sTemp + NumChange(Right(iNum, Len(CStr(iNum)) - (Len(CStr(i)) - 8)))
    Else                            '亿元以内
        cyNum2ChiNum = NumChange(iNum)
    End If
    
    '如果是整数就在最后加整
    Select Case TypeName(Num)
        Case "Long"
            cyNum2ChiNum = cyNum2ChiNum + "整"
        Case "Integer"
            cyNum2ChiNum = cyNum2ChiNum + "整"
        Case "Double"
            cyNum2ChiNum = cyNum2ChiNum
    End Select
End Function









Private Function LShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

Private Function RShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

Private Function RotateLeft(lValue, iShiftBits)
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(lX, lY)
    Dim lX4
    Dim lY4
    Dim lX8
    Dim lY8
    Dim lResult

    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000

    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If

    AddUnsigned = lResult
End Function

Private Function md5_F(X, Y, z)
    md5_F = (X And Y) Or ((Not X) And z)
End Function

Private Function md5_G(X, Y, z)
    md5_G = (X And z) Or (Y And (Not z))
End Function

Private Function md5_H(X, Y, z)
    md5_H = (X Xor Y Xor z)
End Function

Private Function md5_I(X, Y, z)
    md5_I = (Y Xor (X Or (Not z)))
End Function

Private Sub md5_FF(a, b, c, D, X, S, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_F(b, c, D), X), ac))
    a = RotateLeft(a, S)
    a = AddUnsigned(a, b)
End Sub

Private Sub md5_GG(a, b, c, D, X, S, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, D), X), ac))
    a = RotateLeft(a, S)
    a = AddUnsigned(a, b)
End Sub

Private Sub md5_HH(a, b, c, D, X, S, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, D), X), ac))
    a = RotateLeft(a, S)
    a = AddUnsigned(a, b)
End Sub

Private Sub md5_II(a, b, c, D, X, S, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_I(b, c, D), X), ac))
    a = RotateLeft(a, S)
    a = AddUnsigned(a, b)
End Sub

Private Function ConvertToWordArray(sMessage)
    Dim lMessageLength
    Dim lNumberOfWords
    Dim lWordArray()
    Dim lBytePosition
    Dim lByteCount
    Dim lWordCount

    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448

    lMessageLength = Len(sMessage)

    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)

    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)

    ConvertToWordArray = lWordArray
End Function

Private Function WordToHex(lValue)
    Dim lByte
    Dim lCount

    For lCount = 0 To 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
        WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
    Next
End Function

'函数名称:cyMD5

'作用：MD5加密

'##################################################
'整理时间：2004-2-13 18:20:35
'  参数：
'   sMessage                --要加密的字符串
'##################################################
'  例：
Public Function cyMD5(sMessage) As String
    m_lOnBits(0) = CLng(1)
    m_lOnBits(1) = CLng(3)
    m_lOnBits(2) = CLng(7)
    m_lOnBits(3) = CLng(15)
    m_lOnBits(4) = CLng(31)
    m_lOnBits(5) = CLng(63)
    m_lOnBits(6) = CLng(127)
    m_lOnBits(7) = CLng(255)
    m_lOnBits(8) = CLng(511)
    m_lOnBits(9) = CLng(1023)
    m_lOnBits(10) = CLng(2047)
    m_lOnBits(11) = CLng(4095)
    m_lOnBits(12) = CLng(8191)
    m_lOnBits(13) = CLng(16383)
    m_lOnBits(14) = CLng(32767)
    m_lOnBits(15) = CLng(65535)
    m_lOnBits(16) = CLng(131071)
    m_lOnBits(17) = CLng(262143)
    m_lOnBits(18) = CLng(524287)
    m_lOnBits(19) = CLng(1048575)
    m_lOnBits(20) = CLng(2097151)
    m_lOnBits(21) = CLng(4194303)
    m_lOnBits(22) = CLng(8388607)
    m_lOnBits(23) = CLng(16777215)
    m_lOnBits(24) = CLng(33554431)
    m_lOnBits(25) = CLng(67108863)
    m_lOnBits(26) = CLng(134217727)
    m_lOnBits(27) = CLng(268435455)
    m_lOnBits(28) = CLng(536870911)
    m_lOnBits(29) = CLng(1073741823)
    m_lOnBits(30) = CLng(2147483647)

    m_l2Power(0) = CLng(1)
    m_l2Power(1) = CLng(2)
    m_l2Power(2) = CLng(4)
    m_l2Power(3) = CLng(8)
    m_l2Power(4) = CLng(16)
    m_l2Power(5) = CLng(32)
    m_l2Power(6) = CLng(64)
    m_l2Power(7) = CLng(128)
    m_l2Power(8) = CLng(256)
    m_l2Power(9) = CLng(512)
    m_l2Power(10) = CLng(1024)
    m_l2Power(11) = CLng(2048)
    m_l2Power(12) = CLng(4096)
    m_l2Power(13) = CLng(8192)
    m_l2Power(14) = CLng(16384)
    m_l2Power(15) = CLng(32768)
    m_l2Power(16) = CLng(65536)
    m_l2Power(17) = CLng(131072)
    m_l2Power(18) = CLng(262144)
    m_l2Power(19) = CLng(524288)
    m_l2Power(20) = CLng(1048576)
    m_l2Power(21) = CLng(2097152)
    m_l2Power(22) = CLng(4194304)
    m_l2Power(23) = CLng(8388608)
    m_l2Power(24) = CLng(16777216)
    m_l2Power(25) = CLng(33554432)
    m_l2Power(26) = CLng(67108864)
    m_l2Power(27) = CLng(134217728)
    m_l2Power(28) = CLng(268435456)
    m_l2Power(29) = CLng(536870912)
    m_l2Power(30) = CLng(1073741824)


    Dim X
    Dim k
    Dim AA
    Dim BB
    Dim CC
    Dim DD
    Dim a
    Dim b
    Dim c
    Dim D

    Const S11 = 7
    Const S12 = 12
    Const S13 = 17
    Const S14 = 22
    Const S21 = 5
    Const S22 = 9
    Const S23 = 14
    Const S24 = 20
    Const S31 = 4
    Const S32 = 11
    Const S33 = 16
    Const S34 = 23
    Const S41 = 6
    Const S42 = 10
    Const S43 = 15
    Const S44 = 21

    X = ConvertToWordArray(sMessage)

    a = &H67452301
    b = &HEFCDAB89
    c = &H98BADCFE
    D = &H10325476

    For k = 0 To UBound(X) Step 16
        AA = a
        BB = b
        CC = c
        DD = D

        md5_FF a, b, c, D, X(k + 0), S11, &HD76AA478
        md5_FF D, a, b, c, X(k + 1), S12, &HE8C7B756
        md5_FF c, D, a, b, X(k + 2), S13, &H242070DB
        md5_FF b, c, D, a, X(k + 3), S14, &HC1BDCEEE
        md5_FF a, b, c, D, X(k + 4), S11, &HF57C0FAF
        md5_FF D, a, b, c, X(k + 5), S12, &H4787C62A
        md5_FF c, D, a, b, X(k + 6), S13, &HA8304613
        md5_FF b, c, D, a, X(k + 7), S14, &HFD469501
        md5_FF a, b, c, D, X(k + 8), S11, &H698098D8
        md5_FF D, a, b, c, X(k + 9), S12, &H8B44F7AF
        md5_FF c, D, a, b, X(k + 10), S13, &HFFFF5BB1
        md5_FF b, c, D, a, X(k + 11), S14, &H895CD7BE
        md5_FF a, b, c, D, X(k + 12), S11, &H6B901122
        md5_FF D, a, b, c, X(k + 13), S12, &HFD987193
        md5_FF c, D, a, b, X(k + 14), S13, &HA679438E
        md5_FF b, c, D, a, X(k + 15), S14, &H49B40821

        md5_GG a, b, c, D, X(k + 1), S21, &HF61E2562
        md5_GG D, a, b, c, X(k + 6), S22, &HC040B340
        md5_GG c, D, a, b, X(k + 11), S23, &H265E5A51
        md5_GG b, c, D, a, X(k + 0), S24, &HE9B6C7AA
        md5_GG a, b, c, D, X(k + 5), S21, &HD62F105D
        md5_GG D, a, b, c, X(k + 10), S22, &H2441453
        md5_GG c, D, a, b, X(k + 15), S23, &HD8A1E681
        md5_GG b, c, D, a, X(k + 4), S24, &HE7D3FBC8
        md5_GG a, b, c, D, X(k + 9), S21, &H21E1CDE6
        md5_GG D, a, b, c, X(k + 14), S22, &HC33707D6
        md5_GG c, D, a, b, X(k + 3), S23, &HF4D50D87
        md5_GG b, c, D, a, X(k + 8), S24, &H455A14ED
        md5_GG a, b, c, D, X(k + 13), S21, &HA9E3E905
        md5_GG D, a, b, c, X(k + 2), S22, &HFCEFA3F8
        md5_GG c, D, a, b, X(k + 7), S23, &H676F02D9
        md5_GG b, c, D, a, X(k + 12), S24, &H8D2A4C8A

        md5_HH a, b, c, D, X(k + 5), S31, &HFFFA3942
        md5_HH D, a, b, c, X(k + 8), S32, &H8771F681
        md5_HH c, D, a, b, X(k + 11), S33, &H6D9D6122
        md5_HH b, c, D, a, X(k + 14), S34, &HFDE5380C
        md5_HH a, b, c, D, X(k + 1), S31, &HA4BEEA44
        md5_HH D, a, b, c, X(k + 4), S32, &H4BDECFA9
        md5_HH c, D, a, b, X(k + 7), S33, &HF6BB4B60
        md5_HH b, c, D, a, X(k + 10), S34, &HBEBFBC70
        md5_HH a, b, c, D, X(k + 13), S31, &H289B7EC6
        md5_HH D, a, b, c, X(k + 0), S32, &HEAA127FA
        md5_HH c, D, a, b, X(k + 3), S33, &HD4EF3085
        md5_HH b, c, D, a, X(k + 6), S34, &H4881D05
        md5_HH a, b, c, D, X(k + 9), S31, &HD9D4D039
        md5_HH D, a, b, c, X(k + 12), S32, &HE6DB99E5
        md5_HH c, D, a, b, X(k + 15), S33, &H1FA27CF8
        md5_HH b, c, D, a, X(k + 2), S34, &HC4AC5665

        md5_II a, b, c, D, X(k + 0), S41, &HF4292244
        md5_II D, a, b, c, X(k + 7), S42, &H432AFF97
        md5_II c, D, a, b, X(k + 14), S43, &HAB9423A7
        md5_II b, c, D, a, X(k + 5), S44, &HFC93A039
        md5_II a, b, c, D, X(k + 12), S41, &H655B59C3
        md5_II D, a, b, c, X(k + 3), S42, &H8F0CCC92
        md5_II c, D, a, b, X(k + 10), S43, &HFFEFF47D
        md5_II b, c, D, a, X(k + 1), S44, &H85845DD1
        md5_II a, b, c, D, X(k + 8), S41, &H6FA87E4F
        md5_II D, a, b, c, X(k + 15), S42, &HFE2CE6E0
        md5_II c, D, a, b, X(k + 6), S43, &HA3014314
        md5_II b, c, D, a, X(k + 13), S44, &H4E0811A1
        md5_II a, b, c, D, X(k + 4), S41, &HF7537E82
        md5_II D, a, b, c, X(k + 11), S42, &HBD3AF235
        md5_II c, D, a, b, X(k + 2), S43, &H2AD7D2BB
        md5_II b, c, D, a, X(k + 9), S44, &HEB86D391

        a = AddUnsigned(a, AA)
        b = AddUnsigned(b, BB)
        c = AddUnsigned(c, CC)
        D = AddUnsigned(D, DD)
    Next

    'MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
    cyMD5 = LCase(WordToHex(b) & WordToHex(c)) 'I crop this to fit 16byte database password :D
End Function
'统计结果：
'    注释共：2行
'MD5字加密########################


'连接两个数组的内容(String)
Function cyUniteArrayString(OldArray() As String, ArrayToAdd() As String)
    Dim TempArray() As String
    Dim iUnite As Long
    TempArray = OldArray
    '将附加文件添加到原来的资源列表中
    For iUnite = LBound(ArrayToAdd) To UBound(ArrayToAdd)
        ReDim Preserve TempArray(LBound(TempArray) To UBound(TempArray) + 1)
        TempArray(UBound(TempArray)) = ArrayToAdd(iUnite)
    Next
    cyUniteArray = TempArray
End Function

'连接两个数组的内容(Long)
Function cyUniteArrayLong(OldArray() As Long, ArrayToAdd() As Long)
    Dim TempArray() As Long
    Dim iUnite As Long
    TempArray = OldArray
    '将附加文件添加到原来的资源列表中
    For iUnite = LBound(ArrayToAdd) To UBound(ArrayToAdd)
        ReDim Preserve TempArray(LBound(TempArray) To UBound(TempArray) + 1)
        TempArray(UBound(TempArray)) = ArrayToAdd(iUnite)
    Next
    cyUniteArray = TempArray
End Function

'连接两个数组的内容(Integer)
Function cyUniteArrayInteger(OldArray() As Integer, ArrayToAdd() As Integer)
    Dim TempArray() As Integer
    Dim iUnite As Long
    TempArray = OldArray
    '将附加文件添加到原来的资源列表中
    For iUnite = LBound(ArrayToAdd) To UBound(ArrayToAdd)
        ReDim Preserve TempArray(LBound(TempArray) To UBound(TempArray) + 1)
        TempArray(UBound(TempArray)) = ArrayToAdd(iUnite)
    Next
    cyUniteArray = TempArray
End Function


'连接两个数组的内容(byte)
Function cyUniteArrayByte(OldArray() As Byte, ArrayToAdd() As Byte)
    Dim TempArray() As Byte
    Dim iUnite As Long
    TempArray = OldArray
    '将附加文件添加到原来的资源列表中
    For iUnite = LBound(ArrayToAdd) To UBound(ArrayToAdd)
        ReDim Preserve TempArray(LBound(TempArray) To UBound(TempArray) + 1)
        TempArray(UBound(TempArray)) = ArrayToAdd(iUnite)
    Next
    cyUniteArray = TempArray
End Function


'函数名称:cyGetUID

'作用：生成相对唯一的ID，类似CLSID,很难重复

'##################################################
'整理时间：2004-4-13 上午 08:43:28
'  参数：
'##################################################
'  例：    Debug.Print cyGetUID

'{D4CF34CB-008B-4A62-809C-3CED8B1C7F45}

Function cyGetUID() As String
    Dim uGUID As GUID
    Dim sGUID As String
    Dim bGUID() As Byte
    Dim iReturn As Long
    bGUID = String(40, 0)
    CoCreateGuid uGUID '把结构转换为一个可显示的字符串
    iReturn = StringFromGUID2(uGUID, VarPtr(bGUID(0)), 40)
    cyGetUID = Left$(bGUID, iReturn - 1)
End Function

'将字符串格式化为四位十六进制数字符串
'Debug.Print sS.cyFormatString2HexString("1一")
'00314E00
Function cyFormatString2HexString(sMsg As String) As String
Dim i, iLen As Integer
Dim sTemp As String
iLen = Len(sMsg)
For i = 1 To iLen
    '格式化为四位十六进制数字符串
    sTemp = Hex(AscW(Mid(sMsg, i, 1)))
    sTemp = IIf(Len(sTemp) = 2, "00" & sTemp, sTemp)
    cyFormatString2HexString = cyFormatString2HexString & sTemp
Next
End Function


'将Unicode转换为UTF-8
Function cyStringToUTF8(ByVal sString As String) As String
    Dim i As Long
    Dim b1 As Byte
    Dim b2 As Byte
    Dim UTF16 As Long
    If Len(sString) = 0 Then Exit Function    '空字符时退出
    
     '对每个字符进行处理
    For i = 1 To Len(sString)
       ' Work out the UTF16 value of the Unicode character
       CopyMemory b1, ByVal StrPtr(sString) + ((i - 1) * 2), 1
       CopyMemory b2, ByVal StrPtr(sString) + ((i - 1) * 2) + 1, 1
       
       ' Combine the 2 bytes into the Unicode UTF-16
       UTF16 = b2 * 256 + b1
       ' Convert UTF-16 to 2 or 3 bytes of UTF-8
       cyStringToUTF8 = cyStringToUTF8 & ToUTF8(UTF16)
    Next
End Function

Private Function ToUTF8(ByVal UTF16 As Long) As String
    '对特殊字符进行处理
   If UTF16 = 37 Then
        ToUTF8 = "%EF%BC%85"
        Exit Function
   End If
   If UTF16 < &H80 Then
      ToUTF8 = "%" & Hex(UTF16) ' Use number as is
   ElseIf UTF16 < &H800 Then
      ToUTF8 = "%" & Hex(&H80 + (UTF16 And &H3F)) ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      ToUTF8 = "%" & Trim(Hex(&HC0 + (UTF16 And &H1F))) & ToUTF8  ' Use 5 remaining bits
   Else
      ToUTF8 = "%" & Hex(&H80 + (UTF16 And &H3F))  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      ToUTF8 = "%" & Hex(&H80 + (UTF16 And &H3F)) & ToUTF8 ' Use next 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits again
      ToUTF8 = "%" & Hex(&HE0 + (UTF16 And &HF)) & ToUTF8 ' Use 4 remaining bits
   End If
End Function

Public Function cyURLEncode(ByVal sURL As String) As String
    Dim i As Long
    Dim sStr As String
    '逐个字符进行转换
    For i = 1 To Len(sURL)
        '如果ASC码是负数则认为是中文字符
        If Asc(Mid(sURL, i, 1)) < 0 Then
           sStr = Hex(Asc(Mid(sURL, i, 1)))
           '则将它的HEX码拆分，如HEX=D6D0则转换成%D6%D0
           cyURLEncode = cyURLEncode & "%" & Left(sStr, 2) & "%" & Right(sStr, 2)
        Else
            '否则就检测是否大小英文字母，如果是则直接显示而不编码
            sStr = Asc(Mid(sURL, i, 1))
            If (sStr >= 65 And sStr <= 90) Or (sStr >= 97 And sStr <= 122) Then
                cyURLEncode = cyURLEncode & Mid(sURL, i, 1)
            Else
                '如果不是中、英文，则使用HEX转换其ASC码
                sStr = Hex(Asc(Mid(sURL, i, 1)))
                '使用IIf是为了确保ASC码<16时返回的是两位长度的字符串，否则解码时会出错
                cyURLEncode = cyURLEncode & "%" & IIf(Len(sStr) = 2, sStr, "0" & sStr)
           End If
        End If
    Next
End Function

Public Function cyURLDecode(ByVal sURL As String) As String
    Dim i As Long
    '如果没有%表示没有字符经过编码，则退出
    If InStr(sURL, "%") = 0 Then cyURLDecode = sURL: Exit Function
    For i = 1 To Len(sURL)
        '如果字符串的当前字符是%则认为是编码的一部分，则继续分析
        If Mid(sURL, i, 1) = "%" Then
            If Val("&H" & Mid(sURL, i + 1, 2)) > 127 Then
                '紧跟着%后的两个字符的十六进制如果大于127就认为是中文字符，则要将其后两位的字符加起来再通过CHR函数还原成中文字符
                cyURLDecode = cyURLDecode & Chr(Val("&H" & Mid(sURL, i + 1, 2) & Mid(sURL, i + 4, 2)))
                i = i + 5
            Else
                '带%号且不属于中文字的就是数字或其它除英文字符外的字符
                cyURLDecode = cyURLDecode & Chr(Val("&H" & Mid(sURL, i + 1, 2)))
                i = i + 2
            End If
        Else
            '否则就是英文字符
            cyURLDecode = cyURLDecode & Mid(sURL, i, 1)
        End If
    Next
End Function


Function cyStrConvEx(ByVal sText As String, ByVal ChangeAction As cyChangeAction) As String

Dim iTextLen  As Long

'得到源字符串的真实长度，中文字符占2位
iTextLen = lstrlen(sText)

'为返回值设置长度
cyStrConvEx = Space(iTextLen)

'根据不同的要求进行转换

Select Case ChangeAction
    Case GBToGBK
        LCMapString &H804, &H4000000, sText, iTextLen, cyStrConvEx, iTextLen
        
    Case GBToBig5
        cyStrConvEx = StrConv(StrConv(sText, vbFromUnicode, &H404), vbUnicode, &H804)
        iTextLen = lstrlen(sText)
        cyStrConvEx = Space(iTextLen)
        LCMapString &H804, &H4000000, sText, iTextLen, cyStrConvEx, iTextLen
        cyStrConvEx = StrConv(StrConv(cyStrConvEx, vbFromUnicode, &H404), vbUnicode, &H804)
        
    Case GBKToGB
        LCMapString &H804, &H2000000, sText, iTextLen, cyStrConvEx, iTextLen
        
    Case GBKToBig5
        LCMapString &H804, &H2000000, sText, iTextLen, cyStrConvEx, iTextLen
        cyStrConvEx = StrConv(StrConv(sText, vbFromUnicode, &H404), vbUnicode, &H804)
        
    Case Big5ToGB
         sText = StrConv(StrConv(sText, vbFromUnicode, &H804), vbUnicode, &H404)
         iTextLen = lstrlen(sText)
         cyStrConvEx = Space(iTextLen)
         LCMapString &H804, &H2000000, sText, iTextLen, cyStrConvEx, iTextLen
         
    Case Big5ToGBK
         cyStrConvEx = StrConv(StrConv(sText, vbFromUnicode, &H804), vbUnicode, &H404)
         
End Select

End Function




'-----------------------------------------------------------------------------
'32位MD5
'-----------------------------------------------------------------------------
Private Function LShiftEx(lValue, iShiftBits)
  If iShiftBits = 0 Then
    LShiftEx = lValue
    Exit Function
  ElseIf iShiftBits = 31 Then
    If lValue And 1 Then
      LShiftEx = &H80000000
    Else
      LShiftEx = 0
    End If
    Exit Function
  ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
    Err.Raise 6
  End If

  If (lValue And m_l2Power_NET(31 - iShiftBits)) Then
    LShiftEx = ((lValue And m_lOnBits_NET(31 - (iShiftBits + 1))) * m_l2Power_NET(iShiftBits)) Or &H80000000
  Else
    LShiftEx = ((lValue And m_lOnBits_NET(31 - iShiftBits)) * m_l2Power_NET(iShiftBits))
  End If
End Function

Private Function RShiftEx(lValue, iShiftBits)
  If iShiftBits = 0 Then
    RShiftEx = lValue
    Exit Function
  ElseIf iShiftBits = 31 Then
    If lValue And &H80000000 Then
      RShiftEx = 1
    Else
      RShiftEx = 0
    End If
    Exit Function
  ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
    Err.Raise 6
  End If
  
  RShiftEx = (lValue And &H7FFFFFFE) \ m_l2Power_NET(iShiftBits)

  If (lValue And &H80000000) Then
    RShiftEx = (RShiftEx Or (&H40000000 \ m_l2Power_NET(iShiftBits - 1)))
  End If
End Function

Private Function RotateLeftEx(lValue, iShiftBits)
  RotateLeftEx = LShiftEx(lValue, iShiftBits) Or RShiftEx(lValue, (32 - iShiftBits))
End Function

Private Function F(X, Y, z)
  F = (X And Y) Or ((Not X) And z)
End Function

Private Function G(X, Y, z)
  G = (X And z) Or (Y And (Not z))
End Function

Private Function H(X, Y, z)
  H = (X Xor Y Xor z)
End Function

Private Function i(X, Y, z)
  i = (Y Xor (X Or (Not z)))
End Function

Private Sub FF(a, b, c, D, X, S, ac)
  a = AddUnsigned(a, AddUnsigned(AddUnsigned(F(b, c, D), X), ac))
  a = RotateLeftEx(a, S)
  a = AddUnsigned(a, b)
End Sub

Private Sub GG(a, b, c, D, X, S, ac)
  a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, D), X), ac))
  a = RotateLeftEx(a, S)
  a = AddUnsigned(a, b)
End Sub

Private Sub HH(a, b, c, D, X, S, ac)
  a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, D), X), ac))
  a = RotateLeftEx(a, S)
  a = AddUnsigned(a, b)
End Sub

Private Sub II(a, b, c, D, X, S, ac)
  a = AddUnsigned(a, AddUnsigned(AddUnsigned(i(b, c, D), X), ac))
  a = RotateLeftEx(a, S)
  a = AddUnsigned(a, b)
End Sub

Private Function ConvertToWordArrayEx(sMessage)
  Dim lMessageLength
  Dim lNumberOfWords
  Dim lWordArray()
  Dim lBytePosition
  Dim lByteCount
  Dim lWordCount
  
  Const MODULUS_BITS = 512
  Const CONGRUENT_BITS = 448
  
  lMessageLength = Len(sMessage)
  
  m_lOnBits_NET(0) = CLng(1)
m_lOnBits_NET(1) = CLng(3)
m_lOnBits_NET(2) = CLng(7)
m_lOnBits_NET(3) = CLng(15)
m_lOnBits_NET(4) = CLng(31)
m_lOnBits_NET(5) = CLng(63)
m_lOnBits_NET(6) = CLng(127)
m_lOnBits_NET(7) = CLng(255)
m_lOnBits_NET(8) = CLng(511)
m_lOnBits_NET(9) = CLng(1023)
m_lOnBits_NET(10) = CLng(2047)
m_lOnBits_NET(11) = CLng(4095)
m_lOnBits_NET(12) = CLng(8191)
m_lOnBits_NET(13) = CLng(16383)
m_lOnBits_NET(14) = CLng(32767)
m_lOnBits_NET(15) = CLng(65535)
m_lOnBits_NET(16) = CLng(131071)
m_lOnBits_NET(17) = CLng(262143)
m_lOnBits_NET(18) = CLng(524287)
m_lOnBits_NET(19) = CLng(1048575)
m_lOnBits_NET(20) = CLng(2097151)
m_lOnBits_NET(21) = CLng(4194303)
m_lOnBits_NET(22) = CLng(8388607)
m_lOnBits_NET(23) = CLng(16777215)
m_lOnBits_NET(24) = CLng(33554431)
m_lOnBits_NET(25) = CLng(67108863)
m_lOnBits_NET(26) = CLng(134217727)
m_lOnBits_NET(27) = CLng(268435455)
m_lOnBits_NET(28) = CLng(536870911)
m_lOnBits_NET(29) = CLng(1073741823)
m_lOnBits_NET(30) = CLng(2147483647)

m_l2Power_NET(0) = CLng(1)
m_l2Power_NET(1) = CLng(2)
m_l2Power_NET(2) = CLng(4)
m_l2Power_NET(3) = CLng(8)
m_l2Power_NET(4) = CLng(16)
m_l2Power_NET(5) = CLng(32)
m_l2Power_NET(6) = CLng(64)
m_l2Power_NET(7) = CLng(128)
m_l2Power_NET(8) = CLng(256)
m_l2Power_NET(9) = CLng(512)
m_l2Power_NET(10) = CLng(1024)
m_l2Power_NET(11) = CLng(2048)
m_l2Power_NET(12) = CLng(4096)
m_l2Power_NET(13) = CLng(8192)
m_l2Power_NET(14) = CLng(16384)
m_l2Power_NET(15) = CLng(32768)
m_l2Power_NET(16) = CLng(65536)
m_l2Power_NET(17) = CLng(131072)
m_l2Power_NET(18) = CLng(262144)
m_l2Power_NET(19) = CLng(524288)
m_l2Power_NET(20) = CLng(1048576)
m_l2Power_NET(21) = CLng(2097152)
m_l2Power_NET(22) = CLng(4194304)
m_l2Power_NET(23) = CLng(8388608)
m_l2Power_NET(24) = CLng(16777216)
m_l2Power_NET(25) = CLng(33554432)
m_l2Power_NET(26) = CLng(67108864)
m_l2Power_NET(27) = CLng(134217728)
m_l2Power_NET(28) = CLng(268435456)
m_l2Power_NET(29) = CLng(536870912)
m_l2Power_NET(30) = CLng(1073741824)

  
  
  lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE_NET)) \ (MODULUS_BITS \ BITS_TO_A_BYTE_NET)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD_NET)
  ReDim lWordArray(lNumberOfWords - 1)
  
  lBytePosition = 0
  lByteCount = 0
  Do Until lByteCount >= lMessageLength
    lWordCount = lByteCount \ BYTES_TO_A_WORD_NET
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD_NET) * BITS_TO_A_BYTE_NET
    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShiftEx(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
    lByteCount = lByteCount + 1
  Loop

  lWordCount = lByteCount \ BYTES_TO_A_WORD_NET
  lBytePosition = (lByteCount Mod BYTES_TO_A_WORD_NET) * BITS_TO_A_BYTE_NET

  lWordArray(lWordCount) = lWordArray(lWordCount) Or LShiftEx(&H80, lBytePosition)

  lWordArray(lNumberOfWords - 2) = LShiftEx(lMessageLength, 3)
  lWordArray(lNumberOfWords - 1) = RShiftEx(lMessageLength, 29)
  
  ConvertToWordArrayEx = lWordArray
End Function

Private Function WordToHexEx(lValue)
  Dim lByte
  Dim lCount
  
  For lCount = 0 To 3
    lByte = RShiftEx(lValue, lCount * BITS_TO_A_BYTE_NET) And m_lOnBits_NET(BITS_TO_A_BYTE_NET - 1)
    WordToHexEx = WordToHexEx & Right("0" & Hex(lByte), 2)
  Next
End Function

Public Function cyMD5_NET(sMessage)
  Dim X
  Dim k
  Dim AA
  Dim BB
  Dim CC
  Dim DD
  Dim a
  Dim b
  Dim c
  Dim D
  
  Const S11 = 7
  Const S12 = 12
  Const S13 = 17
  Const S14 = 22
  Const S21 = 5
  Const S22 = 9
  Const S23 = 14
  Const S24 = 20
  Const S31 = 4
  Const S32 = 11
  Const S33 = 16
  Const S34 = 23
  Const S41 = 6
  Const S42 = 10
  Const S43 = 15
  Const S44 = 21

  X = ConvertToWordArrayEx(sMessage)
  
  a = &H67452301
  b = &HEFCDAB89
  c = &H98BADCFE
  D = &H10325476

  For k = 0 To UBound(X) Step 16
    AA = a
    BB = b
    CC = c
    DD = D

    FF a, b, c, D, X(k + 0), S11, &HD76AA478
    FF D, a, b, c, X(k + 1), S12, &HE8C7B756
    FF c, D, a, b, X(k + 2), S13, &H242070DB
    FF b, c, D, a, X(k + 3), S14, &HC1BDCEEE
    FF a, b, c, D, X(k + 4), S11, &HF57C0FAF
    FF D, a, b, c, X(k + 5), S12, &H4787C62A
    FF c, D, a, b, X(k + 6), S13, &HA8304613
    FF b, c, D, a, X(k + 7), S14, &HFD469501
    FF a, b, c, D, X(k + 8), S11, &H698098D8
    FF D, a, b, c, X(k + 9), S12, &H8B44F7AF
    FF c, D, a, b, X(k + 10), S13, &HFFFF5BB1
    FF b, c, D, a, X(k + 11), S14, &H895CD7BE
    FF a, b, c, D, X(k + 12), S11, &H6B901122
    FF D, a, b, c, X(k + 13), S12, &HFD987193
    FF c, D, a, b, X(k + 14), S13, &HA679438E
    FF b, c, D, a, X(k + 15), S14, &H49B40821

    GG a, b, c, D, X(k + 1), S21, &HF61E2562
    GG D, a, b, c, X(k + 6), S22, &HC040B340
    GG c, D, a, b, X(k + 11), S23, &H265E5A51
    GG b, c, D, a, X(k + 0), S24, &HE9B6C7AA
    GG a, b, c, D, X(k + 5), S21, &HD62F105D
    GG D, a, b, c, X(k + 10), S22, &H2441453
    GG c, D, a, b, X(k + 15), S23, &HD8A1E681
    GG b, c, D, a, X(k + 4), S24, &HE7D3FBC8
    GG a, b, c, D, X(k + 9), S21, &H21E1CDE6
    GG D, a, b, c, X(k + 14), S22, &HC33707D6
    GG c, D, a, b, X(k + 3), S23, &HF4D50D87
    GG b, c, D, a, X(k + 8), S24, &H455A14ED
    GG a, b, c, D, X(k + 13), S21, &HA9E3E905
    GG D, a, b, c, X(k + 2), S22, &HFCEFA3F8
    GG c, D, a, b, X(k + 7), S23, &H676F02D9
    GG b, c, D, a, X(k + 12), S24, &H8D2A4C8A
        
    HH a, b, c, D, X(k + 5), S31, &HFFFA3942
    HH D, a, b, c, X(k + 8), S32, &H8771F681
    HH c, D, a, b, X(k + 11), S33, &H6D9D6122
    HH b, c, D, a, X(k + 14), S34, &HFDE5380C
    HH a, b, c, D, X(k + 1), S31, &HA4BEEA44
    HH D, a, b, c, X(k + 4), S32, &H4BDECFA9
    HH c, D, a, b, X(k + 7), S33, &HF6BB4B60
    HH b, c, D, a, X(k + 10), S34, &HBEBFBC70
    HH a, b, c, D, X(k + 13), S31, &H289B7EC6
    HH D, a, b, c, X(k + 0), S32, &HEAA127FA
    HH c, D, a, b, X(k + 3), S33, &HD4EF3085
    HH b, c, D, a, X(k + 6), S34, &H4881D05
    HH a, b, c, D, X(k + 9), S31, &HD9D4D039
    HH D, a, b, c, X(k + 12), S32, &HE6DB99E5
    HH c, D, a, b, X(k + 15), S33, &H1FA27CF8
    HH b, c, D, a, X(k + 2), S34, &HC4AC5665

    II a, b, c, D, X(k + 0), S41, &HF4292244
    II D, a, b, c, X(k + 7), S42, &H432AFF97
    II c, D, a, b, X(k + 14), S43, &HAB9423A7
    II b, c, D, a, X(k + 5), S44, &HFC93A039
    II a, b, c, D, X(k + 12), S41, &H655B59C3
    II D, a, b, c, X(k + 3), S42, &H8F0CCC92
    II c, D, a, b, X(k + 10), S43, &HFFEFF47D
    II b, c, D, a, X(k + 1), S44, &H85845DD1
    II a, b, c, D, X(k + 8), S41, &H6FA87E4F
    II D, a, b, c, X(k + 15), S42, &HFE2CE6E0
    II c, D, a, b, X(k + 6), S43, &HA3014314
    II b, c, D, a, X(k + 13), S44, &H4E0811A1
    II a, b, c, D, X(k + 4), S41, &HF7537E82
    II D, a, b, c, X(k + 11), S42, &HBD3AF235
    II c, D, a, b, X(k + 2), S43, &H2AD7D2BB
    II b, c, D, a, X(k + 9), S44, &HEB86D391

    a = AddUnsigned(a, AA)
    b = AddUnsigned(b, BB)
    c = AddUnsigned(c, CC)
    D = AddUnsigned(D, DD)
  Next
  
  cyMD5_NET = UCase(WordToHexEx(a) & WordToHexEx(b) & WordToHexEx(c) & WordToHexEx(D))
End Function
'-----------------------------------------------------------------------------
'32位MD5
'-----------------------------------------------------------------------------

Function cyMD5_NETtoMD5(ByVal sMD5_NET As String) As String
    cyMD5_NETtoMD5 = LCase(Mid(sMD5_NET, 9, 16))
End Function

'---------------------------------------------------------------------------------------
' 过程名: cyCheckKeywords
' 生成日期: 2006-3-4 08:35
' 作者: Dragonschool
' 过程作用:检测整张帖子，如果包含在关键词列表中则返回真,另外设置三个属性记录出现的关键字是哪个，出现的位置及长度

'属性：
'   KeywordFoundWord As String      --出现的关键词
'   KeywordFoundPos As Long         --出现的位置
'   KeywordFoundLen As Long         --长度

'参数：
'   sString             --待检测的内容
'   sKeywordList        --关键词列表，多个关键词间用;来分隔
'   iMaxInterval        --关键词中关键字的最大间隔距离如　测试是关键词，如果间隔为３则“检测和期中试”将不会返为真
'---------------------------------------------------------------------------------------

Function cyCheckKeywords(ByVal sString As String, ByVal sKeywordList As String, Optional ByVal iMaxInterval As Long = 5) As Boolean
 
'处理转义字符
'==================HTML
    sString = Replace(sString, "&amp;", "&")
    sString = Replace(sString, "&#38;", "&")
 
    sString = Replace(sString, "&lt;", "<")
    sString = Replace(sString, "&#60;", "<")
 
    sString = Replace(sString, "&gt;", ">")
    sString = Replace(sString, "&#62;", ">")
 
    sString = Replace(sString, "&quot;", """")
    sString = Replace(sString, "&nbsp;", " ")
    sString = Replace(sString, "&copy;", " ") '版权暂时不能解释
    sString = Replace(sString, "&reg;", " ")  '商标暂时不能解释
    
 '=================XML
    sString = Replace(sString, "&lt;", "<")
    sString = Replace(sString, "&gt;", ">")
    sString = Replace(sString, "&amp;", "&")
    sString = Replace(sString, "&apos;", "'")
    sString = Replace(sString, "&quot;", """")
 
'==================繁体 转 简体
    Dim iStrLen  As Long
    Dim sStr As String
    
    iStrLen = lstrlen(sString)
    sStr = Space(iStrLen)
    LCMapString &H804, &H2000000, sString, iStrLen, sStr, iStrLen
    sString = sStr
 
    '初始值为空
    KeywordFoundWord = ""
    KeywordFoundPos = 0
    
    Dim sA() As String
    Dim i As Long
    '将关键词列表分割成数组,数组大小取决于关键词数量
    sA = Split(sKeywordList, ";")
    
    '对数组进行循环检测
    For i = 0 To UBound(sA)

        '检测单个关键词,如果跟某个关键词匹配则返回真
        If CheckKeywordEx(sString, sA(i), iMaxInterval) Then
            cyCheckKeywords = True
            Exit Function
            
        Else
            KeywordFoundWord = ""
            KeywordFoundPos = 0
        
        End If
    
    Next

End Function

Private Function CheckKeywordEx(ByVal sString As String, ByVal sKeyword As String, ByVal iMaxInterval As Long) As Boolean
    
    Dim i As Long
    Dim iKeywordPos  As Long    '记录内容中关键词所处的位置
    Dim iLastPos As Long        '最后一个关键出现的位置
    Dim iNextFoundPos As Long   '刻录发现第一个关键字的位置，后面可能还有其它关键字
    
    '避免空关键字时返回真
    If Trim(sKeyword) = "" Then Exit Function
    
    iMaxInterval = iMaxInterval + 1
    
ContinueCheck:
    
    '检查内容中第一个关键字的位置
    iKeywordPos = InStr(1, sString, Mid(sKeyword, 1, 1))
    KeywordFoundPos = KeywordFoundPos + iKeywordPos
    If iKeywordPos > 0 Then
    
        '处理关键词只有一个字的情况
        If Len(sKeyword) = 1 Then
            iLastPos = KeywordFoundPos
            GoTo FoundProc
            
        End If
    
    
        '记录最后位置
        iLastPos = iKeywordPos
        '如果所有的关键字都存在于内容中,则表示完全匹配,则返回真
        If Len(sKeyword) = i Then
            KeywordFoundPos = 1
            GoTo FoundProc
        End If
        
                For i = 2 To Len(sKeyword)
                    '检查内容中第一个关键字的位置
                    iKeywordPos = InStr(1, Mid(sString, iLastPos + 1, iMaxInterval), Mid(sKeyword, i, 1))
                    '位置>0才继续检查
                    If iKeywordPos > 0 Then
                        '记录最后位置
                        iLastPos = iLastPos + iKeywordPos
                        '如果所有的关键字都存在于内容中,则表示完全匹配,则返回真
                        If Len(sKeyword) = i Then
                            GoTo FoundProc
                        End If
                    Else
                    '在循环中任意一个关键字找不到则返回假并退出
                    
                        '继续检测剩下的内容中是否存在第一个关键字，如果还有则继续对余下内容进行检测，没有则真正退出
                        iKeywordPos = InStr(1, Right(sString, Len(sString) - iLastPos), Mid(sKeyword, 1, 1))
                        
                        If iKeywordPos > 0 Then
                            '将剩下的内容作为完整内容再进行检测，并将关键字位置设为1
                            sString = Right(sString, Len(sString) - iLastPos)
                            i = 1
                            GoTo ContinueCheck
                        
                        End If
                        
                        CheckKeywordEx = False
                        Exit Function
                
                    End If
                
                Next
        
    Else
        '第一个关键字也找不到
        CheckKeywordEx = False
        Exit Function
    
    End If
    Exit Function
    
'发现
FoundProc:
    KeywordFoundWord = sKeyword
    KeywordFoundLen = iLastPos - KeywordFoundPos + 1
    CheckKeywordEx = True
End Function
'-----------------------------------------------------------------------------
'范式检测关键字
'-----------------------------------------------------------------------------

