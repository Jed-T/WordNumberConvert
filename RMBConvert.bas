Attribute VB_Name = "模块1"
Option Explicit


'==================================================
' 宏1
' 1234 -> 壹贰叁肆
'==================================================
Public Sub NumToChineseDigits()

    Dim s As String
    Dim result As String
    Dim i As Long
    Dim c As String
    Dim nums As Variant

    nums = Array("零", "壹", "贰", "叁", "肆", _
                 "伍", "陆", "柒", "捌", "玖")

    s = Selection.Text

    If Len(s) = 0 Then
        MsgBox "请先选中需要转换的内容。"
        Exit Sub
    End If

    result = ""

    For i = 1 To Len(s)

        c = Mid$(s, i, 1)

        If c >= "0" And c <= "9" Then
            result = result & nums(CInt(c))
        Else
            result = result & c
        End If

    Next i

    Selection.Text = result

End Sub


'==================================================
' 宏2
' 1234 -> 壹仟贰佰叁拾肆
'==================================================
Public Sub NumToChineseValue()

    Dim s As String
    Dim n As Double

    s = Trim$(Selection.Text)

    If s = "" Then
        MsgBox "请先选中数字。"
        Exit Sub
    End If

    If Not IsNumeric(s) Then
        MsgBox "选中的内容必须是数字。"
        Exit Sub
    End If

    n = CDbl(s)

    If n < 0 Then
        MsgBox "暂不支持负数。"
        Exit Sub
    End If

    Selection.Text = CNInteger(CStr(Fix(n)))

End Sub


'==================================================
' 宏3
' 1234.56 -> 壹仟贰佰叁拾肆元伍角陆分
'==================================================
Public Sub NumToRMB()

    Dim s As String
    Dim amount As Double

    Dim integerText As String
    Dim cents As Long
    Dim jiao As Integer
    Dim fen As Integer

    Dim result As String

    s = Trim$(Selection.Text)

    If s = "" Then
        MsgBox "请先选中金额。"
        Exit Sub
    End If

    If Not IsNumeric(s) Then
        MsgBox "选中的内容必须是有效金额。"
        Exit Sub
    End If

    amount = CDbl(s)

    If amount < 0 Then
        MsgBox "暂不支持负数金额。"
        Exit Sub
    End If

    integerText = CStr(Fix(amount))

    cents = CLng(Round((amount - Fix(amount)) * 100, 0))

    If cents = 100 Then
        integerText = CStr(CDbl(integerText) + 1)
        cents = 0
    End If

    jiao = cents \ 10
    fen = cents Mod 10

    result = CNInteger(integerText) & "元"

    If jiao = 0 And fen = 0 Then

        result = result & "整"

    Else

        If jiao > 0 Then
            result = result & CNDigit(jiao) & "角"
        ElseIf fen > 0 Then
            result = result & "零"
        End If

        If fen > 0 Then
            result = result & CNDigit(fen) & "分"
        End If

    End If

    Selection.Text = result

End Sub


'==================================================
' 以下为辅助函数
' 不需要手动运行
'==================================================

Public Function CNDigit(ByVal n As Integer) As String

    Dim nums As Variant

    nums = Array("零", "壹", "贰", "叁", "肆", _
                 "伍", "陆", "柒", "捌", "玖")

    CNDigit = nums(n)

End Function


Public Function CNFour(ByVal n As Integer) As String

    Dim nums As Variant
    Dim units As Variant

    Dim result As String
    Dim pos As Integer
    Dim digit As Integer
    Dim divisor As Integer
    Dim needZero As Boolean

    nums = Array("零", "壹", "贰", "叁", "肆", _
                 "伍", "陆", "柒", "捌", "玖")

    units = Array("", "拾", "佰", "仟")

    result = ""
    needZero = False

    For pos = 3 To 0 Step -1

        divisor = 10 ^ pos
        digit = (n \ divisor) Mod 10

        If digit <> 0 Then

            If needZero And result <> "" Then
                result = result & "零"
            End If

            result = result & nums(digit) & units(pos)

            needZero = False

        Else

            If result <> "" Then
                needZero = True
            End If

        End If

    Next pos

    CNFour = result

End Function


Public Function CNInteger(ByVal numberText As String) As String

    Dim bigUnits As Variant

    Dim result As String
    Dim part As String

    Dim groupValue As Integer
    Dim groupIndex As Integer

    Dim needZero As Boolean

    bigUnits = Array("", "万", "亿", "兆")

    numberText = Trim$(numberText)

    Do While Len(numberText) > 1 And Left$(numberText, 1) = "0"
        numberText = Mid$(numberText, 2)
    Loop

    If numberText = "" Or numberText = "0" Then
        CNInteger = "零"
        Exit Function
    End If

    result = ""
    groupIndex = 0
    needZero = False

    Do While Len(numberText) > 0

        If Len(numberText) > 4 Then

            part = Right$(numberText, 4)
            numberText = Left$(numberText, Len(numberText) - 4)

        Else

            part = numberText
            numberText = ""

        End If

        groupValue = CInt(part)

        If groupValue <> 0 Then

            If needZero And result <> "" Then
                result = "零" & result
            End If

            result = CNFour(groupValue) & _
                     bigUnits(groupIndex) & result

            If groupValue < 1000 Then
                needZero = True
            Else
                needZero = False
            End If

        Else

            If result <> "" Then
                needZero = True
            End If

        End If

        groupIndex = groupIndex + 1

    Loop

    CNInteger = result

End Function

