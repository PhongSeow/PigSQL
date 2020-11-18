''' <summary>
'''************************************
'''* Name: SQLStatement
'''* Author: Seow Phong
'''* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'''* Describe: Responsible for SQL statement assembly and SQL injection prevention. 
'''* 描述: 负责SQL语句组装和防SQL注入等。
'''* Home Url: https://www.seowphong.com or https://en.seowphong.com
'''* Version: 1.0.1
'''* Create Time: 17/11/2020
'''* 1.0.2  18/11/2020  
'''************************************
''' </summary>
Public Class SQLStatement
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.1"

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    ''' <summary>
    ''' SQL字符串处理，有一定防SQL注入效果
    ''' SQL string processing, has a certain effect of preventing SQL injection
    ''' </summary>
    ''' <param name="StrValue">字符串值|String value</param>
    ''' <param name="IsNotNull">是否不允许空值|Is null not allowed</param>
    ''' <returns>返回处理过的字符串|Returns the processed string</returns>
    Public Function SQLStr(StrValue As String, Optional IsNotNull As Boolean = False) As String
        StrValue = Replace(StrValue, "'", "''")
        If UCase(StrValue) = "NULL" And IsNotNull = False Then
            SQLStr = "NULL"
        Else
            SQLStr = "'" & StrValue & "'"
        End If
    End Function

    Public Function SQLStr(DteValue As DateTime) As String
        SQLStr = "'" & Format(DteValue, "yyyy-MM-dd HH:mm:ss.fff") & "'"
    End Function

    ''' <summary>
    ''' SQL数值处理，有一定防SQL注入效果
    ''' SQL string processing, has a certain effect of preventing SQL injection
    ''' </summary>
    ''' <param name="NumValue">数值值|Numerical value</param>
    ''' <returns></returns>
    Public Function SQLNum(NumValue As String) As String
        If UCase(NumValue) = "NULL" Then
            SQLNum = "NULL"
        ElseIf IsNumeric(NumValue) = True Then
            SQLNum = NumValue
        Else
            SQLNum = ""
        End If
    End Function

    Public Function SQLNum(IntValue As Long) As String
        SQLNum = IntValue.ToString
    End Function

    Public Function SQLNum(DecValue As Decimal) As String
        SQLNum = DecValue.ToString
    End Function

    Public Function SQLNum(BolValue As Boolean) As String
        SQLNum = CInt(BolValue).ToString
    End Function


End Class
