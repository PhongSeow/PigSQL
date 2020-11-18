''' <summary>
'''************************************
'''* Name: SQLStateParaItem
'''* Author: Seow Phong
'''* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'''* Describe: Parameter processing of SQL statement.
'''* 描述: SQL语句的参数处理。
'''* Home Url: https://www.seowphong.com or https://en.seowphong.com
'''* Version: 1.0.1
'''* Create Time: 18/11/2020
'''************************************
''' </summary>
Public Class SQLStateParaItem
	''' <summary>
	''' 参数名称
	''' </summary>
	Private mstrParaName As String
	Public Property ParaName() As String
		Get
			Return mstrParaName
		End Get
		Friend Set(ByVal value As String)
			mstrParaName = value
		End Set
	End Property

	''' <summary>
	''' SQL数据类型
	''' </summary>
	Private mstrSQLDataType As String
	Public ReadOnly Property SQLDataType() As String
		Get
			Return mstrSQLDataType
		End Get
	End Property

	''' <summary>
	''' 参数字符串值
	''' </summary>
	Private mstrParaValue As String
	Public ReadOnly Property ParaValue() As String
		Get
			Return mstrParaValue
		End Get
	End Property

	''' <summary>
	''' 父对象
	''' </summary>
	Private mobjParent As Object
	Public Property Parent() As SQLStatement
		Get
			Return mobjParent
		End Get
		Friend Set(ByVal value As SQLStatement)
			mobjParent = value
		End Set
	End Property

	''' <summary>
	''' 新建
	''' </summary>
	''' <param name="ParaName">参数名称</param>
	''' <param name="Parent">父对象</param>
	Public Sub New(ParaName As String, Parent As SQLStatement)
		With Me
			.ParaName = ParaName
			.Parent = Parent
		End With
	End Sub


	''' <summary>
	''' 设置参数值-字符串
	''' </summary>
	''' <param name="StrValue">字符串值</param>
	''' <param name="SQLDataType">指定的SQL数据类型</param>
	Public Sub SetParaValue(StrValue As String, Optional SQLDataType As String = "")
		If SQLDataType = "" Then
			mstrSQLDataType = "nvarchar(" & Len(StrValue) * 2 & ")"
		Else
			mstrSQLDataType = SQLDataType
		End If
		mstrParaValue = Me.Parent.SQLStr(StrValue)
	End Sub


	''' <summary>
	''' 设置参数值-整型
	''' </summary>
	''' <param name="IntValue">整型值</param>
	''' <param name="SQLDataType">指定的SQL数据类型</param>
	Public Sub SetParaValue(IntValue As Long, Optional SQLDataType As String = "")
		If SQLDataType = "" Then
			mstrSQLDataType = "int"
		Else
			mstrSQLDataType = SQLDataType
		End If
		mstrParaValue = Me.Parent.SQLNum(IntValue)
	End Sub



	''' <summary>
	''' 设置参数值-浮点型
	''' </summary>
	''' <param name="DecValue">浮点型值</param>
	''' <param name="SQLDataType">指定的SQL数据类型</param>
	Public Sub SetParaValue(DecValue As Decimal, Optional SQLDataType As String = "")
		If SQLDataType = "" Then
			mstrSQLDataType = "float"
		Else
			mstrSQLDataType = SQLDataType
		End If
		mstrParaValue = Me.Parent.SQLNum(DecValue)
	End Sub

	''' <summary>
	''' 设置参数值-布尔型
	''' </summary>
	''' <param name="BolValue">布尔型值</param>
	''' <param name="SQLDataType">指定的SQL数据类型</param>
	Public Sub SetParaValue(BolValue As Boolean, Optional SQLDataType As String = "")
		If SQLDataType = "" Then
			mstrSQLDataType = "bit"
		Else
			mstrSQLDataType = SQLDataType
		End If
		mstrParaValue = Me.Parent.SQLNum(BolValue)
	End Sub
End Class
