''' <summary>
'''************************************
'''* Name: DBConnItem
'''* Author: Seow Phong
'''* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'''* Describe: Database connection item
'''* Home Url: https://www.seowphong.com or https://en.seowphong.com
'''* Version: 1.0.4
'''* Create Time: 13/11/2020
'''* 1.0.2    13/11/2020  Add some Property
'''* 1.0.3    20/11/2020  Add GetConnStr
'''* 1.0.4    22/11/2020  Add RefConn,mNew
'''************************************
''' </summary>

Imports System.Data.SqlClient

Public Class DBConnItem
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.4"
	Private moSqlConnection As SqlConnection

	''' <summary>
	''' 连接状态	''' </summary>
	Public Enum enmConnState
		''' <summary>
		''' 未知
		''' </summary>
		Unknow = 0
		''' <summary>
		''' 主数据库在线
		''' </summary>
		MasterDBOnline = 10
		''' <summary>
		''' 镜像数据库在线		''' </summary>
		MirrorDBOnline = 20
		''' <summary>
		''' 全部数据库离线		''' </summary>
		AllDBOffLine = 40
	End Enum

	''' <summary>
	''' 连接名称
	''' </summary>
	Private mstrDBConnName As String
	Public ReadOnly Property DBConnName() As String
		Get
			Return mstrDBConnName
		End Get
	End Property

	''' <summary>
	''' 连接超时时间
	''' </summary>
	Private mvarConnectionTimeout As Long
	Public ReadOnly Property ConnectionTimeout() As Long
		Get
			Return mvarConnectionTimeout
		End Get
	End Property

	''' <summary>
	''' SqlConnection
	''' </summary>
	Private mobjSqlConnection As SqlConnection
	Public ReadOnly Property SqlConnection() As SqlConnection
		Get
			Return mobjSqlConnection
		End Get
	End Property

	''' <summary>
	''' 执行超时时间，单位：秒	''' </summary>
	Private mintCommandTimeout As Integer = 60
	Public ReadOnly Property CommandTimeout() As Integer
		Get
			Return mintCommandTimeout
		End Get
	End Property

	''' <summary>
	''' 是否信任连接
	''' </summary>
	Private mbolIsTrustConn As Boolean = False
	Public ReadOnly Property IsTrustConn() As Boolean
		Get
			Return mbolIsTrustConn
		End Get
	End Property

	''' <summary>
	''' 是否运行在镜像数据库
	''' </summary>
	Private mbolIsRenAtMirror As Boolean
	Public ReadOnly Property IsRenAtMirror() As Boolean
		Get
			Return mbolIsRenAtMirror
		End Get
	End Property

	''' <summary>
	''' 数据库服务器
	''' </summary>
	Private mstrDBSrv As String
	Public ReadOnly Property DBSrv() As String
		Get
			Return mstrDBSrv
		End Get
	End Property

	''' <summary>
	''' 镜像数据库服务器，如果设置则
	''' </summary>
	Private mstrMirDBSrv As String
	Public ReadOnly Property MirDBSrv() As String
		Get
			Return mstrMirDBSrv
		End Get
	End Property

	''' <summary>
	''' 数据库用户名
	''' </summary>
	Private mstrDBUser As String
	Public ReadOnly Property DBUser() As String
		Get
			Return mstrDBUser
		End Get
	End Property

	''' <summary>
	''' 默认数据库	''' </summary>
	Private mstrDefaultDB As String
	Public ReadOnly Property DefaultDB() As String
		Get
			Return mstrDefaultDB
		End Get
	End Property

	''' <summary>
	''' 数据库密码	''' </summary>
	Private mstrDBPwd As String
	Public ReadOnly Property DBPwd() As String
		Get
			Return mstrDBPwd
		End Get
	End Property

	''' <summary>
	''' 备注
	''' </summary>
	Private mstrRemark As String
	Public Property Remark() As String
		Get
			Return mstrRemark
		End Get
		Friend Set(ByVal value As String)
			mstrRemark = value
		End Set
	End Property

	''' <summary>
	''' 数据库是否忙
	''' </summary>
	Public ReadOnly Property IsDBBusy() As Boolean
		Get
			If Not moSqlConnection Is Nothing Then
				Select Case moSqlConnection.State
					Case Data.ConnectionState.Open, Data.ConnectionState.Closed
						Return False
					Case Else
						Return True
				End Select
			Else
				Return False
			End If
		End Get
	End Property

	''' <summary>
	''' 数据库是否就绪		''' </summary>
	Public ReadOnly Property IsDBReady() As Boolean
		Get
			If Not moSqlConnection Is Nothing Then
				Select Case moSqlConnection.State
					Case Data.ConnectionState.Open
						If Me.IsHAMode = False Then
							Return True
						ElseIf Me.IsDefaDBOffLine = True Then
							Return False
						Else
							Return True
						End If
					Case Else
						Return False
				End Select
			Else
				Return False
			End If
		End Get
	End Property

	''' <summary>
	''' 是否默认数据库不在线
	''' </summary>
	Private mbolIsDefaDBOffLine As Boolean
	Public ReadOnly Property IsDefaDBOffLine() As Boolean
		Get
			Return mbolIsDefaDBOffLine
		End Get
	End Property

	''' <summary>
	''' 是否高可用模式	''' </summary>
	Private mbolIsHAMode As Boolean
	Public ReadOnly Property IsHAMode() As Boolean
		Get
			Return mbolIsHAMode
		End Get
	End Property

	''' <summary>
	''' 最近刷新连接时间	''' </summary>
	Private mdteLastRefTime As Date
	Public ReadOnly Property LastRefTime() As Date
		Get
			Return mdteLastRefTime
		End Get
	End Property

	''' <summary>
	''' 刷新连接时间间隔，单位：秒。	''' </summary>
	Private mvarRefTimeInterval As Integer = 30
	Public ReadOnly Property RefTimeInterval() As Integer
		Get
			Return mvarRefTimeInterval
		End Get
	End Property

	''' <summary>
	''' 最近的连接失败信息
	''' </summary>
	Private mstrLastConnErr As String
	Public ReadOnly Property LastConnErr() As String
		Get
			Return mstrLastConnErr
		End Get
	End Property

	''' <summary>
	''' 最近的执行失败信息
	''' </summary>
	Private mstrLastExecErr As String
	Public ReadOnly Property LastExecErr() As String
		Get
			Return mstrLastExecErr
		End Get
	End Property

	''' <summary>
	''' 连接状态

	''' </summary>
	Private mvarConnState As enmConnState = enmConnState.Unknow
	Public ReadOnly Property ConnState() As enmConnState
		Get
			Return mvarConnState
		End Get
	End Property

	''' <summary>
	''' 最近是否连接到镜像数据库

	''' </summary>
	Private mbolIsLastConn2Mirror As Boolean = False
	Public ReadOnly Property IsLastConn2Mirror() As Boolean
		Get
			Return mbolIsLastConn2Mirror
		End Get
	End Property

	''' <summary>
	''' SQL Server 单机模式
	''' </summary>
	''' <param name="DBConnName">数据库连接名称</param>
	''' <param name="DBSrv">数据库服务器</param>
	''' <param name="DBUser">数据库用户</param>
	''' <param name="DBPwd">数据库用户密码</param>
	''' <param name="DefaultDB">默认数据库</param>
	Public Sub New(DBConnName As String, DBSrv As String, DBUser As String, DBPwd As String, DefaultDB As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(DBConnName, DBSrv, DefaultDB, False, True, DBUser, DBPwd)
	End Sub

	''' <summary>
	''' SQL Server 单机模式-信任连接
	''' </summary>
	''' <param name="DBConnName">数据库连接名称</param>
	''' <param name="DBSrv">数据库服务器</param>
	''' <param name="DefaultDB">默认数据库</param>
	Public Sub New(DBConnName As String, DBSrv As String, DefaultDB As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(DBConnName, DBSrv, DefaultDB, False, True)
	End Sub

	''' <summary>
	''' SQL Server 双机镜像模式
	''' </summary>
	''' <param name="DBConnName">数据库连接名称</param>
	''' <param name="MasterDBSrv">主数据库服务器</param>
	''' <param name="MirrorDBSrv">镜像数据库服务器</param>
	''' <param name="DBUser">数据库用户</param>
	''' <param name="DBPwd">数据库用户密码</param>
	''' <param name="DefaultDB">默认数据库</param>
	Public Sub New(DBConnName As String, DBSrv As String, MirrorDBSrv As String, DBUser As String, DBPwd As String, DefaultDB As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(DBConnName, DBSrv, DefaultDB, True, False, DBUser, DBPwd, MirrorDBSrv)
	End Sub

	''' <summary>
	''' SQL Server 双机镜像模式-信任连接
	''' </summary>
	''' <param name="DBConnName">数据库连接名称</param>
	''' <param name="MasterDBSrv">主数据库服务器</param>
	''' <param name="MirrorDBSrv">镜像数据库服务器</param>
	''' <param name="DefaultDB">默认数据库</param>
	Public Sub New(DBConnName As String, MasterDBSrv As String, MirrorDBSrv As String, DefaultDB As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(DBConnName, DBSrv, DefaultDB, True, True, , , MirrorDBSrv)
	End Sub


	Private Sub mNew(DBConnName As String, DBSrv As String, DefaultDB As String, IsHAMode As Boolean, IsTrustConn As Boolean, Optional DBUser As String = "", Optional DBPwd As String = "", Optional MirDBSrv As String = "")
		With Me
			mstrDBConnName = DBConnName
			mstrDBSrv = DBSrv
			mbolIsHAMode = IsHAMode
			If .IsHAMode = True Then
				mstrMirDBSrv = MirDBSrv
			Else
				mstrMirDBSrv = ""
			End If
			mbolIsTrustConn = IsTrustConn
			mstrDefaultDB = DefaultDB
			If .IsTrustConn = True Then
				mstrDBUser = DBUser
				mstrDBPwd = DBPwd
			Else
				mstrDBUser = ""
				mstrDBPwd = ""
			End If
		End With
	End Sub

	Private Function mGetConnStr(IsMirror As Boolean) As String
		mGetConnStr = "Data Source=" & Me.DBSrv & ";"
		If IsMirror = False Then
			mGetConnStr &= Me.DBSrv & ";"
		Else
			mGetConnStr &= Me.MirDBSrv & ";"
		End If
		mGetConnStr &= "Initial Catalog=" & Me.DefaultDB & ";"
		If Me.IsTrustConn = True Then
		Else
			mGetConnStr &= "User ID=" & Me.DBUser & ";"
			mGetConnStr &= "Password=" & Me.mSQLStr(Me.DBPwd) & ";"
		End If
		mGetConnStr &= "Application Name=PigSQL;"
		mGetConnStr &= "Connect Timeout=" & Me.ConnectionTimeout.ToString & ";"
	End Function

	Private Function mSQLStr(StrValue As String, Optional IsNotNull As Boolean = False) As String
		StrValue = Replace(StrValue, "'", "''")
		If UCase(StrValue) = "NULL" And IsNotNull = False Then
			mSQLStr = "NULL"
		Else
			mSQLStr = "'" & StrValue & "'"
		End If
	End Function

	''' <summary>
	''' 刷新数据库连接

	''' </summary>
	Public Sub RefConn()
		Dim strStepName As String = ""
		Try
			If Math.Abs(DateDiff(DateInterval.Second, Me.LastRefTime, Now)) > Me.RefTimeInterval Then
				Dim bolIsNew As Boolean = False
				If moSqlConnection Is Nothing Then bolIsNew = True

				Select Case Me.ConnState
					Case enmConnState.Unknow
						bolIsNew = True
					Case enmConnState.MasterDBOnline, enmConnState.MirrorDBOnline
					Case enmConnState.AllDBOffLine
						bolIsNew = True
				End Select
				If bolIsNew = True Then
					moSqlConnection = Nothing
					moSqlConnection = New SqlConnection
					With moSqlConnection
						.ConnectionString = Me.mGetConnStr(mbolIsLastConn2Mirror)
						.Open()

					End With
				End If

				mdteLastRefTime = Now
				Me.ClearErr()
			End If
		Catch ex As Exception
			Me.SetSubErrInf("RefConn", strStepName, ex)
		End Try
	End Sub


End Class
