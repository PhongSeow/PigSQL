''' <summary>
'''************************************
'''* Name: DBConnItem
'''* Author: Seow Phong
'''* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'''* Describe: Database connection item
'''* Home Url: https://www.seowphong.com or https://en.seowphong.com
'''* Version: 1.0.2
'''* Create Time: 13/11/2020
'''* 1.0.2    13/11/2020  Add some Property
'''************************************
''' </summary>
Public Class DBConnItem
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.2"

	''' <summary>
	''' 连接名称
	''' </summary>
	Private mstrDBConnName As String
	Public Property DBConnName() As String
		Get
			Return mstrDBConnName
		End Get
		Friend Set(ByVal value As String)
			mstrDBConnName = value
		End Set
	End Property

	''' <summary>
	''' 连接超时时间
	''' </summary>
	Private mvarConnectionTimeout As Long
	Public Property ConnectionTimeout() As Long
		Get
			Return mvarConnectionTimeout
		End Get
		Friend Set(ByVal value As Long)
			mvarConnectionTimeout = value
		End Set
	End Property

	''' <summary>
	''' 执行超时时间
	''' </summary>
	Private mvarCommandTimeout As Long
	Public Property CommandTimeout() As Long
		Get
			Return mvarCommandTimeout
		End Get
		Friend Set(ByVal value As Long)
			mvarCommandTimeout = value
		End Set
	End Property

	''' <summary>
	''' 数据库服务器
	''' </summary>
	Private mstrDBSrv As String
	Public Property DBSrv() As String
		Get
			Return mstrDBSrv
		End Get
		Friend Set(ByVal value As String)
			mstrDBSrv = value
		End Set
	End Property

	''' <summary>
	''' 镜像数据库服务器，如果设置则
	''' </summary>
	Private mstrMirDBSrv As String
	Public Property MirDBSrv() As String
		Get
			Return mstrMirDBSrv
		End Get
		Friend Set(ByVal value As String)
			mstrMirDBSrv = value
		End Set
	End Property

	''' <summary>
	''' 备用数据库服务器，如果运行在此，数据库只能读，不可以写

	''' </summary>
	Private mstrSpareDBSrv As String
	Public Property SpareDBSrv() As String
		Get
			Return mstrSpareDBSrv
		End Get
		Friend Set(ByVal value As String)
			mstrSpareDBSrv = value
		End Set
	End Property

	''' <summary>
	''' 数据库用户名
	''' </summary>
	Private mstrDBUser As String
	Public Property DBUser() As String
		Get
			Return mstrDBUser
		End Get
		Friend Set(ByVal value As String)
			mstrDBUser = value
		End Set
	End Property

	''' <summary>
	''' 默认数据库

	''' </summary>
	Private mstrDefaultDB As String
	Public Property DefaultDB() As String
		Get
			Return mstrDefaultDB
		End Get
		Friend Set(ByVal value As String)
			mstrDefaultDB = value
		End Set
	End Property

	''' <summary>
	''' 数据库密码

	''' </summary>
	Private mstrDBPwd As String
	Public Property DBPwd() As String
		Get
			Return mstrDBPwd
		End Get
		Friend Set(ByVal value As String)
			mstrDBPwd = value
		End Set
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
	''' SQL Server 单机模式
	''' </summary>
	''' <param name="DBConnName"></param>
	''' <param name="DBSrv"></param>
	''' <param name="DBUser"></param>
	''' <param name="DBPwd"></param>
	''' <param name="DefaultDB"></param>
	Public Sub New(DBConnName As String, DBSrv As String, DBUser As String, DBPwd As String, DefaultDB As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(DBConnName, DBSrv, DBUser, DBPwd, DefaultDB)
	End Sub

	''' <summary>
	''' SQL Server 双机镜像模式
	''' </summary>
	''' <param name="DBConnName"></param>
	''' <param name="DBSrv"></param>
	''' <param name="MirDBSrv"></param>
	''' <param name="DBUser"></param>
	''' <param name="DBPwd"></param>
	''' <param name="DefaultDB"></param>
	Public Sub New(DBConnName As String, DBSrv As String, MirDBSrv As String, DBUser As String, DBPwd As String, DefaultDB As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(DBConnName, DBSrv, DBUser, DBPwd, DefaultDB, MirDBSrv)
	End Sub

	''' <summary>
	''' SQL Server 双机镜像加备用数据库模式，备用数据库为副本，只能读

	''' </summary>
	''' <param name="DBConnName"></param>
	''' <param name="DBSrv"></param>
	''' <param name="MirDBSrv"></param>
	''' <param name="SpareDBSrv"></param>
	''' <param name="DBUser"></param>
	''' <param name="DBPwd"></param>
	''' <param name="DefaultDB"></param>
	Public Sub New(DBConnName As String, DBSrv As String, MirDBSrv As String, SpareDBSrv As String, DBUser As String, DBPwd As String, DefaultDB As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(DBConnName, DBSrv, DBUser, DBPwd, DefaultDB, MirDBSrv, SpareDBSrv)
	End Sub

	Private Sub mNew(DBConnName As String, DBSrv As String, DBUser As String, DBPwd As String, DefaultDB As String, Optional MirDBSrv As String = "", Optional SpareDBSrv As String = "")
		With Me
			.DBConnName = DBConnName
			.DBSrv = DBSrv
			.MirDBSrv = MirDBSrv
			.SpareDBSrv = SpareDBSrv
			.DBUser = DBUser
			.DBPwd = DBPwd
			.DefaultDB = DefaultDB
		End With
	End Sub

End Class
