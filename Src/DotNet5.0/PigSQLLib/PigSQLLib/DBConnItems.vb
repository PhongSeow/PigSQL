'**********************************
'* Name: DBConnItems
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Database connection collection
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.3
'* Create Time: 13/11/2020
'* 1.0.2    13/11/2020  Update Function Add 
'* 1.0.3    23/11/2020  Update Function Add 
'************************************
Public Class DBConnItems
    Inherits PigItemsBase
    ''' <summary>
    ''' SQL Server 单机模式-信任连接
    ''' </summary>
    ''' <param name="DBConnName">数据库连接名称</param>
    ''' <param name="DBSrv">数据库服务器</param>
    ''' <param name="DefaultDB">默认数据库</param>
    Public Shadows Function Add(DBConnName As String, DBSrv As String, DefaultDB As String) As DBConnItem
        Dim oNewItem As New DBConnItem(DBConnName, DBSrv, DefaultDB)
        Return MyBase.Add(DBConnName, oNewItem)
    End Function

    ''' <summary>
    ''' SQL Server 单机模式
    ''' </summary>
    ''' <param name="DBConnName">数据库连接名称</param>
    ''' <param name="DBSrv">数据库服务器</param>
    ''' <param name="DBUser">数据库用户</param>
    ''' <param name="DBPwd">数据库用户密码</param>
    ''' <param name="DefaultDB">默认数据库</param>
    Public Shadows Function Add(DBConnName As String, DBSrv As String, DBUser As String, DBPwd As String, DefaultDB As String) As DBConnItem
        Dim oNewItem As New DBConnItem(DBConnName, DBSrv, DBUser, DBUser, DefaultDB)
        Return MyBase.Add(DBConnName, oNewItem)
    End Function

    ''' <summary>
    ''' SQL Server 双机镜像模式
    ''' </summary>
    ''' <param name="DBConnName">数据库连接名称</param>
    ''' <param name="MasterDBSrv">主数据库服务器</param>
    ''' <param name="MirrorDBSrv">镜像数据库服务器</param>
    ''' <param name="DBUser">数据库用户</param>
    ''' <param name="DBPwd">数据库用户密码</param>
    ''' <param name="DefaultDB">默认数据库</param>
    Public Shadows Function Add(DBConnName As String, DBSrv As String, MirrorDBSrv As String, DBUser As String, DBPwd As String, DefaultDB As String) As DBConnItem
        Dim oNewItem As New DBConnItem(DBConnName, DBSrv, MirrorDBSrv, DBUser, DBUser, DefaultDB)
        Return MyBase.Add(DBConnName, oNewItem)
    End Function

    ''' <summary>
    ''' SQL Server 双机镜像模式-信任连接
    ''' </summary>
    ''' <param name="DBConnName">数据库连接名称</param>
    ''' <param name="MasterDBSrv">主数据库服务器</param>
    ''' <param name="MirrorDBSrv">镜像数据库服务器</param>
    ''' <param name="DefaultDB">默认数据库</param>
    Public Shadows Function Add(DBConnName As String, MasterDBSrv As String, MirrorDBSrv As String, DefaultDB As String) As DBConnItem
        Dim oNewItem As New DBConnItem(DBConnName, MirrorDBSrv, DefaultDB)
        Return MyBase.Add(DBConnName, oNewItem)
    End Function


    Public Shadows ReadOnly Property Item(DBConnName As String) As DBConnItem
        Get
            Return MyBase.Item(DBConnName)
        End Get
    End Property
    Public Shadows ReadOnly Property Item(Index As Integer) As DBConnItem
        Get
            Return MyBase.Item(Index)
        End Get
    End Property

End Class
