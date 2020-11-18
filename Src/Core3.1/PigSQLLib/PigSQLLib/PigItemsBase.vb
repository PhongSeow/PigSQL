'**********************************
'* Name: PigItemsBase
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: An interface is very close to the VB6 set base class. 一个非常接近VB6接口的集合基本类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.3
'* Create Time: 12/11/2020
'* 1.0.2 13/11/2020 Add Item 增加 Item
'* 1.0.3 14/11/2020 Update function Add 修改函数 Add
'************************************
Imports System.Collections
Public Class PigItemsBase
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.3"
    Private mCol As New Collection

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public ReadOnly Property Item(Key As String) As Object
        Get
            Try
                Return mCol.Item(Key)
                Me.ClearErr()
            Catch ex As Exception
                Me.SetSubErrInf("Item", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(Index As Integer) As Object
        Get
            Try
                Return mCol.Item(Index)
                Me.ClearErr()
            Catch ex As Exception
                Me.SetSubErrInf("Item", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Protected Overrides Sub Finalize()
        mCol = Nothing
        MyBase.Finalize()
    End Sub

    Public ReadOnly Property GetItems As Object()
        Get
            Dim strStepName As String = ""
            Try
                strStepName = "Dim"
                Dim geAny As IEnumerator = mCol.GetEnumerator()
                Dim aoItems As Object()
                ReDim aoItems(mCol.Count - 1)
                strStepName = "Reset"
                geAny.Reset()
                Dim i As Integer = 0
                strStepName = "While.MoveNext"
                While geAny.MoveNext()
                    Dim oItem As Object
                    oItem = geAny.Current()
                    aoItems(i) = oItem
                    i += 1
                End While
                GetItems = aoItems
                aoItems = Nothing
                Me.ClearErr()
            Catch ex As Exception
                Me.SetSubErrInf("GetItems", strStepName, ex)
                GetItems = Nothing
            End Try
        End Get
    End Property


    Public Function Add(Key As String, ObjItem As Object) As Object
        Try
            mCol.Add(ObjItem, Key)
            Me.ClearErr()
            Return mCol.Item(Key)
        Catch ex As Exception
            Me.SetSubErrInf("Add", ex)
            Return Nothing
        End Try
    End Function

    Public ReadOnly Property Count As Integer
        Get
            Try
                Count = mCol.Count
                Me.ClearErr()
            Catch ex As Exception
                Me.SetSubErrInf("Count", ex)
                Count = -1
            End Try
        End Get
    End Property

    Public Sub Remove(Index As Integer)
        Try
            mCol.Remove(Index)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Remove", ex)
        End Try
    End Sub

    Public Sub Remove(Key As String)
        Try
            mCol.Remove(Key)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Remove", ex)
        End Try
    End Sub

    Public Sub Clear()
        Try
            mCol.Clear()
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Clear", ex)
        End Try
    End Sub

End Class
