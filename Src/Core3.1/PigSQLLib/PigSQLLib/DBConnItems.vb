'**********************************
'* Name: DBConnItems
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Database connection collection
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.2
'* Create Time: 13/11/2020
'* 1.0.2    13/11/2020  Update Function Add 
'************************************
Public Class DBConnItems
    Inherits PigItemsBase
    Public Shadows Function Add(DBConnName As String) As DBConnItem
        Dim oNewItem As New DBConnItem(DBConnName)
        Return MyBase.Add(Name, oNewItem)
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
