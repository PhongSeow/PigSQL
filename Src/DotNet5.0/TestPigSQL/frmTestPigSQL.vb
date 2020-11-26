Imports PigSQLLib
Public Class frmTestPigSQL
    Dim cnMain As DBConnItem

    Private Sub tsmExit_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub mnuDBCommItemNew_MouseEnter(sender As Object, e As EventArgs) Handles mnuDBCommItemNew.MouseEnter
        cnMain = New DBConnItem("db4_TestUser", "SEOWGDPENGINE\SQLEXPRESS", "")

    End Sub
End Class
