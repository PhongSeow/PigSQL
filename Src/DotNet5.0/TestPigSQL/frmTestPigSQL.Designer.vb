<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmTestPigSQL
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.msMain = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDBCommItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDBCommItemNew = New System.Windows.Forms.ToolStripMenuItem()
        Me.tbMain = New System.Windows.Forms.TextBox()
        Me.信任连接 = New System.Windows.Forms.ToolStripMenuItem()
        Me.msMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'msMain
        '
        Me.msMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuDBCommItem})
        Me.msMain.Location = New System.Drawing.Point(0, 0)
        Me.msMain.Name = "msMain"
        Me.msMain.Size = New System.Drawing.Size(800, 25)
        Me.msMain.TabIndex = 0
        Me.msMain.Text = "MenuStrip1"
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(39, 21)
        Me.mnuFile.Text = "&File"
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Name = "mnuFileExit"
        Me.mnuFileExit.Size = New System.Drawing.Size(96, 22)
        Me.mnuFileExit.Text = "&Exit"
        '
        'mnuDBCommItem
        '
        Me.mnuDBCommItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuDBCommItemNew})
        Me.mnuDBCommItem.Name = "mnuDBCommItem"
        Me.mnuDBCommItem.Size = New System.Drawing.Size(101, 21)
        Me.mnuDBCommItem.Text = "&DBCommItem"
        '
        'mnuDBCommItemNew
        '
        Me.mnuDBCommItemNew.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.信任连接})
        Me.mnuDBCommItemNew.Name = "mnuDBCommItemNew"
        Me.mnuDBCommItemNew.Size = New System.Drawing.Size(102, 22)
        Me.mnuDBCommItemNew.Text = "New"
        '
        'tbMain
        '
        Me.tbMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbMain.Location = New System.Drawing.Point(0, 25)
        Me.tbMain.Multiline = True
        Me.tbMain.Name = "tbMain"
        Me.tbMain.Size = New System.Drawing.Size(800, 425)
        Me.tbMain.TabIndex = 1
        '
        '信任连接
        '
        Me.信任连接.Name = "信任连接"
        Me.信任连接.Size = New System.Drawing.Size(220, 22)
        Me.信任连接.Text = "mnuDBCommItemNewTr"
        '
        'frmTestPigSQL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.tbMain)
        Me.Controls.Add(Me.msMain)
        Me.MainMenuStrip = Me.msMain
        Me.Name = "frmTestPigSQL"
        Me.Text = "Test PigSQL"
        Me.msMain.ResumeLayout(False)
        Me.msMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents msMain As MenuStrip
    Friend WithEvents mnuFile As ToolStripMenuItem
    Friend WithEvents tsmExit As ToolStripMenuItem
    Friend WithEvents tbMain As TextBox
    Friend WithEvents mnuFileExit As ToolStripMenuItem
    Friend WithEvents mnuDBCommItem As ToolStripMenuItem
    Friend WithEvents mnuDBCommItemNew As ToolStripMenuItem
    Friend WithEvents 信任连接 As ToolStripMenuItem
End Class
