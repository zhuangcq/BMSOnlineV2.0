<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmspread
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmspread))
        Me.grdspread = New AxMSFlexGridLib.AxMSFlexGrid()
        CType(Me.grdspread, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grdspread
        '
        Me.grdspread.Location = New System.Drawing.Point(53, 47)
        Me.grdspread.Name = "grdspread"
        Me.grdspread.OcxState = CType(resources.GetObject("grdspread.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grdspread.Size = New System.Drawing.Size(374, 292)
        Me.grdspread.TabIndex = 0
        '
        'frmspread
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(503, 419)
        Me.Controls.Add(Me.grdspread)
        Me.Name = "frmspread"
        Me.Text = "frmspread"
        CType(Me.grdspread, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grdspread As AxMSFlexGridLib.AxMSFlexGrid
End Class
