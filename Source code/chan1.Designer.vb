<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class chan1
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Command3 = New System.Windows.Forms.Button()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(92, 153)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(370, 31)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Label2"
        '
        'Command2
        '
        Me.Command2.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command2.Location = New System.Drawing.Point(250, 350)
        Me.Command2.Name = "Command2"
        Me.Command2.Size = New System.Drawing.Size(116, 41)
        Me.Command2.TabIndex = 9
        Me.Command2.Text = "Clean"
        Me.Command2.UseVisualStyleBackColor = True
        '
        'Command4
        '
        Me.Command4.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command4.Location = New System.Drawing.Point(556, 350)
        Me.Command4.Name = "Command4"
        Me.Command4.Size = New System.Drawing.Size(116, 41)
        Me.Command4.TabIndex = 10
        Me.Command4.Text = "Restore"
        Me.Command4.UseVisualStyleBackColor = True
        '
        'Command3
        '
        Me.Command3.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command3.Location = New System.Drawing.Point(407, 350)
        Me.Command3.Name = "Command3"
        Me.Command3.Size = New System.Drawing.Size(116, 41)
        Me.Command3.TabIndex = 8
        Me.Command3.Text = "Cancel"
        Me.Command3.UseVisualStyleBackColor = True
        '
        'Command1
        '
        Me.Command1.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.Location = New System.Drawing.Point(105, 350)
        Me.Command1.Name = "Command1"
        Me.Command1.Size = New System.Drawing.Size(116, 41)
        Me.Command1.TabIndex = 7
        Me.Command1.Text = "Confirm"
        Me.Command1.UseVisualStyleBackColor = True
        '
        'Text1
        '
        Me.Text1.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text1.Location = New System.Drawing.Point(481, 153)
        Me.Text1.Multiline = True
        Me.Text1.Name = "Text1"
        Me.Text1.Size = New System.Drawing.Size(100, 31)
        Me.Text1.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Georgia", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(90, 71)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(361, 25)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "The parameters you want to change is"
        '
        'chan1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(762, 450)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Command2)
        Me.Controls.Add(Me.Command4)
        Me.Controls.Add(Me.Command3)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "chan1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Change parameters"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Command2 As System.Windows.Forms.Button
    Friend WithEvents Command4 As System.Windows.Forms.Button
    Friend WithEvents Command3 As System.Windows.Forms.Button
    Friend WithEvents Command1 As System.Windows.Forms.Button
    Friend WithEvents Text1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
