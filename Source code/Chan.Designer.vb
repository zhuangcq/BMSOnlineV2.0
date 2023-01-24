<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Chan
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
        Me.components = New System.ComponentModel.Container()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.Text2 = New System.Windows.Forms.TextBox()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Command3 = New System.Windows.Forms.Button()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Georgia", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(49, 42)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(376, 25)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "The parameters you want to change are"
        '
        'Text1
        '
        Me.Text1.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text1.Location = New System.Drawing.Point(440, 124)
        Me.Text1.Multiline = True
        Me.Text1.Name = "Text1"
        Me.Text1.Size = New System.Drawing.Size(100, 31)
        Me.Text1.TabIndex = 1
        '
        'Text2
        '
        Me.Text2.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text2.Location = New System.Drawing.Point(440, 200)
        Me.Text2.Multiline = True
        Me.Text2.Name = "Text2"
        Me.Text2.Size = New System.Drawing.Size(100, 31)
        Me.Text2.TabIndex = 1
        '
        'Command1
        '
        Me.Command1.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.Location = New System.Drawing.Point(64, 321)
        Me.Command1.Name = "Command1"
        Me.Command1.Size = New System.Drawing.Size(116, 42)
        Me.Command1.TabIndex = 2
        Me.Command1.Text = "Confirm"
        Me.Command1.UseVisualStyleBackColor = True
        '
        'Command3
        '
        Me.Command3.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command3.Location = New System.Drawing.Point(366, 321)
        Me.Command3.Name = "Command3"
        Me.Command3.Size = New System.Drawing.Size(116, 42)
        Me.Command3.TabIndex = 2
        Me.Command3.Text = "Cancel"
        Me.Command3.UseVisualStyleBackColor = True
        '
        'Command2
        '
        Me.Command2.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command2.Location = New System.Drawing.Point(209, 321)
        Me.Command2.Name = "Command2"
        Me.Command2.Size = New System.Drawing.Size(116, 42)
        Me.Command2.TabIndex = 2
        Me.Command2.Text = "Clean"
        Me.Command2.UseVisualStyleBackColor = True
        '
        'Command4
        '
        Me.Command4.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command4.Location = New System.Drawing.Point(515, 321)
        Me.Command4.Name = "Command4"
        Me.Command4.Size = New System.Drawing.Size(116, 42)
        Me.Command4.TabIndex = 2
        Me.Command4.Text = "Restore"
        Me.Command4.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(51, 124)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(370, 31)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Label2"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(50, 200)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(370, 31)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Label3"
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'Chan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(711, 414)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Command2)
        Me.Controls.Add(Me.Command4)
        Me.Controls.Add(Me.Command3)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.Text2)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Chan"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Change parameters"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Text1 As System.Windows.Forms.TextBox
    Friend WithEvents Text2 As System.Windows.Forms.TextBox
    Friend WithEvents Command1 As System.Windows.Forms.Button
    Friend WithEvents Command3 As System.Windows.Forms.Button
    Friend WithEvents Command2 As System.Windows.Forms.Button
    Friend WithEvents Command4 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
