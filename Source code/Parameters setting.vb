Imports System.IO.Directory
Public Class Parameters_setting
    Dim Ac As Integer
    Dim CurPath As String
    Dim i, j As Integer

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Parameters_setting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = False

        'Me.ControlBox = False
        '    MaximizeBox = False
        ' MinimizeBox = True
        '  Me.Height = My.Computer.Screen.WorkingArea.Height
        '   Me.Width = My.Computer.Screen.WorkingArea.Width
        '   Me.WindowState = FormWindowState.Maximized
        '  Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None

        'Label1.Parent = PictureBox1
        'Label2.Parent = PictureBox1
        ' TextBox1.Parent = PictureBox1
        ' HScrollBar1.Parent = PictureBox1
        '  GroupBox1.Parent = PictureBox2
        ' GroupBox3.Parent = PictureBox2
        ' GroupBox2.Parent = PictureBox2
        '   Label3.BackColor = Color.FromArgb(100, 0, 0, 0)
        TextBox1.Text = Str$(10)
        HScrollBar1.Value = 10
        RadioButton1.Checked = True
        RadioButton4.Checked = True
        RadioButton8.Checked = True

        Dim a_width, b_height, c_font, c_font1, c_font2 As Integer
        a_width = 965
        b_height = 653
        c_font1 = CInt(Me.Height / b_height * 12)
        c_font2 = CInt(Me.Width / a_width * 12)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If
        Button1.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Button2.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Button3.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Button4.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Ac = Val(TextBox1.Text)

        If Ac >= 1 And Ac <= 20 Then
            HScrollBar1.Value = Str$(TextBox1.Text)
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '初始化输入参数
        Dim SourceFile3 As String = GetCurrentDirectory() + "\bmsinbak.dat"
        Dim DestinationFile3 As String = GetCurrentDirectory() + "\bmsin.dat"
        FileCopy(SourceFile3, DestinationFile3)

        Ac = Val(TextBox1.Text)
        If Ac < 1 Or Ac > 20 Then
            'Load(warn5)
            warn5.Show()
        Else
            FileOpen(2, GetCurrentDirectory() + "\acceler.dat", OpenMode.Output)
            'Open "C:\BMStrain\acceler.dat" For Output As #2
            'Open "acceler.dat" For Output As #2
            Print(2, Ac)
            FileClose(2)
            'Unload(acceler)
            If RadioButton1.Checked = True And RadioButton4.Checked = True And RadioButton8.Checked = True Then
                Z = 1
                W = 1
                period = 2.0#
            ElseIf RadioButton1.Checked = True And RadioButton4.Checked = True And RadioButton9.Checked = True Then
                Z = 2
                W = 2
                period = 4.0#
            ElseIf RadioButton1.Checked = True And RadioButton4.Checked = True And RadioButton10.Checked = True Then
                Z = 3
                W = 3
                period = 8.0#
            ElseIf RadioButton1.Checked = True And RadioButton4.Checked = True And RadioButton11.Checked = True Then
                Z = 4
                W = 4
                period = 12.0#
            ElseIf RadioButton1.Checked = True And RadioButton5.Checked = True And RadioButton8.Checked = True Then
                Z = 5
                W = 1
                period = 2.0#
            ElseIf RadioButton1.Checked = True And RadioButton5.Checked = True And RadioButton9.Checked = True Then
                Z = 6
                W = 2
                period = 4.0#
            ElseIf RadioButton1.Checked = True And RadioButton5.Checked = True And RadioButton10.Checked = True Then
                Z = 7
                W = 3
                period = 8.0#
            ElseIf RadioButton1.Checked = True And RadioButton5.Checked = True And RadioButton11.Checked = True Then
                Z = 8
                W = 4
                period = 12.0#
            ElseIf RadioButton2.Checked = True And RadioButton4.Checked = True And RadioButton8.Checked = True Then
                Z = 9
                W = 1
                period = 2.0#
            ElseIf RadioButton2.Checked = True And RadioButton4.Checked = True And RadioButton9.Checked = True Then
                Z = 10
                W = 2
                period = 4.0#
            ElseIf RadioButton2.Checked = True And RadioButton4.Checked = True And RadioButton10.Checked = True Then
                Z = 11
                W = 3
                period = 8.0#
            ElseIf RadioButton2.Checked = True And RadioButton4.Checked = True And RadioButton11.Checked = True Then
                Z = 12
                W = 4
                period = 12.0#
            ElseIf RadioButton2.Checked = True And RadioButton5.Checked = True And RadioButton8.Checked = True Then
                Z = 13
                W = 1
                period = 2.0#
            ElseIf RadioButton2.Checked = True And RadioButton5.Checked = True And RadioButton9.Checked = True Then
                Z = 14
                W = 2
                period = 4.0#
            ElseIf RadioButton2.Checked = True And RadioButton5.Checked = True And RadioButton10.Checked = True Then
                Z = 15
                W = 3
                period = 8.0#
            ElseIf RadioButton2.Checked = True And RadioButton5.Checked = True And RadioButton11.Checked = True Then
                Z = 16
                W = 4
                period = 12.0#
            ElseIf RadioButton3.Checked = True And RadioButton4.Checked = True And RadioButton8.Checked = True Then
                Z = 17
                W = 1
                period = 2.0#
            ElseIf RadioButton3.Checked = True And RadioButton4.Checked = True And RadioButton9.Checked = True Then
                Z = 18
                W = 2
                period = 4.0#
            ElseIf RadioButton3.Checked = True And RadioButton4.Checked = True And RadioButton10.Checked = True Then
                Z = 19
                W = 3
                period = 8.0#
            ElseIf RadioButton3.Checked = True And RadioButton4.Checked = True And RadioButton11.Checked = True Then
                Z = 20
                W = 4
                period = 12.0#
            ElseIf RadioButton3.Checked = True And RadioButton5.Checked = True And RadioButton8.Checked = True Then
                Z = 21
                W = 1
                period = 2.0#
            ElseIf RadioButton3.Checked = True And RadioButton5.Checked = True And RadioButton9.Checked = True Then
                Z = 22
                W = 2
                period = 4.0#
            ElseIf RadioButton3.Checked = True And RadioButton5.Checked = True And RadioButton10.Checked = True Then
                Z = 23
                W = 3
                period = 8.0#
            ElseIf RadioButton3.Checked = True And RadioButton5.Checked = True And RadioButton11.Checked = True Then
                Z = 24
                W = 4
                period = 12.0#
            Else
            End If

            Menubms.Button2.Visible = True
            Menubms.Show()
            Me.Close()

        End If

    End Sub

    Private Sub HScrollBar1_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar1.Scroll
        TextBox1.Text = Str$(HScrollBar1.Value)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
        Menubms.Show()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Text = 10
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Text = " "
        HScrollBar1.Value = 1

    End Sub
    Private Sub RadioButton5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton5.Click
        If RadioButton5.Checked = True Then
            RadioButton4.Checked = False

        End If
    End Sub

    Private Sub RadioButton4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton4.Click
        If RadioButton4.Checked = True Then
            RadioButton5.Checked = False

        End If
    End Sub

    Private Sub RadioButton3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton3.Click
        If RadioButton3.Checked = True Then
            RadioButton1.Checked = False
            RadioButton2.Checked = False
        End If
    End Sub

    Private Sub RadioButton2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton2.Click
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            RadioButton3.Checked = False
        End If
    End Sub

    Private Sub RadioButton11_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton11.Click
        If RadioButton11.Checked = True Then
            RadioButton9.Checked = False
            RadioButton8.Checked = False
            RadioButton10.Checked = False
        End If
    End Sub

    Private Sub RadioButton10_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton10.Click
        If RadioButton10.Checked = True Then
            RadioButton9.Checked = False
            RadioButton8.Checked = False
            RadioButton11.Checked = False
        End If
    End Sub

    Private Sub RadioButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.Click
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            RadioButton3.Checked = False
        End If
    End Sub

    Private Sub RadioButton8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton8.CheckedChanged

    End Sub

    Private Sub RadioButton8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton8.Click
        If RadioButton8.Checked = True Then
            RadioButton9.Checked = False
            RadioButton10.Checked = False
            RadioButton11.Checked = False
        End If
    End Sub

    Private Sub RadioButton9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton9.CheckedChanged

    End Sub

    Private Sub RadioButton9_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton9.Click
        If RadioButton9.Checked = True Then
            RadioButton8.Checked = False
            RadioButton10.Checked = False
            RadioButton11.Checked = False
        End If
    End Sub

    Private Sub Label3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
End Class