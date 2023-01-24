Public Class Welcome
    Dim CurPath As String
    Dim i, j As Integer
    'Dim hactive As Long  ' handle to the active window
    'Dim retval As Long  ' return value
    Public Sub killprocess(ByRef strprocessTokill As String)
        Dim proc() As Process = Process.GetProcesses
        For i As Integer = 0 To proc.GetUpperBound(0)
            If proc(i).ProcessName = strprocessTokill Then
                proc(i).Kill()
            End If

        Next

    End Sub
    Private Sub Welcome_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = False
        MaximizeBox = False
        Me.Height = My.Computer.Screen.WorkingArea.Height
        Me.Width = My.Computer.Screen.WorkingArea.Width
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.PictureBox1.Height = My.Computer.Screen.Bounds.Height
        Me.PictureBox1.Width = My.Computer.Screen.Bounds.Width

        Button1.Parent = PictureBox1
        Button2.Parent = PictureBox1
     
        Me.Button1.Width = My.Computer.Screen.Bounds.Width * (100 / 1139)
        Me.Button1.Height = My.Computer.Screen.Bounds.Height * (43 / 724)
        Me.Button2.Width = Me.Button1.Width
        Me.Button2.Height = Me.Button1.Height
        Me.Button1.Location = New Point(My.Computer.Screen.Bounds.Width - Me.Button2.Width * 2.8, My.Computer.Screen.Bounds.Height - Me.Button2.Height * 1.8)
        Me.Button2.Location = New Point(My.Computer.Screen.Bounds.Width - Me.Button2.Width * 1.5, My.Computer.Screen.Bounds.Height - Me.Button2.Height * 1.8)
        Dim c_font, b_height As Single
        b_height = 701
        c_font = CInt(Me.Height / b_height * 12)
        Button1.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Button2.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Menubms.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        killprocess(Label31.Text)
        End
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub
End Class