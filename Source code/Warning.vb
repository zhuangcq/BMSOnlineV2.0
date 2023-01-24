Public Class Warning

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
      
    End Sub
    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer1.Enabled = True
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Warning_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = False
        MaximizeBox = False
        MinimizeBox = False
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Vmax1 = 10.0#
        Vmin1 = 0.0#
        Vmin2 = 0.0#
        Vmax2 = 30.0#

        Dim a As Integer

        a = 0
        On Error GoTo MyError
        If a = 0 Then
            'Open "C:\BMStrain\scope.dat" For Output As #1
            'Open "scope.dat" For Output As #1
            FileOpen(1, Application.StartupPath + "\scope.dat", OpenMode.Output)
            Print(1, Vmax1, Vmin1, Vmax2, Vmin2)
            FileClose(1)

            'rewrite the value in selec.dat
            'Open "selec.dat" For Output As #5
            FileOpen(5, Application.StartupPath + "\selec.dat", OpenMode.Output)
            Print(5, 2, 2, 2)
            FileClose(5)

        End If
        Timer1.Enabled = False
MyError:
        a = 1
        Resume Next
        Welcome.Show()
        Menubms.Close()
        Me.Close()
    End Sub
End Class