Public Class chan1
    Dim out(37) As Single
    Private Sub Command1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command1.Click
        GA = Val(Text1.Text)

        ' Unload(chan1)
        Me.Close()

    End Sub

    Private Sub Command2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command2.Click
        Text1.Text = " "
    End Sub

    Private Sub Command3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command3.Click
        Me.Close()
    End Sub

    Private Sub Command4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command4.Click
        ' Open "bmsinbak.dat" For Input As #2
        FileOpen(2, Application.StartupPath + "\bmsinbak.dat", OpenMode.Input)
        For j = 1 To 37
            Input(2, out(j))
        Next j
        FileClose(2)

        Text1.Text = out(M)

    End Sub

    Private Sub Text1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text1.TextChanged
        Dim a As Single
        Dim Msg As String

        On Error GoTo MyError

        a = Text1.Text
        GoTo 111

MyError:
        Beep()
        Text1.Text = " "
        Resume Next

111:    Msg = " "

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub chan1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MaximizeBox = False
        MinimizeBox = False
        Me.TopMost = False
    End Sub
End Class