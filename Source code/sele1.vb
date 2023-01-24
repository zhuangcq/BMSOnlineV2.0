Public Class sele1
    Dim out(37) As Single
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Timer1.Enabled = True

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            RadioButton3.Checked = False
            RadioButton4.Checked = False
            chan1.Close()
            chan1.Show()
            chan1.Label2.text = "Manual setpoint of fresh air(kg/s)"

            M = 28
            'On Error GoTo MyError

            ' Open "BMSin.dat" For Input As #1
            FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
            For i = 1 To 37
                Input(1, out(i))
            Next i
            FileClose(1)
            'MyError:
            '     Resume

            chan1.Text1.Text = Str$(out(28))

        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            RadioButton3.Checked = False
            RadioButton4.Checked = False
        End If

    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            RadioButton4.Checked = False
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            RadioButton1.Checked = False
            RadioButton3.Checked = False
            RadioButton2.Checked = False
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim a As Integer

        a = 0
        On Error GoTo MyError
        If a = 0 Then
            If (RadioButton1.Checked = True) Then
                'Open "BMSin.dat" For Input As #1
                FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
                For i = 1 To 37
                    Input(1, out(i))
                Next i
                FileClose(1)

                out(28) = GA

                ' Open "BMSin.dat" For Output As #1
                FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Output)
                For i = 1 To 37
                    Print(1, out(i) & vbCrLf)
                Next i
                FileClose(1)

                ' Open "selec.dat" For Input As #3
                FileOpen(3, Application.StartupPath + "\selec.dat", OpenMode.Input)
                For j = 0 To 2
                    Input(3, Mb(j))
                Next j
                FileClose(3)

                Mb(1) = 1

                'Open "selec.dat" For Output As #2
                FileOpen(2, Application.StartupPath + "\selec.dat", OpenMode.Output)
                For j = 0 To 2
                    Print(2, Mb(j))
                Next j
                FileClose(2)
                Timer1.Enabled = False
                Me.Close()
            ElseIf (RadioButton2.Checked = True) Then

                ' Open "selec.dat" For Input As #3
                '       For j = 0 To 2
                ' Input #3, Mb(j)
                '       Next j
                '  Close #3
                FileOpen(3, Application.StartupPath + "\selec.dat", OpenMode.Input)
                For j = 0 To 2
                    Input(3, Mb(j))
                Next j
                FileClose(3)
                Mb(1) = 3

                ' Open "selec.dat" For Output As #2
                FileOpen(2, Application.StartupPath + "\selec.dat", OpenMode.Output)
                For j = 0 To 2
                    Print(2, Mb(j))
                Next j
                FileClose(2)
                Timer1.Enabled = False
                Me.Close()
            ElseIf (RadioButton3.Checked = True) Then
                FileOpen(3, Application.StartupPath + "\selec.dat", OpenMode.Input)
                For j = 0 To 2
                    Input(3, Mb(j))
                Next j
                FileClose(3)

                Mb(1) = 4

                ' Open "selec.dat" For Output As #2
                FileOpen(2, Application.StartupPath + "\selec.dat", OpenMode.Output)
                For j = 0 To 2
                    Print(2, Mb(j))
                Next j
                FileClose(2)
                Timer1.Enabled = False
                Me.Close()


            ElseIf (RadioButton4.Checked = True) Then
                'Open "selec.dat" For Input As #3
                FileOpen(3, Application.StartupPath + "\selec.dat", OpenMode.Input)
                For j = 0 To 2
                    Input(3, Mb(j))
                Next j
                FileClose(3)

                Mb(1) = 2

                '  Open "selec.dat" For Output As #2
                FileOpen(2, Application.StartupPath + "\selec.dat", OpenMode.Output)
                For j = 0 To 2
                    Print(2, Mb(j))
                Next j
                FileClose(2)
                Timer1.Enabled = False
                Me.Close()
            End If
        Else
        End If


MyError:
        a = 1
        Resume Next

    End Sub

    Private Sub sele1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = False
        MaximizeBox = False
        MinimizeBox = True
    End Sub
End Class