Public Class sele
    'Dim Mb(3) As Integer
    Dim out(37) As Single

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            chan1.Close()
            chan1.Show()
        End If
        chan1.Label2.Text = "Manual supply air temperature setpoint (°C)"

        M = 31

        'On Error GoTo MyError

        ' Open "BMSin.dat" For Input As #1
        FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(1, out(i))
        Next i
        FileClose(1)

        'MyError:
        '     Resume

        chan1.Text1.Text = Str$(out(31))

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
        End If
    End Sub

    Private Sub sele_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = False
        MaximizeBox = False
        MinimizeBox = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer1.Enabled = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()

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

                out(31) = GA

                'Open "BMSin.dat" For Output As #1
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

                Mb(0) = 1

                '  Open "selec.dat" For Output As #3
                FileOpen(3, Application.StartupPath + "\selec.dat", OpenMode.Output)
                For j = 0 To 2
                    Print(3, Mb(j))
                Next j
                FileClose(3)
                Timer1.Enabled = False
                Me.Close()

            ElseIf (RadioButton2.Checked = True) Then

                ' Open "selec.dat" For Input As #2
                FileOpen(2, Application.StartupPath + "\selec.dat", OpenMode.Input)
                For j = 0 To 2
                    Input(2, Mb(j))
                Next j
                FileClose(2)

                Mb(0) = 2

                ' Open "selec.dat" For Output As #3
                FileOpen(3, Application.StartupPath + "\selec.dat", OpenMode.Output)
                For j = 0 To 2
                    Print(3, Mb(j))
                Next j
                FileClose(3)
            End If
            Timer1.Enabled = False
            Me.Close()
        Else
        End If

MyError:
        a = 1
        Resume Next

    End Sub
End Class