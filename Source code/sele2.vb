Public Class sele2
    Dim M7, N17, L7 As Integer
    Dim out(37) As Single

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer1.Enabled = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Mb(1) = 2
        'Mb(2) = 2
        'Mb(3) = 2

        'Open "selec.dat" For Output As #1
        'For i = 1 To 3
        'Print #1, Mb(i)
        'Next i
        'Close #1

        Me.Close()


    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            chan1.Close()
            chan1.Show()
            chan1.Label2.Text = "Manual setpoint of Static pressure(pa)"

            M = 30

            'On Error GoTo MyError

            'Open "BMSin.dat" For Input As #1
            FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
            For i = 1 To 37
                Input(1, out(i))
            Next i
            FileClose(1)

            'MyError:
            '    Resume

            chan1.Text1.Text = Str$(out(30))

        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
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

                out(30) = GA

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

                Mb(2) = 1

                ' Open "selec.dat" For Output As #2
                FileOpen(2, Application.StartupPath + "\selec.dat", OpenMode.Output)
                For j = 0 To 2
                    Print(2, Mb(j))
                Next j
                FileClose(2)
                Timer1.Enabled = False
                Me.Close()


            ElseIf (RadioButton2.Checked = True) Then

                'Open "selec.dat" For Input As #3
                FileOpen(3, Application.StartupPath + "\selec.dat", OpenMode.Input)
                For j = 0 To 2
                    Input(3, Mb(j))
                Next j
                FileClose(3)

                Mb(2) = 2

                'Open "selec.dat" For Output As #2
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

    Private Sub sele2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = False
        MaximizeBox = False
        MinimizeBox = True
    End Sub
End Class