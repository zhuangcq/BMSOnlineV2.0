Imports System.IO.Directory
Public Class chan7
    Dim out(37) As Single
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

    Private Sub Command1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command1.Click
        Timer1.Enabled = True
    End Sub

    Private Sub Command2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command2.Click
        Text1.Text = " "
        Text2.Text = " "
        Text3.Text = " "
        Text4.Text = " "
        Text5.Text = " "
        Text6.Text = " "
        Text7.Text = " "

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
        Text2.Text = out(n)
        Text3.Text = out(m1)
        Text4.Text = out(m2)
        Text5.Text = out(m3)
        Text6.Text = out(m4)
        Text7.Text = out(m5)

    End Sub

    Private Sub Text2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text2.TextChanged
        Dim a As Single
        Dim Msg As String

        On Error GoTo MyError

        a = Text2.Text
        GoTo 111

MyError:
        Beep()
        Text2.Text = " "
        Resume Next

111:    Msg = " "

    End Sub

    Private Sub Text3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text3.TextChanged
        Dim a As Single
        Dim Msg As String

        On Error GoTo MyError

        a = Text3.Text
        GoTo 111

MyError:
        Beep()
        Text3.Text = " "
        Resume Next

111:    Msg = " "

    End Sub

    Private Sub Text4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text4.TextChanged
        Dim a As Single
        Dim Msg As String

        On Error GoTo MyError

        a = Text4.Text
        GoTo 111

MyError:
        Beep()
        Text4.Text = " "
        Resume Next

111:    Msg = " "

    End Sub

    Private Sub Text5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text5.TextChanged
        Dim a As Single
        Dim Msg As String

        On Error GoTo MyError

        a = Text5.Text
        GoTo 111

MyError:
        Beep()
        Text5.Text = " "
        Resume Next

111:    Msg = " "

    End Sub

    Private Sub Text6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text6.TextChanged
        Dim a As Single
        Dim Msg As String

        On Error GoTo MyError

        a = Text6.Text
        GoTo 111

MyError:
        Beep()
        Text6.Text = " "
        Resume Next

111:    Msg = " "

    End Sub

    Private Sub Text7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text7.TextChanged
        Dim a As Single
        Dim Msg As String

        On Error GoTo MyError

        a = Text7.Text
        GoTo 111

MyError:
        Beep()
        Text7.Text = " "
        Resume Next

111:    Msg = " "

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim a As Integer

        a = 0
        On Error GoTo MyError
        If a = 0 Then
            ' Open "bmsin.dat" For Input As #2
            '  Dim SourceFile3 As String = GetCurrentDirectory() + "\bmsincopy.dat"
            '   Dim DestinationFile3 As String = GetCurrentDirectory() + "\bmsin.dat"

            FileOpen(2, Application.StartupPath + "\bmsin.dat", OpenMode.Input)
            For i = 1 To 37
                Input(2, out(i))
            Next i
            FileClose(2)
            'Shell("cmd /c copy bmsincopy.dat bmsin.dat")
            'FileCopy(SourceFile3, DestinationFile3)
            out(M) = Val(Text1.Text)
            out(n) = Val(Text2.Text)
            out(m1) = Val(Text3.Text)
            out(m2) = Val(Text4.Text)
            out(m3) = Val(Text5.Text)
            out(m4) = Val(Text6.Text)
            out(m5) = Val(Text7.Text)

            ' Open "bmsin.dat" For Output As #1
            ' Dim SourceFile1 As String = GetCurrentDirectory() + "\bmsin.dat"
            '  Dim DestinationFile1 As String = GetCurrentDirectory() + "\bmsincopy.dat"
            '  FileCopy(SourceFile1, DestinationFile1)
            'Shell("cmd /c copy bmsin.dat bmsincopy.dat")
            FileOpen(1, Application.StartupPath + "\bmsin.dat", OpenMode.Output)
            For j = 1 To 37
                Print(1, out(j) & vbCrLf)
            Next j
            FileClose(1)
            Timer1.Enabled = False
            Me.Close()
        Else
        End If

MyError:
        a = 1
        Resume Next

    End Sub

    Private Sub chan7_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MaximizeBox = False
        MinimizeBox = False
        Me.TopMost = False
    End Sub
End Class