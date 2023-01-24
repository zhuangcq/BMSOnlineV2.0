Imports System.IO.Directory
Public Class Chan
    Dim out(37) As Single

    Private Sub Command1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command1.Click
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim a As Integer

        a = 0
        On Error GoTo MyError
        If a = 0 Then
            '  Dim SourceFile3 As String = GetCurrentDirectory() + "\bmsincopy.dat"
            '   Dim DestinationFile3 As String = GetCurrentDirectory() + "\bmsin.dat"

            FileOpen(2, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
            'Open "bmsin.dat" For Input As #2
            For i = 1 To 37
                Input(2, out(i))
            Next i
            FileClose(2)
            '  Shell("cmd /c copy bmsincopy.dat bmsin.dat")
            '  FileCopy(SourceFile3, DestinationFile3)
            out(M) = Val(Text1.Text)
            out(n) = Val(Text2.Text)
            '  Dim SourceFile1 As String = GetCurrentDirectory() + "\bmsin.dat"
            '  Dim DestinationFile1 As String = GetCurrentDirectory() + "\bmsincopy.dat"
            ' FileCopy(SourceFile1, DestinationFile1)
            ' Shell("cmd /c copy bmsin.dat bmsincopy.dat")
            FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Output)
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

    Private Sub Command2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command2.Click
        Text1.Text = " "
        Text2.Text = " "

    End Sub

    Private Sub Command3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command3.Click
        Me.Hide()
    End Sub

    Private Sub Command4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command4.Click
        'Open "bmsinbak.dat" For Input As #2
        FileOpen(2, Application.StartupPath + "\bmsinbak.dat", OpenMode.Input)
        For j = 1 To 37
            Input(2, out(j))
        Next j
        FileClose(2)

        Text1.Text = out(M)
        Text2.Text = out(n)

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

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Chan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MaximizeBox = False
        MinimizeBox = False
        Me.TopMost = False
    End Sub
End Class