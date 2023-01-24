Imports System.IO.Directory
Public Class warn6
    Dim Sta, Fin, DT As Single
    Dim Ru As Integer
    Public Sub killprocess(ByRef strprocessTokill As String)
        Dim proc() As Process = Process.GetProcesses
        For i As Integer = 0 To proc.GetUpperBound(0)
            If proc(i).ProcessName = strprocessTokill Then
                proc(i).Kill()
            End If

        Next

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        killprocess(Label31.Text)
        Timer1.Enabled = True

        Menubms.Button4.Visible = True
        Menubms.Button5.Visible = True
        For i1 As Integer = 0 To 399
            For j1 As Integer = 0 To 177
                frmspread.grdspread.Col = i1
                frmspread.grdspread.Row = j1
                frmspread.grdspread.Text = Nothing
                'frmspread.Close()

            Next j1
        Next i1
        curve.Close()
        curves.Close()
        flow1.Close()
        flow2.Close()
        '  Control.Close()
        'Menubms.Button4.Visible = True
        'Menubms.Button5.Visible = True

        flow1.Timer1.Enabled = True
        flow1.Label4.Visible = False
        Ra1 = 1
        Ra = 1
        Select Case Z
            Case 1
                iTask = Shell("doit3.bat", 6)
                '   flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 2
                iTask = Shell("doit4.bat", 6)
                'Load(flow1)
                '   flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 3
                iTask = Shell("doit5.bat", 6)
                ' Load(flow1)
                '    flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 4
                iTask = Shell("doit6.bat", 6)
                '  Load(flow1)
                '   flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 5
                iTask = Shell("doit9.bat", 6)
                '   Load(flow1)
                '    flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 6
                iTask = Shell("doit10.bat", 6)
                '  Load(flow1)
                '   flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 7
                iTask = Shell("doit11.bat", 6)
                '  Load(flow1)
                '    flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 8
                iTask = Shell("doit12.bat", 6)
                '  Load(flow1)
                '     flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 9
                iTask = Shell("doit15.bat", 6)
                '  Load(flow1)
                '  flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 10
                iTask = Shell("doit16.bat", 6)
                '  Load(flow1)
                '  flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 11
                iTask = Shell("doit17.bat", 6)
                '  Load(flow1)
                '    flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 12
                iTask = Shell("doit18.bat", 6)
                '  Load(flow1)
                '    flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 13
                iTask = Shell("doit21.bat", 6)
                '  Load(flow1)
                '   flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 14
                iTask = Shell("doit22.bat", 6)
                '  Load(flow1)
                '    flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 15
                iTask = Shell("doit23.bat", 6)
                '  Load(flow1)
                '    flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 16
                iTask = Shell("doit24.bat", 6)
                '  Load(flow1)
                '      flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 17
                iTask = Shell("doit27.bat", 6)
                ' Load(flow1)
                '   flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 18
                iTask = Shell("doit28.bat", 6)
                '  Load(flow1)
                '   flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 19
                iTask = Shell("doit29.bat", 6)
                '   Load(flow1)
                '  flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 20
                iTask = Shell("doit30.bat", 6)
                ' Load(flow1)
                '  flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()
            Case 21
                iTask = Shell("doit33.bat", 6)
                '  Load(flow1)
                '   flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 22
                iTask = Shell("doit34.bat", 6)
                ' Load(flow1)
                ' flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()

            Case 23
                iTask = Shell("doit35.bat", 6)
                'Load(flow1)
                ' flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()
            Case 24
                iTask = Shell("doit36.bat", 6)
                'Load(flow1)
                '  flow2.Show()
                flow1.Show()
                Menubms.Hide()
                Me.Close()
            Case Else
                'Load(warn1)
                warn1.Close()
                warn1.Show()
        End Select

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Vmax1 = 10.0#
        Vmin1 = 0.0#
        Vmax2 = 30.0#
        Vmin2 = 0.0#

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

            'copy the default value into BMSin.dat
            'SourceFile = "Bmsinbak.dat"
            'DestinationFile = "Bmsin.dat"
            Dim SourceFile As String = GetCurrentDirectory() + "\Bmsinbak.dat"
            Dim DestinationFile As String = GetCurrentDirectory() + "\Bmsin.dat"
            FileCopy(SourceFile, DestinationFile)
            'Shell("cmd /c copy Bmsinbak.dat bmsin.dat")
            ' read the RUNNING.DAT, see whether the Fortran program is over or not
            ' Open "RUNNING.DAT" For Input As #9
            FileOpen(9, Application.StartupPath + "\RUNNING.DAT", OpenMode.Output)
            Input(9, Ru)
            FileClose(9)

            If (Ru > 1) Then
                hact = OpenProcess(SYNCHRONIZE, False, iTask)
                retval = WaitForSingleObject(hact, 500)
                Select Case Z
                    Case 1
                        AppActivate("DOIT3")
                    Case 2
                        AppActivate("DOIT4")
                    Case 3
                        AppActivate("DOIT5")
                    Case 4
                        AppActivate("DOIT6")
                    Case 5
                        AppActivate("DOIT9")
                    Case 6
                        AppActivate("DOIT10")
                    Case 7
                        AppActivate("DOIT11")
                    Case 8
                        AppActivate("DOIT12")
                    Case 9
                        AppActivate("DOIT15")
                    Case 10
                        AppActivate("DOIT16")
                    Case 11
                        AppActivate("DOIT17")
                    Case 12
                        AppActivate("DOIT18")
                    Case 13
                        AppActivate("DOIT21")
                    Case 14
                        AppActivate("DOIT22")
                    Case 15
                        AppActivate("DOIT23")
                    Case 16
                        AppActivate("DOIT24")
                    Case 17
                        AppActivate("DOIT27")
                    Case 18
                        AppActivate("DOIT28")
                    Case 19
                        AppActivate("DOIT29")
                    Case 20
                        AppActivate("DOIT30")
                    Case 21
                        AppActivate("DOIT33")
                    Case 22
                        AppActivate("DOIT34")
                    Case 23
                        AppActivate("DOIT35")
                    Case 24
                        AppActivate("DOIT36")
                    Case Else
                End Select
                retval = TerminateProcess(hact, 0)
                retval = CloseHandle(hact)

                Timer1.Enabled = False

                ' Unload(warn)
                Me.Hide()
                ' Unload(flow1)
                '  Unload(flow2)
                flow1.Close()
                flow2.Close()
                '    Load(Welcome)
                Welcome.Show()

            Else
                Timer1.Enabled = False

                ' Unload(warn)
                Me.Hide()
                ' Unload(flow1)
                '  Unload(flow2)
                flow1.Close()
                flow2.Close()

                'Load(Welcome)
                Welcome.Show()
                ' End
            End If

        Else
        End If

MyError:
        a = 1
        Resume Next
    End Sub


    Private Sub warn6_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = False
        MaximizeBox = False
        MinimizeBox = False
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub
End Class