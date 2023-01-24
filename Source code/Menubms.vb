Imports System.IO.Directory
Public Class Menubms
    Dim Re, ReT, ReT1, ReT2 As Single
    Dim TheFileNameb As String
    Dim TheNewPathb As String

    Private Sub Panel4_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub Panel3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = False
        MaximizeBox = False
        Me.Height = My.Computer.Screen.WorkingArea.Height
        Me.Width = My.Computer.Screen.WorkingArea.Width
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        '  MenuStrip1.Height = 62
        Me.PictureBox1.Height = Me.Height * 0.88
        Me.PictureBox1.Width = Me.Width
        Me.PictureBox1.Location = New Point(0, MenuStrip1.Height)

        ' Me.Button1.Width = Me.Button2.Width = Me.Button3.Width = Me.Button4.Width = Me.Button5.Width = My.Computer.Screen.Bounds.Width * (130 / 1139)
        '  Me.Button1.Height = Me.Button2.Height = Me.Button3.Height = Me.Button4.Height = Me.Button5.Height = My.Computer.Screen.Bounds.Height * (43 / 724)
        Dim x1, y1, x2, y2, a_width, b_height, c_font, c_font1, c_font2 As Integer

        '窗口大小 a为宽，b为高
        a_width = 1149
        b_height = 734
        x1 = 36
        y1 = 650
        x2 = 160
        y2 = 45

        Me.Button1.Width = x2 / a_width * Me.Width
        Me.Button1.Height = y2 / b_height * Me.Height
        ' Me.Button1.Location = New Point(Me.Width * 0.03, y1 / b_height * Me.Height)
        c_font1 = CInt(Me.Height / b_height * 10)
        c_font2 = CInt(Me.Width / a_width * 10)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If
        Button1.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Button2.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Button3.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Button4.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Button5.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Me.Button2.Width = Me.Button1.Width
        Me.Button2.Height = Me.Button1.Height
        Me.Button3.Width = Me.Button1.Width
        Me.Button3.Height = Me.Button1.Height
        Me.Button4.Width = Me.Button1.Width
        Me.Button4.Height = Me.Button1.Height
        Me.Button5.Width = Me.Button1.Width * 1.1
        Me.Button5.Height = Me.Button1.Height

        '  Me.Button2.Location = New Point(Me.Width * 0.03 + (Me.Width + Button1.Width) / 6, y1 / b_height * Me.Height)
        '   Me.Button4.Location = New Point(Me.Width * 0.03 + (Me.Width + Button1.Width) / 6 * 2, y1 / b_height * Me.Height)
        '  Me.Button5.Location = New Point(Me.Width * 0.03 + (Me.Width + Button1.Width) / 6 * 3, y1 / b_height * Me.Height)
        '  Button3.Location = New Point(Me.Width * 0.03 + (Me.Width + Button1.Width * 1.1) / 6 * 4, y1 / b_height * Me.Height)

        Button1.Location = New Point((My.Computer.Screen.Bounds.Width - Me.Button1.Width * 5) / 6, My.Computer.Screen.Bounds.Height * 0.905)
        Button2.Location = New Point((My.Computer.Screen.Bounds.Width - Me.Button1.Width * 5) / 6 * 2 + Me.Button1.Width, My.Computer.Screen.Bounds.Height * 0.905)
        Button4.Location = New Point((My.Computer.Screen.Bounds.Width - Me.Button1.Width * 5) / 6 * 3 + Me.Button1.Width * 2, My.Computer.Screen.Bounds.Height * 0.905)
        Button5.Location = New Point((My.Computer.Screen.Bounds.Width - Me.Button1.Width * 5) / 6 * 4 + Me.Button1.Width * 3, My.Computer.Screen.Bounds.Height * 0.905)
        Button3.Location = New Point((My.Computer.Screen.Bounds.Width - Me.Button1.Width * 5) / 6 * 5 + Me.Button1.Width * 4.1, My.Computer.Screen.Bounds.Height * 0.905)

        Button2.Visible = False
        Button4.Visible = False
        Button5.Visible = False
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Parameters_setting.Close()
        Parameters_setting.Show()

    End Sub

    Private Sub AboutToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        flow1.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Warning.Close()
        Warning.Show()


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        warn6.Close()
        warn6.Show()
        

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        results_display2.Show()
        Me.Hide()
    End Sub

    Private Sub ResultsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '  CreateDirectory(GetCurrentDirectory() + "\result files")
        Dim a As String = GetCurrentDirectory() + "\result files\Result.txt"
        Dim b As String = GetCurrentDirectory() + "\result files\Data1.dat"
        Dim c As String = GetCurrentDirectory() + "\result files\Data2.dat"
        Dim d As String = GetCurrentDirectory() + "\result files\Data3.dat"
        Dim f As String = GetCurrentDirectory() + "\result files\Data4.dat"
        Dim g As String = GetCurrentDirectory() + "\result files\Detail Description of Data Files.txt"
        Dim Msg As String
        FileOpen(1, Application.StartupPath + "\Result.dat", OpenMode.Output)
        ' MousePointer = vbHourglass
        frmspread.grdspread.Col = 199
        frmspread.grdspread.Row = 118
        ReT2 = Val(frmspread.grdspread.Text)
        If ReT2 <> 0.0# Then
            For i = 1 To 199
                For j = 1 To 59
                    frmspread.grdspread.Col = i
                    frmspread.grdspread.Row = j
                    Re = Val(frmspread.grdspread.Text)
                    If j < 59 Then
                        Write(1, Re)
                    ElseIf j = 59 Then
                        Print(1, Re & vbCrLf)
                    End If
                Next j
            Next i
            For i = 1 To 199
                For j = 60 To 118
                    frmspread.grdspread.Col = i
                    frmspread.grdspread.Row = j
                    Re = Val(frmspread.grdspread.Text)
                    If j < 118 Then
                        Write(1, Re)
                    ElseIf j = 118 Then
                        Print(1, Re & vbCrLf)
                    End If
                Next j
            Next i
            For i = 1 To 199
                frmspread.grdspread.Col = i
                frmspread.grdspread.Row = 177
                ReT = Val(frmspread.grdspread.Text)
                If ReT <> 0.0# Then
                    For j = 118 To 177
                        frmspread.grdspread.Col = i
                        frmspread.grdspread.Row = j
                        Re = Val(frmspread.grdspread.Text)
                        If j < 177 Then
                            Write(1, Re)
                        ElseIf j = 177 Then
                            Print(1, Re & vbCrLf)
                        End If
                    Next j
                Else
                    GoTo 111
                End If
            Next i
        ElseIf ReT2 = 0.0# Then
            frmspread.grdspread.Col = 199
            frmspread.grdspread.Row = 59
            ReT1 = Val(frmspread.grdspread.Text)
            If ReT1 <> 0.0# Then
                For i = 1 To 199
                    For j = 1 To 59
                        frmspread.grdspread.Col = i
                        frmspread.grdspread.Row = j
                        Re = Val(frmspread.grdspread.Text)
                        If j < 59 Then
                            Write(1, Re)
                        ElseIf j = 59 Then
                            Print(1, Re & vbCrLf)
                        End If
                    Next j
                Next i
                For i = 1 To 199
                    frmspread.grdspread.Col = i
                    frmspread.grdspread.Row = 118
                    ReT = Val(frmspread.grdspread.Text)
                    If ReT <> 0.0# Then
                        For j = 60 To 118
                            frmspread.grdspread.Col = i
                            frmspread.grdspread.Row = j
                            Re = Val(frmspread.grdspread.Text)
                            If j < 118 Then
                                Write(1, Re)
                            ElseIf j = 118 Then
                                Print(1, Re & vbCrLf)
                            End If
                        Next j
                    Else
                        GoTo 111
                    End If
                Next i
            Else
                For i = 1 To 199
                    frmspread.grdspread.Col = i
                    frmspread.grdspread.Row = 59
                    ReT = Val(frmspread.grdspread.Text)
                    If ReT <> 0.0# Then
                        For j = 1 To 59
                            frmspread.grdspread.Col = i
                            frmspread.grdspread.Row = j
                            Re = Val(frmspread.grdspread.Text)
                            If j < 59 Then
                                Write(1, Re)
                            ElseIf j = 59 Then
                                Print(1, Re & vbCrLf)
                            End If
                        Next j
                    Else
                        GoTo 111
                    End If
                Next i
            End If
        Else
        End If

111:    FileClose(1)

        FileCopy("Result.dat", a)
        FileCopy("DATA1.OUT", b)
        FileCopy("DATA2.OUT", c)
        FileCopy("DATA3.OUT", d)
        FileCopy("DATA4.OUT", f)
        FileCopy("Detail Description of Data Files.txt", g)
        '  MousePointer = vbDefault

        'Unload Save1
        ' Load(read1)
        'read1.Show()
        GoTo 222

        'MyError:
        ' MousePointer = vbDefault
        ' Load(warn5)
        'warn5.Show()

222:    Msg = " "

        System.Diagnostics.Process.Start(GetCurrentDirectory() + "\Result files")

    End Sub



    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ABOUT.Close()
        ABOUT.Show()
    End Sub

    Private Sub OutputFileDescriptionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        System.Diagnostics.Process.Start(GetCurrentDirectory() + "\result files\Detail Description of Data Files.txt")
    End Sub

    Private Sub FileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        ABOUT.Close()
        ABOUT.Show()
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        Dim a As String = GetCurrentDirectory() + "\result files\Result.txt"
        Dim b As String = GetCurrentDirectory() + "\result files\Data1.dat"
        Dim c As String = GetCurrentDirectory() + "\result files\Data2.dat"
        Dim d As String = GetCurrentDirectory() + "\result files\Data3.dat"
        Dim f As String = GetCurrentDirectory() + "\result files\Data4.dat"
        Dim g As String = GetCurrentDirectory() + "\result files\Detail Description of Data Files.txt"
        Dim Msg As String
        ' Dim dingwei1, dingwei2, dingwei3, dingwei4, dingwei5, dingwei6 As Single
        FileOpen(1, Application.StartupPath + "\Result.dat", OpenMode.Output)
        ' MousePointer = vbHourglass
        Dim markhang, marklie As Integer
        Dim re000, re001, re002 As Single

        For i = 1 To 399
            frmspread.grdspread.Col = i
            frmspread.grdspread.Row = 59
            re000 = Val(frmspread.grdspread.Text)
            If re000 = 0 Then
                markhang = i
                marklie = 59
                GoTo 788
            End If

        Next i

        For i = 1 To 399
            frmspread.grdspread.Col = i
            frmspread.grdspread.Row = 118
            re001 = Val(frmspread.grdspread.Text)
            If re001 = 0 Then
                markhang = i
                marklie = 118
                GoTo 788
            End If
        Next i

        For i = 1 To 399
            frmspread.grdspread.Col = i
            frmspread.grdspread.Row = 176
            re002 = Val(frmspread.grdspread.Text)
            If re002 = 0 Then
                markhang = i
                marklie = 176
                GoTo 788
            End If
        Next i
         
788:
        If marklie = 59 Then
            
            For i = 1 To markhang - 1
                For j = 1 To 59
                    frmspread.grdspread.Col = i
                    frmspread.grdspread.Row = j
                    Re = Val(frmspread.grdspread.Text)
                    If j < 59 Then
                        Write(1, Re)
                    ElseIf j = 59 Then
                        Print(1, Re & vbCrLf)
                    End If
                Next j
            Next i
             
        End If

        If marklie = 118 Then
            For i = 1 To 399
                For j = 1 To 59
                    frmspread.grdspread.Col = i
                    frmspread.grdspread.Row = j
                    Re = Val(frmspread.grdspread.Text)
                    If j < 59 Then
                        Write(1, Re)
                    ElseIf j = 59 Then
                        Print(1, Re & vbCrLf)
                    End If
                Next j
            Next i
      
            For i = 1 To markhang - 1
                For j = 60 To 118
                    frmspread.grdspread.Col = i
                    frmspread.grdspread.Row = j
                    Re = Val(frmspread.grdspread.Text)
                    If j < 118 Then
                        Write(1, Re)
                    ElseIf j = 118 Then
                        Print(1, Re & vbCrLf)
                    End If
                Next j
            Next i
        End If

        If marklie = 176 Then
            For i = 1 To 399
                For j = 1 To 59
                    frmspread.grdspread.Col = i
                    frmspread.grdspread.Row = j
                    Re = Val(frmspread.grdspread.Text)
                    If j < 59 Then
                        Write(1, Re)
                    ElseIf j = 59 Then
                        Print(1, Re & vbCrLf)
                    End If
                Next j
            Next i
            For i = 1 To 399
                For j = 60 To 118
                    frmspread.grdspread.Col = i
                    frmspread.grdspread.Row = j
                    Re = Val(frmspread.grdspread.Text)
                    If j < 118 Then
                        Write(1, Re)
                    ElseIf j = 118 Then
                        Print(1, Re & vbCrLf)
                    End If
                Next j
            Next i

            For i = 1 To markhang - 1
                For j = 118 To 176
                    frmspread.grdspread.Col = i
                    frmspread.grdspread.Row = j
                    Re = Val(frmspread.grdspread.Text)
                    If j < 176 Then
                        Write(1, Re)
                    ElseIf j = 176 Then
                        Print(1, Re & vbCrLf)
                    End If
                Next j
            Next i

        End If



111:    FileClose(1)

        FileCopy("Result.dat", a)
        FileCopy("DATA1.OUT", b)
        FileCopy("DATA2.OUT", c)
        FileCopy("DATA3.OUT", d)
        FileCopy("DATA4.OUT", f)
        FileCopy("Detail Description of Data Files.txt", g)
        '  MousePointer = vbDefault

        'Unload Save1
        ' Load(read1)
        'read1.Show()
        GoTo 222

        'MyError:
        ' MousePointer = vbDefault
        ' Load(warn5)
        'warn5.Show()

222:    Msg = " "

        System.Diagnostics.Process.Start(GetCurrentDirectory() + "\Result files")
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        System.Diagnostics.Process.Start(GetCurrentDirectory() + "\result files\Detail Description of Data Files.txt")
    End Sub

    Private Sub FileToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FileToolStripMenuItem1.Click

    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub
End Class