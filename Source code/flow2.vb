Public Class flow2
    Dim out(37) As Single

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        flow1.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        curves.Show()
    End Sub

    Private Sub Label39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label39.Click
        ' Load(Chan)
        Chan.Close()

        Chan.Show()

        Chan.Label2.Text = "Proportional Gain for Zone1 VAV Damper Control(%/C)"
        Chan.Label3.Text = "Integral Time for Zone1 VAV Damper Control(Second)"

        M = 3
        n = 4

        'On Error GoTo MyError

        'Open "BMSin.dat" For Input As #1
        FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(1, out(i))
        Next i
        FileClose(1)

        'MyError:
        '    Resume

        Chan.Text1.Text = Str$(out(3))
        Chan.Text2.Text = Str$(out(4))

    End Sub

   

    Private Sub Label20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label20.Click
        '  Load(chan7)
        chan7.Close()
        chan7.Show()

        chan7.Label2.Text = "Minimum Temperature Setpoint (°C)"
        chan7.Label3.Text = "Maximum Temperature Setpoint (°C)"
        chan7.Label4.Text = "Maximum change rate of temperature setpoint(°C/S)"
        chan7.Label5.Text = "Proportional Gain for Flow limit check(C/%)"
        chan7.Label6.Text = "Integral Time for Flow limit check(Second)"
        chan7.Label7.Text = "Proportional Gain for Flow modification(C/%)"
        chan7.Label8.Text = "Integral Time for Flow modification(Second)"

        M = 20
        n = 21
        m1 = 22
        m2 = 23
        m3 = 24
        m4 = 25
        m5 = 26

        'On Error GoTo MyError

        'Open "BMSin.dat" For Input As #1
        FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(1, out(i))
        Next i
        FileClose(1)

        'MyError:
        '     Resume

        chan7.Text1.Text = Str$(out(20))
        chan7.Text2.Text = Str$(out(21))
        chan7.Text3.Text = Str$(out(22))
        chan7.Text4.Text = Str$(out(23))
        chan7.Text5.Text = Str$(out(24))
        chan7.Text6.Text = Str$(out(25))
        chan7.Text7.Text = Str$(out(26))

    End Sub

    Private Sub Label21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label21.Click
        ' Load(Chan)
        Chan.Close()
        Chan.Show()

        Chan.Label2.Text = "Proportional Gain for Zone2 VAV Damper Control(%/C)"
        Chan.Label3.Text = "Integral Time for Zone2 VAV Damper Control(Second)"

        M = 36
        n = 37

        'On Error GoTo MyError

        'Open "BMSin.dat" For Input As #1
        FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(1, out(i))
        Next i
        FileClose(1)

        'MyError:
        '    Resume

        Chan.Text1.Text = Str$(out(36))
        Chan.Text2.Text = Str$(out(37))

    End Sub

    Private Sub Label23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label23.Click
        ' Load(chan3)
        chan3.Close()
        chan3.Show()

        chan3.Label2.Text = "Proportional Gain for Zone2 Temperature Control(%/C)"
        chan3.Label3.Text = "Integral Time for Zone2 Temperature Control(Second)"
        chan3.Label4.Text = "Room Temperature Setpoint of 8 Zones(°C)"

        M = 32
        n = 33
        m1 = 27

        'On Error GoTo MyError

        'Open "BMSin.dat" For Input As #1
        FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(1, out(i))
        Next i
        FileClose(1)

        'MyError:
        '    Resume

        chan3.Text1.Text = Str$(out(32))
        chan3.Text2.Text = Str$(out(33))
        chan3.Text3.Text = Str$(out(27))


    End Sub

    Private Sub Label24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label24.Click
        ' Load(chan7)
        chan7.Close()
        chan7.Show()

        chan7.Label2.Text = "Minimum Pressure Setpoint (pa)"
        chan7.Label3.Text = "Maximum Pressure Setpoint (pa)"
        chan7.Label4.Text = "Proportional Gain for Reducing Pset (Pa/%)"
        chan7.Label5.Text = "Integral Time for Reducing Pset(Second)"
        chan7.Label6.Text = "Proportional Gain for Increasing Pset((Pa/%)"
        chan7.Label7.Text = "Integral Time for Increasing Pset(Second)"
        chan7.Label8.Text = "Proportional Gain for Setpoint Change(Pa/%)"

        M = 13
        n = 14
        m1 = 15
        m2 = 16
        m3 = 17
        m4 = 18
        m5 = 19

        'On Error GoTo MyError

        'Open "BMSin.dat" For Input As #1
        FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(1, out(i))
        Next i
        FileClose(1)

        'MyError:
        '    Resume

        chan7.Text1.Text = Str$(out(13))
        chan7.Text2.Text = Str$(out(14))
        chan7.Text3.Text = Str$(out(15))
        chan7.Text4.Text = Str$(out(16))
        chan7.Text5.Text = Str$(out(17))
        chan7.Text6.Text = Str$(out(18))
        chan7.Text7.Text = Str$(out(19))

    End Sub

    Private Sub Label25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label25.Click
        ' Load(chan3)
        chan3.Close()
        chan3.Show()

        chan3.Label2.Text = "Proportional Gain for Zone1 Temperature Control(%/C)"
        chan3.Label3.Text = "Integral Time for Zone1 Temperature Control(Second)"
        chan3.Label4.Text = "Room Temperature Setpoint of 8 Zones(°C)"

        M = 1
        n = 2
        m1 = 27

        'On Error GoTo MyError

        'Open "BMSin.dat" For Input As #1
        FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(1, out(i))
        Next i
        FileClose(1)

        'MyError:
        '    Resume

        chan3.Text1.Text = Str$(out(1))
        chan3.Text2.Text = Str$(out(2))
        chan3.Text3.Text = Str$(out(27))

    End Sub

    Private Sub flow2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.ControlBox = False
        '  MaximizeBox = False
        '  MinimizeBox = True
        MaximizeBox = False
        Me.TopMost = False
        Me.Height = My.Computer.Screen.WorkingArea.Height
        Me.Width = My.Computer.Screen.WorkingArea.Width
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        '比例缩放
        Dim x1, y1, x2, y2, a_width, b_height, c_font As Integer

        '窗口大小 a为宽，b为高
        a_width = 1041
        b_height = 653
        ''
        x2 = 1028
        y2 = 605
        Me.PictureBox1.Width = Me.Width
        Me.PictureBox1.Height = Me.Height * (y2 / b_height)
        PictureBox1.Location = New Point(0, MenuStrip1.Height)
        x1 = 13
        y1 = 628
        x2 = 103
        y2 = 20
        Me.Label27.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label27.Width = x2 / a_width * Me.Width
        Me.Label27.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Label27.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        '  x1 = 108
        y1 = 628
        x2 = 56
        y2 = 20
        Me.Label45.Location = New Point(x1 / a_width * Me.Width + Label27.Width * 1.2, y1 / b_height * Me.Height)
        Me.Label45.Width = x2 / a_width * Me.Width
        Me.Label45.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Label45.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        '      x1 = 170
        y1 = 628
        x2 = 64
        y2 = 20
        Me.Label26.Location = New Point(x1 / a_width * Me.Width + Label27.Width * 1.45 + Label45.Width, y1 / b_height * Me.Height)
        Me.Label26.Width = x2 / a_width * Me.Width
        Me.Label26.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Label26.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 927
        y1 = 623
        x2 = 85
        y2 = 27
        Me.Button3.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Button3.Width = x2 / a_width * Me.Width
        Me.Button3.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Button3.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 534
        y1 = 81
        x2 = 21
        y2 = 16
        Me.Label4.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label4.Width = x2 / a_width * Me.Width
        Me.Label4.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label4.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 699
        y1 = 81
        x2 = 21
        y2 = 16
        Me.Label6.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label6.Width = x2 / a_width * Me.Width
        Me.Label6.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label6.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 872
        y1 = 81
        x2 = 21
        y2 = 16
        Me.Label8.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label8.Width = x2 / a_width * Me.Width
        Me.Label8.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label8.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 534
        y1 = 100
        x2 = 34
        y2 = 16
        Me.Label3.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label3.Width = x2 / a_width * Me.Width
        Me.Label3.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label3.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 699
        y1 = 100
        x2 = 34
        y2 = 16
        Me.Label5.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label5.Width = x2 / a_width * Me.Width
        Me.Label5.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label5.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 872
        y1 = 100
        x2 = 34
        y2 = 16
        Me.Label7.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label7.Width = x2 / a_width * Me.Width
        Me.Label7.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label7.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 586
        y1 = 132
        x2 = 52
        y2 = 16
        Me.Label30.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label30.Width = x2 / a_width * Me.Width
        Me.Label30.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label30.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 745
        y1 = 132
        x2 = 52
        y2 = 16
        Me.Label32.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label32.Width = x2 / a_width * Me.Width
        Me.Label32.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label32.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 925
        y1 = 132
        x2 = 52
        y2 = 16
        Me.Label34.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label34.Width = x2 / a_width * Me.Width
        Me.Label34.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label34.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 586
        y1 = 271
        x2 = 52
        y2 = 16
        Me.Label31.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label31.Width = x2 / a_width * Me.Width
        Me.Label31.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label31.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 745
        y1 = 271
        x2 = 52
        y2 = 16
        Me.Label33.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label33.Width = x2 / a_width * Me.Width
        Me.Label33.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label33.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 925
        y1 = 271
        x2 = 52
        y2 = 16
        Me.Label35.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label35.Width = x2 / a_width * Me.Width
        Me.Label35.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label35.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 571
        y1 = 327
        x2 = 21
        y2 = 16
        Me.Label12.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label12.Width = x2 / a_width * Me.Width
        Me.Label12.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label12.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 742
        y1 = 327
        x2 = 21
        y2 = 16
        Me.Label1.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label1.Width = x2 / a_width * Me.Width
        Me.Label1.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label1.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 918
        y1 = 327
        x2 = 21
        y2 = 16
        Me.Label2.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label2.Width = x2 / a_width * Me.Width
        Me.Label2.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label2.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)



        x1 = 571
        y1 = 347
        x2 = 34
        y2 = 16
        Me.Label22.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label22.Width = x2 / a_width * Me.Width
        Me.Label22.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label22.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 106
        y1 = 161
        x2 = 52
        y2 = 16
        Me.Label28.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label28.Width = x2 / a_width * Me.Width
        Me.Label28.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label28.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 350
        y1 = 121
        x2 = 21
        y2 = 16
        Me.Label10.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label10.Width = x2 / a_width * Me.Width
        Me.Label10.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label10.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 123
        y1 = 365
        x2 = 52
        y2 = 16
        Me.Label29.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label29.Width = x2 / a_width * Me.Width
        Me.Label29.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label29.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 300
        y1 = 313
        x2 = 34
        y2 = 16
        Me.Label19.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label19.Width = x2 / a_width * Me.Width
        Me.Label19.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label19.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 350
        y1 = 298
        x2 = 21
        y2 = 16
        Me.Label11.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label11.Width = x2 / a_width * Me.Width
        Me.Label11.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label11.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 6
        y1 = 446
        x2 = 59
        y2 = 16
        Me.Label13.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label13.Width = x2 / a_width * Me.Width
        Me.Label13.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label13.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 6
        y1 = 465
        x2 = 42
        y2 = 16
        Me.Label14.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label14.Width = x2 / a_width * Me.Width
        Me.Label14.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label14.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 6
        y1 = 491
        x2 = 65
        y2 = 16
        Me.Label15.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label15.Width = x2 / a_width * Me.Width
        Me.Label15.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label15.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 6
        y1 = 512
        x2 = 35
        y2 = 16
        Me.Label16.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label16.Width = x2 / a_width * Me.Width
        Me.Label16.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label16.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 6
        y1 = 541
        x2 = 90
        y2 = 16
        Me.Label18.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label18.Width = x2 / a_width * Me.Width
        Me.Label18.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label18.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 6
        y1 = 564
        x2 = 35
        y2 = 16
        Me.Label17.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label17.Width = x2 / a_width * Me.Width
        Me.Label17.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label17.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 171
        y1 = 161
        x2 = 101
        y2 = 31
        Me.Label39.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label39.Width = x2 / a_width * Me.Width
        Me.Label39.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label39.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 300
        y1 = 180
        x2 = 34
        y2 = 16
        Me.Label9.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label9.Width = x2 / a_width * Me.Width
        Me.Label9.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label9.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 357
        y1 = 158
        x2 = 199
        y2 = 36
        Me.Label25.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label25.Width = x2 / a_width * Me.Width
        Me.Label25.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label25.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 357
        y1 = 343
        x2 = 100
        y2 = 32
        Me.Label23.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label23.Width = x2 / a_width * Me.Width
        Me.Label23.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label23.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 176
        y1 = 343
        x2 = 101
        y2 = 32
        Me.Label21.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label21.Width = x2 / a_width * Me.Width
        Me.Label21.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label21.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 184
        y1 = 533
        x2 = 163
        y2 = 45
        Me.Label20.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label20.Width = x2 / a_width * Me.Width
        Me.Label20.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label20.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 652
        y1 = 529
        x2 = 152
        y2 = 54
        Me.Label24.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label24.Width = x2 / a_width * Me.Width
        Me.Label24.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label24.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        Label1.TextAlign = ContentAlignment.MiddleCenter
        Label2.TextAlign = ContentAlignment.MiddleCenter
        Label3.TextAlign = ContentAlignment.MiddleCenter
        Label4.TextAlign = ContentAlignment.MiddleCenter
        Label5.TextAlign = ContentAlignment.MiddleCenter
        Label6.TextAlign = ContentAlignment.MiddleCenter
        Label7.TextAlign = ContentAlignment.MiddleCenter
        Label8.TextAlign = ContentAlignment.MiddleCenter
        Label9.TextAlign = ContentAlignment.MiddleCenter
        Label10.TextAlign = ContentAlignment.MiddleCenter

        Label11.TextAlign = ContentAlignment.MiddleCenter
        Label12.TextAlign = ContentAlignment.MiddleCenter
        Label13.TextAlign = ContentAlignment.MiddleCenter
        Label14.TextAlign = ContentAlignment.MiddleCenter
        Label15.TextAlign = ContentAlignment.MiddleCenter
        Label16.TextAlign = ContentAlignment.MiddleCenter
        Label17.TextAlign = ContentAlignment.MiddleCenter
        Label18.TextAlign = ContentAlignment.MiddleCenter
        Label19.TextAlign = ContentAlignment.MiddleCenter
        Label20.TextAlign = ContentAlignment.MiddleCenter

        Label21.TextAlign = ContentAlignment.MiddleCenter
        Label22.TextAlign = ContentAlignment.MiddleCenter
        Label23.TextAlign = ContentAlignment.MiddleCenter
        Label24.TextAlign = ContentAlignment.MiddleCenter
        Label25.TextAlign = ContentAlignment.MiddleCenter
        Label26.TextAlign = ContentAlignment.MiddleCenter
        Label27.TextAlign = ContentAlignment.MiddleCenter
        Label28.TextAlign = ContentAlignment.MiddleCenter
        Label29.TextAlign = ContentAlignment.MiddleCenter
        Label30.TextAlign = ContentAlignment.MiddleCenter

        Label31.TextAlign = ContentAlignment.MiddleCenter
        Label32.TextAlign = ContentAlignment.MiddleCenter
        Label33.TextAlign = ContentAlignment.MiddleCenter
        Label34.TextAlign = ContentAlignment.MiddleCenter
        Label35.TextAlign = ContentAlignment.MiddleCenter
     
        Label45.TextAlign = ContentAlignment.MiddleCenter






        Label1.Parent = PictureBox1
        Label2.Parent = PictureBox1
        Label3.Parent = PictureBox1
        Label4.Parent = PictureBox1
        Label5.Parent = PictureBox1
        Label6.Parent = PictureBox1
        Label7.Parent = PictureBox1
        Label8.Parent = PictureBox1
        Label9.Parent = PictureBox1
        Label10.Parent = PictureBox1
        Label11.Parent = PictureBox1
        Label12.Parent = PictureBox1
        Label13.Parent = PictureBox1
        Label14.Parent = PictureBox1
        Label15.Parent = PictureBox1
        Label16.Parent = PictureBox1
        Label17.Parent = PictureBox1
        Label18.Parent = PictureBox1
        Label19.Parent = PictureBox1
        Label20.Parent = PictureBox1
        Label21.Parent = PictureBox1
        Label22.Parent = PictureBox1
        Label23.Parent = PictureBox1
        Label24.Parent = PictureBox1
        Label25.Parent = PictureBox1
        '  Label26.Parent = PictureBox1
        ' Label27.Parent = PictureBox1
        Label28.Parent = PictureBox1
        Label29.Parent = PictureBox1
        Label30.Parent = PictureBox1
        Label31.Parent = PictureBox1
        Label32.Parent = PictureBox1
        Label33.Parent = PictureBox1
        Label34.Parent = PictureBox1
        Label35.Parent = PictureBox1
        Label39.Parent = PictureBox1

        Label1.BackColor = Color.Transparent
        Label2.BackColor = Color.Transparent
        Label3.BackColor = Color.Transparent
        Label4.BackColor = Color.Transparent
        Label5.BackColor = Color.Transparent
        Label6.BackColor = Color.Transparent
        Label7.BackColor = Color.Transparent
        Label8.BackColor = Color.Transparent
        Label9.BackColor = Color.Transparent
        Label10.BackColor = Color.Transparent
        Label11.BackColor = Color.Transparent
        Label12.BackColor = Color.Transparent
        Label13.BackColor = Color.Transparent
        Label14.BackColor = Color.Transparent
        Label15.BackColor = Color.Transparent
        Label16.BackColor = Color.Transparent
        Label17.BackColor = Color.Transparent
        Label18.BackColor = Color.Transparent
        Label19.BackColor = Color.Transparent
        Label20.BackColor = Color.Transparent
        Label21.BackColor = Color.Transparent
        Label22.BackColor = Color.Transparent
        Label23.BackColor = Color.Transparent
        Label24.BackColor = Color.Transparent
        Label25.BackColor = Color.Transparent
        Label26.BackColor = Color.Transparent
        Label27.BackColor = Color.Transparent
        Label28.BackColor = Color.Transparent
        Label29.BackColor = Color.Transparent
        Label30.BackColor = Color.Transparent
        Label31.BackColor = Color.Transparent
        Label32.BackColor = Color.Transparent
        Label33.BackColor = Color.Transparent
        Label34.BackColor = Color.Transparent
        Label35.BackColor = Color.Transparent
        Label39.BackColor = Color.Transparent
        Label45.BackColor = Color.Transparent
   
        'TextBox14.BorderStyle = BorderStyle.None
        ' TextBox18.BorderStyle = BorderStyle.None
        '  TextBox22.BorderStyle = BorderStyle.None
        '  TextBox23.BorderStyle = BorderStyle.None
        '  TextBox24.BorderStyle = BorderStyle.None
        '  TextBox25.BorderStyle = BorderStyle.None
        'TextBox26.BorderStyle = BorderStyle.None
        'TextBox27.BorderStyle = BorderStyle.None
        'TextBox28.BorderStyle = BorderStyle.None
        ' TextBox29.BorderStyle = BorderStyle.None
        ' TextBox30.BorderStyle = BorderStyle.None
        ' TextBox31.BorderStyle = BorderStyle.None
        ' TextBox32.BorderStyle = BorderStyle.None
        'TextBox33.BorderStyle = BorderStyle.None
        '  TextBox34.BorderStyle = BorderStyle.None
        'TextBox35.BorderStyle = BorderStyle.None
        '   TextBox36.BorderStyle = BorderStyle.None
        ' TextBox37.BorderStyle = BorderStyle.None

        ' TextBox39.BorderStyle = BorderStyle.None
        ' TextBox40.BorderStyle = BorderStyle.None
        'TextBox41.BorderStyle = BorderStyle.None
        'TextBox42.BorderStyle = BorderStyle.None
        'TextBox43.BorderStyle = BorderStyle.None
        'TextBox44.BorderStyle = BorderStyle.None
        'TextBox45.BorderStyle = BorderStyle.None
        ' TextBox46.BorderStyle = BorderStyle.None

        '    TextBox14.BackColor = Color.FromArgb(240, 240, 240)
        '  TextBox18.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox22.BackColor = Color.FromArgb(240, 240, 240)
        '  TextBox23.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox24.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox25.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox26.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox27.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox28.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox29.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox30.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox31.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox32.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox33.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox34.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox35.BackColor = Color.FromArgb(240, 240, 240)
        '   TextBox36.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox37.BackColor = Color.FromArgb(240, 240, 240)

        'TextBox39.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox40.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox41.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox42.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox43.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox44.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox45.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox46.BackColor = Color.FromArgb(240, 240, 240)
        Label39.Cursor = System.Windows.Forms.Cursors.Hand
        Label25.Cursor = System.Windows.Forms.Cursors.Hand
        Label21.Cursor = System.Windows.Forms.Cursors.Hand
        Label23.Cursor = System.Windows.Forms.Cursors.Hand
        Label20.Cursor = System.Windows.Forms.Cursors.Hand
        Label24.Cursor = System.Windows.Forms.Cursors.Hand


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        flow1.TopMost = False
        Me.TopMost = False
        Menubms.Show()
        ' Control.Hide()
        Me.Hide()
        flow1.Hide()
    End Sub

    Private Sub Label27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label27.Click

    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 98

            ' Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 3 (°C)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 98

            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 3 (°C)"
        End If
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 85

            '   Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Air flow setpoint of Zone 3 (kg/s)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 85

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Air flow setpoint of Zone 3 (kg/s)"
        End If
    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 96
            curve.Close()
            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 5 (°C)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 96

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 5 (°C)"
        End If
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 94
            curve.Close()
            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 7 (°C)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 94

            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 7 (°C)"
        End If
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 83
            curve.Close()
            'Load(curve)
            curve.Show()
            curve.Label19.Text = "Air flow setpoint of Zone 5 (kg/s)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 83

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Air flow setpoint of Zone 5 (kg/s)"
        End If
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        If Fcurve = 0 Then
            Vmax = 2.0#
            Vmin = 0.0#
            TK = 81
            curve.Close()
            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Air flow setpoint of Zone 7 (kg/s)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 2.0#
            Vmin = 0.0#
            TK = 81

            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Air flow setpoint of Zone 7 (kg/s)"
        End If
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 97
            curve.Close()
            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 4 (°C)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 97

            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 4 (°C)"
        End If
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 95
            curve.Close()
            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 6 (°C)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 95

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 6 (°C)"
        End If
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 93
            curve.Close()
            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 8 (°C)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 93

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 8 (°C)"
        End If
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 100
            curve.Close()
            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 1 (°C)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 100

            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 1 (°C)"
        End If
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 99
            curve.Close()
            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 2 (°C)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 99

            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of Zone 2 (°C)"
        End If
    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click
        If Fcurve = 0 Then
            Vmax = 2.0#
            Vmin = 0.0#
            TK = 87
            curve.Close()
            'Load(curve)
            curve.Show()
            curve.Label19.Text = "Air flow setpoint of Zone 1 (kg/s)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 2.0#
            Vmin = 0.0#
            TK = 87

            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Air flow setpoint of Zone 1 (kg/s)"
        End If
    End Sub

    Private Sub Label19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label19.Click
        If Fcurve = 0 Then
            Vmax = 2.0#
            Vmin = 0.0#
            TK = 86
            curve.Close()
            '    Load(curve)
            curve.Show()
            curve.Label19.Text = "Air flow setpoint of Zone 2 (kg/s)"

        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 2.0#
            Vmin = 0.0#
            TK = 86

            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Air flow setpoint of Zone 2 (kg/s)"
        End If
    End Sub

    Private Sub Label22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label22.Click
        If Fcurve = 0 Then
            Vmax = 2.0#
            Vmin = 0.0#
            TK = 84
            curve.Close()
            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Air flow setpoint of Zone 4 (kg/s)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 2.0#
            Vmin = 0.0#
            TK = 84

            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Air flow setpoint of Zone 4 (kg/s)"
        End If
    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click
        If Fcurve = 0 Then
            Vmax = 0.02
            Vmin = 0.0#
            TK = 92
            curve.Close()
            '    Load(curve)
            curve.Show()
            curve.Label19.Text = "Average humidity of Zone (kg/kg)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 0.02
            Vmin = 0.0#
            TK = 92

            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Average humidity of Zone (kg/kg)"
        End If
    End Sub

    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label16.Click
        If Fcurve = 0 Then
            Vmax = 1500.0#
            Vmin = 0.0#
            TK = 91
            curve.Close()
            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Average CO2 concentration of Zone (ppm)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1500.0#
            Vmin = 0.0#
            TK = 91

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Average CO2 concentration of Zone (ppm)"
        End If
    End Sub

    Private Sub Label17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label17.Click
        If Fcurve = 0 Then
            Vmax = 6.0#
            Vmin = 0.0#
            TK = 90
            curve.Close()
            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Average pollutant concentration of Zone (ppm)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 6.0#
            Vmin = 0.0#
            TK = 90

            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Average pollutant concentration of Zone (ppm)"
        End If
    End Sub

    Private Sub Label28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label28.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 79
            curve.Close()
            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 1 (0-1)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 79

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 1 (0-1)"
        End If
    End Sub

    Private Sub Label29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label29.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 78
            curve.Close()
            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 2 (0-1)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 78

            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 2 (0-1)"
        End If
    End Sub

    Private Sub Label30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label30.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 77
            curve.Close()
            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 3 (0-1)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 77

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 3 (0-1)"
        End If
    End Sub

    Private Sub Label31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label31.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 76
            curve.Close()
            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 4 (0-1)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 76

            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 4 (0-1)"
        End If
    End Sub

    Private Sub Label32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label32.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 75
            curve.Close()
            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 5 (0-1)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 75

            'Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 5 (0-1)"
        End If
    End Sub

    Private Sub Label33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label33.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 74
            curve.Close()
            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 6 (0-1)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 74

            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 6 (0-1)"
        End If
    End Sub

    Private Sub Label34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label34.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 73
            curve.Close()
            '    Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 7 (0-1)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 73

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 7 (0-1)"
        End If
    End Sub

    Private Sub Label35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label35.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 72
            curve.Close()
            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 8 (0-1)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 72

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV damper position demand of Zone 8 (0-1)"
        End If
    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        '  flow1.TopMost = True
        ' Me.TopMost = False
        ' flow1.TopMost = False
        flow1.Show()
        Me.Hide()
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        Me.Show()
        flow1.Hide()
        '  Me.TopMost = True
        ' flow1.TopMost = False
        ' Me.TopMost = False
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub
End Class