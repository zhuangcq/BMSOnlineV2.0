Imports System.IO.Directory
Public Class flow1
    Dim Wfvav, Wfcav, Wfrtn, Qcvav, Qccav As Single
    Dim Wfvav1, Wfcav1, Wfrtn1, Qcvav1, Qccav1 As Single
    Dim out(37) As Single


    Private Sub flow1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = False
        MaximizeBox = False
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
        x1 = 135
        y1 = 9
        x2 = 21
        y2 = 16
        Me.Label19.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label19.Width = x2 / a_width * Me.Width
        Me.Label19.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label19.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 214
        y1 = 8
        x2 = 34
        y2 = 16
        Me.Label21.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label21.Width = x2 / a_width * Me.Width
        Me.Label21.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label21.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 381
        y1 = 13
        x2 = 28
        y2 = 16
        Me.Label26.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label26.Width = x2 / a_width * Me.Width
        Me.Label26.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label26.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 725
        y1 = 13
        x2 = 34
        y2 = 16
        Me.Label8.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label8.Width = x2 / a_width * Me.Width
        Me.Label8.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label8.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)



        x1 = 800
        y1 = 13
        x2 = 21
        y2 = 16
        Me.Label7.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label7.Width = x2 / a_width * Me.Width
        Me.Label7.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label7.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 641
        y1 = 75
        x2 = 28
        y2 = 16
        Me.Label11.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label11.Width = x2 / a_width * Me.Width
        Me.Label11.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label11.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 103
        y1 = 125
        x2 = 42
        y2 = 16
        Me.Label18.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label18.Width = x2 / a_width * Me.Width
        Me.Label18.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label18.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 285
        y1 = 121
        x2 = 34
        y2 = 16
        Me.Label6.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label6.Width = x2 / a_width * Me.Width
        Me.Label6.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label6.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 510
        y1 = 106
        x2 = 56
        y2 = 16
        Me.Label32.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label32.Width = x2 / a_width * Me.Width
        Me.Label32.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label32.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 630
        y1 = 182
        x2 = 56
        y2 = 16
        Me.Label42.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label42.Width = x2 / a_width * Me.Width
        Me.Label42.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label42.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 510
        y1 = 190
        x2 = 56
        y2 = 16
        Me.Label41.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label41.Width = x2 / a_width * Me.Width
        Me.Label41.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label41.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 694
        y1 = 240
        x2 = 34
        y2 = 16
        Me.Label13.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label13.Width = x2 / a_width * Me.Width
        Me.Label13.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label13.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 806
        y1 = 240
        x2 = 21
        y2 = 16
        Me.Label12.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label12.Width = x2 / a_width * Me.Width
        Me.Label12.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label12.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 952
        y1 = 240
        x2 = 25
        y2 = 16
        Me.Label9.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label9.Width = x2 / a_width * Me.Width
        Me.Label9.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label9.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 645
        y1 = 261
        x2 = 28
        y2 = 16
        Me.Label24.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label24.Width = x2 / a_width * Me.Width
        Me.Label24.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label24.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 381
        y1 = 278
        x2 = 28
        y2 = 16
        Me.Label25.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label25.Width = x2 / a_width * Me.Width
        Me.Label25.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label25.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 99
        y1 = 306
        x2 = 56
        y2 = 16
        Me.Label31.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label31.Width = x2 / a_width * Me.Width
        Me.Label31.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label31.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 354
        y1 = 380
        x2 = 34
        y2 = 16
        Me.Label22.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label22.Width = x2 / a_width * Me.Width
        Me.Label22.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label22.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 650
        y1 = 358
        x2 = 42
        y2 = 16
        Me.Label10.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label10.Width = x2 / a_width * Me.Width
        Me.Label10.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label10.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 419
        y1 = 399
        x2 = 28
        y2 = 16
        Me.Label15.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label15.Width = x2 / a_width * Me.Width
        Me.Label15.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label15.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 539
        y1 = 402
        x2 = 21
        y2 = 16
        Me.Label30.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label30.Width = x2 / a_width * Me.Width
        Me.Label30.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label30.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 381
        y1 = 420
        x2 = 56
        y2 = 16
        Me.Label43.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label43.Width = x2 / a_width * Me.Width
        Me.Label43.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label43.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 462
        y1 = 414
        x2 = 34
        y2 = 16
        Me.Label14.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label14.Width = x2 / a_width * Me.Width
        Me.Label14.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label14.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 128
        y1 = 425
        x2 = 34
        y2 = 16
        Me.Label29.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label29.Width = x2 / a_width * Me.Width
        Me.Label29.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label29.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 128
        y1 = 460
        x2 = 34
        y2 = 16
        Me.Label16.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label16.Width = x2 / a_width * Me.Width
        Me.Label16.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label16.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 220
        y1 = 479
        x2 = 34
        y2 = 16
        Me.Label23.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label23.Width = x2 / a_width * Me.Width
        Me.Label23.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label23.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 373
        y1 = 478
        x2 = 56
        y2 = 16
        Me.Label44.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label44.Width = x2 / a_width * Me.Width
        Me.Label44.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label44.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 219
        y1 = 515
        x2 = 34
        y2 = 16
        Me.Label17.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label17.Width = x2 / a_width * Me.Width
        Me.Label17.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label17.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 748
        y1 = 481
        x2 = 21
        y2 = 16
        Me.Label28.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label28.Width = x2 / a_width * Me.Width
        Me.Label28.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label28.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 917
        y1 = 477
        x2 = 25
        y2 = 16
        Me.Label27.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label27.Width = x2 / a_width * Me.Width
        Me.Label27.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label27.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 748
        y1 = 564
        x2 = 21
        y2 = 16
        Me.Label3.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label3.Width = x2 / a_width * Me.Width
        Me.Label3.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label3.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 917
        y1 = 564
        x2 = 25
        y2 = 16
        Me.Label5.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label5.Width = x2 / a_width * Me.Width
        Me.Label5.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label5.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 99
        y1 = 306
        x2 = 56
        y2 = 16
        Me.Label31.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label31.Width = x2 / a_width * Me.Width
        Me.Label31.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 9.75)
        Label31.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 13
        y1 = 628
        x2 = 103
        y2 = 20
        Me.Label1.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label1.Width = x2 / a_width * Me.Width
        Me.Label1.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Label1.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        'x1 = 108
        y1 = 628
        x2 = 56
        y2 = 20
        Me.Label45.Location = New Point(x1 / a_width * Me.Width + Label1.Width * 1.2, y1 / b_height * Me.Height)
        Me.Label45.Width = x2 / a_width * Me.Width
        Me.Label45.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Label45.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        '   x1 = 170
        y1 = 628
        x2 = 64
        y2 = 20
        Me.Label2.Location = New Point(x1 / a_width * Me.Width + Label1.Width * 1.2 + Label45.Width, y1 / b_height * Me.Height)
        Me.Label2.Width = x2 / a_width * Me.Width
        Me.Label2.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Label2.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 89
        y1 = 554
        x2 = 35
        y2 = 16
        Me.Label20.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label20.Width = x2 / a_width * Me.Width
        Me.Label20.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Label20.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 675
        y1 = 407
        x2 = 158
        y2 = 49
        Label4.AutoSize = True
        Label4.Text = " End! Review the" & vbCrLf & "energy consumption"
        Me.Label4.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label4.Width = x2 / a_width * Me.Width
        Me.Label4.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Label4.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 736
        y1 = 65
        x2 = 87
        y2 = 36
        Me.Label35.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label35.Width = x2 / a_width * Me.Width
        Me.Label35.Height = y2 / b_height * Me.Height

        x1 = 619
        y1 = 130
        x2 = 83
        y2 = 36
        Me.Label34.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label34.Width = x2 / a_width * Me.Width
        Me.Label34.Height = y2 / b_height * Me.Height

        x1 = 740
        y1 = 169
        x2 = 83
        y2 = 36
        Me.Label36.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label36.Width = x2 / a_width * Me.Width
        Me.Label36.Height = y2 / b_height * Me.Height

        x1 = 130
        y1 = 384
        x2 = 87
        y2 = 36
        Me.Label39.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label39.Width = x2 / a_width * Me.Width
        Me.Label39.Height = y2 / b_height * Me.Height

        x1 = 276
        y1 = 397
        x2 = 100
        y2 = 45
        Me.Label40.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label40.Width = x2 / a_width * Me.Width
        Me.Label40.Height = y2 / b_height * Me.Height

        x1 = 130
        y1 = 486
        x2 = 87
        y2 = 36
        Me.Label38.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label38.Width = x2 / a_width * Me.Width
        Me.Label38.Height = y2 / b_height * Me.Height

        x1 = 694
        y1 = 510
        x2 = 94
        y2 = 44
        Me.Label33.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label33.Width = x2 / a_width * Me.Width
        Me.Label33.Height = y2 / b_height * Me.Height

        x1 = 864
        y1 = 512
        x2 = 94
        y2 = 42
        Me.Label37.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label37.Width = x2 / a_width * Me.Width
        Me.Label37.Height = y2 / b_height * Me.Height


        x1 = 927
        y1 = 623
        x2 = 85
        y2 = 27
        Me.Button3.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Button3.Width = x2 / a_width * Me.Width
        Me.Button3.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Button3.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        ' Me.ControlBox = False
        ' MaximizeBox = False
        ' MinimizeBox = True
        Label34.Cursor = System.Windows.Forms.Cursors.Hand
        Label35.Cursor = System.Windows.Forms.Cursors.Hand
        Label36.Cursor = System.Windows.Forms.Cursors.Hand
        Label37.Cursor = System.Windows.Forms.Cursors.Hand
        Label38.Cursor = System.Windows.Forms.Cursors.Hand
        Label39.Cursor = System.Windows.Forms.Cursors.Hand
        Label40.Cursor = System.Windows.Forms.Cursors.Hand
        Label33.Cursor = System.Windows.Forms.Cursors.Hand

        '  TextBox20.BorderStyle = BorderStyle.None
        ' TextBox21.BorderStyle = BorderStyle.None
        'TextBox22.BorderStyle = BorderStyle.None
        'TextBox23.BorderStyle = BorderStyle.None
        ' TextBox24.BorderStyle = BorderStyle.None
        ' TextBox25.BorderStyle = BorderStyle.None
        '  TextBox26.BorderStyle = BorderStyle.None
        ' TextBox27.BorderStyle = BorderStyle.None
        '  TextBox28.BorderStyle = BorderStyle.None
        '   TextBox29.BorderStyle = BorderStyle.None
        '  TextBox30.BorderStyle = BorderStyle.None
        '  TextBox31.BorderStyle = BorderStyle.None
        'TextBox32.BorderStyle = BorderStyle.None
        '  TextBox33.BorderStyle = BorderStyle.None
        'TextBox34.BorderStyle = BorderStyle.None
        ' TextBox35.BorderStyle = BorderStyle.None
        '   TextBox36.BorderStyle = BorderStyle.None
        ' TextBox37.BorderStyle = BorderStyle.None
        '   TextBox38.BorderStyle = BorderStyle.None
        'TextBox39.BorderStyle = BorderStyle.None
        '   TextBox40.BorderStyle = BorderStyle.None
        'TextBox41.BorderStyle = BorderStyle.None
        '      TextBox42.BorderStyle = BorderStyle.None
        ' TextBox43.BorderStyle = BorderStyle.None
        '  TextBox44.BorderStyle = BorderStyle.None
        '  TextBox45.BorderStyle = BorderStyle.None
        'TextBox46.BorderStyle = BorderStyle.None
        'TextBox47.BorderStyle = BorderStyle.None
        '  TextBox48.BorderStyle = BorderStyle.None
        'TextBox49.BorderStyle = BorderStyle.None
        '  TextBox50.BorderStyle = BorderStyle.None
        '  TextBox51.BorderStyle = BorderStyle.None
        'TextBox52.BorderStyle = BorderStyle.None
        '   TextBox53.BorderStyle = BorderStyle.None
        '  TextBox54.BorderStyle = BorderStyle.None
        '   TextBox55.BorderStyle = BorderStyle.None



        '  TextBox20.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox21.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox22.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox23.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox24.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox25.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox26.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox27.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox28.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox29.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox30.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox31.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox32.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox33.BackColor = Color.FromArgb(240, 240, 240)
        '  TextBox34.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox35.BackColor = Color.FromArgb(240, 240, 240)
        '   TextBox36.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox37.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox38.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox39.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox40.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox41.BackColor = Color.FromArgb(240, 240, 240)
        '    TextBox42.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox43.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox44.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox45.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox46.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox47.BackColor = Color.FromArgb(240, 240, 240)
        '  TextBox48.BackColor = Color.FromArgb(240, 240, 240)
        'TextBox50.BackColor = Color.FromArgb(240, 240, 240)
        '  TextBox51.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox52.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox53.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox54.BackColor = Color.FromArgb(240, 240, 240)
        ' TextBox55.BackColor = Color.FromArgb(240, 240, 240)


        Label1.TextAlign = ContentAlignment.MiddleCenter
        Label2.TextAlign = ContentAlignment.MiddleCenter
        Label3.TextAlign = ContentAlignment.MiddleCenter
        Label4.TextAlign = ContentAlignment.MiddleLeft
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
        Label36.TextAlign = ContentAlignment.MiddleCenter
        Label37.TextAlign = ContentAlignment.MiddleCenter
        Label38.TextAlign = ContentAlignment.MiddleCenter
        Label39.TextAlign = ContentAlignment.MiddleCenter
        Label40.TextAlign = ContentAlignment.MiddleCenter

        Label41.TextAlign = ContentAlignment.MiddleCenter
        Label42.TextAlign = ContentAlignment.MiddleCenter
        Label43.TextAlign = ContentAlignment.MiddleCenter
        Label44.TextAlign = ContentAlignment.MiddleCenter
        Label45.TextAlign = ContentAlignment.MiddleCenter


    

        ' Label1.Parent = PictureBox1
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
        Label26.Parent = PictureBox1
        Label27.Parent = PictureBox1
        Label28.Parent = PictureBox1
        Label29.Parent = PictureBox1
        Label3.Parent = PictureBox1
        Label30.Parent = PictureBox1
        Label31.Parent = PictureBox1
        Label32.Parent = PictureBox1
        Label35.Parent = PictureBox1
        Label34.Parent = PictureBox1
        Label36.Parent = PictureBox1
        Label33.Parent = PictureBox1
        Label37.Parent = PictureBox1
        Label38.Parent = PictureBox1
        Label39.Parent = PictureBox1
        Label40.Parent = PictureBox1
        Label41.Parent = PictureBox1
        Label42.Parent = PictureBox1
        Label43.Parent = PictureBox1
        Label44.Parent = PictureBox1


        R1 = 0
        Ra = 5
        Rb = 5
        R = 0
        Dim Ac As Integer
        '    Open "acceler.dat" For Input As #1
        FileOpen(1, Application.StartupPath + "\acceler.dat", OpenMode.Input)
        Input(1, Ac)
        FileClose(1)

        Timer1.Enabled = True
        Timer2.Enabled = True

        Select Case W
            Case 1
                If Ac >= 10 Then
                    Timer2.Interval = 1000
                ElseIf Ac >= 5 And Ac < 10 Then
                    Timer2.Interval = 3000
                ElseIf Ac < 5 And Ac >= 2 Then
                    Timer2.Interval = 5000
                Else
                    Timer2.Interval = 10000
                End If
            Case 2
                If Ac >= 10 Then
                    Timer2.Interval = 2000
                ElseIf Ac >= 5 And Ac < 10 Then
                    Timer2.Interval = 4000
                ElseIf Ac < 5 And Ac >= 2 Then
                    Timer2.Interval = 10000
                Else
                    Timer2.Interval = 20000
                End If
            Case 3
                If Ac >= 10 Then
                    Timer2.Interval = 4000
                ElseIf Ac >= 5 And Ac < 10 Then
                    Timer2.Interval = 8000
                ElseIf Ac < 5 And Ac >= 2 Then
                    Timer2.Interval = 20000
                Else
                    Timer2.Interval = 40000
                End If
            Case 4
                If Ac >= 10 Then
                    Timer2.Interval = 5000
                ElseIf Ac >= 5 And Ac < 10 Then
                    Timer2.Interval = 10000
                ElseIf Ac >= 2 And Ac < 5 Then
                    Timer2.Interval = 25000
                Else
                    Timer2.Interval = 50000
                End If
            Case Else
        End Select

        '    Me.BackColor = System.Drawing.Color.FromArgb(100, 0, 5, 210)

        Label4.Visible = False


        'Open "selec.dat" For Input As #7
        FileOpen(7, Application.StartupPath + "\selec.dat", OpenMode.Input)
        For j = 0 To 2
            Input(7, Mb(j))
        Next j
        FileClose(7)

       
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        flow2.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' curves.Hide()
        curves.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        flow2.TopMost = False
        Me.TopMost = False
        Menubms.Show()
        Me.Hide()
        flow2.Hide()
       
    End Sub
    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Power.Show()

        '  Open "power.dat " For Input As #8
        FileOpen(8, Application.StartupPath + "\power.dat", OpenMode.Input)
        Dim i As Integer
        Dim abcc(5) As Single
        For i = 1 To 5
            Input(8, abcc(i))
            'Input(8, Wfvav, Wfcav, Wfrtn, Qcvav, Qccav)

        Next i
        FileClose(8)

        Wfvav = abcc(1)
        Wfcav = abcc(2)
        Wfrtn = abcc(3)
        Qcvav = abcc(4)
        Qccav = abcc(5)
        Wfvav1 = Wfvav / 3600.0#
        Wfcav1 = Wfcav / 3600.0#
        Wfrtn1 = Wfrtn / 3600.0#
        Qcvav1 = Qcvav / 1000.0#
        Qccav1 = Qccav / 1000.0#
        'Load(Power)
        Power.Close()
        Power.Show()
        Power.TextBox1.Text = Str$(Wfvav1)
        Power.TextBox2.Text = Str$(Wfcav1)
        Power.TextBox3.Text = Str$(Wfrtn1)
        Power.TextBox4.Text = Str$(Qcvav1)
        Power.TextBox5.Text = Str$(Qccav1)
        Power.TextBox1.Text = Format(Val(Power.TextBox1.Text), "###0.00")
        Power.TextBox2.Text = Format(Val(Power.TextBox2.Text), "###0.00")
        Power.TextBox3.Text = Format(Val(Power.TextBox3.Text), "###0.00")
        Power.TextBox4.Text = Format(Val(Power.TextBox4.Text), "###0.00")
        Power.TextBox5.Text = Format(Val(Power.TextBox5.Text), "###0.00")
    End Sub

    Private Sub Label33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label33.Click
        sele.Close()
        sele.Show()
    End Sub

    Private Sub Label34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label34.Click
        Chan.Close()
        Chan.Show()
        Chan.Label2.Text = "Proportional Gain for Supply Fan Realistic Control(%/C)"
        Chan.Label3.Text = "Integral Time for Supply Fan Realistic Control(Second)"

        M = 9
        n = 10

        'On Error GoTo MyError

        'Open "BMSin.dat" For Input As #1
        '  Dim SourceFile3 As String = GetCurrentDirectory() + "\bmsincopy.dat"
        '  Dim DestinationFile3 As String = GetCurrentDirectory() + "\bmsin.dat"

        FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(1, out(i))
        Next i
        FileClose(1)
        '  Shell("cmd /c copy bmsincopy.dat bmsin.dat")
        '  FileCopy(SourceFile3, DestinationFile3)
        'MyError:
        '    Resume

        Chan.Text1.Text = Str$(out(9))
        Chan.Text2.Text = Str$(out(10))

    End Sub

    Private Sub Label35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label35.Click
        Chan.Close()
        Chan.Show()

        Chan.Label2.Text = "Proportional Gain for CAV Coil Air Temperature Control(%/C)"
        Chan.Label3.Text = "Integral Time for CAV Coil Air Temperature Control(Second)"

        M = 34
        n = 35

        'On Error GoTo MyError

        'Open "BMSin.dat" For Input As #1
        ' Dim SourceFile3 As String = GetCurrentDirectory() + "\bmsincopy.dat"
        '  Dim DestinationFile3 As String = GetCurrentDirectory() + "\bmsin.dat"
        FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(1, out(i))
        Next i
        FileClose(1)
        'Shell("cmd /c copy bmsincopy.dat bmsin.dat")
        '  FileCopy(SourceFile3, DestinationFile3)
        'MyError:
        '  Resume

        Chan.Text1.Text = Str$(out(34))
        Chan.Text2.Text = Str$(out(35))

    End Sub

    Private Sub Label36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label36.Click
        Chan.Close()
        Chan.Show()

        Chan.Label2.Text = "Proportional Gain for VAV Coil Air Temperature Control(%/C)"
        Chan.Label3.Text = "Integral Time for VAV Coil Air Temperature Control(Second)"

        M = 7
        n = 8

        'On Error GoTo MyError
        ' Dim SourceFile3 As String = GetCurrentDirectory() + "\bmsincopy.dat"
        '  Dim DestinationFile3 As String = GetCurrentDirectory() + "\bmsin.dat"
        FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(1, out(i))
        Next i
        FileClose(1)
        'Shell("cmd /c copy bmsincopy.dat bmsin.dat")
        ' FileCopy(SourceFile3, DestinationFile3)
        'MyError:
        '    Resume

        Chan.Text1.Text = Str$(out(7))
        Chan.Text2.Text = Str$(out(8))

    End Sub

    Private Sub Label37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label37.Click
        sele2.Close()
        sele2.Show()
    End Sub

   

    Private Sub TextBox2_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label38.Click
        sele1.Close()
        sele1.Show()
    End Sub

    Private Sub TextBox43_Click(ByVal sender As Object, ByVal e As System.EventArgs)
       

    End Sub

    Private Sub Label39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label39.Click
        'Load(Chan)
        Chan.Close()
        Chan.Show()

        Chan.Label2.Text = "Proportional Gain for Outdoor Air Flow Rate Control(%/C)"
        Chan.Label3.Text = "Integral Time for Outdoor Air Flow Rate Control(Second)"

        M = 11
        n = 12

        'On Error GoTo MyError

        'Open "BMSin.dat" For Input As #1
        '   Dim SourceFile3 As String = GetCurrentDirectory() + "\bmsincopy.dat"
        '   Dim DestinationFile3 As String = GetCurrentDirectory() + "\bmsin.dat"
        FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(1, out(i))
        Next i
        FileClose(1)
        'Shell("cmd /c copy bmsincopy.dat bmsin.dat")
        ' FileCopy(SourceFile3, DestinationFile3)
        'MyError:
        '    Resume

        Chan.Text1.Text = Str$(out(11))
        Chan.Text2.Text = Str$(out(12))

    End Sub

    Private Sub Label40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label40.Click
        ' Load(chan3)
        chan3.Close()
        chan3.Show()

        chan3.Label2.Text = "Exfiltration flow setpoint(kg/s)"
        chan3.Label3.Text = "Proportional Gain for Return Fan Control(%/C)"
        chan3.Label4.Text = "Integral Time for Return Fan Control(Second)"

        M = 29
        n = 5
        m1 = 6

        'On Error GoTo MyError

        'Open "BMSin.dat" For Input As #1
        'Dim SourceFile3 As String = GetCurrentDirectory() + "\bmsincopy.dat"
        ' Dim DestinationFile3 As String = GetCurrentDirectory() + "\bmsin.dat"

        FileOpen(1, Application.StartupPath + "\BMSin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(1, out(i))
        Next i
        FileClose(1)
        'Shell("cmd /c copy bmsincopy.dat bmsin.dat")
        ' FileCopy(SourceFile3, DestinationFile3)
        'MyError:
        '    Resume

        chan3.Text1.Text = Str$(out(29))
        chan3.Text2.Text = Str$(out(5))
        chan3.Text3.Text = Str$(out(6))

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim NXIN(59), MXIN(59) As Single
        Dim a As Integer

        a = 0
        R1 = R1 + 1

        'Open "MXINa.dat" For Input As #2
        FileOpen(2, Application.StartupPath + "\MXINa.dat", OpenMode.Input)
        For k = 1 To 59
            Input(2, MXIN(k))
        Next k
        FileClose(2)

        On Error GoTo MyError

        'Open "BMSout.dat" For Input As #1
        ' Open "bmsin.dat" For Input As #9
        'Dim SourceFile2 As String = GetCurrentDirectory() + "\bmsout.dat"
        ' Dim DestinationFile2 As String = GetCurrentDirectory() + "\bmsoutcopy.dat"
        'FileCopy(SourceFile2, DestinationFile2)
        'Shell("cmd /c copy bmsout.dat bmsoutcopy.dat")
        FileOpen(1, Application.StartupPath + "\BMSout.dat", OpenMode.Input)
        For i = 1 To 59
            Input(1, NXIN(i))
        Next i
        FileClose(1)


        If a = 0 Then

            ' Open "MXINa.dat" For Output As #3
            FileOpen(3, Application.StartupPath + "\MXINa.dat", OpenMode.Output)
            For j = 1 To 59
                Print(3, NXIN(j))
            Next j
            FileClose(3)
        Else
            For j = 1 To 59
                NXIN(j) = MXIN(j)
            Next j
        End If


MyError:
        a = 1
        Resume Next

        Select Case W
            Case 1
                If R1 > 30 And NXIN(59) = 7200 Then
                    Beep()
                    Timer1.Enabled = False
                    Label4.Visible = True
                    'Open "power.dat " For Input As #8
                    FileOpen(8, Application.StartupPath + "\power.dat", OpenMode.Input)
                    Dim in1 As Integer
                    Dim abb(5) As Single
                    For in1 = 1 To 5
                        Input(8, abb(in1))

                    Next in1

                    FileClose(8)
                    'Input #8, Wfvav, Wfcav, Wfrtn, Qcvav, Qccav
                    Wfvav = abb(1)
                    Wfcav = abb(2)
                    Wfrtn = abb(3)
                    Qcvav = abb(4)
                    Qccav = abb(5)
                    ' FileClose(8)
                    Wfvav1 = Wfvav / 3600.0#
                    Wfcav1 = Wfcav / 3600.0#
                    Wfrtn1 = Wfrtn / 3600.0#
                    Qcvav1 = Qcvav / 1000.0#
                    Qccav1 = Qccav / 1000.0#
                    'Load(Power)
                    Power.Show()
                    Power.TextBox1.Text = Str$(Wfvav1)
                    Power.TextBox2.Text = Str$(Wfcav1)
                    Power.TextBox3.Text = Str$(Wfrtn1)
                    Power.TextBox4.Text = Str$(Qcvav1)
                    Power.TextBox5.Text = Str$(Qccav1)
                    Power.TextBox1.Text = Format(Val(Power.TextBox1.Text), "###0.00")
                    Power.TextBox2.Text = Format(Val(Power.TextBox2.Text), "###0.00")
                    Power.TextBox3.Text = Format(Val(Power.TextBox3.Text), "###0.00")
                    Power.TextBox4.Text = Format(Val(Power.TextBox4.Text), "###0.00")
                    Power.TextBox5.Text = Format(Val(Power.TextBox5.Text), "###0.00")
                    '关闭任务管理器的TRNExe.exe'
                    'killprocess(Label31.Text)


                    ' hact = OpenProcess(SYNCHRONIZE, False, iTask)
                    'retval = WaitForSingleObject(hact, 500)
                    Select Case Z
                        Case 1
                            AppActivate("DOIT3")
                        Case 5
                            AppActivate("DOIT9")
                        Case 9
                            AppActivate("DOIT15")
                        Case 13
                            AppActivate("DOIT21")
                        Case 17
                            AppActivate("DOIT27")
                        Case 21
                            AppActivate("DOIT33")
                        Case Else
                    End Select
                    ' retval = TerminateProcess(hact, 0)
                    '  retval = CloseHandle(hact)
                End If
            Case 2
                If R1 > 30 And NXIN(59) = 14400 Then
                    Beep()
                    Timer1.Enabled = False
                    Label4.Visible = True
                    'Open "power.dat " For Input As #8
                    FileOpen(8, Application.StartupPath + "\power.dat", OpenMode.Input)
                    Dim in11 As Integer
                    Dim abb1(5) As Single
                    For in11 = 1 To 5
                        Input(8, abb1(in11))

                    Next in11

                    FileClose(8)
                    'Input #8, Wfvav, Wfcav, Wfrtn, Qcvav, Qccav
                    '  Close #8
                    Wfvav = abb1(1)
                    Wfcav = abb1(2)
                    Wfrtn = abb1(3)
                    Qcvav = abb1(4)
                    Qccav = abb1(5)
                    Wfvav1 = Wfvav / 3600.0#
                    Wfcav1 = Wfcav / 3600.0#
                    Wfrtn1 = Wfrtn / 3600.0#
                    Qcvav1 = Qcvav / 1000.0#
                    Qccav1 = Qccav / 1000.0#
                    ' Load(Power)
                    Power.Show()
                    Power.TextBox1.Text = Str$(Wfvav1)
                    Power.TextBox2.Text = Str$(Wfcav1)
                    Power.TextBox3.Text = Str$(Wfrtn1)
                    Power.TextBox4.Text = Str$(Qcvav1)
                    Power.TextBox5.Text = Str$(Qccav1)
                    Power.TextBox1.Text = Format(Val(Power.TextBox1.Text), "###0.00")
                    Power.TextBox2.Text = Format(Val(Power.TextBox2.Text), "###0.00")
                    Power.TextBox3.Text = Format(Val(Power.TextBox3.Text), "###0.00")
                    Power.TextBox4.Text = Format(Val(Power.TextBox4.Text), "###0.00")
                    Power.TextBox5.Text = Format(Val(Power.TextBox5.Text), "###0.00")
                    '关闭任务管理器的TRNExe.exe'
                    ' killprocess(Label31.Text)

                    ' hact = OpenProcess(SYNCHRONIZE, False, iTask)
                    '  retval = WaitForSingleObject(hact, 500)
                    Select Case Z
                        Case 2
                            AppActivate("DOIT4")
                        Case 6
                            AppActivate("DOIT10")
                        Case 10
                            AppActivate("DOIT16")
                        Case 14
                            AppActivate("DOIT22")
                        Case 18
                            AppActivate("DOIT28")
                        Case 22
                            AppActivate("DOIT34")
                        Case Else
                    End Select
                    '   retval = TerminateProcess(hact, 0)
                    '   retval = CloseHandle(hact)
                End If
            Case 3
                If R1 > 30 And NXIN(59) = 28800 Then
                    Beep()
                    Timer1.Enabled = False
                    Label4.Visible = True
                    ' Open "power.dat " For Input As #8
                    FileOpen(8, Application.StartupPath + "\power.dat", OpenMode.Input)
                    Dim in12 As Integer
                    Dim abb2(5) As Single
                    For in12 = 1 To 5
                        Input(8, abb2(in12))

                    Next in12

                    FileClose(8)
                    ' Input #8, Wfvav, Wfcav, Wfrtn, Qcvav, Qccav
                    '  Close #8
                    Wfvav = abb2(1)
                    Wfcav = abb2(2)
                    Wfrtn = abb2(3)
                    Qcvav = abb2(4)
                    Qccav = abb2(5)
                    Wfvav1 = Wfvav / 3600.0#
                    Wfcav1 = Wfcav / 3600.0#
                    Wfrtn1 = Wfrtn / 3600.0#
                    Qcvav1 = Qcvav / 1000.0#
                    Qccav1 = Qccav / 1000.0#
                    ' Load(Power)
                    Power.Show()
                    Power.TextBox1.Text = Str$(Wfvav1)
                    Power.TextBox2.Text = Str$(Wfcav1)
                    Power.TextBox3.Text = Str$(Wfrtn1)
                    Power.TextBox4.Text = Str$(Qcvav1)
                    Power.TextBox5.Text = Str$(Qccav1)
                    Power.TextBox1.Text = Format(Val(Power.TextBox1.Text), "###0.00")
                    Power.TextBox2.Text = Format(Val(Power.TextBox2.Text), "###0.00")
                    Power.TextBox3.Text = Format(Val(Power.TextBox3.Text), "###0.00")
                    Power.TextBox4.Text = Format(Val(Power.TextBox4.Text), "###0.00")
                    Power.TextBox5.Text = Format(Val(Power.TextBox5.Text), "###0.00")
                  '关闭任务管理器的TRNExe.exe'
                    'killprocess(Label31.Text)
                    '   hact = OpenProcess(SYNCHRONIZE, False, iTask)
                    '   retval = WaitForSingleObject(hact, 500)
                    Select Case Z
                        Case 3
                            AppActivate("DOIT5")
                        Case 7
                            AppActivate("DOIT11")
                        Case 11
                            AppActivate("DOIT17")
                        Case 15
                            AppActivate("DOIT23")
                        Case 19
                            AppActivate("DOIT29")
                        Case 23
                            AppActivate("DOIT35")
                        Case Else
                    End Select
                    '   retval = TerminateProcess(hact, 0)
                    '    retval = CloseHandle(hact)
                End If
            Case 4
                If R1 > 30 And NXIN(59) = 43200 Then
                    Beep()
                    Timer1.Enabled = False
                    Label4.Visible = True
                    ' Open "power.dat " For Input As #8
                    'Input #8, Wfvav, Wfcav, Wfrtn, Qcvav, Qccav
                    'Close #8
                    FileOpen(8, Application.StartupPath + "\power.dat", OpenMode.Input)
                    Dim in13 As Integer
                    Dim abb3(5) As Single
                    For in13 = 1 To 5
                        Input(8, abb3(in13))

                    Next in13

                    FileClose(8)
                    'Input #8, Wfvav, Wfcav, Wfrtn, Qcvav, Qccav
                    '  Close #8
                    Wfvav = abb3(1)
                    Wfcav = abb3(2)
                    Wfrtn = abb3(3)
                    Qcvav = abb3(4)
                    Qccav = abb3(5)
                    Wfvav1 = Wfvav / 3600.0#
                    Wfcav1 = Wfcav / 3600.0#
                    Wfrtn1 = Wfrtn / 3600.0#
                    Qcvav1 = Qcvav / 1000.0#
                    Qccav1 = Qccav / 1000.0#
                    '  Load(Power)
                    Power.Show()
                    Power.TextBox1.Text = Str$(Wfvav1)
                    Power.TextBox2.Text = Str$(Wfcav1)
                    Power.TextBox3.Text = Str$(Wfrtn1)
                    Power.TextBox4.Text = Str$(Qcvav1)
                    Power.TextBox5.Text = Str$(Qccav1)
                    Power.TextBox1.Text = Format(Val(Power.TextBox1.Text), "###0.00")
                    Power.TextBox2.Text = Format(Val(Power.TextBox2.Text), "###0.00")
                    Power.TextBox3.Text = Format(Val(Power.TextBox3.Text), "###0.00")
                    Power.TextBox4.Text = Format(Val(Power.TextBox4.Text), "###0.00")
                    Power.TextBox5.Text = Format(Val(Power.TextBox5.Text), "###0.00")
                    '关闭任务管理器的TRNExe.exe'
                    '  killprocess(Label31.Text)

                    '   hact = OpenProcess(SYNCHRONIZE, False, iTask)
                    ' retval = WaitForSingleObject(hact, 500)


                    Select Case Z
                        Case 4
                            AppActivate("DOIT6")
                        Case 8
                            AppActivate("DOIT12")
                        Case 12
                            AppActivate("DOIT18")
                        Case 16
                            AppActivate("DOIT24")
                        Case 20
                            AppActivate("DOIT30")
                        Case 24
                            AppActivate("DOIT36")
                        Case Else
                    End Select
                    ' retval = TerminateProcess(hact, 0)
                    '  retval = CloseHandle(hact)
                End If
            Case Else
        End Select

        'TextBox23.Text = Str$(NXIN(36))
        Label7.Text = Str$(NXIN(36))
        l(36) = NXIN(36)
        ' TextBox24.Text = Str$(NXIN(38))
        Label32.Text = Str$(NXIN(38))
        l(38) = NXIN(38)
        '  TextBox31.Text = Str$(NXIN(39))
        Label12.Text = Str$(NXIN(39))
        l(39) = NXIN(39)
        'TextBox32.Text = Str$(NXIN(40))
        Label41.Text = Str$(NXIN(40))
        l(40) = NXIN(40)
        ' TextBox26.Text = Str$(NXIN(41))
        Label9.Text = Str$(NXIN(41))
        l(41) = NXIN(41)
        'TextBox34.Text = Str$(NXIN(43))
        Label42.Text = Str$(NXIN(43))
        l(43) = NXIN(43)
        '  TextBox22.Text = Str$(NXIN(44))
        Label6.Text = Str$(NXIN(44))
        l(44) = NXIN(44)
        '  TextBox48.Text = Str$(NXIN(45))
        Label22.Text = Str$(NXIN(45))
        l(45) = NXIN(45)
        ' TextBox35.Text = Str$(NXIN(46))
        Label43.Text = Str$(NXIN(46))
        l(46) = NXIN(46)
        '  TextBox47.Text = Str$(NXIN(47))
        Label21.Text = Str$(NXIN(47))
        l(47) = NXIN(47)
        '  TextBox39.Text = Str$(NXIN(48))
        Label15.Text = Str$(NXIN(48))
        l(48) = NXIN(48)
        'TextBox46.Text = Str$(NXIN(49))
        Label31.Text = Str$(NXIN(49))
        l(49) = NXIN(49)
        ' TextBox37.Text = Str$(NXIN(50))
        Label14.Text = Str$(NXIN(50))
        l(50) = NXIN(50)
        '  TextBox41.Text = Str$(NXIN(51))
        Label30.Text = Str$(NXIN(51))
        l(51) = NXIN(51)
        '   TextBox29.Text = Str$(NXIN(52))
        Label10.Text = Str$(NXIN(52))
        l(52) = NXIN(52)
        ' TextBox43.Text = Str$(NXIN(53))
        Label18.Text = Str$(NXIN(53))
        l(53) = NXIN(53)
        ' TextBox44.Text = Str$(NXIN(54))
        Label19.Text = Str$(NXIN(54))
        l(54) = NXIN(54)
        ' TextBox45.Text = Str$(NXIN(55))
        Label20.Text = Str$(NXIN(55))
        l(55) = NXIN(55)
        ' TextBox38.Text = Str$(NXIN(56))
        Label44.Text = Str$(NXIN(56))
        l(56) = NXIN(56)
        ' TextBox33.Text = Str$(NXIN(57))
        Label13.Text = Str$(NXIN(57))
        l(57) = NXIN(57)
        ' TextBox25.Text = Str$(NXIN(58))
        Label8.Text = Str$(NXIN(58))
        l(58) = NXIN(58)
        'TextBox20.Text = Str$(NXIN(6))
        Label3.Text = Str$(NXIN(6))
        l(6) = NXIN(6)
        '  TextBox21.Text = Str$(NXIN(5))
        Label5.Text = Str$(NXIN(5))
        l(5) = NXIN(5)
        ' TextBox42.Text = Str$(NXIN(4))
        Label17.Text = Str$(NXIN(4))
        l(4) = NXIN(4)
        'TextBox50.Text = Str$(NXIN(3))
        Label23.Text = Str$(NXIN(3))
        l(3) = NXIN(3)
        ' TextBox40.Text = Str$(NXIN(1))
        Label16.Text = Str$(NXIN(1))
        l(2) = NXIN(2)
        l(1) = NXIN(1)
        '  TextBox30.Text = Str$(NXIN(19))
        Label11.Text = Str$(NXIN(19))
        l(19) = NXIN(19)
        'TextBox53.Text = Str$(NXIN(18))
        Label24.Text = Str$(NXIN(18))
        l(18) = NXIN(18)
        'TextBox54.Text = Str$(NXIN(27))
        Label25.Text = Str$(NXIN(25))
        l(25) = NXIN(25)
        ' TextBox55.Text = Str$(NXIN(25))
        Label26.Text = Str$(NXIN(27))
        l(27) = NXIN(27)

        

        'On Error GoTo MyErrora

        ' Open "bmsin.dat" For Input As #9
        ' Dim SourceFile1 As String = GetCurrentDirectory() + "\bmsin.dat"
        '  Dim DestinationFile1 As String = GetCurrentDirectory() + "\bmsincopy.dat"


        FileOpen(9, Application.StartupPath + "\bmsin.dat", OpenMode.Input)
        For i = 1 To 37
            Input(9, out(i))
        Next i
        FileClose(9)
        'Shell("cmd /c copy bmsincopy.dat bmsin.dat")
        ' FileCopy(SourceFile1, DestinationFile1)
        'MyErrora:
        '  Resume

        If Mb(1) <= 1 Then
            out(28) = out(28)
        ElseIf Mb(1) <= 2 Then
            out(28) = NXIN(1)
        ElseIf Mb(1) <= 3 Then
            out(28) = NXIN(3)
        Else
            out(28) = NXIN(4)
        End If
        'Textbox52.Text = Str$(out(28))
        Label29.Text = Str$(out(28))
        If Mb(2) <= 1 Then
            out(30) = out(30)
        Else
            out(30) = NXIN(5)
        End If
        ' Textbox27.Text = Str$(out(30))
        Label27.Text = Str$(out(30))
        If Mb(0) <= 1 Then
            out(31) = out(31)
        Else
            out(31) = NXIN(6)
        End If
        'Textbox28.Text = Str$(out(31))
        Label28.Text = Str$(out(31))

        l(42) = NXIN(42)
        l(37) = NXIN(37)

        'Textbox51.Text = Str$(NXIN(59))
        Label45.Text = Str$(NXIN(59))
        l(59) = NXIN(59)

        'flow2.Textbox22.Text = Str$(NXIN(7))
        flow2.Label10.Text = Str$(NXIN(7))
        l(7) = NXIN(7)
        ' flow2.Textbox23.Text = Str$(NXIN(8))
        flow2.Label11.Text = Str$(NXIN(8))
        l(8) = NXIN(8)
        ' flow2.Textbox24.Text = Str$(NXIN(9))
        flow2.Label4.Text = Str$(NXIN(9))
        l(9) = NXIN(9)
        'flow2.Textbox25.Text = Str$(NXIN(10))
        flow2.Label12.Text = Str$(NXIN(10))
        l(10) = NXIN(10)
        'flow2.Textbox26.Text = Str$(NXIN(11))
        flow2.Label6.Text = Str$(NXIN(11))
        l(11) = NXIN(11)
        '  flow2.Textbox27.Text = Str$(NXIN(12))
        flow2.Label1.Text = Str$(NXIN(12))
        l(12) = NXIN(12)
        ' flow2.Textbox28.Text = Str$(NXIN(13))
        flow2.Label8.Text = Str$(NXIN(13))
        l(13) = NXIN(13)
        ' flow2.Textbox29.Text = Str$(NXIN(14))
        flow2.Label2.Text = Str$(NXIN(14))
        l(14) = NXIN(14)
        'flow2.Textbox30.Text = Str$(NXIN(15))
        flow2.Label14.Text = Str$(NXIN(15))
        l(15) = NXIN(15)
        ' flow2.Textbox31.Text = Str$(NXIN(16))
        flow2.Label16.Text = Str$(NXIN(16))
        l(16) = NXIN(16)
        'flow2.Textbox32.Text = Str$(NXIN(17))
        flow2.Label17.Text = Str$(NXIN(17))
        l(17) = NXIN(17)
        ' flow2.Textbox14.Text = Str$(NXIN(20))
        flow2.Label9.Text = Str$(NXIN(20))
        l(20) = NXIN(20)
        ' flow2.Textbox18.Text = Str$(NXIN(21))
        flow2.Label19.Text = Str$(NXIN(21))
        l(21) = NXIN(21)
        ' flow2.Textbox33.Text = Str$(NXIN(22))
        flow2.Label3.Text = Str$(NXIN(22))
        l(22) = NXIN(22)
        'flow2.Textbox34.Text = Str$(NXIN(23))
        flow2.Label22.Text = Str$(NXIN(23))
        l(23) = NXIN(23)
        ' flow2.Textbox35.Text = Str$(NXIN(24))
        flow2.Label5.Text = Str$(NXIN(24))
        l(24) = NXIN(24)
        ' flow2.Textbox37.Text = Str$(NXIN(26))
        flow2.Label7.Text = Str$(NXIN(26))
        l(26) = NXIN(26)
        '  flow2.Textbox39.Text = Str$(NXIN(28))
        flow2.Label28.Text = Str$(NXIN(28))
        l(28) = NXIN(28)
        '   flow2.Textbox40.Text = Str$(NXIN(29))
        flow2.Label29.Text = Str$(NXIN(29))
        l(29) = NXIN(29)
        '  flow2.Textbox41.Text = Str$(NXIN(30))
        flow2.Label30.Text = Str$(NXIN(30))
        l(30) = NXIN(30)
        'flow2.Textbox42.Text = Str$(NXIN(31))
        flow2.Label31.Text = Str$(NXIN(31))
        l(31) = NXIN(31)
        ' flow2.Textbox43.Text = Str$(NXIN(32))
        flow2.Label32.Text = Str$(NXIN(32))
        l(32) = NXIN(32)
        ' flow2.Textbox44.Text = Str$(NXIN(33))
        flow2.Label33.Text = Str$(NXIN(33))
        l(33) = NXIN(33)
        ' flow2.Textbox45.Text = Str$(NXIN(34))
        flow2.Label34.Text = Str$(NXIN(34))
        l(34) = NXIN(34)
        ' flow2.Textbox46.Text = Str$(NXIN(35))
        flow2.Label35.Text = Str$(NXIN(35))
        l(35) = NXIN(35)
        flow2.Label45.Text = Str$(NXIN(59))


        '两位有效数字
        ' TextBox20.Text = Format(Val(TextBox20.Text), "###0.00")
        Label3.Text = Format(Val(Label3.Text), "###0.00")
        Label3.Text = Label3.Text & "°C"
        ' TextBox21.Text = Format(Val(TextBox21.Text), "###0.00")
        Label5.Text = Format(Val(Label5.Text), "###0.00")
        Label5.Text = Label5.Text & "Pa"
        '   TextBox22.Text = Format(Val(TextBox22.Text), "###0.00")
        Label6.Text = Format(Val(Label6.Text), "###0.00")
        Label6.Text = Label6.Text & "kg/s"
        'TextBox23.Text = Format(Val(TextBox23.Text), "###0.00")
        Label7.Text = Format(Val(Label7.Text), "###0.00")
        Label7.Text = Label7.Text & "°C"
        ' TextBox24.Text = Format(Val(TextBox24.Text), "###0.00")
        Label32.Text = Format(Val(Label32.Text), "###0.00")
        ' TextBox25.Text = Format(Val(TextBox25.Text), "###0.00")
        Label8.Text = Format(Val(Label8.Text), "###0.00")
        Label8.Text = Label8.Text & "kg/s"
        ' TextBox26.Text = Format(Val(TextBox26.Text), "###0.00")
        Label9.Text = Format(Val(Label9.Text), "###0.00")
        Label9.Text = Label9.Text & "Pa"
        ' TextBox27.Text = Format(Val(TextBox27.Text), "###0.00")
        Label27.Text = Format(Val(Label27.Text), "###0.00")
        Label27.Text = Label27.Text & "Pa"

        '  TextBox28.Text = Format(Val(TextBox28.Text), "###0.00")
        Label28.Text = Format(Val(Label28.Text), "###0.00")
        Label28.Text = Label28.Text & "°C"
        ' TextBox29.Text = Format(Val(TextBox29.Text), "###0.00")
        Label10.Text = Format(Val(Label10.Text), "###0.00")
        Label10.Text = Label10.Text & "kg/kg"
        '  TextBox30.Text = Format(Val(TextBox30.Text), "###0.00")
        Label11.Text = Format(Val(Label11.Text), "###0.00")
        Label11.Text = Label11.Text & "kW"
        '  TextBox31.Text = Format(Val(TextBox31.Text), "###0.00")
        Label12.Text = Format(Val(Label12.Text), "###0.00")
        Label12.Text = Label12.Text & "°C"
        ' TextBox32.Text = Format(Val(TextBox32.Text), "###0.00")
        Label41.Text = Format(Val(Label41.Text), "###0.00")
        '  TextBox33.Text = Format(Val(TextBox33.Text), "###0.00")
        Label13.Text = Format(Val(Label13.Text), "###0.00")
        Label13.Text = Label13.Text & "kg/s"
        ' TextBox34.Text = Format(Val(TextBox34.Text), "###0.00")
        Label42.Text = Format(Val(Label42.Text), "###0.00")
        'TextBox35.Text = Format(Val(TextBox35.Text), "###0.00")
        Label43.Text = Format(Val(Label43.Text), "###0.00")
        'TextBox36.Text =  Format(Val(TextBox26.Text), “###0.00”)
        ' TextBox37.Text = Format(Val(TextBox37.Text), "###0.00")
        Label14.Text = Format(Val(Label14.Text), "###0.00")
        Label14.Text = Label14.Text & "ppm"
        ' TextBox38.Text = Format(Val(TextBox38.Text), "###0.00")
        Label44.Text = Format(Val(Label44.Text), "###0.00")
        '  TextBox39.Text = Format(Val(TextBox39.Text), "###0.00")
        Label15.Text = Format(Val(Label15.Text), "###0.00")
        Label15.Text = Label15.Text & "kW"
        ' TextBox40.Text = Format(Val(TextBox40.Text), "###0.00")
        Label16.Text = Format(Val(Label16.Text), "###0.00")
        Label16.Text = Label16.Text & "kg/s"
        '  TextBox41.Text = Format(Val(TextBox41.Text), "###0.00")
        Label30.Text = Format(Val(Label30.Text), "###0.00")
        Label30.Text = Label30.Text & "°C"
        ' TextBox42.Text = Format(Val(TextBox42.Text), "###0.00")
        Label17.Text = Format(Val(Label17.Text), "###0.00")
        Label17.Text = Label17.Text & "kg/s"
        'TextBox43.Text = Format(Val(TextBox43.Text), "###0.00")
        Label18.Text = Format(Val(Label18.Text), "###0.00")
        Label18.Text = Label18.Text & "kg/kg"
        'TextBox44.Text = Format(Val(TextBox44.Text), "###0.00")
        Label19.Text = Format(Val(Label19.Text), "###0.00")
        Label19.Text = Label19.Text & "°C"
        '  TextBox45.Text = Format(Val(TextBox45.Text), "###0.00")
        Label20.Text = Format(Val(Label20.Text), "###0.00")
        Label20.Text = Label20.Text & "ppm"
        '  TextBox46.Text = Format(Val(TextBox46.Text), "###0.00")
        Label31.Text = Format(Val(Label31.Text), "###0.00")
        'TextBox47.Text = Format(Val(TextBox47.Text), "###0.00")
        Label21.Text = Format(Val(Label21.Text), "###0.00")
        Label21.Text = Label21.Text & "kg/s"
        'TextBox48.Text = Format(Val(TextBox48.Text), "###0.00")
        Label22.Text = Format(Val(Label22.Text), "###0.00")
        Label22.Text = Label22.Text & "kg/s"
        '  TextBox49.Text =  Format(Val(TextBox29.Text), “###0.00”)
        '  TextBox50.Text = Format(Val(TextBox50.Text), "###0.00")
        Label23.Text = Format(Val(Label23.Text), "###0.00")
        Label23.Text = Label23.Text & "kg/s"
        '   TextBox51.Text = Format(Val(TextBox51.Text), "###0.00")
        'TextBox52.Text = Format(Val(TextBox52.Text), "###0.00")
        Label29.Text = Format(Val(Label29.Text), "###0.00")
        Label29.Text = Label29.Text & "kg/s"
        '   TextBox53.Text = Format(Val(TextBox53.Text), "###0.00")
        Label24.Text = Format(Val(Label24.Text), "###0.00")
        Label24.Text = Label24.Text & "kW"
        'TextBox54.Text = Format(Val(TextBox54.Text), "###0.00")
        Label25.Text = Format(Val(Label25.Text), "###0.00")
        Label25.Text = Label25.Text & "kW"
        'TextBox55.Text = Format(Val(TextBox55.Text), "###0.00")
        Label26.Text = Format(Val(Label26.Text), "###0.00")
        Label26.Text = Label26.Text & "kW"
        ' flow2.TextBox14.Text = Format(Val(flow2.TextBox14.Text), "###0.00")
        flow2.Label9.Text = Format(Val(flow2.Label9.Text), "###0.00")
        flow2.Label9.Text = flow2.Label9.Text & "kg/s"
        'flow2.TextBox18.Text = Format(Val(flow2.TextBox18.Text), "###0.00")
        flow2.Label19.Text = Format(Val(flow2.Label19.Text), "###0.00")
        flow2.Label19.Text = flow2.Label19.Text & "kg/s"
        ' flow2.TextBox22.Text = Format(Val(flow2.TextBox22.Text), "###0.00")
        flow2.Label10.Text = Format(Val(flow2.Label10.Text), "###0.00")
        flow2.Label10.Text = flow2.Label10.Text & "°C"
        '   flow2.TextBox23.Text = Format(Val(flow2.TextBox23.Text), "###0.00")
        flow2.Label11.Text = Format(Val(flow2.Label11.Text), "###0.00")
        flow2.Label11.Text = flow2.Label11.Text & "°C"
        ' flow2.TextBox24.Text = Format(Val(flow2.TextBox24.Text), "###0.00")
        flow2.Label4.Text = Format(Val(flow2.Label4.Text), "###0.00")
        flow2.Label4.Text = flow2.Label4.Text & "°C"
        ' flow2.TextBox25.Text = Format(Val(flow2.TextBox25.Text), "###0.00")
        flow2.Label12.Text = Format(Val(flow2.Label12.Text), "###0.00")
        flow2.Label12.Text = flow2.Label12.Text & "°C"
        'flow2.TextBox26.Text = Format(Val(flow2.TextBox26.Text), "###0.00")
        flow2.Label6.Text = Format(Val(flow2.Label6.Text), "###0.00")
        flow2.Label6.Text = flow2.Label6.Text & "°C"
        ' flow2.TextBox27.Text = Format(Val(flow2.TextBox27.Text), "###0.00")
        flow2.Label1.Text = Format(Val(flow2.Label1.Text), "###0.00")
        flow2.Label1.Text = flow2.Label1.Text & "°C"
        ' flow2.TextBox28.Text = Format(Val(flow2.TextBox28.Text), "###0.00")
        flow2.Label8.Text = Format(Val(flow2.Label8.Text), "###0.00")
        flow2.Label8.Text = flow2.Label8.Text & "°C"
        ' flow2.TextBox29.Text = Format(Val(flow2.TextBox29.Text), "###0.00")
        flow2.Label2.Text = Format(Val(flow2.Label2.Text), "###0.00")
        flow2.Label2.Text = flow2.Label2.Text & "°C"
        ' flow2.TextBox30.Text = Format(Val(flow2.TextBox30.Text), "###0.00")
        flow2.Label14.Text = Format(Val(flow2.Label14.Text), "###0.00")
        flow2.Label14.Text = flow2.Label14.Text & "kg/kg"
        'flow2.TextBox31.Text = Format(Val(flow2.TextBox31.Text), "###0.00")
        flow2.Label16.Text = Format(Val(flow2.Label16.Text), "###0.00")
        flow2.Label16.Text = flow2.Label16.Text & "ppm"
        '   flow2.TextBox32.Text = Format(Val(flow2.TextBox32.Text), "###0.00")
        flow2.Label17.Text = Format(Val(flow2.Label17.Text), "###0.00")
        flow2.Label17.Text = flow2.Label17.Text & "ppm"
        ' flow2.TextBox33.Text = Format(Val(flow2.TextBox33.Text), "###0.00")
        flow2.Label3.Text = Format(Val(flow2.Label3.Text), "###0.00")
        flow2.Label3.Text = flow2.Label3.Text & "kg/s"
        'flow2.TextBox34.Text = Format(Val(flow2.TextBox34.Text), "###0.00")
        flow2.Label22.Text = Format(Val(flow2.Label22.Text), "###0.00")
        flow2.Label22.Text = flow2.Label22.Text & "kg/s"
        'flow2.TextBox35.Text = Format(Val(flow2.TextBox35.Text), "###0.00")
        flow2.Label5.Text = Format(Val(flow2.Label5.Text), "###0.00")
        flow2.Label5.Text = flow2.Label5.Text & "kg/s"
        '  flow2.TextBox37.Text = Format(Val(flow2.TextBox37.Text), "###0.00")
        flow2.Label7.Text = Format(Val(flow2.Label7.Text), "###0.00")
        flow2.Label7.Text = flow2.Label7.Text & "kg/s"
        '  flow2.TextBox39.Text = Format(Val(flow2.TextBox39.Text), "###0.00")
        flow2.Label28.Text = Format(Val(flow2.Label28.Text), "###0.00")
        ' flow2.TextBox40.Text = Format(Val(flow2.TextBox40.Text), "###0.00")
        flow2.Label29.Text = Format(Val(flow2.Label29.Text), "###0.00")
        ' flow2.TextBox41.Text = Format(Val(flow2.TextBox41.Text), "###0.00")
        flow2.Label30.Text = Format(Val(flow2.Label30.Text), "###0.00")
        ' flow2.TextBox42.Text = Format(Val(flow2.TextBox42.Text), "###0.00")
        flow2.Label31.Text = Format(Val(flow2.Label31.Text), "###0.00")
        ' flow2.TextBox43.Text = Format(Val(flow2.TextBox43.Text), "###0.00")
        flow2.Label32.Text = Format(Val(flow2.Label32.Text), "###0.00")
        ' flow2.TextBox44.Text = Format(Val(flow2.TextBox44.Text), "###0.00")
        flow2.Label33.Text = Format(Val(flow2.Label33.Text), "###0.00")
        '  flow2.TextBox45.Text = Format(Val(flow2.TextBox45.Text), "###0.00")
        flow2.Label34.Text = Format(Val(flow2.Label34.Text), "###0.00")
        '  flow2.TextBox46.Text = Format(Val(flow2.TextBox46.Text), "###0.00")
        flow2.Label35.Text = Format(Val(flow2.Label35.Text), "###0.00")
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Dim i, j As Integer

        R1 = R1 + 1

        ' frmspread.Load()
        ' frmspread.Visible = True
        frmspread.Hide()
        'frmspread.Show()
        'frmspread.Visible = False

        For j = 1 To 399
            frmspread.grdspread.set_ColWidth(j, 2000)
        Next j

        For k = 1 To 59
            frmspread.grdspread.Col = 0
            frmspread.grdspread.Row = k
            frmspread.grdspread.Text = k
        Next k
        Dim Ro As Integer
        If R > 798 Then
            Ro = R - 798
            frmspread.grdspread.Col = Ro
            For i = 1 To 59
                j = i + 118
                frmspread.grdspread.Row = j
                frmspread.grdspread.Text = Str$(l(i))
            Next i
        ElseIf R > 399 Then
            Ro = R - 399
            frmspread.grdspread.Col = Ro
            For i = 1 To 59
                j = i + 59
                frmspread.grdspread.Row = j
                frmspread.grdspread.Text = Str$(l(i))
            Next i
        Else
            frmspread.grdspread.Col = R
            For i = 1 To 59
                frmspread.grdspread.Row = i
                frmspread.grdspread.Text = Str$(l(i))
            Next i
        End If

        Select Case W
            Case 1
                If R1 > 30 And l(59) = 7200 Then
                    Timer2.Enabled = False
                End If
            Case 2
                If R1 > 30 And l(59) = 14400 Then
                    Timer2.Enabled = False
                End If
            Case 3
                If R1 > 30 And l(59) = 28800 Then
                    Timer2.Enabled = False
                End If
            Case 4
                If R1 > 30 And l(59) = 43200 Then
                    Timer2.Enabled = False
                End If
            Case Else
        End Select

        R = R + 1

    End Sub
    Private Sub Label19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label19.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 44

            'Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Temperature of fresh air (°C)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 44

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of fresh air (°C)"
        End If
    End Sub

    Private Sub Label21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label21.Click
        If Fcurve = 0 Then
            Vmax = 6.0#
            Vmin = 0.0#
            TK = 47

            'Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Flow rate of fresh air (m3/s)"

        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 6.0#
            Vmin = 0.0#
            TK = 47

            ' Load(curve)
            ' curve.Close()
            curve.Show()
            curve.Label19.Text = "Flow rate of fresh air (m3/s)"

        End If
    End Sub

    Private Sub Label18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label18.Click
        If Fcurve = 0 Then
            Vmax = 0.02
            Vmin = 0.0#
            TK = 43

            'Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Humidity of fresh air (kg/kg)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 0.02
            Vmin = 0.0#
            TK = 43

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Humidity of fresh air (kg/kg)"
        End If
    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        If Fcurve = 0 Then
            Vmax = 8.0#
            Vmin = 0.0#
            TK = 22

            ' Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Total supply air flow rate (kg/s)"
        ElseIf Fcurve = 1 Then
            curve.Close()
            Vmax = 8.0#
            Vmin = 0.0#
            TK = 22

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Total supply air flow rate (kg/s)"
        End If
    End Sub

    Private Sub Label26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label26.Click
        If Fcurve = 0 Then
            Vmax = 10.0#
            Vmin = -50.0#
            TK = 54
            ' Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "CAV total heat transfer from the air (kw)"

        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 10.0#
            Vmin = -50.0#
            TK = 54
            'Load(curve)
            curve.Show()
            curve.Label19.Text = "CAV total heat transfer from the air (kw)"
        End If
    End Sub

    Private Sub Label25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label25.Click
        If Fcurve = 0 Then
            Vmax = 10.0#
            Vmin = -50.0#
            TK = 55
            '  Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "VAV total heat transfer from the air (kw)"

        ElseIf Fcurve = 1 Then

            '  Unload(curve)
            curve.Close()
            Vmax = 10.0#
            Vmin = -50.0#
            TK = 55
            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV total heat transfer from the air (kw)"
        End If
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click
        If Fcurve = 0 Then
            Vmax = 10.0#
            Vmin = 0.0#
            TK = 30

            '   Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Electrical power consumption of CAV supply fan (kw)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 10.0#
            Vmin = 0.0#
            TK = 30

            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Electrical power consumption of CAV supply fan (kw)"
        End If
    End Sub

    Private Sub Label24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label24.Click
        If Fcurve = 0 Then
            Vmax = 15.0#
            Vmin = 0.0#
            TK = 53

            '  Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Electrical power consumption of VAV supply fan (kw)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 15.0#
            Vmin = 0.0#
            TK = 53

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Electrical power consumption of VAV supply fan (kw)"
        End If
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click
        If Fcurve = 0 Then
            Vmax = 5.0#
            Vmin = 0.0#
            TK = 25

            'Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "CAV air flow rate (kg/s)"
        ElseIf Fcurve = 1 Then
            'Unload(curve)
            curve.Close()
            Vmax = 5.0#
            Vmin = 0.0#
            TK = 25

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "CAV air flow rate (kg/s)"
        End If
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 23
            ' Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "CAV cooling coil outlet temperature (°C)"

        ElseIf Fcurve = 1 Then
            'Unload(curve)
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 23

            '  Load(curve)

            curve.Show()
            curve.Label19.Text = "CAV cooling coil outlet temperature (°C)"
        End If
    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click
        If Fcurve = 0 Then
            Vmax = 6.0#
            Vmin = 0.0#
            TK = 33

            'Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "VAV air flow rate (kg/s)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 6.0#
            Vmin = 0.0#
            TK = 33

            'Load(curve)
            '   curve.Close()
            curve.Show()
            curve.Label19.Text = "VAV air flow rate (kg/s)"
        End If
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 31

            '   Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "VAV cooling coil outlet temperature (°C)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 31

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV cooling coil outlet temperature (°C)"
        End If
    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click
        If Fcurve = 0 Then
            Vmax = 1500.0#
            Vmin = 0.0#
            TK = 26

            ' Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Static pressure (Pa)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1500.0#
            Vmin = 0.0#
            TK = 26

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Static pressure (Pa)"
        End If
    End Sub

    Private Sub Label22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label22.Click
        If Fcurve = 0 Then
            Vmax = 10.0#
            Vmin = 0.0#
            TK = 48
            '    Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Total return air flow rate (kg/s)"

        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 10.0#
            Vmin = 0.0#
            TK = 48
            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Total return air flow rate (kg/s)"
        End If
    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click
        If Fcurve = 0 Then
            Vmax = 10.0#
            Vmin = 0.0#
            TK = 39

            '  Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Electrical power consumption of return fan (kw)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 10.0#
            Vmin = 0.0#
            TK = 39

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Electrical power consumption of return fan (kw)"
        End If
    End Sub

    Private Sub Label30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label30.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 41

            '    Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Temperature of return air (°C)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 41

            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "Temperature of return air (°C)"
        End If
    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click
        If Fcurve = 0 Then
            Vmax = 1500.0#
            Vmin = 0.0#
            TK = 37
            ' Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "CO2 concentration of return air (ppm)"

        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1500.0#
            Vmin = 0.0#
            TK = 37
            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "CO2 concentration of return air (ppm)"
        End If
    End Sub

    Private Sub Label20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label20.Click
        If Fcurve = 0 Then
            Vmax = 500.0#
            Vmin = 0.0#
            TK = 45

            ' Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "CO2 concentration of fresh air (ppm)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 500.0#
            Vmin = 0.0#
            TK = 45

            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "CO2 concentration of fresh air (ppm)"
        End If
    End Sub

    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label16.Click
        If Fcurve = 0 Then
            Vmax = 8.0#
            Vmin = 0.0#
            TK = 40

            ' Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Actually used fresh air flow rate setpoint (kg/s)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 8.0#
            Vmin = 0.0#
            TK = 40

            'Load(curve)
            curve.Show()
            curve.Label19.Text = "Actually used fresh air flow rate setpoint (kg/s)"
        End If
    End Sub

    Private Sub Label23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label23.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 50
            curve.Close()
            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Fresh air flow rate setpoint by DVC only (kg/s)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 50

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Fresh air flow rate setpoint by DVC only (kg/s)"
        End If
    End Sub

    Private Sub Label17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label17.Click
        If Fcurve = 0 Then
            Vmax = 8.0#
            Vmin = 0.0#
            TK = 42
            curve.Close()
            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Fresh air flow setpoint by enthalpy only (kg/s)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 8.0#
            Vmin = 0.0#
            TK = 42

            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Fresh air flow setpoint by enthalpy only (kg/s)"
        End If
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        If Fcurve = 0 Then
            Vmax = 0.02
            Vmin = 0.0#
            TK = 29
            curve.Close()
            'Load(curve)
            curve.Show()
            curve.Label19.Text = "Humidity of return air (kg/kg)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 0.02
            Vmin = 0.0#
            TK = 29

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Humidity of return air (kg/kg)"
        End If
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        If Fcurve = 0 Then
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 20
            curve.Close()
            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Optimal AHU supply air temperature setpoint (°C)"
        ElseIf Fcurve = 1 Then
            curve.Close()
            Vmax = 30.0#
            Vmin = 0.0#
            TK = 20

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Optimal AHU supply air temperature setpoint (°C)"
        End If
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        If Fcurve = 0 Then
            Vmax = 800.0#
            Vmin = 0.0#
            TK = 21
            curve.Close()
            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Optimal static pressure setpoint (Pa)"
        ElseIf Fcurve = 1 Then
            'Unload(curve)
            curve.Close()
            Vmax = 800.0#
            Vmin = 0.0#
            TK = 21

            'Load(curve)
            curve.Show()
            curve.Label19.Text = "Optimal static pressure setpoint (Pa)"
        End If
    End Sub

    Private Sub Label31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label31.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 46
            curve.Close()
            'Load(curve)
            curve.Show()
            curve.Label19.Text = "Outdoor air damper position demand (0-1)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 46

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Outdoor air damper position demand (0-1)"
        End If
    End Sub

    Private Sub Label32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label32.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 24
            curve.Close()
            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Actually used CAV cooling coil outlet damper position demand (0-1)"

        ElseIf Fcurve = 1 Then
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 24

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Actually used CAV cooling coil outlet damper position demand (0-1)"
        End If
    End Sub

    Private Sub Label41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label41.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 32
            curve.Close()
            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV cooling coil outlet damper position demand (0-1)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 32

            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV cooling coil outlet damper position demand (0-1)"
        End If
    End Sub

    Private Sub Label42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label42.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 34
            curve.Close()
            '   Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV fan control demand (0-1)"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 34

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "VAV fan control demand (0-1)"
        End If
    End Sub

    Private Sub Label43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label43.Click
        If Fcurve = 0 Then
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 35
            curve.Close()
            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Return fan control demand (angle)"

        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 1.0#
            Vmin = 0.0#
            TK = 35

            ' Load(curve)
            curve.Show()
            curve.Label19.Text = "Return fan control demand (angle)"
        End If
    End Sub

    Private Sub Label44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label44.Click
        If Fcurve = 0 Then
            Vmax = 200.0#
            Vmin = 0.0#
            TK = 38
            '   Load(curve)
            curve.Close()
            curve.Show()
            curve.Label19.Text = "Number of people"
        ElseIf Fcurve = 1 Then
            '  Unload(curve)
            curve.Close()
            Vmax = 200.0#
            Vmin = 0.0#
            TK = 38

            '  Load(curve)
            curve.Show()
            curve.Label19.Text = "Number of people"
        End If
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Me.Show()
        ' flow2.TopMost = False
        '  Me.TopMost = False
        flow2.Hide()
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        flow2.Show()
        ' flow2.TopMost = True
        Me.Hide()
        ' flow2.TopMost = False
    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub Label4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        ' Power.Show()

        '  Open "power.dat " For Input As #8
        FileOpen(8, Application.StartupPath + "\power.dat", OpenMode.Input)
        Dim i As Integer
        Dim abcc(5) As Single
        For i = 1 To 5
            Input(8, abcc(i))
            'Input(8, Wfvav, Wfcav, Wfrtn, Qcvav, Qccav)

        Next i
        FileClose(8)

        Wfvav = abcc(1)
        Wfcav = abcc(2)
        Wfrtn = abcc(3)
        Qcvav = abcc(4)
        Qccav = abcc(5)
        Wfvav1 = Wfvav / 3600.0#
        Wfcav1 = Wfcav / 3600.0#
        Wfrtn1 = Wfrtn / 3600.0#
        Qcvav1 = Qcvav / 1000.0#
        Qccav1 = Qccav / 1000.0#
        'Load(Power)
        Power.Show()
        Power.TextBox1.Text = Str$(Wfvav1)
        Power.TextBox2.Text = Str$(Wfcav1)
        Power.TextBox3.Text = Str$(Wfrtn1)
        Power.TextBox4.Text = Str$(Qcvav1)
        Power.TextBox5.Text = Str$(Qccav1)
        Power.TextBox1.Text = Format(Val(Power.TextBox1.Text), "###0.00")
        Power.TextBox2.Text = Format(Val(Power.TextBox2.Text), "###0.00")
        Power.TextBox3.Text = Format(Val(Power.TextBox3.Text), "###0.00")
        Power.TextBox4.Text = Format(Val(Power.TextBox4.Text), "###0.00")
        Power.TextBox5.Text = Format(Val(Power.TextBox5.Text), "###0.00")
    End Sub
End Class
