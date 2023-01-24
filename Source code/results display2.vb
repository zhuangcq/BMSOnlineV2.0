Public Class results_display2

    Private Sub results_display2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = False
        'Me.ControlBox = False
        MaximizeBox = False
        MinimizeBox = True
        'Me.Width = 966
        ' Me.Height = 660
        Me.Height = My.Computer.Screen.WorkingArea.Height
        Me.Width = My.Computer.Screen.WorkingArea.Width
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.Panel1.Width = My.Computer.Screen.WorkingArea.Width
        Me.Panel1.Height = My.Computer.Screen.WorkingArea.Width

        Dim x1, y1, x2, y2, a_width, b_height, c_font, c_font1, c_font2 As Integer

        '窗口大小 a为宽，b为高
        a_width = 998
        b_height = 660

        x1 = 39
        y1 = 8
        x2 = 856
        y2 = 46
        Label1.Text = "There are two online graphics showing the curves of selected variables (the maximum number of variables that can beshown in each graphic is 5)," & vbCrLf & "you could select the variables you like to monitor: GraphicA and GraphicB"

        ' c_font = CInt(Me.Height / b_height * 12)
        '  Me.Label1.Width = x2 / a_width * Me.Width
        '  Me.Label1.Height = y2 / b_height * Me.Height
        Me.Label1.Location = New Point((Me.Width - Label1.Width) / 4, y1 / b_height * Me.Height)

        c_font1 = CInt(Me.Height / b_height * 12)
        c_font2 = CInt(Me.Width / a_width * 12)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If

        Label1.Font = New System.Drawing.Font("arial narrow", c_font, FontStyle.Italic)


        x1 = 212
        y1 = 70
        x2 = 256
        y2 = 42
        Label3.Text = " Select variables showing on GraphicA " & vbCrLf & "(The maximum number of variables is 5)"
        Me.Label3.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 10)
        Me.Label3.Width = x2 / a_width * Me.Width
        Me.Label3.Height = y2 / b_height * Me.Height
        c_font1 = CInt(Me.Height / b_height * 10)
        c_font2 = CInt(Me.Width / a_width * 10)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If
        '     Label3.Font = New System.Drawing.Font(" arial narrow", c_font, FontStyle.Italic)
        Label3.Font = New System.Drawing.Font("arial narrow", c_font, FontStyle.Italic)

        x1 = 705
        y1 = 70
        x2 = 256
        y2 = 42
        Label2.Text = " Select variables showing on GraphicB " & vbCrLf & "(The maximum number of variables is 5)"
        Me.Label2.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 10)
        Me.Label2.Width = x2 / a_width * Me.Width
        Me.Label2.Height = y2 / b_height * Me.Height
        c_font1 = CInt(Me.Height / b_height * 10)
        c_font2 = CInt(Me.Width / a_width * 10)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If
        Label2.Font = New System.Drawing.Font("arial narrow", c_font, FontStyle.Italic)

        x1 = 24
        y1 = 120
        x2 = 251
        y2 = 17
        Me.CheckBox1.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox1.Width = x2 / a_width * Me.Width
        Me.CheckBox1.Height = y2 / b_height * Me.Height
        CheckBox1.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 277
        y2 = 17
        Me.CheckBox2.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox2.Width = x2 / a_width * Me.Width
        Me.CheckBox2.Height = y2 / b_height * Me.Height
        CheckBox2.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 277
        y2 = 17
        Me.CheckBox3.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox3.Width = x2 / a_width * Me.Width
        Me.CheckBox3.Height = y2 / b_height * Me.Height
        CheckBox3.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 172
        y2 = 17
        Me.CheckBox4.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox4.Width = x2 / a_width * Me.Width
        Me.CheckBox4.Height = y2 / b_height * Me.Height
        CheckBox4.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 172
        y2 = 17
        Me.CheckBox5.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox5.Width = x2 / a_width * Me.Width
        Me.CheckBox5.Height = y2 / b_height * Me.Height
        CheckBox5.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 172
        y2 = 17
        Me.CheckBox6.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox6.Width = x2 / a_width * Me.Width
        Me.CheckBox6.Height = y2 / b_height * Me.Height
        CheckBox6.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 172
        y2 = 17
        Me.CheckBox7.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox7.Width = x2 / a_width * Me.Width
        Me.CheckBox7.Height = y2 / b_height * Me.Height
        CheckBox7.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 172
        y2 = 17
        Me.CheckBox8.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox8.Width = x2 / a_width * Me.Width
        Me.CheckBox8.Height = y2 / b_height * Me.Height
        CheckBox8.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)


        x1 = 24
        y1 = y1 + 20
        x2 = 213
        y2 = 17
        Me.CheckBox9.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox9.Width = x2 / a_width * Me.Width
        Me.CheckBox9.Height = y2 / b_height * Me.Height
        CheckBox9.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 172
        y2 = 17
        Me.CheckBox10.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox10.Width = x2 / a_width * Me.Width
        Me.CheckBox10.Height = y2 / b_height * Me.Height
        CheckBox10.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)




        x1 = 24
        y1 = y1 + 20
        x2 = 213
        y2 = 17
        Me.CheckBox11.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox11.Width = x2 / a_width * Me.Width
        Me.CheckBox11.Height = y2 / b_height * Me.Height
        CheckBox11.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox12.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox12.Width = x2 / a_width * Me.Width
        Me.CheckBox12.Height = y2 / b_height * Me.Height
        CheckBox12.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        Me.CheckBox13.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox13.Width = x2 / a_width * Me.Width
        Me.CheckBox13.Height = y2 / b_height * Me.Height
        CheckBox13.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox14.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox14.Width = x2 / a_width * Me.Width
        Me.CheckBox14.Height = y2 / b_height * Me.Height
        CheckBox14.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox15.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox15.Width = x2 / a_width * Me.Width
        Me.CheckBox15.Height = y2 / b_height * Me.Height
        CheckBox15.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox16.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox16.Width = x2 / a_width * Me.Width
        Me.CheckBox16.Height = y2 / b_height * Me.Height
        CheckBox16.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)


        x1 = 24
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox17.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox17.Width = x2 / a_width * Me.Width
        Me.CheckBox17.Height = y2 / b_height * Me.Height
        CheckBox17.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox18.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox18.Width = x2 / a_width * Me.Width
        Me.CheckBox18.Height = y2 / b_height * Me.Height
        CheckBox18.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox19.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox19.Width = x2 / a_width * Me.Width
        Me.CheckBox19.Height = y2 / b_height * Me.Height
        CheckBox19.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 24
        y1 = y1 + 20
        x2 = 221
        y2 = 17
        Me.CheckBox20.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox20.Width = x2 / a_width * Me.Width
        Me.CheckBox20.Height = y2 / b_height * Me.Height
        CheckBox20.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 24
        y1 = y1 + 20
        x2 = 194
        y2 = 17
        Me.CheckBox21.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox21.Width = x2 / a_width * Me.Width
        Me.CheckBox21.Height = y2 / b_height * Me.Height
        CheckBox21.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 24
        y1 = y1 + 20
        x2 = 227
        y2 = 17
        Me.CheckBox22.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox22.Width = x2 / a_width * Me.Width
        Me.CheckBox22.Height = y2 / b_height * Me.Height
        CheckBox22.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = 120
        x2 = 251
        y2 = 17
        Me.CheckBox61.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox61.Width = x2 / a_width * Me.Width
        Me.CheckBox61.Height = y2 / b_height * Me.Height
        CheckBox61.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 196
        y2 = 17
        Me.CheckBox62.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox62.Width = x2 / a_width * Me.Width
        Me.CheckBox62.Height = y2 / b_height * Me.Height
        CheckBox62.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)


        x1 = 374
        y1 = y1 + 20
        x2 = 135
        y2 = 17
        Me.CheckBox63.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox63.Width = x2 / a_width * Me.Width
        Me.CheckBox63.Height = y2 / b_height * Me.Height
        CheckBox63.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 135
        y2 = 17
        Me.CheckBox64.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox64.Width = x2 / a_width * Me.Width
        Me.CheckBox64.Height = y2 / b_height * Me.Height
        CheckBox64.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox65.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox65.Width = x2 / a_width * Me.Width
        Me.CheckBox65.Height = y2 / b_height * Me.Height
        CheckBox65.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)


        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox66.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox66.Width = x2 / a_width * Me.Width
        Me.CheckBox66.Height = y2 / b_height * Me.Height
        CheckBox66.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox67.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox67.Width = x2 / a_width * Me.Width
        Me.CheckBox67.Height = y2 / b_height * Me.Height
        CheckBox67.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox68.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox68.Width = x2 / a_width * Me.Width
        Me.CheckBox68.Height = y2 / b_height * Me.Height
        CheckBox68.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox69.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox69.Width = x2 / a_width * Me.Width
        Me.CheckBox69.Height = y2 / b_height * Me.Height
        CheckBox69.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox70.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox70.Width = x2 / a_width * Me.Width
        Me.CheckBox70.Height = y2 / b_height * Me.Height
        CheckBox70.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox71.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox71.Width = x2 / a_width * Me.Width
        Me.CheckBox71.Height = y2 / b_height * Me.Height
        CheckBox71.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)



        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox72.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox72.Width = x2 / a_width * Me.Width
        Me.CheckBox72.Height = y2 / b_height * Me.Height
        CheckBox72.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox73.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox73.Width = x2 / a_width * Me.Width
        Me.CheckBox73.Height = y2 / b_height * Me.Height
        CheckBox73.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox74.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox74.Width = x2 / a_width * Me.Width
        Me.CheckBox74.Height = y2 / b_height * Me.Height
        CheckBox74.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox75.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox75.Width = x2 / a_width * Me.Width
        Me.CheckBox75.Height = y2 / b_height * Me.Height
        CheckBox75.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox76.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox76.Width = x2 / a_width * Me.Width
        Me.CheckBox76.Height = y2 / b_height * Me.Height
        CheckBox76.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox77.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox77.Width = x2 / a_width * Me.Width
        Me.CheckBox77.Height = y2 / b_height * Me.Height
        CheckBox77.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox78.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox78.Width = x2 / a_width * Me.Width
        Me.CheckBox78.Height = y2 / b_height * Me.Height
        CheckBox78.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox79.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox79.Width = x2 / a_width * Me.Width
        Me.CheckBox79.Height = y2 / b_height * Me.Height
        CheckBox79.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox80.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox80.Width = x2 / a_width * Me.Width
        Me.CheckBox80.Height = y2 / b_height * Me.Height
        CheckBox80.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 374
        y1 = y1 + 20
        x2 = 259
        y2 = 17
        Me.CheckBox81.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox81.Width = x2 / a_width * Me.Width
        Me.CheckBox81.Height = y2 / b_height * Me.Height
        CheckBox81.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 714
        y1 = 120
        x2 = 211
        y2 = 17
        Me.CheckBox31.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox31.Width = x2 / a_width * Me.Width
        Me.CheckBox31.Height = y2 / b_height * Me.Height
        CheckBox31.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 714
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox32.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox32.Width = x2 / a_width * Me.Width
        Me.CheckBox32.Height = y2 / b_height * Me.Height
        CheckBox32.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 714
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox33.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox33.Width = x2 / a_width * Me.Width
        Me.CheckBox33.Height = y2 / b_height * Me.Height
        CheckBox33.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 714
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox34.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox34.Width = x2 / a_width * Me.Width
        Me.CheckBox34.Height = y2 / b_height * Me.Height
        CheckBox34.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 714
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox35.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox35.Width = x2 / a_width * Me.Width
        Me.CheckBox35.Height = y2 / b_height * Me.Height
        CheckBox35.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 714
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox36.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox36.Width = x2 / a_width * Me.Width
        Me.CheckBox36.Height = y2 / b_height * Me.Height
        CheckBox36.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 714
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox37.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox37.Width = x2 / a_width * Me.Width
        Me.CheckBox37.Height = y2 / b_height * Me.Height
        CheckBox37.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 714
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox38.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox38.Width = x2 / a_width * Me.Width
        Me.CheckBox38.Height = y2 / b_height * Me.Height
        CheckBox38.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 714
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox39.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox39.Width = x2 / a_width * Me.Width
        Me.CheckBox39.Height = y2 / b_height * Me.Height
        CheckBox39.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 714
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox40.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox40.Width = x2 / a_width * Me.Width
        Me.CheckBox40.Height = y2 / b_height * Me.Height
        CheckBox40.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 714
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox41.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox41.Width = x2 / a_width * Me.Width
        Me.CheckBox41.Height = y2 / b_height * Me.Height
        CheckBox41.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 714
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox42.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox42.Width = x2 / a_width * Me.Width
        Me.CheckBox42.Height = y2 / b_height * Me.Height
        CheckBox42.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
        x1 = 714
        y1 = y1 + 20
        x2 = 211
        y2 = 17
        Me.CheckBox43.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        Me.CheckBox43.Width = x2 / a_width * Me.Width
        Me.CheckBox43.Height = y2 / b_height * Me.Height
        CheckBox43.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)


        x1 = 485
        y1 = 589
        x2 = 118
        y2 = 39
        Me.Button1.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        c_font1 = CInt(Me.Height / b_height * 8.25)
        c_font2 = CInt(Me.Width / a_width * 8.25)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If

        Me.Button1.Width = x2 / a_width * Me.Width
        Me.Button1.Height = y2 / b_height * Me.Height
        Button1.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 624
        y1 = 589
        x2 = 118
        y2 = 39
        Me.Button2.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        c_font1 = CInt(Me.Height / b_height * 8.25)
        c_font2 = CInt(Me.Width / a_width * 8.25)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If
        Me.Button2.Width = x2 / a_width * Me.Width
        Me.Button2.Height = y2 / b_height * Me.Height
        Button2.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)


        x1 = 756
        y1 = 589
        x2 = 86
        y2 = 39
        Me.Button3.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        c_font1 = CInt(Me.Height / b_height * 8.25)
        c_font2 = CInt(Me.Width / a_width * 8.25)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If
        Me.Button3.Width = x2 / a_width * Me.Width
        Me.Button3.Height = y2 / b_height * Me.Height
        Button3.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 860
        y1 = 589
        x2 = 86
        y2 = 39

        Me.Button4.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        c_font1 = CInt(Me.Height / b_height * 8.25)
        c_font2 = CInt(Me.Width / a_width * 8.25)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If
        Me.Button4.Width = x2 / a_width * Me.Width
        Me.Button4.Height = y2 / b_height * Me.Height
        Button4.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)



        x1 = 7
        y1 = 115
        x2 = 670
        y2 = 448
        Me.PictureBox1.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        c_font1 = CInt(Me.Height / b_height * 8.25)
        c_font2 = CInt(Me.Width / a_width * 8.25)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If
        Me.PictureBox1.Width = x2 / a_width * Me.Width
        Me.PictureBox1.Height = y2 / b_height * Me.Height
        Me.PictureBox1.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)

        x1 = 698
        y1 = 115
        x2 = 248
        y2 = 448
        Me.PictureBox2.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        c_font = CInt(Me.Height / b_height * 8.25)
        c_font1 = CInt(Me.Height / b_height * 8.25)
        c_font2 = CInt(Me.Width / a_width * 8.25)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If
        Me.PictureBox2.Width = x2 / a_width * Me.Width
        Me.PictureBox2.Height = y2 / b_height * Me.Height
        Me.PictureBox2.Font = New System.Drawing.Font("  Georgia", c_font, FontStyle.Regular)
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Z17 = 1
            Counter = Counter + 1
        ElseIf CheckBox1.Checked = False Then
            Z17 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If

    End Sub


    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Z18 = 1
            Counter = Counter + 1
        ElseIf CheckBox2.Checked = False Then
            Z18 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub




    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Z19 = 1
            Counter = Counter + 1
        ElseIf CheckBox3.Checked = False Then
            Z19 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            Z20 = 1
            Counter = Counter + 1
        ElseIf CheckBox4.Checked = False Then
            Z20 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            Z21 = 1
            Counter = Counter + 1
        ElseIf CheckBox5.Checked = False Then
            Z21 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            Z22 = 1
            Counter = Counter + 1
        ElseIf CheckBox6.Checked = False Then
            Z22 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            Z23 = 1
            Counter = Counter + 1
        ElseIf CheckBox7.Checked = False Then
            Z23 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            Z24 = 1
            Counter = Counter + 1
        ElseIf CheckBox8.Checked = False Then
            Z24 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = True Then
            Z27 = 1
            Counter = Counter + 1
        ElseIf CheckBox9.Checked = False Then
            Z27 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = True Then
            Z26 = 1
            Counter = Counter + 1
        ElseIf CheckBox10.Checked = False Then
            Z26 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If

    End Sub

    Private Sub CheckBox11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.Checked = True Then
            Z25 = 1
            Counter = Counter + 1
        ElseIf CheckBox11.Checked = False Then
            Z25 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox12.CheckedChanged
        If CheckBox12.Checked = True Then
            Z28 = 1
            Counter = Counter + 1
        ElseIf CheckBox12.Checked = False Then
            Z28 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox13_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox13.CheckedChanged
        If CheckBox13.Checked = True Then
            Z29 = 1
            Counter = Counter + 1
        ElseIf CheckBox13.Checked = False Then
            Z29 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox14_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox14.CheckedChanged
        If CheckBox14.Checked = True Then
            Z30 = 1
            Counter = Counter + 1
        ElseIf CheckBox14.Checked = False Then
            Z30 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox15_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox15.CheckedChanged
        If CheckBox15.Checked = True Then
            Z31 = 1
            Counter = Counter + 1
        ElseIf CheckBox15.Checked = False Then
            Z31 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox16_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox16.CheckedChanged
        If CheckBox16.Checked = True Then
            Z32 = 1
            Counter = Counter + 1
        ElseIf CheckBox16.Checked = False Then
            Z32 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox17_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox17.CheckedChanged
        If CheckBox17.Checked = True Then
            Z33 = 1
            Counter = Counter + 1
        ElseIf CheckBox17.Checked = False Then
            Z33 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox18_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox18.CheckedChanged
        If CheckBox18.Checked = True Then
            Z34 = 1
            Counter = Counter + 1
        ElseIf CheckBox18.Checked = False Then
            Z34 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox19_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox19.CheckedChanged
        If CheckBox19.Checked = True Then
            Z35 = 1
            Counter = Counter + 1
        ElseIf CheckBox19.Checked = False Then
            Z35 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox20_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox20.CheckedChanged
        If CheckBox20.Checked = True Then
            Z38 = 1
            Counter = Counter + 1
        ElseIf CheckBox20.Checked = False Then
            Z38 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub
    Private Sub CheckBox21_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox21.CheckedChanged
        If CheckBox21.Checked = True Then
            Z5 = 1
            Counter = Counter + 1
        ElseIf CheckBox21.Checked = False Then
            Z5 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox22_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox22.CheckedChanged
        If CheckBox22.Checked = True Then
            Z6 = 1
            Counter = Counter + 1
        ElseIf CheckBox22.Checked = False Then
            Z6 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub
    Private Sub CheckBox61_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox61.CheckedChanged
        If CheckBox61.Checked = True Then
            Z48 = 1
            Counter = Counter + 1
        ElseIf CheckBox61.Checked = False Then
            Z48 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox62_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox62.CheckedChanged
        If CheckBox62.Checked = True Then
            Z49 = 1
            Counter = Counter + 1
        ElseIf CheckBox62.Checked = False Then
            Z49 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox63_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox63.CheckedChanged
        If CheckBox63.Checked = True Then
            Z57 = 1
            Counter = Counter + 1
        ElseIf CheckBox63.Checked = False Then
            Z57 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox64_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox64.CheckedChanged
        If CheckBox64.Checked = True Then
            Z58 = 1
            Counter = Counter + 1
        ElseIf CheckBox64.Checked = False Then
            Z58 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox65_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox65.CheckedChanged
        If CheckBox65.Checked = True Then
            Z15 = 1
            Counter = Counter + 1
        ElseIf CheckBox65.Checked = False Then
            Z15 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox66_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox66.CheckedChanged
        If CheckBox66.Checked = True Then
            Z52 = 1
            Counter = Counter + 1
        ElseIf CheckBox66.Checked = False Then
            Z52 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox67_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox67.CheckedChanged
        If CheckBox67.Checked = True Then
            Z53 = 1
            Counter = Counter + 1
        ElseIf CheckBox67.Checked = False Then
            Z53 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox68_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox68.CheckedChanged
        If CheckBox68.Checked = True Then
            Z16 = 1
            Counter = Counter + 1
        ElseIf CheckBox68.Checked = False Then
            Z16 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox69_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox69.CheckedChanged
        If CheckBox69.Checked = True Then
            Z50 = 1
            Counter = Counter + 1
        ElseIf CheckBox69.Checked = False Then
            Z50 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox70_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox70.CheckedChanged
        If CheckBox70.Checked = True Then
            Z55 = 1
            Counter = Counter + 1
        ElseIf CheckBox70.Checked = False Then
            Z55 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox71_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox71.CheckedChanged
        If CheckBox71.Checked = True Then
            Z1 = 1
            Counter = Counter + 1
        ElseIf CheckBox71.Checked = False Then
            Z1 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox72_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox72.CheckedChanged
        If CheckBox72.Checked = True Then
            Z40 = 1
            Counter = Counter + 1
        ElseIf CheckBox72.Checked = False Then
            Z40 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox73_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox73.CheckedChanged
        If CheckBox73.Checked = True Then
            Z43 = 1
            Counter = Counter + 1
        ElseIf CheckBox73.Checked = False Then
            Z43 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox74_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox74.CheckedChanged
        If CheckBox74.Checked = True Then
            Z44 = 1
            Counter = Counter + 1
        ElseIf CheckBox74.Checked = False Then
            Z44 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox75_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox75.CheckedChanged
        If CheckBox75.Checked = True Then
            Z45 = 1
            Counter = Counter + 1
        ElseIf CheckBox75.Checked = False Then
            Z45 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox76_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox76.CheckedChanged
        If CheckBox76.Checked = True Then
            Z46 = 1
            Counter = Counter + 1
        ElseIf CheckBox76.Checked = False Then
            Z46 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox77_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox77.CheckedChanged
        If CheckBox77.Checked = True Then
            Z47 = 1
            Counter = Counter + 1
        ElseIf CheckBox77.Checked = False Then
            Z47 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox78_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox78.CheckedChanged
        If CheckBox18.Checked = True Then
            Z41 = 1
            Counter = Counter + 1
        ElseIf CheckBox78.Checked = False Then
            Z41 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox79_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox79.CheckedChanged
        If CheckBox79.Checked = True Then
            Z56 = 1
            Counter = Counter + 1
        ElseIf CheckBox79.Checked = False Then
            Z56 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox80_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox80.CheckedChanged
        If CheckBox80.Checked = True Then
            Z3 = 1
            Counter = Counter + 1
        ElseIf CheckBox80.Checked = False Then
            Z3 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox81_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox81.CheckedChanged
        If CheckBox81.Checked = True Then
            Z4 = 1
            Counter = Counter + 1
        ElseIf CheckBox81.Checked = False Then
            Z4 = 0
            If Counter = 0 Then
                Counter = 0
            ElseIf Counter > 0 Then
                Counter = Counter - 1
            End If
        Else
        End If
    End Sub
    Private Sub CheckBox31_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox31.CheckedChanged
        If CheckBox31.Checked = True Then
            Z7 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox31.Checked = False Then
            Z7 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox32_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox32.CheckedChanged
        If CheckBox32.Checked = True Then
            Z8 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox32.Checked = False Then
            Z8 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox33_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox33.CheckedChanged
        If CheckBox33.Checked = True Then
            Z9 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox33.Checked = False Then
            Z9 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox34_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox34.CheckedChanged
        If CheckBox34.Checked = True Then
            Z10 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox34.Checked = False Then
            Z10 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox35_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox35.CheckedChanged
        If CheckBox35.Checked = True Then
            Z11 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox35.Checked = False Then
            Z11 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox36_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox36.CheckedChanged
        If CheckBox36.Checked = True Then
            Z12 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox36.Checked = False Then
            Z12 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox37_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox37.CheckedChanged
        If CheckBox37.Checked = True Then
            Z13 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox37.Checked = False Then
            Z13 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox38_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox38.CheckedChanged
        If CheckBox38.Checked = True Then
            Z14 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox38.Checked = False Then
            Z14 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox39_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox39.CheckedChanged
        If CheckBox39.Checked = True Then
            Z36 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox39.Checked = False Then
            Z36 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox40_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox40.CheckedChanged
        If CheckBox40.Checked = True Then
            Z37 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox40.Checked = False Then
            Z37 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox41_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox41.CheckedChanged
        If CheckBox41.Checked = True Then
            Z39 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox41.Checked = False Then
            Z39 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox42_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox42.CheckedChanged
        If CheckBox42.Checked = True Then
            Z51 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox42.Checked = False Then
            Z51 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub CheckBox43_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox43.CheckedChanged
        If CheckBox43.Checked = True Then
            Z54 = 1
            Counter1 = Counter1 + 1
        ElseIf CheckBox43.Checked = False Then
            Z54 = 0
            If Counter1 = 0 Then
                Counter1 = 0
            ElseIf Counter1 > 0 Then
                Counter1 = Counter1 - 1
            End If
        Else
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
        CheckBox6.Checked = False
        CheckBox7.Checked = False
        CheckBox8.Checked = False
        CheckBox9.Checked = False
        CheckBox10.Checked = False
        CheckBox11.Checked = False
        CheckBox12.Checked = False
        CheckBox13.Checked = False
        CheckBox14.Checked = False
        CheckBox15.Checked = False
        CheckBox16.Checked = False
        CheckBox17.Checked = False
        CheckBox18.Checked = False
        CheckBox19.Checked = False
        CheckBox20.Checked = False
        CheckBox21.Checked = False
        CheckBox22.Checked = False
        CheckBox61.Checked = False
        CheckBox62.Checked = False
        CheckBox63.Checked = False
        CheckBox64.Checked = False
        CheckBox65.Checked = False
        CheckBox66.Checked = False
        CheckBox67.Checked = False
        CheckBox68.Checked = False
        CheckBox69.Checked = False
        CheckBox70.Checked = False
        CheckBox71.Checked = False
        CheckBox72.Checked = False
        CheckBox73.Checked = False
        CheckBox74.Checked = False
        CheckBox75.Checked = False
        CheckBox76.Checked = False
        CheckBox77.Checked = False
        CheckBox78.Checked = False
        CheckBox79.Checked = False
        CheckBox80.Checked = False
        CheckBox81.Checked = False
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        CheckBox31.Checked = False
        CheckBox32.Checked = False
        CheckBox33.Checked = False
        CheckBox34.Checked = False
        CheckBox35.Checked = False
        CheckBox36.Checked = False
        CheckBox37.Checked = False
        CheckBox38.Checked = False
        CheckBox39.Checked = False
        CheckBox40.Checked = False
        CheckBox41.Checked = False
        CheckBox42.Checked = False
        CheckBox43.Checked = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Flag1(1) = 0
        Flag1(2) = 0
        Flag1(3) = 0
        Flag1(4) = 0
        Flag1(5) = 0

        Counterb = Counter1
        Countera = Counter
        If Counter1 > 5 Then
            warn2.Close()
            warn2.Show()
        ElseIf Counter1 <= 5 Then

            If Counter > 5 Then
                warn2.Close()
                warn2.Show()

            ElseIf Counter <= 5 Then


                'Else
                curves.Show()
            End If


        

            ' parselect1.Close()
            ' parselect4.Show()
            ' Me.Close()

            'parselect3.Hide




            '  Else
        End If

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Menubms.Show()
        Me.Close()
    End Sub

  
    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class