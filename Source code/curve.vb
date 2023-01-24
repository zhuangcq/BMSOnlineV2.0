Public Class curve
    Dim frmWidth
    Dim frmHeight
    Dim Yo, Ho As Single
    Dim TInter As Single
    Dim Rm1 As Single
    Dim Ac As Integer
    Dim a As Single
    Dim b As Single
    Dim G As Single
    Dim CurrentX, CurrentY, CurrentXo, CurrentYo As Single

    Public Sub drawline(ByRef n11 As Object)

        If Ra1 > 798 Then
            Ro1 = Ra1 - 798
            frmspread.grdspread.Col = Ro1
            frmspread.grdspread.Row = 118 + n11
            Ver = Val(frmspread.grdspread.Text)
        ElseIf Ra1 > 399 Then
            Ro1 = Ra1 - 399
            frmspread.grdspread.Col = Ro1
            frmspread.grdspread.Row = 59 + n11
            Ver = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra1
            frmspread.grdspread.Row = n11
            Ver = Val(frmspread.grdspread.Text)
        End If
        Dim currentx2, currenty2, currentx3, currenty3 As Single
        currentx2 = Me.Width * 0.05
        currenty3 = Me.Height * 0.1
        currenty2 = currenty3 + Me.Height * 0.75
        currentx3 = currentx2 + Me.Width * 0.8

        If Ra1 > 5 Then
            ' CurrentX = (Hn1 / 3600.0#) * 6000.0# / period + 660.0#
            'CurrentY = 3720.0# - (Ver - Vmin) * 3000.0# / (Vmax - Vmin)

            '  CurrentXo = (Ho1 / 3600.0#) * 6000.0# / period + 660.0#
            '  CurrentYo = 3720.0# - (Vero - Vmin) * 3000.0# / (Vmax - Vmin)
            'Line (CurrentXo, CurrentYo)-(CurrentX, CurrentY), QBColor(12)
            Dim gra As Graphics = Me.CreateGraphics
            Dim myPen As Pen = New Pen(Color.Red, 3)

            CurrentX = (Hn1 / 3600.0#) * (currentx3 - currentx2) / period + currentx2
            CurrentY = currenty2 - (Ver - Vmin) * (currenty2 - currenty3) / (Vmax - Vmin)

            CurrentXo = (Ho1 / 3600.0#) * (currentx3 - currentx2) / period + currentx2
            CurrentYo = currenty2 - (Vero - Vmin) * (currenty2 - currenty3) / (Vmax - Vmin)
            gra.DrawLine(myPen, CurrentXo, CurrentYo, CurrentX, CurrentY)

        Else
        End If
        Vero = Ver
    End Sub
    Public Sub readH()

        If Ra1 > 798 Then
            Ro1 = Ra1 - 798
            frmspread.grdspread.Col = Ro1
            frmspread.grdspread.Row = 177
            Hn1 = Val(frmspread.grdspread.Text)
        ElseIf Ra1 > 399 Then
            Ro1 = Ra1 - 399
            frmspread.grdspread.Col = Ro1
            frmspread.grdspread.Row = 118
            Hn1 = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra1
            frmspread.grdspread.Row = 59
            Hn1 = Val(frmspread.grdspread.Text)
        End If
    End Sub


    Private Sub curve_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Me.ControlBox = False
        '设置文本框大小
        Me.TopMost = False
        MaximizeBox = False
        MinimizeBox = False
        Me.Height = My.Computer.Screen.WorkingArea.Height * 0.55
        Me.Width = My.Computer.Screen.WorkingArea.Width * 0.55
        ' Me.WindowState = FormWindowState.Maximized
        ' Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.Top = (My.Computer.Screen.Bounds.Height - Me.Height) / 2
        Me.Left = (My.Computer.Screen.Bounds.Width - Me.Width) / 2

        Label1.TextAlign = ContentAlignment.MiddleRight
        Label2.TextAlign = ContentAlignment.MiddleRight
        Label3.TextAlign = ContentAlignment.MiddleRight
        Label4.TextAlign = ContentAlignment.MiddleRight
        Label5.TextAlign = ContentAlignment.MiddleRight
        Label6.TextAlign = ContentAlignment.MiddleRight
        Label7.TextAlign = ContentAlignment.MiddleLeft
        Label8.TextAlign = ContentAlignment.MiddleLeft
        Label9.TextAlign = ContentAlignment.MiddleLeft
        Label10.TextAlign = ContentAlignment.MiddleLeft
        Label11.TextAlign = ContentAlignment.MiddleLeft
        Label12.TextAlign = ContentAlignment.MiddleLeft
        Label13.TextAlign = ContentAlignment.MiddleLeft
        Label14.TextAlign = ContentAlignment.MiddleLeft
        Label15.TextAlign = ContentAlignment.MiddleLeft
        Label16.TextAlign = ContentAlignment.MiddleLeft
        Label17.TextAlign = ContentAlignment.MiddleLeft
        Label18.TextAlign = ContentAlignment.MiddleLeft
        Label19.TextAlign = ContentAlignment.Middlecenter
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


        Dim x1, y1, x2, y2, a_width, b_height, c_font As Integer

        '窗口大小 a为宽，b为高
        a_width = 792
        b_height = 451
        Dim currentx2, currenty2, currentx3, currenty3, gapx, gapy As Single
        currentx2 = Me.Width * 0.05
        currenty3 = Me.Height * 0.1
        currenty2 = currenty3 + Me.Height * 0.75
        currentx3 = currentx2 + Me.Width * 0.8


        gapx = (currentx3 - currentx2) / 10
        gapy = (currenty2 - currenty3) / 5

   
        x2 = 80
        y2 = 20
        Me.Label1.Width = x2 / a_width * Me.Width
        Me.Label1.Height = y2 / b_height * Me.Height
        Me.Label2.Width = x2 / a_width * Me.Width
        Me.Label2.Height = y2 / b_height * Me.Height
        Me.Label3.Width = x2 / a_width * Me.Width
        Me.Label3.Height = y2 / b_height * Me.Height
        Me.Label4.Width = x2 / a_width * Me.Width
        Me.Label4.Height = y2 / b_height * Me.Height
        Me.Label5.Width = x2 / a_width * Me.Width
        Me.Label5.Height = y2 / b_height * Me.Height
        Me.Label6.Width = x2 / a_width * Me.Width
        Me.Label6.Height = y2 / b_height * Me.Height
        Me.Label1.Location = New Point(Width * 0.055 - Label1.Width, currenty3 * 0.8)
        Me.Label2.Location = New Point(Width * 0.055 - Label2.Width, currenty3 * 0.8 + gapy)
        Me.Label3.Location = New Point(Width * 0.055 - Label3.Width, currenty3 * 0.8 + gapy * 2)
        Me.Label4.Location = New Point(Width * 0.055 - Label4.Width, currenty3 * 0.8 + gapy * 3)
        Me.Label5.Location = New Point(Width * 0.055 - Label5.Width, currenty3 * 0.8 + gapy * 4)
        Me.Label6.Location = New Point(Width * 0.055 - Label6.Width, currenty3 * 0.8 + gapy * 5)
        c_font = CInt(Me.Height / b_height * 12)
        Label1.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label2.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label3.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label4.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label5.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label6.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label7.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label8.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label9.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label10.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label11.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label12.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label13.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label14.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label15.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label16.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label17.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label18.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 72
        y1 = 0
        x2 = 636
        y2 = 27
        Label19.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Label19.Width = x2 / a_width * Me.Width
        Label19.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 15)
        Label19.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        Label7.Location = New Point(currentx2 * 0.8, currenty3 + gapy * 5 + Me.Height * 0.02)
        Label8.Location = New Point(currentx2 * 0.8 + gapx, currenty3 + gapy * 5 + Me.Height * 0.02)
        Label9.Location = New Point(currentx2 * 0.8 + gapx * 2, currenty3 + gapy * 5 + Me.Height * 0.02)
        Label10.Location = New Point(currentx2 * 0.8 + gapx * 3, currenty3 + gapy * 5 + Me.Height * 0.02)
        Label11.Location = New Point(currentx2 * 0.8 + gapx * 4, currenty3 + gapy * 5 + Me.Height * 0.02)
        Label12.Location = New Point(currentx2 * 0.8 + gapx * 5, currenty3 + gapy * 5 + Me.Height * 0.02)
        Label13.Location = New Point(currentx2 * 0.8 + gapx * 6, currenty3 + gapy * 5 + Me.Height * 0.02)
        Label14.Location = New Point(currentx2 * 0.8 + gapx * 7, currenty3 + gapy * 5 + Me.Height * 0.02)
        Label15.Location = New Point(currentx2 * 0.8 + gapx * 8, currenty3 + gapy * 5 + Me.Height * 0.02)
        Label16.Location = New Point(currentx2 * 0.8 + gapx * 9, currenty3 + gapy * 5 + Me.Height * 0.02)
        Label17.Location = New Point(currentx2 * 0.8 + gapx * 10, currenty3 + gapy * 5 + Me.Height * 0.02)
        Label18.Location = New Point(currentx2 * 0.7 + gapx * 10.5, currenty3 + gapy * 5 + Me.Height * 0.02)

        x1 = 688
        y1 = 352
        x2 = 80
        y2 = 34
        Me.Button1.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Button1.Width = x2 / a_width * Me.Width
        Me.Button1.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Button1.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 688
        y1 = 303
        x2 = 80
        y2 = 34
        Me.Button2.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Button2.Width = x2 / a_width * Me.Width
        Me.Button2.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Button2.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        Fcurve = 1
        Dim C, D As Single
        Rm1 = R

        If Rm1 > 10 Then
            Ra1 = 5
            Timer1.Enabled = False
            Timer2.Enabled = False
            Timer3.Enabled = True
        ElseIf Rm1 < 10 Then
            Ra1 = 5
            Timer1.Enabled = False
            Timer2.Enabled = True
            Timer3.Enabled = False
        End If

        ' Open "acceler.dat" For Input As #1
        'Input #1, Ac
        'Close #1
        FileOpen(1, Application.StartupPath + "\acceler.dat", OpenMode.Input)
        Input(1, Ac)
        FileClose(1)

        Select Case W
            Case 1
                a = (7200.0# - 0.0#) / 3600.0# / 10.0#
                If Ac >= 10 Then
                    Timer1.Interval = 1000
                    Timer2.Interval = 1000
                ElseIf Ac >= 5 And Ac < 10 Then
                    Timer1.Interval = 3000
                    Timer2.Interval = 3000
                ElseIf Ac < 5 And Ac >= 2 Then
                    Timer1.Interval = 5000
                    Timer2.Interval = 5000
                Else
                    Timer1.Interval = 10000
                    Timer2.Interval = 10000
                End If
            Case 2
                a = (14400.0# - 0.0#) / 3600.0# / 10.0#
                If Ac >= 10 Then
                    Timer1.Interval = 2000
                    Timer2.Interval = 2000
                ElseIf Ac >= 5 And Ac < 10 Then
                    Timer1.Interval = 4000
                    Timer2.Interval = 4000
                ElseIf Ac < 5 And Ac >= 2 Then
                    Timer1.Interval = 10000
                    Timer2.Interval = 10000
                Else
                    Timer1.Interval = 20000
                    Timer2.Interval = 20000
                End If
            Case 3
                a = (28800.0# - 0.0#) / 3600.0# / 10.0#
                If Ac >= 10 Then
                    Timer1.Interval = 4000
                    Timer2.Interval = 4000
                ElseIf Ac >= 5 And Ac < 10 Then
                    Timer1.Interval = 8000
                    Timer2.Interval = 8000
                ElseIf Ac < 5 And Ac >= 2 Then
                    Timer1.Interval = 20000
                    Timer2.Interval = 20000
                Else
                    Timer1.Interval = 40000
                    Timer2.Interval = 40000
                End If
            Case 4
                a = (43200.0# - 0.0#) / 3600.0# / 10.0#
                If Ac >= 10 Then
                    Timer1.Interval = 5000
                    Timer2.Interval = 5000
                ElseIf Ac >= 5 And Ac < 10 Then
                    Timer1.Interval = 10000
                    Timer2.Interval = 10000
                ElseIf Ac >= 2 And Ac < 5 Then
                    Timer1.Interval = 25000
                    Timer2.Interval = 25000
                Else
                    Timer1.Interval = 50000
                    Timer2.Interval = 50000
                End If
            Case Else
        End Select

        For i = 1 To 11
            b = 7.0# + a * (i - 1)

            Select Case i
                Case 1
                    Label7.Text = Str$(b)
                Case 2
                    Label8.Text = Str$(b)
                Case 3
                    Label9.Text = Str$(b)
                Case 4
                    Label10.Text = Str$(b)
                Case 5
                    Label11.Text = Str$(b)
                Case 6
                    Label12.Text = Str$(b)
                Case 7
                    Label13.Text = Str$(b)
                Case 8
                    Label14.Text = Str$(b)
                Case 9
                    Label15.Text = Str$(b)
                Case 10
                    Label16.Text = Str$(b)
                Case Else
                    Label17.Text = Str$(b)

            End Select
        Next i

        C = (Vmax - Vmin) / 5.0#
        C = Mid(C, 1, 5)
        For j = 1 To 6
            D = Vmin + C * (j - 1)
            Select Case j
                Case 1
                    Label6.Text = Str$(D)
                Case 2
                    Label5.Text = Str$(D)
                Case 3
                    Label4.Text = Str$(D)
                Case 4
                    Label3.Text = Str$(D)
                Case 5
                    Label2.Text = Str$(D)
                Case Else
                    Label1.Text = Str$(D)
            End Select
        Next j







    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '  Fcurve = 0
        ' Me.Close()
        Me.Hide()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer2.Enabled = False

        Call readH()

        Select Case TK
            Case 23
                n11 = 36
                Call drawline(n11)
            Case 20
                n11 = 6
                Call drawline(n11)
            Case 24
                n11 = 38
                Call drawline(n11)
            Case 31
                n11 = 39
                Call drawline(n11)
            Case 32
                n11 = 40
                Call drawline(n11)
            Case 26
                n11 = 41
                Call drawline(n11)
            Case 21
                n11 = 42
                Call drawline(n11)
            Case 34
                n11 = 43
                Call drawline(n11)
            Case 22
                n11 = 44
                Call drawline(n11)
            Case 48
                n11 = 45
                Call drawline(n11)
            Case 35
                n11 = 46
                Call drawline(n11)
            Case 47
                n11 = 47
                Call drawline(n11)
            Case 36
                n11 = 47
                Call drawline(n11)
            Case 39
                n11 = 48
                Call drawline(n11)
            Case 46
                n11 = 49
                Call drawline(n11)
            Case 37
                n11 = 50
                Call drawline(n11)
            Case 41
                n11 = 51
                Call drawline(n11)
            Case 29
                n11 = 52
                Call drawline(n11)
            Case 43
                n11 = 53
                Call drawline(n11)
            Case 44
                n11 = 54
                Call drawline(n11)
            Case 45
                n11 = 55
                Call drawline(n11)
            Case 38
                n11 = 56
                Call drawline(n11)
            Case 33
                n11 = 57
                Call drawline(n11)
            Case 25
                n11 = 58
                Call drawline(n11)
            Case 30
                n11 = 19
                Call drawline(n11)
            Case 53
                n11 = 18
                Call drawline(n11)
            Case 54
                n11 = 27
                Call drawline(n11)
            Case 55
                n11 = 25
                Call drawline(n11)
            Case 42
                n11 = 4
                Call drawline(n11)
            Case 40
                n11 = 1
                Call drawline(n11)
            Case 50
                n11 = 3
                Call drawline(n11)

            Case 100
                n11 = 7
                Call drawline(n11)
            Case 99
                n11 = 8
                Call drawline(n11)
            Case 98
                n11 = 9
                Call drawline(n11)
            Case 97
                n11 = 10
                Call drawline(n11)
            Case 96
                n11 = 11
                Call drawline(n11)
            Case 95
                n11 = 12
                Call drawline(n11)
            Case 94
                n11 = 13
                Call drawline(n11)
            Case 93
                n11 = 14
                Call drawline(n11)
            Case 92
                n11 = 15
                Call drawline(n11)
            Case 91
                n11 = 16
                Call drawline(n11)
            Case 90
                n11 = 17
                Call drawline(n11)
            Case 87
                n11 = 20
                Call drawline(n11)
            Case 86
                n11 = 21
                Call drawline(n11)
            Case 85
                n11 = 22
                Call drawline(n11)
            Case 84
                n11 = 23
                Call drawline(n11)
            Case 83
                n11 = 24
                Call drawline(n11)
            Case 81
                n11 = 26
                Call drawline(n11)
            Case 79
                n11 = 28
                Call drawline(n11)
            Case 78
                n11 = 29
                Call drawline(n11)
            Case 77
                n11 = 30
                Call drawline(n11)
            Case 76
                n11 = 31
                Call drawline(n11)
            Case 75
                n11 = 32
                Call drawline(n11)
            Case 74
                n11 = 33
                Call drawline(n11)
            Case 73
                n11 = 34
                Call drawline(n11)
            Case 72
                n11 = 35
                Call drawline(n11)
            Case Else
        End Select

        Select Case W
            Case 1
                TInter = 2.0#
                If Ra1 > 30 And Hn1 = 7200 Then
                    Timer1.Enabled = False
                End If
            Case 2
                TInter = 4.0#
                If Ra1 > 30 And Hn1 = 14400 Then
                    Timer1.Enabled = False
                End If
            Case 3
                TInter = 8.0#
                If Ra1 > 30 And Hn1 = 28800 Then
                    Timer1.Enabled = False
                End If
            Case 4
                TInter = 12.0#
                If Ra1 > 30 And Hn1 = 43200 Then
                    Timer1.Enabled = False
                End If
            Case Else
        End Select

        Ho1 = Hn1
        Ra1 = Ra1 + 1

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        frmspread.grdspread.Row = 7
        frmspread.grdspread.Col = 10
        G = Val(frmspread.grdspread.Text)
        If G <> 0.0# Then
            Timer1.Enabled = True
        Else
        End If

    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        If Ra1 < Rm1 Then
            Call readH()

            Select Case TK
                Case 23
                    n11 = 36
                    Call drawline(n11)
                Case 20
                    n11 = 37
                    Call drawline(n11)
                Case 24
                    n11 = 38
                    Call drawline(n11)
                Case 31
                    n11 = 39
                    Call drawline(n11)
                Case 32
                    n11 = 40
                    Call drawline(n11)
                Case 26
                    n11 = 41
                    Call drawline(n11)
                Case 21
                    n11 = 42
                    Call drawline(n11)
                Case 34
                    n11 = 43
                    Call drawline(n11)
                Case 22
                    n11 = 44
                    Call drawline(n11)
                Case 48
                    n11 = 45
                    Call drawline(n11)
                Case 35
                    n11 = 46
                    Call drawline(n11)
                Case 47
                    n11 = 47
                    Call drawline(n11)
                Case 36
                    n11 = 47
                    Call drawline(n11)
                Case 39
                    n11 = 48
                    Call drawline(n11)
                Case 46
                    n11 = 49
                    Call drawline(n11)
                Case 37
                    n11 = 50
                    Call drawline(n11)
                Case 41
                    n11 = 51
                    Call drawline(n11)
                Case 29
                    n11 = 52
                    Call drawline(n11)
                Case 43
                    n11 = 53
                    Call drawline(n11)
                Case 44
                    n11 = 54
                    Call drawline(n11)
                Case 45
                    n11 = 55
                    Call drawline(n11)
                Case 38
                    n11 = 56
                    Call drawline(n11)
                Case 33
                    n11 = 57
                    Call drawline(n11)
                Case 25
                    n11 = 58
                    Call drawline(n11)
                Case 30
                    n11 = 19
                    Call drawline(n11)
                Case 53
                    n11 = 18
                    Call drawline(n11)
                Case 54
                    n11 = 27
                    Call drawline(n11)
                Case 55
                    n11 = 25
                    Call drawline(n11)
                Case 42
                    n11 = 4
                    Call drawline(n11)
                Case 40
                    n11 = 1
                    Call drawline(n11)
                Case 50
                    n11 = 3
                    Call drawline(n11)
                    'Case 52
                    '      If Mb(1) > 1 And Mb(1) <= 2 Then
                    '          n11 = 1
                    '          Call drawline(n11)
                    '       ElseIf Mb(1) <= 3 Then
                    '          n11 = 3
                    '          Call drawline(n11)
                    '       Else
                    '          n11 = 4
                    '          Call drawline(n11)
                    '       End If
                    'Case 27
                    '       If Mb(2) > 1 And Mb(2) <= 1 Then
                    '          n11 = 5
                    '          Call drawline(n11)
                    '       End If
                    'Case 28
                    '       If Mb(3) > 1 And Mb(3) <= 1 Then
                    '          n11 = 6
                    '          Call drawline(n11)
                    '       End If

                Case 100
                    n11 = 7
                    Call drawline(n11)
                Case 99
                    n11 = 8
                    Call drawline(n11)
                Case 98
                    n11 = 9
                    Call drawline(n11)
                Case 97
                    n11 = 10
                    Call drawline(n11)
                Case 96
                    n11 = 11
                    Call drawline(n11)
                Case 95
                    n11 = 12
                    Call drawline(n11)
                Case 94
                    n11 = 13
                    Call drawline(n11)
                Case 93
                    n11 = 14
                    Call drawline(n11)
                Case 92
                    n11 = 15
                    Call drawline(n11)
                Case 91
                    n11 = 16
                    Call drawline(n11)
                Case 90
                    n11 = 17
                    Call drawline(n11)
                Case 89
                    n11 = 18
                    Call drawline(n11)
                Case 88
                    n11 = 19
                    Call drawline(n11)
                Case 87
                    n11 = 20
                    Call drawline(n11)
                Case 86
                    n11 = 21
                    Call drawline(n11)
                Case 85
                    n11 = 22
                    Call drawline(n11)
                Case 84
                    n11 = 23
                    Call drawline(n11)
                Case 83
                    n11 = 24
                    Call drawline(n11)
                Case 82
                    n11 = 25
                    Call drawline(n11)
                Case 81
                    n11 = 26
                    Call drawline(n11)
                Case 80
                    n11 = 27
                    Call drawline(n11)
                Case 79
                    n11 = 28
                    Call drawline(n11)
                Case 78
                    n11 = 29
                    Call drawline(n11)
                Case 77
                    n11 = 30
                    Call drawline(n11)
                Case 76
                    n11 = 31
                    Call drawline(n11)
                Case 75
                    n11 = 32
                    Call drawline(n11)
                Case 74
                    n11 = 33
                    Call drawline(n11)
                Case 73
                    n11 = 34
                    Call drawline(n11)
                Case 72
                    n11 = 35
                    Call drawline(n11)
                Case Else
            End Select

            Select Case W
                Case 1
                    TInter = 2.0#
                    If Ra1 > 30 And Hn1 = 7200 Then
                        Timer3.Enabled = False
                    End If
                Case 2
                    TInter = 4.0#
                    If Ra1 > 30 And Hn1 = 14400 Then
                        Timer3.Enabled = False
                    End If
                Case 3
                    TInter = 8.0#
                    If Ra1 > 30 And Hn1 = 28800 Then
                        Timer3.Enabled = False
                    End If
                Case 4
                    TInter = 12.0#
                    If Ra1 > 30 And Hn1 = 43200 Then
                        Timer3.Enabled = False
                    End If
                Case Else
            End Select

            Ho1 = Hn1
            Ra1 = Ra1 + 1
        ElseIf Ra1 >= Rm1 Then
            Timer3.Enabled = False
            Timer1.Enabled = True
            Ra1 = Rm1
        End If

    End Sub

    Private Sub curve_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Dim p As Pen = New Pen(Color.Gray, 2)
        '  Dim i As Integer
        p.DashStyle = Drawing2D.DashStyle.Dash
        Dim px As Pen = New Pen(Color.Black, 2)
        Dim currentx2, currenty2, currentx3, currenty3, gapx, gapy As Single
        currentx2 = Me.Width * 0.05

        currenty3 = Me.Height * 0.1
        currenty2 = currenty3 + Me.Height * 0.75
        currentx3 = currentx2 + Me.Width * 0.8
        currentx2 = Me.Width * 0.056

        gapx = (currentx3 - currentx2) / 10
        gapy = (currenty2 - currenty3) / 5
        Me.CreateGraphics.DrawLine(px, currentx2, currenty2, currentx3, currenty2)
        Me.CreateGraphics.DrawLine(px, currentx2, currenty3, currentx3, currenty3)
        Me.CreateGraphics.DrawLine(px, currentx2, currenty3, currentx2, currenty2)
        Me.CreateGraphics.DrawLine(px, currentx3, currenty3, currentx3, currenty2)
        For i = 1 To 4
            Me.CreateGraphics.DrawLine(p, currentx2, currenty3 + gapy * i, currentx3, currenty3 + gapy * i)
        Next i

        For i = 1 To 9
            Me.CreateGraphics.DrawLine(p, currentx2 + gapx * i, currenty2, currentx2 + gapx * i, currenty3)
        Next i

    End Sub

    Private Sub Label19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label19.Click

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim g As Graphics

        g = Me.CreateGraphics
        g.Clear(Me.BackColor)
        g.Dispose()
        Call curve_Load(Nothing, Nothing)
        Dim p As Pen = New Pen(Color.Gray, 2)
        '  Dim i As Integer
        p.DashStyle = Drawing2D.DashStyle.Dash
        Dim px As Pen = New Pen(Color.Black, 2)
        Dim currentx2, currenty2, currentx3, currenty3, gapx, gapy As Single
        currentx2 = Me.Width * 0.06
        currenty3 = Me.Height * 0.1
        currenty2 = currenty3 + Me.Height * 0.75
        currentx3 = currentx2 + Me.Width * 0.8
        currentx2 = Me.Width * 0.066

        gapx = (currentx3 - currentx2) / 10
        gapy = (currenty2 - currenty3) / 5
        Me.CreateGraphics.DrawLine(px, currentx2, currenty2, currentx3, currenty2)
        Me.CreateGraphics.DrawLine(px, currentx2, currenty3, currentx3, currenty3)
        Me.CreateGraphics.DrawLine(px, currentx2, currenty3, currentx2, currenty2)
        Me.CreateGraphics.DrawLine(px, currentx3, currenty3, currentx3, currenty2)
        For i = 1 To 4
            Me.CreateGraphics.DrawLine(p, currentx2, currenty3 + gapy * i, currentx3, currenty3 + gapy * i)
        Next i

        For i = 1 To 9
            Me.CreateGraphics.DrawLine(p, currentx2 + gapx * i, currenty2, currentx2 + gapx * i, currenty3)
        Next i
    End Sub

  
End Class