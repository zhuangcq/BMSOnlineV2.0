Public Class curves

    Dim frmWidth
    Dim frmHeight
    Dim Counter2 As Integer
    Dim CurrentXb, CurrentY1b, CurrentXbo, CurrentY1bo As Single
    Dim CurrentY2b, CurrentY2bo, CurrentY3b, CurrentY3bo, CurrentY4b, CurrentY4bo As Single
    Dim CurrentY5b, CurrentY5bo As Single
    Dim CurrentXo, CurrentY1, CurrentY1o, CurrentY2, CurrentY2o As Single
    Dim CurrentY3, CurrentY3o, CurrentY4, CurrentY4o, CurrentY5, CurrentY5o As Single
    Dim CurrentX As Single
    Dim c1, c2, V1, V2 As Single
    Dim i, j As Integer
    Dim s As Integer
    Dim RO As Integer
    Public Sub drawline1(ByRef n1 As Object)

        If Ra > 798 Then
            RO = Ra - 798
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 118 + n1
            Ver1 = Val(frmspread.grdspread.Text)
        ElseIf Ra > 399 Then
            RO = Ra - 399
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 59 + n1
            Ver1 = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra
            frmspread.grdspread.Row = n1
            Ver1 = Val(frmspread.grdspread.Text)
        End If
        Dim currentwidth0, Currentheight1, currentwidth1, Currentheight0 As Single
        currentwidth0 = Me.Width * 0.05
        Currentheight0 = Me.Height * 0.02
        currentwidth1 = currentwidth0 + Me.Width * 0.6
        Currentheight1 = Currentheight0 + Me.Height * 0.45



        If Ra > 5 Then
            CurrentX = (Hn / 3600.0#) * (currentwidth1 - currentwidth0) / period + currentwidth0
            CurrentY1 = Currentheight1 - (Ver1 - Vmin1) * (Currentheight1 - Currentheight0) * 1.0# / (Vmax1 - Vmin1)

            CurrentXo = (Ho / 3600.0#) * (currentwidth1 - currentwidth0) / period + currentwidth0
            CurrentY1o = Currentheight1 - (Ver1o - Vmin1) * (Currentheight1 - Currentheight0) * 1.0# / (Vmax1 - Vmin1)
            Dim gra As Graphics = Me.CreateGraphics
            Dim myPen As Pen = New Pen(Color.Red, 2)


            ' Line (CurrentXo, CurrentY1o)-(CurrentX, CurrentY1), QBColor(12)
            gra.DrawLine(myPen, CurrentXo, CurrentY1o, CurrentX, CurrentY1)
        Else
        End If
        Ver1o = Ver1

    End Sub
  
    Public Sub drawline2(ByRef n2 As Object)

        If Ra > 798 Then
            Ro = Ra - 798
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 118 + n2
            Ver2 = Val(frmspread.grdspread.Text)
        ElseIf Ra > 399 Then
            Ro = Ra - 399
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 59 + n2
            Ver2 = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra
            frmspread.grdspread.Row = n2
            Ver2 = Val(frmspread.grdspread.Text)
        End If
        Dim currentwidth0, Currentheight1, currentwidth1, Currentheight0 As Single
        currentwidth0 = Me.Width * 0.05
        Currentheight0 = Me.Height * 0.02
        currentwidth1 = currentwidth0 + Me.Width * 0.6
        Currentheight1 = Currentheight0 + Me.Height * 0.45
        If Ra > 5 Then
            CurrentX = (Hn / 3600.0#) * (currentwidth1 - currentwidth0) / period + currentwidth0
            CurrentY2 = Currentheight1 - (Ver2 - Vmin1) * (Currentheight1 - Currentheight0) * 1.0# / (Vmax1 - Vmin1)

            CurrentXo = (Ho / 3600.0#) * (currentwidth1 - currentwidth0) / period + currentwidth0
            CurrentY2o = Currentheight1 - (Ver2o - Vmin1) * (Currentheight1 - Currentheight0) * 1.0# / (Vmax1 - Vmin1)
            Dim gra As Graphics = Me.CreateGraphics
            Dim myPen As Pen = New Pen(Color.Orange, 2)

            ' Line (CurrentXo, CurrentY2o)-(CurrentX, CurrentY2), QBColor(14)
            gra.DrawLine(myPen, CurrentXo, CurrentY2o, CurrentX, CurrentY2)
        Else
        End If
        Ver2o = Ver2
        'Ho = Hn
        'Label2.ForeColor = QBColor(6)
    End Sub

    Public Sub drawline3(ByRef n3 As Object)

        If Ra > 798 Then
            Ro = Ra - 798
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 118 + n3
            Ver3 = Val(frmspread.grdspread.Text)
        ElseIf Ra > 399 Then
            Ro = Ra - 399
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 59 + n3
            Ver3 = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra
            frmspread.grdspread.Row = n3
            Ver3 = Val(frmspread.grdspread.Text)
        End If
        Dim currentwidth0, Currentheight1, currentwidth1, Currentheight0 As Single
        currentwidth0 = Me.Width * 0.05
        Currentheight0 = Me.Height * 0.02
        currentwidth1 = currentwidth0 + Me.Width * 0.6
        Currentheight1 = Currentheight0 + Me.Height * 0.45
        If Ra > 5 Then
            CurrentX = (Hn / 3600.0#) * (currentwidth1 - currentwidth0) / period + currentwidth0
            CurrentY3 = Currentheight1 - (Ver3 - Vmin1) * (Currentheight1 - Currentheight0) * 1.0# / (Vmax1 - Vmin1)

            CurrentXo = (Ho / 3600.0#) * (currentwidth1 - currentwidth0) / period + currentwidth0
            CurrentY3o = Currentheight1 - (Ver3o - Vmin1) * (Currentheight1 - Currentheight0) * 1.0# / (Vmax1 - Vmin1)
            Dim gra As Graphics = Me.CreateGraphics
            Dim myPen As Pen = New Pen(Color.Cyan, 2)

            gra.DrawLine(myPen, CurrentXo, CurrentY3o, CurrentX, CurrentY3)

            '  Line (CurrentXo, CurrentY3o)-(CurrentX, CurrentY3), QBColor(11)
        Else
        End If
        Ver3o = Ver3

    End Sub

    Public Sub drawline4(ByRef n4 As Object)

        If Ra > 798 Then
            Ro = Ra - 798
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 118 + n4
            Ver4 = Val(frmspread.grdspread.Text)
        ElseIf Ra > 399 Then
            Ro = Ra - 399
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 59 + n4
            Ver4 = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra
            frmspread.grdspread.Row = n4
            Ver4 = Val(frmspread.grdspread.Text)
        End If
        Dim currentwidth0, Currentheight1, currentwidth1, Currentheight0 As Single
        currentwidth0 = Me.Width * 0.05
        Currentheight0 = Me.Height * 0.02
        currentwidth1 = currentwidth0 + Me.Width * 0.6
        Currentheight1 = Currentheight0 + Me.Height * 0.45
        If Ra > 5 Then
            CurrentX = (Hn / 3600.0#) * (currentwidth1 - currentwidth0) / period + currentwidth0
            CurrentY4 = Currentheight1 - (Ver4 - Vmin1) * (Currentheight1 - Currentheight0) * 1.0# / (Vmax1 - Vmin1)

            CurrentXo = (Ho / 3600.0#) * (currentwidth1 - currentwidth0) / period + currentwidth0
            CurrentY4o = Currentheight1 - (Ver4o - Vmin1) * (Currentheight1 - Currentheight0) * 1.0# / (Vmax1 - Vmin1)

            ' Line (CurrentXo, CurrentY4o)-(CurrentX, CurrentY4), QBColor(15)
            Dim gra As Graphics = Me.CreateGraphics
            Dim myPen As Pen = New Pen(Color.Black, 2)
            gra.DrawLine(myPen, CurrentXo, CurrentY4o, CurrentX, CurrentY4)
        Else
        End If
        Ver4o = Ver4

    End Sub
    Public Sub drawline5(ByRef n5 As Object)

        If Ra > 798 Then
            Ro = Ra - 798
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 118 + n5
            Ver5 = Val(frmspread.grdspread.Text)
        ElseIf Ra > 399 Then
            Ro = Ra - 399
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 59 + n5
            Ver5 = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra
            frmspread.grdspread.Row = n5
            Ver5 = Val(frmspread.grdspread.Text)
        End If
        Dim currentwidth0, Currentheight1, currentwidth1, Currentheight0 As Single
        currentwidth0 = Me.Width * 0.05
        Currentheight0 = Me.Height * 0.02
        currentwidth1 = currentwidth0 + Me.Width * 0.6
        Currentheight1 = Currentheight0 + Me.Height * 0.45
        If Ra > 5 Then
            CurrentX = (Hn / 3600.0#) * (currentwidth1 - currentwidth0) / period + currentwidth0
            CurrentY5 = Currentheight1 - (Ver5 - Vmin1) * (Currentheight1 - Currentheight0) * 1.0# / (Vmax1 - Vmin1)

            CurrentXo = (Ho / 3600.0#) * (currentwidth1 - currentwidth0) / period + currentwidth0
            CurrentY5o = Currentheight1 - (Ver5o - Vmin1) * (Currentheight1 - Currentheight0) * 1.0# / (Vmax1 - Vmin1)
            Dim gra As Graphics = Me.CreateGraphics
            Dim myPen As Pen = New Pen(Color.Green, 2)
            gra.DrawLine(myPen, CurrentXo, CurrentY5o, CurrentX, CurrentY5)

            '  Line (CurrentXo, CurrentY5o)-(CurrentX, CurrentY5), QBColor(10)
        Else
        End If
        Ver5o = Ver5

    End Sub


    Public Sub drawline6(ByRef n6 As Object)

        If Ra > 798 Then
            Ro = Ra - 798
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 118 + n6
            Ver1b = Val(frmspread.grdspread.Text)
        ElseIf Ra > 399 Then
            Ro = Ra - 399
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 59 + n6
            Ver1b = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra
            frmspread.grdspread.Row = n6
            Ver1b = Val(frmspread.grdspread.Text)
        End If
        Dim currentwidth0, Currentheight1, currentwidth1, Currentheight0 As Single
        currentwidth0 = Me.Width * 0.05
        Currentheight0 = Me.Height * 0.02
        currentwidth1 = currentwidth0 + Me.Width * 0.6
        Currentheight1 = Currentheight0 + Me.Height * 0.45
        Dim currentwidth0n, Currentheight1n, currentwidth1n, Currentheight0n As Single
        currentwidth0n = currentwidth0
        Currentheight0n = Me.Height * 0.035 + Currentheight1
        Currentheight1n = (Currentheight1 - Currentheight0) + Currentheight0n
        currentwidth1n = currentwidth1
    

        If Ra > 5 Then
            CurrentX = (Hn / 3600.0#) * (currentwidth1n - currentwidth0n) / period + currentwidth0
            CurrentY1b = (Currentheight1n) - (Ver1b - Vmin2) * (Currentheight1n - Currentheight0n) * 1.0# / (Vmax2 - Vmin2)

            CurrentXo = (Ho / 3600.0#) * (currentwidth1n - currentwidth0n) / period + currentwidth0
            CurrentY1bo = (Currentheight1n) - (Ver1bo - Vmin2) * (Currentheight1n - Currentheight0n) * 1.0# / (Vmax2 - Vmin2)
            Dim gra As Graphics = Me.CreateGraphics
            Dim myPen As Pen = New Pen(Color.Red, 2)

            gra.DrawLine(myPen, CurrentXo, CurrentY1bo, CurrentX, CurrentY1b)

            'Line (CurrentXo, CurrentY1bo)-(CurrentX, CurrentY1b), QBColor(12)
        Else
        End If
        Ver1bo = Ver1b


    End Sub

    Public Sub drawline7(ByRef n7 As Object)

        If Ra > 798 Then
            Ro = Ra - 798
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 118 + n7
            Ver2b = Val(frmspread.grdspread.Text)
        ElseIf Ra > 399 Then
            Ro = Ra - 399
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 59 + n7
            Ver2b = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra
            frmspread.grdspread.Row = n7
            Ver2b = Val(frmspread.grdspread.Text)
        End If
        Dim currentwidth0, Currentheight1, currentwidth1, Currentheight0 As Single
        currentwidth0 = Me.Width * 0.05
        Currentheight0 = Me.Height * 0.02
        currentwidth1 = currentwidth0 + Me.Width * 0.6
        Currentheight1 = Currentheight0 + Me.Height * 0.45
        Dim currentwidth0n, Currentheight1n, currentwidth1n, Currentheight0n As Single
        currentwidth0n = currentwidth0
        Currentheight0n = Me.Height * 0.035 + Currentheight1
        Currentheight1n = (Currentheight1 - Currentheight0) + Currentheight0n
        currentwidth1n = currentwidth1

        If Ra > 5 Then
            CurrentX = (Hn / 3600.0#) * (currentwidth1n - currentwidth0n) / period + currentwidth0
            CurrentY2b = (Currentheight1n) - (Ver2b - Vmin2) * (Currentheight1n - Currentheight0n) * 1.0# / (Vmax2 - Vmin2)

            CurrentXo = (Ho / 3600.0#) * (currentwidth1n - currentwidth0n) / period + currentwidth0
            CurrentY2bo = (Currentheight1n) - (Ver2bo - Vmin2) * (Currentheight1n - Currentheight0n) * 1.0# / (Vmax2 - Vmin2)
            Dim gra As Graphics = Me.CreateGraphics
            Dim myPen As Pen = New Pen(Color.Orange, 2)
            gra.DrawLine(myPen, CurrentXo, CurrentY2bo, CurrentX, CurrentY2b)

            'Line (CurrentXo, CurrentY2bo)-(CurrentX, CurrentY2b), QBColor(14)
        Else
        End If
        Ver2bo = Ver2b
    End Sub

    Public Sub drawline8(ByRef n8 As Object)

        If Ra > 798 Then
            Ro = Ra - 798
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 118 + n8
            Ver3b = Val(frmspread.grdspread.Text)
        ElseIf Ra > 399 Then
            Ro = Ra - 399
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 59 + n8
            Ver3b = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra
            frmspread.grdspread.Row = n8
            Ver3b = Val(frmspread.grdspread.Text)
        End If
        Dim currentwidth0, Currentheight1, currentwidth1, Currentheight0 As Single
        currentwidth0 = Me.Width * 0.05
        Currentheight0 = Me.Height * 0.02
        currentwidth1 = currentwidth0 + Me.Width * 0.6
        Currentheight1 = Currentheight0 + Me.Height * 0.45
        Dim currentwidth0n, Currentheight1n, currentwidth1n, Currentheight0n As Single
        currentwidth0n = currentwidth0
        Currentheight0n = Me.Height * 0.035 + Currentheight1
        Currentheight1n = (Currentheight1 - Currentheight0) + Currentheight0n
        currentwidth1n = currentwidth1

        If Ra > 5 Then
            CurrentX = (Hn / 3600.0#) * (currentwidth1n - currentwidth0n) / period + currentwidth0
            CurrentY3b = Currentheight1n - (Ver3b - Vmin2) * (Currentheight1n - Currentheight0n) * 1.0# / (Vmax2 - Vmin2)

            CurrentXo = (Ho / 3600.0#) * (currentwidth1n - currentwidth0n) / period + currentwidth0
            CurrentY3bo = Currentheight1n - (Ver3bo - Vmin2) * (Currentheight1n - Currentheight0n) * 1.0# / (Vmax2 - Vmin2)


            Dim gra As Graphics = Me.CreateGraphics
            Dim myPen As Pen = New Pen(Color.Cyan, 2)
            gra.DrawLine(myPen, CurrentXo, CurrentY3bo, CurrentX, CurrentY3b)
            'Line (CurrentXo, CurrentY3bo)-(CurrentX, CurrentY3b), QBColor(11)
        Else
        End If
        Ver3bo = Ver3b
    End Sub


    Public Sub drawline9(ByRef n9 As Object)

        If Ra > 798 Then
            Ro = Ra - 798
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 118 + n9
            Ver4b = Val(frmspread.grdspread.Text)
        ElseIf Ra > 399 Then
            Ro = Ra - 399
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 59 + n9
            Ver4b = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra
            frmspread.grdspread.Row = n9
            Ver4b = Val(frmspread.grdspread.Text)
        End If
        Dim currentwidth0, Currentheight1, currentwidth1, Currentheight0 As Single
        currentwidth0 = Me.Width * 0.05
        Currentheight0 = Me.Height * 0.02
        currentwidth1 = currentwidth0 + Me.Width * 0.6
        Currentheight1 = Currentheight0 + Me.Height * 0.45
        Dim currentwidth0n, Currentheight1n, currentwidth1n, Currentheight0n As Single
        currentwidth0n = currentwidth0
        Currentheight0n = Me.Height * 0.035 + Currentheight1
        Currentheight1n = (Currentheight1 - Currentheight0) + Currentheight0n
        currentwidth1n = currentwidth1

        If Ra > 5 Then
            CurrentX = (Hn / 3600.0#) * (currentwidth1n - currentwidth0n) / period + currentwidth0
            CurrentY4b = Currentheight1n - (Ver4b - Vmin2) * (Currentheight1n - Currentheight0n) * 1.0# / (Vmax2 - Vmin2)

            CurrentXo = (Ho / 3600.0#) * (currentwidth1n - currentwidth0n) / period + currentwidth0
            CurrentY4bo = Currentheight1n - (Ver4bo - Vmin2) * (Currentheight1n - Currentheight0n) * 1.0# / (Vmax2 - Vmin2)
            Dim gra As Graphics = Me.CreateGraphics
            Dim myPen As Pen = New Pen(Color.Black, 2)
            gra.DrawLine(myPen, CurrentXo, CurrentY4bo, CurrentX, CurrentY4b)

            ' Line (CurrentXo, CurrentY4bo)-(CurrentX, CurrentY4b), QBColor(15)
        Else
        End If
        Ver4bo = Ver4b
    End Sub
    Public Sub drawline10(ByRef n10 As Object)

        If Ra > 798 Then
            Ro = Ra - 798
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 118 + n10
            Ver5b = Val(frmspread.grdspread.Text)
        ElseIf Ra > 399 Then
            Ro = Ra - 399
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 59 + n10
            Ver5b = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra
            frmspread.grdspread.Row = n10
            Ver5b = Val(frmspread.grdspread.Text)
        End If
        Dim currentwidth0, Currentheight1, currentwidth1, Currentheight0 As Single
        currentwidth0 = Me.Width * 0.05
        Currentheight0 = Me.Height * 0.02
        currentwidth1 = currentwidth0 + Me.Width * 0.6
        Currentheight1 = Currentheight0 + Me.Height * 0.45
        Dim currentwidth0n, Currentheight1n, currentwidth1n, Currentheight0n As Single
        currentwidth0n = currentwidth0
        Currentheight0n = Me.Height * 0.035 + Currentheight1
        Currentheight1n = (Currentheight1 - Currentheight0) + Currentheight0n
        currentwidth1n = currentwidth1

        If Ra > 5 Then
            CurrentX = (Hn / 3600.0#) * (currentwidth1n - currentwidth0n) / period + currentwidth0
            CurrentY5b = (Currentheight1n) - (Ver5b - Vmin2) * (Currentheight1n - Currentheight0n) * 1.0# / (Vmax2 - Vmin2)

            CurrentXo = (Hn / 3600.0#) * (currentwidth1n - currentwidth0n) / period + currentwidth0
            CurrentY5bo = Currentheight1n - (Ver5bo - Vmin2) * (Currentheight1n - Currentheight0n) * 1.0# / (Vmax2 - Vmin2)


            ' Line (CurrentXo, CurrentY5bo)-(CurrentX, CurrentY5b), QBColor(10)
            Dim gra As Graphics = Me.CreateGraphics
            Dim myPen As Pen = New Pen(Color.Green, 2)
            myPen.DashStyle = Drawing2D.DashStyle.Solid
            gra.DrawLine(myPen, CurrentXo, CurrentY5bo, CurrentX, CurrentY5b)
        Else
        End If
        Ver5bo = Ver5b
    End Sub
    Public Sub readH()

        If Ra > 798 Then
            RO = Ra - 798
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 177
            Hn = Val(frmspread.grdspread.Text)
        ElseIf Ra > 399 Then
            RO = Ra - 399
            frmspread.grdspread.Col = RO
            frmspread.grdspread.Row = 118
            Hn = Val(frmspread.grdspread.Text)
        Else
            frmspread.grdspread.Col = Ra
            frmspread.grdspread.Row = 59
            Hn = Val(frmspread.grdspread.Text)
        End If

    End Sub


    Private Sub curves_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Dim Hmax, Hmin As Single
        Dim a, b As Single
        Dim i As Single
        Dim en As Integer
        Hmin = 0.0#
        Select Case W
            Case 1
                Hmax = 2.0#
            Case 2
                Hmax = 4.0#
            Case 3
                Hmax = 8.0#
            Case Else
                Hmax = 12.0#
        End Select

        a = (Hmax - Hmin) / 10.0#

        For i = 1 To 11
            b = 7.0# + a * (i - 1)

            Select Case i
                Case 1
                    Label33.Text = Str$(b)
                Case 2
                    Label34.Text = Str$(b)
                Case 3
                    Label35.Text = Str$(b)
                Case 4
                    Label36.Text = Str$(b)
                Case 5
                    Label37.Text = Str$(b)
                Case 6
                    Label38.Text = Str$(b)
                Case 7
                    Label39.Text = Str$(b)
                Case 8
                    Label40.Text = Str$(b)
                Case 9
                    Label41.Text = Str$(b)
                Case 10
                    Label42.Text = Str$(b)
                Case Else
                    Label43.Text = Str$(b)
            End Select

        Next i
     
        Countera = Counter
        If Z17 = 1 Then
            Countera = Countera - 1

            en = Counter - Countera
            Flag(en) = 17
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Aver.P"
            Panel2.Controls.Item(Index).Text = "Average Pollutant Concentration of Zones(ppm)"
        Else
        End If
        If Z18 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 18
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Powvav"
            Panel2.Controls.Item(Index).Text = "Power consumption of VAV supply fan(kw)"
        Else
        End If
        If Z19 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 19
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Powcav"
            Panel2.Controls.Item(Index).Text = "Power consumption of CAV supply fan(kw)"
        Else
        End If
        If Z20 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 20
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "FlowZ1"
            Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone1 "
        Else
        End If
        If Z21 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 21
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "FlowZ2"
            Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone2 "
        Else
        End If
        If Z22 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 22
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "FlowZ3"
            Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone3 "
        Else
        End If
        If Z23 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 23
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "FlowZ4"
            Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone4 "
        Else
        End If
        If Z24 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 24
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "FlowZ5"
            Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone5 "
        Else
        End If
        If Z25 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 25
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Qvav"
            Panel2.Controls.Item(Index).Text = "VAV Total heat transfer from the air "
        Else
        End If
        If Z26 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 26
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "FlowZ7"
            Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone7 "
        Else
        End If
        If Z27 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 27
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Qcav"
            Panel2.Controls.Item(Index).Text = "CAV Total heat transfer from the air "
        Else
        End If
        If Z28 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 28
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Dam.Z1"
            Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone1 "
        Else
        End If
        If Z29 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 29
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Dam.Z2"
            Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone2 "
        Else
        End If
        If Z30 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 30
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Dam.Z3"
            Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone3 "
        Else
        End If
        If Z31 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 31
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Dam.Z4"
            Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone4 "
        Else
        End If
        If Z32 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 32
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Dam.Z5"
            Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone5 "
        Else
        End If
        If Z33 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 33
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Dam.Z6"
            Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone6 "
        Else
        End If
        If Z34 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 34
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Dam.Z7"
            Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone7 "
        Else
        End If
        If Z35 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 35
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Dam.Z8"
            Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone8 "
        Else
        End If
        If Z38 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 38
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "CAVcoil"
            Panel2.Controls.Item(Index).Text = "CAV cooling coil demand "
        Else
        End If
        If Z40 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 40
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "VAVcoil"
            Panel2.Controls.Item(Index).Text = "VAV cooling coil demand "
        Else
        End If
        If Z43 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 43
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Sup.Fan"
            Panel2.Controls.Item(Index).Text = "VAV supply fan control"
        Else
        End If
        If Z44 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 44
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Sup.F"
            Panel2.Controls.Item(Index).Text = "Total supply air flow rate "
        Else
        End If
        If Z45 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 45
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Retu.F"
            Panel2.Controls.Item(Index).Text = "Total return air flow rate(kg/s)"
        Else
        End If
        If Z46 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 46
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Re.Fan"
            Panel2.Controls.Item(Index).Text = "Return fan control demand "
        Else
        End If
        If Z47 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 47
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Fres.F"
            Panel2.Controls.Item(Index).Text = "Flow rate of fresh air(m3/s) "
        Else
        End If
        If Z48 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 48
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Powrtn"
            Panel2.Controls.Item(Index).Text = "Power consumption of return fan(kw)"
        Else
        End If
        If Z49 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 49
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Out.Dam"
            Panel2.Controls.Item(Index).Text = "Outdoor air damper position demand "
        Else
        End If
        If Z57 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 57
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Wat.V"
            Panel2.Controls.Item(Index).Text = "Water valve control of VAV coil"
        Else
        End If
        If Z58 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 58
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Wat.C"
            Panel2.Controls.Item(Index).Text = "Water valve of CAV coil"
        Else
        End If
        If Z15 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 15
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Aver.H"
            Panel2.Controls.Item(Index).Text = "Average humidity of Zones (kg/kg)"
        Else
        End If
        If Z52 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 52
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Retu.H"
            Panel2.Controls.Item(Index).Text = "Humidity of return air(kg/kg)"
        Else
        End If
        If Z53 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 53
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Fres.H"
            Panel2.Controls.Item(Index).Text = "Humidity of fresh air(kg/kg)"
        Else
        End If
        If Z41 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 41
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Sup.P"
            Panel2.Controls.Item(Index).Text = "VAV supply pressure(Pa) "
        Else
        End If
        'If Z42 = 1 Then
        '   Countera = Countera - 1
        '   E = Counter - Countera
        '   Flag(E) = 42
        '   Index = E - 1
        '   Label1(Index).Caption = "P.set"
        '   Label2(Index).Caption = "Static pressure setpoint(pa)"
        'Else
        'End If
        If Z16 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 16
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Ave.CO2"
            Panel2.Controls.Item(Index).Text = "Average CO2 concentration of Zones(ppm)"
        Else
        End If
        If Z50 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 50
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Re.CO2"
            Panel2.Controls.Item(Index).Text = "CO2 concentration of return air(ppm)"
        Else
        End If
        If Z55 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 55
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Fre.CO2"
            Panel2.Controls.Item(Index).Text = "CO2 concentration of fresh air((ppm)"
        Else
        End If
        If Z56 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 56
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Number"
            Panel2.Controls.Item(Index).Text = "Number of people "
        Else
        End If
        If Z1 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 1
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Fre.flow"
            Panel2.Controls.Item(Index).Text = "Optimal fresh air flow rate setpoint(kg/s)"
        Else
        End If
        If Z3 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 3
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Flow.DVC"
            Panel2.Controls.Item(Index).Text = "Fresh air flow rate setpoint by DVC (kg/s) "
        Else
        End If
        If Z4 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 4
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "Flow.En"
            Panel2.Controls.Item(Index).Text = "Fresh air flow rate setpoint by enthapy (kg/s)"
        Else
        End If
        If Z5 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 5
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "P.set"
            Panel2.Controls.Item(Index).Text = "Optimal static pressure setpoint (pa)"
        Else
        End If
        If Z6 = 1 Then
            Countera = Countera - 1
            en = Counter - Countera
            Flag(en) = 6
            Index = en - 1
            Panel1.Controls.Item(Index).Text = "T.set"
            Panel2.Controls.Item(Index).Text = "Optimal supply air temperature setpoint (°C)"
        Else
        End If

        Dim e1n As Integer

        Counterb = Counter1
        If Z7 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 7
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "Tz1"
            Panel4.Controls.Item(Index).Text = "Temperature of Zone 1 (°C)"
        Else
        End If
        If Z8 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 8
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "Tz2"
            Panel4.Controls.Item(Index).Text = "Temperature of Zone 2 (°C)"
        Else
        End If
        If Z9 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 9
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "Tz3"
            Panel4.Controls.Item(Index).Text = "Temperature of Zone 3 (°C)"
        Else
        End If
        If Z10 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 10
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "Tz4"
            Panel4.Controls.Item(Index).Text = "Temperature of Zone 4 (°C)"
        Else
        End If
        If Z11 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 11
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "Tz5"
            Panel4.Controls.Item(Index).Text = "Temperature of Zone 5 (°C)"
        Else
        End If
        If Z12 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 12
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "Tz6"
            Panel4.Controls.Item(Index).Text = "Temperature of Zone 6 (°C)"
        Else
        End If
        If Z13 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 13
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "Tz7"
            Panel4.Controls.Item(Index).Text = "Temperature of Zone 7 (°C)"
        Else
        End If
        If Z14 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 14
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "Tz8"
            Panel4.Controls.Item(Index).Text = "Temperature of Zone 8 (°C)"
        Else
        End If
        If Z36 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 36
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "Tout.C"
            Panel4.Controls.Item(Index).Text = "CAV cooling coil outlet temperature (°C)"
        Else
        End If
        If Z37 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 37
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "Tset.Sup"
            Panel4.Controls.Item(Index).Text = "Supply air temperature setpoint(°C)"
        Else
        End If
        If Z39 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 39
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "Tout.V"
            Panel4.Controls.Item(Index).Text = "VAV cooling coil outlet temperature(°C)"
        Else
        End If
        If Z51 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 51
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "T.retu"
            Panel4.Controls.Item(Index).Text = "Temperature of return air (°C)"
        Else
        End If
        If Z54 = 1 Then
            Counterb = Counterb - 1
            e1n = Counter1 - Counterb
            Flag1(e1n) = 54
            Index = e1n - 1
            Panel3.Controls.Item(Index).Text = "T.fres"
            Panel4.Controls.Item(Index).Text = "Temperature of fresh air (°C)"
        Else
        End If


    End Sub

    Private Sub Panel1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseMove
        

    End Sub
   
    Public Sub New()

        InitializeComponent()

        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen

    End Sub

    Private Sub curves_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MaximizeBox = False
        Me.TopMost = False
        Me.Height = My.Computer.Screen.WorkingArea.Height
        Me.Width = My.Computer.Screen.WorkingArea.Width
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None

        Label1.TextAlign = ContentAlignment.MiddleLeft
         

       
        Label18.TextAlign = ContentAlignment.MiddleLeft

        Label20.TextAlign = ContentAlignment.MiddleLeft

        Label21.TextAlign = ContentAlignment.TopRight
        Label22.TextAlign = ContentAlignment.TopRight
        Label23.TextAlign = ContentAlignment.TopRight
        Label24.TextAlign = ContentAlignment.TopRight
        Label25.TextAlign = ContentAlignment.TopRight
        Label26.TextAlign = ContentAlignment.TopRight
        Label27.TextAlign = ContentAlignment.TopRight
        Label28.TextAlign = ContentAlignment.TopRight
        Label29.TextAlign = ContentAlignment.TopRight
        Label30.TextAlign = ContentAlignment.TopRight

        Label31.TextAlign = ContentAlignment.TopRight
        Label32.TextAlign = ContentAlignment.TopRight
        Label33.TextAlign = ContentAlignment.MiddleLeft
        Label34.TextAlign = ContentAlignment.MiddleLeft
        Label35.TextAlign = ContentAlignment.MiddleLeft
        Label36.TextAlign = ContentAlignment.MiddleLeft
        Label37.TextAlign = ContentAlignment.MiddleLeft
        Label38.TextAlign = ContentAlignment.MiddleLeft
        Label39.TextAlign = ContentAlignment.MiddleLeft
        Label40.TextAlign = ContentAlignment.MiddleLeft

        Label41.TextAlign = ContentAlignment.MiddleLeft
        Label42.TextAlign = ContentAlignment.MiddleLeft
        Label43.TextAlign = ContentAlignment.MiddleLeft
        Label44.TextAlign = ContentAlignment.MiddleLeft
        Label45.TextAlign = ContentAlignment.MiddleLeft
        Label46.TextAlign = ContentAlignment.MiddleLeft
        Label47.TextAlign = ContentAlignment.MiddleLeft
        Label48.TextAlign = ContentAlignment.MiddleLeft
        Label49.TextAlign = ContentAlignment.MiddleLeft
        Label50.TextAlign = ContentAlignment.MiddleLeft
        Label51.TextAlign = ContentAlignment.MiddleLeft
        Label52.TextAlign = ContentAlignment.MiddleLeft
        Label53.TextAlign = ContentAlignment.MiddleLeft
        Label71.TextAlign = ContentAlignment.MiddleLeft
        Label72.TextAlign = ContentAlignment.MiddleLeft
        Label73.TextAlign = ContentAlignment.MiddleLeft
        Label74.TextAlign = ContentAlignment.MiddleLeft
        Label75.TextAlign = ContentAlignment.MiddleLeft
        Label76.TextAlign = ContentAlignment.MiddleLeft
        Label77.TextAlign = ContentAlignment.MiddleLeft
        Label78.TextAlign = ContentAlignment.MiddleLeft
        Label79.TextAlign = ContentAlignment.MiddleLeft



        Label1.BackColor = Color.Transparent



        Label18.BackColor = Color.Transparent

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
        Label36.BackColor = Color.Transparent
        Label37.BackColor = Color.Transparent
        Label38.BackColor = Color.Transparent
        Label39.BackColor = Color.Transparent
        Label40.BackColor = Color.Transparent

        Label41.BackColor = Color.Transparent
        Label42.BackColor = Color.Transparent
        Label43.BackColor = Color.Transparent
        Label44.BackColor = Color.Transparent
        Label45.BackColor = Color.Transparent
        Label46.BackColor = Color.Transparent
        Label47.BackColor = Color.Transparent
        Label48.BackColor = Color.Transparent
        Label49.BackColor = Color.Transparent
        Label50.BackColor = Color.Transparent
        Label51.BackColor = Color.Transparent
        Label52.BackColor = Color.Transparent
        Label53.BackColor = Color.Transparent
        Label71.BackColor = Color.Transparent
        Label72.BackColor = Color.Transparent
        Label73.BackColor = Color.Transparent
        Label74.BackColor = Color.Transparent
        Label75.BackColor = Color.Transparent
        Label76.BackColor = Color.Transparent
        Label77.BackColor = Color.Transparent
        Label78.BackColor = Color.Transparent
        Label79.BackColor = Color.Transparent



        frmWidth = Me.Width
        frmHeight = Me.Height
        'Label(Size)
        '比例缩放
        Dim x1, y1, x2, y2, a_width, b_height, c_font As Integer

        '窗口大小 a为宽，b为高
        a_width = 1008
        b_height = 672

        x1 = 675
        y1 = 9
        x2 = 80
        y2 = 20
        Me.Label47.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label47.Width = x2 / a_width * Me.Width
        Me.Label47.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Label47.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 675
        y1 = 330
        x2 = 80
        y2 = 20
        Me.Label18.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Label18.Width = x2 / a_width * Me.Width
        Me.Label18.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 12)
        Label18.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        x1 = 678
        y1 = 18
        x2 = 61
        y2 = 149
        Me.Panel1.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Panel1.Width = x2 / a_width * Me.Width
        Me.Panel1.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 8)
        Panel1.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Dim width_p1, height_p1 As Single
        width_p1 = Panel1.Width / 5
        height_p1 = Panel1.Height / 6
        Label44.Location = New Point(width_p1, height_p1)
        Label45.Location = New Point(width_p1, height_p1 * 2)
        Label46.Location = New Point(width_p1, height_p1 * 3)
        Label88.Location = New Point(width_p1, height_p1 * 4)
        Label48.Location = New Point(width_p1, height_p1 * 5)
        x2 = 45
        y2 = 13
        Me.Label44.Width = x2 / a_width * Me.Width
        Me.Label44.Height = y2 / b_height * Me.Height
        Me.Label45.Width = x2 / a_width * Me.Width
        Me.Label45.Height = y2 / b_height * Me.Height
        Me.Label46.Width = x2 / a_width * Me.Width
        Me.Label46.Height = y2 / b_height * Me.Height
        Me.Label88.Width = x2 / a_width * Me.Width
        Me.Label88.Height = y2 / b_height * Me.Height
        Me.Label48.Width = x2 / a_width * Me.Width
        Me.Label48.Height = y2 / b_height * Me.Height




        x1 = 678
        y1 = 345
        x2 = 61
        y2 = 149
        Me.Panel3.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Panel3.Width = x2 / a_width * Me.Width
        Me.Panel3.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 8)
        Panel3.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Dim width_p3, height_p3 As Single
        width_p3 = Panel3.Width / 5
        height_p3 = Panel3.Height / 6
        Label49.Location = New Point(width_p3, height_p3)
        Label50.Location = New Point(width_p3, height_p3 * 2)
        Label51.Location = New Point(width_p3, height_p3 * 3)
        Label52.Location = New Point(width_p3, height_p3 * 4)
        Label53.Location = New Point(width_p3, height_p3 * 5)
        x2 = 45
        y2 = 13
        Me.Label49.Width = x2 / a_width * Me.Width
        Me.Label49.Height = y2 / b_height * Me.Height
        Me.Label50.Width = x2 / a_width * Me.Width
        Me.Label50.Height = y2 / b_height * Me.Height
        Me.Label51.Width = x2 / a_width * Me.Width
        Me.Label51.Height = y2 / b_height * Me.Height
        Me.Label52.Width = x2 / a_width * Me.Width
        Me.Label52.Height = y2 / b_height * Me.Height
        Me.Label53.Width = x2 / a_width * Me.Width
        Me.Label53.Height = y2 / b_height * Me.Height


        x1 = 740
        y1 = 22
        x2 = 205
        y2 = 153
        Me.Panel2.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Panel2.Width = x2 / a_width * Me.Width
        Me.Panel2.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 8)
        Panel2.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Dim width_p2, height_p2 As Single
        width_p2 = 0
        height_p2 = Panel2.Height / 6
        Label71.Location = New Point(width_p2, height_p2)
        Label72.Location = New Point(width_p2, height_p2 * 2)
        Label73.Location = New Point(width_p2, height_p2 * 3)
        Label74.Location = New Point(width_p2, height_p2 * 4)
        Label75.Location = New Point(width_p2, height_p2 * 5)
        x2 = 254
        y2 = 17
        Me.Label71.Width = x2 / a_width * Me.Width
        Me.Label71.Height = y2 / b_height * Me.Height
        Me.Label72.Width = x2 / a_width * Me.Width
        Me.Label72.Height = y2 / b_height * Me.Height
        Me.Label73.Width = x2 / a_width * Me.Width
        Me.Label73.Height = y2 / b_height * Me.Height
        Me.Label74.Width = x2 / a_width * Me.Width
        Me.Label74.Height = y2 / b_height * Me.Height
        Me.Label75.Width = x2 / a_width * Me.Width
        Me.Label75.Height = y2 / b_height * Me.Height



        x1 = 740
        y1 = 349
        x2 = 205
        y2 = 153
        Me.Panel4.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Panel4.Width = x2 / a_width * Me.Width
        Me.Panel4.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 8)
        Panel4.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Dim width_p4, height_p4 As Single
        width_p4 = 0
        height_p4 = Panel4.Height / 6
        Label20.Location = New Point(width_p4, height_p4)
        Label76.Location = New Point(width_p4, height_p4 * 2)
        Label77.Location = New Point(width_p4, height_p4 * 3)
        Label78.Location = New Point(width_p4, height_p4 * 4)
        Label79.Location = New Point(width_p4, height_p4 * 5)
        x2 = 300
        y2 = 17
        Me.Label20.Width = x2 / a_width * Me.Width
        Me.Label20.Height = y2 / b_height * Me.Height
        Me.Label76.Width = x2 / a_width * Me.Width
        Me.Label76.Height = y2 / b_height * Me.Height
        Me.Label77.Width = x2 / a_width * Me.Width
        Me.Label77.Height = y2 / b_height * Me.Height
        Me.Label78.Width = x2 / a_width * Me.Width
        Me.Label78.Height = y2 / b_height * Me.Height
        Me.Label79.Width = x2 / a_width * Me.Width
        Me.Label79.Height = y2 / b_height * Me.Height

        Dim c_font1, c_font2 As Integer
        x1 = 900
        y1 = 540
        x2 = 94
        y2 = 31
        Me.Button1.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Button1.Width = x2 / a_width * Me.Width
        Me.Button1.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 10)
        c_font1 = CInt(Me.Height / b_height * 10)
        c_font2 = CInt(Me.Width / a_width * 10)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If
        Button1.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 900
        y1 = 580
        x2 = 94
        y2 = 31
        Me.Button2.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Button2.Width = x2 / a_width * Me.Width
        Me.Button2.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 10)
        c_font1 = CInt(Me.Height / b_height * 10)
        c_font2 = CInt(Me.Width / a_width * 10)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If
        Button2.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)

        x1 = 900
        y1 = 620
        x2 = 94
        y2 = 31
        Me.Button3.Location = New Point(x1 / a_width * Me.Width, y1 / b_height * Me.Height)
        Me.Button3.Width = x2 / a_width * Me.Width
        Me.Button3.Height = y2 / b_height * Me.Height
        c_font = CInt(Me.Height / b_height * 10)
        c_font1 = CInt(Me.Height / b_height * 10)
        c_font2 = CInt(Me.Width / a_width * 10)
        If c_font1 < c_font2 Then
            c_font = c_font1
        Else
            c_font = c_font2
        End If
        Button3.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)


        Dim currentx2, currenty2, currentx3, currenty3, gapx, gapy As Single
        currentx2 = Me.Width * 0.04
        currenty3 = Me.Height * 0.01
        currenty2 = currenty3 + Me.Height * 0.45
        currentx3 = currentx2 + Me.Width * 0.6
        gapx = (currentx3 - currentx2) / 10
        gapy = (currenty2 - currenty3) / 5
        c_font = CInt(Me.Height / b_height * 9.75)
        Label21.Width = x2 / a_width * Me.Width
        Label21.Height = y2 / b_height * Me.Height
        Label22.Width = x2 / a_width * Me.Width
        Label22.Height = y2 / b_height * Me.Height
        Label23.Width = x2 / a_width * Me.Width
        Label23.Height = y2 / b_height * Me.Height
        Label24.Width = x2 / a_width * Me.Width
        Label24.Height = y2 / b_height * Me.Height
        Label25.Width = x2 / a_width * Me.Width
        Label25.Height = y2 / b_height * Me.Height
        Label26.Width = x2 / a_width * Me.Width
        Label26.Height = y2 / b_height * Me.Height
        Label27.Width = x2 / a_width * Me.Width
        Label27.Height = y2 / b_height * Me.Height
        Label28.Width = x2 / a_width * Me.Width
        Label28.Height = y2 / b_height * Me.Height
        Label29.Width = x2 / a_width * Me.Width
        Label29.Height = y2 / b_height * Me.Height
        Label30.Width = x2 / a_width * Me.Width
        Label30.Height = y2 / b_height * Me.Height
        Label31.Width = x2 / a_width * Me.Width
        Label31.Height = y2 / b_height * Me.Height
        Label32.Width = x2 / a_width * Me.Width
        Label32.Height = y2 / b_height * Me.Height


        Label21.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label21.Location = New Point(Width * 0.055 - Label21.Width, currenty3)
        Label22.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label22.Location = New Point(Width * 0.055 - Label21.Width, currenty3 + gapy)
        Label23.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label23.Location = New Point(Width * 0.055 - Label21.Width, currenty3 + gapy * 2)
        Label24.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label24.Location = New Point(Width * 0.055 - Label21.Width, currenty3 + gapy * 3)
        Label25.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label25.Location = New Point(Width * 0.055 - Label21.Width, currenty3 + gapy * 4)
        Label26.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label26.Location = New Point(Width * 0.055 - Label21.Width, currenty3 + gapy * 5)

        Dim currentx2n, currenty2n, currentx3n, currenty3n, gapxn, gapyn As Single
        currentx2n = currentx2
        currenty3n = Me.Height * 0.035 + currenty2
        currenty2n = (currenty2 - currenty3) + currenty3n
        currentx3n = currentx3

        gapxn = (currentx3n - currentx2n) / 10
        gapyn = (currenty2n - currenty3n) / 5
        Label27.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label27.Location = New Point(Width * 0.055 - Label21.Width, currenty3n)
        Label28.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label28.Location = New Point(Width * 0.055 - Label21.Width, currenty3n + gapyn)
        Label29.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label29.Location = New Point(Width * 0.055 - Label21.Width, currenty3n + gapyn * 2)
        Label30.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label30.Location = New Point(Width * 0.055 - Label21.Width, currenty3n + gapyn * 3)
        Label31.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label31.Location = New Point(Width * 0.055 - Label21.Width, currenty3n + gapyn * 4)
        Label32.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label32.Location = New Point(Width * 0.055 - Label21.Width, currenty3n + gapyn * 5)


        Label33.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label33.Location = New Point(currentx2 + gapxn * 0.1, currenty3n + gapyn * 5 + Me.Height * 0.02)
        Label34.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label34.Location = New Point(currentx2 + gapxn, currenty3n + gapyn * 5 + Me.Height * 0.02)
        Label35.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label35.Location = New Point(currentx2 + gapxn * 2, currenty3n + gapyn * 5 + Me.Height * 0.02)
        Label36.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label36.Location = New Point(currentx2 + gapxn * 3, currenty3n + gapyn * 5 + Me.Height * 0.02)
        Label37.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label37.Location = New Point(currentx2 + gapxn * 4, currenty3n + gapyn * 5 + Me.Height * 0.02)
        Label38.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label38.Location = New Point(currentx2 + gapxn * 5, currenty3n + gapyn * 5 + Me.Height * 0.02)
        Label39.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label39.Location = New Point(currentx2 + gapxn * 6, currenty3n + gapyn * 5 + Me.Height * 0.02)
        Label40.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label40.Location = New Point(currentx2 + gapxn * 7, currenty3n + gapyn * 5 + Me.Height * 0.02)
        Label41.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label41.Location = New Point(currentx2 + gapxn * 8, currenty3n + gapyn * 5 + Me.Height * 0.02)
        Label42.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label42.Location = New Point(currentx2 + gapxn * 9, currenty3n + gapyn * 5 + Me.Height * 0.02)
        Label43.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label43.Location = New Point(currentx2 + gapxn * 10, currenty3n + gapyn * 5 + Me.Height * 0.02)
        Label70.Font = New System.Drawing.Font(" Georgia", c_font, FontStyle.Regular)
        Label70.Location = New Point(currentx2 + gapxn * 10.5, currenty3n + gapyn * 5 + Me.Height * 0.02)

        Panel2.Controls.Item(0).Visible = False
        Panel2.Controls.Item(1).Visible = False
        Panel2.Controls.Item(2).Visible = False
        Panel2.Controls.Item(3).Visible = False
        Panel2.Controls.Item(4).Visible = False
        Panel4.Controls.Item(0).Visible = False
        Panel4.Controls.Item(1).Visible = False
        Panel4.Controls.Item(2).Visible = False
        Panel4.Controls.Item(3).Visible = False
        Panel4.Controls.Item(4).Visible = False
        Rm = R
        If Rm > 10 Then
            Ra = 5
            Timer1.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = True
        ElseIf Rm < 10 Then
            Ra = 5
            Timer1.Enabled = False
            Timer2.Enabled = True
            Timer4.Enabled = False
        End If

        'Open "C:\BMStrain\acceler.dat" For Input As #1
        FileOpen(1, Application.StartupPath + "\acceler.dat", OpenMode.Input)
        Dim Ac As Integer
        Input(1, Ac)
        FileClose(1)


        Select Case W
            Case 1
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

        'Open "C:\BMStrain\scope.dat" For Input As 6#
        FileOpen(6, Application.StartupPath + "\scope.dat", OpenMode.Input)
        ' Open "scope.dat" For Input As 6#
        Dim in1 As Integer
        Dim abb(4) As Single
        For in1 = 1 To 4
            Input(6, abb(in1))
        Next
        '  Input(6, Vmax1, Vmin1, Vmax2, Vmin2)
        FileClose(6)
        Vmax1 = abb(1)
        Vmin1 = abb(2)
        Vmax2 = abb(3)
        Vmin2 = abb(4)


        c1 = (Vmax1 - Vmin1) / 5.0#
        For Me.i = 1 To 6
            V1 = Vmin1 + c1 * (i - 1)
            Select Case i
                Case 1
                    Label26.Text = Str$(V1)
                Case 2
                    Label25.Text = Str$(V1)
                Case 3
                    Label24.Text = Str$(V1)
                Case 4
                    Label23.Text = Str$(V1)
                Case 5
                    Label22.Text = Str$(V1)
                Case Else
                    Label21.Text = Str$(V1)

            End Select
        Next i

        c2 = (Vmax2 - Vmin2) / 5.0#
        For Me.j = 1 To 6
            V2 = Vmin2 + c2 * (j - 1)
            Select Case j
                Case 1
                    Label32.Text = Str$(V2)
                Case 2
                    Label31.Text = Str$(V2)
                Case 3
                    Label30.Text = Str$(V2)
                Case 4
                    Label29.Text = Str$(V2)
                Case 5
                    Label28.Text = Str$(V2)
                Case Else
                    Label27.Text = Str$(V2)

            End Select
        Next j

        'Index = -1
       
        Timer3.Enabled = False
        Select Case Counter
            Case 1
                Panel1.Controls.Item(0).Visible = True
                Panel1.Controls.Item(1).Visible = False
                Panel1.Controls.Item(2).Visible = False
                Panel1.Controls.Item(3).Visible = False
                Panel1.Controls.Item(4).Visible = False
            Case 2
                Panel1.Controls.Item(0).Visible = True
                Panel1.Controls.Item(1).Visible = True
                Panel1.Controls.Item(2).Visible = False
                Panel1.Controls.Item(3).Visible = False
                Panel1.Controls.Item(4).Visible = False
            Case 3
                Panel1.Controls.Item(0).Visible = True
                Panel1.Controls.Item(1).Visible = True
                Panel1.Controls.Item(2).Visible = True
                Panel1.Controls.Item(3).Visible = False
                Panel1.Controls.Item(4).Visible = False
            Case 4
                Panel1.Controls.Item(0).Visible = True
                Panel1.Controls.Item(1).Visible = True
                Panel1.Controls.Item(2).Visible = True
                Panel1.Controls.Item(3).Visible = True
                Panel1.Controls.Item(4).Visible = False
            Case 5
                Panel1.Controls.Item(0).Visible = True
                Panel1.Controls.Item(1).Visible = True
                Panel1.Controls.Item(2).Visible = True
                Panel1.Controls.Item(3).Visible = True
                Panel1.Controls.Item(4).Visible = True
            Case Else
                Panel1.Controls.Item(0).Visible = False
                Panel1.Controls.Item(1).Visible = False
                Panel1.Controls.Item(2).Visible = False
                Panel1.Controls.Item(3).Visible = False
                Panel1.Controls.Item(4).Visible = False

        End Select

        Select Case Counter1
            Case 1
                Panel3.Controls.Item(0).Visible = True
                Panel3.Controls.Item(1).Visible = False
                Panel3.Controls.Item(2).Visible = False
                Panel3.Controls.Item(3).Visible = False
                Panel3.Controls.Item(4).Visible = False
            Case 2
                Panel3.Controls.Item(0).Visible = True
                Panel3.Controls.Item(1).Visible = True
                Panel3.Controls.Item(2).Visible = False
                Panel3.Controls.Item(3).Visible = False
                Panel3.Controls.Item(4).Visible = False
            Case 3
                Panel3.Controls.Item(0).Visible = True
                Panel3.Controls.Item(1).Visible = True
                Panel3.Controls.Item(2).Visible = True
                Panel3.Controls.Item(3).Visible = False
                Panel3.Controls.Item(4).Visible = False
            Case 4
                Panel3.Controls.Item(0).Visible = True
                Panel3.Controls.Item(1).Visible = True
                Panel3.Controls.Item(2).Visible = True
                Panel3.Controls.Item(3).Visible = True
                Panel3.Controls.Item(4).Visible = False
            Case 5
                Panel3.Controls.Item(0).Visible = True
                Panel3.Controls.Item(1).Visible = True
                Panel3.Controls.Item(2).Visible = True
                Panel3.Controls.Item(3).Visible = True
                Panel3.Controls.Item(4).Visible = True
            Case Else
                Panel3.Controls.Item(0).Visible = False
                Panel3.Controls.Item(1).Visible = False
                Panel3.Controls.Item(2).Visible = False
                Panel3.Controls.Item(3).Visible = False
                Panel3.Controls.Item(4).Visible = False

        End Select



    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        results_display2.Show()
        Counter1 = 0
        Counter = 0
        results_display2.CheckBox1.Checked = False
        results_display2.CheckBox2.Checked = False
        results_display2.CheckBox3.Checked = False
        results_display2.CheckBox4.Checked = False
        results_display2.CheckBox5.Checked = False
        results_display2.CheckBox6.Checked = False
        results_display2.CheckBox7.Checked = False
        results_display2.CheckBox8.Checked = False
        results_display2.CheckBox9.Checked = False
        results_display2.CheckBox10.Checked = False
        results_display2.CheckBox11.Checked = False
        results_display2.CheckBox12.Checked = False
        results_display2.CheckBox13.Checked = False
        results_display2.CheckBox14.Checked = False
        results_display2.CheckBox15.Checked = False
        results_display2.CheckBox16.Checked = False
        results_display2.CheckBox17.Checked = False
        results_display2.CheckBox18.Checked = False
        results_display2.CheckBox19.Checked = False
        results_display2.CheckBox20.Checked = False
        results_display2.CheckBox21.Checked = False
        results_display2.CheckBox22.Checked = False
        results_display2.CheckBox61.Checked = False
        results_display2.CheckBox62.Checked = False
        results_display2.CheckBox63.Checked = False
        results_display2.CheckBox64.Checked = False
        results_display2.CheckBox65.Checked = False
        results_display2.CheckBox66.Checked = False
        results_display2.CheckBox67.Checked = False
        results_display2.CheckBox68.Checked = False
        results_display2.CheckBox69.Checked = False
        results_display2.CheckBox70.Checked = False
        results_display2.CheckBox71.Checked = False
        results_display2.CheckBox72.Checked = False
        results_display2.CheckBox73.Checked = False
        results_display2.CheckBox74.Checked = False
        results_display2.CheckBox75.Checked = False
        results_display2.CheckBox76.Checked = False
        results_display2.CheckBox77.Checked = False
        results_display2.CheckBox78.Checked = False
        results_display2.CheckBox79.Checked = False
        results_display2.CheckBox80.Checked = False
        results_display2.CheckBox81.Checked = False
        results_display2.CheckBox31.Checked = False
        results_display2.CheckBox32.Checked = False
        results_display2.CheckBox33.Checked = False
        results_display2.CheckBox34.Checked = False
        results_display2.CheckBox35.Checked = False
        results_display2.CheckBox36.Checked = False
        results_display2.CheckBox37.Checked = False
        results_display2.CheckBox38.Checked = False
        results_display2.CheckBox39.Checked = False
        results_display2.CheckBox40.Checked = False
        results_display2.CheckBox41.Checked = False
        results_display2.CheckBox42.Checked = False
        results_display2.CheckBox43.Checked = False
        Me.Close()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        sca.Close()
        sca.Show()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim g As Graphics

        g = Me.CreateGraphics
        g.Clear(Me.BackColor)
        g.Dispose()

        Call curves_Load(Nothing, Nothing)

        Dim p As Pen = New Pen(Color.Gray, 1)
        '  Dim i As Integer
        p.DashStyle = Drawing2D.DashStyle.Dash
        Dim px As Pen = New Pen(Color.Black, 2)
        Dim currentx2, currenty2, currentx3, currenty3, gapx, gapy As Single
        currentx2 = Me.Width * 0.05
        currenty3 = Me.Height * 0.02
        currenty2 = currenty3 + Me.Height * 0.45
        currentx3 = currentx2 + Me.Width * 0.6
        currentx2 = Me.Width * 0.056

        gapx = (currentx3 - currentx2) / 10
        gapy = (currenty2 - currenty3) / 5
        Me.CreateGraphics.DrawLine(px, currentx2, currenty2, currentx3, currenty2)
        Me.CreateGraphics.DrawLine(px, currentx2, currenty3, currentx3, currenty3)
        Me.CreateGraphics.DrawLine(px, currentx2, currenty3, currentx2, currenty2)
        Me.CreateGraphics.DrawLine(px, currentx3, currenty3, currentx3, currenty2)
        For Me.i = 1 To 4
            Me.CreateGraphics.DrawLine(p, currentx2, currenty3 + gapy * i, currentx3, currenty3 + gapy * i)
        Next i

        For Me.i = 1 To 9
            Me.CreateGraphics.DrawLine(p, currentx2 + gapx * i, currenty2, currentx2 + gapx * i, currenty3)
        Next i


        Dim currentx2n, currenty2n, currentx3n, currenty3n, gapxn, gapyn As Single
        currentx2n = currentx2
        currenty3n = Me.Height * 0.035 + currenty2
        currenty2n = (currenty2 - currenty3) + currenty3n
        currentx3n = currentx3
        '  currenty3n = Me.Height * 0.05 + currenty2

        gapxn = (currentx3n - currentx2n) / 10
        gapyn = (currenty2n - currenty3n) / 5
        Me.CreateGraphics.DrawLine(px, currentx2n, currenty2n, currentx3n, currenty2n)
        Me.CreateGraphics.DrawLine(px, currentx2n, currenty3n, currentx3, currenty3n)
        Me.CreateGraphics.DrawLine(px, currentx2n, currenty3n, currentx2n, currenty2n)
        Me.CreateGraphics.DrawLine(px, currentx3n, currenty3n, currentx3n, currenty2n)
        For Me.i = 1 To 4
            Me.CreateGraphics.DrawLine(p, currentx2n, currenty3n + gapyn * i, currentx3n, currenty3n + gapyn * i)
        Next i

        For Me.i = 1 To 9
            Me.CreateGraphics.DrawLine(p, currentx2n + gapxn * i, currenty2n, currentx2 + gapxn * i, currenty3n)
        Next i
    End Sub

    Private Sub curves_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        'Dim index As Integer
        'For index = 0 To 4
        Panel2.Controls.Item(0).Visible = False
        Panel2.Controls.Item(1).Visible = False
        Panel2.Controls.Item(2).Visible = False
        Panel2.Controls.Item(3).Visible = False
        Panel2.Controls.Item(4).Visible = False
        Panel4.Controls.Item(0).Visible = False
        Panel4.Controls.Item(1).Visible = False
        Panel4.Controls.Item(2).Visible = False
        Panel4.Controls.Item(3).Visible = False
        Panel4.Controls.Item(4).Visible = False
        ' Next index

    End Sub

    Private Sub Label44_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label44.MouseMove
        Panel2.Controls.Item(0).Visible = True
    End Sub

    Private Sub Label45_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label45.MouseMove
        Panel2.Controls.Item(1).Visible = True
    End Sub

    Private Sub Label46_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label46.MouseMove
        Panel2.Controls.Item(2).Visible = True
    End Sub

    Private Sub Label88_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label88.MouseMove
        Panel2.Controls.Item(3).Visible = True
    End Sub

    Private Sub Label48_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label48.MouseMove
        Panel2.Controls.Item(4).Visible = True
    End Sub

    Private Sub Label49_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label49.MouseMove
        Panel4.Controls.Item(0).Visible = True
    End Sub

    Private Sub Label50_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label50.MouseMove
        Panel4.Controls.Item(1).Visible = True
    End Sub

    Private Sub Label51_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label51.MouseMove
        Panel4.Controls.Item(2).Visible = True
    End Sub

    Private Sub Label52_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label52.MouseMove
        Panel4.Controls.Item(3).Visible = True
    End Sub

    Private Sub Label53_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label53.MouseMove
        Panel4.Controls.Item(4).Visible = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer2.Enabled = False

        Call readH()
        Select Case Counter
            Case 1
                n1 = Flag(1)
                Call drawline1(n1)
                Panel1.Controls.Item(0).ForeColor = System.Drawing.Color.Red
            Case 2
                n1 = Flag(1)
                Call drawline1(n1)
                n2 = Flag(2)
                Call drawline2(n2)
                Panel1.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                Panel1.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
            Case 3
                n1 = Flag(1)
                Call drawline1(n1)
                n2 = Flag(2)
                Call drawline2(n2)
                n3 = Flag(3)
                Call drawline3(n3)
                Panel1.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                Panel1.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                Panel1.Controls.Item(2).ForeColor = System.Drawing.Color.Cyan
            Case 4
                n1 = Flag(1)
                Call drawline1(n1)
                n2 = Flag(2)
                Call drawline2(n2)
                n3 = Flag(3)
                Call drawline3(n3)
                n4 = Flag(4)
                Call drawline4(n4)
                Panel1.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                Panel1.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                Panel1.Controls.Item(2).ForeColor = System.Drawing.Color.Cyan
                Panel1.Controls.Item(3).ForeColor = System.Drawing.Color.Black
            Case 5
                n1 = Flag(1)
                Call drawline1(n1)
                n2 = Flag(2)
                Call drawline2(n2)
                n3 = Flag(3)
                Call drawline3(n3)
                n4 = Flag(4)
                Call drawline4(n4)
                n5 = Flag(5)
                Call drawline5(n5)
                Panel1.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                Panel1.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                Panel1.Controls.Item(2).ForeColor = System.Drawing.Color.Cyan
                Panel1.Controls.Item(3).ForeColor = System.Drawing.Color.Black
                Panel1.Controls.Item(4).ForeColor = System.Drawing.Color.Green
            Case Else
        End Select

        Select Case Counter1
            Case 1
                n6 = Flag1(1)
                Call drawline6(n6)
                Panel3.Controls.Item(0).ForeColor = System.Drawing.Color.Red
            Case 2
                n6 = Flag1(1)
                Call drawline6(n6)
                n7 = Flag1(2)
                Call drawline7(n7)
                Panel3.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                Panel3.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
            Case 3
                n6 = Flag1(1)
                Call drawline6(n6)
                n7 = Flag1(2)
                Call drawline7(n7)
                n8 = Flag1(3)
                Call drawline8(n8)
                Panel3.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                Panel3.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                Panel3.Controls.Item(2).ForeColor = System.Drawing.Color.Cyan
            Case 4
                n6 = Flag1(1)
                Call drawline6(n6)
                n7 = Flag1(2)
                Call drawline7(n7)
                n8 = Flag1(3)
                Call drawline8(n8)
                n9 = Flag1(4)
                Call drawline9(n9)
                Panel3.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                Panel3.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                Panel3.Controls.Item(2).ForeColor = System.Drawing.Color.Cyan
                Panel3.Controls.Item(3).ForeColor = System.Drawing.Color.Black
            Case 5
                n6 = Flag1(1)
                Call drawline6(n6)
                n7 = Flag1(2)
                Call drawline7(n7)
                n8 = Flag1(3)
                Call drawline8(n8)
                n9 = Flag1(4)
                Call drawline9(n9)
                n10 = Flag1(5)
                Call drawline10(n10)
                Panel3.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                Panel3.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                Panel3.Controls.Item(2).ForeColor = System.Drawing.Color.Cyan
                Panel3.Controls.Item(3).ForeColor = System.Drawing.Color.Black
                Panel3.Controls.Item(4).ForeColor = System.Drawing.Color.Green
            Case Else
        End Select

        Select Case W
            Case 1
                If Ra > 30 And Hn = 7200 Then
                    Timer1.Enabled = False
                End If
            Case 2
                If Ra > 30 And Hn = 14400 Then
                    Timer1.Enabled = False
                End If
            Case 3
                If Ra > 30 And Hn = 28800 Then
                    Timer1.Enabled = False
                End If
            Case 4
                If Ra > 30 And Hn = 43200 Then
                    Timer1.Enabled = False
                End If
            Case Else
        End Select

        Ho = Hn
        Ra = Ra + 1

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Dim G As Single

        frmspread.grdspread.Row = 7
        frmspread.grdspread.Col = 10
        G = Val(frmspread.grdspread.Text)
        If G <> 0.0# Then
            Timer1.Enabled = True
        Else
        End If

    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick

    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        If Ra < Rm Then
            Call readH()
            Select Case Counter
                Case 1
                    n1 = Flag(1)
                    Call drawline1(n1)
                    Panel1.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                Case 2
                    n1 = Flag(1)
                    Call drawline1(n1)
                    n2 = Flag(2)
                    Call drawline2(n2)
                    Panel1.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                    Panel1.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                Case 3
                    n1 = Flag(1)
                    Call drawline1(n1)
                    n2 = Flag(2)
                    Call drawline2(n2)
                    n3 = Flag(3)
                    Call drawline3(n3)
                    Panel1.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                    Panel1.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                    Panel1.Controls.Item(2).ForeColor = System.Drawing.Color.Cyan
                Case 4
                    n1 = Flag(1)
                    Call drawline1(n1)
                    n2 = Flag(2)
                    Call drawline2(n2)
                    n3 = Flag(3)
                    Call drawline3(n3)
                    n4 = Flag(4)
                    Call drawline4(n4)
                    Panel1.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                    Panel1.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                    Panel1.Controls.Item(2).ForeColor = System.Drawing.Color.Cyan
                    Panel1.Controls.Item(3).ForeColor = System.Drawing.Color.Black
                Case 5
                    n1 = Flag(1)
                    Call drawline1(n1)
                    n2 = Flag(2)
                    Call drawline2(n2)
                    n3 = Flag(3)
                    Call drawline3(n3)
                    n4 = Flag(4)
                    Call drawline4(n4)
                    n5 = Flag(5)
                    Call drawline5(n5)
                    Panel1.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                    Panel1.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                    Panel1.Controls.Item(2).ForeColor = System.Drawing.Color.Cyan
                    Panel1.Controls.Item(3).ForeColor = System.Drawing.Color.Black
                    Panel1.Controls.Item(4).ForeColor = System.Drawing.Color.LightGreen
                Case Else
            End Select

            Select Case Counter1
                Case 1
                    n6 = Flag1(1)
                    Call drawline6(n6)
                    Panel3.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                Case 2
                    n6 = Flag1(1)
                    Call drawline6(n6)
                    n7 = Flag1(2)
                    Call drawline7(n7)
                    Panel3.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                    Panel3.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                Case 3
                    n6 = Flag1(1)
                    Call drawline6(n6)
                    n7 = Flag1(2)
                    Call drawline7(n7)
                    n8 = Flag1(3)
                    Call drawline8(n8)
                    Panel3.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                    Panel3.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                    Panel3.Controls.Item(2).ForeColor = System.Drawing.Color.Cyan
                Case 4
                    n6 = Flag1(1)
                    Call drawline6(n6)
                    n7 = Flag1(2)
                    Call drawline7(n7)
                    n8 = Flag1(3)
                    Call drawline8(n8)
                    n9 = Flag1(4)
                    Call drawline9(n9)
                    Panel3.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                    Panel3.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                    Panel3.Controls.Item(2).ForeColor = System.Drawing.Color.Cyan
                    Panel3.Controls.Item(3).ForeColor = System.Drawing.Color.Black
                Case 5
                    n6 = Flag1(1)
                    Call drawline6(n6)
                    n7 = Flag1(2)
                    Call drawline7(n7)
                    n8 = Flag1(3)
                    Call drawline8(n8)
                    n9 = Flag1(4)
                    Call drawline9(n9)
                    n10 = Flag1(5)
                    Call drawline10(n10)
                    Panel3.Controls.Item(0).ForeColor = System.Drawing.Color.Red
                    Panel3.Controls.Item(1).ForeColor = System.Drawing.Color.Orange
                    Panel3.Controls.Item(2).ForeColor = System.Drawing.Color.Cyan
                    Panel3.Controls.Item(3).ForeColor = System.Drawing.Color.Black
                    Panel3.Controls.Item(4).ForeColor = System.Drawing.Color.Green
                Case Else
            End Select

            Select Case W
                Case 1
                    If Ra > 30 And Hn = 7200 Then
                        Timer4.Enabled = False
                    End If
                Case 2
                    If Ra > 30 And Hn = 14400 Then
                        Timer4.Enabled = False
                    End If
                Case 3
                    If Ra > 30 And Hn = 28800 Then
                        Timer4.Enabled = False
                    End If
                Case 4
                    If Ra > 30 And Hn = 43200 Then
                        Timer4.Enabled = False
                    End If
                Case Else
            End Select

            Ho = Hn
            Ra = Ra + 1

        ElseIf Ra = Rm Then
            Ra = Rm
            Timer4.Enabled = False
            Timer1.Enabled = True

        End If

    End Sub

    Private Sub curves_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Dim p As Pen = New Pen(Color.Gray, 1)
        '  Dim i As Integer
        p.DashStyle = Drawing2D.DashStyle.Dash
        Dim px As Pen = New Pen(Color.Black, 2)
        Dim currentx2, currenty2, currentx3, currenty3, gapx, gapy As Single
        'add
        currentx2 = Me.Width * 0.05
        currenty3 = Me.Height * 0.02
        currenty2 = currenty3 + Me.Height * 0.45
        currentx3 = currentx2 + Me.Width * 0.6
        currentx2 = Me.Width * 0.056

        gapx = (currentx3 - currentx2) / 10
        gapy = (currenty2 - currenty3) / 5
        Me.CreateGraphics.DrawLine(px, currentx2, currenty2, currentx3, currenty2)
        Me.CreateGraphics.DrawLine(px, currentx2, currenty3, currentx3, currenty3)
        Me.CreateGraphics.DrawLine(px, currentx2, currenty3, currentx2, currenty2)
        Me.CreateGraphics.DrawLine(px, currentx3, currenty3, currentx3, currenty2)
        For Me.i = 1 To 4
            Me.CreateGraphics.DrawLine(p, currentx2, currenty3 + gapy * i, currentx3, currenty3 + gapy * i)
        Next i

        For Me.i = 1 To 9
            Me.CreateGraphics.DrawLine(p, currentx2 + gapx * i, currenty2, currentx2 + gapx * i, currenty3)
        Next i
        'add
        ' currentx2 = Me.Width * 0.05
        Dim currentx2n, currenty2n, currentx3n, currenty3n, gapxn, gapyn As Single
        currentx2n = currentx2
        currenty3n = Me.Height * 0.035 + currenty2
        currenty2n = (currenty2 - currenty3) + currenty3n
        currentx3n = currentx3
        ' currentx2n = Me.Width * 0.058
        '  currenty3n = Me.Height * 0.05 + currenty2

        gapxn = (currentx3n - currentx2n) / 10
        gapyn = (currenty2n - currenty3n) / 5
        Me.CreateGraphics.DrawLine(px, currentx2n, currenty2n, currentx3n, currenty2n)
        Me.CreateGraphics.DrawLine(px, currentx2n, currenty3n, currentx3n, currenty3n)
        Me.CreateGraphics.DrawLine(px, currentx2n, currenty3n, currentx2n, currenty2n)
        Me.CreateGraphics.DrawLine(px, currentx3n, currenty3n, currentx3n, currenty2n)
        For Me.i = 1 To 4
            Me.CreateGraphics.DrawLine(p, currentx2n, currenty3n + gapyn * i, currentx3n, currenty3n + gapyn * i)
        Next i

        For Me.i = 1 To 9
            Me.CreateGraphics.DrawLine(p, currentx2n + gapxn * i, currenty2n, currentx2 + gapxn * i, currenty3n)
        Next i

    End Sub

    Private Sub curves_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        '  On Error Resume Next
        Me.Width = 973
        Me.Height = 672
    End Sub

    Private Sub Panel3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub Label47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label47.Click

    End Sub
End Class