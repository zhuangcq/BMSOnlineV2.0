Public Class sca
    Dim c1, c2, V1, V2 As Single
    Dim i, j As Integer

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        curves.Timer1.Enabled = False
        Rm = R
        Ra = 5

        Vmax1 = Val(Text1.Text)
        Vmin1 = Val(Text2.Text)
        Vmax2 = Val(Text3.Text)
        Vmin2 = Val(Text4.Text)

        'Open "C:\BMStrain\scope.dat" For Output As #1
        ' Open "scope.dat" For Output As #1
        FileOpen(1, Application.StartupPath + "\scope.dat", OpenMode.Output)
        Print(1, Vmax1, Vmin1, Vmax2, Vmin2)
        FileClose(1)

        c1 = (Vmax1 - Vmin1) / 5.0#
        For Me.i = 1 To 6
            V1 = Vmin1 + c1 * (Me.i - 1)
            Select Case Me.i
                Case 1
                    curves.Label26.Text = Str$(V1)
                Case 2
                    curves.Label25.Text = Str$(V1)
                Case 3
                    curves.Label24.Text = Str$(V1)
                Case 4
                    curves.Label23.Text = Str$(V1)
                Case 5
                    curves.Label22.Text = Str$(V1)
                Case Else
                    curves.Label21.Text = Str$(V1)
            End Select
        Next i

        c2 = (Vmax2 - Vmin2) / 5.0#
        For Me.j = 1 To 6
            V2 = Vmin2 + c2 * (Me.j - 1)
            Select Case Me.j
                Case 1
                    curves.Label32.Text = Str$(V2)
                Case 2
                    curves.Label31.Text = Str$(V2)
                Case 3
                    curves.Label30.Text = Str$(V2)
                Case 4
                    curves.Label29.Text = Str$(V2)
                Case 5
                    curves.Label28.Text = Str$(V2)
                Case Else
                    curves.Label27.Text = Str$(V2)
            End Select
        Next j

        'Select Case W
        'Case 1
        '   If Rm > 30 And Hn = 1800 Then
        '      curves!Timer1.Enabled = False
        '      curves!Timer4.Enabled = True
        '    ElseIf Hn < 1800 Then
        '       curves!Timer4.Enabled = True
        '    End If
        'Case 2
        '    If Rm > 30 And Hn = 3600 Then
        '       curves!Timer1.Enabled = False
        '       curves!Timer4.Enabled = True
        '    ElseIf Hn < 1800 Then
        '       curves!Timer4.Enabled = True
        '    End If
        'Case 3
        '    If Rm > 30 And Hn = 7200 Then
        '       curves!Timer1.Enabled = False
        '       curves!Timer4.Enabled = True
        '    ElseIf Hn < 1800 Then
        '       'curves!Timer1.Enabled = False
        '       curves!Timer4.Enabled = True
        '    End If
        'Case 4
        '    If Rm > 30 And Hn = 14400 Then
        '       curves!Timer1.Enabled = False
        '       curves!Timer4.Enabled = True
        '    ElseIf Hn < 1800 Then
        '       'curves!Timer1.Enabled = False
        '       curves!Timer4.Enabled = True
        '    End If
        'Case 5
        '    If Rm > 30 And Hn = 28800 Then
        '       curves!Timer1.Enabled = False
        '       curves!Timer4.Enabled = True
        '    ElseIf Hn < 1800 Then
        '       'curves!Timer1.Enabled = False
        '       curves!Timer4.Enabled = True
        '    End If
        'Case 6
        '    If Rm > 30 And Hn = 43200 Then
        '       curves!Timer1.Enabled = False
        '       curves!Timer4.Enabled = True
        '    ElseIf Hn < 1800 Then
        '       'curves!Timer1.Enabled = False
        '       curves!Timer4.Enabled = True
        '    End If
        'Case Else
        'End Select
        '
        ' Unload(sca)
        Me.Hide()
        curves.Button1.PerformClick()
        'curves.Cls
        ' Unload(curves)
        ' curves.Hide()
        ' Load(curves)
        ' curves.Show()

    End Sub

    Private Sub sca_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = False
        Text1.Text = Str$(10.0#)
        Text2.Text = Str$(0.0#)
        Text3.Text = Str$(30.0#)
        Text4.Text = Str$(0.0#)

        MaximizeBox = False
        MinimizeBox = False



    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        curves.Show()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Text1.Text = Str$(10.0#)
        Text2.Text = Str$(0.0#)
        Text3.Text = Str$(30.0#)
        Text4.Text = Str$(0.0#)


    End Sub
End Class