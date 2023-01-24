Public Class frmspread

    Private Sub grdspread_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grdspread.Enter
        



    End Sub

    Private Sub frmspread_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Static Item(59) As String
        '    Dim i, k, l As Integer
        '     Dim j As Integer
        Dim i As Integer

        For i = 1 To 399
            grdspread.Col = 0
            grdspread.Row = i
            grdspread.Text = Str$(i)
        Next i


        grdspread.Row = 0
        For i = 1 To 59
            grdspread.Col = i
            grdspread.Text = Str$(i)
        Next i

    End Sub
End Class