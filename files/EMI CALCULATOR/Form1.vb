Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim a, b, c As Double
        Dim d As Integer

        If IsNumeric(TextBox1.Text) Then
            a = Convert.ToDouble(TextBox1.Text)
        Else
            MsgBox("Invalid amount, please re-enter")
            TextBox1.Focus()
            TextBox1.SelectAll()

        End If
        If IsNumeric(TextBox2.Text) Then
            d = Convert.ToInt32(TextBox2.Text)
        Else
            MsgBox("Please re-enter the number in months")
            TextBox1.Focus()
            TextBox1.SelectAll()

        End If
        If IsNumeric(TextBox3.Text) Then
            b = 0.01 * Convert.ToDouble(TextBox3.Text) / 12
        Else
            MsgBox("Invalid interest rate, please re-enter")
            TextBox1.Focus()
            TextBox1.SelectAll()

        End If
        c = Pmt(b, d, -a, 0)
        TextBox4.Text = c.ToString("#.00")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()

    End Sub
End Class
