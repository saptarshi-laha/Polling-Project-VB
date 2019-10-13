Public Class Form11

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not TextBox1.Text = Nothing Then
            If Button1.Text = "Add Criteria" Then
                Form6.ListBox1.Items.Add(TextBox1.Text)
                TextBox1.Text = ""
                Me.Hide()
            ElseIf Button1.Text = "Add Candidate" Then
                Form6.ListBox2.Items.Add(TextBox1.Text)
                TextBox1.Text = ""
                Form6.TextBox2.Text = "NewCand"
                Me.Hide()
            End If
        Else
            MsgBox("Textbox Is Empty", MsgBoxStyle.Critical)
            Return
        End If
    End Sub
End Class