Public Class Form7
    Dim Con As New OleDb.OleDbConnection
    Dim cmd As New OleDb.OleDbCommand
    Dim Da As New OleDb.OleDbDataAdapter
    Dim Dt As New DataTable

    Private Sub Form7_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Funcs.threadstate = 1 Then
            Funcs.threadstate = 0
            Funcs.thrd.Abort()
        End If
        Application.Exit()
    End Sub

    Private Sub Form7_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Con.ConnectionString = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & Application.StartupPath & "\Data\UserList.mdb; Jet OLEDB:Database Password=encrypted"
        cmd.Connection = Con

        Da = New Data.OleDb.OleDbDataAdapter("SELECT * from UserList", Con)
        Da.Fill(Dt)

        Label5.Text = "Empty!"

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
        Form6.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("1 Or More Entry Fields Have Been Left Blank", MsgBoxStyle.Critical)
            Return
        End If


        Dim x As Integer = Dt.Rows.Count
        Dim y As Integer
        Dim PW As String

        Dim caught As Integer = 0

        For y = 0 To (x - 1)
            If Dt.Rows(y).Item(1).ToString = TextBox1.Text Then
                caught = 1
                PW = Dt.Rows(y).Item(2).ToString
            End If
        Next

        If caught = 1 And Label5.Text = "New Passwords Match" And TextBox2.Text = PW Then

            If Not Con.State = ConnectionState.Open Then
                Con.Open()
            End If


            cmd.CommandText = "UPDATE UserList SET [Password] = @Password WHERE [Username] = @Username"
            cmd.Parameters.AddWithValue("@Password", TextBox3.Text)
            cmd.Parameters.AddWithValue("@Username", TextBox1.Text)
            cmd.ExecuteNonQuery()
            MsgBox("Password Entry Successfully Updated", MsgBoxStyle.Information)
            cmd.Dispose()
            Me.Hide()
            Form6.Show()
        Else

            MsgBox("Invalid Details Provided", MsgBoxStyle.Critical)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Tries Added In Next Update'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If



    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

        If TextBox3.Text = "" Or TextBox4.Text = "" Then
            Label5.Text = "Empty!"
            Return
        End If

        If TextBox4.Text = TextBox3.Text Then
            Label5.Text = "New Passwords Match"
        Else
            Label5.Text = "Do Not Match"
        End If

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Or TextBox4.Text = "" Then
            Label5.Text = "Empty!"
            Return
        End If

        If TextBox4.Text = TextBox3.Text Then
            Label5.Text = "New Passwords Match"
        Else
            Label5.Text = "Do Not Match"
        End If
    End Sub
End Class