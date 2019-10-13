Public Class Form2

    Dim Con2 As New OleDb.OleDbConnection
    Dim Da1 As New OleDb.OleDbDataAdapter
    Dim Dt As New DataTable

    Private Sub Form2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If funcs.threadstate = 1 Then
            funcs.threadstate = 0
            funcs.thrd.Abort()
        End If
        Application.Exit()
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Label4.Text = "Empty!"

        Con2 = New OleDb.OleDbConnection
        Con2.ConnectionString = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & Application.StartupPath & "\Data\UserList.mdb; Jet OLEDB:Database Password=encrypted"

        If Not Con2.State = ConnectionState.Open Then
            Con2.Open()
        End If

        Da1 = New Data.OleDb.OleDbDataAdapter("SELECT * from UserList", Con2)
        Da1.Fill(Dt)



    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        DateTimePicker1.Text = Today.Date
        TextBox1.Focus()


    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

        If TextBox1.Text = "" Then
            Label4.Text = "Empty!"



        Else

            Dim x As Integer = Dt.Rows.Count
            Dim y As Integer
            Dim caught As Integer = 0

            For y = 0 To (x - 1)
                If Dt.Rows(y).Item(1).ToString = TextBox1.Text Then
                    Label4.Text = "Username Already Taken"
                    caught = 1
                End If
            Next

            If Not caught = 1 Then
                Label4.Text = "Username Available"
            End If


        End If



    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Label4.Text = "Username Available" Then
            If Not TextBox2.Text = "" Then
                If Not TextBox3.Text = "" Then
                    If Not TextBox4.Text = "" Then
                        If Not DateTimePicker1.Text < Today Then
                            MsgBox("Invalid Date Of Birth, Please Enter A Valid Date Of Birth", MsgBoxStyle.Critical)
                            Return
                        Else
                            Dim cmd As New OleDb.OleDbCommand

                            cmd.Connection = Con2
                            cmd.CommandText = "INSERT INTO [UserList] ([Username], [Password], [DB_Name], [Ori_Name], [DOB]) VALUES(@username, @password, @database, @name, @dob)"
                            cmd.Parameters.AddWithValue("@username", TextBox1.Text)
                            cmd.Parameters.AddWithValue("@password", TextBox2.Text)
                            cmd.Parameters.AddWithValue("@database", TextBox3.Text)
                            cmd.Parameters.AddWithValue("@name", TextBox4.Text)
                            cmd.Parameters.AddWithValue("@dob", DateTimePicker1.Text)
                            cmd.ExecuteNonQuery()

                            If Not System.IO.Directory.Exists(Application.StartupPath & "\Data\" & TextBox1.Text & "\") Then
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Data\" & TextBox1.Text & "\")
                            End If

                            If Not System.IO.Directory.Exists(Application.StartupPath & "\Data\" & TextBox1.Text & "\" & TextBox3.Text & "_ImageData\") Then
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Data\" & TextBox1.Text & "\" & TextBox3.Text & "_ImageData\")
                            End If

                            System.IO.File.Copy(Application.StartupPath & "\Data\Templates\CandidateList.mdb", Application.StartupPath & "\Data\" & TextBox1.Text & "\CandidateList" & "_" & TextBox3.Text & ".mdb", True)
                            System.IO.File.Copy(Application.StartupPath & "\Data\Templates\VoterList_Students.mdb", Application.StartupPath & "\Data\" & TextBox1.Text & "\VoterList_Students" & "_" & TextBox3.Text & ".mdb", True)
                            System.IO.File.Copy(Application.StartupPath & "\Data\Templates\VoterList_Teachers.mdb", Application.StartupPath & "\Data\" & TextBox1.Text & "\VoterList_Teachers" & "_" & TextBox3.Text & ".mdb", True)

                            MsgBox("Account Successfully Created", MsgBoxStyle.Information)
                            Me.Hide()
                            Form1.Show()

                        End If
                    Else
                        MsgBox("Name Field Is Empty, Please Enter A Valid Name", MsgBoxStyle.Critical)
                        Return
                    End If
                Else
                    MsgBox("Database Name Field Is Empty, Please Enter A Valid Database Name", MsgBoxStyle.Critical)
                    Return
                End If
            Else
                MsgBox("Password Field Is Empty, Please Enter A Valid Password", MsgBoxStyle.Critical)
                Return
            End If
        ElseIf Label4.Text = "Empty!" Then
            MsgBox("Username Field Is Empty, Please Enter A Valid Username", MsgBoxStyle.Critical)
            Return
        ElseIf Label4.Text = "Username Already Taken" Then
            MsgBox("Username Already Taken, Please Try A Different Username", MsgBoxStyle.Critical)
            Return
        End If

    End Sub
End Class