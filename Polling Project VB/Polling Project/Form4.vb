Public Class Form4

    Dim Con As New OleDb.OleDbConnection
    Dim Da As New OleDb.OleDbDataAdapter
    Dim Dt As New DataTable

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If Con.State = ConnectionState.Open Then
            Con.Close()
        End If

        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub Form4_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Funcs.threadstate = 1 Then
            Funcs.threadstate = 0
            Funcs.thrd.Abort()
        End If
        Application.Exit()
    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Funcs.ConStuStr = ""
        Funcs.ConTchStr1 = ""
        Funcs.ConTchStr2 = ""

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Button1.Enabled = False
        Button2.Enabled = False

        If ComboBox1.SelectedItem Is Nothing Then
            MsgBox("Please Select A Database!", MsgBoxStyle.Exclamation)
            Button1.Enabled = True
            Button2.Enabled = True
            Return
        End If

        If Form1.RadioButton2.Checked = True Then

            If Form1.ComboBox3.SelectedItem = "Teacher" Then

                Funcs.ConTchStr1 = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & Application.StartupPath & "\Data\" & Form1.TextBox4.Text & "\" & ComboBox1.SelectedItem.ToString & ".mdb; Jet OLEDB:Database Password=encrypted"
                Me.Hide()
                Button1.Enabled = True
                Button2.Enabled = True
                Form5.Show()

            ElseIf Form1.ComboBox3.SelectedItem = "Student" Then

                Dt.Clear()

                Dim x As String

                x = ComboBox1.SelectedItem.ToString.Remove(0, 19)

                Funcs.ConStuStr = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & Application.StartupPath & "\Data\" & x & "\" & ComboBox1.SelectedItem.ToString & ".mdb; Jet OLEDB:Database Password=encrypted"
                Con = New OleDb.OleDbConnection
                Con.ConnectionString = Funcs.ConStuStr

                If Not Con.State = ConnectionState.Open Then
                    Con.Open()
                End If

                Da = New Data.OleDb.OleDbDataAdapter("SELECT * from VoterList", Con)
                Da.Fill(Dt)


                Dim d As Integer = Dt.Rows.Count
                Dim y As Integer
                Dim caught As Integer = 0

                For y = 0 To (d - 1)
                    If Dt.Rows(y).Item(0).ToString = Form1.TextBox3.Text Then
                        MsgBox("Enrollment Number Entry Already Present, Please Try A Different Database", MsgBoxStyle.Critical)
                        caught = 1
                        Button1.Enabled = True
                        Button2.Enabled = True
                        Return
                    End If
                Next

                If Not caught = 1 Then
                    Me.Hide()
                    Button1.Enabled = True
                    Button2.Enabled = True
                    Form5.Show()
                End If


            End If

        ElseIf Form1.RadioButton1.Checked = True Then

            

            Dt.Clear()
            Funcs.ConTchStr2 = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & Application.StartupPath & "\Data\" & Form1.TextBox1.Text & "\" & ComboBox1.SelectedItem.ToString & ".mdb; Jet OLEDB:Database Password=encrypted"

            Con = New OleDb.OleDbConnection
            Con.ConnectionString = Funcs.ConTchStr2

            If Not Con.State = ConnectionState.Open Then
                Con.Open()
            End If

            Da = New Data.OleDb.OleDbDataAdapter("SELECT * from CandidateList", Con)
            Da.Fill(Dt)

            Dim d As Integer = Dt.Rows.Count

            If d = 0 Then
                Form6.ListBox1.Refresh()
                Form6.ListBox2.Refresh()
            End If

            Dim y As Integer

            For y = 0 To (d - 1)
                If Not String.IsNullOrEmpty(Dt.Rows(y).Item(6).ToString()) Then
                    Form6.ListBox1.Items.Add(Dt.Rows(y).Item(6).ToString())
                End If
            Next

            Dim index As Integer
            Dim itemcount As Integer = Form6.ListBox1.Items.Count

            If itemcount > 1 Then
                Dim lastitem As String = Form6.ListBox1.Items(itemcount - 1)

                For index = itemcount - 2 To 0 Step -1
                    If Form6.ListBox1.Items(index) = lastitem Then
                        Form6.ListBox1.Items.RemoveAt(index)
                    Else
                        lastitem = Form6.ListBox1.Items(index)
                    End If
                Next
            End If

            Form6.ListBox1.Sorted = True
            Form6.ListBox1.Refresh()

            Me.Hide()
            Button1.Enabled = True
            Button2.Enabled = True
            Form6.Show()

        End If


    End Sub
End Class