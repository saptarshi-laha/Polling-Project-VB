Public Class Form6

    Dim Con As New OleDb.OleDbConnection
    Dim Da As New OleDb.OleDbDataAdapter
    Dim Dt As New DataTable
    Dim cmd As New OleDb.OleDbCommand
    Dim ID As Integer
    Dim Votes As Integer

    Private Sub Form6_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Funcs.threadstate = 1 Then
            Funcs.threadstate = 0
            Funcs.thrd.Abort()
        End If
        Application.Exit()
    End Sub

    Private Sub Form6_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox2.Hide()
        TextBox2.Text = ""
        Label6.Text = Form1.TextBox1.Text
        Label10.Text = Form1.ComboBox2.SelectedItem.ToString

        Dim files() As String
        files = System.IO.Directory.GetFiles(Application.StartupPath & "\Data\" & Label6.Text & "\", "*.mdb", IO.SearchOption.TopDirectoryOnly)
        Dim total As Integer = 0

        For Each FileName As String In files
            total = total + 1
        Next

        Label12.Text = total.ToString

        Con = New OleDb.OleDbConnection
        Con.ConnectionString = Funcs.ConTchStr2

        If Not Con.State = ConnectionState.Open Then
            Con.Open()
        End If

        Da = New Data.OleDb.OleDbDataAdapter("SELECT * from CandidateList", Con)
        Da.Fill(Dt)

        ComboBox1.Items.Add("Air")
        ComboBox1.Items.Add("Earth")
        ComboBox1.Items.Add("Fire")
        ComboBox1.Items.Add("Water")

        DateTimePicker1.Text = Today

        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False


    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

        ListBox2.Items.Clear()

        Dim x As Integer = 0
        Dim y As Integer = Dt.Rows.Count

        For x = 0 To (y - 1)
            If Dt.Rows(x).Item(6).ToString = ListBox1.SelectedItem.ToString Then
                ListBox2.Items.Add(Dt.Rows(x).Item(1).ToString)
            End If
        Next

        ListBox2.Sorted = True

        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False

    End Sub

    Public Sub UpdateList()

        ListBox2.Items.Clear()
        Dt.Clear()
        Da.Fill(Dt)

        Dim x As Integer = 0
        Dim y As Integer = Dt.Rows.Count

        For x = 0 To (y - 1)
            If Dt.Rows(x).Item(6).ToString = ListBox1.SelectedItem.ToString Then
                ListBox2.Items.Add(Dt.Rows(x).Item(1).ToString)
            End If
        Next

        ListBox2.Sorted = True


    End Sub

    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged

        TextBox1.Text = ""
        DateTimePicker1.Text = Today
        ComboBox1.SelectedItem = Nothing
        TextBox4.Text = ""

        Dim x As Integer = 0
        Dim y As Integer = Dt.Rows.Count

        For x = 0 To (y - 1)
            If Dt.Rows(x).Item(6).ToString = ListBox1.SelectedItem.ToString Then
                TextBox1.Text = Dt.Rows(x).Item(1).ToString
                DateTimePicker1.Text = Dt.Rows(x).Item(4).ToString
                ComboBox1.SelectedItem = Dt.Rows(x).Item(3).ToString
                TextBox4.Text = Dt.Rows(x).Item(2).ToString
                ID = Dt.Rows(x).Item(0)
                Votes = Dt.Rows(x).Item(5)
            End If
        Next

        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        cmd = New OleDb.OleDbCommand
        cmd.Connection = Con

        If Not Con.State = ConnectionState.Open Then
            Con.Open()
        End If

        If TextBox2.Text = "" Then
            If Not String.IsNullOrEmpty(TextBox1.Text) And Not String.IsNullOrEmpty(TextBox4.Text) And Not String.IsNullOrEmpty(ComboBox1.SelectedIndex.ToString) And Not DateTimePicker1.Text >= Today Then

                cmd.CommandText = "UPDATE CandidateList SET [Cand_Name] = @Name, [Cand_About_Me] = @About, [Cand_House] = @House, [Cand_DOB] = @DOB, [Cand_Votes] = @Votes, [Cand_Criteria] = @Criteria WHERE [ID] = @ID"
                cmd.Parameters.AddWithValue("@Name", TextBox1.Text)
                cmd.Parameters.AddWithValue("@About", TextBox4.Text)
                cmd.Parameters.AddWithValue("@House", ComboBox1.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@DOB", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@Votes", Votes)
                cmd.Parameters.AddWithValue("@Criteria", ListBox1.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@ID", ID)

                cmd.ExecuteNonQuery()

                cmd.Dispose()

                UpdateList()
                ListBox2.SelectedIndex = 0

                MsgBox("Entry Successfully Updated", MsgBoxStyle.Information)


            Else

                MsgBox("Please Check All Field's And Enter Valid Values!", MsgBoxStyle.Critical)

            End If
        Else

            If Not String.IsNullOrEmpty(TextBox1.Text) And Not String.IsNullOrEmpty(TextBox4.Text) And Not String.IsNullOrEmpty(ComboBox1.SelectedIndex.ToString) And Not DateTimePicker1.Text >= Today Then

                Votes = 0

                cmd.CommandText = "INSERT INTO [CandidateList] ([Cand_Name], [Cand_About_Me], [Cand_House], [Cand_DOB], [Cand_Votes], [Cand_Criteria]) VALUES(@Name, @About, @House, @DOB, @Votes, @Criteria)"
                cmd.Parameters.AddWithValue("@Name", TextBox1.Text)
                cmd.Parameters.AddWithValue("@About", TextBox4.Text)
                cmd.Parameters.AddWithValue("@House", ComboBox1.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@DOB", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@Votes", Votes)
                cmd.Parameters.AddWithValue("@Criteria", ListBox1.SelectedItem.ToString)

                cmd.ExecuteNonQuery()

                cmd.Dispose()

                UpdateList()
                ListBox2.SelectedIndex = 0
                MsgBox("Entry Successfully Added", MsgBoxStyle.Information)
                TextBox2.Text = ""

            Else

                MsgBox("Please Check All Field's And Enter Valid Values!", MsgBoxStyle.Critical)

            End If

        End If


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        cmd = New OleDb.OleDbCommand
        cmd.Connection = Con

        If Not String.IsNullOrEmpty(ListBox2.SelectedItem.ToString) Then
            cmd.CommandText = "DELETE FROM CandidateList WHERE ID = @ID"
            cmd.Parameters.AddWithValue("@ID", ID)

            cmd.ExecuteNonQuery()
            cmd.Dispose()

            UpdateList()
            ListBox2.SelectedIndex = 0
            MsgBox("Entry Successfully Deleted", MsgBoxStyle.Information)
            cmd.Dispose()

        End If



    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click

        Me.Enabled = False
        Me.Text = "Processing.. Please Wait"

        If System.IO.Directory.Exists(Application.StartupPath & "\Screens\") Then

            My.Computer.FileSystem.CopyDirectory(Application.StartupPath & "\Screens\", My.Computer.FileSystem.SpecialDirectories.Desktop & "\Screenshots\", True)
            Application.DoEvents()
            MsgBox("Screenshots Successfully Copied To Desktop", MsgBoxStyle.Information)

        Else
            MsgBox("Directory Missing", MsgBoxStyle.Critical)
        End If


        Me.Text = "Details Area"
        Me.Enabled = True

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.Hide()
        Form7.Show()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Me.Hide()
        Form4.Show()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click

        Me.Hide()
        Form1.TextBox1.Text = ""
        Form1.TextBox2.Text = ""
        Form1.ComboBox2.Text = ""
        Form1.Show()

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Me.Hide()
        Form8.Show()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Form11.Text = "Add New Criteria"
        Form11.Show()
        Form11.Label1.Text = "Criteria :"
        Form11.Button1.Text = "Add Criteria"
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Form11.Text = "Add New Candidate"
        Form11.Show()
        Form11.Label1.Text = "Candidate Name :"
        Form11.Button1.Text = "Add Candidate"
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        cmd = New OleDb.OleDbCommand
        cmd.Connection = Con

        If MsgBox("Are You Sure You Want To Delete This Criteria And All Its Members?", MsgBoxStyle.YesNo, "Confirmation") = vbYes Then
           
            If Not String.IsNullOrEmpty(ListBox1.SelectedItem.ToString) Then
                cmd.CommandText = "DELETE FROM CandidateList WHERE Cand_Criteria = @Criteria"
                cmd.Parameters.AddWithValue("@Criteria", ListBox1.SelectedItem.ToString)

                cmd.ExecuteNonQuery()
                cmd.Dispose()

                UpdateList()
                MsgBox("Entries Successfully Deleted", MsgBoxStyle.Information)
                cmd.Dispose()

            End If

        Else
            Return
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim x As Image
        Dim y As New Size(291, 255)
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            PictureBox1.Image = Image.FromFile(OpenFileDialog1.FileName)
            x = New Bitmap(PictureBox1.Image, y)
            PictureBox1.Image = x
        End If
    End Sub
End Class