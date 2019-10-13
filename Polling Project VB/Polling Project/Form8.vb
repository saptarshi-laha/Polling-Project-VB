Public Class Form8

    Dim Con As New OleDb.OleDbConnection
    Dim Da As New OleDb.OleDbDataAdapter
    Dim Dt As New DataTable


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        Form6.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim x As Integer = Dt.Rows.Count
        Dim y As Integer
        Dim PW As String

        For y = 0 To (x - 1)
            If Dt.Rows(y).Item(1).ToString = Form6.Label6.Text Then
                PW = Dt.Rows(y).Item(2).ToString
            End If
        Next



        If Not TextBox1.Text = "" And Not TextBox2.Text = "" Then

            If System.IO.File.Exists(Application.StartupPath & "\Data\" & TextBox1.Text & "\CandidateList" & "_" & TextBox1.Text & ".mdb") Or System.IO.File.Exists(Application.StartupPath & "\Data\" & Form6.Label6.Text & "\VoterList_Students" & "_" & TextBox1.Text & ".mdb") Or System.IO.File.Exists(Application.StartupPath & "\Data\" & Form6.Label6.Text & "\VoterList_Teachers" & "_" & TextBox1.Text & ".mdb") Then

                MsgBox("Database Already Exists By That Name! Please Change New Database Name To Create A New Database", MsgBoxStyle.Critical)
                Return



            ElseIf PW = TextBox2.Text = False Then

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Tries & Locked & Thread Run''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                MsgBox("Invalid Password Entered", MsgBoxStyle.Critical)
                Return

            Else

                If System.IO.Directory.Exists(Application.StartupPath & "\Data\" & Form6.Label6.Text & "\" & TextBox1.Text & "_ImageData\") = False Then
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Data\" & Form6.Label6.Text & "\" & TextBox1.Text & "_ImageData\")
                End If

                System.IO.File.Copy(Application.StartupPath & "\Data\Templates\CandidateList.mdb", Application.StartupPath & "\Data\" & Form6.Label6.Text & "\CandidateList" & "_" & TextBox1.Text & ".mdb", True)
                System.IO.File.Copy(Application.StartupPath & "\Data\Templates\VoterList_Students.mdb", Application.StartupPath & "\Data\" & Form6.Label6.Text & "\VoterList_Students" & "_" & TextBox1.Text & ".mdb", True)
                System.IO.File.Copy(Application.StartupPath & "\Data\Templates\VoterList_Teachers.mdb", Application.StartupPath & "\Data\" & Form6.Label6.Text & "\VoterList_Teachers" & "_" & TextBox1.Text & ".mdb", True)


                MsgBox("Database Created Successfully!", MsgBoxStyle.Information)
                Me.Hide()
                Form6.Show()

            End If
        Else
            MsgBox("1 Or More Fields Have Been Left Empty! Please Fill In All Fields With Valid Values", MsgBoxStyle.Critical)
            Return
        End If

    End Sub

    Private Sub Form8_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Con.ConnectionString = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & Application.StartupPath & "\Data\UserList.mdb; Jet OLEDB:Database Password=encrypted"

        If Not Con.State = ConnectionState.Open Then
            Con.Open()
        End If

        Da = New Data.OleDb.OleDbDataAdapter("SELECT * from UserList", Con)
        Da.Fill(Dt)

    End Sub
End Class