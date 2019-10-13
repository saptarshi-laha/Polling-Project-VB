Imports System
Imports System.IO

Public Class Form1

    Dim Con1 As New OleDb.OleDbConnection
    Dim tries1 As Integer = 5
    Dim tries2 As Integer = 5
    Dim Con2 As New OleDb.OleDbConnection



    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If funcs.threadstate = 1 Then
            funcs.threadstate = 0
            funcs.thrd.Abort()
        End If
        Application.Exit()
    End Sub



    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If File.Exists(Application.StartupPath & "\Data\Tries\Locked.txt") Then
            MsgBox("The program has been locked, Please contact Administrator to Unlock Application", MsgBoxStyle.Critical)
            Application.Exit()
        End If

        If File.Exists(Application.StartupPath & "\Data\Tries\Tries1.txt") Then
            RichTextBox1.LoadFile(Application.StartupPath & "\Data\Tries\Tries1.txt")
            tries1 = RichTextBox1.Text
            RichTextBox1.Text = ""

            If funcs.threadstate = 0 Then
                funcs.threadstate = 1
                funcs.thrd.Start()
                funcs.thrd.IsBackground = True
            End If
        End If


        If File.Exists(Application.StartupPath & "\Data\Tries\Tries2.txt") Then

            If funcs.threadstate = 0 Then
                funcs.threadstate = 1
                funcs.thrd.Start()
                funcs.thrd.IsBackground = True
            End If

        End If


        If File.Exists(Application.StartupPath & "\Data\Tries\Tries3.txt") Then

            RichTextBox1.LoadFile(Application.StartupPath & "\Data\Tries\Tries3.txt")
            tries2 = RichTextBox1.Text
            RichTextBox1.Text = ""

            If funcs.threadstate = 0 Then
                funcs.threadstate = 1
                funcs.thrd.Start()
                funcs.thrd.IsBackground = True
            End If

        End If

        If Not Directory.Exists(Application.StartupPath & "\Data\Tries\") Then
            Directory.CreateDirectory(Application.StartupPath & "\Data\Tries\")
        End If

        ' Me.Height = 84

        ComboBox1.Items.Add("Air")
        ComboBox1.Items.Add("Earth")
        ComboBox1.Items.Add("Fire")
        ComboBox1.Items.Add("Water")

        ComboBox3.Items.Add("Student")
        ComboBox3.Items.Add("Teacher")

        ComboBox2.Items.Add("Air")
        ComboBox2.Items.Add("Earth")
        ComboBox2.Items.Add("Fire")
        ComboBox2.Items.Add("Water")
        ComboBox2.Items.Add("No House")

        ComboBox4.Items.Add("Air")
        ComboBox4.Items.Add("Earth")
        ComboBox4.Items.Add("Fire")
        ComboBox4.Items.Add("Water")
        ComboBox4.Items.Add("No House")

        GroupBox1.Hide()
        GroupBox2.Hide()
        GroupBox3.Hide()
        GroupBox4.Hide()
        Label9.Text = tries1
        Label14.Text = tries2
        RichTextBox1.Hide()

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        ' Me.Height = 286

        GroupBox3.Hide()
        GroupBox4.Hide()
        GroupBox2.Hide()
        GroupBox1.Show()


        Con1 = New OleDb.OleDbConnection
        Con1.ConnectionString = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & Application.StartupPath & "\Data\UserList.mdb; Jet OLEDB:Database Password=encrypted"

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        ' Me.Height = 198

        GroupBox1.Hide()
        GroupBox2.Show()
        GroupBox4.Hide()
        GroupBox3.Hide()


    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Me.Hide()
        Form2.Show()
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        
        '\\ My.Computer.Audio.Play(Root Of Program, Sounds, Click.wav)
        '\\ For All Buttons

        If ComboBox3.SelectedItem = "Teacher" Then

            Con2 = New OleDb.OleDbConnection
            Con2.ConnectionString = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & Application.StartupPath & "\Data\UserList.mdb; Jet OLEDB:Database Password=encrypted;"

            If Not Con2.State = ConnectionState.Open Then
                Con2.Open()
            End If

            Dim Da As New OleDb.OleDbDataAdapter
            Dim Dt As New DataTable

            Da = New Data.OleDb.OleDbDataAdapter("SELECT * from UserList", Con2)

            Da.Fill(Dt)

            Dim x As Integer = Dt.Rows.Count

            If x = 0 Then
                MsgBox("No User In Database, Please Register As A New User", MsgBoxStyle.Exclamation)
                Return
            End If

            If String.IsNullOrEmpty(ComboBox4.SelectedItem()) Then
                MsgBox("Please Select A House", MsgBoxStyle.Exclamation)
                Return
            End If

            If TextBox4.Text = "" Then
                MsgBox("Please Enter Username", MsgBoxStyle.Exclamation)
                Return
            ElseIf TextBox5.Text = "" Then
                MsgBox("Please Enter Password", MsgBoxStyle.Exclamation)
                Return
            End If

            Dim y As Integer
            Dim PW As String

            Dim caught As Integer = 0

            For y = 0 To (x - 1)
                If Dt.Rows(y).Item(1).ToString = TextBox4.Text Then
                    caught = 1
                    PW = Dt.Rows(y).Item(2).ToString
                End If
            Next

            If caught = 1 Then
                If PW = TextBox5.Text Then

                    If File.Exists(Application.StartupPath & "\Data\Tries\Tries3.txt") Then
                        File.Delete(Application.StartupPath & "\Data\Tries\Tries3.txt")
                    End If

                    If Funcs.threadstate = 1 Then
                        Funcs.threadstate = 0
                        Funcs.thrd.Abort()
                    End If

                    tries2 = 5

                    Form4.ComboBox1.Items.Clear()

                    Dim files() As String
                    files = Directory.GetFiles(Application.StartupPath & "\Data\" & TextBox4.Text & "\", "VoterList_Teachers_*.mdb", SearchOption.TopDirectoryOnly)
                    Dim total As Integer = 0

                    For Each FileName As String In files
                        Dim iof As Integer = FileName.IndexOf("VoterList_Teachers_")
                        FileName = FileName.Substring(iof)
                        iof = FileName.Length
                        FileName = FileName.Substring(0, (iof - 4))
                        Form4.ComboBox1.Items.Add(FileName)
                        total = total + 1
                    Next

                    Form4.Label4.Text = TextBox4.Text
                    Form4.Label2.Text = ComboBox4.SelectedItem.ToString

                    If Not total > 0 Then
                        MsgBox("No Student Database Entery Present, Please Register As A New User To Create A Default Student Database", MsgBoxStyle.Critical)
                        Return
                    End If

                    Form4.Label7.Text = total.ToString

                    Me.Hide()
                    Form4.Show()
                Else

                    MsgBox("Wrong Details Entered, Please Enter Valid Info", MsgBoxStyle.Critical)
                    tries2 = tries2 - 1
                    RichTextBox1.Text = tries2
                    RichTextBox1.SaveFile(Application.StartupPath & "\Data\Tries\Tries3.txt")
                    RichTextBox1.Text = ""

                    If Funcs.threadstate = 0 Then
                        Funcs.threadstate = 1
                        Funcs.thrd.Start()
                        Funcs.thrd.IsBackground = True
                    End If

                End If
            Else

                MsgBox("Wrong Details Entered, Please Enter Valid Info", MsgBoxStyle.Critical)
                tries2 = tries2 - 1
                RichTextBox1.Text = tries2
                RichTextBox1.SaveFile(Application.StartupPath & "\Data\Tries\Tries3.txt")
                RichTextBox1.Text = ""

                If Funcs.threadstate = 0 Then
                    Funcs.threadstate = 1
                    Funcs.thrd.Start()
                    Funcs.thrd.IsBackground = True
                End If


            End If

            Label14.Text = tries2

            If tries2 = 0 Then
                RichTextBox1.Text = "Locked on " + Today.Date.ToString & " at " & Now.TimeOfDay.ToString + " for excessive input of wrong login initials"
                RichTextBox1.SaveFile(Application.StartupPath & "\Data\Tries\Locked.txt")
                MsgBox("Application is now Locked. Please contact Administrator to Unlock Application", MsgBoxStyle.Critical)
                Application.Exit()
            End If

        ElseIf ComboBox3.SelectedItem = "Student" Then

            If String.IsNullOrEmpty(ComboBox1.SelectedItem()) Then
                MsgBox("Please Select A House", MsgBoxStyle.Exclamation)
                Return
            End If

            If TextBox3.Text = "" Then
                MsgBox("Please Enter A Valid Enrollment Number", MsgBoxStyle.Exclamation)
                Return
            End If

            Form4.Label4.Text = "Student"
            Form4.Label2.Text = ComboBox1.SelectedItem.ToString

            Form4.ComboBox1.Items.Clear()

            Dim files() As String
            files = Directory.GetFiles(Application.StartupPath & "\Data\", "VoterList_Students_*.mdb", SearchOption.AllDirectories)
            Dim total As Integer = 0
            For Each FileName As String In files
                Dim iof As Integer = FileName.IndexOf("VoterList_Students_")
                FileName = FileName.Substring(iof)
                iof = FileName.Length
                FileName = FileName.Substring(0, (iof - 4))
                Form4.ComboBox1.Items.Add(FileName)
                total = total + 1
            Next

            If Not total > 0 Then
                MsgBox("No Student Database Entery Present, Please Register As A New User To Create A Default Student Database", MsgBoxStyle.Critical)
                Return
            End If

            Form4.Label7.Text = total.ToString

            Me.Hide()
            Form4.Show()

        End If


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If Not Con1.State = ConnectionState.Open Then
            Con1.Open()
        End If

        Dim Da As New OleDb.OleDbDataAdapter
        Dim Dt As New DataTable

        Da = New Data.OleDb.OleDbDataAdapter("SELECT * from UserList", Con1)

        Da.Fill(Dt)

        Dim x As Integer = Dt.Rows.Count

        If x = 0 Then
            MsgBox("No User In Database, Please Register As A New User", MsgBoxStyle.Exclamation)
            Return
        End If

        If String.IsNullOrEmpty(ComboBox2.SelectedItem()) Then
            MsgBox("Please Select A House", MsgBoxStyle.Exclamation)
            Return
        End If

        If TextBox1.Text = "" Then
            MsgBox("Please Enter Username", MsgBoxStyle.Exclamation)
            Return
        ElseIf TextBox2.Text = "" Then
            MsgBox("Please Enter Password", MsgBoxStyle.Exclamation)
            Return
        End If


        Dim y As Integer
        Dim PW As String

        Dim caught As Integer = 0

        For y = 0 To (x - 1)
            If Dt.Rows(y).Item(1).ToString = TextBox1.Text Then
                caught = 1
                PW = Dt.Rows(y).Item(2).ToString
            End If
        Next


        If caught = 1 Then

            If TextBox2.Text = PW Then

                tries1 = 5

                If File.Exists(Application.StartupPath & "\Data\Tries\Tries1.txt") Then
                    File.Delete(Application.StartupPath & "\Data\Tries\Tries1.txt")
                End If

                If funcs.threadstate = 1 Then
                    funcs.threadstate = 0
                    funcs.thrd.Abort()
                End If

                Form4.ComboBox1.Items.Clear()

                Dim files() As String
                files = Directory.GetFiles(Application.StartupPath & "\Data\" & TextBox1.Text & "\", "CandidateList_*.mdb", SearchOption.TopDirectoryOnly)
                Dim total As Integer = 0

                For Each FileName As String In files
                    Dim iof As Integer = FileName.IndexOf("CandidateList_")
                    FileName = FileName.Substring(iof)
                    iof = FileName.Length
                    FileName = FileName.Substring(0, (iof - 4))
                    Form4.ComboBox1.Items.Add(FileName)
                    total = total + 1
                Next

                Form4.Label2.Text = ComboBox2.SelectedItem.ToString
                Form4.Label4.Text = TextBox1.Text
                Form4.Label7.Text = total.ToString

                Me.Hide()
                Form4.Show()

            Else
                MsgBox("Wrong Details Entered, Please Enter Valid Info", MsgBoxStyle.Critical)
                tries1 = tries1 - 1
                RichTextBox1.Text = tries1
                RichTextBox1.SaveFile(Application.StartupPath & "\Data\Tries\Tries1.txt")
                RichTextBox1.Text = ""

                If funcs.threadstate = 0 Then
                    funcs.threadstate = 1
                    funcs.thrd.Start()
                    funcs.thrd.IsBackground = True
                End If

            End If

        Else

            MsgBox("Wrong Details Entered, Please Enter Valid Info", MsgBoxStyle.Critical)
            tries1 = tries1 - 1
            RichTextBox1.Text = tries1
            RichTextBox1.SaveFile(Application.StartupPath & "\Data\Tries\Tries1.txt")
            RichTextBox1.Text = ""

            If funcs.threadstate = 0 Then
                funcs.threadstate = 1
                funcs.thrd.Start()
                funcs.thrd.IsBackground = True
            End If

        End If

        Label9.Text = tries1

        If tries1 = 0 Then
            RichTextBox1.Text = "Locked on " + Today.Date.ToString & " at " & Now.TimeOfDay.ToString + " for excessive input of wrong login initials"
            RichTextBox1.SaveFile(Application.StartupPath & "\Data\Tries\Locked.txt")
            MsgBox("Application is now Locked. Please contact Administrator to Unlock Application", MsgBoxStyle.Critical)
            Application.Exit()
        End If

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged

        If ComboBox3.SelectedItem = "Teacher" Then
            GroupBox3.Show()
            GroupBox4.Hide()
        ElseIf ComboBox3.SelectedItem = "Student" Then
            GroupBox4.Show()
            GroupBox3.Hide()
            GroupBox4.Visible = True
        End If

    End Sub
End Class
