Public Class Form3

    Dim Con2 As New OleDb.OleDbConnection
    Dim Da1 As New OleDb.OleDbDataAdapter
    Dim Dt As New DataTable
    Dim tries As Integer = 5

    Private Sub Form3_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If funcs.threadstate = 1 Then
            funcs.threadstate = 0
            funcs.thrd.Abort()
        End If
        Application.Exit()
    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label2.Hide()
        RichTextBox1.Hide()
        Label5.Hide()

        If System.IO.File.Exists(Application.StartupPath & "\Data\Tries\Tries2.txt") Then
            RichTextBox1.LoadFile(Application.StartupPath & "\Data\Tries\Tries2.txt")
            tries = RichTextBox1.Text
            RichTextBox1.Text = ""

            If Funcs.threadstate = 0 Then
                Funcs.threadstate = 1
                Funcs.thrd.Start()
                Funcs.thrd.IsBackground = True
            End If

        End If

        Label5.Text = tries
        Label5.Show()

        Con2 = New OleDb.OleDbConnection
        Con2.ConnectionString = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & Application.StartupPath & "\Data\UserList.mdb; Jet OLEDB:Database Password=encrypted"

        If Not Con2.State = ConnectionState.Open Then
            Con2.Open()
        End If

        Da1 = New Data.OleDb.OleDbDataAdapter("SELECT * from UserList", Con2)
        Da1.Fill(Dt)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not TextBox1.Text = "" Then
            If Not TextBox3.Text = "" Then
                If Not TextBox4.Text = "" Then
                    If Not DateTimePicker1.Text < Today Then
                        MsgBox("Date Field Is Invalid, Please Enter A Valid Date", MsgBoxStyle.Critical)
                    Else

                        Dim x As Integer = 0
                        Dim y As Integer = Dt.Rows.Count

                        Dim PW As String
                        Dim NM As String
                        Dim DN As String
                        Dim DOB As String
                        Dim caught As Integer = 0

                        For x = 0 To (y - 1)
                            If Dt.Rows(x).Item(1).ToString = TextBox1.Text Then
                                PW = Dt.Rows(x).Item(2).ToString
                                DN = Dt.Rows(x).Item(3).ToString
                                NM = Dt.Rows(x).Item(4).ToString
                                DOB = Dt.Rows(x).Item(5).ToString
                                caught = 1
                            End If
                        Next

                        If caught = 1 Then

                            If DN = TextBox3.Text And NM = TextBox4.Text And DOB = DateTimePicker1.Text Then
                                Label2.Text = PW
                                Label2.Show()

                                If funcs.threadstate = 1 Then
                                    funcs.threadstate = 0
                                    funcs.thrd.Abort()
                                End If

                                If System.IO.File.Exists(Application.StartupPath & "\Data\Tries\Tries2.txt") Then
                                    System.IO.File.Delete(Application.StartupPath & "\Data\Tries\Tries2.txt")
                                End If

                            End If

                        Else

                            MsgBox("Wrong Initials Entered", MsgBoxStyle.Critical)

                            tries = tries - 1
                            RichTextBox1.Text = tries
                            RichTextBox1.SaveFile(Application.StartupPath & "\Data\Tries\Tries2.txt")
                            RichTextBox1.Text = ""

                            Label5.Text = tries

                            If funcs.threadstate = 0 Then
                                funcs.threadstate = 1
                                funcs.thrd.Start()
                                funcs.thrd.IsBackground = True
                            End If

                            If tries = 0 Then
                                RichTextBox1.Text = "Locked on " + Today.Date.ToString & " at " & Now.TimeOfDay.ToString + " for excessive input of wrong login initials"
                                RichTextBox1.SaveFile(Application.StartupPath & "\Data\Tries\Locked.txt")
                                MsgBox("Application is now Locked. Please contact Administrator to Unlock Application", MsgBoxStyle.Critical)
                                Application.Exit()
                            End If

                        End If


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
            MsgBox("Username Field Is Empty, Please Enter A Valid Username", MsgBoxStyle.Critical)
            Return
        End If

    End Sub
End Class