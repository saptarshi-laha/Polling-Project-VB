Public Class Funcs

    Public Shared threadstate As Integer = 0
    Public Shared thrd As New System.Threading.Thread(AddressOf Funcs.StartPrinting)
    Public Shared ConTchStr1 As String
    Public Shared ConTchStr2 As String
    Public Shared ConStuStr As String

    Public Shared Sub StartPrinting()



        Dim todaydate As String = Today.Date.ToString.Replace("/", ".")
        todaydate = todaydate.Replace(":", ".")


        Do Until Funcs.threadstate = 0


            If Not System.IO.Directory.Exists(Application.StartupPath & "\Screens\" & todaydate & "\") Then

                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Screens\" & todaydate & "\")

            End If


            Try
                Dim ScreenSize As Size = New Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
                Dim BMP As New Bitmap(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
                Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(BMP)
                g.CopyFromScreen(New Point(0, 0), New Point(0, 0), ScreenSize)

                Dim DirectoryA As String = Application.StartupPath & "\Screens\" & todaydate
                Dim Time1 As String = Now.TimeOfDay.ToString.Replace(":", ".")
                Dim img1 As String = ".jpg"

                BMP.Save(DirectoryA & "\" & Time1 & img1)

            Catch ex As Exception
                Continue Do
            End Try



        Loop


    End Sub

End Class
