Public Class ADD
    Dim DarkMode As Boolean = False
    Private Sub Splash_Load(
    sender As Object,
    e As EventArgs) _
    Handles MyBase.Load

        Guna2ProgressBar1.Value = 0

        Timer1.Start()
        Me.Opacity = 0
        Guna2ProgressBar1.FillColor = Color.WhiteSmoke

        Guna2ProgressBar1.ProgressColor = Color.MediumPurple

        Guna2ProgressBar1.ProgressColor2 = Color.DeepPink




    End Sub

    Private Sub Timer1_Tick(
    sender As Object,
    e As EventArgs) _
    Handles Timer1.Tick

        Guna2ProgressBar1.Value += 1

        If Guna2ProgressBar1.Value >= 100 Then

            Timer1.Stop()

            Guna2ProgressBar1.Value = 100

        End If

        If Me.Opacity < 1 Then

            Me.Opacity += 0.05

        End If
    End Sub


    Private Sub Guna2GradientButton1_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton1.Click
        user.Show()
         Me.Hide()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        Dim DarkMode As Boolean = False



        If DarkMode = False Then

            Me.BackColor = Color.Black
            DarkMode = True

        Else

            Me.BackColor = Color.White
            DarkMode = False

        End If


    End Sub

End Class