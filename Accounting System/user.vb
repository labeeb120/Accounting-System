
Imports System.Data.OleDb
Imports System.Windows.Forms

Public Class user
    Private LoginAttempts As Integer = 0
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox3.Enabled = False
        Me.Opacity = 0
        Timer1.Start()
        TextBox1.Focus()
    End Sub

    Private Sub Guna2GradientButton2_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton2.Click
        CurrentUserID = 0
        CurrentUserName = ""
        CurrentUserRole = ""
        Me.Close()
    End Sub

    Private Sub Guna2GradientButton1_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton1.Click
        If TextBox1.Text.Trim() = "" Or TextBox2.Text.Trim() = "" Then
            MsgBox("الرجاء إدخال اسم المستخدم وكلمة المرور أولاً", MsgBoxStyle.Exclamation, "تنبيه")
            Exit Sub
        End If

        Try
            OpenConnection()

            ' الخطوة 1: فحص هل اسم المستخدم موجود أصلاً في النظام؟
            Dim checkUserCmd As New OleDbCommand("SELECT * FROM Users WHERE Username=?", conn)
            checkUserCmd.Parameters.AddWithValue("?", TextBox1.Text.Trim())

            Dim dr As OleDbDataReader = checkUserCmd.ExecuteReader()

            If dr.Read() Then
                ' اسم المستخدم صحيح وموجود، الآن ننتقل للخطوة 2: فحص كلمة المرور المرافقة له
                ' قمنا بتغليف [Password] بأقواس مربعة لتجنب الكلمات المحجوزة
                Dim checkPassCmd As New OleDbCommand("SELECT * FROM Users WHERE Username=? AND [Password]=?", conn)
                checkPassCmd.Parameters.AddWithValue("?", TextBox1.Text.Trim())
                checkPassCmd.Parameters.AddWithValue("?", TextBox2.Text)

                Dim drPass As OleDbDataReader = checkPassCmd.ExecuteReader()

                If drPass.Read() Then
                    ' كلمة المرور صحيحة، نجاح تسجيل الدخول
                    CurrentUserID = Convert.ToInt32(drPass("UserID"))
                    CurrentUserName = drPass("Username").ToString()
                    CurrentUserRole = drPass("ROLE").ToString()

                    If Application.OpenForms().OfType(Of MINE).Any() Then
                        Application.OpenForms("MINE").BringToFront()
                    Else
                        MINE.Show()
                    End If
                    Me.Hide()
                Else
                    ' اسم المستخدم صح لكن كلمة المرور هي الغلط
                    MsgBox("عذراً، كلمة المرور التي أدخلتها غير صحيحة بضبط! يرجى إعادة المحاولة.", MsgBoxStyle.Critical, "خطأ في كلمة المرور")
                    TextBox2.Clear()
                    TextBox2.Focus()
                End If
                drPass.Close()
            Else
                ' اسم المستخدم غير موجود نهائياً في قاعدة البيانات
                MsgBox("عذراً، اسم المستخدم هذا غير مسجل بالنظام بتاتاً!", MsgBoxStyle.Critical, "خطأ في اسم المستخدم")
                TextBox1.SelectAll()
                TextBox1.Focus()
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox("حدث خطأ أثناء الاتصال: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Try
            OpenConnection()
            Dim cmd As New OleDbCommand("SELECT HintQuestion FROM Users WHERE Username=?", conn)
            cmd.Parameters.AddWithValue("@p1", TextBox1.Text)

            Dim Question As Object = cmd.ExecuteScalar()

            If Question IsNot Nothing Then
                MsgBox(Question.ToString(), MsgBoxStyle.Information, "سؤال التحقق")
            Else
                MsgBox("اسم المستخدم غير موجود")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click



        If TextBox3.Text.Trim() = "" Then
            MsgBox("الرجاء كتابة كلمة المرور الجديدة أولاً", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Try
            OpenConnection()

            ' تم التعديل بوضع [Password] لتفادي الكلمات المحجوزة بالأكسس
            Dim cmd As New OleDbCommand("UPDATE Users SET [Password]=? WHERE Username=?", conn)
            cmd.Parameters.AddWithValue("@p1", TextBox3.Text)
            cmd.Parameters.AddWithValue("@p2", TextBox1.Text)

            cmd.ExecuteNonQuery()

            MsgBox("تم تغيير كلمة المرور بنجاح في معهد الخنساء ✅", MsgBoxStyle.Information, "تحديث البيانات")

            TextBox3.Clear()
            TextBox4.Clear()
            TextBox3.Enabled = False

        Catch ex As Exception
            MsgBox("فشل التحديث: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub Guna2GradientButton3_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton3.Click


        ' التحقق من أن الحقول ليست فارغة
        If TextBox1.Text.Trim() = "" Or TextBox4.Text.Trim() = "" Then
            MsgBox("الرجاء إدخال اسم المستخدم وإجابة السؤال أولاً", MsgBoxStyle.Exclamation, "تنبيه")
            Exit Sub
        End If

        Try
            OpenConnection()

            ' قمنا باستخدام دالة UCase أو مقارنة النصوص المباشرة مع تنظيف المدخلات من المسافات
            Dim cmd As New OleDbCommand("SELECT UserID FROM Users WHERE Username=? AND HintAnswer=?", conn)

            ' نمرر اسم المستخدم والإجابة بعد تنظيفهما تماماً من أي مسافات مخفية في البداية والنهاية
            cmd.Parameters.AddWithValue("@p1", TextBox1.Text.Trim())
            cmd.Parameters.AddWithValue("@p2", TextBox4.Text.Trim())

            Dim Result = cmd.ExecuteScalar()

            If Result IsNot Nothing Then
                TextBox3.Enabled = True
                MsgBox("تم التحقق بنجاح في معهد الخنساء، يمكنك الآن كتابة كلمة المرور الجديدة ✅", MsgBoxStyle.Information, "نجاح التحقق")
                TextBox3.Focus()
            Else
                ' الرسالة تظهر هنا لو لم تطابق الكلمة بالكامل مخزون قاعدة البيانات
                MsgBox("الإجابة غير مطابقة للمسجل لدينا! تأكدي من طريقة كتابة الكلمة (بالهمزة أو بدونها).", MsgBoxStyle.Critical, "فشل التحقق")
                TextBox4.SelectAll()
                TextBox4.Focus()
            End If

        Catch ex As Exception
            MsgBox("حدث خطأ أثناء معالجة البيانات: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Me.Opacity < 1 Then
            Me.Opacity += 0.05
        Else
            Timer1.Stop()
        End If
    End Sub

    Private Sub TextBox1_KeyDown(
sender As Object,
e As KeyEventArgs) _
Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then

            TextBox2.Focus()

        End If
    End Sub

    Private Sub TextBox2_KeyDown(
sender As Object,
e As KeyEventArgs) _
Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then

            Guna2GradientButton1.PerformClick()

        End If
    End Sub

    Private Sub Guna2GradientButton4_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton4.Click
        TextBox1.Clear()

        TextBox2.Clear()

        TextBox3.Clear()

        TextBox4.Clear()

        TextBox1.Focus()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
