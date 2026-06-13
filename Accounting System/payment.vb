Imports System.Data.OleDb
Imports System.Drawing.Printing
Public Class payment
    Dim WithEvents PD As New PrintDocument
    Dim PPD As New PrintPreviewDialog
    Sub GenerateReceiptNo()

        Try

            OpenConnection()

            Dim cmd As New OleDbCommand(
        "SELECT MAX(ReceiptID) FROM ReceiptVoucher", conn)

            Dim x = cmd.ExecuteScalar()

            If IsDBNull(x) Then

                TextBox1.Text = "1"

            Else

                TextBox1.Text = Val(x) + 1

            End If

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub


    Private Sub ComboBox4_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

        ComboBox4.Items.Clear()

        If ComboBox4.Text = "إيراد" Then
            ComboBox4.Items.AddRange(New Object() {"إيراد تنسيق", "إيراد تسجيل"})
        ElseIf ComboBox4.Text = "دعم" Then
            ComboBox4.Items.AddRange(New Object() {"دعم المكتب", "فاعلين خير", "منظمات"})
        ElseIf ComboBox4.Text = "مصروف" Then
            ComboBox4.Items.Add("نفقات تشغيلية")
        End If


    End Sub

    Private Sub حفظToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles حفظToolStripMenuItem.Click

        Try

            OpenConnection()

            Dim sql As String =
        "INSERT INTO ReceiptVoucher " &
        "(ReceiptDate, VoucherNo, Amount, PaymentMethod, PayerName, Notes) " &
        "VALUES (?, ?, ?, ?, ?, ?)"

            Dim cmd As New OleDbCommand(sql, conn)

            cmd.Parameters.AddWithValue("?", Date.Now)

            cmd.Parameters.AddWithValue("?", TextBox1.Text)

            cmd.Parameters.AddWithValue("?", Val(TextBox3.Text))

            cmd.Parameters.AddWithValue("?", ComboBox4.Text)

            cmd.Parameters.AddWithValue("?", TextBox2.Text)

            cmd.Parameters.AddWithValue("?", TextBox4.Text)

            cmd.ExecuteNonQuery()

            MsgBox("تم حفظ السند بنجاح")

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub payment_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker2.Value = DateTime.Now
        ComboBox4.Items.Clear()

        ComboBox4.Items.Add("إيراد تنسيق")
        ComboBox4.Items.Add("إيراد تسجيل")
        ComboBox4.Items.Add("نفقات تشغيلية")
        ComboBox4.Items.Add("دعم المكتب")
        ComboBox4.Items.Add("فاعلين خير")
        ComboBox4.Items.Add("منظمات")



        GenerateReceiptNo()


    End Sub

    Private Sub إضافةToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles إضافةToolStripMenuItem.Click
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()

        ComboBox4.SelectedIndex = -1

        GenerateReceiptNo()
    End Sub

    Private Sub حذفToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles حذفToolStripMenuItem.Click
        Try

            OpenConnection()

            Dim cmd As New OleDbCommand(
            "DELETE FROM ReceiptVoucher WHERE VoucherNo=?",
            conn)

            cmd.Parameters.AddWithValue("?", TextBox1.Text)

            cmd.ExecuteNonQuery()

            MsgBox("تم حذف السند")

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try
    End Sub

    Private Sub تعديلToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles تعديلToolStripMenuItem.Click

        Try

            OpenConnection()

            Dim sql As String =
        "UPDATE ReceiptVoucher SET " &
        "Amount=?, PaymentMethod=?, PayerName=?, Notes=? " &
        "WHERE VoucherNo=?"

            Dim cmd As New OleDbCommand(sql, conn)

            cmd.Parameters.AddWithValue("?", Val(TextBox3.Text))

            cmd.Parameters.AddWithValue("?", ComboBox4.Text)

            cmd.Parameters.AddWithValue("?", TextBox2.Text)

            cmd.Parameters.AddWithValue("?", TextBox4.Text)

            cmd.Parameters.AddWithValue("?", TextBox1.Text)

            cmd.ExecuteNonQuery()

            MsgBox("تم تعديل السند")

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub بحثToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles بحثToolStripMenuItem.Click
        Try

            OpenConnection()

            Dim cmd As New OleDbCommand(
            "SELECT * FROM ReceiptVoucher WHERE VoucherNo=?",
            conn)

            cmd.Parameters.AddWithValue("?", TextBox1.Text)

            Dim dr As OleDbDataReader

            dr = cmd.ExecuteReader()

            If dr.Read Then

                TextBox1.Text = dr("VoucherNo").ToString()

                TextBox2.Text = dr("PayerName").ToString()

                TextBox3.Text = dr("Amount").ToString()

                ComboBox4.Text = dr("PaymentMethod").ToString()

                TextBox4.Text = dr("Notes").ToString()

            Else

                MsgBox("السند غير موجود")

            End If

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub طباعةToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles طباعةToolStripMenuItem.Click

        PPD.Document = PD
        PPD.ShowDialog()
    End Sub
    Private Sub PD_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PD.PrintPage

        Dim f1 As New Font("Arial", 18, FontStyle.Bold)
        Dim f2 As New Font("Arial", 14, FontStyle.Bold)
        Dim f3 As New Font("Arial", 12)

        e.Graphics.DrawString("سند قبض", f1, Brushes.Purple, 320, 60)

        e.Graphics.DrawLine(New Pen(Color.Purple, 3),
                            250,
                            100,
                            520,
                            100)

        e.Graphics.DrawString("رقم السند : " & TextBox1.Text,
                              f2,
                              Brushes.Black,
                              100,
                              150)

        e.Graphics.DrawString("الاسم : " & TextBox2.Text,
                              f2,
                              Brushes.Black,
                              100,
                              200)

        e.Graphics.DrawString("المبلغ : " & TextBox3.Text,
                              f2,
                              Brushes.Black,
                              100,
                              250)

        e.Graphics.DrawString("نوع الإيراد : " & ComboBox4.Text,
                              f2,
                              Brushes.Black,
                              100,
                              300)

        e.Graphics.DrawString("البيان : " & TextBox4.Text,
                              f2,
                              Brushes.Black,
                              100,
                              350)

        e.Graphics.DrawString("التاريخ : " & Date.Now.ToShortDateString(),
                              f2,
                              Brushes.Black,
                              100,
                              400)

    End Sub

    Private Sub خروجToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles خروجToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub
End Class