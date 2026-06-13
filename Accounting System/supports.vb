Imports System.Data.OleDb
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Imports System.Drawing.Printing
Imports Microsoft.Office.Interop

Public Class supports

    Dim WithEvents pdSupport As New PrintDocument

    Dim WithEvents pdStatistics As New PrintDocument


    Private Sub pdSupport_PrintPage(
sender As Object,
e As Printing.PrintPageEventArgs) _
Handles pdSupport.PrintPage

        Dim HeaderFont As New System.Drawing.Font(
    "Arial", 12, FontStyle.Bold)

        Dim TitleFont As New System.Drawing.Font(
    "Arial", 18, FontStyle.Bold)

        Dim NormalFont As New System.Drawing.Font(
    "Arial", 10)

        e.Graphics.DrawString(
    "الجمهورية اليمنية" & vbCrLf &
    "وزارة التعليم الفني والتدريب المهني" & vbCrLf &
    "معهد الخنساء",
    HeaderFont,
    Brushes.Black,
    520,
    20)

        e.Graphics.DrawString(
    "Republic Of Yemen" & vbCrLf &
    "Ministry Of Technical Education" & vbCrLf &
    "Al-Khansa Institute",
    HeaderFont,
    Brushes.Black,
    20,
    20)

        e.Graphics.DrawLine(
    New Pen(Color.Purple, 3),
    20,
    120,
    780,
    120)

        e.Graphics.DrawString(
    "تقرير الدعم العام",
    TitleFont,
    Brushes.Purple,
    280,
    150)

        Dim y As Integer = 230

        e.Graphics.DrawString(
    "رقم الدعم",
    HeaderFont,
    Brushes.Black,
    680,
    y)

        e.Graphics.DrawString(
    "اسم الداعم",
    HeaderFont,
    Brushes.Black,
    530,
    y)

        e.Graphics.DrawString(
    "المصدر",
    HeaderFont,
    Brushes.Black,
    400,
    y)

        e.Graphics.DrawString(
    "المبلغ",
    HeaderFont,
    Brushes.Black,
    260,
    y)

        e.Graphics.DrawString(
    "طريقة الدفع",
    HeaderFont,
    Brushes.Black,
    100,
    y)

        y += 40

        For Each row As DataGridViewRow In DataGridView2.Rows

            If Not row.IsNewRow Then

                e.Graphics.DrawString(
            row.Cells(0).Value.ToString(),
            NormalFont,
            Brushes.Black,
            680,
            y)

                e.Graphics.DrawString(
            row.Cells(1).Value.ToString(),
            NormalFont,
            Brushes.Black,
            530,
            y)

                e.Graphics.DrawString(
            row.Cells(5).Value.ToString(),
            NormalFont,
            Brushes.Black,
            400,
            y)

                e.Graphics.DrawString(
            row.Cells(6).Value.ToString(),
            NormalFont,
            Brushes.Black,
            260,
            y)

                e.Graphics.DrawString(
            row.Cells(7).Value.ToString(),
            NormalFont,
            Brushes.Black,
            100,
            y)

                y += 30

            End If

        Next

        e.Graphics.DrawString(
    "تاريخ الطباعة : " &
    Date.Now.ToShortDateString(),
    NormalFont,
    Brushes.Black,
    550,
    1000)

        e.Graphics.DrawString(
    "التوقيع : __________________",
    NormalFont,
    Brushes.Black,
    500,
    1030)

    End Sub


    Private Sub FillDataGridView()

        Dim dt As New DataTable

        OpenConnection()

        Dim da As New OleDbDataAdapter(
    "SELECT * FROM Supporters",
    conn)

        da.Fill(dt)

        DataGridView2.DataSource = dt

        CloseConnection()

        ' أسماء الأعمدة بالعربي
        DataGridView2.Columns("supporter_id").HeaderText = "رقم الدعم"

        DataGridView2.Columns("supporter_name").HeaderText = "اسم الداعم"

        DataGridView2.Columns("phone").HeaderText = "الهاتف"

        DataGridView2.Columns("email").HeaderText = "البريد الإلكتروني"

        DataGridView2.Columns("address").HeaderText = "العنوان"

        DataGridView2.Columns("support_source").HeaderText = "مصدر الدعم"

        DataGridView2.Columns("amount").HeaderText = "المبلغ"

        DataGridView2.Columns("payment_method").HeaderText = "طريقة الدفع"

        DataGridView2.Columns("description").HeaderText = "الوصف"

        DataGridView2.Columns("support_date").HeaderText = "تاريخ الدعم"
        DataGridView2.AutoSizeColumnsMode =
DataGridViewAutoSizeColumnsMode.Fill

        DataGridView2.ColumnHeadersDefaultCellStyle.Font =
New System.Drawing.Font("Tahoma", 10, FontStyle.Bold)


        DataGridView2.RowTemplate.Height = 30

        DataGridView2.SelectionMode =
DataGridViewSelectionMode.FullRowSelect

        DataGridView2.ReadOnly = True

        DataGridView2.AllowUserToAddRows = False

    End Sub
    Private Sub LoadSupportData()

        Dim dt As New DataTable

        OpenConnection()

        Dim da As New OleDbDataAdapter("SELECT

    
    Support.SupporterID,
    Support.SupporterName,
    Support.phone,
    Support.email,
    Support.Address,
    SupportPayments.SupportDate,
    SupportPayments.Amount,
    SupportPayments.PaymentMethod

    FROM Support

    INNER JOIN SupportPayments

    ON Support.SupporterID =
    SupportPayments.SupporterID", conn)



        da.Fill(dt)

        DataGridView2.DataSource = dt

        CloseConnection()

    End Sub

    Private Sub supports_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadSupportData()
        LoadStatistics()
        TextBox6.Text = "ابحث هنا..."
        TextBox6.ForeColor = Color.Gray

        ComboBox4.Items.Clear()

        ComboBox4.Items.Add("فاعلين خير")
        ComboBox4.Items.Add("منظمات")
        ComboBox4.Items.Add("دعم المكتب")
        ComboBox4.Items.Add("إيراد تسجيل")
        ComboBox4.Items.Add("إيراد تنسيق")

        ComboBox1.Items.Clear()

        ComboBox1.Items.Add("نقداً")
        ComboBox1.Items.Add("تحويل بنكي")
        ComboBox1.Items.Add("شيك")

        If CurrentUserRole = "سكرتارية" Then

            Button4.Enabled = False

            Button1.Enabled = False

        End If
        Me.Opacity = 0
        Timer1.Start()
    End Sub
    Private Sub pdStatistics_PrintPage(
sender As Object,
e As Printing.PrintPageEventArgs) _
Handles pdStatistics.PrintPage

        Dim HeaderFont As New System.Drawing.Font(
    "Arial", 12, FontStyle.Bold)

        Dim TitleFont As New System.Drawing.Font(
    "Arial", 18, FontStyle.Bold)

        Dim NormalFont As New System.Drawing.Font(
    "Arial", 12)

        e.Graphics.DrawString(
    "الجمهورية اليمنية" & vbCrLf &
    "وزارة التعليم الفني والتدريب المهني" & vbCrLf &
    "معهد الخنساء",
    HeaderFont,
    Brushes.Black,
    520,
    20)

        e.Graphics.DrawString(
    "Republic Of Yemen" & vbCrLf &
    "Ministry Of Technical Education" & vbCrLf &
    "Al-Khansa Institute",
    HeaderFont,
    Brushes.Black,
    20,
    20)

        e.Graphics.DrawLine(
    New Pen(Color.Purple, 3),
    20,
    120,
    780,
    120)

        e.Graphics.DrawString(
    "تقرير إحصائيات الدعم",
    TitleFont,
    Brushes.Purple,
    250,
    150)

        Dim y As Integer = 250

        e.Graphics.DrawString(
    "إجمالي عدد الداعمين : " &
    Label24.Text,
    HeaderFont,
    Brushes.Black,
    150,
    y)

        y += 60

        e.Graphics.DrawString(
    "إجمالي الدعم : " &
    Label18.Text,
    HeaderFont,
    Brushes.Black,
    150,
    y)

        y += 60

        e.Graphics.DrawString(
    "عدد داعمي الطلاب : " &
    Label10.Text,
    HeaderFont,
    Brushes.Black,
    150,
    y)

        y += 60

        e.Graphics.DrawString(
    "عدد داعمي الموظفين : " &
    Label11.Text,
    HeaderFont,
    Brushes.Black,
    150,
    y)

        y += 60

        e.Graphics.DrawString(
    "عدد داعمي المعهد : " &
    Label16.Text,
    HeaderFont,
    Brushes.Black,
    150,
    y)

        e.Graphics.DrawString(
    "تاريخ الطباعة : " &
    Date.Now.ToShortDateString(),
    NormalFont,
    Brushes.Black,
    550,
    1000)

        e.Graphics.DrawString(
    "التوقيع : __________________",
    NormalFont,
    Brushes.Black,
    500,
    1030)

    End Sub
    Private Sub LoadStatistics()

        Try

            OpenConnection()

            Dim cmd As New OleDbCommand(
        "SELECT COUNT(*) FROM Support",
        conn)

            Label24.Text =
        cmd.ExecuteScalar().ToString()

            cmd = New OleDbCommand(
        "SELECT SUM(Amount) FROM SupportPayments",
        conn)

            If IsDBNull(cmd.ExecuteScalar()) Then
                Label18.Text = "0"
            Else
                Label18.Text =
            cmd.ExecuteScalar().ToString()
            End If

            cmd = New OleDbCommand(
        "SELECT COUNT(*) FROM SupportPayments WHERE BeneficiaryTypeID=1",
        conn)

            Label10.Text =
        cmd.ExecuteScalar().ToString()

            cmd = New OleDbCommand(
        "SELECT COUNT(*) FROM SupportPayments WHERE BeneficiaryTypeID=2",
        conn)

            Label11.Text =
        cmd.ExecuteScalar().ToString()

            cmd = New OleDbCommand(
        "SELECT COUNT(*) FROM SupportPayments WHERE BeneficiaryTypeID=3",
        conn)

            Label16.Text =
        cmd.ExecuteScalar().ToString()

        Catch ex As Exception

        Finally

            CloseConnection()

        End Try

    End Sub


    Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem8.Click
        eemp.Show()
    End Sub

    Private Sub صرفToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles صرفToolStripMenuItem.Click
        payment.ShowDialog()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Try

            OpenConnection()

            Dim cmd As New OleDbCommand("UPDATE Support SET SupporterName=?,email=?, phone=?, Address=?    WHERE SupporterID=?", conn)


            cmd.Parameters.AddWithValue(
    "@p1",
    TextBox2.Text)

            cmd.Parameters.AddWithValue(
    "@p2",
    TextBox5.Text)

            cmd.Parameters.AddWithValue(
    "@p3",
    TextBox3.Text)

            cmd.Parameters.AddWithValue(
    "@p4",
    TextBox7.Text)

            cmd.Parameters.AddWithValue(
    "@p5",
    Val(TextBox1.Text))

            cmd.ExecuteNonQuery()

            MsgBox("تم التعديل")
            Dashboard.LoadStatistics()
            FillDataGridView()

            LoadStatistics()
            LoadSupportData()

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Try

            OpenConnection()

            Dim cmd As New OleDbCommand("DELETE FROM Support WHERE SupporterID=?", conn)






            cmd.Parameters.AddWithValue(
    "@p1",
    Val(TextBox1.Text))

            cmd.ExecuteNonQuery()

            MsgBox("تم الحذف")
            LoadSupportData()
            Dashboard.LoadStatistics()

            FillDataGridView()

            LoadStatistics()

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        Try

            OpenConnection()

            '========================
            ' حفظ الداعم
            '========================

            Dim cmd As New OleDbCommand(
        "INSERT INTO Support
        (
        SupporterName,
        SupporterTypeID,
        email,
        phone,
        Address
        )
        VALUES
        (?,?,?,?,?)", conn)

            cmd.Parameters.AddWithValue("@p1", TextBox2.Text)
            cmd.Parameters.AddWithValue("@p2", ComboBox4.SelectedIndex + 1)
            cmd.Parameters.AddWithValue("@p3", TextBox5.Text)
            cmd.Parameters.AddWithValue("@p4", TextBox3.Text)
            cmd.Parameters.AddWithValue("@p5", TextBox7.Text)

            cmd.ExecuteNonQuery()

            '========================
            ' آخر رقم داعم
            '========================

            cmd = New OleDbCommand(
        "SELECT MAX(SupporterID) FROM Support",
        conn)

            Dim SupporterID As Integer =
        cmd.ExecuteScalar()

            '========================
            ' تحديد الجهة المستفيدة
            '========================

            Dim BeneficiaryID As Integer = 0

            If RadioButton1.Checked Then
                BeneficiaryID = 1     ' الطلاب
            End If

            If RadioButton3.Checked Then
                BeneficiaryID = 2     ' الموظفين
            End If

            If RadioButton2.Checked Then
                BeneficiaryID = 3     ' المعهد
            End If

            '========================
            ' حفظ مبلغ الدعم
            '========================

            Dim cmdPay As New OleDbCommand(
        "INSERT INTO SupportPayments
        (
        PaymentMethod,
        SupporterID,
        Amount,
        SupportDate,
        BeneficiaryTypeID
        )
        VALUES
        (?,?,?,?,?)", conn)

            cmdPay.Parameters.AddWithValue("@p1",
        ComboBox1.Text)

            cmdPay.Parameters.AddWithValue("@p2",
        SupporterID)

            cmdPay.Parameters.AddWithValue("@p3",
        Val(TextBox4.Text))

            cmdPay.Parameters.AddWithValue("@p4",
        DateTimePicker1.Value)

            cmdPay.Parameters.AddWithValue("@p5",
        BeneficiaryID)

            cmdPay.ExecuteNonQuery()

            '========================
            ' إنشاء سند قبض تلقائي
            '========================

            Dim cmdReceipt As New OleDbCommand(
        "INSERT INTO ReceiptVoucher
        (
        ReceiptDate,
        VoucherNo,
        Amount,
        PaymentMethod,
        PayerName,
        Notes
        )
        VALUES
        (?,?,?,?,?,?)", conn)

            cmdReceipt.Parameters.AddWithValue("@p1",
        DateTimePicker1.Value)

            cmdReceipt.Parameters.AddWithValue("@p2",
        "SUP" & SupporterID)

            cmdReceipt.Parameters.AddWithValue("@p3",
        Val(TextBox4.Text))

            cmdReceipt.Parameters.AddWithValue("@p4",
        ComboBox1.Text)

            cmdReceipt.Parameters.AddWithValue("@p5",
        TextBox2.Text)

            cmdReceipt.Parameters.AddWithValue("@p6",
        TextBox8.Text)

            cmdReceipt.ExecuteNonQuery()

            '========================
            ' فتح سند القبض
            '========================

            Dim frm As New payment

            frm.TextBox1.Text =
        "SUP" & SupporterID

            frm.TextBox2.Text =
        TextBox2.Text

            frm.TextBox3.Text =
        TextBox4.Text

            frm.ComboBox4.Text =
        ComboBox4.Text

            frm.DateTimePicker2.Value =
        DateTimePicker1.Value

            frm.TextBox4.Text =
        TextBox8.Text

            frm.ShowDialog()

            '========================
            ' تحديث الشاشة
            '========================

            FillDataGridView()

            Dashboard.LoadStatistics()

            MsgBox("تم حفظ الدعم بنجاح")

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FillDataGridView()

        LoadStatistics()

        MsgBox("تم تحديث البيانات")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox7.Clear()
        TextBox8.Clear()

        ComboBox1.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1

        DateTimePicker1.Value = Date.Today

        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If e.RowIndex >= 0 Then

            TextBox1.Text =
            DataGridView2.CurrentRow.Cells("supporter_id").Value.ToString()

            TextBox2.Text =
            DataGridView2.CurrentRow.Cells("supporter_name").Value.ToString()

            TextBox3.Text =
            DataGridView2.CurrentRow.Cells("phone").Value.ToString()

            TextBox5.Text =
            DataGridView2.CurrentRow.Cells("email").Value.ToString()

            TextBox7.Text =
            DataGridView2.CurrentRow.Cells("address").Value.ToString()

            ComboBox4.Text =
            DataGridView2.CurrentRow.Cells("support_source").Value.ToString()

            TextBox4.Text =
            DataGridView2.CurrentRow.Cells("amount").Value.ToString()

            ComboBox1.Text =
            DataGridView2.CurrentRow.Cells("payment_method").Value.ToString()

            TextBox8.Text =
            DataGridView2.CurrentRow.Cells("description").Value.ToString()

        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        TextBox6.Text = TextBox6.Text.Trim
        Try

            If TextBox6.Text.Trim = "" Then

                FillDataGridView()

                Exit Sub

            End If

            Dim dt As New DataTable

            OpenConnection()

            Dim da As New OleDbDataAdapter(
    "SELECT * FROM Support
    WHERE SupporterName LIKE ?
    OR SupporterID LIKE ?
    OR Phone LIKE ?
    OR Email LIKE ?
    OR Address LIKE ?",
    conn)

            da.SelectCommand.Parameters.AddWithValue(
    "@p1",
    "%" & TextBox6.Text & "%")

            da.SelectCommand.Parameters.AddWithValue(
    "@p2",
    "%" & TextBox6.Text & "%")

            da.SelectCommand.Parameters.AddWithValue(
    "@p3",
    "%" & TextBox6.Text & "%")

            da.SelectCommand.Parameters.AddWithValue(
    "@p4",
    "%" & TextBox6.Text & "%")

            da.SelectCommand.Parameters.AddWithValue(
    "@p5",
    "%" & TextBox6.Text & "%")

            da.Fill(dt)

            DataGridView2.DataSource = dt

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try


    End Sub
    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Dashboard.Show()
        LoadStatistics()

        MsgBox("تم تحديث الإحصائيات")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim pp As New PrintPreviewDialog

        pp.Document = pdSupport

        pp.ShowDialog()


    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim pp As New PrintPreviewDialog

        pp.Document = pdStatistics

        pp.ShowDialog()


    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Try

            Dim app As New Excel.Application

            Dim wb As Excel.Workbook
            Dim ws As Excel.Worksheet

            wb = app.Workbooks.Add()
            ws = wb.Sheets(1)

            ws.Cells(1, 1) = "تقرير الدعم"

            ws.Range("A1:J1").Merge()

            ws.Range("A1").Font.Bold = True
            ws.Range("A1").Font.Size = 18

            For c As Integer = 0 To DataGridView2.Columns.Count - 1

                ws.Cells(3, c + 1) =
            DataGridView2.Columns(c).HeaderText

            Next

            For r As Integer = 0 To DataGridView2.Rows.Count - 1

                If Not DataGridView2.Rows(r).IsNewRow Then

                    For c As Integer = 0 To DataGridView2.Columns.Count - 1

                        ws.Cells(r + 4, c + 1) =
                    DataGridView2.Rows(r).Cells(c).Value

                    Next

                End If

            Next

            ws.Columns.AutoFit()

            app.Visible = True

            MsgBox("تم تصدير التقرير إلى Excel")

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click


        Try

            Dim save As New SaveFileDialog

            save.Filter = "PDF Files|*.pdf"

            If save.ShowDialog() <> DialogResult.OK Then Exit Sub

            Dim doc As New iTextSharp.text.Document(
        iTextSharp.text.PageSize.A4,
        20,
        20,
        20,
        20)

            PdfWriter.GetInstance(
        doc,
        New IO.FileStream(
        save.FileName,
        IO.FileMode.Create))

            doc.Open()

            '=====================
            ' الترويسة
            '=====================

            Dim TitleFont As New iTextSharp.text.Font(
iTextSharp.text.Font.BOLD,
16)

            Dim NormalFont As New iTextSharp.text.Font(
iTextSharp.text.Font.NORMAL,
11)

            doc.Add(New Paragraph(
        "Republic Of Yemen"))

            doc.Add(New Paragraph(
        "Ministry Of Technical Education"))

            doc.Add(New Paragraph(
        "Al-Khansa Institute"))

            doc.Add(New Paragraph(
        "----------------------------------------"))

            doc.Add(New Paragraph(
        "تقرير الدعم العام",
        TitleFont))

            doc.Add(New Paragraph(" "))

            '=====================
            ' البيانات
            '=====================

            For Each row As DataGridViewRow In DataGridView2.Rows

                If Not row.IsNewRow Then

                    doc.Add(New Paragraph(
                "رقم الدعم : " &
                row.Cells(0).Value.ToString(),
                NormalFont))

                    doc.Add(New Paragraph(
                "اسم الداعم : " &
                row.Cells(1).Value.ToString(),
                NormalFont))

                    doc.Add(New Paragraph(
                "مصدر الدعم : " &
                row.Cells(5).Value.ToString(),
                NormalFont))

                    doc.Add(New Paragraph(
                "المبلغ : " &
                row.Cells(6).Value.ToString(),
                NormalFont))

                    doc.Add(New Paragraph(
                "طريقة الدفع : " &
                row.Cells(7).Value.ToString(),
                NormalFont))

                    doc.Add(New Paragraph(
                "----------------------------------------"))

                End If

            Next

            doc.Add(New Paragraph(" "))

            doc.Add(New Paragraph(
        "عدد الداعمين : " &
        Label24.Text))

            doc.Add(New Paragraph(
        "إجمالي الدعم : " &
        Label18.Text))

            doc.Add(New Paragraph(" "))

            doc.Add(New Paragraph(
        "تاريخ الطباعة : " &
        Date.Now.ToShortDateString()))

            doc.Close()

            MsgBox("تم إنشاء ملف PDF بنجاح")

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub بحثToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles بحثToolStripMenuItem.Click
        ADD.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        user.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        MINE.Show()
        Me.Close()
    End Sub

    Private Sub حفظToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles حفظToolStripMenuItem.Click
        studnt.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem7.Click
        Expenses.Show()
        Me.Close()
    End Sub

    Private Sub الاعداداتToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles الاعداداتToolStripMenuItem1.Click
        payment.Show()
        Me.Close()
    End Sub

    Private Sub خروجToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles خروجToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub TextBox6_Enter(sender As Object, e As EventArgs) Handles TextBox6.Enter
        If TextBox6.Text = "ابحث هنا..." Then

            TextBox6.Text = ""
            TextBox6.ForeColor = Color.Black

        End If

    End Sub

    Private Sub TextBox6_Leave(sender As Object, e As EventArgs) Handles TextBox6.Leave
        If TextBox6.Text = "" Then

            TextBox6.Text = "ابحث هنا..."
            TextBox6.ForeColor = Color.Gray

        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Me.Opacity < 1 Then
            Me.Opacity += 0.05
        Else
            Timer1.Stop()
        End If
    End Sub
End Class