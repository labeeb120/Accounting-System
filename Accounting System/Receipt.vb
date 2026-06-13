Imports System.Data.OleDb
Imports System.Drawing.Printing
Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports Microsoft.Office.Interop
Public Class Receipt
    Public Sub LoadData()


    End Sub
    Private Sub pd_PrintPage(
sender As Object,
e As PrintPageEventArgs) _
Handles pd.PrintPage

        '========================
        ' الخطوط
        '========================

        Dim titleFont As New System.Drawing.Font(
    "Arial",
    20,
    FontStyle.Bold)

        Dim normalFont As New System.Drawing.Font(
    "Arial",
    12)

        Dim headerFont As New System.Drawing.Font(
    "Arial",
    12,
    FontStyle.Bold)

        '========================
        ' الترويسة
        '========================

        Dim arabicHeader As String = "الجمهورية اليمنية" & vbCrLf &
    "وزارة التعليم الفني والتدريب المهني" & vbCrLf &
    "معهد الخنساء"

        e.Graphics.DrawString(
    arabicHeader,
    headerFont,
    Brushes.Black,
    New RectangleF(500, 20, 250, 100),
    New StringFormat(
    StringFormatFlags.DirectionRightToLeft))

        Dim englishHeader As String = "Republic Of Yemen" & vbCrLf &
    "Ministry Of Technical Education" & vbCrLf &
    "Al-Khansa Institute"

        e.Graphics.DrawString(
    englishHeader,
    headerFont,
    Brushes.Black,
    30,
    20)

        '========================
        ' الشعار
        '========================

        Dim logoPath As String = "C:\Users\ENJAZ\Desktop\مشروع التخرج\IMG-20251103-WA0004.jpg"



        If System.IO.File.Exists(logoPath) Then

            Dim logo As System.Drawing.Image

            logo = System.Drawing.Image.FromFile(logoPath)

            e.Graphics.DrawImage(
        logo,
        330,
        20,
        90,
        90)

        End If

        '========================
        ' خط فاصل
        '========================

        e.Graphics.DrawLine(
    New Pen(Color.Purple, 3),
    20,
    120,
    780,
    120)

        '========================
        ' عنوان السند
        '========================

        e.Graphics.DrawString(
    "سند صرف",
    titleFont,
    Brushes.Purple,
    320,
    150)

        '========================
        ' بيانات السند
        '========================

        Dim y As Integer = 250

        e.Graphics.DrawString(
    "رقم السند : " & TextBox1.Text,
    normalFont,
    Brushes.Black,
    500,
    y)

        y += 50

        e.Graphics.DrawString(
    "اسم المستفيد : " & TextBox2.Text,
    normalFont,
    Brushes.Black,
    500,
    y)

        y += 50

        e.Graphics.DrawString(
    "المبلغ : " & TextBox3.Text,
    normalFont,
    Brushes.Black,
    500,
    y)

        y += 50

        e.Graphics.DrawString(
    "نوع الصرف : " & ComboBox4.Text,
    normalFont,
    Brushes.Black,
    500,
    y)

        y += 50

        e.Graphics.DrawString(
    "التاريخ : " &
    DateTimePicker2.Value.ToShortDateString(),
    normalFont,
    Brushes.Black,
    500,
    y)

        y += 50

        e.Graphics.DrawString(
    "الملاحظات : " & TextBox4.Text,
    normalFont,
    Brushes.Black,
    500,
    y)

        y += 100

        '========================
        ' التوقيع
        '========================

        e.Graphics.DrawString(
    "التوقيع : ____________________",
    normalFont,
    Brushes.Black,
    500,
    y)

    End Sub




    Private Sub Receipt_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox4.Items.Clear()

        ComboBox4.Items.Add("راتب")
        ComboBox4.Items.Add("كهرباء")
        ComboBox4.Items.Add("صيانة اجهزة")
        ComboBox4.Items.Add("صيانة مبنى")
        ComboBox4.Items.Add("إيجار")
        ComboBox4.Items.Add("نثريات")
        ComboBox4.Items.Add("قرطاسية")
        ComboBox4.Items.Add("خدمات اخرى")

    End Sub

    Private Sub حفظToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles حفظToolStripMenuItem.Click
        Try

            OpenConnection()

            Dim sql As String = "INSERT INTO ExpenseVouchers (VoucherDate, VoucherNo, Amount,payment_Method,    ExpenseTypeID, Notes)
   

    VALUES (?, ?, ?, ?, ?, ?, ?)"


            Dim cmd As New OleDbCommand(sql, conn)

            cmd.Parameters.AddWithValue(
    "?",
    DateTimePicker2.Value)

            cmd.Parameters.AddWithValue(
    "?",
    TextBox1.Text)

            cmd.Parameters.AddWithValue(
    "?",
    Val(TextBox3.Text))

            cmd.Parameters.AddWithValue(
    "?",
    ComboBox4.Text)

            cmd.Parameters.AddWithValue(
    "?",
    1)

            cmd.Parameters.AddWithValue(
    "?",
    TextBox2.Text)

            cmd.Parameters.AddWithValue(
    "?",
    TextBox4.Text)

            cmd.ExecuteNonQuery()

            MsgBox("تم حفظ سند الصرف بنجاح")
            Dim frm As New Receipt

            frm.TextBox2.Text = TextBox2.Text
            frm.TextBox3.Text = TextBox4.Text

            frm.ComboBox4.Text = "راتب موظف"

            frm.DateTimePicker2.Value =
DateTimePicker1.Value

            frm.TextBox4.Text = "صرف راتب الموظف"


            frm.ShowDialog()

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub
    Dim WithEvents pd As New PrintDocument
    Public Property DateTimePicker1 As Object

    Private Sub طباعةToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles طباعةToolStripMenuItem.Click
        Dim pp As New PrintPreviewDialog

        pp.Document = pd

        pp.WindowState =
        FormWindowState.Maximized

        pp.ShowDialog()
    End Sub

    Private Sub تعديلToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles تعديلToolStripMenuItem.Click
        Try

            OpenConnection()

            Dim sql As String = "UPDATE ExpenseVouchers SET

    VoucherDate=?,
    Amount=?,
    payment_Method=?,
    PayeeName=?,
    Notes=?

    WHERE VoucherNo=?"

            Dim cmd As New OleDbCommand(sql, conn)

            cmd.Parameters.AddWithValue(
            "?",
            DateTimePicker2.Value)

            cmd.Parameters.AddWithValue(
            "?",
            Val(TextBox3.Text))

            cmd.Parameters.AddWithValue(
            "?",
            ComboBox4.Text)

            cmd.Parameters.AddWithValue(
            "?",
            TextBox2.Text)

            cmd.Parameters.AddWithValue(
            "?",
            TextBox4.Text)

            cmd.Parameters.AddWithValue(
            "?",
            TextBox1.Text)

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

            Dim sql As String = "SELECT * FROM ExpenseVouchers  WHERE VoucherNo=?"


            Dim cmd As New OleDbCommand(sql, conn)

            cmd.Parameters.AddWithValue(
            "?",
            TextBox1.Text)

            Dim dr As OleDbDataReader

            dr = cmd.ExecuteReader()

            If dr.Read() Then

                DateTimePicker2.Value =
                Convert.ToDateTime(
                dr("VoucherDate"))

                TextBox3.Text =
                dr("Amount").ToString()

                ComboBox4.Text =
                dr("payment_Method").ToString()

                TextBox2.Text =
                dr("PayeeName").ToString()

                TextBox4.Text =
                dr("Notes").ToString()

            Else

                MsgBox("السند غير موجود")

            End If

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try
    End Sub

    Private Sub حذفToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles حذفToolStripMenuItem.Click
        Try

            OpenConnection()

            Dim cmd As New OleDbCommand("DELETE FROM ExpenseVouchers
    WHERE VoucherNo=?",
    conn)

            cmd.Parameters.AddWithValue(
    "?",
    TextBox1.Text)

            cmd.ExecuteNonQuery()

            MsgBox("تم حذف السند")

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try
    End Sub

    Private Sub إضافةToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PDFToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PDFToolStripMenuItem.Click

        Try

            Dim save As New SaveFileDialog

            save.Filter =
        "PDF Files|*.pdf"

            If save.ShowDialog =
        DialogResult.OK Then

                Dim doc As New iTextSharp.text.Document(
            iTextSharp.text.PageSize.A4)

                PdfWriter.GetInstance(
            doc,
            New System.IO.FileStream(
            save.FileName,
            System.IO.FileMode.Create))

                doc.Open()

                doc.Add(New Paragraph(
            "سند صرف"))

                doc.Add(New Paragraph(" "))

                doc.Add(New Paragraph(
            "رقم السند : " &
            TextBox1.Text))

                doc.Add(New Paragraph(
            "اسم المستفيد : " &
            TextBox2.Text))

                doc.Add(New Paragraph(
            "المبلغ : " &
            TextBox3.Text))

                doc.Add(New Paragraph(
            "نوع الصرف : " &
            ComboBox4.Text))

                doc.Add(New Paragraph(
            "التاريخ : " &
            DateTimePicker2.Text))

                doc.Add(New Paragraph(
            "الملاحظات : " &
            TextBox4.Text))

                doc.Close()

                MsgBox("تم إنشاء ملف PDF")

            End If

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub EXELToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EXELToolStripMenuItem.Click
        Try

            Dim app As New Excel.Application

            Dim workbook As Excel.Workbook

            Dim worksheet As Excel.Worksheet

            workbook = app.Workbooks.Add()

            worksheet = workbook.Sheets(1)

            '========================
            ' عنوان التقرير
            '========================

            worksheet.Cells(1, 1) =
            "سند صرف"

            worksheet.Range("A1:F1").Merge()

            worksheet.Range("A1").Font.Bold = True

            worksheet.Range("A1").Font.Size = 18

            '========================
            ' البيانات
            '========================

            worksheet.Cells(3, 1) =
            "رقم السند"

            worksheet.Cells(3, 2) =
            TextBox1.Text

            worksheet.Cells(4, 1) =
            "اسم المستفيد"

            worksheet.Cells(4, 2) =
            TextBox2.Text

            worksheet.Cells(5, 1) =
            "المبلغ"

            worksheet.Cells(5, 2) =
            TextBox3.Text

            worksheet.Cells(6, 1) =
            "نوع الصرف"

            worksheet.Cells(6, 2) =
            ComboBox4.Text

            worksheet.Cells(7, 1) =
            "التاريخ"

            worksheet.Cells(7, 2) =
            DateTimePicker2.Value.ToShortDateString()

            worksheet.Cells(8, 1) =
            "الملاحظات"

            worksheet.Cells(8, 2) =
            TextBox4.Text

            '========================
            ' تنسيق
            '========================

            worksheet.Columns.AutoFit()

            worksheet.Range("A3:A8").Font.Bold = True

            worksheet.Range("A1:B8").Borders.LineStyle = 1

            '========================
            ' إظهار الإكسل
            '========================

            app.Visible = True

            MsgBox("تم تصدير السند إلى Excel بنجاح")

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        MINE.Show()
    End Sub

    Private Sub تصديرToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles تصديرToolStripMenuItem.Click

    End Sub

    Private Sub خروجToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles خروجToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

    End Sub
End Class