Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.Office.Interop
Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.Drawing

Public Class mane
    Public Property PrintPreviewDialog1 As Object

    Sub LoadStatistics()
        Try
            OpenConnection()

            ' 1. إجمالي الإيرادات
            Dim cmd1 As New OleDbCommand("SELECT SUM(paid_Amount) FROM payment", conn)
            Dim res1 = cmd1.ExecuteScalar()
            Label8.Text = If(res1 IsNot DBNull.Value AndAlso res1 IsNot Nothing, Convert.ToDouble(res1).ToString("N0"), "0")

            ' 2. عدد الطالبات
            Dim cmd2 As New OleDbCommand("SELECT COUNT(*) FROM Students", conn)
            Label4.Text = cmd2.ExecuteScalar().ToString()

            ' 3. عدد السندات
            Dim cmd3 As New OleDbCommand("SELECT COUNT(*) FROM ReceiptVoucher", conn)
            Label5.Text = cmd3.ExecuteScalar().ToString()

            ' 4. إيراد التسجيل (Label6) - تم توحيد اسم الحقل إلى RevenueTypeID
            Dim cmd4 As New OleDbCommand("SELECT SUM(paid_Amount) FROM payment WHERE RevenueTypeID='إيراد تسجيل'", conn)
            Dim res4 = cmd4.ExecuteScalar()
            Label6.Text = If(res4 IsNot DBNull.Value AndAlso res4 IsNot Nothing, Convert.ToDouble(res4).ToString("N0"), "0")

            ' 5. إيراد التنسيق (Label11)
            Dim cmd5 As New OleDbCommand("SELECT SUM(paid_Amount) FROM payment WHERE RevenueTypeID='إيراد تنسيق'", conn)
            Dim res5 = cmd5.ExecuteScalar()
            Label11.Text = If(res5 IsNot DBNull.Value AndAlso res5 IsNot Nothing, Convert.ToDouble(res5).ToString("N0"), "0")

        Catch ex As Exception
            MsgBox("خطأ في تحميل الإحصائيات: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Try
            Dim titleFont As New System.Drawing.Font("Arial", 18, System.Drawing.FontStyle.Bold)
            Dim headerFont As New System.Drawing.Font("Arial", 11, System.Drawing.FontStyle.Bold)
            Dim bodyFont As New System.Drawing.Font("Arial", 10)
            Dim formatRight As New System.Drawing.StringFormat()
            formatRight.Alignment = System.Drawing.StringAlignment.Far
            formatRight.FormatFlags = System.Drawing.StringFormatFlags.DirectionRightToLeft
            Dim y As Integer = 40

            Dim arabicHeader As String = "الجمهورية اليمنية" & vbCrLf & "وزارة التعليم الفني والتدريب المهني" & vbCrLf & "معهد الخنساء الفني للتدريب"
            e.Graphics.DrawString(arabicHeader, headerFont, System.Drawing.Brushes.Black, New System.Drawing.RectangleF(520, y, 250, 100), formatRight)

            Dim englishHeader As String = "Republic of Yemen" & vbCrLf & "Ministry of Technical Education" & vbCrLf & "Al-Khansa Institute"
            e.Graphics.DrawString(englishHeader, headerFont, System.Drawing.Brushes.Black, 40, y)

            Dim logoPath As String = "C:\Users\ENJAZ\Desktop\مشروع التخرج\IMG-20251103-WA0004.jpg"
            If System.IO.File.Exists(logoPath) Then
                Using logo As System.Drawing.Image = System.Drawing.Image.FromFile(logoPath)
                    e.Graphics.DrawImage(logo, 360, y - 10, 90, 90)
                End Using
            End If

            y += 95
            e.Graphics.DrawLine(New System.Drawing.Pen(System.Drawing.Color.Purple, 3), 30, y, 780, y)
            y += 15
            Dim formatCenter As New System.Drawing.StringFormat() With {.Alignment = System.Drawing.StringAlignment.Center}
            e.Graphics.DrawString("تقرير كشف حركة الإيرادات والمدفوعات العامة", titleFont, System.Drawing.Brushes.Purple, New System.Drawing.RectangleF(30, y, 750, 40), formatCenter)

            y += 45
            e.Graphics.FillRectangle(System.Drawing.Brushes.Lavender, 20, y, 760, 30)
            e.Graphics.DrawRectangle(System.Drawing.Pens.Black, 20, y, 760, 30)

            e.Graphics.DrawString("رقم الدفع", headerFont, System.Drawing.Brushes.Black, 90, y + 5, formatRight)
            e.Graphics.DrawString("رقم الطالبة", headerFont, System.Drawing.Brushes.Black, 200, y + 5, formatRight)
            e.Graphics.DrawString("اسم الطالبة التفصيلي", headerFont, System.Drawing.Brushes.Black, 420, y + 5, formatRight)
            e.Graphics.DrawString("المبلغ المدفوع", headerFont, System.Drawing.Brushes.Black, 540, y + 5, formatRight)
            e.Graphics.DrawString("طريقة الدفع", headerFont, System.Drawing.Brushes.Black, 650, y + 5, formatRight)
            e.Graphics.DrawString("التاريخ", headerFont, System.Drawing.Brushes.Black, 760, y + 5, formatRight)

            y += 30
            Dim totalSum As Double = 0

            For Each row As DataGridViewRow In DataGridView2.Rows
                If row.IsNewRow Then Continue For
                e.Graphics.DrawLine(System.Drawing.Pens.LightGray, 20, y + 25, 780, y + 25)
                e.Graphics.DrawString(row.Cells("payment_ID").Value.ToString(), bodyFont, System.Drawing.Brushes.Black, 90, y + 5, formatRight)
                e.Graphics.DrawString(row.Cells("StudentID").Value.ToString(), bodyFont, System.Drawing.Brushes.Black, 200, y + 5, formatRight)
                e.Graphics.DrawString(row.Cells("StudentName").Value.ToString(), bodyFont, System.Drawing.Brushes.Black, 420, y + 5, formatRight)

                Dim amt As Double = Val(row.Cells("paid_Amount").Value)
                totalSum += amt
                e.Graphics.DrawString(amt.ToString("N0"), bodyFont, System.Drawing.Brushes.DarkGreen, 540, y + 5, formatRight)
                e.Graphics.DrawString(row.Cells("payment_Method").Value.ToString(), bodyFont, System.Drawing.Brushes.Black, 650, y + 5, formatRight)
                e.Graphics.DrawString(Convert.ToDateTime(row.Cells("payment_Date").Value).ToShortDateString(), bodyFont, System.Drawing.Brushes.Black, 760, y + 5, formatRight)

                y += 28
                If y > 1050 Then
                    e.HasMorePages = True
                    Exit Sub
                End If
            Next

            y += 20
            e.Graphics.DrawLine(New System.Drawing.Pen(System.Drawing.Color.Black, 2), 20, y, 780, y)
            y += 10
            e.Graphics.DrawString("إجمالي المبالغ المحصلة في الكشف: " & totalSum.ToString("N0") & " ريال يمني", headerFont, System.Drawing.Brushes.DarkBlue, 780, y, formatRight)
            e.Graphics.DrawString("تاريخ وتوقيت الطباعة الفورية: " & Date.Now.ToString(), bodyFont, System.Drawing.Brushes.Gray, 40, y + 5)
        Catch ex As Exception
            MsgBox("خطأ في معالجة رسم تقرير الطباعة: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub RefreshGrid()
        Try
            OpenConnection()
            Dim sql As String = "SELECT payment_ID, StudentID, StudentName, paid_Amount, payment_Date, payment_Method, RevenueTypeID FROM payment ORDER BY payment_ID DESC"
            Dim da As New OleDbDataAdapter(sql, conn)
            Dim dt As New DataTable
            da.Fill(dt)

            DataGridView2.DataSource = dt

            ' تسمية الأعمدة بالعربية وتطابقها التام مع الـ SELECT
            DataGridView2.Columns("payment_ID").HeaderText = "رقم الدفع"
            DataGridView2.Columns("StudentID").HeaderText = "رقم الطالبة"
            DataGridView2.Columns("StudentName").HeaderText = "اسم الطالبة"
            DataGridView2.Columns("paid_Amount").HeaderText = "المبلغ المدفوع"
            DataGridView2.Columns("payment_Date").HeaderText = "تاريخ الدفع"
            DataGridView2.Columns("payment_Method").HeaderText = "طريقة الدفع"
            DataGridView2.Columns("RevenueTypeID").HeaderText = "مصدر الإيراد"

            ' التنسيقات العامة للجدول
            DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            DataGridView2.MultiSelect = False
            DataGridView2.RightToLeft = RightToLeft.Yes
            DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            DataGridView2.RowTemplate.Height = 35
            DataGridView2.DefaultCellStyle.Font = New System.Drawing.Font("Tahoma", 10)
            DataGridView2.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Tahoma", 10, System.Drawing.FontStyle.Bold)

        Catch ex As Exception
            MsgBox("خطأ في عرض البيانات بالجدول: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub




    Private Sub mane_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' تهيئة وتطوير بنية قاعدة البيانات لدعم الميزات الجديدة
        Try
            OpenConnection()
            ' 1. تحويل رقم الطالبة إلى نص ليدعم الأرقام الخارجية (EXT-...)
            Try
                Dim cmd1 As New OleDbCommand("ALTER TABLE payment ALTER COLUMN StudentID TEXT(50)", conn)
                cmd1.ExecuteNonQuery()
            Catch : End Try

            ' 2. إضافة حقل الهاتف لجدول الدفع (إذا لم يكن موجوداً) لحفظ هاتف الداعم
            Try
                Dim cmd2 As New OleDbCommand("ALTER TABLE payment ADD COLUMN Phone TEXT(20)", conn)
                cmd2.ExecuteNonQuery()
            Catch : End Try

            ' 3. التأكد من حقول جدول سندات القبض
            Try
                Dim cmd3 As New OleDbCommand("ALTER TABLE ReceiptVoucher ADD COLUMN PayerName TEXT(100)", conn)
                cmd3.ExecuteNonQuery()
            Catch : End Try

        Catch ex As Exception
            ' معالجة عامة للأخطاء غير المتوقعة
        Finally
            CloseConnection()
        End Try

        RefreshGrid()
        LoadStatistics() ' تشغيل الإحصائيات فور الفتح لتظهر في لابل 6 ولابل 11
        TextBox9.Focus()
        ' تعبئة قوائم الخيارات يدوياً لضمان الاستقرار
        ComboBox4.Items.Clear()
        ComboBox4.Items.AddRange(New Object() {"إيراد تنسيق", "إيراد تسجيل", "نفقات تشغيلية", "دعم المكتب", "فاعلين خير", "منظمات"})

        ComboBox1.Items.Clear()
        ComboBox1.Items.AddRange(New Object() {"نقدي", "تحويل", "شبكة"})

        ' صلاحيات السكرتارية
        If CurrentUserRole = "سكرتارية" Then
            Button4.Enabled = False ' زر الحذف
            Button1.Enabled = False ' زر التعديل
        End If

        Me.Opacity = 0
        Timer1.Start()
    End Sub

    Function GenerateExternalID() As String
        Dim nextID As Integer = 1
        Try
            OpenConnection()
            ' جلب أقصى رقم دافع مسجل يبدأ بترميز الخارجي
            Dim sql As String = "SELECT MAX(VAL(MID(StudentID, 5))) FROM payment WHERE StudentID LIKE 'EXT-%'"
            Dim cmd As New OleDbCommand(sql, conn)
            Dim result = cmd.ExecuteScalar()
            If Not IsDBNull(result) AndAlso result IsNot Nothing Then
                nextID = Convert.ToInt32(result) + 1
            End If
        Catch ex As Exception
        Finally
            CloseConnection()
        End Try
        Return "EXT-" & nextID.ToString("D4")
    End Function

    Private Sub rbExternal_CheckedChanged(sender As Object, e As EventArgs) Handles rbExternal.CheckedChanged
        If rbExternal.Checked Then
            TextBox9.Enabled = False
            TextBox9.Text = GenerateExternalID()
            TextBox7.Clear()
            TextBox6.Clear()
            TextBox7.Focus()
            ComboBox4.Text = "دعم المكتب"
        End If
    End Sub

    Private Sub rbInternal_CheckedChanged(sender As Object, e As EventArgs) Handles rbInternal.CheckedChanged
        If rbInternal.Checked Then
            TextBox9.Enabled = True
            TextBox9.Clear()
            TextBox7.Clear()
            TextBox6.Clear()
            TextBox9.Focus()
        End If
    End Sub






    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            If rbInternal.Checked AndAlso TextBox9.Text = "" Then
                MsgBox("أدخل رقم الطالبة أولاً", MsgBoxStyle.Exclamation)
                Exit Sub
            ElseIf rbExternal.Checked AndAlso TextBox7.Text = "" Then
                MsgBox("أدخل اسم الداعم أولاً", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            If TextBox4.Text = "" OrElse Not IsNumeric(TextBox4.Text) Then
                MsgBox("يرجى إدخال مبلغ صحيح", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            OpenConnection()

            ' 1. حفظ في جدول الدفع المالي (payment) - تم إضافة حقل Phone
            Dim sql As String = "INSERT INTO payment (StudentID, StudentName, Phone, paid_Amount, payment_Date, payment_Method, RevenueTypeID, Notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd As New OleDbCommand(sql, conn)

            ' في حالة الدفع الخارجي نستخدم المعرف المولد وبيانات الداعم المدخلة
            cmd.Parameters.AddWithValue("?", TextBox9.Text) ' StudentID (أو معرف الداعم)
            cmd.Parameters.AddWithValue("?", TextBox7.Text) ' StudentName (أو اسم الداعم)
            cmd.Parameters.AddWithValue("?", TextBox6.Text) ' Phone (رقم الهاتف)

            cmd.Parameters.AddWithValue("?", Val(TextBox4.Text))
            cmd.Parameters.AddWithValue("?", DateTimePicker1.Value)
            cmd.Parameters.AddWithValue("?", ComboBox1.Text)
            cmd.Parameters.AddWithValue("?", ComboBox4.Text)
            cmd.Parameters.AddWithValue("?", TextBox8.Text)
            cmd.ExecuteNonQuery()

            ' 2. حفظ في جدول سندات القبض (ReceiptVoucher)
            ' نستخدم الحقول الأكثر استقراراً في قاعدة البيانات الحالية لضمان عدم الانهيار
            Dim sqlReceipt As String = "INSERT INTO ReceiptVoucher (ReceiptDate, VoucherNo, Amount, PayerName, Notes) VALUES (?, ?, ?, ?, ?)"
            Dim cmdReceipt As New OleDbCommand(sqlReceipt, conn)
            cmdReceipt.Parameters.AddWithValue("?", DateTimePicker1.Value)
            cmdReceipt.Parameters.AddWithValue("?", "V-" & TextBox9.Text)
            cmdReceipt.Parameters.AddWithValue("?", Val(TextBox4.Text))
            cmdReceipt.Parameters.AddWithValue("?", TextBox7.Text)
            cmdReceipt.Parameters.AddWithValue("?", TextBox8.Text)
            cmdReceipt.ExecuteNonQuery()

            MsgBox("تم حفظ عملية الدفع بنجاح وصدر سند القبض تلقائياً.", MsgBoxStyle.Information, "نجاح الحفظ")

            If Dashboard IsNot Nothing Then Dashboard.LoadStatistics()
            RefreshGrid()
            LoadStatistics()
            Button6_Click(Nothing, Nothing)

        Catch ex As Exception
            MsgBox("خطأ أثناء الحفظ: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click


        TextBox9.Clear()
        TextBox7.Clear()
        TextBox6.Clear()
        TextBox4.Clear()
        TextBox8.Clear()

        ComboBox1.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1

    End Sub



    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        If rbExternal.Checked Then Exit Sub ' تخطي البحث في حالة الدفع الخارجي
        Try
            If String.IsNullOrEmpty(TextBox9.Text) OrElse Not IsNumeric(TextBox9.Text) Then
                TextBox7.Clear()
                TextBox6.Clear()
                TextBox4.Clear()
                Exit Sub
            End If

            OpenConnection()

            Dim sql As String = "SELECT S.StudentName, S.Phone, Sy.FeeAmount " &
                                "FROM Students AS S " &
                                "LEFT JOIN Study AS Sy ON S.StudentID = Sy.StudentID " &
                                "WHERE S.StudentID=?"

            Dim cmd As New OleDbCommand(sql, conn)
            cmd.Parameters.AddWithValue("?", Val(TextBox9.Text))

            Using dr As OleDbDataReader = cmd.ExecuteReader()
                If dr.Read() Then
                    TextBox7.Text = dr("StudentName").ToString()
                    TextBox6.Text = dr("Phone").ToString()

                    If dr("FeeAmount") IsNot DBNull.Value Then
                        TextBox4.Text = Convert.ToDouble(dr("FeeAmount")).ToString("F0")
                    Else
                        TextBox4.Text = ""
                    End If
                Else
                    TextBox7.Clear()
                    TextBox6.Clear()
                    TextBox4.Text = ""
                End If
            End Using

        Catch ex As Exception
            ' صامت لمنع الإزعاج أثناء الكتابة السريعة
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If TextBox9.Text = "" Then
            MsgBox("حدد السجل من الجدول لحذفه", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        If MsgBox("هل أنت متأكد من حذف هذا السجل نهائياً؟", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
        Try
            OpenConnection()
            Dim cmd As New OleDbCommand("DELETE FROM payment WHERE StudentID=?", conn)
            cmd.Parameters.AddWithValue("?", TextBox9.Text)
            cmd.ExecuteNonQuery()

            MsgBox("تم الحذف بنجاح.", MsgBoxStyle.Information)
            If Dashboard IsNot Nothing Then Dashboard.LoadStatistics()
            RefreshGrid()
            LoadStatistics()
            Button6_Click(Nothing, Nothing)
        Catch ex As Exception
            MsgBox("خطأ أثناء الحذف: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try

            OpenConnection()

            Dim sql As String = "UPDATE payment SET
            StudentName=?,
            paid_Amount=?,
            payment_Date=?,
            payment_Method=?,
            payment_Source=?,
            Notes=?

            WHERE StudentID=?"

            Dim cmd As New OleDbCommand(sql, conn)

            cmd.Parameters.AddWithValue("?", TextBox7.Text)

            cmd.Parameters.AddWithValue("?", Val(TextBox4.Text))

            cmd.Parameters.AddWithValue("?", DateTimePicker1.Value)

            cmd.Parameters.AddWithValue("?", ComboBox1.Text)

            cmd.Parameters.AddWithValue("?", ComboBox4.Text)

            cmd.Parameters.AddWithValue("?", TextBox8.Text)

            cmd.Parameters.AddWithValue("?", TextBox9.Text)

            cmd.ExecuteNonQuery()

            MsgBox("تم التعديل")

            RefreshGrid()

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            If TextBox9.Text = "" Then
                MsgBox("حدد السجل المراد تعديله من الجدول أولاً", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            OpenConnection()

            ' تحديث جدول الدفع - تم تصحيح اسم الحقل لـ RevenueTypeID
            Dim sql As String = "UPDATE payment SET StudentName=?, paid_Amount=?, payment_Date=?, payment_Method=?, RevenueTypeID=?, Notes=? WHERE StudentID=?"
            Dim cmd As New OleDbCommand(sql, conn)
            cmd.Parameters.AddWithValue("?", TextBox7.Text)
            cmd.Parameters.AddWithValue("?", Val(TextBox4.Text))
            cmd.Parameters.AddWithValue("?", DateTimePicker1.Value)
            cmd.Parameters.AddWithValue("?", ComboBox1.Text)
            cmd.Parameters.AddWithValue("?", ComboBox4.Text)
            cmd.Parameters.AddWithValue("?", TextBox8.Text)
            cmd.Parameters.AddWithValue("?", TextBox9.Text)
            cmd.ExecuteNonQuery()

            ' تحديث جدول السندات تلقائياً
            Dim cmdReceipt As New OleDbCommand("UPDATE ReceiptVoucher SET Amount=?, ReceiptDate=?, Notes=? WHERE StudentID=?", conn)
            cmdReceipt.Parameters.AddWithValue("?", Val(TextBox4.Text))
            cmdReceipt.Parameters.AddWithValue("?", DateTimePicker1.Value)
            cmdReceipt.Parameters.AddWithValue("?", ComboBox4.Text)
            cmdReceipt.Parameters.AddWithValue("?", TextBox9.Text)
            cmdReceipt.ExecuteNonQuery()

            MsgBox("تم تعديل البيانات بنجاح في سجلات الدفع والسندات.", MsgBoxStyle.Information, "نجاح التعديل")
            If Dashboard IsNot Nothing Then Dashboard.LoadStatistics()
            RefreshGrid()
            LoadStatistics()

        Catch ex As Exception
            MsgBox("خطأ أثناء التعديل: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub


    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView2.Rows(e.RowIndex)
            TextBox9.Text = row.Cells("StudentID").Value.ToString()
            TextBox7.Text = row.Cells("StudentName").Value.ToString()
            TextBox4.Text = row.Cells("paid_Amount").Value.ToString()
            ComboBox1.Text = row.Cells("payment_Method").Value.ToString()
            ComboBox4.Text = row.Cells("RevenueTypeID").Value.ToString()
        End If
    End Sub


    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Try
            OpenConnection()
            ' تم إزالة الحقل الوهمي واستبداله بـ RevenueTypeID في البحث
            Dim sql As String = "SELECT payment_ID, StudentID, StudentName, paid_Amount, payment_Date, payment_Method, RevenueTypeID FROM payment " &
                                "WHERE StudentID LIKE ? OR StudentName LIKE ? OR RevenueTypeID LIKE ? OR payment_Method LIKE ? ORDER BY payment_ID DESC"
            Dim da As New OleDbDataAdapter(sql, conn)

            For i As Integer = 1 To 4
                da.SelectCommand.Parameters.AddWithValue("?", "%" & TextBox5.Text.Trim() & "%")
            Next

            Dim dt As New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
        Catch ex As Exception
        Finally
            CloseConnection()
        End Try
    End Sub



    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.WindowState = FormWindowState.Maximized
        PrintPreviewDialog1.ShowDialog()

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim save As New SaveFileDialog With {.Filter = "PDF Files|*.pdf", .FileName = "تقرير_مدفوعات_" & TextBox7.Text}
        If save.ShowDialog() = DialogResult.OK Then
            Try
                Dim doc As New iTextSharp.text.Document(iTextSharp.text.PageSize.A4)
                iTextSharp.text.pdf.PdfWriter.GetInstance(doc, New IO.FileStream(save.FileName, IO.FileMode.Create))
                doc.Open()
                doc.Add(New iTextSharp.text.Paragraph("Accounting System Report - Payments"))
                doc.Add(New iTextSharp.text.Paragraph("=========================================="))
                doc.Add(New iTextSharp.text.Paragraph("Student ID: " & TextBox9.Text))
                doc.Add(New iTextSharp.text.Paragraph("Student Name: " & TextBox7.Text))
                doc.Add(New iTextSharp.text.Paragraph("Amount Paid: " & TextBox4.Text))
                doc.Add(New iTextSharp.text.Paragraph("Payment Method: " & ComboBox1.Text))
                doc.Add(New iTextSharp.text.Paragraph("RevenueTypeID: " & ComboBox4.Text))
                doc.Add(New iTextSharp.text.Paragraph("Date: " & DateTimePicker1.Value.ToShortDateString()))
                doc.Add(New iTextSharp.text.Paragraph("=========================================="))
                doc.Close()
                MsgBox("تم تصدير مستند الـ PDF بنجاح.", MsgBoxStyle.Information)
            Catch ex As Exception
                MsgBox("خطأ أثناء تصدير الـ PDF: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Try
            Dim app As New Excel.Application
            Dim workbook As Excel.Workbook = app.Workbooks.Add()
            Dim worksheet As Excel.Worksheet = workbook.Sheets(1)
            worksheet.Cells.RightToLeft = True
            For i As Integer = 0 To DataGridView2.Columns.Count - 1
                worksheet.Cells(1, i + 1) = DataGridView2.Columns(i).HeaderText
                worksheet.Cells(1, i + 1).Font.Bold = True
            Next
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                For j As Integer = 0 To DataGridView2.Columns.Count - 1
                    worksheet.Cells(i + 2, j + 1) = DataGridView2.Rows(i).Cells(j).Value
                Next
            Next
            app.Visible = True
            MsgBox("تم تصدير البيانات بالكامل وبنجاح إلى ملف Excel.", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("خطأ أثناء التصدير إلى Excel: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Guna2GradientButton2_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton2.Click

        If TextBox9.Text = "" Then
            MsgBox("يرجى تحديد أو إدخال بيانات عملية الدفع أولاً لطباعة السند الخاص بها", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        PrintPreviewDialog1.Document = PrintDocument2
        PrintPreviewDialog1.WindowState = FormWindowState.Maximized
        PrintPreviewDialog1.ShowDialog()


    End Sub
    Private Sub PrintDocument2_PrintPage(sender As Object, e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument2.PrintPage

        Try
            ' 1. تعريف الخطوط والتنسيقات (نفس خطوط التقرير تماماً)
            Dim titleFont As New System.Drawing.Font("Arial", 18, System.Drawing.FontStyle.Bold)
            Dim headerFont As New System.Drawing.Font("Arial", 11, System.Drawing.FontStyle.Bold)
            Dim itemFont As New System.Drawing.Font("Arial", 11, System.Drawing.FontStyle.Bold)
            Dim valFont As New System.Drawing.Font("Arial", 11, System.Drawing.FontStyle.Regular)

            Dim formatRight As New System.Drawing.StringFormat()
            formatRight.Alignment = System.Drawing.StringAlignment.Far
            formatRight.FormatFlags = System.Drawing.StringFormatFlags.DirectionRightToLeft

            Dim formatLeft As New System.Drawing.StringFormat()
            formatLeft.Alignment = System.Drawing.StringAlignment.Near

            Dim formatCenter As New System.Drawing.StringFormat() With {.Alignment = System.Drawing.StringAlignment.Center}

            Dim y As Integer = 40

            '======================================================
            ' 2. رسم ترويسة المعهد الموحدة (نفس كود التقرير)
            '======================================================
            ' النصوص باللغة العربية (جهة اليمين)
            Dim arabicHeader As String = "الجمهورية اليمنية" & vbCrLf & "وزارة التعليم الفني والتدريب المهني" & vbCrLf & "معهد الخنساء الفني للتدريب"
            e.Graphics.DrawString(arabicHeader, headerFont, System.Drawing.Brushes.Black, New System.Drawing.RectangleF(520, y, 250, 100), formatRight)

            ' النصوص باللغة الإنجليزية (جهة اليسار)
            Dim englishHeader As String = "Republic of Yemen" & vbCrLf & "Ministry of Technical Education" & vbCrLf & "Al-Khansa Institute"
            e.Graphics.DrawString(englishHeader, headerFont, System.Drawing.Brushes.Black, 40, y)

            ' رسم الشعار في المنتصف من مساركِ الخاص
            Dim logoPath As String = "C:\Users\ENJAZ\Desktop\مشروع التخرج\IMG-20251103-WA0004.jpg"
            If System.IO.File.Exists(logoPath) Then
                Using logo As System.Drawing.Image = System.Drawing.Image.FromFile(logoPath)
                    e.Graphics.DrawImage(logo, 360, y - 10, 90, 90)
                End Using
            End If

            ' الخط الفاصل البنفسجي الشهير للتقرير
            y += 95
            e.Graphics.DrawLine(New System.Drawing.Pen(System.Drawing.Color.Purple, 3), 30, y, 780, y)

            '======================================================
            ' 3. رسم هيكل وعنوان السند المالي (يبدأ بعد الترويسة الموحدة)
            '======================================================
            y += 20
            ' إطار داخلي يغلف السند ليعطيه مظهر الوصولات الرسمية
            Dim rectBorder As New System.Drawing.Rectangle(40, y, 740, 420)
            e.Graphics.DrawRectangle(New System.Drawing.Pen(System.Drawing.Color.Purple, 2), rectBorder)

            ' مستطيل عنوان السند الملون بالخلفية اللافندر
            e.Graphics.FillRectangle(System.Drawing.Brushes.Lavender, 280, y + 15, 260, 40)
            e.Graphics.DrawRectangle(System.Drawing.Pens.Purple, 280, y + 15, 260, 40)
            e.Graphics.DrawString("سَنَدْ قَبْضْ مَالِيِّ", titleFont, System.Drawing.Brushes.Purple, New System.Drawing.RectangleF(280, y + 20, 260, 35), formatCenter)

            ' تجهيز إحداثيات طباعة الحقول التفصيلية للسند
            y += 85
            Dim startXLabels As Integer = 740
            Dim startXValues As Integer = 550

            ' المصفوفات لجلب قيم الحقول الخاصة بالسند الحالي من الفورم
            Dim items() As String = {"رقم الطالبة :", "اسم الطالبة :", "المبلغ المدفوع :", "طريقة الدفع :", "مصدر الإيراد :", "تاريخ السند :"}
            Dim values() As String = {
                TextBox9.Text,
                TextBox7.Text,
                Val(TextBox4.Text).ToString("N0") & " ريال يمني",
                If(ComboBox1.Text = "", "نقدي", ComboBox1.Text),
                If(ComboBox4.Text = "", "إيراد عام", ComboBox4.Text),
                DateTimePicker1.Value.ToLongDateString()
            }

            ' دورة رسم بيانات السند والخطوط المنقطة المنسقة
            For i As Integer = 0 To items.Length - 1
                e.Graphics.DrawString(items(i), itemFont, System.Drawing.Brushes.Black, startXLabels, y, formatRight)

                ' تمييز المبلغ بالخط العريض واللون الأخضر الغامق لتسهيل المراجعة المالية
                Dim currentFont As System.Drawing.Font = If(i = 2, New System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold), valFont)
                Dim currentBrush As System.Drawing.Brush = If(i = 2, System.Drawing.Brushes.DarkGreen, System.Drawing.Brushes.MidnightBlue)

                e.Graphics.DrawString(values(i), currentFont, currentBrush, startXValues, y, formatRight)

                Dim penStyle As New System.Drawing.Pen(System.Drawing.Color.LightGray, 1)
                penStyle.DashStyle = Drawing2D.DashStyle.Dot
                e.Graphics.DrawLine(penStyle, 60, y + 22, 740, y + 22)
                y += 38
            Next

            '======================================================
            ' 4. التواقيع والختم في قاع السند
            '======================================================
            y += 35
            e.Graphics.DrawString("توقيع أمين الصندوق: ........................", itemFont, System.Drawing.Brushes.Black, 280, y, formatRight)
            e.Graphics.DrawString("ختم المعهد الرسمي", itemFont, System.Drawing.Brushes.Black, 710, y, formatRight)

        Catch ex As Exception
            MsgBox("خطأ في معالجة طباعة السند الحالية: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Guna2GradientButton1_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton1.Click
        LoadStatistics() : MsgBox("تم تحديث شريط الإحصائيات الفوري بنجاح.", MsgBoxStyle.Information)

    End Sub

    Private Sub ToolStripMenuItem10_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem10.Click
        ADD.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem11_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem11.Click
        user.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem12_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem12.Click
        MINE.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem13_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem13.Click
        studnt.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem14_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem14.Click
        Expenses.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem15_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem15.Click
        eemp.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem16_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem16.Click
        supports.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem20_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem20.Click
        Application.Exit()
    End Sub

    Private Sub MenuStrip2_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip2.ItemClicked

    End Sub

    Private Sub واجهةالتحكمToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles واجهةالتحكمToolStripMenuItem.Click
        Dashboard.Show()
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Me.Opacity < 1 Then
            Me.Opacity += 0.05
        Else
            Timer1.Stop()
        End If
    End Sub

    Private Sub Guna2GroupBox3_Click(sender As Object, e As EventArgs) Handles Guna2GroupBox3.Click

    End Sub

    Private Sub Guna2GradientButton3_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton3.Click

        Try
            OpenConnection()
            Dim sql As String = "SELECT payment_ID, StudentID, StudentName, paid_Amount, payment_Date, payment_Method, RevenueTypeID " &
                                "FROM payment WHERE payment_Date BETWEEN ? AND ? ORDER BY payment_ID DESC"

            Dim da As New OleDbDataAdapter(sql, conn)
            da.SelectCommand.Parameters.AddWithValue("?", DateTimePicker2.Value.Date)
            da.SelectCommand.Parameters.AddWithValue("?", DateTimePicker3.Value.Date)

            Dim dt As New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt

            MsgBox("تم فلترة الجدول لعرض التقرير المخصص بنجاح. يمكنك الطباعة الآن.", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("خطأ في جلب التقرير: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub Guna2GradientButton4_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton4.Click
        '======================================================
        ' زر تصفية البيانات للتقرير السنوي (للسنة الحالية كاملة)
        '======================================================
        Try
            Dim currentYear As Integer = Date.Now.Year ' جلب السنة الحالية تلقائياً
            Dim startDate As New Date(currentYear, 1, 1)   ' 1 يناير
            Dim endDate As New Date(currentYear, 12, 31) ' 31 ديسمبر

            OpenConnection()
            Dim sql As String = "SELECT payment_ID, StudentID, StudentName, paid_Amount, payment_Date, payment_Method, " &
                                "FROM payment WHERE payment_Date BETWEEN ? AND ? ORDER BY payment_ID DESC"

            Dim da As New OleDbDataAdapter(sql, conn)
            da.SelectCommand.Parameters.AddWithValue("?", startDate)
            da.SelectCommand.Parameters.AddWithValue("?", endDate)

            Dim dt As New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt

            MsgBox("تم فلترة الجدول لعرض التقرير السنوي لعام " & currentYear & " بنجاح جاهز للطباعة.", MsgBoxStyle.Information)

        Catch ex As Exception
            MsgBox("خطأ في جلب التقرير السنوي: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub

End Class