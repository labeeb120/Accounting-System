Imports System.Data.OleDb
Imports System.Drawing.Printing
Imports System.IO
Imports Microsoft.Office.Interop
Imports iTextSharp.text.pdf
Imports System.Drawing



Public Class Expenses
    Dim WithEvents pd As New PrintDocument
    Dim pp As New PrintPreviewDialog
    Dim WithEvents pdReport As New PrintDocument
    Dim ppReport As New PrintPreviewDialog

    Private Sub Expenses_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FilldataGrideView()
        Dim dal As New ExpenseTypeDAL()

        ' 2. جلب البيانات في قائمة
        Dim typesList As List(Of ExpenseType) = dal.GetRecords()

        ' 3. ربط القائمة بالـ ComboBox
        CbExpenseTypeID.DataSource = typesList
        CbExpenseTypeID.ValueMember = "Id"
        CbExpenseTypeID.DisplayMember = "TypeName"

        CbExpenseTypeID.SelectedIndex = -1
        Timer1.Start()
        If CurrentUserRole = "سكرتارية" Then

            BtnEdt.Enabled = False

            BtnDel.Enabled = False

        End If
        Me.Opacity = 0
        Timer2.Start()

    End Sub
    Private Sub FilldataGrideView()
        Dim dal As New ExpensesDAL()
        Dim g = dal.GetAll()
        DataGridView2.AutoGenerateColumns = False
        DataGridView2.DataSource = g

    End Sub
    Private Sub ClearFields()
        TxAmount.Clear()
        TxExpense_ID.Clear()
        TxNotes.Clear()
        CbExpenseTypeID.SelectedIndex = -1
        DtpExpense_Date.Value = Date.Now

    End Sub
    Function GetLastExpenseID() As Integer

        Dim id As Integer = 0

        Try

            OpenConnection()

            Dim cmd As New OleDbCommand(
        "SELECT MAX(Expense_ID) FROM Expenses", conn)

            Dim result = cmd.ExecuteScalar()

            If Not IsDBNull(result) Then
                id = CInt(result)
            End If




        Catch ex As Exception

        Finally

            CloseConnection()

        End Try

        Return id

    End Function
    Private Sub pd_PrintPage(sender As Object, e As PrintPageEventArgs) Handles pd.PrintPage

        Dim titleFont As New Font("Arial", 20, FontStyle.Bold)
        Dim normalFont As New Font("Arial", 12, FontStyle.Bold)

        '==============================
        ' العنوان العربي
        '==============================

        e.Graphics.DrawString("الجمهورية اليمنية", normalFont, Brushes.Black, 600, 40)
        e.Graphics.DrawString("وزارة التعليم الفني والتدريب المهني", normalFont, Brushes.Black, 500, 70)
        e.Graphics.DrawString("معهد الخنساء", normalFont, Brushes.Black, 620, 100)

        '==============================
        ' العنوان الإنجليزي
        '==============================

        e.Graphics.DrawString("Republic Of Yemen", normalFont, Brushes.Black, 40, 40)
        e.Graphics.DrawString("Ministry Of Technical Education", normalFont, Brushes.Black, 40, 70)
        e.Graphics.DrawString("Al-Khansa Institute", normalFont, Brushes.Black, 40, 100)

        '==============================
        ' الشعار
        '==============================

        Dim logoPath As String =
    "C:\Users\ENJAZ\Desktop\مشروع التخرج\IMG-20251103-WA0004.jpg"

        If IO.File.Exists(logoPath) Then
            Dim logo As System.Drawing.Image =
System.Drawing.Image.FromFile(logoPath)
        End If

        '==============================
        ' خط فاصل
        '==============================

        e.Graphics.DrawLine(Pens.Purple, 40, 170, 760, 170)

        '==============================
        ' عنوان السند
        '==============================

        e.Graphics.DrawString("سند صرف", titleFont, Brushes.Purple, 320, 190)

        '==============================
        ' البيانات
        '==============================

        Dim y As Integer = 270

        e.Graphics.DrawString("رقم السند : " & TxExpense_ID.Text, normalFont, Brushes.Black, 550, y)

        y += 40

        e.Graphics.DrawString("التاريخ : " & DtpExpense_Date.Text, normalFont, Brushes.Black, 550, y)

        y += 40

        e.Graphics.DrawString("نوع المصروف : " & CbExpenseTypeID.Text, normalFont, Brushes.Black, 550, y)

        y += 40

        e.Graphics.DrawString("المبلغ : " & TxAmount.Text & " ريال", normalFont, Brushes.Black, 550, y)

        y += 40

        e.Graphics.DrawString("الملاحظات : " & TxNotes.Text, normalFont, Brushes.Black, 550, y)

        y += 100

        e.Graphics.DrawString("التوقيع", normalFont, Brushes.Black, 100, y)

        e.Graphics.DrawString("الختم", normalFont, Brushes.Black, 650, y)

    End Sub






    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ClearFields()

        BtnSave.Enabled = False
        BtnEdt.Enabled = False
        BtnDel.Enabled = False

    End Sub


    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click
        ClearFields()
        BtnSave.Enabled = True
        TxAmount.Text = "0.00"
        TxAmount.SelectAll()
        TxAmount.Focus()
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click




        Try

            Dim newExp As New Expense()

            Dim Dal As New ExpensesDAL()

            '========================
            ' تعبئة البيانات
            '========================

            newExp.Expense_Date =
        CDate(DtpExpense_Date.Value)

            newExp.ExpenseTypeID =
        CInt(CbExpenseTypeID.SelectedValue)

            newExp.Amount =
        CDec(TxAmount.Text)

            newExp.Notes =
        TxNotes.Text

            '========================
            ' حفظ المصروف
            '========================

            Dim t As Boolean

            t = Dal.Insert(newExp)

            If t = True Then

                '========================
                ' حفظ سند الصرف
                '========================

                SaveExpenseVoucher()

                MessageBox.Show("تم حفظ المصروف بنجاح")

                FilldataGrideView()

                ClearFields()

                BtnSave.Enabled = False

                BtnEdt.Enabled = False

                BtnDel.Enabled = False

                '========================
                ' فتح سند الصرف
                '========================

                Dim frm As New Receipt

                ' المبلغ
                frm.TextBox1.Text =
            TxAmount.Text

                ' نوع المصروف
                frm.TextBox2.Text =
            CbExpenseTypeID.Text

                ' الملاحظات
                frm.TextBox3.Text =
            TxNotes.Text

                ' التاريخ
                frm.DateTimePicker2.Value =
            DtpExpense_Date.Value

                frm.ShowDialog()

            Else

                MessageBox.Show("حصل خطأ أثناء الحفظ")

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub





    Private Sub DataGridView2_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        If e.RowIndex > -1 And e.ColumnIndex = -1 Then
            If DataGridView2("Expense_ID", e.RowIndex).Value <> vbNull Then
                ClearFields()
                TxExpense_ID.Text = DataGridView2("Expense_ID", e.RowIndex).Value
                TxAmount.Text = DataGridView2("Amount", e.RowIndex).Value
                TxNotes.Text = DataGridView2("notes", e.RowIndex).Value
                DtpExpense_Date.Value = DataGridView2("Expense_Date", e.RowIndex).Value
                CbExpenseTypeID.SelectedValue = DataGridView2("ExpenseTypeID", e.RowIndex).Value
                BtnEdt.Enabled = True
                BtnDel.Enabled = True
                BtnSave.Enabled = False
            End If
        End If

    End Sub

    Private Sub BtnEdt_Click(sender As Object, e As EventArgs) Handles BtnEdt.Click
        Dim newExp As New Expense()
        Dim Dal As New ExpensesDAL()

        ' 3. تعبئة الكائن من أدوات التحكم في الفورم
        ' لاحظ التحويلات (CInt, CDec) لضمان توافق الأنواع
        newExp.Expense_ID = CInt(TxExpense_ID.Text.Trim()) ' DateTimePicker

        newExp.Expense_Date = CDate(DtpExpense_Date.Value) ' DateTimePicker
        newExp.ExpenseTypeID = CInt(CbExpenseTypeID.SelectedValue) ' القيمة المخفية (ID) من الكومبو بوكس
        newExp.Amount = CDec(TxAmount.Text) ' تحويل النص إلى رقم عشري
        newExp.Notes = TxNotes.Text ' الملاحظات

        Dim t = Dal.Update(newExp)
        If t = True Then
            MessageBox.Show("تمت عملية التعديل بنجاح")
            ClearFields()
            FilldataGrideView()
            BtnSave.Enabled = False
            BtnEdt.Enabled = False
            BtnDel.Enabled = False
        Else
            MessageBox.Show("حصل خطا اثناء التعديل")

        End If
    End Sub

    Private Sub BtnDel_Click(sender As Object, e As EventArgs) Handles BtnDel.Click
        If String.IsNullOrWhiteSpace(TxExpense_ID.Text.Trim()) <> True Then
            Dim d As DialogResult
            d = MessageBox.Show("هل متاكد من عملية الحذف؟؟", "تحذير", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
            If d = DialogResult.Yes Then
                Dim dal As New ExpensesDAL()
                Dim dl = dal.Delete(CInt(TxExpense_ID.Text))
                If dl Then
                    MessageBox.Show("تمت عملية الحذف بنجاح")
                    ClearFields()
                    FilldataGrideView()
                    BtnSave.Enabled = False
                    BtnEdt.Enabled = False
                    BtnDel.Enabled = False
                Else
                    MessageBox.Show("حصل خطا اثناء الحذف")
                End If
            End If
        End If
    End Sub

    Private Sub TxAmount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxAmount.KeyPress
        ' 1. السماح بالأرقام وأزرار التحكم والنقطة العشرية
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If

        ' 2. منع تكرار العلامة العشرية (استخدمنا الاسم الكامل هنا لحل المشكلة)
        If e.KeyChar = "." Then
            ' قمنا بتحديد النوع بدقة: System.Windows.Forms.TextBox
            Dim tBox As System.Windows.Forms.TextBox = DirectCast(sender, System.Windows.Forms.TextBox)
            If tBox.Text.Contains(".") Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Label18.Text = DateTime.Now.ToString("hh:mm:ss tt")

        Label17.Text = Today.ToString("yyyy/MM/dd")

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        FilldataGrideView()

        MsgBox("تم تحديث البيانات")
    End Sub
    Sub SaveExpenseVoucher()

        Try

            OpenConnection()

            '========================
            ' جلب آخر رقم مصروف
            '========================

            Dim expenseID As Integer = 0

            Dim cmdMax As New OleDbCommand(
            "SELECT MAX(Expense_ID) FROM Expenses", conn)

            Dim result = cmdMax.ExecuteScalar()

            If Not IsDBNull(result) Then
                expenseID = CInt(result)
            End If

            '========================
            ' حفظ سند الصرف
            '========================

            Dim sql As String =
            "INSERT INTO ExpenseVouchers
        (VoucherDate, VoucherNo,
        Amount, payment_Method,
        ExpenseTypeID, PayeeName,
        Expense_ID, Notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)"

            Dim cmd As New OleDbCommand(sql, conn)

            ' التاريخ
            cmd.Parameters.AddWithValue("?", DtpExpense_Date.Value)

            ' رقم السند
            cmd.Parameters.AddWithValue("?", "EXP-" & expenseID)

            ' المبلغ
            cmd.Parameters.AddWithValue("?", Val(TxAmount.Text))

            ' طريقة الدفع
            cmd.Parameters.AddWithValue("?", "نقدي")

            ' نوع المصروف
            cmd.Parameters.AddWithValue("?", CInt(CbExpenseTypeID.SelectedValue))

            ' اسم المستلم
            cmd.Parameters.AddWithValue("?", "مصروفات المعهد")

            ' رقم المصروف
            cmd.Parameters.AddWithValue("?", expenseID)

            ' الملاحظات
            cmd.Parameters.AddWithValue("?", TxNotes.Text)

            cmd.ExecuteNonQuery()

        Catch ex As Exception

            MsgBox("خطأ في حفظ سند الصرف: " & ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        pp.Document = pd
        pp.WindowState = FormWindowState.Maximized
        pp.ShowDialog()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Try

            Dim dt As New DataTable

            OpenConnection()

            Dim sql As String =
            "SELECT Expenses.*, ExpenseType.TypeName " &
            "FROM Expenses " &
            "INNER JOIN ExpenseType " &
            "ON Expenses.ExpenseTypeID = ExpenseType.Id " &
            "WHERE Expense_ID LIKE ? " &
            "OR TypeName LIKE ? " &
            "OR Notes LIKE ? " &
            "OR Amount LIKE ?"

            Dim da As New OleDbDataAdapter(sql, conn)

            da.SelectCommand.Parameters.AddWithValue("?", "%" & TextBox1.Text & "%")

            da.SelectCommand.Parameters.AddWithValue("?", "%" & TextBox1.Text & "%")

            da.SelectCommand.Parameters.AddWithValue("?", "%" & TextBox1.Text & "%")

            da.SelectCommand.Parameters.AddWithValue("?", "%" & TextBox1.Text & "%")

            da.Fill(dt)

            DataGridView2.DataSource = dt

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub





    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click



        Try

            Dim exApp As New Excel.Application

            Dim exBook As Excel.Workbook

            Dim exSheet As Excel.Worksheet

            exBook = exApp.Workbooks.Add()

            exSheet = exBook.ActiveSheet

            '========================
            ' أسماء الأعمدة
            '========================

            For c As Integer = 0 To DataGridView2.Columns.Count - 1

                exSheet.Cells(1, c + 1) =
            DataGridView2.Columns(c).HeaderText

            Next

            '========================
            ' البيانات
            '========================

            For r As Integer = 0 To DataGridView2.Rows.Count - 1

                If DataGridView2.Rows(r).IsNewRow Then Continue For

                For c As Integer = 0 To DataGridView2.Columns.Count - 1

                    exSheet.Cells(r + 2, c + 1) =
                DataGridView2.Rows(r).Cells(c).Value

                Next

            Next

            exApp.Visible = True

            MessageBox.Show("تم تصدير البيانات إلى Excel")

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub


    Private Sub Guna2GradientButton2_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton2.Click
        ppReport.Document = pdReport

        ppReport.WindowState = FormWindowState.Maximized

        ppReport.ShowDialog()

    End Sub
    Private Sub pdReport_PrintPage(sender As Object, e As PrintPageEventArgs) Handles pdReport.PrintPage

        Dim titleFont As New Font("Arial", 18, FontStyle.Bold)

        Dim headerFont As New Font("Arial", 11, FontStyle.Bold)

        Dim normalFont As New Font("Arial", 10)

        Dim y As Integer = 40

        '========================
        ' عنوان التقرير
        '========================

        e.Graphics.DrawString(
        "تقرير المصروفات",
        titleFont,
        Brushes.DarkBlue,
        300,
        y)

        y += 60

        '========================
        ' رؤوس الأعمدة
        '========================

        e.Graphics.DrawRectangle(Pens.Black, 40, y, 100, 30)
        e.Graphics.DrawString("رقم", headerFont, Brushes.Black, 70, y + 5)

        e.Graphics.DrawRectangle(Pens.Black, 140, y, 180, 30)
        e.Graphics.DrawString("نوع المصروف", headerFont, Brushes.Black, 170, y + 5)

        e.Graphics.DrawRectangle(Pens.Black, 320, y, 120, 30)
        e.Graphics.DrawString("المبلغ", headerFont, Brushes.Black, 350, y + 5)

        e.Graphics.DrawRectangle(Pens.Black, 440, y, 140, 30)
        e.Graphics.DrawString("التاريخ", headerFont, Brushes.Black, 470, y + 5)

        e.Graphics.DrawRectangle(Pens.Black, 580, y, 180, 30)
        e.Graphics.DrawString("الملاحظات", headerFont, Brushes.Black, 620, y + 5)

        y += 30

        '========================
        ' بيانات الجدول
        '========================

        For Each row As DataGridViewRow In DataGridView2.Rows

            If row.IsNewRow Then Continue For

            Dim expenseID As String =
            row.Cells("Expense_ID").Value.ToString()

            Dim amount As String =
            row.Cells("Amount").Value.ToString()

            Dim notes As String =
            row.Cells("Notes").Value.ToString()

            Dim expDate As String =
            Convert.ToDateTime(
            row.Cells("Expense_Date").Value).ToShortDateString()

            Dim expType As String =
CbExpenseTypeID.Text

            '========================
            ' رسم الصف
            '========================

            e.Graphics.DrawRectangle(Pens.Black, 40, y, 100, 30)
            e.Graphics.DrawString(expenseID, normalFont, Brushes.Black, 70, y + 5)

            e.Graphics.DrawRectangle(Pens.Black, 140, y, 180, 30)
            e.Graphics.DrawString(expType, normalFont, Brushes.Black, 170, y + 5)

            e.Graphics.DrawRectangle(Pens.Black, 320, y, 120, 30)
            e.Graphics.DrawString(amount, normalFont, Brushes.Black, 350, y + 5)

            e.Graphics.DrawRectangle(Pens.Black, 440, y, 140, 30)
            e.Graphics.DrawString(expDate, normalFont, Brushes.Black, 460, y + 5)

            e.Graphics.DrawRectangle(Pens.Black, 580, y, 180, 30)
            e.Graphics.DrawString(notes, normalFont, Brushes.Black, 590, y + 5)

            y += 30

        Next

        '========================
        ' إجمالي المصروفات
        '========================

        y += 40

        Dim total As Decimal = 0

        For Each row As DataGridViewRow In DataGridView2.Rows

            If row.IsNewRow Then Continue For

            total += Convert.ToDecimal(row.Cells("Amount").Value)

        Next

        e.Graphics.DrawString(
        "إجمالي المصروفات : " & total.ToString("N0") & " ريال",
        headerFont,
        Brushes.DarkRed,
        500,
        y)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        Try

            Dim sfd As New SaveFileDialog()

            sfd.Filter = "PDF Files|*.pdf"
            sfd.FileName = "ExpensesReport.pdf"

            If sfd.ShowDialog() = DialogResult.OK Then

                ' إنشاء الملف
                Dim fs As New System.IO.FileStream(
                sfd.FileName,
                System.IO.FileMode.Create)

                ' إنشاء المستند
                Dim doc As New iTextSharp.text.Document(
                iTextSharp.text.PageSize.A4,
                20, 20, 20, 20)

                ' إنشاء الكاتب
                iTextSharp.text.pdf.PdfWriter.GetInstance(doc, fs)

                doc.Open()

                '========================
                ' العنوان
                '========================

                Dim title As New itextsharp.text.Paragraph(
                "تقرير المصروفات")

                title.Alignment =
                itextsharp.text.Element.ALIGN_CENTER

                doc.Add(title)

                doc.Add(New itextsharp.text.Paragraph(" "))

                '========================
                ' الجدول
                '========================

                Dim table As New itextsharp.text.pdf.PdfPTable(
                DataGridView2.Columns.Count)

                table.WidthPercentage = 100

                ' رؤوس الأعمدة
                For Each col As DataGridViewColumn In DataGridView2.Columns

                    table.AddCell(col.HeaderText)

                Next

                ' البيانات
                For Each row As DataGridViewRow In DataGridView2.Rows

                    If row.IsNewRow Then Continue For

                    For Each cell As DataGridViewCell In row.Cells

                        If cell.Value IsNot Nothing Then

                            table.AddCell(cell.Value.ToString())

                        Else

                            table.AddCell("")

                        End If

                    Next

                Next

                doc.Add(table)

                doc.Close()

                fs.Close()

                MessageBox.Show("تم إنشاء التقرير بنجاح")

                Process.Start(sfd.FileName)

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub صرفToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles صرفToolStripMenuItem.Click
        Receipt.Show()

    End Sub

    Private Sub الاعداداتToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub بحثToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles بحثToolStripMenuItem.Click
        ADD.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        user.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Dashboard.Show()
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
        mane.Show()
        Me.Close()
    End Sub

    Private Sub الاعداداتToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles الاعداداتToolStripMenuItem1.Click
        eemp.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem8.Click
        supports.Show()
        Me.Close()
    End Sub

    Private Sub خروجToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles خروجToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        If Me.Opacity < 1 Then
            Me.Opacity += 0.05
        Else
            Timer2.Stop()
        End If
    End Sub
End Class