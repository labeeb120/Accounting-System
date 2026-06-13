Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.Drawing.Printing
Imports System.Drawing

Public  Class eemp
    Dim WithEvents pdEmployees As New PrintDocument
    Dim WithEvents pdEmployeesSummary As New PrintDocument
    Dim WithEvents pdEmployeesGeneral As New PrintDocument
    Dim WithEvents pdMonthlySalary As New PrintDocument
    Dim WithEvents pdMonthlySalaary As New PrintDocument
    Dim dtMonthly As New DataTable
    Private save As Object


    Private Sub CalculateNet()

        TextBox4.Text =
    Val(TextBox7.Text) +
    Val(TextBox11.Text)
    End Sub
    Private Sub ClearData()

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()

        TextBox4.Text = "0"
        TextBox5.Text = "0"
        TextBox7.Text = "0"
        TextBox8.Clear()
        TextBox10.Text = "0"
        TextBox11.Text = "0"

        ComboBox1.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        ComboBox5.SelectedIndex = -1

    End Sub
    Private Sub LoadMonthlyReport()

        Try

            dtMonthly.Clear()

            OpenConnection()

            Dim sql As String = "SELECT Employee.employee_name,

       
                Employee.job_title,
                EmployeePayments.BasicSalary,
                EmployeePayments.bonuses,
                EmployeePayments.NetSalary

        FROM Employee

        INNER JOIN EmployeePayments

        ON Employee.employee_id =
        EmployeePayments.employee_id

        WHERE EmployeePayments.Month=?
        AND EmployeePayments.Year=?"

            Dim da As New OleDbDataAdapter(sql, conn)

            da.SelectCommand.Parameters.AddWithValue(
        "@p1",
        ComboBox1.Text)

            da.SelectCommand.Parameters.AddWithValue(
        "@p2",
        TextBox8.Text)

            da.Fill(dtMonthly)

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub
    Private Sub pdMonthlySaalary_PrintPage(
sender As Object,
e As PrintPageEventArgs) _
Handles pdMonthlySalary.PrintPage

        Dim HeaderFont As New System.Drawing.Font(
    "Arial", 12, FontStyle.Bold)

        Dim TitleFont As New System.Drawing.Font(
    "Arial", 18, FontStyle.Bold)

        Dim NormalFont As New System.Drawing.Font(
    "Arial", 10)

        e.Graphics.DrawString(
    "الجمهورية اليمنية",
    HeaderFont,
    Brushes.Black,
    560,
    20)

        e.Graphics.DrawString(
    "وزارة التعليم الفني والتدريب المهني",
    HeaderFont,
    Brushes.Black,
    460,
    45)

        e.Graphics.DrawString(
    "معهد الخنساء",
    HeaderFont,
    Brushes.Black,
    600,
    70)

        e.Graphics.DrawString(
    "Republic Of Yemen",
    HeaderFont,
    Brushes.Black,
    20,
    20)

        e.Graphics.DrawString(
    "Ministry Of Technical Education",
    HeaderFont,
    Brushes.Black,
    20,
    45)

        e.Graphics.DrawString(
    "Al-Khansa Institute",
    HeaderFont,
    Brushes.Black,
    20,
    70)

        e.Graphics.DrawLine(
    New Pen(Color.Purple, 2),
    20,
    110,
    780,
    110)

        e.Graphics.DrawString(
    "تقرير الرواتب الشهرية",
    TitleFont,
    Brushes.Purple,
    260,
    140)

        e.Graphics.DrawString(
    "الشهر : " & ComboBox1.Text,
    HeaderFont,
    Brushes.Black,
    550,
    180)

        e.Graphics.DrawString(
    "السنة : " & TextBox8.Text,
    HeaderFont,
    Brushes.Black,
    350,
    180)

        Dim y As Integer = 240

        e.Graphics.DrawString(
    "الموظف",
    HeaderFont,
    Brushes.Black,
    120,
    y)

        e.Graphics.DrawString(
    "المسمى",
    HeaderFont,
    Brushes.Black,
    320,
    y)

        e.Graphics.DrawString(
    "الراتب",
    HeaderFont,
    Brushes.Black,
    470,
    y)

        e.Graphics.DrawString(
    "المكافأة",
    HeaderFont,
    Brushes.Black,
    580,
    y)

        e.Graphics.DrawString(
    "الصافي",
    HeaderFont,
    Brushes.Black,
    690,
    y)

        y += 40

        Dim TotalNet As Decimal = 0

        For Each row As DataRow In dtMonthly.Rows

            e.Graphics.DrawString(
        row("employee_name").ToString,
        NormalFont,
        Brushes.Black,
        120,
        y)

            e.Graphics.DrawString(
        row("job_title").ToString,
        NormalFont,
        Brushes.Black,
        320,
        y)

            e.Graphics.DrawString(
        row("BasicSalary").ToString,
        NormalFont,
        Brushes.Black,
        470,
        y)

            e.Graphics.DrawString(
        row("bonuses").ToString,
        NormalFont,
        Brushes.Black,
        580,
        y)

            e.Graphics.DrawString(
        row("NetSalary").ToString,
        NormalFont,
        Brushes.Black,
        690,
        y)

            TotalNet +=
        Val(row("NetSalary"))

            y += 25

        Next

        y += 30

        e.Graphics.DrawString(
    "عدد الموظفين : " &
    dtMonthly.Rows.Count,
    HeaderFont,
    Brushes.Black,
    50,
    y)

        e.Graphics.DrawString(
    "إجمالي الرواتب : " &
    TotalNet.ToString("N0"),
    HeaderFont,
    Brushes.Black,
    500,
    y)

    End Sub
    Private Sub pdMonthlySalary_PrintPage(
sender As Object,
e As PrintPageEventArgs) _
Handles pdMonthlySalary.PrintPage

        Dim HeaderFont As New System.Drawing.Font("Arial", 12, FontStyle.Bold)
        Dim TitleFont As New System.Drawing.Font("Arial", 18, FontStyle.Bold)
        Dim NormalFont As New System.Drawing.Font("Arial", 10)

        '=========================
        ' الترويسة
        '=========================

        e.Graphics.DrawString(
    "الجمهورية اليمنية",
    HeaderFont,
    Brushes.Black,
    560,
    20)

        e.Graphics.DrawString(
    "وزارة التعليم الفني والتدريب المهني",
    HeaderFont,
    Brushes.Black,
    460,
    45)

        e.Graphics.DrawString(
    "معهد الخنساء",
    HeaderFont,
    Brushes.Black,
    600,
    70)

        e.Graphics.DrawString(
    "Republic Of Yemen",
    HeaderFont,
    Brushes.Black,
    20,
    20)

        e.Graphics.DrawString(
    "Ministry Of Technical Education",
    HeaderFont,
    Brushes.Black,
    20,
    45)

        e.Graphics.DrawString(
    "Al-Khansa Institute",
    HeaderFont,
    Brushes.Black,
    20,
    70)

        e.Graphics.DrawLine(
    New Pen(Color.Purple, 2),
    20,
    110,
    780,
    110)

        '=========================
        ' عنوان التقرير
        '=========================

        e.Graphics.DrawString(
    "تقرير الرواتب الشهرية",
    TitleFont,
    Brushes.Purple,
    250,
    140)

        e.Graphics.DrawString(
    "الشهر : " & ComboBox1.Text,
    HeaderFont,
    Brushes.Black,
    550,
    180)

        e.Graphics.DrawString(
    "السنة : " & TextBox8.Text,
    HeaderFont,
    Brushes.Black,
    350,
    180)

        '=========================
        ' رؤوس الأعمدة
        '=========================

        Dim y As Integer = 240

        e.Graphics.DrawString("رقم", HeaderFont, Brushes.Black, 40, y)
        e.Graphics.DrawString("اسم الموظف", HeaderFont, Brushes.Black, 130, y)
        e.Graphics.DrawString("المسمى", HeaderFont, Brushes.Black, 320, y)
        e.Graphics.DrawString("الراتب", HeaderFont, Brushes.Black, 470, y)
        e.Graphics.DrawString("المكافآت", HeaderFont, Brushes.Black, 580, y)
        e.Graphics.DrawString("الصافي", HeaderFont, Brushes.Black, 700, y)

        y += 30

        e.Graphics.DrawLine(
    Pens.Black,
    20,
    y,
    780,
    y)

        y += 20

        Dim TotalSalary As Decimal = 0

        '=========================
        ' البيانات
        '=========================

        For Each row As DataGridViewRow In DataGridView2.Rows

            If Not row.IsNewRow Then

                e.Graphics.DrawString(
            row.Cells("employee_id").Value.ToString(),
            NormalFont,
            Brushes.Black,
            40,
            y)

                e.Graphics.DrawString(
            row.Cells("employee_name").Value.ToString(),
            NormalFont,
            Brushes.Black,
            130,
            y)

                e.Graphics.DrawString(
            row.Cells("job_title").Value.ToString(),
            NormalFont,
            Brushes.Black,
            320,
            y)

                e.Graphics.DrawString(
            row.Cells("paid_amount").Value.ToString(),
            NormalFont,
            Brushes.Black,
            470,
            y)

                e.Graphics.DrawString(
            row.Cells("bonus").Value.ToString(),
            NormalFont,
            Brushes.Black,
            580,
            y)

                e.Graphics.DrawString(
            row.Cells("net_salary").Value.ToString(),
            NormalFont,
            Brushes.Black,
            700,
            y)

                TotalSalary +=
            Val(row.Cells("net_salary").Value)

                y += 25

            End If

        Next

        '=========================
        ' الإجمالي
        '=========================

        y += 30

        e.Graphics.DrawLine(
    Pens.Black,
    20,
    y,
    780,
    y)

        y += 25

        e.Graphics.DrawString(
    "إجمالي الرواتب : " &
    TotalSalary.ToString("N0"),
    HeaderFont,
    Brushes.Black,
    500,
    y)

        y += 40

        e.Graphics.DrawString(
    "عدد الموظفين : " &
    (DataGridView2.Rows.Count - 1).ToString(),
    HeaderFont,
    Brushes.Black,
    40,
    y)

        e.Graphics.DrawString(
    "تاريخ الطباعة : " &
    Date.Now.ToShortDateString(),
    HeaderFont,
    Brushes.Black,
    500,
    y)

        y += 60

        e.Graphics.DrawString(
    "التوقيع : ________________________",
    HeaderFont,
    Brushes.Black,
    500,
    y)

    End Sub
    Private Sub pdEmployeesSummary_PrintPage(
sender As Object,
e As PrintPageEventArgs) _
Handles pdEmployeesSummary.PrintPage

        ' كود التقرير المختصر








        Dim HeaderFont As New Drawing.Font("Arial", 12, FontStyle.Bold)
        Dim TitleFont As New Drawing.Font("Arial", 18, FontStyle.Bold)
        Dim NormalFont As New Drawing.Font("Arial", 10)

        ' الترويسة اليمنية
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

        ' عنوان التقرير
        e.Graphics.DrawString(
    "تقرير الموظفين",
    TitleFont,
    Brushes.Purple,
    280,
    120)

        ' خط فاصل
        e.Graphics.DrawLine(
    Pens.Purple,
    20,
    160,
    780,
    160)

        Dim y As Integer = 200

        ' رؤوس الأعمدة
        e.Graphics.DrawString("رقم", HeaderFont, Brushes.Black, 700, y)
        e.Graphics.DrawString("الاسم", HeaderFont, Brushes.Black, 560, y)
        e.Graphics.DrawString("المسمى", HeaderFont, Brushes.Black, 400, y)
        e.Graphics.DrawString("التعاقد", HeaderFont, Brushes.Black, 250, y)
        e.Graphics.DrawString("الصافي", HeaderFont, Brushes.Black, 100, y)

        y += 40

        For Each row As DataGridViewRow In DataGridView2.Rows

            If Not row.IsNewRow Then

                e.Graphics.DrawString(
            row.Cells("employee_id").Value.ToString,
            NormalFont,
            Brushes.Black,
            700,
            y)

                e.Graphics.DrawString(
            row.Cells("employee_name").Value.ToString,
            NormalFont,
            Brushes.Black,
            560,
            y)

                e.Graphics.DrawString(
            row.Cells("job_title").Value.ToString,
            NormalFont,
            Brushes.Black,
            400,
            y)

                e.Graphics.DrawString(
            row.Cells("contract_type").Value.ToString,
            NormalFont,
            Brushes.Black,
            250,
            y)

                e.Graphics.DrawString(
            row.Cells("net_salary").Value.ToString,
            NormalFont,
            Brushes.Black,
            100,
            y)

                y += 30

            End If

        Next

        y += 50

        e.Graphics.DrawString(
    "تاريخ الطباعة : " &
    Date.Now.ToShortDateString(),
    NormalFont,
    Brushes.Black,
    550,
    y)

        y += 60

        e.Graphics.DrawString(
    "التوقيع : __________________",
    NormalFont,
    Brushes.Black,
    500,
    y)

    End Sub
    Private Sub pdEmployeesGeneral_PrintPage(
sender As Object,
e As PrintPageEventArgs) _
Handles pdEmployeesGeneral.PrintPage

        ' كود التقرير العام



        Dim TitleFont As New Drawing.Font("Arial", 18, FontStyle.Bold)
        Dim HeaderFont As New Drawing.Font("Arial", 11, FontStyle.Bold)
        Dim NormalFont As New Drawing.Font("Arial", 10)

        '====================
        ' الترويسة
        '====================

        e.Graphics.DrawString(
    "الجمهورية اليمنية",
    HeaderFont,
    Brushes.Black,
    550,
    20)

        e.Graphics.DrawString(
    "وزارة التعليم الفني والتدريب المهني",
    HeaderFont,
    Brushes.Black,
    470,
    45)

        e.Graphics.DrawString(
    "معهد الخنساء",
    HeaderFont,
    Brushes.Black,
    600,
    70)

        e.Graphics.DrawString(
    "Republic Of Yemen",
    HeaderFont,
    Brushes.Black,
    20,
    20)

        e.Graphics.DrawString(
    "Ministry Of Technical Education",
    HeaderFont,
    Brushes.Black,
    20,
    45)

        e.Graphics.DrawString(
    "Al-Khansa Institute",
    HeaderFont,
    Brushes.Black,
    20,
    70)

        e.Graphics.DrawLine(
    Pens.Purple,
    20,
    110,
    780,
    110)

        '====================
        ' عنوان التقرير
        '====================

        e.Graphics.DrawString(
    "تقرير الموظفين العام",
    TitleFont,
    Brushes.Purple,
    280,
    130)

        '====================
        ' رؤوس الأعمدة
        '====================

        Dim y As Integer = 220

        e.Graphics.DrawString("الرقم", HeaderFont, Brushes.Black, 40, y)
        e.Graphics.DrawString("الاسم", HeaderFont, Brushes.Black, 120, y)
        e.Graphics.DrawString("المسمى", HeaderFont, Brushes.Black, 300, y)
        e.Graphics.DrawString("نوع التعاقد", HeaderFont, Brushes.Black, 450, y)
        e.Graphics.DrawString("صافي الراتب", HeaderFont, Brushes.Black, 650, y)

        y += 30

        e.Graphics.DrawLine(
    Pens.Black,
    20,
    y,
    780,
    y)

        y += 20

        '====================
        ' بيانات الموظفين
        '====================

        For i As Integer = 0 To DataGridView2.Rows.Count - 1

            If DataGridView2.Rows(i).Cells(0).Value IsNot Nothing Then

                e.Graphics.DrawString(
            DataGridView2.Rows(i).Cells(0).Value.ToString(),
            NormalFont,
            Brushes.Black,
            40,
            y)

                e.Graphics.DrawString(
            DataGridView2.Rows(i).Cells(1).Value.ToString(),
            NormalFont,
            Brushes.Black,
            120,
            y)

                e.Graphics.DrawString(
            DataGridView2.Rows(i).Cells(5).Value.ToString(),
            NormalFont,
            Brushes.Black,
            300,
            y)

                e.Graphics.DrawString(
            DataGridView2.Rows(i).Cells(4).Value.ToString(),
            NormalFont,
            Brushes.Black,
            450,
            y)

                e.Graphics.DrawString(
            DataGridView2.Rows(i).Cells(8).Value.ToString(),
            NormalFont,
            Brushes.Black,
            650,
            y)

                y += 25

            End If

        Next

        '====================
        ' التذييل
        '====================

        y += 30

        e.Graphics.DrawString(
    "عدد الموظفين : " &
    (DataGridView2.Rows.Count - 1).ToString(),
    HeaderFont,
    Brushes.Black,
    40,
    y)

        e.Graphics.DrawString(
    "تاريخ الطباعة : " &
    Date.Now.ToShortDateString(),
    HeaderFont,
    Brushes.Black,
    550,
    y)

    End Sub
    Private Sub LoadStatistics()

        Try

            OpenConnection()

            Dim cmd As New OleDbCommand(
        "SELECT COUNT(*) FROM Employee",
        conn)

            Label21.Text =
        cmd.ExecuteScalar().ToString()

            cmd = New OleDbCommand(
        "SELECT COUNT(*) FROM Employee WHERE gender='ذكر'",
        conn)

            Label19.Text =
        cmd.ExecuteScalar().ToString()

            cmd = New OleDbCommand(
        "SELECT COUNT(*) FROM Employee WHERE gender='أنثى'",
        conn)

            Label22.Text =
        cmd.ExecuteScalar().ToString()

            cmd = New OleDbCommand(
        "SELECT SUM(net_salary) FROM EmployeePayments",
        conn)

            If IsDBNull(cmd.ExecuteScalar()) Then

                Label15.Text = "0"

            Else

                Label15.Text =
            cmd.ExecuteScalar().ToString()

            End If

            cmd = New OleDbCommand(
        "SELECT COUNT(*) FROM Employee WHERE job_title='متعاقد'",
        conn)

            Label13.Text =
        cmd.ExecuteScalar().ToString()

        Catch ex As Exception

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub FillDataGridView()

        Dim dt As New DataTable

        OpenConnection()

        Dim da As New OleDbDataAdapter(
        "SELECT * FROM Employee",
        conn)

        da.Fill(dt)

        DataGridView2.DataSource = dt

        CloseConnection()

        'العناوين العربية
        DataGridView2.Columns("employee_id").HeaderText = "رقم الموظف"
        DataGridView2.Columns("employee_name").HeaderText = "اسم الموظف"
        DataGridView2.Columns("phone").HeaderText = "الهاتف"
        DataGridView2.Columns("gender").HeaderText = "الجنس"
        DataGridView2.Columns("contract_type").HeaderText = "نوع التعاقد"
        DataGridView2.Columns("job_title").HeaderText = "المسمى الوظيفي"

    End Sub
    Private Sub SaveSalary()

        Dim sql As String = "INSERT INTO EmployeePayments(employee_id,
bonuses,

BasicSalary,
NetSalary,
payment_date,
LectureCount,
AbsenceCount,
Month,
Year
)

VALUES
(?,?,?,?,?,?,?,?,?)"

        Dim cmd As New OleDbCommand(sql, conn)

        cmd.Parameters.AddWithValue("@p1", Val(TextBox1.Text))
        cmd.Parameters.AddWithValue("@p2", Val(TextBox11.Text))
        cmd.Parameters.AddWithValue("@p3", Val(TextBox7.Text))
        cmd.Parameters.AddWithValue("@p4", Val(TextBox4.Text))
        cmd.Parameters.AddWithValue("@p5", DateTimePicker1.Value)
        cmd.Parameters.AddWithValue("@p6", Val(TextBox10.Text))
        cmd.Parameters.AddWithValue("@p7", Val(TextBox5.Text))
        cmd.Parameters.AddWithValue("@p8", ComboBox1.Text)
        cmd.Parameters.AddWithValue("@p9", TextBox8.Text)

        cmd.ExecuteNonQuery()

    End Sub




    Private Sub eemp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Width = 100
        Me.Height = 100
        Timer1.Start()

        FillDataGridView()


        LoadStatistics()

        ' الجنس
        ComboBox3.Items.Clear()

        ComboBox3.Items.Add("ذكر")
        ComboBox3.Items.Add("أنثى")

        ' المسمى الوظيفي
        ComboBox4.Items.Clear()

        ComboBox4.Items.Add("متعاقد")
        ComboBox4.Items.Add("إداري")
        ComboBox4.Items.Add("عامل نظافة")
        ComboBox4.Items.Add("أساسي")

        ' نوع التعاقد
        ComboBox5.Items.Clear()

        ComboBox5.Items.Add("رسمي / دائم")
        ComboBox5.Items.Add("بالساعات")
        ComboBox5.Items.Add("بالحاضرة")
        ComboBox5.Items.Add("تعاقد سنوي")
        ComboBox5.Items.Add("شهري")
        ComboBox5.Items.Add("متعاون")
        ComboBox5.Items.Add("مؤقت")
        ComboBox5.Items.Add("تحت التجربة")
        ComboBox1.Items.Clear()

        ComboBox1.Items.Add("يناير")
        ComboBox1.Items.Add("فبراير")
        ComboBox1.Items.Add("مارس")
        ComboBox1.Items.Add("ابريل")
        ComboBox1.Items.Add("مايو")
        ComboBox1.Items.Add("يونيو")
        ComboBox1.Items.Add("يوليو")
        ComboBox1.Items.Add("اغسطس")
        ComboBox1.Items.Add("سبتمبر")
        ComboBox1.Items.Add("اكتوبر")
        ComboBox1.Items.Add("نوفمبر")
        ComboBox1.Items.Add("ديسمبر")

        FillDataGridView()




        TextBox10.Enabled = False
        TextBox5.Enabled = False

        TextBox10.Text = "0"
        TextBox5.Text = "0"

        If CurrentUserRole = "سكرتارية" Then

            Button1.Enabled = False

            Button4.Enabled = False

        End If
    End Sub




    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

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
        mane.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem8.Click
        supports.Show()
        Me.Close()
    End Sub

    Private Sub خروجToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles خروجToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Dashboard.Show()
        Me.Close()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Try

            Dim save As New SaveFileDialog

            save.Filter = "PDF Files|*.pdf"
            save.Title = "حفظ تقرير الموظفين"
            save.FileName = "تقرير_الموظفين_" &
                    Date.Now.ToString("yyyyMMdd")

            If save.ShowDialog() <> DialogResult.OK Then Exit Sub

            Dim doc As New iTextSharp.text.Document(
    iTextSharp.text.PageSize.A4)

            PdfWriter.GetInstance(
    doc,
    New IO.FileStream(
    save.FileName,
    IO.FileMode.Create))

            doc.Open()

            doc.Add(New Paragraph(
    "الجمهورية اليمنية"))

            doc.Add(New Paragraph(
    "وزارة التعليم الفني والتدريب المهني"))

            doc.Add(New Paragraph(
    "معهد الخنساء"))

            doc.Add(New Paragraph(" "))

            doc.Add(New Paragraph(
    "تقرير الموظفين"))

            doc.Add(New Paragraph(
    "تاريخ التقرير : " &
    Date.Now.ToString("yyyy/MM/dd")))

            doc.Add(New Paragraph(
    "===================================="))

            doc.Add(New Paragraph(" "))

            For Each row As DataGridViewRow In DataGridView2.Rows

                If Not row.IsNewRow Then

                    Dim empName As String =
            row.Cells("employee_name").Value.ToString()

                    Dim jobTitle As String =
            row.Cells("job_title").Value.ToString()

                    doc.Add(New Paragraph(
            "الموظف : " &
            empName &
            "     |     المسمى : " &
            jobTitle))

                End If

            Next

            doc.Add(New Paragraph(" "))

            doc.Add(New Paragraph(
    "عدد الموظفين : " &
    (DataGridView2.Rows.Count - 1).ToString()))

            doc.Close()

            MsgBox("تم إنشاء ملف PDF بنجاح")

            Process.Start(save.FileName)

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ComboBox4.Text = "متعاقد" Then

            TextBox10.Enabled = True
            TextBox5.Enabled = True

        Else

            TextBox10.Text = "0"
            TextBox5.Text = "0"

            TextBox10.Enabled = False
            TextBox5.Enabled = False

        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        CalculateNet()
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        CalculateNet()
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

        If e.RowIndex >= 0 Then

            TextBox1.Text =
        DataGridView2.CurrentRow.Cells("employee_id").Value.ToString()

            TextBox2.Text =
        DataGridView2.CurrentRow.Cells("employee_name").Value.ToString()

            TextBox3.Text =
        DataGridView2.CurrentRow.Cells("phone").Value.ToString()

            ComboBox3.Text =
        DataGridView2.CurrentRow.Cells("gender").Value.ToString()

            ComboBox5.Text =
        DataGridView2.CurrentRow.Cells("contract_type").Value.ToString()

            ComboBox4.Text =
        DataGridView2.CurrentRow.Cells("job_title").Value.ToString()

        End If
        DataGridView2.Columns("employee_id").HeaderText = "رقم الموظف"

        DataGridView2.Columns("employee_name").HeaderText = "اسم الموظف"

        DataGridView2.Columns("phone").HeaderText = "الهاتف"

        DataGridView2.Columns("gender").HeaderText = "الجنس"

        DataGridView2.Columns("contract_type").HeaderText = "نوع التعاقد"

        DataGridView2.Columns("job_title").HeaderText = "المسمى الوظيفي"

        DataGridView2.Columns("lectures_count").HeaderText = "عدد المحاضرات"

        DataGridView2.Columns("absence_count").HeaderText = "عدد الغياب"

        DataGridView2.Columns("paid_amount").HeaderText = "المبلغ المدفوع"

        DataGridView2.Columns("bonus").HeaderText = "المكافآت"

        DataGridView2.Columns("net_salary").HeaderText = "صافي المبلغ"

        DataGridView2.Columns("salary_month").HeaderText = "الشهر"

        DataGridView2.Columns("salary_year").HeaderText = "السنة"

        DataGridView2.Columns("payment_date").HeaderText = "تاريخ الدفع"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try

            OpenConnection()

            ' حفظ الموظف

            SaveSalary()

            ' إنشاء سند الصرف
            Dim cmd2 As New OleDbCommand(
    "INSERT INTO ExpenseVouchers
    (VoucherDate,VoucherNo,Amount,payment_Method,PayeeName,employee_id,Notes)
    VALUES (?,?,?,?,?,?,?)",
    conn)

            cmd2.Parameters.AddWithValue("@p1", DateTimePicker1.Value)

            cmd2.Parameters.AddWithValue("@p2",
    "EMP" & TextBox1.Text)

            cmd2.Parameters.AddWithValue("@p3",
    Val(TextBox4.Text))

            cmd2.Parameters.AddWithValue("@p4",
    "راتب موظف")

            cmd2.Parameters.AddWithValue("@p5",
    TextBox2.Text)

            cmd2.Parameters.AddWithValue("@p6",
    Val(TextBox1.Text))

            cmd2.Parameters.AddWithValue("@p7",
    "صرف راتب الموظف")

            cmd2.ExecuteNonQuery()

            ' فتح سند الصرف
            Dim frm As New Receipt

            frm.TextBox1.Text =
    "EMP" & TextBox1.Text

            frm.TextBox2.Text =
    TextBox2.Text

            frm.TextBox3.Text =
    TextBox4.Text

            frm.ComboBox4.Text =
    "راتب موظف"

            frm.DateTimePicker2.Value =
    DateTimePicker1.Value

            frm.TextBox4.Text =
    "صرف راتب الموظف"

            LoadStatistics()

            MsgBox("تم الحفظ بنجاح")

            frm.ShowDialog()

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click


        Try

            OpenConnection()

            Dim cmd As New OleDbCommand(
        "DELETE FROM Employee WHERE employee_id=?",
        conn)

            cmd.Parameters.AddWithValue("@p1",
        Val(TextBox1.Text))

            cmd.ExecuteNonQuery()

            MsgBox("تم الحذف")
            LoadStatistics()

            FillDataGridView()

            ClearData()

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()

        TextBox4.Text = "0"
        TextBox5.Text = "0"
        TextBox7.Text = "0"
        TextBox8.Clear()
        TextBox10.Text = "0"
        TextBox11.Text = "0"

        ComboBox1.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        ComboBox5.SelectedIndex = -1

        DateTimePicker1.Value = Date.Today
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Try

            OpenConnection()

            Dim sql As String = "UPDATE Employee SET
        employee_name=?,
        phone=?,
        gender=?,
        contract_type=?,
        job_title=?
        WHERE employee_id=?"

            Dim cmd As New OleDbCommand(sql, conn)

            cmd.Parameters.AddWithValue("@p1", TextBox2.Text)
            cmd.Parameters.AddWithValue("@p2", TextBox3.Text)
            cmd.Parameters.AddWithValue("@p3", ComboBox3.Text)
            cmd.Parameters.AddWithValue("@p4", ComboBox5.Text)
            cmd.Parameters.AddWithValue("@p5", ComboBox4.Text)
            cmd.Parameters.AddWithValue("@p6", Val(TextBox1.Text))

            cmd.ExecuteNonQuery()

            MsgBox("تم التعديل")
            LoadStatistics()

            FillDataGridView()

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try


    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged



        If TextBox6.Text = "" Then

            FillDataGridView()

            Exit Sub

        End If

        Try

            Dim dt As New DataTable

            OpenConnection()

            Dim da As New OleDbDataAdapter(
        "SELECT * FROM Employee WHERE employee_name LIKE ?",
        conn)

            da.SelectCommand.Parameters.AddWithValue(
        "@p1",
        "%" & TextBox6.Text & "%")

            da.Fill(dt)

            DataGridView2.DataSource = dt

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub صرفToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles صرفToolStripMenuItem.Click

        Receipt.ShowDialog()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FillDataGridView()
        LoadStatistics()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim pp As New PrintPreviewDialog

        pp.Document = pdEmployeesSummary

        pp.WindowState =
    FormWindowState.Maximized

        pp.ShowDialog()

    End Sub



    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Try

            Dim app As New Excel.Application

            Dim wb As Excel.Workbook

            Dim ws As Excel.Worksheet

            wb = app.Workbooks.Add()

            ws = wb.Sheets(1)

            For c As Integer = 0 To DataGridView2.Columns.Count - 1

                ws.Cells(1, c + 1) =
                DataGridView2.Columns(c).HeaderText

            Next

            For r As Integer = 0 To DataGridView2.Rows.Count - 1

                For c As Integer = 0 To DataGridView2.Columns.Count - 1

                    ws.Cells(r + 2, c + 1) =
                    DataGridView2.Rows(r).Cells(c).Value

                Next

            Next

            app.Visible = True

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim pp As New PrintPreviewDialog

        pp.Document = pdEmployeesGeneral

        pp.WindowState =
    FormWindowState.Maximized

        pp.ShowDialog()

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        LoadStatistics()
    End Sub

    Private Sub Guna2GradientButton1_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton1.Click
        Dim pp As New PrintPreviewDialog

        pp.Document = pdMonthlySalary

        pp.WindowState =
        FormWindowState.Maximized

        pp.ShowDialog()
    End Sub

    Private Sub Guna2GradientButton2_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton2.Click
        LoadMonthlyReport()

        Dim pp As New PrintPreviewDialog

        pp.Document = pdMonthlySalary

        pp.WindowState =
        FormWindowState.Maximized

        pp.ShowDialog()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Me.Width < 800 Then
            Me.Width += 1309
            Me.Height += 928
        Else
            Timer1.Stop()
        End If
    End Sub
End Class