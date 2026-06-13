Imports System.Data.OleDb
Imports System.Drawing.Printing
Imports Microsoft.Office.Interop
Imports iTextSharp.text.pdf
Public Class studnt
    Public Property StringDirection As Object

    Sub LoadDataGrid()
        Try
            OpenConnection()

            Dim sql As String =
                "SELECT " &
                "S.StudentID AS [رقم الطالبة], " &
                "S.StudentName AS [اسم الطالبة], " &
                "S.Phone AS [الهاتف], " &
                "D.DepartmentName AS [التخصص], " &
                "L.LevelName AS [المستوى], " &
                "ST.StudyTypeName AS [نوع الدراسة], " &
                "Sy.AcademicYear AS [السنة الدراسية], " &
                "S.EnrollmentDate AS [تاريخ الالتحاق], " &
                "S.StudentStatus AS [الحالة], " &
                "NZ(Sy.FeeAmount, 0) AS [إجمالي الرسوم], " &
                "NZ(P.paid_Amount, 0) AS [المبلغ المدفوع], " &
                "P.payment_Date AS [تاريخ الدفع] " &
                "FROM (((((Students AS S " &
                "LEFT JOIN Departments AS D ON S.DepartmentID = D.DepartmentID) " &
                "LEFT JOIN Levels AS L ON S.LevelID = L.LevelID) " &
                "LEFT JOIN Study AS Sy ON S.StudentID = Sy.StudentID) " &
                "LEFT JOIN StudyType AS ST ON Sy.StudyTypeID = ST.StudyTypeID) " &
                "LEFT JOIN payment AS P ON S.StudentID = P.StudentID) " &
                "ORDER BY S.StudentID"

            Dim da As New OleDbDataAdapter(sql, conn)
            Dim dt As New DataTable()
            da.Fill(dt)

            DataGridView2.DataSource = dt
            FormatGrid()

        Catch ex As Exception
            MsgBox("خطأ في تحميل البيانات : " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub
    ' دالة لتوليد رقم الطالبة القادم تلقائياً
    Sub GenerateNextStudentID()
        Try
            OpenConnection()
            ' جلب أعلى رقم موجود في جدول الطلاب
            Dim cmd As New OleDbCommand("SELECT MAX(StudentID) FROM Students", conn)
            Dim result = cmd.ExecuteScalar()

            If result Is DBNull.Value Or result Is Nothing Then
                ' إذا كان الجدول فارغاً تماماً، نبدأ من الرقم 1
                TextBox1.Text = "1"
            Else
                ' إذا وجدنا أرقاماً، نأخذ الأعلى ونزيد عليه 1
                TextBox1.Text = (Convert.ToInt32(result) + 1).ToString()
            End If

            ' جعل التكست بوكس للقراءة فقط حتى لا يقوم المستخدم بتعديله يدوياً
            TextBox1.ReadOnly = True
            TextBox1.BackColor = Color.LightGray ' تلميح بصري للمستخدم أنه تلقائي

        Catch ex As Exception
            MsgBox("خطأ في توليد رقم الطالبة التلقائي: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub PrintStudentsReport(sender As Object, e As PrintPageEventArgs)

        Try

            Dim titleFont As New Font("Tahoma", 16, FontStyle.Bold)
            Dim headerFont As New Font("Tahoma", 10, FontStyle.Bold)
            Dim bodyFont As New Font("Tahoma", 9)

            Dim formatRight As New StringFormat()
            formatRight.Alignment = StringAlignment.Far
            formatRight.FormatFlags = StringFormatFlags.DirectionRightToLeft

            Dim formatCenter As New StringFormat()
            formatCenter.Alignment = StringAlignment.Center

            Dim y As Integer = 30

            '=========================
            ' الترويسة
            '=========================

            Dim rightText As String =
        "الجمهورية اليمنية" & vbCrLf &
        "وزارة التعليم الفني والتدريب المهني" & vbCrLf &
        "معهد الخنساء للتدريب الفني"

            e.Graphics.DrawString(
        rightText,
        headerFont,
        Brushes.Black,
        New RectangleF(520, y, 250, 80),
        formatRight)

            Dim leftText As String =
        "Republic of Yemen" & vbCrLf &
        "Ministry of Technical Education" & vbCrLf &
        "Al-Khansa Institute"

            e.Graphics.DrawString(
        leftText,
        headerFont,
        Brushes.Black,
        New RectangleF(30, y, 250, 80))

            Dim imagePath As String =
        "C:\Users\ENJAZ\Desktop\مشروع التخرج\IMG-20251103-WA0004.jpg"

            If IO.File.Exists(imagePath) Then

                Using logo As Image =
            Image.FromFile(imagePath)

                    e.Graphics.DrawImage(
                logo,
                350,
                y,
                90,
                90)

                End Using

            End If

            y += 95

            e.Graphics.DrawLine(
        New Pen(Color.Purple, 3),
        20,
        y,
        800,
        y)

            y += 15

            e.Graphics.DrawString(
        "كشف بيانات الطالبات التفصيلي",
        titleFont,
        Brushes.Purple,
        New RectangleF(20, y, 780, 40),
        formatCenter)

            y += 50

            '=========================
            ' رؤوس الأعمدة
            '=========================

            e.Graphics.FillRectangle(
        Brushes.Plum,
        20,
        y,
        790,
        30)

            e.Graphics.DrawString("الرقم", headerFont, Brushes.White, 790, y + 5, formatRight)
            e.Graphics.DrawString("الاسم", headerFont, Brushes.White, 710, y + 5, formatRight)
            e.Graphics.DrawString("التخصص", headerFont, Brushes.White, 620, y + 5, formatRight)
            e.Graphics.DrawString("المستوى", headerFont, Brushes.White, 520, y + 5, formatRight)
            e.Graphics.DrawString("نوع الدراسة", headerFont, Brushes.White, 430, y + 5, formatRight)
            e.Graphics.DrawString("السنة", headerFont, Brushes.White, 340, y + 5, formatRight)
            e.Graphics.DrawString("الرسوم", headerFont, Brushes.White, 250, y + 5, formatRight)
            e.Graphics.DrawString("المدفوع", headerFont, Brushes.White, 150, y + 5, formatRight)
            e.Graphics.DrawString("الحالة", headerFont, Brushes.White, 70, y + 5, formatRight)

            y += 35

            '=========================
            ' البيانات
            '=========================

            Dim count As Integer = 0

            For Each row As DataGridViewRow In DataGridView2.Rows

                If row.IsNewRow Then Continue For

                Dim id As String =
            If(row.Cells("رقم الطالبة").Value IsNot Nothing,
            row.Cells("رقم الطالبة").Value.ToString(),
            "")

                Dim name As String =
            If(row.Cells("اسم الطالبة").Value IsNot Nothing,
            row.Cells("اسم الطالبة").Value.ToString(),
            "")

                Dim dept As String =
            If(row.Cells("التخصص").Value IsNot Nothing,
            row.Cells("التخصص").Value.ToString(),
            "")

                Dim level As String =
            If(row.Cells("المستوى").Value IsNot Nothing,
            row.Cells("المستوى").Value.ToString(),
            "")

                Dim studyType As String =
            If(row.Cells("نوع الدراسة").Value IsNot Nothing,
            row.Cells("نوع الدراسة").Value.ToString(),
            "")

                Dim year As String =
            If(row.Cells("السنة الدراسية").Value IsNot Nothing,
            row.Cells("السنة الدراسية").Value.ToString(),
            "")

                Dim fees As String =
            If(row.Cells("إجمالي الرسوم").Value IsNot Nothing,
            row.Cells("إجمالي الرسوم").Value.ToString(),
            "0")

                Dim paid As String =
            If(row.Cells("المبلغ المدفوع").Value IsNot Nothing,
            row.Cells("المبلغ المدفوع").Value.ToString(),
            "0")

                e.Graphics.DrawString(id, bodyFont, Brushes.Black, 790, y, formatRight)
                e.Graphics.DrawString(name, bodyFont, Brushes.Black, 710, y, formatRight)
                e.Graphics.DrawString(dept, bodyFont, Brushes.Black, 620, y, formatRight)
                e.Graphics.DrawString(level, bodyFont, Brushes.Black, 520, y, formatRight)
                e.Graphics.DrawString(studyType, bodyFont, Brushes.Black, 430, y, formatRight)
                e.Graphics.DrawString(year, bodyFont, Brushes.Black, 340, y, formatRight)
                e.Graphics.DrawString(fees, bodyFont, Brushes.DarkBlue, 250, y, formatRight)
                e.Graphics.DrawString(paid, bodyFont, Brushes.DarkGreen, 150, y, formatRight)

                y += 25

                e.Graphics.DrawLine(
            Pens.LightGray,
            20,
            y,
            810,
            y)

                y += 5

                count += 1

            Next

            y += 20

            e.Graphics.DrawString(
        "إجمالي عدد الطالبات : " & count,
        headerFont,
        Brushes.DarkBlue,
        790,
        y,
        formatRight)

        Catch ex As Exception

            MsgBox(
        "خطأ في طباعة التقرير : " &
        ex.Message,
        MsgBoxStyle.Critical)

        End Try

    End Sub


    Sub FormatGrid()
        With DataGridView2
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            .AllowUserToAddRows = False
            .ReadOnly = True
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 10, FontStyle.Bold)
            .DefaultCellStyle.Font = New Font("Tahoma", 10)
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .RowTemplate.Height = 30
        End With
    End Sub

    Sub LoadDepartmentReport()
        Try
            OpenConnection()

            ' تم تصحيح أسماء الحقول ومكان حقل FeeAmount لربطه بجدول Study الصحيح
            Dim sql As String =
"SELECT
D.DepartmentName AS [التخصص],
COUNT(S.StudentID) AS [عدد الطالبات],
SUM(Sy.FeeAmount) AS [إجمالي الرسوم],
SUM(P.paid_Amount) AS [إجمالي المدفوع]
FROM (((
Students AS S
LEFT JOIN Departments AS D
ON S.DepartmentID=D.DepartmentID)
LEFT JOIN Study AS Sy
ON S.StudentID=Sy.StudentID)
LEFT JOIN payment AS P
ON S.StudentID=P.StudentID)
GROUP BY D.DepartmentName"

            Dim da As New OleDbDataAdapter(sql, conn)
            Dim dt As New DataTable
            da.Fill(dt)

            DataGridView2.DataSource = dt

        Catch ex As Exception
            MsgBox("خطأ في تحميل تقرير الأقسام: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try

    End Sub
    Private Sub PrintDepartmentReport(sender As Object, e As PrintPageEventArgs)

        Try
            Dim titleFont As New Font("Tahoma", 16, FontStyle.Bold)
            Dim bodyFont As New Font("Tahoma", 10)

            Dim headerFont As New Font("Arial", 11, FontStyle.Bold)

            Dim formatRight As New StringFormat()
            formatRight.Alignment = StringAlignment.Far
            formatRight.FormatFlags = StringFormatFlags.DirectionRightToLeft

            Dim formatCenter As New StringFormat()
            formatCenter.Alignment = StringAlignment.Center

            Dim y As Integer = 40

            '=========================
            ' الترويسة
            '=========================

            Dim rightText As String =
        "الجمهورية اليمنية" & vbCrLf &
        "وزارة التعليم الفني والتدريب المهني" & vbCrLf &
        "معهد الخنساء للتدريب الفني"

            e.Graphics.DrawString(
        rightText,
        headerFont,
        Brushes.Black,
        New RectangleF(530, y, 250, 80),
        formatRight)

            Dim leftText As String =
        "Republic of Yemen" & vbCrLf &
        "Ministry of Technical Education" & vbCrLf &
        "Al-Khansa Institute"

            e.Graphics.DrawString(
        leftText,
        headerFont,
        Brushes.Black,
        New RectangleF(40, y, 250, 80))

            Dim imagePath As String =
        "C:\Users\ENJAZ\Desktop\مشروع التخرج\IMG-20251103-WA0004.jpg"

            If IO.File.Exists(imagePath) Then

                Using logo As Image =
            Image.FromFile(imagePath)

                    e.Graphics.DrawImage(
                logo,
                360,
                y - 10,
                90,
                90)

                End Using

            End If

            y += 90

            e.Graphics.DrawLine(
        New Pen(Color.Purple, 3),
        30,
        y,
        780,
        y)

            y += 15

            e.Graphics.DrawString(
        "تقرير الأقسام والإيرادات المالية",
        titleFont,
        Brushes.Purple,
        New RectangleF(30, y, 750, 40),
        formatCenter)

            '=========================
            ' رؤوس الجدول
            '=========================

            y += 50

            e.Graphics.FillRectangle(
        Brushes.Plum,
        30,
        y,
        750,
        30)

            e.Graphics.DrawString(
        "التخصص",
        headerFont,
        Brushes.Black,
        740,
        y + 5,
        formatRight)

            e.Graphics.DrawString(
        "عدد الطالبات",
        headerFont,
        Brushes.Black,
        520,
        y + 5,
        formatRight)

            e.Graphics.DrawString(
        "إجمالي الرسوم",
        headerFont,
        Brushes.Black,
        320,
        y + 5,
        formatRight)

            e.Graphics.DrawString(
        "إجمالي المدفوع",
        headerFont,
        Brushes.Black,
        140,
        y + 5,
        formatRight)

            y += 35

            '=========================
            ' البيانات
            '=========================

            For Each row As DataGridViewRow In DataGridView2.Rows

                If row.IsNewRow Then Continue For

                Dim deptName As String =
            If(row.Cells("التخصص").Value IsNot Nothing,
            row.Cells("التخصص").Value.ToString(),
            "")

                Dim studentCount As String =
            If(row.Cells("عدد الطالبات").Value IsNot Nothing,
            row.Cells("عدد الطالبات").Value.ToString(),
            "0")

                Dim totalFees As String =
            If(row.Cells("إجمالي الرسوم").Value IsNot Nothing,
            row.Cells("إجمالي الرسوم").Value.ToString(),
            "0")

                Dim totalPaid As String =
            If(row.Cells("إجمالي المدفوع").Value IsNot Nothing,
            row.Cells("إجمالي المدفوع").Value.ToString(),
            "0")

                e.Graphics.DrawString(deptName, bodyFont, Brushes.Black, 740, y, formatRight)
                e.Graphics.DrawString(studentCount, bodyFont, Brushes.Black, 520, y, formatRight)
                e.Graphics.DrawString(totalFees, bodyFont, Brushes.Black, 320, y, formatRight)
                e.Graphics.DrawString(totalPaid, bodyFont, Brushes.DarkGreen, 140, y, formatRight)

                y += 30

                e.Graphics.DrawLine(
            Pens.LightGray,
            30,
            y,
            780,
            y)

            Next

        Catch ex As Exception

            MsgBox(
        "خطأ في طباعة تقرير الأقسام : " &
        ex.Message,
        MsgBoxStyle.Critical)

        End Try

    End Sub
    ' 2. دالة تعبئة القوائم المنسدلة

    Sub FillComboBoxes()
        Try
            Dim daDept As New OleDbDataAdapter("SELECT DepartmentID, DepartmentName FROM Departments", conn)
            Dim dtDept As New DataTable
            daDept.Fill(dtDept)
            ComboBox3.DataSource = dtDept
            ComboBox3.DisplayMember = "DepartmentName"
            ComboBox3.ValueMember = "DepartmentID"
            ComboBox3.SelectedIndex = -1

            Dim daLev As New OleDbDataAdapter("SELECT LevelID, LevelName FROM Levels", conn)
            Dim dtLev As New DataTable
            daLev.Fill(dtLev)
            ComboBox5.DataSource = dtLev
            ComboBox5.DisplayMember = "LevelName"
            ComboBox5.ValueMember = "LevelID"
            ComboBox5.SelectedIndex = -1

            Dim daST As New OleDbDataAdapter("SELECT StudyTypeID, StudyTypeName FROM StudyType", conn)
            Dim dtST As New DataTable
            daST.Fill(dtST)
            ComboBox1.DataSource = dtST
            ComboBox1.DisplayMember = "StudyTypeName"
            ComboBox1.ValueMember = "StudyTypeID"
            ComboBox1.SelectedIndex = -1

            ComboBox2.Items.Clear()
            ComboBox2.Items.AddRange(New Object() {"الفصل الأول", "الفصل الثاني", "الفصل الثالث", "الفصل الرابع"})
            ComboBox2.SelectedIndex = -1

            ComboBox4.Items.Clear()
            ComboBox4.Items.AddRange(New Object() {"نشط", "مستجد", "خريجة", "موقف", "باق", "منقول"})
            ComboBox4.SelectedIndex = -1
        Catch ex As Exception
            MsgBox("خطأ في تحميل القوائم المنسدلة: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub


    Private Sub studnt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            OpenConnection()
            FillComboBoxes()
            LoadDataGrid()
            GenerateNextStudentID()
            If CurrentUserRole = "سكرتارية" Then
                Button12.Enabled = False
                Button14.Enabled = False
            End If
            TextBox1.Focus()
            Me.Opacity = 1
        Catch ex As Exception
            MsgBox("خطأ أثناء تهيئة الشاشة: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub



    ' 3. كود زر الحفظ (تم إضافة TextBox7)
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        If TextBox1.Text.Trim() = "" Or TextBox2.Text.Trim() = "" Then
            MsgBox("الرجاء إدخال رقم واسم الطالبة أولاً", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If ComboBox1.SelectedIndex = -1 Or ComboBox3.SelectedIndex = -1 Or ComboBox5.SelectedIndex = -1 Then
            MsgBox("الرجاء اختيار (نوع الدراسة، التخصص، والمستوى) قبل الحفظ", MsgBoxStyle.Critical)
            Exit Sub
        End If

        Try
            OpenConnection()

            ' التحقق من عدم تكرار رقم الطالبة
            Dim checkCmd As New OleDbCommand("SELECT COUNT(*) FROM Students WHERE StudentID=?", conn)
            checkCmd.Parameters.AddWithValue("?", Val(TextBox1.Text.Trim()))

            If Convert.ToInt32(checkCmd.ExecuteScalar()) > 0 Then
                MsgBox("رقم الطالبة موجود مسبقاً", MsgBoxStyle.Critical)
                Exit Sub
            End If

            ' أولاً: جدول Students (تم وضع افتراضي للمستخدم رقم 1 لتفادي خطأ التكامل)
            Dim sql1 As String = "INSERT INTO Students (StudentID, StudentName, Phone, Notes, DepartmentID, LevelID, EnrollmentDate, StudentStatus, UserID) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Using cmd1 As New OleDbCommand(sql1, conn)
                cmd1.Parameters.AddWithValue("?", Val(TextBox1.Text))
                cmd1.Parameters.AddWithValue("?", TextBox2.Text.Trim())
                cmd1.Parameters.AddWithValue("?", TextBox3.Text.Trim())
                cmd1.Parameters.AddWithValue("?", TextBox5.Text.Trim())
                cmd1.Parameters.AddWithValue("?", ComboBox3.SelectedValue)
                cmd1.Parameters.AddWithValue("?", ComboBox5.SelectedValue)
                cmd1.Parameters.AddWithValue("?", DateTimePicker1.Value.Date)
                cmd1.Parameters.AddWithValue("?", ComboBox4.Text)
                cmd1.Parameters.AddWithValue("?", 1) ' رقم الـ User المسؤول عن الربط الحتمي
                cmd1.ExecuteNonQuery()
            End Using

            ' ثانياً: جدول Study
            Dim sql2 As String = "INSERT INTO Study (StudentID, StudyTypeID, Semester, AcademicYear, FeeAmount) VALUES (?, ?, ?, ?, ?)"
            Using cmd2 As New OleDbCommand(sql2, conn)
                cmd2.Parameters.AddWithValue("?", Val(TextBox1.Text))
                cmd2.Parameters.AddWithValue("?", ComboBox1.SelectedValue)
                cmd2.Parameters.AddWithValue("?", ComboBox2.Text)
                cmd2.Parameters.AddWithValue("?", TextBox10.Text.Trim())
                cmd2.Parameters.AddWithValue("?", Val(TextBox7.Text))
                cmd2.ExecuteNonQuery()
            End Using

            ' ثالثاً: جدول payment
            Dim amount As Double = Val(TextBox6.Text)
            Dim sql3 As String = "INSERT INTO payment (StudentID, paid_Amount, payment_Date, payment_Method, Notes, UserID) VALUES (?, ?, ?, ?, ?, ?)"
            Using cmd3 As New OleDbCommand(sql3, conn)
                cmd3.Parameters.AddWithValue("?", Val(TextBox1.Text))
                cmd3.Parameters.AddWithValue("?", amount)
                cmd3.Parameters.AddWithValue("?", DateTimePicker2.Value.Date)
                cmd3.Parameters.AddWithValue("?", "نقداً")
                cmd3.Parameters.AddWithValue("?", "رسوم تسجيل طالبة")
                cmd3.Parameters.AddWithValue("?", 1)
                cmd3.ExecuteNonQuery()
            End Using

            ' رابعاً: جدول ReceiptVoucher
            Dim sql4 As String = "INSERT INTO ReceiptVoucher (ReceiptDate, VoucherNo, Amount, PaymentMethod, PayerName, Notes, CreatedBy, UserID, RevenueTypeID, StudentID) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Using cmd4 As New OleDbCommand(sql4, conn)
                cmd4.Parameters.AddWithValue("?", DateTimePicker2.Value.Date)
                cmd4.Parameters.AddWithValue("?", "V" & TextBox1.Text.Trim())
                cmd4.Parameters.AddWithValue("?", amount)
                cmd4.Parameters.AddWithValue("?", "نقداً")
                cmd4.Parameters.AddWithValue("?", TextBox2.Text.Trim())
                cmd4.Parameters.AddWithValue("?", "رسوم تسجيل طالبة")
                cmd4.Parameters.AddWithValue("?", "ADMIN")
                cmd4.Parameters.AddWithValue("?", 1)
                cmd4.Parameters.AddWithValue("?", 1)
                cmd4.Parameters.AddWithValue("?", Val(TextBox1.Text))
                cmd4.ExecuteNonQuery()
            End Using

            MsgBox("تم حفظ بيانات الطالبة بنجاح وإنشاء القيود المالية لها 🌟", MsgBoxStyle.Information)
            LoadDataGrid()
            ' امسحي الحقول القديمة وجهزي الشاشة للطالبة التالية بالرقم الجديد
            TextBox2.Clear() ' اسم الطالبة
            TextBox3.Clear() ' الهاتف
            TextBox5.Clear() ' الملاحظات
            TextBox6.Clear() ' المدفوع
            TextBox7.Clear() ' الرسوم
            TextBox10.Clear() ' السنة
            GenerateNextStudentID() ' 🌟 توليد الرقم الجديد للطالبة القادمة تلقائياً

        Catch ex As Exception
            MsgBox("خطأ في الحفظ والتسجيل:" & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub







    ' هذا هو الحدث الصحيح للرسم (هنا نستخدم e.Graphics)
    Private Sub PrintDocument_PrintPage(sender As Object, e As PrintPageEventArgs)
        Try
            Dim fontHeader As New Font("Arial", 11, FontStyle.Bold)
            Dim fontTitle As New Font("Arial", 16, FontStyle.Bold)

            Dim formatRight As New StringFormat()
            formatRight.Alignment = StringAlignment.Far
            formatRight.FormatFlags = StringFormatFlags.DirectionRightToLeft

            Dim formatLeft As New StringFormat() With {.Alignment = StringAlignment.Near}
            Dim y As Integer = 40

            Dim textRight As String = "الجمهورية اليمنية" & vbCrLf & "وزارة التعليم الفني والتدريب المهني" & vbCrLf & "معهد الخنساء"
            e.Graphics.DrawString(textRight, fontHeader, Brushes.Black, New RectangleF(530, y, 250, 80), formatRight)

            Dim textLeft As String = "Republic of Yemen" & vbCrLf & "Ministry of Technical Education" & vbCrLf & "Al-Khansa Institute"
            e.Graphics.DrawString(textLeft, fontHeader, Brushes.Black, New RectangleF(40, y, 250, 80), formatLeft)

            Dim imagePath As String = "C:\Users\ENJAZ\Desktop\مشروع التخرج\IMG-20251103-WA0004.jpg"
            If System.IO.File.Exists(imagePath) Then
                Using logo As Image = Image.FromFile(imagePath)
                    e.Graphics.DrawImage(logo, 360, y - 10, 90, 90)
                End Using
            End If

            y += 90
            e.Graphics.DrawLine(New Pen(Color.Black, 2), 30, y, 780, y)

            y += 15
            Dim formatCenter As New StringFormat() With {.Alignment = StringAlignment.Center}
            e.Graphics.DrawString("سجل الطالبات العام المعروض بالشاشة", fontTitle, Brushes.DarkBlue, New RectangleF(30, y, 750, 40), formatCenter)

            y += 50
            If DataGridView2.Rows.Count > 0 Then
                Dim bm As New Bitmap(Me.DataGridView2.Width, Me.DataGridView2.Height)
                DataGridView2.DrawToBitmap(bm, New Rectangle(0, 0, Me.DataGridView2.Width, Me.DataGridView2.Height))
                e.Graphics.DrawImage(bm, 40, y, 730, Me.DataGridView2.Height)
            Else
                e.Graphics.DrawString("لا توجد بيانات حالية لعرضها في الجدول.", fontHeader, Brushes.Red, 40, y)
            End If

        Catch ex As Exception
            MsgBox("خطأ في الرسم والطباعة الفورية: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        ' التأكد من أن السطر المحدد يحتوي على بيانات وليس فارغاً
        If e.RowIndex >= 0 Then
            Try
                Dim row As DataGridViewRow = DataGridView2.Rows(e.RowIndex)
                TextBox1.Text = If(row.Cells("رقم الطالبة").Value, "").ToString()
                TextBox2.Text = If(row.Cells("اسم الطالبة").Value, "").ToString()
                TextBox3.Text = If(row.Cells("الهاتف").Value, "").ToString()
                ComboBox3.Text = If(row.Cells("التخصص").Value, "").ToString()
                ComboBox5.Text = If(row.Cells("المستوى").Value, "").ToString()
                ComboBox1.Text = If(row.Cells("نوع الدراسة").Value, "").ToString()

                If row.Cells("تاريخ الالتحاق").Value IsNot DBNull.Value AndAlso row.Cells("تاريخ الالتحاق").Value IsNot Nothing Then
                    DateTimePicker1.Value = Convert.ToDateTime(row.Cells("تاريخ الالتحاق").Value)
                End If

                TextBox7.Text = If(row.Cells("إجمالي الرسوم").Value, "0").ToString()
                TextBox6.Text = If(row.Cells("المبلغ المدفوع").Value, "0").ToString()
                TextBox10.Text = If(row.Cells("السنة الدراسية").Value, "").ToString()
                ComboBox4.Text = If(row.Cells(" الحالة").Value, "").ToString()
            Catch
            End Try
        End If
    End Sub




    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Application.Exit()

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If TextBox1.Text.Trim() = "" Then
            MsgBox("حدد الطالبة من الجدول أولاً", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        If MsgBox("هل أنت متأكد من حذف هذه الطالبة وجميع السجلات والملفات المالية المرتبطة بها نهائياً؟", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "تأكيد الحذف") = MsgBoxResult.No Then
            Exit Sub
        End If
        Try
            OpenConnection()
            Dim id As String = TextBox1.Text.Trim()
            Dim cmd1 As New OleDbCommand("DELETE FROM ReceiptVoucher WHERE StudentID=?", conn)
            cmd1.Parameters.AddWithValue("?", id)
            cmd1.ExecuteNonQuery()

            Dim cmd2 As New OleDbCommand("DELETE FROM payment WHERE StudentID=?", conn)
            cmd2.Parameters.AddWithValue("?", id)
            cmd2.ExecuteNonQuery()

            Dim cmd3 As New OleDbCommand("DELETE FROM Study WHERE StudentID=?", conn)
            cmd3.Parameters.AddWithValue("?", id)
            cmd3.ExecuteNonQuery()

            Dim cmd4 As New OleDbCommand("DELETE FROM Students WHERE StudentID=?", conn)
            cmd4.Parameters.AddWithValue("?", id)
            cmd4.ExecuteNonQuery()

            MsgBox("تم حذف سجل الطالبة وقيودها المالية بالكامل بنجاح 🗑️", MsgBoxStyle.Information)
            LoadDataGrid()
            If Dashboard IsNot Nothing Then Dashboard.LoadStatistics()
        Catch ex As Exception
            MsgBox("خطأ أثناء الحذف: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub


    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        If TextBox1.Text.Trim() = "" Then
            MsgBox("الرجاء اختيار الطالبة من الجدول أولاً", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Try
            OpenConnection()

            ' تحديث جدول Students (تم ترتيب المعاملات بالتطابق التام مع علامات الاستهام)
            Dim sqlStudent As String = "UPDATE Students SET StudentName=?, Phone=?, Notes=?, DepartmentID=?, LevelID=?, EnrollmentDate=?, StudentStatus=? WHERE StudentID=?"
            Using cmdStudent As New OleDbCommand(sqlStudent, conn)
                cmdStudent.Parameters.AddWithValue("?", TextBox2.Text.Trim())
                cmdStudent.Parameters.AddWithValue("?", TextBox3.Text.Trim())
                cmdStudent.Parameters.AddWithValue("?", TextBox5.Text.Trim())
                cmdStudent.Parameters.AddWithValue("?", If(ComboBox3.SelectedValue IsNot Nothing, ComboBox3.SelectedValue, DBNull.Value))
                cmdStudent.Parameters.AddWithValue("?", If(ComboBox5.SelectedValue IsNot Nothing, ComboBox5.SelectedValue, DBNull.Value))
                cmdStudent.Parameters.AddWithValue("?", DateTimePicker1.Value.Date)
                cmdStudent.Parameters.AddWithValue("?", ComboBox4.Text)
                cmdStudent.Parameters.AddWithValue("?", Val(TextBox1.Text)) ' المعرّف في نهاية الشرط
                cmdStudent.ExecuteNonQuery()
            End Using

            ' تحديث جدول Study
            Dim sqlStudy As String = "UPDATE Study SET StudyTypeID=?, Semester=?, AcademicYear=?, FeeAmount=? WHERE StudentID=?"
            Using cmdStudy As New OleDbCommand(sqlStudy, conn)
                cmdStudy.Parameters.AddWithValue("?", If(ComboBox1.SelectedValue IsNot Nothing, ComboBox1.SelectedValue, DBNull.Value))
                cmdStudy.Parameters.AddWithValue("?", ComboBox2.Text)
                cmdStudy.Parameters.AddWithValue("?", TextBox10.Text.Trim())
                cmdStudy.Parameters.AddWithValue("?", Val(TextBox7.Text))
                cmdStudy.Parameters.AddWithValue("?", Val(TextBox1.Text))
                cmdStudy.ExecuteNonQuery()
            End Using

            ' تحديث جدول payment
            Dim sqlPay As String = "UPDATE payment SET paid_Amount=?, payment_Date=? WHERE StudentID=?"
            Using cmdPay As New OleDbCommand(sqlPay, conn)
                cmdPay.Parameters.AddWithValue("?", Val(TextBox6.Text))
                cmdPay.Parameters.AddWithValue("?", DateTimePicker2.Value.Date)
                cmdPay.Parameters.AddWithValue("?", Val(TextBox1.Text))
                cmdPay.ExecuteNonQuery()
            End Using

            ' تحديث جدول ReceiptVoucher
            Dim sqlVoucher As String = "UPDATE ReceiptVoucher SET ReceiptDate=?, Amount=?, PayerName=? WHERE StudentID=?"
            Using cmdVoucher As New OleDbCommand(sqlVoucher, conn)
                cmdVoucher.Parameters.AddWithValue("?", DateTimePicker2.Value.Date)
                cmdVoucher.Parameters.AddWithValue("?", Val(TextBox6.Text))
                cmdVoucher.Parameters.AddWithValue("?", TextBox2.Text.Trim())
                cmdVoucher.Parameters.AddWithValue("?", Val(TextBox1.Text))
                cmdVoucher.ExecuteNonQuery()
            End Using

            MsgBox("تم تعديل بيانات الطالبة وتحديث سجلاتها المالية بنجاح", MsgBoxStyle.Information)
            LoadDataGrid()

            If Dashboard IsNot Nothing Then Dashboard.LoadStatistics()

        Catch ex As Exception
            MsgBox("خطأ أثناء التعديل : " & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        Finally
            CloseConnection()
        End Try
    End Sub





    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        LoadDataGrid()
        MsgBox("تم تحديث وجل البيانات الحالية بنجاح", MsgBoxStyle.Information)
    End Sub



    Private Sub بحثToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ADD.Show()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        user.Show()
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs)
        MINE.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Clear() : TextBox2.Clear() : TextBox3.Clear()
        TextBox6.Clear() : TextBox7.Clear() : TextBox10.Clear()
        ComboBox1.SelectedIndex = -1 : ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1 : ComboBox4.SelectedIndex = -1
        ComboBox5.SelectedIndex = -1
        TextBox2.Focus()



    End Sub




    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label12.Text = DateTime.Now.ToString("dddd - hh:mm:ss tt")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' هذا الزر فقط يستدعي نافذة الطباعة
        Dim pd As New PrintDocument
        AddHandler pd.PrintPage, AddressOf PrintDocument_PrintPage
        Dim ppd As New PrintDialog
        ppd.Document = pd
        If ppd.ShowDialog() = DialogResult.OK Then pd.Print()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Try

            OpenConnection()

            Dim searchVal As String = TextBox4.Text.Trim()

            If searchVal = "" Then
                LoadDataGrid()
                Exit Sub
            End If

            Dim sql As String =
    "SELECT " &
    "S.StudentID AS [رقم الطالبة], " &
    "S.StudentName AS [اسم الطالبة], " &
    "S.Phone AS [الهاتف], " &
    "D.DepartmentName AS [التخصص], " &
    "L.LevelName AS [المستوى], " &
    "ST.StudyTypeName AS [نوع الدراسة], " &
    "Sy.AcademicYear AS [السنة الدراسية], " &
    "S.EnrollmentDate AS [تاريخ الالتحاق], " &
    "S.StudentStatus AS [الحالة], " &
    "Sy.FeeAmount AS [إجمالي الرسوم], " &
    "P.paid_Amount AS [المبلغ المدفوع], " &
    "P.payment_Date AS [تاريخ الدفع] " &
    "FROM (((((Students AS S " &
    "LEFT JOIN Departments AS D ON S.DepartmentID = D.DepartmentID) " &
    "LEFT JOIN Levels AS L ON S.LevelID = L.LevelID) " &
    "LEFT JOIN Study AS Sy ON S.StudentID = Sy.StudentID) " &
    "LEFT JOIN StudyType AS ST ON Sy.StudyTypeID = ST.StudyTypeID) " &
    "LEFT JOIN payment AS P ON S.StudentID = P.StudentID) " &
    "WHERE S.StudentID LIKE ? " &
    "OR S.StudentName LIKE ? " &
    "OR S.Phone LIKE ?"

            Dim da As New OleDbDataAdapter(sql, conn)

            da.SelectCommand.Parameters.AddWithValue("?", "%" & searchVal & "%")
            da.SelectCommand.Parameters.AddWithValue("?", "%" & searchVal & "%")
            da.SelectCommand.Parameters.AddWithValue("?", "%" & searchVal & "%")

            Dim dt As New DataTable

            da.Fill(dt)

            DataGridView2.DataSource = dt

            FormatGrid()

        Catch ex As Exception

            MsgBox(
    "خطأ في البحث : " &
    vbCrLf &
    ex.Message,
    MsgBoxStyle.Critical)

        Finally

            CloseConnection()

        End Try
    End Sub






    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim pd As New PrintDocument
        AddHandler pd.PrintPage, AddressOf PrintStudentsReport
        Dim dlg As New PrintDialog
        dlg.Document = pd
        If dlg.ShowDialog() = DialogResult.OK Then pd.Print()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        LoadDepartmentReport() ' جلب البيانات المحدثة أولاً
        Dim pd As New PrintDocument
        AddHandler pd.PrintPage, AddressOf PrintDepartmentReport
        Dim dlg As New PrintDialog
        dlg.Document = pd
        If dlg.ShowDialog() = DialogResult.OK Then pd.Print()
    End Sub



    Private Sub بحثToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles بحثToolStripMenuItem.Click
        ADD.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem1_Click_1(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        user.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem6_Click_1(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        MINE.Show()
        Me.Close()
    End Sub

    Private Sub حفظToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles حفظToolStripMenuItem.Click
        Dashboard.Show()
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

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Expenses.Show()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If Me.Opacity < 1 Then
            Me.Opacity += 0.05
        Else
            Timer2.Stop()
            Timer1.Start() ' تشغيل ساعة النظام بعد ظهور النموذج
        End If

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Try
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
            saveFileDialog.FileName = "تقرير_معهد_الخنساء_" & DateTime.Now.ToString("yyyyMMdd") & ".pdf"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                ' تحديد حجم الصفحة والهوامش باستخدام المسار الكامل لـ iTextSharp لمنع التعارض
                Dim doc As New iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 25, 25, 30, 30)
                iTextSharp.text.pdf.PdfWriter.GetInstance(doc, New System.IO.FileStream(saveFileDialog.FileName, System.IO.FileMode.Create))
                doc.Open()

                ' إعداد الخطوط العربية عبر خطوط النظام الافتراضية
                Dim arialFontPath As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf")
                Dim bfArabic As iTextSharp.text.pdf.BaseFont = iTextSharp.text.pdf.BaseFont.CreateFont(arialFontPath, iTextSharp.text.pdf.BaseFont.IDENTITY_H, iTextSharp.text.pdf.BaseFont.EMBEDDED)

                ' تعريف الخطوط مع تحديد مرجع iTextSharp صراحةً
                Dim fontHeader As New iTextSharp.text.Font(bfArabic, 10, iTextSharp.text.Font.NORMAL)
                Dim fontTitle As New iTextSharp.text.Font(bfArabic, 14, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.Blue)

                ' --- رسم الترويسة الفوقية (جدول خفي) ---
                Dim headerTable As New iTextSharp.text.pdf.PdfPTable(3)
                headerTable.WidthPercentage = 100
                headerTable.RunDirection = iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL
                headerTable.SetWidths({35, 30, 35})

                ' 1. القسم الأيمن (عربي)
                Dim cellRight As New iTextSharp.text.pdf.PdfPCell(New iTextSharp.text.Phrase("الجمهورية اليمنية" & vbCrLf & "وزارة التعليم الفني والتدريب المهني" & vbCrLf & "معهد الخنساء للتدريب الفني", fontHeader))
                cellRight.Border = iTextSharp.text.Rectangle.NO_BORDER
                cellRight.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT
                headerTable.AddCell(cellRight)

                ' 2. الشعار بالمنتصف
                Dim imagePath As String = "C:\Users\ENJAZ\Desktop\مشروع التخرج\IMG-20251103-WA0004.jpg"
                If System.IO.File.Exists(imagePath) Then
                    ' استخدام مسار iTextSharp الصريح للصورة للتخلص من تعارض System.Drawing.Image
                    Dim logo As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(imagePath)
                    logo.ScaleToFit(70, 70)
                    Dim cellLogo As New iTextSharp.text.pdf.PdfPCell(logo)
                    cellLogo.Border = iTextSharp.text.Rectangle.NO_BORDER
                    cellLogo.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER
                    headerTable.AddCell(cellLogo)
                Else
                    Dim emptyCell As New iTextSharp.text.pdf.PdfPCell(New iTextSharp.text.Phrase(""))
                    emptyCell.Border = iTextSharp.text.Rectangle.NO_BORDER
                    headerTable.AddCell(emptyCell)
                End If

                ' 3. القسم الأيسر (إنجليزي)
                Dim cellLeft As New iTextSharp.text.pdf.PdfPCell(New iTextSharp.text.Phrase("Republic of Yemen" & vbCrLf & "Ministry of Technical Education" & vbCrLf & "Al-Khansa Institute", fontHeader))
                cellLeft.Border = iTextSharp.text.Rectangle.NO_BORDER
                cellLeft.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT
                cellLeft.RunDirection = iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_LTR
                headerTable.AddCell(cellLeft)

                doc.Add(headerTable)

                ' خط فاصل تحت الترويسة
                Dim line As New iTextSharp.text.pdf.draw.LineSeparator(2.0F, 100.0F, iTextSharp.text.BaseColor.Black, iTextSharp.text.Element.ALIGN_CENTER, -5)
                doc.Add(New iTextSharp.text.Chunk(line))
                doc.Add(New iTextSharp.text.Paragraph(" ")) ' سطر فارغ للمسافة

                ' عنوان التقرير
                Dim pTitle As New iTextSharp.text.Paragraph("كشف بيانات الطالبات المالي والإداري التفصيلي", fontTitle)
                pTitle.Alignment = iTextSharp.text.Element.ALIGN_CENTER
                doc.Add(pTitle)
                doc.Add(New iTextSharp.text.Paragraph(" "))

                ' --- رسم جدول البيانات ---
                Dim visibleCols As Integer = 0
                For Each c As DataGridViewColumn In DataGridView2.Columns
                    If c.Visible Then visibleCols += 1
                Next

                Dim dataTable As New iTextSharp.text.pdf.PdfPTable(visibleCols)
                dataTable.WidthPercentage = 100
                dataTable.RunDirection = iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL

                ' عناوين أعمدة الجدول
                For Each col As DataGridViewColumn In DataGridView2.Columns
                    If col.Visible Then
                        Dim cell As New iTextSharp.text.pdf.PdfPCell(New iTextSharp.text.Phrase(col.HeaderText, New iTextSharp.text.Font(bfArabic, 11, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.White)))
                        cell.BackgroundColor = New iTextSharp.text.BaseColor(128, 0, 128) ' اللون البنفسجي للمعهد
                        cell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER
                        cell.Padding = 5
                        dataTable.AddCell(cell)
                    End If
                Next

                ' تعبئة بيانات أسطر الجدول
                For Each row As DataGridViewRow In DataGridView2.Rows
                    If row.IsNewRow Then Continue For
                    For Each col As DataGridViewColumn In DataGridView2.Columns
                        If col.Visible Then
                            Dim txt As String = If(row.Cells(col.Index).Value IsNot Nothing, row.Cells(col.Index).Value.ToString(), "")
                            Dim cell As New iTextSharp.text.pdf.PdfPCell(New iTextSharp.text.Phrase(txt, New iTextSharp.text.Font(bfArabic, 10)))
                            cell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER
                            cell.Padding = 4
                            dataTable.AddCell(cell)
                        End If
                    Next
                Next

                doc.Add(dataTable)
                doc.Close()
                MsgBox("تم تصدير ملف الـ PDF بنجاح تام بدون تعارض مراجع!", MsgBoxStyle.Information, "تم التصدير")
            End If
        Catch ex As Exception
            MsgBox("خطأ في تصدير الـ PDF: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Try
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel Files (*.xls)|*.xls"
            saveFileDialog.FileName = "تقرير_معهد_الخنساء_" & DateTime.Now.ToString("yyyyMMdd") & ".xls"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Using sw As New System.IO.StreamWriter(saveFileDialog.FileName, False, System.Text.Encoding.Unicode)
                    ' ترويسة ملف الإكسل (أول 3 أسطر)
                    sw.WriteLine("الجمهورية اليمنية")
                    sw.WriteLine("وزارة التعليم الفني والتدريب المهني")
                    sw.WriteLine("معهد الخنساء للتدريب الفني - تقرير بيانات الطالبات")
                    sw.WriteLine("تاريخ التصدير: " & DateTime.Now.ToString("yyyy/MM/dd"))
                    sw.WriteLine("") ' سطر فارغ للفصل

                    ' كتابة رؤوس الأعمدة
                    Dim headers As New List(Of String)
                    For Each column As DataGridViewColumn In DataGridView2.Columns
                        If column.Visible Then headers.Add(column.HeaderText)
                    Next
                    sw.WriteLine(String.Join(vbTab, headers))

                    ' كتابة البيانات
                    For Each row As DataGridViewRow In DataGridView2.Rows
                        If row.IsNewRow Then Continue For
                        Dim cells As New List(Of String)
                        For Each column As DataGridViewColumn In DataGridView2.Columns
                            If column.Visible Then
                                cells.Add(If(row.Cells(column.Index).Value IsNot Nothing, row.Cells(column.Index).Value.ToString(), ""))
                            End If
                        Next
                        sw.WriteLine(String.Join(vbTab, cells))
                    Next
                End Using
                MsgBox("تم تصدير ملف الإكسل مع الترويسة بنجاح!", MsgBoxStyle.Information)
            End If
        Catch ex As Exception
            MsgBox("خطأ في تصدير إكسل: " & ex.Message)
        End Try
    End Sub

End Class