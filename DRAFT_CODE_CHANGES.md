# مسودة مقترحة لتعديلات الأكواد (Draft Code Changes)
*ملاحظة: هذه مسودة برمجية توضح الهيكلية والمنطق المقترح، وليست للتطبيق المباشر قبل التأكد من مطابقتها لكامل بيئة النظام.*

## 1. إضافة متغيرات واجهة المستخدم (UI Controls)
يُقترح إضافة خيارين (RadioButtons) لتحديد نوع الدفع:
*   `rbInternal` (دفع داخلي - طالبات)
*   `rbExternal` (دفع خارجي - داعمين)

### منطق التبديل بين الأنواع (UI Toggling Logic)
```vb.net
' عند اختيار دفع خارجي
Private Sub rbExternal_CheckedChanged(sender As Object, e As EventArgs) Handles rbExternal.CheckedChanged
    If rbExternal.Checked Then
        ' تعطيل حقول الطالبة
        TextBox9.Enabled = False ' رقم الطالبة
        TextBox9.Text = GenerateExternalID() ' توليد رقم تلقائي للدافع

        ' تمكين حقول الداعم (بفرض إضافة هذه الحقول)
        txtDonorName.Enabled = True
        txtDonorPhone.Enabled = True

        ' تغيير نص مصدر الإيراد تلقائياً
        ComboBox4.Text = "دعم خارجي / داعمين"
    End If
End Sub

' عند اختيار دفع داخلي
Private Sub rbInternal_CheckedChanged(sender As Object, e As EventArgs) Handles rbInternal.CheckedChanged
    If rbInternal.Checked Then
        TextBox9.Enabled = True
        TextBox9.Clear()
        txtDonorName.Enabled = False
        txtDonorPhone.Enabled = False
    End If
End Sub
```

## 2. دالة توليد رقم الدافع الخارجي آلياً
```vb.net
Function GenerateExternalID() As String
    Dim nextID As Integer = 1
    Try
        OpenConnection()
        ' جلب أقصى رقم دفع مسجل يبدأ بترميز الخارجي
        Dim sql As String = "SELECT MAX(VAL(MID(StudentID, 5))) FROM payment WHERE StudentID LIKE 'EXT-%'"
        Dim cmd As New OleDbCommand(sql, conn)
        Dim result = cmd.ExecuteScalar()
        If Not IsDBNull(result) AndAlso result IsNot Nothing Then
            nextID = Convert.ToInt32(result) + 1
        End If
    Catch ex As Exception
        ' معالجة الخطأ
    Finally
        CloseConnection()
    End Try
    Return "EXT-" & nextID.ToString("D4") ' مثال: EXT-0001
End Function
```

## 3. تعديل منطق الحفظ (Modified Save Logic)
تعديل الكود في `Button2_Click` ليتعامل مع النوعين:

```vb.net
Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    Try
        ' التحقق من المدخلات بناءً على النوع
        If rbInternal.Checked AndAlso TextBox9.Text = "" Then
            MsgBox("أدخل رقم الطالبة أولاً")
            Exit Sub
        ElseIf rbExternal.Checked AndAlso txtDonorName.Text = "" Then
            MsgBox("أدخل اسم الداعم")
            Exit Sub
        End If

        OpenConnection()

        ' جملة SQL المعدلة (تم إضافة حقل Phone)
        Dim sql As String = "INSERT INTO payment (StudentID, StudentName, Phone, paid_Amount, payment_Date, payment_Method, RevenueTypeID, Notes) " & _
                            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)"

        Dim cmd As New OleDbCommand(sql, conn)

        ' تحديد القيم بناءً على الاختيار
        If rbInternal.Checked Then
            cmd.Parameters.AddWithValue("?", TextBox9.Text) ' رقم الطالبة
            cmd.Parameters.AddWithValue("?", TextBox7.Text) ' اسم الطالبة (من البحث)
            cmd.Parameters.AddWithValue("?", TextBox6.Text) ' هاتف الطالبة
        Else
            cmd.Parameters.AddWithValue("?", TextBox9.Text) ' رقم الدافع (المولد آلياً EXT-XXX)
            cmd.Parameters.AddWithValue("?", txtDonorName.Text) ' اسم الداعم المكتوب يدوياً
            cmd.Parameters.AddWithValue("?", txtDonorPhone.Text) ' هاتف الداعم المكتوب يدوياً
        End If

        cmd.Parameters.AddWithValue("?", Val(TextBox4.Text))
        cmd.Parameters.AddWithValue("?", DateTimePicker1.Value)
        cmd.Parameters.AddWithValue("?", ComboBox1.Text)
        cmd.Parameters.AddWithValue("?", ComboBox4.Text)
        cmd.Parameters.AddWithValue("?", TextBox8.Text)

        cmd.ExecuteNonQuery()

        MsgBox("تم الحفظ بنجاح")
        ' استكمال منطق تحديث الشبكة والإحصائيات...

    Catch ex As Exception
        MsgBox("خطأ أثناء الحفظ: " & ex.Message)
    Finally
        CloseConnection()
    End Try
End Sub
```

## 4. التغييرات الجوهرية المطلوبة في قاعدة البيانات
*   يجب التأكد أن حقل `StudentID` في جدول `payment` من نوع **نص (Short Text)** وليس رقماً، لكي يستوعب ترميز `EXT-0001`.
*   إذا كان `StudentID` مفتاحاً أجنبياً (Foreign Key) يفرض وجوده في جدول الطالبات، يجب إزالة هذا الارتباط الإجباري أو إنشاء سجل افتراضي في جدول الطالبات برقم 0 لاستخدامه لكل العمليات الخارجية.

---
**إعداد:** قسم التطوير البرمجي - خبير VB.NET.
