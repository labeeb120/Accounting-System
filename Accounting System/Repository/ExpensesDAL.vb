' ملف اسم ExpensesDAL.vb
Imports System.Data.OleDb

Public Class ExpensesDAL
    ' نص الاتصال ثابت هنا لسهولة الصيانة
    Private ReadOnly connString As New OleDbConnection(DatabaseModule.connString) 'String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\users.accdb"


    ' ميثود الإضافة
    Public Function Insert(ByVal exp As Expense) As Boolean
        Dim sql As String = "INSERT INTO Expenses (Expense_Date, ExpenseTypeID, Amount, notes, employee_id, UserID) " &
                            "VALUES (?, ?, ?, ?, ?, ?)"


        Using cmd As New OleDbCommand(sql, connString)
            Try
                ' نأخذ القيم من الكائن (exp) الذي مررناه للميثود
                'cmd.Parameters.AddWithValue("@d", exp.Expense_Date)
                cmd.Parameters.Add("@d", OleDbType.Date).Value = exp.Expense_Date
                cmd.Parameters.AddWithValue("@t", exp.ExpenseTypeID)
                cmd.Parameters.AddWithValue("@a", exp.Amount)
                cmd.Parameters.AddWithValue("@n", If(String.IsNullOrEmpty(exp.Notes), DBNull.Value, exp.Notes))
                cmd.Parameters.AddWithValue("@e", DBNull.Value) ' exp.Employee_ID)
                cmd.Parameters.AddWithValue("@u", DBNull.Value) 'exp.UserID)

                connString.Open()
                Return cmd.ExecuteNonQuery() > 0
            Catch ex As Exception
                Throw New Exception("خطأ في قاعدة البيانات: " & ex.Message)
            End Try
        End Using

    End Function


    ' ميثود التعديل (Update)
    Public Function Update(ByVal exp As Expense) As Boolean
        ' لاحظ أننا نعدل البيانات بناءً على Expense_ID
        Dim sql As String = "UPDATE Expenses SET Expense_Date=?, ExpenseTypeID=?, Amount=?, notes=?, employee_id=?, UserID=? " &
                            "WHERE Expense_ID=?"

        Using conn = connString 'As New OleDbConnection(connString)
            Using cmd As New OleDbCommand(sql, conn)
                Try
                    ' الترتيب هنا حرج جداً ويجب أن يطابق ترتيب علامات الاستفهام في جملة SQL
                    cmd.Parameters.Add("@d", OleDbType.Date).Value = exp.Expense_Date

                    'cmd.Parameters.AddWithValue("@d", exp.Expense_Date)
                    cmd.Parameters.AddWithValue("@t", exp.ExpenseTypeID)
                    cmd.Parameters.AddWithValue("@a", exp.Amount)
                    cmd.Parameters.AddWithValue("@n", If(String.IsNullOrEmpty(exp.Notes), DBNull.Value, exp.Notes))
                    cmd.Parameters.AddWithValue("@e", DBNull.Value)
                    cmd.Parameters.AddWithValue("@u", DBNull.Value)
                    ' البارامتر الأخير هو المعرف الموجود في شرط WHERE
                    cmd.Parameters.AddWithValue("@id", exp.Expense_ID)

                    conn.Open()

                    Return cmd.ExecuteNonQuery() > 0
                Catch ex As Exception
                    Throw New Exception("خطأ أثناء التعديل: " & ex.Message)
                End Try
            End Using
        End Using
    End Function

    ' ميثود الحذف (Delete)
    Public Function Delete(ByVal id As Integer) As Boolean
        Dim sql As String = "DELETE FROM Expenses WHERE Expense_ID=?"

        Using conn = connString
            Using cmd As New OleDbCommand(sql, conn)
                Try
                    cmd.Parameters.AddWithValue("@id", id)

                    conn.Open()
                    Return cmd.ExecuteNonQuery() > 0
                Catch ex As Exception
                    Throw New Exception("خطأ أثناء الحذف: " & ex.Message)
                End Try
            End Using
        End Using
    End Function

    ' ميثود استرجاع كافة بيانات الجدول
    Public Function GetAll() As List(Of Expense)
        Dim expensesList As New List(Of Expense)
        Dim sql As String = "SELECT Expense_ID, Expense_Date, e.ExpenseTypeID, Amount, notes, employee_id, UserID,ExpenseTypeNAME as TypeName FROM Expenses e left join Expensetype t on e.ExpenseTypeID = t.ExpenseTypeID  ORDER BY Expense_ID asc"

        ' نستخدم كائن الاتصال المعرف في الكلاس (الذي سميته أنت connString)
        Using cmd As New OleDbCommand(sql, connString)
            Try
                If connString.State = ConnectionState.Closed Then connString.Open()

                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim item As New Expense With {
                            .Expense_ID = Convert.ToInt32(reader("Expense_ID")),
                            .Expense_Date = Convert.ToDateTime(reader("Expense_Date")),
                            .ExpenseTypeID = Convert.ToInt32(reader("ExpenseTypeID")),
                            .Amount = Convert.ToDecimal(reader("Amount")),
                        .Notes = If(reader("notes") Is DBNull.Value, String.Empty, reader("notes").ToString()),
                            .Employee_ID = If(reader("employee_id") Is DBNull.Value, 0, Convert.ToInt32(reader("employee_id"))),
                            .UserID = If(reader("UserID") Is DBNull.Value, 0, Convert.ToInt32(reader("UserID"))),
                             .TypeName = If(reader("TypeName") Is DBNull.Value, "", reader("TypeName").ToString())
                        }
                        expensesList.Add(item)
                    End While
                End Using
            Catch ex As Exception
                Throw New Exception("فشل في استرجاع بيانات المصروفات: " & ex.Message)
            Finally
                ' التأكد من إغلاق الاتصال بعد الانتهاء
                If connString.State = ConnectionState.Open Then connString.Close()
            End Try
        End Using

        Return expensesList
    End Function
    ' يمكنك هنا إضافة ميثود الحذف والتعديل بنفس الطريقة
End Class
