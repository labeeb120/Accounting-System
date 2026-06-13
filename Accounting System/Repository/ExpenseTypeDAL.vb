Imports System.Data.OleDb

Public Class ExpenseTypeDAL
    ' مسار قاعدة بيانات اكسس (تأكد من تغيير المسار لمكان ملفك)
    Private ReadOnly ConnectionString As New OleDbConnection(DatabaseModule.connString) 'String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\YourPath\YourDatabase.accdb;Persist Security Info=False;"

    Public Function GetRecords() As List(Of ExpenseType)
        Dim list As New List(Of ExpenseType)()

        ' استخدام OleDbConnection بدلاً من OracleConnection
        Using localConn = ConnectionString
            Dim sql As String = "SELECT * FROM EXPENSETYPE"
            Dim cmd As New OleDbCommand(sql, localConn)

            Try
                localConn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        list.Add(New ExpenseType With {
                            .Id = Convert.ToInt32(reader("ExpenseTypeID")),
                            .TypeName = reader("ExpenseTypeNAME").ToString()
                        })
                    End While
                End Using
            Catch ex As Exception
                ' يفضل عرض رسالة الخطأ هنا للتأكد من مسار الملف أو الصلاحيات
                MsgBox("Error: " & ex.Message)
            End Try
        End Using

        Return list
    End Function
End Class
