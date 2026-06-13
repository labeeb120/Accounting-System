Imports System.Data.OleDb
Imports System.Windows.Forms.DataVisualization.Charting
Public Class Dashboard
    Private Sub EnableManager()

    End Sub
    Private Sub EnableSecretary()

    End Sub
    Sub LoadNetProfit()

        Try

            OpenConnection()

            Dim Fees As Decimal = 0
            Dim Support As Decimal = 0
            Dim Expenses As Decimal = 0

            Dim cmd1 As New OleDbCommand(
        "SELECT SUM(paid_Amount) FROM payment",
        conn)

            If Not IsDBNull(cmd1.ExecuteScalar()) Then

                Fees =
            Convert.ToDecimal(cmd1.ExecuteScalar())

            End If

            Dim cmd2 As New OleDbCommand(
        "SELECT SUM(Amount) FROM pportPayments",
        conn)

            If Not IsDBNull(cmd2.ExecuteScalar()) Then

                Support =
            Convert.ToDecimal(cmd2.ExecuteScalar())

            End If

            Dim cmd3 As New OleDbCommand(
        "SELECT SUM(Amount) FROM Expenses",
        conn)

            If Not IsDBNull(cmd3.ExecuteScalar()) Then

                Expenses =
            Convert.ToDecimal(cmd3.ExecuteScalar())

            End If

            Label20.Text =
        (Fees + Support - Expenses).ToString("N0")

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub
    Private Sub LoadDashboard()

        OpenConnection()

        Dim cmd As New OleDbCommand("SELECT COUNT(*)  FROM Support", conn)



        Label11.Text =
    cmd.ExecuteScalar().ToString()

        CloseConnection()

    End Sub

    Sub LoadStudentsCount()

        Try

            OpenConnection()

            Dim cmd As New OleDbCommand("SELECT COUNT(*) FROM Students", conn)

            Label2.Text = cmd.ExecuteScalar().ToString()

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub
    Public Sub LoadStatistics()
        Try

            OpenConnection()

            ' عدد الطالبات
            Dim cmdStudents As New OleDbCommand(
        "SELECT COUNT(*) FROM Students",
        conn)

            Label2.Text =
        cmdStudents.ExecuteScalar().ToString()

            ' عدد الموظفين
            Dim cmdEmployees As New OleDbCommand(
        "SELECT COUNT(*) FROM Employee",
        conn)

            Label13.Text =
        cmdEmployees.ExecuteScalar().ToString()

            ' عدد الأقسام
            Dim cmdDepartments As New OleDbCommand(
        "SELECT COUNT(*) FROM Departments",
        conn)

            Label7.Text =
        cmdDepartments.ExecuteScalar().ToString()

            ' عدد الداعمين
            Dim cmdSupporters As New OleDbCommand(
        "SELECT COUNT(*) FROM Support",
        conn)

            Label11.Text =
        cmdSupporters.ExecuteScalar().ToString()

            ' إجمالي الدعم
            Dim SupportTotal As Decimal = 0

            Dim cmdSupport As New OleDbCommand(
        "SELECT SUM(Amount) FROM  SupportPayments",
        conn)

            If Not IsDBNull(cmdSupport.ExecuteScalar()) Then

                SupportTotal =
            Convert.ToDecimal(cmdSupport.ExecuteScalar())

            End If

            Label9.Text =
        SupportTotal.ToString("N0")

            ' إجمالي الرسوم
            Dim FeesTotal As Decimal = 0

            Dim cmdFees As New OleDbCommand(
        "SELECT SUM(paid_Amount) FROM payment",
        conn)

            If Not IsDBNull(cmdFees.ExecuteScalar()) Then

                FeesTotal =
            Convert.ToDecimal(cmdFees.ExecuteScalar())

            End If

            ' إجمالي الإيرادات
            Label3.Text =
        (FeesTotal + SupportTotal).ToString("N0")

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub
    Sub LoadEmployeesCount()

        Try

            OpenConnection()

            Dim cmd As New OleDbCommand("SELECT COUNT(*) FROM Employee", conn)

            Label13.Text = cmd.ExecuteScalar().ToString()

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub
    Sub LoadRevenue()

        Try

            OpenConnection()

            Dim studentRevenue As Decimal = 0
            Dim supportRevenue As Decimal = 0

            Dim cmd1 As New OleDbCommand("SELECT SUM(paid_Amount) FROM payment", conn)

            If Not IsDBNull(cmd1.ExecuteScalar()) Then
                studentRevenue = Convert.ToDecimal(cmd1.ExecuteScalar())
            End If

            Dim cmd2 As New OleDbCommand("SELECT SUM(Amount) FROM Support", conn)

            If Not IsDBNull(cmd2.ExecuteScalar()) Then
                supportRevenue = Convert.ToDecimal(cmd2.ExecuteScalar())
            End If

            Label3.Text = (studentRevenue + supportRevenue).ToString("N0")

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Sub LoadExpenses()

        Try

            OpenConnection()

            Dim cmd As New OleDbCommand("SELECT SUM(Amount) FROM Expenses", conn)

            If IsDBNull(cmd.ExecuteScalar()) Then

                Label5.Text = "0"

            Else

                Label5.Text = Convert.ToDecimal(cmd.ExecuteScalar()).ToString("N0")

            End If

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub
    Sub LoadStudentsChart()

        Try

            Chart1.Series.Clear()

            OpenConnection()

            Dim sql As String =
            "SELECT Departments.DepartmentName, COUNT(Students.StudentID) AS TotalStudents " &
            "FROM Departments INNER JOIN Students " &
            "ON Departments.DepartmentID = Students.[Department ID] " &
            "GROUP BY Departments.DepartmentName"

            Dim cmd As New OleDbCommand(sql, conn)

            Dim dr As OleDbDataReader = cmd.ExecuteReader()

            Dim s As New Series()

            s.ChartType = SeriesChartType.Column

            s.IsValueShownAsLabel = True

            While dr.Read()

                s.Points.AddXY(dr("DepartmentName").ToString(),
                           Convert.ToInt32(dr("TotalStudents")))

            End While

            Chart1.Series.Add(s)

            Chart1.Titles.Clear()

            Chart1.Titles.Add("عدد الطالبات حسب القسم")

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub
    Sub LoadMoneyChart()

        Try

            Chart2.Series.Clear()

            OpenConnection()

            Dim Fees As Decimal = 0
            Dim Support As Decimal = 0
            Dim Expenses As Decimal = 0

            ' الرسوم الدراسية
            Dim cmd1 As New OleDbCommand(
        "SELECT SUM(paid_Amount) FROM payment",
        conn)

            If Not IsDBNull(cmd1.ExecuteScalar()) Then

                Fees =
            Convert.ToDecimal(cmd1.ExecuteScalar())

            End If

            ' الدعم
            Dim cmd2 As New OleDbCommand(
        "SELECT SUM(Amount) FROM SupportPayments",
        conn)

            If Not IsDBNull(cmd2.ExecuteScalar()) Then

                Support =
            Convert.ToDecimal(cmd2.ExecuteScalar())

            End If

            ' المصروفات
            Dim cmd3 As New OleDbCommand(
        "SELECT SUM(Amount) FROM Expenses",
        conn)

            If Not IsDBNull(cmd3.ExecuteScalar()) Then

                Expenses =
            Convert.ToDecimal(cmd3.ExecuteScalar())

            End If

            Dim TotalRevenue As Decimal =
        Fees + Support

            Dim s As New Series()

            s.ChartType = SeriesChartType.Column

            s.IsValueShownAsLabel = True

            s.Points.AddXY(
        "الإيرادات",
        TotalRevenue)

            s.Points.AddXY(
        "المصروفات",
        Expenses)

            Chart2.Series.Add(s)

            Chart2.Titles.Clear()

            Chart2.Titles.Add(
        "الإيرادات والمصروفات")

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub
    Sub LoadSupportChart()

        Try

            Chart4.Series.Clear()

            OpenConnection()

            Dim sql As String =
            "SELECT SupportTypeName, SUM(Amount) AS TotalSupport " &
            "FROM Support INNER JOIN SupportType " &
            "ON Support.SupportTypeID = SupportType.SupportTypeID " &
            "GROUP BY SupportTypeName"

            Dim cmd As New OleDbCommand(sql, conn)

            Dim dr As OleDbDataReader = cmd.ExecuteReader()

            Dim s As New Series()

            s.ChartType = SeriesChartType.Pie

            s.IsValueShownAsLabel = True

            While dr.Read()

                s.Points.AddXY(dr("SupportTypeName").ToString(),
                           Convert.ToDecimal(dr("TotalSupport")))

            End While

            Chart4.Series.Add(s)

            Chart4.Titles.Clear()

            Chart4.Titles.Add("أنواع الدعم")

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Sub LoadStatusChart()

        Try

            Chart3.Series.Clear()

            OpenConnection()

            Dim sql As String =
            "SELECT StudentStatus, COUNT(StudentID) AS TotalStatus FROM Students GROUP BY StudentStatus"

            Dim cmd As New OleDbCommand(sql, conn)

            Dim dr As OleDbDataReader = cmd.ExecuteReader()

            Dim s As New Series()

            s.ChartType = SeriesChartType.Pie

            s.IsValueShownAsLabel = True

            While dr.Read()

                s.Points.AddXY(dr("StudentStatus").ToString(),
                           Convert.ToInt32(dr("TotalStatus")))

            End While

            Chart3.Series.Add(s)

            Chart3.Titles.Clear()

            Chart3.Titles.Add("حالة الطالبات")

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Sub LoadDepartmentsCount()

        Try

            OpenConnection()

            Dim cmd As New OleDbCommand("SELECT COUNT(*) FROM Departments", conn)

            Label7.Text = cmd.ExecuteScalar().ToString()

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub
    Sub LoadSupportersCount()

        Try

            OpenConnection()

            Dim cmd As New OleDbCommand("SELECT COUNT(*) FROM Supporters", conn)

            Label11.Text = cmd.ExecuteScalar().ToString()

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub
    Sub LoadSupportTotal()

        Try

            OpenConnection()

            Dim cmd As New OleDbCommand("SELECT SUM(Amount) FROM Support", conn)

            If IsDBNull(cmd.ExecuteScalar()) Then

                Label9.Text = "0"

            Else

                Label9.Text =
                Convert.ToDecimal(cmd.ExecuteScalar()).ToString("N0")

            End If

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub





    Private Sub Chart1_Click(sender As Object, e As EventArgs) Handles Chart1.Click
        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Pie
    End Sub


    Private Sub Dashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadStatistics()

        LoadStudentsChart()
        LoadMoneyChart()
        LoadStatusChart()
        LoadSupportChart()
        LoadNetProfit()

        LoadExpenses()

        LoadStudentsChart()

        LoadMoneyChart()

        LoadStatusChart()

        LoadSupportersCount()
        LoadDepartmentsCount()
        LoadSupportChart()
        Label25.Text =
"المستخدم : " &
CurrentUserName

        Label23.Text =
        "الصلاحية : " &
        CurrentUserRole
        If CurrentUserRole = "مدير" Then

            EnableManager()

        Else

            EnableSecretary()

        End If
        Me.Width = 100
        Me.Height = 100
        Timer2.Start()

    End Sub



    Private Sub Guna2GradientButton2_Click(sender As Object, e As EventArgs)
        MINE.Show()
    End Sub

    Private Sub Guna2GradientButton1_Click(sender As Object, e As EventArgs)
        Application.Exit()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Label18.Text = TimeOfDay.ToString("hh:mm:ss tt")

        Label17.Text = Today.ToString("yyyy/MM/dd")



    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If Me.Width < 800 Then
            Me.Width += 30
            Me.Height += 20
        Else
            Timer2.Stop()
        End If

    End Sub
End Class