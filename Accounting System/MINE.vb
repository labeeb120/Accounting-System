Imports System.IO
Imports System.Data.OleDb
Public Class MINE
    Private IsCollapsed As Boolean = False
    Private TargetWidth As Integer
    Private ReadOnly نسخاحتياطيToolStripMenuItem As Object
    Dim DarkMode As Boolean = False
    Dim OriginalPanelColor As Color
    Public Property استعادةالنسخةالاحتياطيةToolStripMenuItem As Object

    Private Sub واجههالطلابToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub قبضToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        eemp.Show()
    End Sub

    Private Sub ملفToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ملفToolStripMenuItem.Click

    End Sub

    Private Sub حذفToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub MINE_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        OriginalPanelColor = Panel1.BackColor

        If CurrentUserRole <> "مدير" AndAlso
   CurrentUserRole <> "سكرتارية" Then

            MsgBox("غير مصرح لك بالدخول")

            Me.Close()

            Exit Sub

        End If

        Me.Left = Screen.PrimaryScreen.Bounds.Width

        Timer1.Start()

        ' If CurrentUserRole = "سكرتارية" Then

        'الاعداداتToolStripMenuItem.Enabled = False

        ' نسخاحتياطيToolStripMenuItem.Enabled = False

        'استعادةالنسخةالاحتياطيةToolStripMenuItem.Enabled = False

        ' End If
    End Sub

    Private Sub خروجToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles خروجToolStripMenuItem.Click

        If MsgBox(
    "هل تريد الخروج ؟",
    MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

            Application.Exit()

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Expenses.Show()
        Me.Close()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        studnt.Show()
        Me.Close()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)



    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ADD.Show()
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        mane.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        supports.Show()
        Me.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        user.Show()
        Me.Close()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dashboard.Show()
        Me.Close()
    End Sub

    Private Sub Guna2CirclePictureBox1_Click(sender As Object, e As EventArgs) Handles Guna2CirclePictureBox1.Click

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Me.Left > (Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2 Then
            Me.Left -= 20
        Else
            Timer1.Stop()
        End If

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        MsgBox(
   "نظام إدارة معهد الخنساء" &
   vbCrLf &
   "الإصدار 1.0" &
   vbCrLf &
   "تم تطوير النظام باستخدام VB.NET و Access",
   MsgBoxStyle.Information,
   "المساعدة")
    End Sub

    Private Sub بحثToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles بحثToolStripMenuItem.Click
        Dashboard.Show()
    End Sub

    Private Sub حفظToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles حفظToolStripMenuItem.Click

        Try

            Dim SaveDlg As New SaveFileDialog

            SaveDlg.Filter =
        "Backup Files|*.accdb"

            SaveDlg.FileName =
        "Backup_" &
        Date.Now.ToString("yyyyMMdd")

            If SaveDlg.ShowDialog = DialogResult.OK Then

                File.Copy(
            Application.StartupPath &
            "\users.accdb",
            SaveDlg.FileName,
            True)

                MsgBox(
            "تم إنشاء النسخة الاحتياطية بنجاح")

            End If

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub استعادهالنسخةالاحتياطيهToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles استعادهالنسخةالاحتياطيهToolStripMenuItem.Click
        Try

            Dim OpenDlg As New OpenFileDialog

            OpenDlg.Filter =
        "Backup Files|*.accdb"

            If OpenDlg.ShowDialog = DialogResult.OK Then

                File.Copy(
            OpenDlg.FileName,
            Application.StartupPath &
            "\users.accdb",
            True)

                MsgBox(
            "تمت استعادة النسخة بنجاح")

                Application.Restart()

            End If

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub الاعداداتToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles الاعداداتToolStripMenuItem1.Click

        MsgBox(
        "نظام إدارة معهد الخنساء" &
        vbCrLf &
        "الإصدار : 1.0" &
        vbCrLf &
        "قاعدة البيانات : Access" &
        vbCrLf &
        "لغة البرمجة : VB.NET",
        MsgBoxStyle.Information,
        "الإعدادات")
    End Sub
End Class