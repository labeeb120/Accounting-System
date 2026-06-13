<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Expenses
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Expenses))
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TxAmount = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TxExpense_ID = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CbExpenseTypeID = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.DtpExpense_Date = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.TxNotes = New System.Windows.Forms.TextBox()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.Expense_ID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Expense_Date = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TypeName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Amount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.notes = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ExpenseTypeID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.بحثToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem6 = New System.Windows.Forms.ToolStripMenuItem()
        Me.حفظToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem7 = New System.Windows.Forms.ToolStripMenuItem()
        Me.الاعداداتToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem8 = New System.Windows.Forms.ToolStripMenuItem()
        Me.حذفToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.صرفToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.خروجToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ملفToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem4 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem5 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ملفToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem9 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Guna2GradientButton2 = New Guna.UI2.WinForms.Guna2GradientButton()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.PictureBox6 = New System.Windows.Forms.PictureBox()
        Me.Button13 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.BtnDel = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.BtnAdd = New System.Windows.Forms.Button()
        Me.BtnSave = New System.Windows.Forms.Button()
        Me.BtnEdt = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arabic11 BT", 13.8!, System.Drawing.FontStyle.Bold)
        Me.Label9.ForeColor = System.Drawing.Color.Purple
        Me.Label9.Location = New System.Drawing.Point(385, 76)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(116, 39)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "نوع الصرفيات"
        '
        'TxAmount
        '
        Me.TxAmount.Font = New System.Drawing.Font("Tahoma", 8.0!)
        Me.TxAmount.ForeColor = System.Drawing.Color.Purple
        Me.TxAmount.Location = New System.Drawing.Point(172, 17)
        Me.TxAmount.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxAmount.Multiline = True
        Me.TxAmount.Name = "TxAmount"
        Me.TxAmount.Size = New System.Drawing.Size(207, 43)
        Me.TxAmount.TabIndex = 7
        Me.TxAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arabic11 BT", 13.8!, System.Drawing.FontStyle.Bold)
        Me.Label4.ForeColor = System.Drawing.Color.Purple
        Me.Label4.Location = New System.Drawing.Point(395, 25)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(106, 39)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "مبلغ الصرف"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arabic11 BT", 13.8!, System.Drawing.FontStyle.Bold)
        Me.Label2.ForeColor = System.Drawing.Color.Purple
        Me.Label2.Location = New System.Drawing.Point(768, 30)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(115, 39)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "رقم الصرفيات"
        '
        'TxExpense_ID
        '
        Me.TxExpense_ID.Font = New System.Drawing.Font("Tahoma", 8.0!)
        Me.TxExpense_ID.ForeColor = System.Drawing.Color.Purple
        Me.TxExpense_ID.Location = New System.Drawing.Point(559, 22)
        Me.TxExpense_ID.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxExpense_ID.Multiline = True
        Me.TxExpense_ID.Name = "TxExpense_ID"
        Me.TxExpense_ID.ReadOnly = True
        Me.TxExpense_ID.Size = New System.Drawing.Size(201, 44)
        Me.TxExpense_ID.TabIndex = 1
        Me.TxExpense_ID.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arabic11 BT", 13.8!, System.Drawing.FontStyle.Bold)
        Me.Label1.ForeColor = System.Drawing.Color.Purple
        Me.Label1.Location = New System.Drawing.Point(766, 69)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(113, 39)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "تاريخ الصرف"
        '
        'CbExpenseTypeID
        '
        Me.CbExpenseTypeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold)
        Me.CbExpenseTypeID.ForeColor = System.Drawing.Color.Purple
        Me.CbExpenseTypeID.FormattingEnabled = True
        Me.CbExpenseTypeID.Location = New System.Drawing.Point(172, 72)
        Me.CbExpenseTypeID.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CbExpenseTypeID.Name = "CbExpenseTypeID"
        Me.CbExpenseTypeID.Size = New System.Drawing.Size(207, 37)
        Me.CbExpenseTypeID.TabIndex = 16
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Simple Outline Pat", 24.0!, System.Drawing.FontStyle.Bold)
        Me.Label13.ForeColor = System.Drawing.Color.Purple
        Me.Label13.Location = New System.Drawing.Point(432, 55)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(367, 67)
        Me.Label13.TabIndex = 52
        Me.Label13.Text = "قائمة صرفيات المعهد"
        '
        'DtpExpense_Date
        '
        Me.DtpExpense_Date.Location = New System.Drawing.Point(559, 80)
        Me.DtpExpense_Date.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DtpExpense_Date.Name = "DtpExpense_Date"
        Me.DtpExpense_Date.Size = New System.Drawing.Size(201, 24)
        Me.DtpExpense_Date.TabIndex = 55
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.ForeColor = System.Drawing.Color.Purple
        Me.GroupBox1.Location = New System.Drawing.Point(0, 114)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(114, 158)
        Me.GroupBox1.TabIndex = 83
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "تصدير"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Thistle
        Me.GroupBox2.Controls.Add(Me.PictureBox4)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.TxNotes)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.GroupBox1)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.TxExpense_ID)
        Me.GroupBox2.Controls.Add(Me.DtpExpense_Date)
        Me.GroupBox2.Controls.Add(Me.CbExpenseTypeID)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.TxAmount)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.ForeColor = System.Drawing.Color.Purple
        Me.GroupBox2.Location = New System.Drawing.Point(178, 146)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(939, 258)
        Me.GroupBox2.TabIndex = 87
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "المصروفات"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Arabic11 BT", 13.8!, System.Drawing.FontStyle.Bold)
        Me.Label12.ForeColor = System.Drawing.Color.Purple
        Me.Label12.Location = New System.Drawing.Point(790, 140)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(91, 39)
        Me.Label12.TabIndex = 88
        Me.Label12.Text = "الملاحضات"
        '
        'TxNotes
        '
        Me.TxNotes.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold)
        Me.TxNotes.ForeColor = System.Drawing.Color.Purple
        Me.TxNotes.Location = New System.Drawing.Point(172, 128)
        Me.TxNotes.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxNotes.Multiline = True
        Me.TxNotes.Name = "TxNotes"
        Me.TxNotes.Size = New System.Drawing.Size(601, 55)
        Me.TxNotes.TabIndex = 87
        Me.TxNotes.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.AllowUserToOrderColumns = True
        Me.DataGridView2.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Expense_ID, Me.Expense_Date, Me.TypeName, Me.Amount, Me.notes, Me.ExpenseTypeID})
        Me.DataGridView2.GridColor = System.Drawing.Color.White
        Me.DataGridView2.Location = New System.Drawing.Point(184, 403)
        Me.DataGridView2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.DataGridView2.RowHeadersWidth = 51
        Me.DataGridView2.RowTemplate.Height = 29
        Me.DataGridView2.Size = New System.Drawing.Size(709, 340)
        Me.DataGridView2.TabIndex = 86
        '
        'Expense_ID
        '
        Me.Expense_ID.DataPropertyName = "Expense_ID"
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Tahoma", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Expense_ID.DefaultCellStyle = DataGridViewCellStyle1
        Me.Expense_ID.FillWeight = 150.0!
        Me.Expense_ID.HeaderText = "رقم المصروف"
        Me.Expense_ID.MinimumWidth = 6
        Me.Expense_ID.Name = "Expense_ID"
        Me.Expense_ID.ReadOnly = True
        Me.Expense_ID.Width = 125
        '
        'Expense_Date
        '
        Me.Expense_Date.DataPropertyName = "Expense_Date"
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Tahoma", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Expense_Date.DefaultCellStyle = DataGridViewCellStyle2
        Me.Expense_Date.HeaderText = "التاريخ"
        Me.Expense_Date.MinimumWidth = 6
        Me.Expense_Date.Name = "Expense_Date"
        Me.Expense_Date.ReadOnly = True
        Me.Expense_Date.Width = 80
        '
        'TypeName
        '
        Me.TypeName.DataPropertyName = "TypeName"
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Tahoma", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TypeName.DefaultCellStyle = DataGridViewCellStyle3
        Me.TypeName.FillWeight = 150.0!
        Me.TypeName.HeaderText = "نوع المصروف"
        Me.TypeName.MinimumWidth = 6
        Me.TypeName.Name = "TypeName"
        Me.TypeName.ReadOnly = True
        Me.TypeName.Width = 150
        '
        'Amount
        '
        Me.Amount.DataPropertyName = "Amount"
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Tahoma", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Amount.DefaultCellStyle = DataGridViewCellStyle4
        Me.Amount.HeaderText = "المبلغ"
        Me.Amount.MinimumWidth = 6
        Me.Amount.Name = "Amount"
        Me.Amount.ReadOnly = True
        Me.Amount.Width = 120
        '
        'notes
        '
        Me.notes.DataPropertyName = "notes"
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Tahoma", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.notes.DefaultCellStyle = DataGridViewCellStyle5
        Me.notes.HeaderText = "ملاحظات"
        Me.notes.MinimumWidth = 6
        Me.notes.Name = "notes"
        Me.notes.ReadOnly = True
        Me.notes.Width = 300
        '
        'ExpenseTypeID
        '
        Me.ExpenseTypeID.DataPropertyName = "ExpenseTypeID"
        Me.ExpenseTypeID.HeaderText = "رقم النوع"
        Me.ExpenseTypeID.MinimumWidth = 6
        Me.ExpenseTypeID.Name = "ExpenseTypeID"
        Me.ExpenseTypeID.Visible = False
        Me.ExpenseTypeID.Width = 125
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.Thistle
        Me.MenuStrip1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold)
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.بحثToolStripMenuItem, Me.ToolStripMenuItem1, Me.ToolStripMenuItem3, Me.ToolStripMenuItem6, Me.حفظToolStripMenuItem, Me.ToolStripMenuItem7, Me.الاعداداتToolStripMenuItem1, Me.ToolStripMenuItem8, Me.حذفToolStripMenuItem, Me.خروجToolStripMenuItem, Me.ToolStripMenuItem2, Me.ToolStripMenuItem4, Me.ToolStripMenuItem5, Me.ToolStripMenuItem9})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(5, 2, 0, 2)
        Me.MenuStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.MenuStrip1.Size = New System.Drawing.Size(1244, 39)
        Me.MenuStrip1.TabIndex = 88
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'بحثToolStripMenuItem
        '
        Me.بحثToolStripMenuItem.Font = New System.Drawing.Font("Arabic11 BT", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.بحثToolStripMenuItem.ForeColor = System.Drawing.Color.Purple
        Me.بحثToolStripMenuItem.Name = "بحثToolStripMenuItem"
        Me.بحثToolStripMenuItem.Size = New System.Drawing.Size(110, 35)
        Me.بحثToolStripMenuItem.Text = "الوجهة الرئسيه"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Font = New System.Drawing.Font("Arabic11 BT", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.ToolStripMenuItem1.ForeColor = System.Drawing.Color.Purple
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(108, 35)
        Me.ToolStripMenuItem1.Text = "وجهة المستخدم"
        '
        'ToolStripMenuItem3
        '
        Me.ToolStripMenuItem3.Font = New System.Drawing.Font("Arabic11 BT", 10.8!, System.Drawing.FontStyle.Bold)
        Me.ToolStripMenuItem3.ForeColor = System.Drawing.Color.Purple
        Me.ToolStripMenuItem3.Name = "ToolStripMenuItem3"
        Me.ToolStripMenuItem3.Size = New System.Drawing.Size(101, 35)
        Me.ToolStripMenuItem3.Text = "لوحة التحكم"
        '
        'ToolStripMenuItem6
        '
        Me.ToolStripMenuItem6.Font = New System.Drawing.Font("Arabic11 BT", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.ToolStripMenuItem6.ForeColor = System.Drawing.Color.Purple
        Me.ToolStripMenuItem6.Name = "ToolStripMenuItem6"
        Me.ToolStripMenuItem6.Size = New System.Drawing.Size(98, 35)
        Me.ToolStripMenuItem6.Text = "واجهة التنقل"
        '
        'حفظToolStripMenuItem
        '
        Me.حفظToolStripMenuItem.Font = New System.Drawing.Font("Arabic11 BT", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.حفظToolStripMenuItem.ForeColor = System.Drawing.Color.Purple
        Me.حفظToolStripMenuItem.Name = "حفظToolStripMenuItem"
        Me.حفظToolStripMenuItem.Size = New System.Drawing.Size(112, 35)
        Me.حفظToolStripMenuItem.Text = "واجهة الطالبات"
        '
        'ToolStripMenuItem7
        '
        Me.ToolStripMenuItem7.Font = New System.Drawing.Font("Arabic11 BT", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.ToolStripMenuItem7.ForeColor = System.Drawing.Color.Purple
        Me.ToolStripMenuItem7.Name = "ToolStripMenuItem7"
        Me.ToolStripMenuItem7.Size = New System.Drawing.Size(126, 35)
        Me.ToolStripMenuItem7.Text = "واجهة الدفع المالي"
        '
        'الاعداداتToolStripMenuItem1
        '
        Me.الاعداداتToolStripMenuItem1.Font = New System.Drawing.Font("Arabic11 BT", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.الاعداداتToolStripMenuItem1.ForeColor = System.Drawing.Color.Purple
        Me.الاعداداتToolStripMenuItem1.Name = "الاعداداتToolStripMenuItem1"
        Me.الاعداداتToolStripMenuItem1.Size = New System.Drawing.Size(109, 35)
        Me.الاعداداتToolStripMenuItem1.Text = "واجهة الموظفين"
        '
        'ToolStripMenuItem8
        '
        Me.ToolStripMenuItem8.Font = New System.Drawing.Font("Arabic11 BT", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.ToolStripMenuItem8.ForeColor = System.Drawing.Color.Purple
        Me.ToolStripMenuItem8.Name = "ToolStripMenuItem8"
        Me.ToolStripMenuItem8.Size = New System.Drawing.Size(114, 35)
        Me.ToolStripMenuItem8.Text = "واجهة الداعمين"
        '
        'حذفToolStripMenuItem
        '
        Me.حذفToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.صرفToolStripMenuItem})
        Me.حذفToolStripMenuItem.Font = New System.Drawing.Font("Arabic11 BT", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.حذفToolStripMenuItem.ForeColor = System.Drawing.Color.Purple
        Me.حذفToolStripMenuItem.Name = "حذفToolStripMenuItem"
        Me.حذفToolStripMenuItem.Size = New System.Drawing.Size(66, 35)
        Me.حذفToolStripMenuItem.Text = "سندات"
        '
        'صرفToolStripMenuItem
        '
        Me.صرفToolStripMenuItem.Name = "صرفToolStripMenuItem"
        Me.صرفToolStripMenuItem.Size = New System.Drawing.Size(134, 36)
        Me.صرفToolStripMenuItem.Text = "صرف"
        '
        'خروجToolStripMenuItem
        '
        Me.خروجToolStripMenuItem.Font = New System.Drawing.Font("Arabic11 BT", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.خروجToolStripMenuItem.ForeColor = System.Drawing.Color.Purple
        Me.خروجToolStripMenuItem.Name = "خروجToolStripMenuItem"
        Me.خروجToolStripMenuItem.Size = New System.Drawing.Size(60, 35)
        Me.خروجToolStripMenuItem.Text = "خروج"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ملفToolStripMenuItem})
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(14, 35)
        '
        'ملفToolStripMenuItem
        '
        Me.ملفToolStripMenuItem.Name = "ملفToolStripMenuItem"
        Me.ملفToolStripMenuItem.Size = New System.Drawing.Size(135, 34)
        Me.ملفToolStripMenuItem.Text = "ملف"
        '
        'ToolStripMenuItem4
        '
        Me.ToolStripMenuItem4.Name = "ToolStripMenuItem4"
        Me.ToolStripMenuItem4.Size = New System.Drawing.Size(14, 35)
        '
        'ToolStripMenuItem5
        '
        Me.ToolStripMenuItem5.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ملفToolStripMenuItem1})
        Me.ToolStripMenuItem5.Name = "ToolStripMenuItem5"
        Me.ToolStripMenuItem5.Size = New System.Drawing.Size(14, 35)
        '
        'ملفToolStripMenuItem1
        '
        Me.ملفToolStripMenuItem1.Name = "ملفToolStripMenuItem1"
        Me.ملفToolStripMenuItem1.Size = New System.Drawing.Size(135, 34)
        Me.ملفToolStripMenuItem1.Text = "ملف"
        '
        'ToolStripMenuItem9
        '
        Me.ToolStripMenuItem9.Name = "ToolStripMenuItem9"
        Me.ToolStripMenuItem9.Size = New System.Drawing.Size(14, 35)
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.BackColor = System.Drawing.Color.White
        Me.Label17.Font = New System.Drawing.Font("Arabic11 BT", 13.8!, System.Drawing.FontStyle.Bold)
        Me.Label17.ForeColor = System.Drawing.Color.Purple
        Me.Label17.Location = New System.Drawing.Point(64, 55)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(65, 39)
        Me.Label17.TabIndex = 93
        Me.Label17.Text = "التاريخ"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.Color.White
        Me.Label18.Font = New System.Drawing.Font("Arabic11 BT", 13.8!, System.Drawing.FontStyle.Bold)
        Me.Label18.ForeColor = System.Drawing.Color.Purple
        Me.Label18.Location = New System.Drawing.Point(64, 99)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(67, 39)
        Me.Label18.TabIndex = 92
        Me.Label18.Text = "الوقت "
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'Guna2GradientButton2
        '
        Me.Guna2GradientButton2.BorderRadius = 10
        Me.Guna2GradientButton2.DisabledState.BorderColor = System.Drawing.Color.DarkGray
        Me.Guna2GradientButton2.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray
        Me.Guna2GradientButton2.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer))
        Me.Guna2GradientButton2.DisabledState.FillColor2 = System.Drawing.Color.FromArgb(CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer))
        Me.Guna2GradientButton2.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer))
        Me.Guna2GradientButton2.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Guna2GradientButton2.ForeColor = System.Drawing.Color.White
        Me.Guna2GradientButton2.Location = New System.Drawing.Point(893, 581)
        Me.Guna2GradientButton2.Name = "Guna2GradientButton2"
        Me.Guna2GradientButton2.Size = New System.Drawing.Size(224, 45)
        Me.Guna2GradientButton2.TabIndex = 97
        Me.Guna2GradientButton2.Text = "تقرير المصروفات"
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Font = New System.Drawing.Font("Tahoma", 8.0!)
        Me.TextBox1.ForeColor = System.Drawing.Color.Purple
        Me.TextBox1.Location = New System.Drawing.Point(893, 690)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(224, 53)
        Me.TextBox1.TabIndex = 90
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Timer2
        '
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 20
        '
        'PictureBox6
        '
        Me.PictureBox6.Image = Global.Accounting_System.My.Resources.Resources.search_24px
        Me.PictureBox6.Location = New System.Drawing.Point(893, 702)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(35, 35)
        Me.PictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox6.TabIndex = 104
        Me.PictureBox6.TabStop = False
        '
        'Button13
        '
        Me.Button13.BackColor = System.Drawing.Color.Purple
        Me.Button13.ForeColor = System.Drawing.Color.White
        Me.Button13.Image = CType(resources.GetObject("Button13.Image"), System.Drawing.Image)
        Me.Button13.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button13.Location = New System.Drawing.Point(1013, 403)
        Me.Button13.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(106, 54)
        Me.Button13.TabIndex = 88
        Me.Button13.Text = "تحديث"
        Me.Button13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button13.UseVisualStyleBackColor = False
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.Purple
        Me.Button6.ForeColor = System.Drawing.Color.White
        Me.Button6.Image = CType(resources.GetObject("Button6.Image"), System.Drawing.Image)
        Me.Button6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button6.Location = New System.Drawing.Point(1013, 520)
        Me.Button6.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(104, 56)
        Me.Button6.TabIndex = 81
        Me.Button6.Text = "الغاء"
        Me.Button6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button6.UseVisualStyleBackColor = False
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(745, 134)
        Me.PictureBox4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(44, 43)
        Me.PictureBox4.TabIndex = 89
        Me.PictureBox4.TabStop = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Purple
        Me.Button2.ForeColor = System.Drawing.Color.White
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(0, 89)
        Me.Button2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(108, 55)
        Me.Button2.TabIndex = 28
        Me.Button2.Text = "اكسل"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.Purple
        Me.Button3.ForeColor = System.Drawing.Color.White
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(0, 30)
        Me.Button3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(108, 55)
        Me.Button3.TabIndex = 31
        Me.Button3.Text = "PDF"
        Me.Button3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button11
        '
        Me.Button11.BackColor = System.Drawing.Color.Purple
        Me.Button11.ForeColor = System.Drawing.Color.White
        Me.Button11.Image = CType(resources.GetObject("Button11.Image"), System.Drawing.Image)
        Me.Button11.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button11.Location = New System.Drawing.Point(893, 631)
        Me.Button11.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(226, 55)
        Me.Button11.TabIndex = 84
        Me.Button11.Text = "طباعة سند "
        Me.Button11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button11.UseVisualStyleBackColor = False
        '
        'BtnDel
        '
        Me.BtnDel.BackColor = System.Drawing.Color.Purple
        Me.BtnDel.Enabled = False
        Me.BtnDel.ForeColor = System.Drawing.Color.White
        Me.BtnDel.Image = CType(resources.GetObject("BtnDel.Image"), System.Drawing.Image)
        Me.BtnDel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDel.Location = New System.Drawing.Point(1015, 460)
        Me.BtnDel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnDel.Name = "BtnDel"
        Me.BtnDel.Size = New System.Drawing.Size(104, 56)
        Me.BtnDel.TabIndex = 79
        Me.BtnDel.Text = "حذف"
        Me.BtnDel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnDel.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(1107, 41)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(137, 142)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 75
        Me.PictureBox1.TabStop = False
        '
        'BtnAdd
        '
        Me.BtnAdd.BackColor = System.Drawing.Color.Purple
        Me.BtnAdd.ForeColor = System.Drawing.Color.White
        Me.BtnAdd.Image = CType(resources.GetObject("BtnAdd.Image"), System.Drawing.Image)
        Me.BtnAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAdd.Location = New System.Drawing.Point(893, 462)
        Me.BtnAdd.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.Size = New System.Drawing.Size(104, 56)
        Me.BtnAdd.TabIndex = 82
        Me.BtnAdd.Text = "اضافه"
        Me.BtnAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnAdd.UseVisualStyleBackColor = False
        '
        'BtnSave
        '
        Me.BtnSave.BackColor = System.Drawing.Color.Purple
        Me.BtnSave.Enabled = False
        Me.BtnSave.ForeColor = System.Drawing.Color.White
        Me.BtnSave.Image = CType(resources.GetObject("BtnSave.Image"), System.Drawing.Image)
        Me.BtnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSave.Location = New System.Drawing.Point(893, 403)
        Me.BtnSave.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(104, 56)
        Me.BtnSave.TabIndex = 76
        Me.BtnSave.Text = "حفظ"
        Me.BtnSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnSave.UseVisualStyleBackColor = False
        '
        'BtnEdt
        '
        Me.BtnEdt.BackColor = System.Drawing.Color.Purple
        Me.BtnEdt.Enabled = False
        Me.BtnEdt.ForeColor = System.Drawing.Color.White
        Me.BtnEdt.Image = CType(resources.GetObject("BtnEdt.Image"), System.Drawing.Image)
        Me.BtnEdt.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEdt.Location = New System.Drawing.Point(893, 521)
        Me.BtnEdt.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnEdt.Name = "BtnEdt"
        Me.BtnEdt.Size = New System.Drawing.Size(104, 56)
        Me.BtnEdt.TabIndex = 77
        Me.BtnEdt.Text = "تعديل"
        Me.BtnEdt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnEdt.UseVisualStyleBackColor = False
        '
        'Expenses
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1244, 744)
        Me.Controls.Add(Me.PictureBox6)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Guna2GradientButton2)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.Button13)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Button11)
        Me.Controls.Add(Me.BtnDel)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.BtnAdd)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.BtnSave)
        Me.Controls.Add(Me.BtnEdt)
        Me.ForeColor = System.Drawing.Color.Purple
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "Expenses"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Expenses"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label9 As Label
    Friend WithEvents TxAmount As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents TxExpense_ID As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents CbExpenseTypeID As ComboBox
    Friend WithEvents Label13 As Label
    Friend WithEvents DtpExpense_Date As DateTimePicker
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Button11 As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents BtnAdd As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents BtnDel As Button
    Friend WithEvents BtnEdt As Button
    Friend WithEvents BtnSave As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents PictureBox4 As PictureBox
    Friend WithEvents Label12 As Label
    Friend WithEvents TxNotes As TextBox
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents Button13 As Button
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents حفظToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents حذفToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents صرفToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents بحثToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents خروجToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As ToolStripMenuItem
    Friend WithEvents ملفToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents الاعداداتToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem3 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem4 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem5 As ToolStripMenuItem
    Friend WithEvents ملفToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem6 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem7 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem8 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem9 As ToolStripMenuItem
    Friend WithEvents Label17 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents Expense_ID As DataGridViewTextBoxColumn
    Friend WithEvents Expense_Date As DataGridViewTextBoxColumn
    Friend WithEvents TypeName As DataGridViewTextBoxColumn
    Friend WithEvents Amount As DataGridViewTextBoxColumn
    Friend WithEvents notes As DataGridViewTextBoxColumn
    Friend WithEvents ExpenseTypeID As DataGridViewTextBoxColumn
    Friend WithEvents Timer1 As Timer
    Friend WithEvents Guna2GradientButton2 As Guna.UI2.WinForms.Guna2GradientButton
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents PictureBox6 As PictureBox
    Friend WithEvents Timer2 As Timer
End Class
