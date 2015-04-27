<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form1
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents chkDontSendMail As System.Windows.Forms.CheckBox
	Public WithEvents cmdSetDates As System.Windows.Forms.Button
	Public WithEvents txtDueDate As System.Windows.Forms.TextBox
	Public WithEvents txtCalendar As System.Windows.Forms.TextBox
    Public WithEvents cmdDueDate As System.Windows.Forms.Button
	Public WithEvents txtNotice As System.Windows.Forms.TextBox
	Public WithEvents cmdEmailNotices As System.Windows.Forms.Button
	Public WithEvents cmdEmailTest As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtNotice = New System.Windows.Forms.TextBox
        Me.cmdEmailTest = New System.Windows.Forms.Button
        Me.chkSMTPtest = New System.Windows.Forms.CheckBox
        Me.chkDontSendMail = New System.Windows.Forms.CheckBox
        Me.chkDontUpdateDatabase = New System.Windows.Forms.CheckBox
        Me.cmdSetDates = New System.Windows.Forms.Button
        Me.txtDueDate = New System.Windows.Forms.TextBox
        Me.txtCalendar = New System.Windows.Forms.TextBox
        Me.cmdDueDate = New System.Windows.Forms.Button
        Me.cmdEmailNotices = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkShowMessages = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'txtNotice
        '
        Me.txtNotice.AcceptsReturn = True
        Me.txtNotice.BackColor = System.Drawing.SystemColors.Window
        Me.txtNotice.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNotice.Font = New System.Drawing.Font("Tahoma", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNotice.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNotice.Location = New System.Drawing.Point(270, 180)
        Me.txtNotice.MaxLength = 0
        Me.txtNotice.Name = "txtNotice"
        Me.txtNotice.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNotice.Size = New System.Drawing.Size(131, 27)
        Me.txtNotice.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtNotice, "Enter date to run program for.")
        '
        'cmdEmailTest
        '
        Me.cmdEmailTest.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEmailTest.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdEmailTest.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEmailTest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEmailTest.Location = New System.Drawing.Point(40, 30)
        Me.cmdEmailTest.Name = "cmdEmailTest"
        Me.cmdEmailTest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdEmailTest.Size = New System.Drawing.Size(181, 41)
        Me.cmdEmailTest.TabIndex = 1
        Me.cmdEmailTest.Text = "Send Test Email to User"
        Me.ToolTip1.SetToolTip(Me.cmdEmailTest, "Send an email to the test user.")
        Me.cmdEmailTest.UseVisualStyleBackColor = False
        '
        'chkSMTPtest
        '
        Me.chkSMTPtest.BackColor = System.Drawing.SystemColors.Control
        Me.chkSMTPtest.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSMTPtest.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSMTPtest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSMTPtest.Location = New System.Drawing.Point(270, 258)
        Me.chkSMTPtest.Name = "chkSMTPtest"
        Me.chkSMTPtest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSMTPtest.Size = New System.Drawing.Size(213, 21)
        Me.chkSMTPtest.TabIndex = 8
        Me.chkSMTPtest.Text = "SMTP Test (use GMail)"
        Me.ToolTip1.SetToolTip(Me.chkSMTPtest, "Send all e-mails via Gordon's GMail SMTP server.")
        Me.chkSMTPtest.UseVisualStyleBackColor = False
        '
        'chkDontSendMail
        '
        Me.chkDontSendMail.BackColor = System.Drawing.SystemColors.Control
        Me.chkDontSendMail.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDontSendMail.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDontSendMail.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDontSendMail.Location = New System.Drawing.Point(270, 231)
        Me.chkDontSendMail.Name = "chkDontSendMail"
        Me.chkDontSendMail.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDontSendMail.Size = New System.Drawing.Size(213, 21)
        Me.chkDontSendMail.TabIndex = 7
        Me.chkDontSendMail.Text = "Don't Send Emails"
        Me.ToolTip1.SetToolTip(Me.chkDontSendMail, "Don't send emails.")
        Me.chkDontSendMail.UseVisualStyleBackColor = False
        '
        'chkDontUpdateDatabase
        '
        Me.chkDontUpdateDatabase.BackColor = System.Drawing.SystemColors.Control
        Me.chkDontUpdateDatabase.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDontUpdateDatabase.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDontUpdateDatabase.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDontUpdateDatabase.Location = New System.Drawing.Point(270, 285)
        Me.chkDontUpdateDatabase.Name = "chkDontUpdateDatabase"
        Me.chkDontUpdateDatabase.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDontUpdateDatabase.Size = New System.Drawing.Size(241, 21)
        Me.chkDontUpdateDatabase.TabIndex = 9
        Me.chkDontUpdateDatabase.Text = "Don't Update Database"
        Me.chkDontUpdateDatabase.UseVisualStyleBackColor = False
        '
        'cmdSetDates
        '
        Me.cmdSetDates.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSetDates.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSetDates.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSetDates.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSetDates.Location = New System.Drawing.Point(410, 12)
        Me.cmdSetDates.Name = "cmdSetDates"
        Me.cmdSetDates.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSetDates.Size = New System.Drawing.Size(128, 79)
        Me.cmdSetDates.TabIndex = 0
        Me.cmdSetDates.Text = "Set DueDate and NoticeDate from This Date"
        Me.cmdSetDates.UseVisualStyleBackColor = False
        '
        'txtDueDate
        '
        Me.txtDueDate.AcceptsReturn = True
        Me.txtDueDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtDueDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDueDate.Font = New System.Drawing.Font("Tahoma", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDueDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDueDate.Location = New System.Drawing.Point(270, 110)
        Me.txtDueDate.MaxLength = 0
        Me.txtDueDate.Name = "txtDueDate"
        Me.txtDueDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDueDate.Size = New System.Drawing.Size(131, 27)
        Me.txtDueDate.TabIndex = 4
        '
        'txtCalendar
        '
        Me.txtCalendar.AcceptsReturn = True
        Me.txtCalendar.BackColor = System.Drawing.SystemColors.Window
        Me.txtCalendar.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCalendar.Font = New System.Drawing.Font("Tahoma", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCalendar.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCalendar.Location = New System.Drawing.Point(270, 40)
        Me.txtCalendar.MaxLength = 0
        Me.txtCalendar.Name = "txtCalendar"
        Me.txtCalendar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCalendar.Size = New System.Drawing.Size(131, 27)
        Me.txtCalendar.TabIndex = 2
        '
        'cmdDueDate
        '
        Me.cmdDueDate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDueDate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDueDate.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDueDate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDueDate.Location = New System.Drawing.Point(40, 100)
        Me.cmdDueDate.Name = "cmdDueDate"
        Me.cmdDueDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDueDate.Size = New System.Drawing.Size(181, 41)
        Me.cmdDueDate.TabIndex = 3
        Me.cmdDueDate.Text = "Send DueDate Emails"
        Me.cmdDueDate.UseVisualStyleBackColor = False
        '
        'cmdEmailNotices
        '
        Me.cmdEmailNotices.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEmailNotices.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdEmailNotices.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEmailNotices.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEmailNotices.Location = New System.Drawing.Point(40, 170)
        Me.cmdEmailNotices.Name = "cmdEmailNotices"
        Me.cmdEmailNotices.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdEmailNotices.Size = New System.Drawing.Size(181, 41)
        Me.cmdEmailNotices.TabIndex = 5
        Me.cmdEmailNotices.Text = "Send Notice Emails"
        Me.cmdEmailNotices.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(330, 361)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(181, 21)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "2014-May-12  8:50"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'chkShowMessages
        '
        Me.chkShowMessages.BackColor = System.Drawing.SystemColors.Control
        Me.chkShowMessages.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkShowMessages.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShowMessages.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkShowMessages.Location = New System.Drawing.Point(270, 312)
        Me.chkShowMessages.Name = "chkShowMessages"
        Me.chkShowMessages.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkShowMessages.Size = New System.Drawing.Size(241, 21)
        Me.chkShowMessages.TabIndex = 12
        Me.chkShowMessages.Text = "Show Messages"
        Me.chkShowMessages.UseVisualStyleBackColor = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(550, 391)
        Me.Controls.Add(Me.chkShowMessages)
        Me.Controls.Add(Me.chkDontUpdateDatabase)
        Me.Controls.Add(Me.chkSMTPtest)
        Me.Controls.Add(Me.chkDontSendMail)
        Me.Controls.Add(Me.cmdSetDates)
        Me.Controls.Add(Me.txtDueDate)
        Me.Controls.Add(Me.txtCalendar)
        Me.Controls.Add(Me.cmdDueDate)
        Me.Controls.Add(Me.txtNotice)
        Me.Controls.Add(Me.cmdEmailNotices)
        Me.Controls.Add(Me.cmdEmailTest)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 29)
        Me.Name = "Form1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Docket Control Email"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents chkSMTPtest As System.Windows.Forms.CheckBox
    Public WithEvents chkDontUpdateDatabase As System.Windows.Forms.CheckBox
    Public WithEvents chkShowMessages As System.Windows.Forms.CheckBox
#End Region
End Class