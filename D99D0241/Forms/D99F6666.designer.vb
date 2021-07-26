<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D99F6666
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D99F6666))
        Me.txtStandardReportID = New System.Windows.Forms.TextBox()
        Me.lblStandardForm = New System.Windows.Forms.Label()
        Me.tdbcReportID = New C1.Win.C1List.C1Combo()
        Me.lblReportID = New System.Windows.Forms.Label()
        Me.lblPrefixReport = New System.Windows.Forms.Label()
        Me.chkViewPathReport = New System.Windows.Forms.CheckBox()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.txtReportName = New System.Windows.Forms.TextBox()
        Me.txtStandardReportName = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblPaperID = New System.Windows.Forms.Label()
        Me.tdbcPaperID = New C1.Win.C1List.C1Combo()
        Me.lblPrinterID = New System.Windows.Forms.Label()
        Me.tdbcPrinterID = New C1.Win.C1List.C1Combo()
        CType(Me.tdbcReportID, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.tdbcPaperID, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tdbcPrinterID, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtStandardReportID
        '
        Me.txtStandardReportID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtStandardReportID.Location = New System.Drawing.Point(88, 19)
        Me.txtStandardReportID.Name = "txtStandardReportID"
        Me.txtStandardReportID.ReadOnly = True
        Me.txtStandardReportID.Size = New System.Drawing.Size(177, 20)
        Me.txtStandardReportID.TabIndex = 1
        '
        'lblStandardForm
        '
        Me.lblStandardForm.AutoSize = True
        Me.lblStandardForm.Location = New System.Drawing.Point(10, 23)
        Me.lblStandardForm.Name = "lblStandardForm"
        Me.lblStandardForm.Size = New System.Drawing.Size(61, 13)
        Me.lblStandardForm.TabIndex = 0
        Me.lblStandardForm.Text = "Mẫu chuẩn"
        Me.lblStandardForm.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tdbcReportID
        '
        Me.tdbcReportID.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.tdbcReportID.AllowColMove = False
        Me.tdbcReportID.AllowSort = False
        Me.tdbcReportID.AlternatingRows = True
        Me.tdbcReportID.AutoCompletion = True
        Me.tdbcReportID.AutoDropDown = True
        Me.tdbcReportID.Caption = ""
        Me.tdbcReportID.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.tdbcReportID.ColumnWidth = 100
        Me.tdbcReportID.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.tdbcReportID.DisplayMember = "ReportID"
        Me.tdbcReportID.DropdownPosition = C1.Win.C1List.DropdownPositionEnum.LeftDown
        Me.tdbcReportID.DropDownWidth = 350
        Me.tdbcReportID.EditorBackColor = System.Drawing.SystemColors.Window
        Me.tdbcReportID.EditorFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcReportID.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.tdbcReportID.EmptyRows = True
        Me.tdbcReportID.ExtendRightColumn = True
        Me.tdbcReportID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcReportID.Images.Add(CType(resources.GetObject("tdbcReportID.Images"), System.Drawing.Image))
        Me.tdbcReportID.Location = New System.Drawing.Point(88, 47)
        Me.tdbcReportID.MatchEntryTimeout = CType(2000, Long)
        Me.tdbcReportID.MaxDropDownItems = CType(8, Short)
        Me.tdbcReportID.MaxLength = 32767
        Me.tdbcReportID.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.tdbcReportID.Name = "tdbcReportID"
        Me.tdbcReportID.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.tdbcReportID.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbcReportID.Size = New System.Drawing.Size(177, 21)
        Me.tdbcReportID.TabIndex = 4
        Me.tdbcReportID.ValueMember = "ReportID"
        Me.tdbcReportID.PropBag = resources.GetString("tdbcReportID.PropBag")
        '
        'lblReportID
        '
        Me.lblReportID.AutoSize = True
        Me.lblReportID.Location = New System.Drawing.Point(10, 51)
        Me.lblReportID.Name = "lblReportID"
        Me.lblReportID.Size = New System.Drawing.Size(45, 13)
        Me.lblReportID.TabIndex = 3
        Me.lblReportID.Text = "Đặc thù"
        Me.lblReportID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPrefixReport
        '
        Me.lblPrefixReport.AutoSize = True
        Me.lblPrefixReport.Location = New System.Drawing.Point(268, 51)
        Me.lblPrefixReport.Name = "lblPrefixReport"
        Me.lblPrefixReport.Size = New System.Drawing.Size(22, 13)
        Me.lblPrefixReport.TabIndex = 5
        Me.lblPrefixReport.Text = ".rpt"
        Me.lblPrefixReport.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkViewPathReport
        '
        Me.chkViewPathReport.AutoSize = True
        Me.chkViewPathReport.Checked = True
        Me.chkViewPathReport.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkViewPathReport.Location = New System.Drawing.Point(13, 109)
        Me.chkViewPathReport.Name = "chkViewPathReport"
        Me.chkViewPathReport.Size = New System.Drawing.Size(263, 17)
        Me.chkViewPathReport.TabIndex = 11
        Me.chkViewPathReport.Text = "Hiển thị màn hình đường dẫn báo cáo cho lần sau"
        Me.chkViewPathReport.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(462, 141)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(70, 22)
        Me.btnPrint.TabIndex = 1
        Me.btnPrint.Text = "&In"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(538, 141)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(70, 22)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Đó&ng"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'txtReportName
        '
        Me.txtReportName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtReportName.Location = New System.Drawing.Point(301, 47)
        Me.txtReportName.Name = "txtReportName"
        Me.txtReportName.ReadOnly = True
        Me.txtReportName.Size = New System.Drawing.Size(290, 20)
        Me.txtReportName.TabIndex = 6
        '
        'txtStandardReportName
        '
        Me.txtStandardReportName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtStandardReportName.Location = New System.Drawing.Point(271, 19)
        Me.txtStandardReportName.Name = "txtStandardReportName"
        Me.txtStandardReportName.ReadOnly = True
        Me.txtStandardReportName.Size = New System.Drawing.Size(320, 20)
        Me.txtStandardReportName.TabIndex = 2
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblPaperID)
        Me.GroupBox1.Controls.Add(Me.tdbcPaperID)
        Me.GroupBox1.Controls.Add(Me.lblPrinterID)
        Me.GroupBox1.Controls.Add(Me.tdbcPrinterID)
        Me.GroupBox1.Controls.Add(Me.txtStandardReportID)
        Me.GroupBox1.Controls.Add(Me.lblPrefixReport)
        Me.GroupBox1.Controls.Add(Me.txtStandardReportName)
        Me.GroupBox1.Controls.Add(Me.tdbcReportID)
        Me.GroupBox1.Controls.Add(Me.lblReportID)
        Me.GroupBox1.Controls.Add(Me.chkViewPathReport)
        Me.GroupBox1.Controls.Add(Me.lblStandardForm)
        Me.GroupBox1.Controls.Add(Me.txtReportName)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(601, 134)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'lblPaperID
        '
        Me.lblPaperID.AutoSize = True
        Me.lblPaperID.Location = New System.Drawing.Point(301, 80)
        Me.lblPaperID.Name = "lblPaperID"
        Me.lblPaperID.Size = New System.Drawing.Size(48, 13)
        Me.lblPaperID.TabIndex = 9
        Me.lblPaperID.Text = "Khổ giấy"
        Me.lblPaperID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tdbcPaperID
        '
        Me.tdbcPaperID.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.tdbcPaperID.AllowColMove = False
        Me.tdbcPaperID.AllowColSelect = True
        Me.tdbcPaperID.AllowSort = False
        Me.tdbcPaperID.AlternatingRows = True
        Me.tdbcPaperID.AutoCompletion = True
        Me.tdbcPaperID.AutoDropDown = True
        Me.tdbcPaperID.Caption = ""
        Me.tdbcPaperID.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.tdbcPaperID.ColumnHeaders = False
        Me.tdbcPaperID.ColumnWidth = 100
        Me.tdbcPaperID.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.tdbcPaperID.DisplayMember = "PaperName"
        Me.tdbcPaperID.DropdownPosition = C1.Win.C1List.DropdownPositionEnum.LeftDown
        Me.tdbcPaperID.DropDownWidth = 240
        Me.tdbcPaperID.EditorBackColor = System.Drawing.SystemColors.Window
        Me.tdbcPaperID.EditorFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcPaperID.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.tdbcPaperID.EmptyRows = True
        Me.tdbcPaperID.ExtendRightColumn = True
        Me.tdbcPaperID.FlatStyle = C1.Win.C1List.FlatModeEnum.Standard
        Me.tdbcPaperID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcPaperID.Images.Add(CType(resources.GetObject("tdbcPaperID.Images"), System.Drawing.Image))
        Me.tdbcPaperID.Location = New System.Drawing.Point(387, 76)
        Me.tdbcPaperID.MatchEntryTimeout = CType(2000, Long)
        Me.tdbcPaperID.MaxDropDownItems = CType(8, Short)
        Me.tdbcPaperID.MaxLength = 32767
        Me.tdbcPaperID.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.tdbcPaperID.Name = "tdbcPaperID"
        Me.tdbcPaperID.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.tdbcPaperID.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbcPaperID.Size = New System.Drawing.Size(204, 21)
        Me.tdbcPaperID.TabIndex = 10
        Me.tdbcPaperID.ValueMember = "PaperID"
        Me.tdbcPaperID.PropBag = resources.GetString("tdbcPaperID.PropBag")
        '
        'lblPrinterID
        '
        Me.lblPrinterID.AutoSize = True
        Me.lblPrinterID.Location = New System.Drawing.Point(10, 80)
        Me.lblPrinterID.Name = "lblPrinterID"
        Me.lblPrinterID.Size = New System.Drawing.Size(38, 13)
        Me.lblPrinterID.TabIndex = 7
        Me.lblPrinterID.Text = "Máy in"
        Me.lblPrinterID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tdbcPrinterID
        '
        Me.tdbcPrinterID.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.tdbcPrinterID.AllowColMove = False
        Me.tdbcPrinterID.AllowColSelect = True
        Me.tdbcPrinterID.AllowSort = False
        Me.tdbcPrinterID.AlternatingRows = True
        Me.tdbcPrinterID.AutoCompletion = True
        Me.tdbcPrinterID.AutoDropDown = True
        Me.tdbcPrinterID.Caption = ""
        Me.tdbcPrinterID.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.tdbcPrinterID.ColumnHeaders = False
        Me.tdbcPrinterID.ColumnWidth = 100
        Me.tdbcPrinterID.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.tdbcPrinterID.DisplayMember = "PrinterID"
        Me.tdbcPrinterID.DropdownPosition = C1.Win.C1List.DropdownPositionEnum.LeftDown
        Me.tdbcPrinterID.DropDownWidth = 240
        Me.tdbcPrinterID.EditorBackColor = System.Drawing.SystemColors.Window
        Me.tdbcPrinterID.EditorFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcPrinterID.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.tdbcPrinterID.EmptyRows = True
        Me.tdbcPrinterID.ExtendRightColumn = True
        Me.tdbcPrinterID.FlatStyle = C1.Win.C1List.FlatModeEnum.Standard
        Me.tdbcPrinterID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcPrinterID.Images.Add(CType(resources.GetObject("tdbcPrinterID.Images"), System.Drawing.Image))
        Me.tdbcPrinterID.Location = New System.Drawing.Point(88, 76)
        Me.tdbcPrinterID.MatchEntryTimeout = CType(2000, Long)
        Me.tdbcPrinterID.MaxDropDownItems = CType(8, Short)
        Me.tdbcPrinterID.MaxLength = 32767
        Me.tdbcPrinterID.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.tdbcPrinterID.Name = "tdbcPrinterID"
        Me.tdbcPrinterID.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.tdbcPrinterID.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbcPrinterID.Size = New System.Drawing.Size(177, 21)
        Me.tdbcPrinterID.TabIndex = 8
        Me.tdbcPrinterID.ValueMember = "PrinterID"
        Me.tdbcPrinterID.PropBag = resources.GetString("tdbcPrinterID.PropBag")
        '
        'D99F6666
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(614, 169)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnPrint)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D99F6666"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Chãn ¢§éng dÉn bÀo cÀo"
        CType(Me.tdbcReportID, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.tdbcPaperID, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tdbcPrinterID, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents txtStandardReportID As System.Windows.Forms.TextBox
    Private WithEvents lblStandardForm As System.Windows.Forms.Label
    Private WithEvents tdbcReportID As C1.Win.C1List.C1Combo
    Private WithEvents lblReportID As System.Windows.Forms.Label
    Private WithEvents lblPrefixReport As System.Windows.Forms.Label
    Private WithEvents chkViewPathReport As System.Windows.Forms.CheckBox
    Private WithEvents btnPrint As System.Windows.Forms.Button
    Private WithEvents btnClose As System.Windows.Forms.Button
    Private WithEvents txtReportName As System.Windows.Forms.TextBox
    Private WithEvents txtStandardReportName As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Private WithEvents lblPaperID As System.Windows.Forms.Label
    Private WithEvents tdbcPaperID As C1.Win.C1List.C1Combo
    Private WithEvents lblPrinterID As System.Windows.Forms.Label
    Private WithEvents tdbcPrinterID As C1.Win.C1List.C1Combo
End Class