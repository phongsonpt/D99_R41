<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D91F5554
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D91F5554))
        Dim Style1 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style2 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style3 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style4 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style5 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style6 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style7 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Dim Style8 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style()
        Me.tdbg = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnCreate = New System.Windows.Forms.Button()
        Me.tdbdPeriod = New C1.Win.C1TrueDBGrid.C1TrueDBDropdown()
        CType(Me.tdbg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tdbdPeriod, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tdbg
        '
        Me.tdbg.AllowColMove = False
        Me.tdbg.AllowColSelect = False
        Me.tdbg.AllowFilter = False
        Me.tdbg.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbg.AllowSort = False
        Me.tdbg.AlternatingRows = True
        Me.tdbg.CaptionHeight = 17
        Me.tdbg.ColumnFooters = True
        Me.tdbg.EmptyRows = True
        Me.tdbg.ExtendRightColumn = True
        Me.tdbg.FilterBar = True
        Me.tdbg.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbg.GroupByCaption = "Drag a column header here to group by that column"
        Me.tdbg.Images.Add(CType(resources.GetObject("tdbg.Images"), System.Drawing.Image))
        Me.tdbg.Location = New System.Drawing.Point(1, 0)
        Me.tdbg.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.FloatingEditor
        Me.tdbg.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbg.Name = "tdbg"
        Me.tdbg.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbg.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbg.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbg.PrintInfo.PageSettings = CType(resources.GetObject("tdbg.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbg.RowHeight = 15
        Me.tdbg.Size = New System.Drawing.Size(449, 260)
        Me.tdbg.SplitDividerSize = New System.Drawing.Size(1, 1)
        Me.tdbg.TabAcrossSplits = True
        Me.tdbg.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbg.TabIndex = 0
        Me.tdbg.Tag = "COL"
        Me.tdbg.WrapCellPointer = True
        Me.tdbg.PropBag = resources.GetString("tdbg.PropBag")
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(374, 268)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 26)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Đó&ng"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(287, 268)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(81, 26)
        Me.btnCreate.TabIndex = 2
        Me.btnCreate.Text = "Khóa sổ"
        Me.btnCreate.UseVisualStyleBackColor = True
        '
        'tdbdPeriod
        '
        Me.tdbdPeriod.AllowColMove = False
        Me.tdbdPeriod.AllowColSelect = False
        Me.tdbdPeriod.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbdPeriod.AllowSort = False
        Me.tdbdPeriod.AlternatingRows = True
        Me.tdbdPeriod.CaptionHeight = 17
        Me.tdbdPeriod.CaptionStyle = Style1
        Me.tdbdPeriod.ColumnCaptionHeight = 17
        Me.tdbdPeriod.ColumnFooterHeight = 17
        Me.tdbdPeriod.ColumnHeaders = False
        Me.tdbdPeriod.DisplayMember = "Period"
        Me.tdbdPeriod.EmptyRows = True
        Me.tdbdPeriod.EvenRowStyle = Style2
        Me.tdbdPeriod.ExtendRightColumn = True
        Me.tdbdPeriod.FetchRowStyles = False
        Me.tdbdPeriod.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbdPeriod.FooterStyle = Style3
        Me.tdbdPeriod.HeadingStyle = Style4
        Me.tdbdPeriod.HighLightRowStyle = Style5
        Me.tdbdPeriod.Images.Add(CType(resources.GetObject("tdbdPeriod.Images"), System.Drawing.Image))
        Me.tdbdPeriod.Location = New System.Drawing.Point(329, 49)
        Me.tdbdPeriod.Name = "tdbdPeriod"
        Me.tdbdPeriod.OddRowStyle = Style6
        Me.tdbdPeriod.RecordSelectorStyle = Style7
        Me.tdbdPeriod.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.tdbdPeriod.RowDivider.Style = C1.Win.C1TrueDBGrid.LineStyleEnum.[Single]
        Me.tdbdPeriod.RowHeight = 15
        Me.tdbdPeriod.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbdPeriod.ScrollTips = False
        Me.tdbdPeriod.Size = New System.Drawing.Size(111, 147)
        Me.tdbdPeriod.Style = Style8
        Me.tdbdPeriod.TabIndex = 133
        Me.tdbdPeriod.TabStop = False
        Me.tdbdPeriod.ValueMember = "Period"
        Me.tdbdPeriod.Visible = False
        Me.tdbdPeriod.PropBag = resources.GetString("tdbdPeriod.PropBag")
        '
        'D91F5554
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(452, 301)
        Me.Controls.Add(Me.tdbdPeriod)
        Me.Controls.Add(Me.btnCreate)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.tdbg)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D91F5554"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Khóa sổ - D91F5554"
        CType(Me.tdbg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tdbdPeriod, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Private WithEvents btnClose As System.Windows.Forms.Button
    Private WithEvents btnCreate As System.Windows.Forms.Button
    Private WithEvents tdbdPeriod As C1.Win.C1TrueDBGrid.C1TrueDBDropdown
End Class