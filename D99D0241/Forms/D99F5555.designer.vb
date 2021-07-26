<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D99F5555
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D99F5555))
        Me.tdbg1 = New C1.Win.C1TrueDBGrid.C1TrueDBGrid
        Me.pnlInfo = New System.Windows.Forms.Panel
        Me.lblEnter = New System.Windows.Forms.Label
        Me.lblEscap = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        CType(Me.tdbg1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlInfo.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tdbg1
        '
        Me.tdbg1.AllowColMove = False
        Me.tdbg1.AllowColSelect = False
        Me.tdbg1.AllowFilter = False
        Me.tdbg1.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbg1.AllowUpdate = False
        Me.tdbg1.AlternatingRows = True
        Me.tdbg1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tdbg1.CaptionHeight = 17
        Me.tdbg1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tdbg1.EmptyRows = True
        Me.tdbg1.ExtendRightColumn = True
        Me.tdbg1.FilterBar = True
        Me.tdbg1.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbg1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbg1.GroupByCaption = "Drag a column header here to group by that column"
        Me.tdbg1.Images.Add(CType(resources.GetObject("tdbg1.Images"), System.Drawing.Image))
        Me.tdbg1.Location = New System.Drawing.Point(0, 0)
        Me.tdbg1.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbg1.Name = "tdbg1"
        Me.tdbg1.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbg1.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbg1.PreviewInfo.ZoomFactor = 75
        Me.tdbg1.PrintInfo.PageSettings = CType(resources.GetObject("tdbg1.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbg1.RecordSelectors = False
        Me.tdbg1.RowHeight = 15
        Me.tdbg1.Size = New System.Drawing.Size(427, 165)
        Me.tdbg1.TabAcrossSplits = True
        Me.tdbg1.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbg1.TabIndex = 1
        Me.tdbg1.PropBag = resources.GetString("tdbg1.PropBag")
        '
        'pnlInfo
        '
        Me.pnlInfo.Controls.Add(Me.lblEnter)
        Me.pnlInfo.Controls.Add(Me.lblEscap)
        Me.pnlInfo.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlInfo.Location = New System.Drawing.Point(0, 165)
        Me.pnlInfo.Name = "pnlInfo"
        Me.pnlInfo.Size = New System.Drawing.Size(427, 46)
        Me.pnlInfo.TabIndex = 3
        '
        'lblEnter
        '
        Me.lblEnter.AutoSize = True
        Me.lblEnter.Location = New System.Drawing.Point(12, 5)
        Me.lblEnter.Name = "lblEnter"
        Me.lblEnter.Size = New System.Drawing.Size(257, 13)
        Me.lblEnter.TabIndex = 0
        Me.lblEnter.Text = "- Nhấn phím Enter: Trả dữ liệu về màn hình nhập liệu"
        '
        'lblEscap
        '
        Me.lblEscap.AutoSize = True
        Me.lblEscap.Location = New System.Drawing.Point(12, 24)
        Me.lblEscap.Name = "lblEscap"
        Me.lblEscap.Size = New System.Drawing.Size(124, 13)
        Me.lblEscap.TabIndex = 0
        Me.lblEscap.Text = "- Nhấn phím ESC: Thoát"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.tdbg1)
        Me.Panel1.Controls.Add(Me.pnlInfo)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(429, 213)
        Me.Panel1.TabIndex = 4
        '
        'D99F5555
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(429, 213)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D99F5555"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Tìm kiếm dữ liệu - D99F5555"
        CType(Me.tdbg1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlInfo.ResumeLayout(False)
        Me.pnlInfo.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents tdbg1 As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Friend WithEvents pnlInfo As System.Windows.Forms.Panel
    Friend WithEvents lblEnter As System.Windows.Forms.Label
    Friend WithEvents lblEscap As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
End Class