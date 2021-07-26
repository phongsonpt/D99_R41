<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D99F5556
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D99F5556))
        Me.tdbg1 = New C1.Win.C1TrueDBGrid.C1TrueDBGrid
        Me.pnlInfo = New System.Windows.Forms.Panel
        Me.lblSpace = New System.Windows.Forms.Label
        Me.lblEnter = New System.Windows.Forms.Label
        Me.lblEscap = New System.Windows.Forms.Label
        CType(Me.tdbg1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlInfo.SuspendLayout()
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
        Me.tdbg1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tdbg1.CaptionHeight = 17
        Me.tdbg1.EmptyRows = True
        Me.tdbg1.ExtendRightColumn = True
        Me.tdbg1.FilterBar = True
        Me.tdbg1.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbg1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbg1.GroupByCaption = "Drag a column header here to group by that column"
        Me.tdbg1.Images.Add(CType(resources.GetObject("tdbg1.Images"), System.Drawing.Image))
        Me.tdbg1.Location = New System.Drawing.Point(1, 1)
        Me.tdbg1.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbg1.Name = "tdbg1"
        Me.tdbg1.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbg1.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbg1.PreviewInfo.ZoomFactor = 75
        Me.tdbg1.PrintInfo.PageSettings = CType(resources.GetObject("tdbg1.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbg1.RecordSelectors = False
        Me.tdbg1.RowHeight = 15
        Me.tdbg1.Size = New System.Drawing.Size(547, 280)
        Me.tdbg1.TabAcrossSplits = True
        Me.tdbg1.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbg1.TabIndex = 1
        Me.tdbg1.PropBag = resources.GetString("tdbg1.PropBag")
        '
        'pnlInfo
        '
        Me.pnlInfo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlInfo.Controls.Add(Me.lblSpace)
        Me.pnlInfo.Controls.Add(Me.lblEnter)
        Me.pnlInfo.Controls.Add(Me.lblEscap)
        Me.pnlInfo.Location = New System.Drawing.Point(1, 283)
        Me.pnlInfo.Name = "pnlInfo"
        Me.pnlInfo.Size = New System.Drawing.Size(547, 64)
        Me.pnlInfo.TabIndex = 2
        '
        'lblSpace
        '
        Me.lblSpace.AutoSize = True
        Me.lblSpace.Location = New System.Drawing.Point(12, 24)
        Me.lblSpace.Name = "lblSpace"
        Me.lblSpace.Size = New System.Drawing.Size(188, 13)
        Me.lblSpace.TabIndex = 0
        Me.lblSpace.Text = "- Nhấn phím Space: Chọn nhiều giá trị"
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
        Me.lblEscap.Location = New System.Drawing.Point(12, 43)
        Me.lblEscap.Name = "lblEscap"
        Me.lblEscap.Size = New System.Drawing.Size(124, 13)
        Me.lblEscap.TabIndex = 0
        Me.lblEscap.Text = "- Nhấn phím ESC: Thoát"
        '
        'D99F5556
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(550, 348)
        Me.Controls.Add(Me.pnlInfo)
        Me.Controls.Add(Me.tdbg1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D99F5556"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Tìm kiếm dữ liệu - D99F5556"
        CType(Me.tdbg1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlInfo.ResumeLayout(False)
        Me.pnlInfo.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents tdbg1 As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Friend WithEvents pnlInfo As System.Windows.Forms.Panel
    Friend WithEvents lblSpace As System.Windows.Forms.Label
    Friend WithEvents lblEnter As System.Windows.Forms.Label
    Friend WithEvents lblEscap As System.Windows.Forms.Label
End Class