<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReConnect
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
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
        Me.components = New System.ComponentModel.Container
        Me.picProcess = New System.Windows.Forms.PictureBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.lblMessager = New System.Windows.Forms.Label
        Me.tmrWait = New System.Windows.Forms.Timer(Me.components)
        CType(Me.picProcess, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'picProcess
        '
        Me.picProcess.BackColor = System.Drawing.Color.Transparent
        Me.picProcess.Image = Global.D99D0241.My.Resources.Resources.reloader
        Me.picProcess.InitialImage = Nothing
        Me.picProcess.Location = New System.Drawing.Point(139, 7)
        Me.picProcess.Name = "picProcess"
        Me.picProcess.Size = New System.Drawing.Size(49, 46)
        Me.picProcess.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picProcess.TabIndex = 121
        Me.picProcess.TabStop = False
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(253, 92)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(64, 22)
        Me.btnClose.TabIndex = 122
        Me.btnClose.Text = "Đóng"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'lblMessager
        '
        Me.lblMessager.AutoSize = True
        Me.lblMessager.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMessager.Location = New System.Drawing.Point(79, 62)
        Me.lblMessager.Name = "lblMessager"
        Me.lblMessager.Size = New System.Drawing.Size(171, 17)
        Me.lblMessager.TabIndex = 123
        Me.lblMessager.Text = "Mạng yếu xin vui lòng chờ"
        Me.lblMessager.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tmrWait
        '
        Me.tmrWait.Enabled = True
        Me.tmrWait.Interval = 1000
        '
        'frmReConnect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ClientSize = New System.Drawing.Size(323, 119)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblMessager)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.picProcess)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmReConnect"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.picProcess, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents picProcess As System.Windows.Forms.PictureBox
    Private WithEvents btnClose As System.Windows.Forms.Button
    Private WithEvents lblMessager As System.Windows.Forms.Label
    Friend WithEvents tmrWait As System.Windows.Forms.Timer

End Class
