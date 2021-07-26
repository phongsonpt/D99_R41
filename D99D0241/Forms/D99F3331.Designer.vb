<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D99F3331
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D99F3331))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtFormular = New System.Windows.Forms.TextBox
        Me.lblOutputWindow = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.btnPercent = New CustomButton
        Me.btnClose = New CustomButton
        Me.btnOpen = New CustomButton
        Me.btnDel = New CustomButton
        Me.btnBack = New CustomButton
        Me.btnClear = New CustomButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnEqual = New CustomButton
        Me.btnPlus = New CustomButton
        Me.btnSub = New CustomButton
        Me.btnMulti = New CustomButton
        Me.btnDivide = New CustomButton
        Me.btn9 = New CustomButton
        Me.btn8 = New CustomButton
        Me.btn7 = New CustomButton
        Me.btn6 = New CustomButton
        Me.btn5 = New CustomButton
        Me.btn4 = New CustomButton
        Me.btn3 = New CustomButton
        Me.btn2 = New CustomButton
        Me.btn1 = New CustomButton
        Me.btnPoint = New CustomButton
        Me.btn00 = New CustomButton
        Me.btn0 = New CustomButton
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtFormular)
        Me.GroupBox1.Controls.Add(Me.lblOutputWindow)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(408, 328)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'txtFormular
        '
        Me.txtFormular.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFormular.Location = New System.Drawing.Point(5, 19)
        Me.txtFormular.Name = "txtFormular"
        Me.txtFormular.Size = New System.Drawing.Size(391, 29)
        Me.txtFormular.TabIndex = 0
        '
        'lblOutputWindow
        '
        Me.lblOutputWindow.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblOutputWindow.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOutputWindow.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lblOutputWindow.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOutputWindow.ForeColor = System.Drawing.Color.Black
        Me.lblOutputWindow.Location = New System.Drawing.Point(6, 55)
        Me.lblOutputWindow.Name = "lblOutputWindow"
        Me.lblOutputWindow.Size = New System.Drawing.Size(389, 29)
        Me.lblOutputWindow.TabIndex = 1
        Me.lblOutputWindow.Text = "0"
        Me.lblOutputWindow.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox3.Controls.Add(Me.btnPercent)
        Me.GroupBox3.Controls.Add(Me.btnClose)
        Me.GroupBox3.Controls.Add(Me.btnOpen)
        Me.GroupBox3.Controls.Add(Me.btnDel)
        Me.GroupBox3.Controls.Add(Me.btnBack)
        Me.GroupBox3.Controls.Add(Me.btnClear)
        Me.GroupBox3.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(270, 89)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(126, 227)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        '
        'btnPercent
        '
        Me.btnPercent.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btnPercent.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btnPercent.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btnPercent.ButtonImage = Nothing
        Me.btnPercent.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnPercent.ButtonRadius = 0
        Me.btnPercent.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnPercent.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btnPercent.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPercent.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btnPercent.Location = New System.Drawing.Point(16, 181)
        Me.btnPercent.Name = "btnPercent"
        Me.btnPercent.Size = New System.Drawing.Size(96, 35)
        Me.btnPercent.TabIndex = 5
        Me.btnPercent.Text = "%"
        Me.btnPercent.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnClose
        '
        Me.btnClose.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btnClose.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btnClose.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btnClose.ButtonImage = Nothing
        Me.btnClose.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnClose.ButtonRadius = 0
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnClose.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btnClose.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btnClose.Location = New System.Drawing.Point(67, 140)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(45, 29)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = ")"
        Me.btnClose.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnOpen
        '
        Me.btnOpen.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btnOpen.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btnOpen.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btnOpen.ButtonImage = Nothing
        Me.btnOpen.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnOpen.ButtonRadius = 0
        Me.btnOpen.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnOpen.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btnOpen.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpen.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btnOpen.Location = New System.Drawing.Point(16, 140)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(45, 29)
        Me.btnOpen.TabIndex = 3
        Me.btnOpen.Text = "("
        Me.btnOpen.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnDel
        '
        Me.btnDel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btnDel.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btnDel.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btnDel.ButtonImage = Nothing
        Me.btnDel.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnDel.ButtonRadius = 0
        Me.btnDel.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnDel.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btnDel.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDel.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btnDel.Location = New System.Drawing.Point(16, 99)
        Me.btnDel.Name = "btnDel"
        Me.btnDel.Size = New System.Drawing.Size(96, 35)
        Me.btnDel.TabIndex = 2
        Me.btnDel.Text = "Delete"
        Me.btnDel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnBack
        '
        Me.btnBack.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btnBack.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btnBack.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btnBack.ButtonImage = Nothing
        Me.btnBack.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnBack.ButtonRadius = 0
        Me.btnBack.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnBack.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btnBack.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBack.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btnBack.Location = New System.Drawing.Point(16, 58)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(96, 35)
        Me.btnBack.TabIndex = 1
        Me.btnBack.Text = "Backspace"
        Me.btnBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnClear
        '
        Me.btnClear.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btnClear.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btnClear.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btnClear.ButtonImage = Nothing
        Me.btnClear.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnClear.ButtonRadius = 0
        Me.btnClear.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnClear.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btnClear.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClear.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btnClear.Location = New System.Drawing.Point(16, 17)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(96, 35)
        Me.btnClear.TabIndex = 0
        Me.btnClear.Text = "C"
        Me.btnClear.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnEqual)
        Me.GroupBox2.Controls.Add(Me.btnPlus)
        Me.GroupBox2.Controls.Add(Me.btnSub)
        Me.GroupBox2.Controls.Add(Me.btnMulti)
        Me.GroupBox2.Controls.Add(Me.btnDivide)
        Me.GroupBox2.Controls.Add(Me.btn9)
        Me.GroupBox2.Controls.Add(Me.btn8)
        Me.GroupBox2.Controls.Add(Me.btn7)
        Me.GroupBox2.Controls.Add(Me.btn6)
        Me.GroupBox2.Controls.Add(Me.btn5)
        Me.GroupBox2.Controls.Add(Me.btn4)
        Me.GroupBox2.Controls.Add(Me.btn3)
        Me.GroupBox2.Controls.Add(Me.btn2)
        Me.GroupBox2.Controls.Add(Me.btn1)
        Me.GroupBox2.Controls.Add(Me.btnPoint)
        Me.GroupBox2.Controls.Add(Me.btn00)
        Me.GroupBox2.Controls.Add(Me.btn0)
        Me.GroupBox2.Location = New System.Drawing.Point(11, 88)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(251, 228)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        '
        'btnEqual
        '
        Me.btnEqual.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btnEqual.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btnEqual.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btnEqual.ButtonImage = Nothing
        Me.btnEqual.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnEqual.ButtonRadius = 0
        Me.btnEqual.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnEqual.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btnEqual.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEqual.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btnEqual.Location = New System.Drawing.Point(194, 143)
        Me.btnEqual.Name = "btnEqual"
        Me.btnEqual.Size = New System.Drawing.Size(45, 77)
        Me.btnEqual.TabIndex = 16
        Me.btnEqual.Text = "="
        Me.btnEqual.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnPlus
        '
        Me.btnPlus.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btnPlus.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btnPlus.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btnPlus.ButtonImage = Nothing
        Me.btnPlus.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnPlus.ButtonRadius = 0
        Me.btnPlus.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnPlus.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btnPlus.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPlus.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btnPlus.Location = New System.Drawing.Point(194, 60)
        Me.btnPlus.Name = "btnPlus"
        Me.btnPlus.Size = New System.Drawing.Size(45, 77)
        Me.btnPlus.TabIndex = 15
        Me.btnPlus.Text = "+"
        Me.btnPlus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnSub
        '
        Me.btnSub.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btnSub.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btnSub.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btnSub.ButtonImage = Nothing
        Me.btnSub.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnSub.ButtonRadius = 0
        Me.btnSub.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnSub.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btnSub.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSub.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btnSub.Location = New System.Drawing.Point(194, 18)
        Me.btnSub.Name = "btnSub"
        Me.btnSub.Size = New System.Drawing.Size(45, 35)
        Me.btnSub.TabIndex = 14
        Me.btnSub.Text = "-"
        Me.btnSub.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnMulti
        '
        Me.btnMulti.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btnMulti.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btnMulti.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btnMulti.ButtonImage = Nothing
        Me.btnMulti.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnMulti.ButtonRadius = 0
        Me.btnMulti.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnMulti.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btnMulti.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMulti.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btnMulti.Location = New System.Drawing.Point(134, 18)
        Me.btnMulti.Name = "btnMulti"
        Me.btnMulti.Size = New System.Drawing.Size(45, 35)
        Me.btnMulti.TabIndex = 13
        Me.btnMulti.Text = "*"
        Me.btnMulti.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnDivide
        '
        Me.btnDivide.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btnDivide.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btnDivide.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btnDivide.ButtonImage = Nothing
        Me.btnDivide.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnDivide.ButtonRadius = 0
        Me.btnDivide.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnDivide.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btnDivide.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDivide.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btnDivide.Location = New System.Drawing.Point(74, 18)
        Me.btnDivide.Name = "btnDivide"
        Me.btnDivide.Size = New System.Drawing.Size(45, 35)
        Me.btnDivide.TabIndex = 12
        Me.btnDivide.Text = "/"
        Me.btnDivide.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn9
        '
        Me.btn9.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btn9.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btn9.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btn9.ButtonImage = Nothing
        Me.btn9.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn9.ButtonRadius = 0
        Me.btn9.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btn9.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btn9.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn9.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btn9.Location = New System.Drawing.Point(134, 60)
        Me.btn9.Name = "btn9"
        Me.btn9.Size = New System.Drawing.Size(45, 35)
        Me.btn9.TabIndex = 11
        Me.btn9.Text = "9"
        Me.btn9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn8
        '
        Me.btn8.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btn8.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btn8.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btn8.ButtonImage = Nothing
        Me.btn8.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn8.ButtonRadius = 0
        Me.btn8.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btn8.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btn8.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn8.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btn8.Location = New System.Drawing.Point(74, 60)
        Me.btn8.Name = "btn8"
        Me.btn8.Size = New System.Drawing.Size(45, 35)
        Me.btn8.TabIndex = 10
        Me.btn8.Text = "8"
        Me.btn8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn7
        '
        Me.btn7.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btn7.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btn7.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btn7.ButtonImage = Nothing
        Me.btn7.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn7.ButtonRadius = 0
        Me.btn7.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btn7.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btn7.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn7.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btn7.Location = New System.Drawing.Point(14, 60)
        Me.btn7.Name = "btn7"
        Me.btn7.Size = New System.Drawing.Size(45, 35)
        Me.btn7.TabIndex = 9
        Me.btn7.Text = "7"
        Me.btn7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn6
        '
        Me.btn6.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btn6.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btn6.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btn6.ButtonImage = Nothing
        Me.btn6.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn6.ButtonRadius = 0
        Me.btn6.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btn6.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btn6.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn6.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btn6.Location = New System.Drawing.Point(134, 102)
        Me.btn6.Name = "btn6"
        Me.btn6.Size = New System.Drawing.Size(45, 35)
        Me.btn6.TabIndex = 8
        Me.btn6.Text = "6"
        Me.btn6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn5
        '
        Me.btn5.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btn5.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btn5.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btn5.ButtonImage = Nothing
        Me.btn5.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn5.ButtonRadius = 0
        Me.btn5.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btn5.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btn5.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn5.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btn5.Location = New System.Drawing.Point(74, 102)
        Me.btn5.Name = "btn5"
        Me.btn5.Size = New System.Drawing.Size(45, 35)
        Me.btn5.TabIndex = 7
        Me.btn5.Text = "5"
        Me.btn5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn4
        '
        Me.btn4.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btn4.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btn4.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btn4.ButtonImage = Nothing
        Me.btn4.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn4.ButtonRadius = 0
        Me.btn4.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btn4.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btn4.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn4.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btn4.Location = New System.Drawing.Point(14, 102)
        Me.btn4.Name = "btn4"
        Me.btn4.Size = New System.Drawing.Size(45, 35)
        Me.btn4.TabIndex = 6
        Me.btn4.Text = "4"
        Me.btn4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn3
        '
        Me.btn3.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btn3.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btn3.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btn3.ButtonImage = Nothing
        Me.btn3.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn3.ButtonRadius = 0
        Me.btn3.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btn3.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btn3.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn3.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btn3.Location = New System.Drawing.Point(134, 143)
        Me.btn3.Name = "btn3"
        Me.btn3.Size = New System.Drawing.Size(45, 35)
        Me.btn3.TabIndex = 5
        Me.btn3.Text = "3"
        Me.btn3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn2
        '
        Me.btn2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btn2.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btn2.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btn2.ButtonImage = Nothing
        Me.btn2.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn2.ButtonRadius = 0
        Me.btn2.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btn2.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btn2.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn2.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btn2.Location = New System.Drawing.Point(74, 143)
        Me.btn2.Name = "btn2"
        Me.btn2.Size = New System.Drawing.Size(45, 35)
        Me.btn2.TabIndex = 4
        Me.btn2.Text = "2"
        Me.btn2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn1
        '
        Me.btn1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btn1.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btn1.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btn1.ButtonImage = Nothing
        Me.btn1.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn1.ButtonRadius = 0
        Me.btn1.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btn1.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btn1.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn1.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btn1.Location = New System.Drawing.Point(14, 143)
        Me.btn1.Name = "btn1"
        Me.btn1.Size = New System.Drawing.Size(45, 35)
        Me.btn1.TabIndex = 3
        Me.btn1.Text = "1"
        Me.btn1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnPoint
        '
        Me.btnPoint.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btnPoint.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btnPoint.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btnPoint.ButtonImage = Nothing
        Me.btnPoint.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnPoint.ButtonRadius = 0
        Me.btnPoint.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnPoint.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btnPoint.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPoint.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btnPoint.Location = New System.Drawing.Point(134, 185)
        Me.btnPoint.Name = "btnPoint"
        Me.btnPoint.Size = New System.Drawing.Size(45, 35)
        Me.btnPoint.TabIndex = 2
        Me.btnPoint.Text = "."
        Me.btnPoint.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn00
        '
        Me.btn00.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btn00.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btn00.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btn00.ButtonImage = Nothing
        Me.btn00.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn00.ButtonRadius = 0
        Me.btn00.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btn00.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btn00.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn00.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btn00.Location = New System.Drawing.Point(74, 185)
        Me.btn00.Name = "btn00"
        Me.btn00.Size = New System.Drawing.Size(45, 35)
        Me.btn00.TabIndex = 1
        Me.btn00.Text = "00"
        Me.btn00.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn0
        '
        Me.btn0.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly
        Me.btn0.ButtonBackColor = System.Drawing.Color.Transparent
        Me.btn0.ButtonBorderColor = System.Drawing.Color.RoyalBlue
        Me.btn0.ButtonImage = Nothing
        Me.btn0.ButtonImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn0.ButtonRadius = 0
        Me.btn0.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btn0.FlatStyle = CustomButton.EnumFlatStyle.Standard
        Me.btn0.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn0.Gradient.Orientation = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.btn0.Location = New System.Drawing.Point(14, 185)
        Me.btn0.Name = "btn0"
        Me.btn0.Size = New System.Drawing.Size(45, 35)
        Me.btn0.TabIndex = 0
        Me.btn0.Text = "0"
        Me.btn0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'D99F3331
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(428, 345)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D99F3331"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "M¿y t€nh"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblOutputWindow As System.Windows.Forms.Label
    Friend WithEvents btnClear As CustomButton
    Friend WithEvents btnDel As CustomButton
    Friend WithEvents btnBack As CustomButton
    Friend WithEvents btnOpen As CustomButton
    Friend WithEvents btnClose As CustomButton
    Friend WithEvents btnPercent As CustomButton
    Friend WithEvents btn1 As CustomButton
    Friend WithEvents btnPoint As CustomButton
    Friend WithEvents btn00 As CustomButton
    Friend WithEvents btn0 As CustomButton
    Friend WithEvents btnEqual As CustomButton
    Friend WithEvents btnPlus As CustomButton
    Friend WithEvents btnSub As CustomButton
    Friend WithEvents btnMulti As CustomButton
    Friend WithEvents btnDivide As CustomButton
    Friend WithEvents btn9 As CustomButton
    Friend WithEvents btn8 As CustomButton
    Friend WithEvents btn7 As CustomButton
    Friend WithEvents btn6 As CustomButton
    Friend WithEvents btn5 As CustomButton
    Friend WithEvents btn4 As CustomButton
    Friend WithEvents btn3 As CustomButton
    Friend WithEvents btn2 As CustomButton
    Friend WithEvents txtFormular As System.Windows.Forms.TextBox
End Class
