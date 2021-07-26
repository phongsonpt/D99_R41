'#------------------------------------------------------
'# Title: D99F0002
'# CreateUser: NGUYEN NGOC THANH
'# CreateDate: 24/03/2004
'# ModifiedUser: Minh Hòa
'# ModifiedDate: 05/06/2013
'# Description: Phân quyền In
'# Xử lý tác vụ form Print
'# Sửa lỗi không mở được PrintSetup của WIN7 64bit
'# Hỗ trợ in 2 mặt (cho máy in hai mặt)
'# Sửa design form: FormBorderStyle = Sizable
'# Sửa lỗi không in được ra giấy trên máy in Canon (máy Canon có 2 ngăn giấy nhưng không lấy được ngăn giấy mặc định)
'#------------------------------------------------------
Imports System.Drawing
Imports System.Diagnostics
Imports System.Drawing.Printing
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports System

Friend Class D99F0002

    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
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
    Friend WithEvents cboExportTypesList As System.Windows.Forms.ComboBox
    Friend WithEvents btnExport As System.Windows.Forms.Button
    Friend WithEvents tip1 As System.Windows.Forms.ToolTip
    Private WithEvents btnPrint As System.Windows.Forms.Button
    Private WithEvents btnPrintSetup As System.Windows.Forms.Button
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents CRViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D99F0002))
        Me.CRViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.cboExportTypesList = New System.Windows.Forms.ComboBox
        Me.tip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnPrintSetup = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnExport = New System.Windows.Forms.Button
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.SuspendLayout()
        '
        'CRViewer1
        '
        Me.CRViewer1.ActiveViewIndex = -1
        Me.CRViewer1.AllowDrop = True
        Me.CRViewer1.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange
        Me.CRViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CRViewer1.DisplayBackgroundEdge = False
        Me.CRViewer1.DisplayGroupTree = False
        Me.CRViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CRViewer1.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.CRViewer1.Location = New System.Drawing.Point(0, 0)
        Me.CRViewer1.Name = "CRViewer1"
        Me.CRViewer1.SelectionFormula = " "
        Me.CRViewer1.ShowCloseButton = False
        Me.CRViewer1.ShowExportButton = False
        Me.CRViewer1.Size = New System.Drawing.Size(790, 541)
        Me.CRViewer1.TabIndex = 4
        Me.CRViewer1.ViewTimeSelectionFormula = ""
        '
        'cboExportTypesList
        '
        Me.cboExportTypesList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboExportTypesList.FormattingEnabled = True
        Me.cboExportTypesList.Location = New System.Drawing.Point(319, 4)
        Me.cboExportTypesList.Name = "cboExportTypesList"
        Me.cboExportTypesList.Size = New System.Drawing.Size(142, 21)
        Me.cboExportTypesList.TabIndex = 5
        Me.cboExportTypesList.Visible = False
        '
        'btnPrintSetup
        '
        Me.btnPrintSetup.BackgroundImage = CType(resources.GetObject("btnPrintSetup.BackgroundImage"), System.Drawing.Image)
        Me.btnPrintSetup.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnPrintSetup.Location = New System.Drawing.Point(535, 2)
        Me.btnPrintSetup.Name = "btnPrintSetup"
        Me.btnPrintSetup.Size = New System.Drawing.Size(28, 22)
        Me.btnPrintSetup.TabIndex = 8
        Me.tip1.SetToolTip(Me.btnPrintSetup, "Print Setup (Ctrl + S)")
        Me.btnPrintSetup.UseVisualStyleBackColor = True
        Me.btnPrintSetup.Visible = False
        '
        'btnPrint
        '
        Me.btnPrint.BackgroundImage = CType(resources.GetObject("btnPrint.BackgroundImage"), System.Drawing.Image)
        Me.btnPrint.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnPrint.Location = New System.Drawing.Point(501, 2)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(28, 22)
        Me.btnPrint.TabIndex = 7
        Me.tip1.SetToolTip(Me.btnPrint, "Print (Ctrl + O)")
        Me.btnPrint.UseVisualStyleBackColor = True
        Me.btnPrint.Visible = False
        '
        'btnExport
        '
        Me.btnExport.BackgroundImage = CType(resources.GetObject("btnExport.BackgroundImage"), System.Drawing.Image)
        Me.btnExport.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnExport.Location = New System.Drawing.Point(467, 2)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(28, 22)
        Me.btnExport.TabIndex = 6
        Me.btnExport.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.tip1.SetToolTip(Me.btnExport, "Export (Ctrl + E)")
        Me.btnExport.UseVisualStyleBackColor = True
        Me.btnExport.Visible = False
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'D99F0002
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(790, 541)
        Me.Controls.Add(Me.btnPrintSetup)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.cboExportTypesList)
        Me.Controls.Add(Me.CRViewer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "D99F0002"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Crystal Reports"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Event PreviewClick(ByVal bExport As Boolean)

    Private myDiskFileDestinationOptions As DiskFileDestinationOptions
    Private myExportOptions As ExportOptions
    Private sExportFileName As String

    'Private PathTemp As String = Application.StartupPath & "\Temp\"
    Private PathTemp As String = gsApplicationPath & "\Temp\"
    Private PathExport As String = PathTemp & "Exported\"
    Private arr() As String

    ' 6/5/2014 id 65095 - khi dung fomshow phải giữ lại giả trị form gọi
    Private _formParent As Form = Nothing
    Public Property FormParent() As Form
        Get
            Return _formParent
        End Get
        Set(ByVal Value As Form)
            _formParent = Value
        End Set
    End Property

    Private Sub D99F0002_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        If My.Computer.FileSystem.FileExists(gsSaveFileName1) = True Then
            My.Computer.FileSystem.DeleteFile(gsSaveFileName1)
        End If

    End Sub

    Private Sub D99F0002_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        'Thiên Huỳnh Update 11/05/2009
        If e.Control And e.KeyCode = Keys.O Then
            If btnPrint.Enabled = True Then btnPrint_Click(Nothing, Nothing)
        ElseIf e.Control And e.KeyCode = Keys.S Then
            'CRViewer1.PrintReport()
            If btnPrintSetup.Enabled = True Then btnPrintSetup_Click(Nothing, Nothing)
        ElseIf e.Control And e.KeyCode = Keys.E Then
            DoExport()
        End If
    End Sub

    Private Sub D99F0002_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'Thiên Huỳnh Update 11/05/2009
            giNumberOfPrint = 0 'Khai báo toàn cục

            cboExportTypesList.Items.Add("RichTextFormat - (*.rtf)")
            cboExportTypesList.Items.Add("Word - (*.doc)")
            cboExportTypesList.Items.Add("Excel - (*.xls)")
            cboExportTypesList.Items.Add("ExcelRecord - (*.xls)")
            cboExportTypesList.Items.Add("PortableDocFormat - (*.pdf)")

            cboExportTypesList.Text = "Excel - (*.xls)"

            Application.DoEvents()
            'CRViewer1.Zoom(1)
            CRViewer1.Dock = DockStyle.Fill
            btnExport.Visible = True
            cboExportTypesList.Visible = True
            btnPrint.Visible = True
            btnPrintSetup.Visible = True

            If giPermissionPrint < 2 Then '10/04/2013, Ngọc Thoại: Lưu lại số lần in - Phân quyền In
                If giPrintNumber > 0 Then
                    btnPrint.Enabled = False
                    btnPrintSetup.Enabled = False
                End If
            End If

            'SetResolutionForm(Me)
            Application.DoEvents()
        Catch ex As Exception
            MessageBox.Show("Error D99F0002_Load")
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Function ExistColumnNew(ByVal sTableName As String, ByVal sFieldName As String) As Boolean
        'Kiểm tra tồn tại cột mới thêm vào bảng
        Dim sSQL As String
        sSQL = "SELECT top 1 1 FROM syscolumns col INNER JOIN sysobjects tab On col.id = tab.id"
        sSQL = " WHERE  tab.name = " & SQLString(sTableName)
        sSQL = "    AND  col.name = " & SQLString(sFieldName)
        Return ExistRecord(sSQL)
    End Function

    Private Function AllowExportFile() As Boolean
        'If Not ExistColumnNew("D91T9102", "CheckPermissionExportFile") Then Return True

        Dim bExport As Boolean = True
        If ReturnScalar("Select top 1 CheckPermissionExportFile from D91T9102") = "1" Then
            bExport = ReturnPermission("D91F5708") > 0
        End If
        If Not bExport Then
            D99C0008.MsgL3(rL3("Ban_khong_co_quyen_xuat_file"))
        End If

        Return bExport
    End Function

    Private Sub DoExport()
        'Update 18/09/2014: Incident 67401 Kiểm tra phân quyền khi file khi in báo cáo
        If Not AllowExportFile() Then Exit Sub

        Try
            If cboExportTypesList.Text = "" Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor

            arr = Split(My.Computer.FileSystem.GetName(gsMainReportName1), ".")

            ExportSetup()
            ExportSelection()
            rpt1.Export()

            Me.Cursor = Cursors.Default

            'MessageBox.Show("File exported sucessfully.", "Exporting done", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Process.Start(sExportFileName)

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            If cboExportTypesList.Text = "Excel - (*.xls)" Or cboExportTypesList.Text = "ExcelRecord - (*.xls)" Then
                'If nRowMaximum > RowMaximumExcel Then
                '    MessageBox.Show("Kh¤ng thÓ xuÊt dö liÖu ra Excel vØ sç dßng ¢º v§ít quÀ " & RowMaximumExcel, "Exporting done", MessageBoxButtons.OK, MessageBoxIcon.Error)
                'End If
                If ex.Message.Contains("OutOfMemoryException") Then
                    D99C0008.MsgL3("Lỗi tràn bộ nhớ khi xuất excel từ report." & vbCrLf & ex.Message & " - " & ex.Source)
                Else
                    GoTo ErrExport
                End If
            Else
ErrExport:

                MessageBox.Show("Error export: " & ex.Message & vbCrLf & ex.Source, "Exporting done", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If


        End Try
    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        DoExport()
    End Sub

    Private Sub ExportSetup()

        'Tạo thư mục chứa file Export tạm
        If Not My.Computer.FileSystem.DirectoryExists(PathTemp) Then
            My.Computer.FileSystem.CreateDirectory(PathTemp)
            My.Computer.FileSystem.CreateDirectory(PathExport)
        Else
            If Not My.Computer.FileSystem.DirectoryExists(PathExport) Then
                My.Computer.FileSystem.CreateDirectory(PathExport)
            End If
        End If

        myDiskFileDestinationOptions = New DiskFileDestinationOptions()
        myExportOptions = rpt1.ExportOptions
        myExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
        myExportOptions.FormatOptions = Nothing

    End Sub

    Private Sub ExportSelection()
        Select Case cboExportTypesList.SelectedIndex
            Case 0
                ConfigureExportToRichTextFormat()
            Case 1
                ConfigureExportToWord()
            Case 2
                ConfigureExportToExcel()
            Case 3
                ConfigureExportToExcelRecord()
            Case 4
                ConfigureExportToPdf()
        End Select
    End Sub

    Private Sub ConfigureExportToRichTextFormat()

        sExportFileName = PathExport & arr(0) & ".rtf"

        myExportOptions.ExportFormatType = ExportFormatType.RichText
        myDiskFileDestinationOptions.DiskFileName = sExportFileName
        myExportOptions.DestinationOptions = myDiskFileDestinationOptions
    End Sub

    Private Sub ConfigureExportToWord()

        sExportFileName = PathExport & arr(0) & ".doc"

        myExportOptions.ExportFormatType = ExportFormatType.WordForWindows
        myDiskFileDestinationOptions.DiskFileName = sExportFileName
        myExportOptions.DestinationOptions = myDiskFileDestinationOptions
    End Sub

    Private Sub ConfigureExportToExcel()
        sExportFileName = PathExport & arr(0) & ".xls"

        myExportOptions.ExportFormatType = ExportFormatType.Excel
        myDiskFileDestinationOptions.DiskFileName = sExportFileName
        myExportOptions.DestinationOptions = myDiskFileDestinationOptions
    End Sub

    Private Sub ConfigureExportToExcelRecord()

        sExportFileName = PathExport & arr(0) & ".xls"

        myExportOptions.ExportFormatType = ExportFormatType.ExcelRecord
        myDiskFileDestinationOptions.DiskFileName = sExportFileName
        myExportOptions.DestinationOptions = myDiskFileDestinationOptions
    End Sub

    Private Sub ConfigureExportToPdf()

        sExportFileName = PathExport & arr(0) & ".pdf"

        myExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat
        myDiskFileDestinationOptions.DiskFileName = sExportFileName
        myExportOptions.DestinationOptions = myDiskFileDestinationOptions
    End Sub

    'Thiên Huỳnh Update 11/05/2009
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        giNumberOfPrint += 1 'Khai báo toàn cục

        'rpt1.PrintOptions.PaperSource = CrystalDecisions.Shared.PaperSource.Auto'theo Incident 64818
        '
        '        MessageBox.Show(" rpt1.PrintOptions.PrinterName=" & rpt1.PrintOptions.PrinterName)
        '        MessageBox.Show(" rpt1.PrintOptions.PaperSize=" & rpt1.PrintOptions.PaperSize)

        rpt1.PrintToPrinter(1, True, 1, TotalPage)

        If giPermissionPrint < 2 Then '10/04/2013, Ngọc Thoại: Lưu lại số lần in - Phân quyền In
            If giNumberOfPrint > 0 Then
                btnPrint.Enabled = False
                btnPrintSetup.Enabled = False
            End If
        End If
    End Sub

    Private Sub btnPrintSetup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrintSetup.Click
        'Dim PrintDialog1 As PrintDialog' Đã đưa vào Design

        Dim result As DialogResult

        PrintDialog1.AllowCurrentPage = True
        'Update 12/10/2011: sửa lỗi không mở được PrintSetup của WIN7 64bit
        PrintDialog1.UseEXDialog = True
        If TotalPage > 1 Then
            PrintDialog1.AllowSomePages = True
        Else
            PrintDialog1.AllowSomePages = False
        End If
        PrintDialog1.PrinterSettings.FromPage = 1
        PrintDialog1.PrinterSettings.ToPage = TotalPage


        'Thiên Huỳnh Edit 21/06/2010: Set Hướng In cho PrintSetup theo Design của Report
        If rpt1.PrintOptions.PaperOrientation = PaperOrientation.Portrait Then
            PrintDialog1.PrinterSettings.DefaultPageSettings.Landscape = False
        Else
            PrintDialog1.PrinterSettings.DefaultPageSettings.Landscape = True
        End If


        If gsPrinterName1 <> "" Then 'Nếu có chọn máy thì lấy máy đó ngược lại lấy Default
            PrintDialog1.PrinterSettings.PrinterName = gsPrinterName1
            If giPaperSizes > -1 Then
                Dim settings As PrinterSettings = New PrinterSettings()
                settings.PrinterName = gsPrinterName1
                Dim oPaperSize As New System.Drawing.Printing.PaperSize

                For i As Integer = 0 To settings.PaperSizes.Count - 1
                    'dtPaperSizes.Rows.Add(New Object() {settings.PaperSizes.Item(i).RawKind, settings.PaperSizes.Item(i).PaperName, sPrinter})
                    If settings.PaperSizes.Item(i).PaperName = gsPaperSizesName Then

                        '   'oPaperSize.RawKind = giPaperSizes
                        '   oPaperSize.PaperName = PaperKind.Custom
                        oPaperSize = settings.PaperSizes.Item(i)
                        oPaperSize.RawKind = settings.PaperSizes.Item(i).RawKind
                        PrintDialog1.PrinterSettings.DefaultPageSettings.PaperSize = oPaperSize
                        '                        MessageBox.Show("PrintDialog1 PaperSize --> PaperName = " & oPaperSize.PaperName & " --> Width= " & oPaperSize.Width & "--> Height= " & oPaperSize.Height)

                        'PageSetupDialog.PageSettings.PaperSize = oPaperSize
                        Exit For
                    End If
                Next

            End If
        End If

        'mPrint.PrinterSettings.PaperSizes 
        'Dim iPaperSize As Integer = CrystalDecisions.Shared.PaperSize.PaperA4 

        'MessageBox.Show("Print Setup = " & PrintDialog1.PrinterSettings.PrinterName)

        result = PrintDialog1.ShowDialog()

        If result = Windows.Forms.DialogResult.OK Then
            '10/04/2013, Ngọc Thoại: Lưu lại số lần in - Phân quyền In
            '========
            If giPermissionPrint < 2 Then
                If giPrintNumber = 0 Then
                    PrintDialog1.PrinterSettings.Copies = 1
                    PrintDialog1.PrinterSettings.Collate = False
                End If
            End If
            '=======

            giNumberOfPrint += PrintDialog1.PrinterSettings.Copies 'Khai báo toàn cục
            '  rpt1.PrintOptions.PaperSource = CrystalDecisions.Shared.PaperSource.Auto'theo Incident 64818

            rpt1.PrintOptions.PrinterName = PrintDialog1.PrinterSettings.PrinterName

            'MessageBox.Show(" Print Setup --> rpt1.PrintOptions.PaperSize = " & rpt1.PrintOptions.PaperSize)
            'MessageBox.Show("PrinterName = " & rpt1.PrintOptions.PrinterName)

            'Thiên Huỳnh Edit 21/06/2010: Máy in 2 mặt
            If PrintDialog1.PrinterSettings.CanDuplex = True Then
                Select Case PrintDialog1.PrinterSettings.Duplex
                    Case Duplex.Default
                        rpt1.PrintOptions.PrinterDuplex = PrinterDuplex.Default
                    Case Duplex.Horizontal
                        rpt1.PrintOptions.PrinterDuplex = PrinterDuplex.Horizontal
                    Case Duplex.Simplex
                        rpt1.PrintOptions.PrinterDuplex = PrinterDuplex.Simplex
                    Case Duplex.Vertical
                        rpt1.PrintOptions.PrinterDuplex = PrinterDuplex.Vertical
                End Select
            End If

            'Thiên Huỳnh Edit 21/06/2010: Set lại Hướng In theo PrintSetup
            If PrintDialog1.PrinterSettings.DefaultPageSettings.Landscape = True Then
                rpt1.PrintOptions.PaperOrientation = PaperOrientation.Landscape
            Else
                rpt1.PrintOptions.PaperOrientation = PaperOrientation.Portrait
            End If

            Select Case PrintDialog1.PrinterSettings.PrintRange
                Case PrintRange.AllPages
                    rpt1.PrintToPrinter(PrintDialog1.PrinterSettings.Copies, PrintDialog1.PrinterSettings.Collate, 1, TotalPage)
                Case PrintRange.SomePages
                    rpt1.PrintToPrinter(PrintDialog1.PrinterSettings.Copies, PrintDialog1.PrinterSettings.Collate, PrintDialog1.PrinterSettings.FromPage, PrintDialog1.PrinterSettings.ToPage)
            End Select
        End If
        If giPermissionPrint < 2 Then '10/04/2013, Ngọc Thoại: Lưu lại số lần in - Phân quyền In
            If giNumberOfPrint > 0 Then
                btnPrint.Enabled = False
                btnPrintSetup.Enabled = False
            End If
        End If


    End Sub

    Dim TotalPage As Integer

    'Private Sub CRViewer1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles CRViewer1.Load

    '    'CRViewer1.ShowLastPage()
    '    'TotalPage = CRViewer1.GetCurrentPageNumber
    'End Sub
    Private Sub D99F0002_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        'Update 04/02/2015: Lỗi do thêm 1 paramater nhưng trong code không truyền giá trị vào, sửa lỗi theo incident 72781
        Try
            TotalPage = rpt1.FormatEngine.GetLastPageNumber(New CrystalDecisions.Shared.ReportPageRequestContext)
        Catch ex As Exception
            TotalPage = 1
            D99C0008.MsgL3("Error: " & ex.Message & " - " & ex.Source)
        End Try
    End Sub
End Class