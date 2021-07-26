Imports System
Imports System.Drawing.Printing

'#---------------------------------------------------------------------------------------------------
'# Title: D99F6666
'# Created User: Nguyễn Thị Ánh
'# Created Date: 30/06/2010 14:33
'# Modified User: Nguyễn Thị Minh Hòa
'# Modified Date: 15/02/2011
'# Description: Viết thành Form chung: 
'                                   Đầu vào : _ModuleID, _pathReportID, _reportID, _reportTypeID
'                                   Đầu ra : _pathReportID, _reportID, _reportTitle
'#              Sau khi In gọi thêm hàm SaveOptionReport (tự Viết riêng tại DxxX0002 để lưu lại Tùy chọn)    
'#---------------------------------------------------------------------------------------------------
Public Class D99F6666

    Dim _pressClosed As Boolean = True

#Region "In - Out parameters"
    Private _reportTypeID As String = ""
    Public WriteOnly Property ReportTypeID() As String
        Set(ByVal Value As String)
            _reportTypeID = Value
        End Set
    End Property

    Private _customReport As String = ""
    Public WriteOnly Property CustomReport() As String
        Set(ByVal Value As String)
            _customReport = Value
        End Set
    End Property

    Private _ModuleID As String = "" 'Truyền XX
    Public WriteOnly Property ModuleID() As String
        Set(ByVal Value As String)
            _ModuleID = Value
        End Set
    End Property

    Private _reportPath As String = ""
    Public Property ReportPath() As String
        Get
            Return _reportPath
        End Get
        Set(ByVal Value As String)
            _reportPath = Value
        End Set
    End Property

    Private _reportName As String = ""
    Public Property ReportName() As String
        Get
            Return _reportName
        End Get
        Set(ByVal Value As String)
            _reportName = Value
        End Set
    End Property

    Private _reportTitle As String
    Public Property ReportTitle() As String
        Get
            Return _reportTitle
        End Get
        Set(ByVal Value As String)
            _reportTitle = Value
        End Set
    End Property

    Private _printVNI As Boolean = False
    Public ReadOnly Property PrintVNI() As Boolean 
        Get
            Return _printVNI
        End Get
    End Property

    Private _reportLanguge As Byte = 0
    Public WriteOnly Property ReportLanguge() As Byte 
        Set(ByVal Value As Byte )
            _reportLanguge = Value
        End Set
    End Property
#End Region

    Private _showReportPath As Boolean = True
    Public ReadOnly Property ShowReportPath() As Boolean
        Get
            Return _showReportPath
        End Get
    End Property

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub D99F6666_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If _pressClosed Then _reportName = ""
    End Sub

    Private Sub D99F6666_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
    End Sub

    Private Sub D99F6666_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        LoadLanguage()
        InputbyUnicode(Me, gbUnicode)
        'Update 06/01/2011: Làm theo chuẩn, không lấy hết đường dẫn, chỉ lấy đường dẫn ngắn nhất \XReports\DxxRxxxx.rpt
        'txtStandardReportID.Text = _reportPath & _reportName & ".rpt"
        txtStandardReportID.Text = _reportPath.Substring(_reportPath.LastIndexOf("\"c, _reportPath.Length - 2)) & _reportName & ".rpt"
        LoadTDBCombo()
        tdbcReportID.SelectedValue = _customReport
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub LoadStandardReportName()
        Dim sSQL As String = ""

        'sSQL = "Select " & IIf(geLanguage = EnumLanguage.Vietnamese, "ReportTypeName", " ReportTypeName01").ToString & " As ReportTypeName " & vbCrLf
        'sSQL &= " From D89T0010 Where ModuleID = " & SQLString(_ModuleID) & " And ReportTypeID = " & SQLString(_reportTypeID)

        'Update 15/02/2011: Làm theo chuẩn
        sSQL = "Select "
        If gbUnicode Then
            sSQL &= IIf(geLanguage = EnumLanguage.Vietnamese, "TitleU", "Title01U").ToString & " as Title " & vbCrLf
        Else
            sSQL &= IIf(geLanguage = EnumLanguage.Vietnamese, "Title", "Title01").ToString & " as Title " & vbCrLf
        End If
        sSQL &= " From D89T2000 WITH(NOLOCK) Where ModuleID = " & SQLString(_ModuleID) & " And ReportTypeID = " & SQLString(_reportTypeID)
        sSQL &= " And ReportID = " & SQLString(_reportName)

        txtStandardReportName.Text = ReturnScalar(sSQL)
    End Sub

    Dim dtPaperSizes As DataTable
    Private Sub LoadTDBCombo()
        LoadStandardReportName()
        LoadtdbcCustomizeReport(tdbcReportID, _ModuleID, _reportTypeID, txtReportName)

        '10/3/2021, id 155409
        Dim dtPrinter As New DataTable()
        dtPrinter.Columns.Add("PrinterID", GetType(String))

        'Load nguồn cho khổ giấy
        dtPaperSizes = New DataTable
        dtPaperSizes.Columns.Add("PaperID", GetType(Integer))
        dtPaperSizes.Columns.Add("PaperName", GetType(String))
        dtPaperSizes.Columns.Add("PrinterID", GetType(String))

        tdbcPrinterID.Tag = ""
        Dim sDefaultPrinter As String = ""
        Dim iPaperSizes As Integer = 0
        Dim settings As PrinterSettings = New PrinterSettings()
        For Each printer As String In PrinterSettings.InstalledPrinters
            settings.PrinterName = printer
            If printer <> "Fax" And printer <> "Microsoft XPS Document Writer" Then
                If Not printer.Contains("\\") Then
                    printer = "\\" & My.Computer.Name & "\" & printer
                End If
                dtPrinter.Rows.Add(New Object() {printer})
                If settings.IsDefaultPrinter Then sDefaultPrinter = printer
            End If
        Next

        LoadDataSource(tdbcPrinterID, dtPrinter, True)
    End Sub

    Private Sub LoadtdbcPaperID(ByVal sPrinterID As String)
        '10/3/2021, id 155409
        Try
            If sPrinterID = "" Then Exit Sub

            If dtPaperSizes IsNot Nothing Then dtPaperSizes.Clear()
            If sPrinterID = "-1" Then Exit Sub

            Dim settings As PrinterSettings = New PrinterSettings()
            settings.PrinterName = sPrinterID
            For i As Integer = 0 To settings.PaperSizes.Count - 1
                dtPaperSizes.Rows.Add(New Object() {settings.PaperSizes.Item(i).RawKind, settings.PaperSizes.Item(i).PaperName, sPrinterID})
            Next

            LoadDataSource(tdbcPaperID, dtPaperSizes, True)

        Catch ex As Exception
        End Try
    End Sub

#Region "Events tdbcReportID with txtReportName"

    Private Sub tdbcReportID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcReportID.SelectedValueChanged
        If tdbcReportID.SelectedValue Is Nothing Then
            txtReportName.Text = ""
        Else
            txtReportName.Text = tdbcReportID.Columns("Title").Value.ToString
        End If
    End Sub
#End Region

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim sReportID = "" '10/3/2021, id 155409

        _pressClosed = False
        If tdbcReportID.Text <> "" Then
            '_reportPath = Application.StartupPath & "\XCustom\"
            _reportPath = gsApplicationSetup & "\XCustom\"
            _reportName = tdbcReportID.Text
            _reportTitle = txtReportName.Text

            sReportID = tdbcReportID.Text '10/3/2021, id 155409
        Else
            _reportTitle = txtStandardReportName.Text

            sReportID = System.IO.Path.GetFileNameWithoutExtension(txtStandardReportID.Text) '10/3/2021, id 155409
        End If
        'định nghĩa thêm hàm này tại module dxxx0004
        SaveOptionReportNew(chkViewPathReport.Checked)
        '  _printVNI = chkIsPrintVNI.Checked

        If lblPrefixReport.Text.Contains("rpt") Then SetPaperReport(_reportPath, sReportID) '10/3/2021, id 155409

        Me.Close()
    End Sub

    'Tùy thuộc từng module có biến lưu dưới Registry
    Public Sub SaveOptionReportNew(ByVal bShowReportPath As Boolean)
        D99C0007.SaveModulesSetting("D" & _ModuleID, ModuleOption.lmOptions, "ShowReportPath", bShowReportPath) 'vẫn lưu registry vì có TH chưa có dữ liệu D91T8889 nên biến này chỉ Update tại hàm GetReportPath
        _showReportPath = bShowReportPath 'Biến Tùy chọn
    End Sub

    Private Sub D99F6666_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        tdbcReportID.Focus()
    End Sub

    Private Sub LoadLanguage()
        Me.Text = rl3("Chon_duong_dan_bao_caoF") & UnicodeCaption(gbUnicode) 'Chãn ¢§éng dÉn bÀo cÀo
        'if geLanguage = EnumLanguage.English Then
        If geLanguage <> EnumLanguage.Vietnamese Then
            lblStandardForm.Text = "Report Type of"
            lblReportID.Text = "Customized"

            btnPrint.Text = "&Print"
            btnClose.Text = "&Close"
            tdbcReportID.Columns("ReportID").Caption = "Code"
            tdbcReportID.Columns("Title").Caption = "Description"
        End If
        chkViewPathReport.Text = rL3("MSG000035")
        '  chkIsPrintVNI.Text = rl3("In_bao_cao_VNI") 'Bổ sung 22/02/2013
        lblPrinterID.Text = rL3("May_in")
        lblPaperID.Text = rL3("Kho_giay")
    End Sub

    'Bổ sung 22/02/2013 - Định lại đường dẫn in báo cáo
    Private Sub chkIsPrintVNI_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        _reportPath = UnicodeGetReportPath(True, _reportLanguge, "")
        txtStandardReportID.Text = _reportPath.Substring(_reportPath.LastIndexOf("\"c, _reportPath.Length - 2)) & _reportName & ".rpt"
    End Sub

    Private Sub tdbcPrinterID_Validated(sender As Object, e As EventArgs) Handles tdbcPrinterID.Validated
        If tdbcPrinterID.Tag.ToString = tdbcPrinterID.Text Then Exit Sub
        If tdbcPrinterID.FindStringExact(tdbcPrinterID.Text) = -1 Then
            tdbcPrinterID.Text = ""
            LoadtdbcPaperID("-1")
            Exit Sub
        End If
        LoadtdbcPaperID(tdbcPrinterID.Text)
        tdbcPrinterID.Tag = tdbcPrinterID.Text
    End Sub

    Private Sub tdbcPaperID_Validated(sender As Object, e As EventArgs) Handles tdbcPaperID.Validated
        If tdbcPaperID.FindStringExact(tdbcPaperID.Text) = -1 Then tdbcPaperID.Text = ""
    End Sub

    Private Sub SetPaperReport(sPath As String, sReportID As String)
        '10/3/2021, id 155409
        Try
            If tdbcPrinterID.Text = "" Then Exit Sub
            Dim sReportPath As String = sPath & sReportID & ".rpt"

            Dim rpt1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            rpt1 = New CrystalDecisions.CrystalReports.Engine.ReportDocument
            rpt1.Load(sReportPath)

            Dim iPaperID As Integer = L3Int(ReturnValueC1Combo(tdbcPaperID, "PaperID"))
            rpt1.PrintOptions.PrinterName = tdbcPrinterID.Text
            If iPaperID > -1 Then rpt1.PrintOptions.PaperSize = CType(iPaperID, CrystalDecisions.Shared.PaperSize)

            rpt1.SaveAs(sReportPath)


        Catch ex As Exception

            WriteLogFile(tdbcReportID.Text, "SetupReportPaper.log", True)
        End Try


    End Sub

End Class