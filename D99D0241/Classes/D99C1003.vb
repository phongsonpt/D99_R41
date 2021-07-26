'#------------------------------------------------------
'#Title: D99C1003
'#CreateUser: NGUYEN NGOC THANH
'#CreateDate: 24/03/2004
'#ModifiedUser: Hồ Ngọc Thoại
'#ModifiedDate: 10/04/2013
'#Description: Phân quyền In
'# Sửa tiêu đề form Print Preview có thêm chữ Unicode (nếu form đã chuyển Unicode)
'# Bổ sung chức năng kiểm tra: Nếu SubReport có tồn tại trên Report Design thì mới cho thực thi nguồn.
'# Kiểm tra kết nối khi rớt mạng
'#------------------------------------------------------

Imports System.Data.SqlClient
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing

Public Class D99C1003

    Private WithEvents FormPreview As New D99F0002
    'Mảng sử dụng lưu trữ Parameter của Main
    Private arrPName() As Object
    Private arrPValue() As Object
    'Mảng sử dụng lưu trữ Parameter của  Sub
    Private arrPSubName() As Object
    Private arrPSubValue() As Object
    'Mảng sử dụng lưu trữ của Sub Report truyền vào bằng câu SQL
    Private arrSubSQL() As Object
    Private arrSubName() As Object
    'Mảng sử dụng lưu trữ của Sub Report truyền vào bằng Table
    Private arrSubTable() As Object
    Private arrSubNameTable() As Object
    Private sTableTemp() As String ' Bảng lưu dữ liệu temp
    Private bHostID() As Boolean ' Cờ kiểm tra có dùng HostID

    Public Sub New()
        arrPName = Nothing
        arrPValue = Nothing
        arrPSubName = Nothing
        arrPSubValue = Nothing
        arrSubSQL = Nothing
        arrSubName = Nothing
        arrSubTable = Nothing
        arrSubNameTable = Nothing
        gsMainStrData1 = ""

        rpt1 = New CrystalDecisions.CrystalReports.Engine.ReportDocument

        gbFlagPrint1 = False
    End Sub

    Public Sub OpenConnection(ByVal mConnection As SqlConnection)
        gcConPrint = mConnection
    End Sub

    Private Function CheckExistDefaultPrinter(ByVal sPrinterName As String) As Boolean
        For Each printer As String In PrinterSettings.InstalledPrinters
            If printer = sPrinterName Then
                Return True
            End If
        Next
        Return False
    End Function


    Public Sub PrintReport(ByVal MainReportName As String, Optional ByVal MainReportCaption As String = "Crystal Report", Optional ByVal ModePrint As ReportModeType = ReportModeType.lmPreview, Optional ByVal PrinterName As String = "")
        PrintReport(Nothing, MainReportName, MainReportCaption, ModePrint, PrinterName)
    End Sub

    Public Sub PrintReport(formParent As Form, ByVal MainReportName As String, Optional ByVal MainReportCaption As String = "Crystal Report", Optional ByVal ModePrint As ReportModeType = ReportModeType.lmPreview, Optional ByVal PrinterName As String = "")

        Try

            'Lấy đường dẫn Main report
            gsMainReportName1 = MainReportName

            'Lấy Caption hiển thị
            gsMainReportCaption1 = MainReportCaption
            'Update 22/12/2014: bổ sung tên Máy in được thiết lập ở màn hình Thiết lập cá nhân
            'Test PrinterName = "\\10.0.0.189\HP LaserJet 1160 (Copy 1)"
            'Lấy tên máy in
            If PrinterName = "" Then
                gsPrinterName1 = ""
            Else
                If CheckExistDefaultPrinter(PrinterName) Then
                    gsPrinterName1 = PrinterName
                Else
                    gsPrinterName1 = ""
                    D99C0008.MsgL3("Không tồn tại máy in đang thiết lập.", L3MessageBoxIcon.Exclamation)
                End If
            End If
            'Thực hiện dữ liệu in
            Dim sFormName As String = ""
            If formParent IsNot Nothing Then sFormName = formParent.Name
            Print(sFormName)

            If gbFlagPrint1 = False Then

                Select Case ModePrint
                    Case ReportModeType.lmPreview
                        FormPreview.CRViewer1.ReportSource = rpt1
                        FormPreview.Text = MainReportCaption & UnicodeCaption(gbUnicode) 'Ngọc Thoại update: 14/10/2010
                        FormPreview.ShowDialog() 'Show bị lỗi
                        FormPreview.Dispose()
                    Case ReportModeType.lmPrint
                        'If gsPrinterName1 = "" Then
                        '    Dim localPrinter As System.Drawing.Printing.PrintDocument = New PrintDocument()
                        '    rpt1.PrintOptions.PrinterName = localPrinter.PrinterSettings.PrinterName
                        'End If
                        FormPreview.CRViewer1.ReportSource = rpt1
                        FormPreview.CRViewer1.ShowLastPage()
                        rpt1.PrintToPrinter(1, True, 1, FormPreview.CRViewer1.GetCurrentPageNumber)
                        giNumberOfPrint += 1 '10/04/2013, Ngọc Thoại : id 51975 - Lưu lại số lần in ra giấy
                    Case ReportModeType.lmExport

                        Dim strExportFile As String = "C:\LEMON3\Temp\bad.doc"
                        rpt1.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
                        rpt1.ExportOptions.ExportFormatType = ExportFormatType.WordForWindows

                        Dim objOptions As DiskFileDestinationOptions = New DiskFileDestinationOptions
                        objOptions.DiskFileName = strExportFile
                        rpt1.ExportOptions.DestinationOptions = objOptions
                        rpt1.Export()
                        objOptions = Nothing

                End Select

            End If

            rpt1.Close()

            Exit Sub

        Catch ex As Exception
            MessageBox.Show("Error PrintReport " & vbCrLf & ex.Message, "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End Try

    End Sub

    Public Sub AddParameter(ByVal ParameterName As String, ByVal ParameterValue As Object, Optional ByVal TypeParameter As ReportDataType = ReportDataType.lmReportString)

        Try
            If arrPName Is Nothing Then
                ReDim arrPName(0)
                ReDim arrPValue(0)
            Else
                ReDim Preserve arrPName(UBound(arrPName) + 1)
                ReDim Preserve arrPValue(UBound(arrPValue) + 1)
            End If

            arrPName(UBound(arrPName)) = ParameterName
            Select Case TypeParameter
                Case ReportDataType.lmReportDate
                    arrPValue(UBound(arrPValue)) = CDate(ParameterValue)
                Case ReportDataType.lmReportNumber
                    arrPValue(UBound(arrPValue)) = Val(ParameterValue)
                Case ReportDataType.lmReportString
                    arrPValue(UBound(arrPValue)) = CStr(ParameterValue)
                Case ReportDataType.lmReportBoolean
                    arrPValue(UBound(arrPValue)) = CBool(ParameterValue)
            End Select


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End Try

    End Sub

    Public Sub AddParameterSub(ByVal ParameterName As String, ByVal ParameterValue As Object, Optional ByVal TypeParameter As ReportDataType = ReportDataType.lmReportString)
        Try
            If arrPSubName Is Nothing Then
                ReDim arrPSubName(0)
                ReDim arrPSubValue(0)
            Else
                ReDim Preserve arrPSubName(UBound(arrPSubName) + 1)
                ReDim Preserve arrPSubValue(UBound(arrPSubValue) + 1)
            End If

            arrPSubName(UBound(arrPSubName)) = ParameterName
            Select Case TypeParameter
                Case ReportDataType.lmReportDate
                    arrPSubValue(UBound(arrPSubValue)) = CDate(ParameterValue)
                Case ReportDataType.lmReportNumber
                    arrPSubValue(UBound(arrPSubValue)) = Val(ParameterValue)
                Case ReportDataType.lmReportString
                    arrPSubValue(UBound(arrPSubValue)) = CStr(ParameterValue)
                Case ReportDataType.lmReportBoolean
                    arrPSubValue(UBound(arrPSubValue)) = CBool(ParameterValue)
            End Select

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End Try


    End Sub

    Public Sub AddMain(ByVal MainStrSQL As String)
        gsMainStrData1 = MainStrSQL
    End Sub

    Public Sub AddMain(ByVal dt As DataTable)
        'Minh Hòa update: 20/5/2008
        dtReportMain = dt
    End Sub

    Public Sub AddSub(ByVal SubStrSQL As String, ByVal SubReportName As String)

        Try
            If arrSubSQL Is Nothing Then
                ReDim arrSubSQL(0)
                ReDim arrSubName(0)
            Else
                ReDim Preserve arrSubSQL(UBound(arrSubSQL) + 1)
                ReDim Preserve arrSubName(UBound(arrSubName) + 1)
            End If
            arrSubSQL(UBound(arrSubSQL)) = SubStrSQL
            arrSubName(UBound(arrSubName)) = SubReportName

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End Try

    End Sub

    Public Sub AddSub(ByVal dt As DataTable, ByVal SubReportName As String)
        'Minh Hòa update: 20/5/2008

        Try
            If arrSubTable Is Nothing Then
                ReDim arrSubTable(0)
                ReDim arrSubNameTable(0)
            Else
                ReDim Preserve arrSubTable(UBound(arrSubTable) + 1)
                ReDim Preserve arrSubNameTable(UBound(arrSubNameTable) + 1)
            End If
            arrSubTable(UBound(arrSubTable)) = dt
            arrSubNameTable(UBound(arrSubNameTable)) = SubReportName

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End Try

    End Sub

    Public Sub AddTableTemp(ByVal TableTemp As String, Optional ByVal FlagHostID As Boolean = False)
        'Minh Hòa update: 17/09/2009 lấy tên bảng temp và cờ có gắn HostID
        sTableTemp = New String() {TableTemp}
        bHostID = New Boolean() {FlagHostID}
    End Sub

    Public Sub AddTableTemp(ByVal TableTemp() As String)
        'Minh Hòa update: 17/09/2009 lấy tên bảng temp và cờ có gắn HostID
        sTableTemp = TableTemp
    End Sub

    Public Sub AddTableTemp(ByVal TableTemp() As String, ByVal FlagHostID() As Boolean)
        'Minh Hòa update: 17/09/2009 lấy tên bảng temp và cờ có gắn HostID
        sTableTemp = TableTemp
        bHostID = FlagHostID
    End Sub

    Private Sub Print(formName As String)
        Dim i As Integer
        Dim i1 As Integer
        Dim j As Integer
        Dim j1 As Integer
        'Contain String Error
        Dim StrError As String

        'Main Report ---------------------------------------------------------------
        Try

            If My.Computer.FileSystem.FileExists(gsMainReportName1) = False Then
                'Update 19/10/2010: sửa lỗi Font khi nhập VNI
                If gbUnicode = False Then gsMainReportName1 = ConvertVniToUnicode(gsMainReportName1)
                'If geLanguage = EnumLanguage.Vietnamese Then
                '    'StrError = "Không tồn tại file báo cáo: " & vbCrLf & gsMainReportName1
                '    StrError = "Kh¤ng tän tÁi file bÀo cÀo: " & vbCrLf & gsMainReportName1
                'Else
                '    StrError = "Not exist file report: " & vbCrLf & gsMainReportName1
                'End If
                StrError = rL3("Khong_ton_tai_file_bao_cao") & ": " & vbCrLf & gsMainReportName1
                GoTo Err_Handlling
            Else
                rpt1.Load(gsMainReportName1, CrystalDecisions.[Shared].OpenReportMethod.OpenReportByTempCopy)
            End If

        Catch ex As Exception
            StrError = "Open report error." & vbCrLf & "(" & ex.Message & ")"
            GoTo Err_Handlling
        End Try
        Dim sConnectString As String = gsConnectionString
        If formName <> "" Then ConnectReportServer(formName, sConnectString)
        Try
            Dim dtMain As DataTable
            If gsMainStrData1 <> "" Then 'Truyền vào Main là câu SQL
                dtMain = GetDataTable(gsMainStrData1)
                FilterByDivision(dtMain) 'Chỉ bổ sung khi là SQL, còn tale thì đã qua hàm ReturnDataSet rồi
            Else ' Truyền vào Main là Table
                'Minh Hòa update: 20/5/2008
                dtMain = dtReportMain
            End If
            nRowMaximum = dtMain.Rows.Count
            rpt1.SetDataSource(dtMain)
            ExecuteSQLNoTransaction(" Exec D91P9400", sConnectString) 'Bổ sung 20/04/2016: 86435 : giải phóng file log

        Catch ex As Exception
            StrError = "Set data source error1." & vbCrLf & "(" & ex.Message & ")"
            GoTo Err_Handlling
        End Try

        'Sub Report ---------------------------------------------------------------
        If arrSubName IsNot Nothing Or arrSubNameTable IsNot Nothing Then
            Dim rptSub() As CrystalDecisions.CrystalReports.Engine.ReportDocument
            'Minh Hòa update: 20/5/2008
            If arrSubName IsNot Nothing And arrSubNameTable Is Nothing Then 'Truyền vào Sub là câu SQL
                ReDim rptSub(UBound(arrSubName))
                For i = 0 To UBound(arrSubName)
                    If CheckExistSubReport(arrSubName(i).ToString) = True Then

                        Try
                            rptSub(i) = New CrystalDecisions.CrystalReports.Engine.ReportDocument
                            rptSub(i) = rpt1.OpenSubreport(arrSubName(i).ToString)
                        Catch ex As Exception
                            StrError = "Open sub report <" & rptSub(i).ToString & "> error." & vbCrLf & "(" & ex.Message & ")"
                            GoTo Err_Handlling
                        End Try

                        Try
                            Dim dtSub As DataTable
                            dtSub = GetDataTable(arrSubSQL(i).ToString)
                            rptSub(i).SetDataSource(dtSub)
                        Catch ex As Exception
                            StrError = "Set data source sub error." & vbCrLf & "(" & ex.Message & ")"
                            GoTo Err_Handlling
                        End Try

                    End If
                Next i

            ElseIf arrSubName Is Nothing And arrSubNameTable IsNot Nothing Then 'Truyền vào Sub là Table
                ReDim rptSub(UBound(arrSubNameTable))
                For i = 0 To UBound(arrSubNameTable)
                    If CheckExistSubReport(arrSubNameTable(i).ToString) = True Then

                        Try
                            rptSub(i) = New CrystalDecisions.CrystalReports.Engine.ReportDocument
                            rptSub(i) = rpt1.OpenSubreport(arrSubNameTable(i).ToString)
                        Catch ex As Exception
                            StrError = "Open sub report <" & rptSub(i).ToString & "> error." & vbCrLf & "(" & ex.Message & ")"
                            GoTo Err_Handlling
                        End Try

                        Try
                            Dim dtSub As DataTable
                            dtSub = CType(arrSubTable(i), DataTable)
                            rptSub(i).SetDataSource(dtSub)
                        Catch ex As Exception
                            StrError = "Set data source sub error." & vbCrLf & "(" & ex.Message & ")"
                            GoTo Err_Handlling
                        End Try

                    End If
                Next i

            Else 'Truyền vào Sub vừa là câu SQL vừa là Table
                ReDim rptSub(UBound(arrSubName) + UBound(arrSubNameTable))
                For i = 0 To UBound(arrSubName)
                    If CheckExistSubReport(arrSubName(i).ToString) = True Then

                        Try
                            rptSub(i) = New CrystalDecisions.CrystalReports.Engine.ReportDocument
                            rptSub(i) = rpt1.OpenSubreport(arrSubName(i).ToString)
                        Catch ex As Exception
                            StrError = "Open sub report <" & rptSub(i).ToString & "> error." & vbCrLf & "(" & ex.Message & ")"
                            GoTo Err_Handlling
                        End Try

                        Try
                            Dim dtSub As DataTable
                            dtSub = GetDataTable(arrSubSQL(i).ToString)
                            rptSub(i).SetDataSource(dtSub)
                        Catch ex As Exception
                            StrError = "Set data source sub error." & vbCrLf & "(" & ex.Message & ")"
                            GoTo Err_Handlling
                        End Try

                    End If
                Next i

                For j = 0 To UBound(arrSubNameTable)
                    If CheckExistSubReport(arrSubNameTable(j).ToString) = True Then

                        Try
                            rptSub(i + j) = New CrystalDecisions.CrystalReports.Engine.ReportDocument
                            rptSub(i + j) = rpt1.OpenSubreport(arrSubNameTable(j).ToString)
                        Catch ex As Exception
                            StrError = "Open sub report <" & rptSub(i + j).ToString & "> error." & vbCrLf & "(" & ex.Message & ")"
                            GoTo Err_Handlling
                        End Try

                        Try
                            Dim dtSub As DataTable
                            dtSub = CType(arrSubTable(j), DataTable)
                            rptSub(i + j).SetDataSource(dtSub)
                        Catch ex As Exception
                            StrError = "Set data source sub error." & vbCrLf & "(" & ex.Message & ")"
                            GoTo Err_Handlling
                        End Try

                    End If
                Next j

            End If

            '------------------------------------------------
            'Minh Hòa update: 17/09/2009 Delete bảng temp 
            If sTableTemp IsNot Nothing Then
                Dim sSQL As String = ""
                For i = 0 To sTableTemp.Length - 1
                    sSQL &= vbCrLf & "Delete From " & sTableTemp(i) & "  Where UserID = " & SQLString(gsUserID)
                    If bHostID IsNot Nothing AndAlso bHostID(i) Then
                        sSQL &= " And HostID = " & SQLString(My.Computer.Name)
                    End If
                Next
                If sSQL <> "" Then ExecuteSQLNoTransaction(sSQL)
            End If

            'Sub Parameter ---------------------------------------------------------------
            If arrPSubName IsNot Nothing Then
                Dim pvCollectionSub As New CrystalDecisions.Shared.ParameterValues
                Dim pdvValueIDSub As New CrystalDecisions.Shared.ParameterDiscreteValue

                Try
                    j1 = 0
NextParaSub:
                    For j = j1 To UBound(arrPSubName)
                        pdvValueIDSub.Value = arrPSubValue(i)
                        pvCollectionSub.Add(pdvValueIDSub)
                        rpt1.DataDefinition.ParameterFields(arrPSubName(i).ToString).ApplyCurrentValues(pvCollectionSub)
                        pvCollectionSub.Clear()
                    Next

                Catch ex As Exception
                    j1 = j + 1
                    GoTo NextParaSub
                End Try

            End If

        End If

        'Main Parameter ---------------------------------------------------------------
        If arrPName IsNot Nothing Then
            Dim pvCollection As New CrystalDecisions.Shared.ParameterValues
            Dim pdvValueID As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim iPaperSize As Integer = CrystalDecisions.Shared.PaperSize.PaperA4 'New CrystalDecisions.Shared.PaperSize
            Try
                i1 = 0
NextParaMain:
                For i = i1 To UBound(arrPName)
                    pdvValueID.Value = arrPValue(i)
                    pvCollection.Add(pdvValueID)
                    rpt1.DataDefinition.ParameterFields(arrPName(i).ToString).ApplyCurrentValues(pvCollection)
                    pvCollection.Clear()
                Next
            Catch ex As Exception
                i1 = i + 1
                GoTo NextParaMain
            End Try
        End If

        'Update 27/12/2014
        If gsPrinterName1 <> "" Then 'Nếu có chọn máy thì lấy máy đó ngược lại lấy Default
            rpt1.PrintOptions.PrinterName = gsPrinterName1
            If giPaperSizes > -1 Then rpt1.PrintOptions.PaperSize = giPaperSizes
            'MessageBox.Show("rpt1.PrintOptions.PaperSize = " & rpt1.PrintOptions.PaperSize)

        End If



        Exit Sub

Err_Handlling:

        MessageBox.Show(StrError, rL3("Thong_bao"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
        'Update 17/05/2011: Không dùng thông báo của D99C0008.MsgL3, vì bị lỗi khi chuỗi thông báo quá dài.
        'D99C0008.MsgL3(StrError, L3MessageBoxIcon.Exclamation)
        gbFlagPrint1 = True

    End Sub

    'Private Sub FormPreview_PreviewClick(ByVal bExport As Boolean)

    '    Select Case bExport
    '        Case True
    '            PrintReport(gsMainReportName1, gsMainReportCaption1, ReportModeType.lmExport, gsPrinterName1)
    '        Case False
    '            PrintReport(gsMainReportName1, gsMainReportCaption1, ReportModeType.lmPrint, gsPrinterName1)
    '    End Select

    'End Sub

    Private Function GetDataTable(ByVal Sql As String) As Data.DataTable
        Dim ds As New DataSet

        'Minh Hòa Update 10/08/2012: Đếm số lần bị lỗi
        Dim iCountError As Integer = 0

        Try

ErrorHandles:
            'Modify: MinhHoa 28/02/2008
            '"Data Source=DCSSERVER\SQL2008;Initial Catalog=VOSAFIN;User ID=sa;Password=;Connection Timeout = 0"
            Dim cmd As SqlCommand = New SqlCommand(Sql, gcConPrint)
            'Dim Sda As New SqlDataAdapter(StrSql, gcCon1)
            Dim Sda As New SqlDataAdapter(cmd)
            'cmd.CommandTimeout = 0
            If iCountError > 0 Then 'Minh Hòa Update 10/08/2012: nếu có lỗi trước đó thi gán CommandTimeout = 30
                cmd.CommandTimeout = 30
            Else
                cmd.CommandTimeout = 0
            End If

            Sda.Fill(ds)

            If iCountError > 0 Then
                'Minh Hòa Update 10/08/2012: Nếu có lỗi trước đó thì trả lại thời gian ConnectionTimeout =0
                gcConPrint.ConnectionString = gcConPrint.ConnectionString.Replace(gsConnectionTimeout15, gsConnectionTimeout)
            End If
            Return ds.Tables(0)

        Catch ex As SqlException 'Exception

            '******************************************
            'Minh Hòa Update 10/08/2012: Kiểm tra nếu không kết nối được với server thì thông báo để kết nối lại.
            '            If Err.Number = 10054 OrElse Err.Number = 1231 _
            '            OrElse ex.Message.Contains("Could not open a connection to SQL Server") _
            '            OrElse ex.Message.Contains("The server was not found or was not accessible") _
            '            OrElse ex.Message.Contains("A transport-level error") Then
            '                'OrElse ex.Message.Contains("Login failed") Then 'Lỗi không kết nối được server
            '
            '                If CheckConnectFailed(gcConPrint, iCountError, "FromDataSet", True) Then
            '
            '                    GoTo ErrorHandles
            '                End If
            '            Else
            '                Clipboard.Clear()
            '                Clipboard.SetText(Sql)
            '
            '                MessageBox.Show(Err.Number & " - " & ex.Message & vbCrLf & Sql, "Error Print 1", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            '                Return Nothing
            '            End If
            '******************************************
            'Minh Hòa Update 10/08/2012: Kiểm tra nếu không kết nối được với server thì thông báo để kết nối lại.
            If ex.Number = 10054 OrElse ex.Number = 1231 _
            OrElse ex.Message.Contains("Could not open a connection to SQL Server") _
            OrElse ex.Message.Contains("The server was not found or was not accessible") _
            OrElse ex.Message.Contains("A transport-level error") Then 'Lỗi không kết nối được server
                If CheckConnectFailed(gcConPrint, iCountError, "FromDataSet") Then
                    GoTo ErrorHandles
                End If
            ElseIf ex.Number = 1205 OrElse ex.Message.Contains("deadlocked") Then 'Update 27/01/2015: Lỗi Deadlock
                'If iCountError = 0 Then
                gcConPrint.Close()
                'End If
                iCountError += 1

                WriteLogFile(Sql, "Deadlocked (GetDataTable Print)_" & Now.Year & Now.Month.ToString("00") & Now.Day.ToString("00") & Now.Hour.ToString("00") & Now.Minute.ToString("00") & Now.Second.ToString("00") & ".log")
                'pauses for two second.
                Threading.Thread.Sleep(2000)

                GoTo ErrorHandles
            Else
                gcConPrint.Close()

                Clipboard.Clear()
                Clipboard.SetText(Sql)


                'Update 15/09/2011: Bổ sung kiểm tra mã lỗi từ Server trả ra
                If ex.Number >= 50000 Then 'Mã lỗi thực trả ra từ Server
                    MsgErr(ex.Errors.Item(0).Message)
                Else
                    MessageBox.Show(Err.Number & " - " & ex.Message & vbCrLf & Sql, "Error Print 1", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
                Return Nothing
            End If

            '*****************************************
        End Try

        Return Nothing
    End Function


    'Kiểm tra sự tồn tại của Sub Report trên báo cáo
    Private Function CheckExistSubReport(ByVal sSubNameSource As String) As Boolean
        Dim i As Integer

        For i = 0 To rpt1.Subreports.Count - 1
            If rpt1.Subreports(i).Name = sSubNameSource Then
                Return True
            End If
        Next

        Return False

    End Function

End Class
