'#------------------------------------------------------
'#Title: D99X0003
'#CreateUser: NGUYEN NGOC THANH
'#CreateDate: 24/03/2004
'#ModifiedUser: Hồ Ngọc Thoại
'#ModifiedDate: 10/04/2013
'#Description: Phân quyền In
'# Chứa các biến dùng chung (Print)
'# Sửa lại biến gcCon1 thành gcConPrint để khỏi trùng biến trong dll D99D0041
'#------------------------------------------------------

Imports System.Data.SqlClient

Public Module D99X0003

    'Kết nối khi in báo cáo
    Public gcConPrint As SqlConnection
    'Public gsServerName1 As String
    'Public gsPassword1 As String
    'Public gsDataBase1 As String
    'Public gsUser1 As String

    'Số dòng cho phép export của Excel
    Public Const RowMaximumExcel As Integer = 65528

    Public nRowMaximum As Double

    Public nCountPrint As Integer
    Public bCountPrint As Boolean

    Public rpt1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument

    'Kiểm tra in
    Public gbFlagPrint1 As Boolean

    'Chứa đường dẫn report chính
    Public gsMainReportName1 As String

    'Caption form báo cáo chính
    Public gsMainReportCaption1 As String

    'Nguồn dữ liệu thực thi chính
    Public gsMainStrData1 As String = ""
    'Nguồn dữ liệu thực thi chính
    Public dtReportMain As DataTable
    'Chứa tên máy in
    Public gsPrinterName1 As String
    'Đường dẫn chứa báo cáo tạm
    Public gsSaveFileName1 As String

    'Lưu lại tổng số lần in ra giấy của Hóa đơn
    Public giPrintNumber As Integer = 0
    'Phân quyền nút Print, PrintSetup
    Public giPermissionPrint As Integer = 4
    Public gsConnectionStringReportServer As String = ""
    Public Function GetgcconReportServer() As String
        Dim sConnectionStringReportServer As String = ""
        Dim sConnectionUserReportServer As String
        Dim sPasswordReportServer As String
        Dim sCompanyIDReportServer As String
        Dim sServerNameReportServer As String

        sCompanyIDReportServer = D99C0007.GetOthersSetting("D00", "SystemConfig", "ReportServerDatabase")
        sConnectionUserReportServer = D99C0007.GetOthersSetting("D00", "SystemConfig", "ReportServerUserName")
        sServerNameReportServer = D99C0007.GetOthersSetting("D00", "SystemConfig", "ReportServerName")

        sPasswordReportServer = D00D0041.D00C0001.EncryptString(D99C0007.GetOthersSetting("D00", "SystemConfig", "ReportServerUserPassword", ""), False)

        If sServerNameReportServer = "" Or sConnectionUserReportServer = "" Or sServerNameReportServer = "" Then
            sConnectionStringReportServer = gsConnectionString
        Else
            sConnectionStringReportServer = "Data Source=" & sServerNameReportServer & ";Initial Catalog=" & sCompanyIDReportServer & ";User ID=" & sConnectionUserReportServer & ";Password=" & sPasswordReportServer & ";Application Name=DigiNet ERP Desktop"
        End If

        Return sConnectionStringReportServer
    End Function
    Public Function GetConfigReportServer() As Boolean
        'Khai bÀo cho Report Server
        Dim gsConnectionUserReportServer As String
        Dim gsPasswordReportServer As String
        Dim gsCompanyIDReportServer As String
        Dim gsServerNameReportServer As String
        Dim sSQL As String = ""
        sSQL = "SELECT * FROM D91T2502 WITH (NOLOCK)"
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            With dt.Rows(0)
                gsServerNameReportServer = L3String(.Item("ReportServerName"))
                gsCompanyIDReportServer = L3String(.Item("ReportServerDatabase"))
                gsConnectionUserReportServer = L3String(.Item("ReportServerUserName"))
                gsPasswordReportServer = D00C0001.EncryptString(L3String(.Item("ReportServerUserPassword")), False)
            End With
        Else
            '++++++++++Lay Report Server++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            gsCompanyIDReportServer = GetSetting("D00", "SystemConfig", "ReportServerDatabase")
            gsConnectionUserReportServer = GetSetting("D00", "SystemConfig", "ReportServerUserName")
            gsServerNameReportServer = GetSetting("D00", "SystemConfig", "ReportServerName")
            gsPasswordReportServer = GetSetting("D00", "SystemConfig", "ReportServerUserPassword")
        End If

        If gsCompanyIDReportServer = "" Or gsConnectionUserReportServer = "" Or gsServerNameReportServer = "" Then
            gsConnectionStringReportServer = gsConnectionString
        Else
            gsConnectionStringReportServer = "Data Source=" & gsServerNameReportServer & ";Initial Catalog=" & gsCompanyIDReportServer & ";User ID = " & gsConnectionUserReportServer & ";Password=" & gsPasswordReportServer & ";Application Name=DigiNet ERP Desktop"
        End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        If gsConnectionStringReportServer = "" Then
            Return False
        End If
        Return True
    End Function

    Public Function ConnectReportServer(FormID As String, ByRef ConnectString As String) As Boolean
        If FormID = "" Then Return False 'Bổ sung khi tính năng này chép lên
        'If L3Bool(ReturnScalar("Select Top 1 1 from D89T5000 where IsReportServer =1 AND FormID =" & SQLString(FormID))) Then
        '    'Kiểm tra dữ liệu có tồn tại báo cáo này không?
        '    ConnectString = GetgcconReportServer() 'gsConnectionString 'If bReportServer Then sConnectionString = GetgcconReportServer()
        '    If ConnectString <> "" AndAlso ConnectString <> gsConnectionString Then
        '        gcConPrint = New SqlConnection(ConnectString)
        '        Return True
        '    End If
        'End If
        Return False
    End Function
End Module
