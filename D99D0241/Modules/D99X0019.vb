'#######################################################################################
'#                                     Sinh số phiếu theo kiểu mới cho nhóm G4
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày cập nhật cuối cùng: 08/02/2012
'# Diễn giải: 
'# Người cập nhật cuối cùng: Nguyễn Thị Minh Hòa
'#######################################################################################
Public Module D99X0019


#Region "(DÙNG CHO G4) Sinh số phiếu tự động theo kiểu mới, gọi storeD09P7100 , nếu trùng phiếu thì tự động tăng (Chỉ gọi lúc lưu)"

    ''' <summary>
    ''' (DÙNG CHO G4) Insert dữ liệu vào bảng D09T7100 cho trường hợp số phiếu KHÔNG sinh tự động
    ''' </summary>
    ''' <param name="sVoucherNo">Số phiếu hiện tại</param>
    ''' <param name="sVoucherTableName">Bảng nghiệp vụ chứa số phiếu</param>
    ''' <param name="sVoucherIGE">Giá trị khóa chính của bảng nghiệp vụ</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub InsertVoucherNoD91T9111G4(ByVal sVoucherNo As String, ByVal sVoucherTableName As String, ByVal sVoucherIGE As String)
        'Nhóm G4 thay thế table D91T9111 bằng table D09T7100
        Dim sSQL As String
        sSQL = "INSERT INTO D09T7100 (VoucherNo, VoucherIGE, VoucherTableName)" & vbCrLf
        sSQL &= "SELECT " & SQLString(sVoucherNo) & COMMA
        sSQL &= SQLString(sVoucherIGE) & COMMA
        sSQL &= SQLString(sVoucherTableName) & vbCrLf
        sSQL &= "WHERE NOT EXISTS (" & vbCrLf
        sSQL &= "SELECT TOP 1 1 "
        sSQL &= "FROM D09T7100 "
        sSQL &= "WHERE	VoucherNo = " & SQLString(sVoucherNo)
        sSQL &= ")"

        ExecuteSQLNoTransaction(sSQL)
    End Sub

    ''' <summary>
    ''' (DÙNG CHO G4) Xóa số phiếu của bảng D09T7100, Gọi tại form Truy vấn sau khi xóa thành công (Xóa Phiếu)
    ''' </summary>
    ''' <param name="sVoucherNo">Số phiếu khi xóa</param>
    ''' <param name="sVoucherTableName">Bảng nghiệp vụ chứa số phiếu</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub DeleteVoucherNoD91T9111G4(ByVal sVoucherNo As String, ByVal sVoucherTableName As String)
        'Nhóm G4 thay thế table D91T9111 bằng table D09T7100
        Dim sSQL As String = ""
        sSQL &= "Exec D09P7102 "
        sSQL &= SQLString(sVoucherTableName) & COMMA 'VoucherTableName, varchar[50], NOT NULL
        sSQL &= SQLString(sVoucherNo) 'VoucherNo, varchar[20], NOT NULL
        ExecuteSQLNoTransaction(sSQL)

    End Sub

    ''' <summary>
    ''' (DÙNG CHO G4) Xóa số phiếu của bảng D09T7100, Gọi tại form Truy vấn sau khi xóa thành công (Xóa Hóa đơn)
    ''' </summary>
    ''' <param name="sVoucherNo">Số phiếu khi xóa</param>
    ''' <param name="sVoucherTableName">Bảng nghiệp vụ chứa số phiếu</param>
    ''' <param name="FieldNameVoucherNo"> FieldName của Số phiếu (VD: VoucherNo)</param>
    ''' <remarks>Áp dụng cho TH 1 phiếu có nhiều hóa đơn</remarks>
    <DebuggerStepThrough()> _
    Public Sub DeleteVoucherNoD91T9111G4(ByVal sVoucherNo As String, ByVal sVoucherTableName As String, ByVal FieldNameVoucherNo As String)
        'Nhóm G4 thay thế table D91T9111 bằng table D09T7100
        Dim sSQL As String = ""
        sSQL = "IF NOT EXISTS (SELECT TOP 1 1 FROM " & sVoucherTableName & " WHERE " & FieldNameVoucherNo & " = " & SQLString(sVoucherNo) & ")" & vbCrLf
        sSQL &= "Exec D09P7102 "
        sSQL &= SQLString(sVoucherTableName) & COMMA 'VoucherTableName, varchar[50], NOT NULL
        sSQL &= SQLString(sVoucherNo) 'VoucherNo, varchar[20], NOT NULL
        ExecuteSQLNoTransaction(sSQL)
    End Sub


    ''' <summary>
    ''' (DÙNG CHO G4) Sinh số phiếu theo kiểu mới
    ''' </summary>
    ''' <param name="c1Combo">Combo số phiếu</param>
    ''' <param name="sVoucherTableName">Bảng nghiệp vụ chứa số phiếu</param>
    ''' <param name="sVoucherIGE">Giá trị khóa chính của bảng nghiệp vụ</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function CreateIGEVoucherNoNewG4(ByVal c1Combo As C1.Win.C1List.C1Combo, ByVal sVoucherTableName As String, ByVal sVoucherIGE As String) As String
        Return CreateIGEVoucherNoNewG4(sVoucherTableName, sVoucherIGE, L3Int(c1Combo.Columns("S1Type").Text), c1Combo.Columns("S1").Text, L3Int(c1Combo.Columns("S2Type").Text), c1Combo.Columns("S2").Text, L3Int(c1Combo.Columns("S3Type").Text), c1Combo.Columns("S3").Text, L3Int(c1Combo.Columns("OutputOrder").Text), L3Int(c1Combo.Columns("OutputLength").Text), c1Combo.Columns("Separator").Text)

    End Function

    ''' <summary>
    ''' (DÙNG CHO G4)  Sinh số phiếu theo kiểu mới
    ''' </summary>
    ''' <param name="sVoucherTableName">Bảng nghiệp vụ chứa số phiếu</param>
    ''' <param name="sVoucherIGE">Giá trị khóa chính của bảng nghiệp vụ</param>
    ''' <param name="S1Type"></param>
    ''' <param name="s1"></param>
    ''' <param name="S2Type"></param>
    ''' <param name="S2"></param>
    ''' <param name="S3Type"></param>
    ''' <param name="S3"></param>
    ''' <param name="OutputOrder"></param>
    ''' <param name="OutputLength"></param>
    ''' <param name="Seperator"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateIGEVoucherNoNewG4(ByVal sVoucherTableName As String, ByVal sVoucherIGE As String, _
                ByVal S1Type As Integer, ByVal s1 As String, ByVal S2Type As Integer, ByVal S2 As String, ByVal S3Type As Integer, ByVal S3 As String, _
                ByVal OutputOrder As Integer, ByVal OutputLength As Integer, ByVal Seperator As String) As String

        '*********************************
        'Kiểm tra IGE của khóa chính phải có
        If sVoucherIGE = "" Then
            D99C0008.MsgL3("Do vấn đề về kỹ thuật (Khóa chính chưa được tạo) nên việc tạo phiếu bị lỗi." & vbCrLf & "Chương trình kết thúc.", L3MessageBoxIcon.Err)
            WriteLogFile("Loi Sinh so phieu (tao phieu) cua table TableName " & sVoucherTableName, "LogCreateIGEVoucherNoNewG4.log")
            Application.Exit()
        End If
        '*********************************

        Dim strS1 As String = ""
        Dim strS2 As String = ""
        Dim strS3 As String = ""

        If Not IsDBNull(S1Type) AndAlso S1Type <> 0 Then strS1 = FindSxType(S1Type.ToString, s1.Trim)
        If Not IsDBNull(S2Type) AndAlso S2Type <> 0 Then strS2 = FindSxType(S2Type.ToString, S2.Trim)
        If Not IsDBNull(S3Type) AndAlso S3Type <> 0 Then strS3 = FindSxType(S3Type.ToString, S3.Trim)

        Dim sSQL As String = SQLStoreD09P7100(sVoucherIGE, sVoucherTableName, strS1.Trim, strS2.Trim, strS3.Trim, OutputLength, OutputOrder, Seperator.Trim)
        Dim dt As DataTable
        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0).Item("VoucherNo").ToString
        Else
            D99C0008.MsgL3("Do vấn đề về kỹ thuật (Khóa chính chưa được tạo) nên việc tạo phiếu bị lỗi." & vbCrLf & "Chương trình kết thúc.", L3MessageBoxIcon.Err)
            WriteLogFile("Loi Sinh so phieu (tao phieu) cua store D09P7100 " & sSQL, "LogCreateIGEVoucherNoNewG4.log")
            Application.Exit()
        End If

        Return ""
    End Function

    ''' <summary>
    ''' (DÙNG CHO G4) Kiểm tra trùng số phiếu theo kiểu mới
    ''' </summary>
    ''' <param name="ModuleID">Tên module: Dxx</param>
    ''' <param name="TableName">Tên bảng để lưu Số phiếu</param>
    ''' <param name="VoucherID">Giá trị của Khóa chính trong bảng lưu Số phiếu</param>
    ''' <param name="VoucherNo">Số phiếu</param>
    ''' <returns>True: vi phạm, không làm gì cả</returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function CheckDuplicateVoucherNoNewG4(ByVal ModuleID As String, ByVal TableName As String, ByVal VoucherID As String, ByVal VoucherNo As String) As Boolean
        '*********************************
        'Kiểm tra IGE của khóa chính phải có
        If VoucherID = "" Then
            '            D99C0008.MsgL3("Khóa chính của nghiệp vụ này chưa được tạo", L3MessageBoxIcon.Exclamation)
            '            Return True
            D99C0008.MsgL3("Do vấn đề về kỹ thuật (Khóa chính chưa được tạo) nên việc kiểm tra trùng phiếu bị lỗi." & vbCrLf & "Chương trình kết thúc.", L3MessageBoxIcon.Err)
            WriteLogFile("Loi Sinh so phieu (kiem tra trung phieu) cua table TableName " & TableName, "LogCreateIGEVoucherNoNewG4.log")
            Application.Exit()

        End If
        '*********************************
        Dim sSQL As String = ""
        sSQL = SQLStoreD09P7101(ModuleID.Substring(1, 2), TableName, VoucherID, VoucherNo)
        Dim dt As New DataTable
        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            Select Case CInt(dt.Rows(0).Item("Status"))
                Case 1
                    MessageBox.Show(dt.Rows(0).Item("Message").ToString, MsgAnnouncement, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    dt = Nothing
                    Return True
            End Select
        End If
        dt = Nothing
        Return False
    End Function


    <DebuggerStepThrough()> _
  Private Function FindSxType(ByVal nType As String, ByVal s As String) As String
        Select Case nType.Trim
            Case "1" ' Theo tháng
                Return giTranMonth.ToString("00")
            Case "2" ' Theo năm (YYYY)
                Return giTranYear.ToString
            Case "3" ' Theo loại chứng từ
                Return s
            Case "4" ' Theo đơn vị
                Return gsDivisionID
            Case "5" ' Theo hằng
                Return s

                'Modify date: 02/02/2007: bổ sung thêm 3 loại 
            Case "6" ' Theo năm (YY)
                Return giTranYear.ToString.Substring(2, 2)
            Case "7" ' Theo tháng năm (MMYY)
                Return giTranMonth.ToString("00") & giTranYear.ToString.Substring(2, 2)
            Case "8" ' Theo năm tháng (YYMM)
                Return giTranYear.ToString.Substring(2, 2) & giTranMonth.ToString("00")

            Case Else
                Return ""
        End Select
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Description: Sinh số phiếu theo kiểu mới
    '#---------------------------------------------------------------------------------------------------
    <DebuggerStepThrough()> _
    Private Function SQLStoreD09P7100(ByVal VoucherIGE As String, ByVal VoucherTableName As String, ByVal S1 As String, ByVal S2 As String, ByVal S3 As String, ByVal OutputLength As Integer, ByVal OutputOrder As Integer, ByVal Separator As String) As String
        Dim sSQL As String = ""
        sSQL &= "SET NOCOUNT ON " & vbCrLf
        sSQL &= "DECLARE @VoucherNo AS VARCHAR(20) " & vbCrLf
        sSQL &= "Exec D09P7100 "
        sSQL &= SQLString("D91T0001") & COMMA 'TableName, varchar[8], NOT NULL
        sSQL &= SQLString(S1) & COMMA 'StringKey1, varchar[20], NOT NULL
        sSQL &= SQLString(S2) & COMMA 'StringKey2, varchar[20], NOT NULL
        sSQL &= SQLString(S3) & COMMA 'StringKey3, varchar[20], NOT NULL
        sSQL &= SQLNumber(OutputLength) & COMMA 'OutputLen, int, NOT NULL
        sSQL &= SQLNumber(OutputOrder) & COMMA 'OutputOrder, int, NOT NULL
        If Separator <> "" Then
            sSQL &= SQLNumber(1) & COMMA 'Seperated, int, NOT NULL
            sSQL &= SQLString(Separator) & COMMA 'Seperator, char, NOT NULL
        Else
            sSQL &= SQLNumber(0) & COMMA 'Seperated, int, NOT NULL
            sSQL &= SQLString("") & COMMA 'Seperator, char, NOT NULL
        End If
        sSQL &= SQLString("") & COMMA 'Temp1, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'Temp2, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'Temp3, varchar[20], NOT NULL
        sSQL &= SQLString(VoucherIGE) & COMMA 'VoucherIGE, varchar[20], NOT NULL
        sSQL &= SQLString(VoucherTableName) & COMMA  'VoucherTableName, varchar[20], NOT NULL
        sSQL &= " @VoucherNo  OUTPUT " & vbCrLf 'KeyString, varchar[20], NOT NULL
        sSQL &= "SELECT @VoucherNo AS VoucherNo "
        Return sSQL

    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Description: Kiểm tra trùng phiếu theo kiểu mới
    '#---------------------------------------------------------------------------------------------------
    <DebuggerStepThrough()> _
    Private Function SQLStoreD09P7101(ByVal ModuleID As String, ByVal TableName As String, ByVal VoucherID As String, ByVal VoucherNo As String) As String
        Dim sSQL As String = ""
        sSQL &= "Exec D09P7101 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, VarChar[20], NOT NULL
        sSQL &= SQLString(ModuleID) & COMMA 'ModuleID, VarChar[20], NOT NULL
        sSQL &= SQLString(TableName) & COMMA 'TableName, VarChar[20], NOT NULL
        sSQL &= SQLString(VoucherID) & COMMA 'VoucherID, VarChar[20], NOT NULL
        sSQL &= SQLString(VoucherNo) & COMMA 'VoucherNo, VarChar[20], NOT NULL
        sSQL &= SQLString(gsLanguage) 'Language, VarChar[20], NOT NULL
        Return sSQL
    End Function

#End Region

End Module
