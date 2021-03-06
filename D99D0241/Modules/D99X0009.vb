'#######################################################################################
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày cập nhật cuối cùng: 10/4/2014
'# Người cập nhật cuối cùng: Nguyễn Thị Ánh
'# Diễn giải: Các hàm chung của nhóm G4
'Bổ sung Điều điện IsUseD15/IsUseD21=1 cho combo Nhân viên
'Bổ sung thêm "." vào hàm IndexIdCharactor và đổ nguồn Nhóm nhân viên - EmpGroupID
'Bổ sung thêm đổ nguồn Nhóm nhân viên cho dropdown - EmpGroupID
'Bổ sung thêm ký tự đặc biệt của hàm IndexIdCharactor
'Bổ sung đổ nguồn dropdown Khối, Phòng ban, Tổ nhóm, Nhân viên
'Bổ sung Nhóm nhân viên theo Đơn vị, Khối
'Sửa lại câu đổ nguồn theo Tiếng Anh: Khối, Phòng ban, Tổ nhóm, Nhóm nhân viên
'# Bổ sung WITH (NOLOCK) vào table, trong bảng D91T0000 23/9/2013
'# Bổ sung hàm lấy giá trị Hồ sơ tháng: GetPayRollVoucherID 09/12/2013 sửa lại lấy theo store
'#Bổ sung Tính năng phân quyền theo phòng ban 10/1/2014
'#Bổ sung Nhóm nhân viên theo D09P6868 10/4/2014
'#######################################################################################

Public Module D99X0009
    Public gsCreateByG4 As String = "" 'Lấy giá trị mặc định của Người tạo
    Public gbLockL3UserIDG4 As Boolean = False

#Region "General Function of G4"

#Region "Câu đổ nguồn cho Subreport G4"
    ''' <summary>
    ''' Câu SQL load SubReport cho G4 (có Combo Đơn vị hiện tại)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function SQLSubReportG4(Optional ByVal sDivisionID As String = "") As String
        If sDivisionID = "" Then sDivisionID = gsDivisionID
        Dim sSQL As String
        sSQL = "Select CompanyName  as  Company, CompanyAddress as  Address, CompanyPhone  as  Telephone, CompanyFax  as  Fax, BankAccountName as BankAccountName, BankAccountNo,  VATCode" & vbCrLf
        sSQL &= " FROM D91V0016 " & vbCrLf
        sSQL &= " WHERE   DivisionID = " & SQLString(sDivisionID)

        Return sSQL
    End Function
#End Region

#Region "Đơn vị"
    Public Function ReturnTableDivisionIDD09(ByVal sModuleID As String, Optional ByVal bHavePercent As Boolean = False, Optional ByVal bUseUnicode As Boolean = False) As DataTable
        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************
        Dim sSQL As String = "--Do nguon Don vi" & vbCrLf
        sSQL &= "Select Distinct T99.DivisionID as DivisionID, "
        sSQL &= "T16.DivisionName" & sUnicode & " as DivisionName,"
        sSQL &= " 1 as DisplayOrder" & vbCrLf
        sSQL &= " From " & sModuleID & "T9999 T99 WITH(NOLOCK) Inner Join D91T0016 T16 WITH(NOLOCK) On T99.DivisionID = T16.DivisionID " & vbCrLf
        ' Nếu là nhóm G4 thì lấy table D91T0061
        sSQL &= " Inner Join D91T0061 T60 WITH(NOLOCK) On T99.DivisionID = T60.DivisionID " & vbCrLf
        sSQL &= " Where T16.Disabled = 0 And T60.UserID = '" & gsUserID & "'" & vbCrLf
        If bHavePercent Then
            sSQL &= " Union All " & vbCrLf
            sSQL &= " Select '%' as DivisionID," & sLanguage & " as DivisionName ,0 as DisplayOrder" & vbCrLf
        End If
        sSQL &= " Order By DisplayOrder, DivisionName"

        Dim dt As DataTable
        dt = ReturnDataTable(sSQL)
        Return dt
    End Function

    ''' <summary>
    ''' Load data of combo DivisionID in Inquiry/Transaction/List (no Report) 
    ''' </summary>
    ''' <param name="tdbcDivisionID">combo DivisionID</param>
    ''' <param name="sModuleID">Dxx</param>
    ''' <param name="bHavePercent">True: contain %; False(Default): no contain %</param>
    ''' <remarks></remarks>
    ''' 
    Public Sub LoadCboDivisionIDD09(ByVal tdbcDivisionID As C1.Win.C1List.C1Combo, ByVal sModuleID As String, Optional ByVal bHavePercent As Boolean = False, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbcDivisionID, ReturnTableDivisionIDD09(sModuleID, bHavePercent, bUseUnicode), bUseUnicode)
        tdbcDivisionID.DisplayMember = "DivisionName"
        tdbcDivisionID.ValueMember = "DivisionID"
    End Sub


    ''' <summary>
    ''' Load data of combo DivisionID in Report
    ''' </summary>
    ''' <param name="tdbc">Combo DivisionID</param>
    ''' <param name="sModuleID">ModuleID (Dxx)</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadCboDivisionIDReportD09(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal sModuleID As String, Optional ByVal bUseUnicode As Boolean = False)

        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************

        Dim sSQL As String = ""
        sSQL = "Select Distinct T99.DivisionID as DivisionID, T16.DivisionName" & sUnicode & " as DivisionName,1 as DisplayOrder "
        sSQL &= " From " & sModuleID & "T9999 T99 WITH(NOLOCK) Inner Join D91T0016 T16 WITH(NOLOCK) On T99.DivisionID = T16.DivisionID "
        ' Nếu là nhóm G4 thì lấy table D91T0061
        sSQL &= " Inner Join D91T0061 T60 WITH(NOLOCK) On T99.DivisionID = T60.DivisionID "
        sSQL &= " Where T16.Disabled = 0 And T60.UserID = '" & gsUserID & "' "
        If giModuleAdmin = 1 Then
            sSQL &= " Union All " & vbCrLf
            sSQL &= " Select '%' as DivisionID," & sLanguage & " as DivisionName,0 as DisplayOrder " & vbCrLf
        End If
        sSQL &= " Order By DisplayOrder, DivisionName"

        LoadDataSource(tdbc, sSQL, bUseUnicode)

        tdbc.DisplayMember = "DivisionName"
        tdbc.ValueMember = "DivisionID"
    End Sub
#End Region

#Region "Khối"

    'Load tdbcBlockID: Khối
    Public Function ReturnTableBlockID(Optional ByVal bCurDivision As Boolean = False, Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False) As DataTable

        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************

        Dim sSQL As String = "--Do nguon Khoi" & vbCrLf
        If bHavePercent Then
            'sSQL = "SELECT 	'%' As BlockID, '<" & rl3("Tat_ca") & ">' As BlockName, '%' As DivisionID,0 as DisplayOrder" & vbCrLf
            sSQL &= "SELECT 	'%' As BlockID, " & sLanguage & " As BlockName, '%' As DivisionID,  0 as BlockDisplayOrder, 0 as DisplayOrder" & vbCrLf
            sSQL &= "UNION" & vbCrLf
        End If
        sSQL &= "SELECT BlockID, BlockName" & IIf(geLanguage = EnumLanguage.English, "01", "").ToString & sUnicode & " As BlockName, DivisionID, BlockDisplayOrder, 1 as DisplayOrder" & vbCrLf
        sSQL &= "FROM 	D09T1140 WITH(NOLOCK) WHERE	Disabled = 0 " & vbCrLf
        If bCurDivision Then sSQL &= " And DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        sSQL &= "ORDER BY 	DisplayOrder, BlockDisplayOrder, BlockName"
        Return ReturnDataTable(sSQL)
    End Function

    'Load tdbcBlockID by Current DivisionID: no have tdbcDivisionID
    Public Sub LoadtdbcBlockID(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, ReturnTableBlockID(True, , bUseUnicode), bUseUnicode) 'Load by Current Division
        ' tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
        tdbc.DisplayMember = "BlockName"
        tdbc.ValueMember = "BlockID"
    End Sub

    'Load tdbcBlockID by tdbcDivisionID
    Public Sub LoadtdbcBlockID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        If sDivisionID = "%" Then
            LoadDataSource(tdbc, dtOriginal.Copy, bUseUnicode)
        Else
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DivisionID='%' or DivisionID =" & SQLString(sDivisionID), True), bUseUnicode)
        End If
        Try
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False 'Không cần design
        Catch ex As Exception

        End Try

        tdbc.DisplayMember = "BlockName"
        tdbc.ValueMember = "BlockID"
    End Sub

    Public Sub LoadtdbdBlockID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, ReturnTableBlockID(True, False, bUseUnicode), bUseUnicode) 'Load by Current Division
    End Sub

    Public Sub LoadtdbdBlockID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sDivisionID}
        Dim sField() As String = {"DivisionID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), bUseUnicode)
    End Sub

#Region "Load Khối theo D09P6868"

    Public Function ReturnTableBlockID_D09P6868(ByVal DivisionID As String, ByVal FormID As String, ByVal IsMSS As Integer) As DataTable
        Return ReturnDataTable(SQLStoreD09P6868(DivisionID, FormID, "Block", IsMSS, "Do nguon Khoi"))
    End Function

    Public Sub LoadtdbcBlockID_D09P6868(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal DivisionID As String, ByVal FormID As String, ByVal IsMSS As Integer)
        LoadDataSource(tdbc, ReturnTableBlockID_D09P6868(DivisionID, FormID, IsMSS), gbUnicode) 'Load by Current Division
        tdbc.DisplayMember = "BlockName"
        tdbc.ValueMember = "BlockID"
    End Sub

    Public Sub LoadtdbdBlockID_D09P6868(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal DivisionID As String, ByVal FormID As String, ByVal IsMSS As Integer)
        LoadDataSource(tdbd, ReturnTableFilter(ReturnTableBlockID_D09P6868(DivisionID, FormID, IsMSS), "BlockID<>'%'", True), gbUnicode) 'Load by Current Division
    End Sub
#End Region
#End Region

#Region "Nhóm phòng ban"
    Public Function ReturnTableDepartmentGroupIDNew() As DataTable
        Return ReturnDataTable(SQLStoreD09P2525("DepartmentGroupID", True))
    End Function

#End Region

#Region "Phòng ban"
    Public Function ReturnTableDepartmentID_D09P6868(ByVal DivisionID As String, ByVal FormID As String, ByVal IsMSS As Integer) As DataTable
        Return ReturnDataTable(SQLStoreD09P6868(DivisionID, FormID, "Department", IsMSS, "Do nguon Phong ban"))
    End Function

    '************************

    'Load tdbcDepartmentID: Phòng ban
    'bCurDivision=True: by Current DivisionID
    'bCurDivision=False: by tdbcDivisionID
    Public Function ReturnTableDepartmentID(Optional ByVal bCurDivision As Boolean = False, Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False) As DataTable
        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************
        Dim sSQL As String = "--Do nguon Phong ban" & vbCrLf
        If bHavePercent Then
            sSQL &= "SELECT 	'%' As DepartmentID, " & sLanguage & " As DepartmentName, '%' As DivisionID, '%' As BlockID,0 as DepDisplayOrder, 0 as DisplayOrder " & vbCrLf
            sSQL &= "UNION" & vbCrLf
        End If
        sSQL &= "SELECT DepartmentID, DepartmentName" & IIf(geLanguage = EnumLanguage.English, "01", "").ToString & sUnicode & " As DepartmentName, DivisionID,BlockID, DepDisplayOrder, 1 as DisplayOrder" & vbCrLf
        sSQL &= "FROM 	D91T0012 WITH(NOLOCK) WHERE	Disabled = 0 " & vbCrLf
        If bCurDivision Then sSQL &= " And DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        sSQL &= "ORDER BY DisplayOrder, DepDisplayOrder, DepartmentName"

        Return ReturnDataTable(sSQL)
    End Function

    'ID 103797 11/10/2017 Bổ sung thêm Nhóm Phòng ban trong câu Đổ nguồn PN,thay câu Select =Store D09P2525
    Public Function ReturnTableDepartmentID_HaveGroupID(Optional ByVal bCurDivision As Boolean = False, Optional ByVal bHavePercent As Boolean = True) As DataTable
        Dim dtSource As DataTable
        dtSource = ReturnDataTable(SQLStoreD09P2525("DepartmentID", bHavePercent))
        If bCurDivision Then
            Return ReturnTableFilter(dtSource, "DivisionID=" & SQLString(gsDivisionID), True)
        End If
        Return dtSource
    End Function

    Public Sub LoadtdbcDepartmentID_HaveGroupID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentGroupID As String, Optional ByVal bUseUnicode As Boolean = False)
        If sDepartmentGroupID = "%" And sBlockID = "%" Then 'No Filter
            LoadDataSource(tdbc, dtOriginal.Copy, bUseUnicode)
        ElseIf sDepartmentGroupID = "%" And sBlockID <> "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or BlockID =" & SQLString(sBlockID), True), bUseUnicode)
        ElseIf sDepartmentGroupID <> "%" And sBlockID = "%" Then 'Filter by DivisionID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or DepartmentGroupID =" & SQLString(sDepartmentGroupID), True), bUseUnicode)
        Else
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or (DepartmentGroupID =" & SQLString(sDepartmentGroupID) & " And BlockID =" & SQLString(sBlockID) & ")", True), bUseUnicode)
        End If
        Try
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentGroupID").Visible = False
        Catch ex As Exception

        End Try

        tdbc.DisplayMember = "DepartmentName"
        tdbc.ValueMember = "DepartmentID"
    End Sub

    Public Sub LoadtdbcDepartmentID_HaveGroupID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDivisionID As String, ByVal sDepartmentGroupID As String, Optional ByVal bUseUnicode As Boolean = False)
        If sDepartmentGroupID = "%" And sBlockID = "%" And sDivisionID = "%" Then 'No Filter
            LoadDataSource(tdbc, dtOriginal.Copy, bUseUnicode)
        ElseIf sDepartmentGroupID = "%" And sBlockID = "%" And sDivisionID <> "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or DivisionID =" & SQLString(sDivisionID), True), bUseUnicode)

        ElseIf sDepartmentGroupID = "%" And sBlockID <> "%" And sDivisionID = "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or BlockID =" & SQLString(sBlockID), True), bUseUnicode)

        ElseIf sDepartmentGroupID <> "%" And sBlockID = "%" And sDivisionID = "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or DepartmentGroupID =" & SQLString(sDepartmentGroupID), True), bUseUnicode)

        ElseIf sDepartmentGroupID = "%" And sBlockID <> "%" And sDivisionID <> "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or (BlockID =" & SQLString(sBlockID) & " and DivisionID =" & SQLString(sDivisionID) & ")", True), bUseUnicode)

        ElseIf sDepartmentGroupID <> "%" And sBlockID <> "%" And sDivisionID = "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or (BlockID =" & SQLString(sBlockID) & " and DepartmentGroupID =" & SQLString(sDepartmentGroupID) & ")", True), bUseUnicode)

        ElseIf sDepartmentGroupID <> "%" And sBlockID = "%" And sDivisionID <> "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or (DepartmentGroupID =" & SQLString(sDepartmentGroupID) & " and DivisionID =" & SQLString(sDivisionID) & ")", True), bUseUnicode)

        Else
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or (DepartmentGroupID =" & SQLString(sDepartmentGroupID) & " And BlockID =" & SQLString(sBlockID) & " and DivisionID =" & SQLString(sDivisionID) & ")", True), bUseUnicode)
        End If
        Try
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentGroupID").Visible = False
        Catch ex As Exception

        End Try

        tdbc.DisplayMember = "DepartmentName"
        tdbc.ValueMember = "DepartmentID"
    End Sub

    Public Sub LoadtdbdDepartmentID_HaveGroupID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentGroupID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentGroupID}
        Dim sField() As String = {"BlockID", "DepartmentGroupID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), bUseUnicode)
    End Sub

    Public Sub LoadtdbdDepartmentID_HaveGroupID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDivisionID As String, ByVal sDepartmentGroupID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDivisionID, sDepartmentGroupID}
        Dim sField() As String = {"BlockID", "DivisionID", "DepartmentGroupID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load tdbcDepartmentID by BlockID and gsDivisionID 
    ''' </summary>
    ''' <param name="tdbc">tdbcDepartmentID</param>
    ''' <param name="dtOriginal">get value from ReturnTableDepartmentID(True) function</param>
    ''' <param name="sBlockID">value of tdbcBlockID</param>
    ''' <remarks></remarks>
    Public Sub LoadtdbcDepartmentID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, Optional ByVal bUseUnicode As Boolean = False)
        If sBlockID = "%" Then
            LoadDataSource(tdbc, dtOriginal.Copy, bUseUnicode)
        Else
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or BlockID =" & SQLString(sBlockID)), bUseUnicode)
        End If
        'tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
        Try
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
        Catch ex As Exception

        End Try

        tdbc.DisplayMember = "DepartmentName"
        tdbc.ValueMember = "DepartmentID"
    End Sub
    ''' <summary>
    ''' Load tdbcDepartmentID by BlockID and tdbcDivisionID 
    ''' </summary>
    ''' <param name="tdbc">tdbcDepartmentID</param>
    ''' <param name="dtOriginal">get value from ReturnTableDepartmentID() function</param>
    ''' <param name="sBlockID">tdbcBlockID</param>
    ''' <param name="sDivisionID">tdbcDivisionID</param>
    ''' <remarks></remarks>
    Public Sub LoadtdbcDepartmentID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        If sDivisionID = "%" And sBlockID = "%" Then 'No Filter
            LoadDataSource(tdbc, dtOriginal.Copy, bUseUnicode)
        ElseIf sDivisionID = "%" And sBlockID <> "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or BlockID =" & SQLString(sBlockID), True), bUseUnicode)
        ElseIf sDivisionID <> "%" And sBlockID = "%" Then 'Filter by DivisionID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or DivisionID =" & SQLString(sDivisionID), True), bUseUnicode)
        Else
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or (DivisionID =" & SQLString(sDivisionID) & " And BlockID =" & SQLString(sBlockID) & ")", True), bUseUnicode)
        End If
        Try
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "DepartmentName"
        tdbc.ValueMember = "DepartmentID"
    End Sub

    Public Sub LoadtdbdDepartmentID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID}
        Dim sField() As String = {"BlockID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), bUseUnicode)
    End Sub

    Public Sub LoadtdbdDepartmentID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDivisionID}
        Dim sField() As String = {"BlockID", "DivisionID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)
    End Sub
#End Region

#Region "Tổ nhóm"
    Public Function ReturnTableTeamID_D09P6868(ByVal DivisionID As String, ByVal FormID As String, ByVal IsMSS As Integer) As DataTable
        Return ReturnDataTable(SQLStoreD09P6868(DivisionID, FormID, "Team", IsMSS, "Do nguon To nhom"))
    End Function

    'Load tdbcTeamID: Tổ nhóm
    'bCurDivision=True: by Current DivisionID
    'bCurDivision=False: by tdbcDivisionID
    Public Function ReturnTableTeamID(Optional ByVal bCurDivision As Boolean = False, Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False) As DataTable
        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************

        Dim sSQL As String = "--Do nguon To nhom" & vbCrLf

        If bHavePercent Then
            sSQL &= "SELECT 	'%' As TeamID, " & sLanguage & " As TeamName, '%' As DivisionID , '%' As BlockID, '%' As DepartmentID, 0 As TeamDisplayOrder, 0 As DisplayOrder " & vbCrLf
            sSQL &= "UNION" & vbCrLf
        End If
        sSQL &= "SELECT  D01.TeamID, D01.TeamName" & IIf(geLanguage = EnumLanguage.English, "01", "").ToString & sUnicode & " as TeamName, D02.DivisionID, D02.BlockID, D01.DepartmentID, TeamDisplayOrder, 1 As DisplayOrder" & vbCrLf
        sSQL &= "FROM 	D09T0227 D01 WITH(NOLOCK) " & vbCrLf
        sSQL &= "INNER JOIN D91T0012 D02 WITH(NOLOCK) On D02.DepartmentID = D01.DepartmentID" & vbCrLf
        sSQL &= "WHERE	D01.Disabled = 0 " & vbCrLf
        If bCurDivision Then sSQL &= " And D02.DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        sSQL &= "ORDER BY  DisplayOrder,TeamDisplayOrder, TeamName"
        Return ReturnDataTable(sSQL)
    End Function


    ''' <summary>
    ''' Load tdbcTeamID by BlockID and gsDivisionID 
    ''' </summary>
    ''' <param name="tdbc">tdbcTeamID</param>
    ''' <param name="dtOriginal">get value from ReturnTableTeamID(True) function</param>
    ''' <param name="sBlockID">value of tdbcBlockID</param>
    ''' <remarks></remarks>
    Public Sub LoadtdbcTeamID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, Optional ByVal bUseUnicode As Boolean = False)
        If sDepartmentID = "%" And sBlockID = "%" Then 'No Filter
            LoadDataSource(tdbc, dtOriginal.Copy, bUseUnicode)
        ElseIf sDepartmentID = "%" And sBlockID <> "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or BlockID =" & SQLString(sBlockID), True), bUseUnicode)
        ElseIf sDepartmentID <> "%" And sBlockID = "%" Then 'Filter by Department
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or DepartmentID =" & SQLString(sDepartmentID), True), bUseUnicode)
        Else
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or (DepartmentID =" & SQLString(sDepartmentID) & " And BlockID =" & SQLString(sBlockID) & ")", True), bUseUnicode)
        End If
        'tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
        Try
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "TeamName"
        tdbc.ValueMember = "TeamID"
    End Sub

    Public Sub LoadtdbdTeamID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID}
        Dim sField() As String = {"BlockID", "DepartmentID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load tdbcTeamID by BlockID and tdbcDivisionID 
    ''' </summary>
    ''' <param name="tdbc">tdbcTeamID</param>
    ''' <param name="dtOriginal">get value from ReturnTableTeamID() function</param>
    ''' <param name="sBlockID">tdbcBlockID</param>
    ''' <param name="sDivisionID">tdbcDivisionID</param>
    ''' <remarks></remarks>
    Public Sub LoadtdbcTeamID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)

        If sDivisionID = "%" And sBlockID = "%" And sDepartmentID = "%" Then 'No Filter
            LoadDataSource(tdbc, dtOriginal.Copy, bUseUnicode)

        ElseIf sDivisionID = "%" And sBlockID <> "%" And sDepartmentID = "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or BlockID =" & SQLString(sBlockID), True), bUseUnicode)

        ElseIf sDivisionID = "%" And sBlockID <> "%" And sDepartmentID <> "%" Then 'Filter by BlockID and DepartmentID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or (DepartmentID =" & SQLString(sDepartmentID) & " And BlockID =" & SQLString(sBlockID) & ")", True), bUseUnicode)

        ElseIf sDivisionID = "%" And sBlockID = "%" And sDepartmentID <> "%" Then 'Filter by DepartmentID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or DepartmentID =" & SQLString(sDepartmentID), True), bUseUnicode)

        ElseIf sDivisionID <> "%" And sBlockID = "%" And sDepartmentID = "%" Then 'Filter by DivisionID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or DivisionID =" & SQLString(sDivisionID), True), bUseUnicode)

        ElseIf sDivisionID <> "%" And sBlockID = "%" And sDepartmentID <> "%" Then 'Filter by DivisionID and DepartmentID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or (DepartmentID =" & SQLString(sDepartmentID) & " And DivisionID =" & SQLString(sDivisionID) & ")", True), bUseUnicode)

        ElseIf sDivisionID <> "%" And sBlockID <> "%" And sDepartmentID = "%" Then 'Filter by DivisionID and BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or (DivisionID =" & SQLString(sDivisionID) & " And BlockID =" & SQLString(sBlockID) & ")", True), bUseUnicode)
        Else 'Filter by DivisionID and BlockID and DepartmentID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or (DivisionID =" & SQLString(sDivisionID) & " And BlockID =" & SQLString(sBlockID) & " And DepartmentID =" & SQLString(sDepartmentID) & ")", True), bUseUnicode)
        End If
        Try
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "TeamName"
        tdbc.ValueMember = "TeamID"
    End Sub


    Public Sub LoadtdbdTeamID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sDivisionID}
        Dim sField() As String = {"BlockID", "DepartmentID", "DivisionID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), bUseUnicode)
    End Sub
#End Region

#Region "Hình thức làm việc"
    'Load tdbcWorkingStatusID: Hình thức làm việc
    Public Function ReturnTableWorkingStatusID(Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False) As DataTable
        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************

        Dim sSQL As String = "--Do nguon Hinh thuc lam viec" & vbCrLf
        If bHavePercent Then
            sSQL &= "SELECT 	'%' As WorkingStatusID, " & sLanguage & "  As WorkingStatusName,0 as DisplayOrder" & vbCrLf
            sSQL &= "UNION" & vbCrLf
        End If
        sSQL &= "SELECT  WorkingStatusID, WorkingStatusName" & sUnicode & " as WorkingStatusName,1 as DisplayOrder" & vbCrLf
        sSQL &= "FROM 	D09T0070 WITH(NOLOCK) " & vbCrLf
        sSQL &= "WHERE	Disabled = 0 " & vbCrLf
        sSQL &= "ORDER BY 	DisplayOrder,WorkingStatusName"

        Return ReturnDataTable(sSQL)
    End Function

    'Load tdbcWorkingStatusID 
    Public Sub LoadtdbcWorkingStatusID(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, ReturnTableWorkingStatusID(bHavePercent, bUseUnicode), bUseUnicode) 'Load by Current Division
        tdbc.DisplayMember = "WorkingStatusName"
        tdbc.ValueMember = "WorkingStatusID"
    End Sub

    Public Sub LoadtdbdWorkingStatusID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, ReturnTableWorkingStatusID(False, bUseUnicode), bUseUnicode)
    End Sub
#End Region

#Region "Nhân viên"

#Region "Load Nhân viên theo kỳ"
    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD09P6262
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 24/08/2015 09:47:46
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD09P6262(ByVal FormID As String, ByVal DivisionID As String, ByVal TranMonth As Integer, ByVal TranYear As Integer) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Load nhan vien theo ky" & vbCrLf)
        sSQL &= "Exec D09P6262 "
        sSQL &= SQLString(DivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLString(FormID) & COMMA 'FormID, varchar[50], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[50], NOT NULL
        sSQL &= SQLNumber(TranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(TranYear) 'TranYear, int, NOT NULL
        Return sSQL
    End Function

    ''' <summary>
    ''' Load nhân viên theo kỳ truyền vào
    ''' </summary>
    ''' <param name="sFormID">FormID trong D09P6262</param>
    ''' <param name="DivisionID">Đơn vị</param>
    ''' <param name="TranMonth"></param>
    ''' <param name="TranYear"></param>
    ''' <remarks></remarks>
    Public Function ReturnTableEmployeeIDByPeriod(ByVal sFormID As String, ByVal DivisionID As String, ByVal TranMonth As Integer, ByVal TranYear As Integer) As DataTable
        Return ReturnDataTable(SQLStoreD09P6262(sFormID, DivisionID, TranMonth, TranYear))
    End Function

    ''' <summary>
    ''' Load nhân viên theo Đơn vị và kỳ hiện tại
    ''' </summary>
    ''' <param name="sFormID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnTableEmployeeIDByPeriod(ByVal sFormID As String) As DataTable
        Return ReturnTableEmployeeIDByPeriod(sFormID, gsDivisionID, giTranMonth, giTranYear)
    End Function
#End Region

    Public Function ReturnTableEmployeeID_D09P6868(ByVal DivisionID As String, ByVal FormID As String, ByVal IsMSS As Integer) As DataTable
        Return ReturnDataTable(SQLStoreD09P6868(DivisionID, FormID, "Employee", IsMSS, "Do nguon Nhan vien"))
    End Function

    ''' <summary>
    ''' Table dữ liệu NV không ĐK Disabled = 0
    ''' </summary>
    ''' <param name="bCurDivision">gsDivisionID</param>
    ''' <param name="bHavePercent">%</param>
    ''' <param name="bUseUnicode">gbUnicode</param>
    ''' <param name="sModuleID">Dxx</param>
    ''' <returns>DataTable D09T0201</returns>
    ''' <remarks></remarks>
    Public Function ReturnTableEmployeeID_Inquiry(Optional ByVal bCurDivision As Boolean = False, Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False, Optional ByVal sModuleID As String = "") As DataTable
        Return ReturnDataTable(SQLEmployeID(bCurDivision, bHavePercent, bUseUnicode, sModuleID, False))
    End Function

    Private Function SQLEmployeID(Optional ByVal bCurDivision As Boolean = False, Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False, Optional ByVal sModuleID As String = "", Optional ByVal bAndDisabled As Boolean = True)
        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************
        Dim sSQL As String = "--Do nguon Nhan vien" & vbCrLf
        If bHavePercent Then
            sSQL &= "SELECT     '%' As EmployeeID, " & sLanguage & "  As EmployeeName, '%' As DivisionID , '%' As BlockID, '%' As DepartmentID, '%' as  TeamID, '%' as  WorkingStatusID, '%' as EmpGroupID, 0 as EmpDisplayOrder, 0 as DisplayOrder" & vbCrLf
            sSQL &= "UNION" & vbCrLf
        End If
        sSQL &= "SELECT     D01.EmployeeID,Isnull(D01.LastName" & sUnicode & ",'') + CASE WHEN  D01.MiddleName" & sUnicode & " ='' THEN '' ELSE ' ' + D01.MiddleName" & sUnicode & " END  + ' '+ Isnull(D01.FirstName" & sUnicode & ",'') as EmployeeName, D01.DivisionID, D02.BlockID,  D01.DepartmentID, D01.TeamID, D01.WorkingTypeID AS WorkingStatusID, D01.EmpGroupID, EmpDisplayOrder, 1 as DisplayOrder  " & vbCrLf
        sSQL &= "FROM     D09T0201 D01 WITH (NOLOCK) " & vbCrLf
        sSQL &= "INNER JOIN D91T0012 D02  WITH(NOLOCK) ON D01.DepartmentID = D02.DepartmentID" & vbCrLf
        '*************************
        'ID 86605 27/04/2016
        'sSQL &= "WHERE    D01.Disabled = 0 " & vbCrLf
        'If bCurDivision Then sSQL &= " And D02.DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        'If sModuleID.Contains("21") OrElse sModuleID.Contains("15") Then sSQL &= " And D01.IsUseD" & Microsoft.VisualBasic.Right(sModuleID, 2) & " = 1" & vbCrLf
        Dim sWhere As String = ""
        If bAndDisabled Then sWhere &= " D01.Disabled = 0 " & vbCrLf
        If bCurDivision Then sWhere &= IIf(sWhere = "", "", " And ").ToString & "D02.DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        If sModuleID.Contains("21") OrElse sModuleID.Contains("15") Then sWhere &= IIf(sWhere = "", "", " And ").ToString & "D01.IsUseD" & Microsoft.VisualBasic.Right(sModuleID, 2) & " = 1" & vbCrLf
        If sWhere <> "" Then sSQL &= "WHERE " & sWhere & vbCrLf
        '*************************
        sSQL &= "ORDER BY DisplayOrder, EmpDisplayOrder, EmployeeName"
        Return sSQL
    End Function

    ''' <summary>
    ''' Table dữ liệu NV có ĐK Disabled = 0
    ''' </summary>
    ''' <param name="bCurDivision">gsDivisionID</param>
    ''' <param name="bHavePercent">%</param>
    ''' <param name="bUseUnicode">gbUnicode</param>
    ''' <param name="sModuleID">Dxx</param>
    ''' <returns>DataTable D09T0201</returns>
    ''' <remarks></remarks>
    Public Function ReturnTableEmployeeID(Optional ByVal bCurDivision As Boolean = False, Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False, Optional ByVal sModuleID As String = "") As DataTable
        ''Bổ sung Field Unicode
        'Dim sUnicode As String = ""
        'Dim sLanguage As String = ""
        'UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        ''***************
        'Dim sSQL As String = "--Do nguon Nhan vien" & vbCrLf
        'If bHavePercent Then
        '    sSQL &= "SELECT 	'%' As EmployeeID, " & sLanguage & "  As EmployeeName, '%' As DivisionID , '%' As BlockID, '%' As DepartmentID, '%' as  TeamID, '%' as  WorkingStatusID, '%' as EmpGroupID, 0 as EmpDisplayOrder, 0 as DisplayOrder" & vbCrLf
        '    sSQL &= "UNION" & vbCrLf
        'End If
        'sSQL &= "SELECT 	D01.EmployeeID,Isnull(D01.LastName" & sUnicode & ",'') + CASE WHEN  D01.MiddleName" & sUnicode & " ='' THEN '' ELSE ' ' + D01.MiddleName" & sUnicode & " END  + ' '+ Isnull(D01.FirstName" & sUnicode & ",'') as EmployeeName, D01.DivisionID, D02.BlockID,  D01.DepartmentID, D01.TeamID, D01.WorkingTypeID AS WorkingStatusID, D01.EmpGroupID, EmpDisplayOrder, 1 as DisplayOrder  " & vbCrLf
        'sSQL &= "FROM 	D09T0201 D01 WITH (NOLOCK) " & vbCrLf
        'sSQL &= "INNER JOIN D91T0012 D02  WITH(NOLOCK) ON D01.DepartmentID = D02.DepartmentID" & vbCrLf
        'sSQL &= "WHERE	D01.Disabled = 0 " & vbCrLf
        'If bCurDivision Then sSQL &= " And D02.DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        'If sModuleID.Contains("21") OrElse sModuleID.Contains("15") Then sSQL &= " And D01.IsUseD" & Microsoft.VisualBasic.Right(sModuleID, 2) & " = 1" & vbCrLf
        'sSQL &= "ORDER BY DisplayOrder, EmpDisplayOrder, EmployeeName"
        Return ReturnDataTable(SQLEmployeID(bCurDivision, bHavePercent, bUseUnicode, sModuleID))
    End Function

    Public Sub LoadtdbcEmployeeID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sWorkingStatusID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sFilter As String = ""
        If sBlockID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "BlockID =" & SQLString(sBlockID)
        End If
        If sDepartmentID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "DepartmentID =" & SQLString(sDepartmentID)
        End If
        If sTeamID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "TeamID =" & SQLString(sTeamID)
        End If
        If sWorkingStatusID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "WorkingStatusID =" & SQLString(sWorkingStatusID)
        End If
        If sFilter <> "" Then sFilter = "EmployeeID='%' or (" & sFilter & ")"
        LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, sFilter, True), gbUnicode)

        'tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
        Try
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
            tdbc.Splits(0).DisplayColumns("TeamID").Visible = False
            tdbc.Splits(0).DisplayColumns("WorkingStatusID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "EmployeeName"
        tdbc.ValueMember = "EmployeeID"
    End Sub

    Public Sub LoadtdbdEmployeeID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sWorkingStatusID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sTeamID, sWorkingStatusID}
        Dim sField() As String = {"BlockID", "DepartmentID", "TeamID", "WorkingStatusID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load tdbcEmployeeID by BlockID,DepartmentID,TeamID,WorkingStatusID and tdbcDivisionID 
    ''' </summary>
    ''' <param name="tdbc">tdbcEmployeeID</param>
    ''' <param name="dtOriginal">get value from ReturnTableEmployeeID() function</param>
    ''' <param name="sBlockID">tdbcBlockID</param>
    ''' <param name="sDivisionID">tdbcDivisionID</param>
    ''' <remarks></remarks>
    Public Sub LoadtdbcEmployeeID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sWorkingStatusID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sFilter As String = ""
        If sDivisionID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "DivisionID =" & SQLString(sDivisionID)
        End If
        If sBlockID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "BlockID =" & SQLString(sBlockID)
        End If
        If sDepartmentID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "DepartmentID =" & SQLString(sDepartmentID)
        End If
        If sTeamID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "TeamID =" & SQLString(sTeamID)
        End If
        If sWorkingStatusID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "WorkingStatusID =" & SQLString(sWorkingStatusID)
        End If
        If sFilter <> "" Then sFilter = "EmployeeID='%' or (" & sFilter & ")"
        LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, sFilter, True), gbUnicode)

        Try
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
            tdbc.Splits(0).DisplayColumns("TeamID").Visible = False
            tdbc.Splits(0).DisplayColumns("WorkingStatusID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "EmployeeName"
        tdbc.ValueMember = "EmployeeID"
    End Sub

    Public Sub LoadtdbdEmployeeID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sWorkingStatusID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sTeamID, sWorkingStatusID, sDivisionID}
        Dim sField() As String = {"BlockID", "DepartmentID", "TeamID", "WorkingStatusID", "DivisionID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load tdbcEmployeeID by BlockID,DepartmentID,TeamID,WorkingStatusID, EmpGroupID and tdbcDivisionID 
    ''' </summary>
    ''' <param name="tdbc">tdbcEmployeeID</param>
    ''' <param name="dtOriginal">get value from ReturnTableEmployeeID() function</param>
    ''' <param name="sBlockID">tdbcBlockID</param>
    ''' <param name="sDivisionID">tdbcDivisionID</param>
    ''' <remarks></remarks>
    Public Sub LoadtdbcEmployeeID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sWorkingStatusID As String, ByVal sDivisionID As String, ByVal sEmpGroupID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sFilter As String = ""
        If sDivisionID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "DivisionID =" & SQLString(sDivisionID)
        End If
        If sBlockID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "BlockID =" & SQLString(sBlockID)
        End If
        If sDepartmentID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "DepartmentID =" & SQLString(sDepartmentID)
        End If
        If sTeamID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "TeamID =" & SQLString(sTeamID)
        End If
        If sWorkingStatusID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "WorkingStatusID =" & SQLString(sWorkingStatusID)
        End If
        If sEmpGroupID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "EmpGroupID =" & SQLString(sEmpGroupID)
        End If
        If sFilter <> "" Then sFilter = "EmployeeID='%' or (" & sFilter & ")"
        LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, sFilter, True), gbUnicode)

        Try
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
            tdbc.Splits(0).DisplayColumns("TeamID").Visible = False
            tdbc.Splits(0).DisplayColumns("WorkingStatusID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "EmployeeName"
        tdbc.ValueMember = "EmployeeID"
    End Sub

    Public Sub LoadtdbdEmployeeID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sWorkingStatusID As String, ByVal sDivisionID As String, ByVal sEmpGroupID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sTeamID, sWorkingStatusID, sDivisionID, sEmpGroupID}
        Dim sField() As String = {"BlockID", "DepartmentID", "TeamID", "WorkingStatusID", "DivisionID", "EmpGroupID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), bUseUnicode)
    End Sub
#End Region

#Region "Nhóm nhân viên"
    Public Function ReturnTableEmpGroupID_D09P6868(ByVal DivisionID As String, ByVal FormID As String, ByVal IsMSS As Integer) As DataTable
        Return ReturnDataTable(SQLStoreD09P6868(DivisionID, FormID, "EmpGroup", IsMSS, "Do nguon Nhom nhan vien"))
    End Function

    ''' <summary>
    ''' Đổ nguồn Nhóm nhân viên
    ''' </summary>
    ''' <param name="bHavePercent"></param>
    ''' <param name="bUseUnicode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnTableEmpGroupID(Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False, Optional ByVal bCurDivision As Boolean = False) As DataTable
        Dim sSQL As String = "--Do nguon Nhom nhan vien" & vbCrLf
        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        If bHavePercent Then
            sSQL &= "SELECT 	" & AllCode & " As EmpGroupID, " & sLanguage & "  As EmpGroupName, '%' As DepartmentID, '%' as  TeamID, '%' As DivisionID , '%' As BlockID, 0 as EGDisplayOrder, 0 As DisplayOrder" & vbCrLf
            sSQL &= "UNION" & vbCrLf
        End If
        sSQL &= "SELECT 	EmpGroupID, EmpGroupName" & gsLanguage & UnicodeJoin(gbUnicode) & " As EmpGroupName, T1.DepartmentID, T1. TeamID, T2.DivisionID, T2.BlockID, EGDisplayOrder, 1 As DisplayOrder  " & vbCrLf
        sSQL &= "FROM 	    D09T1210 T1 WITH(NOLOCK) " & vbCrLf
        sSQL &= " INNER JOIN	 D91T0012 T2 WITH(NOLOCK) ON	T1.DepartmentID = T2.DepartmentID" & vbCrLf
        sSQL &= "WHERE	    T1.Disabled = 0 " & vbCrLf
        If bCurDivision Then sSQL &= " And T2.DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        sSQL &= "ORDER BY   DisplayOrder, EGDisplayOrder, EmpGroupName"
        Return ReturnDataTable(sSQL)
    End Function

    Public Sub LoadtdbcEmpGroupID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sTeamID}
        Dim sField() As String = {"BlockID", "DepartmentID", "TeamID"}
        LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)

        Try
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
            tdbc.Splits(0).DisplayColumns("TeamID").Visible = False
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "EmpGroupName"
        tdbc.ValueMember = "EmpGroupID"
    End Sub

    Public Sub LoadtdbcEmpGroupID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sDivisionID, sBlockID, sDepartmentID, sTeamID}
        Dim sField() As String = {"DivisionID", "BlockID", "DepartmentID", "TeamID"}
        LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)

        Try
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
            tdbc.Splits(0).DisplayColumns("TeamID").Visible = False
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "EmpGroupName"
        tdbc.ValueMember = "EmpGroupID"
    End Sub

    Public Sub LoadtdbdEmpGroupID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sTeamID}
        Dim sField() As String = {"BlockID", "DepartmentID", "TeamID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), bUseUnicode)
        Try
            tdbd.DisplayColumns("DivisionID").Visible = False
            tdbd.DisplayColumns("BlockID").Visible = False
            tdbd.DisplayColumns("DepartmentID").Visible = False
            tdbd.DisplayColumns("TeamID").Visible = False
        Catch ex As Exception

        End Try
        '        tdbd.DisplayMember = "EmpGroupName"
        '        tdbd.ValueMember = "EmpGroupID"
    End Sub

    Public Sub LoadtdbdEmpGroupID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sTeamID, sDivisionID}
        Dim sField() As String = {"BlockID", "DepartmentID", "TeamID", "DivisionID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), bUseUnicode)
        Try
            tdbd.DisplayColumns("DivisionID").Visible = False
            tdbd.DisplayColumns("BlockID").Visible = False
            tdbd.DisplayColumns("DepartmentID").Visible = False
            tdbd.DisplayColumns("TeamID").Visible = False
        Catch ex As Exception

        End Try
        '        tdbd.DisplayMember = "EmpGroupName"
        '        tdbd.ValueMember = "EmpGroupID"
    End Sub
#End Region

#Region "Hồ sơ lương tháng"

    ''' <summary>
    ''' Lấy giá trị hồ sơ lương tháng
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPayRollVoucherID() As String
        ' update 19/11/2013 id 61328 
        Return GetPayRollVoucherID(gsDivisionID, giTranMonth, giTranYear)
    End Function

    ''' <summary>
    ''' Lấy giá trị hồ sơ lương tháng
    ''' </summary>
    ''' <param name="iTranMonth"></param>
    ''' <param name="iTranYear"></param>
    ''' <param name="sDivisionID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPayRollVoucherID(ByVal sDivisionID As String, ByVal iTranMonth As Integer, ByVal iTranYear As Integer) As String
        ' update 19/11/2013 id 61328 
        Dim sSQL As String = SQLStoreD09P6001(sDivisionID, iTranMonth, iTranYear)
        Return ReturnScalar(sSQL)
    End Function

    Private Function SQLStoreD09P6001(ByVal sDivisionID As String, ByVal iTranMonth As Integer, ByVal iTranYear As Integer) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Lay ho so luong thang: Do nguon cho bien @PayrollVoucherID" & vbCrLf)
        sSQL &= "Exec D09P6001 "
        sSQL &= SQLString(sDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLNumber(iTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(iTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsUserID) 'UserID, varchar[50], NOT NULL
        Return sSQL
    End Function



#End Region

#Region "Người lập cho nhóm Nhân sự"

    ''' <summary>
    ''' Lấy giá trị mặc định của người lập cho G4
    ''' </summary>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadCreateByG4()
        'Update 24/10/2014: 67897 Bổ sung thông tin người lập
        Dim sSQL As String = "-- Lay gia tri mac dinh cua nguoi lap" & vbCrLf
        sSQL = "Select Top 1 T.HREmployeeID as ObjectID From LEMONSYS..D00T0030  T  Where "
        sSQL &= " T.UserID = " & SQLString(gsUserID)
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            gsCreateByG4 = dt.Rows(0).Item("ObjectID").ToString
            'Tạm thời G4 không lock
            'gbLockL3UserIDG4 = L3Bool(dt.Rows(0).Item("LockL3UserID"))
        End If

    End Sub

    ''' <summary>
    ''' Table load nguồn cho combo Người lập
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function ReturnTableCreateByG4(ByVal FormID As String, ByVal IsMSS As Integer) As DataTable
        Return ReturnDataTable(SQLStoreD09P6868(gsDivisionID, FormID, "CreateByHR", IsMSS, "Do nguon Nguoi lap"))
    End Function

    
    ''' Load Combo Người lập
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <remarks></remarks>
    Public Sub LoadCboCreateByG4(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal FormID As String, ByVal IsMSS As Integer)
        LoadDataSource(tdbc, ReturnTableCreateByG4(FormID, IsMSS), gbUnicode) 'Load by Current Division
    End Sub

    ''' <summary>
    ''' Load Dropdown Người lập
    ''' </summary>
    ''' <param name="tdbd"></param>
    ''' <remarks></remarks>
    Public Sub LoadDropdownCreateByG4(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal FormID As String, ByVal IsMSS As Integer)
        LoadDataSource(tdbd, ReturnTableCreateByG4(FormID, IsMSS), gbUnicode) 'Load by Current Division
    End Sub

    ''' <summary>
    ''' <summary>
    '''  Load Combo Người lập theo table có sẵn
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <param name="dt"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadCboCreateByG4(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dt As DataTable)
        LoadDataSource(tdbc, dt, gbUnicode)
    End Sub

    ''' <summary>
    ''' Gán text cho combo Người lập
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <remarks></remarks>
    Public Sub GetTextCreateByG4(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bDefault As Boolean = True)
        If bDefault Then tdbc.SelectedValue = gsCreateByG4
        'Nếu D91 thiết lập "Khóa người dùng Lemon3" (gbLockL3UserID = True) và combo Người lập có giá trị thì Lock Combo lại.
        'If gbLockL3UserIDG4 And tdbc.SelectedValue IsNot Nothing Then tdbc.ReadOnly = (tdbc.Text <> "")

    End Sub
#End Region


#Region "Nhóm khối"
    Public Function ReturnTableBlockGroupID() As DataTable
        Return ReturnDataTable(SQLStoreD09P1221(0, "Do nguon Nhom khoi"))
    End Function

    Public Sub LoadtdbcBlockGroupID(ByVal tdbc As C1.Win.C1List.C1Combo)
        LoadDataSource(tdbc, ReturnTableBlockGroupID, gbUnicode)
        tdbc.DisplayMember = "BlockGroupName"
        tdbc.ValueMember = "BlockGroupID"
    End Sub

    Public Sub LoadtdbcBlockGroupID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown)
        LoadDataSource(tdbd, ReturnTableFilter(ReturnTableBlockGroupID, "BlockGroupID <>'%'", True), gbUnicode)
    End Sub
#End Region

#Region "Nhóm phòng ban"
    Public Function ReturnTableDepartmentGroupID() As DataTable
        Return ReturnDataTable(SQLStoreD09P1221(1, "Do nguon Nhom phong ban"))
    End Function

    Public Sub LoadtdbcDepartmentGroupID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockGroupID As String)
        If sBlockGroupID = "%" Then
            LoadDataSource(tdbc, dtOriginal.DefaultView.ToTable, gbUnicode)
        Else
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentGroupID='%' or BlockGroupID =" & SQLString(sBlockGroupID)), gbUnicode)
        End If

        Try
            tdbc.Splits(0).DisplayColumns("BlockGroupID").Visible = False
        Catch ex As Exception

        End Try

        tdbc.DisplayMember = "DepartmentGroupName"
        tdbc.ValueMember = "DepartmentGroupID"
    End Sub

    Public Sub LoadtdbdDepartmentGroupID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockGroupID As String)
        Dim sValue() As String = {sBlockGroupID}
        Dim sField() As String = {"BlockGroupID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), gbUnicode)
    End Sub

#End Region

#Region "Nhóm tổ nhóm"
    Public Function ReturnTableTeamGroupID() As DataTable
        Return ReturnDataTable(SQLStoreD09P1221(2, "Do nguon nhom to nhom"))
    End Function

    Public Sub LoadtdbcTeamGroupID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockGroupID As String, ByVal sDepartmentGroupID As String)
        If sDepartmentGroupID = "%" And sBlockGroupID = "%" Then 'No Filter
            LoadDataSource(tdbc, dtOriginal.DefaultView.ToTable, gbUnicode)
        ElseIf sDepartmentGroupID = "%" And sBlockGroupID <> "%" Then 'Filter by BlockGroupID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamGroupID='%' or BlockGroupID =" & SQLString(sBlockGroupID), True), gbUnicode)
        ElseIf sDepartmentGroupID <> "%" And sBlockGroupID = "%" Then 'Filter by DepartmentGroupID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamGroupID='%' or DepartmentGroupID =" & SQLString(sDepartmentGroupID), True), gbUnicode)
        Else
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamGroupID='%' or (DepartmentGroupID =" & SQLString(sDepartmentGroupID) & " And BlockGroupID =" & SQLString(sBlockGroupID) & ")", True), gbUnicode)
        End If
        Try
            tdbc.Splits(0).DisplayColumns("BlockGroupID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentGroupID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "TeamGroupName"
        tdbc.ValueMember = "TeamGroupID"
    End Sub

    Public Sub LoadtdbdTeamGroupID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockGroupID As String, ByVal sDepartmentGroupID As String)
        Dim sValue() As String = {sBlockGroupID, sDepartmentGroupID}
        Dim sField() As String = {"BlockGroupID", "DepartmentGroupID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), gbUnicode)
    End Sub
#End Region

#Region "Nhóm nhóm nhân viên"

    Public Function ReturnTableEmpGroup2ID() As DataTable
        Return ReturnDataTable(SQLStoreD09P1221(3, "Do nguon Nhom nhom nhan vien"))
    End Function

    Public Sub LoadtdbcEmpGroup2ID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockGroupID As String, ByVal sDepartmentGroupID As String, ByVal sTeamGroupID As String)
        Dim sValue() As String = {sBlockGroupID, sDepartmentGroupID, sTeamGroupID}
        Dim sField() As String = {"BlockGroupID", "DepartmentGroupID", "TeamGroupID"}
        LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), gbUnicode)

        Try
            tdbc.Splits(0).DisplayColumns("BlockGroupID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentGroupID").Visible = False
            tdbc.Splits(0).DisplayColumns("TeamGroupID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "EmpGroup2Name"
        tdbc.ValueMember = "EmpGroup2ID"
    End Sub

    Public Sub LoadtdbdEmpGroup2ID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockGroupID As String, ByVal sDepartmentGroupID As String, ByVal sTeamGroupID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockGroupID, sDepartmentGroupID, sTeamGroupID}
        Dim sField() As String = {"BlockGroupID", "DepartmentGroupID", "TeamGroupID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField, False), True), bUseUnicode)
        Try
            tdbd.DisplayColumns("BlockGroupID").Visible = False
            tdbd.DisplayColumns("DepartmentGroupID").Visible = False
            tdbd.DisplayColumns("TeamGroupID").Visible = False
        Catch ex As Exception

        End Try
        tdbd.DisplayMember = "EmpGroup2Name"
        tdbd.ValueMember = "EmpGroup2ID"
    End Sub

#End Region


#End Region

#Region "Kiểm tra nhập Mã, Công thức nhóm G4"
    ''' <summary>
    ''' Thay đổi vị trí Select của chuỗi Vni
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="posFrom">vị trí bắt đầu</param>
    ''' <param name="posTo">Số ký tự được Select</param>
    ''' <remarks>Không cần kiểm tra khi Unicode</remarks>
    Private Sub ChangePositionIndexVNI(ByVal str As String, ByRef posFrom As Integer, ByRef posTo As Integer)
        If str = "" OrElse posFrom < 0 OrElse posFrom >= str.Length - 1 Then Exit Sub

        Dim arrChar() As String = {"Â", "Á", "À", "Å", "Ä", "Ã", "Ù", "Ø", "Û", "Õ", "Ï", "É", "È", "Ú", "Ü", "Ë", "Ê"}
        Dim c As String = (str.Substring(posFrom, 1)).ToUpper
        '"Ö", "Ô"
        Select Case c
            Case "Ö", "Ô" 'Ö: Ư; Ô: Ơ - không tăng vị trí, ngược lại thì tăng thêm 1 vị trí
                If L3FindArrString(arrChar, (str.Substring(posFrom + 1, 1)).ToUpper) Then posTo = 2
            Case Else 'kiểm tra trong danh sách arrChar
                If L3FindArrString(arrChar, c) Then
                    If posFrom > 0 Then posFrom -= 1
                    posTo = 2
                End If
        End Select
    End Sub

    'Kiểm tra Button Đóng có đặt Tên "Close"
    Private Function CheckContinue(ByVal ctrl As Control) As Boolean
        Try
            Dim form As Form = CType(ctrl.TopLevelControl, Form)
            If form.Controls.ContainsKey("btnClose") Then
                Dim btnClose As Control = CType(form.Controls("btnClose"), System.Windows.Forms.Button)
                If btnClose Is Nothing Then Return True 'không có nút đóng
                If btnClose.Focused Then Return False 'Nhấn vào nút Đóng
                Dim arr() As String = ctrl.Tag.ToString.Split(" "c) 'Nhấn ALT + N
                If arr.Length > 2 Then Return False
                '************
            End If
        Catch ex As Exception

        End Try
        Return True
    End Function

    Private Sub txtID_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim txtID As TextBox = CType(sender, TextBox)
        If txtID.ReadOnly OrElse txtID.Enabled = False Then Exit Sub
        'Nếu nhấn đóng thì không cần hiện thông báo
        If CheckContinue(txtID) = False Then Exit Sub
        '**************
        'update 15/06/2015: bỏ ký tự Enter, khoảng trắng
        txtID.Text = ReplaceCharactorEnter(txtID.Text)
        '************
        Select Case geLanguage
            Case EnumLanguage.Vietnamese, EnumLanguage.English
            Case Else 'Nếu ngôn ngữ khác thì không kiểm tra ký tự đặc biệt 06/09/2018
                Exit Sub
        End Select
        Dim posFrom As Integer
        'Bổ sung kiểm tra ký tự đặc biệt chuỗi truyền vào
        Dim arrTag() As String = Nothing
        If txtID.Tag IsNot Nothing Then arrTag = txtID.Tag.ToString.Split(" "c)
        Dim bFormula As Boolean = False
        Dim sCheckID As String = ""
        If arrTag IsNot Nothing AndAlso arrTag.Length > 0 Then
            bFormula = L3Bool(arrTag(0))
            If arrTag.Length > 1 Then sCheckID = arrTag(1)
        End If
        If bFormula Then
            posFrom = IndexFormulaCharactor(txtID.Text, sCheckID)
        Else
            posFrom = IndexIdCharactor(txtID.Text, sCheckID)
        End If
        '***********************
        Dim posTo As Integer = 1 'posFrom
        If posFrom > 0 Then posTo = posFrom

        If txtID.Font.Name.Contains("Lemon3") Then ChangePositionIndexVNI(txtID.Text, posFrom, posTo)
        Select Case posFrom
            Case -1 'thỏa điều kiện
            Case -2 'Vượt chiều dài
                'D99C0008.MsgL3("Chiều dài vượt quá quy định.")
                'txtID.SelectAll()
                'e.Cancel = True
            Case Else 'vi phạm
                D99C0008.MsgL3(rL3("Ma_co_ky_tu_khong_hop_le"))
                e.Cancel = True
                txtID.Select(posFrom, posTo)
        End Select
    End Sub

    Private Sub txtID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.Modifiers <> Keys.Alt Then Exit Sub
        If e.KeyCode = Keys.N Then
            Dim txtID As TextBox = CType(sender, TextBox)
            txtID.Tag = txtID.Tag.ToString & " True"
        End If
    End Sub

    ''' <summary>
    ''' Kiểm tra TextBox Nhập Mã/Công thức
    ''' </summary>
    ''' <param name="txtID">Control cần kiểm tra</param>
    ''' <param name="iLength">Chiều dài nhập liệu</param>
    ''' <param name="bFormula">Theo kiểu công thức (Default: False - cho Mã)</param>
    ''' <remarks>Cho một textbox, đối số bFormula : Kiểu kiểm tra là Mã hay Công thức</remarks>
    Public Sub CheckIdTextBoxG4(ByRef txtID As TextBox, Optional ByVal iLength As Integer = 20, Optional ByVal bFormula As Boolean = False, Optional ByVal sCheckID As String = "")
        ' txtID.Multiline = True
        txtID.CharacterCasing = CharacterCasing.Upper
        txtID.MaxLength = iLength

        'If bFormula = False Then sCheckID &= ":;\{},()""<>=&~" 'TH nhập mã
        'Update 08/05/2015: theo Incident 75272 --> bổ sung thêm chặn 2 ký tự | và ?
        If bFormula = False Then sCheckID &= ":;\{},()""<>=&~|?`!" 'TH nhập mã
        '* + - /        :;\{},()"<>=&~ và Bổ sung chặn 3 ký tự | và ?

        txtID.Tag = bFormula.ToString & IIf(sCheckID = "", "", " " & sCheckID).ToString
        AddHandler txtID.KeyDown, AddressOf txtID_KeyDown 'Khi nhấn ALT + N thì không kiểm tra
        AddHandler txtID.Validating, AddressOf txtID_Validating
    End Sub

    ''' <summary>
    ''' Kiểm tra nhiều TextBox Nhập Mã/Công thức
    ''' </summary>
    ''' <param name="txtID">Control cần kiểm tra</param>
    ''' <param name="iLength">Chiều dài nhập liệu</param>
    ''' <param name="bFormula">Theo kiểu công thức (Default: False - cho Mã)</param>
    ''' <remarks>Cho một textbox, đối số bFormula : Kiểu kiểm tra là Mã hay Công thức</remarks>
    Public Sub CheckIdTextBoxG4(ByRef txtID() As TextBox, Optional ByVal iLength As Integer = 20, Optional ByVal bFormula As Boolean = False, Optional ByVal sCheckID As String = "")
        For i As Integer = 0 To txtID.Length - 1
            CheckIdTextBoxG4(txtID(i), iLength, bFormula, sCheckID)
        Next
    End Sub

    ''' <summary>
    ''' Kiểm tra Mã hợp lệ 
    ''' </summary>
    ''' <param name="str">Chuỗi kiểm tra</param>
    ''' <returns>Vị trí ký tự vi phạm</returns>
    ''' <remarks></remarks>
    Private Function IndexIdCharactor(ByVal str As String, Optional ByVal sCheckID As String = ":;\{},()""<>=&~") As Integer
        'BackSpace: 8
        'Kiem tra theo chuan nhap Ma cho G4: dấu ' % [ ] ^ * + - . / :;\{},()"<>=&~ và Bổ sung chặn 2 ký tự | và ?

        For Each c As Char In str
            Select Case AscW(c)
                Case 13, 10 'Mutiline của textbox và phím Enter
                    Continue For
                    'Update 25/06/2015: tạm thời rem lại theo mail của Lan, bỏ dấu .
                    'Case Is < 33, Is > 127, 37, 39, 42, 43, 45, 46, 47, 91, 93, 94 'Các ký tự đặc biệt: 37(%) 39(') 42 (*) 43 (+) 45 (-) 46 (.) 47 (/) 91([) 93(]) 94(^)
                Case Is < 33, Is > 127, 37, 39, 42, 43, 45, 47, 91, 93, 94 'Các ký tự đặc biệt: 37(%) 39(') 42 (*) 43 (+) 45 (-) 47 (/) 91([) 93(]) 94(^)
                    Return str.IndexOf(c)
            End Select
            If sCheckID <> "" Then
                Dim index As Integer = Strings.InStr(sCheckID, c, CompareMethod.Text)
                If index > 0 Then Return str.IndexOf(c)
            End If
        Next
        Return -1
    End Function


    '''' Kiểm tra công thức hợp lệ
    '''' </summary>
    '''' <param name="str">Chuỗi kiểm tra</param>
    '''' <returns>Vị trí ký tự vi phạm</returns>
    '''' <remarks></remarks>
    Private Function IndexFormulaCharactor(ByVal str As String, Optional ByVal sCheckID As String = "") As Integer
        '  If str.Length > iLength Then Return -2 'vượt chiều dài
        'BackSpace: 8
        For Each c As Char In str
            Select Case AscW(c)
                Case 13, 10 'Mutiline của textbox và phím Enter
                    Continue For
                Case Is < 33, Is > 127, 94 ''Các ký tự đặc biệt: 94(^)
                    Return str.IndexOf(c)
            End Select
            If sCheckID <> "" Then
                Dim index As Integer = Strings.InStr(sCheckID, c, CompareMethod.Text)
                If index > 0 Then Return str.IndexOf(c)
            End If
        Next
        Return -1
    End Function

    Private Function CheckFormulaCharactor(ByVal str As String) As Boolean
        Return IndexFormulaCharactor(str) >= 0
    End Function

    Public Sub CheckIdTextBoxG4_CheckG4(ByRef txtID As TextBox, ByVal iLength As Integer, ByVal bFormula As Boolean, ByVal sCheckID As String)
        '4/6/2020, id 136438-Thêm biến bCheckG4 = False: bỏ chặn 4 ký tự * - / \
        txtID.CharacterCasing = CharacterCasing.Upper
        txtID.MaxLength = iLength

        ''Update 08/05/2015: theo Incident 75272 --> bổ sung thêm chặn 2 ký tự | và ?
        If bFormula = False Then sCheckID = ":;{},()""<>=&~|?"

        txtID.Tag = bFormula.ToString & IIf(sCheckID = "", "", " " & sCheckID).ToString
        AddHandler txtID.KeyDown, AddressOf txtID_KeyDown 'Khi nhấn ALT + N thì không kiểm tra
        AddHandler txtID.Validating, AddressOf txtIDCheckG4_Validating
    End Sub

    Public Sub CheckIdTextBoxG4_CheckG4(ByRef txtID() As TextBox, Optional ByVal iLength As Integer = 20, Optional ByVal bFormula As Boolean = False, Optional ByVal sCheckID As String = "")
        '4/6/2020, id 136438-Thêm biến bCheckG4 = False: bỏ chặn 4 ký tự * - / \, bCheckG4 = True : không chặn 4 ký tự này
        For i As Integer = 0 To txtID.Length - 1
            CheckIdTextBoxG4_CheckG4(txtID(i), iLength, bFormula, sCheckID)
        Next
    End Sub

    Private Sub txtIDCheckG4_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        '4/6/2020, id 136438-Thêm biến bCheckG4 = False: bỏ chặn 4 ký tự * - / \
        Dim txtID As TextBox = CType(sender, TextBox)
        If txtID.ReadOnly OrElse txtID.Enabled = False Then Exit Sub
        'Nếu nhấn đóng thì không cần hiện thông báo
        If CheckContinue(txtID) = False Then Exit Sub
        '**************
        'update 15/06/2015: bỏ ký tự Enter, khoảng trắng
        txtID.Text = ReplaceCharactorEnter(txtID.Text)
        '************
        Select Case geLanguage
            Case EnumLanguage.Vietnamese, EnumLanguage.English
            Case Else 'Nếu ngôn ngữ khác thì không kiểm tra ký tự đặc biệt 06/09/2018
                Exit Sub
        End Select
        Dim posFrom As Integer
        'Bổ sung kiểm tra ký tự đặc biệt chuỗi truyền vào
        Dim arrTag() As String = Nothing
        If txtID.Tag IsNot Nothing Then arrTag = txtID.Tag.ToString.Split(" "c)
        Dim bFormula As Boolean = False
        Dim sCheckID As String = ""
        If arrTag IsNot Nothing AndAlso arrTag.Length > 0 Then
            bFormula = L3Bool(arrTag(0))
            If arrTag.Length > 1 Then sCheckID = arrTag(1)
        End If

        If bFormula Then
            posFrom = IndexFormulaCharactor(txtID.Text, sCheckID)
        Else
            posFrom = IndexIdCharactorCheckG4(txtID.Text, sCheckID, False) '4/6/2020, id 136438-Thêm biến bCheckG4 = False: bỏ chặn 4 ký tự * - / \, bCheckG4 = True : không chặn 4 ký tự này
        End If
        '***********************
        Dim posTo As Integer = 1 'posFrom
        If posFrom > 0 Then posTo = posFrom

        If txtID.Font.Name.Contains("Lemon3") Then ChangePositionIndexVNI(txtID.Text, posFrom, posTo)
        Select Case posFrom
            Case -1 'thỏa điều kiện
            Case -2 'Vượt chiều dài
                ''D99C0008.MsgL3("Chiều dài vượt quá quy định.")
                ''txtID.SelectAll()
                ''e.Cancel = True
            Case Else 'vi phạm
                D99C0008.MsgL3(rL3("Ma_co_ky_tu_khong_hop_le"))
                e.Cancel = True
                txtID.Select(posFrom, posTo)
        End Select
    End Sub

    ''' <summary>
    ''' Kiểm tra Mã hợp lệ 
    ''' </summary>
    ''' <param name="str">Chuỗi kiểm tra</param>
    ''' <returns>Vị trí ký tự vi phạm</returns>
    ''' <remarks></remarks>
    Private Function IndexIdCharactorCheckG4(ByVal str As String, Optional ByVal sCheckID As String = ":;\{},()""<>=&~", Optional bCheckG4 As Boolean = True) As Integer
        'BackSpace: 8
        'Kiem tra theo chuan nhap Ma cho G4: dấu ' % [ ] ^ * + - . / :;\{},()"<>=&~ và Bổ sung chặn 2 ký tự | và ?

        For Each c As Char In str
            '4/6/2020, id 136438-Thêm biến bCheckG4 = False: bỏ chặn 4 ký tự * - / \, bCheckG4 = True : không chặn 4 ký tự này
            If bCheckG4 Then
                Select Case AscW(c)
                    Case 13, 10 'Mutiline của textbox và phím Enter
                        Continue For
                        ''Update 25/06/2015: tạm thời rem lại theo mail của Lan, bỏ dấu .
                        ''Case Is < 33, Is > 127, 37, 39, 42, 43, 45, 46, 47, 91, 93, 94 'Các ký tự đặc biệt: 37(%) 39(') 42 (*) 43 (+) 45 (-) 46 (.) 47 (/) 91([) 93(]) 94(^)
                    Case Is < 33, Is > 127, 37, 39, 42, 43, 45, 47, 91, 93, 94 'Các ký tự đặc biệt: 37(%) 39(') 42 (*) 43 (+) 45 (-) 47 (/) 91([) 93(]) 94(^)
                        Return str.IndexOf(c)
                End Select
            Else
                Select Case AscW(c)
                    Case 13, 10 'Mutiline của textbox và phím Enter
                        Continue For
                        ''Update 25/06/2015: tạm thời rem lại theo mail của Lan, bỏ dấu .
                        ''Case Is < 33, Is > 127, 37, 39, 42, 43, 45, 46, 47, 91, 93, 94 'Các ký tự đặc biệt: 37(%) 39(') 42 (*) 43 (+) 45 (-) 46 (.) 47 (/) 91([) 93(]) 94(^)
                    Case Is < 33, Is > 127, 37, 39, 43, 91, 93, 94 'Các ký tự đặc biệt: 37(%) 39(') 43 (+) 91([) 93(]) 94(^)
                        Return str.IndexOf(c)
                End Select
            End If

            If sCheckID <> "" Then
                Dim index As Integer = Strings.InStr(sCheckID, c, CompareMethod.Text)
                If index > 0 Then Return str.IndexOf(c)
            End If
        Next
        Return -1
    End Function

#End Region

#Region "Các hàm cục bộ"

    Private Function ReturnFilter(ByVal arrValue() As String, ByVal arrField() As String, Optional ByVal bHavePercent As Boolean = True) As String
        Dim sFilter As String = ""
        If arrValue.Length = 0 OrElse arrValue.Length <> arrField.Length Then
            D99C0008.MsgL3("Điều kiện lọc không đúng.")
            Return sFilter
        End If

        For i As Integer = 0 To arrValue.Length - 1
            If arrValue(i) <> "%" Then
                If sFilter <> "" Then sFilter &= " And "
                sFilter &= arrField(i) & " =" & SQLString(arrValue(i))
            End If
        Next
        If sFilter <> "" Then
            If bHavePercent Then
                sFilter = arrField(0) & " ='%' or (" & sFilter & ")"
                '            Else
                '                sFilter = arrField(0) & " (" & sFilter & ")"
            End If
        End If

        Return sFilter
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD09P6868
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 10/01/2014 01:59:24
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD09P6868(ByVal DivisionID As String, ByVal FormID As String, ByVal TypeID As String, ByVal IsMSS As Integer, ByVal sComment As String) As String
        Dim sSQL As String = ""
        sSQL &= ("-- " & sComment & vbCrLf)
        sSQL &= "Exec D09P6868 "
        sSQL &= SQLString(DivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(FormID) & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLString(TypeID) & COMMA 'TypeID, varchar[20], NOT NULL
        sSQL &= SQLNumber(IsMSS) & COMMA 'IsMSS, tinyint, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD09P1221
    '# Created User: Nguyễn Lê Phương
    '# Created Date: 23/01/2015 10:11:03
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD09P1221(ByVal iMode As Byte, ByVal sComment As String) As String
        Dim sSQL As String = ""
        sSQL &= ("-- " & sComment & vbCrLf)
        sSQL &= "Exec D09P1221 "
        sSQL &= SQLNumber(iMode) 'Mode, int, NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD09P2525
    '# Created User: Kim Long
    '# Created Date: 11/10/2017 11:14:27
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD09P2525(ByVal sType As String, ByVal bIsPercent As Boolean) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon combo nhom phong ban, phong ban" & vbCrLf)
        sSQL &= "Exec D09P2525 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(sType) & COMMA 'Type, varchar[50], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[50], NOT NULL
        sSQL &= SQLNumber(bIsPercent) 'IsPercent, TINYINT, NOT NULL
        Return sSQL
    End Function


#End Region

End Module

