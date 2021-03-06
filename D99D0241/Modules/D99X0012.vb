Imports System
'#######################################################################################
'#                                     CHÚ Ý
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày tạo: 28/06/2012
'# Người tạo: Nguyễn Thị Ánh
'# Ngày cập nhật cuối cùng: 06/01/2014
'# Người cập nhật cuối cùng: Nguyễn Thị Minh Hòa
'# Bổ sung Tạo dropdown động
'# Bổ sung Đổ nguồn Dự án, Hạng mục
'# Bổ sung Đổ nguồn Mã XDCB
'# Bổ sung chuẩn hiển thị Dự án - hạng mục cho Truy vấn
'# Sửa hàm ReturnConversionFactor(thêm biến ModuleID)
'# Bổ sung Đổ nguồn Kho hàng cho dropdown
'# Sửa add dropdown động : 14/06/2013
'# Chặn ds cột thêm dropdown vào nothing : 21/06/2013
'# Thêm các hàm Load Tài khoản, Khoản mục phân quyền theo Đơn vị
'#######################################################################################
''' <summary>
''' Module liên quan đến các vấn đề Load nguồn
''' </summary>
''' <remarks></remarks>
Public Module D99X0012

#Region "Đơn vị tính"
    Public Function ReturnTableUnitID(ByVal sInventoryID As String, Optional ByVal sOrderby As String = "") As DataTable
        Dim sSQL As String = ""
        sSQL &= "SELECT T1.UnitID,T2.UnitName" & UnicodeJoin(gbUnicode) & " as UnitName, IsNull(T1.ConversionFactor,1) As ConversionFactor, T1.Tolerance,T1.UseFormula,T1.Formula " & vbCrLf
        sSQL &= " FROM D07T0004 T1   WITH(NOLOCK)" & vbCrLf
        sSQL &= "INNER JOIN  D07T0005 T2  WITH(NOLOCK) ON  T1.UnitID=T2.UnitID" & vbCrLf
        sSQL &= "WHERE T1.Disabled=0 And T1.InventoryID=" & SQLString(sInventoryID)
        If sOrderby <> "" Then sSQL &= vbCrLf & "Order by " & sOrderby
        Return ReturnDataTable(sSQL)
    End Function

    Public Sub LoadtdbdUnitID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sInventoryID As String, Optional ByVal sOrderby As String = "")
        LoadDataSource(tdbd, ReturnTableUnitID(sInventoryID, sOrderby), gbUnicode)
    End Sub

    Private Function GetValueCol(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sField As String, Optional ByVal row As Integer = -1) As Object
        Dim sValue As String = ""
        Try
            If row = -1 Then
                sValue = tdbg.Columns(sField).Text
            Else
                sValue = tdbg(row, sField)
            End If
        Catch ex As Exception
            
        End Try
        Return sValue
    End Function

    '#------------------------------------------------------------------------
    '#Title: SQLStoreD07P7004
    '#Create User: Nguyễn Thị Ánh
    '#Create Date: 25/12/2007 10:11:08
    '#Modified User:
    '#Modified Date:
    '#Description: Lấy ConversionFactor khi UseFormula=1
    '#------------------------------------------------------------------------
    Private Function SQLStoreDxxP7004(ByVal sModuleID As String, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sFieldFormula As String, Optional ByVal row As Integer = -1, Optional ByVal InventoryID As String = "InventoryID", Optional ByVal LocationNo As String = "LocationNo") As String
        Dim sSQL As String = ""
        sSQL = "Exec " & sModuleID & "P7004 "
        Try
            'If row = -1 Then
            '    sSQL &= SQLString(tdbg.Columns("InventoryID").Text) & COMMA  ' InventoryID
            '    sSQL &= SQLString(tdbg.Columns("LocationNo").Text) & COMMA ' LocationNo
            '    sSQL &= SQLString(tdbg.Columns("Spec01ID").Text) & COMMA  ' Spec01ID
            '    sSQL &= SQLString(tdbg.Columns("Spec02ID").Text) & COMMA  ' Spec02ID
            '    sSQL &= SQLString(tdbg.Columns("Spec03ID").Text) & COMMA  ' Spec03ID
            '    sSQL &= SQLString(tdbg.Columns("Spec04ID").Text) & COMMA  ' Spec04ID
            '    sSQL &= SQLString(tdbg.Columns("Spec05ID").Text) & COMMA  ' Spec05ID
            '    sSQL &= SQLString(tdbg.Columns("Spec06ID").Text) & COMMA  ' Spec06ID
            '    sSQL &= SQLString(tdbg.Columns("Spec07ID").Text) & COMMA  ' Spec07ID
            '    sSQL &= SQLString(tdbg.Columns("Spec08ID").Text) & COMMA  ' Spec08ID
            '    sSQL &= SQLString(tdbg.Columns("Spec09ID").Text) & COMMA  ' Spec09ID
            '    sSQL &= SQLString(tdbg.Columns("Spec10ID").Text) & COMMA  ' Spec10ID
            '    sSQL &= SQLString(tdbg.Columns(sFieldFormula).Text)  ' Formula
            'Else 'Dùng cho HeadClick
            '    sSQL &= SQLString(tdbg(row, "InventoryID")) & COMMA  ' InventoryID
            '    sSQL &= SQLString(tdbg(row, "LocationNo")) & COMMA ' LocationNo
            '    sSQL &= SQLString(tdbg(row, "Spec01ID")) & COMMA  ' Spec01ID
            '    sSQL &= SQLString(tdbg(row, "Spec02ID")) & COMMA  ' Spec02ID
            '    sSQL &= SQLString(tdbg(row, "Spec03ID")) & COMMA  ' Spec03ID
            '    sSQL &= SQLString(tdbg(row, "Spec04ID")) & COMMA  ' Spec04ID
            '    sSQL &= SQLString(tdbg(row, "Spec05ID")) & COMMA  ' Spec05ID
            '    sSQL &= SQLString(tdbg(row, "Spec06ID")) & COMMA  ' Spec06ID
            '    sSQL &= SQLString(tdbg(row, "Spec07ID")) & COMMA  ' Spec07ID
            '    sSQL &= SQLString(tdbg(row, "Spec08ID")) & COMMA  ' Spec08ID
            '    sSQL &= SQLString(tdbg(row, "Spec09ID")) & COMMA  ' Spec09ID
            '    sSQL &= SQLString(tdbg(row, "Spec10ID")) & COMMA  ' Spec10ID
            '    sSQL &= SQLString(tdbg(row, sFieldFormula))  ' Formula
            'End If
            sSQL &= SQLString(GetValueCol(tdbg, InventoryID, row)) & COMMA  ' InventoryID
            sSQL &= SQLString(GetValueCol(tdbg, LocationNo, row)) & COMMA ' LocationNo
            sSQL &= SQLString(GetValueCol(tdbg, "Spec01ID", row)) & COMMA  ' Spec01ID
            sSQL &= SQLString(GetValueCol(tdbg, "Spec02ID", row)) & COMMA  ' Spec02ID
            sSQL &= SQLString(GetValueCol(tdbg, "Spec03ID", row)) & COMMA  ' Spec03ID
            sSQL &= SQLString(GetValueCol(tdbg, "Spec04ID", row)) & COMMA  ' Spec04ID
            sSQL &= SQLString(GetValueCol(tdbg, "Spec05ID", row)) & COMMA  ' Spec05ID
            sSQL &= SQLString(GetValueCol(tdbg, "Spec06ID", row)) & COMMA  ' Spec06ID
            sSQL &= SQLString(GetValueCol(tdbg, "Spec07ID", row)) & COMMA  ' Spec07ID
            sSQL &= SQLString(GetValueCol(tdbg, "Spec08ID", row)) & COMMA  ' Spec08ID
            sSQL &= SQLString(GetValueCol(tdbg, "Spec09ID", row)) & COMMA  ' Spec09ID
            sSQL &= SQLString(GetValueCol(tdbg, "Spec10ID", row)) & COMMA  ' Spec10ID
            sSQL &= SQLString(GetValueCol(tdbg, sFieldFormula, row))  ' Formula
        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try
        Return sSQL
    End Function

    ''' <summary>
    ''' Trả về Hệ số quy đổi theo D07P7004 khi UseConversionFormula =1
    ''' </summary>
    ''' <param name="tdbg"></param>
    ''' <param name="sFieldFormula">Field ConversionFormula trên lưới</param>
    ''' <returns>Hệ số quy đổi theo Result của D07P7004</returns>
    ''' <remarks>Mặc định trả về 1</remarks>
    Public Function ReturnConversionFactor(ByVal sModuleID As String, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sFieldFormula As String, Optional ByVal Row As Integer = -1) As Object
        Dim dt As DataTable = ReturnDataTable(SQLStoreDxxP7004(sModuleID, tdbg, sFieldFormula, Row, "InventoryID", "LocationNo"))
        If dt.Rows.Count > 0 Then Return dt.Rows(0).Item("Result")
        Return 1
    End Function

    Public Function ReturnConversionFactor(ByVal sModuleID As String, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sFieldFormula As String, ByVal Row As Integer, ByVal InventoryID As String, Optional ByVal LocationNo As String = "LocationNo") As Object
        Dim dt As DataTable = ReturnDataTable(SQLStoreDxxP7004(sModuleID, tdbg, sFieldFormula, Row, InventoryID, LocationNo))
        If dt.Rows.Count > 0 Then Return dt.Rows(0).Item("Result")
        Return 1
    End Function
#End Region

#Region "Đổ nguồn Kho hàng"
    ''' <summary>
    ''' Trả về Table Kho hàng của tất cả Đơn vị
    ''' </summary>
    ''' <param name="bAll">Sử dụng Tất cả. Mặc định True</param>
    ''' <param name="sWhere">Điều kiện lọc. Mặc định ""</param>
    ''' <returns>Table Kho hàng</returns>
    ''' <remarks>Nghiệp vụ thêm sWhere = Disabled = 0. Ngược lại thì không có</remarks>
    ''' 
    Private Function ReturnTableWareHouseID(Optional ByVal bAll As Boolean = True, Optional ByVal sWhere As String = "") As DataTable
        ''Private Function ReturnTableWareHouseID(Optional ByVal bAll As Boolean = True, Optional ByVal bAllowAddNew As Boolean = False,, Optional ByVal sWhere As String = "") As DataTable
        ''Bổ sung ObjectTypeID, ObjectID theo 82383 Thu Thảo
        Dim sSQL As String = SQLFinance.SQLWareHouseID(IIf(bAll, EnumUnionAll.All, EnumUnionAll.None), "", sWhere)
        'sSQL &= ("-- Do nguon Kho hang" & vbCrLf)
        'If bAll Then sSQL &= "Select " & AllCode & " as WareHouseID, " & AllName & " as WareHouseName, '%' as DivisionID, '%' as ObjectTypeID, '%' as ObjectID, 0 as DisplayOrder " & vbCrLf & "Union All" & vbCrLf
        'sSQL &= "Select WareHouseID, WareHouseName" & UnicodeJoin(gbUnicode) & " as WareHouseName, DivisionID, ObjectTypeID, ObjectID, 1 as DisplayOrder " & vbCrLf
        'sSQL &= "From D07T0007  WITH(NOLOCK)" & vbCrLf
        'sSQL &= "Where ( DAGroupID = '' OR DAGroupID IN (SELECT DAGroupID From LEMONSYS.DBO.D00V0080  WHERE UserID = " & SQLString(gsUserID) & ") OR 'LEMONADMIN' =" & SQLString(gsUserID) & ")" & vbCrLf
        'If sWhere <> "" Then sSQL &= " And " & sWhere & vbCrLf
        'sSQL &= "Order by DisplayOrder, WareHouseID"
        ''        'Update 15/06/2015: update theo mail của Tiến
        ''        sSQL = SQLStoreD91P3305(bAll, bAllowAddNew, swhere)

        Return ReturnDataTable(sSQL)
    End Function

    '    '#---------------------------------------------------------------------------------------------------
    '    '# Title: SQLStoreD91P3305
    '    '# Created User: Nguyễn Thị Minh Hòa
    '    '# Created Date: 15/06/2015 10:06:19
    '    '#---------------------------------------------------------------------------------------------------
    '    Private Function SQLStoreD91P3305(Optional ByVal bAll As Boolean = True, Optional ByVal bAllowAddNew As Boolean = False, Optional ByVal sWhere As String = "") As String
    '        Dim sSQL As String = ""
    '        sSQL &= ("-- Do nguon kho hang" & vbCrLf)
    '        sSQL &= "Exec D91P3305 "
    '        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
    '        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
    '        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[20], NOT NULL
    '        sSQL &= SQLString(XXXXX) & COMMA 'TypeID, varchar[20], NOT NULL
    '        sSQL &= SQLString(XXXXX) & COMMA 'FormID, varchar[20], NOT NULL
    '        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
    '        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
    '        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
    '        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
    '        sSQL &= SQLString("") & COMMA 'ListType, varchar[50], NOT NULL
    '        sSQL &= SQLNumber(bAllowAddNew) & COMMA 'AllowAddNew, tinyint, NOT NULL
    '        sSQL &= SQLNumber(bAll) & COMMA 'HaveAll, tinyint, NOT NULL
    '        sSQL &= SQLNumber(XXXXX) & COMMA 'ShowDisabled, tinyint, NOT NULL
    '        sSQL &= SQLStringUnicode(XXXXX, gbUnicode, True) 'StrFind, nvarchar[4000], NOT NULL
    '        Return sSQL
    '    End Function

    ''' <summary>
    ''' Đổ nguồn cho nhiều combo Kho hàng phụ thuộc Đơn vị
    ''' </summary>
    ''' <param name="tdbc">Danh sách combo</param>
    ''' <param name="dtWareHouse">Có thể truyền Nothing</param>
    ''' <param name="bDisableZero">Sử dụng Nghiệp vụ (True)/ Truy vấn/ Báo cáo (False). Nghiệp vụ thêm Disabled = 0. Ngược lại thì không có</param>
    ''' <param name="sDivsionID">Đơn vị. Mặc định rỗng</param>
    ''' <param name="bAll">Sử dụng Tất cả. Mặc định True</param>
    ''' <remarks>Có thể lấy được dtWareHouse nếu truyền Nothing. LoadtdbcWareHouseID(New C1.Win.C1List.C1Combo() {tdbcWareHouseIDFrom, tdbcWareHouseIDTo}, dtWareHouse, tdbcDivisionID.SelectedValue.ToString)</remarks>
    Public Sub LoadtdbcWareHouseID(ByVal tdbc() As C1.Win.C1List.C1Combo, ByRef dtWareHouse As DataTable, ByVal bDisableZero As Boolean, Optional ByVal sDivsionID As String = "", Optional ByVal bAll As Boolean = True)
        If dtWareHouse Is Nothing Then dtWareHouse = ReturnTableWareHouseID(bAll, IIf(bDisableZero, " Disabled = 0", "").ToString)

        Dim sFilter As String = ""
        If sDivsionID <> "%" And sDivsionID <> "" Then sFilter = " DivisionID = " & SQLString(sDivsionID) & " or DivisionID = '%'"
        Dim dtTemp As DataTable = ReturnTableFilter(dtWareHouse, sFilter, True)
        For i As Integer = 0 To tdbc.Length - 1
            LoadDataSource(tdbc(i), dtTemp.DefaultView.ToTable, gbUnicode)
            If bAll Then tdbc(i).SelectedIndex = 0
            Try
                tdbc(i).Columns("DivisionID").Caption = rL3("Don_vi")
                tdbc(i).Splits(0).DisplayColumns("DivisionID").Visible = sDivsionID = "%"
            Catch ex As Exception

            End Try
        Next
    End Sub

    ''' <summary>
    ''' Đổ nguồn cho nhiều combo Kho hàng phụ thuộc Đơn vị
    ''' </summary>
    ''' <param name="tdbc">Danh sách combo</param>
    ''' <param name="dtWareHouse">Có thể truyền Nothing</param>
    ''' <param name="bDisableZero">Sử dụng Nghiệp vụ (True)/ Truy vấn/ Báo cáo (False). Nghiệp vụ thêm Disabled = 0. Ngược lại thì không có</param>
    ''' <param name="sDivsionID">Đơn vị. Mặc định rỗng</param>
    ''' <param name="bAll">Sử dụng Tất cả. Mặc định True</param>
    ''' <remarks>Có thể lấy được dtWareHouse nếu truyền Nothing. LoadtdbcWareHouseID(New C1.Win.C1List.C1Combo() {tdbcWareHouseIDFrom, tdbcWareHouseIDTo}, dtWareHouse, tdbcDivisionID.SelectedValue.ToString)</remarks>
    Public Sub LoadtdbcWareHouseID(ByRef dtWareHouse As DataTable, ByVal bDisableZero As Boolean, ByVal sDivsionID As String, ByVal bAll As Boolean, ByVal ParamArray tdbc() As C1.Win.C1List.C1Combo)
        LoadtdbcWareHouseID(tdbc, dtWareHouse, bDisableZero, sDivsionID, bAll)
    End Sub

    Public Sub LoadtdbcWareHouseID(ByVal bDisableZero As Boolean, ByVal bAll As Boolean, ByVal ParamArray tdbc() As C1.Win.C1List.C1Combo)
        LoadtdbcWareHouseID(tdbc, bDisableZero, bAll)
    End Sub

    ''' <summary>
    ''' Đổ nguồn cho nhiều combo Kho hàng theo Đơn vị hiện tại
    ''' </summary>
    ''' <param name="tdbc">Danh sách combo</param>
    ''' <param name="bAll">Sử dụng Tất cả. Mặc định True</param>
    ''' <remarks>LoadtdbcWareHouseID(New C1.Win.C1List.C1Combo() {tdbcWareHouseIDFrom, tdbcWareHouseIDTo})</remarks>
    Public Sub LoadtdbcWareHouseID(ByVal tdbc() As C1.Win.C1List.C1Combo, ByVal bDisableZero As Boolean, Optional ByVal bAll As Boolean = True)
        Dim dtWareHouse As DataTable = ReturnTableWareHouseID(bAll, " DivisionID = " & SQLString(gsDivisionID) & IIf(bDisableZero, " And Disabled = 0", "").ToString)

        For i As Integer = 0 To tdbc.Length - 1
            LoadDataSource(tdbc(i), dtWareHouse.DefaultView.ToTable, gbUnicode)
            If bAll Then tdbc(i).SelectedIndex = 0
        Next
    End Sub

    Public Sub LoadtdbdWareHouseID(ByVal tdbd() As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByRef dtWareHouse As DataTable, ByVal bDisableZero As Boolean, Optional ByVal sDivsionID As String = "")
        If dtWareHouse Is Nothing Then dtWareHouse = ReturnTableWareHouseID(False, IIf(bDisableZero, " Disabled = 0", "").ToString)

        Dim sFilter As String = ""
        If sDivsionID <> "%" And sDivsionID <> "" Then sFilter = " DivisionID = " & SQLString(sDivsionID) & " or DivisionID = '%'"
        Dim dtTemp As DataTable = ReturnTableFilter(dtWareHouse, sFilter, True)
        For i As Integer = 0 To tdbd.Length - 1
            LoadDataSource(tdbd(i), dtTemp.DefaultView.ToTable, gbUnicode)
        Next
    End Sub

    Public Sub LoadtdbdWareHouseID(ByVal tdbd() As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal bDisableZero As Boolean)
        Dim dtWareHouse As DataTable = ReturnTableWareHouseID(False, " DivisionID = " & SQLString(gsDivisionID) & IIf(bDisableZero, " And Disabled = 0", "").ToString)

        For i As Integer = 0 To tdbd.Length - 1
            LoadDataSource(tdbd(i), dtWareHouse.DefaultView.ToTable, gbUnicode)
        Next
    End Sub

    Public Sub LoadtdbdWareHouseID(ByVal bDisableZero As Boolean, ByVal ParamArray tdbd() As C1.Win.C1TrueDBGrid.C1TrueDBDropdown)
        Dim dtWareHouse As DataTable = ReturnTableWareHouseID(False, " DivisionID = " & SQLString(gsDivisionID) & IIf(bDisableZero, " And Disabled = 0", "").ToString)

        For i As Integer = 0 To tdbd.Length - 1
            LoadDataSource(tdbd(i), dtWareHouse.DefaultView.ToTable, gbUnicode)
        Next
    End Sub
#End Region

#Region "Đổ nguồn Dự án và Hạng mục"
    Public gbUseD54 As Boolean = False

    Public Sub UseModuleD54()
        If geDisplayProduct = EnumDisplayProduct.L3ERP Then 'Update 18/11/2015
            'If Not CheckModuleG4(GetModuleWhenRuntime(2)) Then gbUseD54 = ExistRecord("Select top 1 1 From D54T0000  WITH(NOLOCK)")
            gbUseD54 = ExistRecord("Select top 1 1 From D54T0000  WITH(NOLOCK)")
        End If
    End Sub


    Public Function GetModuleWhenRuntime(ByVal iIndexMethod As Integer) As String
        Dim _st As StackTrace = New StackTrace()
        Dim sModule As String = ""
        Try
            If _st.GetFrames.Length > 0 Then
                Dim _sf As System.Diagnostics.StackFrame
                _sf = _st.GetFrames(iIndexMethod)
                Dim _mt As Reflection.MethodBase = _sf.GetMethod()
                sModule = _mt.ReflectedType.FullName
                If sModule.Length >= 3 Then sModule = sModule.Substring(0, 3)
            End If

        Catch ex As Exception
            Return ""
        End Try

        Return sModule

    End Function


    Public Sub VisibleProjectID(ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iSplit As Integer)
        tdbg.Splits(iSplit).DisplayColumns("ProjectID").Visible = gbUseD54
        tdbg.Splits(iSplit).DisplayColumns("ProjectName").Visible = gbUseD54
        tdbg.Splits(iSplit).DisplayColumns("TaskID").Visible = gbUseD54
        tdbg.Splits(iSplit).DisplayColumns("TaskName").Visible = gbUseD54
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD54P9000
    '# Created User: Nguyễn Thị Minh Hòa
    '# Created Date: 12/06/2015 02:12:40
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD54P9000(ByVal bIsPercent As Boolean, Optional ByVal sFormID As String = "", Optional ByVal sFormPermission As String = "", Optional IsTask As Byte = 1) As String
        Dim sSQL As String = ""
        If IsTask = 0 Then
            sSQL &= ("-- Do nguon Du an" & vbCrLf)
        Else
            sSQL &= ("-- Do nguon hang muc" & vbCrLf)
        End If
        sSQL &= "Exec D54P9000 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLNumber(bIsPercent) & COMMA 'IsPercent, tinyint, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'FromProjectID, varchar[50], NOT NULL
        sSQL &= SQLString("") & COMMA 'ToProjectID, varchar[50], NOT NULL
        sSQL &= SQLString(sFormID) & COMMA 'FormID, varchar[50], NOT NULL
        If sFormID <> "" Then
            sSQL &= SQLString(L3Left(sFormID, 3)) & COMMA 'ModuleID, varchar[50], NOT NULL
        Else
            sSQL &= SQLString("") & COMMA 'ModuleID, varchar[50], NOT NULL
        End If
        sSQL &= SQLString(sFormPermission) & COMMA 'FormPermission, varchar[50], NOT NULL
        sSQL &= SQLNumber(IsTask) 'IsTask, tinyint, NOT NULL
        Return sSQL
    End Function


    Public Function ReturnTableProject_TaskID(Optional ByVal bIsPercent As Boolean = False) As DataTable
        Return ReturnTableProject_TaskID(bIsPercent, "", "")
    End Function

    Public Function ReturnTableProject_TaskID(ByVal bIsPercent As Boolean, ByVal sFormID As String, ByVal sFormPermission As String) As DataTable
        Return ReturnDataTable(SQLStoreD54P9000(bIsPercent, sFormID, sFormPermission))
    End Function

    Public Sub LoadProject(ByVal ctrlProject As Control, Optional ByRef dtSource As DataTable = Nothing, Optional ByVal bIsPercent As Boolean = False)
        LoadProject(ctrlProject, dtSource, bIsPercent, "", "")

    End Sub

    Public Sub LoadProject(ByVal ctrlProject As Control, ByRef dtSource As DataTable, ByVal bIsPercent As Boolean, ByVal sFormID As String, ByVal sFormPermission As String)
        'If dtSource Is Nothing Then dtSource = ReturnDataTable(SQLStoreD54P9000(bIsPercent, sFormID, sFormPermission)) 
        'Dim dtProject As DataTable = Nothing
        'If dtSource.Rows.Count > 0 Then
        '    dtProject = dtSource.DefaultView.ToTable(True, New String() {"ProjectID", "ProjectName"})
        'Else
        '    dtProject = dtSource
        'End If
        Dim dtProject As DataTable = ReturnDataTable(SQLStoreD54P9000(bIsPercent, sFormID, sFormPermission, 0)) 'Tách câu đổ nguồn project và TaskID ra thành 2 SQL
        If TypeOf (ctrlProject) Is C1.Win.C1List.C1Combo Then
            LoadDataSource(CType(ctrlProject, C1.Win.C1List.C1Combo), dtProject, gbUnicode)
        ElseIf TypeOf (ctrlProject) Is C1.Win.C1TrueDBGrid.C1TrueDBDropdown Then
            LoadDataSource(CType(ctrlProject, C1.Win.C1TrueDBGrid.C1TrueDBDropdown), dtProject, gbUnicode)
        End If
    End Sub


    Public Sub LoadTask(ByVal ctrlTask As Control, Optional ByRef dtSource As DataTable = Nothing, Optional ByVal sProjectID As String = "%", Optional ByVal bIsPercent As Boolean = False)
        If dtSource Is Nothing Then dtSource = ReturnDataTable(SQLStoreD54P9000(bIsPercent))
        Dim dtTaskID As DataTable = Nothing
        If sProjectID = "%" Then
            dtTaskID = dtSource
        Else
            Dim sFilter As String = "ProjectID = '%' or ProjectID=" & SQLString(sProjectID)
            dtTaskID = ReturnTableFilter(dtSource, sFilter, True)
        End If
        If TypeOf (ctrlTask) Is C1.Win.C1List.C1Combo Then
            LoadDataSource(CType(ctrlTask, C1.Win.C1List.C1Combo), dtTaskID, gbUnicode)
        ElseIf TypeOf (ctrlTask) Is C1.Win.C1TrueDBGrid.C1TrueDBDropdown Then
            LoadDataSource(CType(ctrlTask, C1.Win.C1TrueDBGrid.C1TrueDBDropdown), dtTaskID, gbUnicode)
        End If
    End Sub
#End Region

#Region "Đổ nguồn Ngân sách và Hạng mục ngân sách"
    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD10P9000
    '# Created User: Nguyễn Lê Phương
    '# Created Date: 23/10/2013 11:42:28
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD10P9000(ByVal sFormID As String, Optional ByVal sKey01ID As String = "", Optional ByVal sKey02ID As String = "", Optional ByVal sKey03ID As String = "", Optional ByVal sKey04ID As String = "", Optional ByVal sKey05ID As String = "")
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon Ngan sach, hang muc ngan sach" & vbCrLf)
        sSQL &= "Exec D10P9000 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(sFormID) & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(sKey01ID) & COMMA 'Key01ID, varchar[150], NOT NULL
        sSQL &= SQLString(sKey02ID) & COMMA 'Key02ID, varchar[150], NOT NULL
        sSQL &= SQLString(sKey03ID) & COMMA 'Key03ID, varchar[150], NOT NULL
        sSQL &= SQLString(sKey04ID) & COMMA 'Key04ID, varchar[150], NOT NULL
        sSQL &= SQLString(sKey05ID) 'Key05ID, varchar[150], NOT NULL
        Return sSQL
    End Function

    Public Function ReturnTableBudgetID_BudgetItemID(ByVal sFormID As String) As DataTable
        Return ReturnDataTable(SQLStoreD10P9000(sFormID, "", "", "", "", ""))
    End Function

    Public Function ReturnTableBudgetID_BudgetItemID(ByVal sFormID As String, ByVal sKey01ID As String, ByVal sKey02ID As String, ByVal sKey03ID As String, ByVal sKey04ID As String, ByVal sKey05ID As String) As DataTable
        Return ReturnDataTable(SQLStoreD10P9000(sFormID, sKey01ID, sKey02ID, sKey03ID, sKey04ID, sKey05ID))
    End Function


    Public Sub LoadBudget(ByVal ctrlBudgetID As Control, ByVal sFormID As String, ByVal sKey01ID As String, ByVal sKey02ID As String, ByVal sKey03ID As String, ByVal sKey04ID As String, ByVal sKey05ID As String, Optional ByRef dtSource As DataTable = Nothing)
        If dtSource Is Nothing Then dtSource = ReturnDataTable(SQLStoreD10P9000(sFormID, sKey01ID, sKey02ID, sKey03ID, sKey04ID, sKey05ID))
        Dim dtBudgetID As DataTable = Nothing
        If dtSource.Rows.Count > 0 Then
            dtBudgetID = dtSource.DefaultView.ToTable(True, New String() {"BudgetID", "BudgetName"})
        Else
            dtBudgetID = dtSource
        End If
        If TypeOf (ctrlBudgetID) Is C1.Win.C1List.C1Combo Then
            LoadDataSource(CType(ctrlBudgetID, C1.Win.C1List.C1Combo), dtBudgetID, gbUnicode)
        ElseIf TypeOf (ctrlBudgetID) Is C1.Win.C1TrueDBGrid.C1TrueDBDropdown Then
            LoadDataSource(CType(ctrlBudgetID, C1.Win.C1TrueDBGrid.C1TrueDBDropdown), dtBudgetID, gbUnicode)
        End If
    End Sub

    Public Sub LoadBudget(ByVal ctrlBudgetID As Control, ByVal sFormID As String, Optional ByRef dtSource As DataTable = Nothing)
        LoadBudget(ctrlBudgetID, sFormID, "", "", "", "", "", dtSource)

    End Sub

    Public Sub LoadBudgetItem(ByVal ctrlBudgetItemID As Control, ByVal sFormID As String, ByVal sKey01ID As String, ByVal sKey02ID As String, ByVal sKey03ID As String, ByVal sKey04ID As String, ByVal sKey05ID As String, Optional ByRef dtSource As DataTable = Nothing, Optional ByVal sBudgetID As String = "%")
        If dtSource Is Nothing Then dtSource = ReturnDataTable(SQLStoreD10P9000(sFormID, sKey01ID, sKey02ID, sKey03ID, sKey04ID, sKey05ID))
        Dim dtBudgetItemID As DataTable = Nothing
        If sBudgetID = "%" Then
            dtBudgetItemID = dtSource
        Else
            Dim sFilter As String = "BudgetID = '%' or BudgetID=" & SQLString(sBudgetID)
            dtBudgetItemID = ReturnTableFilter(dtSource, sFilter, True)
        End If
        If TypeOf (ctrlBudgetItemID) Is C1.Win.C1List.C1Combo Then
            LoadDataSource(CType(ctrlBudgetItemID, C1.Win.C1List.C1Combo), dtBudgetItemID, gbUnicode)
        ElseIf TypeOf (ctrlBudgetItemID) Is C1.Win.C1TrueDBGrid.C1TrueDBDropdown Then
            LoadDataSource(CType(ctrlBudgetItemID, C1.Win.C1TrueDBGrid.C1TrueDBDropdown), dtBudgetItemID, gbUnicode)
        End If
    End Sub

    Public Sub LoadBudgetItem(ByVal ctrlBudgetItemID As Control, ByVal sFormID As String, Optional ByRef dtSource As DataTable = Nothing, Optional ByVal sBudgetID As String = "%")
        LoadBudgetItem(ctrlBudgetItemID, sFormID, "", "", "", "", "", dtSource, sBudgetID)
    End Sub

#End Region

#Region "Đổ nguồn cho Mã XDCB"

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P9000
    '# Created User: Nguyễn Lê Phương
    '# Created Date: 13/11/2012 02:52:06
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P9000() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon XDCB" & vbCrLf)
        sSQL &= "Exec D02P9000 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function

    Public Function ReturnTableCIP() As DataTable
        Return ReturnDataTable(SQLStoreD02P9000)
    End Function

    'Update 13/11/2012 by Lê Phương
    Public Sub LoadCIP(ByVal ctrlCIP As Control, Optional ByRef dtCIP As DataTable = Nothing)
        If dtCIP Is Nothing Then dtCIP = ReturnTableCIP()
        If TypeOf (ctrlCIP) Is C1.Win.C1List.C1Combo Then
            LoadDataSource(CType(ctrlCIP, C1.Win.C1List.C1Combo), dtCIP, gbUnicode)
        ElseIf TypeOf (ctrlCIP) Is C1.Win.C1TrueDBGrid.C1TrueDBDropdown Then
            LoadDataSource(CType(ctrlCIP, C1.Win.C1TrueDBGrid.C1TrueDBDropdown), dtCIP, gbUnicode)
        End If
    End Sub

#End Region

#Region "Tạo dropdown động"
    '''' <summary>
    '''' Trả về dropdown
    '''' </summary>
    '''' <param name="frm">form</param>
    '''' <param name="sDropdownName">tên dropdown cần tạo</param>
    ''' <param name="sValueMember">ValueMember của dropdown cần tạo</param>
    ''' <param name="sDisplayMember">DisplayMember của dropdown cần tạo</param>
    ''' <param name="sColumns">mảng các cột của dropdown cần tạo (khác ValueMember)</param>
    Public Function CreateDropDownID(ByVal frm As Form, ByVal sDropdownName As String, ByVal sValueMember As String, ByVal sDisplayMember As String, ByVal ParamArray sColumns() As String) As C1.Win.C1TrueDBGrid.C1TrueDBDropdown
        Dim tdbdID As New C1.Win.C1TrueDBGrid.C1TrueDBDropdown

        Dim Style1 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style
        Dim Style2 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style
        Dim Style3 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style
        Dim Style4 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style
        Dim Style5 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style
        Dim Style6 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style
        Dim Style7 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style
        Dim Style8 As C1.Win.C1TrueDBGrid.Style = New C1.Win.C1TrueDBGrid.Style
        'Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager() '(GetType(D09F2180))

        ''***********************************************
        CType(tdbdID, System.ComponentModel.ISupportInitialize).BeginInit()
        '***********************************************
        tdbdID.AllowColMove = False
        tdbdID.AllowColSelect = False
        tdbdID.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        tdbdID.AllowSort = False
        tdbdID.AlternatingRows = True
        tdbdID.CaptionHeight = 17
        tdbdID.CaptionStyle = Style1
        tdbdID.ColumnCaptionHeight = 17
        tdbdID.ColumnFooterHeight = 17

        tdbdID.EmptyRows = True
        tdbdID.ExtendRightColumn = True
        tdbdID.FetchRowStyles = False
        tdbdID.FooterStyle = Style3
        tdbdID.HeadingStyle = Style4
        tdbdID.HighLightRowStyle = Style5
        tdbdID.Location = New System.Drawing.Point(541, 90)
        '  tdbdID.Images.Add(CType(resources.GetObject("tdbdID.Images"), System.Drawing.Image))
        '  tdbdID.PropBag = resources.GetString("tdbdID.PropBag")

        tdbdID.RecordSelectorStyle = Style7
        tdbdID.RowDivider.Color = System.Drawing.Color.DarkGray
        tdbdID.RowDivider.Style = C1.Win.C1TrueDBGrid.LineStyleEnum.[Single]
        tdbdID.RowHeight = 15
        tdbdID.RowSubDividerColor = System.Drawing.Color.DarkGray
        tdbdID.ScrollTips = False
        tdbdID.Size = New System.Drawing.Size(400, 147)
        tdbdID.Style = Style8
        tdbdID.TabIndex = 30
        tdbdID.TabStop = False
        tdbdID.Visible = False
        tdbdID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        '***********************************************
        tdbdID.OddRowStyle.BackColor = Color.Beige
        tdbdID.EvenRowStyle.BackColor = Color.White
        '***********************************************
        tdbdID.Name = "tdbd" & sDropdownName
        tdbdID.ValueMember = sValueMember
        tdbdID.DisplayMember = sDisplayMember
        If sValueMember = sDisplayMember Then 'Hiển thị Mã, Lưu Mã
            tdbdID.ValueTranslate = False
        Else
            tdbdID.ValueTranslate = True 'Hiển thị Tên, Lưu Mã
        End If
        ''***********************************************
        'Me.Controls.Add(tdbdID)
        frm.Controls.Add(tdbdID)
        CType(tdbdID, System.ComponentModel.ISupportInitialize).EndInit()
        '***********************************************
        'Append 28/6/16 Thi Ánh vì lỗi khi sValueMember=Name, sDisplayMember =Name, sColumns = ID
        'CreateColumnDropdown(tdbdID, sValueMember, rL3("Ma"), 110)
        'If sValueMember <> sDisplayMember Then
        '    CreateColumnDropdown(tdbdID, sDisplayMember, rL3("Ten"), 200)
        'End If
        'If sColumns Is Nothing Then Return tdbdID '21/06/2013

        'For i As Integer = 0 To sColumns.Length - 1
        '    If i = 0 AndAlso sValueMember = sDisplayMember AndAlso sColumns(0) <> sDisplayMember Then 'Cột đầu tiên của mảng cột
        '        CreateColumnDropdown(tdbdID, sColumns(i), rL3("Ten"), 200)
        '    Else
        '        CreateColumnDropdown(tdbdID, sColumns(i))
        '    End If
        'Next'***********************************************
        If sColumns Is Nothing Then
            CreateColumnDropdown(tdbdID, sValueMember, rL3("Ma"), 110)
            CreateColumnDropdown(tdbdID, sDisplayMember, rL3("Ten"), 170)
        Else
            If L3FindArrString(sColumns, sValueMember) = False Then 'VD: "ID, ID, {Name}"; "Name, Name, {ID}"; '"ID, Name, {Name}"
                If sValueMember.EndsWith("ID") Then CreateColumnDropdown(tdbdID, sValueMember, rL3("Ma"), 110) 'Add cột mã trước
            End If
            If L3FindArrString(sColumns, sDisplayMember) = False Then 'VD: "ID, Name, {Value}"
                CreateColumnDropdown(tdbdID, sDisplayMember, rL3("Ten"), 170) 'Add cột mã trước
            End If

            For i As Integer = 0 To sColumns.Length - 1
                CreateColumnDropdown(tdbdID, sColumns(i))
            Next
            Try
                If tdbdID.Columns.IndexOf(tdbdID.Columns(sValueMember)) > -1 Then

                End If
            Catch ex As Exception
                CreateColumnDropdown(tdbdID, sValueMember, "", 170) 'Cho "Name, Name, {ID}"
            End Try
            Try
                If tdbdID.Columns.IndexOf(tdbdID.Columns(sDisplayMember)) > -1 Then

                End If
            Catch ex As Exception
                CreateColumnDropdown(tdbdID, sDisplayMember, "", 170) 'Cho "Name, Name, {ID}"
            End Try
        End If
        Return tdbdID
    End Function

    Private Sub CreateColumnDropdown(ByVal tdbdID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sDataField As String, Optional ByVal sCaption As String = "", Optional ByVal iWidth As Integer = 110, Optional ByVal bVisible As Boolean = True)
        Try
            If tdbdID.Columns.IndexOf(tdbdID.Columns(sDataField)) > -1 Then Exit Sub
        Catch ex As Exception 'Bổ sung kiểm tra tồn tại cột 28/06/20016 Thị Ánh
            Dim dc As New C1.Win.C1TrueDBGrid.C1DropDataColumn
            dc.DataField = sDataField
            tdbdID.Columns.Add(dc)

            Dim sDes As String = sCaption
            If sDes = "" Then
                If sDataField.EndsWith("ID") OrElse sDataField.EndsWith("Code") Then
                    sDes = rL3("Ma")
                    iWidth = 110
                ElseIf sDataField.EndsWith("Name") OrElse sDataField.EndsWith("Description") OrElse sDataField.EndsWith("Notes") Then
                    sDes = rL3("Ten")
                    iWidth = 170
                Else
                    sDes = sDataField
                End If
            End If

            tdbdID.Columns(dc.DataField).Caption = sDes ' IIf(sCaption = "", sDataField, sCaption).ToString

            tdbdID.DisplayColumns(dc.DataField).HeadingStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
            tdbdID.DisplayColumns(dc.DataField).HeadingStyle.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
            tdbdID.DisplayColumns(dc.DataField).Style.Font = FontUnicode(gbUnicode)
            tdbdID.DisplayColumns(dc.DataField).Visible = bVisible
            tdbdID.DisplayColumns(dc.DataField).Width = iWidth
        End Try


    End Sub

#End Region

#Region "Đổ nguồn cho Khoản mục, kiểm tra theo Phân quyền Đơn vị"

    Public Function ReturnTableAnaIDForDivision(ByVal bAddnew As Boolean, ByVal bUnicode As Boolean, ByVal sAnaCategoryID As String) As DataTable
        Dim sSQL As String = "--Do nguon Khoan muc " & vbCrLf
        If bAddnew Then
            sSQL &= "Select '+' as AnaID, " & NewName & " As AnaName, '+' as AnaCategoryID, 0 AS DisplayOrder " & vbCrLf
            sSQL &= "Union All " & vbCrLf
        End If
        sSQL &= "Select AnaID, AnaName" & UnicodeJoin(bUnicode) & " as AnaName, AnaCategoryID, 1 AS DisplayOrder " & vbCrLf
        sSQL &= "From D91T0051 WITH(NOLOCK) Where Disabled = 0 " & vbCrLf
        If sAnaCategoryID = "" Then
            sSQL &= "And AnaCategoryID like 'K%' " & vbCrLf
        Else
            sSQL &= "And AnaCategoryID = " & SQLString(sAnaCategoryID) & vbCrLf
        End If
        'Update 06/01/2014: Kiểm tra quyền Đơn vị sử dụng
        sSQL &= " And ((ISNULL(StrDivisionID,'')= '' OR CHARINDEX(';" & gsDivisionID & ";' , ';' +ISNULL(StrDivisionID,'') +';') <> 0))	" & vbCrLf
        sSQL &= " Order by DisplayOrder, AnaID"
        Return ReturnDataTable(sSQL)
    End Function

    Public Function ReturnTableAnaIDForDivision(ByVal bAddnew As Boolean, ByVal bUnicode As Boolean) As DataTable
        Return ReturnTableAnaIDForDivision(bAddnew, bUnicode, "")
    End Function

    Public Function ReturnTableAnaIDForDivision(ByVal bAddnew As Boolean) As DataTable
        Return ReturnTableAnaIDForDivision(bAddnew, gbUnicode, "")
    End Function

    Public Function ReturnTableAnaIDForDivision() As DataTable
        Return ReturnTableAnaIDForDivision(False, gbUnicode, "")
    End Function

    '''' <summary>
    '''' Hàm mới: Đổ nguồn cho 10 khoản mục kiểm tra KM nào có dùng thì đổ nguồn cho Dropdown, Cải tiến ít tham số
    '''' </summary>
    '''' <remarks>Đặt sau hàm LoadCaption và gắn Dropdown</remarks>
    '<DebuggerStepThrough()> _
    Public Sub LoadTDBDropDownAnaForDivision(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_Ana01ID As Integer, Optional ByVal bAddNew As Boolean = False, Optional ByRef dt As DataTable = Nothing)
        If dt Is Nothing Then dt = ReturnTableAnaIDForDivision(bAddNew, gbUnicode, "")
        For i As Integer = 0 To 9
            Dim index As Integer = COL_Ana01ID + i
            If tdbg.Columns(index).DropDown Is Nothing Then
                D99C0008.MsgL3(tdbg.Columns(index).DataField & " không có dropdown")
                Continue For
            End If
            If Convert.ToBoolean(tdbg.Columns(index).Tag) Then LoadDataSource(tdbg.Columns(index).DropDown, ReturnTableFilter(dt, "AnaCategoryID='K" & (i + 1).ToString("00") & "'" & " or AnaCategoryID='+'"), gbUnicode)
        Next
    End Sub

    '''' <summary>
    '''' Hàm mới: Đổ nguồn cho 10 khoản mục kiểm tra KM nào có dùng thì đổ nguồn cho Dropdown
    '''' </summary>
    '''' <remarks>Truyền 10 khoản mục cần đổ nguồn vào</remarks>
    '<DebuggerStepThrough()> _
    Public Sub LoadTDBDropDownAnaForDivision(ByVal tdbdAna01ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna02ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna03ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna04ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna05ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna06ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna07ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna08ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna09ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna10ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, _
  ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_Ana01ID As Integer, Optional ByVal bUseUnicode As Boolean = False, Optional ByVal bAddNew As Boolean = False, Optional ByRef dt As DataTable = Nothing, Optional ByVal bPerDivision As Boolean = False)
        If dt Is Nothing Then dt = ReturnTableAnaIDForDivision(bAddNew, bUseUnicode, "")

        If Convert.ToBoolean(tdbg.Columns(COL_Ana01ID).Tag) Then LoadDataSource(tdbdAna01ID, ReturnTableFilter(dt, "AnaCategoryID='K01' or AnaCategoryID='+'"), bUseUnicode)
        If Convert.ToBoolean(tdbg.Columns(COL_Ana01ID + 1).Tag) Then LoadDataSource(tdbdAna02ID, ReturnTableFilter(dt, "AnaCategoryID='K02' or AnaCategoryID='+'"), bUseUnicode)
        If Convert.ToBoolean(tdbg.Columns(COL_Ana01ID + 2).Tag) Then LoadDataSource(tdbdAna03ID, ReturnTableFilter(dt, "AnaCategoryID='K03' or AnaCategoryID='+'"), bUseUnicode)
        If Convert.ToBoolean(tdbg.Columns(COL_Ana01ID + 3).Tag) Then LoadDataSource(tdbdAna04ID, ReturnTableFilter(dt, "AnaCategoryID='K04' or AnaCategoryID='+'"), bUseUnicode)
        If Convert.ToBoolean(tdbg.Columns(COL_Ana01ID + 4).Tag) Then LoadDataSource(tdbdAna05ID, ReturnTableFilter(dt, "AnaCategoryID='K05' or AnaCategoryID='+'"), bUseUnicode)
        If Convert.ToBoolean(tdbg.Columns(COL_Ana01ID + 5).Tag) Then LoadDataSource(tdbdAna06ID, ReturnTableFilter(dt, "AnaCategoryID='K06' or AnaCategoryID='+'"), bUseUnicode)
        If Convert.ToBoolean(tdbg.Columns(COL_Ana01ID + 6).Tag) Then LoadDataSource(tdbdAna07ID, ReturnTableFilter(dt, "AnaCategoryID='K07' or AnaCategoryID='+'"), bUseUnicode)
        If Convert.ToBoolean(tdbg.Columns(COL_Ana01ID + 7).Tag) Then LoadDataSource(tdbdAna08ID, ReturnTableFilter(dt, "AnaCategoryID='K08' or AnaCategoryID='+'"), bUseUnicode)
        If Convert.ToBoolean(tdbg.Columns(COL_Ana01ID + 8).Tag) Then LoadDataSource(tdbdAna09ID, ReturnTableFilter(dt, "AnaCategoryID='K09' or AnaCategoryID='+'"), bUseUnicode)
        If Convert.ToBoolean(tdbg.Columns(COL_Ana01ID + 9).Tag) Then LoadDataSource(tdbdAna10ID, ReturnTableFilter(dt, "AnaCategoryID='K10' or AnaCategoryID='+'"), bUseUnicode)
    End Sub


    '''' <summary>
    '''' Đổ nguồn cho 10 khoản mục (Hàm cũ BỎ)
    '''' </summary>
    '''' <remarks>Truyền 10 khoản mục cần đổ nguồn vào</remarks>
    '<DebuggerStepThrough()> _
    Public Sub LoadTDBDropDownAnaForDivision(ByVal tdbdAna01ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna02ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna03ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna04ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna05ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna06ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna07ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna08ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna09ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdAna10ID As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False, Optional ByVal bPerDivision As Boolean = False)
        Dim dt As DataTable
        Dim sSQL As String = ""

        sSQL = "Select AnaID, AnaName" & UnicodeJoin(bUseUnicode) & " as AnaName, AnaCategoryID " & vbCrLf
        sSQL &= "From D91T0051 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where Disabled = 0 And AnaCategoryID like 'K%' " & vbCrLf
        'Update 06/01/2014: Kiểm tra quyền Đơn vị sử dụng
        ' sSQL &= "And (StrDivisionID = '' OR CHARINDEX(" & SQLString(gsDivisionID) & ",StrDivisionID) > 0) " & vbCrLf
        sSQL &= "  And ((ISNULL(StrDivisionID,'')= '' OR CHARINDEX(';" & gsDivisionID & ";' , ';' +ISNULL(StrDivisionID,'') +';') <> 0))	" & vbCrLf
        sSQL &= "Order by AnaID"
        dt = ReturnDataTable(sSQL)

        LoadDataSource(tdbdAna01ID, ReturnTableFilter(dt, "AnaCategoryID='K01'"), bUseUnicode)
        LoadDataSource(tdbdAna02ID, ReturnTableFilter(dt, "AnaCategoryID='K02'"), bUseUnicode)
        LoadDataSource(tdbdAna03ID, ReturnTableFilter(dt, "AnaCategoryID='K03'"), bUseUnicode)
        LoadDataSource(tdbdAna04ID, ReturnTableFilter(dt, "AnaCategoryID='K04'"), bUseUnicode)
        LoadDataSource(tdbdAna05ID, ReturnTableFilter(dt, "AnaCategoryID='K05'"), bUseUnicode)
        LoadDataSource(tdbdAna06ID, ReturnTableFilter(dt, "AnaCategoryID='K06'"), bUseUnicode)
        LoadDataSource(tdbdAna07ID, ReturnTableFilter(dt, "AnaCategoryID='K07'"), bUseUnicode)
        LoadDataSource(tdbdAna08ID, ReturnTableFilter(dt, "AnaCategoryID='K08'"), bUseUnicode)
        LoadDataSource(tdbdAna09ID, ReturnTableFilter(dt, "AnaCategoryID='K09'"), bUseUnicode)
        LoadDataSource(tdbdAna10ID, ReturnTableFilter(dt, "AnaCategoryID='K10'"), bUseUnicode)
    End Sub

#End Region

#Region "Đổ nguồn Tài khoản (Chỉ gồm Tài khoản trong bảng) kiểm tra Phân quyền theo Đơn vị"

    ''' <summary>
    ''' Trả ra Table Tài khoản có chứa % và kiểm tra Phân quyền theo Đơn vị và load theo Mã hàng
    ''' </summary>
    ''' <param name="sClauseWhere">Nếu có điều kiện lọc thì truyền vào (ví dụ: GroupID in ('1', '13'))</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnTableAccountIDWithInventoryID(ByVal sInventoryID As String, ByVal sClauseWhere As String, ByVal bAll As Boolean) As DataTable
        'Update 23/09/2012: incident 80005 sửa câu SQL sang Store 
        'Lưu ý: nếu load TK theo mã hàng thì công kiểm tra TK trong bảng hay ngoại bảng
        Return ReturnDataTable(SQLFinance.SQLStoreD91P9302("Theo mã hàng", sClauseWhere, L3Byte(bAll), "IN", sInventoryID, "", "", ""))

    End Function

    Public Function ReturnTableAccountIDWithInventoryID(ByVal sInventoryID As String, ByVal sClauseWhere As String, ByVal bAll As Boolean, ByVal sEditAccountID As String) As DataTable
        'Update 23/09/2012: incident 80005 sửa câu SQL sang Store 
        'Lưu ý: nếu load TK theo mã hàng thì công kiểm tra TK trong bảng hay ngoại bảng
        Return ReturnDataTable(SQLFinance.SQLStoreD91P9302("Theo mã hàng", sClauseWhere, L3Byte(bAll), "IN", sInventoryID, sEditAccountID, "", ""))

    End Function

    ''' <summary>
    ''' Trả ra Table Tài khoản có chứa % và kiểm tra Phân quyền theo Đơn vị
    ''' </summary>
    ''' <param name="sClauseWhere">Nếu có điều kiện lọc thì truyền vào (ví dụ: GroupID in ('1', '13'))</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ReturnTableAccountID(ByVal sClauseWhere As String, ByVal bAll As Boolean) As DataTable
        'Dim sUnicode As String = ""
        'Dim sLanguage As String = ""
        'UnicodeAllString(sUnicode, sLanguage, bUseUnicode)

        ''***************
        'Dim sSQL As String = "--Do nguon Tai khoan" & IIf(bAll, " co % ", "").ToString & vbCrLf
        'If bAll Then
        '    sSQL &= "Select '%' as AccountID, " & sLanguage & " as AccountName, '' As GroupID ,0 AS DisplayOrder, 0 AS IsMappingAnalysis " & vbCrLf
        '    sSQL &= "Union All " & vbCrLf
        'End If
        'sSQL &= "Select AccountID, " & IIf(geLanguage = EnumLanguage.Vietnamese, "AccountName", "AccountName01").ToString & sUnicode & " as AccountName, GroupID, 1 AS DisplayOrder, IsMappingAnalysis "
        'sSQL &= "From D90T0001 WITH(NOLOCK) " & vbCrLf
        'sSQL &= "Where Disabled = 0 And OffAccount = 0 " & vbCrLf
        ''sSQL &= "And (StrDivisionID = '' OR CHARINDEX(" & SQLString(gsDivisionID) & ",StrDivisionID) > 0) " & vbCrLf
        'sSQL &= " And ((ISNULL(StrDivisionID,'')= '' OR CHARINDEX(';" & gsDivisionID & ";' , ';' +ISNULL(StrDivisionID,'') +';') <> 0))	" & vbCrLf
        'If sClauseWhere <> "" Then
        '    sSQL &= "And (" & sClauseWhere & ") " & vbCrLf
        'End If
        'sSQL &= "Order By DisplayOrder, AccountID"


       
        'Return ReturnDataTable(sSQL)

        'Update 23/09/2012: incident 80005 sửa câu SQL sang Store 
        If sClauseWhere = "" Then
            sClauseWhere = "OffAccount = 0"
        Else
            sClauseWhere = "OffAccount = 0 And (" & sClauseWhere & ")"
        End If

        Return ReturnDataTable(SQLFinance.SQLStoreD91P9302("Trong bang theo OffAccount = 0", sClauseWhere, L3Byte(bAll), "", "", "", "", ""))

    End Function


    ''#---------------------------------------------------------------------------------------------------
    ''# Title: SQLStoreD91P9302
    ''# Created User: Nguyễn Thị Minh Hòa
    ''# Created Date: 25/09/2015 04:55:02
    ''#---------------------------------------------------------------------------------------------------
    'Private Function SQLStoreD91P9302(ByVal sCommand As String, ByVal sWhere As String, ByVal IsAll As Byte, ByVal sType As String, ByVal sInventoryID As String, ByVal sEditAccountID As String, ByVal sModuleID As String, ByVal sFormID As String) As String

    '    Dim sSQL As String = ""
    '    sSQL &= ("-- Do nguon chuan cho Tai khoan " & sCommand & vbCrLf)
    '    sSQL &= "Exec D91P9302 "
    '    sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
    '    sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
    '    sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
    '    sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[2], NOT NULL
    '    sSQL &= SQLString(sModuleID) & COMMA 'ModuleID, varchar[20], NOT NULL
    '    sSQL &= SQLString(sFormID) & COMMA 'FormID, varchar[20], NOT NULL
    '    sSQL &= SQLString(sWhere) & COMMA 'sWhere, varchar[8000], NOT NULL
    '    sSQL &= SQLNumber(IsAll) & COMMA 'IsAll, tinyint, NOT NULL
    '    sSQL &= SQLString(sType) & COMMA 'Type, varchar[20], NOT NULL
    '    sSQL &= SQLString(sInventoryID) & COMMA 'InventoryID, varchar[50], NOT NULL
    '    sSQL &= SQLString(sEditAccountID) 'AccountID, varchar[50], NOT NULL
    '    Return sSQL
    'End Function




    '''' <summary>
    '''' Trả ra Table Tài khoản kiểm tra Phân quyền theo Đơn vị
    '''' </summary>
    '''' <param name="sClauseWhere">Nếu có điều kiện lọc thì truyền vào (ví dụ: GroupID in ('1', '13')) </param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    '<DebuggerStepThrough()> _
    Public Function ReturnTableAccountIDForDivision(ByVal sClauseWhere As String, ByVal bUseUnicode As Boolean) As DataTable
        Return ReturnTableAccountID(sClauseWhere, False)
    End Function

    Public Function ReturnTableAccountIDForDivision(ByVal sClauseWhere As String) As DataTable
        ' Trả ra Table Tài khoản kiểm tra Phân quyền theo Đơn vị
        Return ReturnTableAccountIDForDivision(sClauseWhere, gbUnicode)
    End Function

    Public Function ReturnTableAccountIDForDivision() As DataTable
        ' Trả ra Table Tài khoản kiểm tra Phân quyền theo Đơn vị
        Return ReturnTableAccountIDForDivision("", gbUnicode)
    End Function

    ''' <summary>
    ''' Trả ra Table Tài khoản có chứa % và kiểm tra Phân quyền theo Đơn vị
    ''' </summary>
    ''' <param name="sClauseWhere">Nếu có điều kiện lọc thì truyền vào (ví dụ: GroupID in ('1', '13'))</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnTableAccountIDAllForDivision(ByVal sClauseWhere As String, ByVal bUseUnicode As Boolean) As DataTable
        Return ReturnTableAccountID(sClauseWhere, True)
    End Function


    Public Function ReturnTableAccountIDAllForDivision(ByVal sClauseWhere As String) As DataTable
        Return ReturnTableAccountIDAllForDivision(sClauseWhere, gbUnicode)
    End Function

    Public Function ReturnTableAccountIDAllForDivision() As DataTable
        Return ReturnTableAccountIDAllForDivision("", gbUnicode)
    End Function

    ''' <summary>
    ''' Load Combo Tài khoản theo Mã hàng
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDWithInventoryID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal sInventoryID As String, ByVal bAll As Boolean, ByVal sClauseWhere As String, ByVal sModuleID As String, ByVal sFormID As String)
        LoadDataSource(tdbc, ReturnTableAccountIDWithInventoryID(sInventoryID, sClauseWhere, bAll), gbUnicode)

    End Sub


    ''' <summary>
    ''' Load 2 Combo Tài khoản theo Mã hàng
    ''' </summary>
    ''' <param name="tdbcFrom"></param>
    ''' <param name="tdbcTo"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
Public Sub LoadAccountIDWithInventoryID(ByVal tdbcFrom As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo, ByVal sInventoryID As String, ByVal bAll As Boolean, ByVal sClauseWhere As String, ByVal sModuleID As String, ByVal sFormID As String)
        Dim dt As DataTable
        dt = ReturnTableAccountIDWithInventoryID(sInventoryID, sClauseWhere, bAll)
        LoadDataSource(tdbcFrom, dt, gbUnicode)
        LoadDataSource(tdbcTo, dt.DefaultView.ToTable, gbUnicode)
    End Sub

    ''' <summary>
    ''' Load Combo Tài khoản
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDForDivision(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, ReturnTableAccountIDForDivision("", bUseUnicode), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load Combo Tài khoản có chứa %
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDAllForDivision(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, ReturnTableAccountIDAllForDivision("", bUseUnicode), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 Combo Tài khoản
    ''' </summary>
    ''' <param name="tdbcFrom"></param>
    ''' <param name="tdbcTo"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDForDivision(ByVal tdbcFrom As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDForDivision("", bUseUnicode)
        LoadDataSource(tdbcFrom, dt, bUseUnicode)
        LoadDataSource(tdbcTo, dt.DefaultView.ToTable, bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 Combo Tài khoản có chứa %
    ''' </summary>
    ''' <param name="tdbcFrom"></param>
    ''' <param name="tdbcTo"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDAllForDivision(ByVal tdbcFrom As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDAllForDivision("", gbUnicode)
        LoadDataSource(tdbcFrom, dt, gbUnicode)
        LoadDataSource(tdbcTo, dt.DefaultView.ToTable, gbUnicode)
    End Sub

    ''' <summary>
    ''' Load Combo Tài khoản có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <param name="sClauseWhere"> Điều kiện where (ví dụ: GroupID in ('1', '13'))</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDForDivision(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, ReturnTableAccountIDForDivision(sClauseWhere, gbUnicode), gbUnicode)
    End Sub

    ''' <summary>
    ''' Load Combo Tài khoản có chứa % có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <param name="sClauseWhere"> Điều kiện where (ví dụ: GroupID in ('1', '13'))</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDAllForDivision(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, ReturnTableAccountIDAllForDivision(sClauseWhere, gbUnicode), gbUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 combo Tài khoản có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbcFrom"></param>
    ''' <param name="tdbcTo"></param>
    ''' <param name="sClauseWhere">Điều kiện where (ví dụ: GroupID in ('1', '13'))</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDForDivision(ByVal tdbcFrom As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDForDivision(sClauseWhere, gbUnicode)
        LoadDataSource(tdbcFrom, dt, gbUnicode)
        LoadDataSource(tdbcTo, dt.DefaultView.ToTable, gbUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 combo Tài khoản có chứa % có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbcFrom"></param>
    ''' <param name="tdbcTo"></param>
    ''' <param name="sClauseWhere">Điều kiện where (ví dụ: GroupID in ('1', '13'))</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDAllForDivision(ByVal tdbcFrom As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDAllForDivision(sClauseWhere, gbUnicode)
        LoadDataSource(tdbcFrom, dt, gbUnicode)
        LoadDataSource(tdbcTo, dt.DefaultView.ToTable, gbUnicode)
    End Sub

    ''' <summary>
    ''' Load combo Tài khoản (áp dụng cho màn hình có nhiều combo có cùng nguồn dữ liệu )
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <param name="dt">Bảng dữ liệu tài khoản</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
        Public Sub LoadAccountIDForDivision(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dt As DataTable, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, dt.DefaultView.ToTable, gbUnicode)
    End Sub

    ''' <summary>
    ''' Load Dropdown Tài khoản
    ''' </summary>
    ''' <param name="tdbd"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDForDivision(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, ReturnTableAccountIDForDivision("", gbUnicode), gbUnicode)
    End Sub

    ''' <summary>
    ''' Load Dropdown Tài khoản có chứa %
    ''' </summary>
    ''' <param name="tdbd"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDAllForDivision(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, ReturnTableAccountIDAllForDivision("", gbUnicode), gbUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 Dropdown Tài khoản
    ''' </summary>
    ''' <param name="tdbdFrom"></param>
    ''' <param name="tdbdTo"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDForDivision(ByVal tdbdFrom As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdTo As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDForDivision("", gbUnicode)
        LoadDataSource(tdbdFrom, dt, gbUnicode)
        LoadDataSource(tdbdTo, dt.DefaultView.ToTable, gbUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 Dropdown Tài khoản có chứa %
    ''' </summary>
    ''' <param name="tdbdFrom"></param>
    ''' <param name="tdbdTo"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDAllForDivision(ByVal tdbdFrom As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdTo As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDAllForDivision("", gbUnicode)
        LoadDataSource(tdbdFrom, dt, gbUnicode)
        LoadDataSource(tdbdTo, dt.DefaultView.ToTable, gbUnicode)
    End Sub

    ''' <summary>
    ''' Load Dropdown Tài khoản có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbd"></param>
    ''' <param name="sClauseWhere"> Điều kiện where (ví dụ: GroupID in ('1', '13'))</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDForDivision(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, ReturnTableAccountIDForDivision(sClauseWhere, gbUnicode), gbUnicode)
    End Sub

    ''' <summary>
    ''' Load Dropdown Tài khoản có chứa % có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbd"></param>
    ''' <param name="sClauseWhere"> Điều kiện where (ví dụ: GroupID in ('1', '13'))</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDAllForDivision(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, ReturnTableAccountIDAllForDivision(sClauseWhere, gbUnicode), gbUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 Dropdown Tài khoản có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbdFrom"></param>
    ''' <param name="tdbdTo"></param>
    ''' <param name="sClauseWhere"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDForDivision(ByVal tdbdFrom As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdTo As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDForDivision(sClauseWhere, gbUnicode)
        LoadDataSource(tdbdFrom, dt, gbUnicode)
        LoadDataSource(tdbdTo, dt.DefaultView.ToTable, gbUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 Dropdown Tài khoản có chứa % có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbdFrom"></param>
    ''' <param name="tdbdTo"></param>
    ''' <param name="sClauseWhere"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDAllForDivision(ByVal tdbdFrom As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdTo As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDAllForDivision(sClauseWhere, gbUnicode)
        LoadDataSource(tdbdFrom, dt, gbUnicode)
        LoadDataSource(tdbdTo, dt.DefaultView.ToTable, gbUnicode)
    End Sub

    ''' <summary>
    ''' Load Dropdown Tài khoản (áp dụng cho màn hình có nhiều Dropdown có cùng nguồn dữ liệu)
    ''' </summary>
    ''' <param name="tdbd"></param>
    ''' <param name="dt">Bảng dữ liệu Tài khoản</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDForDivision(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dt As DataTable, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, dt.Copy, gbUnicode)
    End Sub

#End Region

#Region "Đổ nguồn Tài khoản (Bao gồm Tài khoản trong bảng và ngoại bảng) kiểm tra Phân quyền theo Đơn vị"

    '    ''' <summary>
    '    ''' Trả ra Table Tài khoản có chứa % (bao gồm Tài khoản trong và ngoại bảng)
    '    ''' </summary>
    '    ''' <param name="sClauseWhere"></param>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    <DebuggerStepThrough()> _
    '    Private Function ReturnTableAccountIDAndOffAccountAllGeneralForDivision(ByVal sClauseWhere As String, ByVal bUseUnicode As Boolean, ByVal bAll As Boolean) As DataTable
    '
    '        Dim sUnicode As String = ""
    '        Dim sLanguage As String = ""
    '        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
    '        Dim sSQL As String = "--Do nguon Tai khoan" & IIf(bAll, " co %", "").ToString & " (gom TK trong va ngoai bang) " & vbCrLf
    '        If bAll Then
    '            sSQL &= "Select '%' as AccountID, " & sLanguage & " as AccountName, '' as GroupID, 0 as OffAccount, 0 as DisplayOrder, 0 AS IsMappingAnalysis " & vbCrLf
    '            sSQL &= "Union All " & vbCrLf
    '        End If
    '        sSQL &= "Select AccountID, " & IIf(geLanguage = EnumLanguage.Vietnamese, "AccountName", "AccountName01").ToString & sUnicode & " as AccountName, GroupID, OffAccount, 1 as DisplayOrder, IsMappingAnalysis " & vbCrLf
    '        sSQL &= "From D90T0001 WITH(NOLOCK) " & vbCrLf
    '        sSQL &= "Where Disabled = 0 " & vbCrLf
    '        'sSQL &= "And (StrDivisionID = '' OR CHARINDEX(" & SQLString(gsDivisionID) & ",StrDivisionID) > 0) " & vbCrLf
    '        sSQL &= " And ((ISNULL(StrDivisionID,'')= '' OR CHARINDEX(';" & gsDivisionID & ";' , ';' +ISNULL(StrDivisionID,'') +';') <> 0))	" & vbCrLf
    '        If sClauseWhere <> "" Then
    '            sSQL &= "And (" & sClauseWhere & ") " & vbCrLf
    '        End If
    '        sSQL &= "Order By DisplayOrder, AccountID "
    '
    '        Return ReturnDataTable(sSQL)
    '    End Function


    ''' <summary>
    ''' Trả ra Table Tài khoản có chứa % (bao gồm Tài khoản trong và ngoại bảng)
    ''' </summary>
    ''' <param name="sClauseWhere"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnTableOffAccountID(ByVal sClauseWhere As String, ByVal bAll As Boolean, ByVal sEditAccountID As String) As DataTable
        'Update 23/09/2012: incident 80005 sửa câu SQL sang Store 
        Return ReturnDataTable(SQLFinance.SQLStoreD91P9302("Trong bang va ngoai bang", sClauseWhere, L3Byte(bAll), "", "", sEditAccountID, "", ""))

    End Function


    ''' <summary>
    ''' Trả ra Table Tài khoản (bao gồm Tài khoản trong và ngoại bảng)  kiểm tra Phân quyền theo Đơn vị
    ''' </summary>
    ''' <param name="sClauseWhere"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnTableAccountIDAndOffAccountForDivision(ByVal sClauseWhere As String, ByVal bUseUnicode As Boolean) As DataTable
        'Return ReturnTableAccountIDAndOffAccountAllGeneralForDivision(sClauseWhere, bUseUnicode, False)
        'Update 23/09/2012: incident 80005 sửa câu SQL sang Store
        Return ReturnTableOffAccountID(sClauseWhere, False, "")
    End Function

    Public Function ReturnTableAccountIDAndOffAccountForDivision(ByVal sClauseWhere As String) As DataTable
        Return ReturnTableAccountIDAndOffAccountForDivision(sClauseWhere, False)
        
    End Function

    Public Function ReturnTableAccountIDAndOffAccountForDivision() As DataTable
        Return ReturnTableAccountIDAndOffAccountForDivision("", False)
    End Function

    ''' <summary>
    ''' Trả ra Table Tài khoản có chứa % (bao gồm Tài khoản trong và ngoại bảng)
    ''' </summary>
    ''' <param name="sClauseWhere"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnTableAccountIDAndOffAccountAllForDivision(ByVal sClauseWhere As String, ByVal bUseUnicode As Boolean) As DataTable
        'Return ReturnTableAccountIDAndOffAccountAllGeneralForDivision(sClauseWhere, bUseUnicode, True)
        'Update 23/09/2012: incident 80005 sửa câu SQL sang Store
        Return ReturnTableOffAccountID(sClauseWhere, True, "")
    End Function

    Public Function ReturnTableAccountIDAndOffAccountAllForDivision(ByVal sClauseWhere As String) As DataTable
        Return ReturnTableAccountIDAndOffAccountAllForDivision(sClauseWhere, False)
    End Function

    Public Function ReturnTableAccountIDAndOffAccountAllForDivision() As DataTable
        Return ReturnTableAccountIDAndOffAccountAllForDivision("", False)
    End Function
    ''' <summary>
    ''' Load Combo Tài khoản (bao gồm Tài khoản trong và ngoại bảng)
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDAndOffAccountForDivision(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, ReturnTableAccountIDAndOffAccountForDivision("", bUseUnicode), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load Combo Tài khoản có chứa % (bao gồm Tài khoản trong và ngoại bảng)
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDAndOffAccountAllForDivision(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, ReturnTableAccountIDAndOffAccountAllForDivision("", bUseUnicode), bUseUnicode)
    End Sub

    Public Sub LoadAccountIDAndOffAccountAllForDivision(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, ReturnTableAccountIDAndOffAccountAllForDivision(sClauseWhere, bUseUnicode), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 Combo Tài khoản (bao gồm Tài khoản trong và ngoại bảng)
    ''' </summary>
    ''' <param name="tdbcFrom"></param>
    ''' <param name="tdbcTo"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDAndOffAccountForDivision(ByVal tdbcFrom As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDAndOffAccountForDivision("", bUseUnicode)
        LoadDataSource(tdbcFrom, dt, bUseUnicode)
        LoadDataSource(tdbcTo, dt.DefaultView.ToTable, bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 Combo Tài khoản có chứa % (bao gồm Tài khoản trong và ngoại bảng)
    ''' </summary>
    ''' <param name="tdbcFrom"></param>
    ''' <param name="tdbcTo"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDAndOffAccountAllForDivision(ByVal tdbcFrom As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDAndOffAccountAllForDivision("", bUseUnicode)
        LoadDataSource(tdbcFrom, dt, bUseUnicode)
        LoadDataSource(tdbcTo, dt.DefaultView.ToTable, bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load Combo Tài khoản (bao gồm Tài khoản trong và ngoại bảng) có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <param name="sClauseWhere"> Điều kiện where (ví dụ: GroupID in ('1', '13'))</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDAndOffAccountForDivision(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, ReturnTableAccountIDAndOffAccountForDivision(sClauseWhere, bUseUnicode), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 combo Tài khoản (bao gồm Tài khoản trong và ngoại bảng) có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbcFrom"></param>
    ''' <param name="tdbcTo"></param>
    ''' <param name="sClauseWhere"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDAndOffAccountForDivision(ByVal tdbcFrom As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDAndOffAccountForDivision(sClauseWhere, bUseUnicode)
        LoadDataSource(tdbcFrom, dt, bUseUnicode)
        LoadDataSource(tdbcTo, dt.DefaultView.ToTable, bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 combo Tài khoản có chứa % (bao gồm Tài khoản trong và ngoại bảng) có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbcFrom"></param>
    ''' <param name="tdbcTo"></param>
    ''' <param name="sClauseWhere"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDAndOffAccountAllForDivision(ByVal tdbcFrom As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDAndOffAccountAllForDivision(sClauseWhere, bUseUnicode)
        LoadDataSource(tdbcFrom, dt, bUseUnicode)
        LoadDataSource(tdbcTo, dt.DefaultView.ToTable, bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load Dropdown Tài khoản (bao gồm Tài khoản trong và ngoại bảng)
    ''' </summary>
    ''' <param name="tdbd"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDAndOffAccountForDivision(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, ReturnTableAccountIDAndOffAccountForDivision("", bUseUnicode), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load Dropdown Tài khoản có chứa % (bao gồm Tài khoản trong và ngoại bảng)
    ''' </summary>
    ''' <param name="tdbd"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDAndOffAccountAllForDivision(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, ReturnTableAccountIDAndOffAccountAllForDivision("", bUseUnicode), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 Dropdown Tài khoản (bao gồm Tài khoản trong và ngoại bảng)
    ''' </summary>
    ''' <param name="tdbdFrom"></param>
    ''' <param name="tdbdTo"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDAndOffAccountForDivision(ByVal tdbdFrom As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdTo As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDAndOffAccountForDivision("", bUseUnicode)
        LoadDataSource(tdbdFrom, dt, bUseUnicode)
        LoadDataSource(tdbdTo, dt.DefaultView.ToTable, bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 Dropdown Tài khoản có chứa % (bao gồm Tài khoản trong và ngoại bảng)
    ''' </summary>
    ''' <param name="tdbdFrom"></param>
    ''' <param name="tdbdTo"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAccountIDAndOffAccountAllForDivision(ByVal tdbdFrom As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdTo As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDAndOffAccountAllForDivision("", bUseUnicode)
        LoadDataSource(tdbdFrom, dt, bUseUnicode)
        LoadDataSource(tdbdTo, dt.DefaultView.ToTable, bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load Dropdown Tài khoản (bao gồm Tài khoản trong và ngoại bảng) có truyền thêm điều kiện Where 
    ''' </summary>
    ''' <param name="tdbd"></param>
    ''' <param name="sClauseWhere"> Điều kiện where (ví dụ: GroupID in ('1', '13'))</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDAndOffAccountForDivision(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, ReturnTableAccountIDAndOffAccountForDivision(sClauseWhere, bUseUnicode), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load Dropdown Tài khoản có chứa % (bao gồm Tài khoản trong và ngoại bảng) có truyền thêm điều kiện Where 
    ''' </summary>
    ''' <param name="tdbd"></param>
    ''' <param name="sClauseWhere"> Điều kiện where (ví dụ: GroupID in ('1', '13'))</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDAndOffAccountAllForDivision(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, ReturnTableAccountIDAndOffAccountAllForDivision(sClauseWhere, bUseUnicode), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 Dropdown Tài khoản (bao gồm Tài khoản trong và ngoại bảng) có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbdFrom"></param>
    ''' <param name="tdbdTo"></param>
    ''' <param name="sClauseWhere"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDAndOffAccountForDivision(ByVal tdbdFrom As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdTo As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDAndOffAccountForDivision(sClauseWhere, bUseUnicode)
        LoadDataSource(tdbdFrom, dt, bUseUnicode)
        LoadDataSource(tdbdTo, dt.DefaultView.ToTable, bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load 2 Dropdown Tài khoản có chứa % (bao gồm Tài khoản trong và ngoại bảng) có truyền thêm điều kiện Where
    ''' </summary>
    ''' <param name="tdbdFrom"></param>
    ''' <param name="tdbdTo"></param>
    ''' <param name="sClauseWhere"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
   Public Sub LoadAccountIDAndOffAccountAllForDivision(ByVal tdbdFrom As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal tdbdTo As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sClauseWhere As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable
        dt = ReturnTableAccountIDAndOffAccountAllForDivision(sClauseWhere, bUseUnicode)
        LoadDataSource(tdbdFrom, dt, bUseUnicode)
        LoadDataSource(tdbdTo, dt.DefaultView.ToTable, bUseUnicode)
    End Sub

#End Region

#Region "Đổ nguồn cho Tài khoản thuộc dạng báo cáo"

    ''#---------------------------------------------------------------------------------------------------
    ''# Title: SQLStoreD91P9303
    ''# Created User: 
    ''# Created Date: 20/10/2015 11:26:08
    ''#---------------------------------------------------------------------------------------------------
    'Private Function SQLStoreD91P9303(ByVal sWhere As String, ByVal IsAll As Byte, ByVal sModuleID As String, ByVal sFormID As String) As String
    '    Dim sSQL As String = ""
    '    sSQL &= ("-- Do nguon cho Tai khoan bao cao" & vbCrLf)
    '    sSQL &= "Exec D91P9303 "
    '    sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
    '    sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
    '    sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
    '    sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[2], NOT NULL
    '    sSQL &= SQLString(sModuleID) & COMMA 'ModuleID, varchar[20], NOT NULL
    '    sSQL &= SQLString(sFormID) & COMMA 'FormID, varchar[20], NOT NULL
    '    sSQL &= SQLString(sWhere) & COMMA 'sWhere, varchar[8000], NOT NULL
    '    sSQL &= SQLNumber(IsAll)  'IsAll, tinyint, NOT NULL
    '    Return sSQL
    'End Function

    ''' <summary>
    ''' Trả ra Table Tài khoản có chứa % và kiểm tra Phân quyền theo Đơn vị
    ''' </summary>
    ''' <param name="sClauseWhere">Nếu có điều kiện lọc thì truyền vào (ví dụ: GroupID in ('1', '13'))</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnTableAccountIDReport(ByVal sClauseWhere As String, ByVal bAll As Boolean) As DataTable
        Return ReturnDataTable(SQLFinance.SQLStoreD91P9303(sClauseWhere, L3Byte(bAll), "", ""))

    End Function

    Public Sub LoadAccountIDReport(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal sClauseWhere As String, ByVal bAll As Boolean)
        LoadDataSource(tdbc, ReturnTableAccountIDReport(sClauseWhere, bAll), gbUnicode)
    End Sub

    Public Sub LoadAccountIDReport(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal sClauseWhere As String)
        LoadAccountIDReport(tdbc, sClauseWhere, False)
    End Sub

    Public Sub LoadAccountIDReport(ByVal tdbc As C1.Win.C1List.C1Combo)
        LoadAccountIDReport(tdbc, "")
    End Sub

    Public Sub LoadAccountIDReport(ByVal tdbcFrom As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo, ByVal sClauseWhere As String, ByVal bAll As Boolean)
        Dim dt As DataTable
        dt = ReturnTableAccountIDReport(sClauseWhere, bAll)
        LoadDataSource(tdbcFrom, dt, gbUnicode)
        LoadDataSource(tdbcTo, dt.DefaultView.ToTable, gbUnicode)

    End Sub

    Public Sub LoadAccountIDReport(ByVal tdbcFrom As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo, ByVal sClauseWhere As String)
        LoadAccountIDReport(tdbcFrom, tdbcTo, sClauseWhere, False)
    End Sub

    Public Sub LoadAccountIDReport(ByVal tdbcFrom As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo)
        LoadAccountIDReport(tdbcFrom, tdbcTo, "")
    End Sub
#End Region

#Region "KeyboardCues"

    Private Declare Function SystemParametersInfoSet Lib "user32.dll" Alias "SystemParametersInfoW" (ByVal action As Integer, ByVal param As Integer, ByVal value As Integer, ByVal winini As Boolean) As Boolean
    Private Const SPI_SETKEYBOARDCUES As Integer = &H100B

    Public Sub KeyboardCues()
        SystemParametersInfoSet(SPI_SETKEYBOARDCUES, 0, 1, False)
    End Sub
#End Region

#Region "Liên quan tới In D99C2003"

    Public Function AllowNewD99C2003(ByRef report As D99C2003, ByVal frmParent As Form) As Boolean
        Dim isOpen As Boolean = False
        report = New D99C2003(frmParent, isOpen)
        If isOpen Then Return False
        Return True
    End Function
#End Region

#Region "Đổ nguồn Mã bất động sản"

    'Public Sub LoadPropertyProduct(ByVal ctrlPropertyProduct As Control, ByVal sProjectValue As String, ByRef dtSource As DataTable, ByVal iIsAll As Integer)
    '    If dtSource Is Nothing Then
    '        dtSource = ReturnTablePropertyProduct()
    '        If dtSource Is Nothing Then Exit Sub
    '    End If

    '    Dim dtPropertyProduct As DataTable = Nothing
    '    If sProjectValue <> "" Then
    '        dtPropertyProduct = ReturnTableFilter(dtSource, "ProjectID = " & SQLString(sProjectValue), True)
    '    Else
    '        dtPropertyProduct = dtSource
    '    End If
    '    If TypeOf (ctrlPropertyProduct) Is C1.Win.C1List.C1Combo Then
    '        LoadDataSource(CType(ctrlPropertyProduct, C1.Win.C1List.C1Combo), dtPropertyProduct, gbUnicode)
    '    ElseIf TypeOf (ctrlPropertyProduct) Is C1.Win.C1TrueDBGrid.C1TrueDBDropdown Then
    '        LoadDataSource(CType(ctrlPropertyProduct, C1.Win.C1TrueDBGrid.C1TrueDBDropdown), dtPropertyProduct, gbUnicode)
    '    End If
    'End Sub

    ''' <summary>
    ''' Hàm đổ nguồn combo/dropdown Mã BĐS
    ''' </summary>
    ''' <param name="ctrlPropertyProduct"></param>
    ''' <param name="sProjectValue"></param>
    ''' <param name="dtSource"></param>
    ''' <param name="iIsAll">Tất cả: 0 hoặc 1</param>
    ''' <param name="sFormID"></param>
    ''' <param name="iIsAddNew">Thêm mới: 0 hoặc 1</param>
    ''' <param name="iIsD89F3000">Đổ nguồn màn hình báo cáo động: 0 hoặc 1</param>
    ''' <remarks></remarks>
    Public Sub LoadPropertyProduct(ByVal ctrlPropertyProduct As Control, ByVal sProjectValue As String, ByRef dtSource As DataTable, ByVal iIsAll As Integer, ByVal sFormID As String, ByVal iIsAddNew As Byte, ByVal iIsD89F3000 As Byte)
        If dtSource Is Nothing Then
            'update 15/7/2015
            dtSource = ReturnTablePropertyProduct(iIsAll, sFormID, iIsAddNew, iIsD89F3000)
            If dtSource Is Nothing Then Exit Sub
        End If

        Dim dtPropertyProduct As DataTable = Nothing
        'update 15/7/2015
        If iIsAll = 1 Then
            If sProjectValue = "%" Then
                dtPropertyProduct = ReturnTableFilter(dtSource, "ProjectID = '%' ", True)
            Else
                dtPropertyProduct = ReturnTableFilter(dtSource, "ProjectID = '%' Or ProjectID = " & SQLString(sProjectValue), True)
            End If
        Else
            If sProjectValue <> "" Then
                dtPropertyProduct = ReturnTableFilter(dtSource, "ProjectID = " & SQLString(sProjectValue), True)
            Else
                dtPropertyProduct = dtSource
            End If
        End If

        If TypeOf (ctrlPropertyProduct) Is C1.Win.C1List.C1Combo Then
            LoadDataSource(CType(ctrlPropertyProduct, C1.Win.C1List.C1Combo), dtPropertyProduct, gbUnicode)
        ElseIf TypeOf (ctrlPropertyProduct) Is C1.Win.C1TrueDBGrid.C1TrueDBDropdown Then
            LoadDataSource(CType(ctrlPropertyProduct, C1.Win.C1TrueDBGrid.C1TrueDBDropdown), dtPropertyProduct, gbUnicode)
        End If
    End Sub

    ''' <summary>
    ''' Hàm đổ nguồn combo/dropdown Mã BĐS
    ''' </summary>
    ''' <param name="ctrlPropertyProduct"></param>
    ''' <param name="sProjectValue"></param>
    ''' <param name="dtSource"></param>
    ''' <param name="iIsAll">Tất cả: 0 hoặc 1</param>
    ''' <remarks></remarks>
    Public Sub LoadPropertyProduct(ByVal ctrlPropertyProduct As Control, ByVal sProjectValue As String, ByRef dtSource As DataTable, ByVal iIsAll As Integer)
        LoadPropertyProduct(ctrlPropertyProduct, sProjectValue, dtSource, iIsAll, "", 0, 0)
    End Sub

    ''' <summary>
    ''' Hàm đổ nguồn combo/dropdown Mã BĐS 
    ''' </summary>
    ''' <param name="ctrlPropertyProduct"></param>
    ''' <param name="sProjectValue">Giá trị lọc theo Dự án</param>
    ''' <param name="dtSource"></param>
    ''' <remarks></remarks>
    Public Sub LoadPropertyProduct(ByVal ctrlPropertyProduct As Control, ByVal sProjectValue As String, ByRef dtSource As DataTable)
        '        If dtSource Is Nothing Then
        '            dtSource = ReturnTablePropertyProduct()
        '            If dtSource Is Nothing Then Exit Sub
        '        End If
        '
        '        Dim dtPropertyProduct As DataTable = Nothing
        '        If sProjectValue <> "" Then
        '            dtPropertyProduct = ReturnTableFilter(dtSource, "ProjectID = " & SQLString(sProjectValue), True)
        '        Else
        '            dtPropertyProduct = dtSource
        '        End If
        '        If TypeOf (ctrlPropertyProduct) Is C1.Win.C1List.C1Combo Then
        '            LoadDataSource(CType(ctrlPropertyProduct, C1.Win.C1List.C1Combo), dtPropertyProduct, gbUnicode)
        '        ElseIf TypeOf (ctrlPropertyProduct) Is C1.Win.C1TrueDBGrid.C1TrueDBDropdown Then
        '            LoadDataSource(CType(ctrlPropertyProduct, C1.Win.C1TrueDBGrid.C1TrueDBDropdown), dtPropertyProduct, gbUnicode)
        '        End If
        LoadPropertyProduct(ctrlPropertyProduct, sProjectValue, dtSource, 0)

    End Sub

    Public Sub LoadPropertyProduct(ByVal ctrlPropertyProduct As Control, ByVal sProjectValue As String)
        LoadPropertyProduct(ctrlPropertyProduct, sProjectValue, Nothing)
    End Sub

    ''' <summary>
    ''' Hàm trả về datatable nguồn Mã BĐS 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnTablePropertyProduct() As DataTable
        Return ReturnTablePropertyProduct(0)
    End Function

    ''' <summary>
    ''' Hàm trả về datatable nguồn Mã BĐS
    ''' </summary>
    ''' <param name="iIsAll"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnTablePropertyProduct(ByVal iIsAll As Integer) As DataTable
        Return ReturnTablePropertyProduct(iIsAll, "", 0, 0) 'ReturnDataTable(SQLStoreD27P9000(iIsAll, "", 0, 0))
    End Function

    ''' <summary>
    ''' Hàm trả về datatable nguồn Mã BĐS
    ''' </summary>
    ''' <param name="iIsAll">Tất cả: 0 hoặc 1</param>
    ''' <param name="sFormID"></param>
    ''' <param name="iIsAddNew">Thêm mới: 0 hoặc 1</param>
    ''' <param name="iIsD89F3000">Đổ nguồn màn hình báo cáo động: 0 hoặc 1</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnTablePropertyProduct(ByVal iIsAll As Integer, ByVal sFormID As String, ByVal iIsAddNew As Byte, ByVal iIsD89F3000 As Byte) As DataTable
        Dim sSQL As String = ""
        sSQL = "SELECT TOP 1 1 FROM DBO.SYSOBJECTS WITH(NOLOCK) " & vbCrLf
        sSQL &= "WHERE ID = OBJECT_ID(N'[DBO].[D27P9000]') " & vbCrLf
        sSQL &= "AND OBJECTPROPERTY(ID, N'IsProcedure') = 1 "
        If Not ExistRecord(sSQL) Then Return Nothing
        Return ReturnDataTable(SQLStoreD27P9000(iIsAll, sFormID, iIsAddNew, iIsD89F3000))
    End Function

    'Public Function ReturnTablePropertyProduct(ByVal iIsAll As Integer) As DataTable
    '    Dim sSQL As String = ""
    '    sSQL = "SELECT TOP 1 1 FROM DBO.SYSOBJECTS WITH(NOLOCK) " & vbCrLf
    '    sSQL &= "WHERE ID = OBJECT_ID(N'[DBO].[D27P9000]') " & vbCrLf
    '    sSQL &= "AND OBJECTPROPERTY(ID, N'IsProcedure') = 1 "
    '    If Not ExistRecord(sSQL) Then Return Nothing
    '    Return ReturnDataTable(SQLStoreD27P9000(iIsAll))
    'End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD27P9000
    '# Created User: NGOCTHOAI
    '# Created Date: 15/09/2014 03:16:53
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD27P9000(ByVal iIsAll As Integer, ByVal sFormID As String, ByVal iIsAddNew As Byte, ByVal iIsD89F3000 As Byte) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon combo/dropdown BĐS " & vbCrLf)
        sSQL &= "Exec D27P9000 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLNumber(iIsAll) & COMMA 'IsAll, tinyint, NOT NULL
        sSQL &= SQLNumber(0) & COMMA 'Mode, tinyint, NOT NULL
        sSQL &= SQLString(sFormID) & COMMA 'FormID, varchar[50], NOT NULL
        'update 15/7/2015
        sSQL &= SQLNumber(iIsAddNew) & COMMA 'IsAddNew, tinyint, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(iIsD89F3000) 'IsD89F3000, tinyint, NOT NULL
        Return sSQL
    End Function



#End Region

#Region "Set thuộc tính Anchor cho form có độ rộng 1024, áp dụng cho Desktop"

    Public Enum EnumAnchorStyles
        TopRight = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        TopLeft = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left)
        BottomLeft = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        BottomRight = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)

        TopRightBottom = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right Or System.Windows.Forms.AnchorStyles.Bottom)
        TopLeftBottom = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Bottom)
        TopLeftRight = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right)
        BottomLeftRight = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right)

        TopLeftRightBottom = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right Or System.Windows.Forms.AnchorStyles.Bottom)
    End Enum

    Private Function CheckPrimaryScreen() As Boolean
        'Nếu màn hình có độ phân giải đúng chuẩn 1024x680 thì trả về True

        Dim x_original As Integer = 1024
        Dim y_original As Integer = 768

        Dim workingRectangle As System.Drawing.Rectangle = Screen.PrimaryScreen.Bounds 'Screen.PrimaryScreen.WorkingArea
        Dim tile_x As Double
        Dim tile_y As Double
        Dim y_Location As Integer = 35

        tile_x = workingRectangle.Width / x_original
        tile_y = workingRectangle.Height / y_original

        If tile_x <= 1 And tile_y <= 1 Then
            Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Khi chỉnh Anchor cho tất cả các control
    ''' </summary>
    ''' <param name="EAnchor">Thuộc tính cần chỉnh Anchor</param>
    ''' <param name="AlwaysAnchor">True: luôn luôn chỉnh Anchor, không kiểm tra Call Desktop</param>
    ''' <param name="obj">Tập các control được chỉnh Anchor</param>
    ''' <remarks></remarks>
    Public Sub AnchorForControl(ByVal EAnchor As EnumAnchorStyles, ByVal AlwaysAnchor As Boolean, ByVal ParamArray objControlList() As System.Windows.Forms.Control)

        ' Khi chỉnh Anchor cho tất cả các control
        ' Đối với lưới thì cột cuối giãn ra, nếu muốn chỉnh ra đều thì gọi hàm Anchor_ResizeColumnsGird
        If Not AlwaysAnchor AndAlso Not gbCallDesktop Then Exit Sub

        If objControlList.Length < 1 Then Exit Sub

        'Update 22/06/2015: Nếu màn hình có độ phân giải đúng chuẩn 1024x680 thì không set Anchor
        If CheckPrimaryScreen() Then Exit Sub

        Try
            Dim _anchor As System.Windows.Forms.AnchorStyles = CType(EAnchor, System.Windows.Forms.AnchorStyles)
            For Each ctrl As System.Windows.Forms.Control In objControlList
                ctrl.Anchor = _anchor
            Next
        Catch ex As Exception

        End Try
    End Sub

    Public Sub AnchorForControl(ByVal EAnchor As EnumAnchorStyles, ByVal ParamArray objControlList() As System.Windows.Forms.Control)
        AnchorForControl(EAnchor, False, objControlList)
    End Sub

    Private _dicWith As Dictionary(Of String, Integer)

    Public Sub AnchorResizeColumnsGrid(ByVal EAnchor As EnumAnchorStyles, ByVal AlwaysAnchor As Boolean, ByVal ParamArray objControlList() As System.Windows.Forms.Control)
        ' Dùng cho trường hợp chỉnh Anchor giãn đều các cột trên lưới.
        If Not AlwaysAnchor AndAlso Not gbCallDesktop Then Exit Sub

        'Update 22/06/2015: Nếu màn hình có độ phân giải đúng chuẩn 1024x680 thì không set Anchor
        If CheckPrimaryScreen() Then Exit Sub


        If objControlList.Length < 1 Then Exit Sub

        Try
            '_dicWith = New Dictionary(Of String, Integer)
            Dim _anchor As System.Windows.Forms.AnchorStyles = CType(EAnchor, System.Windows.Forms.AnchorStyles)
            For Each ctrl As System.Windows.Forms.Control In objControlList
                ctrl.Anchor = _anchor
                If TypeOf (ctrl) Is C1.Win.C1TrueDBGrid.C1TrueDBGrid Then
                    Dim sKeyName As String
                    sKeyName = GetFormNameByControl(ctrl)
                    If Not sKeyName.Equals("") Then
                        sKeyName &= "_" & ctrl.Name
                    Else
                        Exit Sub
                    End If

                    If _dicWith Is Nothing Then _dicWith = New Dictionary(Of String, Integer)
                    If Not _dicWith.ContainsKey(sKeyName) Then
                        _dicWith.Add(sKeyName, ctrl.Width)
                    Else
                        _dicWith.Item(sKeyName) = ctrl.Width
                    End If

                    AddHandler ctrl.Resize, AddressOf Grid_Resize
                End If
            Next

        Catch ex As Exception
        End Try

    End Sub

    Public Sub AnchorResizeColumnsGrid(ByVal EAnchor As EnumAnchorStyles, ByVal ParamArray objControlList() As System.Windows.Forms.Control)
        AnchorResizeColumnsGrid(EAnchor, False, objControlList)
    End Sub

    Private Sub Grid_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles tdbg.Resize
        Dim iWWidthOld As Integer

        Dim sKeyName As String

        Try

            Dim c1Gird As C1.Win.C1TrueDBGrid.C1TrueDBGrid

            c1Gird = CType(sender, C1.Win.C1TrueDBGrid.C1TrueDBGrid)

            sKeyName = GetFormNameByControl(c1Gird)

            If Not sKeyName.Equals("") Then

                sKeyName &= "_" & c1Gird.Name

            Else

                Exit Sub

            End If

            If _dicWith.ContainsKey(sKeyName) Then

                iWWidthOld = _dicWith.Item(sKeyName)

            Else

                Exit Sub

            End If

            If iWWidthOld = c1Gird.Width Then

                Exit Sub

            End If

            Dim iNowtdbgWith As Integer

            Dim iSumColGird As Integer = 0


            For i As Integer = 0 To c1Gird.Splits.Count - 1

                For j As Integer = 0 To c1Gird.Columns.Count - 1

                    If c1Gird.Splits(i).DisplayColumns(j).Visible Then

                        iSumColGird += (c1Gird.Splits(i).DisplayColumns(j).Width)

                    End If

                Next

            Next

            iNowtdbgWith = c1Gird.Width

            If iNowtdbgWith <= 0 Then Exit Sub

            Dim iNowWithForGird As Integer

            iNowWithForGird = iNowtdbgWith

            iNowWithForGird -= c1Gird.RecordSelectorWidth

            Dim vs As C1.Win.C1TrueDBGrid.Util.VBar

            vs = c1Gird.VScrollBar

            If vs.Visible AndAlso iNowWithForGird < iSumColGird Then

                iNowWithForGird -= vs.Width

            End If

            If iSumColGird >= iNowWithForGird AndAlso iNowtdbgWith > iWWidthOld Then Exit Sub

            Dim dRate As Double = (iNowWithForGird) / iWWidthOld

            For i As Integer = 0 To c1Gird.Splits.Count - 1

                For j As Integer = 0 To c1Gird.Columns.Count - 1

                    'If c1Gird.Splits(i).DisplayColumns(j).Visible Then

                    c1Gird.Splits(i).DisplayColumns(j).Width = Math.Ceiling(dRate * c1Gird.Splits(i).DisplayColumns(j).Width)

                    'End If

                Next

            Next

            _dicWith.Item(sKeyName) = iNowtdbgWith

        Catch ex As Exception


        End Try
    End Sub

    Public Function GetFormNameByControl(ByVal ctrl As System.Windows.Forms.Control) As String
        Dim frm As Form = GetFormByControl(ctrl)
        If frm Is Nothing Then Return ""
        Return frm.Name
    End Function

    Public Function GetFormByControl(ByVal ctrl As System.Windows.Forms.Control) As Form
        If ctrl Is Nothing Then Return Nothing
        If TypeOf (ctrl) Is Form Then Return CType(ctrl, Form)

        Dim father As Control = ctrl.Parent
        If TypeOf (father) Is Form Then
            Return CType(father, Form)
        Else
            Return GetFormByControl(father)
        End If
    End Function

#End Region

#Region "Hiển thị menu theo phân quyền"
    'Ẩn phân nhóm
    Public Sub SetVisibleDelimiter(ByVal menu As C1.Win.C1Command.C1MainMenu)
        For i As Integer = 0 To menu.CommandLinks.Count - 1
            If TypeOf (menu.CommandLinks(i).Command) Is C1.Win.C1Command.C1CommandMenu Then 'Lấy 5 menu chính
                Dim mnu As C1.Win.C1Command.C1CommandMenu = CType(menu.CommandLinks(i).Command, C1.Win.C1Command.C1CommandMenu)
                SetVisibleDelimiter(mnu) 'Set menu con trong từng menu chính
            End If
        Next
    End Sub

    Private Sub SetVisibleDelimiter(ByVal mnu As C1.Win.C1Command.C1CommandMenu)
        For i As Integer = 0 To mnu.CommandLinks.Count - 1
            If (i > 0 And mnu.CommandLinks(i).Delimiter = True) AndAlso FindVisibleMenu(mnu, i) = False Then
                mnu.CommandLinks(i).Delimiter = False
            End If

            If TypeOf (mnu.CommandLinks(i).Command) Is C1.Win.C1Command.C1CommandMenu Then 'menu chứa menu con
                Dim cmnu As C1.Win.C1Command.C1CommandMenu = CType(mnu.CommandLinks(i).Command, C1.Win.C1Command.C1CommandMenu)
                SetVisibleChildDelimiter(cmnu)
            End If
        Next
    End Sub

    Private Sub SetVisibleChildDelimiter(ByVal mnu As C1.Win.C1Command.C1CommandMenu)
        If mnu.CommandLinks.Count < 1 Then Exit Sub

        Dim iVisible As Integer = 0

        For i As Integer = 0 To mnu.CommandLinks.Count - 1
            If (i > 0 And mnu.CommandLinks(i).Delimiter) AndAlso FindVisibleMenu(mnu, i) = False Then
                mnu.CommandLinks(i).Delimiter = False
            End If

            If TypeOf (mnu.CommandLinks(i).Command) Is C1.Win.C1Command.C1CommandMenu Then 'menu chứa menu con
                Dim cmnu As C1.Win.C1Command.C1CommandMenu = CType(mnu.CommandLinks(i).Command, C1.Win.C1Command.C1CommandMenu)
                SetVisibleChildDelimiter(cmnu)
            End If

            'Ẩn menu cha khi các menu con bị ẩn
            If mnu.CommandLinks(i).Command.Visible = True Then iVisible += 1
        Next

        If mnu.Visible Then mnu.Visible = (iVisible > 0)
    End Sub

    Private Function FindVisibleMenu(ByVal mnu As C1.Win.C1Command.C1CommandMenu, ByVal iCurPosition As Integer) As Boolean
        If iCurPosition = 0 Then Return False
        For i As Integer = iCurPosition - 1 To 0 Step -1
            If TypeOf (mnu.CommandLinks(i).Command) Is C1.Win.C1Command.C1CommandMenu Then
                Dim cmnu As C1.Win.C1Command.C1CommandMenu = CType(mnu.CommandLinks(i).Command, C1.Win.C1Command.C1CommandMenu)
                If cmnu.Visible = True Then Return True
            Else
                If mnu.CommandLinks(i).Command.Visible = True Then Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' Phân quyền menu
    ''' </summary>
    ''' <param name="mnu"></param>
    ''' <param name="Visible">Điều kiện để hiển thị menu: 1: ReturnPermission(FormPermission) >= EnumPermission.View hoặc Nghiệp vụ: ReturnPermission(FormPermission) > EnumPermission.View </param>
    ''' <remarks></remarks>
    Public Sub VisibledMenu(ByVal mnu As C1.Win.C1Command.C1Command, ByVal Visible As Boolean)
        If mnu.Visible Then mnu.Visible = Visible
    End Sub

    ''' <summary>
    ''' Phân quyền menu (ngoại trừ menu Nghiệp vụ)
    ''' </summary>
    ''' <param name="mnu"></param>
    ''' <param name="FormPermission">form dùng để phần quyền</param>
    ''' <remarks>ReturnPermission(FormPermission) >= EnumPermission.View</remarks>
    Public Sub VisibledMenu(ByVal mnu As C1.Win.C1Command.C1Command, ByVal FormPermission As String)
        If mnu.Visible Then mnu.Visible = ReturnPermission(FormPermission) >= EnumPermission.View
    End Sub

    ''' <summary>
    '''  Phân quyền menu Nghiệp vụ
    ''' </summary>
    ''' <param name="mnu"></param>
    ''' <param name="FormPermission">form dùng để phần quyền</param>
    ''' <remarks>ReturnPermission(FormPermission) > EnumPermission.View</remarks>
    Public Sub VisibledMenuTrans(ByVal mnu As C1.Win.C1Command.C1Command, ByVal FormPermission As String)
        If mnu.Visible Then mnu.Visible = ReturnPermission(FormPermission) > EnumPermission.View
    End Sub

#End Region

#Region "Các hàm cho SplitContainer"

    ''' <summary>
    ''' Vẽ lại thanh ngang cho ResetSplitContainer
    ''' </summary>
    ''' <param name="ctrl"></param>
    ''' <remarks></remarks>
    Public Sub ResetSplitContainer(ByVal ctrl As SplitContainer)
        ResetSplitContainer(ctrl, Color.DarkGray, 1)
    End Sub

    ''' <summary>
    ''' Vẽ lại thanh ngang cho ResetSplitContainer
    ''' </summary>
    ''' <param name="ctrl"></param>
    ''' <param name="color">Màu cho thanh ngang (default: Color.DarkGray)</param>
    ''' <param name="iWidth">Độ rộng cho thanh ngang (default: 1)</param>
    ''' <remarks></remarks>
    Public Sub ResetSplitContainer(ByVal ctrl As SplitContainer, ByVal color As Color, ByVal iWidth As Integer)
        Dim penSplitContainer As Pen = New Pen(color, iWidth)
        ctrl.Tag = penSplitContainer
        AddHandler ctrl.Paint, AddressOf SplitContainer1_Paint
    End Sub

    <DebuggerStepThrough()> _
    Private Sub SplitContainer1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)
        Dim splitter As SplitContainer = TryCast(sender, SplitContainer)
        If splitter Is Nothing Then Return
        Dim penSplitContainer As Pen = New Pen(Color.DarkGray, 1)
        If splitter.Tag IsNot Nothing AndAlso TypeOf (splitter.Tag) Is Pen Then
            penSplitContainer = CType(splitter.Tag, Pen)
        End If
        If splitter.Orientation = Orientation.Horizontal Then
            e.Graphics.DrawLine(penSplitContainer, 0, L3Int(splitter.SplitterDistance + (splitter.SplitterWidth / 2)), splitter.Width, L3Int(splitter.SplitterDistance + (splitter.SplitterWidth / 2)))
            ' Vẽ 1 đoạn ngắn
            '  e.Graphics.DrawLine(Pens.DarkGray, L3Int(splitter.Width / 2) - 100, L3Int(splitter.SplitterDistance + (splitter.SplitterWidth / 2)), L3Int(splitter.Width / 2) + 100, L3Int(splitter.SplitterDistance + (splitter.SplitterWidth / 2)))
        Else
            e.Graphics.DrawLine(penSplitContainer, L3Int(splitter.SplitterDistance + (splitter.SplitterWidth / 2)), 0, L3Int(splitter.SplitterDistance + (splitter.SplitterWidth / 2)), splitter.Height)
        End If
    End Sub

#End Region

#Region "Sự kiện của TabControl"

    ''' <summary>
    ''' Đánh số thứ tự cho text của từng tabpage ("1. Text....")
    ''' </summary>
    ''' <param name="tab"></param>
    ''' <remarks></remarks>
    Public Sub ResetIndexTabControl(ByVal tab As TabControl)
        Try
            For i As Integer = 0 To tab.TabPages.Count - 1
                Dim sText As String = tab.TabPages(i).Text
                If sText.Contains(".") AndAlso IsNumeric(L3Left(sText, sText.IndexOf("."c))) Then
                    'Continue For
                    tab.TabPages(i).Text = sText.Replace(L3Left(sText, sText.IndexOf("."c)) & ". ", (i + 1) & ". ")
                Else
                    tab.TabPages(i).Text = (i + 1) & ". " & sText
                End If
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private dicTab As Dictionary(Of String, TabControl)
    ''' <summary>
    ''' Phím nóng Alt + 1,2,3... cho tab (chi hỗ trợ 1 tab cho 1 form)
    ''' </summary>
    ''' <param name="tab"></param>
    ''' <remarks></remarks>
    Public Sub HotkeyAltTabControl(ByVal tab As TabControl)
        Dim frm As Form = GetFormByControl(tab)
        ResetIndexTabControl(tab)
        If dicTab Is Nothing Then dicTab = New Dictionary(Of String, TabControl)

        If Not dicTab.ContainsKey(frm.Name) Then dicTab.Add(frm.Name, tab)
        AddHandler frm.KeyDown, AddressOf frm_KeyDown
        AddHandler frm.FormClosed, AddressOf frm_FormClosed
    End Sub

    <DebuggerStepThrough()> _
    Private Sub frm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.Alt And ((e.KeyValue >= 97 And e.KeyValue <= 105) Or (e.KeyValue >= 49 And e.KeyValue <= 57)) Then
            Dim index As Integer = e.KeyValue - 97
            If index < 0 Then index = e.KeyValue - 49
            Dim frm As Form = CType(sender, Form)
            Dim tab As TabControl = dicTab(frm.Name)
            If tab.Enabled AndAlso index < tab.TabPages.Count AndAlso tab.TabPages(index).Enabled Then
                Application.DoEvents()
                tab.SelectedIndex = index
                tab.Focus()
                Application.DoEvents()
            End If
        End If
    End Sub

    <DebuggerStepThrough()> _
    Private Sub frm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs)
        Try
            Dim frm As Form = CType(sender, Form)
            If dicTab.ContainsKey(frm.Name) Then dicTab.Remove(frm.Name)
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "Filter combo"

    ' Biến nhận dạng có nhập liệu bằng tày hay gán bằng code, mục đích là khi enter (ko thay đồi giá trị) thì qua control kế tiếp mà ko xổ D99F5555
    Private bTextChanged As Boolean = False
    Private dicComboObjectTypeID As New Dictionary(Of String, C1.Win.C1List.C1Combo)

    ''' <summary>
    ''' Khai báo sử dụng tính năng FilterCombo (Mặc định check D91T0025)
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <remarks></remarks>
    Public Sub UseFilterCombo(ByVal ParamArray tdbc() As C1.Win.C1List.C1Combo)
        UseFilterCombo(True, tdbc)
    End Sub

    ''' <summary>
    ''' Khai báo sử dụng tính năng FilterCombo (Mặc định check D91T0025)
    ''' </summary>
    ''' <param name="bCheckD91">Kiểm tra trường LoadFormNotINV của bảng D91T0025</param>
    ''' <param name="tdbc"></param>
    ''' <remarks></remarks>
    Public Sub UseFilterCombo(ByVal bCheckD91 As Boolean, ByVal ParamArray tdbc() As C1.Win.C1List.C1Combo)
        UseFilterCombo(bCheckD91, False, "", tdbc)
    End Sub

    ''' <summary>
    ''' Khai báo sử dụng tính năng FilterCombo cho combo có check chọn nhiều dòng
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <remarks></remarks>
    Public Sub UseFilterCheckCombo(ByVal ParamArray tdbc() As C1.Win.C1List.C1Combo)
        UseFilterCombo(False, True, ";", tdbc)
    End Sub

    ''' <summary>
    ''' Khai báo sử dụng tính năng FilterCombo cho combo có check chọn nhiều dòng
    ''' </summary>
    ''' <param name="sSeparator">Chuỗi phân cách giữa các giá trị (mặc định là ";")</param>
    ''' <param name="tdbc"></param>
    ''' <remarks></remarks>
    Public Sub UseFilterCheckCombo(ByVal sSeparator As String, ByVal ParamArray tdbc() As C1.Win.C1List.C1Combo)
        UseFilterCombo(False, True, sSeparator, tdbc)
    End Sub

    Private Sub UseFilterCombo(ByVal bCheckD91 As Boolean, ByVal bVisibleChoose As Boolean, ByVal sSeparator As String, ByVal ParamArray tdbc() As C1.Win.C1List.C1Combo)
        If bCheckD91 Then ' Kiểm tra bảng D91T0025
            If DxxFormat.LoadFormNotINV = 0 Then Exit Sub
        End If

        For i As Integer = 0 To tdbc.Length - 1
            If bVisibleChoose Then
                tdbc(i).Columns(0).Tag = sSeparator
            Else
                tdbc(i).Columns(0).Tag = bCheckD91
            End If
            tdbc(i).EditorForeColor = Color.Blue
            tdbc(i).AutoCompletion = False
            tdbc(i).AutoDropDown = False

            If bVisibleChoose Then
                AddHandler tdbc(i).SelectedValueChanged, AddressOf tdbc_SelectedValueChanged_check
                AddHandler tdbc(i).TextChanged, AddressOf tdbc_TextChanged_check
                AddHandler tdbc(i).BeforeOpen, AddressOf tdbc_BeforeOpen_Check
            Else
                AddHandler tdbc(i).SelectedValueChanged, AddressOf tdbc_SelectedValueChanged
                AddHandler tdbc(i).TextChanged, AddressOf tdbc_TextChanged
                AddHandler tdbc(i).BeforeOpen, AddressOf tdbc_BeforeOpen
            End If
            AddHandler tdbc(i).KeyDown, AddressOf tdbc_KeyDown
        Next
    End Sub

    Public Sub UseFilterComboObjectID(ByVal tdbcObjectID As C1.Win.C1List.C1Combo, ByVal tdbcObjectTypeID As C1.Win.C1List.C1Combo, ByVal bCheckD91 As Boolean)
        If bCheckD91 Then ' Kiểm tra bảng D91T0025
            If DxxFormat.LoadFormNotINV = 0 Then Exit Sub
        End If

        If dicComboObjectTypeID.ContainsKey(tdbcObjectID.FindForm.Name & "_" & tdbcObjectID.Name) Then
            dicComboObjectTypeID.Remove(tdbcObjectID.FindForm.Name & "_" & tdbcObjectID.Name)
            dicComboObjectTypeID.Add(tdbcObjectID.FindForm.Name & "_" & tdbcObjectID.Name, tdbcObjectTypeID)
        Else
            dicComboObjectTypeID.Add(tdbcObjectID.FindForm.Name & "_" & tdbcObjectID.Name, tdbcObjectTypeID)
        End If

        tdbcObjectID.Columns(0).Tag = bCheckD91
        tdbcObjectID.EditorForeColor = Color.Blue
        tdbcObjectID.AutoCompletion = False
        tdbcObjectID.AutoDropDown = False
        AddHandler tdbcObjectID.SelectedValueChanged, AddressOf tdbc_SelectedValueChanged
        AddHandler tdbcObjectID.TextChanged, AddressOf tdbc_TextChanged
        AddHandler tdbcObjectID.BeforeOpen, AddressOf tdbc_BeforeOpen
        AddHandler tdbcObjectID.KeyDown, AddressOf tdbc_KeyDown
    End Sub

    Public Sub UseFilterComboObjectID(ByVal tdbcObjectID As C1.Win.C1List.C1Combo, ByVal tdbcObjectTypeID As C1.Win.C1List.C1Combo)
        UseFilterComboObjectID(tdbcObjectID, tdbcObjectTypeID, True)
    End Sub

    <DebuggerStepThrough()> _
    Private Sub tdbc_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        bTextChanged = False
    End Sub

    <DebuggerStepThrough()> _
Private Sub tdbc_SelectedValueChanged_check(ByVal sender As System.Object, ByVal e As System.EventArgs)
        bTextChanged = True
        'Dùng cho combo check chọn
        Dim tdbc As C1.Win.C1List.C1Combo = CType(sender, C1.Win.C1List.C1Combo)
        If tdbc.ValueMember <> tdbc.DisplayMember Then
            tdbc.Tag = ""
        End If
    End Sub

    <DebuggerStepThrough()> _
    Private Sub tdbc_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        bTextChanged = True
    End Sub

    <DebuggerStepThrough()> _
Private Sub tdbc_TextChanged_check(ByVal sender As Object, ByVal e As System.EventArgs)
        bTextChanged = True
        Dim tdbc As C1.Win.C1List.C1Combo = CType(sender, C1.Win.C1List.C1Combo)
        If tdbc.ValueMember <> tdbc.DisplayMember Then
            tdbc.Tag = ""
        End If
    End Sub

    <DebuggerStepThrough()> _
    Private Sub tdbc_BeforeOpen(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim bCheckD91 As Boolean = True
        Dim tdbc As C1.Win.C1List.C1Combo = CType(sender, C1.Win.C1List.C1Combo)
        If Not L3Bool(tdbc.Columns(0).Tag) Then bCheckD91 = False
        e.Cancel = True
        FilterCombo(tdbc, e, bCheckD91)
    End Sub

    <DebuggerStepThrough()> _
    Private Sub tdbc_BeforeOpen_Check(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        e.Cancel = True
        FilterCheckCombo(CType(sender, C1.Win.C1List.C1Combo), e)
    End Sub

    <DebuggerStepThrough()> _
    Private Sub tdbc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.F2 Then
            CType(sender, C1.Win.C1List.C1Combo).OpenCombo()
        End If
    End Sub

    ''' <summary>
    ''' Tìm kiếm theo tên cho combo
    ''' </summary>
    ''' <param name="tdbc">Combo cần filter</param>
    ''' <param name="e"></param>
    ''' <param name="bCheckD91">Kiểm tra trường LoadFormNotINV của bảng D91T0025</param>
    ''' <remarks></remarks>
    Public Sub FilterCombo(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal e As System.EventArgs, Optional ByVal bCheckD91 As Boolean = True)
        Try
            ' Kiểm tra bảng D91T0025
            If bCheckD91 Then
                If DxxFormat.LoadFormNotINV = 0 Then Exit Sub 'Return False Then
            End If
            If tdbc.ReadOnly Then Exit Sub ' Return False
            If tdbc.Columns.Count < 2 Then D99C0008.MsgL3("FilterCombo: Invalid columns") : Exit Sub

            ' Kiểm tra gọi từ sự kiện nào?
            Dim bF4 As Boolean = e.GetType().Name = "CancelEventArgs" ' Or e.GetType().Name = "KeyEventArgs"

            Dim listID As New List(Of String)
            Dim listName As New List(Of String)
            GetListField(tdbc, listID, listName)
            Dim sValueMember As String = tdbc.ValueMember
            Dim sValueName As String = ""
            If listName.Count > 0 Then sValueName = listName(0)

            Dim bOld_bTextChanged As Boolean = bTextChanged ' Rightclick -> paste -> click button ->escape =>Lỗi không set '' combo
            If Not bF4 And Not bTextChanged Then
                If tdbc.ValueMember = tdbc.DisplayMember Then tdbc.Text = tdbc.Text.ToUpper ' Khi nhập xong -> CtrL mà có trong ds thi viết hoa Mã lên
                bTextChanged = False : Exit Sub ' Return False
            End If
            bTextChanged = False

            If Not bF4 And (tdbc.Text = "" Or tdbc.Text = "+") Then Exit Sub ' Return False

            Dim dtCombo As DataTable = CType(tdbc.DataSource, DataTable)
            If dtCombo Is Nothing Then D99C0008.MsgL3("FilterCombo: Invalid datasource.") : Exit Sub

            ' Khi thay đổi source tu set lại combo.text = "" nên phải set lại giá trị cũ
            Dim sOldID As String = ""
            For i As Integer = 0 To listName.Count - 1
                If Not dtCombo.Columns.Contains(listName(i) & "_Unsign") Then
                    If sOldID = "" Then sOldID = tdbc.Text ' chỉ lấy 1 lần đầu tiên
                    ConvertDataToUnsign(dtCombo, listName(i), listName(i) & "_Unsign")
                End If
            Next
            If sOldID <> "" Then tdbc.Text = sOldID
            '===============================================

            Dim sText As String = tdbc.Text
            If sText <> "%" Then sText = SQLSpecialCharactor(sText)
            ' Không hiển thị dòng %, + 
            Dim sColumnFind As String = ""
            '            For i As Integer = 0 To listID.Count - 1
            '                If dtCombo.Columns(listID(i)).DataType.Name <> "String" Then Continue For
            '
            '                If sColumnFind <> "" Then sColumnFind &= " And "
            '                sColumnFind &= listID(i) & " <> " & NewCode & " AND " & listID(i) & " <> " & AllCode
            '            Next
            '            sColumnFind = "(" & sColumnFind & ")"
            If dtCombo.Columns(sValueMember).DataType.Name = "String" Then sColumnFind = sValueMember & " <> '' "

            If Not bF4 And sText <> "+" Then
                Dim sSubF As String = ""
                For i As Integer = 0 To listID.Count - 1
                    If dtCombo.Columns(listID(i)).DataType.Name <> "String" Then Continue For
                    If sSubF <> "" Then sSubF &= " OR "
                    sSubF &= listID(i) & " like '%" & sText & "%'"
                Next
                ' Nếu tìm kiếm dạng mới thì luôn tìm theo tên
                '  If DxxFormat.IsFillterName = 1 Or tdbc.ValueMember <> tdbc.DisplayMember Then
                For i As Integer = 0 To listName.Count - 1
                    sSubF &= " Or " & listName(i) & " like '%" & sText & "%'"
                    sSubF &= " Or " & listName(i) & "_Unsign" & " like '%" & sText & "%'"
                Next
                'End If
                If sSubF <> "" Then sColumnFind &= " AND (" & sSubF & ")"
            End If

            dtCombo = ReturnTableFilter(dtCombo, sColumnFind, True)
            Dim iCountRow As Integer = dtCombo.Rows.Count

            If Not bF4 And iCountRow > 0 Then ' Hiện thị ưu tiên dữ liệu (bằng mã -> bằng tên -> bắt đầu mã -> bắt đầu tên -> có chứa)
                dtCombo.Columns.Add(sValueMember & "_SortID")
                Dim sWhere As String = ""
                sWhere &= sValueMember & " like " & SQLString(sText & "%")
                '  Nếu tìm kiếm dạng mới thì luôn tìm theo tên
                If tdbc.ValueMember <> tdbc.DisplayMember And sValueName <> "" Then
                    sWhere &= " OR " & sValueName & " like " & SQLString(sText & "%")
                    sWhere &= " OR " & sValueName & "_Unsign like " & SQLString(sText & "%")
                End If

                Dim dr() As DataRow = dtCombo.Select(sWhere)
                For i As Integer = 0 To dr.Length - 1
                    If dr(i).Item(sValueMember).ToString.ToLower = sText.ToLower Then ' Bẳng với Mã
                        dr(i).Item(sValueMember & "_SortID") = 4
                    ElseIf sValueName <> "" AndAlso (dr(i).Item(sValueName).ToString.ToLower = sText.ToLower Or dr(i).Item(sValueName & "_Unsign").ToString.ToLower = sText.ToLower) Then ' Bằng với Tên
                        dr(i).Item(sValueMember & "_SortID") = 3
                    ElseIf dr(i).Item(sValueMember).ToString.ToLower.StartsWith(sText.ToLower) Then ' Bắt đầu là Mã
                        dr(i).Item(sValueMember & "_SortID") = 2
                    ElseIf sValueName <> "" AndAlso (dr(i).Item(sValueName).ToString.ToLower.StartsWith(sText.ToLower) Or dr(i).Item(sValueName & "_Unsign").ToString.ToLower.StartsWith(sText.ToLower)) Then ' Bắt đầu là Tên
                        dr(i).Item(sValueMember & "_SortID") = 1
                    End If
                Next
                dtCombo.DefaultView.Sort = sValueMember & "_SortID DESC, " & sValueMember & " ASC"
            End If

            If iCountRow = 0 And Not bF4 Then 'Không có dòng nào
                tdbc.Text = ""
            ElseIf iCountRow = 1 And Not bF4 Then 'Chỉ có 1 dòng
                If dicComboObjectTypeID IsNot Nothing AndAlso dicComboObjectTypeID.Count > 0 Then
                    If dicComboObjectTypeID.ContainsKey(tdbc.FindForm.Name & "_" & tdbc.Name) Then
                        Dim C1ComboObjectTypeID As C1.Win.C1List.C1Combo = dicComboObjectTypeID(tdbc.FindForm.Name & "_" & tdbc.Name)
                        If C1ComboObjectTypeID IsNot Nothing Then
                            C1ComboObjectTypeID.SelectedValue = dtCombo.Rows(0).Item(C1ComboObjectTypeID.ValueMember)
                        End If
                    End If
                End If
                tdbc.SelectedValue = dtCombo.Rows(0).Item(tdbc.ValueMember).ToString
                If tdbc.DisplayMember <> tdbc.ValueMember Then
                    tdbc.Text = dtCombo.Rows(0).Item(tdbc.DisplayMember).ToString
                Else
                    tdbc.Text = dtCombo.Rows(0).Item(tdbc.ValueMember).ToString
                End If
            Else
                Dim f As D99F5555
                f = New D99F5555(tdbc, dtCombo)

                If dicComboObjectTypeID IsNot Nothing AndAlso dicComboObjectTypeID.Count > 0 Then
                    If dicComboObjectTypeID.ContainsKey(tdbc.FindForm.Name & "_" & tdbc.Name) Then
                        f.C1ComboObjectTypeID = dicComboObjectTypeID(tdbc.FindForm.Name & "_" & tdbc.Name)
                    End If
                End If

                f.ShowDialog()
                If f.IsEscape Then ' trường hợp nhập giá trị có trong combo mà nhấn Escape thì phải set trắng
                    f.Dispose()
                    If Not bF4 Or bOld_bTextChanged Then tdbc.Text = ""
                    tdbc.Focus()
                    Exit Sub
                End If
                If tdbc.Text = "+" Then
                    UseEnterAsTab(tdbc.FindForm)
                End If
                f.Dispose()
                bTextChanged = False
            End If
        Catch ex As Exception
            D99C0008.MsgL3(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    ''' <summary>
    ''' Tìm kiếm theo tên cho combo có check chọn
    ''' </summary>
    ''' <param name="tdbc">Combo cần filter</param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub FilterCheckCombo(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal e As System.EventArgs)
        Try
            If tdbc.ReadOnly Then Exit Sub ' Return False
            If tdbc.Columns.Count < 2 Then D99C0008.MsgL3("FilterCombo: Invalid columns") : Exit Sub

            ' Kiểm tra gọi từ sự kiện nào?
            Dim bF4 As Boolean = e.GetType().Name = "CancelEventArgs" ' Or e.GetType().Name = "KeyEventArgs"

            Dim listID As New List(Of String)
            Dim listName As New List(Of String)
            GetListField(tdbc, listID, listName)
            Dim sValueMember As String = tdbc.ValueMember
            Dim sValueName As String = ""
            If listName.Count > 0 Then sValueName = listName(0)

            If Not bF4 And Not bTextChanged Then
                If tdbc.ValueMember = tdbc.DisplayMember Then tdbc.Text = tdbc.Text.ToUpper ' Khi nhập xong -> CtrL mà có trong ds thi viết hoa Mã lên
                bTextChanged = False : Exit Sub ' Return False
            End If
            bTextChanged = False

            If Not bF4 And tdbc.Text = "" Then tdbc.Tag = "" : Exit Sub ' Return False

            Dim dtCombo As DataTable = CType(tdbc.DataSource, DataTable)
            If dtCombo Is Nothing Then D99C0008.MsgL3("FilterCombo: Invalid datasource.")

            ' Khi thay đổi source tu set lại combo.text = "" nên phải set lại giá trị cũ
            'Dim sOldID As String = ""
            'For i As Integer = 0 To listName.Count - 1
            '    If Not dtCombo.Columns.Contains(listName(i) & "_Unsign") Then
            '        If sOldID = "" Then sOldID = tdbc.Text ' chỉ lấy 1 lần đầu tiên
            '        ConvertDataToUnsign(dtCombo, listName(i), listName(i) & "_Unsign")
            '    End If
            'Next
            'If sOldID <> "" Then tdbc.Text = sOldID
            '===============================================

            Dim sText As String = tdbc.Text
            If sText = "" Then tdbc.Tag = ""
            If tdbc.Tag IsNot Nothing Then sText = tdbc.Text
            If sText <> "%" Then sText = SQLSpecialCharactor(sText)
            ' Không hiển thị dòng %, + 
            Dim sColumnFind As String = ""
            'For i As Integer = 0 To listID.Count - 1
            '    If dtCombo.Columns(listID(i)).DataType.Name <> "String" Then Continue For

            '    If sColumnFind <> "" Then sColumnFind &= " And "
            '    sColumnFind &= listID(i) & " <> " & NewCode & " AND " & listID(i) & " <> " & AllCode
            'Next
            'sColumnFind = "(" & sColumnFind & ")"
            sColumnFind = sValueMember & " <> '' "

            Dim sArrFilter() As String = Nothing
            If Not bF4 Then
                Dim sFilterF4 As String = ""
                sArrFilter = sText.Split(";"c)
                For k As Integer = 0 To sArrFilter.Length - 1
                    If sArrFilter(k) = "" Then Continue For
                    If sFilterF4 <> "" Then sFilterF4 &= " OR "
                    Dim sFSub As String = ""
                    For i As Integer = 0 To listID.Count - 1
                        If dtCombo.Columns(listID(i)).DataType.Name <> "String" Then Continue For
                        If sFSub <> "" Then sFSub &= " OR "
                        sFSub &= listID(i) & " like '%" & sArrFilter(k) & "%'"
                    Next
                    sFilterF4 &= sFSub
                    ' Nếu tìm kiếm dạng mới thì luôn tìm theo tên
                    'If DxxFormat.IsFillterName = 1 Then
                    For i As Integer = 0 To listName.Count - 1
                        If sFilterF4 <> "" Then sFilterF4 &= " OR "
                        sFilterF4 &= listName(i) & " like '%" & sArrFilter(k) & "%'"
                        'sFilterF4 &= " Or " & listName(i) & "_Unsign" & " like '%" & sArrFilter(k) & "%'"
                    Next
                    'End If
                Next
                If sFilterF4 <> "" Then sColumnFind &= " AND (" & sFilterF4 & ")"
            End If
            '================================================
            dtCombo = ReturnTableFilter(dtCombo, sColumnFind, True)
            Dim iCountRow As Integer = dtCombo.Rows.Count

            If Not bF4 And iCountRow > 0 Then ' Hiện thị ưu tiên dữ liệu (bằng mã -> bằng tên -> bắt đầu mã -> bắt đầu tên -> có chứa)
                dtCombo.Columns.Add(sValueMember & "_SortID")
                Dim sWhere As String = ""
                sWhere &= sValueMember & " like " & SQLString(sText & "%")
                If sValueName <> "" Then
                    sWhere &= " OR " & sValueName & " like " & SQLString(sText & "%")
                    ' sWhere &= " OR " & sValueName & "_Unsign like " & SQLString(sText & "%")
                End If

                Dim dr() As DataRow = dtCombo.Select(sWhere)
                For i As Integer = 0 To dr.Length - 1
                    If dr(i).Item(sValueMember).ToString.ToLower = sText.ToLower Then ' Bẳng với Mã
                        dr(i).Item(sValueMember & "_SortID") = 4
                    ElseIf sValueName <> "" AndAlso (dr(i).Item(sValueName).ToString.ToLower = sText.ToLower) Then 'Or dr(i).Item(sValueName & "_Unsign").ToString.ToLower = sText.ToLower) Then ' Bằng với Tên
                        dr(i).Item(sValueMember & "_SortID") = 3
                    ElseIf dr(i).Item(sValueMember).ToString.ToLower.StartsWith(sText.ToLower) Then ' Bắt đầu là Mã
                        dr(i).Item(sValueMember & "_SortID") = 2
                    ElseIf sValueName <> "" AndAlso (dr(i).Item(sValueName).ToString.ToLower.StartsWith(sText.ToLower)) Then ' Or dr(i).Item(sValueName & "_Unsign").ToString.ToLower.StartsWith(sText.ToLower)) Then ' Bắt đầu là Tên
                        dr(i).Item(sValueMember & "_SortID") = 1
                    End If
                Next
                dtCombo.DefaultView.Sort = sValueMember & "_SortID DESC, " & sValueMember & " ASC"
            End If

            If iCountRow = 0 And Not bF4 Then 'Không có dòng nào
                tdbc.Text = ""
                tdbc.Tag = ""
            ElseIf iCountRow = 1 And Not bF4 Then 'Chỉ có 1 dòng
                'If tdbc.DisplayMember <> tdbc.ValueMember Then
                '    tdbc.Text = rL3("Du_lieu_da_chon") & " (1)"
                'Else
                '    tdbc.Text = dtCombo.Rows(0).Item(tdbc.ValueMember).ToString
                'End If
                'Phải gán tag sau khi gán text
                tdbc.Text = dtCombo.Rows(0).Item(tdbc.DisplayMember).ToString
                tdbc.Tag = dtCombo.Rows(0).Item(tdbc.ValueMember).ToString
            Else
                Dim sSeparator As String = L3String(tdbc.Columns(0).Tag) 'Separator được set từ UseFilterCheckCombo()
                If sSeparator = "" Then sSeparator = ";"
                Dim f As D99F5555
                f = New D99F5555(tdbc, dtCombo, sSeparator)
                f.ShowDialog()
                If f.IsEscape Then ' trường hợp nhập giá trị có trong combo mà nhấn Escape thì phải set trắng
                    If tdbc.DisplayMember = tdbc.ValueMember Then
                        CheckValueExistCombo(tdbc)
                    Else
                        If bTextChanged Then
                            tdbc.Text = ""
                            tdbc.Tag = ""
                            bTextChanged = False
                        End If
                    End If
                    tdbc.Focus()
                    Exit Sub
                End If
                bTextChanged = False
            End If
            If tdbc.DisplayMember = tdbc.ValueMember Then
                CheckValueExistCombo(tdbc)
            End If
        Catch ex As Exception
            D99C0008.MsgL3(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    ''' <summary>
    ''' Kiểm tra giá trị tồn tại trong combo theo mảng check chọn
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <remarks></remarks>
    Public Sub CheckValueExistCombo(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bCount As Boolean = False)
        If tdbc.Tag Is Nothing Then tdbc.Text = "" : tdbc.Tag = "" : Exit Sub
        Dim sSeparator As String = L3String(tdbc.Columns(0).Tag) 'Separator được set từ UseFilterCheckCombo()
        If sSeparator = "" Then sSeparator = ";"
        Dim dtCombo As DataTable = CType(tdbc.DataSource, DataTable)
        If dtCombo Is Nothing Then D99C0008.MsgL3("FilterCombo: Invalid datasource.")
        Dim sArrFilter() As String = Strings.Split(L3String(tdbc.Tag), sSeparator)
        Dim sValue As String = ""
        If bCount Then
            Dim sTag As String = L3String(tdbc.Tag) 'Dùng set tag lại khi set '' trong text_change
            If sArrFilter.Length <= 1 Then
                tdbc.SelectedValue = sTag
            Else
                sValue = rL3("Du_lieu_da_chon") & " (" & sArrFilter.Length & ")"
                tdbc.Text = sValue
            End If
            tdbc.Tag = sTag
        Else
            If sSeparator <> ";" Then
                For i As Integer = 0 To sArrFilter.Length - 1
                    Dim dr() As DataRow = dtCombo.Select(tdbc.ValueMember & "=" & SQLString(sArrFilter(i)))
                    If dr.Length > 0 Then
                        If sValue <> "" Then sValue &= sSeparator
                        sValue &= sArrFilter(i)
                    End If
                Next
            Else
                sValue = L3String(tdbc.Tag)
            End If
            tdbc.Text = sValue 'Phải set text trước khi set tag
            tdbc.Tag = sValue
        End If
    End Sub

    ''' <summary>
    ''' Lấy giá trị trả về của combo có check chọn khi lưu
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <param name="bReplace">Có convert chuỗi trả về theo cấu trúc ('A','B') hay không</param>
    ''' <param name="sSeparator">Dấu phân cách giữa các giá trị trên combo</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SQLCombo(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bReplace As Boolean = False, Optional ByVal sSeparator As String = ";") As String
        Dim sValue As String = ""
        If tdbc.Tag Is Nothing Then
            sValue = tdbc.Text
        Else
            sValue = L3String(tdbc.Tag)
        End If
        If bReplace Then
            sValue = "('" & sValue.Replace(sSeparator, "','") & "')"
        End If
        Return SQLString(sValue)
    End Function

    ''' <summary>
    ''' Gán giá trị cho combo có check chọn theo giá trị truyền vào
    ''' </summary>
    ''' <param name="tdbc"></param>
    ''' <param name="sValue">Giá trị được gán lên combo</param>
    ''' <remarks></remarks>
    Public Sub SetValueCombo(ByRef tdbc As C1.Win.C1List.C1Combo, ByVal sValue As Object)
        If L3String(sValue) = "" Then
            tdbc.Text = ""
            tdbc.Tag = ""
        Else
            tdbc.Tag = L3String(sValue)
            CheckValueExistCombo(tdbc, tdbc.DisplayMember <> tdbc.ValueMember)
        End If
    End Sub

    Private Sub GetListField(ByVal tdbc As C1.Win.C1List.C1Combo, ByRef listID As List(Of String), ByRef listName As List(Of String))
        For i As Integer = 0 To tdbc.Columns.Count - 1
            If tdbc.Splits(0).DisplayColumns(i).Visible Then
                If IsColumnName(tdbc.Columns(i).DataField) Then
                    listName.Add(tdbc.Columns(i).DataField)
                Else
                    listID.Add(tdbc.Columns(i).DataField)
                End If
            End If
        Next
    End Sub

    Private Sub GetListField(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByRef listID As List(Of String), ByRef listName As List(Of String))
        For i As Integer = 0 To tdbd.Columns.Count - 1
            If tdbd.DisplayColumns(i).Visible Then
                If IsColumnName(tdbd.Columns(i).DataField) Then
                    listName.Add(tdbd.Columns(i).DataField)
                Else
                    listID.Add(tdbd.Columns(i).DataField)
                End If
            End If
        Next
    End Sub

    Public Function IsColumnName(ByVal sField As String) As Boolean
        If sField.Contains("Name") Then
            Return True
        ElseIf sField.Contains("Note") Then
            Return True
        ElseIf sField.Contains("Description") Then
            Return True
        ElseIf sField.Contains("Address") Then
            Return True
        End If

        Return False
    End Function
#End Region

#Region "Filter dropdown"

    Private isDisableAfterColUpdate As Boolean = False
    Private isBeforeColUpdate As Boolean = False '  Nhận dạng gọi từ sự kiện nào. True: Khi update dữ liệu, False khi gọi từ Keydown hay ButtonClick

    ''' <summary>
    ''' Khái báo sử dụng tính năng FilterDropdown (Mặc định check D91T0025)
    ''' </summary>
    ''' <param name="C1Grid"></param>
    ''' <param name="iCol">Cột cần filter</param>
    ''' <remarks></remarks>
    Public Sub UseFilterDropdown(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray iCol() As Integer)
        UseFilterDropdown(C1Grid, True, iCol)
    End Sub

    ''' <summary>
    ''' Khái báo sử dụng tính năng FilterDropdown (Mặc định check D91T0025)
    ''' </summary>
    ''' <param name="C1Grid"></param>
    ''' <param name="bCheckD91">Kiểm tra trường LoadFormNotINV của bảng D91T0025</param>
    ''' <param name="iCol">Cột cần filter</param>
    ''' <remarks></remarks>
    Public Sub UseFilterDropdown(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal bCheckD91 As Boolean, ByVal ParamArray iCol() As Integer)
        If iCol Is Nothing Then D99C0008.MsgL3("UseFilterDropdown: Columns invalid") : Exit Sub

        If bCheckD91 Then
            If DxxFormat.LoadFormNotINV = 0 Then Exit Sub
        End If

        For j As Integer = 0 To iCol.Length - 1
            C1Grid.Columns(iCol(j)).DropDown = Nothing
        Next

        For i As Integer = 0 To C1Grid.Splits.Count - 1
            For j As Integer = 0 To iCol.Length - 1
                C1Grid.Splits(i).DisplayColumns(iCol(j)).Button = True
                C1Grid.Splits(i).DisplayColumns(iCol(j)).Style.ForeColor = Color.Blue
            Next
        Next
        AddHandler C1Grid.BeforeColUpdate, AddressOf C1Grid_BeforeColUpdate
    End Sub

    ''' <summary>
    ''' Khái báo sử dụng tính năng FilterDropdown (Mặc định check D91T0025)
    ''' </summary>
    ''' <param name="C1Grid"></param>
    ''' <param name="sCol">Cột cần filter</param>
    ''' <remarks></remarks>
    Public Sub UseFilterDropdown(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray sCol() As String)
        UseFilterDropdown(C1Grid, True, sCol)
    End Sub

    ''' <summary>
    ''' Khái báo sử dụng tính năng FilterDropdown
    ''' </summary>
    ''' <param name="C1Grid"></param>
    ''' <param name="sCol">Cột cần filter</param>
    ''' <remarks></remarks>
    Public Sub UseFilterDropdown(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal bCheckD91 As Boolean, ByVal ParamArray sCol() As String)
        If sCol Is Nothing Then D99C0008.MsgL3("UseFilterDropdown: Columns invalid") : Exit Sub

        Dim iCol(sCol.Length - 1) As Integer
        For i As Integer = 0 To sCol.Length - 1
            iCol(i) = IndexOfColumn(C1Grid, sCol(i))
        Next
        UseFilterDropdown(C1Grid, iCol)
    End Sub

    <DebuggerStepThrough()> _
    Private Sub C1Grid_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs)
        isBeforeColUpdate = True
    End Sub

    ''' <summary>
    ''' Kiểm tra phím gọi từ sự kiện Keydown của lưới
    ''' </summary>
    ''' <param name="C1Grid"></param>
    ''' <param name="e"></param>
    ''' <param name="bCheckD91">Kiểm tra trường LoadFormNotINV của bảng D91T0025</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckKeydownFilterDropdown(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal e As System.Windows.Forms.KeyEventArgs, Optional ByVal bCheckD91 As Boolean = True) As Boolean
        If (e.Alt And e.KeyCode = Keys.Down) Or e.KeyCode = Keys.F2 Or e.KeyCode = Keys.F4 Then ' Bổ sung tiện ích tìm kiếm theo tên
            If bCheckD91 Then
                If DxxFormat.LoadFormNotINV = 0 Then Return False
            End If
            If C1Grid.Splits(C1Grid.SplitIndex).DisplayColumns(C1Grid.Col).Locked Then Return False

            ' e.Handled = True
            Return True
        End If
        Return False
    End Function

    ''' <summary>
    ''' Tìm kiếm theo tên cho dropdown
    ''' </summary>
    ''' <param name="C1Grid"></param>
    ''' <param name="e"></param>
    ''' <param name="tdbd">Dropdown cần filter</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FilterDropdown(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal e As System.EventArgs, ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown) As DataRow()
        Return FilterDropdown(C1Grid, e, tdbd, False)
    End Function

    ''' <summary>
    ''' Tìm kiếm theo tên cho dropdown
    ''' </summary>
    ''' <param name="C1Grid"></param>
    ''' <param name="e"></param>
    ''' <param name="tdbd">Dropdown cần filter</param>
    ''' <param name="bAllowAddRow">Cho phép chọn để thêm dòng (VD: Mã hàng)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FilterDropdown(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal e As System.EventArgs, ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal bAllowAddRow As Boolean) As DataRow()
        If C1Grid.AllowUpdate = False Then Return Nothing

        '======   Nhận dạng gọi từ sự kiện nào   ==============
        Dim bF4 As Boolean = False ' Nhận dạng gọi từ sự kiện nào
        If e.GetType.Name = "KeyEventArgs" Then
            bF4 = True
            Dim e1 As KeyEventArgs = CType(e, KeyEventArgs)
            e1.Handled = True ' Fix lỗi nhấn Alt+down bị nhảy xuống dưới
        ElseIf e.GetType.Name = "ColEventArgs" Then
            If Not isBeforeColUpdate Then ' gọi từ sự kiện ButtonClick
                bF4 = True
            End If
        End If
        isBeforeColUpdate = False

        '======   Chặn lại 1 số sự kiện khi bị chạy nhiều lần   ==============
        If Not bF4 Then 'If e.KeyCode = Keys.Enter Then
            If isDisableAfterColUpdate Then
                isDisableAfterColUpdate = False
                Return Nothing
            End If
            If C1Grid.Columns(C1Grid.Col).Text = "" Or C1Grid.Bookmark <> C1Grid.Row Then Return Nothing
        Else
            isDisableAfterColUpdate = True
            '  4/7/2014 67039 - Trường hợp nhập xong xóa -> Ctrl+L xong -> mất dữ liệu
            C1Grid.UpdateData() ' Trong khi dang update mà nhấn Ctrl+l thì phai update cho xong 1 thao tác
            isDisableAfterColUpdate = False
        End If
        '=======================================================

        Dim iCountRow As Integer

        Dim listID As New List(Of String)
        Dim listName As New List(Of String)
        GetListField(tdbd, listID, listName)
        Dim sValueMember As String = tdbd.ValueMember
        Dim sValueName As String = ""
        If listName.Count > 0 Then sValueName = listName(0)

        Dim dtSource, dtSourceTemp As DataTable
        If tdbd.DataSource Is Nothing Then
            If L3String(tdbd.Tag) = "" Then D99C0008.MsgL3("FilterDropdown: Invalid SQL string.") : Return Nothing
            LoadDataSource(tdbd, L3String(tdbd.Tag), gbUnicode)
        End If

        dtSource = CType(tdbd.DataSource, DataTable)
        If dtSource Is Nothing Then D99C0008.MsgL3("FilterDropdown: Invalid datasource.") : Return Nothing

        '******************************
        Dim sFilter As String = C1Grid.Columns(C1Grid.Col).Text
        If sFilter <> "%" Then sFilter = SQLSpecialCharactor(C1Grid.Columns(C1Grid.Col).Text)
        Dim sArrFilter() As String = Nothing
        'Dim sFieldColName_Unsign As String = sFieldName & "_Unsign" ' Luu chuổi tên hang ko dấu
        Dim sColumnFind As String = ""
        '        For i As Integer = 0 To listID.Count - 1
        '            If dtSource.Columns(listID(i)).DataType.Name <> "String" Then Continue For
        '            If sColumnFind <> "" Then sColumnFind &= " And "
        '
        '            sColumnFind &= listID(i) & " <> " & NewCode & " AND " & listID(i) & " <> " & AllCode
        '        Next
        '        sColumnFind = "(" & sColumnFind & ")"
        ' sColumnFind = sFieldID & " <> " & NewCode & " AND " & sFieldID & " <> " & AllCode
        sColumnFind = sValueMember & " <> '' "

        If Not bF4 Then
            If sFilter = "" Then Return Nothing
            sArrFilter = sFilter.Split(";"c)

            For i As Integer = 0 To sArrFilter.Length - 1
                Dim sSubF As String = ""
                For j As Integer = 0 To listID.Count - 1
                    If Not dtSource.Columns.Contains(listID(j)) OrElse dtSource.Columns(listID(j)).DataType.Name <> "String" Then Continue For
                    If sSubF <> "" Then sSubF &= " OR "
                    sSubF &= listID(j) & " like " & SQLString("%" & sArrFilter(i).Trim & "%")
                Next
                If DxxFormat.IsFillterName = 1 Then
                    For j As Integer = 0 To listName.Count - 1
                        sSubF &= " Or " & listName(j) & " like " & SQLString("%" & sArrFilter(i).Trim & "%")
                        sSubF &= " Or " & listName(j) & "_Unsign" & " like " & SQLString("%" & sArrFilter(i).Trim & "%")
                    Next
                End If
                sColumnFind &= " And " 'If sColumnFind <> "" Then sColumnFind &= " And "

                sColumnFind &= " (" & sSubF & ")"
            Next
        End If

        '======  Bổ sung cột dữ liệu không dấu cho cột Tên ===================
        For i As Integer = 0 To listName.Count - 1
            If Not dtSource.Columns.Contains(listName(i) & "_Unsign") Then
                ConvertDataToUnsign(dtSource, listName(i), listName(i) & "_Unsign")
            End If
        Next
        '  If Not dtSource.Columns.Contains(sFieldColName_Unsign) Then ConvertDataToUnsign(dtSource, sFieldName, sFieldColName_Unsign)

        dtSourceTemp = ReturnTableFilter(dtSource, sColumnFind, True)
        iCountRow = dtSourceTemp.Rows.Count

        If Not bF4 And iCountRow > 0 Then
            '=======  sắp xếp lại theo thứ tự hiển dữ liệu kho show màn hình D99F5556 ==========================
            dtSourceTemp.Columns.Add(sValueMember & "_SortID")
            Dim sWhere As String = ""
            sWhere &= sValueMember & " like " & SQLString(sArrFilter(0) & "%")
            ' Nếu tìm kiếm dạng mới thì luôn tìm theo tên
            '  If DxxFormat.IsFillterName = 1 Then
            sWhere &= " OR " & sValueName & " like " & SQLString(sArrFilter(0) & "%")
            sWhere &= " OR " & sValueName & "_Unsign like " & SQLString(sArrFilter(0) & "%")
            'End If

            Dim dr() As DataRow = dtSourceTemp.Select(sWhere)
            For i As Integer = 0 To dr.Length - 1
                If dr(i).Item(sValueMember).ToString.ToLower = sArrFilter(0).ToLower Then ' Bẳng với Mã
                    dr(i).Item(sValueMember & "_SortID") = 4
                ElseIf sValueName <> "" AndAlso (dr(i).Item(sValueName).ToString.ToLower = sArrFilter(0).ToLower Or dr(i).Item(sValueName & "_Unsign").ToString.ToLower = sArrFilter(0).ToLower) Then ' Bằng với Tên
                    dr(i).Item(sValueMember & "_SortID") = 3
                ElseIf dr(i).Item(sValueMember).ToString.ToLower.StartsWith(sArrFilter(0).ToLower) Then ' Bắt đầu là Mã
                    dr(i).Item(sValueMember & "_SortID") = 2
                ElseIf sValueName <> "" AndAlso (dr(i).Item(sValueName).ToString.ToLower.StartsWith(sArrFilter(0).ToLower) Or dr(i).Item(sValueName & "_Unsign").ToString.ToLower.StartsWith(sArrFilter(0).ToLower)) Then ' Bắt đầu là Tên
                    dr(i).Item(sValueMember & "_SortID") = 1
                End If
            Next
            dtSourceTemp.DefaultView.Sort = sValueMember & "_SortID DESC, " & sValueMember & " ASC"
        End If

        If iCountRow = 0 And Not bF4 Then 'Không có Mã hàng nào
            '  C1Grid.Columns(C1Grid.Col).Text = ""' Lê Anh Vũ: 10/09/2015: Không Set lại = "" - Sẽ bị lõi trong TH Ana cho phép nhập ngoài danh sách.
            Return Nothing
        ElseIf iCountRow = 1 And Not bF4 Then 'Chỉ có 1 Mã hàng
            Dim drF() As DataRow = dtSourceTemp.Select(sValueMember & " = " & SQLString(dtSourceTemp.Rows(0).Item(sValueMember).ToString))
            Return drF
        Else
            Dim f As D99F5556
            f = New D99F5556(C1Grid, tdbd, dtSourceTemp, bAllowAddRow)
            f.ShowDialog()
            If f.DialogResult = Windows.Forms.DialogResult.OK Then
                Dim drF() As DataRow = f.drResultFind ' dtSourceTemp.Select(f.SResultFind)
                f.Dispose()
                Return drF
            Else
                f.Dispose()
                Return Nothing
            End If
        End If
    End Function

#End Region

#Region "Load nguồn co combo, dropdown dạng mới, khi click vao moi đổ nguồn"

    ''' <summary>
    ''' Đổ nguồn cho dropdown (Khi click vào dropdwon mới đổ nguồn). Dùng Tag cua dropdown dể giữ câu đổ nguồn
    ''' </summary>
    ''' <param name="tdbd"></param>
    ''' <param name="sSQL">Câu đổ nguồn cho Dropdown</param>
    ''' <remarks></remarks>
    Public Sub LoadStringSource(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sSQL As String)
        tdbd.Tag = sSQL
        tdbd.DataSource = Nothing
    End Sub


    'Public Sub LoadStringSource(ByVal tdbc As L3Combo, ByVal sSQL As String)
    '    tdbc.SourceString = sSQL
    '    UnicodeConvertFont(tdbc, gbUnicode)

    '    tdbc.IsDataSource = False
    '    tdbc.FilterClient = ""
    '    If tdbc.FullSource IsNot Nothing Then tdbc.FullSource.Clear() : tdbc.FullSource = Nothing

    '    If tdbc.IsDataSource Or tdbc.DataSource IsNot Nothing Then
    '        tdbc.SelectedIndex = -1
    '    End If
    '    '        If Not tdbc.IsDataSource And tdbc.DataSource Is Nothing Then
    '    '            AddHandler tdbc.BeforeOpen, AddressOf L3Combo_BeforeOpen
    '    '            AddHandler tdbc.SelectedValueChanged, AddressOf L3Combo_SelectedValueChanged
    '    '            '  AddHandler tdbc.TextChanged, AddressOf L3Combo_TextChanged
    '    '            AddHandler tdbc.MouseUp, AddressOf L3Combo_MouseUp
    '    '        Else
    '    '            tdbc.SelectedIndex = -1
    '    '        End If
    'End Sub

    'Public Sub LoadStringSource(ByVal tdbc As L3Combo, ByVal sSQL As String)
    '    tdbc.SourceString = sSQL
    '    UnicodeConvertFont(tdbc, gbUnicode)

    '    tdbc.IsDataSource = False
    '    tdbc.FilterClient = ""
    '    If tdbc.FullSource IsNot Nothing Then tdbc.FullSource.Clear() : tdbc.FullSource = Nothing

    '    If Not tdbc.IsDataSource And tdbc.DataSource Is Nothing Then
    '        AddHandler tdbc.BeforeOpen, AddressOf L3Combo_BeforeOpen
    '        AddHandler tdbc.SelectedValueChanged, AddressOf L3Combo_SelectedValueChanged
    '        '  AddHandler tdbc.TextChanged, AddressOf L3Combo_TextChanged
    '        AddHandler tdbc.MouseUp, AddressOf L3Combo_MouseUp
    '    Else
    '        tdbc.SelectedIndex = -1
    '    End If
    'End Sub

    '    <DebuggerStepThrough()> _
    '    Public Function CheckL3ComboLostFocus(ByVal sender As System.Object) As Boolean
    '        ' Kiểm tra nếu combo chưa đổ nguồn thì LostFocus qua vẫn giữ nguyên già trị
    '        If TypeOf (sender) Is L3Combo Then
    '            If CType(sender, L3Combo).DataSource Is Nothing Then Return True
    '        End If
    '
    '        Return False
    '    End Function
    '
    '    <DebuggerStepThrough()> _
    '    Private Sub L3Combo_BeforeOpen(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
    '        Dim tdbc As L3Combo = CType(sender, L3Combo)
    '        If tdbc.IsChangeFilterClient Then
    '            If tdbc.FullSource IsNot Nothing Then
    '                LoadDataSource(tdbc, ReturnTableFilter(tdbc.FullSource, tdbc.FilterClient, True), gbUnicode)
    '                GoTo 1
    '            End If
    '        End If
    '
    '        If tdbc.IsDataSource Then Exit Sub
    '        If tdbc.SourceString = "" Then Exit Sub
    '
    '        If tdbc.FilterClient <> "" Then
    '            tdbc.FullSource = ReturnDataTable(tdbc.SourceString)
    '            LoadDataSource(tdbc, ReturnTableFilter(tdbc.FullSource, tdbc.FilterClient, True), gbUnicode)
    '        Else
    '            tdbc.FullSource = ReturnDataTable(tdbc.SourceString)
    '            LoadDataSource(tdbc, tdbc.FullSource, gbUnicode)
    '        End If
    '
    '1:      tdbc.IsDataSource = True
    '        tdbc.IsChangeFilterClient = False
    '        tdbc.Row = findrowInC1Combo(tdbc, tdbc.Text, tdbc.DisplayMember) ' Focus to current row
    '    End Sub
    '
    '    <DebuggerStepThrough()> _
    '    Private Sub L3Combo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '        Dim tdbc As L3Combo = CType(sender, L3Combo)
    '        If tdbc.DataSource Is Nothing Then tdbc.ValueID = "" : Exit Sub
    '
    '        tdbc.ValueID = L3String(tdbc.SelectedValue)
    '    End Sub
    '
    '    <DebuggerStepThrough()> _
    '    Private Sub L3Combo_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
    '        If e.Button = MouseButtons.Right Then
    '            Dim tdbc As L3Combo = CType(sender, L3Combo)
    '            If tdbc.DataSource Is Nothing Then tdbc.OpenCombo()
    '        End If
    '    End Sub
    '
    '    Public Function ReturnValueC1Combo(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal sField As String = "") As String
    '        If tdbc.SelectedValue Is Nothing OrElse tdbc.Text = "" Then Return ""
    '        If sField = "" Then
    '            Return tdbc.SelectedValue.ToString
    '        Else
    '            Return tdbc.Columns(sField).Text
    '        End If
    '    End Function
    '
    '    Public Function ReturnValueC1Combo(ByVal tdbc As L3Combo, Optional ByVal sField As String = "") As String
    '        If tdbc.DataSource Is Nothing Then
    '            If sField <> "" Then ' Khi chưa đổ nguồn thì không hỗ trợ lấy theo Field truyền vào
    '                D99C0008.MsgL3("L3Combo: Invalid DataSource")
    '                Return ""
    '            Else
    '                If tdbc.DisplayMember = tdbc.ValueMember Then
    '                    Return tdbc.Text
    '                Else
    '                    Return tdbc.ValueID
    '                End If
    '            End If
    '        Else
    '            Return ReturnValueC1Combo(CType(tdbc, C1.Win.C1List.C1Combo), sField)
    '        End If
    '    End Function
#End Region

#Region "Kiểm tra trùng Mã trên lưới khi nhập liệu"

    ''' <summary>
    ''' Kiểm tra trùng Mã trên lưới khi nhập liệu (gọi tại sự kiện BeforeColUpdate )
    ''' </summary>
    ''' <param name="C1Grid"></param>
    ''' <param name="iColumnName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckDuplicateIDOnGrid(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iColumnName As Integer) As Boolean
        'Tại sự kiện tdbgMainProduct_BeforeColUpdate gán e.Cancel = CheckDuplicateRow(dtgrid, tdbg, COL_AccountID)
        Dim sWhereClause As String = C1Grid.Columns(iColumnName).DataField.ToString & "=" & SQLString(C1Grid.Columns(iColumnName).Value)
        Dim dt As DataTable = CType(C1Grid.DataSource, DataTable)
        If dt Is Nothing Then D99C0008.MsgL3("CheckDuplicateIDOnGrid: Invalid Datasource.") : Return True

        Dim iFind As Integer = dt.Select(sWhereClause).Length
        If iFind > 0 AndAlso dt.Rows.IndexOf(dt.Select(sWhereClause)(0)) <> C1Grid.Row Then
            ' "Bạn đã chọn trùng" nối với tên cột kiểm tra
            D99C0008.MsgL3(rL3("MSG000050") & Space(1) & C1Grid.Columns(iColumnName).Caption)
            Return True
        End If
        Return False
    End Function

    ''' <summary>
    ''' Kiểm tra trùng Mã trên lưới khi nhập liệu (gọi tại sự kiện BeforeColUpdate )
    ''' </summary>
    ''' <param name="C1Grid"></param>
    ''' <param name="sColumnName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckDuplicateIDOnGrid(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sColumnName As String) As Boolean
        Return CheckDuplicateIDOnGrid(C1Grid, IndexOfColumn(C1Grid, sColumnName))
    End Function

#End Region


#Region "Lấy Font Caption DGSansSerif  từ hệ thống Control Pane"

    'Public giCheckFontDGSansSerif As Integer = -1 '= -1 là chưa chạy kiểm tra Font; =1 là  font DG Sans Serif"; =0 là font Unicode (Tahoma, ...)

    Public Function CheckFontCaptionIsDGSansSerif() As Boolean
        Dim sFontCaption As String = GetSystemFontCaption()
        Return sFontCaption.Contains("DG Sans Serif")
    End Function

    Private Function GetSystemFontCaption() As String
        Dim KeyName As String = "CaptionFont"
        Dim subKey As String = "Control Panel\Desktop\WindowMetrics"
        Dim baseRegistryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser
        Dim rk As Microsoft.Win32.RegistryKey = baseRegistryKey
        Dim subSk As Microsoft.Win32.RegistryKey = rk.OpenSubKey(subKey)
        If subSk Is Nothing Then
            Return ""
        Else
            Try
                Dim o As Object = subSk.GetValue(KeyName)
                Dim result As String = ""
                Select Case subSk.GetValueKind(KeyName)
                    Case Microsoft.Win32.RegistryValueKind.Binary
                        result = ByteToCharToString(DirectCast(o, Byte()))
                        Exit Select
                End Select
                Return result
            Catch e As Exception
                MessageBox.Show("Errors : " + e.ToString())
                Return Nothing
            End Try
        End If
    End Function

    Private Function ByteToCharToString(ByVal o As Byte()) As String
        Dim result As String = ""
        Dim arrChar As Char() = System.Text.Encoding.UTF8.GetChars(DirectCast(o, Byte()))
        Dim i As Integer = 28
        While i < arrChar.Length - 1
            If arrChar(i) <> ControlChars.NullChar Then 'AndAlso arrChar(i) <> " "c
                result += arrChar(i)
            End If
            System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
        End While
        Return result
    End Function
#End Region


    Public Function CheckSecurity() As Boolean
        'Hàm này dùng cho V4.0
        Dim CommandArgs() As String = Environment.GetCommandLineArgs()
        'MessageBox.Show(" CommandArgs.Length = " & CommandArgs.Length)

        'If CommandArgs.Length <> 3 OrElse CommandArgs(1) <> "/DigiNet" OrElse CommandArgs(2) <> "Corporation" Then
        If CommandArgs.Length <= 1 Then
            If geLanguage = EnumLanguage.Vietnamese Then
                'Update 16/09/2015
                If giCheckFontDGSansSerif = -1 Then  'Khởi tạo lần đầu tiên kiểm tra Font DGSansSerif
                    If CheckFontCaptionIsDGSansSerif() Then
                        giCheckFontDGSansSerif = 1
                    Else
                        giCheckFontDGSansSerif = 0
                    End If
                End If

                If giCheckFontDGSansSerif = 1 Then
                    MessageBox.Show("Thï tóc gãi nèi bè bÊt híp lÖ!", "Th¤ng bÀo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Else
                    MessageBox.Show("Thủ tục gọi nội bộ bất hợp lệ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                MessageBox.Show("Invalid internal system call!", "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End If
            Return False
        End If
        Return True
    End Function


    ''' <summary>
    ''' Gắn giá trị mặc định của table khi lưới thiết lập Thêm mới
    ''' </summary>
    ''' <param name="tdbg"></param>
    ''' <param name="dr">Dòng dữ liệu thêm vào</param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Public Sub SetDefaultValue(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByRef dr As DataRow, ByVal col As String)
        If L3String(tdbg.Columns(col).DefaultValue) = "" Then Exit Sub 'Không default giá trị

        Try
            If tdbg.Columns.IndexOf(tdbg.Columns(col)) >= 0 Then
                If IsDBNull(dr.Item(col)) OrElse dr.Item(col) Is Nothing Then 'chưa có dữ liệu
                    dr.Item(col) = tdbg.Columns(col).DefaultValue
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub SetDefaultValues(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByRef dr As DataRow)
        For i As Integer = 0 To tdbg.Columns.Count - 1
            SetDefaultValue(tdbg, dr, tdbg.Columns(i).DataField)
        Next
    End Sub

    Public Sub SetDataSource(ByVal Ctrl As Windows.Forms.Control, ByRef dtObjectID As DataTable)
        If TypeOf (Ctrl) Is C1.Win.C1List.C1Combo Then
            LoadDataSource(CType(Ctrl, C1.Win.C1List.C1Combo), dtObjectID.DefaultView.ToTable, D99D0241.gbUnicode)
        ElseIf TypeOf (Ctrl) Is C1.Win.C1TrueDBGrid.C1TrueDBDropdown Then
            LoadDataSource(CType(Ctrl, C1.Win.C1TrueDBGrid.C1TrueDBDropdown), dtObjectID.DefaultView.ToTable, D99D0241.gbUnicode)
        End If
    End Sub

    Public Sub SetPeriodDateBySetup(ByRef chkIsPeriod As CheckBox, ByRef chkIsDate As CheckBox)
        If DxxFormat.TimeFilterMode = 0 Then Exit Sub
        chkIsDate.Checked = True
        chkIsPeriod.Checked = False
    End Sub

    Public Sub SetPeriodDateBySetup(ByRef optIsPeriod As RadioButton, ByRef optIsDate As RadioButton)
        If DxxFormat.TimeFilterMode = 0 Then Exit Sub
        optIsDate.Checked = True
    End Sub
End Module
