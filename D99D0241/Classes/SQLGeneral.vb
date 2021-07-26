Public Class SQLGeneral
    ''' <summary>
    ''' SQL lấy dữ liệu Kỳ và Đơn vị ở Desktop
    ''' </summary>
    ''' <param name="iMode">=0 : load nguon cho Don vi; =1 : load nguon cho Ky</param>
    ''' <param name="sDivisionID">chỉ xet cho Mode =1</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SQLStoreD91P0145(ByVal iMode As Integer, Optional ByVal sDivisionID As String = "") As String
        'iMode =0 : load nguon cho Don vi
        'iMode =1 : load nguon cho Ky
        Dim iProduct As Integer = geDisplayProduct '1: la LemonHR
        Dim sSQL As String = ""
        Dim sComment As String = "Load nguon cho Don vi (0)"
        If iMode = 1 Then sComment = " Load nguon cho Ky (1)"
        sSQL &= ("-- " & sComment & vbCrLf)
        sSQL &= "Exec D91P0145 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLString(gsCompanyID) & COMMA 'CompanyID, varchar[50], NOT NULL
        sSQL &= SQLString(sDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(iProduct) & COMMA 'Product, int, NOT NULL
        sSQL &= SQLNumber(iMode) 'Mode, int, NOT NULL
        Return sSQL
    End Function

    ''' <summary>
    ''' Load combo dữ liệu Nhân viên
    ''' </summary>
    ''' <param name="sKeyID"> Điều kiện Where của ObjectID</param>
    ''' <param name="sValueEdit"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SQLEmployeeID(ByVal sKeyID As String, ByVal sValueEdit As String) As String
        'Dim sSQL As String = "--Do nguon combo Nguoi lap" & vbCrLf

        'sSQL &= " SELECT Object.ObjectID as EmployeeID, Object.ObjectName" & UnicodeJoin(gbUnicode) & " as EmployeeName" & vbCrLf
        'sSQL &= " FROM 	Object  WITH(NOLOCK) " & vbCrLf
        'sSQL &= " WHERE (Disabled = 0 And  Object.ObjectTypeID = 'NV'" & vbCrLf
        'sSQL &= " And (	DAG ='' Or DAG In (Select DAGroupID From LemonSys.dbo.D00V0080 " & vbCrLf
        'sSQL &= " Where UserID= " & SQLString(gsUserID) & " ) Or 'LEMONADMIN' = " & SQLString(gsUserID) & " ) " & vbCrLf
        'If sKeyID <> "" Then sSQL &= " And ObjectID = " & SQLString(sKeyID)
        'sSQL &= ")"
        'If sValueEdit <> "" Then sSQL &= " Or ObjectID = " & SQLString(sValueEdit)
        'sSQL &= vbCrLf
        'sSQL &= " ORDER BY ObjectID"
        Return SQLEmployeeID(sKeyID, " =", SQLString(sValueEdit))
    End Function
    ''' <summary>
    ''' Đổ nguồn combo Người lập có bổ sung thêm Select
    ''' </summary>
    ''' <param name="sKeyID"></param>
    ''' <param name="sValueEdit"></param>
    ''' <param name="SQLSelect">bổ sung thêm một số field của table Object. Ví dụ : DebtManagerTypeID, DebtManagerID</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SQLEmployeeID_Select(ByVal sKeyID As String, ByVal sValueEdit As String, SQLSelect As String) As String
        Return SQLEmployeeID(sKeyID, " =", SQLString(sValueEdit), SQLSelect)
    End Function

    Public Shared Function SQLEmployeeID(ByVal sKeyID As String, ByVal lstValueEdit() As String) As String 'bổ sung 22/07/16
        If lstValueEdit Is Nothing Then Return ""

        If lstValueEdit IsNot Nothing AndAlso lstValueEdit.Length = 1 Then
            Return SQLEmployeeID(sKeyID, lstValueEdit(0))
        End If
        Dim sValue As String = ""
        For i As Integer = 0 To lstValueEdit.Length - 1
            If sValue <> "" Then sValue &= ","
            sValue &= SQLString(lstValueEdit(i))
        Next
        Return SQLEmployeeID(sKeyID, " IN (", sValue & ")")
    End Function

    Private Shared Function SQLEmployeeID(ByVal sKeyID As String, ByVal opera As String, ByVal sValueEdit As String, Optional SQLSelect As String = "") As String
        Dim sSQL As String = "--Do nguon combo Nguoi lap" & vbCrLf

        sSQL &= " SELECT Object.ObjectID as EmployeeID, Object.ObjectName" & UnicodeJoin(gbUnicode) & " as EmployeeName" & IIf(SQLSelect <> "", ",", "") & SQLSelect & vbCrLf
        sSQL &= " FROM 	Object  WITH(NOLOCK) " & vbCrLf
        sSQL &= " WHERE (Disabled = 0 And  Object.ObjectTypeID = 'NV'" & vbCrLf
        sSQL &= " And (	DAG ='' Or DAG In (Select DAGroupID From LemonSys.dbo.D00V0080 " & vbCrLf
        sSQL &= " Where UserID= " & SQLString(gsUserID) & " ) Or 'LEMONADMIN' = " & SQLString(gsUserID) & " ) " & vbCrLf
        If sKeyID <> "" Then sSQL &= " And ObjectID = " & SQLString(sKeyID)
        sSQL &= ")"
        If sValueEdit <> "" Then sSQL &= " Or ObjectID " & opera & sValueEdit
        sSQL &= vbCrLf
        sSQL &= " ORDER BY ObjectID"
        Return sSQL
    End Function

    Public Shared Function SQLObjectTypeID(ByVal union As EnumUnionAll, ByVal sValueEdit As String) As String
        Dim sUnicode As String = ""
        Dim sAll As String = ""
        UnicodeAllString(sUnicode, sAll, gbUnicode)
        '***************
        Dim sSQL As String = "--Do nguon Loai doi tuong" & vbCrLf
        sSQL &= "Select ObjectTypeID, " & IIf(geLanguage = EnumLanguage.Vietnamese, "ObjectTypeName", "ObjectTypeName01").ToString & sUnicode & " as ObjectTypeName, 1 as DisplayOrder  "
        sSQL &= "From D91T0005 WITH(NOLOCK) " & vbCrLf
        If sValueEdit <> "" Then sSQL &= " Where 1=1 or ObjectTypeID = " & SQLString(sValueEdit)
        If union = EnumUnionAll.All Then
            sSQL &= "Union All" & vbCrLf
            sSQL &= "Select '%' as ObjectTypeID, " & sAll & " as ObjectTypeName, 0 as DisplayOrder "
        ElseIf union = EnumUnionAll.AddNew Then
            sSQL &= "Union All" & vbCrLf
            sSQL &= "Select " & NewCode & " as ObjectTypeID, " & NewName & " as ObjectTypeName, 0 as DisplayOrder "
        End If
        sSQL &= vbCrLf
        sSQL &= "Order By DisplayOrder, ObjectTypeID"
        Return sSQL
    End Function

    Public Shared Function SQLStoreD91P9106(ByVal ModuleID_xx As String, ByVal sAuditCode As String, _
                           ByVal sEvents As String, _
                           Optional ByVal sDesc1 As String = "", _
                           Optional ByVal sDesc2 As String = "", _
                           Optional ByVal sDesc3 As String = "", _
                           Optional ByVal sDesc4 As String = "", _
                           Optional ByVal sDesc5 As String = "", _
                           Optional ByVal IsAuditDetail As Integer = 0, _
                           Optional ByVal AuditItemID As String = "") As String
        Dim sSQL As String = "--Audit Log" & vbCrLf
        sSQL &= "Exec D91P9106 " 'SQLDateSave(Now.Date)
        sSQL &= "NULL" & COMMA 'AuditDate, datetime, NOT NULL
        sSQL &= SQLString(sAuditCode) & COMMA 'AuditCode, varchar[20], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(ModuleID_xx) & COMMA 'ModuleID, varchar[2], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(sEvents) & COMMA 'EventID, varchar[20], NOT NULL
        sSQL &= "N" & SQLString(sDesc1) & COMMA 'Desc1, varchar[250], NOT NULL
        sSQL &= "N" & SQLString(sDesc2) & COMMA 'Desc2, varchar[250], NOT NULL
        sSQL &= "N" & SQLString(sDesc3) & COMMA 'Desc3, varchar[250], NOT NULL
        sSQL &= "N" & SQLString(sDesc4) & COMMA 'Desc4, varchar[250], NOT NULL
        sSQL &= "N" & SQLString(sDesc5) & COMMA 'Desc5, nvarchar[500], NOT NULL
        sSQL &= SQLNumber(IsAuditDetail) & COMMA 'IsAuditDetail, tinyint, NOT NULL
        sSQL &= SQLString(AuditItemID) 'AuditItemID, varchar[50], NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P9200
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 19/02/2016 03:04:50
    '#---------------------------------------------------------------------------------------------------
    Public Shared Function SQLStoreD91P9200(ByVal TableName As String, ByVal iMode As Integer, Optional ByVal ColVoucherID As String = "", Optional ByVal VoucherID As String = "" _
        , Optional ByVal ColTransID As String = "", Optional ByVal LineType As Integer = 0, Optional ByVal sLineFilter As String = "") As String
        Dim sSQL As String = ""
        sSQL &= ("-- Luu vao bang tam cac gia tri can audit" & vbCrLf)
        sSQL &= "Exec D91P9200 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(TableName) & COMMA 'TableName, varchar[20], NOT NULL
        sSQL &= SQLString(ColVoucherID) & COMMA 'ColVoucherID, varchar[50], NOT NULL
        sSQL &= SQLString(VoucherID) & COMMA 'VoucherID, varchar[50], NOT NULL
        sSQL &= SQLNumber(iMode) & COMMA 'Mode, tinyint, NOT NULL
        sSQL &= SQLString(ColTransID) & COMMA 'ColTransID, varchar[50], NOT NULL
        sSQL &= SQLNumber(LineType) & COMMA 'LineType, int, NOT NULL
        sSQL &= SQLString(sLineFilter) & COMMA 'sLineFilter, varchar[8000], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P5566
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 29/12/2016 01:36:35
    '#---------------------------------------------------------------------------------------------------
    Public Shared Function SQLStoreD91P5566(_Mode As Integer, _ModuleID As String, Optional IsPeriod As Integer = 0) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Load luoi" & vbCrLf)
        sSQL &= "Exec D91P5566 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(gsCompanyID) & COMMA 'CompanyID, varchar[20], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(_ModuleID) & COMMA 'ModuleID, varchar[10], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, varchar[20], NOT NULL
        sSQL &= SQLNumber(_Mode) & COMMA 'Mode, tinyint, NOT NULL
        sSQL &= SQLNumber(IsPeriod) & COMMA 'IsPeriod, varchar[10], NOT NULL
        sSQL &= SQLString(gsLanguage) 'Language, varchar[20], NOT NULL
        Return sSQL
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P5556
    '# Created User: 
    '# Created Date: 29/12/2016 02:03:00
    '#---------------------------------------------------------------------------------------------------
    Public Shared Function SQLStoreD91P5556(_Mode As Integer) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Luu du lieu Khoa so/ Mo so/ Tao ky" & vbCrLf)
        sSQL &= "Exec D91P5556 "
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, tinyint, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, smallint, NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'CreateModuleID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[50], NOT NULL
        sSQL &= SQLNumber(_Mode) & COMMA 'Mode, tinyint, NOT NULL
        sSQL &= SQLNumber(1) & COMMA 'Version, tinyint, NOT NULL
        sSQL &= SQLString(gsLanguage) 'Language, varchar[20], NOT NULL
        Return sSQL
    End Function

    Public Shared Function SQLSelectD91T8888() As String
        Dim sSQL As String = ""
        sSQL = "-- Lay cac gia tri Thiet lap ca nhan" & vbCrLf
        sSQL &= "IF NOT EXISTS (SELECT TOP 1 1 FROM D91T8888 WITH(NOLOCK) WHERE ModuleID ='D00' AND UserID = " & SQLString(gsUserID) & " )" & vbCrLf
        sSQL &= "BEGIN" & vbCrLf
        sSQL &= "SELECT *,  GETDATE() As DateServer FROM D91T8888 WITH(NOLOCK) WHERE ModuleID ='D00' AND UserID = 'LEMONADMIN'" & vbCrLf
        sSQL &= "End" & vbCrLf
        sSQL &= "ELSE" & vbCrLf
        sSQL &= "BEGIN" & vbCrLf
        sSQL &= " SELECT *, GETDATE() As DateServer FROM D91T8888 WITH(NOLOCK) WHERE ModuleID ='D00' AND UserID = " & SQLString(gsUserID) & vbCrLf
        sSQL &= "End" & vbCrLf
        Return sSQL
    End Function
End Class
