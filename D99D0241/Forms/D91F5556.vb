'#-------------------------------------------------------------------------------------
'# Created Date: 30/12/2016 08:52:29 AM
'# Created User: Nguyễn Thị Ánh
'# Description: Tạo kỳ hàng loạt cho nhóm Tài chính và D00E1040
'#-------------------------------------------------------------------------------------
Public Class D91F5556

    Private _moduleID As String = "90" 'xx
    Public WriteOnly Property ModuleID() As String
        Set(ByVal Value As String)
            _moduleID = Value
        End Set
    End Property

    'Private _formIDPermission As String = "D90F5556"
    'Public WriteOnly Property FormIDPermission() As String
    '    Set(ByVal Value As String)
    '        _formIDPermission = Value
    '    End Set
    'End Property

#Region "Const of tdbg - Total of Columns: 8"
    Private Const COL_IsUsed As Integer = 0       ' Chọn
    Private Const COL_ModuleName As Integer = 1   ' Module
    Private Const COL_Period As Integer = 2       ' Kỳ hiện tại
    Private Const COL_PeriodNew As Integer = 3    ' Kỳ muốn tạo
    Private Const COL_ModuleID As Integer = 4     ' ModuleID
    Private Const COL_TableName As Integer = 5    ' TableName
    Private Const COL_TranMonthNew As Integer = 6 ' TranMonthNew
    Private Const COL_TranYearNew As Integer = 7  ' TranYearNew
#End Region



    Private dtGrid As DataTable

    Private Sub D91F5556_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        '  LoadInfoGeneral() 'Load System/ Option /... in DxxD9940
        ResetColorGrid(tdbg)
        gbEnabledUseFind = False
        LoadTDBGrid()
        Loadlanguage()
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Loadlanguage()
        Me.Text = rL3("Tao_ky_moi") & " - " & Me.Name.Replace("D91", "D" & _moduleID) & UnicodeCaption(gbUnicode) 'TÁo kù mìi - D90F5556".Replace("D91", "D" & _moduleID)
        '================================================================ 
        tdbg.Columns(COL_IsUsed).Caption = rL3("Chon")
        tdbg.Columns(COL_Period).Caption = rL3("Ky_hien_tai")
        tdbg.Columns(COL_PeriodNew).Caption = rL3("Ky_muon_tao")
        '================================================================ 
        btnCreate.Text = "&" & rL3("Tao_ky")
        btnClose.Text = rL3("Do_ng") 'Đó&ng
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ''#---------------------------------------------------------------------------------------------------
    ''# Title: SQLStoreD91P5566
    ''# Created User: 
    ''# Created Date: 29/12/2016 01:36:35
    ''#---------------------------------------------------------------------------------------------------
    'Private Function SQLStoreD91P5566() As String
    '    Dim sSQL As String = ""
    '    sSQL &= ("-- Load luoi" & vbCrLf)
    '    sSQL &= "Exec D91P5566 "
    '    sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
    '    sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
    '    sSQL &= SQLString(gsCompanyID) & COMMA 'CompanyID, varchar[20], NOT NULL
    '    sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
    '    sSQL &= SQLString(_moduleID) & COMMA 'ModuleID, varchar[10], NOT NULL
    '    sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, varchar[20], NOT NULL
    '    sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, varchar[20], NOT NULL
    '    sSQL &= SQLNumber(0) 'Mode, tinyint, NOT NULL
    '    Return sSQL
    'End Function

    Private Sub LoadTDBGrid()
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        Dim sSQL As String = SQLGeneral.SQLStoreD91P5566(0, _moduleID)
        dtGrid = ReturnDataTable(sSQL)
        'Cách mới theo chuẩn: Tìm kiếm và Liệt kê tất cả luôn luôn sáng Khi(dt.Rows.Count > 0)
        gbEnabledUseFind = dtGrid.Rows.Count > 0
        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()

        btnCreate.Enabled = tdbg.RowCount > 0
    End Sub

    Private Sub ReLoadTDBGrid()
        Dim strFind As String = ""
        If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter.ToString
        dtGrid.DefaultView.RowFilter = strFind
        ResetGrid()
    End Sub

    Private Sub ResetGrid()
        FooterTotalGrid(tdbg, COL_ModuleName)
    End Sub
    Dim sFilter As New System.Text.StringBuilder()
    'Dim sFilterServer As New System.Text.StringBuilder()
    Dim bRefreshFilter As Boolean = False

    Private Sub tdbg_AfterColUpdate(sender As Object, e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.AfterColUpdate
        Select Case e.ColIndex
            Case COL_PeriodNew
                tdbg.Columns(COL_IsUsed).Text = "1"
        End Select
    End Sub

    Dim bSelect As Boolean = False 'Mặc định Uncheck - tùy thuộc dữ liệu database
    Private Sub HeadClick(ByVal iCol As Integer)
        If tdbg.RowCount <= 0 Then Exit Sub
        Select Case iCol
            Case COL_IsUsed
                L3HeadClick(tdbg, iCol, bSelect)
        End Select
    End Sub

    Private Sub tdbg_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.HeadClick
        HeadClick(e.ColIndex)
    End Sub

    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        Try
            If (dtGrid Is Nothing) Then Exit Sub
            If bRefreshFilter Then Exit Sub
            FilterChangeGrid(tdbg, sFilter) 'Nếu có Lọc khi In
            ReLoadTDBGrid()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        Me.Cursor = Cursors.WaitCursor
        HotKeyCtrlVOnGrid(tdbg, e)
        If e.Control And e.KeyCode = Keys.S Then HeadClick(tdbg.Col)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tdbg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        If tdbg.Columns(tdbg.Col).ValueItems.Presentation = C1.Win.C1TrueDBGrid.PresentationEnum.CheckBox Then
            e.Handled = CheckKeyPress(e.KeyChar)
        ElseIf tdbg.Splits(tdbg.SplitIndex).DisplayColumns(tdbg.Col).Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far Then
            e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P5555
    '# Created User: 
    '# Created Date: 29/12/2016 01:38:22
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD91P5555() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Kiem tra store" & vbCrLf)
        sSQL &= "Exec D91P5555 "
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[50], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[50], NOT NULL
        sSQL &= SQLNumber(0) & COMMA 'Mode, int, NOT NULL
        sSQL &= SQLString(Me.Name) & COMMA 'FormID, varchar[50], NOT NULL
        sSQL &= SQLString(gsCompanyID) & COMMA 'Key01ID, varchar[1000], NOT NULL
        sSQL &= "'','','','',null, null,null, null, null, '',"
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, tinyint, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLNumber(1) 'Version, tinyint, NOT NULL
        Return sSQL
    End Function

    Private Function SQLCreateTable(TableName As String) As String
        Dim sSQL As String = "CREATE TABLE " & TableName & " ("
        sSQL &= "OrderNum		int identity(1,1)," & _
                "ModuleID 		varchar (20)," & _
                "ModuleName 		nvarchar (250)," & _
                "TableName		varchar (20)," & _
                "Period			varchar (20)," & _
                "PeriodNew		varchar (20)," & _
                "TranMonthNew		int," & _
                "TranYearNew		int"
        sSQL &= ")"
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD91T0001s
    '# Created User: 
    '# Created Date: 29/12/2016 01:52:47
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertTables(TableName As String) As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        For i As Integer = 0 To tdbg.RowCount - 1
            If L3Bool(tdbg(i, COL_IsUsed)) = False Then Continue For
            sSQL.Append("Insert Into " & TableName & "(")
            sSQL.Append("ModuleID, ModuleName, TableName, Period, PeriodNew, TranMonthNew, TranYearNew")
            sSQL.Append(") Values(" & vbCrLf)
            sSQL.Append(SQLString(tdbg(i, COL_ModuleID)) & COMMA)
            sSQL.Append(SQLString(tdbg(i, COL_ModuleName)) & COMMA)
            sSQL.Append(SQLString(tdbg(i, COL_TableName)) & COMMA)
            sSQL.Append(SQLString(tdbg(i, COL_Period)) & COMMA)
            sSQL.Append(SQLString(tdbg(i, COL_PeriodNew)) & COMMA)
            sSQL.Append(SQLNumber(tdbg(i, COL_TranMonthNew)) & COMMA)
            sSQL.Append(SQLNumber(tdbg(i, COL_TranYearNew)))
            sSQL.Append(")")

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function


    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        Dim TableName As String = "[#D91P5556]"
        Dim SQLInsert As New StringBuilder
        SQLInsert.Append(SQLInsertTables(TableName))
        If SQLInsert.Length = 0 Then
            D99C0008.MsgL3(rL3("MSG000010"))
            tdbg.Col = COL_IsUsed
            tdbg.Row = 0
            tdbg.Focus()
            Exit Sub
        End If
        SQLInsert.Insert(0, SQLCreateTable(TableName) & vbCrLf)
        Dim SQLCheck As StringBuilder = SQLInsert
        SQLCheck.Append(SQLStoreD91P5555)
        If Not CheckStore(SQLCheck.ToString) Then Exit Sub

        SQLInsert.Append(vbCrLf & SQLGeneral.SQLStoreD91P5556(0))
        If ExecuteSQL(SQLInsert.ToString) Then
            '   RunAuditLogNewPeriod("D" & _moduleID, lblPeriod.Text)
            D99C0008.MsgL3(rL3("Tao_ky_thanh_cong"))
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Else
            D99C0008.MsgL3(rL3("Thuc_thi_khong_thanh_cong"))
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        End If
    End Sub
End Class