'#-------------------------------------------------------------------------------------
'# Created Date: 30/12/2016 08:52:29 AM
'# Created User: Nguyễn Thị Ánh
'# Description: Tạo kỳ hàng loạt cho nhóm Tài chính và D00E1040
'#-------------------------------------------------------------------------------------
Public Class D91F5555

    Private _moduleID As String = "90" 'xx
    Public WriteOnly Property ModuleID() As String
        Set(ByVal Value As String)
            _moduleID = Value
        End Set
    End Property

#Region "Const of tdbg - Total of Columns: 7"
    Private Const COL_IsUsed As Integer = 0     ' Chọn
    Private Const COL_ModuleName As Integer = 1 ' Module
    Private Const COL_Period As Integer = 2     ' Kỳ muốn mở
    Private Const COL_ModuleID As Integer = 3   ' ModuleID
    Private Const COL_TableName As Integer = 4  ' TableName
    Private Const COL_TranMonth As Integer = 5  ' TranMonth
    Private Const COL_TranYear As Integer = 6   ' TranYear
#End Region

    Private dtGrid As DataTable

    Private Sub D91F5555_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        ResetColorGrid(tdbg)
        gbEnabledUseFind = False
        dtPeriod = ReturnDataTable(SQLGeneral.SQLStoreD91P5566(2, _moduleID, 1))
        tdbg_LockedColumns()
        LoadTDBGrid()
        Loadlanguage()
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Dim dtPeriod As DataTable

    Private Sub Loadlanguage()
        Me.Text = rL3("Mo_soF") & " - " & Me.Name.Replace("D91", "D" & _moduleID) & UnicodeCaption(gbUnicode) 'TÁo kù mìi - D90F5556"
        '================================================================ 
        tdbg.Columns(COL_IsUsed).Caption = rL3("Chon")
        tdbg.Columns(COL_Period).Caption = rL3("Ky")
        '================================================================ 
        btnCreate.Text = "&" & rL3("Mo_so")
        btnClose.Text = rL3("Do_ng") 'Đó&ng
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub LoadTDBGrid()
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        Dim sSQL As String = SQLGeneral.SQLStoreD91P5566(2, _moduleID, 0)
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
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tdbg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        If tdbg.Columns(tdbg.Col).ValueItems.Presentation = C1.Win.C1TrueDBGrid.PresentationEnum.CheckBox Then
            e.Handled = CheckKeyPress(e.KeyChar)
        ElseIf tdbg.Splits(tdbg.SplitIndex).DisplayColumns(tdbg.Col).Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far Then
            e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End If
    End Sub

    Private Function SQLCreateTable(TableName As String) As String
        Dim sSQL As String = "CREATE TABLE " & TableName & " ("
        sSQL &= "OrderNum		int identity(1,1)," & _
                "ModuleID 		varchar (20)," & _
                "ModuleName 		nvarchar (250)," & _
                "TableName		varchar (20)," & _
                "Period			varchar (20)," & _
                "TranMonth		int," & _
                "TranYear		int"
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
            sSQL.Append("ModuleID, ModuleName, TableName, Period, TranMonth, TranYear")
            sSQL.Append(") Values(" & vbCrLf)
            sSQL.Append(SQLString(tdbg(i, COL_ModuleID)) & COMMA)
            sSQL.Append(SQLString(tdbg(i, COL_ModuleName)) & COMMA)
            sSQL.Append(SQLString(tdbg(i, COL_TableName)) & COMMA)
            sSQL.Append(SQLString(tdbg(i, COL_Period)) & COMMA)
            sSQL.Append(SQLNumber(tdbg(i, COL_TranMonth)) & COMMA)
            sSQL.Append(SQLNumber(tdbg(i, COL_TranYear)))
            sSQL.Append(")")

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function



    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        Const TableName As String = "[#D91P5556]"
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
        SQLInsert.Append(SQLGeneral.SQLStoreD91P5556(2))
        '   If ExecuteSQL(SQLInsert.ToString) Then'97662 
        If SaveStorebyTrans(SQLInsert.ToString) Then
            gbClosed = False
            '  D99C0008.MsgL3(rL3("Mo_so_thanh_cong"))
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Else
            gbClosed = True
            D99C0008.MsgL3(rL3("Thuc_thi_khong_thanh_cong"))
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        End If
    End Sub

    Private Sub tdbg_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.ComboSelect
        tdbg.UpdateData()
    End Sub


    Private Sub tdbg_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg.BeforeColUpdate
        '--- Kiểm tra giá trị hợp lệ
        Select Case e.ColIndex
            Case COL_Period
                If tdbg.Columns(e.ColIndex).Text <> tdbg.Columns(e.ColIndex).DropDown.Columns(tdbg.Columns(e.ColIndex).DropDown.DisplayMember).Text Then
                    tdbg.Columns(e.ColIndex).Text = ""
                End If
        End Select
    End Sub


    Private Sub tdbg_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.AfterColUpdate
        '--- Gán giá trị cột sau khi tính toán và giá trị phụ thuộc từ Dropdown
        Select Case e.ColIndex
            Case COL_IsUsed
            Case COL_ModuleName
            Case COL_Period
                If tdbg.Columns(e.ColIndex).Text = "" Then
                    'Gắn rỗng các cột liên quan
                    tdbg.Columns(COL_TranMonth).Text = ""
                    tdbg.Columns(COL_TranYear).Text = ""
                    Exit Select
                End If
                tdbg.Columns(COL_TranMonth).Text = tdbdPeriod.Columns("TranMonth").Text
                tdbg.Columns(COL_TranYear).Text = tdbdPeriod.Columns("TranYear").Text
                tdbg.Columns(COL_IsUsed).Text = "1"
            Case COL_ModuleID
            Case COL_TableName
            Case COL_TranMonth
            Case COL_TranYear
        End Select
    End Sub


    Private Sub tdbg_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbg.RowColChange
  If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        '--- Đổ nguồn cho các Dropdown phụ thuộc
        Select Case tdbg.Col
            Case COL_Period
                LoadtdbdPeriod(tdbg(tdbg.Row, COL_ModuleID).ToString)
        End Select
    End Sub

    Private Sub LoadtdbdPeriod(sModule As String)
        LoadDataSource(tdbdPeriod, ReturnTableFilter(dtPeriod, "ModuleID=" & SQLString(sModule)), gbUnicode)
    End Sub

    Private Sub tdbg_LockedColumns()
        tdbg.Splits(SPLIT0).DisplayColumns(COL_ModuleName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
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

End Class