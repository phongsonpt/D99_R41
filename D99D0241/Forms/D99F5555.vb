Imports System
Public Class D99F5555

    Private tdbcO As C1.Win.C1List.C1Combo
    Dim sDisplayFieldName As String = ""
    Dim sDisplayFieldID As String = ""
    Dim dtGrid As DataTable
    Private tdbgO As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Dim sFilter As StringBuilder
    Private ctrName As Control
    Dim bVisibleChoose As Boolean = False
    Dim strSeparator As String = ";"

    Private _isEscape As Boolean = False
    Public ReadOnly Property IsEscape() As Boolean 
        Get
            Return _isEscape
        End Get
    End Property

    Private _c1ComboObjectTypeID As C1.Win.C1List.C1Combo
    Public WriteOnly Property C1ComboObjectTypeID() As C1.Win.C1List.C1Combo
        Set(ByVal Value As C1.Win.C1List.C1Combo)
            _c1ComboObjectTypeID = Value
        End Set
    End Property

    ''' <summary>
    ''' Hàm New dùng cho combo
    ''' </summary>
    ''' <param name="tdbc">Combo cần filter</param>
    ''' ''' <param name="dtCombo">Nguồn load dữ liệu cho Combo cần filter</param>
    ''' <remarks></remarks>
    Public Sub New(ByRef tdbc As C1.Win.C1List.C1Combo, ByVal dtCombo As DataTable)
        ' This call is required by the Windows Form Designer. 
        InitializeComponent()

#If Not Debug Then 'Nếu đang ở trạng thái DEBUG thì ...
        '  If System.IO.File.Exists(Application.StartupPath & "\D99D9241.DLL") Then
        'InitMouseHookListener()
        '  End If
#End If

        'Add các cột tương ứng của combo vào lưới
        '=======================================================
        Dim iCountVisible As Integer = 0
        For i As Integer = 0 To tdbc.Columns.Count - 1
            If tdbc.Splits(0).DisplayColumns(i).Visible Then iCountVisible += 1
            ' chỉ add các cột hiển thị, ValueMember, DisplayMember
            If _c1ComboObjectTypeID IsNot Nothing Then
                If tdbc.Splits(0).DisplayColumns(i).Visible Or tdbc.Columns(i).DataField = tdbc.ValueMember Or tdbc.Columns(i).DataField = tdbc.DisplayMember Or tdbc.Columns(i).DataField = _c1ComboObjectTypeID.ValueMember Then
                    AddColumns(tdbc, i, 0)
                End If
            Else
                If tdbc.Splits(0).DisplayColumns(i).Visible Or tdbc.Columns(i).DataField = tdbc.ValueMember Or tdbc.Columns(i).DataField = tdbc.DisplayMember Then
                    AddColumns(tdbc, i, 0)
                End If
            End If
        Next
        If iCountVisible = 1 Then tdbg1.ColumnHeaders = False

        sDisplayFieldID = tdbc.ValueMember
        sDisplayFieldName = tdbc.DisplayMember

        'Đổ nguồn
        dtGrid = dtCombo 'CType(tdbc.DataSource, DataTable).DefaultView.ToTable
        LoadDataSource(tdbg1, dtGrid, gbUnicode)
        '*******************************************
        If sDisplayFieldName <> "" Then tdbg1.Col = IndexOfColumn(tdbg1, sDisplayFieldName)

        tdbcO = tdbc
        If tdbc.DropDownWidth < lblEnter.Width Then
            Me.Width = lblEnter.Width + 10
        Else
            Me.Width = tdbc.DropDownWidth + 40
        End If

        tdbg1.Row = findrowInGrid(tdbg1, tdbc.Text, tdbc.DisplayMember)
    End Sub

    Public Sub New(ByRef tdbc As C1.Win.C1List.C1Combo, ByVal dtCombo As DataTable, ByVal sSeparator As String)
        ' This call is required by the Windows Form Designer. 
        InitializeComponent()
#If Not Debug Then 'Nếu đang ở trạng thái DEBUG thì ...
        '  If System.IO.File.Exists(Application.StartupPath & "\D99D9241.DLL") Then
        'InitMouseHookListener()
        '  End If
#End If
        'Add các cột tương ứng của combo vào lưới
        '=======================================================
        AddColumnsChoose()
        tdbg1.FetchRowStyles = True
        tdbg1.InsertHorizontalSplit(1)
        tdbg1.Splits(1).DisplayColumns(0).Visible = False
        tdbg1.Splits(1).DisplayColumns(0).AllowSizing = False

        Dim iCountVisible As Integer = 0
        For i As Integer = 0 To tdbc.Columns.Count - 1
            If tdbc.Splits(0).DisplayColumns(i).Visible Then iCountVisible += 1
            ' chỉ add các cột hiển thị, ValueMember, DisplayMember
            If _c1ComboObjectTypeID IsNot Nothing Then
                If tdbc.Splits(0).DisplayColumns(i).Visible Or tdbc.Columns(i).DataField = tdbc.ValueMember Or tdbc.Columns(i).DataField = tdbc.DisplayMember Or tdbc.Columns(i).DataField = _c1ComboObjectTypeID.ValueMember Then
                    AddColumns(tdbc, i, 1)
                End If
            Else
                If tdbc.Splits(0).DisplayColumns(i).Visible Or tdbc.Columns(i).DataField = tdbc.ValueMember Or tdbc.Columns(i).DataField = tdbc.DisplayMember Then
                    AddColumns(tdbc, i, 1)
                End If
            End If
        Next

        tdbg1.Splits(0).SplitSizeMode = C1.Win.C1TrueDBGrid.SizeModeEnum.Exact
        tdbg1.Splits(0).SplitSize = 50

        ResetSplitDividerSize(tdbg1)

        If iCountVisible = 1 Then tdbg1.ColumnHeaders = False
        sDisplayFieldID = tdbc.ValueMember
        sDisplayFieldName = tdbc.DisplayMember
        'Đổ nguồn
        dtGrid = dtCombo 'CType(tdbc.DataSource, DataTable).DefaultView.ToTable
        If dtGrid.Columns.IndexOf("IsChoose") = -1 Then
            dtGrid.Columns.Add("IsChoose", GetType(Boolean))
        End If
        tdbg1.AllowUpdate = True

        '=============================================================
        'Gán check chọn cho các giá trị theo tag
        Dim sFirstValue As String = ""
        If tdbc.Tag IsNot Nothing Then
            Dim sValue() As String = Microsoft.VisualBasic.Split(tdbc.Tag.ToString, sSeparator)
            For i As Integer = 0 To sValue.Length - 1
                Dim drow As DataRow = ReturnDataRow(dtGrid, tdbc.ValueMember & "=" & SQLString(sValue(i)))
                If drow IsNot Nothing Then
                    drow.Item("IsChoose") = True
                    If sFirstValue = "" Then sFirstValue = sValue(i)
                End If
            Next
        End If
        '=============================================================
        LoadDataSource(tdbg1, dtGrid, gbUnicode)
        '*******************************************
        If sDisplayFieldName <> "" Then tdbg1.Col = IndexOfColumn(tdbg1, sDisplayFieldName)

        tdbcO = tdbc
        If tdbc.DropDownWidth < lblEnter.Width Then
            Me.Width = lblEnter.Width + 10
        Else
            Me.Width = tdbc.DropDownWidth + 40
        End If
        bVisibleChoose = True
        strSeparator = sSeparator
        tdbg1.Row = findrowInGrid(tdbg1, sFirstValue, sDisplayFieldID)
    End Sub

    Private Sub AddColumnsChoose()
        Dim ColGrid As C1.Win.C1TrueDBGrid.C1DataColumn
        ColGrid = New C1.Win.C1TrueDBGrid.C1DataColumn
        ColGrid.DataField = "IsChoose"
        ColGrid.Caption = rL3("Chon")
        tdbg1.Columns.Add(ColGrid)
        tdbg1.Splits(0).DisplayColumns(ColGrid).Visible = True
        tdbg1.Splits(0).DisplayColumns(ColGrid).Style.Font = FontUnicode(gbUnicode)
        tdbg1.Splits(0).DisplayColumns(ColGrid).HeadingStyle.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
        tdbg1.Columns(0).ValueItems.Presentation = C1.Win.C1TrueDBGrid.PresentationEnum.CheckBox
        tdbg1.Splits(0).DisplayColumns(0).Width = 45
        tdbg1.Splits(0).DisplayColumns(0).Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
        tdbg1.Splits(0).DisplayColumns(0).Locked = False
    End Sub

    Private Sub AddColumns(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal iCol As Integer, ByVal iSplit As Integer)
        Dim ColGrid As New C1.Win.C1TrueDBGrid.C1DataColumn
        ColGrid.DataField = tdbc.Columns(iCol).DataField
        ColGrid.Caption = tdbc.Columns(iCol).Caption
        tdbg1.Columns.Add(ColGrid)
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).Visible = tdbc.Splits(0).DisplayColumns(ColGrid.DataField).Visible
        If Not tdbg1.Splits(iSplit).DisplayColumns(ColGrid).Visible Then tdbg1.Splits(iSplit).DisplayColumns(ColGrid).AllowSizing = False
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).Width = tdbc.Splits(0).DisplayColumns(ColGrid.DataField).Width
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).Locked = True
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).Style.Font = FontUnicode(gbUnicode)
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).HeadingStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25)
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).HeadingStyle.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).Style.HorizontalAlignment = tdbc.Splits(0).DisplayColumns(ColGrid.DataField).Style.HorizontalAlignment
        tdbg1.Columns(ColGrid.DataField).NumberFormat = tdbc.Columns(iCol).NumberFormat
        tdbg1.Columns(ColGrid.DataField).ValueItems.Presentation = tdbc.Columns(iCol).ValueItems.Presentation
    End Sub

    Dim iHeight As Integer = -1
    Private Sub tdbg1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbg1.MouseClick
        iHeight = e.Y
    End Sub

    Private Sub tdbg1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg1.DoubleClick
        'If tdbg1.FilterActive AndAlso (tdbg1.RowCount = 0 OrElse tdbg1.RowCount > 1) Then Exit Sub
        If tdbg1.RowCount = 0 Then Me.Close() : Exit Sub
        If tdbg1.FilterActive AndAlso tdbg1.RowCount > 1 Then Exit Sub

        If sender IsNot Nothing And iHeight <= tdbg1.Splits(0).CaptionHeight Then Exit Sub

        If bVisibleChoose Then
            Dim sRValue As String = ReturnMultiValue(, sDisplayFieldID)

            If sDisplayFieldName <> sDisplayFieldID Then
                tdbcO.Text = ReturnMultiValue(True) 'Gán text là tập hợp các giá trị theo DisplayMember
            Else
                tdbcO.Text = sRValue
            End If
            tdbcO.Tag = sRValue 'Lưu mã về tag để xử lý
            Me.Close()
        Else
            If _c1ComboObjectTypeID IsNot Nothing Then
                If dtGrid.Columns.Contains(_c1ComboObjectTypeID.ValueMember) Then ' If _c1ComboObjectTypeID.Text = "" Or dtGrid.Columns.Contains(_c1ComboObjectTypeID.ValueMember) Then
                    If tdbg1.Columns(sDisplayFieldID).Text <> "+" Then
                        _c1ComboObjectTypeID.SelectedValue = dtGrid.DefaultView(tdbg1.Row).Row.Item(_c1ComboObjectTypeID.ValueMember).ToString ' tdbg1.Columns(_c1ComboObjectTypeID.ValueMember).Text
                    End If
                End If
            End If
            tdbcO.SelectedValue = tdbg1.Columns(sDisplayFieldID).Text
            tdbcO.Text = tdbg1.Columns(sDisplayFieldName).Text
            If ctrName IsNot Nothing And sDisplayFieldName <> "" Then
                ctrName.Text = tdbg1.Columns(sDisplayFieldName).Text
                ctrName.Focus()
            Else
                Me.Close() 'Me.Dispose()
            End If
        End If
    End Sub

Private Function ReturnMultiValue(Optional ByVal bCount As Boolean = False, Optional ByVal sField As String = "") As String
        tdbg1.UpdateData()
        Dim dr() As DataRow = dtGrid.Select("IsChoose=1")
        Dim sReturn As String = ""
        If dr.Length = 0 Then
            If bCount Then 'DÙng cho gán text khi Display <> Value
                Return tdbg1.Columns(sDisplayFieldName).Text 'rL3("Du_lieu_da_chon") & " (1)"
            End If
            sReturn = tdbg1.Columns(sDisplayFieldID).Text
        Else
            If bCount Then 'DÙng cho gán text khi Display <> Value
                If dr.Length > 1 Then
                    Return rL3("Du_lieu_da_chon") & " (" & dr.Length & ")"
                Else
                    Return L3String(dr(0).Item(sDisplayFieldName)) 'tdbg1.Columns(sDisplayFieldName).Text
                End If
            End If
            For i As Integer = 0 To dr.Length - 1
                If sReturn <> "" Then sReturn &= strSeparator
                sReturn &= dr(i)(sField)
            Next
        End If
        Return sReturn
    End Function

    Private Sub tdbg1_FetchRowStyle(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.FetchRowStyleEventArgs) Handles tdbg1.FetchRowStyle
        If bVisibleChoose Then
            If L3Bool(tdbg1(e.Row, "IsChoose")) Then
                e.CellStyle.ForeColor = Color.Blue
            End If
        End If
    End Sub

    Private Sub tdbg1_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg1.FilterChange
        Try
            If (dtGrid Is Nothing) Then Exit Sub
            sFilter = New StringBuilder("")
            FilterChangeGrid(tdbg1, sFilter) 'Nếu có Lọc khi In
            Dim strFilter As String = sFilter.ToString()
            '            If strFilter <> "" And bVisibleChoose Then
            '                strFilter = "(" & strFilter & ") or IsChoose=1"
            '            End If
            Dim sFind As String = ""
            If bVisibleChoose Then
                If tdbg1.Col > 0 Or L3Bool(tdbg1.Columns("IsChoose").FilterText) Then
                    sFind = " IsChoose = 1 "
                Else
                    sFind = " (IsChoose = 1 OR IsChoose is NULL) "
                End If
            End If

            If sFilter.ToString <> "" Then
                If sFind <> "" Then sFind &= " OR "
                sFind &= sFilter.ToString
            Else
                sFind = ""
            End If

            dtGrid.DefaultView.RowFilter = sFind
        Catch ex As Exception
            WriteLogFile(ex.Message)
        End Try
    End Sub

    Private Sub tdbg1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg1.KeyDown
        HotKeyCtrlVOnGrid(tdbg1, e)
        If (e.KeyCode = Keys.Escape) Then
            'e.Handled = True
            '  tdbcO.Text = ""
            tdbcO.Focus()
            _isEscape = True
            Me.Close() '  Me.Dispose()
        ElseIf (e.KeyCode = Keys.Space) Then
            If tdbg1.FilterActive Then Exit Sub
            If tdbg1.Col <> 0 And dtGrid.Columns.Contains("IsChoose") Then
                tdbg1.Columns(0).Value = Not L3Bool(tdbg1.Columns(0).Value)
            End If
        ElseIf e.KeyCode = Keys.Enter Then
            e.Handled = True
            tdbg1_DoubleClick(Nothing, Nothing)
        End If
    End Sub

    Private Sub D99F5555_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ResetColorGrid(tdbg1, 0, tdbg1.Splits.Count - 1)
        LoadLanguage()
        SetResolutionForm(Me)
        SetLocationForm(tdbcO)
    End Sub

    Private Sub LoadLanguage()
        lblEnter.Text = rL3("Nhan_phim") & " Enter: " & rL3("Tra_du_lieu_ve_man_hinh_nhap_lieu") 'Nhấn phím Enter: Trả dữ liệu về màn hình nhập liệu
        lblEscap.Text = rL3("Nhan_phim") & " ESC: " & rL3("Thoat") 'Nhấn phím ESC: Thoát
    End Sub

    Private Sub SetLocationForm(ByVal tdbc As C1.Win.C1List.C1Combo)
        Me.StartPosition = FormStartPosition.Manual
        Dim pControl As Point = tdbc.PointToScreen(Point.Empty)
        Dim szScreen As Size = Screen.PrimaryScreen.Bounds.Size
        Dim iX, iY As Integer

        If (pControl.X + Me.Width > szScreen.Width) And Me.Width <= pControl.X Then
            iX = pControl.X - (Me.Width - tdbc.Width)
        Else
            iX = pControl.X
        End If
        If (pControl.Y + Me.Height + tdbc.Height + 2 > szScreen.Height) And Me.Height <= pControl.Y Then
            iY = pControl.Y - (Me.Height + 2)
        Else
            iY = pControl.Y + tdbc.Height + 2
        End If
        Me.Location = New Point(iX, iY)
    End Sub

    Dim m_MouseHookManager As MouseHookListener = Nothing
    Private Sub InitMouseHookListener()
        m_MouseHookManager = New MouseHookListener(New GlobalHooker())
        m_MouseHookManager.Enabled = True
        AddHandler m_MouseHookManager.MouseClick, AddressOf HookManager_MouseClick
    End Sub

    <DebuggerStepThrough()> _
    Private Sub HookManager_MouseClick(ByVal sender As Object, ByVal e As MouseEventArgs)
        If _isEscape Then Exit Sub

        ' Click chuột ngoài form thì đóng form
        If e.X >= Me.Left And e.X <= Me.Right And e.Y >= Me.Top And e.Y <= Me.Bottom Then
            Exit Sub
        Else
            m_MouseHookManager.Enabled = False
            ' tdbcO.Focus()
            _isEscape = True
            Me.Close() '  Me.Dispose()
        End If
    End Sub

    Private Sub D99F5555_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If tdbg1.Splits.Count > 1 Then
            If tdbg1.Splits(1).HScrollBar.Visible Then
                tdbg1.Splits(0).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.Always
                tdbg1.Refresh()
            End If
        End If
    End Sub
End Class
