Imports System
Public Class D99F5556
    Dim dtGrid As DataTable
    Private WithEvents tdbgO As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Private tdbDD As C1.Win.C1TrueDBGrid.C1TrueDBDropdown
    Dim sFieldNameFilter As String

    Private _sResultFind As String = ""
    Public ReadOnly Property SResultFind() As String
        Get
            Return _sResultFind
        End Get
    End Property

    Private _drResultFind() As Datarow
    Public ReadOnly Property drResultFind() As DataRow()
        Get
            Return _drResultFind
        End Get
    End Property

    Public Sub New(ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dropdown As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByRef dtDD As DataTable, ByVal bVisibleChoose As Boolean)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        tdbDD = dropdown 'tdbg.Columns(tdbg.Col).DropDown

        If bVisibleChoose And tdbg.AllowAddNew And tdbg.Row >= tdbg.RowCount - 1 Then
            dtGrid = dtDD  'CType(dropdown.DataSource, DataTable).DefaultView.ToTable

            If Not dtGrid.Columns.Contains("IsChoose") Then dtGrid.Columns.Add("IsChoose", GetType(Boolean))
            tdbg1.AllowUpdate = True
            AddColumnsChoose()
           
            tdbg1.FetchRowStyles = True

            tdbg1.InsertHorizontalSplit(1)
            tdbg1.Splits(1).DisplayColumns(0).Visible = False
            tdbg1.Splits(1).DisplayColumns(0).AllowSizing = False
            For i As Integer = 0 To tdbDD.Columns.Count - 1
                ' chỉ add các cột hiển thị, ValueMember, DisplayMember
                If tdbDD.DisplayColumns(i).Visible Or tdbDD.Columns(i).DataField = tdbDD.ValueMember Or tdbDD.Columns(i).DataField = tdbDD.DisplayMember Then
                    AddColumns(tdbDD, i, 1)
                End If
            Next

            tdbg1.Splits(0).SplitSizeMode = C1.Win.C1TrueDBGrid.SizeModeEnum.Exact
            tdbg1.Splits(0).SplitSize = 50

            ResetSplitDividerSize(tdbg1)
        Else
            dtGrid = dtDD  'CType(dropdown.DataSource, DataTable).DefaultView.ToTable
            For i As Integer = 0 To tdbDD.Columns.Count - 1
                ' chỉ add các cột hiển thị, ValueMember, DisplayMember
                If tdbDD.DisplayColumns(i).Visible Or tdbDD.Columns(i).DataField = tdbDD.ValueMember Or tdbDD.Columns(i).DataField = tdbDD.DisplayMember Then
                    AddColumns(tdbDD, i, 0)
                End If
            Next
            lblSpace.Visible = False
            lblEscap.Location = lblSpace.Location
            pnlInfo.Height -= lblSpace.Height + 3
            pnlInfo.Location = New Point(pnlInfo.Location.X, pnlInfo.Location.Y + lblSpace.Height + 2)
            tdbg1.Height += lblSpace.Height + 1
        End If

        ' Xóa dòng đầu tiên (cho trường có Thêm mới hoặc Tất cả)
        '    If dtGrid.Rows.Count > 0 AndAlso dtGrid.Rows(0).Item(tdbDD.ValueMember).ToString = "%" Then dtGrid.Rows.RemoveAt(0)

        LoadDataSource(tdbg1, dtGrid, gbUnicode)
        If tdbDD.Width < lblEnter.Width Then
            Me.Width = lblEnter.Width + 10
        Else
            Me.Width = tdbDD.Width + 40
        End If
        Me.Width = tdbDD.Width + 20

        tdbgO = tdbg

        sFieldNameFilter = tdbDD.ValueMember ' tdbg.Columns(tdbg.Col).DataField

        tdbg1.Col = 1 'tdbg.Col
        tdbg1.Row = findrowInGrid(tdbg1, tdbg(tdbg.Row, tdbg.Col), tdbDD.ValueMember) '  findrowInGrid(tdbg1, tdbg.Columns(tdbg.Col).Value, tdbDD.DisplayMember)
    End Sub

    Public Sub New(ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dropdown As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByRef dtDD As DataTable)
        Me.NEW(tdbg, dropdown, dtdD, True)
    End Sub

    Private Sub AddColumnsChoose()
        Dim ColGrid As C1.Win.C1TrueDBGrid.C1DataColumn

        ColGrid = New C1.Win.C1TrueDBGrid.C1DataColumn
        ColGrid.DataField = "IsChoose"
        ColGrid.Caption = rL3("Chon")
        tdbg1.Columns.Add(ColGrid)
        tdbg1.Splits(0).DisplayColumns(ColGrid).Visible = True
        tdbg1.Splits(0).DisplayColumns(ColGrid).Style.Font = FontUnicode(gbUnicode)
        'tdbg1.Splits(0).DisplayColumns(ColGrid).HeadingStyle.Font = New System.Drawing.Font(tdbg.Splits(0).DisplayColumns(COL_Filter).HeadingStyle.Font, tdbg.Splits(0).DisplayColumns(COL_Filter).HeadingStyle.Font.Style)
        tdbg1.Splits(0).DisplayColumns(ColGrid).HeadingStyle.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
        tdbg1.Columns(0).ValueItems.Presentation = C1.Win.C1TrueDBGrid.PresentationEnum.CheckBox
        tdbg1.Splits(0).DisplayColumns(ColGrid).Width = 45
        tdbg1.Splits(0).DisplayColumns(ColGrid).Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
        tdbg1.Splits(0).DisplayColumns(ColGrid).Locked = False
    End Sub

    Private Sub AddColumns(ByVal dropdown As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal iCol As Integer, ByVal iSplit As Integer)
        Dim ColGrid As New C1.Win.C1TrueDBGrid.C1DataColumn
        ColGrid.DataField = dropdown.Columns(iCol).DataField
        ColGrid.Caption = dropdown.Columns(iCol).Caption
        tdbg1.Columns.Add(ColGrid)
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).Visible = True
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).Width = dropdown.DisplayColumns(iCol).Width
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).Style.Font = FontUnicode(gbUnicode)
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).HeadingStyle.Font = New System.Drawing.Font(dropdown.DisplayColumns(iCol).HeadingStyle.Font, dropdown.DisplayColumns(iCol).HeadingStyle.Font.Style)
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).HeadingStyle.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).Style.HorizontalAlignment = dropdown.DisplayColumns(iCol).Style.HorizontalAlignment
        tdbg1.Columns(ColGrid.DataField).NumberFormat = dropdown.Columns(iCol).NumberFormat
        tdbg1.Columns(ColGrid.DataField).ValueItems = dropdown.Columns(iCol).ValueItems
        tdbg1.BorderStyle = Windows.Forms.BorderStyle.Fixed3D

        tdbg1.Splits(iSplit).DisplayColumns(ColGrid).Locked = True
    End Sub

    Dim sFilter As New StringBuilder

    Private Sub tdbg1_FetchRowStyle(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.FetchRowStyleEventArgs) Handles tdbg1.FetchRowStyle
        If dtGrid.Columns.Contains("IsChoose") Then
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
            Dim sFind As String = ""
            ' Trường hợp fiter nhưng không cho chọn
            'If dtGrid.Columns.Contains("IsChoose") Then sFind = "(IsChoose = 1 OR IsChoose is NULL)"
            If dtGrid.Columns.Contains("IsChoose") Then
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

    Private Sub tdbg1_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg1.HeadClick
        tdbg1.AllowSort = Not (tdbg1.Columns(e.ColIndex).DataField = "IsChoose")
    End Sub

    Dim iHeight As Integer = -1
    Private Sub tdbg1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbg1.MouseClick
        iHeight = e.Y
    End Sub

    Private Sub tdbg1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg1.KeyDown
        HotKeyCtrlVOnGrid(tdbg1, e)
        If (e.KeyCode = Keys.Escape) Then
            ' tdbgO.Columns(sFieldNameFilter).Value = ""
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            tdbgO.Focus()
            Me.Close()
        ElseIf (e.KeyCode = Keys.Space) Then
            If tdbg1.FilterActive Then Exit Sub
            If tdbg1.Col <> 0 And dtGrid.Columns.Contains("IsChoose") Then
                tdbg1.Columns(0).Value = Not L3Bool(tdbg1.Columns(0).Value)
            End If
        ElseIf e.KeyCode = Keys.Enter Then
            tdbg1_DoubleClick(Nothing, Nothing)
        End If

    End Sub

    Private Sub tdbg1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg1.DoubleClick
        If tdbg1.FilterActive AndAlso (tdbg1.RowCount = 0 OrElse tdbg1.RowCount > 1) Then Exit Sub
        If sender IsNot Nothing And iHeight <= tdbg1.Splits(0).CaptionHeight Then Exit Sub

        'If tdbg1.RowCount > 0 Then
        '    tdbgO.Columns(sFieldNameFilter).Value = tdbg1.Columns(sFieldNameFilter).Text
        '    'tdbgO.UpdateData()
        '    tdbgO.Focus()
        '    Me.DialogResult = Windows.Forms.DialogResult.OK
        '    Me.Visible = False
        'End If

        Try
            '            Dim strFind As String = ""
            '            ' Trường hợp fiter nhưng không cho chọn
            '            If tdbg1.RowCount = 1 Or Not dtGrid.Columns.Contains("IsChoose") Then
            '                strFind = tdbg1.Columns(sFieldNameFilter).DataField & " = " & SQLString(tdbg1.Columns(sFieldNameFilter).Text)
            '            Else
            '                tdbg1.UpdateData()
            '                Dim dr1() As DataRow = dtGrid.Select("IsChoose =true")
            '                If dr1.Length > 0 Then
            '                    For i As Integer = 0 To dr1.Length - 1
            '                        If strFind = "" Then
            '                            strFind &= tdbg1.Columns(sFieldNameFilter).DataField & " = " & SQLString(dr1(i).Item(tdbg1.Columns(sFieldNameFilter).DataField).ToString)
            '                        Else
            '                            strFind &= " Or " & tdbg1.Columns(sFieldNameFilter).DataField & " = " & SQLString(dr1(i).Item(tdbg1.Columns(sFieldNameFilter).DataField).ToString)
            '                        End If
            '                    Next
            '                Else
            '                    strFind = tdbg1.Columns(sFieldNameFilter).DataField & " = " & SQLString(tdbg1.Columns(sFieldNameFilter).Text)
            '                End If
            '            End If
            '_sResultFind = strFind


            ' Trường hợp fiter nhưng không cho chọn
            If tdbg1.RowCount = 1 Or Not dtGrid.Columns.Contains("IsChoose") Then
                Dim dr As DataRow = dtGrid.DefaultView(tdbg1.Row).Row
                Dim lst As New List(Of DataRow)
                lst.Add(dr)
                _drResultFind = lst.ToArray
            Else
                tdbg1.UpdateData()
                _drResultFind = dtGrid.Select("IsChoose =true")
                If _drResultFind.Length = 0 Then
                    Dim dr As DataRow = dtGrid.DefaultView(tdbg1.Row).Row
                    Dim lst As New List(Of DataRow)
                    lst.Add(dr)
                    _drResultFind = lst.ToArray
                End If
            End If

            tdbgO.Focus()
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()   ' Me.Visible = False
        Catch ex As Exception

        End Try

    End Sub

    Private Sub tdbg1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg1.KeyPress
        If tdbg1.Columns(tdbg1.Col).ValueItems.Presentation = C1.Win.C1TrueDBGrid.PresentationEnum.CheckBox Then
            e.Handled = CheckKeyPress(e.KeyChar)
        ElseIf tdbg1.Splits(tdbg1.SplitIndex).DisplayColumns(tdbg1.Col).Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far Then
            e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End If
    End Sub

    Private Sub D99F5556_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If _drResultFind Is Nothing OrElse _drResultFind.Length < 1 Then ' If _sResultFind = "" Then
            '  tdbgO.Columns(sFieldNameFilter).Value = ""
            Me.DialogResult = Windows.Forms.DialogResult.Cancel

            tdbgO.Focus()
        End If
    End Sub

    Private Sub D99F5555_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Text = rL3("Tim_kiem_du_lieuF") & " - " & Me.Name & UnicodeCaption(gbUnicode)
        ResetColorGrid(tdbg1, 0, tdbg1.Splits.Count - 1)
        LoadLanguage()
        SetResolutionForm(Me)
    End Sub

    Private Sub LoadLanguage()
        lblEnter.Text = rL3("Nhan_phim") & " Enter: " & rL3("Tra_du_lieu_ve_man_hinh_nhap_lieu") 'Nhấn phím Enter: Trả dữ liệu về màn hình nhập liệu
        lblEscap.Text = rL3("Nhan_phim") & " ESC: " & rL3("Thoat") 'Nhấn phím ESC: Thoát
        lblSpace.Text = rL3("Nhan_phim") & " Space: " & rL3("Chon_nhieu_gia_tri") ' Nhấn phím Space: Chọn nhiều giá trị
    End Sub

    Private Sub D99F5555_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        tdbg1.Col = 0
        tdbg1.Focus()
        If tdbg1.Splits.Count > 1 Then
            If tdbg1.Splits(1).HScrollBar.Visible Then
                tdbg1.Splits(0).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.Always
                tdbg1.Refresh()
            End If
        End If
    End Sub

End Class
