'#######################################################################################
'#                                     CHÚ Ý
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Bổ sung C1Grid.Splits(i).DisplayColumns(dc.DataField).EditorStyle.Font = C1Grid.Font vào hàm LoadDataSource(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dt As DataTable, ByVal bUseUnicode As Boolean)
'# Ngày cập nhật cuối cùng: 18/02/2014
'# Người cập nhật cuối cùng: Nguyễn Thị Minh Hòa
'# Chỉ gắn font 1 lần của các control trong hàm LoadDataSource  26/07/2012
'# Đưa set màu và size chữ vào hàm SetResolutionControl: 18/02/2014
'#######################################################################################
''' <summary>
''' Module liên quan đến các vấn đề của hãng C1
''' </summary>
''' <remarks></remarks>
Public Module D99X0002


#Region "Đổ nguồn cho C1Combo"
    ''' <summary>
    ''' Đổ nguồn cho C1Combo có nguồn truyền vào là câu SQL và có set lại độ rộng của combo
    ''' </summary>
    ''' <param name="C1Combo">Tên C1Combo cần đổ nguồn</param>
    ''' <param name="sSQL">Lệnh SQL đổ nguồn</param>
    ''' <param name="Width">Mảng chiều rộng các cột tương ứng của C1Combo</param>
    '<DebuggerStepThrough()> _
    Public Sub LoadDataSource(ByVal C1Combo As C1.Win.C1List.C1Combo, ByVal sSQL As String, ByVal Width() As Integer, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable = ReturnDataTable(sSQL)
        LoadDataSource(C1Combo, dt, Width, bUseUnicode)
    End Sub

    ''' <summary>
    ''' Đổ nguồn cho C1Combo có nguồn truyền vào là datatable và có set lại độ rộng của combo
    ''' </summary>
    ''' <param name="C1Combo">Tên C1Combo cần đổ nguồn</param>
    ''' <param name="dt">DataTable nguồn</param>
    ''' <param name="Width">Mảng chiều rộng các cột tương ứng của C1Combo</param>
    '<DebuggerStepThrough()> _
    Public Sub LoadDataSource(ByVal C1Combo As C1.Win.C1List.C1Combo, ByVal dt As DataTable, ByVal Width() As Integer, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(C1Combo, dt, bUseUnicode)
        For i As Integer = 0 To C1Combo.Splits(0).DisplayColumns.Count - 1
            C1Combo.Splits(0).DisplayColumns(i).Width = Width(i)
        Next
    End Sub

    ''' <summary>
    ''' Đổ nguồn cho C1Combo có nguồn truyền vào là câu SQL
    ''' </summary>
    ''' <param name="C1Combo">Tên C1Combo cần đổ nguồn</param>
    ''' <param name="sSQL">Lệnh SQL đổ nguồn</param>
    ''' <remarks></remarks>
    '<DebuggerStepThrough()> _
    Public Sub LoadDataSource(ByVal C1Combo As C1.Win.C1List.C1Combo, ByVal sSQL As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable = ReturnDataTable(sSQL)
        LoadDataSource(C1Combo, dt, bUseUnicode)
    End Sub

    ''' <summary>
    ''' Đổ nguồn cho C1Combo có nguồn truyền vào là datatable
    ''' </summary>
    ''' <param name="C1Combo">Tên C1Combo cần đổ nguồn</param>
    ''' <param name="dt">DataTable nguồn</param>
    ''' <remarks></remarks>
    Public Sub LoadDataSource(ByVal C1Combo As C1.Win.C1List.C1Combo, ByVal dt As DataTable, Optional ByVal bUseUnicode As Boolean = False)
        If bUseUnicode = False Then bUseUnicode = gbUnicode 'Fix 11/04/2014 lỗi khi không truyền bUseUnicode hoặc luôn luôn = True
        'Fix 28/06/2011 Thay đổi vị trí cho đúng cột
        Dim i As Integer = 0
1:
        While i < C1Combo.Columns.Count 'Không dùng vòng for vì cột c1combo có thay đổi

            Dim index As Integer = dt.Columns.IndexOf(C1Combo.Columns(i).DataField)
            'Nếu cột trên combo không tồn tại trong table thì xóa cột này
            If index = -1 Then C1Combo.Columns.RemoveAt(i) : GoTo 1
            If index <> i Then dt.Columns(C1Combo.Columns(i).DataField).SetOrdinal(i)
            i += 1
        End While

        '***************
        Dim iMaxDropdownItems As Integer = 9 '8 sửa lại 9 vì có 8 dòng dữ liệu sẽ bị mất 1 dòng 106753
        If dt.Rows.Count < iMaxDropdownItems Then
            Dim dr As DataRow = Nothing
            For i = 0 To iMaxDropdownItems - dt.Rows.Count - 1
                dr = dt.NewRow
                dt.Rows.Add(dr)
            Next
        End If
        Dim arrCaption(C1Combo.Columns.Count - 1) As String
        Dim arrWidth(C1Combo.Splits(0).DisplayColumns.Count - 1) As Integer
        Dim arrVisible(C1Combo.Splits(0).DisplayColumns.Count - 1) As Boolean
        Dim arrHorizontalAlignment(C1Combo.Splits(0).DisplayColumns.Count - 1) As C1.Win.C1List.AlignHorzEnum
        For i = 0 To C1Combo.Columns.Count - 1
            If i >= arrCaption.Length Then Exit For
            arrCaption(i) = C1Combo.Columns(i).Caption
            With C1Combo.Splits(0).DisplayColumns(i)
                arrWidth(i) = .Width
                arrVisible(i) = .Visible
                arrHorizontalAlignment(i) = .Style.HorizontalAlignment
            End With
        Next

        ' Sau khi change DataSource bị mất các giá trị Tag trong cột trước đó, Giữ lại Tag, có thể sử dụng lại (VD: FilterCombo, FilterCheckCombo)
        Dim Tag As Object = Nothing
        If C1Combo.Columns.Count > 0 AndAlso C1Combo.Columns(0).Tag IsNot Nothing Then Tag = C1Combo.Columns(0).Tag
        C1Combo.DataSource = dt
        If C1Combo.Columns.Count > 0 AndAlso Tag IsNot Nothing Then C1Combo.Columns(0).Tag = Tag

        For i = 0 To C1Combo.Columns.Count - 1
            C1Combo.Splits(0).DisplayColumns(i).HeadingStyle.Font = FontUnicode(True)
        Next
        C1Combo.DisplayMember = C1Combo.DisplayMember
        C1Combo.ValueMember = C1Combo.ValueMember
        C1Combo.Font = FontUnicode(bUseUnicode, C1Combo.Font.Style)
        C1Combo.EditorFont = FontUnicode(bUseUnicode, C1Combo.EditorFont.Style)
        'Bổ sung 26/07/2012
        'UnicodeConvertFont(C1Combo, bUseUnicode)
        '**************
        C1Combo.Style.VerticalAlignment = C1.Win.C1TrueDBGrid.AlignVertEnum.Center
        For i = 0 To C1Combo.Columns.Count - 1
            With C1Combo.Splits(0).DisplayColumns(i)
                If i >= arrCaption.Length Then 'Nếu không tồn tại cột này trên combo thì Ẩn đi
                    .Visible = False
                    Continue For
                End If
                .HeadingStyle.HorizontalAlignment = C1.Win.C1List.AlignHorzEnum.Center
                .Width = arrWidth(i)
                .Visible = arrVisible(i)
                .Style.HorizontalAlignment = arrHorizontalAlignment(i)
            End With
            C1Combo.Columns(i).Caption = arrCaption(i)
        Next

        'Update 18/02/2014: Bỏ, đã đưa vào hàm SetResolutionControl
        '        C1Combo.HeadingStyle.Font = FontUnicode()
        '        C1Combo.HighLightRowStyle.BackColor = Color.Green
        '        C1Combo.HighLightRowStyle.ForeColor = SystemColors.HighlightText
        '        C1Combo.SelectedStyle.BackColor = Color.Green
        '        C1Combo.SelectedStyle.ForeColor = SystemColors.HighlightText
    End Sub

#End Region

#Region "Đổ nguồn cho C1DropDown"

    ''' <summary>
    ''' Đổ nguồn cho C1DropDown nhập Unicode
    ''' </summary>
    ''' <param name="C1DropDown">Tên C1DropDown cần đổ nguồn</param>
    ''' <param name="sSQL">Lệnh SQL đổ nguồn</param>
    <DebuggerStepThrough()> _
    Public Sub LoadDataSource(ByVal C1DropDown As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sSQL As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable = ReturnDataTable(sSQL)
        LoadDataSource(C1DropDown, dt, bUseUnicode)
    End Sub

    ''' <summary>
    ''' Đổ nguồn cho C1DropDown nhập Unicode
    ''' </summary>
    ''' <param name="C1DropDown">Tên C1DropDown cần đổ nguồn</param>
    ''' <param name="dt">DataTable nguồn</param>
    <DebuggerStepThrough()> _
    Public Sub LoadDataSource(ByVal C1DropDown As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dt As DataTable, Optional ByVal bUseUnicode As Boolean = False)
        'Modify date: 30/07/2010: Set Font nhập Unicode
        ' If bUseUnicode Then C1DropDown.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25)
        C1DropDown.Font = FontUnicode(bUseUnicode, C1DropDown.Font.Style)
        For i As Integer = 0 To C1DropDown.DisplayColumns.Count - 1
            C1DropDown.DisplayColumns(i).Style.Font = FontUnicode(bUseUnicode, C1DropDown.Font.Style, C1DropDown.Font.Size)
        Next
        C1DropDown.Style.VerticalAlignment = C1.Win.C1TrueDBGrid.AlignVertEnum.Center
        'Bổ sung 26/07/2012
        '   UnicodeConvertFont(C1DropDown, bUseUnicode)

        '*****************
        Dim iMaxDropdownItems As Integer = 8

        If dt.Rows.Count < iMaxDropdownItems Then
            Dim dr As DataRow = Nothing
            For i As Integer = 0 To iMaxDropdownItems - dt.Rows.Count - 1
                dr = dt.NewRow
                dt.Rows.Add(dr)
            Next
        End If
        'Update 18/02/2014: Bỏ, đã đưa vào hàm SetResolutionControl
        'C1DropDown.Styles.Item("Selected").BackColor = Color.Green

        C1DropDown.SetDataBinding(dt, "", True)
    End Sub

    ''' <summary>
    ''' Đổ nguồn cho C1DropDown nhập Unicode
    ''' </summary>
    ''' <param name="C1DropDown">Tên C1DropDown cần đổ nguồn</param>
    ''' <param name="sSQL">Lệnh SQL đổ nguồn</param>
    ''' <param name="Width">Mảng chiều rộng các cột tương ứng của C1DropDown</param>
    <DebuggerStepThrough()> _
    Public Sub LoadDataSource(ByVal C1DropDown As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal sSQL As String, ByVal Width() As Integer, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable = ReturnDataTable(sSQL)
        LoadDataSource(C1DropDown, dt, Width, bUseUnicode)
    End Sub

    ''' <summary>
    ''' Đổ nguồn cho C1DropDown nhập Unicode
    ''' </summary>
    ''' <param name="C1DropDown">Tên C1DropDown cần đổ nguồn</param>
    ''' <param name="dt">DataTable nguồn</param>
    ''' <param name="Width">Mảng chiều rộng các cột tương ứng của C1DropDown</param>
    <DebuggerStepThrough()> _
    Public Sub LoadDataSource(ByVal C1DropDown As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dt As DataTable, ByVal Width() As Integer, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(C1DropDown, dt, bUseUnicode)
        For i As Integer = 0 To C1DropDown.DisplayColumns.Count - 1
            C1DropDown.DisplayColumns(i).Width = Width(i)
        Next
    End Sub
#End Region

#Region "Đổ nguồn cho C1Grid"

    ''' <summary>
    ''' Đổ nguồn cho C1Grid nhập Unicode
    ''' </summary>
    ''' <param name="C1Grid">Tên C1Grid cần đổ nguồn</param>
    ''' <param name="dt">DataTable nguồn</param>
    <DebuggerStepThrough()> _
    Public Sub LoadDataSource(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dt As DataTable, Optional ByVal bUseUnicode As Boolean = False)
        'Modify date: 17/03/2009: Set Font nhập Unicode
        '  If bUseUnicode Then C1Grid.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25)
        'If bUseUnicode Then
        '    C1Grid.Font = FontUnicode(bUseUnicode, C1Grid.Font.Style) 'New System.Drawing.Font("Microsoft Sans Serif", 8.25)
        '    For Each dc As C1.Win.C1TrueDBGrid.C1DataColumn In C1Grid.Columns
        '        For i As Integer = 0 To C1Grid.Splits.ColCount - 1
        '            C1Grid.Splits(i).DisplayColumns(dc.DataField).Style.Font = C1Grid.Font 'FontUnicode(bUseUnicode)
        '            C1Grid.Splits(i).DisplayColumns(dc.DataField).EditorStyle.Font = C1Grid.Font 'Bổ sung 27/05/2011
        '        Next
        '    Next
        'End If
        ' If bUseUnicode = False Then bUseUnicode = gbUnicode 'Fix 11/04/2014 lỗi khi không truyền bUseUnicode hoặc luôn luôn = True
        'Bổ sung 26/07/2012
        '   UnicodeConvertFont(C1Grid, bUseUnicode)
        '*************
        C1Grid.Font = FontUnicode(gbUnicode)
        For Each dc As C1.Win.C1TrueDBGrid.C1DataColumn In C1Grid.Columns
            For i As Integer = 0 To C1Grid.Splits.ColCount - 1
                C1Grid.Splits(i).DisplayColumns(dc.DataField).Style.Font = C1Grid.Font 'FontUnicode(bUseUnicode)
                C1Grid.Splits(i).DisplayColumns(dc.DataField).EditorStyle.Font = C1Grid.Font 'Bổ sung 27/05/2011
            Next
        Next
        C1Grid.SetDataBinding(dt, "", True)

    End Sub

    '''' <summary>
    '''' Đổ nguồn cho C1Grid nhập Unicode
    '''' </summary>
    '''' <param name="C1Grid">Tên C1Grid cần đổ nguồn</param>
    '''' <param name="sSQL">Lệnh SQL đổ nguồn</param>
    '<DebuggerStepThrough()> _
    Public Sub LoadDataSource(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sSQL As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim dt As DataTable = ReturnDataTable(sSQL)
        LoadDataSource(C1Grid, dt, bUseUnicode)
    End Sub
#End Region

End Module
