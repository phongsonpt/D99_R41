
'#######################################################################################
'#                                     CHÚ Ý (Các hàm chung (viết tiếp cho D99X0000))
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày tạo: 29/11/2010 
'# Ngày cập nhật cuối cùng:  21/02/2014
'# Người cập nhật cuối cùng: Minh Hòa
'# Diễn giải: 
'# Đưa hàm CheckConnection vào đây
'# Sửa hàm ClearText
'# Bỏ sự kiện c1dateDate_LostFocus
'# Bổ sung WITH (NOLOCK) vào table, trong bảng D91T0000 23/9/2013
'#Bổ sung GetTextCreateBy_new và LockColums, UnLockColums 4/11/2013
'# Bổ sung Tìm dòng trên lưới :findrowInGrid() 12/11/2013
'# Sửa lại kiểm tra Ngày hóa đơn lớn hơn Ngày phiếu : CheckInvoiceDateWithVoucherDate  20/11/2013
'# Bổ sung các hàm thông báo chung AskDelete: 10/01/2014
'# Bổ sung các hàm InputDateCustomFormat: 15/01/2014
'# Bổ sung các hàm Cho phép rớt dòng trên lưới HotKeyCtrlEnter và resetRowHeight: 18/02/2014
'# Sửa lại hàm CheckInvoiceDateWithVoucherDate  bổ sung kiểm tra ngày hóa đơn phù hợp với Kỳ hiện tại: 21/2/2014
'# Sửa lại hàm CheckInvoiceDateWithVoucherDate  bổ sung kiểm tra ngày hóa đơn nhỏ hơn ngày hiện tại: 09/5/2014

'#######################################################################################
Imports System.IO
Imports System.Threading
Imports System.Security.Cryptography
Imports System
Imports System.Runtime.InteropServices
Imports System.Runtime.CompilerServices

Public Module D99X0011

#Region "Khai báo biến toàn cục"
    ''' <summary>
    ''' Màu nền của Control bắt buộc nhập
    ''' </summary>
    ''' <remarks></remarks>
    Public COLOR_BACKCOLORWARNING As System.Drawing.Color = Color.Beige
    Public gbyD91RefDate As Integer = -1 ' Kiểm tra Ngày hóa đơn với Ngày chứng từ
    Public gbyD91InvoicesDateInPeriod As Integer = -1 'Kiểm tra ngày hóa đơn có phù hợp với kỳ hiện tại
    Public gbyD91InvoicesDateWithGetDate As Integer = -1 ' Kiểm tra Ngày hóa đơn nhỏ hơn Ngày hiện tại

    Public gdDatePeriodFrom As Date
    Public gdDatePeriodTo As Date
    Public gsFormatDateType As String = ""

#End Region

    '#Region "Lấy màu"
    '    ''' <summary>
    '    ''' Lấy màu bắt buộc nhập cho các control
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public Sub GetBackColorObligatory()
    '        COLOR_BACKCOLOROBLIGATORY = ColorTranslator.FromHtml("#" & D99C0007.GetModulesSetting("D00", ModuleOption.lmOptions, "ColorObligatory", "F5F5DC"))
    '        COLOR_BACKCOLORWARNING = ColorTranslator.FromHtml("#" & D99C0007.GetModulesSetting("D00", ModuleOption.lmOptions, "ColorWarning", "FF0000"))
    '    End Sub
    '#End Region

#Region "Nhập ngày trên lưới"

    Public Function GetFormatDateType() As String
        Dim FormatDateType As String = "dd/MM/yyyy"
        Dim sSQL As String = "Select FormatDateType From D09T0000 WITH(NOLOCK) "
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then FormatDateType = dt.Rows(0).Item("FormatDateType").ToString
        dt.Dispose()
        Return FormatDateType
    End Function

    Public Function GetValueD09T0000(ByVal dtD09T0000 As DataTable, ByVal sfield As String) As String
        If dtD09T0000 Is Nothing Then dtD09T0000 = ReturndtD09T0000(sfield)
        Dim sValue As String = ""
        If dtD09T0000.Rows.Count > 0 Then sValue = dtD09T0000.Rows(0).Item(sfield).ToString
        Return sValue
    End Function

    Public Function ReturndtD09T0000(ByVal sFields As String) As DataTable
        Return ReturnDataTable("Select " & sFields & " From D09T0000 WITH(NOLOCK) ")
    End Function

    Public Sub InputDateCustomFormat(ByVal ParamArray c1date() As C1.Win.C1Input.C1DateEdit)
        For i As Integer = 0 To c1date.Length - 1
            InputDateCustomFormat(c1date(i))
        Next
    End Sub

    Private Sub InputDateCustomFormat(ByVal sCustomFormat As String, ByRef c1date As C1.Win.C1Input.C1DateEdit)
        c1date.CustomFormat = sCustomFormat 'TH gắn trên lưới
        c1date.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        c1date.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        c1date.ValueIsDbNull = True
        c1date.EmptyAsNull = True
    End Sub

    'Dùng trên master
    Public Sub InputDateCustomFormat(ByRef c1date As C1.Win.C1Input.C1DateEdit)
        If gsFormatDateType <> "" Then c1date.CustomFormat = c1date.CustomFormat.Replace("dd/MM/yyyy", gsFormatDateType) 'TH có đủ ngày giờ
        Dim sValue As Object = c1date.Value
        c1date.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        c1date.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        c1date.ValueIsDbNull = True
        c1date.EmptyAsNull = True
        c1date.Value = sValue  'Khi thay đổi CustomFormat giá trị bị set ""
    End Sub

    Public Sub InputDateInTrueDBGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray COL_Date() As Integer)
        InputDateInTrueDBGrid_Private(tdbg, True, COL_Date)
    End Sub

    Private Sub InputDateInTrueDBGrid_Private(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, bUpdateAllColumns As Boolean, ByVal ParamArray COL_Date() As Integer)
        For i As Integer = 0 To COL_Date.Length - 1
            tdbg.Columns(COL_Date(i)).EditMask = ""
            tdbg.Columns(COL_Date(i)).EditMaskUpdate = False
            tdbg.Columns(COL_Date(i)).EnableDateTimeEditor = True
            Dim sCustomFormat As String = gsFormatDateType
            If sCustomFormat = "" Then sCustomFormat = "dd/MM/yyyy" 'TH những exe không thiết lập lại ngày thì gắn mặc định

            If tdbg.Columns(COL_Date(i)).NumberFormat = "" OrElse tdbg.Columns(COL_Date(i)).NumberFormat = "d" OrElse tdbg.Columns(COL_Date(i)).NumberFormat = "Short Date" Then tdbg.Columns(COL_Date(i)).NumberFormat = sCustomFormat
            tdbg.Columns(COL_Date(i)).NumberFormat = tdbg.Columns(COL_Date(i)).NumberFormat.Replace("dd/MM/yyyy", sCustomFormat) 'TH có đủ ngày giờ
            Dim c1dateTemp As New C1.Win.C1Input.C1DateEdit
            c1dateTemp.Visible = False
            c1dateTemp.CustomFormat = tdbg.Columns(COL_Date(i)).NumberFormat
            InputDateCustomFormat(c1dateTemp)
            c1dateTemp.TabStop = False
            SetMinMaxC1DateEdit(c1dateTemp) '29/10/2014 id 70210 - Kiểm tra ngày hợp lệ
            AddHandler c1dateTemp.LostFocus, AddressOf c1dateDate_LostFocus
            If bUpdateAllColumns Then
                AddHandler tdbg.AfterColUpdate, AddressOf tdbg_AfterColUpdate 'Bố sung vì lỗi chọn ngày không gắn được trên lưới 22/07/2017
            Else
                AddHandler tdbg.AfterColUpdate, AddressOf tdbg_AfterColUpdate_Private 'Bố sung vì lỗi chọn ngày không gắn được trên lưới 22/07/2017
            End If
            tdbg.Columns(COL_Date(i)).Editor = c1dateTemp
        Next
    End Sub


    Public Sub InputDateInTrueDBGridNew(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray COL_Date() As Integer)
        InputDateInTrueDBGrid_Private(tdbg, False, COL_Date)
    End Sub
    Public Sub tdbg_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs)
        CType(sender, C1.Win.C1TrueDBGrid.C1TrueDBGrid).UpdateData()
    End Sub

    Private Sub tdbg_AfterColUpdate_Private(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs)
        Dim tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid = CType(sender, C1.Win.C1TrueDBGrid.C1TrueDBGrid)
        If tdbg.Columns(e.ColIndex).Editor Is Nothing Then Exit Sub
        If TypeOf (tdbg.Columns(e.ColIndex).Editor) Is C1.Win.C1Input.C1DateEdit Then
            tdbg.UpdateData()
        End If
    End Sub

    Public Sub InputDateInTrueDBGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray COL_Date() As String)
        Dim arrInteger(COL_Date.Length - 1) As Integer
        For i As Integer = 0 To COL_Date.Length - 1
            arrInteger.SetValue(IndexOfColumn(tdbg, COL_Date(i)), i)
        Next
        InputDateInTrueDBGrid(tdbg, arrInteger)
    End Sub

    Public Sub InputDateInTrueDBGridNew(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray COL_Date() As String)
        Dim arrInteger(COL_Date.Length - 1) As Integer
        For i As Integer = 0 To COL_Date.Length - 1
            arrInteger.SetValue(IndexOfColumn(tdbg, COL_Date(i)), i)
        Next
        InputDateInTrueDBGridNew(tdbg, arrInteger)
    End Sub

    'Dùng Focus đúng các cột động dạng Ngày
    Private Sub c1dateDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim c1date As C1.Win.C1Input.C1DateEdit = CType(sender, C1.Win.C1Input.C1DateEdit)
        If c1date.Parent.GetType.Name = "C1TrueDBGrid" Then c1date.Parent.Focus() 'Khi nhập liệu trên lưới Enter không qua cột kế tiếp

        'CType(sender, C1.Win.C1Input.C1DateEdit).Parent.Focus() 'bị lỗi nhập liệu không có button chọn ngày
    End Sub

    Public Sub InputDateInFlexGrid(ByVal tdbg As C1.Win.C1FlexGrid.C1FlexGrid, ByVal ParamArray COL_Date() As String)
        If gsFormatDateType = "dd/MM/yyyy" Then Exit Sub
        For i As Integer = 0 To COL_Date.Length - 1
            tdbg.Cols(COL_Date(i)).Format = gsFormatDateType
        Next
    End Sub

    '29/10/2014 id 70210 - Kiểm tra ngày hợp lệ
    Public Sub SetMinMaxC1DateEdit(ByVal c1date As C1.Win.C1Input.C1DateEdit)
        c1date.Calendar.MinDate = SqlTypes.SqlDateTime.MinValue.Value ' dateMin 
        c1date.Calendar.MaxDate = SqlTypes.SqlDateTime.MaxValue.Value ' dateMax

        AddHandler c1date.ValueChanged, AddressOf c1dateDateFrom_ValueChanged
    End Sub

    Private Sub c1dateDateFrom_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim c1date As C1.Win.C1Input.C1DateEdit = CType(sender, C1.Win.C1Input.C1DateEdit)
        If SQLDateSave(c1date.Value) = "NULL" Then Exit Sub

        Dim dateMin As Date = SqlTypes.SqlDateTime.MinValue.Value ' #1/1/1753#
        Dim dateMax As Date = SqlTypes.SqlDateTime.MaxValue.Value ' #12/31/9999 11:59:59 PM#

        If DateDiff(DateInterval.Year, CDate(c1date.Value), dateMin) > 0 Or DateDiff(DateInterval.Year, CDate(c1date.Value), dateMax) < 0 Then
            '   c1date.Value = dateMin
            c1date.Value = DateAdd(DateInterval.Year, DateDiff(DateInterval.Year, CDate(c1date.Value), Now.Date), CDate(c1date.Value))
        End If
    End Sub

    Public Sub InputDateInTrueDBGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, bdevDate As Boolean, ByVal ParamArray COL_Date() As String)
        Dim arrInteger(COL_Date.Length - 1) As Integer
        For i As Integer = 0 To COL_Date.Length - 1
            arrInteger.SetValue(IndexOfColumn(tdbg, COL_Date(i)), i)
        Next
        If bdevDate Then
            InputDateInTrueDBGrid(tdbg, True, arrInteger)
        Else
            InputDateInTrueDBGrid(tdbg, arrInteger)
        End If
    End Sub

    Public Sub InputDateInTrueDBGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, bdevDate As Boolean, ByVal ParamArray COL_Date() As Integer)
        If bdevDate Then
            For i As Integer = 0 To COL_Date.Length - 1
                tdbg.Columns(COL_Date(i)).EditMask = ""
                tdbg.Columns(COL_Date(i)).EditMaskUpdate = False
                tdbg.Columns(COL_Date(i)).EnableDateTimeEditor = True
                Dim sCustomFormat As String = gsFormatDateType
                If sCustomFormat = "" Then sCustomFormat = "dd/MM/yyyy" 'TH những exe không thiết lập lại ngày thì gắn mặc định

                If tdbg.Columns(COL_Date(i)).NumberFormat = "" OrElse tdbg.Columns(COL_Date(i)).NumberFormat = "d" OrElse tdbg.Columns(COL_Date(i)).NumberFormat = "Short Date" Then tdbg.Columns(COL_Date(i)).NumberFormat = sCustomFormat
                tdbg.Columns(COL_Date(i)).NumberFormat = tdbg.Columns(COL_Date(i)).NumberFormat.Replace("dd/MM/yyyy", sCustomFormat) 'TH có đủ ngày giờ

                Dim x_original As Integer = 1024
                Dim y_original As Integer = 768
                Dim workingRectangle As System.Drawing.Rectangle = Screen.PrimaryScreen.Bounds 'Screen.PrimaryScreen.WorkingArea
                Dim tile_x As Double
                Dim tile_y As Double
                Dim iSizeFont_ As Single = giSizeTextD00 ' 8.25

                If iSizeFont <> SizeTextSmall Then
                    tile_x = workingRectangle.Width / x_original
                    tile_y = workingRectangle.Height / y_original
                    iSizeFont_ = giSizeTextD00 + CSng(tile_x)
                End If

                Dim devdateTemp As New DevExpress.XtraEditors.DateEdit
                devdateTemp.Font = New System.Drawing.Font("Microsoft Sans Serif", iSizeFont_)
                devdateTemp.Properties.AutoHeight = False
                devdateTemp.Height = tdbg.RowHeight '15
                devdateTemp.Visible = False

                With devdateTemp.Properties
                    .BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
                    .EditMask = tdbg.Columns(COL_Date(i)).NumberFormat
                    .Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTimeAdvancingCaret
                    .Mask.UseMaskAsDisplayFormat = True

                    .AllowFocused = False
                    .AllowNullInput = DevExpress.Utils.DefaultBoolean.True
                    .ButtonsStyle = DevExpress.XtraEditors.Controls.BorderStyles.Style3D

                    .Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                    .AppearanceCalendar.DayCellHighlighted.Font = New System.Drawing.Font(devdateTemp.Font.Name, 8.25, .AppearanceCalendar.DayCellHighlighted.Font.Style)
                    .AppearanceCalendar.DayCellHoliday.Font = New System.Drawing.Font(devdateTemp.Font.Name, iSizeFont_, .AppearanceCalendar.DayCellHoliday.Font.Style)
                    .AppearanceCalendar.DayCellSelected.Font = New System.Drawing.Font(devdateTemp.Font.Name, 8.25, .AppearanceCalendar.DayCellSelected.Font.Style)
                    .AppearanceCalendar.DayCellToday.Font = New System.Drawing.Font(devdateTemp.Font.Name, 8.25, .AppearanceCalendar.DayCellToday.Font.Style)
                    .AppearanceDisabled.Font = New System.Drawing.Font(devdateTemp.Font.Name, iSizeFont_, .AppearanceDisabled.Font.Style)
                    .AppearanceFocused.Font = New System.Drawing.Font(devdateTemp.Font.Name, 8.25, .AppearanceFocused.Font.Style)
                    .AppearanceReadOnly.Font = New System.Drawing.Font(devdateTemp.Font.Name, iSizeFont_, .AppearanceReadOnly.Font.Style)
                    .AppearanceFocused.Font = New System.Drawing.Font(devdateTemp.Font.Name, CSng(iSizeFont_ - 0.5), .AppearanceFocused.Font.Style)
                End With

                devdateTemp.TabStop = True

                tdbg.Columns(COL_Date(i)).Editor = devdateTemp

                AddHandler devdateTemp.DoubleClick, AddressOf devdateDate_DoubleClick
                AddHandler devdateTemp.FormatEditValue, AddressOf devdateDate_FormatEditValue
            Next
        Else
            Dim arrInteger(COL_Date.Length - 1) As Integer
            For i As Integer = 0 To COL_Date.Length - 1
                arrInteger.SetValue(COL_Date(i), i)
            Next
            InputDateInTrueDBGrid(tdbg, arrInteger)
        End If
    End Sub

    Private Sub devdateDate_FormatEditValue(sender As Object, e As DevExpress.XtraEditors.Controls.ConvertEditValueEventArgs)
        Dim devdate As DevExpress.XtraEditors.DateEdit = CType(sender, DevExpress.XtraEditors.DateEdit)
        If devdate.EditValue Is Nothing Then Exit Sub
        If devdate.EditValue.ToString() = "" Then Exit Sub
        devdate.DateTime = Convert.ToDateTime(devdate.EditValue)
    End Sub

    Private Sub devdateDate_DoubleClick(sender As Object, e As EventArgs)
        Dim devdate As DevExpress.XtraEditors.DateEdit = CType(sender, DevExpress.XtraEditors.DateEdit)
        If devdate.ReadOnly = False AndAlso (devdate.EditValue Is Nothing OrElse devdate.EditValue.ToString = "") Then devdate.EditValue = Now.Date
    End Sub

#Region "Định dạng Giờ trên master và lưới không theo thiết lập của hệ thống"
    ''' <summary>
    ''' Nhập Giờ không phụ thuộc giờ hệ thống, mặc định 00
    ''' </summary>
    ''' <param name="c1date"></param>
    ''' <remarks>CustomFormat = HH:mm</remarks>
    Public Sub InputTimeHHmm(ByVal ParamArray c1date() As C1.Win.C1Input.C1DateEdit)
        InputTimeCustomFormat("HH:mm", c1date)
    End Sub
    ''' <summary>
    ''' Nhập Giờ không phụ thuộc giờ hệ thống, mặc định 00
    ''' </summary>
    ''' <param name="CustomFormat">Định dạng do người dùng định nghĩa</param>
    ''' <param name="c1date"></param>
    ''' <remarks></remarks>
    Public Sub InputTimeCustomFormat(ByVal CustomFormat As String, ByVal ParamArray c1date() As C1.Win.C1Input.C1DateEdit)
        For i As Integer = 0 To c1date.Length - 1
            InputTimeCustomFormat(CustomFormat, c1date(i))
        Next
    End Sub

    Private Sub InputTimeCustomFormat(ByVal sCustomFormat As String, ByRef c1date As C1.Win.C1Input.C1DateEdit)
        InputDateCustomFormat(sCustomFormat, c1date)
        c1date.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None
        c1date.ParseInfo.DateTimeStyle = CType((((C1.Win.C1Input.DateTimeStyleFlags.AllowLeadingWhite Or C1.Win.C1Input.DateTimeStyleFlags.AllowTrailingWhite) _
                    Or C1.Win.C1Input.DateTimeStyleFlags.AllowInnerWhite) _
                    Or C1.Win.C1Input.DateTimeStyleFlags.NoCurrentDateDefault), C1.Win.C1Input.DateTimeStyleFlags)
    End Sub

    ''' <summary>
    ''' Nhập Giờ trên lưới không phụ thuộc giờ hệ thống, mặc định 00
    ''' </summary>
    ''' <param name="sCustomFormat">Định dạng do người dùng định nghĩa</param>
    ''' <param name="tdbg"></param>
    ''' <param name="COL_Time"></param>
    ''' <remarks></remarks>
    Public Sub InputTimeInTrueDBGrid(ByVal sCustomFormat As String, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray COL_Time() As Integer)
        For i As Integer = 0 To COL_Time.Length - 1
            tdbg.Columns(COL_Time(i)).EditMask = ""
            tdbg.Columns(COL_Time(i)).EditMaskUpdate = False
            tdbg.Columns(COL_Time(i)).EnableDateTimeEditor = True
            ' sCustomFormat = "HH:mm" 'TH những exe không thiết lập lại ngày thì gắn mặc định

            tdbg.Columns(COL_Time(i)).NumberFormat = sCustomFormat

            Dim c1dateTemp As New C1.Win.C1Input.C1DateEdit

            c1dateTemp.Visible = False
            c1dateTemp.CustomFormat = tdbg.Columns(COL_Time(i)).NumberFormat
            InputTimeCustomFormat(sCustomFormat, c1dateTemp)
            c1dateTemp.TabStop = False
            tdbg.Columns(COL_Time(i)).Editor = c1dateTemp
        Next
    End Sub

    Public Sub InputTimeInTrueDBGrid(ByVal sCustomFormat As String, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray COL_Time() As String)
        Dim arrInteger(COL_Time.Length - 1) As Integer
        For i As Integer = 0 To COL_Time.Length - 1
            arrInteger.SetValue(IndexOfColumn(tdbg, COL_Time(i)), i)
        Next
        InputTimeInTrueDBGrid(sCustomFormat, tdbg, arrInteger)
    End Sub
    ''' <summary>
    ''' Nhập Giờ trên lưới không phụ thuộc giờ hệ thống, mặc định 00
    ''' </summary>
    ''' <param name="tdbg"></param>
    ''' <param name="COL_Time"></param>
    ''' <remarks>CustomFormat = HH:mm</remarks>
    Public Sub InputTimeInTrueDBGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray COL_Time() As String)
        InputTimeInTrueDBGrid("HH:mm", tdbg, COL_Time)
    End Sub

    ''' <summary>
    ''' Nhập Giờ trên lưới không phụ thuộc giờ hệ thống, mặc định 00
    ''' </summary>
    ''' <param name="tdbg"></param>
    ''' <param name="COL_Time"></param>
    ''' <remarks>CustomFormat = HH:mm</remarks>
    Public Sub InputTimeInTrueDBGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray COL_Time() As Integer)
        InputTimeInTrueDBGrid("HH:mm", tdbg, COL_Time)
    End Sub
#End Region

#End Region

#Region "Cho phép rớt dòng trên cột của lưới"
    Public Sub HotKeyCtrlEnter(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal e As System.Windows.Forms.KeyEventArgs)
        tdbg.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.IndividualRows
        If e.KeyCode = Keys.Back Then 'Giảm RowHeight'Enter =2 ký tự
            If tdbg.SelectionStart > 1 AndAlso tdbg.Columns(tdbg.Col).Text.Substring(tdbg.SelectionStart - 2, 2) = vbCrLf Then
                For j As Integer = 0 To tdbg.Splits.Count - 1
                    tdbg.Splits(j).Rows(tdbg.Row).Height -= tdbg.RowHeight
                Next
            End If
        ElseIf e.Control And e.KeyCode = Keys.Enter Then
            For j As Integer = 0 To tdbg.Splits.Count - 1
                If tdbg.Splits(j).Rows(tdbg.Row).Height = -1 Then
                    tdbg.Splits(j).Rows(tdbg.Row).Height = tdbg.RowHeight * 2
                Else
                    tdbg.Splits(j).Rows(tdbg.Row).Height += tdbg.RowHeight
                End If
            Next
        End If
    End Sub

    Public Sub resetRowHeight(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray COL_Name() As String)
        tdbg.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.IndividualRows
        Dim tile_y As Double = (Screen.PrimaryScreen.Bounds.Height / 768) 'Tính tile_y =Screen.PrimaryScreen.Bounds.Height / 768 theo hàm SetResolutionForm
        For i As Integer = 0 To tdbg.RowCount - 1
            Dim iCountCtrlEnter As Integer = 0
            For iCol As Integer = 0 To COL_Name.Length - 1
                Dim newCount As Integer = System.Text.RegularExpressions.Regex.Matches(tdbg(i, COL_Name(iCol)).ToString, "\r\n").Count
                If iCountCtrlEnter = 0 Or iCountCtrlEnter < newCount Then iCountCtrlEnter = newCount
            Next
            iCountCtrlEnter += 1 'Số dòng
            For j As Integer = 0 To tdbg.Splits.Count - 1
                If GRID_ROW_HEIGHT > tdbg.RowHeight Then
                    tdbg.RowHeight = GRID_ROW_HEIGHT
                End If
                tdbg.Splits(j).Rows(i).Height = L3Int(tdbg.RowHeight * iCountCtrlEnter * tile_y)
            Next
        Next
    End Sub
#End Region

#Region "Kiểm tra ngày hóa đơn lớn hơn ngày phiếu"
    ''' <summary>
    ''' (Master) Kiểm tra ngày hóa đơn có phù hợp với kỳ hiện tại , Kiểm tra ngày chứng từ và Ngày hóa đơn
    ''' </summary>
    ''' <param name="c1VoucherDate">Required. Date voucher control.</param>
    ''' <param name="c1RefDate">Required. Date invoice control.</param>
    ''' <remarks>Return True : Valid</remarks>
    Public Function CheckInvoiceDateWithVoucherDate(ByVal sModuleID As String, ByVal c1VoucherDate As C1.Win.C1Input.C1DateEdit, ByVal c1RefDate As C1.Win.C1Input.C1DateEdit, ByVal sRefNo As String, ByVal sSerialNo As String, Optional ByVal tabSelection As System.Windows.Forms.TabControl = Nothing, Optional ByVal Index As Integer = -1) As Boolean

        'Ngày hóa đơn bằng rỗng thì không kiểm tra
        If c1RefDate.Text = "" Then Return True

        'Số số hóa đơn và số sêrial bằng rỗng thì không kiểm tra
        If sRefNo = "" And sSerialNo = "" Then Return True
        If gbyD91RefDate = -1 Then
            GetD91RefDateCheck(sModuleID) 'lấy thiết lập từ D91
        End If

        'Update 21/02/2014: incidnet 62702: kiểm tra ngày hóa đơn có phù hợp với kỳ hiện tại.
        If gbyD91InvoicesDateInPeriod > 0 Then
            'Kiểm tra vi phạm
            If CDate(SQLDateShow(c1RefDate.Value.ToString)) < CDate(SQLDateShow(gdDatePeriodFrom.Date)) Or CDate(SQLDateShow(c1RefDate.Value.ToString)) > CDate(SQLDateShow(gdDatePeriodTo.Date)) Then 'Vi phạm
                Select Case gbyD91InvoicesDateInPeriod
                    Case 1 ' Kiểm tra thông báo
                        'Ngày phiếu đã vi phạm kỳ hiện tại
                        If D99C0008.MsgAsk(rL3("MSG000054") & vbCrLf & rL3("MSG000021")) = Windows.Forms.DialogResult.No Then
                            If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                            c1RefDate.Focus()
                            Return False
                        End If
                    Case 2 ' Kiểm tra không cho lưu
                        D99C0008.MsgL3(rL3("MSG000054") & vbCrLf & rL3("MSG000053"), L3MessageBoxIcon.Exclamation)
                        If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                        c1RefDate.Focus()
                        Return False
                End Select
            End If
        End If

        '***********************************************
        'Update 09/05/2014: incidnet 60902: kiểm tra ngày hóa đơn phải nhỏ hơn Ngày hiện tại.
        If gbyD91InvoicesDateWithGetDate > 0 Then
            'Kiểm tra vi phạm
            If c1RefDate.Value < Now.Date Then 'Vi phạm
                Select Case gbyD91InvoicesDateWithGetDate
                    Case 1 ' Kiểm tra thông báo
                        'Ngày phiếu đã vi phạm kỳ hiện tại
                        If D99C0008.MsgAsk(rL3("MSG000055") & vbCrLf & rL3("MSG000021")) = Windows.Forms.DialogResult.No Then
                            If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                            c1RefDate.Focus()
                            Return False
                        End If
                    Case 2 ' Kiểm tra không cho lưu
                        D99C0008.MsgL3(rL3("MSG000055") & vbCrLf & rL3("MSG000053"), L3MessageBoxIcon.Exclamation)
                        If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                        c1RefDate.Focus()
                        Return False
                End Select
            End If
        End If

        '***********************************************
        'Ngày phiếu bằng rỗng thì không kiểm tra
        If c1VoucherDate.Text = "" Then Return True
        'Kiểm tra Ngày hóa đơn phải nhỏ hơn Ngày chứng từ
        If gbyD91RefDate = 0 Then Return True ' Không kiểm tra


        'Sửa lỗi cho TH nhập tiếp gán  thì ngày phiếu
        '  If Convert.ToDateTime(c1InvoiceDate.Value) <= Convert.ToDateTime(c1VoucherDate.Value) Then Return True
        If CDate(SQLDateShow(c1RefDate.Value.ToString)) <= CDate(SQLDateShow(c1VoucherDate.Value.ToString)) Then Return True
        Select Case gbyD91RefDate
            Case 1 ' Kiểm tra thông báo
                If D99C0008.MsgAsk(rL3("MSG000036") & vbCrLf & rL3("MSG000021")) = Windows.Forms.DialogResult.No Then
                    If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                    c1RefDate.Focus()
                    Return False
                End If
            Case 2 ' Kiểm tra không cho lưu
                D99C0008.MsgL3(rL3("MSG000036") & vbCrLf & rL3("MSG000053"), L3MessageBoxIcon.Exclamation)
                If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                c1RefDate.Focus()
                Return False
        End Select
        Return True
    End Function

    ''' <summary>
    ''' (Master) Kiểm tra ngày hóa đơn có phù hợp với kỳ hiện tại , Kiểm tra ngày chứng từ và Ngày hóa đơn
    ''' </summary>
    ''' <param name="devVoucherDate">Required. Date voucher control.</param>
    ''' <param name="devRefDate">Required. Date invoice control.</param>
    ''' <remarks>Return True : Valid</remarks>
    Public Function CheckInvoiceDateWithVoucherDate(ByVal sModuleID As String, ByVal devVoucherDate As DevExpress.XtraEditors.DateEdit, ByVal devRefDate As DevExpress.XtraEditors.DateEdit, ByVal sRefNo As String, ByVal sSerialNo As String, Optional ByVal tabSelection As System.Windows.Forms.TabControl = Nothing, Optional ByVal Index As Integer = -1) As Boolean

        'Ngày hóa đơn bằng rỗng thì không kiểm tra
        If devRefDate.Text = "" Then Return True

        'Số số hóa đơn và số sêrial bằng rỗng thì không kiểm tra
        If sRefNo = "" And sSerialNo = "" Then Return True
        If gbyD91RefDate = -1 Then
            GetD91RefDateCheck(sModuleID) 'lấy thiết lập từ D91
        End If

        'Update 21/02/2014: incidnet 62702: kiểm tra ngày hóa đơn có phù hợp với kỳ hiện tại.
        If gbyD91InvoicesDateInPeriod > 0 Then
            'Kiểm tra vi phạm
            If CDate(SQLDateShow(devRefDate.EditValue.ToString)) < CDate(SQLDateShow(gdDatePeriodFrom.Date)) Or CDate(SQLDateShow(devRefDate.EditValue.ToString)) > CDate(SQLDateShow(gdDatePeriodTo.Date)) Then 'Vi phạm
                Select Case gbyD91InvoicesDateInPeriod
                    Case 1 ' Kiểm tra thông báo
                        'Ngày phiếu đã vi phạm kỳ hiện tại
                        If D99C0008.MsgAsk(rL3("MSG000054") & vbCrLf & rL3("MSG000021")) = Windows.Forms.DialogResult.No Then
                            If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                            devRefDate.Focus()
                            Return False
                        End If
                    Case 2 ' Kiểm tra không cho lưu
                        D99C0008.MsgL3(rL3("MSG000054") & vbCrLf & rL3("MSG000053"), L3MessageBoxIcon.Exclamation)
                        If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                        devRefDate.Focus()
                        Return False
                End Select
            End If
        End If

        '***********************************************
        'Update 09/05/2014: incidnet 60902: kiểm tra ngày hóa đơn phải nhỏ hơn Ngày hiện tại.
        If gbyD91InvoicesDateWithGetDate > 0 Then
            'Kiểm tra vi phạm
            If devRefDate.EditValue < Now.Date Then 'Vi phạm
                Select Case gbyD91InvoicesDateWithGetDate
                    Case 1 ' Kiểm tra thông báo
                        'Ngày phiếu đã vi phạm kỳ hiện tại
                        If D99C0008.MsgAsk(rL3("MSG000055") & vbCrLf & rL3("MSG000021")) = Windows.Forms.DialogResult.No Then
                            If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                            devRefDate.Focus()
                            Return False
                        End If
                    Case 2 ' Kiểm tra không cho lưu
                        D99C0008.MsgL3(rL3("MSG000055") & vbCrLf & rL3("MSG000053"), L3MessageBoxIcon.Exclamation)
                        If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                        devRefDate.Focus()
                        Return False
                End Select
            End If
        End If

        '***********************************************
        'Ngày phiếu bằng rỗng thì không kiểm tra
        If devVoucherDate.Text = "" Then Return True
        'Kiểm tra Ngày hóa đơn phải nhỏ hơn Ngày chứng từ
        If gbyD91RefDate = 0 Then Return True ' Không kiểm tra


        'Sửa lỗi cho TH nhập tiếp gán  thì ngày phiếu
        '  If Convert.ToDateTime(c1InvoiceDate.Value) <= Convert.ToDateTime(c1VoucherDate.Value) Then Return True
        If CDate(SQLDateShow(devRefDate.EditValue.ToString)) <= CDate(SQLDateShow(devVoucherDate.EditValue.ToString)) Then Return True
        Select Case gbyD91RefDate
            Case 1 ' Kiểm tra thông báo
                If D99C0008.MsgAsk(rL3("MSG000036") & vbCrLf & rL3("MSG000021")) = Windows.Forms.DialogResult.No Then
                    If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                    devRefDate.Focus()
                    Return False
                End If
            Case 2 ' Kiểm tra không cho lưu
                D99C0008.MsgL3(rL3("MSG000036") & vbCrLf & rL3("MSG000053"), L3MessageBoxIcon.Exclamation)
                If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                devRefDate.Focus()
                Return False
        End Select
        Return True
    End Function

    ''' <summary>
    ''' (Detail) Kiểm tra ngày chứng từ và Ngày hóa đơn trên lưới (dataField column)
    ''' </summary>
    ''' <param name="c1VoucherDate">Required. Date control.</param>
    ''' <param name="tdbg">lưới</param>
    ''' <param name="COL_RefDate">DataField cột Ngày hóa đơn trên lưới. String</param>
    ''' <param name="iSplit">split chứa cột Ngày hóa đơn. Integer</param>
    ''' <remarks>Return True : Valid</remarks>
    Private Function CheckInvoiceDateWithVoucherDate(ByVal sModuleID As String, ByVal c1VoucherDate As C1.Win.C1Input.C1DateEdit, ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_RefDate As String, ByVal COL_RefNo As String, ByVal COL_SerialNo As String, Optional ByVal iSplit As Integer = 0) As Boolean ', ByRef bCheckOkDate As Boolean, ByRef iIndexError As Integer,
        If gbyD91RefDate = -1 Then
            GetD91RefDateCheck(sModuleID) '23/10/2013, Văn Tâm-Ngọc Thoại: id 79723- 	Kiểm tra ngày hóa đơn lớn hơn ngày phiếu
        End If

        Dim dtGrid As DataTable = CType(tdbg.DataSource, DataTable)
        Dim dr() As DataRow
        'Xóa những dòng vi phạm trước đó
        For Each row As DataRow In dtGrid.GetErrors()
            row.ClearErrors()
        Next

        'Update 21/02/2014: incidnet 62702: kiểm tra ngày hóa đơn có phù hợp với kỳ hiện tại.
        If gbyD91InvoicesDateInPeriod > 0 Then
            'Kiểm tra vi phạm
            Dim sWhere1 As String = ""
            sWhere1 = "(" & COL_RefDate & "< #" & DateSave(gdDatePeriodFrom) & "# " & " Or " & COL_RefDate & "> #" & DateSave(gdDatePeriodTo) & "#)"
            'Số số hóa đơn và số sêrial bằng rỗng thì không kiểm tra
            sWhere1 &= " AND (" & COL_SerialNo & " <> '' or " & COL_RefNo & " <> '')"
            dr = dtGrid.Select(sWhere1)
            If dr.Length > 0 Then 'Vi phạm
                Select Case gbyD91InvoicesDateInPeriod
                    Case 1 ' Kiểm tra thông báo
                        If D99C0008.MsgAsk(rL3("MSG000054") & vbCrLf & rL3("MSG000021")) = Windows.Forms.DialogResult.No Then
                            For i As Integer = 0 To dr.Length - 1
                                dr(i).RowError = "ErrorDate"
                            Next
                            tdbg.Focus()
                            tdbg.SplitIndex = iSplit
                            tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                            tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                            Return False
                        End If

                    Case 2 ' Kiểm tra không cho lưu
                        D99C0008.MsgL3(rL3("MSG000054") & vbCrLf & rL3("MSG000053"), L3MessageBoxIcon.Exclamation)
                        For i As Integer = 0 To dr.Length - 1
                            dr(i).RowError = "ErrorDate"
                        Next
                        tdbg.Focus()
                        tdbg.SplitIndex = iSplit
                        tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                        tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                        Return False
                End Select

            End If
        End If
        '***********************************************
        'Update 09/05/2014: incidnet 60902: kiểm tra ngày hóa đơn phải nhỏ hơn Ngày hiện tại.
        If gbyD91InvoicesDateWithGetDate > 0 Then
            'Kiểm tra vi phạm
            Dim sWhere1 As String = ""
            sWhere1 = "(" & COL_RefDate & "< #" & DateSave(Now.Date) & "#)"
            'Số số hóa đơn và số sêrial bằng rỗng thì không kiểm tra
            sWhere1 &= " AND (" & COL_SerialNo & " <> '' or " & COL_RefNo & " <> '')"
            dr = dtGrid.Select(sWhere1)
            If dr.Length > 0 Then 'Vi phạm
                Select Case gbyD91InvoicesDateWithGetDate
                    Case 1 ' Kiểm tra thông báo
                        If D99C0008.MsgAsk(rL3("MSG000055") & vbCrLf & rL3("MSG000021")) = Windows.Forms.DialogResult.No Then
                            For i As Integer = 0 To dr.Length - 1
                                dr(i).RowError = "ErrorDate"
                            Next
                            tdbg.Focus()
                            tdbg.SplitIndex = iSplit
                            tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                            tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                            Return False
                        End If

                    Case 2 ' Kiểm tra không cho lưu
                        D99C0008.MsgL3(rL3("MSG000055") & vbCrLf & rL3("MSG000053"), L3MessageBoxIcon.Exclamation)
                        For i As Integer = 0 To dr.Length - 1
                            dr(i).RowError = "ErrorDate"
                        Next
                        tdbg.Focus()
                        tdbg.SplitIndex = iSplit
                        tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                        tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                        Return False
                End Select

            End If
        End If


        '***********************************************
        'Ngày phiếu bằng rỗng thì không kiểm tra
        If c1VoucherDate.Text = "" Then Return True

        'Kiểm tra ngày chứng từ và Ngày hóa đơn
        If gbyD91RefDate = 0 Then Return True ' Không kiểm tra

        '*************
        Dim sWhere As String = ""
        sWhere = COL_RefDate & ">= #" & DateSave(CDate(c1VoucherDate.Value).AddDays(1)) & "#"
        'Số số hóa đơn và số sêrial bằng rỗng thì không kiểm tra
        sWhere &= " And (" & COL_RefNo & " <>'' Or " & COL_SerialNo & " <>'' )"

        dr = dtGrid.Select(sWhere)
        If dr.Length = 0 Then Return True
        Select Case gbyD91RefDate
            Case 1 ' Kiểm tra thông báo
                If D99C0008.MsgAsk(rL3("MSG000036") & vbCrLf & rL3("MSG000021")) = Windows.Forms.DialogResult.No Then
                    For i As Integer = 0 To dr.Length - 1
                        dr(i).RowError = "ErrorDate"
                    Next
                    tdbg.Focus()
                    tdbg.SplitIndex = iSplit
                    tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                    tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                    Return False
                End If

            Case 2 ' Kiểm tra không cho lưu
                D99C0008.MsgL3(rL3("MSG000036") & vbCrLf & rL3("MSG000053"), L3MessageBoxIcon.Exclamation)
                For i As Integer = 0 To dr.Length - 1
                    dr(i).RowError = "ErrorDate"
                Next
                tdbg.Focus()
                tdbg.SplitIndex = iSplit
                tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                Return False
        End Select
        Return True
    End Function

    ''' <summary>
    ''' (Detail) Kiểm tra ngày chứng từ và Ngày hóa đơn trên lưới (dataField column)
    ''' </summary>
    ''' <param name="devVoucherDate">Required. Date control.</param>
    ''' <param name="tdbg">lưới</param>
    ''' <param name="COL_RefDate">DataField cột Ngày hóa đơn trên lưới. String</param>
    ''' <param name="iSplit">split chứa cột Ngày hóa đơn. Integer</param>
    ''' <remarks>Return True : Valid</remarks>
    Private Function CheckInvoiceDateWithVoucherDate(ByVal sModuleID As String, ByVal devVoucherDate As DevExpress.XtraEditors.DateEdit, ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_RefDate As String, ByVal COL_RefNo As String, ByVal COL_SerialNo As String, Optional ByVal iSplit As Integer = 0) As Boolean ', ByRef bCheckOkDate As Boolean, ByRef iIndexError As Integer,
        If gbyD91RefDate = -1 Then
            GetD91RefDateCheck(sModuleID) '23/10/2013, Văn Tâm-Ngọc Thoại: id 79723- 	Kiểm tra ngày hóa đơn lớn hơn ngày phiếu
        End If

        Dim dtGrid As DataTable = CType(tdbg.DataSource, DataTable)
        Dim dr() As DataRow
        'Xóa những dòng vi phạm trước đó
        For Each row As DataRow In dtGrid.GetErrors()
            row.ClearErrors()
        Next

        'Update 21/02/2014: incidnet 62702: kiểm tra ngày hóa đơn có phù hợp với kỳ hiện tại.
        If gbyD91InvoicesDateInPeriod > 0 Then
            'Kiểm tra vi phạm
            Dim sWhere1 As String = ""
            sWhere1 = "(" & COL_RefDate & "< #" & DateSave(gdDatePeriodFrom) & "# " & " Or " & COL_RefDate & "> #" & DateSave(gdDatePeriodTo) & "#)"
            'Số số hóa đơn và số sêrial bằng rỗng thì không kiểm tra
            sWhere1 &= " AND (" & COL_SerialNo & " <> '' or " & COL_RefNo & " <> '')"
            dr = dtGrid.Select(sWhere1)
            If dr.Length > 0 Then 'Vi phạm
                Select Case gbyD91InvoicesDateInPeriod
                    Case 1 ' Kiểm tra thông báo
                        If D99C0008.MsgAsk(rL3("MSG000054") & vbCrLf & rL3("MSG000021")) = Windows.Forms.DialogResult.No Then
                            For i As Integer = 0 To dr.Length - 1
                                dr(i).RowError = "ErrorDate"
                            Next
                            tdbg.Focus()
                            tdbg.SplitIndex = iSplit
                            tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                            tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                            Return False
                        End If

                    Case 2 ' Kiểm tra không cho lưu
                        D99C0008.MsgL3(rL3("MSG000054") & vbCrLf & rL3("MSG000053"), L3MessageBoxIcon.Exclamation)
                        For i As Integer = 0 To dr.Length - 1
                            dr(i).RowError = "ErrorDate"
                        Next
                        tdbg.Focus()
                        tdbg.SplitIndex = iSplit
                        tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                        tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                        Return False
                End Select

            End If
        End If
        '***********************************************
        'Update 09/05/2014: incidnet 60902: kiểm tra ngày hóa đơn phải nhỏ hơn Ngày hiện tại.
        If gbyD91InvoicesDateWithGetDate > 0 Then
            'Kiểm tra vi phạm
            Dim sWhere1 As String = ""
            sWhere1 = "(" & COL_RefDate & "< #" & DateSave(Now.Date) & "#)"
            'Số số hóa đơn và số sêrial bằng rỗng thì không kiểm tra
            sWhere1 &= " AND (" & COL_SerialNo & " <> '' or " & COL_RefNo & " <> '')"
            dr = dtGrid.Select(sWhere1)
            If dr.Length > 0 Then 'Vi phạm
                Select Case gbyD91InvoicesDateWithGetDate
                    Case 1 ' Kiểm tra thông báo
                        If D99C0008.MsgAsk(rL3("MSG000055") & vbCrLf & rL3("MSG000021")) = Windows.Forms.DialogResult.No Then
                            For i As Integer = 0 To dr.Length - 1
                                dr(i).RowError = "ErrorDate"
                            Next
                            tdbg.Focus()
                            tdbg.SplitIndex = iSplit
                            tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                            tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                            Return False
                        End If

                    Case 2 ' Kiểm tra không cho lưu
                        D99C0008.MsgL3(rL3("MSG000055") & vbCrLf & rL3("MSG000053"), L3MessageBoxIcon.Exclamation)
                        For i As Integer = 0 To dr.Length - 1
                            dr(i).RowError = "ErrorDate"
                        Next
                        tdbg.Focus()
                        tdbg.SplitIndex = iSplit
                        tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                        tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                        Return False
                End Select

            End If
        End If


        '***********************************************
        'Ngày phiếu bằng rỗng thì không kiểm tra
        If devVoucherDate.Text = "" Then Return True

        'Kiểm tra ngày chứng từ và Ngày hóa đơn
        If gbyD91RefDate = 0 Then Return True ' Không kiểm tra

        '*************
        Dim sWhere As String = ""
        sWhere = COL_RefDate & ">= #" & DateSave(CDate(devVoucherDate.EditValue).AddDays(1)) & "#"
        'Số số hóa đơn và số sêrial bằng rỗng thì không kiểm tra
        sWhere &= " And (" & COL_RefNo & " <>'' Or " & COL_SerialNo & " <>'' )"

        dr = dtGrid.Select(sWhere)
        If dr.Length = 0 Then Return True
        Select Case gbyD91RefDate
            Case 1 ' Kiểm tra thông báo
                If D99C0008.MsgAsk(rL3("MSG000036") & vbCrLf & rL3("MSG000021")) = Windows.Forms.DialogResult.No Then
                    For i As Integer = 0 To dr.Length - 1
                        dr(i).RowError = "ErrorDate"
                    Next
                    tdbg.Focus()
                    tdbg.SplitIndex = iSplit
                    tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                    tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                    Return False
                End If

            Case 2 ' Kiểm tra không cho lưu
                D99C0008.MsgL3(rL3("MSG000036") & vbCrLf & rL3("MSG000053"), L3MessageBoxIcon.Exclamation)
                For i As Integer = 0 To dr.Length - 1
                    dr(i).RowError = "ErrorDate"
                Next
                tdbg.Focus()
                tdbg.SplitIndex = iSplit
                tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                Return False
        End Select
        Return True
    End Function

    ''' <summary>
    ''' Kiểm tra ngày chứng từ và Ngày hóa đơn trên lưới (index column)
    ''' </summary>
    ''' <param name="c1VoucherDate">Required. Date control.</param>
    ''' <param name="tdbg">lưới</param>
    ''' <param name="COL_InvoiceDate">index cột Ngày hóa đơn trên lưới. Integer</param>
    ''' <param name="iSplit">split chứa cột Ngày hóa đơn. Integer</param>
    ''' <remarks>Return ></remarks>
    Public Function CheckInvoiceDateWithVoucherDate(ByVal sModuleID As String, ByVal c1VoucherDate As C1.Win.C1Input.C1DateEdit, ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_InvoiceDate As Integer, ByVal COL_RefNo As Integer, ByVal COL_SerialNo As Integer, Optional ByVal iSplit As Integer = 0) As Boolean
        Return CheckInvoiceDateWithVoucherDate(sModuleID, c1VoucherDate, tdbg, tdbg.Columns(COL_InvoiceDate).DataField, tdbg.Columns(COL_RefNo).DataField, tdbg.Columns(COL_SerialNo).DataField, iSplit)
    End Function

    ''' <summary>
    ''' Kiểm tra ngày chứng từ và Ngày hóa đơn trên lưới (index column)
    ''' </summary>
    ''' <param name="devVoucherDate">Required. Date control.</param>
    ''' <param name="tdbg">lưới</param>
    ''' <param name="COL_InvoiceDate">index cột Ngày hóa đơn trên lưới. Integer</param>
    ''' <param name="iSplit">split chứa cột Ngày hóa đơn. Integer</param>
    ''' <remarks>Return ></remarks>
    Public Function CheckInvoiceDateWithVoucherDate(ByVal sModuleID As String, ByVal devVoucherDate As DevExpress.XtraEditors.DateEdit, ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_InvoiceDate As Integer, ByVal COL_RefNo As Integer, ByVal COL_SerialNo As Integer, Optional ByVal iSplit As Integer = 0) As Boolean
        Return CheckInvoiceDateWithVoucherDate(sModuleID, devVoucherDate, tdbg, tdbg.Columns(COL_InvoiceDate).DataField, tdbg.Columns(COL_RefNo).DataField, tdbg.Columns(COL_SerialNo).DataField, iSplit)
    End Function

    'Lấy giá trị thiết lập từ D91
    Public Sub GetD91RefDateCheck(ByVal sModule As String)
        'Return CByte(ReturnScalar("SELECT Top 1 RefDateCheck FROM D91T9102 WITH(NOLOCK) "))
        If sModule.Length = 2 Then sModule = "D" & sModule
        'gbyD91RefDate = CByte(ReturnScalar("Exec D91P9301 " & SQLString(sModule)))
        Dim sSQL As String
        sSQL = "-- Kiem tra Ngay hoa don va ngay phieu" & vbCrLf
        sSQL &= " Exec D91P9301 " & SQLString(sModule)

        Dim dtCheck As DataTable = ReturnDataTable(sSQL)
        gbyD91RefDate = L3Byte(dtCheck.Rows(0).Item("RefDateCheck"))
        gbyD91InvoicesDateInPeriod = L3Byte(dtCheck.Rows(0).Item("InvoicesDateInPeriod"))
        gbyD91InvoicesDateWithGetDate = L3Byte(dtCheck.Rows(0).Item("InvoiceDateWithGetDate"))

    End Sub

#End Region

#Region "Kiểm tra số sêri trùng số hóa đơn"
     Public Function CheckSerialAndRefNo(ByVal sTableName As String, _
                       ByVal sBatchID As String, _
                       ByVal sRefNo As String, _
                       ByVal sSerialNo As String, _
                       ByVal sVoucherID As String, ByVal sModuleID As String) As Boolean
        Return CheckSerialAndRefNo(sTableName, sBatchID, sRefNo, sSerialNo, sVoucherID, sModuleID, "", "", "")
    End Function

    Public Function CheckSerialAndRefNo(ByVal sTableName As String, _
                      ByVal sBatchID As String, _
                      ByVal sRefNo As String, _
                      ByVal sSerialNo As String, _
                      ByVal sVoucherID As String, ByVal sModuleID As String, sObjectTypeID As String, sObjectID As String, sVATNo As String) As Boolean
        Return CheckSerialAndRefNo(sTableName, sBatchID, sRefNo, sSerialNo, sVoucherID, sModuleID, sObjectTypeID, sObjectID, sVATNo, "")
    End Function

    'bổ sung thêm 31/07/2019-id 122373
    Public Function CheckSerialAndRefNo(ByVal sTableName As String, _
                      ByVal sBatchID As String, _
                      ByVal sRefNo As String, _
                      ByVal sSerialNo As String, _
                      ByVal sVoucherID As String, ByVal sModuleID As String, sObjectTypeID As String, sObjectID As String, sVATNo As String, sTranTypeID As String) As Boolean
        Return CheckSerialAndRefNo(sTableName, sBatchID, sRefNo, sSerialNo, sVoucherID, sModuleID, sObjectTypeID, sObjectID, sVATNo, sTranTypeID, "", "")
    End Function

    Public Function CheckSerialAndRefNo(ByVal sTableName As String, _
                    ByVal sBatchID As String, _
                    ByVal sRefNo As String, _
                    ByVal sSerialNo As String, _
                    ByVal sVoucherID As String, ByVal sModuleID As String, sObjectTypeID As String, sObjectID As String, sVATNo As String, sTranTypeID As String, sVATObjectTypeID As String, sVATObjectID As String) As Boolean
        Dim sSQL As String = SQLFinance.SQLStoreD91P9103(sTableName, sBatchID, sRefNo, sSerialNo, sVoucherID, sModuleID, sObjectTypeID, sObjectID, sVATNo, sTranTypeID, sVATObjectTypeID, sVATObjectID) 'bổ sung thêm 31/07/2019-id 122373
        Return CheckSerialAndRefNo(sSQL)
    End Function

    Public Function CheckSerialAndRefNo(ByVal sSQL As String) As Boolean

        Dim dt As New DataTable

        dt = ReturnDataTable(sSQL)

        If dt.Rows.Count > 0 Then
            Select Case CInt(dt.Rows(0).Item("Status"))
                Case 0
                Case 1
                    'If MessageBox.Show(dt.Rows(0).Item("Message").ToString, MsgAnnouncement, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.No Then
                    If D99C0008.Msg(ConvertVietwareFToUnicode(dt.Rows(0).Item("Message").ToString), MsgAnnouncement, L3MessageBoxButtons.YesNo, L3MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.No Then
                        dt = Nothing
                        Return True
                    End If
                Case 2
                    'MessageBox.Show(dt.Rows(0).Item("Message").ToString, MsgAnnouncement, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    D99C0008.MsgL3(ConvertVietwareFToUnicode(dt.Rows(0).Item("Message").ToString), L3MessageBoxIcon.Exclamation)
                    dt = Nothing
                    Return True
            End Select
        End If
        dt = Nothing
        Return False

        'Return Not CheckStoreNew(sSQL)

    End Function
#End Region

#Region "Đọc tiền bằng chữ"

    Public Function ConvertNumberToString(ByVal sNumber As Object, Optional ByVal CurrencyID As String = "VND", Optional ByVal bEnglish As Boolean = False) As String
        If bEnglish Then ' Đọc tiền bằng tiếng Anh
            Return ConvertNumberToString_English(sNumber, CurrencyID)
        End If

        'Đọc tiền bằng tiếng Việt (dạng Unicode)
        Try
            Dim nLen As Integer
            Dim i As Integer
            Dim j As Integer
            Dim nDigit As Integer
            Dim sTemp As String
            Dim sNumText() As String
            Dim sSeparatorNumber() As String 'Lưu 2 giá trị số phần nguyên và phần lẻ
            Dim sSeparatorValue(1) As String 'Lưu 2 giá trị chuỗi phần nguyên và phần lẻ

            Dim sUnit As String
            Dim sSubUnit As String
            Dim sSeparator As String
            Dim sStartSymbol As String
            Dim sEndSymbol As String
            Dim sResult As String = ""

            If Not IsNumeric(sNumber) Then Return "Số không hợp lệ !"
            If Math.Abs(CDbl(sNumber)) >= 1.0E+15 Then Return "Số quá lớn !"

            Dim NumString As String = sNumber.ToString  '= Format(Math.Abs(WhatNumber), "##############0.00")

            sUnit = ""
            sSubUnit = ""
            If CurrencyID = "Percent" Then
                sSeparator = "phẩy"
                sEndSymbol = "phần trăm"
            Else
                sSeparator = "và"
                sEndSymbol = "chẵn"
            End If

            'Lấy tên đơn vị tiền và đơn vị tiền lẻ
            If CurrencyID <> "Percent" Then GetCurrencyUnitName(CurrencyID, sUnit, sSubUnit)

            Dim bHundred As Boolean = False 'Có hàng trăm hay không
            Dim bUnit As Boolean = False 'Có hàng chục và đơn vị hay không
            Dim iCount As Integer 'Đếm số lần tăng hàng chục và đơn vị

            If Val(NumString) >= 0 Then
                sStartSymbol = ""
            Else
                sStartSymbol = "Âm"
            End If
            sNumText = Split("không;một;hai;ba;bốn;năm;sáu;bảy;tám;chín", ";")

            NumString = Format(Math.Abs(CDbl(sNumber)), "##############0.00")

            'Nếu tiền là VND thì bỏ số lẻ, không đọc phần xu
            If CurrencyID = "VND" Then
                NumString = Math.Round(CDbl(NumString)).ToString
            End If

            sSeparatorNumber = Split(NumString, ".") 'Tách phần nguyên và phần thập phân

            For j = 0 To sSeparatorNumber.Length - 1

                sTemp = ""

                NumString = sSeparatorNumber(j)

                nLen = Len(NumString)

                For i = 1 To nLen

                    nDigit = CInt(Mid(NumString, i, 1))

                    If bHundred Then
                        iCount += 1
                        If nDigit <> 0 Then
                            bUnit = True
                        End If
                    End If
                    sTemp &= " " & sNumText(nDigit)
                    If (nLen = i) Then Exit For

                    Select Case (nLen - i) Mod 9
                        Case 0
                            'sTemp &= " tỷ,"
                            sTemp &= " tỷ"
                            If Mid(NumString, i + 1, 3) = "000" Then
                                If nLen - i > 9 Then
                                    sTemp = Mid(sTemp, 1, sTemp.Length - 1) 'Cắt dấu phầy cuối
                                    'sTemp &= " tỷ,"
                                    sTemp &= " tỷ"
                                End If
                                i = i + 3
                            End If
                            If Mid(NumString, i + 1, 3) = "000" Then i = i + 3
                            If Mid(NumString, i + 1, 3) = "000" Then i = i + 3
                        Case 6
                            'sTemp &= " triệu,"
                            sTemp &= " triệu"
                            If Mid(NumString, i + 1, 3) = "000" Then
                                If nLen - i > 9 Then
                                    sTemp = Mid(sTemp, 1, sTemp.Length - 1) 'Cắt dấu phầy cuối
                                    'sTemp &= " tỷ,"
                                    sTemp &= " tỷ"
                                End If
                                i = i + 3
                            End If
                            If Mid(NumString, i + 1, 3) = "000" Then i = i + 3
                        Case 3
                            'sTemp &= " ngàn,"
                            sTemp &= " ngàn "
                            If Mid(NumString, i + 1, 3) = "000" Then
                                If nLen - i > 9 Then
                                    sTemp = Mid(sTemp, 1, sTemp.Length - 1) 'Cắt dấu phầy cuối
                                    'sTemp &= " tỷ,"
                                    sTemp &= " tỷ"
                                End If
                                i = i + 3
                            End If
                        Case Else
                            Select Case (nLen - i) Mod 3
                                Case 2
                                    bHundred = True
                                    sTemp &= " trăm"
                                Case 1
                                    sTemp &= " mươi"
                            End Select
                    End Select
                    If iCount = 2 And bUnit = False Then
                        iCount = 0
                        bHundred = False
                        bUnit = False
                        Dim sValues() As String = Split(sTemp, ",")
                        If sValues.Length > 1 AndAlso sValues(1).Contains("tỷ") Then
                            'sTemp = sValues(0) & " tỷ,"
                            sTemp = sValues(0) & " tỷ"
                        Else
                            'sTemp = sValues(0) & ","
                            sTemp = sValues(0) '& ","
                        End If
                    End If
                Next i

                sTemp = Replace(sTemp, "không mươi không", "")

                sTemp = Replace(sTemp, "không mươi", "lẻ")

                sTemp = Replace(sTemp, "mươi không", "mươi")

                sTemp = Replace(sTemp, "một mươi", "mười")

                'sTemp = Replace(sTemp, "mươi bốn", "mươi tư")

                'sTemp = Replace(sTemp, "lẻ bốn", "lẻ tư")

                sTemp = Replace(sTemp, "mươi năm", "mươi lăm")

                sTemp = Replace(sTemp, "mươi một", "mươi mốt")

                sTemp = Replace(sTemp, "mười năm", "mười lăm")

                sTemp = Trim(sTemp)

                If sTemp <> "" Then If Mid(sTemp, sTemp.Length, 1) = "," Then sTemp = Mid(sTemp, 1, sTemp.Length - 1)

                sSeparatorValue(j) = sTemp

            Next j
            Select Case sSeparatorValue(1)
                Case Nothing, "không", "không trăm", "không ngàn", "không triệu", "không tỷ"
                    sSeparatorValue(1) = ""
            End Select
            If sSeparatorValue(1) Is Nothing Then sSeparatorValue(1) = ""

            If sSeparatorNumber.Length - 1 > 0 And sSeparatorValue(1).Length > 0 Then
                If sStartSymbol = "" Then
                    sResult = UCase(Mid(sSeparatorValue(0), 1, 1)) & Mid(sSeparatorValue(0), 2) & IIf(sUnit = "", "", " ").ToString & sUnit & " " & sSeparator & " " & Mid(sSeparatorValue(1), 1, 1) & Mid(sSeparatorValue(1), 2) & IIf(sSubUnit = "", "", " ").ToString & sSubUnit
                Else
                    sResult = sStartSymbol & " " & Mid(sSeparatorValue(0), 1, 1) & Mid(sSeparatorValue(0), 2) & " " & IIf(sUnit = "", "", " ").ToString & " " & sSeparator & " " & Mid(sSeparatorValue(1), 1, 1) & Mid(sSeparatorValue(1), 2) & IIf(sSubUnit = "", "", " ").ToString & sSubUnit
                End If
                If CurrencyID = "Percent" Then
                    sResult &= " " & sEndSymbol
                End If
            Else
                If sStartSymbol = "" Then
                    sResult = UCase(Mid(sSeparatorValue(0), 1, 1)) & Mid(sSeparatorValue(0), 2) & IIf(sUnit = "", "", " ").ToString & sUnit & " " & sEndSymbol
                Else
                    sResult = sStartSymbol & " " & Mid(sSeparatorValue(0), 1, 1) & Mid(sSeparatorValue(0), 2) & IIf(sUnit = "", "", " ").ToString & sUnit & " " & sEndSymbol
                End If
            End If

            Return sResult.Replace("  ", " ")
        Catch ex As Exception
            D99C0008.MsgL3("Lỗi đọc tiền", L3MessageBoxIcon.Err)
            Return ""
        End Try

    End Function

    Private Function ConvertNumberToString_English(ByVal sNumber As Object, Optional ByVal CurrencyID As String = "VND") As String
        ' Đọc tiền bằng tiếng Anh
        Dim ToRead, NumString, Group, Word As String
        Dim WhatNumber As Double = CDbl(sNumber)

        If Not IsNumeric(sNumber) Then Return "Number is invalid !"
        If Math.Abs(CDbl(sNumber)) >= 1.0E+15 Then Return "Too long number !"

        If WhatNumber = 0 Then
            ToRead = "None"
        Else
            Dim i As Byte, W, X, Y, Z As Integer
            Dim FristColum() As String = {"None", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", _
                    "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eightteen", "nineteen"}
            Dim SecondColum() As String = {"None", "None", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"}

            Dim sUnit As String = ""
            Dim sSubUnit As String = ""
            'Lấy tên đơn vị tiền và đơn vị tiền lẻ
            GetCurrencyUnitName(CurrencyID, sUnit, sSubUnit)

            Dim ReadMetho() As String = {"None", "trillion", "billion", "million", "thousand", sUnit, sSubUnit}

            If WhatNumber < 0 Then
                ToRead = "Minus" & Space(1)
            Else
                ToRead = Space(0)
            End If
            'Nếu tiền là VND thì bỏ số lẻ, không đọc phần xu
            If CurrencyID = "VND" Then
                NumString = Format(Math.Round(Math.Abs(WhatNumber)), "##############0.00")
            Else
                NumString = Format(Math.Abs(WhatNumber), "##############0.00")
            End If

            NumString = Right(Space(15) & NumString, 18)
            For i = 1 To 6
                Group = Mid(NumString, i * 3 - 2, 3)
                If Group <> Space(3) Then
                    Select Case Group
                        Case "000"
                            If i = 5 And Math.Abs(WhatNumber) > 1 Then
                                'Word = "Vietnamese dong" & Space(1)
                                Word = ReadMetho(i) & Space(1)
                            Else
                                Word = Space(0)
                            End If
                        Case ".00"
                            Word = "only"
                        Case Else
                            X = L3Int(Left(Group, 1))
                            Y = L3Int(Mid(Group, 2, 1))
                            Z = L3Int(Right(Group, 1))
                            W = L3Int(Right(Group, 2))
                            If X = 0 Then
                                Word = Space(0)
                            Else
                                Word = FristColum(X) & Space(1) & "hundred" & Space(1)
                                If W > 0 And W < 21 Then
                                    Word = Word & "and" & Space(1)
                                End If
                            End If
                            If i = 6 And Math.Abs(WhatNumber) > 1 Then
                                Word = "and" & Space(1) & Word
                            End If
                            If W < 20 And W > 0 Then
                                Word = Word & FristColum(W) & Space(1)
                            Else
                                If W >= 20 Then
                                    Word = Word & SecondColum(Y) & Space(1)
                                    If Z > 0 Then
                                        Word = Word & FristColum(Z) & Space(1)
                                    End If
                                End If
                            End If
                            Word = Word & ReadMetho(i) & Space(1)
                    End Select
                    ToRead = ToRead & Word
                End If
            Next i
        End If

        Dim sResult As String = ""
        sResult = UCase(Left(ToRead, 1)) & Mid(ToRead, 2)

        Return sResult
    End Function

    Private Sub GetCurrencyUnitName(ByVal CurrencyID As String, ByRef UnitName As String, ByRef SubUnitName As String)
        ' Lấy Tên tiền và đơn vị tiền
        Dim sSQL As String
        sSQL = "Select UnitNameU, SubUnitNameU From D91T0010 WITH(NOLOCK)  Where (CurrencyID = " & SQLString(CurrencyID) & ")  And (Disabled = 0)"
        Dim dt As DataTable
        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            UnitName = dt.Rows(0).Item("UnitNameU").ToString
            SubUnitName = dt.Rows(0).Item("SubUnitNameU").ToString
        End If
        dt.Dispose()

    End Sub

#End Region

#Region "Xóa text các control khi nhấn Nhập tiếp"

    '''' <summary>
    '''' Xóa tất cả các control
    '''' </summary>
    '''' <param name="ctrl">truyền vào Me</param>
    '''' <remarks>Cách gọi: tại nút Nhập tiếp, gọi ClearText (Me)</remarks>
    '<DebuggerStepThrough()> _
    Public Sub ClearText(ByVal ctrl As Control)
        If TypeOf (ctrl) Is C1.Win.C1Input.C1DateEdit Then
            CType(ctrl, C1.Win.C1Input.C1DateEdit).Value = ""
        ElseIf TypeOf (ctrl) Is DevExpress.XtraEditors.DateEdit Then
            CType(ctrl, DevExpress.XtraEditors.DateEdit).EditValue = ""
        ElseIf TypeOf (ctrl) Is C1.Win.C1Input.C1NumericEdit Then
            CType(ctrl, C1.Win.C1Input.C1NumericEdit).Value = ""
        ElseIf (TypeOf (ctrl) Is CheckBox) Then
            CType(ctrl, CheckBox).Checked = False
        ElseIf (TypeOf (ctrl) Is TextBox Or TypeOf (ctrl) Is C1.Win.C1List.C1Combo) Then
            'Update 15/08/2013: Filter Bar trên lưới có TypeOf là TextBox nên chạy vào điều kiện này
            'ctrl.Text = ""
            If ctrl.Name <> "" Then ctrl.Text = ""
        End If

        For Each childControl As Control In ctrl.Controls
            ClearText(childControl)
        Next
    End Sub


    ''' <summary>
    ''' Xóa tất cả các control trừ các control trong tập truyền vào ctrlExclude
    ''' </summary>
    ''' <param name="ctrl">truyền vào Me</param>
    ''' <param name="ctrlExclude">truyền vào tập các control (txtName, tdbcCurrencyID, c1dateVoucherDate) không gán text =""</param>
    ''' <remarks>Cách gọi: tại nút Nhập tiếp, gọi ClearText (Me, txtName, tdbcCurrencyID, c1dateVoucherDate)</remarks>
    Public Sub ClearText(ByVal ctrl As Control, ByVal ParamArray ctrlExclude() As Control)
        If TypeOf (ctrl) Is C1.Win.C1Input.C1DateEdit Then
            'If ctrl.Visible = True Then CType(ctrl, C1.Win.C1Input.C1DateEdit).Value = ""
            CType(ctrl, C1.Win.C1Input.C1DateEdit).Value = ""
        ElseIf TypeOf (ctrl) Is DevExpress.XtraEditors.DateEdit Then
            CType(ctrl, DevExpress.XtraEditors.DateEdit).EditValue = ""
        ElseIf TypeOf (ctrl) Is C1.Win.C1Input.C1NumericEdit Then
            CType(ctrl, C1.Win.C1Input.C1NumericEdit).Value = ""
        ElseIf (TypeOf (ctrl) Is CheckBox) Then
            CType(ctrl, CheckBox).Checked = False
        ElseIf (TypeOf (ctrl) Is TextBox Or TypeOf (ctrl) Is C1.Win.C1List.C1Combo) Then
            ctrl.Text = ""
        End If

        For Each childControl As Control In ctrl.Controls
            If (TypeOf (childControl) Is TextBox) OrElse (TypeOf (childControl) Is C1.Win.C1List.C1Combo) _
            OrElse (TypeOf (childControl) Is DevExpress.XtraEditors.DateEdit) _
            OrElse (TypeOf (childControl) Is C1.Win.C1Input.C1DateEdit) OrElse (TypeOf (childControl) Is CheckBox) Then
                If FindControl(childControl, ctrlExclude) = False Then
                    ClearText(childControl, ctrlExclude)
                End If
            Else
                ClearText(childControl, ctrlExclude)
            End If
        Next
    End Sub


    Dim sValue As String = ""
    Private Function ContainsValue(ByVal s As Control) As Boolean
        Return s.Name.Equals(sValue)
    End Function

    Private Function FindControl(ByVal ctrl As Control, ByVal ParamArray ArrString() As Control) As Boolean
        sValue = ctrl.Name
        If Array.Exists(ArrString, AddressOf ContainsValue) Then
            Return True
        End If
        Return False
    End Function
#End Region

#Region "Enalbed TabPages"
    ''' <summary>
    ''' Set thuộc tính Enable = False của Tab page
    ''' </summary>
    ''' <param name="tabPage">tập các tab cần set Enabled = False</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub EnabledTabPage(ByVal tabPage() As TabPage, Optional ByVal bEnabled As Boolean = False)
        Dim tabMain As TabControl = CType(tabPage(0).Parent, TabControl)
        For i As Integer = 0 To tabPage.Length - 1
            tabPage(i).Enabled = bEnabled
        Next
        EnabledTabPage(tabMain)
    End Sub


    Public Sub EnabledTabPage(ByVal tabMain As TabControl)
        If tabMain.SelectedTab.Enabled = False Then
            For i As Integer = 0 To tabMain.TabPages.Count - 1
                If tabMain.TabPages(i).Enabled Then tabMain.SelectedTab = tabMain.TabPages(i) : Exit For
            Next
        End If

        '        tabMain.Padding = New Point(10, 3)

        '        If tabMain.DrawMode = TabDrawMode.OwnerDrawFixed Then GoTo refresh
        '        tabMain.DrawMode = TabDrawMode.OwnerDrawFixed
        '        AddHandler tabMain.Selecting, AddressOf tabMain_Selecting
        '        AddHandler tabMain.DrawItem, AddressOf OnDrawItem
        'refresh:
        tabMain.TopLevelControl.Refresh()
    End Sub


    Public Sub SetBoldTabPage(ByVal tabMain As TabControl)
        tabMain.Padding = New Point(15, 3)
        If SCT_BackColor <> "" Then
            For Each page As TabPage In tabMain.TabPages
                page.BackColor = ColorTranslator.FromHtml(SCT_BackColor)
            Next
        End If
        If tabMain.DrawMode = TabDrawMode.OwnerDrawFixed Then GoTo refresh
        tabMain.DrawMode = TabDrawMode.OwnerDrawFixed
        AddHandler tabMain.Selecting, AddressOf tabMain_Selecting
        AddHandler tabMain.DrawItem, AddressOf OnDrawItem
refresh:
        tabMain.TopLevelControl.Refresh()
    End Sub


    Private Sub tabMain_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs)
        Try
            If e.TabPage Is Nothing Then e.Cancel = True : Exit Sub
            If e.TabPage.Enabled = False Then
                e.Cancel = True
            Else
                e.Cancel = False
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub OnDrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)

        Dim tabMain As TabControl = CType(sender, TabControl)
        If tabMain.TabPages.Count = 0 Then Exit Sub
        ' Create pen.
        Dim blackPen As New Pen(tabMain.TabPages(0).BackColor, 3)
        'Get Location tabpage
        Dim myTabRect As Rectangle
        myTabRect = tabMain.GetTabRect(e.Index)
        ' Create coordinates of points that define line.
        Dim x1 As Integer = myTabRect.X
        Dim y1 As Integer = myTabRect.Bottom
        Dim x2 As Integer = myTabRect.X + myTabRect.Width
        ' Draw line to screen.
        'e.Graphics.DrawLine(blackPen, x1, y1, x2, y1)
        '**************
        ' Set format of string.
        Dim drawFormat As New StringFormat
        drawFormat.Alignment = StringAlignment.Center
        drawFormat.LineAlignment = StringAlignment.Center
        If SCT_BackColor <> "" Then
            e.Graphics.FillRectangle(New SolidBrush(ColorTranslator.FromHtml(SCT_BackColor)), New Rectangle(myTabRect.X, myTabRect.Y - 2, myTabRect.Width, myTabRect.Height + 4))
            e.Graphics.DrawRectangle(New Pen(ColorTranslator.FromHtml(SCT_BackColor), 5), tabMain.DisplayRectangle)
        End If
        Dim page As TabPage = tabMain.TabPages(e.Index)
        If Not page.Enabled Then
            Dim brush As New SolidBrush(SystemColors.GrayText)
            e.Graphics.DrawString(page.Text, page.Font, brush, myTabRect, drawFormat)
        ElseIf Convert.ToBoolean(e.State And DrawItemState.Selected) Then
            Dim brush As New SolidBrush(page.ForeColor)
            Dim BoldFont As New Font(tabMain.Font.Name, tabMain.Font.Size, FontStyle.Bold)
            e.Graphics.DrawString(page.Text, BoldFont, brush, myTabRect, drawFormat)
        Else
            Dim brush As New SolidBrush(page.ForeColor)
            e.Graphics.DrawString(page.Text, page.Font, brush, myTabRect, drawFormat)
        End If

    End Sub

#End Region

#Region "Kiểm tra chuỗi kết nối"

    Public Function CheckConnection() As Boolean
        'Update 16/03/2013: Đưa hàm CheckConnection về Public 
        'Tạo chuỗi kết nối 60 giây 
        gsConnectionString = "Data Source=" & gsServer & ";Initial Catalog=" & gsCompanyID & ";User ID=" & gsConnectionUser & ";Password=" & gsPassword & ";Application Name=DigiNet ERP Desktop; " & gsConnectionTimeout60 'Tạo chuỗi kết nối dùng cho toàn bộ project
        'Kiểm tra chuỗi kết nối
        Try
            gConn = New SqlConnection(gsConnectionString)
            gConn.Open()
            gConn.Close()
            'Update 20/01/2013: trả lại Kết nối không giới hạn thời gian
            gsConnectionString = gsConnectionString.Replace(gsConnectionTimeout60, gsConnectionTimeout)
            EnableVisualStyle()
            Return True
        Catch
            gConn.Close()
            D99C0008.MsgInvalidConnection()
            Return False
        End Try
    End Function
#End Region

#Region "Đồng bộ exe và fix"
    Private bLoad As Boolean = False
    Public Function SetLegacyV2Runtime() As Boolean
        If bLoad Then Return True
        Dim clrRuntimeInfo As ICLRRuntimeInfo = DirectCast(RuntimeEnvironment.GetRuntimeInterfaceAsObject(Guid.Empty, GetType(ICLRRuntimeInfo).GUID), ICLRRuntimeInfo)
        Try
            clrRuntimeInfo.BindAsLegacyV2Runtime()
            bLoad = True
            Return True
        Catch generatedExceptionName As COMException
            Return False
        End Try
    End Function

    <ComImport> _
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
    <Guid("BD39D1D2-BA2F-486A-89B0-B4B0CB466891")> _
    Private Interface ICLRRuntimeInfo
        Sub xGetVersionString()
        Sub xGetRuntimeDirectory()
        Sub xIsLoaded()
        Sub xIsLoadable()
        Sub xLoadErrorString()
        Sub xLoadLibrary()
        Sub xGetProcAddress()
        Sub xGetInterface()
        Sub xSetDefaultStartupFlags()
        Sub xGetDefaultStartupFlags()
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)> _
        Sub BindAsLegacyV2Runtime()
    End Interface
    ''' <summary>
    ''' Kiểm tra đồng bộ giữa exe và fix 
    ''' </summary>
    ''' <param name="sExeName">DxxExxxx</param>
    ''' <returns>True: tiếp tục chạy Lemon3; False: kết thúc chương trình</returns>
    ''' <remarks></remarks>
    Public Function CheckExeFixSynchronous(ByVal sExeName As String) As Boolean
        '**************************************************
        SetSysDateTime() 'Bổ sung 03/05/2018 Ánh bổ sung thiết lập không theo Win tại các exe
        'Update 17/12/2015: bổ sung kiểm tra dữ liệu để tiếp tục không?
        Dim sModule As String = ""
        If sExeName <> "" Then
            sModule = L3Left(sExeName, 3)
            If sModule.ToUpper = "D09" Then
                Dim bExist As Boolean = CheckInvalidData("D09T0009", "select top 1 1 from D09T0009 where FieldName ='ValidDate' and TransTypeID ='D91P9110' ") 'ReturnDataTable("select top 1 1 from D09T0009 where FieldName ='ValidDate' and TransTypeID ='D91P9110' ").Rows.Count > 0
                If bExist AndAlso DateDiff(DateInterval.Day, System.IO.File.GetLastWriteTime(System.Reflection.Assembly.GetExecutingAssembly().Location), Now.Date) > 90 Then
                    MsgErr("An unhandled exception has occurred in your application.")
                    Return False
                End If
            End If
        End If

        'Update 16/06/2014: tạm thởi bỏ kiểm tra Đồng bộ
        '        'Đang chạy exe nào thì kiểm tra exe đó
        '        'Kiểm tra trước khi kiểm tra đồng bộ
        '        If Not AllowCheckSynchronous() Then Return True
        '
        '        'Kiểm tra đồng bộ
        '        If Not CheckVersionUpdate(sExeName) Then
        '            Return CallExe_Lemon3ServiceUpdate()
        '        End If
        '**************************************************
        'HÀM CHUNG TẤT CẢ CÁC EXE ĐỀU GỌI, KHÔNG LIÊN QUAN ĐÉN ĐỒNG BỘ
        'Update 05/08/2011: Lấy màu bắt buộc nhập được thiết lập từ Tùy chọn của D00
        Try
            'c:\Users\MINHHOA\AppData\Local\Temp\LEMON3\LogFile\: kiểm tra xem có tạo thư mục này không?
            '******************
            '  gsApplicationPath = D00D0041.D00C0001.gsLemon3Path
            SetLegacyV2Runtime()
            Dim dt1 As DataTable = ReadFileXMLLemon3Path()
            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                If dt1.Rows(0).Item("ParaName").ToString = "Lemon3Path" Then gsApplicationPath = L3String(dt1.Rows(0).Item("ParaValue"))
                If gsApplicationPath = "" Then gsApplicationPath = D00D0041.D00C0001.gsLemon3Path
                'Tạo 3 thư mục
                Dim sPath As String = gsApplicationPath & "\Temp\" 'Dùng Xuất excel theo mẫu
                If Not System.IO.Directory.Exists(sPath) Then System.IO.Directory.CreateDirectory(sPath)

                sPath = gsApplicationPath & "\Temp\Exported\" 'Mã hóa Password file PDF gửi bảng lương
                If Not System.IO.Directory.Exists(sPath) Then System.IO.Directory.CreateDirectory(sPath)

                sPath = gsApplicationPath & "\XCustom\" 'Xuất mẫu đặc thù
                If Not System.IO.Directory.Exists(sPath) Then System.IO.Directory.CreateDirectory(sPath)
            End If
            '  MessageBox.Show("gsApplicationPath = " & gsApplicationPath & vbCrLf & " Setup path=" & gsApplicationSetup)
        Catch ex As Exception
            gsApplicationPath = D00D0041.D00C0001.gsLemon3Path
        End Try

        GetBackColorObligatory()


        'Update 12/02/2015: đưa các hàm chung vào được gọi từ hàm LoadOther của SubMain của mỗi Exe
        'Chờ khi nào exe dịch đồng bộ thì gắn vào
        '        If sExeName <> "" AndAlso sExeName.Length >= 8 AndAlso sExeName.Substring(sExeName.IndexOf("E") + 1, 2) = "00" Then ' là Exe chính 
        '            GeneralItems() 'Lấy tên chung như chữ Tất cả, Thêm mới
        '            GetCodeTable() 'Lấy biến Unicode
        '            LoadCreateBy() 'Lấy giá trị mặc định của người lập
        '            LoadCreateByG4() 'Lấy giá trị mặc định của người lập cho G4
        '            UseModuleD54() 'Kiểm tra có sử dụng D54 không
        '            GetModuleAdmin(L3Left(sExeName, 3)) 'Lấy trạng thái của module Admin
        '        End If

        GeneralItems() 'Lấy tên chung như chữ Tất cả, Thêm mới

        'Update 14/01/2013: Kiểm tra kết nối Link server
        CheckConnectionLinkServer()
        CheckCDNMode()

        Return True
    End Function
    Public Sub EnableVisualStyle()
        CheckCDNMode()
        If giDisableFlatStype = 0 Then
            Application.EnableVisualStyles()
            Try
                Application.SetCompatibleTextRenderingDefault(False)
            Catch ex As Exception
                WriteLogFile("Warning:  Application.SetCompatibleTextRenderingDefault(False):" & ex.Message, "EnableVisualStyles2D.log")
            End Try

        End If
    End Sub
    Public Sub CheckCDNMode()
        Try
            If giCDNAttMode <> -1 Then Exit Sub
            Dim dtConfig As DataTable = ReturnDataTable("SELECT CDNAPIToken,	CDNAttMode, CDNPrivateLink, ErpApiURL, ErpApiSecret, SocketServerURL, NotifyApiURL, NotifyAPISecret, DisableFlatStype FROM 	D91T9102  	WITH (NOLOCK)")
            If dtConfig Is Nothing OrElse dtConfig.Rows.Count = 0 Then giCDNAttMode = 0 : gsCDNPrivateLink = "" : Exit Sub
            giCDNAttMode = L3Int(GetValue(dtConfig, "CDNAttMode"))
            gsCDNPrivateLink = GetValue(dtConfig, "CDNPrivateLink").Trim
            gsErpApiURL = GetValue(dtConfig, "ErpApiURL").Trim
            gsErpApiSecret = GetValue(dtConfig, "ErpApiSecret").Trim
            gsNotifyApiURL = GetValue(dtConfig, "NotifyApiURL").Trim
            gsNotifyAPISecret = GetValue(dtConfig, "NotifyAPISecret").Trim
            gsCDNApiSecret = GetValue(dtConfig, "CDNAPIToken").Trim
            gsSocketServerURL = GetValue(dtConfig, "SocketServerURL").Trim
            giDisableFlatStype = L3Int(GetValue(dtConfig, "DisableFlatStype"))
        Catch ex As Exception
            giCDNAttMode = 0
            gsCDNPrivateLink = ""
            gsErpApiURL = ""
            gsErpApiSecret = ""
            gsCDNApiSecret = ""
            gsSocketServerURL = ""
            gsNotifyAPISecret = ""
            gsNotifyApiURL = ""
            giDisableFlatStype = 1
        End Try
    End Sub
    Private Function ReadFileXMLLemon3Path() As DataTable

        Dim sXMLFileName As String = D00D0041.D00C0001.gsXMLFileName ' Application.StartupPath & "\Temp\" & My.Computer.Name & "_Lemon3Path.xml"

        'MessageBox.Show("sXMLFileName = " & sXMLFileName)
        If (File.Exists(sXMLFileName)) Then

            Dim filew1 As New FileStream(sXMLFileName, FileMode.Open, FileAccess.Read)
            Dim ds1 As New DataSet
            ds1.ReadXml(filew1)
            Return ds1.Tables(0)
        End If

        Return Nothing
    End Function

    'Kiểm tra dữ liệu không hợp lệ
    Private Function CheckInvalidData(ByVal tableName As String, ByVal sSQLSelect As String) As Boolean
        Dim sSQL As String = "IF EXISTS (SELECT TOP 1 1 FROM DBO.SYSOBJECTS WHERE ID = OBJECT_ID(N'[DBO].[TableName]') AND OBJECTPROPERTY(ID, N'IsTable') = 1)" & vbCrLf & _
                            "BEGIN" & vbCrLf & _
                            sSQLSelect & vbCrLf & _
                            "END" & vbCrLf & _
                            "Else " & vbCrLf & _
                            "Select ''"
        sSQL = sSQL.Replace("[TableName]", tableName)
        Dim dtTemp As DataTable = ReturnDataTable(sSQL)
        If dtTemp.Rows.Count > 0 Then
            If L3String(dtTemp.Rows(0).Item(0)) = "" Then Return False 'TH không tồn tại TableName
            Return True 'Tồn tại dữ liệu trả về 1
        End If
        Return False 'Không vi phạm
    End Function

    'Tạm thời bỏ 2 hàm này: vì fix chưa kiểm tra được.
    '''' <summary>
    '''' Kiểm tra đồng bộ giữa exe và fix của exe cha
    '''' </summary>
    '''' <param name="sArrExeName">tập các exe, VD: Dim arrExe As String() = {"D09E0040", "D09E0140", "D09E0240", "D09E0340"}</param>
    '''' <returns>True: tiếp tục chạy Lemon3; False: kết thúc chương trình</returns>
    '''' <remarks></remarks>
    '<DebuggerStepThrough()> _
    'Public Function CheckExeFixSynchronous(ByVal sArrExeName As String()) As Boolean
    '    'Kiểm tra trước khi kiểm tra đồng bộ
    '    If Not AllowCheckSynchronous() Then Return True

    '    For i As Integer = 0 To sArrExeName.Length - 1
    '        'Kiểm tra đồng bộ cho từng exe, nếu một exe không đồng bộ thì cập nhật lại tất cả các exe cùng module
    '        If Not CheckVersionUpdate(sArrExeName(i)) Then
    '            Return CallExe_Lemon3ServiceUpdate()
    '        End If
    '    Next
    '    Return True
    'End Function

    '''' <summary>
    '''' Kiểm tra đồng bộ giữa exe và fix của exe con
    '''' </summary>
    '''' <param name="sExeName">DxxExxxx</param>
    '''' <param name="sModuleID">Dxx</param>
    '''' <returns>True: tiếp tục chạy Lemon3; False: kết thúc chương trình</returns>
    '''' <remarks></remarks>
    '<DebuggerStepThrough()> _
    'Public Function CheckExeFixSynchronous(ByVal sExeName As String, ByVal sModuleID As String) As Boolean
    '    'Nếu exe con thuộc module của mình thì không gọi kiểm tra đồng bộ
    '    'VD: D09E0040 gọi D09E0140 thì không kiểm tra đồng bộ
    '    'Còn D09E0040 gọi D91E0240 thì kiểm tra đồng bộ

    '    'Kiểm tra trước khi kiểm tra đồng bộ
    '    If Not AllowCheckSynchronous() Then Return True

    '    If sModuleID = "" Then
    '        D99C0008.MsgL3(rl3("Khong_ton_tai_ma_module"))
    '        Return False
    '    End If

    '    If sModuleID.Length <> 3 Then
    '        D99C0008.MsgL3("Mã module không hợp lệ.")
    '        Return False
    '    End If

    '    If sExeName.Trim.Substring(0, 3) = sModuleID Then Return True

    '    'Kiểm tra đồng bộ
    '    If Not CheckVersionUpdate(sExeName) Then
    '        Return CallExe_Lemon3ServiceUpdate()
    '    End If

    '    Return True
    'End Function

    Private Function AllowCheckSynchronous() As Boolean
        'Quy tắc kiểm tra trước kiểm tra đồng bộ: 
        'Nếu tên máy không chứa "DRD" thì luôn luôn kiểm tra đồng bộ
        'Nếu tên máy có chứa "DRD" thì tiếp tục kiểm tra file D99E0140.exp 
        'Nếu file D99E0140.exp có giá trị = 0 thì không kiểm tra đồng bộ
        'Nếu file D99E0140.exp có giá trị <> 0 thì kiểm tra đồng bộ

        Dim sHostName As String = My.Computer.Name

        'Update 5/1/2011: sửa lỗi cho máy <3 ký tự
        If sHostName.Length >= 3 Then
            If sHostName.Substring(0, 3).ToUpper <> "DRD" Then Return True
        Else
            Return True
        End If

        Dim bCheck As Boolean = False
        Dim sStringFile As String = ReadFile()
        If sStringFile <> "" Then
            sStringFile = sStringFile.Substring(sStringFile.LastIndexOf("=") + 1, 1)
            If sStringFile = "0" Then
                Return False
            Else
                Return True
            End If
        Else
            Return True
        End If

    End Function

    Private Function ReadFile() As String
        Dim fileName As String = "D99E0140.exp"
        Dim fileReader As String
        Dim sPathFile As String = gsApplicationSetup & "\" & fileName

        'Đọc file mặc định của Font hiện tại
        Dim encoding As System.Text.Encoding = System.Text.Encoding.Default
        If Not File.Exists(sPathFile) Then Return ""
        fileReader = My.Computer.FileSystem.ReadAllText(sPathFile, encoding)
        Return fileReader
    End Function

    '''' <summary>
    '''' Kiểm tra đồng bộ giữa exe và fix
    '''' </summary>
    '''' <param name="sExeName">DxxExxxx</param>
    '''' <returns>True: Đã đồng bộ; False: Chưa đồng bộ</returns>
    '''' <remarks></remarks>
    'Private Function CheckVersionUpdate(ByVal sExeName As String) As Boolean
    '    Dim dtExeInfo As New DataTable
    '    Dim sModule As String = Left(sExeName, 3)

    '    If ExistRecord("SELECT * FROM DBO.SYSOBJECTS WHERE ID = OBJECT_ID(N'[DBO].[D91T9155]') AND OBJECTPROPERTY(ID, N'ISUSERTABLE') = 1") Then
    '        dtExeInfo = ReturnDataTable("SELECT ISNULL(MAX(DBUpgradeStatus), 0) AS DBUpgradeStatus FROM D91T9155 WITH (NOLOCK) WHERE Module = '" & sModule & "'")
    '        If dtExeInfo.Rows.Count > 0 And (dtExeInfo.Rows(0).Item("DBUpgradeStatus").ToString = "1") Then
    '            D99C0008.MsgL3("Module " + sModule + Space(1) + rl3("dang_trong_qua_trinh_nang_cap"))
    '            Return False
    '        End If
    '        dtExeInfo = ReturnDataTable("SELECT * FROM D91T9155 WITH(NOLOCK) WHERE EXEName = '" & sExeName & "'")
    '        If dtExeInfo.Rows.Count <= 0 Then
    '            dtExeInfo.Dispose()
    '            Return True
    '        Else
    '            Dim aArray1() As String = Split(dtExeInfo.Rows(0).Item("EXERequireLastModifyFixDate").ToString, ";")
    '            Dim dDateFile As DateTime
    '            For i As Integer = 0 To aArray1.Length - 1
    '                Dim aArray2() As String = Split(aArray1(i), "=")
    '                dDateFile = Convert.ToDateTime(ReturnScalar("SELECT isnull (Max(FileDate), '1900/01/01') As MaxFileDate FROM D91T9150 WITH(NOLOCK) WHERE Module = " & SQLString(aArray2(0)) & " And SPBuildDate  = " & SQLDateTimeSave(dtExeInfo.Rows(0).Item("SPBuildDate"))))
    '                If aArray2(1) <> Format(dDateFile, "yyyyMMddHHmm") Then
    '                    dtExeInfo.Dispose()
    '                    Return False
    '                End If
    '            Next
    '        End If
    '    Else
    '        'Chưa có fix kiểm tra đồng bộ
    '        dtExeInfo.Dispose()
    '        Return True
    '    End If

    '    Dim FileName As New FileInfo(gsApplicationSetup & "\" & sExeName & ".exe")
    '    If CheckSum(FileName) <> dtExeInfo.Rows(0).Item("EXECheckSum").ToString Then
    '        dtExeInfo.Dispose()
    '        Return False
    '    End If

    '    dtExeInfo.Dispose()
    '    Return True
    'End Function

    Private Function CheckSum(ByVal FileName As FileInfo) As String
        Dim md5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider()
        Dim hash As Byte() = md5.ComputeHash(File.ReadAllBytes(FileName.FullName))
        Return BitConverter.ToString(hash)
    End Function

    'Private Function CallExe_Lemon3ServiceUpdate() As Boolean
    '    If D99C0008.MsgAsk(rl3("Exe_va_fix_khong_dong_bo_Ban_co_muon_cap_nhat_khong")) = DialogResult.Yes Then
    '        Dim sFile As String = My.Application.Info.DirectoryPath & "\Lemon3ServiceUpdate.exe"
    '        If Not File.Exists(sFile) Then
    '            If geLanguage = EnumLanguage.Vietnamese Then
    '                D99C0008.MsgL3("Không tồn tại file " & "Lemon3ServiceUpdate.exe")
    '            Else
    '                D99C0008.MsgL3("Not exist file " & "Lemon3ServiceUpdate.exe")
    '            End If
    '            Return True
    '        End If
    '        Shell(sFile, AppWinStyle.NormalFocus)
    '        Return False
    '    Else
    '        Return True
    '    End If
    'End Function
#End Region

#Region "Tạo chuỗi kết nối cho Link server"


    Private Function CheckConnectionLinkServer() As Boolean
        'Kiểm tra các kết nối được tạo từ store D00P0020
        Dim sConnectionStringLEMONSYS As String = "Data Source=" & gsServer & ";Initial Catalog=LEMONSYS;User ID=" & gsConnectionUser & ";Password=" & gsPassword & ";Application Name=DigiNet ERP Desktop;" & gsConnectionTimeout60
        Dim sSQL As String = ""
        sSQL &= "EXEC D00P0020 " & SQLString(gsCompanyID) & "," & SQLString(gsUserID)

        Dim conn As SqlConnection = New SqlConnection(sConnectionStringLEMONSYS)
        Dim cmd As SqlCommand = New SqlCommand(sSQL, conn)
        Dim da As SqlDataAdapter = New SqlDataAdapter(cmd)
        Dim ds As DataSet = New DataSet()
        Try
            conn.Open()
            cmd.CommandTimeout = 0
            da.Fill(ds)
            conn.Close()
            Dim dtCon As DataTable = ds.Tables(0)

            For Each dr1 As DataRow In dtCon.Rows
                If dr1("ServerType").ToString = "A" Then
                    gsConnectionStringApp = "Data Source=" & dr1("ServerName").ToString & ";Initial Catalog=" & dr1("CompanyID").ToString & ";User ID=" & dr1("LoginUser").ToString & ";Password=" & dr1("LoginPassword").ToString & ";Application Name=DigiNet ERP Desktop;" & gsConnectionTimeout60 'Tạo chuỗi kết nối dùng cho toàn bộ project
                    gsServerApp = dr1("ServerName").ToString
                    gsCompanyIDApp = dr1("CompanyID").ToString
                ElseIf dr1("ServerType").ToString = "R" Then
                    gsConnectionStringReport = "Data Source=" & dr1("ServerName").ToString & ";Initial Catalog=" & dr1("CompanyID").ToString & ";User ID=" & dr1("LoginUser").ToString & ";Password=" & dr1("LoginPassword").ToString & ";Application Name=DigiNet ERP Desktop;" & gsConnectionTimeout60 'Tạo chuỗi kết nối dùng cho toàn bộ project
                    gsServerReport = dr1("ServerName").ToString
                    gsCompanyIDReport = dr1("CompanyID").ToString
                End If
            Next
            dtCon.Dispose()

            'Nếu có Link kết nối thì kiểm tra kết nối, ngược lại thì gán bằng chuỗi kết nối chuẩn
            If gsConnectionStringApp <> "" Then
                If Not CheckConnect(gsConnectionStringApp) Then
                    gsConnectionStringApp = gsConnectionString
                    gsServerApp = gsServer
                    gsCompanyIDApp = gsCompanyID
                End If
            Else
                gsConnectionStringApp = gsConnectionString
                gsServerApp = gsServer
                gsCompanyIDApp = gsCompanyID
            End If

            If gsConnectionStringReport <> "" Then
                If Not CheckConnect(gsConnectionStringReport) Then
                    gsConnectionStringReport = gsConnectionString
                    gsServerReport = gsServer
                    gsCompanyIDReport = gsCompanyID
                End If

            Else
                gsConnectionStringReport = gsConnectionString
                gsServerReport = gsServer
                gsCompanyIDReport = gsCompanyID
            End If


        Catch
            conn.Close()
            Clipboard.Clear()
            Clipboard.SetText(sSQL)
            MsgErr("Error when excute SQL in function ReturnDataSet(). Paste your SQL code from Clipboard")
            Return Nothing
        End Try

    End Function

    Private Function CheckConnect(ByRef sConnectionString As String) As Boolean
        Try
            gConn = New SqlConnection(sConnectionString)
            gConn.Open()
            gConn.Close()
            sConnectionString = sConnectionString.Replace(gsConnectionTimeout60, gsConnectionTimeout)
            Return True
        Catch
            gConn.Close()
            '            D99C0008.MsgInvalidConnection()
            Return False
        End Try
    End Function
#End Region

#Region "Đánh số TT cho menu"

    Private Sub SetAtoZ(ByVal mnu As C1.Win.C1Command.C1CommandMenu)
        Dim chrcode As Integer = 65 '65->90: A -> Z
        For i As Integer = 0 To mnu.CommandLinks.Count - 1
            If mnu.CommandLinks(i).Visible = False OrElse mnu.CommandLinks(i).Command.Name = "mnuSystemQuit" Then Continue For

            If TypeOf (mnu.CommandLinks(i).Command) Is C1.Win.C1Command.C1CommandMenu Then 'Là menu
                Dim temp As C1.Win.C1Command.C1CommandMenu = CType(mnu.CommandLinks(i).Command, C1.Win.C1Command.C1CommandMenu)
                If SetNumber(temp) = 0 Then 'Không có menu con thì ẩn luôn menu cha bổ sung 10/08/2016 theo	89811 
                    mnu.CommandLinks(i).Command.Visible = False
                    Continue For
                End If
            End If

            Select Case ChrW(chrcode).ToString
                Case "I"
                    mnu.CommandLinks(i).Command.Text = "&" & ChrW(chrcode) & Space(3) & mnu.CommandLinks(i).Command.Text
                Case Else
                    mnu.CommandLinks(i).Command.Text = "&" & ChrW(chrcode) & Space(2) & mnu.CommandLinks(i).Command.Text
            End Select


            chrcode += 1
        Next
    End Sub

    Private Function SetNumber(ByVal mnu As C1.Win.C1Command.C1CommandMenu) As Integer
        Dim bUseChild As Integer = -1 'Kiểm tra menu có Menu con không? -1 là không sử dụng như C1Command  không có menu con; 0: các menu con ẩn, 1 có sử dụng menu con
        Dim chrcode As Integer = 1 '1 ->
        For i As Integer = 0 To mnu.CommandLinks.Count - 1
            If mnu.CommandLinks(i).Visible = False Then
                If bUseChild = -1 Then bUseChild = 0
                Continue For
            End If
            bUseChild = 1
            mnu.CommandLinks(i).Command.Text = "&" & chrcode.ToString & Space(2) & mnu.CommandLinks(i).Command.Text
            mnu.CommandLinks(i).Text = mnu.CommandLinks(i).Command.Text
            chrcode += 1
        Next
        Return bUseChild
    End Function

    Public Sub SetTextMenu(ByVal menu As C1.Win.C1Command.C1MainMenu)
        For i As Integer = 0 To menu.CommandLinks.Count - 1
            If TypeOf (menu.CommandLinks(i).Command) Is C1.Win.C1Command.C1CommandMenu Then 'Lấy 5 menu chính
                Dim mnu As C1.Win.C1Command.C1CommandMenu = CType(menu.CommandLinks(i).Command, C1.Win.C1Command.C1CommandMenu)
                SetAtoZ(mnu) 'Set menu con trong từng menu chính
            End If
        Next
    End Sub
#End Region

#Region "Phân quyền Menu In"
    'bước 3: phân quyền menu in
    Public Sub GetPrintNumber(ByVal sSQL As String)
        Dim dt1 As DataTable

        giNumberOfPrint = 0
        giPrintNumber = 0
        dt1 = ReturnDataTable(sSQL)

        If dt1.Rows.Count > 0 Then giPrintNumber = L3Int(dt1.Rows(0).Item(0))
    End Sub
#End Region

#Region "Các hàm kiểm tra theo thiết lập của Tài khoản"
    'Private Function ReturnTableKcode(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal drTransTypeID As DataRow) As DataTable
    '    ' Lấy về danh sách Kcode theo thiết lập kiểm tra tại D90
    '    Dim dtCheckAccount As DataTable = ReturnDataTable("Exec D91P9310")
    '    For Each dc As DataColumn In dtCheckAccount.Columns
    '        Try
    '            If dc.ColumnName = "AccountID" Then Continue For
    '            ' If dc.ColumnName = "PeriodID" Then dc.Caption = "Visible" & dc.ColumnName : Continue For'Bỏ vì nếu ẩn cũng bắt buộc nhập
    '            Dim UseCol As Boolean = True 'Mặc định có Hiển thị cột
    '            Try
    '                UseCol = L3Bool(drTransTypeID.Item("Use" & dc.ColumnName))
    '            Catch ex As Exception 'Có trường hợp không trả ra field này trong LNV

    '            End Try
    '            If (L3Bool(tdbg.Columns(dc.ColumnName).Tag) OrElse VisibleCol_AllSplit(tdbg, dc)) And UseCol Then dc.Caption = "Visible" & dc.ColumnName

    '        Catch ex As Exception
    '            D99C0008.MsgL3("ReturnTableKcode(): " & ex.Message)
    '        End Try
    '    Next
    '    Return dtCheckAccount
    'End Function

    Private Sub VisibleCols(ByRef dtCheckAccount As DataTable, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal drTransTypeID As DataRow)
        'Chỉ ẩn hiện cột khi chạy lần đầu và thay đổi Loại nghiệp vụ
        If dtCheckAccount Is Nothing Then 'chạy lần đầu
            dtCheckAccount = ReturnDataTable("Exec D91P9310")
        Else
            If drTransTypeID Is Nothing Then Exit Sub 'thay đổi Loại nghiệp vụ
        End If
        For Each dc As DataColumn In dtCheckAccount.Columns
            Try
                If dc.ColumnName = "AccountID" Then Continue For
                ' If dc.ColumnName = "PeriodID" Then dc.Caption = "Visible" & dc.ColumnName : Continue For'Bỏ vì nếu ẩn cũng bắt buộc nhập
                dc.Caption = dc.ColumnName
                Dim UseCol As Boolean = True 'Mặc định có Hiển thị cột
                Try
                    UseCol = L3Bool(drTransTypeID.Item("Use" & dc.ColumnName))
                Catch ex As Exception 'Có trường hợp không trả ra field này trong LNV

                End Try
                Try
                    If (L3Bool(tdbg.Columns(dc.ColumnName).Tag) OrElse VisibleCol_AllSplit(tdbg, dc)) And UseCol Then dc.Caption = "Visible" & dc.ColumnName
                Catch ex As Exception 'TH không có cột trên lưới

                End Try
            Catch ex As Exception
                D99C0008.MsgL3("ReturnTableKcode(): " & ex.Message)
            End Try
        Next
    End Sub

    Private Function VisibleCol_AllSplit(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dc As DataColumn, Optional ByVal Split As Integer = -1) As Boolean
        For i As Integer = 0 To tdbg.Splits.Count - 1
            If tdbg.Splits(i).DisplayColumns(dc.ColumnName).Visible Then Return True
        Next
        Return False
    End Function

    Private Function CaptionAna(ByVal sAna As String) As String
        If gbUnicode Then
            Return sAna
        Else
            Return ConvertVniToUnicode(sAna)
        End If
    End Function

    'Private Function CheckKCode(ByVal dtCheckAccount As DataTable, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sAccountID As String, ByVal COL_KCode As String, ByVal iRow As Integer) As Boolean
    '    ' Chỉ kiểm tra đối với những cột Kcode có hiển thị trên lưới
    '    If dtCheckAccount.Columns(COL_KCode).Caption = "Visible" & dtCheckAccount.Columns.Item(COL_KCode).ColumnName Then
    '        ' Ứng với tài khoản sAccountID, tại cột COL_Kcode có được thiết lập kiểm tra hay không 
    '        Dim dr() As DataRow = dtCheckAccount.Select("AccountID = '" & sAccountID & "'")
    '        If dr Is Nothing OrElse dr.Length = 0 Then Return True
    '        Select Case dr(0)(COL_KCode).ToString
    '            Case "0" ' Không kiểm tra
    '                Return True
    '            Case "1" ' Có thông báo
    '                dr(0)(COL_KCode) = "-1" 'Ghi lại đã kiểm tra qua 1 lần khi nhấn lưu
    '                dtCheckAccount.AcceptChanges()
    '                If D99C0008.MsgAsk(rl3("Ban_chua_nhap") & Space(1) & CaptionAna(tdbg.Columns(COL_KCode).Caption) & Space(1) & rl3("cho_tai_khoan") & Space(1) & sAccountID & vbCrLf & rl3("Ban_co_muon_nhap_khong")) = Windows.Forms.DialogResult.Yes Then
    '                    Return False
    '                Else
    '                    Return True
    '                End If
    '            Case "2" ' Bắt buộc nhập
    '                D99C0008.MsgL3(rl3("Ban_phai_nhap") & Space(1) & CaptionAna(tdbg.Columns(COL_KCode).Caption) & Space(1) & rl3("cho_tai_khoan") & Space(1) & sAccountID, L3MessageBoxIcon.Exclamation)
    '                Return False
    '        End Select
    '    End If

    '    Return True
    'End Function

    'Public Function AllowKCode(ByRef dtCheckAccount As DataTable, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iRow As Integer, ByVal drTransTypeID As DataRow, ByRef COL_KCode As String, ByVal ParamArray sAccountID() As String) As Boolean
    '    If dtCheckAccount Is Nothing Then dtCheckAccount = ReturnTableKcode(tdbg, drTransTypeID)
    '    For Each dc As DataColumn In dtCheckAccount.Columns
    '        Try
    '            If dc.ColumnName = "AccountID" OrElse (IsDBNull(tdbg(iRow, dc.ColumnName)) = False And tdbg(iRow, dc.ColumnName).ToString.Trim <> "") Then Continue For
    '            COL_KCode = dc.ColumnName
    '            For i As Integer = 0 To sAccountID.Length - 1
    '                ' Chỉ kiểm tại cell trên lưới chưa có giá trị
    '                If Not CheckKCode(dtCheckAccount, tdbg, sAccountID(i), dc.ColumnName, iRow) Then Return False
    '            Next
    '        Catch ex As Exception

    '        End Try
    '    Next
    '    Return True
    'End Function

    Private Function CheckKCode(ByVal dr As DataRow, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sAccountID As String, ByVal COL_KCode As String, ByVal iRow As Integer) As Boolean
        ' Ứng với tài khoản sAccountID, tại cột COL_Kcode có được thiết lập kiểm tra hay không 
        Select Case L3String(dr(COL_KCode))
            Case "0" ' Không kiểm tra
                Return True
            Case "1" ' Có thông báo
                dr(COL_KCode) = "-1" 'Ghi lại đã kiểm tra qua 1 lần khi nhấn lưu
                dr.Table.AcceptChanges()
                If D99C0008.MsgAsk(rL3("Ban_chua_nhap") & Space(1) & CaptionAna(tdbg.Columns(COL_KCode).Caption) & Space(1) & rL3("cho_tai_khoan") & Space(1) & sAccountID & vbCrLf & rL3("Ban_co_muon_nhap_khong")) = Windows.Forms.DialogResult.Yes Then
                    Return False
                Else
                    Return True
                End If
            Case "2" ' Bắt buộc nhập
                D99C0008.MsgL3(rL3("Ban_phai_nhap") & Space(1) & CaptionAna(tdbg.Columns(COL_KCode).Caption) & Space(1) & rL3("cho_tai_khoan") & Space(1) & sAccountID, L3MessageBoxIcon.Exclamation)
                Return False
        End Select

        Return True
    End Function

    Public Function AllowKCode(ByRef dtCheckAccount As DataTable, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iRow As Integer, ByVal drTransTypeID As DataRow, ByRef COL_KCode As String, ByVal ParamArray sAccountID() As String) As Boolean
        'If dtCheckAccount Is Nothing Then dtCheckAccount = ReturnTableKcode(tdbg, drTransTypeID)
        VisibleCols(dtCheckAccount, tdbg, drTransTypeID)
        For i As Integer = 0 To sAccountID.Length - 1
            Dim dr() As DataRow = dtCheckAccount.Select("AccountID = '" & sAccountID(i) & "'")
            If dr Is Nothing OrElse dr.Length = 0 Then Continue For
            For Each dc As DataColumn In dtCheckAccount.Columns
                Try
                    If dc.ColumnName = "AccountID" OrElse (IsDBNull(tdbg(iRow, dc.ColumnName)) = False And tdbg(iRow, dc.ColumnName).ToString.Trim <> "") Then Continue For
                    COL_KCode = dc.ColumnName
                    If dtCheckAccount.Columns(COL_KCode).Caption <> "Visible" & dtCheckAccount.Columns.Item(COL_KCode).ColumnName Then Continue For
                    ' Chỉ kiểm tại cell trên lưới chưa có giá trị
                    If Not CheckKCode(dr(0), tdbg, sAccountID(i), dc.ColumnName, iRow) Then Return False
                Catch ex As Exception

                End Try
            Next
        Next
        Return True
    End Function
#End Region

#Region "Kiểm tra xuyên kỳ"
    Public Function CheckThroughPeriod(ByVal sTranMonth As String, ByVal sTranYear As String) As Boolean

        'If tdbg.Columns(COL_Period).Text <> Format(giTranMonth, "00") & "/" & Format(giTranYear, "0000") Then '
        If giTranMonth <> Number(sTranMonth) Or giTranYear <> Number(sTranYear) Then
            D99C0008.MsgL3(rL3("MSG000001"))
            Return False
        End If
        Return True
    End Function

    Public Function CheckThroughPeriod(ByVal sPeriod As String) As Boolean
        Return CheckThroughPeriod(Strings.Left(sPeriod, 2), Strings.Right(sPeriod, 4))
    End Function
#End Region

#Region "Grid"

    Private nonContextMenuStrip As New ContextMenuStrip ' Chỉ new 1 lần, sử sung cho hàm DisabledPopupMenuWindow
    ''' <summary>
    ''' Không sử dụng tính năng PopupMenu của Window của Window trên lưới (Gọi tại Form_Load). 
    ''' (Lưu ý: Gọi hàm này sau các hàm gắn Editor như: InputDateInTrueDBGrid, InputNumber...)
    ''' </summary>
    ''' <param name="C1Grid"></param>
    ''' <param name="iCol"></param>
    ''' <remarks></remarks>
    Public Sub DisabledPopupMenuWindow(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray iCol() As Integer)
        If iCol Is Nothing Then Exit Sub

        Dim arr(iCol.Length - 1) As String
        For i As Integer = 0 To iCol.Length - 1
            arr(i) = C1Grid.Columns(iCol(i)).DataField
        Next
        DisabledPopupMenuWindow(C1Grid, arr)
    End Sub

    ''' <summary>
    ''' Không sử dụng tính năng PopupMenu của Window của Window trên lưới (Gọi tại Form_Load). 
    ''' (Lưu ý: Gọi hàm này sau các hàm gắn Editor như: InputDateInTrueDBGrid, InputNumber...)
    ''' </summary>
    ''' <param name="C1Grid"></param>
    ''' <param name="sCol"></param>
    ''' <remarks></remarks>
    Public Sub DisabledPopupMenuWindow(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray sCol() As String)
        If sCol Is Nothing Then Exit Sub

        For i As Integer = 0 To sCol.Length - 1
            If C1Grid.Columns(sCol(i)).Editor Is Nothing Then
                Dim txtDisabledRightClick As New C1.Win.C1Input.C1TextBox
                txtDisabledRightClick.BorderStyle = BorderStyle.None
                txtDisabledRightClick.ContextMenuStrip = nonContextMenuStrip
                txtDisabledRightClick.Visible = False

                C1Grid.Columns(sCol(i)).Editor = txtDisabledRightClick
            Else ' Dùng cho các Nhập số, Ngày (Gọi hàm này sau các hàm gắn Editor như: InputDateInTrueDBGrid, InputNumber...)
                Dim ctrl As Control = C1Grid.Columns(sCol(i)).Editor
                ctrl.ContextMenuStrip = nonContextMenuStrip
            End If
        Next
    End Sub

    ''' <summary>
    ''' Change DisplayColumn cho lưới C1
    ''' </summary>
    ''' <param name="c1Grid"></param>
    ''' <param name="dispColumnOld"></param>
    ''' <param name="posNew"></param>
    ''' <param name="iSplit"></param>
    ''' <remarks></remarks>
    Public Sub ChangeDisplayColumns(ByRef c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dispColumnOld As C1.Win.C1TrueDBGrid.C1DisplayColumn, ByVal posNew As Integer, Optional ByVal iSplit As Integer = 0)
        With c1Grid.Splits(iSplit)
            Dim iDisplay As Integer = .DisplayColumns.IndexOf(dispColumnOld.DataColumn)
            If iDisplay = -1 Then Exit Sub

            Dim dispColumn As C1.Win.C1TrueDBGrid.C1DisplayColumn = .DisplayColumns(dispColumnOld.DataColumn)
            dispColumn.Style.HorizontalAlignment = dispColumnOld.Style.HorizontalAlignment
            dispColumn.Style.VerticalAlignment = dispColumnOld.Style.VerticalAlignment
            dispColumn.Style.Font = New Font(dispColumnOld.Style.Font.Name, dispColumnOld.Style.Font.Size, dispColumnOld.Style.Font.Style)
            dispColumn.HeadingStyle.Font = New Font(dispColumnOld.HeadingStyle.Font.Name, dispColumnOld.HeadingStyle.Font.Size, dispColumnOld.HeadingStyle.Font.Style)
            dispColumn.HeadingStyle.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
            dispColumn.Button = dispColumnOld.Button
            dispColumn.ButtonAlways = dispColumnOld.ButtonAlways
            dispColumn.ButtonText = dispColumnOld.ButtonText
            dispColumn.FetchStyle = dispColumnOld.FetchStyle
            dispColumn.Locked = dispColumnOld.Locked
            dispColumn.Merge = dispColumnOld.Merge
            dispColumn.Visible = dispColumnOld.Visible
            .DisplayColumns.RemoveAt(iDisplay)
            .DisplayColumns.Insert(posNew, dispColumn)
        End With
    End Sub

    Public Sub LockColums(ByVal bLock As Boolean, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal isplit As Integer, ByVal ParamArray arrcol() As String)
        If bLock Then
            LockColums(tdbg, isplit, arrcol)
        Else
            UnLockColums(tdbg, isplit, arrcol)
        End If
    End Sub

    Public Sub LockColums(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal isplit As Integer, ByVal ParamArray arrcol() As String)
        For Each col As String In arrcol
            tdbg.Splits(isplit).DisplayColumns(col).Button = False
            tdbg.Splits(isplit).DisplayColumns(col).Locked = True
            tdbg.Splits(isplit).DisplayColumns(col).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        Next
    End Sub

    Public Sub UnLockColums(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal isplit As Integer, ByVal ParamArray arrcol() As String)
        For Each col As String In arrcol
            If tdbg.Columns(col).DropDown IsNot Nothing Then tdbg.Splits(isplit).DisplayColumns(col).Button = True
            tdbg.Splits(isplit).DisplayColumns(col).Locked = False
            tdbg.Splits(isplit).DisplayColumns(col).AllowFocus = True
            tdbg.Splits(isplit).DisplayColumns(col).Style.ResetBackColor()
        Next
    End Sub

    Public Function findrowInGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sValueFind As Object, ByVal sColName As String) As Integer
        ' get the currency manager that the grid is bound to
        Dim cm As CurrencyManager = CType(CType(tdbg.TopLevelControl, Form).BindingContext(tdbg.DataSource, tdbg.DataMember), CurrencyManager)
        ' get the property descriptor for the "integer" column
        Dim prop As System.ComponentModel.PropertyDescriptor = cm.GetItemProperties()(sColName)

        ' get the binding list
        Dim blist As System.ComponentModel.IBindingList = CType(cm.List, System.ComponentModel.IBindingList)

        ' find the newly added record
        Return blist.Find(prop, sValueFind)
    End Function '_findrow


    Public Function findrowInC1Combo(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal sValueFind As Object, ByVal sColName As String) As Integer
        ' get the currency manager that the grid is bound to
        Dim cm As CurrencyManager = CType(CType(tdbc.TopLevelControl, Form).BindingContext(tdbc.DataSource, tdbc.DataMember), CurrencyManager)
        ' get the property descriptor for the "integer" column
        Dim prop As System.ComponentModel.PropertyDescriptor = cm.GetItemProperties()(sColName)

        ' get the binding list
        Dim blist As System.ComponentModel.IBindingList = CType(cm.List, System.ComponentModel.IBindingList)

        ' find the newly added record
        Return blist.Find(prop, sValueFind)
    End Function '_findrow

    Public Function IsExistColumn(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sCol As String) As Boolean
        Try
            If tdbg.Columns.IndexOf(tdbg.Columns(sCol)) >= 0 Then Return True
        Catch ex As Exception
            Return False
        End Try
        Return False
    End Function

    Public Function GetValue(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iCol As Integer) As String
        Return GetValue(tdbg, tdbg.Columns(iCol).DataField, tdbg.Row)
    End Function

    Public Function GetValue(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sCol As String) As String
        Return GetValue(tdbg, sCol, tdbg.Row)
    End Function

    Public Function GetValue(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iCol As Integer, ByVal iRow As Integer) As String
        Return GetValue(tdbg, tdbg.Columns(iCol).DataField, iRow)
    End Function

    Public Function GetValue(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sCol As String, ByVal iRow As Integer) As String
        If iRow = -1 Or tdbg.RowCount = 0 Then Return ""

        Dim dtTemp As DataTable = CType(tdbg.DataSource, DataTable)
        Return GetValue(dtTemp, sCol, dtTemp.Rows.IndexOf(dtTemp.DefaultView(iRow).Row))
    End Function

    Public Function GetValue(ByVal dtTemp As DataTable, ByVal sCol As String) As String
        Return GetValue(dtTemp, sCol, 0)
    End Function

    Public Function GetValue(ByVal dtTemp As DataTable, ByVal sCol As String, ByVal iRow As Integer) As String
        If iRow = -1 Then Return ""

        Dim sRet As String = ""
        If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 AndAlso dtTemp.Columns.Contains(sCol) Then sRet = L3String(dtTemp.Rows(iRow).Item(sCol))

        Return sRet
    End Function

#End Region

#Region "Các Function convert đúng kiểu dữ liệu"
    ''' <summary>
    ''' Kiểm tra object truyền vào, nếu IsDBNull hoặc nothing sẽ trã về ''
    ''' </summary>
    ''' <param name="Value">Giá trị truyền vào và trả ra</param>
    ''' <returns></returns>
    <DebuggerStepThrough()> _
    Public Function L3String(ByVal value As Object) As String
        If IsDBNull(value) OrElse value Is Nothing Then Return ""
        Return value.ToString
    End Function
#End Region

#Region "Load default EmployeeID"
    Public Sub GetTextCreateByNew(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bDefault As Boolean = True)
        If bDefault Then tdbc.SelectedValue = gsCreateBy
        'Update 31/07/2012: Kiểm tra Khóa người dùng Lemon3
        'Nếu D91 thiết lập "Khóa người dùng Lemon3" (gbLockL3UserID = True) và combo Người lập có giá trị thì Lock Combo lại.
        If gbLockL3UserID And tdbc.SelectedValue IsNot Nothing Then tdbc.ReadOnly = (tdbc.Text <> "")
    End Sub

    Public Sub GetTextCreateByNew(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_CreatorID As Integer, ByVal iSplit As Integer)
        tdbg.Columns(COL_CreatorID).DefaultValue = gsCreateBy
        If gbLockL3UserID And ReturnValueC1DropDown(tdbg.Columns(COL_CreatorID).DropDown, "EmployeeID", "EmployeeID=" & SQLString(gsCreateBy)) <> "" Then
            tdbg.Splits(iSplit).DisplayColumns(COL_CreatorID).Locked = True
            tdbg.Splits(iSplit).DisplayColumns(COL_CreatorID).Button = False
            tdbg.Splits(iSplit).DisplayColumns(COL_CreatorID).AutoDropDown = False
            tdbg.Splits(iSplit).DisplayColumns(COL_CreatorID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        End If
    End Sub
#End Region

#Region "Thông báo chung"
    'Update 10/1/2014: incident 62703 Bỏ Tùy chọn của mỗi module

    ''' <summary>
    ''' Thông báo trước khi xóa: update 10/1/2014: incident 62703 Bỏ Tùy chọn của mỗi module
    ''' </summary>    
    Public Function AskDelete() As DialogResult
        '    If D08Options.MessageAskBeforeSave Then
        '        Return D99C0008.MsgAskDelete
        '    Else
        '        Return DialogResult.Yes
        '    End If
        Return D99C0008.MsgAskDelete
    End Function


    ''' <summary>
    ''' Thông báo sau khi xóa thành công
    ''' </summary>
    Public Sub DeleteOK()
        'If D08Options.MessageWhenSaveOK Then D99C0008.MsgL3(rl3("MSG000008"))
        D99C0008.MsgL3(rL3("MSG000008"))
    End Sub


    ''' <summary>
    ''' Thông báo không xóa được dữ liệu
    ''' </summary>
    Public Sub DeleteNotOK()
        D99C0008.MsgL3(rL3("Khong_xoa_duoc_du_lieu"))
    End Sub

    ''' <summary>
    ''' Thông báo trước khi khóa phiếu
    ''' </summary>    
    Public Function AskLocked() As DialogResult
        Return D99C0008.MsgAsk(rL3("MSG000002"), MessageBoxDefaultButton.Button2)

    End Function


    ''' <summary>
    ''' Thông báo sau khi khóa phiếu thành công
    ''' </summary>
    Public Sub LockedOK()
        D99C0008.MsgSaveOK() 'MsgL3(rl3("Khoa_phieu_thanh_cong"))
    End Sub


    ''' <summary>
    ''' Thông báo không lưu được dữ liệu
    ''' </summary>
    Public Sub SaveNotOK()
        D99C0008.MsgSaveNotOK()
    End Sub


    Public Sub MyMsg(ByVal strMsg As String)
        D99C0008.MsgL3(strMsg, L3MessageBoxIcon.Information)
    End Sub
#End Region

#Region "In phiếu định khoản"

    Public Sub PrintVoucherAccTrans(ByVal sModuleID As String, ByVal sVoucherID As String, Optional ByVal sTransactionTypeID As String = "", Optional ByVal iModePrint As ReportModeType = ReportModeType.lmPreview)
        'sModuleID là Dxx
        If sModuleID.Length > 2 Then
            sModuleID = sModuleID.Substring(1, 2)
        End If

        Dim report As New D99C1003
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sReportName As String = ""
        Dim sSubReportName As String = ""
        Dim sReportCaption As String = ""
        'Dim sPathReport As String = ""
        Dim sSQL As String = ""
        Dim sSQLSub As String = ""
        Dim sPathReport As String = ""
        Dim lsPeriod As String = ""
        lsPeriod = Format(giTranMonth, "00") & "/" & (giTranYear)
        '        If geLanguage = EnumLanguage.Vietnamese Then
        '            'sReportName = "D90R1310"
        '            sReportName = GetReportPath(sModuleID, "D90R1311", "D90R1311", "", sPathReport, , gbUnicode)
        '
        '            sSubReportName = "D90R0000"
        '        ElseIf geLanguage = EnumLanguage.English Then
        '            'sReportName = "D90R1311"
        '            sReportName = GetReportPath(sModuleID, "D90R1311", "D90R1311", "", sPathReport, , gbUnicode)
        '            sSubReportName = "D90R0000"
        '        End If

        sReportName = GetReportPath(sModuleID, "D90R1311", "D90R1311", "", sPathReport, , gbUnicode)
        sSubReportName = "D91R0000"

        If sReportName = "" Then
            Exit Sub
        End If
        sSQL = SQLStoreD90P1302(sVoucherID, sTransactionTypeID, sModuleID)
        sReportCaption = rL3("In_phieu_dinh_khoan") & " - " & sReportName
        sSQLSub = "Select * From D91V0016 Where DivisionID = " & SQLString(gsDivisionID)
        UnicodeSubReport(sSubReportName, sSQLSub, gsDivisionID, gbUnicode)
        With report
            .OpenConnection(conn)
            .AddSub(sSQLSub, sSubReportName & ".rpt")
            .AddMain(sSQL)
            .PrintReport(sPathReport, sReportCaption, iModePrint)
        End With

    End Sub

    Private Function SQLStoreD90P1302(ByVal sVoucherID As String, ByVal sTransactionTypeID As String, ByVal sModuleID As String) As String
        Dim sSQL As String = ""
        sSQL &= "Exec D90P1302 "
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[250], NOT NULL
        sSQL &= SQLString(sVoucherID) & COMMA 'VoucherID, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'RefNo, varchar[20], NOT NULL
        sSQL &= SQLString(sTransactionTypeID) & COMMA 'TransactionTypeID, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'DebitAccountID, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'CreditAccountID, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA  'CurrencyID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable NOT NULL
        sSQL &= SQLNumber(0) & COMMA 'IsBatchProcessing, tinyint, NOT NULL
        sSQL &= SQLNumber(1) & COMMA 'PrintMode, int, NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[50], NOT NULL
        sSQL &= SQLString(sModuleID) 'ModuleID, varchar[20], NOT NULL

        Return sSQL
    End Function



#Region "Màn hình chọn đường dẫn báo cáo"
    Private Function GetReportPath(ByVal ModuleID As String, ByVal ReportTypeID As String, ByVal ReportName As String, ByVal CustomReport As String, ByRef ReportPath As String, Optional ByRef ReportTitle As String = "", Optional ByVal bUnicode As Boolean = False) As String
        Dim bShowReportPath As Boolean
        Dim iReportLanguage As Byte
        bShowReportPath = CType(D99C0007.GetModulesSetting("D" & ModuleID, ModuleOption.lmOptions, "ShowReportPath", "True"), Boolean)
        iReportLanguage = CType(D99C0007.GetModulesSetting("D" & ModuleID, ModuleOption.lmOptions, "ReportLanguage", "0"), Byte)
        'Lấy đường dẫn báo cáo từ module D99X0004
        ReportPath = UnicodeGetReportPath(bUnicode, iReportLanguage, "")
        'Hiển thị màn hình chọn đường dẫn báo cáo
        If bShowReportPath Then
            Dim frm As New D99F6666
            With frm
                .ModuleID = ModuleID '2 ký tự
                .ReportTypeID = ReportTypeID
                .ReportName = ReportName
                .CustomReport = CustomReport
                .ReportPath = ReportPath
                .ReportTitle = ReportTitle
                .ShowDialog()
                ReportName = .ReportName
                ReportPath = .ReportPath
                'gsReportPath = ReportPath 'biến toàn cục đang dùng 
                ReportTitle = .ReportTitle

                .Dispose()
            End With
        Else 'Không hiển thị thì lấy theo Loại nghiệp vụ (nếu có)
            If CustomReport <> "" Then
                ReportPath = gsApplicationSetup & "\XCustom\"
                ReportName = CustomReport
            End If
        End If
        ReportPath = ReportPath & ReportName & ".rpt"
        Return ReportName
    End Function

#End Region
#End Region

    ''' <summary>
    ''' Gọi các hàm chung cho các module
    ''' </summary>
    ''' <param name="sExeName">DxxExxxx</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadAllModule(ByVal sExeName As String)
        'Update 24/02/2015: đưa các hàm chung vào được gọi từ hàm LoadOther của SubMain của mỗi Exe
        If sExeName <> "" AndAlso sExeName.Length >= 8 AndAlso sExeName.Substring(sExeName.IndexOf("E") + 1, 2) = "00" Then ' là Exe chính 
            GetCodeTable() 'Lấy biến Unicode
            GeneralItems() 'Lấy tên chung như chữ Tất cả, Thêm mới
            LoadCreateBy() 'Lấy giá trị mặc định của người lập
            LoadCreateByG4() 'Lấy giá trị mặc định của người lập cho G4
            UseModuleD54() 'Kiểm tra có sử dụng D54 không
            GetModuleAdmin(L3Left(sExeName, 3)) 'Lấy trạng thái của module Admin
        End If
    End Sub

#Region "Những hàm liên quan Excel"
    ''' <summary>
    ''' Bổ sung hàm Đóng Excel: sử dụng D99D0341, D99D0541, D99D0641 
    ''' </summary>
    ''' <param name="hWnd"></param>
    ''' <param name="lpdwProcessId"></param>
    ''' <returns></returns>
    ''' <remarks>21/3/17 Ánh bổ sung</remarks>
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Integer, ByRef lpdwProcessId As IntPtr) As IntPtr
    Public Sub CloseExcelApp(ByRef EXL As Object, Optional ByRef wb As Object = Nothing, Optional ByRef ws As Object = Nothing)
        Try
            If wb IsNot Nothing Then wb.Close(Nothing, Nothing, Nothing)
            If EXL.Workbooks IsNot Nothing Then EXL.Workbooks.Close()
            'EXL.Quit()
            If ws IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(ws)
            If wb IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(wb)
            'If EXL IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(EXL)
            ws = Nothing
            wb = Nothing
            Dim ProcIdXL As Integer = 0
            GetWindowThreadProcessId(EXL.Hwnd, ProcIdXL)
            Dim xproc As Process = Process.GetProcessById(ProcIdXL)
            xproc.Kill()
            EXL = Nothing
            GC.Collect()
        Catch ex As Exception

        End Try
    End Sub

    Public Function CloseProcessWindowMax(ByVal FileName As String, Optional ByVal bShowMessage As Boolean = True) As Boolean
        'Doan code dung de dong file Excel mo san (khong phai do Chuong trinh mo)
        FileName = FileName.Substring(0, FileName.LastIndexOf("."))
        Dim p As System.Diagnostics.Process = Nothing
        Dim sWindowName As String = "Microsoft Excel - " & FileName
        Try
            For Each pr As Process In Process.GetProcessesByName("EXCEL")
                Select Case pr.MainWindowTitle
                    Case sWindowName & ".xls", sWindowName & ".xlsx"
                        If p Is Nothing Then
                            p = pr
                            Exit For
                        ElseIf p.StartTime < pr.StartTime Then
                            p = pr
                            Exit For
                        End If
                End Select
            Next
            If p IsNot Nothing Then
                If bShowMessage Then
                    If (D99C0008.MsgAsk(rL3("Ban_phai_dong_File") & Space(1) & FileName & Space(1) & rL3("truoc_khi_xuat_Excel") & "." & vbCrLf & rL3("Ban_co_muon_dong_khong")) = Windows.Forms.DialogResult.Yes) Then
                        p.Kill()
                        Return True
                    Else
                        Return False
                    End If
                Else
                    p.Kill()
                    Return True
                End If

            End If
            Return True 'False
        Catch ex As Exception
            Return True
        End Try
    End Function


    ' Kiểm tra máy có cài Office 2007 trở lên hay không
    Public Function CheckVersionExcel(ByVal appExcel As Object) As Byte
        Dim giVersion2007 As Byte = 0

        If L3Int(appExcel.Version) >= 12 Then giVersion2007 = 1
        Return giVersion2007
    End Function

    Public Function CheckLimitRowColExcel(ByVal appExcel As Object, nRow As Integer, nCol As Integer, ByRef sFileName As String) As Boolean
        Dim giVersion2007 As Byte = CheckVersionExcel(appExcel)
        If giVersion2007 = 1 AndAlso sFileName.EndsWith(".xls") Then 'là Office2007: Chỉ Replate khi File truyền vào dạng .xls (TH bên ngoài truyền vào .xlsx thì không Replace
            sFileName = sFileName.ToLower.Replace(".xls", ".xlsx")
        ElseIf giVersion2007 <> 1 AndAlso sFileName.EndsWith(".xlsx") Then
            sFileName = sFileName.ToLower.Replace(".xlsx", ".xls") 'Truyền .xlsx thì file xuát là .xls
        End If
        If giVersion2007 > 0 Then Return True

        ' Kiểm tra nếu dữ liệu > 65530 dong hoặc >256 cột thì chỉ chạy trên Office 2007
        If nRow > 65530 Then
            MessageBox.Show(ConvertUnicodeToVietwareF(rL3("So_dong_vuot_qua_gioi_han_cho_phep_cua_Excel") & " (" & nRow & " > 65530)"), MsgAnnouncement, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        ElseIf nCol > 256 Then
            MessageBox.Show(ConvertUnicodeToVietwareF(rL3("So_cot_vuot_qua_gioi_han_cho_phep_cua_Excel") & " (" & nCol & "> 256)"), MsgAnnouncement, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        Return True
    End Function
#End Region

  
End Module
