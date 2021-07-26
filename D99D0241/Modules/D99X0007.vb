'#######################################################################################
'#                                     CHÚ Ý
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Người tạo: Nguyễn Thị Ánh
'# Ngày cập nhật cuối cùng: 11/03/2014
'# Người cập nhật cuối cùng: Nguyễn Thị Ánh
'# Diễn giải: Kiểm tra nhập số C1NumericEdit
'# Bỏ  C1NumericEdit.LostFocus tại InputNumber(C1NumericEdit)
'# Bổ sung Try Catch tại hàm LoadCustomFormat
'# Bổ sung thêm gắn tag của cột trên lưới trong hàm AddColFormatText(): 13/08/2013
'# Bổ sung thêm LoadFormNotINV vào LoadCustomFormat()
'# Bổ sung FooterSum cho cột Percent
'#######################################################################################

Imports System

Public Module D99X0007

    ''' <summary>
    ''' Kiểu format số của C1NumericEdit 
    ''' </summary>
    ''' <param name="Number"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CustomFormat(ByVal Number As Object, Optional ByVal sRefix As String = "N") As String
        If Number Is Nothing OrElse Number.ToString = "" OrElse L3Int(Number) < 0 Then Return sRefix & "0"
        Return sRefix & Number.ToString
    End Function

#Region "Nhập số trên C1NumericEdit"

    ''' <summary>
    ''' Nhập số cho control C1NumbericEdit có cùng DataType dưới SQL
    ''' </summary>
    ''' <param name="C1NumericEdit">tên control C1NumbericEdit</param>
    ''' <param name="SQLTypeName">Kiểu dữ liệu của SQL (VD: SqlDbType.Decimal )</param>
    ''' <param name="sCustomFormat">Định dạng số lẻ (VD: "N2" là 2 số lẻ) </param>
    ''' <param name="bSign"> Cho phép nhập số âm (-)</param>
    ''' <param name="iPrecision">Lấy định dạng phần nguyên của Decimal(28,8)->truyền vào 28</param>
    ''' <param name="iScale">Lấy định dạng phần thập phân của Decimal(28,8)->truyền vào 8</param>
    ''' <remarks></remarks>
    Public Sub InputNumber(ByRef C1NumericEdit As C1.Win.C1Input.C1NumericEdit, ByVal SQLTypeName As System.Data.SqlDbType, Optional ByVal sCustomFormat As String = "N2", Optional ByVal bSign As Boolean = False, Optional ByVal iPrecision As Integer = -1, Optional ByVal iScale As Integer = -1)
        If SQLTypeName = SqlDbType.Decimal Then
            InputDecimalInC1Numeric(C1NumericEdit, sCustomFormat, iPrecision, iScale, bSign)
        Else
            InputNumberInC1Numeric(C1NumericEdit, SQLTypeName, sCustomFormat, bSign)
        End If
        'Bị sai khi C1NumericEdit đặt trong groupbox, TabControl
        '        AddHandler C1NumericEdit.LostFocus, AddressOf C1NumericEdit_LostFocus 'Focus đúng cột trên lưới khi nhấn Enter(cột kế là Lock và Unlock)
    End Sub

    ''' <summary>
    ''' Nhập số cho nhiều control C1NumbericEdit có cùng DataType dưới SQL
    ''' </summary>
    ''' <param name="C1NumericEdit">tên control C1NumbericEdit</param>
    ''' <param name="SQLTypeName">Kiểu dữ liệu của SQL (VD: SqlDbType.Decimal )</param>
    ''' <param name="sCustomFormat">Định dạng số lẻ (VD: "N2" là 2 số lẻ) </param>
    ''' <param name="bSign"> Cho phép nhập số âm (-)</param>
    ''' <param name="iPrecision">Lấy định dạng phần nguyên của Decimal(28,8)->truyền vào 28</param>
    ''' <param name="iScale">Lấy định dạng phần thập phân của Decimal(28,8)->truyền vào 8</param>
    ''' <remarks></remarks>
    Public Sub InputNumber(ByRef C1NumericEdit() As C1.Win.C1Input.C1NumericEdit, ByVal SQLTypeName As System.Data.SqlDbType, Optional ByVal sCustomFormat As String = "N2", Optional ByVal bSign As Boolean = False, Optional ByVal iPrecision As Integer = -1, Optional ByVal iScale As Integer = -1)
        If SQLTypeName = SqlDbType.Decimal Then
            For i As Integer = 0 To C1NumericEdit.Length - 1
                InputDecimalInC1Numeric(C1NumericEdit(i), sCustomFormat, iPrecision, iScale, bSign)
            Next
        Else
            For i As Integer = 0 To C1NumericEdit.Length - 1
                InputNumberInC1Numeric(C1NumericEdit(i), SQLTypeName, sCustomFormat, bSign)
            Next
        End If
    End Sub

    ''' <summary>
    ''' Nhập Tỷ lệ dùng control C1NumericEdit
    ''' </summary>
    ''' <param name="C1NumericEdit">tên control C1NumbericEdit</param>
    ''' <remarks></remarks>
    Public Sub InputPercent(ByRef C1NumericEdit As C1.Win.C1Input.C1NumericEdit)
        'InputNumber(C1NumericEdit, SqlDbType.Money, "##0.00%") 'Nhập Tỷ lệ %
        InputPercent(C1NumericEdit, 2)
    End Sub

    ''' <summary>
    ''' Nhập Tỷ lệ dùng control C1NumericEdit
    ''' </summary>
    ''' <param name="C1NumericEdit">tên control C1NumbericEdit</param>
    ''' <remarks></remarks>
    Public Sub InputPercent(ByRef C1NumericEdit As C1.Win.C1Input.C1NumericEdit, InterestRateDecimals As Integer)
        Dim sCusFormat As String = GetFormatPercent(InterestRateDecimals)
        InputNumber(C1NumericEdit, SqlDbType.Money, sCusFormat) 'Nhập Tỷ lệ %
    End Sub
    Private Function GetFormatPercent(InterestRateDecimals As Integer) As String
        Dim sCusFormat As String = "##0"
        If InterestRateDecimals > 0 Then
            sCusFormat &= "."
            For i As Integer = 1 To InterestRateDecimals
                sCusFormat &= "0"
            Next
        End If
        sCusFormat &= "%"
        Return sCusFormat
    End Function
    ''' <summary>
    ''' Thay đổi số lẻ của c1NumberEdit
    ''' </summary>
    ''' <param name="C1NumericEdit"></param>
    ''' <param name="sCustomFormat"></param>
    ''' <remarks></remarks>
    Public Sub ReFormatNumber(ByRef C1NumericEdit As C1.Win.C1Input.C1NumericEdit, Optional ByVal sCustomFormat As String = "N2")
        C1NumericEdit.CustomFormat = sCustomFormat
        C1NumericEdit.DisplayFormat.CustomFormat = sCustomFormat
        C1NumericEdit.EditFormat.CustomFormat = sCustomFormat
        C1NumericEdit.DisplayFormat.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
    End Sub

#End Region

#Region "Nhập số trên C1TrueDBGrid"

    ''' <summary>
    ''' Cấu trúc định dạng một cột số trên lưới
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure FormatColumn
        Public FieldName As String
        Public NumberFormat As String  'Exp: "N2"
        Public DataType As System.Data.SqlDbType
        Public Sign As Boolean 'Cho phép nhập số Âm. Default = False
        'Chỉ sử dụng khi dataType dưới SQL là Decimal. Exp: Decimal(18,2)
        Public Precision As Integer '=18
        Public Scale As Integer '=2
    End Structure

    ''' <summary>
    ''' Thay đổi NumberFormat của 1 cột trên lưới
    ''' </summary>
    ''' <param name="tdbg"></param>
    ''' <param name="sCustomFormat">chuỗi format</param>
    ''' <param name="sColumnName">mảng DataField của cột có cùng NumberFormat</param>
    ''' <remarks></remarks>
    Public Sub ReFormatNumber(ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sCustomFormat As String, ByVal ParamArray sColumnName() As String)
        For i As Integer = 0 To sColumnName.Length - 1
            Dim c1NumberEdit As C1.Win.C1Input.C1NumericEdit = CType(tdbg.Columns(sColumnName(i)).Editor, C1.Win.C1Input.C1NumericEdit)
            ' Trường hợp chưa gọi NumberFormat thì bị lỗi
            If c1NumberEdit Is Nothing Then Continue For
            ReFormatNumber(c1NumberEdit, sCustomFormat)
            tdbg.Columns(sColumnName(i)).NumberFormat = sCustomFormat
        Next
    End Sub

    ''' <summary>
    ''' Thêm cột (có DataType dưới SQL Decimal) vào mảng FormatColumn
    ''' </summary>
    ''' <param name="arrCol">mảng các cột</param>
    ''' <param name="sColumnName">DataField của cột</param>
    ''' <remarks></remarks>
    Public Sub AddPercentColumns(ByRef arrCol() As FormatColumn, ByVal sColumnName As String)

        AddPercentColumns(arrCol, sColumnName, 2)
    End Sub

    ''' <summary>
    ''' Thêm cột (có DataType dưới SQL Decimal) vào mảng FormatColumn
    ''' </summary>
    ''' <param name="arrCol">mảng các cột</param>
    ''' <param name="sColumnName">DataField của cột</param>
    ''' <remarks></remarks>
    Public Sub AddPercentColumns(ByRef arrCol() As FormatColumn, ByVal sColumnName As String, InterestRateDecimals As Integer) 
        AddNumberColumns(arrCol, SqlDbType.Money, sColumnName, GetFormatPercent(InterestRateDecimals), False)
    End Sub

    ''' <summary>
    ''' Thêm cột (có DataType dưới SQL Decimal) vào mảng FormatColumn
    ''' </summary>
    ''' <param name="arrCol">mảng các cột</param>
    ''' <param name="sColumnName">DataField của cột</param>
    ''' <param name="sCustomFormat">Định dạng số lẻ. Exp: N2 -> nhập 2 số lẻ</param>
    ''' <param name="iPrecision">Decimal(28,8)->truyền vào 28</param>
    ''' <param name="iScale">Decimal(28,8)->truyền vào 8</param>
    ''' <param name="bSign">Default is False. Cho phép nhập số âm (-)</param>
    ''' <remarks></remarks>
    Public Sub AddDecimalColumns(ByRef arrCol() As FormatColumn, ByVal sColumnName As String, ByVal sCustomFormat As String, ByVal iPrecision As Integer, ByVal iScale As Integer, Optional ByVal bSign As Boolean = False)
        AddNumberColumns(arrCol, SqlDbType.Decimal, sColumnName, sCustomFormat, bSign)
        arrCol(arrCol.Length - 1).Precision = iPrecision
        arrCol(arrCol.Length - 1).Scale = iScale
    End Sub

    ''' <summary>
    ''' Thêm cột (có DataType dưới SQL không phải Decimal) vào mảng FormatColumn
    ''' </summary>
    ''' <param name="arrCol">mảng các cột</param>
    ''' <param name="SQLTypeName">DataType dưới SQL của cột</param>
    ''' <param name="sColumnName">DataField của cột</param>
    ''' <param name="sCustomFormat">Định dạng số lẻ. Exp: N2 -> nhập 2 số lẻ</param>
    ''' <param name="bSign">Default is False. Cho phép nhập số âm (-)</param>
    ''' <remarks></remarks>
    Public Sub AddNumberColumns(ByRef arrCol() As FormatColumn, ByVal SQLTypeName As System.Data.SqlDbType, ByVal sColumnName As String, ByVal sCustomFormat As String, Optional ByVal bSign As Boolean = False)
        If arrCol Is Nothing Then
            ReDim arrCol(0)
        Else
            ReDim Preserve arrCol(arrCol.Length)
        End If
        Dim index As Integer = arrCol.Length - 1
        arrCol(index).DataType = SQLTypeName
        arrCol(index).Sign = bSign
        arrCol(index).FieldName = sColumnName
        arrCol(index).NumberFormat = sCustomFormat
    End Sub

    ''' <summary>
    ''' Nhập số trên lưới dạng mới có dùng control C1NumericEdit
    ''' </summary>
    ''' <param name="tdbg">tên lưới</param>
    ''' <param name="arrCols"> Tập mảng cấu trúc các cột số </param>
    ''' <remarks></remarks>
    Public Sub InputNumber(ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal arrCols() As FormatColumn)
        InputNumber(tdbg, arrCols, True)
    End Sub
    Public Sub InputNumber(ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal arrCols() As FormatColumn, bUpdateDataForGrid As Boolean)
        If (tdbg Is Nothing) Then Return
        If (arrCols Is Nothing) Then Return

        ' Duyệt theo mảng cột số trên lưới truyền vào
        ' Gán control C1NumericEdit vào từng cột tương ứng
        ' Có thể dùng 1 đối tượng C1NumericEdit cho tât cả các cột
        ' nhưng vì để nhấn Tab, Enter trên lưới không bị ảnh hưởng
        ' và có thể các cột khác nhau vể số lẻ nên mỗi cột nên gán cho một đối tượng C1NumericEdit
        For i As Integer = 0 To arrCols.Length - 1
            Dim C1NumericEdit As New C1.Win.C1Input.C1NumericEdit()
            C1NumericEdit.Name = "C1NumericEdit" + i.ToString()
            ' Để nhập số không bị giới hạn nên truyền vào chuỗi format dạng N0,N1,N2,N3,N4 ...
            ' Gọi hàm khởi tạo một số thuộc tính cho C1NumericEdit
            Select Case arrCols(i).DataType
                Case SqlDbType.Decimal
                    InputDecimalInC1Numeric(C1NumericEdit, arrCols(i).NumberFormat, arrCols(i).Precision, arrCols(i).Scale, arrCols(i).Sign)
                Case Else
                    InputNumberInC1Numeric(C1NumericEdit, arrCols(i).DataType, arrCols(i).NumberFormat, arrCols(i).Sign)
            End Select

            C1NumericEdit.Visible = False
            AddHandler C1NumericEdit.LostFocus, AddressOf C1NumericEdit_LostFocus 'Focus đúng cột trên lưới khi nhấn Enter(cột kế là Lock và Unlock)
            If bUpdateDataForGrid Then AddHandler tdbg.AfterColUpdate, AddressOf tdbg_AfterColUpdate 'Bố sung vì lỗi không gắn được giá trị trên lưới 22/07/2017
            ' Gán control InitC1NumericEdit cho cột trên lưới
            If arrCols(i).FieldName Is Nothing Then Continue For

            tdbg.Columns(arrCols(i).FieldName).NumberFormat = arrCols(i).NumberFormat
            tdbg.Columns(arrCols(i).FieldName).Editor = C1NumericEdit
        Next
    End Sub


#End Region

#Region "Chuyển DataType dưới SQL server sang DataType .Net Framework"
    ''' <summary>
    ''' Chuyển dataType dưới SQL server (không phải Decimal) sang DataType .Net Framework
    ''' </summary>
    ''' <param name="SQLTypeName">DataType dưới Server</param>
    ''' <param name="dMin">Trả về Giới hạn nhỏ nhất</param>
    ''' <param name="dMax">Trả về Giới hạn lớn nhất</param>
    ''' <return>DataType .Net Framework</return>
    ''' <remarks></remarks>
    Public Function GetProviderData(ByVal SQLTypeName As System.Data.SqlDbType, ByRef dMin As Double, ByRef dMax As Double) As System.Type
        Dim typeNet As System.Type = GetType(System.Double)
        Select Case SQLTypeName
            Case SqlDbType.TinyInt 'Đã kiểm tra
                typeNet = GetType(System.Byte)
                dMin = System.Data.SqlTypes.SqlByte.MinValue.Value '0
                dMax = System.Data.SqlTypes.SqlByte.MaxValue.Value '255
            Case SqlDbType.BigInt 'Đã kiểm tra
                typeNet = GetType(System.Int64)
                dMin = System.Data.SqlTypes.SqlInt64.MinValue.Value '-9,223,372,036,854,775,808
                dMax = System.Data.SqlTypes.SqlInt64.MaxValue.Value '9,223,372,036,854,775,807
            Case SqlDbType.Real 'Đã kiểm tra
                typeNet = GetType(System.Int64)
                dMin = System.Data.SqlTypes.SqlSingle.MinValue.Value '-3.40282347E+38
                dMax = System.Data.SqlTypes.SqlSingle.MaxValue.Value '3.40282347E+38
            Case SqlDbType.Float 'Đã kiểm tra
                typeNet = GetType(System.Int64)
                dMin = System.Data.SqlTypes.SqlDouble.MinValue.Value '-1.7976931348623157E+308
                dMax = System.Data.SqlTypes.SqlDouble.MaxValue.Value '1.7976931348623157E+308
            Case SqlDbType.Bit
                typeNet = GetType(System.Boolean)
            Case SqlDbType.SmallMoney  'Đã kiểm tra
                typeNet = GetType(System.Decimal)
                dMin = -214748.3648
                dMax = 214748.3647
            Case SqlDbType.Money  'Đã kiểm tra
                typeNet = GetType(System.Decimal)
                dMin = System.Data.SqlTypes.SqlMoney.MinValue.Value '-922337203685477.5808D
                dMax = System.Data.SqlTypes.SqlMoney.MaxValue.Value '922337203685477.5807D
            Case SqlDbType.Int 'Đã kiểm tra
                typeNet = GetType(System.Int32)
                dMin = System.Data.SqlTypes.SqlInt32.MinValue.Value '-2147483648
                dMax = System.Data.SqlTypes.SqlInt32.MaxValue.Value '2147483647
            Case SqlDbType.SmallInt 'Đã kiểm tra
                typeNet = GetType(System.Int16)
                dMin = System.Data.SqlTypes.SqlInt16.MinValue.Value '-32768
                dMax = System.Data.SqlTypes.SqlInt16.MaxValue.Value '32767
        End Select
        Return typeNet
    End Function

    ''' <summary>
    ''' Chuyển dataType Decimal dưới SQL server sang DataType .Net Framework
    ''' </summary>
    ''' <param name="dMin">Trả về Giới hạn nhỏ nhất</param>
    ''' <param name="dMax">Trả về Giới hạn lớn nhất</param>
    ''' <return>DataType .Net Framework</return>
    ''' <remarks></remarks>
    Public Function GetProviderDataDecimal(ByRef dMin As Decimal, ByRef dMax As Decimal, Optional ByVal iPrecision As Integer = 18, Optional ByVal iScale As Integer = 8) As System.Type
        Dim sTemp As String = "".PadRight(iPrecision, "9"c)
        If iScale > 0 Then sTemp = sTemp.Insert(iPrecision - iScale, ".")
        dMin = -CDec(sTemp)
        dMax = CDec(sTemp)
        Return GetType(System.Decimal)
    End Function
#End Region

#Region "Các hàm Private"

    ''' <summary>
    ''' Khởi tạo cho control C1NumbericEdit dạng số
    ''' </summary>
    ''' <param name="C1NumericEdit">control C1NumbericEdit</param>
    ''' <param name="sCustomFormat">Optional. Default is N2. Định dạng số lẻ. Exp: N2 -> nhập 2 số lẻ</param>
    ''' <param name="bSign">Optional. Default is False. Cho phép nhập số âm (-)</param>
    ''' <remarks></remarks>
    Private Sub InitC1Numeric(ByRef C1NumericEdit As C1.Win.C1Input.C1NumericEdit, Optional ByVal sCustomFormat As String = "N2", Optional ByVal bSign As Boolean = False)
        C1NumericEdit.AcceptsTab = True
        'C1NumericEdit.ErrorInfo.ErrorMessageCaption = "ERROR"
        C1NumericEdit.ErrorInfo.ShowErrorMessage = False
        C1NumericEdit.ErrorInfo.ErrorAction = C1.Win.C1Input.ErrorActionEnum.None
        C1NumericEdit.EmptyAsNull = True 'Append 08/11/2011 => cho phép nhấn Delete để xóa giá trị
        ReFormatNumber(C1NumericEdit, sCustomFormat)

        C1NumericEdit.DisplayFormat.Inherit = CType((((C1.Win.C1Input.FormatInfoInheritFlags.NullText Or C1.Win.C1Input.FormatInfoInheritFlags.EmptyAsNull) _
                    Or C1.Win.C1Input.FormatInfoInheritFlags.TrimStart) _
                    Or C1.Win.C1Input.FormatInfoInheritFlags.TrimEnd), C1.Win.C1Input.FormatInfoInheritFlags)
        C1NumericEdit.EditFormat.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        C1NumericEdit.EditFormat.Inherit = CType(((((C1.Win.C1Input.FormatInfoInheritFlags.CustomFormat Or C1.Win.C1Input.FormatInfoInheritFlags.NullText) _
                    Or C1.Win.C1Input.FormatInfoInheritFlags.EmptyAsNull) _
                    Or C1.Win.C1Input.FormatInfoInheritFlags.TrimStart) _
                    Or C1.Win.C1Input.FormatInfoInheritFlags.TrimEnd), C1.Win.C1Input.FormatInfoInheritFlags)
        C1NumericEdit.Tag = Nothing
        C1NumericEdit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        C1NumericEdit.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None

        If Not bSign Then C1NumericEdit.NumericInputKeys = CType(((C1.Win.C1Input.NumericInputKeyFlags.Plus Or C1.Win.C1Input.NumericInputKeyFlags.[Decimal]) _
                                   Or C1.Win.C1Input.NumericInputKeyFlags.X), C1.Win.C1Input.NumericInputKeyFlags)

    End Sub

    ''' <summary>
    ''' Nhập số cho control C1NumbericEdit (có DataType dưới SQL khác Decimal)
    ''' </summary>
    ''' <param name="C1NumericEdit">control C1NumbericEdit</param>
    ''' <param name="SQLTypeName">DataType dưới SQL</param>
    ''' <param name="sCustomFormat">Định dạng số lẻ. Exp: N2 -> nhập 2 số lẻ</param>
    ''' <param name="bSign">Optional. Default is False. Cho phép nhập số âm (-)</param>
    ''' <remarks></remarks>
    Private Sub InputNumberInC1Numeric(ByRef C1NumericEdit As C1.Win.C1Input.C1NumericEdit, ByVal SQLTypeName As System.Data.SqlDbType, Optional ByVal sCustomFormat As String = "N2", Optional ByVal bSign As Boolean = False)
        InitC1Numeric(C1NumericEdit, sCustomFormat, bSign)

        Dim dMin As Double = -1, dMax As Double = -1
        C1NumericEdit.DataType = GetProviderData(SQLTypeName, dMin, dMax)
        If dMax = -1 And dMin = -1 Then Exit Sub
        C1NumericEdit.PostValidation.Intervals.Add(New C1.Win.C1Input.ValueInterval(dMin, dMax, True, True))
    End Sub

    ''' <summary>
    ''' Nhập số cho control C1NumbericEdit (có DataType dưới SQL = Decimal)
    ''' </summary>
    ''' <param name="C1NumericEdit">control C1NumbericEdit</param>
    ''' <param name="sCustomFormat">Định dạng số lẻ. Exp: N2 -> nhập 2 số lẻ</param>
    ''' <param name="iPrecision">Decimal(28,8)->truyền vào 28</param>
    ''' <param name="iScale">Decimal(28,8)->truyền vào 8</param>
    ''' <param name="bSign">Default is False. Cho phép nhập số âm (-)</param>
    ''' <remarks></remarks>
    Private Sub InputDecimalInC1Numeric(ByRef C1NumericEdit As C1.Win.C1Input.C1NumericEdit, Optional ByVal sCustomFormat As String = "N2", Optional ByVal iPrecision As Integer = 18, Optional ByVal iScale As Integer = 8, Optional ByVal bSign As Boolean = False)
        InitC1Numeric(C1NumericEdit, sCustomFormat, bSign)
        Dim dMin As Decimal = -1, dMax As Decimal = -1
        C1NumericEdit.DataType = GetProviderDataDecimal(dMin, dMax, iPrecision, iScale)
        If dMax = -1 And dMin = -1 Then Exit Sub
        C1NumericEdit.PostValidation.Intervals.Add(New C1.Win.C1Input.ValueInterval(dMin, dMax, True, True))
    End Sub

    '''' <summary>
    '''' Sự kiện keyPress của C1NumericEdit, chặn phím nhấn khi giá trị nhập vượt quá giới hạn cho phép
    '''' </summary>
    '''' <param name="C1NumericEdit"> đối tượng C1NumericEdit </param>
    'Dim ErrorProvider As New System.Windows.Forms.ErrorProvider
    Private Sub C1NumericEdit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim c1NumericEdit As C1.Win.C1Input.C1NumericEdit = CType(sender, C1.Win.C1Input.C1NumericEdit)
        c1NumericEdit.ValidateText()
        If (Not c1NumericEdit.Focused) Then c1NumericEdit.Parent.Focus()
    End Sub

#End Region

#Region "Format theo D91P9300"

    Public Structure StructureCustomFormat
        Public ExchangeRateDecimals As String
        Public BaseCurrencyID As String
        Public DecimalPlaces As String
        Public D90_ConvertedDecimals As String
        Public D07_QuantityDecimals As String
        Public D07_UnitCostDecimals As String
        Public DecimalSeparator As String
        Public ThousandSeparator As String
        Public D07_MaxOriginalDecimals As String
        Public UnitPriceDecimalPlaces As String
        Public D08_QuantityDecimals As String
        Public D08_UnitCostDecimals As String
        Public D08_RatioDecimals As String
        Public BOMQtyDecimals As String
        Public BOMPriceDecimals As String
        Public BOMAmtDecimals As String
        '************************
        Public iExchangeRateDecimals As Integer
        Public iDecimalPlaces As Integer
        Public iD90_ConvertedDecimals As Integer
        Public iD07_QuantityDecimals As Integer
        'Format cột đơn giá theo loại tiền
        Public iD07_UnitCostDecimals As Integer 'Mặc định
        Public iUnitPriceDecimalPlaces As Integer 'Theo yêu cầu
        '----------------------------------------
        Public iD07_MaxOriginalDecimals As Integer
        Public iD08_QuantityDecimals As Integer
        Public iD08_UnitCostDecimals As Integer
        Public iD08_RatioDecimals As Integer
        Public iBOMQtyDecimals As Integer
        Public iBOMPriceDecimals As Integer
        Public iBOMAmtDecimals As Integer

        Public iDefaultNumber0 As Integer
        Public iDefaultNumber1 As Integer
        Public iDefaultNumber2 As Integer
        Public iDefaultNumber3 As Integer
        Public iDefaultNumber4 As Integer
        Public iDefaultNumber5 As Integer
        '************************
        Public DefaultNumber0 As String
        Public DefaultNumber1 As String
        Public DefaultNumber2 As String
        Public DefaultNumber3 As String
        Public DefaultNumber4 As String
        Public DefaultNumber5 As String

        Public IsDependSpecbyInv As Byte 'Kiểm tra Quy cách theo mặt hàng
        Public LoadFormNotINV As Byte 'Kiểm tra nhập liệu mặt hàng hay tên hàng kiểu mới =1: nhập liệu dạng mới (ko có Dropdown Mã hàng)
        Public IsFillterName As Byte 'Update 16/06/2014 incident 64388 : Kiểm tra có Filter theo Tên hàng khi LoadFormNotINV =1
        Public SimType As Integer
        Public TimeFilterMode As Integer
        Public IsFilterNewFormat As Boolean
        Public IsUseDevExpressDocument As Boolean
    End Structure

    Public DxxFormat As StructureCustomFormat

    Public Sub LoadCustomFormat()
        Dim dt As DataTable = ReturnDataTable("Exec D91P9300 ")
        Try
            With DxxFormat
                .iDefaultNumber0 = 0
                .iDefaultNumber1 = 1
                .iDefaultNumber2 = 2
                .iDefaultNumber3 = 3
                .iDefaultNumber4 = 4
                .iDefaultNumber5 = 5

                .DefaultNumber0 = CustomFormat(0)
                .DefaultNumber1 = CustomFormat(1)
                .DefaultNumber2 = CustomFormat(2)
                .DefaultNumber3 = CustomFormat(3)
                .DefaultNumber4 = CustomFormat(4)
                .DefaultNumber5 = CustomFormat(5)

                .IsDependSpecbyInv = 0
                .LoadFormNotINV = 0
                .IsFillterName = 0
                .IsFilterNewFormat = True 'Mặc định Filter dạng mới
                .IsUseDevExpressDocument = False
                If dt.Rows.Count > 0 Then
                    .iBOMAmtDecimals = L3Int(dt.Rows(0).Item("BOMAmtDecimals"))
                    .iBOMPriceDecimals = L3Int(dt.Rows(0).Item("BOMPriceDecimals"))
                    .iBOMQtyDecimals = L3Int(dt.Rows(0).Item("BOMQtyDecimals"))
                    .iD07_MaxOriginalDecimals = L3Int(dt.Rows(0).Item("D07_MaxOriginalDecimals"))
                    .iD07_QuantityDecimals = L3Int(dt.Rows(0).Item("D07_QuantityDecimals"))
                    .iD07_UnitCostDecimals = L3Int(dt.Rows(0).Item("D07_UnitCostDecimals"))
                    .iD08_QuantityDecimals = L3Int(dt.Rows(0).Item("D08_QuantityDecimals"))
                    .iD08_RatioDecimals = L3Int(dt.Rows(0).Item("D08_RatioDecimals"))
                    .iD08_UnitCostDecimals = L3Int(dt.Rows(0).Item("D08_UnitCostDecimals"))
                    .iD90_ConvertedDecimals = L3Int(dt.Rows(0).Item("D90_ConvertedDecimals"))
                    .iDecimalPlaces = L3Int(dt.Rows(0).Item("DecimalPlaces"))
                    .iExchangeRateDecimals = L3Int(dt.Rows(0).Item("ExchangeRateDecimals"))
                    .iUnitPriceDecimalPlaces = L3Int(dt.Rows(0).Item("UnitPriceDecimalPlaces"))

                    .BaseCurrencyID = dt.Rows(0).Item("BaseCurrencyID").ToString
                    .ThousandSeparator = dt.Rows(0).Item("ThousandSeparator").ToString
                    .DecimalSeparator = dt.Rows(0).Item("DecimalSeparator").ToString

                    .IsDependSpecbyInv = L3Byte(dt.Rows(0).Item("IsDependSpecbyInv").ToString)
                    .LoadFormNotINV = L3Byte(dt.Rows(0).Item("LoadFormNotINV").ToString)
                    .IsFillterName = L3Byte(dt.Rows(0).Item("IsFillterName").ToString)
                    .SimType = L3Int(dt.Rows(0).Item("SimType"))
                    .TimeFilterMode = L3Int(dt.Rows(0).Item("TimeFilterMode"))
                    .IsFilterNewFormat = L3Bool(dt.Rows(0).Item("IsFilterNewFormat"))
                    .IsUseDevExpressDocument = L3Bool(dt.Rows(0).Item("IsUseDevExpressDocument"))
                End If
                .BOMAmtDecimals = CustomFormat(.iBOMAmtDecimals)
                .BOMPriceDecimals = CustomFormat(.iBOMPriceDecimals)
                .BOMQtyDecimals = CustomFormat(.iBOMQtyDecimals)
                .D07_MaxOriginalDecimals = CustomFormat(.iD07_MaxOriginalDecimals)
                .D07_QuantityDecimals = CustomFormat(.iD07_QuantityDecimals)
                .D07_UnitCostDecimals = CustomFormat(.iD07_UnitCostDecimals)
                .D08_QuantityDecimals = CustomFormat(.iD08_QuantityDecimals)
                .D08_RatioDecimals = CustomFormat(.iD08_RatioDecimals)
                .D08_UnitCostDecimals = CustomFormat(.iD08_UnitCostDecimals)
                .D90_ConvertedDecimals = CustomFormat(.iD90_ConvertedDecimals)
                .DecimalPlaces = CustomFormat(.iDecimalPlaces)
                .ExchangeRateDecimals = CustomFormat(.iExchangeRateDecimals)
                .UnitPriceDecimalPlaces = CustomFormat(.iUnitPriceDecimalPlaces)
            End With
        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try
    End Sub
#End Region

#Region "FooterText của lưới"
    Public Sub FooterSumNew(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray oColumns() As Integer)
        FooterSumFilter(tdbg, "", oColumns)
    End Sub

    Public Sub FooterSumNew(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray oColumns() As String)
        FooterSumFilter(tdbg, "", oColumns)
    End Sub


    Public Function CalSUMCol(ByVal dtGrid As DataTable, ByVal sFieldName As String, ByVal sNumberFormat As String, Optional ByVal sFilter As String = "") As String
        Dim dTotal As Double
        If sFieldName.Contains("[") Then
            dTotal = Number(dtGrid.Compute("SUM(" & sFieldName & ")", sFilter))
        Else
            dTotal = Number(dtGrid.Compute("SUM([" & sFieldName & "])", sFilter))
        End If


        Return FormatRoundNumber(dTotal, ReturnNumDigits(sNumberFormat)) 'Lấy số thập phân để làm tròn
    End Function


    Public Sub FooterSumFilter(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sFilter As String, ByVal ParamArray oColumns() As String)
        Try
            If tdbg.FilterBar = False Then tdbg.UpdateData()

            Dim dt As DataTable = CType(tdbg.DataSource, DataTable)
            If dt Is Nothing Then Exit Sub
            'ĐK lọc
            Dim strFilter As String = sFilter
            If strFilter <> "" And dt.DefaultView.RowFilter <> "" Then strFilter &= " And "
            strFilter &= dt.DefaultView.RowFilter
            '***************
            For j As Integer = 0 To oColumns.Length - 1
                If dt.Columns.Contains(oColumns(j)) = False Then Continue For
                Try
                    tdbg.Columns(oColumns(j)).FooterText = CalSUMCol(dt, oColumns(j), tdbg.Columns(oColumns(j)).NumberFormat, strFilter)
                    If tdbg.Columns(oColumns(j)).NumberFormat = "Percent" Or (tdbg.Columns(oColumns(j)).NumberFormat IsNot Nothing AndAlso tdbg.Columns(oColumns(j)).NumberFormat.Contains("%")) Then
                        tdbg.Columns(oColumns(j)).FooterText = SQLNumber(tdbg.Columns(oColumns(j)).FooterText, tdbg.Columns(oColumns(j)).NumberFormat).ToString
                    End If
                Catch ex As Exception
                    MessageBox.Show("Error FooterSumFilter: " & oColumns(j) & vbCrLf & ex.Message & vbCrLf & ex.Source)
                End Try
            Next
        Catch ex As Exception
            MessageBox.Show("Error FooterSumFilter: " & vbCrLf & ex.Message & vbCrLf & ex.Source)
        End Try
    End Sub

    '    Public Sub FooterSumFilter(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sFilter As String, ByVal ParamArray oColumns() As String)
    '        If tdbg.FilterBar = False Then tdbg.UpdateData()
    '
    '        Dim dt As DataTable = CType(tdbg.DataSource, DataTable)
    '        If dt Is Nothing Then Exit Sub
    '        'ĐK lọc
    '        Dim strFilter As String = sFilter
    '        If strFilter <> "" And dt.DefaultView.RowFilter <> "" Then strFilter &= " And "
    '        strFilter &= dt.DefaultView.RowFilter
    '        '***************
    '        For j As Integer = 0 To oColumns.Length - 1
    '            If dt.Columns.Contains(oColumns(j)) = False Then Continue For
    '            tdbg.Columns(oColumns(j)).FooterText = CalSUMCol(dt, oColumns(j), tdbg.Columns(oColumns(j)).NumberFormat, strFilter)
    '            If tdbg.Columns(oColumns(j)).NumberFormat = "Percent" Then
    '                tdbg.Columns(oColumns(j)).FooterText = tdbg.Columns(oColumns(j)).FooterText & "%"
    '            End If
    '        Next
    '    End Sub

    Public Sub FooterSumFilter(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sFilter As String, ByVal ParamArray iColumns() As Integer)
        Dim sColumns(iColumns.Length - 1) As String
        For i As Integer = 0 To iColumns.Length - 1
            sColumns.SetValue(tdbg.Columns(iColumns(i)).DataField, i)
        Next
        FooterSumFilter(tdbg, sFilter, sColumns)
    End Sub

    'Private Function CastObjToStrCol(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal oColumn As Object, ByRef sCol As String) As Boolean
    '    If oColumn Is Nothing Then Return False
    '    If IsNumeric(oColumn) Then
    '        sCol = tdbg.Columns(L3Int(oColumn)).DataField
    '    Else
    '        sCol = oColumn.ToString
    '    End If
    '    Return True
    'End Function

#Region "Set FormatTextEvent và Cal footer"
    Public Structure NumberFormatColumn
        Public FieldName As String
        Public NumberFormat As Integer  'Có thể dùng Integer hoặc String - Exp: N2 hoặc #,##0.00
        Public IsSum As Boolean 'Tính tổng hay không
    End Structure

    ''' <summary>
    ''' Total columns of grid by FormatText Event
    ''' </summary>
    ''' <param name="tdbg"></param>
    ''' <param name="sColumns">List of columns have construct</param>
    ''' <remarks>Exp: tham khảo D07F6040</remarks>

    Public Sub FooterSumNew(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray sColumns() As NumberFormatColumn)
        If tdbg.FilterBar = False Then tdbg.UpdateData()

        Dim dt As DataTable = CType(tdbg.DataSource, DataTable)
        If dt Is Nothing Then Exit Sub
        For j As Integer = 0 To sColumns.Length - 1
            If dt.Columns.Contains(sColumns(j).FieldName) = False Then Continue For
            Dim dTotal As Double = Number(dt.Compute("SUM(" & sColumns(j).FieldName & ")", dt.DefaultView.RowFilter))
            '*************
            'Lấy số thập phân để làm tròn
            Dim iNumDigits As Integer = sColumns(j).NumberFormat
            '**************************
            tdbg.Columns(sColumns(j).FieldName).FooterText = FormatRoundNumber(dTotal, iNumDigits)
        Next
    End Sub

    Private Function CreateTableFormatText() As DataTable
        Dim table As New DataTable
        Dim column As DataColumn
        ' Create new DataColumn, set DataType, ColumnName 
        ' and add to DataTable.    
        column = New DataColumn("FieldName", System.Type.GetType("System.String"))
        table.Columns.Add(column)
        ' Create second column.
        column = New DataColumn("NumberFormat", Type.GetType("System.String"))
        table.Columns.Add(column)
        ' Create third column.
        column = New DataColumn("IsSum", Type.GetType("System.Boolean"))
        table.Columns.Add(column)
        Return table
    End Function

    Public Sub AddColFormatText(ByRef dtFormat As DataTable, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal NumberFormat As String, ByVal sCol As String, ByVal IsSum As Boolean, Optional ByVal IsTag As Boolean = False)
        If dtFormat Is Nothing Then dtFormat = CreateTableFormatText()

        tdbg.Columns(sCol).NumberFormat = "FormatText Event" 'Gắn sự kiện cho lưới
        If IsTag Then tdbg.Columns(sCol).Tag = ReturnNumDigits(NumberFormat)
        Dim row As DataRow = dtFormat.NewRow
        row("FieldName") = sCol
        row("NumberFormat") = NumberFormat
        row("IsSum") = IsSum
        dtFormat.Rows.Add(row)
    End Sub

    Public Sub AddColFormatText(ByRef dtFormat As DataTable, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal NumberFormat As String, ByVal ParamArray iColumn() As Integer)
        For i As Integer = 0 To iColumn.Length - 1
            AddColFormatText(dtFormat, tdbg, NumberFormat, tdbg.Columns(iColumn(i)).DataField, True)
        Next
    End Sub

    Public Sub AddColFormatText(ByRef dtFormat As DataTable, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal NumberFormat As String, ByVal ParamArray sColumn() As String)
        For i As Integer = 0 To sColumn.Length - 1
            AddColFormatText(dtFormat, tdbg, NumberFormat, sColumn(i), True)
        Next
    End Sub

    Public Sub AddColFormatText(ByRef dtFormat As DataTable, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal NumberFormat As String, ByVal IsTag As Boolean, ByVal ParamArray sColumn() As String)
        For i As Integer = 0 To sColumn.Length - 1
            AddColFormatText(dtFormat, tdbg, NumberFormat, sColumn(i), True, IsTag)
        Next
    End Sub

    Public Sub AddColFormatText(ByRef dtFormat As DataTable, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal NumberFormat As String, ByVal IsTag As Boolean, ByVal ParamArray sColumn() As Integer)
        For i As Integer = 0 To sColumn.Length - 1
            AddColFormatText(dtFormat, tdbg, NumberFormat, tdbg.Columns(sColumn(i)).DataField, True, IsTag)
        Next
    End Sub

    Public Sub FooterSumNew(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dtFormat As DataTable, Optional ByVal sFilter As String = "")
        If dtFormat Is Nothing Then Exit Sub

        If tdbg.FilterBar = False Then tdbg.UpdateData()
        Dim dt As DataTable = CType(tdbg.DataSource, DataTable)
        If dt Is Nothing Then Exit Sub
        'ĐK lọc
        Dim strFilter As String = sFilter
        If strFilter <> "" And dt.DefaultView.RowFilter <> "" Then strFilter &= " And "
        strFilter &= dt.DefaultView.RowFilter
        '***************
        For j As Integer = 0 To dtFormat.Rows.Count - 1
            Dim sFieldName As String = dtFormat.Rows(j).Item("FieldName").ToString
            If dt.Columns.Contains(sFieldName) = False OrElse L3Bool(dtFormat.Rows(j).Item("IsSum")) = False Then Continue For
            tdbg.Columns(sFieldName).FooterText = CalSUMCol(dt, sFieldName, dtFormat.Rows(j).Item("NumberFormat").ToString, strFilter)
        Next
    End Sub

    Public Sub EventFormatText(ByVal dtFormat As DataTable, ByRef e As C1.Win.C1TrueDBGrid.FormatTextEventArgs)
        If dtFormat Is Nothing OrElse dtFormat.Rows.Count = 0 Then Exit Sub

        If Number(e.Value) = 0 Then
            e.Value = ""
        Else
            Dim dr() As DataRow = dtFormat.Select("FieldName=" & SQLString(e.Column.DataField))
            If dr.Length = 0 Then Exit Sub
            e.Value = SQLNumber(e.Value, dr(0).Item("NumberFormat").ToString)
        End If
    End Sub
#End Region
#End Region
End Module

