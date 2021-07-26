Public Class SQLFinance
    Public Shared Function SQLCurrencyID(ByVal sDate As String, ByVal UnionAll As EnumUnionAll, ByVal sValueEdit As String) As String
        Dim sUnicode As String = ""
        Dim sAll As String = ""
        UnicodeAllString(sUnicode, sAll, gbUnicode)

        Dim sExchangeRate As String = "ExchangeRate"
        If sDate <> "" Then
            sExchangeRate = "ISNULL(( SELECT TOP 1 D01.ExchangeRate FROM D01T0060 D01 WITH(NOLOCK) WHERE (D01.CurrencyID = D91.CurrencyID) "
            sExchangeRate &= "AND D01.HisCurrencyDate <= " & SQLDateSave(sDate) & " "
            sExchangeRate &= " ORDER BY HisCurrencyDate DESC), ExchangeRate)"
        End If
        Dim sSQL As String = "--Do nguon cho loai tien" & vbCrLf
        sSQL &= "Select CurrencyID,  CurrencyName" & sUnicode & " As CurrencyName, " & sExchangeRate & " AS ExchangeRate, Operator, MethodID, OriginalDecimal, ExchangeRateDecimal, UnitPriceDecimals, ConvertedDecimal "
        sSQL &= ", 1 AS DisplayOrder " & vbCrLf
        sSQL &= "From D91V0010 D91" & vbCrLf
        sSQL &= "Where Disabled =0 "
        If sValueEdit <> "" Then sSQL &= " Or CurrencyID=" & SQLString(sValueEdit)
        sSQL &= vbCrLf
        If UnionAll = EnumUnionAll.All Then
            sSQL &= "Union All" & vbCrLf
            sSQL &= "Select '%' as CurrencyID, " & sAll & " As CurrencyName, 1 as ExchangeRate, 0 as Operator, '' as MethodID, 0 as OriginalDecimal, 0 as  ExchangeRateDecimal, 0 as UnitPriceDecimals, 0 as ConvertedDecimal" & vbCrLf
            sSQL &= ", 0 AS DisplayOrder " & vbCrLf
        End If
        sSQL &= "Order By DisplayOrder, CurrencyID "

        Return sSQL
    End Function


    Public Shared Function SQLWareHouseID(ByVal union As EnumUnionAll, ByVal sValueEdit As String, ByVal sWhere As String) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon Kho hang" & vbCrLf)
        If union = EnumUnionAll.All Then
            sSQL &= "Select " & AllCode & " as WareHouseID, " & AllName & " as WareHouseName, '%' as DivisionID, '%' as ObjectTypeID, '%' as ObjectID, 0 as DisplayOrder " & vbCrLf
            sSQL &= "Union All" & vbCrLf
        ElseIf union = EnumUnionAll.AddNew Then
            sSQL &= "Select " & NewCode & " as WareHouseID, " & NewName & " as WareHouseName, '%' as DivisionID, '%' as ObjectTypeID, '%' as ObjectID, 0 as DisplayOrder " & vbCrLf
            sSQL &= "Union All" & vbCrLf
        End If
        sSQL &= "Select WareHouseID, WareHouseName" & UnicodeJoin(gbUnicode) & " as WareHouseName, DivisionID, ObjectTypeID, ObjectID, 1 as DisplayOrder " & vbCrLf
        sSQL &= "From D07T0007  WITH(NOLOCK)" & vbCrLf
        sSQL &= "Where (( DAGroupID = '' OR DAGroupID IN (SELECT DAGroupID From LEMONSYS.DBO.D00V0080  WHERE UserID = " & SQLString(gsUserID) & ") OR 'LEMONADMIN' =" & SQLString(gsUserID) & ")" & vbCrLf
        If sWhere <> "" Then sSQL &= " And " & sWhere
        sSQL &= ")"
        If sValueEdit <> "" Then sSQL &= " Or WareHouseID =" & SQLString(sValueEdit)
        sSQL &= vbCrLf
        sSQL &= "Order by DisplayOrder, WareHouseID"
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P9302
    '# Created User: Nguyễn Thị Minh Hòa
    '# Created Date: 25/09/2015 04:55:02
    '#---------------------------------------------------------------------------------------------------
    Public Shared Function SQLStoreD91P9302(ByVal sCommand As String, ByVal sWhere As String, ByVal IsAll As Byte, ByVal sType As String, ByVal sInventoryID As String, ByVal sEditAccountID As String, Optional ByVal sModuleID As String = "", Optional ByVal sFormID As String = "") As String
        'Biến ModuleID và FormID không thấy sử dụng
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon chuan cho Tai khoan " & sCommand & vbCrLf)
        sSQL &= "Exec D91P9302 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[2], NOT NULL
        sSQL &= SQLString(sModuleID) & COMMA 'ModuleID, varchar[20], NOT NULL
        sSQL &= SQLString(sFormID) & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLString(sWhere) & COMMA 'sWhere, varchar[8000], NOT NULL
        sSQL &= SQLNumber(IsAll) & COMMA 'IsAll, tinyint, NOT NULL
        sSQL &= SQLString(sType) & COMMA 'Type, varchar[20], NOT NULL
        sSQL &= SQLString(sInventoryID) & COMMA 'InventoryID, varchar[50], NOT NULL
        sSQL &= SQLString(sEditAccountID) 'AccountID, varchar[50], NOT NULL
        Return sSQL
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P9303
    '# Created User: 
    '# Created Date: 20/10/2015 11:26:08
    '#---------------------------------------------------------------------------------------------------
    Public Shared Function SQLStoreD91P9303(ByVal sWhere As String, ByVal IsAll As Byte, Optional ByVal sModuleID As String = "", Optional ByVal sFormID As String = "") As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon cho Tai khoan bao cao" & vbCrLf)
        sSQL &= "Exec D91P9303 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[2], NOT NULL
        sSQL &= SQLString(sModuleID) & COMMA 'ModuleID, varchar[20], NOT NULL
        sSQL &= SQLString(sFormID) & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLString(sWhere) & COMMA 'sWhere, varchar[8000], NOT NULL
        sSQL &= SQLNumber(IsAll)  'IsAll, tinyint, NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P9103
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 13/08/2018 08:59:05
    '#---------------------------------------------------------------------------------------------------
    Public Shared Function SQLStoreD91P9103(ByVal sTableName As String, _
                        ByVal sBatchID As String, _
                        ByVal sRefNo As String, _
                        ByVal sSerialNo As String, _
                        ByVal sVoucherID As String, ByVal sModuleID As String, ByVal ObjectTypeID As String, ByVal ObjectID As String) As String
        Return SQLStoreD91P9103(sTableName, sBatchID, sRefNo, sSerialNo, sVoucherID, sModuleID, ObjectTypeID, ObjectID, "")
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P9103
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 13/08/2018 08:59:05
    '#---------------------------------------------------------------------------------------------------
    Public Shared Function SQLStoreD91P9103(ByVal sTableName As String, _
        ByVal sBatchID As String, _
        ByVal sRefNo As String, _
        ByVal sSerialNo As String, _
       ByVal sVoucherID As String, ByVal sModuleID As String, ByVal ObjectTypeID As String, ByVal ObjectID As String, sVATNo As String) As String
        Return SQLStoreD91P9103(sTableName, sBatchID, sRefNo, sSerialNo, sVoucherID, sModuleID, ObjectTypeID, ObjectID, sVATNo, "")
    End Function

    'bổ sung thêm 31/07/2019-id 122373
    Public Shared Function SQLStoreD91P9103(ByVal sTableName As String, _
        ByVal sBatchID As String, _
        ByVal sRefNo As String, _
        ByVal sSerialNo As String, _
        ByVal sVoucherID As String, ByVal sModuleID As String, ByVal ObjectTypeID As String, ByVal ObjectID As String, sVATNo As String, sTranTypeID As String) As String
        Return SQLStoreD91P9103(sTableName, sBatchID, sRefNo, sSerialNo, sVoucherID, sModuleID, ObjectTypeID, ObjectID, sVATNo, sTranTypeID, "", "")
    End Function

    Public Shared Function SQLStoreD91P9103(ByVal sTableName As String, _
      ByVal sBatchID As String, _
      ByVal sRefNo As String, _
      ByVal sSerialNo As String, _
      ByVal sVoucherID As String, ByVal sModuleID As String, ByVal ObjectTypeID As String, ByVal ObjectID As String, sVATNo As String, sTranTypeID As String, ByVal VATObjectTypeID As String, ByVal VATObjectID As String) As String
        sModuleID = L3Right(sModuleID, 2)
        Dim sSQL As String = ""
        sSQL &= ("-- Kiem tra so seri trung so hd" & vbCrLf)
        sSQL &= "Exec D91P9103 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(sModuleID) & COMMA 'ModuleID, varchar[20], NOT NULL
        sSQL &= SQLString(sTableName) & COMMA 'TableName, varchar[20], NOT NULL
        sSQL &= SQLString(sBatchID) & COMMA 'BatchID, varchar[20], NOT NULL
        sSQL &= SQLString(sRefNo) & COMMA 'RefNo, varchar[100], NOT NULL
        sSQL &= SQLString(sSerialNo) & COMMA 'serialNo, varchar[100], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLString(sVoucherID) & COMMA 'VoucherID, varchar[20], NOT NULL
        sSQL &= SQLString(ObjectID) & COMMA 'ObjectID, varchar[50], NOT NULL
        sSQL &= SQLString(ObjectTypeID) & COMMA 'ObjectTypeID, varchar[50], NOT NULL
        sSQL &= SQLString(sVATNo) & COMMA 'VATNo, varchar[50], NOT NULL
        'bổ sung thêm 31/07/2019-id 122373
        sSQL &= SQLString(sTranTypeID) & COMMA 'TranTypeID, varchar[50], NOT NULL
        sSQL &= SQLNumber(0) & COMMA 'Mode, tinyint, NOT NULL
        sSQL &= SQLNumber(0) & COMMA 'OrderNum, int, NOT NULL
        sSQL &= SQLString(VATObjectTypeID) & COMMA 'ObjectID, varchar[50], NOT NULL
        sSQL &= SQLString(VATObjectID)  'ObjectTypeID, varchar[50], NOT NULL
        Return sSQL
    End Function


End Class
