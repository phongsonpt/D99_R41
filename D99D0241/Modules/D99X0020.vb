Imports System.Threading
Imports System.Reflection
Imports System
Imports Lemon3
''' <summary>
''' Module này dùng để khai báo các Sub và Function toàn cục liên quan gọi DLL
''' </summary>
''' <remarks>
''' ở các module D99Xxxxx
''' </remarks>
''' 

Public Module D99X0020
    Public Structure StructureProperties
        Public properties As String
        Public value As Object
    End Structure

#Region "DLL"
    'iStatus: 0 - Show; 1- ShowDialog; 2- Thread
    Public Sub SetProperties(ByRef arr() As StructureProperties, ByVal properties As String, ByVal value As Object)
        Try
            If arr Is Nothing Then
                ReDim arr(0)
            Else
                ReDim Preserve arr(arr.Length)
            End If
            arr(arr.Length - 1).properties = properties
            arr(arr.Length - 1).value = value
        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try
    End Sub

    Dim FormID As String = "" 'Dùng để lưu file LogParaDLL

    Public Sub SetVariable()
        '   If L3.IsLoaded Then Exit Sub
        GeneralItems()
        L3.ConnectionString = gsConnectionString
        L3.Closed = gbClosed
        L3.CompanyID = gsCompanyID
        L3.ConnectionUser = gsConnectionUser
        L3.DivisionID = gsDivisionID
        L3.Password = gsPassword
        L3.Server = gsServer
        L3.TranMonth = giTranMonth
        L3.TranYear = giTranYear
        L3.UserID = gsUserID
        L3.Language = geLanguage
        L3.STRLanguage = gsLanguage
        L3.IsUniCode = gbUnicode
        L3.PrintVNI = gbPrintVNI
        L3.DisplayProduct = geDisplayProduct
        L3.IsDesktop = gbIsDesktop
        L3.IsCallDesktop = gbCallDesktop
        L3.Grid_BackColor_HighLigh = GRID_BACK_COLOR_HighLigh
        L3.SizeTextD00 = giSizeTextD00
        L3.IsAdjustResolution = gbAdjustResolution
        L3.Grid_Back_Color = GRID_BACK_COLOR
        L3.Grid_Fore_Color = GRID_FORE_COLOR
        L3.Grid_Row_Height = GRID_ROW_HEIGHT
        L3.IsSetupPerson = gbIsSetupPerson
        L3.Color_BackColorObliratory = COLOR_BACKCOLOROBLIGATORY
        L3.CreateBy = gsCreateBy
        L3.IsUseD54 = gbUseD54
        L3.Color_BackColorWarning = COLOR_BACKCOLORWARNING
        L3.IsLockL3UserID = gbLockL3UserID
        L3.ApplicationPath = gsApplicationPath
        L3.ApplicationSetup = gsApplicationSetup
        L3.CDNAttMode = giCDNAttMode
        L3.CDNPrivateLink = gsCDNPrivateLink
        L3.ErpApiSecret = gsErpApiSecret
        L3.ErpApiURL = gsErpApiURL
        L3.CDNApiSecret = gsCDNApiSecret
        L3.NewCode = NewCode
        L3.NewName = NewName
        L3.AllName = AllName
        L3.AllCode = AllCode
        L3.SocketServerURL = gsSocketServerURL
        L3.NotifyAPISecret = gsNotifyAPISecret
        L3.NotifyApiURL = gsNotifyApiURL
        L3.IsLoaded = True
    End Sub
    Public Sub SetProperties(ByVal frm As Form, ByVal properties As String, ByVal value As Object)
        SetProperties(CType(frm, Object), properties, value)
        'Try
        '    Dim info As System.Reflection.PropertyInfo = frm.GetType.GetProperty(properties.Trim)
        '    If info IsNot Nothing Then
        '        info.SetValue(frm, value, Nothing)
        '        If FormID <> frm.Name Then
        '            WriteLogFile("The parameters of " & frm.Name & " (" & frm.GetType.Namespace & ") :", "LogParaDLL.log", False)
        '            '  SetVariable()
        '        End If
        '        WriteLogFile("'" & properties & "' = '" & L3String(value) & "'", "LogParaDLL.log", True)
        '        FormID = frm.Name
        '    Else 'Vi phạm tham số
        '        WriteLogFile("'" & properties & "' of " & frm.Name & " (" & frm.GetType.Namespace & "): doesn't exist.", "LogParaDLL.log", True)
        '    End If

        'Catch ex As Exception
        '    D99C0008.Msg("DLL: " & ex.Message)
        'End Try
    End Sub

    Public Function GetProperties(ByVal frm As Form, ByVal properties As String) As Object
        Return GetProperties(CType(frm, Object), properties)
    End Function

    Public Function GetProperties(ByVal ctrl As Control, ByVal properties As String) As Object
        Return GetProperties(CType(ctrl, Object), properties)
    End Function

    Public Function GetProperties(ByVal ctrl As Object, ByVal properties As String) As Object
        If ctrl Is Nothing Then Return Nothing
        Try
            Dim info As System.Reflection.PropertyInfo = ctrl.GetType.GetProperty(properties)
            If info Is Nothing Then Return Nothing
            Return info.GetValue(ctrl, Nothing)
        Catch ex As Exception
            D99C0008.Msg(ex.Message)
        End Try
        Return Nothing
    End Function

    'iStatus: 0 - Show; 1- ShowDialog; 2- Thread
    Private Function CallForms(ByVal frmCall As Form, ByVal sDLLName As String, ByVal sFormActive As String, ByVal iStatus As Integer, ByVal arrPro() As StructureProperties) As Form
        If sFormActive = "" Then
            D99C0008.Msg(sFormActive & " invalid")
            Return Nothing
        End If
        Select Case sFormActive
            Case "D99F2222", "D89F3000" 'Xuất Excel bổ sung thêm D89F3000 21/11/2016 vì form này được gọi nhiều lần
            Case Else 'Không cần kiểm tra
                If CheckExistForm(sFormActive) Then Return Nothing
        End Select
        Try


            Dim file As String = (New System.Uri(Assembly.GetExecutingAssembly().CodeBase)).LocalPath
            Dim dirName As String = New System.IO.DirectoryInfo(file).Parent.FullName
            'MessageBox.Show("B1: " & dirName & "\" & sDLLName & ".dll")
            Dim ass As Reflection.Assembly = Reflection.Assembly.LoadFrom(dirName & "\" & sDLLName & ".dll")

            'MessageBox.Show("B2: " & dirName & "\" & sDLLName & ".dll")
            'Dim frmAss As Reflection.Assembly
            'frmAss = System.Reflection.Assembly.LoadFrom(dirName & "\" & sDLLName & ".dll")

            'MessageBox.Show("B3: ")

            Dim FormInstanceType = ass.GetType(sDLLName & "." & sFormActive)
            Dim frm As Form = Nothing
            Try
                frm = CType(Activator.CreateInstance(FormInstanceType), Form)
            Catch ex As Exception
                If ex.Message.Contains("CreateInstance") Then
                    D99C0008.MsgL3("Not exist form " & sFormActive & " in " & sDLLName & " dll", L3MessageBoxIcon.Exclamation)
                Else
                    D99C0008.MsgL3(ex.Message, L3MessageBoxIcon.Exclamation)
                End If
                Return Nothing
            End Try

            Select Case sFormActive
                Case "D99F2222", "D89F3000" 'Xuất Excel
                Case Else 'Không cần kiểm tra
                    frm.ShowInTaskbar = True
            End Select
            If arrPro IsNot Nothing Then 'Gắn giá trị tham số WriteOnly
                For i As Integer = 0 To arrPro.Length - 1
                    SetProperties(frm, arrPro(i).properties, arrPro(i).value)
                Next
            Else
                WriteLogFile(frm.Name & " (" & frm.GetType.Namespace & ") : No parameter", "LogParaDLL.log", False)
                '               SetVariable()
            End If
            Select Case iStatus
                Case 0
                    If frmCall IsNot Nothing Then frm.Owner = frmCall
                    frm.Show()
                Case 1
                    frm.ShowDialog()
                    frm.Dispose()
                Case 2 'Thread
                    Control.CheckForIllegalCrossThreadCalls = False
                    Dim thr As Thread = New Thread(AddressOf CreateThread)
                    thr.SetApartmentState(ApartmentState.STA) 'Để test
                    thr.IsBackground = True
                    thr.Start(frm)
            End Select
            Return frm
        Catch ex As Exception
            If System.IO.File.Exists(gsApplicationSetup & "\" & sDLLName & ".dll") = False Then
                D99C0008.MsgL3("Not exist file " & gsApplicationSetup & "\" & sDLLName & ".dll", L3MessageBoxIcon.Exclamation)
            Else
                D99C0008.Msg(ex.Message)
            End If
        End Try
        Return Nothing
    End Function

    Private Sub CreateThread(ByVal frm As Object)
        Dim f As Form = DirectCast(frm, Form)
        f.ShowDialog()
        '   f.Dispose()
    End Sub

    'Status =0: Show
    Public Function CallFormShow(ByVal frmCall As Form, ByVal sDLLName As String, ByVal sFormActive As String) As Form
        Return CallForms(frmCall, sDLLName, sFormActive, 0, Nothing)
    End Function

    Public Function CallFormShow(ByVal sDLLName As String, ByVal sFormActive As String) As Form
        Return CallForms(Nothing, sDLLName, sFormActive, 0, Nothing)
    End Function

    Public Function CallFormShow(ByVal frmCall As Form, ByVal sDLLName As String, ByVal sFormActive As String, ByVal arrPro() As StructureProperties) As Form
        Return CallForms(frmCall, sDLLName, sFormActive, 0, arrPro)
    End Function

    Public Function CallFormShow(ByVal sDLLName As String, ByVal sFormActive As String, ByVal arrPro() As StructureProperties) As Form
        Return CallForms(Nothing, sDLLName, sFormActive, 0, arrPro)
    End Function

    'Status =1: ShowDialog
    Public Function CallFormShowDialog(ByVal sDLLName As String, ByVal sFormActive As String) As Form
        Return CallForms(Nothing, sDLLName, sFormActive, 1, Nothing)
    End Function


    Public Function CallFormShowDialog(ByVal sDLLName As String, ByVal sFormActive As String, ByVal arrPro() As StructureProperties) As Form
        Return CallForms(Nothing, sDLLName, sFormActive, 1, arrPro)
    End Function

    'Status =2: Thread
    Public Function CallFormThread(ByVal frmCall As Form, ByVal sDLLName As String, ByVal sFormActive As String) As Form
        'Return CallForms(frmCall, sDLLName, sFormActive, 2, Nothing) 'Gọi giống Show
        Return CallFormShow(frmCall, sDLLName, sFormActive)
    End Function

    Public Function CallFormThread(ByVal sDLLName As String, ByVal sFormActive As String) As Form
        ' Return CallForms(Nothing, sDLLName, sFormActive, 2, Nothing) 'Gọi giống Show
        Return CallFormShow(sDLLName, sFormActive)
    End Function

    Public Function CallFormThread(ByVal frmCall As Form, ByVal sDLLName As String, ByVal sFormActive As String, ByVal arrPro() As StructureProperties) As Form
        'Return CallForms(frmCall, sDLLName, sFormActive, 2, arrPro) 'Gọi giống Show: Thread : 2
        Return CallFormShow(frmCall, sDLLName, sFormActive, arrPro)
    End Function

    Public Function CallFormThread(ByVal sDLLName As String, ByVal sFormActive As String, ByVal arrPro() As StructureProperties) As Form
        'Return CallForms(Nothing, sDLLName, sFormActive, 2, arrPro) 'Gọi giống Show: Thread : 2
        Return CallFormShow(sDLLName, sFormActive, arrPro)
    End Function


    Public Function CallFormThreadNew(ByVal frmCall As Form, ByVal sDLLName As String, ByVal sFormActive As String) As Form
        Return CallForms(frmCall, sDLLName, sFormActive, 2, Nothing) 'Gọi giống Show
        ' Return CallFormShow(frmCall, sDLLName, sFormActive)
    End Function

    Public Function CallFormThreadNew(ByVal sDLLName As String, ByVal sFormActive As String) As Form
        Return CallForms(Nothing, sDLLName, sFormActive, 2, Nothing) 'Gọi giống Show
        'Return CallFormShow(sDLLName, sFormActive)
    End Function

    Public Function CallFormThreadNew(ByVal frmCall As Form, ByVal sDLLName As String, ByVal sFormActive As String, ByVal arrPro() As StructureProperties) As Form
        Return CallForms(frmCall, sDLLName, sFormActive, 2, arrPro) 'Gọi giống Show: Thread : 2
        'Return CallFormShow(frmCall, sDLLName, sFormActive, arrPro)
    End Function

    Public Function CallFormThreadNew(ByVal sDLLName As String, ByVal sFormActive As String, ByVal arrPro() As StructureProperties) As Form
        Return CallForms(Nothing, sDLLName, sFormActive, 2, arrPro) 'Gọi giống Show: Thread : 2
        'Return CallFormShow(sDLLName, sFormActive, arrPro)
    End Function
#Region "Call D99F2222"
    Public Sub CallShowD99F2222(ByVal frmCall As Form, ByVal dtCaptionExcel As DataTable, ByVal dtExportTable As DataTable, ByVal arrGroupColumns() As String)
        Dim dtTemp As DataTable = Nothing
        CallShowD99F2222(frmCall, dtCaptionExcel, dtExportTable, arrGroupColumns, dtTemp)
    End Sub

    Private Function InitTable() As DataTable
        Dim dtM As New DataTable
        dtM.Columns.Add("Description")
        dtM.Columns.Add("FieldExcel")
        dtM.Columns.Add("Value")
        dtM.Columns.Add("TabIndex", Type.GetType("System.Int64"))
        Return dtM
    End Function

    Public Sub AddDataRowTable(ByRef dtM As DataTable, ByVal text As String, ByVal sField As String, ByVal value As String, ByVal tabIndex As Integer)
        If dtM Is Nothing OrElse dtM.Columns.Count = 0 Then dtM = InitTable()
        Dim dr As DataRow = dtM.NewRow
        dr("Description") = text
        dr("FieldExcel") = "<#" & sField & "#>"
        dr("Value") = value
        dr("TabIndex") = tabIndex
        dtM.Rows.Add(dr)
    End Sub

    Public Sub AddDataRowTable(ByRef dtM As DataTable, ByVal text As String, ByVal ctrl As Control)
        Dim sValue As String = ctrl.Text
        Dim sField As String = ""
        GetInfoCtrl(ctrl, sValue, sField)
        AddDataRowTable(dtM, text, sField, sValue, ctrl.TabIndex)
    End Sub

    Public Sub AddDataRowTable(ByRef dtM As DataTable, ByVal lbl As Label, ByVal value As String)
        AddDataRowTable(dtM, lbl.Text, lbl.Name.Replace("lbl", ""), value, lbl.TabIndex)
    End Sub
    Public Sub CallShowD99F2222(ByVal frmCall As Form, ByVal dtCaptionExcel As DataTable, ByVal dtExportTable As DataTable, ByVal arrGroupColumns() As String, ByVal dtMaster As DataTable)
        If Not AllOpenForm(frmCall.Name, "D99F2222", "FormID") Then Exit Sub
        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "UseUnicode", gbUnicode)
        SetProperties(arrPro, "FormID", frmCall.Name)
        SetProperties(arrPro, "dtLoadGrid", dtCaptionExcel)
        SetProperties(arrPro, "GroupColumns", arrGroupColumns)
        SetProperties(arrPro, "dtExportTable", dtExportTable)
        SetProperties(arrPro, "dtMaster", dtMaster)
        CallFormShow(frmCall, "D99D0341", "D99F2222", arrPro)
    End Sub

    Private Sub GetInfoCtrl(ByVal ctrl As Control, ByRef sValue As String, ByRef sField As String)
        sValue = ctrl.Text
        Dim refix As String = Strings.Left(ctrl.Name, 3)
        If TypeOf (ctrl) Is C1.Win.C1List.C1Combo Then
            sValue = L3String(GetProperties(ctrl, "SelectedValue"))
            refix = "tdbc"
        ElseIf TypeOf (ctrl) Is C1.Win.C1Input.C1DateEdit OrElse TypeOf (ctrl) Is C1.Win.C1Input.C1NumericEdit Then
            sValue = L3String(GetProperties(ctrl, "Value"))
            If TypeOf (ctrl) Is C1.Win.C1Input.C1DateEdit Then refix = "c1date"
        ElseIf TypeOf (ctrl) Is DevExpress.XtraEditors.DateEdit Then
            sValue = L3String(GetProperties(ctrl, "EditValue"))
            refix = "devdate"
        End If
        sField = ctrl.Name.Replace(refix, "")
    End Sub

    Public Function AddDataRowTable(ByRef dtM As DataTable, ByVal ctrl As Control, ByVal bTabControl As Boolean) As Boolean
        If TypeOf (ctrl) Is Label OrElse TypeOf (ctrl) Is CheckBox OrElse TypeOf (ctrl) Is RadioButton OrElse TypeOf (ctrl) Is Button Then 'Không kiểm các control 
            Return True
        ElseIf ctrl.Controls.Count = 0 OrElse TypeOf (ctrl) Is C1.Win.C1List.C1Combo OrElse TypeOf (ctrl) Is C1.Win.C1Input.C1DateEdit OrElse TypeOf (ctrl) Is C1.Win.C1Input.C1NumericEdit OrElse TypeOf (ctrl) Is DevExpress.XtraEditors.DateEdit Then
            If bTabControl = False AndAlso ctrl.Visible = False Then Return True 'TH tab control không xét ẩn. Ngược lại ẩn thì không thêm vào
            Dim sValue As String = ctrl.Text
            Dim sField As String = ""
            GetInfoCtrl(ctrl, sValue, sField)
            'Tìm control đi kèm là label, chekbox và radiobutton
            Dim sCaption As String = GetCaption(ctrl, sField, "lbl")
            If sCaption = "" Then sCaption = GetCaption(ctrl, sField, "opt")
            If sCaption = "" Then sCaption = GetCaption(ctrl, sField, "chk")
            '******************
            AddDataRowTable(dtM, sCaption, sField, sValue, ctrl.TabIndex)
            Return True
        ElseIf TypeOf (ctrl.Parent) Is TabControl Then
            bTabControl = True
        End If
        For i As Integer = 0 To ctrl.Controls.Count - 1
            AddDataRowTable(dtM, ctrl.Controls(i), bTabControl)
        Next
        Return False
    End Function

    Private Function GetCaption(ByVal ctrl As Control, ByVal sField As String, ByVal refix As String) As String
        Dim ctrlTemp() As Control = ctrl.Parent.Controls.Find(refix & sField, True)
        If ctrlTemp.Length = 0 Then Return ""
        Return ctrlTemp(0).Text
    End Function

    '    Private Sub AddDataRowTables(ByRef dtM As DataTable, ByVal ctrl As Control)
    '        Dim bTabControl As Boolean = TypeOf (ctrl.Parent) Is TabControl
    '        If AddDataRowTable(dtM, ctrl, bTabControl) Then Exit Sub
    '        For Each control As Control In ctrl.Controls
    '            If AddDataRowTable(dtM, control, bTabControl) Then Continue For
    '        Next
    '    End Sub

    Public Sub CallShowD99F2222(ByVal frmCall As Form, ByVal dtCaptionExcel As DataTable, ByVal dtExportTable As DataTable, ByVal arrGroupColumns() As String, ByVal ctrlContain As Control)
        Dim dtMaster As New DataTable
        AddDataRowTable(dtMaster, ctrlContain, TypeOf (ctrlContain.Parent) Is TabControl)
        CallShowD99F2222(frmCall, dtCaptionExcel, dtExportTable, arrGroupColumns, dtMaster)
    End Sub
#End Region

    Public Function CallShowDialogD80F2090(ByVal sModuleID As String, ByVal sFormIDPermission As String, ByVal sTransTypeID As String) As Boolean
        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "ModuleID", sModuleID)
        SetProperties(arrPro, "FormIDPermission", sFormIDPermission)
        SetProperties(arrPro, "TransTypeID", sTransTypeID)
        SetProperties(arrPro, "sFont", IIf(gbUnicode, "UNICODE", "VNI").ToString)
        Dim frm As Form = CallFormShowDialog("D80D0040", "D80F2090", arrPro)
        Return L3Bool(GetProperties(frm, "bSaveOK"))
    End Function

    Public Function CallShowDialogD80F2090(ByVal sModuleID As String, ByVal sFormIDPermission As String, ByVal sTransTypeID As String, ByVal dtData As DataTable) As Boolean
        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "ModuleID", sModuleID)
        SetProperties(arrPro, "FormIDPermission", sFormIDPermission)
        SetProperties(arrPro, "TransTypeID", sTransTypeID)
        SetProperties(arrPro, "dtData", dtData)
        SetProperties(arrPro, "sFont", IIf(gbUnicode, "UNICODE", "VNI").ToString)
        Dim frm As Form = CallFormShowDialog("D80D0040", "D80F2090", arrPro)
        Return L3Bool(GetProperties(frm, "bSaveOK"))
    End Function

    Public Function CallFormD91F6010(ByVal InListID As String) As String
        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "InListID", InListID)
        Dim frm As Form = CallFormShowDialog("D91D0240", "D91F6010", arrPro)
        Return GetProperties(frm, "Output01").ToString
    End Function


    Public Function CheckExistForm(ByVal sFormName As String) As Boolean
        Dim f As Form = CheckExist(sFormName)
        If f IsNot Nothing Then
            f.Activate()
            f.WindowState = FormWindowState.Normal
            Return True
        End If
        Return False
    End Function

    Private Function CheckExist(ByVal sFormName As String) As Form
        Return Application.OpenForms.Item(sFormName)
    End Function

    Public Function GetChildFormState(ByVal formParent As Form) As EnumFormState
        Dim formChild As Form = Nothing
        If formParent.OwnedForms.Length > 0 Then formChild = formParent.OwnedForms(0)
        If formChild IsNot Nothing Then Return CType(GetProperties(formChild, "FormState"), EnumFormState)
    End Function

    Public Function AllOpenForm(ByVal sFormParent As String, ByVal sCheckForm As String, ByVal properties As String) As Boolean
        Dim frmCollection As FormCollection = Application.OpenForms
        For Each frm As Form In frmCollection
            If frm.Name = sCheckForm Then
                Dim sFm As Object = GetProperties(frm, properties)
                If sFm Is Nothing Then Continue For
                Dim sName As String = ""
                If sFm.GetType().Name.Contains(".Form") Then
                    sName = CType(sFm, Form).Name
                Else
                    sName = sFm.ToString
                End If
                If sName.ToString.Contains(sFormParent) Then frm.Activate() : frm.WindowState = FormWindowState.Normal : Return False
            End If
        Next
        Return True
    End Function

    Private Function CallWPF(ByVal frmCall As Form, ByVal sDll As String, ByVal FormActive As String, showDialog As Boolean, arrPro() As StructureProperties) As Object
        If FormActive = "" Then
            D99C0008.Msg(FormActive & " invalid")
            Return Nothing
        End If
        'If CheckExistForm(FormActive) Then Return Nothing
        SetVariable()
        Try
            Dim file As String = (New System.Uri(Assembly.GetExecutingAssembly().CodeBase)).LocalPath
            Dim dirName As String = New System.IO.DirectoryInfo(file).Parent.FullName
            Dim frmAss As System.Reflection.Assembly

            frmAss = System.Reflection.Assembly.LoadFrom(dirName & "\" & sDll & ".dll")
            Dim FormInstanceType As Object = frmAss.GetType(sDll & "." & FormActive)
            '  Dim frm As PageFrame = New PageFrame()
            Dim frm As Object 'System.Windows.Window = New System.Windows.Window
            Dim page As Boolean = False
            Try
                If frmAss.GetType(sDll & "." & FormActive).BaseType.Name.Contains("Page") Then
                    frm = New System.Windows.Controls.Page 'PageFrame()
                    frm.FPage = CType(Activator.CreateInstance(FormInstanceType), System.Windows.Controls.Page)
                    page = True
                Else
                    frm = New System.Windows.Window
                    frm = CType(Activator.CreateInstance(FormInstanceType), System.Windows.Window)
                    Integration.ElementHost.EnableModelessKeyboardInterop(frm)
                    '  frm.SizeToContent = Windows.SizeToContent.WidthAndHeight
                    frm.WindowStartupLocation = Windows.WindowStartupLocation.CenterScreen
                End If
            Catch ex As Exception
                If ex.Message.Contains("CreateInstance") Then
                    D99C0008.MsgL3("Not exist form " & FormActive & " in " & sDll & " dll", L3MessageBoxIcon.Exclamation)
                Else
                    D99C0008.MsgL3("L3: " & ex.Message & " - " & ex.Source, L3MessageBoxIcon.Exclamation)
                End If

                Return Nothing
            End Try

            If arrPro IsNot Nothing Then 'Gắn giá trị tham số WriteOnly
                For i As Integer = 0 To arrPro.Length - 1
                    SetProperties(frm, arrPro(i).properties, arrPro(i).value)
                Next
            Else
                WriteLogFile(frm.Name & " (" & frm.GetType.Namespace & ") : No parameter", "LogParaDLL.log", False)
            End If
            If showDialog Then
                frm.ShowDialog()
                ' bSaved = L3Bool(GetProperties(frm, "bSaved"))
                frm.Close()
            Else
                If page Then
                    frm.MdiParent = frmCall
                    frm.AutoScroll = True
                    'Update 10/06/2015: gắn lại tooltip sau khi load form
                End If
                frm.Show()

            End If

            Return frm

        Catch ex As Exception
            If System.IO.File.Exists(gsApplicationSetup & "\" & sDll & ".dll") = False Then
                D99C0008.MsgL3("Not exist file " & gsApplicationSetup & "\" & sDll & ".dll", L3MessageBoxIcon.Exclamation)
            Else
                D99C0008.MsgL3(ex.Message, L3MessageBoxIcon.Exclamation)
            End If

        Finally

        End Try

        Return Nothing
    End Function

    Public Sub SetProperties(ByVal frm As Object, ByVal properties As String, ByVal value As Object)
        Try
            Dim info As System.Reflection.PropertyInfo = frm.GetType.GetProperty(properties)
            If info IsNot Nothing Then
                info.SetValue(frm, value, Nothing)
                If FormID <> frm.Name Then
                    WriteLogFile("The parameters of " & frm.Name & " (" & frm.GetType.Namespace & ") :", "LogParaDLL.log", False)
                    '  SetVariable()
                End If
                WriteLogFile("'" & properties & "' = '" & L3String(value) & "'", "LogParaDLL.log", True)
                FormID = frm.Name
            Else 'Vi phạm tham số
                WriteLogFile("'" & properties & "' of " & frm.Name & " (" & frm.GetType.Namespace & "): doesn't exist.", "LogParaDLL.log", True)
            End If

        Catch ex As Exception
            D99C0008.Msg("DLL: " & ex.Message)
        End Try
    End Sub

    Public Function CallWPF(ByVal sDLLName As String, ByVal sFormActive As String, ByVal arrPro() As StructureProperties) As System.Windows.Window
        Return CallWPF(Nothing, sDLLName, sFormActive, True, arrPro)
    End Function
#End Region
    Private Class DxxExx40
        Dim sModuleID As String = ""

        Private _exeName As String = ""
        Public WriteOnly Property ExeName() As String
            Set(ByVal Value As String)
                _exeName = Value
            End Set
        End Property

        Public Sub New()

        End Sub
        ''' <summary>
        ''' Khởi tạo exe con DxxExx40
        ''' </summary>
        ''' <param name="Server">Server kết nối đến hệ thống</param>
        ''' <param name="Database">Database kết nối đến hệ thống</param>
        ''' <param name="UserDatabaseID">User Database kết nối đến hệ thống</param>
        ''' <param name="Password">Password kết nối đến hệ thống</param>
        ''' <param name="UserID">User Lemon3 kết nối đến hệ thống</param>
        ''' <param name="Language">Ngôn ngữ sử dụng</param>
        Public Sub New(ByVal ExeName As String, ByVal Server As String, ByVal Database As String, ByVal UserDatabaseID As String, ByVal Password As String, ByVal UserID As String, ByVal Language As String)
            _exeName = ExeName
            sModuleID = L3Left(ExeName, 3)

            D99C0007.SaveOthersSetting(sModuleID, _exeName, "ServerName", Server, CodeOption.lmCode)
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "DBName", Database, CodeOption.lmCode)
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "ConnectionUserID", UserDatabaseID, CodeOption.lmCode)
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "Password", Password, CodeOption.lmCode)
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "UserID", UserID, CodeOption.lmCode)
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "Language", Language)

            'gbUnicode: True(Nhập liệu Unicode)/False(Nhập liệu VNI)
            'gbUseD54: True(Sử dụng Dự án, Hạng mục)/False (Không sử dụng Dự án, Hạng mục)
            'gbPrintVNI: Có sử dụng In báo cáo VNI hay không khi nhập liệu Unicode
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "CodeTable", gbUnicode.ToString)
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "UseD54", gbUseD54.ToString)
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "IsPrintVNI", gbPrintVNI.ToString)
        End Sub

        ''' <summary>
        ''' Khởi tạo exe con DxxExx40
        ''' </summary>
        ''' <param name="Server">Server kết nối đến hệ thống</param>
        ''' <param name="Database">Database kết nối đến hệ thống</param>
        ''' <param name="UserDatabaseID">User Database kết nối đến hệ thống</param>
        ''' <param name="Password">Password kết nối đến hệ thống</param>
        ''' <param name="UserID">User Lemon3 kết nối đến hệ thống</param>
        ''' <param name="Language">Ngôn ngữ sử dụng</param>
        ''' <param name="DivisionID">Đơn vị hiện tại</param>
        ''' <param name="TranMonth">Tháng kế toán hiện tại</param>
        ''' <param name="TranYear">Năm kế toán hiện tại</param>
        Public Sub New(ByVal ExeName As String, ByVal Server As String, ByVal Database As String, ByVal UserDatabaseID As String, ByVal Password As String, ByVal UserID As String, ByVal Language As String, ByVal DivisionID As String, ByVal TranMonth As Integer, ByVal TranYear As Integer)
            Me.New(ExeName, Server, Database, UserDatabaseID, Password, UserID, Language)

            D99C0007.SaveOthersSetting(sModuleID, _exeName, "DivisionID", DivisionID)
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "TranMonth", TranMonth.ToString)
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "TranYear", TranYear.ToString)

            D99C0007.SaveOthersSetting(sModuleID, _exeName, "Closed", gbClosed.ToString)
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "CreateBy", gsCreateBy) 'gsCreateBy: Lấy giá Người tạo
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "LockL3UserID", gbLockL3UserID.ToString) 'gbLockL3UserID: Lấy giá trị Khóa người dùng Lemon3
            '*********************************
            Dim sExe As String = System.Reflection.Assembly.GetExecutingAssembly.GetName.Name
            D99C0007.SaveOthersSetting(sModuleID, _exeName, "ModuleID", sExe.Substring(1, 2))
        End Sub

        ''' <summary>
        ''' Màn hình cần hiển thị cho exe con
        ''' </summary>
        Public WriteOnly Property FormActive() As String
            Set(ByVal Value As String)
                D99C0007.SaveOthersSetting(sModuleID, _exeName, "Ctrl01", Value)
            End Set
        End Property

        ''' <summary>
        ''' Form Phân quyền cho màn hình được gọi 
        ''' </summary>
        Public WriteOnly Property FormPermission() As String
            Set(ByVal Value As String)
                D99C0007.SaveOthersSetting(sModuleID, _exeName, "Ctrl03", Value)
            End Set
        End Property


        Public WriteOnly Property FormState() As EnumFormState
            Set(ByVal Value As EnumFormState)
                D99C0007.SaveOthersSetting(sModuleID, _exeName, "FormState", CType(Value, Integer).ToString)
                D99C0007.SaveOthersSetting(sModuleID, _exeName, "Ctrl02", CType(Value, Integer).ToString)
            End Set
        End Property

        Public WriteOnly Property IDxx(ByVal sKeyID As String) As String
            Set(ByVal Value As String)
                D99C0007.SaveOthersSetting(sModuleID, _exeName, sKeyID, Value)
            End Set
        End Property

        Public ReadOnly Property OutputXX(ByVal sModuleID As String, ByVal _exeName As String, ByVal sKeyID As String) As String
            Get
                Return D99C0007.GetOthersSetting(sModuleID, _exeName, sKeyID, "")
            End Get
        End Property

        ''' <summary>
        ''' Thực thi exe con
        ''' </summary>
        Public Sub Run()
            If Not ExistFile(gsApplicationSetup & "\" & _exeName & ".exe") Then Exit Sub
            Dim pInfo As New System.Diagnostics.ProcessStartInfo(gsApplicationSetup & "\" & _exeName & ".exe")
            pInfo.Arguments = "/DigiNet Corporation"
            pInfo.WindowStyle = ProcessWindowStyle.Normal
            '*************************
            Dim sExe As String = System.Reflection.Assembly.GetExecutingAssembly.GetName.Name
            SaveRunningExeSettings(sExe, _exeName)
            '*************************
            Process.Start(pInfo)
        End Sub

        ''' <summary>
        ''' Kiểm tra tồn tại exe con không ?
        ''' </summary>
        Private Function ExistFile(ByVal Path As String) As Boolean
            If System.IO.File.Exists(Path) Then Return True
            If geLanguage = EnumLanguage.Vietnamese Then
                D99C0008.MsgL3("Không tồn tại file " & _exeName & ".exe")
            Else
                D99C0008.MsgL3("Not exist file " & _exeName & ".exe")
            End If
            Return False
        End Function

    End Class
End Module

