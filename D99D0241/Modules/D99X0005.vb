Imports System.Text
'#######################################################################################
'#                                     CHÚ Ý
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày cập nhật cuối cùng: 25/12/2013
'# Người cập nhật cuối cùng: Minh Hòa
'# Diễn giải: Thêm hàm SetShortcutPopupMenu chỉ cho ContextMenuStrip
'# Bổ sung icon và phím nóng cho các menu trên ToolStrip
'# Bổ sung menu Xuất Excel trực tiếp. mnsExportOut
'# Bổ sung menu Xuất dữ liệu, tsmExportDataScript
'# Bổ sung tham số oDisplayStyle vào hàm SetShortcutPopupMenu cho dạng màn hình mới (hiển thị ImageAndText)
'#######################################################################################

''' <summary>
''' Module quản lý các vấn đề về Shortcut và Image của Popupmenu
''' </summary>
''' <remarks>Chỉ gắn vào exe nào có sử dụng popupmenu</remarks>
Public Module D99X0005

#Region "Set Image và Shortcut của Popupmenu"
    ''' <summary>
    ''' Set Image và Shortcut của Popupmenu
    ''' </summary>
    ''' <param name="C1CmdHolder">C1CommandHolder của popupmenu</param>
    ''' <remarks>Gọi hàm SetShortcutPopupMenu(C1CommandHolder)tại form_load</remarks>
    <DebuggerStepThrough()> _
    Public Sub SetShortcutPopupMenu(ByVal C1CmdHolder As C1.Win.C1Command.C1CommandHolder)
        'On Error Resume Next

        With C1CmdHolder
            .SmoothImages = False
            '************************************************
            'Image và Shortcut chuẩn
            'Thêm
            For i As Integer = 0 To .Commands.Count - 1
                Select Case .Commands(i).Name
                    Case "mnuAdd"
                        .Commands("mnuAdd").Image = GetImageMenu(ImageMenuType.Add)
                        .Commands("mnuAdd").Shortcut = Shortcut.CtrlN
                    Case "mnuInherit"
                        .Commands("mnuInherit").Image = GetImageMenu(ImageMenuType.Copy)
                        .Commands("mnuInherit").Shortcut = Shortcut.CtrlY
                    Case "mnuView"
                        .Commands("mnuView").Image = GetImageMenu(ImageMenuType.View)
                        .Commands("mnuView").Shortcut = Shortcut.CtrlW
                    Case "mnuEdit"
                        .Commands("mnuEdit").Image = GetImageMenu(ImageMenuType.Edit)
                        .Commands("mnuEdit").Shortcut = Shortcut.CtrlE
                    Case "mnuDelete"
                        .Commands("mnuDelete").Image = GetImageMenu(ImageMenuType.Delete)
                        .Commands("mnuDelete").Shortcut = Shortcut.CtrlD
                    Case "mnuFind"
                        .Commands("mnuFind").Image = GetImageMenu(ImageMenuType.Find)
                        .Commands("mnuFind").Shortcut = Shortcut.CtrlF
                    Case "mnuListAll"
                        .Commands("mnuListAll").Image = GetImageMenu(ImageMenuType.ListAlL)
                        .Commands("mnuListAll").Shortcut = Shortcut.CtrlA
                    Case "mnuSysInfo"
                        .Commands("mnuSysInfo").Image = GetImageMenu(ImageMenuType.SysInfo)
                        .Commands("mnuSysInfo").Shortcut = Shortcut.CtrlI
                    Case "mnuPrint"
                        .Commands("mnuPrint").Image = GetImageMenu(ImageMenuType.PrintReport)
                        .Commands("mnuPrint").Shortcut = Shortcut.CtrlP
                    Case "mnuExportToExcel"
                        .Commands("mnuExportToExcel").Image = GetImageMenu(ImageMenuType.ExportToExcel)
                        .Commands("mnuExportToExcel").Shortcut = Shortcut.CtrlX
                    Case "mnuEditOther"
                        .Commands("mnuEditOther").Image = GetImageMenu(ImageMenuType.EditOther)
                        .Commands("mnuEditOther").Shortcut = Shortcut.CtrlO
                    Case "mnuEditVoucher"
                        .Commands("mnuEditVoucher").Image = GetImageMenu(ImageMenuType.EditVoucher)
                        .Commands("mnuEditVoucher").Shortcut = Shortcut.CtrlU
                    Case "mnuIssueStock"
                        .Commands("mnuIssueStock").Image = My.Resources.IssueStock
                        .Commands("mnuIssueStock").Shortcut = Shortcut.CtrlK
                    Case "mnuLockVoucher"
                        .Commands("mnuLockVoucher").Image = GetImageMenu(ImageMenuType.LockVoucher)
                        .Commands("mnuLockVoucher").Shortcut = Shortcut.CtrlL
                    Case "mnuShowDetail"
                        .Commands("mnuShowDetail").Image = GetImageMenu(ImageMenuType.ShowDetail)
                        .Commands("mnuShowDetail").Shortcut = Shortcut.CtrlS
                    Case "mnuApprove"
                        .Commands("mnuApprove").Image = GetImageMenu(ImageMenuType.Approve)
                        .Commands("mnuApprove").Shortcut = Shortcut.CtrlR
                    Case "mnuCancelVoucher"
                        .Commands("mnuCancelVoucher").Image = GetImageMenu(ImageMenuType.CancelVoucher)
                        .Commands("mnuCancelVoucher").Shortcut = Shortcut.CtrlH
                    Case "mnuImportData"
                        .Commands("mnuImportData").Image = GetImageMenu(ImageMenuType.ImportData)
                        .Commands("mnuImportData").Shortcut = Shortcut.CtrlM
                    Case "mnuExportOut"
                        .Commands("mnuExportOut").Image = GetImageMenu(ImageMenuType.ExportToExcelOut)
                    Case "mnuExportDataScript"
                        .Commands("mnuExportDataScript").Image = GetImageMenu(ImageMenuType.ExportDataScript)
                    Case "C1CommandMenu1"
                        .Commands("C1CommandMenu1").Image = GetImageMenu(ImageMenuType.PrintReport)
                End Select

            Next
        End With
    End Sub
#End Region

#Region "Set Image và Shortcut của ToolStrip"

    Public Function GetImageMenu(typeMenu As ImageMenuType) As Image
        If SCT_BackColor = "" Then
            Return GetImageMenuClassis(typeMenu)
        Else
            Return GetImageMenuBlue(typeMenu)
        End If
    End Function
    Private Function GetImageMenuClassis(typeMenu As ImageMenuType) As Image
        Select Case typeMenu
            Case ImageMenuType.Add
                Return My.Resources.add
            Case ImageMenuType.CloseForm
                Return My.Resources.CloseForm
            Case ImageMenuType.Delete
                Return My.Resources.delete
            Case ImageMenuType.Edit
                Return My.Resources.edit
            Case ImageMenuType.ExportDataScript
                Return My.Resources.ExportDataScript
            Case ImageMenuType.ExportToExcel
                Return My.Resources.ExportToExcel
            Case ImageMenuType.Find
                Return My.Resources.find
            Case ImageMenuType.ImportData
                Return My.Resources.ImportData
            Case ImageMenuType.ListAlL
                Return My.Resources.ListAll
            Case ImageMenuType.PrintReport
                Return My.Resources.PrintReport
            Case ImageMenuType.SysInfo
                Return My.Resources.SysInfo
            Case ImageMenuType.View
                Return My.Resources.view
            Case ImageMenuType.Approve
                Return My.Resources.approve
            Case ImageMenuType.CancelVoucher
                Return My.Resources.CancelVoucher
            Case ImageMenuType.Copy
                Return My.Resources.copy
            Case ImageMenuType.EditOther
                Return My.Resources.EditOther
            Case ImageMenuType.EditVoucher
                Return My.Resources.EditVoucher
            Case ImageMenuType.ExportToExcelOut
                Return My.Resources.ExportToExcelOut
            Case ImageMenuType.LockVoucher
                Return My.Resources.LockVoucher
            Case ImageMenuType.ShowDetail
                Return My.Resources.ShowDetail
            Case Else
                Return Nothing
        End Select
    End Function

    Private Function GetImageMenuBlue(typeMenu As ImageMenuType) As Image
        Select Case typeMenu
            Case ImageMenuType.Add
                Return My.Resources.add_blue
            Case ImageMenuType.CloseForm
                Return My.Resources.CloseForm_blue
            Case ImageMenuType.Delete
                Return My.Resources.delete_blue
            Case ImageMenuType.Edit
                Return My.Resources.edit_blue
            Case ImageMenuType.ExportDataScript
                Return My.Resources.ExportDataScript_blue
            Case ImageMenuType.ExportToExcel
                Return My.Resources.ExportToExcel_blue
            Case ImageMenuType.Find
                Return My.Resources.find_blue
            Case ImageMenuType.ImportData
                Return My.Resources.ImportData_blue
            Case ImageMenuType.ListAlL
                Return My.Resources.ListAll_blue
            Case ImageMenuType.PrintReport
                Return My.Resources.PrintReport_blue
            Case ImageMenuType.SysInfo
                Return My.Resources.SysInfo_blue
            Case ImageMenuType.View
                Return My.Resources.view_blue

            Case ImageMenuType.Approve
                Return My.Resources.approve_blue
            Case ImageMenuType.CancelVoucher
                Return My.Resources.CancelVoucher_blue
            Case ImageMenuType.Copy
                Return My.Resources.copy_blue
            Case ImageMenuType.EditOther
                Return My.Resources.EditOther_blue
            Case ImageMenuType.EditVoucher
                Return My.Resources.EditVoucher_blue
            Case ImageMenuType.ExportToExcelOut
                Return My.Resources.ExportToExcelOut_blue
            Case ImageMenuType.LockVoucher
                Return My.Resources.LockVoucher_blue
            Case ImageMenuType.ShowDetail
                Return My.Resources.ShowDetail_blue
            Case Else
                Return Nothing
        End Select
    End Function

    ''' <summary>
    ''' Set Image và Shortcut của ContextMenuStrip
    ''' </summary>
    ''' <param name="bTransactionHistory">Nếu = True thì menu tsbSysInfo có tên là Lịch sử tác động</param>
    ''' <param name="ContextMenuStrip"></param>
    ''' <remarks>Gọi hàm  SetShortcutPopupMenu(ContextMenuStrip1) tại form_load</remarks>
    Public Sub SetShortcutPopupMenu(ByVal ContextMenuStrip As System.Windows.Forms.ContextMenuStrip, Optional ByVal bTransactionHistory As Boolean = False)
        'On Error Resume Next
        'Update 19/07/2011: Set ShortCut cho ContextMenuStrip (Áp dụng cho một số đối tượng có sử dụng đến Menu ngắn: Lưới, treeview...)
        If ContextMenuStrip IsNot Nothing Then
            With ContextMenuStrip
                For i As Integer = 0 To .Items.Count - 1
                    Select Case .Items(i).Name
                        Case "mnsAdd"
                            .Items("mnsAdd").Image = GetImageMenu(ImageMenuType.Add)
                            .Items("mnsAdd").Text = rL3("_Them") '&Thêm
                            CType(.Items("mnsAdd"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
                        Case "mnsInherit"
                            .Items("mnsInherit").Image = GetImageMenu(ImageMenuType.Copy)
                            .Items("mnsInherit").Text = rL3("Ke_thu_a") 'Kế thừ&a
                            CType(.Items("mnsInherit"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Y), System.Windows.Forms.Keys)
                        Case "mnsView"
                            .Items("mnsView").Image = GetImageMenu(ImageMenuType.View)
                            .Items("mnsView").Text = rL3("Xe_m") 'Xe&m
                            CType(.Items("mnsView"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.W), System.Windows.Forms.Keys)
                        Case "mnsEdit"
                            .Items("mnsEdit").Image = GetImageMenu(ImageMenuType.Edit)
                            .Items("mnsEdit").Text = rL3("_Sua") '&Sửa
                            CType(.Items("mnsEdit"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
                        Case "mnsDelete"
                            .Items("mnsDelete").Image = GetImageMenu(ImageMenuType.Delete)
                            .Items("mnsDelete").Text = rL3("_Xoa") '&Xóa
                            CType(.Items("mnsDelete"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D), System.Windows.Forms.Keys)
                        Case "mnsFind"
                            .Items("mnsFind").Image = GetImageMenu(ImageMenuType.Find)
                            .Items("mnsFind").Text = rL3("Tim__kiem") 'Tìm &kiếm
                            CType(.Items("mnsFind"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
                        Case "mnsListAll"
                            .Items("mnsListAll").Image = GetImageMenu(ImageMenuType.ListAlL)
                            .Items("mnsListAll").Text = rL3("_Liet_ke_tat_ca") '&Liệt kế tất cả
                            CType(.Items("mnsListAll"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
                        Case "mnsSysInfo"
                            .Items("mnsSysInfo").Image = GetImageMenu(ImageMenuType.SysInfo)
                            If bTransactionHistory Then
                                .Items("mnsSysInfo").Text = rL3("Lich_su_tac_dong") 'Lịch sử tác động
                            Else
                                .Items("mnsSysInfo").Text = rL3("Thong_tin__he_thong") 'Thông tin &hệ thống
                            End If
                            CType(.Items("mnsSysInfo"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.I), System.Windows.Forms.Keys)
                        Case "mnsPrint"
                            .Items("mnsPrint").Image = GetImageMenu(ImageMenuType.PrintReport)
                            .Items("mnsPrint").Text = rL3("_In") '&In
                            CType(.Items("mnsPrint"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
                        Case "mnsExportToExcel"
                            .Items("mnsExportToExcel").Image = GetImageMenu(ImageMenuType.ExportToExcel)
                            .Items("mnsExportToExcel").Text = rL3("Xuat__Excel") 'Xuất &Excel
                            CType(.Items("mnsExportToExcel"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
                        Case "mnsEditOther"
                            .Items("mnsEditOther").Image = GetImageMenu(ImageMenuType.EditOther)
                            .Items("mnsEditOther").Text = rL3("Sua_khac_(_O)") 'Sửa khác (&O)
                            CType(.Items("mnsEditOther"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
                        Case "mnsEditVoucher"
                            .Items("mnsEditVoucher").Image = GetImageMenu(ImageMenuType.EditVoucher)
                            .Items("mnsEditVoucher").Text = rL3("Sua_so_phie_u") 'Sửa số phiế&u
                            CType(.Items("mnsEditVoucher"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.U), System.Windows.Forms.Keys)
                        Case "mnsIssueStock"
                            .Items("mnsIssueStock").Image = My.Resources.IssueStock
                            .Items("mnsIssueStock").Text = rL3("Xuat_kho") 'Xuất kho
                            CType(.Items("mnsIssueStock"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.K), System.Windows.Forms.Keys)
                        Case "mnsLocked"
                            .Items("mnsLocked").Image = GetImageMenu(ImageMenuType.LockVoucher)
                            .Items("mnsLocked").Text = rL3("Khoa__phieu") 'Khóa &phiếu
                            CType(.Items("mnsLocked"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.L), System.Windows.Forms.Keys)
                        Case "mnsLockVoucher"
                            .Items("mnsLockVoucher").Image = GetImageMenu(ImageMenuType.LockVoucher)
                            .Items("mnsLockVoucher").Text = rL3("Khoa__phieu") 'Khóa &phiếu
                            CType(.Items("mnsLockVoucher"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.L), System.Windows.Forms.Keys)
                        Case "mnsShowDetail"
                            .Items("mnsShowDetail").Image = GetImageMenu(ImageMenuType.ShowDetail)
                            .Items("mnsShowDetail").Text = rL3("Hien_thi__chi_tiet") 'Hiển thị &chi tiết
                            CType(.Items("mnsShowDetail"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
                        Case "mnsApprove"
                            .Items("mnsApprove").Image = GetImageMenu(ImageMenuType.Approve)
                            .Items("mnsApprove").Text = rL3("_Duyet") '&Duyệt
                            CType(.Items("mnsApprove"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
                        Case "mnsCancelVoucher"
                            .Items("mnsCancelVoucher").Image = GetImageMenu(ImageMenuType.CancelVoucher)
                            .Items("mnsCancelVoucher").Text = rL3("Huy_phieu") '&Hủy phiếu
                            CType(.Items("mnsCancelVoucher"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.H), System.Windows.Forms.Keys)
                        Case "mnsImportData"
                            .Items("mnsImportData").Image = GetImageMenu(ImageMenuType.ImportData)
                            .Items("mnsImportData").Text = rL3("Import__du_lieu") 'Import &dữ liệu 
                            CType(.Items("mnsImportData"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.M), System.Windows.Forms.Keys)
                        Case "mnsExportOut"
                            .Items("mnsExportOut").Image = GetImageMenu(ImageMenuType.ExportToExcelOut)
                            .Items("mnsExportOut").Text = rL3("MSG000052") 'Xuất Excel trực tiếp
                        Case "mnuExportDataScript"
                            .Items("mnsExportDataScript").Image = GetImageMenu(ImageMenuType.ExportDataScript)
                            .Items("mnsExportDataScript").Text = rL3("MSG000051") 'X&uất dữ liệu 
                    End Select
                Next
            End With
        End If

        'Các phím chung của sự kiện Form_Keydown
        'AddHandler Form.KeyDown, AddressOf Form_KeyDown
    End Sub


    ''' <summary>
    ''' Set Image và Shortcut của ToolStripButton
    ''' </summary>
    ''' <param name="Form">Me</param>
    ''' <param name="TableToolStrip">System.Windows.Forms.ToolStri</param>
    ''' <param name="bTransactionHistory">Nếu = True thì menu tsbSysInfo có tên là Lịch sử tác động</param>
    ''' <param name="ContextMenuStrip"></param>
    ''' <remarks>Gọi hàm  SetShortcutPopupMenu(Me, TableToolStrip, ContextMenuStrip1) tại form_load</remarks>
    Public Sub SetShortcutPopupMenu(ByVal Form As Form, ByVal TableToolStrip As System.Windows.Forms.ToolStrip, ByVal ContextMenuStrip As System.Windows.Forms.ContextMenuStrip, ByVal bTransactionHistory As Boolean, ByVal oDisplayStyle As ToolStripItemDisplayStyle)
        'On Error Resume Next

        If TableToolStrip IsNot Nothing Then  ' Set ShortCut cho TableToolStrip



            With TableToolStrip 'Khi bổ sung menu mới thì bổ sung phím nóng vào hàm Form_Keydown
                If SCT_BackColor <> "" Then
                    .BackColor = ColorTranslator.FromHtml(SCT_BackColorMenu)
                Else
                    .BackColor = System.Drawing.SystemColors.ActiveBorder
                End If

                Dim bStandard As Boolean = True  'Nếu menu chuẩn thì có một số thuộc tính chung
                For i As Integer = 0 To .Items.Count - 1
                    bStandard = True
                    Select Case .Items(i).Name
                        Case "tsbAdd"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.Add)
                            .Items(.Items(i).Name).Text = rL3("_Them")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbAdd").Text.Replace("&", "") & " (Ctrl+N)"
                        Case "tsbInherit"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.Copy)
                            .Items(.Items(i).Name).Text = rL3("Ke_thu_a")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbInherit").Text.Replace("&", "") & " (Ctrl+Y)"
                        Case "tsbView"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.View)
                            .Items(.Items(i).Name).Text = rL3("Xe_m")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbView").Text.Replace("&", "") & " (Ctrl+W)"
                        Case "tsbEdit"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.Edit)
                            .Items(.Items(i).Name).Text = rL3("_Sua")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbEdit").Text.Replace("&", "") & " (Ctrl+E)"
                        Case "tsbDelete"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.Delete)
                            .Items(.Items(i).Name).Text = rL3("_Xoa")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbDelete").Text.Replace("&", "") & " (Ctrl+D)"
                        Case "tsbFind"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.Find)
                            .Items(.Items(i).Name).Text = rL3("Tim__kiem")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbFind").Text.Replace("&", "") & " (Ctrl+F)"
                        Case "tsbListAll"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.ListAlL)
                            .Items(.Items(i).Name).Text = rL3("_Liet_ke_tat_ca")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbListAll").Text.Replace("&", "") & " (Ctrl+A)"
                        Case "tsbSysInfo"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.SysInfo)
                            If bTransactionHistory Then
                                .Items(.Items(i).Name).Text = rL3("Lich_su_tac_dong") 'Lịch sử tác động
                            Else
                                .Items(.Items(i).Name).Text = rL3("Thong_tin__he_thong") 'Thông tin &hệ thống
                            End If
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbSysInfo").Text.Replace("&", "") & " (Ctrl+I)"
                        Case "tsbPrint"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.PrintReport)
                            .Items(.Items(i).Name).Text = rL3("_In")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbPrint").Text.Replace("&", "") & " (Ctrl+P)"
                        Case "tsbExportToExcel"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.ExportToExcel)
                            .Items(.Items(i).Name).Text = rL3("Xuat__Excel")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbExportToExcel").Text.Replace("&", "") & " (Ctrl+X)"
                        Case "tsbEditOther"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.EditOther)
                            .Items(.Items(i).Name).Text = rL3("Sua_khac_(_O)")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbEditOther").Text.Replace("&", "") & " (Ctrl+O)"
                        Case "tsbEditVoucher"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.EditVoucher)
                            .Items(.Items(i).Name).Text = rL3("Sua_so_phie_u")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbEditVoucher").Text.Replace("&", "") & " (Ctrl+U)"
                        Case "tsbIssueStock"
                            .Items(.Items(i).Name).Image = My.Resources.IssueStock
                            .Items(.Items(i).Name).Text = rL3("Xuat_kho")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbIssueStock").Text.Replace("&", "") & " (Ctrl+K)"
                        Case "tsbLocked"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.LockVoucher)
                            .Items(.Items(i).Name).Text = rL3("Khoa__phieu")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbLocked").Text.Replace("&", "") & " (Ctrl+L)"
                        Case "tsbLockVoucher"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.LockVoucher)
                            .Items(.Items(i).Name).Text = rL3("Khoa__phieu")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbLockVoucher").Text.Replace("&", "") & " (Ctrl+L)"
                        Case "tsbShowDetail"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.ShowDetail)
                            .Items(.Items(i).Name).Text = rL3("Hien_thi__chi_tiet")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbShowDetail").Text.Replace("&", "") & " (Ctrl+S)"
                        Case "tsbApprove"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.Approve)
                            .Items(.Items(i).Name).Text = rL3("_Duyet")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbApprove").Text.Replace("&", "") & " (Ctrl+R)"
                        Case "tsbCancelVoucher"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.CancelVoucher)
                            .Items(.Items(i).Name).Text = rL3("Huy_phieu")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbCancelVoucher").Text.Replace("&", "") & " (Ctrl+H)"
                        Case "tsbImportData"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.ImportData)
                            .Items(.Items(i).Name).Text = rL3("Import__du_lieu")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbImportData").Text.Replace("&", "") & " (Ctrl+M)"
                        Case "tsbClose"
                            .Items(.Items(i).Name).Image = GetImageMenu(ImageMenuType.CloseForm)
                            .Items(.Items(i).Name).Text = rL3("Do_ng")
                            .Items(.Items(i).Name).ToolTipText = .Items("tsbClose").Text.Replace("&", "") & " (Ctrl+Q)"
                        Case "tsdActive"
                            bStandard = False
                            .Items("tsdActive").DisplayStyle = ToolStripItemDisplayStyle.Text
                            .Items("tsdActive").Text = rL3("_Thuc_hien_") '&Thực hiện...
                        Case "tsdTranstion"
                            bStandard = False
                            .Items("tsdTranstion").DisplayStyle = ToolStripItemDisplayStyle.Text
                            .Items("tsdTranstion").Text = rL3("_Nghiep_vu") '&Nghiệp vụ
                        Case Else
                            bStandard = False
                    End Select
                    If bStandard Then .Items(.Items(i).Name).DisplayStyle = oDisplayStyle
                Next

            End With

            Dim tsdActive As ToolStripDropDownButton = CType(TableToolStrip.Items("tsdActive"), ToolStripDropDownButton)
            '        tsdActive.Alignment = ToolStripItemAlignment.Right
            '        tsdActive.DropDownDirection = ToolStripDropDownDirection.BelowLeft

            '  Dim tsdTranstion As ToolStripDropDownButton = CType(TableToolStrip.Items("tsdTranstion"), ToolStripDropDownButton)
            If tsdActive IsNot Nothing Then 'Tồn tại
                Dim mnu As ToolStripMenuItem

                With tsdActive.DropDownItems

                    For i As Integer = 0 To .Count - 1
                        If Not TypeOf (.Item(tsdActive.DropDownItems(i).Name)) Is ToolStripMenuItem Then Continue For
                        mnu = CType(.Item(tsdActive.DropDownItems(i).Name), ToolStripMenuItem)
                        Select Case tsdActive.DropDownItems(i).Name
                            Case "tsmAdd"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.Add)
                                mnu.Text = rL3("_Them") '&Thêm
                            Case "tsmEdit"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.Edit)
                                mnu.Text = rL3("_Sua") '&Sửa
                            Case "tsmInherit"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Y), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.Copy)
                                mnu.Text = rL3("Ke_thu_a") 'Kế thừ&a
                            Case "tsmView"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.W), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.View)
                                mnu.Text = rL3("Xe_m") 'Xe&m
                            Case "tsmDelete"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.Delete)
                                mnu.Text = rL3("_Xoa") '&Xóa
                            Case "tsmFind"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.Find)
                                mnu.Text = rL3("Tim__kiem") 'Tìm &kiếm
                            Case "tsmListAll"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.ListAlL)
                                mnu.Text = rL3("_Liet_ke_tat_ca") '&Liệt kế tất cả
                            Case "tsmSysInfo"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.I), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.SysInfo)
                                If bTransactionHistory Then
                                    mnu.Text = rL3("Lich_su_tac_dong") 'Lịch sử tác động
                                Else
                                    mnu.Text = rL3("Thong_tin__he_thong") 'Thông tin &hệ thống
                                End If
                            Case "tsmPrint"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.PrintReport)
                                mnu.Text = rL3("_In") '&In
                            Case "tsmExportToExcel"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.ExportToExcel)
                                mnu.Text = rL3("Xuat__Excel") 'Xuất &Excel
                            Case "tsmImportData"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.M), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.ImportData)
                                mnu.Text = rL3("Import__du_lieu") 'Import &dữ liệu 
                            Case "tsmEditOther"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.EditOther)
                                mnu.Text = rL3("Sua_khac_(_O)") 'Sửa khác (&O)
                            Case "tsmEditVoucher"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.U), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.EditVoucher)
                                mnu.Text = rL3("Sua_so_phie_u") 'Sửa số phiế&u
                            Case "tsmIssueStock"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.K), System.Windows.Forms.Keys)
                                mnu.Image = My.Resources.IssueStock
                                mnu.Text = rL3("Xuat_kho") 'Xuất kho
                            Case "tsmLocked"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.L), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.LockVoucher)
                                mnu.Text = rL3("Khoa__phieu") 'Khóa &phiếu
                            Case "tsmLockVoucher"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.L), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.LockVoucher)
                                mnu.Text = rL3("Khoa__phieu") 'Khóa &phiếu
                            Case "tsmShowDetail"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.ShowDetail)
                                mnu.Text = rL3("Hien_thi__chi_tiet") 'Hiển thị &chi tiết
                            Case "tsmApprove"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.Approve)
                                mnu.Text = rL3("_Duyet") '&Duyệt
                            Case "tsmCancelVoucher"
                                mnu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.H), System.Windows.Forms.Keys)
                                mnu.Image = GetImageMenu(ImageMenuType.CancelVoucher)
                                mnu.Text = rL3("Huy_phieu") '&Hủy phiếu
                            Case "tsmExportOut"
                                mnu.Image = GetImageMenu(ImageMenuType.ExportToExcelOut)
                                mnu.Text = rL3("MSG000052") 'Xuất Excel trực tiếp
                            Case "tsmExportDataScript"
                                mnu.Image = GetImageMenu(ImageMenuType.ExportDataScript)
                                mnu.Text = rL3("MSG000051") 'X&uất dữ liệu
                        End Select
                    Next
                End With
            End If
        End If
        'Update 19/07/2011: Set ShortCut cho ContextMenuStrip (Áp dụng cho một số đối tượng có sử dụng đến Menu nóng: Lưới, treeview...)
        If ContextMenuStrip IsNot Nothing Then
            With ContextMenuStrip
                For i As Integer = 0 To .Items.Count - 1
                    Select Case .Items(i).Name
                        Case "mnsAdd"
                            .Items("mnsAdd").Image = GetImageMenu(ImageMenuType.Add)
                            .Items("mnsAdd").Text = rL3("_Them") '&Thêm
                            CType(.Items("mnsAdd"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
                        Case "mnsInherit"
                            .Items("mnsInherit").Image = GetImageMenu(ImageMenuType.Copy)
                            .Items("mnsInherit").Text = rL3("Ke_thu_a") 'Kế thừ&a
                            CType(.Items("mnsInherit"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Y), System.Windows.Forms.Keys)
                        Case "mnsView"
                            .Items("mnsView").Image = GetImageMenu(ImageMenuType.View)
                            .Items("mnsView").Text = rL3("Xe_m") 'Xe&m
                            CType(.Items("mnsView"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.W), System.Windows.Forms.Keys)
                        Case "mnsEdit"
                            .Items("mnsEdit").Image = GetImageMenu(ImageMenuType.Edit)
                            .Items("mnsEdit").Text = rL3("_Sua") '&Sửa
                            CType(.Items("mnsEdit"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
                        Case "mnsDelete"
                            .Items("mnsDelete").Image = GetImageMenu(ImageMenuType.Delete)
                            .Items("mnsDelete").Text = rL3("_Xoa") '&Xóa
                            CType(.Items("mnsDelete"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D), System.Windows.Forms.Keys)
                        Case "mnsFind"
                            .Items("mnsFind").Image = GetImageMenu(ImageMenuType.Find)
                            .Items("mnsFind").Text = rL3("Tim__kiem") 'Tìm &kiếm
                            CType(.Items("mnsFind"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
                        Case "mnsListAll"
                            .Items("mnsListAll").Image = GetImageMenu(ImageMenuType.ListAlL)
                            .Items("mnsListAll").Text = rL3("_Liet_ke_tat_ca") '&Liệt kế tất cả
                            CType(.Items("mnsListAll"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
                        Case "mnsSysInfo"
                            .Items("mnsSysInfo").Image = GetImageMenu(ImageMenuType.SysInfo)
                            If bTransactionHistory Then
                                .Items("mnsSysInfo").Text = rL3("Lich_su_tac_dong") 'Lịch sử tác động
                            Else
                                .Items("mnsSysInfo").Text = rL3("Thong_tin__he_thong") 'Thông tin &hệ thống
                            End If
                            CType(.Items("mnsSysInfo"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.I), System.Windows.Forms.Keys)
                        Case "mnsPrint"
                            .Items("mnsPrint").Image = GetImageMenu(ImageMenuType.PrintReport)
                            .Items("mnsPrint").Text = rL3("_In") '&In
                            CType(.Items("mnsPrint"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
                        Case "mnsExportToExcel"
                            .Items("mnsExportToExcel").Image = GetImageMenu(ImageMenuType.ExportToExcel)
                            .Items("mnsExportToExcel").Text = rL3("Xuat__Excel") 'Xuất &Excel
                            CType(.Items("mnsExportToExcel"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
                        Case "mnsEditOther"
                            .Items("mnsEditOther").Image = GetImageMenu(ImageMenuType.EditOther)
                            .Items("mnsEditOther").Text = rL3("Sua_khac_(_O)") 'Sửa khác (&O)
                            CType(.Items("mnsEditOther"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
                        Case "mnsEditVoucher"
                            .Items("mnsEditVoucher").Image = GetImageMenu(ImageMenuType.EditVoucher)
                            .Items("mnsEditVoucher").Text = rL3("Sua_so_phie_u") 'Sửa số phiế&u
                            CType(.Items("mnsEditVoucher"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.U), System.Windows.Forms.Keys)
                        Case "mnsIssueStock"
                            .Items("mnsIssueStock").Image = My.Resources.IssueStock
                            .Items("mnsIssueStock").Text = rL3("Xuat_kho") 'Xuất kho
                            CType(.Items("mnsIssueStock"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.K), System.Windows.Forms.Keys)
                        Case "mnsLocked"
                            .Items("mnsLocked").Image = GetImageMenu(ImageMenuType.LockVoucher)
                            .Items("mnsLocked").Text = rL3("Khoa__phieu") 'Khóa &phiếu
                            CType(.Items("mnsLocked"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.L), System.Windows.Forms.Keys)
                        Case "mnsLockVoucher"
                            .Items("mnsLockVoucher").Image = GetImageMenu(ImageMenuType.LockVoucher)
                            .Items("mnsLockVoucher").Text = rL3("Khoa__phieu") 'Khóa &phiếu
                            CType(.Items("mnsLockVoucher"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.L), System.Windows.Forms.Keys)
                        Case "mnsShowDetail"
                            .Items("mnsShowDetail").Image = GetImageMenu(ImageMenuType.ShowDetail)
                            .Items("mnsShowDetail").Text = rL3("Hien_thi__chi_tiet") 'Hiển thị &chi tiết
                            CType(.Items("mnsShowDetail"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
                        Case "mnsApprove"
                            .Items("mnsApprove").Image = GetImageMenu(ImageMenuType.Approve)
                            .Items("mnsApprove").Text = rL3("_Duyet") '&Duyệt
                            CType(.Items("mnsApprove"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
                        Case "mnsCancelVoucher"
                            .Items("mnsCancelVoucher").Image = GetImageMenu(ImageMenuType.CancelVoucher)
                            .Items("mnsCancelVoucher").Text = rL3("Huy_phieu") '&Hủy phiếu
                            CType(.Items("mnsCancelVoucher"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.H), System.Windows.Forms.Keys)
                        Case "mnsImportData"
                            .Items("mnsImportData").Image = GetImageMenu(ImageMenuType.ImportData)
                            .Items("mnsImportData").Text = rL3("Import__du_lieu") 'Import &dữ liệu 
                            CType(.Items("mnsImportData"), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.M), System.Windows.Forms.Keys)
                        Case "mnsExportOut"
                            .Items("mnsExportOut").Image = GetImageMenu(ImageMenuType.ExportToExcelOut)
                            .Items("mnsExportOut").Text = rL3("MSG000052") 'Xuất Excel trực tiếp
                        Case "mnuExportDataScript"
                            .Items("mnsExportDataScript").Image = GetImageMenu(ImageMenuType.ExportDataScript)
                            .Items("mnsExportDataScript").Text = rL3("MSG000051") 'X&uất dữ liệu 
                    End Select
                Next
            End With
        End If

        'Các phím chung của sự kiện Form_Keydown
        If Form IsNot Nothing Then AddHandler Form.KeyDown, AddressOf Form_KeyDown
    End Sub


    Public Sub SetShortcutPopupMenu(ByVal Form As Form, ByVal TableToolStrip As System.Windows.Forms.ToolStrip, ByVal ContextMenuStrip As System.Windows.Forms.ContextMenuStrip)
        SetShortcutPopupMenu(Form, TableToolStrip, ContextMenuStrip, False, ToolStripItemDisplayStyle.Image)
    End Sub

    Public Sub SetShortcutPopupMenu(ByVal Form As Form, ByVal TableToolStrip As System.Windows.Forms.ToolStrip, ByVal ContextMenuStrip As System.Windows.Forms.ContextMenuStrip, ByVal bTransactionHistory As Boolean)
        SetShortcutPopupMenu(Form, TableToolStrip, ContextMenuStrip, bTransactionHistory, ToolStripItemDisplayStyle.Image)
    End Sub

    Private Sub HotKeyMenutsb(itemName As String, MainControl As Form)
        For Each ctrl As Control In MainControl.Controls
            If ctrl.GetType.Name = "ToolStrip" Then
                CallClick_ToolStripButton(CType(ctrl, ToolStrip), itemName)
                Exit Sub
            End If
        Next
    End Sub

    Private Sub Form_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.Control Then
            Dim MainControl As Form = CType(sender, Form)
            Select Case e.KeyCode
                '******************************************
                'Update 25/12/2013: Do làm màn hình mới không dùng ContextMenuStrip nên TableToolStrip không hiểu được phiếu nóng, nên bổ sung đoạn code này
                Case Keys.N 'Thêm mới
                    HotKeyMenutsb("tsbAdd", MainControl)
                Case Keys.Y 'Kế thừa
                    HotKeyMenutsb("tsbInherit", MainControl)
                Case Keys.W 'Xem
                    HotKeyMenutsb("tsbView", MainControl)
                Case Keys.E 'Sửa
                    HotKeyMenutsb("tsbEdit", MainControl)
                Case Keys.D 'Xóa
                    HotKeyMenutsb("tsbDelete", MainControl)
                Case Keys.F 'Tìm kiếm
                    HotKeyMenutsb("tsbFind", MainControl)
                Case Keys.A 'Liệt kê tất cả
                    HotKeyMenutsb("tsbListAll", MainControl)
                Case Keys.I 'Thông tin hệ thống
                    HotKeyMenutsb("tsbSysInfo", MainControl)
                Case Keys.P 'In
                    HotKeyMenutsb("tsbPrint", MainControl)
                Case Keys.O 'Sửa khác
                    HotKeyMenutsb("tsbEditOther", MainControl)
                Case Keys.U ''Sửa số phiếu
                    HotKeyMenutsb("tsbEditVoucher", MainControl)
                Case Keys.K ''Xuất kho
                    HotKeyMenutsb("tsbIssueStock", MainControl)
                Case Keys.L ''Khóa phiếu
                    HotKeyMenutsb("tsbLocked", MainControl)
                Case Keys.S 'Hiển thị chi tiết
                    HotKeyMenutsb("tsbShowDetail", MainControl)
                Case Keys.R 'Duyệt
                    HotKeyMenutsb("tsbApprove", MainControl)
                Case Keys.H 'Hủy phiếu
                    HotKeyMenutsb("tsbCancelVoucher", MainControl)
                Case Keys.X 'Xuất Excel
                    HotKeyMenutsb("tsbExportToExcel", MainControl)
                Case Keys.M 'Import Excel
                    HotKeyMenutsb("tsbImportData", MainControl)
                    '******************************************
                Case Keys.Q 'Đóng form
                    'Hình như không có tác dụng
                    ' Dim tsbcon As Control = MainControl.Controls.Find("tsbClose", True)(0)
                    HotKeyMenutsb("tsbClose", MainControl)
            End Select
        End If

    End Sub
    'Private Sub Form_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If e.Control Then
    '        Dim MainControl As Form
    '        MainControl = CType(sender, Form)
    '        Select Case e.KeyCode
    '            '******************************************
    '            'Update 25/12/2013: Do làm màn hình mới không dùng ContextMenuStrip nên TableToolStrip không hiểu được phiếu nóng, nên bổ sung đoạn code này
    '            Case Keys.N 'Thêm mới
    '                For Each ctrl As Control In MainControl.Controls
    '                    If ctrl.GetType.Name = "ToolStrip" Then
    '                        CallClick_ToolStripButton(CType(ctrl, ToolStrip), "tsbAdd")
    '                        Exit Sub
    '                    End If
    '                Next
    '            Case Keys.E 'Sửa
    '                For Each ctrl As Control In MainControl.Controls
    '                    If ctrl.GetType.Name = "ToolStrip" Then
    '                        CallClick_ToolStripButton(CType(ctrl, ToolStrip), "tsbEdit")
    '                        Exit Sub
    '                    End If
    '                Next
    '            Case Keys.D 'Xóa
    '                For Each ctrl As Control In MainControl.Controls
    '                    If ctrl.GetType.Name = "ToolStrip" Then
    '                        CallClick_ToolStripButton(CType(ctrl, ToolStrip), "tsbDelete")
    '                        Exit Sub
    '                    End If
    '                Next
    '            Case Keys.F 'Tìm kiếm
    '                For Each ctrl As Control In MainControl.Controls
    '                    If ctrl.GetType.Name = "ToolStrip" Then
    '                        CallClick_ToolStripButton(CType(ctrl, ToolStrip), "tsbFind")
    '                        Exit Sub
    '                    End If
    '                Next
    '            Case Keys.A 'Liệt kê tất cả
    '                For Each ctrl As Control In MainControl.Controls
    '                    If ctrl.GetType.Name = "ToolStrip" Then
    '                        CallClick_ToolStripButton(CType(ctrl, ToolStrip), "tsbListAll")
    '                        Exit Sub
    '                    End If
    '                Next
    '            Case Keys.P 'In
    '                For Each ctrl As Control In MainControl.Controls
    '                    If ctrl.GetType.Name = "ToolStrip" Then
    '                        CallClick_ToolStripButton(CType(ctrl, ToolStrip), "tsbPrint")
    '                        Exit Sub
    '                    End If
    '                Next
    '            Case Keys.X 'Xuất Excel
    '                For Each ctrl As Control In MainControl.Controls
    '                    If ctrl.GetType.Name = "ToolStrip" Then
    '                        CallClick_ToolStripButton(CType(ctrl, ToolStrip), "tsbExportToExcel")
    '                        Exit Sub
    '                    End If
    '                Next
    '            Case Keys.M 'Import Excel
    '                For Each ctrl As Control In MainControl.Controls
    '                    If ctrl.GetType.Name = "ToolStrip" Then
    '                        CallClick_ToolStripButton(CType(ctrl, ToolStrip), "tsbImportData")
    '                        Exit Sub
    '                    End If
    '                Next

    '                '******************************************
    '            Case Keys.Q 'Đóng form
    '                'Hình như không có tác dụng
    '                ' Dim tsbcon As Control = MainControl.Controls.Find("tsbClose", True)(0)
    '                For Each ctrl As Control In MainControl.Controls
    '                    If ctrl.GetType.Name = "ToolStrip" Then
    '                        CallClick_ToolStripButton(CType(ctrl, ToolStrip), "tsbClose")
    '                        '  Exit Sub
    '                    End If
    '                Next
    '        End Select




    'End Sub

    Private Sub CallClick_ToolStripButton(ByVal ctrl As ToolStrip, ByVal ItemName As String)
        'Update 06/08/2014: incident 67929 - Lỗi khi dùng phím tắt cho những menu có menu con
        '  If ctrl.GetType.Name <> "ToolStrip" And ctrl.Items.Count > 0 Then Exit Sub

        Try
            '4/4/2015 id 74304 - Lỗi khi thêm mới ứng viên bằng phím Ctrl+N
            ' Các menu có menu con thì không sử dụng phím tắt
            If TypeOf (ctrl.Items(ItemName)) Is ToolStripDropDownButton Then
                If CType(ctrl.Items(ItemName), ToolStripDropDownButton).HasDropDownItems Then
                    Exit Sub
                End If
            End If
            If TypeOf (ctrl.Items(ItemName)) Is ToolStripSplitButton Then
                If CType(ctrl.Items(ItemName), ToolStripSplitButton).HasDropDownItems Then
                    Exit Sub
                End If
            End If

            Dim tsbButton As ToolStripButton = CType(ctrl.Items(ItemName), ToolStripButton)
            If tsbButton IsNot Nothing AndAlso tsbButton.Enabled AndAlso tsbButton.Visible Then
                ctrl.Items(ItemName).PerformClick()
            End If
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "Chỉ cho gõ phím tắt trên lưới"
    ''' <summary>
    ''' Chặn phím tắt của popup menu
    ''' </summary>
    ''' <param name="tdbg"></param>
    ''' <param name="e"></param>
    ''' <returns></returns>
    ''' <remarks>Tại sự kiện của các menu kiểm tra: If CallMenuFromGrid(tdbg, e) = False Then Exit Sub</remarks>
    <DebuggerStepThrough()> _
    Public Function CallMenuFromGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal e As C1.Win.C1Command.ClickEventArgs) As Boolean
        If e Is Nothing Then
            Return True
        End If

        Select Case e.ClickSource
            Case C1.Win.C1Command.ClickSourceEnum.External
                Return False
            Case C1.Win.C1Command.ClickSourceEnum.Shortcut
                If tdbg.Focused = False Then
                    Return False
                End If
            Case C1.Win.C1Command.ClickSourceEnum.None
                Return False
        End Select
        Return True
    End Function

    Public Function CallMenuFromGrid(ByVal sender As Object, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid) As Boolean
        If sender IsNot Nothing Then
            Select Case sender.GetType.Name
                Case "ToolStripMenuItem"
                    Dim ctrl As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
                    If ctrl.Enabled = False Then Return False
                    'Khi không có dữ liệu và đang ở Filter Bar vẫn cho Thêm mới
                    If tdbg.Focused And (ctrl.Name.Contains("Add") OrElse ctrl.Name.Contains("Find") OrElse ctrl.Name.Contains("ListAll")) Then Return True
                Case "ToolStripButton"
                    Dim ctrl As ToolStripButton = CType(sender, ToolStripButton)
                    If tdbg.Focused And (ctrl.Name.Contains("Add") OrElse ctrl.Name.Contains("Find") OrElse ctrl.Name.Contains("ListAll")) Then Return True
            End Select
        End If
        If tdbg.RowCount <= 0 OrElse tdbg.FilterActive OrElse tdbg.Focused = False Then Return False
        Return True
    End Function

#End Region

End Module
