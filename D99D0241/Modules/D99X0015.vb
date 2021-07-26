'#######################################################################################
'#                                     CHÚ Ý
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày cập nhật cuối cùng: 24/12/2013
'# Người cập nhật cuối cùng: Minh Hòa
'# Diễn giải: Thêm hàm SetShortcutPopupMenu chỉ cho ContextMenuStrip cho màn hình dạng mới
'#######################################################################################

''' <summary>
''' Module quản lý các vấn đề về Shortcut và Image của Popupmenu
''' </summary>
''' <remarks>Chỉ gắn vào exe nào có sử dụng popupmenu</remarks>
Public Module D99X0015
#Region "Set Image và Shortcut của ToolStrip"

    ''' <summary>
    ''' Set Image và Shortcut của ToolStripButton cho màn hình dạng mới
    ''' </summary>
    ''' <param name="Form">Me</param>
    ''' <param name="TableToolStrip">System.Windows.Forms.ToolStri</param>
    ''' <param name="bTransactionHistory">Nếu = True thì menu tsbSysInfo có tên là Lịch sử tác động</param>
    ''' <param name="ContextMenuStrip"></param>
    ''' <remarks>Gọi hàm  SetShortcutPopupMenu(Me, TableToolStrip, ContextMenuStrip1) tại form_load</remarks>
    Public Sub SetShortcutPopupMenuNew(ByVal Form As Form, ByVal TableToolStrip As System.Windows.Forms.ToolStrip, ByVal ContextMenuStrip As System.Windows.Forms.ContextMenuStrip, Optional ByVal bTransactionHistory As Boolean = False)
        SetShortcutPopupMenu(Form, TableToolStrip, ContextMenuStrip, bTransactionHistory, ToolStripItemDisplayStyle.ImageAndText)

        'On Error Resume Next
        If TableToolStrip IsNot Nothing Then
            ' Set ShortCut cho TableToolStrip
            With TableToolStrip
                If SCT_BackColor <> "" Then
                    .BackColor = ColorTranslator.FromHtml(SCT_BackColorMenu)
                Else
                    .BackColor = System.Drawing.SystemColors.ActiveBorder
                End If
                For i As Integer = 0 To .Items.Count - 1
                    Select Case .Items(i).Name
                        Case "tsbSave"
                            'Lưu
                            .Items("tsbSave").DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                            .Items("tsbSave").Image = My.Resources.Save
                            .Items("tsbSave").Text = rL3("_Luu") '&Lưu
                            .Items("tsbSave").ToolTipText = .Items("tsbSave").Text.Replace("&", "")
                        Case "tsbNotSave"
                            'Không Lưu
                            .Items("tsbNotSave").DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                            .Items("tsbNotSave").Image = My.Resources.NoSave
                            .Items("tsbNotSave").Text = rL3("_Khong_luu") '&Không lưu
                            .Items("tsbNotSave").ToolTipText = .Items("tsbNotSave").Text.Replace("&", "")
                        Case "tsbBegin"
                            'Đầu tiên
                            .Items("tsbBegin").DisplayStyle = ToolStripItemDisplayStyle.Image
                            .Items("tsbBegin").Image = My.Resources.navigate_beginning
                            .Items("tsbBegin").Text = rL3("Dau_tien") 'Đầu tiên
                            .Items("tsbBegin").ToolTipText = .Items("tsbBegin").Text.Replace("&", "")
                        Case "tsbPrevious"
                            'Trước
                            .Items("tsbPrevious").DisplayStyle = ToolStripItemDisplayStyle.Image
                            .Items("tsbPrevious").Image = My.Resources.navigate_left2
                            .Items("tsbPrevious").Text = rL3("Truoc") 'Trước
                            .Items("tsbPrevious").ToolTipText = .Items("tsbPrevious").Text.Replace("&", "")
                        Case "tsbNext"
                            'Sau
                            .Items("tsbNext").DisplayStyle = ToolStripItemDisplayStyle.Image
                            .Items("tsbNext").Image = My.Resources.navigate_right2
                            .Items("tsbNext").Text = rL3("SauUNI") 'Sau
                            .Items("tsbNext").ToolTipText = .Items("tsbNext").Text.Replace("&", "")
                        Case "tsbEnd"
                            'Cuối cùng
                            .Items("tsbEnd").DisplayStyle = ToolStripItemDisplayStyle.Image
                            .Items("tsbEnd").Image = My.Resources.navigate_end
                            .Items("tsbEnd").Text = rL3("Cuoi_cung") 'Cuối cùng
                            .Items("tsbEnd").ToolTipText = .Items("tsbEnd").Text.Replace("&", "")
                    End Select
                Next

            End With
        End If
    End Sub

#End Region
End Module
