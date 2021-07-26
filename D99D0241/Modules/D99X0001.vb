'#######################################################################################
'#                                     CHÚ Ý
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày cập nhật cuối cùng: 18/06/2013
'# Người cập nhật cuối cùng: Nguyễn Thị Minh Hòa
'# Diễn giả: bổ sung resource thay thế cho mỗi khách hàng: vd  replace tên "Phòng" thành "Xưởng"
'#######################################################################################
Imports System.Resources
Imports System.IO

''' <summary>
''' Module quản lý các vấn đề về Resource
''' </summary>
Public Module D99X0001
    ''' <summary>
    ''' Lưu trữ Resource tiếng Việt
    ''' </summary>
    Private rV As ResourceManager '= ResourceManager.CreateFileBasedResourceManager("Vietnamese", Application.StartupPath, Nothing)
    ''' <summary>
    ''' Lưu trữ Resource tiếng Anh
    ''' </summary>
    Private rE As ResourceManager '= ResourceManager.CreateFileBasedResourceManager("English", Application.StartupPath, Nothing)
    Private rJ As ResourceManager
    Private rC As ResourceManager
    Private rO As ResourceManager

    '    ''' <summary>
    '    ''' Trả về chuỗi resource ứng với ResourceID truyền vào
    '    ''' </summary>
    '    ''' <param name="ResourceID">Mã Resource</param>
    '    '''<remarks>Nếu không tìm thấy ResourceID ở file Tiếng Anh thì sẽ
    '    ''' trả về Resouce ở file Tiếng Việt, và nếu không tìm thấy Resource ở 
    '    ''' file Tiếng Việt sẽ trả về Mã ResourceID này
    '    ''' </remarks>
    '    Public Function rL3(ByVal ResourceID As String) As String
    '        Try
    '
    '            Dim sRes As String = ""
    '            If geLanguage = EnumLanguage.Vietnamese Then
    '                sRes = rV.GetString(ResourceID).ToString
    '            Else
    '                sRes = rE.GetString(ResourceID).ToString
    '            End If
    '
    '            'Update 17/06/2013: Nếu có thay đổi tên resource
    '            If giReplacResource <> 0 AndAlso gsConnectionString <> "" Then
    '                sRes = ReplaceResourceCustom(sRes)
    '            End If
    '
    '            Return sRes
    '
    '        Catch
    '            Try
    '                'Update 17/06/2013: Nếu đã thông báo lỗi 1 lần thì không cần thông báo nữa
    '                If Not gbResourceError Then
    '                    gbResourceError = True
    '                    If geLanguage = EnumLanguage.Vietnamese Then
    '                        If MessageBox.Show("˜Ò nghÜ li£n hÖ nhª cung cÊp ¢Ó cËp nhËt mìi 2 file ng¤n ngö: Vietnamese.resources vª English.resources" & vbCrLf & _
    '                            "BÁn câ muçn tiÕp tóc kh¤ng?" & vbCrLf & _
    '                            "Læi: [" + ResourceID + "]", MsgAnnouncement, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
    '                            Return "[" + rV.GetString(ResourceID).ToString + "]"
    '                        Else
    '                            Application.Exit()
    '                        End If
    '                    Else
    '                        If MessageBox.Show("Please contact to provider for update two resource file: Vietnamese.resource and English.resource" & vbCrLf & _
    '                            "Do you want to continue?" & vbCrLf & _
    '                            "Error: [" + ResourceID + "]", MsgAnnouncement, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
    '                            Return "[" + rV.GetString(ResourceID).ToString + "]"
    '                        Else
    '                            Application.Exit()
    '                        End If
    '                    End If
    '                Else
    '                    Return "[" + rV.GetString(ResourceID).ToString + "]"
    '                End If
    '            Catch ex As Exception
    '                If Len(ResourceID) > 74 Then
    '                    Return "[" + ResourceID.Substring(0, 70) + "...]"
    '                Else
    '                    Return "[" + ResourceID + "]"
    '                End If
    '            End Try
    '        End Try
    '        Return ResourceID
    '        'CheckFontCaptionIsDGSansSerif
    '    End Function
    Private Sub GetResource(ByRef RS As ResourceManager, resources As String)
        If File.Exists(gsApplicationSetup & "\" & resources & ".resources") Then
            RS = ResourceManager.CreateFileBasedResourceManager(resources, gsApplicationSetup, Nothing)
        Else
            RS = rE
        End If
    End Sub
    ''' <summary>
    ''' Trả về chuỗi resource ứng với ResourceID truyền vào
    ''' </summary>
    ''' <param name="ResourceID">Mã Resource</param>
    '''<remarks>Nếu không tìm thấy ResourceID ở file Tiếng Anh thì sẽ
    ''' trả về Resouce ở file Tiếng Việt, và nếu không tìm thấy Resource ở 
    ''' file Tiếng Việt sẽ trả về Mã ResourceID này
    ''' </remarks>
    Public Function rL3(ByVal ResourceID As String) As String 'Hàm này sửa lại bổ sung thêm ngôn ngữ J, C, Other 26/07/2018
        Try
            'CheckFontResourceDG()
            If giCheckFontDGSansSerif = -1 Then
                giCheckFontDGSansSerif = 0
                rV = ResourceManager.CreateFileBasedResourceManager("Vietnamese", gsApplicationSetup, Nothing)
                rE = ResourceManager.CreateFileBasedResourceManager("English", gsApplicationSetup, Nothing)
                GetResource(rC, "Chinese")
                GetResource(rJ, "Japanese")
                GetResource(rO, "Other")
            End If
            Dim sRes As String = ""
            Select Case geLanguage
                Case EnumLanguage.Vietnamese
                    sRes = rV.GetString(ResourceID).ToString
                Case EnumLanguage.Japanese
                    sRes = rJ.GetString(ResourceID).ToString
                Case EnumLanguage.Chinese
                    sRes = rC.GetString(ResourceID).ToString
                Case EnumLanguage.Other
                    sRes = rO.GetString(ResourceID).ToString
                Case Else
                    sRes = rE.GetString(ResourceID).ToString
            End Select

            'Update 17/06/2013: Nếu có thay đổi tên resource
            If (giReplacResource <> 0 OrElse giReplacResourceERP <> 0) AndAlso gsConnectionString <> "" Then
                sRes = ReplaceResourceCustom(sRes)
            End If

            Return sRes

        Catch ex1 As Exception
            WriteLogFile("rl3:" & ex1.Message)
            Try
                'Update 17/06/2013: Nếu đã thông báo lỗi 1 lần thì không cần thông báo nữa
                If Not gbResourceError Then
                    gbResourceError = True
                    Dim Message As String = "Cập nhật mới các file ngôn ngữ (*.resources)" & vbCrLf & _
                                            "Bạn có muốn tiếp tục không?" & vbCrLf & _
                                            "Lỗi: [" + ResourceID + "]"
                    If geLanguage <> EnumLanguage.Vietnamese Then
                        Message = "Please update resources file (*.resources)" & vbCrLf & _
                            "Do you want to continue?" & vbCrLf & _
                            "Error: [" + ResourceID + "]"
                    End If
                    If MessageBox.Show(Message, MsgAnnouncement, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                        Return "[" + ResourceID + "]"
                    Else
                        Application.Exit()
                    End If
                    '            If geLanguage = EnumLanguage.Vietnamese Then
                    '                'If MessageBox.Show("˜Ò nghÜ li£n hÖ nhª cung cÊp ¢Ó cËp nhËt mìi 2 file ng¤n ngö: Vietnamese.resources vª English.resources" & vbCrLf & _
                    '                '    "BÁn câ muçn tiÕp tóc kh¤ng?" & vbCrLf & _
                    '                '    "Læi: [" + ResourceID + "]", MsgAnnouncement, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                    '                '    Return "[" + rV.GetString(ResourceID).ToString + "]"
                    '                'Else
                    '                '    Application.Exit()
                    '                'End If
                    '                'If File.Exists(Application.StartupPath & "\VietnameseDG.resources") And File.Exists(Application.StartupPath & "\EnglishDG.resources") Then

                    '                'If giCheckFontDGSansSerif = 1 Then
                    '                '    If MessageBox.Show("CËp nhËt mìi cÀc file ng¤n ngö (Vietnamese*.resources, English*.resources)" & vbCrLf & _
                    '                '        "BÁn câ muçn tiÕp tóc kh¤ng?" & vbCrLf & _
                    '                '        "Læi: [" + ResourceID + "]", MsgAnnouncement, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                    '                '        Return "[" + rV.GetString(ResourceID).ToString + "]"
                    '                '    Else
                    '                '        Application.Exit()
                    '                '    End If

                    '                'Else
                    '                If MessageBox.Show("Cập nhật mới các file ngôn ngữ (Vietnamese*.resources, English*.resources)" & vbCrLf & _
                    '"Bạn có muốn tiếp tục không?" & vbCrLf & _
                    '"Lỗi: [" + ResourceID + "]", MsgAnnouncement, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                    '                    Return "[" + rV.GetString(ResourceID).ToString + "]"
                    '                Else
                    '                    Application.Exit()
                    '                End If
                    '                'End If
                    '            Else
                    '                If MessageBox.Show("Please update resources file (Vietnamese*.resources, English*.resources)" & vbCrLf & _
                    '                    "Do you want to continue?" & vbCrLf & _
                    '                    "Error: [" + ResourceID + "]", MsgAnnouncement, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                    '                    Return "[" + rV.GetString(ResourceID).ToString + "]"
                    '                Else
                    '                    Application.Exit()
                    '                End If
                    '            End If
                Else
                    ' Return "[" + rV.GetString(ResourceID).ToString + "]"
                    Return "[" + ResourceID + "]"
                End If
            Catch ex As Exception
                If Len(ResourceID) > 74 Then
                    Return "[" + ResourceID.Substring(0, 70) + "...]"
                Else
                    Return "[" + ResourceID + "]"
                End If
            End Try

        End Try
        Return ResourceID
        'CheckFontCaptionIsDGSansSerif
    End Function

    'Dim iCheckFontDGSansSerif As Integer = -1

    'Private Sub CheckFontResourceDG()
    '    'If giCheckFontDGSansSerif <> -1 Then Exit Sub 'chỉ khởi tạo load file resource lần đầu tiên
    '    ''        If iCheckFontDGSansSerif <> -1 Then Exit Sub 'Khởi tạo lần đầu tiên kiểm tra Font DGSansSerif

    '    ''        'MessageBox.Show("iCheckFontDGSansSerif = " & iCheckFontDGSansSerif)

    '    ''        If CheckFontCaptionIsDGSansSerif() Then
    '    ''            iCheckFontDGSansSerif = 1
    '    ''        Else
    '    ''            iCheckFontDGSansSerif = 0
    '    ''        End If
    '    'giCheckFontDGSansSerif = 0 ' iCheckFontDGSansSerif luôn dùng MS Sanserif

    '    ''        'Thay Application.StartupPath thành gsApplicationPath 22/03/2016
    '    ''        If iCheckFontDGSansSerif = 1 Then
    '    ''            If File.Exists(gsApplicationSetup & "\VietnameseDG.resources") And File.Exists(gsApplicationSetup & "\EnglishDG.resources") Then
    '    ''                rV = ResourceManager.CreateFileBasedResourceManager("VietnameseDG", gsApplicationSetup, Nothing)
    '    ''                rE = ResourceManager.CreateFileBasedResourceManager("EnglishDG", gsApplicationSetup, Nothing)
    '    ''                Exit Sub
    '    ''            Else
    '    ''                GoTo RsUnicode
    '    ''            End If
    '    ''        End If

    '    ''RsUnicode:
    '    'rV = ResourceManager.CreateFileBasedResourceManager("Vietnamese", gsApplicationSetup, Nothing)
    '    'rE = ResourceManager.CreateFileBasedResourceManager("English", gsApplicationSetup, Nothing)
    '    'rC = ResourceManager.CreateFileBasedResourceManager("Chinese", gsApplicationSetup, Nothing)
    '    'rJ = ResourceManager.CreateFileBasedResourceManager("Japanese", gsApplicationSetup, Nothing)
    '    'rO = ResourceManager.CreateFileBasedResourceManager("Other", gsApplicationSetup, Nothing)

    'End Sub

End Module
