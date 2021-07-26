Imports System.Data.SqlClient
Imports System.Net.Mail 'Add EASendMail namespace

Public Class SendMailLemon3_123
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UserDatabaseID As String, ByVal Password As String, ByVal UserID As String, ByVal Language As EnumLanguage, ByVal Unicode As Boolean)

        gsServer = Server
        gsCompanyID = Database
        gsConnectionUser = UserDatabaseID
        gsUserID = UserID
        gsPassword = Password
        gbUnicode = Unicode

        geLanguage = Language
        gsLanguage = IIf(geLanguage = EnumLanguage.Vietnamese, "84", "01").ToString
        D99C0008.Language = geLanguage

        gsConnectionString = "Data Source=" & gsServer & ";Initial Catalog=" & gsCompanyID & ";User ID=" & gsConnectionUser & ";Password=" & gsPassword & ";Application Name=DigiNet ERP Desktop;Connect Timeout = 0" 'Tạo chuỗi kết nối dùng cho toàn bộ project
        If Not CheckConnection() Then Application.Exit() 'Kiểm tra nối không kết nối được với Server thì END

        GetBackColorObligatory()

    End Sub

    Public Sub New()
        '        gsConnectionString = "Data Source=" & gsServer & ";Initial Catalog=" & gsCompanyID & ";User ID=" & gsConnectionUser & ";Password=" & gsPassword & ";Connect Timeout = 0" 'Tạo chuỗi kết nối dùng cho toàn bộ project
        '        If Not CheckConnection() Then Application.Exit() 'Kiểm tra nối không kết nối được với Server thì END

        GetBackColorObligatory()

    End Sub

    'Private Function CheckConnection() As Boolean
    '    gConn = New SqlConnection(gsConnectionString)
    '    Try
    '        gConn.Open()
    '        gConn.Close()
    '        Return True
    '    Catch ex As Exception
    '        gConn.Close()
    '        D99C0008.MsgInvalidConnection()
    '        MessageBox.Show("Error: " & ex.Message & " - " & ex.Source)

    '        Return False
    '    End Try
    'End Sub


    Public Sub SetupMailServer()
        Dim Frm As New D91F3330
        Frm.ShowDialog()
        Frm.Dispose()
    End Sub


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    '    Public Function SendMail(ByVal sSenderAddress As String, ByVal sReceivedAddress As String, ByVal sCCAddress As String, ByVal sBCCAddress As String, ByVal sSubject As String, ByVal sEmailContent As String, Optional ByVal sAttachmentFile As String = "", Optional ByVal sEmailUser As String = "", Optional ByVal sPassEmailUser As String = "") As Boolean
    '        Cursor.Current = Cursors.WaitCursor
    '
    '        Dim oMail As New MailMessage()
    '        Dim oSmtp As New SmtpClient()
    '        Dim oAtt As Attachment
    '        '
    '        '        Dim oMail As New SmtpMail("TryIt")
    '        '        Dim oSmtp As New SmtpClient()
    '
    '        ' Set sender email address, please change it to yours
    '        oMail.From = sSenderAddress
    '        oMail.To = sReceivedAddress
    '        oMail.Cc = sCCAddress
    '        oMail.Bcc = sBCCAddress
    '        oMail.Subject = sSubject
    '        oMail.TextBody = sEmailContent ' Set email body
    '        If sAttachmentFile <> "" Then oMail.AddAttachment(sAttachmentFile) ' Set Attachment
    '
    '        'Get data from D91T2600
    '        Dim sSQL As String = "select ServerIPAddress, UserAdminEmail, UserPassword, TypeMailServer, PortMailServer from D91T2600"
    '        Dim dt As DataTable = ReturnDataTable(sSQL)
    '        If dt.Rows.Count <= 0 Then
    '            D99C0008.MsgL3(rl3("Mail_server_chua_duoc_thiet_lap"), L3MessageBoxIcon.Err)
    '            dt.Dispose()
    '            Cursor.Current = Cursors.Default
    '            Return False
    '        End If
    '
    '        Dim dr As DataRow = dt.Rows(0)
    '        ' Your SMTP server address
    '
    '        If optNORMAL.Checked = True Then
    '
    '            oSmtp.Host = txtSMTPServer.Text
    '            oSmtp.Port = CType(txtNORMAL_PORT.Text, Integer)
    '
    '        ElseIf optSSL.Checked = True Then
    '
    '            oSmtp.Host = txtSMTPServer.Text
    '            oSmtp.Port = CType(txtSSL_PORT.Text, Integer)
    '            oSmtp.UseDefaultCredentials = False
    '            oSmtp.Credentials = New Net.NetworkCredential(txtUser.Text, txtPassword.Text)
    '            oSmtp.EnableSsl = True
    '
    '            MsgBox("1")
    '
    '        End If
    '
    '        'Dim oServer As New SmtpServer("smtp.gmail.com")
    '        Dim oServer As New SmtpServer(dr("ServerIPAddress").ToString)
    '        Dim sUser As String = sEmailUser
    '        Dim sPass As String = sPassEmailUser
    '        If sUser = "" Then
    '            sUser = dr("UserAdminEmail").ToString
    '        End If
    '        If sPass = "" Then
    '            sPass = D00D0041.D00C0001.EncryptString(dr("UserPassword").ToString, False)
    '        End If
    '
    '        If dr("PortMailServer").ToString <> "" Then
    '            oServer.Port = CInt(dr("PortMailServer"))
    '        Else
    '            oServer.Port = 25
    '        End If
    '        Select Case dr("TypeMailServer").ToString
    '            Case "TLS"
    '                oServer.ConnectType = SmtpConnectType.ConnectSTARTTLS
    '                oServer.User = sUser
    '                oServer.Password = sPass
    '            Case "SSL"
    '                oServer.ConnectType = SmtpConnectType.ConnectDirectSSL
    '                oServer.User = sUser
    '                oServer.Password = sPass
    '            Case Else ' "NO"
    '                oServer.ConnectType = SmtpConnectType.ConnectNormal
    '                'TH Normail không cần User
    '        End Select
    '
    '
    '        ''Nếu người gởi rỗng thì lấy UserAdminEmail của thiết lập
    '        'If sSenderAddress = "" Then sSenderAddress = sUser
    '        '' Set sender email address, please change it to yours
    '        'oMail.From = sSenderAddress
    '        'oMail.To = sReceivedAddress
    '        'oMail.Cc = sCCAddress
    '        'oMail.Bcc = sBCCAddress
    '        'oMail.Subject = sSubject
    '        'oMail.TextBody = sEmailContent ' Set email body
    '        'If sAttachmentFile <> "" Then oMail.AddAttachment(sAttachmentFile) ' Set Attachment
    '
    '        Try
    '
    '            oSmtp.SendMail(oServer, oMail)
    '            'MsgBox("email was sent successfully!", MsgBoxStyle.Information)
    '            Cursor.Current = Cursors.Default
    '            Return True
    '        Catch ex As Exception
    '            D99C0008.MsgL3(rl3("Loi_khi_goi_mail") & vbCrLf & ex.Message, L3MessageBoxIcon.Err)
    '            Cursor.Current = Cursors.Default
    '            Return False
    '        End Try
    '    End Function


    Public Function SendMail(ByVal sSenderAddress As String, ByVal sReceivedAddress As String, ByVal sCCAddress As String, ByVal sBCCAddress As String, ByVal sSubject As String, ByVal sEmailContent As String) As Boolean
        Return SendMail(sSenderAddress, sReceivedAddress, sCCAddress, sBCCAddress, sSubject, sEmailContent, "", "", "", False)
    End Function


    Public Function SendMail(ByVal sSenderAddress As String, ByVal sReceivedAddress As String, ByVal sCCAddress As String, ByVal sBCCAddress As String, ByVal sSubject As String, ByVal sEmailContent As String, ByVal sAttachmentFile As String) As Boolean
        Return SendMail(sSenderAddress, sReceivedAddress, sCCAddress, sBCCAddress, sSubject, sEmailContent, sAttachmentFile, "", "", False)
    End Function

    Public Function SendMail(ByVal sSenderAddress As String, ByVal sReceivedAddress As String, ByVal sCCAddress As String, ByVal sBCCAddress As String, ByVal sSubject As String, ByVal sEmailContent As String, ByVal sAttachmentFile As String, ByVal sEmailUser As String) As Boolean
        Return SendMail(sSenderAddress, sReceivedAddress, sCCAddress, sBCCAddress, sSubject, sEmailContent, sAttachmentFile, sEmailUser, "", False)
    End Function

    Public Function SendMail(ByVal sSenderAddress As String, ByVal sReceivedAddress As String, ByVal sCCAddress As String, ByVal sBCCAddress As String, ByVal sSubject As String, ByVal sEmailContent As String, ByVal sAttachmentFile As String, ByVal sEmailUser As String, ByVal sPassEmailUser As String) As Boolean
        Return SendMail(sSenderAddress, sReceivedAddress, sCCAddress, sBCCAddress, sSubject, sEmailContent, sAttachmentFile, sEmailUser, sPassEmailUser, False)
    End Function


    'Public Function SendMail(ByVal sSenderAddress As String, ByVal sReceivedAddress As String, ByVal sCCAddress As String, ByVal sBCCAddress As String, ByVal sSubject As String, ByVal sEmailContent As String, Optional ByVal sAttachmentFile As String = "", Optional ByVal sEmailUser As String = "", Optional ByVal sPassEmailUser As String = "") As Boolean
    Public Function SendMail(ByVal sSenderAddress As String, ByVal sReceivedAddress As String, ByVal sCCAddress As String, ByVal sBCCAddress As String, ByVal sSubject As String, ByVal sEmailContent As String, ByVal sAttachmentFile As String, ByVal sEmailUser As String, ByVal sPassEmailUser As String, ByVal bMsg As Boolean) As Boolean
        Dim oMail As New MailMessage()
        Dim oSmtp As New SmtpClient()
        Dim oAtt As Attachment
        Dim sUser As String = sEmailUser '"minhhoa@diginet.com.vn" '
        Dim sPass As String = sPassEmailUser

        Try
            'Get data from D91T2600
            Dim sSQL As String = "select ServerIPAddress, UserAdminEmail, UserPassword, TypeMailServer, PortMailServer from D91T2600"
            Dim dt As DataTable = ReturnDataTable(sSQL)
            If dt.Rows.Count <= 0 Then
                D99C0008.MsgL3(rL3("Mail_server_chua_duoc_thiet_lap"), L3MessageBoxIcon.Err)
                dt.Dispose()
                Return False
            End If
            Dim dr As DataRow = dt.Rows(0)

            If sUser = "" Then
                sUser = dr("UserAdminEmail").ToString
            End If
            If sPass = "" Then
                sPass = D00D0041.D00C0001.EncryptString(dr("UserPassword").ToString, False)
            End If

            oSmtp.Host = dr("ServerIPAddress").ToString '"10.0.0.14" '
            Select Case dr("TypeMailServer").ToString
                Case "SSL"
                    If dr("PortMailServer").ToString <> "" Then
                        oSmtp.Port = CInt(dr("PortMailServer"))
                    Else
                        oSmtp.Port = 587
                    End If
                    oSmtp.UseDefaultCredentials = False
                    oSmtp.Credentials = New Net.NetworkCredential(sUser, sPass)
                    oSmtp.EnableSsl = True
                Case "NOSSL"
                    If dr("PortMailServer").ToString <> "" Then
                        oSmtp.Port = CInt(dr("PortMailServer"))
                    Else
                        oSmtp.Port = 25
                    End If
                    oSmtp.UseDefaultCredentials = False
                    oSmtp.Credentials = New Net.NetworkCredential(sUser, sPass)
                    oSmtp.EnableSsl = False
                Case Else ' "NO"
                    If dr("PortMailServer").ToString <> "" Then
                        oSmtp.Port = CInt(dr("PortMailServer"))
                    Else
                        oSmtp.Port = 25
                    End If
                    'TH Normail không cần User
            End Select

            oMail = New MailMessage()
            'Nếu người gởi rỗng thì lấy UserAdminEmail của thiết lập
            If sSenderAddress = "" Then sSenderAddress = sUser

            sReceivedAddress = sReceivedAddress.Replace(";", ",")
            sCCAddress = sCCAddress.Replace(";", ",")

            oMail.From = New MailAddress(sSenderAddress)
            oMail.To.Add(sReceivedAddress)

            If sCCAddress <> "" Then oMail.CC.Add(sCCAddress)
            If sBCCAddress <> "" Then oMail.Bcc.Add(sBCCAddress)
            oMail.Subject = sSubject
            oMail.Body = sEmailContent
            oMail.IsBodyHtml = True
            If sAttachmentFile <> "" Then
                Dim sfile() As String = sAttachmentFile.Split(";"c)
                For i As Integer = 0 To sfile.Length - 1
                    oAtt = New Attachment(sfile(i))
                    oMail.Attachments.Add(oAtt)
                Next
            End If

            'MessageBox.Show("Send(oMail123)")

            oSmtp.Send(oMail)
            'MessageBox.Show("Send (End)")

            'Cursor.Current = Cursors.Default
            Return True
        Catch ex As Exception
            If bMsg Then
                MessageBox.Show(ex.Message & " - " & ex.Source)
                '& vbCrLf & " oSmtp.Host = " & oSmtp.Host & " oSmtp.Port = " & oSmtp.Port _
                '& vbCrLf & " sUser = " & sUser & " sPass = " & sPass & vbCrLf & "sSenderAddress = " & sSenderAddress & "sReceivedAddress = " & sReceivedAddress & "sCCAddress = " & sCCAddress)
            End If
            'Cursor.Current = Cursors.Default
            Return False

        Finally
            oMail.Dispose()
            oMail = Nothing
            oSmtp = Nothing
        End Try


    End Function

End Class
