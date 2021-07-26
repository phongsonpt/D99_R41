'#######################################################################################
'#                                     CHÚ Ý
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày cập nhật cuối cùng: 01/10/2010 
'# Diễn giải: Kiểm tra sản phẩm Lemon3
'# Người cập nhật cuối cùng: Nguyễn Thị Minh Hòa
'#######################################################################################

Imports System.IO

Public Module D99X0010

#Region "Khai báo Enum"
    ''' <summary>
    ''' Enum trả về sản phẩm của Lemon3
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum EnumDisplayProduct
        ''' <summary>
        ''' LEMON3-ERP
        ''' </summary>
        ''' <remarks></remarks>
        L3ERP = 0
        ''' <summary>
        ''' LemonHR
        ''' </summary>
        ''' <remarks></remarks>
        L3HR = 1
        ''' <summary>
        ''' LEMON-LemonFinance
        ''' </summary>
        ''' <remarks></remarks>
        L3F = 2
    End Enum
#End Region

#Region "Khai báo biến toàn cục"
    ''' <summary>
    ''' Trả về sản phẩm của Lemon3
    ''' </summary>
    ''' <remarks></remarks>
    Public geDisplayProduct As EnumDisplayProduct
#End Region

#Region "Public Sub and Function"

#Region "Hàm trả về sản phẩm của Lemon"


    Public Sub ReturnDisplayProduct() 'Dùng cho exe chính
        'Update 27/06/2014 Bỏ đoạn code này, do đã gán biến này trong hàm GetCodeTable
        ' geDisplayProduct = CType(GetSetting("Lemon3", "Settings", "DisplayProduct", "0"), EnumDisplayProduct)
        Dim dt As DataTable = ReturnDataTable(SQLGeneral.SQLSelectD91T8888())
        If dt.Rows.Count > 0 Then
            geDisplayProduct = CType(dt.Rows(0).Item("DisplayProduct"), EnumDisplayProduct)
        End If
        UseModuleD54()
    End Sub

    'Public Sub LoadBitmapNew(ByVal BVietname As String, ByVal BEnglish As String, ByRef picMain As PictureBox, ByRef lblBitmap As Label, ByVal sCaptionForm As String, Optional ByVal bLoadFist As Boolean = False)
    '    'Update 11/11/2010: Nếu sản phẩm là Lemon3Finance thì mới load Bitmap mới
    '    If geDisplayProduct <> EnumDisplayProduct.L3F Then
    '        lblBitmap.Visible = False
    '        LoadBitmapOld(BVietname, BEnglish, picMain)
    '        Exit Sub
    '    End If
    '    If bLoadFist = True OrElse lblBitmap.Visible = False Then
    '        SetLableBitmap(lblBitmap)
    '        LoadBitmapOld(BVietname, BEnglish, picMain)

    '        ''Dim bmp As New Bitmap(sBitmap)
    '        'Dim bmp As Bitmap
    '        'bmp = ShowImage(sBitmap)
    '        ''picMain.BackgroundImage = bmp

    '        Dim sPathBitmap As String = My.Application.Info.DirectoryPath & "\Bitmap\"
    '        Dim sBitmap As String = sPathBitmap
    '        'Dim workingRectangle As System.Drawing.Rectangle = Screen.PrimaryScreen.Bounds
    '        'Dim sSize As String = workingRectangle.Width.ToString & "x" & workingRectangle.Height.ToString & ".jpg"
    '        'Select Case geDisplayProduct
    '        '    Case EnumDisplayProduct.L3ERP
    '        '        sBitmap &= "L3ERP" & sSize
    '        '    Case EnumDisplayProduct.L3HR
    '        '        sBitmap &= "L3HR" & sSize
    '        '    Case EnumDisplayProduct.L3F
    '        '        sBitmap &= "L3F" & sSize
    '        'End Select
    '        'Update 11/11/2010: Load bitmap theo kiểu mới, lấy theo ảnh có kích thước 1024x768
    '        sBitmap &= "L3F1024x768.jpg"

    '        Dim bmp As Bitmap
    '        If Not My.Computer.FileSystem.FileExists(sBitmap) Then
    '            lblBitmap.Visible = False
    '            D99C0008.MsgL3("Không tồn tại bitmap" & Space(1) & sBitmap, L3MessageBoxIcon.Exclamation)
    '            Exit Sub
    '        Else
    '            lblBitmap.Visible = True
    '            picMain.SizeMode = PictureBoxSizeMode.StretchImage
    '            bmp = New Bitmap(sBitmap)
    '            picMain.Image = bmp
    '        End If
    '    End If

    '    If geLanguage = EnumLanguage.Vietnamese Then
    '        lblBitmap.Text = "Coâng cuï duïng cuï"
    '    Else
    '        Dim sCaptionForm As String = Me.Text.Substring(0, Me.Text.IndexOf("-"))
    '        lblBitmap.Text = sCaptionForm
    '    End If

    '    ''Update 12/08/2010: Load bitmap theo kiểu mới
    '    'Dim workingRectangle As System.Drawing.Rectangle = Screen.PrimaryScreen.Bounds
    '    'Dim sSize As String = workingRectangle.Width.ToString & "x" & workingRectangle.Height.ToString & ".jpg"

    '    'picMain.SizeMode = PictureBoxSizeMode.StretchImage
    '    'LoadBitmapOld(BVietname, BEnglish, picMain)

    '    ''Dim sPathBitmap As String = My.Application.Info.DirectoryPath & "\Bitmap\"
    '    'Dim sBitmap As String '= sPathBitmap
    '    'sBitmap = ""
    '    ''Dim iDisplayProduct As Integer = CType(GetSetting("Lemon3", "Settings", "DisplayProduct", "0"), Integer)

    '    'sBitmap &= geDisplayProduct.ToString & sSize

    '    ''Select Case iDisplayProduct
    '    ''    Case 0 ' ERP
    '    ''        'sBitmap &= "L3ERP1024x768.jpg"
    '    ''        sBitmap &= "L3ERP" & sSize
    '    ''    Case 1 'HR
    '    ''        'sBitmap &= "L3HR1024x768.jpg"
    '    ''        sBitmap &= "L3HR" & sSize
    '    ''    Case 2 'Financials
    '    ''        'sBitmap &= "L3F1024x768.jpg"
    '    ''        sBitmap &= "L3F" & sSize
    '    ''End Select

    'End Sub

    'Private Sub SetLableBitmap(ByRef lblBitmap As Label)
    '    lblBitmap.Font = New System.Drawing.Font("VNI-Ariston", 39.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    '    lblBitmap.Location = New System.Drawing.Point(0, 103)
    '    lblBitmap.Size = New System.Drawing.Size(759, 75)
    '    lblBitmap.Text = ""
    '    lblBitmap.AutoSize = False
    '    lblBitmap.BackColor = ColorTranslator.FromHtml("#eba102")
    '    lblBitmap.ForeColor = Color.White
    '    lblBitmap.TextAlign = ContentAlignment.MiddleCenter
    '    lblBitmap.Dock = DockStyle.Bottom
    'End Sub

    'Private Sub LoadBitmapOld(ByVal BVietname As String, ByVal BEnglish As String, ByRef picMain As PictureBox)
    '    If geLanguage = EnumLanguage.Vietnamese Then
    '        If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & BVietname) Then
    '            Dim bmp As New Bitmap(My.Application.Info.DirectoryPath & BVietname)
    '            'picMain.BackgroundImage = bmp
    '            picMain.Image = bmp
    '        End If
    '    ElseIf geLanguage = EnumLanguage.English Then
    '        If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & BEnglish) Then
    '            Dim bmp As New Bitmap(My.Application.Info.DirectoryPath & BEnglish)
    '            'picMain.BackgroundImage = bmp
    '            picMain.Image = bmp
    '        End If
    '    End If
    'End Sub

    'Private Function ShowImage(ByVal FileName As String) As Bitmap
    '    Dim FILE_PATH As String = Application.StartupPath & "\filebitmap.txt"
    '    Dim FILE_INDEX_PATH As String = Application.StartupPath & "\filei.txt"

    '    'Get buffer index of image content
    '    Dim fileID As String = FileName 'cboImageList.Text
    '    If String.IsNullOrEmpty(fileID) Then
    '        'MessageBox.Show("Chưa chọn id ảnh nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        'MessageBox.Show("There's not image's id choose", "Anounment", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Return Nothing
    '    End If

    '    If Not File.Exists(FILE_INDEX_PATH) Then Return Nothing

    '    Dim rStream As New StreamReader(FILE_INDEX_PATH)

    '    Dim nStart, nCount As Integer

    '    While Not rStream.EndOfStream
    '        Dim sIndex As String
    '        sIndex = rStream.ReadLine()
    '        If (sIndex.Split(",")(0) = fileID) Then
    '            nStart = sIndex.Split(",")(1)
    '            nCount = sIndex.Split(",")(2)
    '            Exit While
    '        End If
    '    End While

    '    'picImageReader.Image = Nothing
    '    Dim buffer(nStart + nCount) As Byte
    '    Dim stream As New FileStream(FILE_PATH, IO.FileMode.Open)

    '    stream.Read(buffer, 0, nStart + nCount)

    '    stream = New FileStream(Application.StartupPath + "\\temp.txt", IO.FileMode.Create)

    '    Dim writer As New BinaryWriter(stream)
    '    writer.Write(buffer, nStart, nCount)
    '    Dim image As New System.Drawing.Bitmap(stream)
    '    'picImageReader.Image = image
    '    stream.Close()
    '    stream = Nothing
    '    rStream = Nothing
    '    GC.Collect()
    '    If File.Exists(Application.StartupPath + "\\temp.txt") Then
    '        File.Delete(Application.StartupPath + "\\temp.txt")
    '    End If
    '    Return image
    'End Function


#End Region

#End Region

End Module
