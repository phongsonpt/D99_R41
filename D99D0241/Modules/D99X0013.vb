'#######################################################################################
'#                                     Hàm đọc cấu hình máy
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày tạo: 05/06/2013 
'# Ngày cập nhật cuối cùng:  05/06/2013
'# Người cập nhật cuối cùng: Minh Hòa
'# Diễn giải: Hàm đọc cấu hình máy cần Reference đến dll System.System.Management
'# Thêm hàm GetInformationPC

'#######################################################################################
Imports system.IO
Imports System.Security.Cryptography
'Imports System.Management

Public Module D99X0013

#Region "Lấy cấu hình máy hiện tại: RAM, OS, CPU"

    ''' <summary>
    ''' Lấy thông tin cấu hình của máy tính hiện tại
    ''' </summary>
    ''' <param name="sRAM">Lấy RAM</param>
    ''' <param name="sOS">Lấy hệ điều hành</param>
    ''' <param name="sCPU">Lấy CPU</param>
    ''' <remarks>Nếu truyền vào tham số nào thì sẽ trả ra tham số đó</remarks>
    Public Sub GetInformationPC(Optional ByRef sRAM As String = "", Optional ByRef sOS As String = "", Optional ByRef sCPU As String = "")
        sRAM = Math.Round(My.Computer.Info.TotalPhysicalMemory / 1024 / 1024 / 1024)

        'Dim mbs As ManagementObjectSearcher
        'Try
        '    sOS = ""
        '    mbs = New ManagementObjectSearcher("Select * from Win32_OperatingSystem")
        '    For Each mo As ManagementObject In mbs.Get
        '        sOS = sOS & mo("Caption").ToString & ", " & mo("CSDVersion").ToString & " - "
        '    Next
        '    sOS = Microsoft.VisualBasic.Strings.Left(sOS, Len(sOS) - Len(" - "))
        'Catch ex As Exception
        '    'sOS = sOS & "<Empty>" & " - "
        '    'sOS = Left(sOS, Len(sOS) - Len(" - "))
        'End Try

        'Try
        '    sCPU = ""
        '    mbs = New ManagementObjectSearcher("Select * from Win32_Processor")
        '    For Each mo As ManagementObject In mbs.Get
        '        If sCPU <> mo("Name").ToString & " - " Then
        '            sCPU = sCPU & mo("Name").ToString & " - "
        '        End If
        '    Next
        '    sCPU = sCPU.Substring(0, sCPU.Length - Len(" - "))
        'Catch ex As Exception
        '    'sCPU = sCPU & "<Empty>" & " - "
        '    'sCPU = Left(sCPU, Len(sCPU) - Len(" - "))
        'End Try

        'Try
        '    sRAM = ""
        '    mbs = New ManagementObjectSearcher("Select * from Win32_LogicalMemoryConfiguration")
        '    For Each mo As ManagementObject In mbs.Get
        '        sRAM = sRAM & mo("TotalPhysicalMemory").ToString & " - "
        '    Next
        '    sRAM = sRAM.Substring(0, sRAM.Length - Len(" - "))
        '    'Đổi KB ra GB
        '    If sRAM <> "" Then
        '        sRAM = Math.Round((CDbl(sRAM) / 1024) / 1024)
        '    End If


        'Catch ex As Exception
        '    'sRAM = sRAM & "<Empty>" & " - "
        '    'sRAM = Left(sRAM, Len(sRAM) - Len(" - "))
        'End Try

    End Sub
#End Region

End Module
