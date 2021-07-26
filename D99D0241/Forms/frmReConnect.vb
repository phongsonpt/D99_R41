Imports System
Public NotInheritable Class frmReConnect

    'TODO: This form can easily be set as the splash screen for the application by going to the "Application" tab
    '  of the Project Designer ("Properties" under the "Project" menu).

    Private _eSC As Boolean = False
    Public ReadOnly Property ESC() As Boolean
        Get
            Return _eSC
        End Get
    End Property

    Private _connection As SqlConnection
    Public WriteOnly Property Connection() As SqlConnection
        Set(ByVal value As SqlConnection)
            _connection = value
        End Set
    End Property

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        _eSC = True
        tmrWait.Enabled = False
        Me.Close()
    End Sub

    Private bClose As Boolean = False
    Dim bRun As Boolean = True
    Private Sub ThreadProc() 'Làm cho cột Ca thực tế
        Try
ReConect:
            Dim conn As New SqlConnection(_connection.ConnectionString)
            conn.Open()
            _eSC = False
            bClose = True
        Catch ex As SqlException
            If ex.Number = 10054 OrElse ex.Number = 1231 _
              OrElse ex.Message.Contains("Could not open a connection to SQL Server") _
              OrElse ex.Message.Contains("The server was not found or was not accessible") _
              OrElse ex.Message.Contains("A transport-level error") Then 'Lỗi không kết nối được server
                GoTo ReConect
            Else
                _eSC = False
                bClose = True
            End If
        End Try
    End Sub

    Private Sub tmrWait_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrWait.Tick
        If Not bRun Then
            If bClose Then
                Me.Close()
                tmrWait.Enabled = False
            End If
            Exit Sub
        End If
        bRun = False
        Dim t As New Threading.Thread(AddressOf ThreadProc)
        t.IsBackground = True
        t.Start()
    End Sub

    Private Sub frmReConnect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '_SQLConnectionString = gsConnectionString.Replace("Connection Timeout = 0", "Connection Timeout = 15")
    End Sub
End Class
