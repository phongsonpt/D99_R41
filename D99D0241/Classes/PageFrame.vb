Imports System
Imports System.ComponentModel
Imports System.ComponentModel.Design.Serialization
Imports System.Drawing
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Forms
Imports System.Windows.Markup
Imports System.Windows.Forms.Integration
Public Class PageFrame

    ' Fields...

    Private _fPage As Page

    Public Property FPage As Page
        Get
            Return _fPage
        End Get
        Set(ByVal Value As Page)
            _fPage = Value
        End Set
    End Property
    
    

    Private Sub PageFrame_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Add any initialization after the InitializeComponent() call.
        Dim host As New ElementHost
        Dim fFrame As New Frame

        fFrame.Navigate(_fPage)
        If _fPage.Width > 1000 Then 'Set kích thước Auto cho form lớn
            _fPage.Width = Double.NaN
            _fPage.Height = Double.NaN
            host.Dock = DockStyle.Fill
        Else 'Đối với form nhỏ thì ko Dock + Set kích thước = kích thước Page cần gọi tới
            host.Dock = DockStyle.None
            host.Height = L3Int(_fPage.Height)
            host.Width = L3Int(_fPage.Width)
            Me.Width = L3Int(_fPage.Width)
        End If
        host.Child = fFrame
        Me.Controls.Add(host)
        Me.Text = _fPage.Title
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        
    End Sub
End Class