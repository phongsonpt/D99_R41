Namespace Lemon3.Controls.Customs
    Public Class L3WDateEdit : Inherits DevExpress.XtraEditors.DateEdit

        Public Sub New()
            MyBase.New()
            InputDate("dd/MM/yyyy")
        End Sub

        Public Sub InputDate(ByVal numberFormat As String)
            Me.Properties.EditMask = numberFormat
            Me.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTime
            Me.Properties.Mask.EditMask = numberFormat
            Me.Properties.EditFormat.FormatString = numberFormat
            'Me.MaskUseAsDisplayFormat = True
            Me.Properties.Mask.UseMaskAsDisplayFormat = True
            AddHandler Me.DoubleClick, AddressOf Me_DoubleClick
        End Sub

        Private Sub Me_DoubleClick(sender As Object, e As EventArgs) Handles Me.DoubleClick
            If Me.ReadOnly = False AndAlso (Me.EditValue Is Nothing OrElse Me.EditValue.ToString = "") Then Me.EditValue = Now.Date
        End Sub
    End Class
End Namespace



