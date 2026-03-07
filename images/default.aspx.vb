Public Class _default1
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim objconn As SqlConnection = DbAccess.GetConnection()
        Call TIMS.sUtl_404NOTFOUND(Me, objconn)
    End Sub

End Class