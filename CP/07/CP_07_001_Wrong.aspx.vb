Public Class CP_07_001_Wrong
    Inherits AuthBasePage 'System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then cCreate1()

    End Sub

    Sub cCreate1()
        If Session("MyWrongTable") Is Nothing Then Exit Sub
        Dim dt As DataTable = Session("MyWrongTable")
        PageControler1.SSSDTRID = TIMS.GetRnd6Eng()
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
        Session("MyWrongTable") = Nothing
    End Sub

End Class