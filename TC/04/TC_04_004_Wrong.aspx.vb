Partial Class TC_04_004_Wrong
    Inherits AuthBasePage

    'Dim PageControler1 As New PageControler

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'PageControler1 = Me.FindControl("PageControler1")
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            create()
            Session("MyWrongTable") = Nothing
        End If
    End Sub

    Sub create()
        Dim dt As DataTable = Session("MyWrongTable")
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub
End Class