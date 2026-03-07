Partial Class SD_02_001_Wrong
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        'PageControler1 = Me.FindControl("PageControler1")
        pagecontroler1.PageDataGrid = datagrid1
        If Not IsPostBack Then
            create()
            Session("MyWrongTable") = Nothing
        End If
    End Sub

    Sub create()
        Dim dt As DataTable = Session("MyWrongTable")
        pagecontroler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub
End Class