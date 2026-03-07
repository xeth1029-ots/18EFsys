Partial Class SD_04_002_Wrong
    Inherits AuthBasePage

    'Dim PageControler1 As New PageControler
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

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