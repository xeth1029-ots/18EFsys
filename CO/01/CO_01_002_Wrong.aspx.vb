Partial Class CO_01_002_Wrong
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
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
        'If pagecontroler1.SSSDTRID <> "" Then
        '    If Not Session(pagecontroler1.SSSDTRID) Is Nothing Then Session(pagecontroler1.SSSDTRID) = Nothing
        'End If
        If Session("MyWrongTable") Is Nothing Then Exit Sub
        Dim dt As DataTable = Session("MyWrongTable")
        pagecontroler1.SSSDTRID = TIMS.GetRnd6Eng()
        pagecontroler1.PageDataTable = dt
        pagecontroler1.ControlerLoad()
    End Sub

End Class
