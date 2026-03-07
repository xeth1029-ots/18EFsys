Partial Class CP_05_002_R
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not Page.IsPostBack Then

            yearlist = TIMS.GetSyear(yearlist)
            Common.SetListItem(yearlist, sm.UserInfo.Years)

            planlist = TIMS.Get_TPlan(planlist, , 1)
            Common.SetListItem(planlist, sm.UserInfo.TPlanID)

            Print.Attributes("onclick") = "CheckPrint();return false;"
        End If
    End Sub

End Class
