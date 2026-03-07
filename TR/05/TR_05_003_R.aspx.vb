Partial Class TR_05_003_R
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            CreateItem()
        End If

        Button1.Attributes("onclick") = "return search();"
    End Sub

    Sub CreateItem()
        Syear = TIMS.GetSyear(Syear)
        Common.SetListItem(Syear, Now.Year)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ReportQuery.PrintReport(Me, "TR", "TR_05_003_R", "Years=" & Syear.SelectedValue & "&PlanType=" & PlanType.SelectedValue & "&PlanTypeName=" & PlanType.SelectedItem.Text)
    End Sub
End Class
