Partial Class TR_04_016_R
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            DistID = TIMS.Get_DistID(DistID)
            DistID.SelectedValue = sm.UserInfo.DistID
        End If

        Button1.Attributes("onclick") = "return print();"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim FTDate3 As String = FormatDateTime(DateAdd(DateInterval.Month, -3, Now.Date), 2)
        ReportQuery.PrintReport(Me, "TR", "TR_04_016_R", "FTDate1=" & FTDate1.Text & "&FTDate2=" & FTDate2.Text & "&FTDate3=" & FTDate3 & "&TPlanID=" & sm.UserInfo.TPlanID & "&DistID=" & DistID.SelectedValue)
    End Sub
End Class
