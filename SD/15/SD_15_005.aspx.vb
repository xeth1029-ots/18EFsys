Partial Class SD_15_005
    Inherits AuthBasePage

    '1.轄區 2.縣市 3.訓練職類 4.課程類別
    'SD_15_005_1
    'SD_15_005_2
    'SD_15_005_3
    'SD_15_005_4

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            DistID = TIMS.Get_DistID(DistID)
            Button1.Attributes("onclick") = "OpenOrg('" & sm.UserInfo.TPlanID & "');"
        End If

    End Sub

End Class
