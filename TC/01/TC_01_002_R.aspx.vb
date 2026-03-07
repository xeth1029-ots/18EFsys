Partial Class TC_01_002_R
    Inherits AuthBasePage

    'TC_01_002_R

    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not Me.IsPostBack Then
            yearlist = TIMS.GetSyear(yearlist)
            '2005/4/8--Ellen年度帶預設值Ellen 
            yearlist.SelectedValue = DateTime.Today.Year
            Common.SetListItem(Me.yearlist, sm.UserInfo.TPlanID)
            Me.planlist = TIMS.Get_TPlan(Me.planlist) '取得訓練計畫
            Me.DistrictList = TIMS.Get_DistID(Me.DistrictList, Nothing, objconn) '取得轄區
            Me.OrgTypeList = TIMS.Get_OrgType(Me.OrgTypeList, objconn) '取得機構別
            Common.SetListItem(Me.planlist, sm.UserInfo.TPlanID)
            Common.SetListItem(Me.DistrictList, sm.UserInfo.DistID)
            If sm.UserInfo.LID <> "000" Then
                Me.planlist.Enabled = False
                DistrictList.Enabled = False
            End If
            btnPrint.Attributes("onclick") = "return ReportPrint();"
        End If
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim sMyValue As String = ""
        sMyValue = ""
        sMyValue &= "&Years=" & Me.yearlist.SelectedValue
        sMyValue &= "&TPlanID=" & Me.planlist.SelectedValue
        sMyValue &= "&DistID=" & Me.DistrictList.SelectedValue
        sMyValue &= "&CTID=" & Me.city_code.Value
        sMyValue &= "&ComIDNO=" & Me.ComIDNO.Text
        sMyValue &= "&OrgKind=" & Me.OrgTypeList.SelectedValue
        sMyValue &= "&OrgName=" & Server.UrlEncode(Me.OrgName.Text)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "MultiBlock", "TC_01_002_R", sMyValue)
    End Sub
End Class