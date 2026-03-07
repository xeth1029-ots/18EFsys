Partial Class SD_05_018_R
    Inherits AuthBasePage

    'SD_05_018_R
    Const cst_printFN1 As String = "SD_05_018_R"
    'Dim objconn As SqlConnection
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End
        'objconn = DbAccess.GetConnection()

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button1.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If
    End Sub

    Sub Utl_Print1()
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim MyValue As String = ""
        MyValue &= "RID=" & RIDValue.Value
        MyValue &= "&Years=" & sm.UserInfo.Years
        ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        Utl_Print1()
    End Sub
End Class