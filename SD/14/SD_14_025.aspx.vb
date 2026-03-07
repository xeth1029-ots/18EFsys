Partial Class SD_14_025
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SD_14_025" '2024

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            Dim s_javascript_btn2 As String = ""
            Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
            s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
            Button2.Attributes("onclick") = s_javascript_btn2
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        '列印
        'btnPrint1.Attributes("onclick") = "return CheckPrint();"
    End Sub

    Protected Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim v_DistID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        If v_DistID = "" Then v_DistID = sm.UserInfo.DistID
        If v_DistID = "000" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim prtstrVal As String = ""
        prtstrVal &= "&DistID=" & v_DistID ' sm.UserInfo.DistID
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, prtstrVal)
    End Sub
End Class