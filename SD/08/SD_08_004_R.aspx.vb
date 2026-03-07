Partial Class SD_08_004_R
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button1.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Dim s_ReportScript_Js As String = ReportQuery.ReportScript(Me, "MultiBlock", "Subsidy_Report", "RID='+document.getElementById('RIDValue').value+'&OCID='+document.getElementById('OCIDValue').value+'")
        Button_print.Attributes("onclick") = String.Concat("if(ReportPrint()){", s_ReportScript_Js, "}return false;")

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
    End Sub

    Private Sub Button_print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_print.Click
        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=MultiBlock&path=TIMS&filename=Subsidy_Report&RID=" & RIDValue.Value & "&OCID=" & OCIDValue.Value & "');" + vbCrLf
        'strScript += "</script>"

        'Page.RegisterStartupScript("window_onload", strScript)
    End Sub
End Class
