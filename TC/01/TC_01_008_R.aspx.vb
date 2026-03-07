Partial Class TC_01_008_R
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'objconn = DbAccess.GetConnection()
        'AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If

        Dim s_ReportScript_JS1 As String = ReportQuery.ReportScript(Me, "Teach_Report", String.Concat("PlanID=", sm.UserInfo.PlanID, "&RID='+document.getElementById('RIDValue').value+'&OCID='+document.getElementById('OCIDValue1').value+'&TMID='+document.getElementById('TMIDValue1').value+'"))
        Button1.Attributes("onclick") = String.Concat("if(ReportPrint()){", s_ReportScript_JS1, "}return false;")

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        If center.Text = sm.UserInfo.OrgName Then
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim PlanID As String
        'PlanID = sm.UserInfo.PlanID
        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=MultiBlock&path=TIMS&filename=Teach_Report&PlanID=" & PlanID & "&RID=" & RIDValue.Value & "&OCID=" & OCIDValue1.Value & "&TMID=" & TMIDValue1.Value & "');" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
        '&PlanID=" & PlanID & "&RID=" & RIDValue.Value & "&OCID=" & OCIDValue1.Value & "&TMID=" & TMIDValue1.Value & "');" + vbCrLf
        Dim sMyValue As String = ""
        sMyValue &= "&PlanID=" & sm.UserInfo.PlanID
        sMyValue &= "&RID=" & RIDValue.Value
        sMyValue &= "&OCID=" & OCIDValue1.Value
        sMyValue &= "&TMID=" & TMIDValue1.Value
        ReportQuery.PrintReport(Me, "MultiBlock", "Teach_Report", sMyValue)
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class
