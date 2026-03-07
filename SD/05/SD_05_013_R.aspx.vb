Partial Class SD_05_013_R
    Inherits AuthBasePage

    'SD_05_013_R

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        'Dim dr As DataRow
        'conn = DbAccess.GetConnection
        msg.Text = ""
        Button2.Attributes("onclick") = "javascript:return search()"
        If Not IsPostBack Then
            RIDValue.Value = sm.UserInfo.RID
            '若只有管理一個班級，自動協助帶出班級--by AMU 2009-02
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Job As String = ""
        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)

        If Me.IsOnJob.SelectedIndex = 0 Then
            Job = "N"
        End If
        If Me.IsOnJob.SelectedIndex = 1 Then
            Job = "Y"
        End If

        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=Member&filename=SD_05_013_R&path=TIMS&STDate1=" & STDate1.Text & "&STDate2=" & STDate2.Text & "&FTDate1=" & FTDate1.Text & "&FTDate2=" & FTDate2.Text & "&OCID=" & OCIDValue1.Value & "&IsOnJob=" & Job & "' );" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
        Dim myvalue As String = ""
        myvalue += "&STDate1=" & STDate1.Text
        myvalue += "&STDate2=" & STDate2.Text
        myvalue += "&FTDate1=" & FTDate1.Text
        myvalue += "&FTDate2=" & FTDate2.Text
        myvalue += "&OCID=" & OCIDValue1.Value
        myvalue += "&IsOnJob=" & Job
        myvalue += "&PLANID=" & sm.UserInfo.PlanID
        myvalue += "&CJOB_UNKEY=" & Me.cjobValue.Value
        ReportQuery.PrintReport(Me, "Member", "SD_05_013_R", myvalue)
    End Sub

End Class
